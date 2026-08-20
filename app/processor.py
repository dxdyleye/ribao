# -*- coding: utf-8 -*-
"""核心数据处理流水线：全局预处理、BI+SSI 专项处理、ADI 专项处理。"""
import datetime as dtm
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from . import config as C
from . import utils as U


class ProcessingError(Exception):
    """处理过程中可向用户展示的错误。"""


class NoDataError(ProcessingError):
    """目标日期无数据。"""


class ColumnError(ProcessingError):
    """列名不匹配。"""


BI_SHEET_ORDER = [
    "BI+SSI表",
    "BI+SSI(不重复)",
    "BI+SSI(重复)+取较大值处理",
    "重复数据删除",
    "地址区分处理",
    "最终表",
]
ADI_SHEET_ORDER = [
    "ADI表",
    "ADI(不重复)",
    "ADI(重复)+取较大值处理",
    "ADI(重复数据删除)",
    "ADI(地址区分处理)",
    "ADI(最终表)",
]


@dataclass
class PipelineResult:
    """一次处理的所有输出数据。"""
    target: dtm.date
    exclude_field: str
    raw: pd.DataFrame                       # 原始数据（含全部输入列）
    base: pd.DataFrame                      # 基础数据集
    bi_sheets: dict = field(default_factory=dict)
    bi_final: pd.DataFrame = field(default_factory=pd.DataFrame)
    adi_sheets: dict = field(default_factory=dict)
    adi_final: pd.DataFrame = field(default_factory=pd.DataFrame)
    deletions: pd.DataFrame = field(default_factory=pd.DataFrame)  # 删除数据情况说明


# ---------------------------------------------------------------------------
# 读取与校验
# ---------------------------------------------------------------------------
def load_and_validate(path):
    """读取源 Excel 并校验必需列，返回仅含必需列的 DataFrame（含 _row_id）。"""
    try:
        raw = pd.read_excel(path, dtype=object)
    except Exception as exc:  # noqa: BLE001
        raise ProcessingError(f"无法读取Excel文件“{path}”：{exc}") from exc
    if raw is None or len(raw) == 0 or len(raw.columns) == 0:
        raise ColumnError("Excel文件中没有数据（请确保第一行为表头且包含必需列）。")
    cols = [str(c).strip() for c in raw.columns]
    raw.columns = cols
    missing = [c for c in C.REQUIRED_COLUMNS if c not in cols]
    if missing:
        raise ColumnError(
            "源Excel文件缺少以下必需列：\n" + "\n".join("· " + m for m in missing)
            + "\n\n请检查列名是否与要求完全一致（包括括号、引号、顿号等字符）。"
        )
    df = raw[C.REQUIRED_COLUMNS].copy()
    df[C.ID_COL] = np.arange(len(df))
    return df


# ---------------------------------------------------------------------------
# 3.1 全局预处理
# ---------------------------------------------------------------------------
def preprocess(raw, target, exclude_field=None):
    """执行 3.1 全局预处理，返回 (基础数据集, 删除记录列表)。"""
    df = raw[C.REQUIRED_COLUMNS].copy()
    df[C.ID_COL] = np.arange(len(df))
    deleted = []

    def drop_rows(mask, reason):
        """删除 mask 为 True 的行，并记录删除原因（记录源文件原始值）。"""
        nonlocal df
        ids = df.loc[mask, C.ID_COL].tolist()
        for rid in ids:
            row = raw.iloc[rid]
            rec = {c: row[c] for c in C.REQUIRED_COLUMNS}
            rec[C.ID_COL] = rid
            rec["删除原因"] = reason
            deleted.append(rec)
        df = df[~mask].copy()

    # 11. 按目标日期筛选
    dt = U.parse_date_column(df[C.COL_TIME])
    t = pd.Timestamp(target)
    drop_rows(dt.isna() | (dt.dt.normalize() != t), "监测时间与目标日期不符")
    if df.empty:
        raise NoDataError(
            f"目标日期 {target.strftime('%Y-%m-%d')} 无任何监测数据，请检查日期或源文件内容。"
        )

    # 12. 删除监测指标值为空的记录
    val = pd.to_numeric(df[C.COL_VALUE], errors="coerce")
    drop_rows(val.isna(), "监测指标值为空")

    # 13. 监测指标值四舍五入保留 1 位小数
    df[C.COL_VALUE] = val.round(1)

    # 14. 排除字段过滤
    if exclude_field:
        drop_rows(
            df[C.COL_REGION].astype(str).str.contains(exclude_field, na=False),
            "被排除字段过滤",
        )

    # 15. 距末例天数：<=5 或 >40000
    days = pd.to_numeric(df[C.COL_DAYS], errors="coerce")
    keep = (days <= C.DAYS_LE) | (days > C.DAYS_GT)
    drop_rows(~keep, "距末例天数不在范围内")

    # 16. 仅保留核心区 / 警戒区
    zone = df[C.COL_ZONE].astype(str).str.strip()
    drop_rows(~zone.isin(list(C.ZONES)), "防控区类型不符合")

    # 17. 新增“监测地点”列
    village = df[C.COL_VILLAGE].astype(str).str.strip()
    df["监测地点"] = village + "（" + zone + "）"

    # 内部辅助列
    df[C.KIND_COL] = df[C.COL_METHOD].map(U.method_kind)
    df[C.KEY_COL] = df[C.KEY_COLS].apply(U.key_tuple, axis=1)

    if df.empty:
        raise NoDataError("目标日期下没有符合条件的数据，请检查排除字段及筛选条件。")
    return df, deleted


# ---------------------------------------------------------------------------
# 地址区分处理（3.2.5，BI 与 ADI 共用）
# ---------------------------------------------------------------------------
def address_distinguish(df):
    """对 (地市, 社区/村居, 防控区类型) 相同但地址不同的记录修改“监测地点”。

    返回 (地址区分处理表, 处理后的完整数据集)。
    """
    if df is None or df.empty:
        out = pd.DataFrame(columns=C.SHEET1_COLS + ["备注"])
        out[C.FILL_COL] = None
        return out, df.copy() if df is not None else pd.DataFrame(columns=C.SHEET1_COLS)

    def _safe_str(v):
        if v is None or (isinstance(v, float) and v != v):
            return ""
        return str(v).strip()

    def eff_addr(row):
        a1 = _safe_str(row[C.COL_ADDR1])
        a2 = _safe_str(row[C.COL_ADDR2])
        if a1 and a1 != "nan":
            return a1
        if a2 and a2 != "nan":
            return a2
        return ""

    df = df.copy()
    df["_eff"] = df.apply(eff_addr, axis=1)

    parts = []
    for (region, village, zone), g in df.groupby([C.COL_REGION, C.COL_VILLAGE, C.COL_ZONE], dropna=False):
        combos = g.groupby("_eff", dropna=False)
        combo_list = [(name, cg) for name, cg in combos]
        if len(combo_list) <= 1:
            rows = g.drop(columns=["_eff"]).copy()
            rows["备注"] = ""
            rows[C.FILL_COL] = None
            parts.append(rows)
            continue
        remove_parts = U.remove_parts_for(region, village)
        suffix_map = {}
        for name, _cg in combo_list:
            if name is None or (isinstance(name, float) and name != name) or str(name).strip() in ("", "nan"):
                suffix_map[name] = ""
                continue
            others = [n for n, _ in combo_list if n != name
                      and not (n is None or (isinstance(n, float) and n != n) or str(n).strip() in ("", "nan"))]
            suffix_map[name] = U.minimal_suffix(name, others, remove_parts)
        for name, cg in combo_list:
            rows = cg.drop(columns=["_eff"]).copy()
            suf = suffix_map.get(name, "")
            if suf:
                old_loc = rows["监测地点"].astype(str)
                rows["监测地点"] = f"{village}{suf}（{zone}）"
                rows["备注"] = "原监测地点：" + old_loc
            else:
                rows["备注"] = ""
            rows[C.FILL_COL] = None
            parts.append(rows)

    out = pd.concat(parts, ignore_index=True)
    out = out[C.SHEET1_COLS + ["备注"] + [C.ID_COL]].copy()
    out[C.FILL_COL] = None
    return out, out[C.SHEET1_COLS].copy()


def _empty_sheet(extra=()):
    d = pd.DataFrame(columns=list(C.SHEET1_COLS) + list(extra) + [C.ID_COL, C.KEY_COL])
    d[C.FILL_COL] = None
    return d


def _dedupe_keep_max(merged, deleted, raw, label):
    """按完整键值去重、每组保留最大值；被删除的行标红并记入删除日志。"""
    if merged.empty:
        sheet = _empty_sheet()
        return sheet, merged.copy()
    merged = merged.copy()
    merged["_rank"] = merged.groupby(C.KEY_COL, dropna=False)[C.COL_VALUE].rank(
        method="first", ascending=False
    )
    keep_mask = merged["_rank"] == 1
    sheet = merged[C.SHEET1_COLS].copy()
    sheet[C.FILL_COL] = np.where(keep_mask, None, "red")
    # 记录被删除的行（以源文件原始内容为准）
    for rid in merged.loc[~keep_mask, C.ID_COL].tolist():
        rid = int(rid)
        row = raw.loc[rid]
        rec = {c: row[c] for c in C.REQUIRED_COLUMNS}
        rec[C.ID_COL] = rid
        rec["删除原因"] = f"重复数据保留最大值（{label}）"
        deleted.append(rec)
    kept = merged[keep_mask].copy()
    return sheet, kept


# ---------------------------------------------------------------------------
# 3.2 BI+SSI 专项处理
# ---------------------------------------------------------------------------
def bi_pipeline(base, deleted):
    """执行 3.2.1 ~ 3.2.6，返回 (sheet字典, 最终BI表, 删除记录追加)。"""
    sub = base[base[C.KIND_COL].isin(("BI", "SSI"))].copy()
    ssi_mask = sub[C.KIND_COL] == "SSI"
    sub[C.ORIG_COL] = sub[C.COL_VALUE]  # 原始值（SSI 为 ×2 前）
    # 方法名统一为规范写法
    sub[C.COL_METHOD] = np.where(ssi_mask, C.METHOD_SSI, C.METHOD_BI)
    sub.loc[ssi_mask, C.COL_VALUE] = sub.loc[ssi_mask, C.COL_VALUE] * 2  # SSI 转换

    # ---- Sheet1: BI+SSI表 ----
    sheet1 = sub[C.SHEET1_COLS].copy()
    sheet1[C.FILL_COL] = None

    bi = sub[~ssi_mask]
    ssi = sub[ssi_mask]
    bi_keys = set(bi[C.KEY_COL]) if len(bi) else set()
    ssi_keys = set(ssi[C.KEY_COL]) if len(ssi) else set()
    dup_keys = bi_keys & ssi_keys

    # ---- Sheet2: BI+SSI(不重复) ----
    bi_alone = bi[~bi[C.KEY_COL].isin(ssi_keys)]
    ssi_alone = ssi[~ssi[C.KEY_COL].isin(bi_keys)].copy()
    ssi_alone[C.COL_METHOD] = C.METHOD_BI  # 合并为 BI
    sheet2 = pd.concat([bi_alone, ssi_alone], ignore_index=True)
    sheet2 = sheet2[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]].copy()
    ssi_alone_ids = set(ssi_alone[C.ID_COL].tolist()) if len(ssi_alone) else set()
    sheet2[C.FILL_COL] = sheet2[C.ID_COL].map(lambda rid: "yellow" if rid in ssi_alone_ids else None)

    # ---- Sheet3: BI+SSI(重复)+取较大值处理 ----
    dup_records = []
    for k in sorted(dup_keys, key=str):
        k_bi = bi[bi[C.KEY_COL] == k]
        k_ssi = ssi[ssi[C.KEY_COL] == k]
        orig_bi = float(k_bi[C.COL_VALUE].max())
        orig_ssi = float(k_ssi[C.ORIG_COL].max())
        conv_ssi = float(k_ssi[C.COL_VALUE].max())
        final_val = max(orig_bi, conv_ssi)
        r = k_bi.iloc[0].copy()
        r[C.COL_VALUE] = final_val
        r[C.COL_METHOD] = C.METHOD_BI
        r["原BI值"] = round(orig_bi, 1)
        r["原SSI值"] = round(orig_ssi, 1)
        r["转换后的SSI值"] = round(conv_ssi, 1)
        dup_records.append(r)
    if dup_records:
        sheet3 = pd.DataFrame(dup_records)
        sheet3 = sheet3[C.SHEET1_COLS + ["原BI值", "原SSI值", "转换后的SSI值"]
                        + [C.ID_COL, C.KEY_COL]].copy()
    else:
        sheet3 = _empty_sheet(["原BI值", "原SSI值", "转换后的SSI值"])
    sheet3[C.FILL_COL] = None

    # ---- Sheet4: 重复数据删除（完整键值去重，保留最大值） ----
    merged = pd.concat([
        sheet2[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]],
        sheet3[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]],
    ], ignore_index=True)
    sheet4, kept = _dedupe_keep_max(merged, deleted, base, "BI")

    # ---- Sheet5: 地址区分处理 ----
    sheet5, final = address_distinguish(kept)

    # ---- Sheet6: 最终表 ----
    sheet6 = final[C.SHEET1_COLS].copy() if len(final) else _empty_sheet()
    sheet6[C.FILL_COL] = None

    sheets = {
        "BI+SSI表": sheet1,
        "BI+SSI(不重复)": sheet2,
        "BI+SSI(重复)+取较大值处理": sheet3,
        "重复数据删除": sheet4,
        "地址区分处理": sheet5,
        "最终表": sheet6,
    }
    return sheets, final


# ---------------------------------------------------------------------------
# 3.3 ADI 专项处理（复用 BI 的去重、取最大值、地址区分逻辑）
# ---------------------------------------------------------------------------
def adi_pipeline(base, deleted):
    """执行 ADI 处理，返回 (sheet字典, 最终ADI表, 删除记录追加)。"""
    sub = base[base[C.KIND_COL] == "ADI"].copy()
    if len(sub):
        sub[C.COL_METHOD] = C.METHOD_ADI  # 方法名统一为规范写法

    # ---- ADI表 ----
    sheet_all = sub[C.SHEET1_COLS].copy()
    sheet_all[C.FILL_COL] = None

    if sub.empty:
        sheets = {
            "ADI表": sheet_all,
            "ADI(不重复)": _empty_sheet(),
            "ADI(重复)+取较大值处理": _empty_sheet(["原ADI值"]),
            "ADI(重复数据删除)": _empty_sheet(),
            "ADI(地址区分处理)": _empty_sheet(["备注"]),
            "ADI(最终表)": _empty_sheet(),
        }
        return sheets, pd.DataFrame(columns=C.SHEET1_COLS)

    counts = sub[C.KEY_COL].value_counts()
    dup_keys = set(counts[counts > 1].index)
    unique = sub[~sub[C.KEY_COL].isin(dup_keys)]

    # ---- ADI(不重复) ----
    sheet_nodup = unique[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]].copy()
    sheet_nodup[C.FILL_COL] = None

    # ---- ADI(重复)+取较大值处理 ----
    dup_records = []
    for k in sorted(dup_keys, key=str):
        rows = sub[sub[C.KEY_COL] == k]
        vals = rows[C.COL_VALUE].tolist()
        best = max(vals)
        r = rows.iloc[0].copy()
        r[C.COL_VALUE] = best
        r["原ADI值"] = "、".join(f"{v:.1f}" for v in vals)
        dup_records.append(r)
    if dup_records:
        sheet_dup = pd.DataFrame(dup_records)
        sheet_dup = sheet_dup[C.SHEET1_COLS + ["原ADI值"] + [C.ID_COL, C.KEY_COL]].copy()
    else:
        sheet_dup = _empty_sheet(["原ADI值"])
    sheet_dup[C.FILL_COL] = None

    # ---- ADI(重复数据删除) ----
    merged = pd.concat([
        sheet_nodup[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]],
        sheet_dup[C.SHEET1_COLS + [C.ID_COL, C.KEY_COL]],
    ], ignore_index=True)
    sheet_dedup, kept = _dedupe_keep_max(merged, deleted, base, "ADI")

    # ---- ADI(地址区分处理) 与 ADI(最终表) ----
    sheet_addr, final = address_distinguish(kept)
    sheet_final = final[C.SHEET1_COLS].copy() if len(final) else _empty_sheet()
    sheet_final[C.FILL_COL] = None

    sheets = {
        "ADI表": sheet_all,
        "ADI(不重复)": sheet_nodup,
        "ADI(重复)+取较大值处理": sheet_dup,
        "ADI(重复数据删除)": sheet_dedup,
        "ADI(地址区分处理)": sheet_addr,
        "ADI(最终表)": sheet_final,
    }
    return sheets, final


# ---------------------------------------------------------------------------
# 总入口
# ---------------------------------------------------------------------------
def process_all(input_path, year, month, day, exclude_field=None, log=None):
    """完整处理流程，返回 PipelineResult。"""
    def say(msg, step=None):
        if log:
            log(msg, step)

    say("正在读取Excel文件并校验列名…", 1)
    raw = load_and_validate(input_path)

    say("正在校验目标日期…", 2)
    target = U.validate_date(year, month, day)
    if exclude_field is None or str(exclude_field).strip() == "":
        exclude_field = ""
    else:
        exclude_field = str(exclude_field).strip()

    say("正在执行全局预处理（筛选日期、空值、排除字段、距末例天数、防控区类型）…", 3)
    base, deleted = preprocess(raw, target, exclude_field or None)

    say("正在处理BI+SSI专项数据（转换、合并、去重、地址区分）…", 4)
    bi_sheets, bi_final = bi_pipeline(base, deleted)

    say("正在处理ADI专项数据（去重、取最大值、地址区分）…", 5)
    adi_sheets, adi_final = adi_pipeline(base, deleted)

    deletions = _deletions_frame(deleted)
    return PipelineResult(
        target=target,
        exclude_field=exclude_field,
        raw=raw,
        base=base,
        bi_sheets=bi_sheets,
        bi_final=bi_final,
        adi_sheets=adi_sheets,
        adi_final=adi_final,
        deletions=deletions,
    )


def _deletions_frame(deleted):
    """删除记录 -> DataFrame（按源文件顺序排列）。"""
    if not deleted:
        return pd.DataFrame(columns=C.REQUIRED_COLUMNS + ["删除原因"])
    df = pd.DataFrame(deleted)
    if C.ID_COL in df.columns:
        df = df.sort_values(C.ID_COL, kind="stable")
    cols = [c for c in C.REQUIRED_COLUMNS if c in df.columns] + ["删除原因"]
    return df[cols].copy()
