# -*- coding: utf-8 -*-
"""工具函数：日期解析、地市/区县/街道解析、监测方法识别、风险分级、地址最小区分后缀等。"""
import datetime as dtm
import re

import pandas as pd

from . import config as C


# ---------------------------------------------------------------------------
# 日期解析
# ---------------------------------------------------------------------------
def parse_date_column(series):
    """将“监测时间”列解析为 datetime64（归一化到当天 0 点）。

    支持 yyyy/mm/dd、yyyy-mm-dd、yyyy年mm月dd日、Excel 日期对象及 Excel 序列号。
    """
    s = pd.Series(series)
    if pd.api.types.is_numeric_dtype(s):
        num = pd.to_numeric(s, errors="coerce")
        return (pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")).dt.normalize()
    dt = pd.to_datetime(s, errors="coerce")
    if dt.notna().any():
        return dt.dt.normalize()
    num = pd.to_numeric(s, errors="coerce")
    if num.notna().any():
        return (pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")).dt.normalize()
    return dt.dt.normalize()


def validate_date(year, month, day):
    """校验年月日并返回 datetime.date，非法时抛出 ValueError。"""
    try:
        y, m, d = int(year), int(month), int(day)
    except (TypeError, ValueError):
        raise ValueError("年、月、日必须为整数。")
    if not (2000 <= y <= 2100):
        raise ValueError(f"年份 {y} 超出合理范围（2000~2100）。")
    if not (1 <= m <= 12):
        raise ValueError(f"月份 {m} 不合法（应为 1~12）。")
    import calendar
    max_day = calendar.monthrange(y, m)[1]
    if not (1 <= d <= max_day):
        raise ValueError(f"日期 {y}-{m}-{d} 不合法（{m} 月只有 {max_day} 天）。")
    return dtm.date(y, m, d)


# ---------------------------------------------------------------------------
# 监测方法识别
# ---------------------------------------------------------------------------
def norm_method(value):
    """去除空白并转为小写，便于匹配方法别名。"""
    if value is None:
        return ""
    if isinstance(value, float) and value != value:  # NaN
        return ""
    return str(value).strip().replace(" ", "").lower()


def method_kind(value):
    """返回方法类别：'BI' / 'SSI' / 'ADI'，无法识别返回 None。"""
    m = norm_method(value)
    if m in ("布雷图指数bi", "bi"):
        return "BI"
    if m in ("标准间指数ssi", "ssi"):
        return "SSI"
    if m in ("成蚊密度指数法adi", "adi"):
        return "ADI"
    return None


# ---------------------------------------------------------------------------
# 键值（用于“一致”判断与去重）
# ---------------------------------------------------------------------------
def key_tuple(row):
    """把一行（仅含 KEY_COLS 的序列）转换为可比较的元组，NaN/None 统一为空串。"""
    out = []
    for v in row:
        if v is None or (isinstance(v, float) and v != v):
            out.append("")
        else:
            out.append(str(v).strip())
    return tuple(out)


# ---------------------------------------------------------------------------
# 地市-区县-街道 解析
# ---------------------------------------------------------------------------
def parse_region(region_str):
    """解析“广东省肇庆市-怀集县-大岗镇” -> (地市, 区县, 街道)。

    地市、区县、街道分别取第一段、中间段（可多段）、最后一段。
    """
    if region_str is None or (isinstance(region_str, float) and region_str != region_str):
        return "", "", ""
    parts = [p.strip() for p in str(region_str).split("-") if p.strip()]
    if not parts:
        return "", "", ""
    city = parts[0]
    street = parts[-1]
    if len(parts) >= 3:
        district = "-".join(parts[1:-1])
    elif len(parts) == 2:
        district = parts[1]
    else:
        district = ""
    return city, district, street


def _strip_province(city):
    """去掉省份前缀：广东省肇庆市 -> 肇庆市。"""
    c = str(city).strip()
    for marker in ("特别行政区", "自治区", "省"):
        if marker in c:
            c = c[c.index(marker) + len(marker):]
            break
    return c


def city_with_suffix(city):
    """地市带“市”后缀：广东省肇庆市 -> 肇庆市。"""
    c = _strip_province(city)
    if c and not c.endswith("市"):
        c = c + "市"
    return c


def city_short(city):
    """地市不带“市”后缀：广东省肇庆市 -> 肇庆。"""
    return city_with_suffix(city).rstrip("市")


def district_short(district):
    """区县转换：市辖区 -> /；去除“市/县/区”字。天河区 -> 天河。"""
    d = str(district).strip()
    if d in ("", "nan"):
        return ""
    d = d.replace("市辖区", "/")
    for ch in ("市", "县", "区"):
        d = d.replace(ch, "")
    return d


# ---------------------------------------------------------------------------
# 风险分级
# ---------------------------------------------------------------------------
def risk_level(value):
    """返回“安全 / 低风险 / 中风险 / 高风险”。"""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "安全"
    if v < C.TH_LOW:
        return "安全"
    if v < C.TH_MID:
        return "低风险"
    if v < C.TH_HIGH:
        return "中风险"
    return "高风险"


def risk_level_word(value):
    """日报风险列表用的等级词：低 / 中 / 高。"""
    lv = risk_level(value)
    return {"低风险": "低", "中风险": "中", "高风险": "高"}.get(lv, "")


# ---------------------------------------------------------------------------
# 3.2.5 地址最小区分后缀
# ---------------------------------------------------------------------------
_ADDR_MARKERS = [
    "广场", "大厦", "大楼", "商场", "市场", "公园", "酒店", "医院", "学校",
    "路", "街", "道", "巷", "小区", "花园", "苑", "园", "桥", "馆", "城",
    "镇", "村", "站", "场", "楼", "里", "号",
]


def _clean_addr_segments(addr, remove_parts):
    """去除与地市/区县/街道/村居重复的部分，并按标点切分为若干段。"""
    segs = []
    for s in re.split(r"[\s，。、,．.；;：:]", str(addr)):
        t = s
        for p in remove_parts:
            if p:
                t = t.replace(str(p), "")
        t = t.strip(" \t，。、,．.；;：:（）()“”\"'")
        if t:
            segs.append(t)
    return segs


def minimal_suffix(addr, others, remove_parts):
    """计算能区分本地址的最小区分后缀。

    addr      : 本记录的有效地址
    others    : 同组其它地址（列表）
    remove_parts : 需要从地址中去除的重复部分（地市、区县、街道、村居等）
    """
    segs = _clean_addr_segments(addr, remove_parts)
    if not segs:
        return str(addr)
    cleaned = "".join(segs)

    others_clean = set()
    for o in others:
        os_ = _clean_addr_segments(o, remove_parts)
        oc = "".join(os_)
        if oc and oc != cleaned:
            others_clean.add(oc)

    # 候选：按段的后缀 与 单个段（优先包含地点标志词、且不与其他地址重复的最短段）
    cands = set()
    for i in range(len(segs)):
        cands.add("".join(segs[i:]))
    for s in segs:
        cands.add(s)
    cands = sorted(cands, key=lambda c: (len(c), c))
    for cand in cands:
        if len(cand) >= 2 and any(m in cand for m in _ADDR_MARKERS) \
                and not any(cand in o for o in others_clean):
            return cand
    # 回退：最短的字符级后缀（至少 2 个字符），不与其他地址重复
    for length in range(2, len(cleaned) + 1):
        suf = cleaned[-length:]
        if not any(suf in o for o in others_clean):
            return suf
    return cleaned


def remove_parts_for(region, village):
    """生成地址清理时需要去除的部分：地市、区县、街道、村居及完整地市串。"""
    city, district, street = parse_region(region)
    parts = [city, district, street, str(region)]
    cs = city_with_suffix(city)
    if cs and cs != city:
        parts.append(cs)
    parts.append(village)
    seen = []
    for p in parts:
        if p and str(p).strip() not in ("", "nan") and p not in seen:
            seen.append(p)
    return sorted(seen, key=lambda p: len(str(p)), reverse=True)
