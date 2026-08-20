# -*- coding: utf-8 -*-
"""全局常量配置：列名、监测方法、地市顺序、风险阈值等。"""

# ============ 输入列名（必须与源文件完全一致） ============
COL_TIME = "监测时间（年/月/日）"
COL_VALUE = "监测指标值"
COL_REGION = "地市-区/县/市-街道/乡/镇"
COL_DAYS = "距末例天数（自动计算）"
COL_ZONE = "防控区类型"
COL_VILLAGE = "社区/村居"
COL_METHOD = "监测方法（如BI/RI/MOI/ADI等）"
COL_ADDR1 = "监测地址（地图定位版）"
COL_ADDR2 = "监测地址（如“监测地址”定位字段不可用，可手填；如定位可用，不需要重复填写）"

REQUIRED_COLUMNS = [
    COL_TIME, COL_VALUE, COL_REGION, COL_DAYS, COL_ZONE,
    COL_VILLAGE, COL_METHOD, COL_ADDR1, COL_ADDR2,
]

# ============ 监测方法 ============
METHOD_BI = "布雷图指数BI"
METHOD_SSI = "标准间指数SSI"
METHOD_ADI = "成蚊密度指数法ADI"

# ============ 防控区类型 ============
ZONES = ("核心区", "警戒区")

# ============ 计算过程表的固定列 ============
KEY_COLS = [COL_REGION, COL_VILLAGE, COL_ADDR1, COL_ADDR2, COL_ZONE]
SHEET1_COLS = [
    COL_REGION, COL_VILLAGE, COL_ADDR1, COL_ADDR2, COL_ZONE,
    "监测地点", COL_METHOD, COL_VALUE,
]

# ============ 地市顺序（日报与汇总表的排序依据） ============
REGION_ORDER = [
    "广州", "深圳", "珠海", "汕头", "佛山", "韶关", "河源", "梅州", "惠州",
    "汕尾", "东莞", "中山", "江门", "阳江", "湛江", "茂名", "肇庆", "清远",
    "潮州", "揭阳", "云浮",
]

# ============ 风险分级阈值 ============
TH_LOW = 5.0    # <5 安全
TH_MID = 10.0   # 5~10 低风险
TH_HIGH = 20.0  # 10~20 中风险；>=20 高风险

# ============ 距末例天数筛选条件 ============
DAYS_LE = 5      # 保留 <=5
DAYS_GT = 40000  # 或 >40000

# ============ 内部辅助列（不写入输出） ============
ID_COL = "_row_id"       # 原始行号（用于删除日志溯源）
KEY_COL = "_key"         # 完整键值元组
KIND_COL = "_method_kind"  # 方法类别：BI / SSI / ADI / None
FILL_COL = "_FILL"       # 单元格填充：yellow / red / None
ORIG_COL = "_orig_value" # SSI 转换前的原始值

# ============ Excel 填充颜色（8位ARGB，FF为不透明） ============
FILL_YELLOW = "FFFFEB9C"
FILL_RED = "FFFF7C80"
FILL_GREEN = "FFC6EFCE"
FILL_ORANGE = "FFFFC000"
FILL_HEADER = "FFDDEBF7"
