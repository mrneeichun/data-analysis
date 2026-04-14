import pandas as pd


# ==================== 配置参数 ====================
# 肿瘤标志物项目阈值配置
# 阈值按照「阴性 / 阳性」两段配置：
#   阴性：结果 < neg_max
#   阳性：结果 >= pos_min
TUMOR_RULES = {
    "AFP":        {"neg_max": 6.13, "pos_min": 6.13},
    "CEA":        {"neg_max": 4.1, "pos_min": 4.1},
    "CA 125":     {"neg_max": 37.22, "pos_min": 37.22},
    "CA 15-3":    {"neg_max": 27.96, "pos_min": 27.96},
    "CA 19-9":    {"neg_max": 27.5, "pos_min": 27.5},
    "tPSA":       {"neg_max": 4.03, "pos_min": 4.03},
    "fPSA":       {"neg_max": 0.83, "pos_min": 0.83},
    "Tg":         {"neg_max": 79.67, "pos_min": 79.67}, 
    "CA50":      {"neg_max": 21.962, "pos_min": 21.962},
    "CYFRA21-1":  {"neg_max": 3.24, "pos_min": 3.24},
    "CA 242":     {"neg_max": 10.0, "pos_min": 10.0},  
    "CA72-4":    {"neg_max": 6.9, "pos_min": 6.9},
    "SCC":        {"neg_max": 1.5, "pos_min": 1.5},
    "NSE":        {"neg_max": 16.01, "pos_min": 16.01},
    "HE4":        {"neg_max": 69.0, "pos_min": 69.0},
}

# 单位配置（用于阈值设置界面显示）
UNITS = {
    "AFP": "IU/mL",
    "CEA": "ng/mL",
    "CA 125": "U/mL",
    "CA 15-3": "U/mL",
    "CA 19-9": "U/mL",
    "tPSA": "ng/mL",
    "fPSA": "ng/mL",
    "Tg": "ng/mL",
    "CA50": "U/mL",
    "CYFRA21-1": "ng/mL",
    "CA 242": "U/mL",
    "CA72-4": "IU/mL",
    "SCC": "ng/mL",
    "NSE": "ng/mL",
    "HE4": "pmol/L",
}

DISPLAY_ORDER = list(TUMOR_RULES.keys())

# ==================== 判定函数 ====================
def _judge_肿瘤(row, rules):
    """
    根据 TUMOR_RULES 判定阴性/阳性
    未配置或无法解析数值时返回 None（不参与占比统计）
    """
    p, v = row.get("测试项目"), row.get("结果_num")
    if pd.isna(v):
        return None
    
    # 获取阈值规则
    rule = rules.get(p) if isinstance(rules, dict) else None
    if not rule or not isinstance(rule, dict):
        return None
    
    neg_max = rule.get("neg_max")
    pos_min = rule.get("pos_min")
    
    # 判定阴性
    if neg_max is not None and v < neg_max:
        return "阴性"
    
    # 判定阳性
    elif pos_min is not None and v >= pos_min:
        return "阳性"
    
    # 其他情况返回None
    return None

# ==================== 主分析函数 ====================
def analyze_肿瘤(df):
    """
    肿瘤标志物数据分析主函数
    按项目、试剂批号统计 阴性/阳性 占比
    返回：(占比统计表, 空DataFrame, None)
    """
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame(), None

    rules = TUMOR_RULES
    df = df.copy()
    
    # -------------------- 步骤1：提取数值并判定 --------------------
    # 从结果列中提取数字
    df["结果_num"] = pd.to_numeric(df["结果"].astype(str).str.extract(r"(\d+\.?\d*)", expand=False), errors="coerce")
    df["判定结果"] = df.apply(lambda row: _judge_肿瘤(row, rules), axis=1)
    
    # 只统计有判定结果的数据
    calc_df = df[df["判定结果"].notna()].copy()
    calc_df = calc_df[calc_df["测试项目"].isin(DISPLAY_ORDER)]

    if calc_df.empty:
        summary = pd.DataFrame(columns=["测试项目", "试剂批号", "阴性", "阴性率", "阳性", "阳性率"])
        return summary, pd.DataFrame(), None

    # -------------------- 步骤2：生成占比统计表 --------------------
    # 按项目、批号、结果分组统计
    summary = calc_df.groupby(["测试项目", "试剂批号", "判定结果"]).size().unstack(fill_value=0)
    
    # 确保所有必要列存在
    for c in ["阴性", "阳性"]:
        if c not in summary.columns:
            summary[c] = 0
    
    # 计算总数和百分比
    summary["总数"] = summary["阴性"] + summary["阳性"]
    summary["阴性率"] = (summary["阴性"] / summary["总数"] * 100).map("{:.2f}%".format)
    summary["阳性率"] = (summary["阳性"] / summary["总数"] * 100).map("{:.2f}%".format)
    
    # 按配置顺序排序，选择需要的列
    summary = summary.reset_index()
    summary["sort"] = summary["测试项目"].apply(lambda x: DISPLAY_ORDER.index(x) if x in DISPLAY_ORDER else 99)
    summary = summary.sort_values(by=["sort", "试剂批号"]).drop(columns=["sort"])
    summary = summary[["测试项目", "阳性", "阳性率", "总数", "试剂批号"]]

    return summary, pd.DataFrame(), None