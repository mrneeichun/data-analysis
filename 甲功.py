import pandas as pd


# ==================== 配置参数 ====================
# 甲功项目阈值配置：TSH / FT3 / FT4 / TT3 / TT4 / Anti-TPO / Anti-TG
# 阈值按照「偏低 / 正常 / 偏高」三段配置：
#   偏低：结果 < low_max
#   正常：normal_low ≤ 结果 < normal_high
#   偏高：结果 ≥ high_min
THYROID_RULES = {
    "TSH":       {"low_max": 0.30, "normal_low": 0.30, "normal_high": 4.3, "high_min": 4.3},
    "FT3":       {"low_max": 3.3, "normal_low": 3.3, "normal_high": 7.3, "high_min": 7.3},
    "FT4":       {"low_max": 12.06, "normal_low": 12.06, "normal_high": 23.0, "high_min": 23.0},
    "TT3":       {"low_max": 0.85, "normal_low": 0.85, "normal_high": 2.68, "high_min": 2.68},
    "TT4":       {"low_max": 60.72, "normal_low": 60.72, "normal_high": 170.0, "high_min": 170.0},
    "Anti-TPO":  {"low_max": 32, "normal_low": None, "normal_high": None, "high_min": 32},
    "Anti-TG":   {"low_max": 113, "normal_low": None, "normal_high": None, "high_min": 113},
}

# 单位配置（用于阈值设置界面显示）
UNITS = {
    "TSH": "μIU/L",
    "FT3": "pmol/L",
    "FT4": "pmol/L",
    "TT3": "nmol/L",
    "TT4": "nmol/L",
    "Anti-TPO": "IU/mL",
    "Anti-TG": "IU/mL",
}

DISPLAY_ORDER = list(THYROID_RULES.keys())

# ==================== 判定函数 ====================
def _judge_甲功(row, rules):
    """
    根据 THYROID_RULES 判定偏低/正常/偏高
    未配置或无法解析数值时返回 None（不参与占比统计）
    """
    p, v = row.get("测试项目"), row.get("结果_num")
    if pd.isna(v):
        return None
    
    # 获取阈值规则
    rule = rules.get(p) if isinstance(rules, dict) else None
    if not rule or not isinstance(rule, dict):
        return None
    
    low_max = rule.get("low_max")
    normal_low = rule.get("normal_low")
    normal_high = rule.get("normal_high")
    high_min = rule.get("high_min")
    
    # 判定正常区间
    if normal_low is not None and normal_high is not None:
        if normal_low <= v < normal_high:
            return "正常"
    
    # 判定偏高
    if high_min is not None and v >= high_min:
        return "偏高"
    
    # 判定偏低
    if low_max is not None and v < low_max:
        return "偏低"
    
    # 其他情况返回None
    return None


# ==================== 主分析函数 ====================
def analyze_甲功(df):
    """
    甲功数据分析主函数
    按项目、试剂批号统计 偏低/正常/偏高 占比
    返回：(占比统计表, 空DataFrame, None)
    """
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame(), None

    rules = THYROID_RULES
    df = df.copy()
    
    # -------------------- 步骤1：提取数值并判定 --------------------
    # 从结果列中提取数字
    df["结果_num"] = pd.to_numeric(df["结果"].astype(str).str.extract(r"(\d+\.?\d*)", expand=False), errors="coerce")
    df["判定结果"] = df.apply(lambda row: _judge_甲功(row, rules), axis=1)
    
    # 只统计有判定结果的数据
    calc_df = df[df["判定结果"].notna()].copy()
    calc_df = calc_df[calc_df["测试项目"].isin(DISPLAY_ORDER)]

    if calc_df.empty:
        summary = pd.DataFrame(columns=["测试项目", "试剂批号", "偏低", "偏低率", "正常", "正常率", "偏高", "偏高率"])
        return summary, pd.DataFrame(), None

    # -------------------- 步骤2：生成占比统计表 --------------------
    # 按项目、批号、结果分组统计
    summary = calc_df.groupby(["测试项目", "试剂批号", "判定结果"]).size().unstack(fill_value=0)
    
    # 确保所有必要列存在
    for c in ["偏低", "正常", "偏高"]:
        if c not in summary.columns:
            summary[c] = 0
    
    # 计算总数和百分比
    summary["总数"] = summary["偏低"] + summary["正常"] + summary["偏高"]
    summary["偏低率"] = (summary["偏低"] / summary["总数"] * 100).map("{:.2f}%".format)
    summary["正常率"] = (summary["正常"] / summary["总数"] * 100).map("{:.2f}%".format)
    summary["偏高率"] = (summary["偏高"] / summary["总数"] * 100).map("{:.2f}%".format)
    
    # 按配置顺序排序，选择需要的列
    summary = summary.reset_index()
    summary["sort"] = summary["测试项目"].apply(lambda x: DISPLAY_ORDER.index(x) if x in DISPLAY_ORDER else 99)
    summary = summary.sort_values(by=["sort", "试剂批号"]).drop(columns=["sort"])
    summary = summary[["测试项目", "偏低", "偏低率", "正常", "正常率", "偏高", "偏高率", "总数", "试剂批号"]]

    return summary, pd.DataFrame(), None