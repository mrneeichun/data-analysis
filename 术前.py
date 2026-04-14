import pandas as pd


# ==================== 配置参数 ====================
# 阈值配置规则：判读规则唯一来源
# 注意：改阈值请用程序里的「设置」或改配置文件，不要改本字典默认值（会被配置文件覆盖）
# 每个项目可以配置「阴性上限、灰区范围、阳性下限」，示意形式为：
#   阴性：1    灰区：1-10    阳性：10
# 若不需要灰区，则将灰区留空即可，例如：
#   阴性：0.05 灰区：（空）  阳性：0.05
THRESHOLD_RULES = {
    "HBsAg":      {"neg_max": 0.05, "gray_low": None, "gray_high": None, "pos_min": 0.05},
    "HBsAb":      {"neg_max": 10.0, "gray_low": None, "gray_high": None, "pos_min": 10.0},
    "HBeAg":      {"neg_max": 0.1,  "gray_low": None, "gray_high": None, "pos_min": 0.1},
    "HBeAb":      {"neg_max": 0.2,  "gray_low": None, "gray_high": None, "pos_min": 0.2},
    "HBcAb":      {"neg_max": 0.6,  "gray_low": None, "gray_high": None, "pos_min": 0.6},
    "HIV Ag/Ab":  {"neg_max": 1.0,  "gray_low": 1.0, "gray_high": 10.0, "pos_min": 10.0},
    "Anti-HCV":   {"neg_max": 1.0,  "gray_low": 1.0, "gray_high": 10.0, "pos_min": 10.0},
    "Anti-TP":    {"neg_max": 1.0,  "gray_low": 1.0, "gray_high": 10.0, "pos_min": 10.0},
}

# 单位配置（用于阈值设置界面显示）
UNITS = {
    "HBsAg": "IU/mL",
    "HBsAb": "mIU/mL",
    "HBeAg": "IU/mL",
    "HBeAb": "IU/mL",
    "HBcAb": "IU/mL",
    "HIV Ag/Ab": "S/CO",
    "Anti-HCV": "S/CO",
    "Anti-TP": "S/CO",
}

# 显示顺序和乙肝五项配置
DISPLAY_ORDER = ["HBsAg", "HBsAb", "HBeAg", "HBeAb", "HBcAb", "HIV Ag/Ab", "Anti-HCV", "Anti-TP"]
HBV_ORDER = ["HBsAg", "HBsAb", "HBeAg", "HBeAb", "HBcAb"]
HBV_MAP = {"HBsAg": "1", "HBsAb": "2", "HBeAg": "3", "HBeAb": "4", "HBcAb": "5"}


# ==================== 主分析函数 ====================
def analyze_术前(df):
    """
    术前八项数据分析主函数
    返回：(阳性率统计表, 乙肝模式分布表, 样本模式映射, 索引列名)
    """
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.Series(dtype=object), '样本ID'

    # -------------------- 步骤1：提取数值并判定阴阳性 --------------------
    # 从结果列中提取数字
    df['结果_num'] = pd.to_numeric(df['结果'].str.extract(r'(\d+\.?\d*)')[0], errors='coerce')

    def judge(row):
        """根据配置的阈值规则判定阴性/灰区/阳性"""
        p, v = row['测试项目'], row.get('结果_num')
        if pd.isna(v):
            return "阴性"

        # 从配置中读取阈值规则
        rule = THRESHOLD_RULES.get(p)
        if rule is not None and isinstance(rule, dict):
            neg_max = rule.get("neg_max")
            gray_low = rule.get("gray_low")
            gray_high = rule.get("gray_high")
            pos_min = rule.get("pos_min")

            # 有灰区的情况
            if gray_low is not None and gray_high is not None:
                if pos_min is None:
                    pos_min = gray_high
                if v >= pos_min:
                    return "阳性"
                if gray_low <= v < gray_high:
                    return "灰区"
                return "阴性"

            # 无灰区的情况：简单阴阳判定
            if pos_min is None and neg_max is not None:
                pos_min = neg_max
            if pos_min is None:
                return "阴性"
            return "阳性" if v >= pos_min else "阴性"

        # 未配置的项目默认返回阴性
        return "阴性"

    df['判定结果'] = df.apply(judge, axis=1)

    # -------------------- 步骤2：生成阳性率统计表 --------------------
    # 筛选配置的项目并按项目、批号、结果分组统计
    calc_df = df[df['测试项目'].isin(DISPLAY_ORDER)].copy()
    summary = calc_df.groupby(['测试项目', '试剂批号', '判定结果']).size().unstack(fill_value=0)

    # 确保所有必要列存在
    for c in ["阳性", "阴性", "灰区"]:
        if c not in summary.columns:
            summary[c] = 0

    # 计算总数和百分比
    summary['总数'] = summary["阳性"] + summary["阴性"] + summary["灰区"]
    summary['灰区率'] = (summary["灰区"] / summary['总数'] * 100).map('{:.2f}%'.format)
    summary['阳性率'] = (summary["阳性"] / summary['总数'] * 100).map('{:.2f}%'.format)

    # 按配置顺序排序，选择需要的列
    summary = summary.reset_index()
    summary['sort'] = summary['测试项目'].apply(lambda x: DISPLAY_ORDER.index(x) if x in DISPLAY_ORDER else 99)
    summary = summary.sort_values(by=['sort', '试剂批号']).drop(columns=['sort'])
    summary = summary[['测试项目', '灰区', '灰区率', '阳性', '阳性率', '总数', '试剂批号']]

    # -------------------- 步骤3：生成乙肝五项模式分布 --------------------
    # 提取乙肝五项数据
    # i3000使用序号作为唯一标识，i6000使用样本ID
    hbv_raw = calc_df[calc_df['测试项目'].isin(HBV_ORDER)].copy()
    has_serial = '序号' in hbv_raw.columns and (hbv_raw['序号'].astype(str).str.strip() != '').any()
    index_col = '序号' if has_serial else '样本ID'
    
    # 数据透视：每个样本的五项结果
    pivot = hbv_raw.pivot_table(index=index_col, columns='测试项目', values='判定结果', aggfunc='last')

    # 初始化返回值
    mode_stats = pd.DataFrame(columns=['模式', '样本数', '占比'])
    sample_mode_series = pd.Series(dtype=object)
    
    # 检查五项是否都存在（避免删除数据后缺失项目导致KeyError）
    existing_hbv_cols = [col for col in HBV_ORDER if col in pivot.columns]
    if len(existing_hbv_cols) == len(HBV_ORDER):
        # 五项全齐才进行模式分析，筛选出五项都有结果的样本
        valid_pivot = pivot.dropna(subset=HBV_ORDER)
    else:
        # 有项目缺失，返回空结果
        valid_pivot = pd.DataFrame()

    # 构建模式统计
    if not valid_pivot.empty:
        def build_mode(row):
            """构建乙肝模式：阳性用数字表示，阴性用-表示，例如：1----表示只有HBsAg阳性"""
            return "".join([HBV_MAP.get(p, "-") if row.get(p) == "阳性" else "-" for p in HBV_ORDER])

        valid_pivot['最终模式'] = valid_pivot.apply(build_mode, axis=1)
        sample_mode_series = valid_pivot['最终模式']
        
        # 统计各模式的数量和占比
        counts = valid_pivot['最终模式'].value_counts().reset_index()
        counts.columns = ['模式', '样本数']
        total_valid = len(valid_pivot)
        counts['占比'] = (counts['样本数'] / total_valid * 100).map('{:.2f}%'.format)
        mode_stats = counts

    return summary, mode_stats, sample_mode_series, index_col