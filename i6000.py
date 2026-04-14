import pandas as pd
# ==================== i6000设备数据清洗 ====================
def clean_i6000(raw_df):
    df = raw_df.copy()
    # -------------------- 步骤1：修正标题行 --------------------
    if '样本条码' not in df.columns:
        df.columns = df.iloc[0].astype(str).str.strip()
        df = df.iloc[1:].reset_index(drop=True)
    # -------------------- 步骤2：统一列名 --------------------
    # 将i6000的列名映射到统一的列名
    df.rename(columns={'样本条码': '样本ID', '项目名称': '测试项目', '检测结果': '结果'}, inplace=True)
    # -------------------- 步骤3：样本ID清洗和填充 --------------------
    df['样本ID'] = df['样本ID'].astype(str).str.replace(r'[=\"]', '', regex=True).str.strip()
    df['样本ID'] = df['样本ID'].ffill()
    # -------------------- 步骤4：数据清洗 --------------------
    for c in ['测试项目', '结果', '试剂批号']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.replace('=', '').str.replace('"', '').str.strip()
    # -------------------- 步骤5：提取数值 -------------------
    if '结果' in df.columns:
        df['结果_num'] = pd.to_numeric(df['结果'], errors='coerce')
    # i6000原始数据无序号列，术前八项模式统计使用样本ID作为标识
    return df