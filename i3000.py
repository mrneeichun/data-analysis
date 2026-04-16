import pandas as pd
# ==================== i3000设备数据清洗 ====================
def clean_i3000(raw_df):
    df = raw_df.copy()
    # -------------------- 步骤1：列名清洗 --------------------
    df.columns = df.columns.astype(str).str.strip()

    # -------------------- 步骤2：标识符填充 --------------------
    # 向下填充
    for col in ['序号', '样本ID', '样本号']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaN', 'None', '', ' '], pd.NA)
            df[col] = df[col].ffill()
    # -------------------- 步骤3：数据清洗 --------------------
    def clean_val(x):
        if pd.isna(x):
            return ""
        return str(x).replace('=', '').replace('...', '').replace('"', '').strip()

    for c in ['试剂批号', '测试项目', '结果']:
        if c in df.columns:
            df[c] = df[c].apply(clean_val)

    # -------------------- 步骤4：唯一标识 序号+项目--------------------
    if '序号' in df.columns and '测试项目' in df.columns:
        df['UID'] = df['序号'].astype(str) + "_" + df['测试项目'].astype(str)
        
    return df