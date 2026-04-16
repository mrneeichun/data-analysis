import streamlit as st
import pandas as pd
import io

# 导入核心业务模块
from i3000 import clean_i3000
from i6000 import clean_i6000
from 术前 import analyze_术前, THRESHOLD_RULES as PRE_RULES, UNITS as PRE_UNITS, DISPLAY_ORDER as PRE_ORDER
from 甲功 import analyze_甲功, THYROID_RULES as THYROID_RULES, UNITS as THYROID_UNITS, DISPLAY_ORDER as THYROID_ORDER
from 肿瘤 import analyze_肿瘤, TUMOR_RULES as TUMOR_RULES, UNITS as TUMOR_UNITS, DISPLAY_ORDER as TUMOR_ORDER
from 阈值 import ConfigManager

# 初始化配置管理器
config_mgr = ConfigManager()

# 加载阈值配置
thresholds_pre, thresholds_thyroid, thresholds_tumor = config_mgr.load_thresholds(
    dict(PRE_RULES), dict(THYROID_RULES), dict(TUMOR_RULES)
)

# Streamlit 页面配置
st.set_page_config(
    page_title="数据分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 全局状态管理
if 'full_df' not in st.session_state:
    st.session_state.full_df = None
if 'summary_map' not in st.session_state:
    st.session_state.summary_map = pd.DataFrame()
if 'mode_map' not in st.session_state:
    st.session_state.mode_map = pd.DataFrame()
if 'sample_mode_map' not in st.session_state:
    st.session_state.sample_mode_map = None
if 'current_project' not in st.session_state:
    st.session_state.current_project = "术前八项"
if 'thresholds_pre' not in st.session_state:
    st.session_state.thresholds_pre = thresholds_pre
if 'thresholds_thyroid' not in st.session_state:
    st.session_state.thresholds_thyroid = thresholds_thyroid
if 'thresholds_tumor' not in st.session_state:
    st.session_state.thresholds_tumor = thresholds_tumor

# 更新模块全局变量
import 术前
import 甲功
import 肿瘤
术前.THRESHOLD_RULES = st.session_state.thresholds_pre
甲功.THYROID_RULES = st.session_state.thresholds_thyroid
肿瘤.TUMOR_RULES = st.session_state.thresholds_tumor

def detect_file_format(file):
    """探测文件格式并读取数据"""
    df_probe = None
    
    # 尝试多种编码和引擎读取文件首行
    probe_readers = [
        lambda: pd.read_csv(file, sep='\t', encoding='utf-16', header=None, nrows=1),
        lambda: pd.read_csv(file, sep='\t', encoding='ANSI', header=None, nrows=1),
        lambda: pd.read_excel(file, header=None, nrows=1),
    ]
    
    for reader in probe_readers:
        try:
            file.seek(0)
            df_probe = reader()
            if df_probe is not None:
                break
        except:
            continue
    
    if df_probe is None:
        return None, None
    
    # 检测文件是否有表头
    l1_val = df_probe.iloc[0, 11] if df_probe.shape[1] > 11 else None
    has_l1 = pd.notna(l1_val) and str(l1_val).strip() != ""
    
    return df_probe, has_l1

def read_file(file, machine_type):
    """根据仪器类型读取文件"""
    skiprows = 2 if machine_type == "i3000" else 1
    df_raw = None
    
    readers = [
        lambda: pd.read_csv(file, sep='\t', encoding='utf-16', skiprows=skiprows, dtype=object),
        lambda: pd.read_csv(file, sep='\t', encoding='ANSI', skiprows=skiprows, dtype=object),
        lambda: pd.read_excel(file, skiprows=skiprows, dtype=object),
    ]
    
    for reader in readers:
        try:
            file.seek(0)
            df_raw = reader()
            if df_raw is not None:
                break
        except:
            continue
    
    return df_raw

def process_data(file, machine_type, project_name, progress_bar):
    """处理数据的核心函数"""
    try:
        progress_bar.progress(5)
        st.session_state.status_text = "正在探测文件格式..."
        
        # 文件格式探测
        df_probe, has_l1 = detect_file_format(file)
        if df_probe is None:
            raise Exception("探测失败：无法识别文件编码或格式")
        
        # 仪器类型验证
        if machine_type == "i3000" and not has_l1:
            raise Exception("仪器不匹配：检测到的数据格式更像i6000")
        if machine_type == "i6000" and has_l1:
            raise Exception("仪器不匹配：检测到的数据格式更像i3000")
        
        progress_bar.progress(15)
        st.session_state.status_text = "正在读取文件..."
        
        # 读取完整文件
        df_raw = read_file(file, machine_type)
        if df_raw is None:
            raise Exception("无法读取该文件格式")
        
        progress_bar.progress(35)
        st.session_state.status_text = "正在清洗数据..."
        
        # 数据清洗
        if machine_type == "i3000":
            processed_df = clean_i3000(df_raw)
        else:
            processed_df = clean_i6000(df_raw)
        
        progress_bar.progress(60)
        st.session_state.status_text = "正在分析统计..."
        
        # 项目分析
        if project_name == "术前八项":
            summary_map, mode_map, sample_mode_map, index_col = analyze_术前(processed_df)
            if sample_mode_map is not None and not sample_mode_map.empty:
                sample_mode_map.index = sample_mode_map.index.astype(str).str.strip()
        elif project_name == "肿瘤":
            summary_map, mode_map, _ = analyze_肿瘤(processed_df)
            sample_mode_map = None
        elif project_name == "甲功":
            summary_map, mode_map, _ = analyze_甲功(processed_df)
            sample_mode_map = None
        else:
            summary_map, mode_map = pd.DataFrame(), pd.DataFrame()
            sample_mode_map = None
        
        progress_bar.progress(95)
        st.session_state.status_text = "即将完成..."
        
        # 保存结果到会话状态
        st.session_state.full_df = processed_df
        st.session_state.summary_map = summary_map
        st.session_state.mode_map = mode_map
        st.session_state.sample_mode_map = sample_mode_map
        st.session_state.current_project = project_name
        
        progress_bar.progress(100)
        st.session_state.status_text = "分析完成"
        return True
        
    except Exception as e:
        st.session_state.status_text = f"处理失败: {str(e)}"
        st.error(f"处理失败: {str(e)}")
        return False

def export_data():
    """导出数据为CSV"""
    if st.session_state.full_df is not None:
        csv = st.session_state.full_df.to_csv(index=False, encoding='utf-8-sig')
        return csv
    return None

def save_thresholds():
    """保存阈值配置"""
    config_mgr.thresholds_pre = st.session_state.thresholds_pre
    config_mgr.thresholds_thyroid = st.session_state.thresholds_thyroid
    config_mgr.thresholds_tumor = st.session_state.thresholds_tumor
    config_mgr.save_thresholds()
    
    # 更新模块全局变量
    术前.THRESHOLD_RULES = st.session_state.thresholds_pre
    甲功.THYROID_RULES = st.session_state.thresholds_thyroid
    肿瘤.TUMOR_RULES = st.session_state.thresholds_tumor
    
    st.success("阈值已更新并保存，将在下一次分析时生效")

# ==================== 主界面 ====================
st.title("📊 数据分析系统")

# 侧边栏
with st.sidebar:
    st.header("数据导入")
    
    # 文件上传
    uploaded_file = st.file_uploader("选择数据文件", type=['txt', 'xlsx', 'xls', 'csv'])
    
    # 仪器型号选择
    machine_type = st.selectbox("选择仪器型号", ["i3000", "i6000"])
    
    # 项目选择
    project_name = st.selectbox("选择项目名称", ["术前八项", "甲功", "肿瘤"])
    
    # 运行按钮
    run_button = st.button("开始分析", type="primary")
    
    # 导出按钮
    if st.session_state.full_df is not None:
        csv_data = export_data()
        if csv_data:
            st.download_button(
                label="导出原始数据",
                data=csv_data,
                file_name="分析结果.csv",
                mime="text/csv"
            )

# 主内容区
if run_button and uploaded_file is not None:
    progress_bar = st.progress(0)
    st.session_state.status_text = "准备中..."
    
    status_placeholder = st.empty()
    with status_placeholder:
        st.write(st.session_state.status_text)
    
    def update_status():
        status_placeholder.write(st.session_state.status_text)
    
    import threading
    thread = threading.Thread(target=lambda: process_data(uploaded_file, machine_type, project_name, progress_bar))
    thread.start()
    thread.join()
    
    update_status()

# 状态显示
if 'status_text' in st.session_state:
    st.write(f"**状态:** {st.session_state.status_text}")

# 选项卡内容
tab1, tab2, tab3, tab4 = st.tabs(["术前八项", "甲功", "肿瘤", "原始数据"])

with tab1:
    st.subheader("术前八项分析")
    
    if st.session_state.current_project == "术前八项":
        if not st.session_state.summary_map.empty:
            st.write("### 阳性率统计")
            st.dataframe(st.session_state.summary_map, use_container_width=True)
        
        if not st.session_state.mode_map.empty:
            st.write("### 乙肝模式分布")
            st.dataframe(st.session_state.mode_map, use_container_width=True)
    else:
        st.info("请选择术前八项项目并运行分析")

with tab2:
    st.subheader("甲功分析")
    
    if st.session_state.current_project == "甲功":
        if not st.session_state.summary_map.empty:
            st.write("### 占比分析")
            st.dataframe(st.session_state.summary_map, use_container_width=True)
    else:
        st.info("请选择甲功项目并运行分析")

with tab3:
    st.subheader("肿瘤分析")
    
    if st.session_state.current_project == "肿瘤":
        if not st.session_state.summary_map.empty:
            st.write("### 占比分析")
            st.dataframe(st.session_state.summary_map, use_container_width=True)
    else:
        st.info("请选择肿瘤项目并运行分析")

with tab4:
    st.subheader("原始数据")
    
    if st.session_state.full_df is not None:
        # 搜索功能
        search_query = st.text_input("搜索样本ID")
        display_df = st.session_state.full_df.copy()
        
        if search_query:
            display_df = display_df[display_df['样本ID'].astype(str).str.contains(search_query, na=False)]
        
        st.dataframe(display_df, use_container_width=True, height=500)
    else:
        st.info("请先导入数据并运行分析")

# 阈值设置
with st.expander("⚙️ 阈值设置"):
    st.write("修改阈值后点击保存按钮，下次分析将使用新阈值")
    
    threshold_tab1, threshold_tab2, threshold_tab3 = st.tabs(["术前八项", "甲功", "肿瘤"])
    
    with threshold_tab1:
        st.write("#### 术前八项（阴性 / 灰区 / 阳性）")
        for proj, rule in st.session_state.thresholds_pre.items():
            unit = PRE_UNITS.get(proj, "")
            st.write(f"**{proj}** ({unit})")
            col1, col2, col3 = st.columns(3)
            with col1:
                neg_max = st.number_input(f"{proj}_neg_max", value=rule.get('neg_max') or 0.0, key=f"pre_{proj}_neg")
            with col2:
                gray_low = st.number_input(f"{proj}_gray_low", value=rule.get('gray_low') or 0.0, key=f"pre_{proj}_gray_low")
                gray_high = st.number_input(f"{proj}_gray_high", value=rule.get('gray_high') or 0.0, key=f"pre_{proj}_gray_high")
            with col3:
                pos_min = st.number_input(f"{proj}_pos_min", value=rule.get('pos_min') or 0.0, key=f"pre_{proj}_pos")
            
            st.session_state.thresholds_pre[proj] = {
                'neg_max': neg_max,
                'gray_low': gray_low if gray_low != 0 else None,
                'gray_high': gray_high if gray_high != 0 else None,
                'pos_min': pos_min
            }
    
    with threshold_tab2:
        st.write("#### 甲功（偏低 / 正常 / 偏高）")
        for proj, rule in st.session_state.thresholds_thyroid.items():
            unit = THYROID_UNITS.get(proj, "")
            st.write(f"**{proj}** ({unit})")
            col1, col2, col3 = st.columns(3)
            with col1:
                low_max = st.number_input(f"{proj}_low_max", value=rule.get('low_max') or 0.0, key=f"thy_{proj}_low")
            with col2:
                normal_low = st.number_input(f"{proj}_normal_low", value=rule.get('normal_low') or 0.0, key=f"thy_{proj}_norm_low")
                normal_high = st.number_input(f"{proj}_normal_high", value=rule.get('normal_high') or 0.0, key=f"thy_{proj}_norm_high")
            with col3:
                high_min = st.number_input(f"{proj}_high_min", value=rule.get('high_min') or 0.0, key=f"thy_{proj}_high")
            
            st.session_state.thresholds_thyroid[proj] = {
                'low_max': low_max,
                'normal_low': normal_low if normal_low != 0 else None,
                'normal_high': normal_high if normal_high != 0 else None,
                'high_min': high_min
            }
    
    with threshold_tab3:
        st.write("#### 肿瘤（阴性 / 阳性）")
        for proj, rule in st.session_state.thresholds_tumor.items():
            unit = TUMOR_UNITS.get(proj, "")
            st.write(f"**{proj}** ({unit})")
            threshold = st.number_input(f"{proj}_threshold", value=max(rule.get('neg_max') or 0, rule.get('pos_min') or 0), key=f"tumor_{proj}")
            
            st.session_state.thresholds_tumor[proj] = {
                'neg_max': threshold,
                'pos_min': threshold
            }
    
    st.button("保存阈值配置", on_click=save_thresholds)

# 页脚
st.markdown("---")
st.write("数据分析系统 v1.0 | 支持 i3000/i6000 仪器数据")
