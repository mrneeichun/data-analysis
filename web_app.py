# ==================== 数据分析Web版 ====================
# 使用Streamlit构建 - 完整功能版
import streamlit as st
import pandas as pd
import io
import os
import json
from datetime import datetime

# 导入原有的数据处理模块
from i3000 import clean_i3000
from i6000 import clean_i6000
from 术前 import analyze_术前, THRESHOLD_RULES as PRE_THRESHOLD_RULES, UNITS as PRE_UNITS, DISPLAY_ORDER as PRE_DISPLAY_ORDER
from 肿瘤 import analyze_肿瘤, TUMOR_RULES, UNITS as TUMOR_UNITS
from 甲功 import analyze_甲功, THYROID_RULES, UNITS as THYROID_UNITS

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="数据分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义样式 ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #3498db;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #3498db;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 初始化Session State ====================
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = {}
if 'thresholds' not in st.session_state:
    st.session_state.thresholds = {
        '术前': PRE_THRESHOLD_RULES.copy(),
        '甲功': THYROID_RULES.copy(),
        '肿瘤': TUMOR_RULES.copy()
    }
if 'full_df' not in st.session_state:
    st.session_state.full_df = None

# ==================== 侧边栏配置 ====================
with st.sidebar:
    st.markdown("## ⚙️ 配置面板")
    
    # 仪器选择
    machine_type = st.selectbox(
        "选择仪器型号",
        ["i3000", "i6000"],
        help="选择数据来源的仪器型号"
    )
    
    # 项目选择
    project_type = st.selectbox(
        "选择分析项目",
        ["术前八项", "甲功", "肿瘤"],
        help="选择要分析的项目类型"
    )
    
    st.markdown("---")
    
    # 阈值设置折叠面板
    with st.expander("🔧 阈值设置", expanded=False):
        st.info("调整各项目的判定阈值")
        
        if project_type == "术前八项":
            st.markdown("**术前八项阈值**")
            for item, rule in st.session_state.thresholds['术前'].items():
                col1, col2, col3 = st.columns(3)
                with col1:
                    neg_max = st.number_input(
                        f"{item} 阴性上限",
                        value=float(rule.get('neg_max', 0)) if rule.get('neg_max') is not None else 0.0,
                        step=0.01,
                        key=f"pre_{item}_neg"
                    )
                with col2:
                    gray_range = st.text_input(
                        f"{item} 灰区(如:1-10)",
                        value=f"{rule.get('gray_low', '')}-{rule.get('gray_high', '')}" if rule.get('gray_low') is not None else "",
                        key=f"pre_{item}_gray"
                    )
                with col3:
                    pos_min = st.number_input(
                        f"{item} 阳性下限",
                        value=float(rule.get('pos_min', 0)) if rule.get('pos_min') is not None else 0.0,
                        step=0.01,
                        key=f"pre_{item}_pos"
                    )
                st.session_state.thresholds['术前'][item]['neg_max'] = neg_max if neg_max > 0 else None
                st.session_state.thresholds['术前'][item]['pos_min'] = pos_min if pos_min > 0 else None
                # 解析灰区
                if gray_range and '-' in gray_range:
                    try:
                        parts = gray_range.split('-')
                        st.session_state.thresholds['术前'][item]['gray_low'] = float(parts[0].strip())
                        st.session_state.thresholds['术前'][item]['gray_high'] = float(parts[1].strip())
                    except:
                        st.session_state.thresholds['术前'][item]['gray_low'] = None
                        st.session_state.thresholds['术前'][item]['gray_high'] = None
                else:
                    st.session_state.thresholds['术前'][item]['gray_low'] = None
                    st.session_state.thresholds['术前'][item]['gray_high'] = None
                    
        elif project_type == "甲功":
            st.markdown("**甲功项目阈值**")
            for item, rule in st.session_state.thresholds['甲功'].items():
                cols = st.columns(3)
                with cols[0]:
                    low_max = st.number_input(
                        f"{item} 偏低上限",
                        value=float(rule.get('low_max', 0)) if rule.get('low_max') is not None else 0.0,
                        step=0.01,
                        key=f"thy_{item}_low"
                    )
                with cols[1]:
                    normal_range = st.text_input(
                        f"{item} 正常范围(如:0.3-4.3)",
                        value=f"{rule.get('normal_low', '')}-{rule.get('normal_high', '')}" if rule.get('normal_low') is not None else "",
                        key=f"thy_{item}_normal"
                    )
                with cols[2]:
                    high_min = st.number_input(
                        f"{item} 偏高下限",
                        value=float(rule.get('high_min', 0)) if rule.get('high_min') is not None else 0.0,
                        step=0.01,
                        key=f"thy_{item}_high"
                    )
                st.session_state.thresholds['甲功'][item]['low_max'] = low_max if low_max > 0 else None
                st.session_state.thresholds['甲功'][item]['high_min'] = high_min if high_min > 0 else None
                # 解析正常范围
                if normal_range and '-' in normal_range:
                    try:
                        parts = normal_range.split('-')
                        st.session_state.thresholds['甲功'][item]['normal_low'] = float(parts[0].strip())
                        st.session_state.thresholds['甲功'][item]['normal_high'] = float(parts[1].strip())
                    except:
                        st.session_state.thresholds['甲功'][item]['normal_low'] = None
                        st.session_state.thresholds['甲功'][item]['normal_high'] = None
                else:
                    st.session_state.thresholds['甲功'][item]['normal_low'] = None
                    st.session_state.thresholds['甲功'][item]['normal_high'] = None
                    
        elif project_type == "肿瘤":
            st.markdown("**肿瘤标志物阈值**")
            for item, rule in st.session_state.thresholds['肿瘤'].items():
                col1, col2 = st.columns(2)
                with col1:
                    neg_max = st.number_input(
                        f"{item} 阴性上限",
                        value=float(rule.get('neg_max', 0)) if rule.get('neg_max') is not None else 0.0,
                        step=0.01,
                        key=f"tumor_{item}_neg"
                    )
                with col2:
                    pos_min = st.number_input(
                        f"{item} 阳性下限",
                        value=float(rule.get('pos_min', 0)) if rule.get('pos_min') is not None else 0.0,
                        step=0.01,
                        key=f"tumor_{item}_pos"
                    )
                st.session_state.thresholds['肿瘤'][item]['neg_max'] = neg_max if neg_max > 0 else None
                st.session_state.thresholds['肿瘤'][item]['pos_min'] = pos_min if pos_min > 0 else None

# ==================== 主页面 ====================
st.markdown('<div class="main-header">📊 数据分析系统</div>', unsafe_allow_html=True)

# 文件上传区域
st.markdown('<div class="section-header">📁 数据导入</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "上传Excel数据文件",
    type=['xlsx', 'xls'],
    help="支持 .xlsx 和 .xls 格式的Excel文件"
)

# 分析按钮
col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    analyze_btn = st.button("🔍 开始分析", type="primary", use_container_width=True)
with col2:
    clear_btn = st.button("🔄 清空数据", use_container_width=True)
with col3:
    export_placeholder = st.empty()

# 清空数据
if clear_btn:
    st.session_state.processed_data = None
    st.session_state.analysis_results = {}
    st.session_state.full_df = None
    st.rerun()

# ==================== 数据处理函数 ====================
def process_data(file, machine_type, project_type):
    """处理上传的数据文件"""
    try:
        # 读取上传的文件内容
        file_bytes = file.read()
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        
        if machine_type == "i3000":
            df = clean_i3000(df)
        elif machine_type == "i6000":
            df = clean_i6000(df)
        
        # 保存完整数据到session state
        st.session_state.full_df = df.copy()
        
        if project_type == "术前八项":
            import 术前
            术前.THRESHOLD_RULES = st.session_state.thresholds['术前']
            summary_df, mode_df, sample_map, index_col = analyze_术前(df)
            return {
                'raw_data': df,
                'summary': summary_df,
                'mode_distribution': mode_df,
                'sample_map': sample_map,
                'type': '术前八项'
            }
            
        elif project_type == "甲功":
            import 甲功
            甲功.THYROID_RULES = st.session_state.thresholds['甲功']
            summary_df = analyze_甲功(df)
            return {
                'raw_data': df,
                'summary': summary_df,
                'type': '甲功'
            }
            
        elif project_type == "肿瘤":
            import 肿瘤
            肿瘤.TUMOR_RULES = st.session_state.thresholds['肿瘤']
            summary_df = analyze_肿瘤(df)
            return {
                'raw_data': df,
                'summary': summary_df,
                'type': '肿瘤'
            }
            
    except Exception as e:
        st.error(f"数据处理错误: {str(e)}")
        return None

# ==================== 导出功能 ====================
def export_results():
    """导出分析结果到Excel"""
    result = st.session_state.analysis_results
    if not result:
        return None
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 写入统计汇总
        if 'summary' in result and not result['summary'].empty:
            result['summary'].to_excel(writer, sheet_name='统计汇总', index=False)
        
        # 写入模式分布（仅术前八项）
        if 'mode_distribution' in result and not result['mode_distribution'].empty:
            result['mode_distribution'].to_excel(writer, sheet_name='乙肝模式分布', index=False)
        
        # 写入原始数据
        if 'raw_data' in result and not result['raw_data'].empty:
            result['raw_data'].to_excel(writer, sheet_name='原始数据', index=False)
    
    output.seek(0)
    return output

# ==================== 执行分析 ====================
if analyze_btn and uploaded_file is not None:
    with st.spinner("正在分析数据，请稍候..."):
        result = process_data(uploaded_file, machine_type, project_type)
        if result:
            st.session_state.analysis_results = result
            st.success("✅ 分析完成！")

# 更新导出按钮
if st.session_state.analysis_results:
    export_data = export_results()
    if export_data:
        with col3:
            st.download_button(
                label="📥 导出结果",
                data=export_data,
                file_name=f"分析结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ==================== 结果显示 ====================
if st.session_state.analysis_results:
    result = st.session_state.analysis_results
    
    st.markdown('<div class="section-header">📈 分析结果</div>', unsafe_allow_html=True)
    
    # 显示统计指标
    if 'raw_data' in result and not result['raw_data'].empty:
        cols = st.columns(4)
        with cols[0]:
            unique_samples = result['raw_data']['样本ID'].nunique() if '样本ID' in result['raw_data'].columns else 0
            st.metric("总样本数", unique_samples)
        with cols[1]:
            total_items = len(result['raw_data'])
            st.metric("总检测项", total_items)
        with cols[2]:
            if 'summary' in result and not result['summary'].empty and '阳性' in result['summary'].columns:
                total_positive = result['summary']['阳性'].sum() if '阳性' in result['summary'].columns else 0
                st.metric("阳性总数", int(total_positive))
        with cols[3]:
            if result['type'] == '术前八项' and 'mode_distribution' in result and not result['mode_distribution'].empty:
                total_modes = len(result['mode_distribution'])
                st.metric("乙肝模式数", total_modes)
    
    if result['type'] == '术前八项':
        tab1, tab2, tab3, tab4 = st.tabs(["📊 阳性率统计", "🧬 乙肝模式分布", "🔍 样本查询", "📋 原始数据"])
        
        with tab1:
            st.markdown("**阳性率统计**")
            if not result['summary'].empty:
                # 添加样式显示
                styled_summary = result['summary'].style.background_gradient(
                    subset=['阳性率(%)'], cmap='Reds', vmin=0, vmax=100
                )
                st.dataframe(styled_summary, use_container_width=True, height=400)
            else:
                st.info("暂无统计数据")
        
        with tab2:
            st.markdown("**乙肝模式分布**")
            if not result['mode_distribution'].empty:
                col_chart, col_table = st.columns([1, 1])
                with col_table:
                    st.dataframe(result['mode_distribution'], use_container_width=True, height=400)
                with col_chart:
                    st.markdown("**模式分布图**")
                    chart_data = result['mode_distribution'].set_index('模式')['数量']
                    st.bar_chart(chart_data)
            else:
                st.info("暂无模式分布数据")
        
        with tab3:
            st.markdown("**样本查询**")
            search_id = st.text_input("输入样本ID进行搜索", placeholder="例如: 2024001")
            if search_id and 'full_df' in st.session_state and st.session_state.full_df is not None:
                filtered_data = st.session_state.full_df[
                    st.session_state.full_df['样本ID'].astype(str).str.contains(search_id, case=False, na=False)
                ]
                if not filtered_data.empty:
                    st.dataframe(filtered_data, use_container_width=True, height=400)
                    
                    # 显示该样本的乙肝模式
                    if 'sample_map' in result and search_id in result['sample_map']:
                        st.info(f"样本 **{search_id}** 的乙肝模式: **{result['sample_map'][search_id]}**")
                else:
                    st.warning("未找到匹配的样本")
        
        with tab4:
            st.markdown("**原始数据**")
            if 'full_df' in st.session_state and st.session_state.full_df is not None:
                # 添加筛选功能
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    if '测试项目' in st.session_state.full_df.columns:
                        projects = ['全部'] + list(st.session_state.full_df['测试项目'].unique())
                        project_filter = st.selectbox("筛选项目", projects)
                with filter_col2:
                    search_sample = st.text_input("搜索样本ID", placeholder="输入样本ID...")
                
                display_df = st.session_state.full_df.copy()
                if project_filter != '全部':
                    display_df = display_df[display_df['测试项目'] == project_filter]
                if search_sample:
                    display_df = display_df[display_df['样本ID'].astype(str).str.contains(search_sample, case=False, na=False)]
                
                st.dataframe(display_df, use_container_width=True, height=500)
                
                # 删除功能
                st.markdown("---")
                st.markdown("**数据删除**")
                st.info("在表格中查看数据，如需删除特定样本，请使用上方筛选功能定位后，在导出前手动清理原始Excel文件后重新上传。")
    
    elif result['type'] == '甲功':
        tab1, tab2, tab3 = st.tabs(["📊 占比分析", "📈 分布图表", "📋 原始数据"])
        
        with tab1:
            st.markdown("**甲功项目统计**")
            if not result['summary'].empty:
                # 为甲功添加颜色渐变
                if '偏高率(%)' in result['summary'].columns:
                    styled_summary = result['summary'].style.background_gradient(
                        subset=['偏高率(%)'], cmap='Reds', vmin=0, vmax=100
                    ).background_gradient(
                        subset=['偏低率(%)'], cmap='Blues', vmin=0, vmax=100
                    )
                else:
                    styled_summary = result['summary']
                st.dataframe(styled_summary, use_container_width=True, height=400)
            else:
                st.info("暂无统计数据")
        
        with tab2:
            st.markdown("**数据分布**")
            if not result['summary'].empty:
                # 准备图表数据
                chart_cols = [c for c in result['summary'].columns if '率' in c and '%' in c]
                if chart_cols:
                    chart_data = result['summary'].set_index('项目')[chart_cols[:3]]  # 取前3个率
                    st.bar_chart(chart_data)
        
        with tab3:
            st.markdown("**原始数据**")
            if 'full_df' in st.session_state and st.session_state.full_df is not None:
                st.dataframe(st.session_state.full_df, use_container_width=True, height=500)
    
    elif result['type'] == '肿瘤':
        tab1, tab2, tab3 = st.tabs(["📊 占比分析", "📈 阳性分布", "📋 原始数据"])
        
        with tab1:
            st.markdown("**肿瘤标志物统计**")
            if not result['summary'].empty:
                styled_summary = result['summary'].style.background_gradient(
                    subset=['阳性率(%)'], cmap='Reds', vmin=0, vmax=100
                )
                st.dataframe(styled_summary, use_container_width=True, height=400)
            else:
                st.info("暂无统计数据")
        
        with tab2:
            st.markdown("**阳性率分布**")
            if not result['summary'].empty and '阳性率(%)' in result['summary'].columns:
                chart_data = result['summary'].set_index('项目')['阳性率(%)']
                st.bar_chart(chart_data)
                
                # 显示阳性率排序
                st.markdown("**阳性率排序**")
                sorted_df = result['summary'].sort_values('阳性率(%)', ascending=False)
                st.dataframe(sorted_df[['项目', '阳性', '阳性率(%)', '总数']], use_container_width=True)
        
        with tab3:
            st.markdown("**原始数据**")
            if 'full_df' in st.session_state and st.session_state.full_df is not None:
                st.dataframe(st.session_state.full_df, use_container_width=True, height=500)

else:
    # 显示使用说明
    st.markdown('<div class="section-header">📖 使用说明</div>', unsafe_allow_html=True)
    
    st.markdown("""
    ### 操作步骤：
    
    1. **选择配置** - 在左侧边栏选择仪器型号和分析项目
    2. **上传文件** - 点击上方上传区域选择Excel数据文件
    3. **调整阈值**（可选）- 在侧边栏展开"阈值设置"进行自定义
    4. **开始分析** - 点击"🔍 开始分析"按钮
    5. **查看结果** - 在下方选项卡中查看统计汇总、模式分布等
    6. **导出结果** - 点击"导出结果"按钮下载Excel报告
    
    ### 支持的项目：
    
    | 项目 | 说明 |
    |------|------|
    | **术前八项** | 乙肝五项 + HIV + HCV + TP，含乙肝模式识别 |
    | **甲功** | TSH、FT3、FT4、TT3、TT4、Anti-TPO、Anti-TG |
    | **肿瘤** | AFP、CEA、CA125、CA153、CA199等标志物 |
    
    ### 支持的仪器：
    
    - i3000
    - i6000
    """)
    
    # 显示示例数据格式
    with st.expander("📋 查看数据格式要求"):
        st.markdown("""
        **i3000数据格式示例：**
        | 序号 | 样本ID | 样本号 | 测试项目 | 结果 | 试剂批号 |
        |------|--------|--------|----------|------|----------|
        | 1 | S001 | 001 | HBsAg | 0.02 | 20240101 |
        | 1 | S001 | 001 | HBsAb | 12.5 | 20240101 |
        
        **i6000数据格式示例：**
        | 样本条码 | 项目名称 | 检测结果 |
        |----------|----------|----------|
        | S001 | HBsAg | 0.02 |
        | S001 | HBsAb | 12.5 |
        """)

# ==================== 页脚 ====================
st.markdown("---")
st.markdown("<div style='text-align: center; color: #7f8c8d;'>数据分析系统 Web版 | Powered by Streamlit</div>", unsafe_allow_html=True)
