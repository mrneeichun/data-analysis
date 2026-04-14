# ==================== 库、数据模块导入 ====================
import tkinter as tk                   # 主窗口
from tkinter import filedialog         # 文件
from tkinter import ttk                # ui控件
from tkinter import messagebox         # 弹窗
import pandas as pd                    # 数据分析核心库
import os                              # 系统操作
import threading                       # 多线程支持
from i3000 import clean_i3000          # i3000仪器数据清洗
from i6000 import clean_i6000          # i6000仪器数据清洗
from 术前 import analyze_术前          # 术前八项(含乙肝模式识别)
from 肿瘤 import analyze_肿瘤          # 肿瘤
from 甲功 import analyze_甲功          # 甲状腺
import 术前                           # 阈值配置
import 甲功                           
import 肿瘤                            
from 阈值 import ConfigManager         # 配置管理器

# ==================== 主程序====================
class ModernApp:     
    def __init__(self, root):
        # -------------------- 窗口配置 --------------------
        self.root = root
        self.root.title("数据分析")              # 窗口标题
        self.root.geometry("1280x700")           # 窗口尺寸
        self.root.configure(bg="#F0F5F9")        # 背景色
        
        # -------------------- 数据变量 --------------------
        # （1）数据存储
        self.full_df = None                      # DataFrame,存储处理后的原始数据
        
        # （2）UI控制变量
        self.file_path = tk.StringVar()          # 文件路径
        self.machine_type = tk.StringVar()       # 仪器型号
        self.project_name = tk.StringVar()       # 项目名称
        self.search_var = tk.StringVar()         # 样本ID搜索
        self.status_var = tk.StringVar(value="待分析")  # 状态栏文本
        
        # （3）分析结果存储
        self.summary_map = pd.DataFrame({})      # 统计汇总表
        self.mode_map = pd.DataFrame({})         # 模式分布表
        
        # （4）功能状态标记
        self._last_hint = None                   # 未实现功能的提示信息缓存
        self.sample_mode_map = None              # 乙肝模式映射表
        self.sample_mode_index_col = '样本ID'    # 样本唯一标识列名
        
        # （5）筛选和排序控制
        self.mode_filter_var = tk.StringVar(value="全部")    # 乙肝模式筛选条件
        self.project_filter_var = tk.StringVar(value="全部") # 项目筛选条件
        self.tree0_sort_col = None               # 原始数据表当前排序列
        self.tree0_sort_reverse = False          # 原始数据表排序方向False=升序

        # -------------------- 阈值配置加载 --------------------
        # 加载阈值（从本地json文件或使用模块默认值）
        self.config_mgr = ConfigManager()
        self.thresholds_pre, self.thresholds_thyroid, self.thresholds_tumor = self.config_mgr.load_thresholds(
            dict(getattr(术前, "THRESHOLD_RULES", {})),      # 术前阈值
            dict(getattr(甲功, "THYROID_RULES", {})),        # 甲功
            dict(getattr(肿瘤, "TUMOR_RULES", {}))           # 肿瘤
        )
        
        # 同步阈值
        术前.THRESHOLD_RULES = self.thresholds_pre
        甲功.THYROID_RULES = self.thresholds_thyroid
        肿瘤.TUMOR_RULES = self.thresholds_tumor

        # -------------------- 构建UI界面 --------------------
        self.setup_ui()
    
    # ==================== UI辅助方法 ====================
    def _add_hover_effect(self, widget, enter_bg, leave_bg, enter_fg=None, leave_fg=None):

        def on_enter(e):
            widget['background'] = enter_bg
            if enter_fg:
                widget['foreground'] = enter_fg
        
        def on_leave(e):
            widget['background'] = leave_bg
            if leave_fg:
                widget['foreground'] = leave_fg
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
        
    # ==================== UI界面 ====================
    def setup_ui(self):
           
        # ==================== 全局 ====================
        self.root.configure(bg="#E8ECF1")        # 设置主窗口背景色
        style = ttk.Style()
        
        # 选项卡样式
        style.configure("TNotebook", background="#E8ECF1")                    # 选项卡背景
        style.configure("TNotebook.Tab", padding=[24, 10], font=("微软雅黑", 10))  # 选项卡内边距和字体
        style.map("TNotebook.Tab", 
                  background=[("selected", "#fff"), ("active", "#E3F2FD")],   # 选中时白色,悬停时浅蓝
                  expand=[("selected", [1, 1, 1, 0])])                        # 选中时微调位置
        
        # 顶层选项卡样式（术前/甲功/肿瘤）
        style.configure("Top.TNotebook", background="#E8ECF1")
        style.configure("Top.TNotebook.Tab", padding=[28, 12], font=("微软雅黑", 11, "bold"))
        style.map("Top.TNotebook.Tab", 
                  background=[("selected", "#fff"), ("active", "#E3F2FD")], 
                  expand=[("selected", [1, 1, 1, 0])])
        
        # 其他组件样式
        style.configure("TFrame", background="#E8ECF1")                       # 框架背景色
        style.configure("TLabelframe", background="#E8ECF1", font=("微软雅黑", 9))  # 标签框架
        style.configure("TLabelframe.Label", background="#E8ECF1", foreground="#2C3E50", font=("微软雅黑", 9, "bold"))  # 标签文字
        style.configure("TProgressbar", troughcolor="#D5DBE0", background="#3498DB", thickness=8)  # 进度条
        style.configure("Treeview", font=("微软雅黑", 9), rowheight=22, fieldbackground="white")   # 表格
        style.configure("Treeview.Heading", font=("微软雅黑", 9, "bold"), background="#ECF0F1")    # 表头
        style.map("Treeview.Heading", background=[("active", "#D5DBE0")])  # 表头悬停效果

        # ==================== 顶部区 ====================
        opt_frame = tk.Frame(self.root, bg="#E8ECF1", pady=14)  # 创建顶部操作区域框架
        opt_frame.pack(fill="x", padx=24)                       # 水平填充,左右留白24像素
        
        # 导入数据区域
        f_lf = ttk.LabelFrame(opt_frame, text=" 1. 导入原始数据 ", padding=10)  # 标签框架容器
        f_lf.pack(side="left", fill="x", expand=True, padx=6)                   # 左对齐,可扩展
        
        # 文件路径输入框
        e1 = tk.Entry(f_lf, textvariable=self.file_path, bd=1, relief="solid", font=("微软雅黑", 9))
        e1.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=2)        # 左侧填充,右侧留8像素间距
        
        # 打开文件按钮
        open_btn = tk.Button(f_lf, text="打开", command=lambda: self.file_path.set(filedialog.askopenfilename()),
                 bg="#3498DB", fg="white", font=("微软雅黑", 9), bd=0, padx=14, pady=2,
                 activebackground="#2980B9", activeforeground="white", cursor="hand2")
        open_btn.pack(side="right")
        self._add_hover_effect(open_btn, "#2980B9", "#3498DB")
        
        # 仪器型号选择区域
        m_lf = ttk.LabelFrame(opt_frame, text=" 2. 选择仪器型号 ", padding=10)
        m_lf.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Combobox(m_lf, textvariable=self.machine_type, values=["i3000", "i6000"], state="readonly", font=("微软雅黑", 9)).pack(fill="x", pady=2)
        
        # 项目选择区域
        p_lf = ttk.LabelFrame(opt_frame, text=" 3. 选择项目名称 ", padding=10)
        p_lf.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Combobox(p_lf, textvariable=self.project_name, values=["术前八项", "甲功", "肿瘤"],state="readonly", font=("微软雅黑", 9)).pack(fill="x", pady=2)

        # ==================== 主操作按钮区 ====================
        btn_f = tk.Frame(self.root, bg="#E8ECF1")                # 按钮容器框架
        btn_f.pack(fill="x", padx=24, pady=8)                    # 水平填充,上下留白
        
        # 开始分析按钮
        self.run_btn = tk.Button(btn_f, text="开始分析", command=self.run, bg="#3498DB", fg="white",
                                 font=("微软雅黑", 11, "bold"), height=2, bd=0, padx=24, cursor="hand2",
                                 activebackground="#2980B9", activeforeground="white")
        self.run_btn.pack(side="left", fill="x", expand=True, padx=6)  # 左侧填充,平均分配空间
        self._add_hover_effect(self.run_btn, "#2980B9", "#3498DB")
        
        # 导出数据按钮
        self.export_btn = tk.Button(btn_f, text="导出当前结果", command=self.export, bg="#95A5A6", fg="white",
                                    font=("微软雅黑", 11), height=2, bd=0, padx=24, state="disabled", cursor="hand2",
                                    activebackground="#7F8C8D", activeforeground="white")
        self.export_btn.pack(side="left", fill="x", expand=True, padx=6)
        self._add_hover_effect(self.export_btn, "#7F8C8D", "#95A5A6")
        
        # 阈值设置按钮
        self.threshold_btn = tk.Button(btn_f,text="阈值设置",command=self.open_threshold_dialog,bg="#ECF0F1",fg="#2C3E50",
                                       font=("微软雅黑", 9),height=2,bd=0,padx=18, cursor="hand2",
                                       activebackground="#D0D7DE",activeforeground="#2C3E50",)
        self.threshold_btn.pack(side="right", padx=6)  # 右侧对齐
        self._add_hover_effect(self.threshold_btn, "#D0D7DE", "#ECF0F1")

        # ==================== 进度条和状态栏 ====================
        prog_frame = tk.Frame(self.root, bg="#E8ECF1")           # 进度条容器
        prog_frame.pack(fill="x", padx=24, pady=(0, 8))          # 底部留白8像素
        self.progress = ttk.Progressbar(prog_frame, orient="horizontal", mode="determinate")  # 确定模式进度条
        self.progress.pack(fill="x", side="top", pady=(0, 4))    # 顶部对齐,底部留白4像素
        # 创建状态标签（待分析状态为红色）
        self.status_label = tk.Label(prog_frame, textvariable=self.status_var, bg="#E8ECF1", fg="#E67E22", font=("微软雅黑", 9, "bold"))
        self.status_label.pack(side="left")

        # ==================== 主选项卡区 ====================
        # 顶层选项卡
        # 使用自定义按钮实现四个主选项卡
        top_tab_bar = tk.Frame(self.root, bg="#D5DBE0", height=42)  # 选项卡栏容器,高度42像素
        top_tab_bar.pack(fill="x", padx=24, pady=(0, 0))             # 顶部对齐,无垂直间距
        top_tab_bar.pack_propagate(False)                            # 固定高度不随内容变化
        self._top_frames = []                                        # 存储各选项卡内容框架的列表
        
        # 四个选项卡按钮
        for i, label in enumerate(["术前八项", "甲功", "肿瘤", "原始数据"]):
            btn = tk.Button(top_tab_bar, text=label, font=("微软雅黑", 10), bd=0, cursor="hand2",
                            bg="#BDC3C7", fg="#2C3E50", activebackground="#ECF0F1", activeforeground="#2C3E50",
                            command=lambda idx=i: self._switch_top_tab(idx))  # 点击切换选项卡
            btn.pack(side="left", fill="x", expand=True, padx=1, ipady=8)     # 平均分配宽度,垂直内边距8像素
            if i == 0:
                # 默认选中术前八项：蓝色背景,白色文字
                btn.config(bg="#3498DB", fg="white", activebackground="#2980B9", activeforeground="white")
            setattr(self, "_top_tab_btn_%d" % i, btn)  # 动态创建属性引用
        self._top_tab_selected = 0  # 记录当前选中的选项卡索引

        # 主内容区
        top_content = tk.Frame(self.root, bg="#E8ECF1")          # 内容区域容器
        top_content.pack(fill="both", expand=True, padx=24, pady=(2, 16))  # 四周留白

        # ==================== 术前八项选项卡内容 ====================
        self.frame_术前 = tk.Frame(top_content, bg="#E8ECF1")    # 术前八项主框架
        self.frame_术前.pack(fill="both", expand=True)           # 填充整个内容区域
        self._top_frames.append(self.frame_术前)                 # 添加到选项卡框架列表
        self.notebook = ttk.Notebook(self.frame_术前)            # 创建子选项卡容器
        self.notebook.pack(fill="both", expand=True)             # 填充术前框架

        # 术前-阳性率统计表
        self.tab1 = tk.Frame(self.notebook, bg="white", padx=12, pady=12)  # 白色背景,内边距12像素
        self.notebook.add(self.tab1, text="  阳性率统计  ")                # 添加到选项卡
        self.tree1 = ttk.Treeview(self.tab1, columns=("项目", "灰区", "灰区率", "阳性", "阳性率", "总数", "批号"), show="headings")
        for col in self.tree1["columns"]:
            self.tree1.heading(col, text=col)                    # 设置表头文字
            self.tree1.column(col, width=120, anchor="center")   # 设置列宽和对齐方式
        self.tree1.pack(fill="both", expand=True)                # 填充整个选项卡

        # 乙肝五项模式分布
        self.tab2 = tk.Frame(self.notebook, bg="white", padx=12, pady=12)
        self.notebook.add(self.tab2, text="  乙肝模式分布  ")
        self.tree2 = ttk.Treeview(self.tab2, columns=("M", "C", "Pct"), show="headings")
        for c, h in zip(self.tree2["columns"], ["乙肝模式", "数量", "占比"]):
            self.tree2.heading(c, text=h)
            self.tree2.column(c, width=200, anchor="center")
        self.tree2.pack(fill="both", expand=True)

        # ==================== 甲功选项卡内容 ====================
        self.frame_甲功 = tk.Frame(top_content, bg="#E8ECF1")    # 甲功主框架
        self.frame_甲功.place(relx=0, rely=0, relwidth=1, relheight=1)  # 覆盖整个内容区域（叠放布局）
        self._top_frames.append(self.frame_甲功)

        # ==================== 肿瘤选项卡内容 ====================
        self.frame_肿瘤 = tk.Frame(top_content, bg="#E8ECF1")    # 肿瘤主框架
        self.frame_肿瘤.place(relx=0, rely=0, relwidth=1, relheight=1)  # 覆盖整个内容区域（叠放布局）
        self._top_frames.append(self.frame_肿瘤)
        
        # 原始数据选项卡框架
        self.frame_原始数据 = tk.Frame(top_content, bg="#E8ECF1")
        self.frame_原始数据.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._top_frames.append(self.frame_原始数据)
        
        # 肿瘤：占比分析 + 原始数据
        self.notebook_肿瘤 = ttk.Notebook(self.frame_肿瘤)       # 肿瘤子选项卡容器
        self.notebook_肿瘤.pack(fill="both", expand=True)
        
        self.notebook_甲功 = ttk.Notebook(self.frame_甲功)       # 甲功子选项卡容器
        self.notebook_甲功.pack(fill="both", expand=True)

        # 甲功 - 占比分析
        self.tab_甲功1 = tk.Frame(self.notebook_甲功, bg="white", padx=12, pady=12)
        self.notebook_甲功.add(self.tab_甲功1, text="  占比分析  ")
        cols_甲功1 = ("项目", "偏低", "偏低率", "正常", "正常率", "偏高", "偏高率", "试剂批号")
        self.tree_甲功1 = ttk.Treeview(self.tab_甲功1, columns=cols_甲功1, show="headings")
        for c in cols_甲功1:
            self.tree_甲功1.heading(c, text=c)
            self.tree_甲功1.column(c, width=100, anchor="center")
        self.tree_甲功1.pack(fill="both", expand=True)

        # 肿瘤 - 占比分析
        self.tab_肿瘤1 = tk.Frame(self.notebook_肿瘤, bg="white", padx=12, pady=12)
        self.notebook_肿瘤.add(self.tab_肿瘤1, text="  占比分析  ")
        cols_肿瘤1 = ("项目", "阳性", "阳性率", "总数", "试剂批号")
        self.tree_肿瘤1 = ttk.Treeview(self.tab_肿瘤1, columns=cols_肿瘤1, show="headings")
        for c in cols_肿瘤1:
            self.tree_肿瘤1.heading(c, text=c)
            self.tree_肿瘤1.column(c, width=100, anchor="center")
        self.tree_肿瘤1.pack(fill="both", expand=True)

        # ==================== 原始数据选项卡内容 ====================
        # 原始数据浏览容器
        raw_container = tk.Frame(self.frame_原始数据, bg="white")  # 白色背景容器
        raw_container.pack(fill="both", expand=True)
        
        # 筛选工具栏
        filter_bar = tk.Frame(raw_container, bg="#F4F6F7", height=44, bd=0)  # 浅灰色筛选栏
        filter_bar.pack(fill="x")                                            # 水平填充
        filter_bar.pack_propagate(False)                                     # 固定高度
        
        # 分隔线（视觉分隔）
        sep_inner = tk.Frame(filter_bar, width=1, bg="#D5DBE0")              # 竖直分隔线
        sep_inner.place(relx=0.32, rely=0.15, relheight=0.7)                 # 相对定位
        
        # 样本ID搜索框
        tk.Label(filter_bar, text="样本ID", bg="#F4F6F7", fg="#2C3E50", font=("微软雅黑", 9)).pack(side="left", padx=(16, 6), pady=10)
        s_entry = tk.Entry(filter_bar, textvariable=self.search_var, width=28, font=("微软雅黑", 9), bd=1, relief="solid")
        s_entry.pack(side="left", padx=(0, 8), pady=8)
        s_entry.bind("<KeyRelease>", lambda e: self.refresh_ui())            # 实时刷新搜索结果
        
        # 重置按钮
        tk.Button(filter_bar, text="重置", command=lambda: [self.search_var.set(""), self.refresh_ui()],
                  bg="white", fg="#5D6D7E", font=("微软雅黑", 9), bd=1, relief="solid", padx=12, pady=2, cursor="hand2").pack(side="left", padx=(0, 24))
        
        # 乙肝模式筛选（仅术前八项可用）
        tk.Label(filter_bar, text="乙肝模式筛选", bg="#F4F6F7", fg="#2C3E50", font=("微软雅黑", 9)).pack(side="left", padx=(0, 6), pady=10)
        self.mode_filter_combo = ttk.Combobox(filter_bar, textvariable=self.mode_filter_var, width=14, state="readonly", values=["全部"], font=("微软雅黑", 9))
        self.mode_filter_combo.pack(side="left", padx=(0, 16), pady=8)
        self.mode_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_ui())
        
        # 项目筛选（通用功能）
        tk.Label(filter_bar, text="项目筛选", bg="#F4F6F7", fg="#2C3E50", font=("微软雅黑", 9)).pack(side="left", padx=(0, 6), pady=10)
        # 存储选中的项目集合
        self.selected_projects_set = set()
        self.project_filter_combo = ttk.Combobox(filter_bar, textvariable=self.project_filter_var, width=14, 
                                               state="readonly", font=("微软雅黑", 9))
        self.project_filter_combo.pack(side="left", padx=(0, 16), pady=8)
        self.project_filter_combo.bind("<<ComboboxSelected>>", self._on_project_selection)

        # 原始数据表格
        self.cols0 = ("idx", "sn", "sid", "proj", "res", "unit", "dilution", "time", "batch", "recheck")  # 列标识符
        self.heads0 = ["序号", "样本号", "样本ID", "测试项目", "结果", "单位", "稀释倍数", "检测时间", "试剂批号", "复查结果"]  # 表头文字
        self.tree0 = ttk.Treeview(raw_container, columns=self.cols0, show="headings")  # 创建表格控件
        self._col_to_df = dict(zip(self.cols0, self.heads0))  # 列标识符到表头文字的映射字典
        for c, h in zip(self.cols0, self.heads0):
            self.tree0.heading(c, text=h, command=lambda col=c: self._sort_tree0(col))  # 设置表头（带排序功能）
            self.tree0.column(c, width=100, anchor="center")  # 设置列宽和对齐
        self.tree0.pack(side="left", fill="both", expand=True)  # 左侧对齐,填充剩余空间
        
        # 排序列映射字典（用于排序时查找真实列名）
        self._tree0_sort_col_map = {"idx": "序号", "sn": "样本号", "sid": "样本ID", "proj": "测试项目",
                                    "res": "结果", "unit": "单位", "time": "检测完成时间", "batch": "试剂批号",
                                    "dilution": "稀释倍数", "recheck": "复查结果"}
        
        # 滚动条
        sb0 = ttk.Scrollbar(raw_container, orient="vertical", command=self.tree0.yview)  # 垂直滚动条
        self.tree0.configure(yscrollcommand=sb0.set)  # 关联表格和滚动条
        sb0.pack(side="right", fill="y")              # 右侧对齐,垂直填充
        
        # 右键菜单绑定
        self.tree0.bind("<Button-3>", self._show_tree0_context_menu)  # 绑定右键点击事件

        # 默认显示术前八项选项卡
        self._top_frames[0].lift()

    # ====================阈值设置对话框 ====================
    def open_threshold_dialog(self):
        if not getattr(术前, "THRESHOLD_RULES", None) and not getattr(甲功, "THYROID_RULES", None) and not getattr(肿瘤, "TUMOR_RULES", None):
            messagebox.showwarning("提示", "当前版本未找到可配置的阈值。")
            return

        # 创建模态对话框窗口
        top = tk.Toplevel(self.root)
        top.title("阈值设置")
        top.configure(bg="#F7F9FB")              # 浅蓝灰色背景
        top.resizable(False, False)              # 禁止调整大小

        # 主容器框架
        frm = tk.Frame(top, bg="#F7F9FB", padx=16, pady=16)
        frm.pack(fill="both", expand=True)

        # 创建Notebook控件实现分页功能
        notebook = ttk.Notebook(frm)
        notebook.pack(fill="both", expand=True)

        # 创建三个分页标签
        tab_pre = tk.Frame(notebook, bg="#F7F9FB", padx=12, pady=12)     # 术前八项分页
        notebook.add(tab_pre, text="  术前八项  ")

        tab_thyroid = tk.Frame(notebook, bg="#F7F9FB", padx=12, pady=12) # 甲功分页
        notebook.add(tab_thyroid, text="  甲功  ")

        tab_tumor = tk.Frame(notebook, bg="#F7F9FB", padx=12, pady=12)   # 肿瘤分页
        notebook.add(tab_tumor, text="  肿瘤  ")

        # ==================== 术前八项分页 ====================
        # 标题和说明
        tk.Label(tab_pre,text="术前八项（阴性 / 灰区 / 阳性）",bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))

        tk.Label(tab_pre,text="灰区可自由设置",bg="#F7F9FB",fg="#7B8A8B",justify="left",font=("微软雅黑", 8),
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 6))

        # 表头
        headers_pre = ["项目", "阴性上限", "灰区范围（可留空）", "阳性下限"]
        for j, h in enumerate(headers_pre):
            tk.Label(tab_pre,text=h,bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),anchor="w",
            ).grid(row=2, column=j, sticky="w", pady=(0, 4), padx=(0 if j == 0 else 8, 0))

        # 获取当前阈值配置
        current_pre = self.thresholds_pre or dict(术前.THRESHOLD_RULES)
        pre_units = getattr(术前,"UNITS",{})  # 获取单位配置
        pre_neg_vars, pre_gray_vars, pre_pos_vars = {}, {}, {}  # 存储各项目阈值的变量

        # 动态创建输入框（每个项目一行）
        row_offset = 3
        for i, (proj, rule) in enumerate(current_pre.items(), start=row_offset):
            # 提取当前阈值
            neg_max = rule.get("neg_max")        # 阴性上限
            gray_low = rule.get("gray_low")      # 灰区下限
            gray_high = rule.get("gray_high")    # 灰区上限
            pos_min = rule.get("pos_min")        # 阳性下限
            
            # 获取单位（如果没有配置则留空）
            unit = pre_units.get(proj, "")

            # 项目名称标签（带单位）
            proj_text = f"{proj} ({unit})" if unit else proj
            tk.Label(tab_pre,text=proj_text,bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9),anchor="w",width=16,
            ).grid(row=i, column=0, sticky="w", pady=3)

            # 创建输入变量
            neg_var = tk.StringVar(value="" if neg_max is None else str(neg_max))  # 阴性上限输入框
            gray_text = ""
            if gray_low is not None and gray_high is not None:
                gray_text = f"{gray_low}-{gray_high}"  # 灰区范围格式：下限-上限
            gray_var = tk.StringVar(value=gray_text)   # 灰区范围输入框
            pos_var = tk.StringVar(value="" if pos_min is None else str(pos_min))  # 阳性下限输入框

            # 创建输入控件
            tk.Entry(tab_pre, textvariable=neg_var, width=10, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=1, sticky="w", pady=3, padx=(8, 0))
            tk.Entry(tab_pre, textvariable=gray_var, width=14, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=2, sticky="w", pady=3, padx=(8, 0))
            tk.Entry(tab_pre, textvariable=pos_var, width=10, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=3, sticky="w", pady=3, padx=(8, 0))

            # 保存变量引用以便后续读取
            pre_neg_vars[proj] = neg_var
            pre_gray_vars[proj] = gray_var
            pre_pos_vars[proj] = pos_var

        last_pre_row = row_offset + len(current_pre)

        # ==================== 甲功分页 ====================
        # 标题和说明
        tk.Label(tab_thyroid,text="甲功（偏低 / 正常 / 偏高）",bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))
        tk.Label(tab_thyroid,text="请确认单位",bg="#F7F9FB",fg="#7B8A8B",justify="left",font=("微软雅黑", 8),
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 6))
        # 表头
        headers_th = ["项目", "偏低上限", "正常范围", "偏高下限"]
        for j, h in enumerate(headers_th):
            tk.Label(tab_thyroid,text=h,bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),anchor="w",
            ).grid(row=2, column=j, sticky="w", pady=(0, 4), padx=(0 if j == 0 else 8, 0))

        # 获取当前阈值配置
        current_th = getattr(甲功,"THYROID_RULES", {}).copy()
        th_units = getattr(甲功,"UNITS", {})  # 获取单位配置
        th_low_vars, th_norm_vars, th_high_vars = {}, {}, {}
        
        # 动态创建输入框
        base_row_th = 3
        for i, (proj, rule) in enumerate(current_th.items(), start=base_row_th):
            # 提取当前阈值
            low_max = rule.get("low_max")        # 偏低上限
            normal_low = rule.get("normal_low")  # 正常范围下限
            normal_high = rule.get("normal_high") # 正常范围上限
            high_min = rule.get("high_min")      # 偏高下限
                    
            # 获取单位（如果没有配置则留空）
            unit = th_units.get(proj, "")
        
            # 项目名称标签（带单位）
            proj_text = f"{proj} ({unit})" if unit else proj
            tk.Label(tab_thyroid,text=proj_text, bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9),anchor="w",width=16,
            ).grid(row=i, column=0, sticky="w", pady=3)

            # 创建输入变量
            low_var = tk.StringVar(value="" if low_max is None else str(low_max))
            norm_text = ""
            if normal_low is not None and normal_high is not None:
                norm_text = f"{normal_low}-{normal_high}"  # 正常范围格式
            norm_var = tk.StringVar(value=norm_text)
            high_var = tk.StringVar(value="" if high_min is None else str(high_min))

            # 创建输入控件
            tk.Entry(tab_thyroid, textvariable=low_var, width=10, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=1, sticky="w", pady=3, padx=(8, 0))
            tk.Entry(tab_thyroid, textvariable=norm_var, width=14, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=2, sticky="w", pady=3, padx=(8, 0))
            tk.Entry(tab_thyroid, textvariable=high_var, width=10, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=3, sticky="w", pady=3, padx=(8, 0))

            # 保存变量引用
            th_low_vars[proj] = low_var
            th_norm_vars[proj] = norm_var
            th_high_vars[proj] = high_var

        # ==================== 肿瘤分页 ====================
        import 肿瘤  # 导入肿瘤模块
        current_tumor = getattr(肿瘤,"TUMOR_RULES", {}).copy() if hasattr(肿瘤,"TUMOR_RULES") else {}
        tumor_units = getattr(肿瘤,"UNITS", {})  # 获取单位配置
                
        # 即使肿瘤项目不存在,也显示分页（保持界面一致性）
        tk.Label(tab_tumor,text="肿瘤（阴性 / 阳性）",bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))
        
        tk.Label(tab_tumor,text="请确认单位",bg="#F7F9FB",fg="#7B8A8B",justify="left",font=("微软雅黑", 8),
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 6))
        
        # 表头
        headers_tumor = ["项目", "阈值"]
        for j, h in enumerate(headers_tumor):
            tk.Label(tab_tumor,text=h,bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9, "bold"),anchor="w",
            ).grid(row=2, column=j, sticky="w", pady=(0, 4), padx=(0 if j == 0 else 8, 0))
        
        tumor_threshold_vars = {}  # 存储肿瘤阈值变量
                
        # 动态创建输入框
        base_row_tumor = 3
        for i, (proj, rule) in enumerate(current_tumor.items(), start=base_row_tumor):
            # 取阴性上限和阳性下限的最大值作为阈值（简化处理）
            neg_max = rule.get("neg_max")
            pos_min = rule.get("pos_min")
            threshold = max(neg_max or 0, pos_min or 0) if (neg_max is not None or pos_min is not None) else None
                    
            # 获取单位（如果没有配置则留空）
            unit = tumor_units.get(proj, "")
        
            # 项目名称标签（带单位）
            proj_text = f"{proj} ({unit})" if unit else proj
            tk.Label(tab_tumor,text=proj_text,bg="#F7F9FB",fg="#2C3E50",font=("微软雅黑", 9),anchor="w",width=16,
            ).grid(row=i, column=0, sticky="w", pady=3)

            # 阈值输入框
            threshold_var = tk.StringVar(value="" if threshold is None else str(threshold))
            tk.Entry(tab_tumor, textvariable=threshold_var, width=20, font=("微软雅黑", 9), bd=1, relief="solid").grid(
                row=i, column=1, sticky="w", pady=3, padx=(8, 0))

            # 保存变量引用
            tumor_threshold_vars[proj] = threshold_var

        final_row = base_row_tumor + len(current_tumor)

        # ==================== 按钮区域 ====================
        # 按钮容器（放置在Notebook下方）
        btn_frame = tk.Frame(frm, bg="#F7F9FB")
        btn_frame.pack(side="bottom", pady=(12, 0), fill="x", anchor="e")  # 右下角对齐

        # 保存并关闭函数
        def save_and_close():
                # ---- 保存术前八项阈值 ----
            new_pre = {}
            for proj in current_pre.keys():
                neg_txt = pre_neg_vars[proj].get().strip()    # 获取阴性上限输入
                gray_txt = pre_gray_vars[proj].get().strip()  # 获取灰区范围输入
                pos_txt = pre_pos_vars[proj].get().strip()    # 获取阳性下限输入
                # 验证阴性上限
                try:
                    neg_v = float(neg_txt) if neg_txt != "" else None
                except ValueError:
                    messagebox.showerror("错误", f"{proj} 的阴性上限必须为数字。")
                    return
                # 验证阳性下限
                try:
                    pos_v = float(pos_txt) if pos_txt != "" else None
                except ValueError:
                    messagebox.showerror("错误", f"{proj} 的阳性下限必须为数字。")
                    return
                # 验证非负数约束
                if neg_v is not None and neg_v < 0:
                    messagebox.showerror("错误", f"{proj} 的阴性上限不能为负数。")
                    return
                if pos_v is not None and pos_v < 0:
                    messagebox.showerror("错误", f"{proj} 的阳性下限不能为负数。")
                    return
                # 解析灰区范围
                gray_low = gray_high = None
                if gray_txt:
                    parts = gray_txt.replace(",", ",").replace("－", "-").split("-")  # 支持中文逗号和减号
                    if len(parts) != 2:
                        messagebox.showerror("错误", f"{proj} 的灰区范围格式应为 类似 1-10。")
                        return
                    try:
                        gray_low = float(parts[0].strip())
                        gray_high = float(parts[1].strip())
                    except ValueError:
                        messagebox.showerror("错误", f"{proj} 的灰区范围必须为数字,例如 1-10。")
                        return
                    if gray_low < 0 or gray_high < 0 or gray_low >= gray_high:
                        messagebox.showerror("错误", f"{proj} 的灰区范围必须满足 0 ≤ 下限 < 上限。")
                        return
                # 保存该项目的阈值规则
                new_pre[proj] = {
                    "neg_max": neg_v,      # 阴性上限
                    "gray_low": gray_low,  # 灰区下限
                    "gray_high": gray_high, # 灰区上限
                    "pos_min": pos_v,      # 阳性下限
                }
            # ---- 保存甲功阈值 ----
            new_th = {}
            for proj in current_th.keys():
                low_txt = th_low_vars[proj].get().strip()     # 偏低上限
                norm_txt = th_norm_vars[proj].get().strip()   # 正常范围
                high_txt = th_high_vars[proj].get().strip()   # 偏高下限
                # 验证偏低上限
                try:
                    low_v = float(low_txt) if low_txt != "" else None
                except ValueError:
                    messagebox.showerror("错误", f"{proj} 的偏低上限必须为数字。")
                    return
                # 验证偏高下限
                try:
                    high_v = float(high_txt) if high_txt != "" else None
                except ValueError:
                    messagebox.showerror("错误", f"{proj} 的偏高下限必须为数字。")
                    return
                # 验证非负数约束
                if low_v is not None and low_v < 0:
                    messagebox.showerror("错误", f"{proj} 的偏低上限不能为负数。")
                    return
                if high_v is not None and high_v < 0:
                    messagebox.showerror("错误", f"{proj} 的偏高下限不能为负数。")
                    return
                # 解析正常范围
                normal_low = normal_high = None
                if norm_txt:
                    parts = norm_txt.replace(",", ",").replace("－", "-").split("-")
                    if len(parts) != 2:
                        messagebox.showerror("错误", f"{proj} 的正常范围格式应为 类似 2-4。")
                        return
                    try:
                        normal_low = float(parts[0].strip())
                        normal_high = float(parts[1].strip())
                    except ValueError:
                        messagebox.showerror("错误", f"{proj} 的正常范围必须为数字,例如 2-4。")
                        return
                    if normal_low < 0 or normal_high < 0 or normal_low >= normal_high:
                        messagebox.showerror("错误", f"{proj} 的正常范围必须满足 0 ≤ 下限 < 上限。")
                        return
                # 保存该项目的阈值规则
                new_th[proj] = {
                    "low_max": low_v,        # 偏低上限
                    "normal_low": normal_low, # 正常范围下限
                    "normal_high": normal_high, # 正常范围上限
                    "high_min": high_v,      # 偏高下限
                }
            # ---- 保存肿瘤阈值 ----
            new_tumor = {}
            for proj in current_tumor.keys():
                threshold_txt = tumor_threshold_vars[proj].get().strip()
                # 验证阈值输入
                try:
                    threshold = float(threshold_txt) if threshold_txt != "" else None
                except ValueError:
                    messagebox.showerror("错误", f"{proj} 的阈值必须为数字。")
                    return
                # 验证非负数约束
                if threshold is not None and threshold < 0:
                    messagebox.showerror("错误", f"{proj} 的阈值不能为负数。")
                    return
                # 保存该项目的阈值规则（阴性上限和阳性下限设为相同值）
                new_tumor[proj] = {
                    "neg_max": threshold,  # 阴性上限
                    "pos_min": threshold,  # 阳性下限
                }
            # ---- 同步配置到各个层级 ----
            # 更新实例变量
            self.thresholds_pre = new_pre
            self.thresholds_thyroid = new_th
            self.thresholds_tumor = new_tumor
            # 更新配置管理器
            self.config_mgr.thresholds_pre = new_pre
            self.config_mgr.thresholds_thyroid = new_th
            self.config_mgr.thresholds_tumor = new_tumor
            #更新各模块的全局变量（确保分析时使用新阈值）
            肿瘤.TUMOR_RULES = new_tumor
            术前.THRESHOLD_RULES = new_pre
            甲功.THYROID_RULES = new_th
            #保存到配置文件,实现持久化
            self.config_mgr.save_thresholds()
            # 提示成功并关闭对话框
            messagebox.showinfo("成功", "阈值已更新并保存,将在下一次分析时生效。")
            top.destroy()
        # 保存按钮
        tk.Button(btn_frame,text="保存",command=save_and_close,bg="#3498DB",fg="white",font=("微软雅黑", 9),padx=18,pady=2,
                  bd=0,cursor="hand2",activebackground="#2980B9",activeforeground="white",).pack(side="right", padx=(6, 0))
        # 取消按钮
        tk.Button(btn_frame,text="取消",command=top.destroy,bg="#ECF0F1",fg="#2C3E50",font=("微软雅黑", 9),padx=14,pady=2,
                  bd=0,cursor="hand2",activebackground="#D0D7DE",activeforeground="#2C3E50",).pack(side="right")

    # ==================== 选项卡切换方法 ====================
    def _switch_top_tab(self, idx):
        """
            idx: 选项卡索引（0=术前八项,1=甲功,2=肿瘤,3=原始数据）

        """
        self._top_tab_selected = idx  # 记录当前选中索引
        
        # 更新所有选项卡按钮的样式
        for i in range(4):
            btn = getattr(self, "_top_tab_btn_%d" % i)  # 获取对应按钮对象
            if i == idx:
                # 选中状态：蓝色背景,白色文字
                btn.config(bg="#3498DB", fg="white", activebackground="#2980B9", activeforeground="white")
            else:
                # 未选中状态：灰色背景,深色文字
                btn.config(bg="#BDC3C7", fg="#2C3E50", activebackground="#ECF0F1", activeforeground="#2C3E50")
        
        # 根据选中索引控制各选项卡内容的显示/隐藏
        if idx == 0:
            # 显示术前八项：显示其Notebook,隐藏其他
            self.notebook.pack(fill="both", expand=True)
            self.notebook_甲功.pack_forget()
            self.notebook_肿瘤.pack_forget()
        elif idx == 1:
            # 显示甲功：显示其Notebook,隐藏其他
            self.notebook.pack_forget()
            self.notebook_甲功.pack(fill="both", expand=True)
            self.notebook_肿瘤.pack_forget()
        elif idx == 2:
            # 显示肿瘤：显示其Notebook,隐藏其他
            self.notebook.pack_forget()
            self.notebook_甲功.pack_forget()
            self.notebook_肿瘤.pack(fill="both", expand=True)
        else:
            # 显示原始数据：隐藏所有Notebook
            self.notebook.pack_forget()
            self.notebook_甲功.pack_forget()
            self.notebook_肿瘤.pack_forget()
            
        # 将选中选项卡的内容框架提升到最顶层显示
        self._top_frames[idx].lift()

    # ==================== 数据处理辅助 ====================
    def _find_col_by_keyword(self, columns, keyword):
        for c in columns:
            if keyword in str(c):
                return c
        return None
    
    def _get_dilution_value(self, row, display_df):
       
        # 查找所有包含"稀释倍数"的列
        dilution_cols = [c for c in display_df.columns if '稀释倍数' in str(c)]
        
        if not dilution_cols:
            return ''  # 未找到相关列,返回空字符串
        
        # 如果只有一列,直接返回
        if len(dilution_cols) == 1:
            val = row.get(dilution_cols[0], '')  # 获取该列的值
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return ''
            # 清理格式：去除等号和引号
            return str(val).strip().replace('=', '').replace('"', '').replace("'", '')
        
        # 如果有多列（i6000的情况）,取最大值
        values = []
        for col in dilution_cols:
            val = row.get(col, '')
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            # 尝试转换为数字
            try:
                num_val = float(str(val).strip().replace('=', '').replace('"', '').replace("'", ''))
                values.append(num_val)
            except:
                continue  # 转换失败则跳过该值
        
        if not values:
            return ''  # 没有有效数值,返回空字符串
        
        # 返回最大值（i6000没稀释时显示1,有稀释时显示实际倍数）
        return str(int(max(values)) if max(values).is_integer() else max(values))

    # ====================右键菜单和数据删除功能 ====================
    def _show_tree0_context_menu(self, event):
    
        menu = tk.Menu(self.root, tearoff=0)                    # 创建无边框菜单
        menu.add_command(label="删除数据", command=self._delete_selected_tree0_rows)  # 添加删除命令
        try:
            menu.tk_popup(event.x_root, event.y_root)           # 在鼠标位置显示菜单
        finally:
            menu.grab_release()                                 # 释放菜单焦点

    def _delete_selected_tree0_rows(self):
       
        sel = self.tree0.selection()                            # 获取选中的行ID列表
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的行")
            return
        
        # 收集要删除的记录标识
        to_drop = set()
        for iid in sel:
            vals = self.tree0.item(iid)["values"]               # 获取该行的所有列值
            if len(vals) >= 4:
                # 检查是否为i6000数据（没有真正序号列的情况）
                if self.full_df is not None and '样本ID' in self.full_df.columns and '序号' not in self.full_df.columns:
                    # i6000使用样本ID, 测试项目作为唯一标识
                    to_drop.add((str(vals[2]).strip(), str(vals[3]).strip()))  # 样本ID, 测试项目
                else:
                    # 其他设备使用序号, 测试项目作为唯一标识
                    to_drop.add((str(vals[0]).strip(), str(vals[3]).strip()))  # 序号, 测试项目
        
        if not to_drop:
            return

        # 定义保留行的判断函数
        def keep_row(r):
            # 根据数据类型选择不同的唯一标识方式
            if '样本ID' in r and '序号' not in r:
                # i6000数据：使用样本ID + 测试项目作为唯一标识
                key = (str(r.get("样本ID", "")).strip(), str(r.get("测试项目", "")).strip())
            else:
                # 其他数据：使用序号 + 测试项目作为唯一标识
                key = (str(r.get("序号", "")).strip(), str(r.get("测试项目", "")).strip())
            return key not in to_drop  # 返回True表示保留该行

        # 从完整数据集中删除选中的行
        self.full_df = self.full_df[self.full_df.apply(keep_row, axis=1)].reset_index(drop=True)
        
        # 重新执行当前项目的分析（更新统计结果）
        current_proj = self.project_name.get()
        if current_proj == "术前八项":
            self.summary_map, self.mode_map, self.sample_mode_map, self.sample_mode_index_col = analyze_术前(self.full_df)
            if self.sample_mode_map is not None and not self.sample_mode_map.empty:
                self.sample_mode_map.index = self.sample_mode_map.index.astype(str).str.strip()
        elif current_proj == "肿瘤":
            self.summary_map, self.mode_map, _ = analyze_肿瘤(self.full_df)
            self.sample_mode_map = None
        elif current_proj == "甲功":
            self.summary_map, self.mode_map, _ = analyze_甲功(self.full_df)
            self.sample_mode_map = None
        else:
            self.summary_map, self.mode_map = pd.DataFrame(), pd.DataFrame()
            self.sample_mode_map = None
            
        # 刷新界面显示更新后的数据
        self.refresh_ui()

    def _show_tree_甲功0_context_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="删除数据", command=self._delete_selected_tree_甲功0_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _delete_selected_tree_甲功0_rows(self):
        sel = self.tree_甲功0.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的行")
            return
        to_drop = set()
        for iid in sel:
            vals = self.tree_甲功0.item(iid)["values"]
            if len(vals) >= 4:
                to_drop.add((str(vals[0]).strip(), str(vals[3]).strip()))  # 序号, 测试项目
        if not to_drop:
            return
        def keep_row(r):
            key = (str(r.get("序号", "")).strip(), str(r.get("测试项目", "")).strip())
            return key not in to_drop
        self.full_df = self.full_df[self.full_df.apply(keep_row, axis=1)].reset_index(drop=True)
        current_proj = self.project_name.get()
        if current_proj == "术前八项":
            self.summary_map, self.mode_map, self.sample_mode_map, self.sample_mode_index_col = analyze_术前(self.full_df)
            if self.sample_mode_map is not None and not self.sample_mode_map.empty:
                self.sample_mode_map.index = self.sample_mode_map.index.astype(str).str.strip()
        elif current_proj == "肿瘤":
            self.summary_map, self.mode_map, _ = analyze_肿瘤(self.full_df)
            self.sample_mode_map = None
        elif current_proj == "甲功":
            self.summary_map, self.mode_map, _ = analyze_甲功(self.full_df)
            self.sample_mode_map = None
        else:
            self.summary_map, self.mode_map = pd.DataFrame(), pd.DataFrame()
            self.sample_mode_map = None
        self._refresh_甲功_raw()

    def _show_tree_肿瘤0_context_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="删除数据", command=self._delete_selected_tree_肿瘤0_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _delete_selected_tree_肿瘤0_rows(self):
        sel = self.tree_肿瘤0.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的行")
            return
        to_drop = set()
        for iid in sel:
            vals = self.tree_肿瘤0.item(iid)["values"]
            if len(vals) >= 4:
                to_drop.add((str(vals[0]).strip(), str(vals[3]).strip()))  # 序号, 测试项目
        if not to_drop:
            return
        def keep_row(r):
            key = (str(r.get("序号", "")).strip(), str(r.get("测试项目", "")).strip())
            return key not in to_drop
        self.full_df = self.full_df[self.full_df.apply(keep_row, axis=1)].reset_index(drop=True)
        current_proj = self.project_name.get()
        if current_proj == "术前八项":
            self.summary_map, self.mode_map, self.sample_mode_map, self.sample_mode_index_col = analyze_术前(self.full_df)
            if self.sample_mode_map is not None and not self.sample_mode_map.empty:
                self.sample_mode_map.index = self.sample_mode_map.index.astype(str).str.strip()
        elif current_proj == "肿瘤":
            self.summary_map, self.mode_map, _ = analyze_肿瘤(self.full_df)
            self.sample_mode_map = None
        elif current_proj == "甲功":
            self.summary_map, self.mode_map, _ = analyze_甲功(self.full_df)
            self.sample_mode_map = None
        else:
            self.summary_map, self.mode_map = pd.DataFrame(), pd.DataFrame()
            self.sample_mode_map = None
        self._refresh_肿瘤_raw()

    def _set_progress(self, percent, status_text):
        """在后台线程中调用,通过主线程更新进度条和状态文字。"""
        self.root.after(0, lambda: (self.progress.__setitem__("value", percent), self.status_var.set(status_text)))


    # ==================== 核心业务方法 ====================
    
    def run(self):
        
        path = self.file_path.get()              # 获取文件路径
        machine = self.machine_type.get()        # 获取仪器型号
        
        # 参数验证：检查必要信息是否完整
        if not path or not machine:
            return messagebox.showwarning("警告", "请先选择数据文件和仪器型号")
        
        # UI状态更新：禁用按钮防止重复点击
        self.run_btn.config(state="disabled")
        self.progress['value'] = 0               # 重置进度条
        self.status_var.set("准备中...")         # 更新状态文字
        # 重置状态标签颜色为红色（待分析状态）
        self.status_label.config(fg="#E67E22")
        
        # 启动后台处理线程（daemon=True确保主程序退出时线程自动结束）
        threading.Thread(target=self._heavy_process, args=(path, machine), daemon=True).start()

    def _heavy_process(self, path, machine):
        
        try:
            # -------------------- 文件格式探测 --------------------
            self._set_progress(5, "正在探测文件格式...")
            
            # 尝试多种编码和引擎读取文件首行以确定格式
            df_probe = None
            for probe_reader in [
                lambda: pd.read_csv(path, sep='\t', encoding='utf-16', header=None, nrows=1),  # UTF-16编码TSV
                lambda: pd.read_csv(path, sep='\t', encoding='ANSI', header=None, nrows=1),  # ANSI编码TSV
                lambda: pd.read_excel(path, header=None, nrows=1),                             # Excel默认引擎
                lambda: pd.read_excel(path, header=None, nrows=1, engine='calamine')          # Calamine引擎
            ]:
                try:
                    df_probe = probe_reader()
                    if df_probe is not None: break  # 成功读取则跳出循环
                except:
                    continue  # 失败则尝试下一种方式

            if df_probe is None:
                raise Exception("探测失败：无法识别文件编码或格式。")

            # 检测文件是否有表头（通过检查第12列是否有数据）
            l1_val = df_probe.iloc[0, 11] if df_probe.shape[1] > 11 else None
            has_l1 = pd.notna(l1_val) and str(l1_val).strip() != ""

            # -------------------- 2. 仪器类型验证 --------------------
            # i3000有表头,i6000无表头,通过这个特征验证用户选择的仪器类型是否正确
            if machine == "i3000" and not has_l1:
                raise Exception("仪器不匹配：检测到的数据格式更像i6000,请检查仪器选择")
            if machine == "i6000" and has_l1:
                raise Exception("仪器不匹配：检测到的数据格式更像i3000,请检查仪器选择")

            # -------------------- 3. 读取完整文件 --------------------
            self._set_progress(15, "正在读取文件...")
            
            # 根据仪器类型设置跳过的行数（表头行数）
            skiprows = 2 if machine == "i3000" else 1
            df_raw = None
            
            # 尝试多种方式读取完整文件
            for reader in [
                lambda: pd.read_csv(path, sep='\t', encoding='utf-16', skiprows=2, dtype=object),  # UTF-16 TSV
                lambda: pd.read_csv(path, sep='\t', encoding='ANSI', skiprows=2, dtype=object),  # ANSI TSV
                lambda: pd.read_excel(path, skiprows=skiprows, dtype=object),                     # Excel默认
                lambda: pd.read_excel(path, skiprows=skiprows, dtype=object, engine='calamine')   # Calamine引擎
            ]:
                try:
                    df_raw = reader()
                    if df_raw is not None: break
                except:
                    continue

            if df_raw is None: 
                raise Exception("无法读取该文件格式,请确认文件完整性")

            # -------------------- 数据清洗 --------------------
            self._set_progress(35, "正在清洗数据...")
            
            # 根据仪器类型调用对应的清洗函数
            if machine == "i3000":
                processed_df = clean_i3000(df_raw)    # 调用i3000专用清洗函数
            else:
                processed_df = clean_i6000(df_raw)    # 调用i6000专用清洗函数


            # -------------------- 项目分析 --------------------
            self._set_progress(60, "正在分析统计...")
            current_proj = self.project_name.get()    # 获取当前选择的项目类型

            # 根据项目类型调用对应的分析函数
            if current_proj == "术前八项":
                # 术前八项分析：返回统计汇总、乙肝模式分布、样本-模式映射
                self.summary_map, self.mode_map, self.sample_mode_map, self.sample_mode_index_col = analyze_术前(processed_df)
                # 规范化样本ID索引格式（去除首尾空格）
                if self.sample_mode_map is not None and not self.sample_mode_map.empty:
                    self.sample_mode_map.index = self.sample_mode_map.index.astype(str).str.strip()
                self._last_hint = None  # 清除提示信息
                
            elif current_proj == "肿瘤":
                # 肿瘤标志物分析
                self.summary_map, self.mode_map, self._last_hint = analyze_肿瘤(processed_df)
                self.sample_mode_map = None  # 肿瘤项目不需要模式映射
                
            elif current_proj == "甲功":
                # 甲状腺功能分析
                self.summary_map, self.mode_map, self._last_hint = analyze_甲功(processed_df)
                self.sample_mode_map = None  # 甲功项目不需要模式映射
                
            else:
                # 未知项目,初始化空数据
                self.summary_map, self.mode_map = pd.DataFrame(), pd.DataFrame()
                self._last_hint = None
                self.sample_mode_map = None

            # -------------------- 完成处理 --------------------
            self.full_df = processed_df              # 保存处理后的完整数据
            self._set_progress(95, "即将完成...")
            self.root.after(0, self._process_complete)  # 回到主线程更新UI

        except Exception as e:
            # 异常处理：在主线程中显示错误信息
            self.root.after(0, lambda: messagebox.showerror("运行故障", f"处理失败: {str(e)}"))
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            # 发生错误时也重置状态标签颜色为红色
            self.root.after(0, lambda: self.status_label.config(fg="#E67E22"))

    def _process_complete(self):
     
        self.progress['value'] = 100              # 进度条满格
        self.refresh_ui()                         # 刷新所有表格显示分析结果
        self.status_var.set("分析完成")           # 更新状态文字
        # 将状态标签颜色改为绿色表示成功
        self.status_label.config(fg="#27AE60")
        self.run_btn.config(state="normal")       # 恢复运行按钮
        self.export_btn.config(state="normal")    # 启用导出按钮
        
        # 同步导出按钮颜色为蓝色（与开始分析按钮一致）
        self.export_btn.config(bg="#3498DB", fg="white", activebackground="#2980B9", activeforeground="white")
        # 重新绑定悬停效果以匹配新的颜色配置
        self._add_hover_effect(self.export_btn, "#2980B9", "#3498DB", "white", "white")
        
    # ==================== UI辅助方法 ====================
    
    def _sort_tree0(self, col):
       
        df_col = self._tree0_sort_col_map.get(col)
        if df_col is None and col == 'dilution':
            df_col = getattr(self, '_tree0_dilution_col', None)
        if df_col is None and col == 'recheck':
            df_col = getattr(self, '_tree0_recheck_col', None)
        if not df_col: return
        
        # 检查是否为同一列的再次点击（用于切换升序/降序）
        if hasattr(self, 'tree0_sort_col') and self.tree0_sort_col == df_col:
            # 同一列,切换排序方向
            self.tree0_sort_reverse = not self.tree0_sort_reverse
        else:
            # 不同列或首次排序,设置新列并默认升序
            self.tree0_sort_col = df_col
            self.tree0_sort_reverse = False
        
        # 对所有列都使用数字排序
        # 设置一个特殊标志,表明这是数字排序
        self.tree0_numeric_sort = True
                
        self.refresh_ui()

    def _refresh_甲功_raw(self):
        
        for i in self.tree_甲功0.get_children():
            self.tree_甲功0.delete(i)
        if self.full_df is None or self.project_name.get() != "甲功":
            return
        display_df = self.full_df.copy()
        ffill_cols = [c for c in ['序号', '样本号', '样本ID', '检测完成时间', '试剂批号', '单位'] if c in display_df.columns]
        for col in ffill_cols:
            s = display_df[col].astype(object).replace('', pd.NA)
            display_df[col] = s.ffill().fillna('')
        def _str(v):
            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
            s = str(v).strip()
            # 移除等号和双引号
            s = s.replace('=', '').replace('"', '').replace('\'', '')
            return s
        kw = self.search_var_甲功.get().strip().lower()
        if kw:
            display_df = display_df[display_df['样本ID'].astype(str).str.lower().str.contains(kw, na=False)]
        # 检查是否有手动排序设置
        if hasattr(self, 'tree_甲功0_sort_col') and self.tree_甲功0_sort_col and self.tree_甲功0_sort_col in display_df.columns:
            # 使用手动排序
            if hasattr(self, 'tree_甲功0_numeric_sort') and self.tree_甲功0_numeric_sort and self.tree_甲功0_sort_col in ['序号', '样本号', '样本ID','结果']:
                # 对ID列使用数字排序
                try:
                    # 定义清理函数
                    def clean_value(v):
                        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                        s = str(v).strip()
                        # 移除等号和双引号
                        s = s.replace('=', '').replace('"', '').replace('\'', '')
                        return s
                    
                    # 先清理要排序的列数据
                    cleaned_col_data = display_df[self.tree_甲功0_sort_col].apply(clean_value)
                    # 先保存原始列用于次级排序（处理NaN值和相同数字值的情况）
                    display_df['_original_value'] = cleaned_col_data
                    display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')
                    # 按数字排序,原值作为次级排序键以确保稳定性
                    display_df = display_df.sort_values(
                        by=['_sort_key', '_original_value'], 
                        ascending=[not self.tree_甲功0_sort_reverse, True], 
                        na_position='last')
                    # 删除临时列
                    display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                except:
                    # 如果数字转换失败,使用普通排序
                    display_df = display_df.sort_values(
                        by=[self.tree_甲功0_sort_col], ascending=not self.tree_甲功0_sort_reverse, na_position='last')
            else:
                display_df = display_df.sort_values(
                    by=[self.tree_甲功0_sort_col], ascending=not self.tree_甲功0_sort_reverse, na_position='last')
        else:
            # 自动排序 - 尝试按多列进行数字排序
            sort_columns = ['序号', '样本号', '样本ID', '结果']
            for col in sort_columns:
                if col in display_df.columns:
                    try:
                        # 定义清理函数
                        def clean_value(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                            s = str(v).strip()
                            # 移除等号和双引号
                            s = s.replace('=', '').replace('"', '').replace('\'', '')
                            return s
                        
                        # 先清理要排序的列数据
                        cleaned_col_data = display_df[col].apply(clean_value)
                        # 先保存原始列用于次级排序（处理NaN值和相同数字值的情况）
                        display_df['_original_value'] = cleaned_col_data
                        display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')
                        # 按数字排序,原值作为次级排序键以确保稳定性
                        display_df = display_df.sort_values(
                            by=['_sort_key', '_original_value'], 
                            ascending=[True, True], 
                            na_position='last')
                        # 删除临时列
                        display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                        break  # 找到合适的排序列后停止
                    except:
                        # 如果转换失败,尝试下一列
                        continue
        
        recheck_col = self._find_col_by_keyword(display_df.columns, '复查结果')
        for _, r in display_df.iterrows():
            sid = _str(r.get('样本ID', ''))
            dilution_val = self._get_dilution_value(r, display_df)  # 使用新方法获取稀释倍数
            recheck_val = _str(r.get(recheck_col, '')) if recheck_col else ''
            self.tree_甲功0.insert("", "end", values=(
                _str(r.get('序号', '')), _str(r.get('样本号', '')), sid, _str(r.get('测试项目', '')),
                _str(r.get('结果', '')), _str(r.get('单位', '')), dilution_val, _str(r.get('检测完成时间', '')),
                _str(r.get('试剂批号', '')), recheck_val
            ))

    def _sort_tree_甲功0(self, col):
        # 为甲功原始数据表头添加排序功能
        df_col = self._tree0_sort_col_map.get(col)
        if df_col is None and col == 'dilution':
            df_col = getattr(self, '_tree0_dilution_col', None)
        if df_col is None and col == 'recheck':
            df_col = getattr(self, '_tree0_recheck_col', None)
        if not df_col: return
        
        # 检查是否为同一列的再次点击（用于切换升序/降序）
        if hasattr(self, 'tree_甲功0_sort_col') and self.tree_甲功0_sort_col == df_col:
            # 同一列,切换排序方向
            self.tree_甲功0_sort_reverse = not self.tree_甲功0_sort_reverse
        else:
            # 不同列或首次排序,设置新列并默认升序
            self.tree_甲功0_sort_col = df_col
            self.tree_甲功0_sort_reverse = False
        
        # 检查是否为 ID 列或结果列,如果是则始终使用数字排序
        id_columns = ['序号', '样本号', '样本 ID', '结果']
        if df_col in id_columns:
            # 设置一个特殊标志,表明这是数字排序
            self.tree_甲功0_numeric_sort = True
        else:
            # 清除数字排序标志
            if hasattr(self, 'tree_甲功0_numeric_sort'):
                delattr(self, 'tree_甲功0_numeric_sort')
                
        self._refresh_甲功_raw()

    def _refresh_肿瘤_raw(self):        
        for i in self.tree_肿瘤0.get_children():
            self.tree_肿瘤0.delete(i)
        if self.full_df is None or self.project_name.get() != "肿瘤":
            return
        display_df = self.full_df.copy()
        ffill_cols = [c for c in ['序号', '样本号', '样本ID', '检测完成时间', '试剂批号', '单位'] if c in display_df.columns]
        for col in ffill_cols:
            s = display_df[col].astype(object).replace('', pd.NA)
            display_df[col] = s.ffill().fillna('')
        def _str(v):
            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
            s = str(v).strip()
            # 移除等号和双引号
            s = s.replace('=', '').replace('"', '').replace('\'', '')
            return s
        kw = self.search_var_肿瘤.get().strip().lower()
        if kw:
            display_df = display_df[display_df['样本ID'].astype(str).str.lower().str.contains(kw, na=False)]
        # 检查是否有手动排序设置
        if hasattr(self, 'tree_肿瘤0_sort_col') and self.tree_肿瘤0_sort_col and self.tree_肿瘤0_sort_col in display_df.columns:
            # 使用手动排序
            if hasattr(self, 'tree_肿瘤0_numeric_sort') and self.tree_肿瘤0_numeric_sort and self.tree_肿瘤0_sort_col in ['序号', '样本号', '样本ID','结果']:
                # 对ID列使用数字排序
                try:
                    # 定义清理函数
                    def clean_value(v):
                        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                        s = str(v).strip()
                        # 移除等号和双引号
                        s = s.replace('=', '').replace('"', '').replace('\'', '')
                        return s
                    
                    # 先清理要排序的列数据
                    cleaned_col_data = display_df[self.tree_肿瘤0_sort_col].apply(clean_value)
                    # 先保存原始列用于次级排序（处理NaN值和相同数字值的情况）
                    display_df['_original_value'] = cleaned_col_data
                    display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')
                    # 按数字排序,原值作为次级排序键以确保稳定性
                    display_df = display_df.sort_values(
                        by=['_sort_key', '_original_value'], 
                        ascending=[not self.tree_肿瘤0_sort_reverse, True], 
                        na_position='last')
                    # 删除临时列
                    display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                except:
                    # 如果数字转换失败,使用普通排序
                    display_df = display_df.sort_values(
                        by=[self.tree_肿瘤0_sort_col], ascending=not self.tree_肿瘤0_sort_reverse, na_position='last')
            else:
                display_df = display_df.sort_values(
                    by=[self.tree_肿瘤0_sort_col], ascending=not self.tree_肿瘤0_sort_reverse, na_position='last')
        else:
            # 自动排序 - 尝试按多列进行数字排序
            sort_columns = ['序号', '样本号', '样本ID', '结果']
            for col in sort_columns:
                if col in display_df.columns:
                    try:
                        # 定义清理函数
                        def clean_value(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                            s = str(v).strip()
                            # 移除等号和双引号
                            s = s.replace('=', '').replace('"', '').replace('\'', '')
                            return s
                        
                        # 先清理要排序的列数据
                        cleaned_col_data = display_df[col].apply(clean_value)
                        # 先保存原始列用于次级排序（处理NaN值和相同数字值的情况）
                        display_df['_original_value'] = cleaned_col_data
                        display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')
                        # 按数字排序,原值作为次级排序键以确保稳定性
                        display_df = display_df.sort_values(
                            by=['_sort_key', '_original_value'], 
                            ascending=[True, True], 
                            na_position='last')
                        # 删除临时列
                        display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                        break  # 找到合适的排序列后停止
                    except:
                        # 如果转换失败,尝试下一列
                        continue
        
        recheck_col = self._find_col_by_keyword(display_df.columns, '复查结果')
        for _, r in display_df.iterrows():
            sid = _str(r.get('样本ID', ''))
            dilution_val = self._get_dilution_value(r, display_df)  # 使用新方法获取稀释倍数
            recheck_val = _str(r.get(recheck_col, '')) if recheck_col else ''
            self.tree_肿瘤0.insert("", "end", values=(
                _str(r.get('序号', '')), _str(r.get('样本号', '')), sid, _str(r.get('测试项目', '')),
                _str(r.get('结果', '')), _str(r.get('单位', '')), dilution_val, _str(r.get('检测完成时间', '')),
                _str(r.get('试剂批号', '')), recheck_val
            ))

    def _sort_tree_肿瘤0(self, col):
        # 为肿瘤原始数据表头添加排序功能
        df_col = self._tree0_sort_col_map.get(col)
        if df_col is None and col == 'dilution':
            df_col = getattr(self, '_tree0_dilution_col', None)
        if df_col is None and col == 'recheck':
            df_col = getattr(self, '_tree0_recheck_col', None)
        if not df_col: return
        
        # 检查是否为同一列的再次点击（用于切换升序/降序）
        if hasattr(self, 'tree_肿瘤0_sort_col') and self.tree_肿瘤0_sort_col == df_col:
            # 同一列,切换排序方向
            self.tree_肿瘤0_sort_reverse = not self.tree_肿瘤0_sort_reverse
        else:
            # 不同列或首次排序,设置新列并默认升序
            self.tree_肿瘤0_sort_col = df_col
            self.tree_肿瘤0_sort_reverse = False
        
        # 检查是否为 ID 列或结果列,如果是则始终使用数字排序
        id_columns = ['序号', '样本号', '样本 ID', '结果']
        if df_col in id_columns:
            # 设置一个特殊标志,表明这是数字排序
            self.tree_肿瘤0_numeric_sort = True
        else:
            # 清除数字排序标志
            if hasattr(self, 'tree_肿瘤0_numeric_sort'):
                delattr(self, 'tree_肿瘤0_numeric_sort')
                
        self._refresh_肿瘤_raw()


    # ==================== UI刷新方法 ====================
    
    def refresh_ui(self):
        # 数据有效性检查
        if self.full_df is None: 
            return
        
        # -------------------- 清空所有表格 --------------------
        # 清空主选项卡的表格
        for tree in [self.tree0, self.tree1, self.tree2]:
            for i in tree.get_children(): 
                tree.delete(i)
        # 清空甲功和肿瘤选项卡的表格
        for i in self.tree_甲功1.get_children():
            self.tree_甲功1.delete(i)
        for i in self.tree_肿瘤1.get_children():
            self.tree_肿瘤1.delete(i)

        proj = self.project_name.get()  # 获取当前项目类型

        # -------------------- 更新筛选控件状态 --------------------
        if proj == "术前八项":
            # 术前八项：启用乙肝模式筛选
            if not self.mode_map.empty and '模式' in self.mode_map.columns:
                mode_opts = ["全部"] + self.mode_map['模式'].astype(str).tolist()  # 生成模式选项
                self.mode_filter_combo['values'] = mode_opts
                if self.mode_filter_var.get() not in mode_opts:
                    self.mode_filter_var.set("全部")
            else:
                self.mode_filter_combo['values'] = ["全部"]
                self.mode_filter_var.set("全部")
            self.mode_filter_combo.config(state="readonly")  # 启用下拉框
            
            # 项目筛选选项（用于多选）
            if self.full_df is not None and '测试项目' in self.full_df.columns:
                # 获取所有项目及其计数
                self.project_options_with_count = self.full_df['测试项目'].value_counts().to_dict()
                # 更新下拉框选项
                self._update_project_combo_options()
            else:
                self.project_options_with_count = {}
                self.project_filter_combo['values'] = ["全部"]
                self.project_filter_var.set("全部")
        else:
            # 非术前八项：禁用乙肝模式筛选
            self.mode_filter_combo['values'] = ["全部"]
            self.mode_filter_var.set("全部")
            self.mode_filter_combo.config(state="disabled")
            
            # 项目筛选选项（用于多选）
            if self.full_df is not None and '测试项目' in self.full_df.columns:
                # 获取所有项目及其计数
                self.project_options_with_count = self.full_df['测试项目'].value_counts().to_dict()
                # 更新下拉框选项
                self._update_project_combo_options()
            else:
                self.project_options_with_count = {}
                self.project_filter_combo['values'] = ["全部"]
                self.project_filter_var.set("全部")

        # -------------------- 预处理 --------------------
        display_df = self.full_df.copy()  # 创建副本避免修改原数据
        
        # 前向填充必要列（处理合并单元格情况）
        ffill_cols = [c for c in ['序号', '样本号', '样本ID', '检测完成时间', '试剂批号', '单位'] if c in display_df.columns]
        for col in ffill_cols:
            s = display_df[col].astype(object).replace('', pd.NA)
            display_df[col] = s.ffill().fillna('')

        # 字符串清理函数
        def _str(v):
            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
            s = str(v).strip()
            s = s.replace('=', '').replace('"', '').replace('\'', '')  # 去除等号和引号
            return s

        # -------------------- 应用筛选条件 --------------------
        # （1）样本ID搜索筛选
        kw = self.search_var.get().strip().lower()
        if kw:
            display_df = display_df[display_df['样本ID'].astype(str).str.lower().str.contains(kw, na=False)]
        
        # （2）项目筛选（多选支持）
        if self.selected_projects_set and '测试项目' in display_df.columns:
            display_df = display_df[display_df['测试项目'].isin(self.selected_projects_set)]
        
        # （3）乙肝模式筛选（仅术前八项）
        if proj == "术前八项":
            mode_sel = self.mode_filter_var.get()
            if self.sample_mode_map is not None and mode_sel and mode_sel != "全部":
                idx_col = self.sample_mode_index_col
                if idx_col in display_df.columns:
                    key_series = display_df[idx_col].astype(str).str.strip()
                    mode_mapped = key_series.map(self.sample_mode_map)  # 映射样本到乙肝模式
                    display_df = display_df[mode_mapped == mode_sel]    # 筛选指定模式

        # -------------------- 数据排序 --------------------
        # 如果没有手动排序设置,则执行自动排序
        if not self.tree0_sort_col:
            sort_columns = ['序号', '样本号', '样本ID', '结果']  # 优先级排序列
            for col in sort_columns:
                if col in display_df.columns:
                    try:
                        # 定义清理函数用于排序
                        def clean_value(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                            s = str(v).strip()
                            s = s.replace('=', '').replace('"', '').replace('\'', '')
                            return s
                        
                        # 清理数据并转换为数值用于排序
                        cleaned_col_data = display_df[col].apply(clean_value)
                        display_df['_original_value'] = cleaned_col_data  # 保存原始值用于次级排序
                        display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')  # 转换为数值
                        
                        # 按数字排序,原值作为次级排序键以确保稳定性
                        display_df = display_df.sort_values(
                            by=['_sort_key', '_original_value'],
                            ascending=[True, True],  # 升序排列
                            na_position='last')      # 空值排在最后
                        
                        # 删除临时列
                        display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                        break  # 找到合适的排序列后停止
                    except:
                        continue  # 转换失败则尝试下一列
        else:
            # 有手动排序设置时使用指定列排序
            if self.tree0_sort_col in display_df.columns:
                # 对所有列都使用数字排序
                try:
                    def clean_value(v):
                        if v is None or (isinstance(v, float) and pd.isna(v)): return ''
                        s = str(v).strip()
                        s = s.replace('=', '').replace('"', '').replace('\'', '')
                        return s
                            
                    # 清理要排序的列数据
                    cleaned_col_data = display_df[self.tree0_sort_col].apply(clean_value)
                    display_df['_original_value'] = cleaned_col_data
                    display_df['_sort_key'] = pd.to_numeric(cleaned_col_data, errors='coerce')
                            
                    # 按数字排序,原值作为次级排序键
                    display_df = display_df.sort_values(
                        by=['_sort_key', '_original_value'],
                        ascending=[not self.tree0_sort_reverse, True],  # 根据排序方向
                        na_position='last')
                            
                    # 删除临时列
                    display_df = display_df.drop(['_sort_key', '_original_value'], axis=1)
                except:
                    # 如果数字转换失败,使用普通排序
                    display_df = display_df.sort_values(
                        by=[self.tree0_sort_col], ascending=not self.tree0_sort_reverse, na_position='last')

        # -------------------- 填充原始数据表格 --------------------
        recheck_col = self._find_col_by_keyword(display_df.columns, '复查结果')  # 查找复查结果列
        for _, r in display_df.iterrows():
            sid = _str(r.get('样本ID', ''))                          # 获取样本ID
            dilution_val = self._get_dilution_value(r, display_df)   # 获取稀释倍数
            recheck_val = _str(r.get(recheck_col, '')) if recheck_col else ''  # 获取复查结果
            self.tree0.insert("", "end", values=(                   # 插入表格行
                _str(r.get('序号', '')), _str(r.get('样本号', '')), sid, _str(r.get('测试项目', '')),
                _str(r.get('结果', '')), _str(r.get('单位', '')), dilution_val, _str(r.get('检测完成时间', '')),
                _str(r.get('试剂批号', '')), recheck_val
            ))

        # -------------------- 填充统计表格 --------------------
        if proj == "术前八项":
            # 填充阳性率统计表
            prev_proj = None
            for _, r in self.summary_map.iterrows():
                proj_name = r.get('测试项目', '')
                show_proj = '' if proj_name == prev_proj else proj_name  # 合并相同项目
                prev_proj = proj_name
                self.tree1.insert("", "end", values=(show_proj, r.get('灰区', 0),
                                                     r.get('灰区率', ''), r.get('阳性', 0), r.get('阳性率', ''),
                                                     r.get('总数', 0), r.get('试剂批号', '')))

            # 填充乙肝模式分布表
            for _, r in self.mode_map.iterrows():
                self.tree2.insert("", "end", values=(r.get('模式', ''), r.get('样本数', 0), r.get('占比', '')))

        # 仅当当前项目为「甲功」时填充甲功占比分析表
        if proj == "甲功" and not self.summary_map.empty and "偏低" in self.summary_map.columns:
            prev_proj = None
            for _, r in self.summary_map.iterrows():
                proj_name = r.get('测试项目', '')
                show_proj = '' if proj_name == prev_proj else proj_name
                prev_proj = proj_name
                self.tree_甲功1.insert("", "end", values=(
                    show_proj, r.get("偏低", 0), r.get("偏低率", ""),
                    r.get("正常", 0), r.get("正常率", ""), r.get("偏高", 0), r.get("偏高率", ""),
                    r.get("试剂批号", "")
                ))

        # 仅当当前项目为「肿瘤」时填充肿瘤占比分析表
        if proj == "肿瘤" and not self.summary_map.empty and "阳性" in self.summary_map.columns:
            prev_proj = None
            for _, r in self.summary_map.iterrows():
                proj_name = r.get('测试项目', '')
                show_proj = '' if proj_name == prev_proj else proj_name
                prev_proj = proj_name
                self.tree_肿瘤1.insert("", "end", values=(
                    show_proj, r.get("阳性", 0), r.get("阳性率", ""),
                    r.get("总数", 0), r.get("试剂批号", "")
                ))


    # ==================== 十、数据导出方法 ====================
    def export(self):
        """
        导出当前显示的数据（经过所有筛选条件过滤后的数据）
        """
        # 检查是否有数据
        if self.full_df is None:
            messagebox.showwarning("提示", "请先加载数据")
            return
        
        # 获取当前显示的数据（应用所有筛选条件）
        display_df = self._get_current_display_data()
        
        if display_df.empty:
            messagebox.showwarning("提示", "当前筛选条件下无数据可导出")
            return
        
        # 选择保存路径
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        if not path: 
            return
        
        try:
            # 根据文件扩展名选择导出格式
            if path.endswith('.csv'):
                display_df.to_csv(path, index=False, encoding='utf-8-sig')  # CSV格式导出
            else:
                # 写入Excel文件
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    display_df.to_excel(writer, sheet_name="筛选后数据", index=False)
                    # 如果有统计汇总数据,也一并导出
                    if not self.summary_map.empty:
                        self.summary_map.to_excel(writer, sheet_name="统计汇总", index=False)
                    if not self.mode_map.empty:
                        self.mode_map.to_excel(writer, sheet_name="模式分布", index=False)
            
            # 自动打开文件
            os.startfile(path)
            messagebox.showinfo("成功", f"数据已导出到:\n{path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败:\n{str(e)}")
    
    def _get_current_display_data(self):
        
        if self.full_df is None:
            return pd.DataFrame()
        
        display_df = self.full_df.copy()  # 创建数据副本
        
        # 前向填充必要的列（处理合并单元格情况）
        ffill_cols = [c for c in ['序号', '样本号', '样本ID', '检测完成时间', '试剂批号', '单位'] if c in display_df.columns]
        for col in ffill_cols:
            s = display_df[col].astype(object).replace('', pd.NA)
            display_df[col] = s.ffill().fillna('')
        
        # 应用样本ID搜索筛选
        kw = self.search_var.get().strip().lower()
        if kw:
            display_df = display_df[display_df['样本ID'].astype(str).str.lower().str.contains(kw, na=False)]
        
        # 应用项目筛选（多选支持）
        if self.selected_projects_set and '测试项目' in display_df.columns:
            display_df = display_df[display_df['测试项目'].isin(self.selected_projects_set)]
        
        # 应用乙肝模式筛选（仅术前八项）
        proj = self.project_name.get()
        if proj == "术前八项":
            mode_sel = self.mode_filter_var.get()
            if self.sample_mode_map is not None and mode_sel and mode_sel != "全部":
                idx_col = self.sample_mode_index_col
                if idx_col in display_df.columns:
                    key_series = display_df[idx_col].astype(str).str.strip()
                    mode_mapped = key_series.map(self.sample_mode_map)
                    display_df = display_df[mode_mapped == mode_sel]
        
        return display_df

    # ==================== 项目筛选功能 ====================
    def _on_project_selection(self, event):
        
        selected = self.project_filter_var.get()  # 获取当前选择的文本
        if selected == "全部":
            self.selected_projects_set.clear()     # 清空选中集合
        elif selected and selected != "全部":
            # 从"项目名 (数量)"格式中提取项目名
            if " (" in selected and selected.endswith(")"):
                project_name = selected.split(" (")[0]
                # 去掉可能的✓符号
                if project_name.startswith("✓ "):
                    project_name = project_name[2:]
            else:
                project_name = selected
            
            # 切换选中状态
            if project_name in self.selected_projects_set:
                self.selected_projects_set.discard(project_name)  # 取消选中
            else:
                self.selected_projects_set.add(project_name)      # 添加选中
        
        # 更新下拉框显示和选项
        self._update_project_combo_options()
        self.refresh_ui()  # 重新刷新界面应用筛选
    
    def _update_project_combo_options(self):
        
        if not hasattr(self, 'project_options_with_count') or not self.project_options_with_count:
            self.project_filter_combo['values'] = ["全部"]
            self.project_filter_var.set("全部")
            return
        
        # 生成选项列表
        options = ["全部"]
        # 按数量降序排列项目
        sorted_projects = sorted(self.project_options_with_count.items(), key=lambda x: x[1], reverse=True)
        
        for project, count in sorted_projects:
            if project in self.selected_projects_set:
                # 已选中项目添加✓符号
                options.append(f"✓ {project} ({count})")
            else:
                # 未选中项目
                options.append(f"{project} ({count})")
        
        self.project_filter_combo['values'] = options
        
        # 设置当前显示文本
        if not self.selected_projects_set:
            self.project_filter_var.set("全部")  # 无选中项目时显示"全部"
        else:
            # 显示选中项目数量
            self.project_filter_var.set(f"已选 {len(self.selected_projects_set)} 项")


# ==================== 程序入口 ====================
if __name__ == "__main__":
    root = tk.Tk()                    # 创建主窗口
    app = ModernApp(root)             # 初始化应用程序
    root.mainloop()                   # 启动事件循环
