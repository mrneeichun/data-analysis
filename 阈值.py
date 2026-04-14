"""
阈值配置管理模块
负责加载、保存和管理三个项目（术前八项、甲功、肿瘤）的阈值配置
"""
import json
import os


class ConfigManager:
    """配置管理器：处理阈值的加载、保存和持久化"""
    
    def __init__(self):
        # 配置文件路径：保存在用户目录，移动exe位置也不会丢失配置
        app_data_dir = os.getenv("LOCALAPPDATA")
        if not app_data_dir:
            app_data_dir = os.path.expanduser("~")
        config_dir = os.path.join(app_data_dir, "数据分析")
        os.makedirs(config_dir, exist_ok=True)
        self.config_file = os.path.join(config_dir, "迈克阈值配置.json")
        
        # 存储三个项目的阈值
        self.thresholds_pre = {}
        self.thresholds_thyroid = {}
        self.thresholds_tumor = {}
    
    def load_thresholds(self, default_pre, default_thyroid, default_tumor):
        """
        从配置文件加载阈值，如果文件不存在则使用模块默认值
        
        参数:
            default_pre: 术前八项的默认阈值
            default_thyroid: 甲功的默认阈值
            default_tumor: 肿瘤的默认阈值
        
        返回:
            (术前阈值, 甲功阈值, 肿瘤阈值) 三元组
        """
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    pre = config.get("术前八项", default_pre)
                    th = config.get("甲功", default_thyroid)
                    tumor = config.get("肿瘤", default_tumor)
                    
                    # 防止配置文件被改坏：必须是字典
                    if not isinstance(pre, dict):
                        pre = default_pre
                    if not isinstance(th, dict):
                        th = default_thyroid
                    if not isinstance(tumor, dict):
                        tumor = default_tumor
                    
                    # 确保所有项目都存在（兼容新增项目的情况）
                    for k, v in default_pre.items():
                        if k not in pre or not isinstance(pre.get(k), dict):
                            pre[k] = v
                    for k, v in default_thyroid.items():
                        if k not in th or not isinstance(th.get(k), dict):
                            th[k] = v
                    for k, v in default_tumor.items():
                        if k not in tumor or not isinstance(tumor.get(k), dict):
                            tumor[k] = v
                    
                    self.thresholds_pre = pre
                    self.thresholds_thyroid = th
                    self.thresholds_tumor = tumor
                    return pre, th, tumor
            except Exception:
                # 如果读取失败，使用默认值
                self.thresholds_pre = default_pre
                self.thresholds_thyroid = default_thyroid
                self.thresholds_tumor = default_tumor
                return default_pre, default_thyroid, default_tumor
        
        # 配置文件不存在，使用默认值
        self.thresholds_pre = default_pre
        self.thresholds_thyroid = default_thyroid
        self.thresholds_tumor = default_tumor
        return default_pre, default_thyroid, default_tumor
    
    def save_thresholds(self):
        """将当前阈值保存到配置文件"""
        try:
            config = {
                "术前八项": self.thresholds_pre,
                "甲功": self.thresholds_thyroid,
                "肿瘤": self.thresholds_tumor
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            return False, str(e)
