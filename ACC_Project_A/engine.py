import yaml
import pandas as pd
import os


class ProjectConfig:
    def __init__(self, project_path):
        self.project_path = project_path
        self.load()
    
    def load(self):
        """加载所有配置文件"""
        self.project = self._load_yaml("project.yaml")
        self.functions = self._load_yaml("functions.yaml")["functions"]
        self.safety = self._load_yaml("safety.yaml")
        self.state_machine = self._load_csv("state_machine.csv")
        self.signals = self._load_excel("signals.xlsx")

    def _load_yaml(self, filename):
        """加载 YAML 文件"""
        path = os.path.join(self.project_path, filename)
        print("project_path =", self.project_path)
        print("filename     =", filename)
        print("full path    =", path)

        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)

    def _load_csv(self, filename):
        """加载 CSV 文件"""
        path = os.path.join(self.project_path, filename)
        return pd.read_csv(path)

    def _load_excel(self, filename):
        """加载 Excel 文件"""
        path = os.path.join(self.project_path, filename)
        return pd.read_excel(path)

    # 便捷接口
    def get_project_info(self):
        return self.project

    def get_functions(self):
        return self.functions

    def get_safety(self):
        return self.safety

    def get_state_machine(self):
        return self.state_machine

    def get_signals(self):
        return self.signals
