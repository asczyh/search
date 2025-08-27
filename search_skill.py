import pandas as pd
import tkinter as tk
from tkinter import ttk
import re
import sys
import os

def get_exe_dir():
    """获取EXE所在目录（开发环境返回脚本目录）"""
    if getattr(sys, 'frozen', False):  # 打包环境
        return os.path.dirname(sys.executable)  # EXE所在目录
    return os.path.dirname(os.path.abspath(__file__))  # 开发环境脚本目录

# 获取Excel路径（与EXE同级）
FILE_PATH = os.path.join(get_exe_dir(), "Z战法.xlsx")
COLUMNS = ['Id', 'NameDesc', 'EffectParam1', 'AdditionPerAdd1']
#FILE_PATH = "Z战法.xlsx"
SHEET1_NAME = '战法表'
SHEET2_NAME = '战法Buff表'


class AutocompleteEntry:
    def __init__(self, root, options, row=0, column=0):
        self.root = root
        self.options = options

        # 创建输入框
        self.entry = tk.Entry(root, width=40)
        self.entry.grid(row=row, column=column, padx=10, pady=10, sticky="ew")

        # 创建预选列表框
        self.listbox = tk.Listbox(root, width=38, height=5)
        self.listbox.grid(row=row + 1, column=column, padx=10, sticky="nsew")
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        self.listbox.grid_remove()  # 初始隐藏

        # 设置默认提示文本
        self.entry.insert(0, "输入技能名...")
        self.entry.bind("<Button-1>", self.on_entry_click)

    def on_entry_click(self, event):
        """点击输入框时清空默认文本"""
        if self.entry.get() == "输入技能名...":
            self.entry.delete(0, tk.END)

    def on_key_release(self, event):
        """键盘输入时动态匹配选项"""
        current_text = self.entry.get().lower()
        self.listbox.delete(0, tk.END)

        if current_text:
            matches = [opt for opt in self.options if current_text in opt.lower()]
            if matches:
                for match in matches:
                    self.listbox.insert(tk.END, match)
                self.listbox.grid()  # 显示列表框
            else:
                self.listbox.grid_remove()
        else:
            self.listbox.grid_remove()

    def on_select(self, event):
        """选择列表框项后更新输入框并触发查询"""
        if self.listbox.curselection():
            selected = self.listbox.get(self.listbox.curselection()[0])
            self.entry.delete(0, tk.END)
            self.entry.insert(0, selected)
            self.listbox.grid_remove()
            # 触发查询
            self.root.event_generate("<<AutocompleteSelected>>")


class DataManager:
    def __init__(self, columns):
        self.columns = columns
        # 读取Excel数据
        self.sheet1 = pd.read_excel(
            FILE_PATH,
            sheet_name=SHEET1_NAME,
            dtype={'NameDesc': 'string', 'Id': 'string'}  # 确保ID为字符串
        )
        self.sheet2 = pd.read_excel(
            FILE_PATH,
            sheet_name=SHEET2_NAME,
            dtype={'Id': 'string', 'NameDesc': 'string',
                   'EffectParam1': 'string', 'AdditionPerAdd1': 'string'},
            na_values=["", "NA", "N/A", "ARRAY(INT32)"],
            keep_default_na=False,
            skiprows=range(1, 4)
        )
        # 构建技能名称到ID的映射
        self.id_name_map = dict(zip(
            self.sheet1['NameDesc'].astype(str),
            self.sheet1['Id'].astype(str)  # 确保ID为字符串
        ))

    def get_skill_options(self):
        """获取所有技能名称列表"""
        return list(self.id_name_map.keys())

    def query_data(self, skill_name):
        """根据技能名称查询数据"""
        if not skill_name or skill_name == "输入技能名...":
            return pd.DataFrame(columns=self.columns)  # 返回空DataFrame

        # 查找匹配的技能ID (确保为字符串)
        matched_ids = [str(id_val) for name, id_val in self.id_name_map.items()
                       if skill_name in name]

        if not matched_ids:
            return pd.DataFrame(columns=self.columns)

        # 安全处理正则特殊字符
        escaped_ids = [re.escape(id_str) for id_str in matched_ids]
        pattern = '|'.join(escaped_ids)  # 创建正则模式

        # 执行模糊匹配
        mask = self.sheet2['Id'].astype(str).str.contains(
            pattern, case=False, na=False
        )
        return self.sheet2.loc[mask, self.columns]


class SkillQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("技能查询系统")
        self.root.geometry("800x600")
        self.data_manager = DataManager(COLUMNS)

        # 配置网格布局权重
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # 创建UI组件
        self.create_widgets()

        # 绑定自定义事件
        self.root.bind("<<AutocompleteSelected>>", self.update_results)

    def create_widgets(self):
        """创建所有界面组件"""
        # 输入区域框架
        input_frame = ttk.Frame(self.root)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(0, weight=1)  # 输入框列可扩展

        # 自动完成输入框
        self.autocomplete = AutocompleteEntry(
            input_frame,
            self.data_manager.get_skill_options(),
            row=0,
            column=0
        )

        # 添加ID筛选复选框
        self.id_filter_var = tk.BooleanVar(value=False)
        self.id_filter_check = ttk.Checkbutton(
            input_frame,
            text="仅显示ID以10结尾的技能",
            variable=self.id_filter_var,
            command=self.update_results  # 勾选时自动更新结果
        )
        self.id_filter_check.grid(row=0, column=1, padx=10, sticky="w")

        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="查询结果")
        result_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

        # 创建Treeview表格
        self.result_tree = ttk.Treeview(
            result_frame,
            columns=COLUMNS,
            show="headings",
            height=15  # 固定高度
        )

        # 设置列标题
        for col in COLUMNS:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=100, anchor="center")

        # 添加滚动条
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # 布局Treeview和滚动条
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def update_results(self, event=None):
        """更新查询结果"""
        # 清空现有数据
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 获取查询结果
        skill_name = self.autocomplete.entry.get()
        result_df = self.data_manager.query_data(skill_name)

        # 应用ID结尾筛选
        if not result_df.empty and self.id_filter_var.get():
            result_df = result_df[result_df['Id'].astype(str).str.endswith('0')]

        # 显示结果
        if not result_df.empty:
            for _, row in result_df.iterrows():
                self.result_tree.insert("", "end", values=list(row))
        else:
            self.result_tree.insert("", "end", values=("无匹配数据", "", "", ""))


if __name__ == "__main__":
    root = tk.Tk()
    app = SkillQueryApp(root)
    root.mainloop()