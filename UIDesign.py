import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk

from function import selectExcel,dealData,excelOutput

# 是否成功标识，0正在处理，1处理成功，-1处理失败
class ExcelGUI:

    # 主界面
    def main_window(self):
        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)  # 用frame标签
        # 状态提示
        self.status_lable = tk.Label(self.master, textvariable=self.str)
        self.status_lable.pack(side="bottom")
        # 设置导入文件按钮键
        import_button = tk.Button(frame, text="1、选择文件", command=self.import_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        import_button.pack(side=tk.LEFT,padx=10, pady=10,anchor='nw')


        # 创建一个文本标签，并设置其显示的文本
        label = tk.Label(frame, text="选择年限")
        # 创建下拉框
        combo_box = ttk.Combobox(frame, width=10, exportselection=False)
        options = ['一年', '两年', '三年','四年','五年']  # 下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.combox = combo_box
        label.pack(side=tk.LEFT,padx=5)
        combo_box.pack(side=tk.LEFT, padx=5)
        frame.pack(padx=1, pady=1, anchor='nw')

        # 设置分析数据按钮键
        analyse_button = tk.Button(self.master, text="2、分析数据", command=self.analyse_excel, relief=tk.RAISED,
                                   bd=1,
                                   bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                   highlightthickness=2)
        analyse_button.pack(padx=10, pady=10,anchor='nw')
        # 设置导出数据按钮键
        export_button = tk.Button(self.master, text="3、导出表格", command=self.export_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        export_button.pack(padx=10, pady=10,anchor='nw')
        # 设置退出主界面按钮键
        # exit_button = tk.Button(self.master, text="4、退出程序", command=exit, relief=tk.RAISED, bd=1,
        #                         bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
        #                         highlightthickness=2)
        # exit_button.pack(padx=10, pady=10)


    # 数据导出,导出格式为Excel
    def export_file(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        excelOutput(export_path, self.dataResult)
        self.str.set("数据导出成功，位置在：" + export_path)

    # 导入数据源文件
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        # if file_path:
        #     self.data_frame = pd.read_excel(file_path)
        #     self.data_text.delete(1.0, tk.END)
        #     self.data_text.insert(tk.END, self.data_frame.to_string())
        self.expth = file_path
        self.str.set("选中：" + self.expth+"   点击步骤2，请耐心等待。。。")

    # 处理分析数据
    def analyse_excel(self):
        Y_date = self.combox.get() #获取下拉框值
        self.dataResult = dealData(self.expth,Y_date)
        if self.dataResult is not None:
            self.str.set("恭喜!处理成功。点击步骤3，导出表格。")
        else:
            self.str.set("处理失败！请重试或咨询技术人员。")

    def __init__(self, master):
        self.master = master
        self.status_lable = tk.Label(self.master, text="")
        self.master.title("按揭贷款数据分析")
        self.master.geometry("800x500")  # 固定窗口大小
        self.expth = ''  # 初始化路径字段
        self.str = tk.StringVar()
        self.dataResult = 0
        # 启动主界面
        self.main_window()
