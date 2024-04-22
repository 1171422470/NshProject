import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
from DKAnalysis_AJDK import AJDK_import_file, AJDK_analyse_excel, AJDK_export_file
from DKAnalysis_ZJGH import (ZJGH_import_dk202312_file, ZJGH_import_ck202312_file, ZJGH_import_ck_now_file,
                             ZJGH_dk_ck_file_analyse, ZJGH_export)
from DKAnalysis_LSKH import LSKH_wjq_import_file, LSKH_yjq_import_file, LSKH_analysis, LSKH_export

from function import selectExcel, dealData, excelOutput


# 是否成功标识，0正在处理，1处理成功，-1处理失败
class DKAnalysis:

    # 主界面
    def main_window(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()
        # 按揭贷款分析按揭
        AJDK_analyse_button = tk.Button(self.master, text="1、正常类按揭贷款数据分析", command=self.DKAnalysis_AJDK,
                                        relief=tk.RAISED,
                                        bd=1,
                                        bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                        highlightthickness=2)
        AJDK_analyse_button.pack(padx=10, pady=10, anchor='nw')
        # 贷款客户资金归行情况分析
        ZJGH_analyse_button = tk.Button(self.master, text="2、2023年贷款客户资金归行情况分析",
                                        command=self.DKAnalysis_ZJGH,
                                        relief=tk.RAISED,
                                        bd=1,
                                        bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                        highlightthickness=2)
        ZJGH_analyse_button.pack(padx=10, pady=10, anchor='nw')
        # 近5年有贷款往来，现在已经结清了的客户
        LSKH_analyse_button = tk.Button(self.master, text="3、通过近五年的贷款数据，分析流失客户",
                                        command=self.DKAnalysis_LSKH,
                                        relief=tk.RAISED,
                                        bd=1,
                                        bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                        highlightthickness=2)
        LSKH_analyse_button.pack(padx=10, pady=10, anchor='nw')

    # 按揭贷款分析界面
    def DKAnalysis_AJDK(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()
        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)  # 用frame标签
        # 设置导入文件按钮键
        import_button = tk.Button(frame, text="1、选择未结清贷款明细文件（数据来源：信贷查询系统）",
                                  command=self.DKAnalysis_AJDK_import_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        import_button.pack(side=tk.LEFT, padx=10, pady=10, anchor='nw')

        # 创建一个文本标签，并设置其显示的文本
        label = tk.Label(frame, text="还款年限")
        # 创建下拉框
        combo_box = ttk.Combobox(frame, width=10, exportselection=False)
        options = ['一年及以上', '二年及以上', '三年及以上', '四年及以上', '五年及以上']  # 下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.AJDK_combox = combo_box
        label.pack(side=tk.LEFT, padx=5)
        combo_box.pack(side=tk.LEFT, padx=5)
        frame.pack(padx=1, pady=1, anchor='nw')

        # 设置分析数据按钮键
        analyse_button = tk.Button(self.master, text="2、分析数据", command=self.DKAnalysis_AJDK_analyse_excel,
                                   relief=tk.RAISED,
                                   bd=1,
                                   bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                   highlightthickness=2)
        analyse_button.pack(padx=10, pady=10, anchor='nw')
        # 设置导出数据按钮键
        export_button = tk.Button(self.master, text="3、导出表格", command=self.DKAnalysis_AJDK_export_file,
                                  relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        export_button.pack(padx=10, pady=10, anchor='nw')
        # 返回主界面按钮键
        return_button = tk.Button(self.master, text="4、返回主界面", command=self.main_window, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        return_button.pack(padx=10, pady=10, anchor='nw')

    # 贷款客户资金归行情况分析
    def DKAnalysis_ZJGH(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()

        # 202312贷款余额导入文件
        import_dk202312_button = tk.Button(self.master, text="1、导入2023年底贷款余额表(数据来源:ODS)",
                                           command=self.DKAnalysis_ZJGH_import_dk202312_file,
                                           relief=tk.RAISED,
                                           bd=1,
                                           bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                           highlightthickness=2)
        import_dk202312_button.pack(padx=10, pady=10, anchor='nw')
        # 202312存款余额导入文件
        import_ck202312_button = tk.Button(self.master, text="2、导入2023年底存款余额表(数据来源:ODS)",
                                           command=self.DKAnalysis_ZJGH_import_ck202312_file,
                                           relief=tk.RAISED,
                                           bd=1,
                                           bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                           highlightthickness=2)
        import_ck202312_button.pack(padx=10, pady=10, anchor='nw')

        # 现在贷款余额导入文件
        # import_dk_now_button = tk.Button(self.master, text="3、导入当前贷款余额表",
        #                                  command=self.import_dk_now_file,
        #                                  relief=tk.RAISED,
        #                                  bd=1,
        #                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
        #                                  highlightthickness=2)
        # import_dk_now_button.pack(padx=10, pady=10, anchor='nw')
        # 现在存款余额导入文件
        import_ck_now_button = tk.Button(self.master, text="3、导入当前存款余额表",
                                         command=self.DKAnalysis_ZJGH_import_ck_now_file,
                                         relief=tk.RAISED,
                                         bd=1,
                                         bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                         highlightthickness=2)
        import_ck_now_button.pack(padx=10, pady=10, anchor='nw')

        # 2023年贷款客户资金归行情况
        output_ck_dk_2023_button = tk.Button(self.master, text="4、分析数据",
                                             command=self.DKAnalysis_ZJGH_dk_ck_file_analyse,
                                             relief=tk.RAISED,
                                             bd=1,
                                             bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                             highlightthickness=2)
        output_ck_dk_2023_button.pack(padx=10, pady=10, anchor='nw')

        # 导出数据
        output_ck_dk_2023_button = tk.Button(self.master, text="5、导出数据",
                                             command=self.DKAnalysis_ZJGH_export,
                                             relief=tk.RAISED,
                                             bd=1,
                                             bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                             highlightthickness=2)
        output_ck_dk_2023_button.pack(padx=10, pady=10, anchor='nw')

        # 返回主界面按钮键
        return_button = tk.Button(self.master, text="6、返回主界面", command=self.main_window, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        return_button.pack(padx=10, pady=10, anchor='nw')

    # 近5年有贷款往来，现在已经结清了的客户
    def DKAnalysis_LSKH(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()

        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)  # 用frame标签
        # 设置导入文件按钮键
        import_button = tk.Button(frame, text="1、选择已结清贷款明细文件（数据来源：信贷查询系统）",
                                  command=self.DKAnalysis_LSKH_yjq_import_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        import_button.pack(side=tk.LEFT, padx=10, pady=10, anchor='nw')

        # 创建一个文本标签，并设置其显示的文本
        label = tk.Label(frame, text="未发生贷款业务时长")
        # 创建下拉框
        combo_box = ttk.Combobox(frame, width=10, exportselection=False)
        options = ['一年及以上', '二年及以上', '三年及以上', '四年及以上', '五年']  # 下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.LSKH_combox = combo_box
        label.pack(side=tk.LEFT, padx=5)
        combo_box.pack(side=tk.LEFT, padx=5)
        frame.pack(padx=1, pady=1, anchor='nw')
        # 设置数据按钮键
        import_button_2 = tk.Button(self.master, text="2、选择未结清贷款明细文件（数据来源：信贷查询系统）",
                                    command=self.DKAnalysis_LSKH_wjq_import_file,
                                    relief=tk.RAISED,
                                    bd=1,
                                    bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                    highlightthickness=2)
        import_button_2.pack(padx=10, pady=10, anchor='nw')

        # 分析数据按钮键
        analysis_button = tk.Button(self.master, text="3、分析数据",
                                    command=self.DKAnalysis_LSKH_analysis,
                                    relief=tk.RAISED,
                                    bd=1,
                                    bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                    highlightthickness=2)
        analysis_button.pack(padx=10, pady=10, anchor='nw')

        # 导出数据按钮键
        analysis_button = tk.Button(self.master, text="4、导出数据",
                                    command=self.DKAnalysis_LSKH_export,
                                    relief=tk.RAISED,
                                    bd=1,
                                    bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                    highlightthickness=2)
        analysis_button.pack(padx=10, pady=10, anchor='nw')

        # 返回主界面按钮键
        return_button = tk.Button(self.master, text="5、返回主界面", command=self.main_window, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        return_button.pack(padx=10, pady=10, anchor='nw')

    # 按揭贷款数据分析
    # 按揭贷款数据导入
    def DKAnalysis_AJDK_import_file(self):
        AJDK_import_file(self)

    # 按揭贷款数据分析
    def DKAnalysis_AJDK_analyse_excel(self):
        AJDK_analyse_excel(self)

    # 按揭贷款分析结果导出
    def DKAnalysis_AJDK_export_file(self):
        AJDK_export_file(self)

    # 贷款客户资金归行情况分析
    # 202312贷款余额数据导入
    def DKAnalysis_ZJGH_import_dk202312_file(self):
        ZJGH_import_dk202312_file(self)

    # 202312存款余额数据导入
    def DKAnalysis_ZJGH_import_ck202312_file(self):
        ZJGH_import_ck202312_file(self)

    # 当前存款余额数据导入
    def DKAnalysis_ZJGH_import_ck_now_file(self):
        ZJGH_import_ck_now_file(self)

    # 数据分析
    def DKAnalysis_ZJGH_dk_ck_file_analyse(self):
        ZJGH_dk_ck_file_analyse(self)

    # 结果导出
    def DKAnalysis_ZJGH_export(self):
        ZJGH_export(self)

    # 近5年有贷款往来，现在已经结清了的客户
    # 导入未结清数据源文件
    def DKAnalysis_LSKH_wjq_import_file(self):
        LSKH_wjq_import_file(self)

    # 导入已结清数据源文件
    def DKAnalysis_LSKH_yjq_import_file(self):
        LSKH_yjq_import_file(self)

    # 导入已结清数据源文件
    def DKAnalysis_LSKH_analysis(self):
        LSKH_analysis(self)

    # 导出结果
    def DKAnalysis_LSKH_export(self):
        LSKH_export(self)
    def __init__(self, master):
        self.master = master
        self.status_lable = tk.Label(self.master, text="")
        self.master.title("辰溪农商银行数据分析小程序")
        self.master.geometry("800x500")  # 固定窗口大小
        # 按揭贷款数据分析
        self.expth = ''  # 初始化按揭贷款分析导入文件路径字段
        self.AJDK_dataResult = 0
        # 资金归行情况分析
        self.dk202312 = ''  # 初始化202312贷款余额导入文件路径字段
        self.ck202312 = ''  # 初始化202312存款余额导入文件路径字段
        self.dk_now = ''  # 初始化当前贷款余额导入文件路径字段
        self.ck_now = ''  # 初始化当前存款余额导入文件路径字段
        self.dkht = ''  # 初始化贷款合同导入文件路径字段
        self.ck_dk_202312_result_mx = 0
        self.ck_dk_202312_result_hz = 0
        # 近5年有贷款往来，现在已经结清了的客户
        self.LSKH_wjq_addr = ''  # 初始化未结清贷款明细路径
        self.LSKH_yjq_addr = ''  # 初始化已结清贷款明细路径
        self.LSKH_result_data = 0   # 初始化生产的结果
        # 启动主界面
        self.main_window()
