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
    # 按揭贷款数据分析
    # 按揭贷款数据导入
    def DKAnalysis_AJDK_import_file(self):
        AJDK_import_file(self)

    # 按揭贷款数据分析
    def DKAnalysis_AJDK_analyse_excel(self,Y_date):
        AJDK_analyse_excel(self,Y_date)

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
    def DKAnalysis_LSKH_analysis(self,Y_date):
        LSKH_analysis(self,Y_date)

    # 导出结果
    def DKAnalysis_LSKH_export(self):
        LSKH_export(self)
    def __init__(self, master):
        self.master = master
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
