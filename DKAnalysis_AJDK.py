import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import datetime


# 是否成功标识，0正在处理，1处理成功，-1处理失败


# 导入数据源文件
def AJDK_import_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xls"), ("All files", "*.*")))

    self.expth = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + self.expth + "  请点击步骤2")


# 处理分析数据
def AJDK_analyse_excel(self, Y_date):
    if self.expth == '':
        messagebox.showinfo("Message", "未选中文件，请点击步骤1！")
        return None
    else:
        data = pd.read_excel(self.expth, sheet_name=0, usecols='A:H, L, O:P, S', header=0, dtype=str)
        # 贷款形态为正常
        data = data[data["五级分类状态"] == "1-正常"]
        # 剔除无用数据
        data.drop([len(data) - 1], inplace=True)
        # 贷款期限大于5年的，判断出是按揭贷款
        data = data[
            (pd.to_datetime(data["到期日期"]) - pd.to_datetime(data["贷款日期"])) > datetime.timedelta(days=365 * 5)]
        # 获取现在时间
        now = datetime.datetime.now()

        if Y_date == '一年及以上':
            data = data[now - pd.to_datetime(data["贷款日期"]) >= datetime.timedelta(days=365)]
        elif Y_date == '二年及以上':
            data = data[now - pd.to_datetime(data["贷款日期"]) >= datetime.timedelta(days=365 * 2)]
        elif Y_date == '三年及以上':
            data = data[now - pd.to_datetime(data["贷款日期"]) >= datetime.timedelta(days=365 * 3)]
        elif Y_date == '四年及以上':
            data = data[now - pd.to_datetime(data["贷款日期"]) >= datetime.timedelta(days=365 * 4)]
        elif Y_date == '五年及以上':
            data = data[now - pd.to_datetime(data["贷款日期"]) >= datetime.timedelta(days=365 * 5)]
        # 计算可用贷款额度
        data.insert(loc=6, column="可用贷款余额",
                    value=data['贷款金额'].str.replace(',', '').astype(float) - data['本金余额'].str.replace(',', '').astype(float))

        self.AJDK_dataResult = data

        if self.AJDK_dataResult is not None:
            messagebox.showinfo("Message", "数据处理完毕！点击步骤3，可以导出数据！")
        else:
            messagebox.showinfo("Message", "处理失败！请重试或咨询技术人员。")


# 数据导出,导出格式为Excel
def AJDK_export_file(self):
    export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if export_path == '':
        messagebox.showinfo("Message", "导出失败，请重新导出！")
        return None
    else:
        writer = pd.ExcelWriter(export_path, engine='openpyxl')
        # 导出明细表格
        df1 = pd.DataFrame(self.AJDK_dataResult)
        df1.to_excel(writer, sheet_name='明细', index=False)
        # 计算客户户数
        count = self.AJDK_dataResult["机构名称"].value_counts().to_frame().reset_index()
        # 导出汇总表格
        df2 = pd.DataFrame({'机构名称': count["机构名称"], '户数': count["count"]})
        print(df2)
        df2.to_excel(writer, sheet_name='汇总', index=False)
        writer.close()
        messagebox.showinfo("Message", "数据导出成功，位置在：" + export_path)
