# 需求4：近5年在我行有业务往来(有贷款但是已经结清了)，贷款已结清表和贷款未结算表(流失客户)
# 字段名：客户名称，注册号码，贷款金额，贷款日期，到期日期，贷款到期前利率，五级分类状态，主客户经理，第一责任人，行政村组，联系电话，本金结清日期
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import datetime


# 导入未结清数据源文件
def LSKH_wjq_import_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))

    self.LSKH_wjq_addr = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        data = pd.read_excel(file_path, sheet_name=0, header=0, dtype=str).head(1)
        if data.empty:
            messagebox.showinfo("Message", "文件有误，请检查文件！")
            return False
        else:
            messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤3")


# 导入已结清数据源文件
def LSKH_yjq_import_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))

    self.LSKH_yjq_addr = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        data = pd.read_excel(file_path, sheet_name=0, header=0, dtype=str).head(1)
        if data.empty:
            messagebox.showinfo("Message", "文件有误，请检查文件！")
            return False
        else:
            messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤2")


# 分析数据
def LSKH_analysis(self):
    # 清空页面
    if self.LSKH_wjq_addr == '':
        messagebox.showinfo("Message", "未选中文件，请点击步骤2！")
        return None
    elif self.LSKH_yjq_addr == '':
        messagebox.showinfo("Message", "未选中文件，请点击步骤1！")
        return None
    else:
        wjq_data = pd.read_excel(self.LSKH_wjq_addr, sheet_name=0, usecols='A, C:E, G:H, J, L, O:Q, S, V',
                                 header=0, dtype=str)
        yjq_data = pd.read_excel(self.LSKH_yjq_addr, sheet_name=0, usecols='A, C:E, G:H, J, L, O:Q, S, V',
                                 header=0, dtype=str)
        # 获取当前时间
        now = datetime.datetime.now()
        # 删除为空的数据
        wjq_data.drop([len(wjq_data) - 1], inplace=True)
        yjq_data.drop([len(yjq_data) - 1], inplace=True)
        #
        # 已结清贷款还款日期在5年以内的
        yjq_data = yjq_data[(pd.to_datetime(yjq_data["本金结清日期/最近还款日期"]).dt.year >= now.year - 5) & (
                    pd.to_datetime(yjq_data["本金结清日期/最近还款日期"]).dt.year < now.year)]
        # 筛选不良态的客户信息
        bad_customers_data = yjq_data[
            (yjq_data["五级分类状态"] == '3-次级') | (yjq_data["五级分类状态"] == '4-可疑') | (
                        yjq_data["五级分类状态"] == '5-损失')]
        # print(bad_customers_data)
        # 删除不良客户和注册号码为空的
        yjq_data = yjq_data[~yjq_data["注册号码"].isin(bad_customers_data["注册号码"])]
        # 按证件号码排序
        yjq_data.sort_values(by=['注册号码', '本金结清日期/最近还款日期'], inplace=True, ascending=True)
        # 删除在未结清贷款中的客户
        result_date = yjq_data[~yjq_data["注册号码"].isin(wjq_data["注册号码"])]

        Y_date = self.LSKH_combox.get()  # 获取下拉框值
        if Y_date == '二年及以上':
            result_date = result_date[pd.to_datetime(result_date["本金结清日期/最近还款日期"]).dt.year <= now.year - 2]
        elif Y_date == '三年及以上':
            result_date = result_date[pd.to_datetime(result_date["本金结清日期/最近还款日期"]).dt.year <= now.year - 3]
        elif Y_date == '四年及以上':
            result_date = result_date[pd.to_datetime(result_date["本金结清日期/最近还款日期"]).dt.year <= now.year - 4]
        elif Y_date == '五年':
            result_date = result_date[pd.to_datetime(result_date["本金结清日期/最近还款日期"]).dt.year <= now.year - 5]

        self.LSKH_result_data = result_date

        if result_date is not None:
            messagebox.showinfo("Message", "数据处理完毕！点击步骤4，可以导出数据！")
        else:
            messagebox.showinfo("Message", "处理失败！请重试或咨询技术人员。")


# 数据导出,导出格式为Excel
def LSKH_export(self):
    export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if export_path == '':
        messagebox.showinfo("Message", "导出失败，请重新导出！")
        return None
    else:
        writer = pd.ExcelWriter(export_path, engine='openpyxl')
        # 导出明细表格
        df1 = pd.DataFrame(self.LSKH_result_data)
        df1.to_excel(writer, sheet_name='明细', index=False)
        # 计算客户户数
        count = self.LSKH_result_data["机构名称"].value_counts().to_frame().reset_index()
        # 导出汇总表格
        df2 = pd.DataFrame({'机构名称': count["机构名称"], '笔数': count["count"]})
        print(df2)
        df2.to_excel(writer, sheet_name='汇总', index=False)
        writer.close()
        messagebox.showinfo("Message", "数据导出成功，位置在：" + export_path)
