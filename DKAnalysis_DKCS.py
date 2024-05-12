import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import datetime
import string


# 1、余额占比=区间贷款余额/各项贷款余额
# 2、收息占比=区间贷款实收利息/实收利息总额
# 3、不良贷款余额占比=区间不良贷款余额/区间贷款余额
# 4、不良率=区间贷款不良贷款余额/各项贷款余额
# 5、加权利率=∑（区间每笔贷款余额*对应贷款利率）/区间每笔贷款余额合计
# 6、收息率=当期实收利息/当期应收利息
# 各项数据来源：
# 1、当前实收利息：ODS贷款回收登记薄
# 2、收息率=当期实收利息/（当期实收利息+当期期末应收未收利息）


# 导入贷款余额表
def DKCS_DKYEB_import_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))

    self.DKCS_dkyeb_addr = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤2")


# 导入贷款回收登记簿
def DKCS_DKHSDJB_import_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))

    self.DKCS_dkhsdjb_addr = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤3")


# 处理分析数据
def DKCS_analyse_excel(self):
    if self.DKCS_dkyeb_addr == '':
        messagebox.showinfo("Message", "未选中文件，请点击步骤1！")
        return None
    elif self.DKCS_dkhsdjb_addr == '':
        messagebox.showinfo("Message", "未选中文件，请点击步骤2！")
        return None
    else:
        # 打开贷款余额表
        dkyeb_data = pd.read_excel(self.DKCS_dkyeb_addr, sheet_name=0, header=4, dtype=str)
        # 删除不需要的列
        dkyeb_data.drop([len(dkyeb_data) - 1], inplace=True)
        # 打开贷款回收登记薄
        dkhsdjb_data = pd.read_excel(self.DKCS_dkhsdjb_addr, sheet_name=0, header=4, dtype=str)
        # 删除不需要的列
        dkhsdjb_data.drop([len(dkhsdjb_data) - 1], inplace=True)
        # 数据类型转换
        dkyeb_data['贷款余额(元)'] = dkyeb_data['贷款余额(元)'].apply(lambda x: float(x.replace(',', '')))
        dkyeb_data['表内应收利息（元）'] = dkyeb_data['表内应收利息（元）'].apply(lambda x: float(x.replace(',', '')))
        dkyeb_data['利率(%)'] = dkyeb_data['利率(%)'].apply(lambda x: float(x.replace(',', '')))
        dkhsdjb_data['收回利息(元)'] = dkhsdjb_data['收回利息(元)'].apply(lambda x: float(x.replace(',', '')))
        # 初始化结果
        self.LSKH_result_data = pd.DataFrame(
            {'贷款金额': ['50万元（含）以下', '50万元-500万元（含）以下', '500万元以上', '贷款余额合计'],
             '贷款余额（万元）': None, '余额占比': None, '收息占比': None, '不良贷款余额占比': None, '不良率': None,
             '贷款加权利率': None, '收息率': None
             })
        # 聚合
        dkyeb_data_groupID = dkyeb_data.groupby('证件号码', as_index=False)['贷款余额(元)'].sum()

        # 50万以下余额求和
        ye_0to50 = dkyeb_data[
            dkyeb_data['证件号码'].isin(dkyeb_data_groupID[dkyeb_data_groupID['贷款余额(元)'] <= 500000]['证件号码'])]
        # print(sum(ye_0to50['贷款余额(元)']))
        self.LSKH_result_data.loc[0, '贷款余额（万元）'] = sum(ye_0to50['贷款余额(元)']) / 10000
        # 50万元-500万元（含）以下余额求和
        ye_50to500 = dkyeb_data[
            dkyeb_data['证件号码'].isin(dkyeb_data_groupID[dkyeb_data_groupID['贷款余额(元)'] > 500000]['证件号码']) &
            dkyeb_data['证件号码'].isin(dkyeb_data_groupID[dkyeb_data_groupID['贷款余额(元)'] <= 5000000]['证件号码'])]
        self.LSKH_result_data.loc[1, '贷款余额（万元）'] = sum(ye_50to500['贷款余额(元)']) / 10000
        # 500万元以上余额求和
        ye_500 = dkyeb_data[
            dkyeb_data['证件号码'].isin(dkyeb_data_groupID[dkyeb_data_groupID['贷款余额(元)'] > 5000000]['证件号码'])]
        self.LSKH_result_data.loc[2, '贷款余额（万元）'] = sum(ye_500['贷款余额(元)']) / 10000
        # 贷款余额总和
        self.LSKH_result_data.loc[3, '贷款余额（万元）'] = sum(dkyeb_data['贷款余额(元)']) / 10000

        # 余额占比计算
        for i in range(0, 3):
            self.LSKH_result_data.loc[i, '余额占比'] = (self.LSKH_result_data.loc[i, '贷款余额（万元）']
                                                        / self.LSKH_result_data.loc[3, '贷款余额（万元）'])

        # 收息占比计算
        # 0-50万
        self.LSKH_result_data.loc[0, '收息占比'] = sum(
            dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_0to50['贷款账号'])]['收回利息(元)']) / sum(
            dkhsdjb_data['收回利息(元)'])
        # print(sum(dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_0to50['贷款账号'])]['收回利息(元)']))
        # print(sum(dkhsdjb_data['收回利息(元)']))
        # 50-500万
        self.LSKH_result_data.loc[1, '收息占比'] = sum(
            dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_50to500['贷款账号'])]['收回利息(元)']) / sum(
            dkhsdjb_data['收回利息(元)'])
        # 500万以上
        self.LSKH_result_data.loc[2, '收息占比'] = sum(
            dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_500['贷款账号'])]['收回利息(元)']) / sum(
            dkhsdjb_data['收回利息(元)'])

        # 不良贷款余额占比
        # 0-50万
        if sum(ye_0to50['贷款余额(元)']) == 0:
            self.LSKH_result_data.loc[0, '不良贷款余额占比'] = 0
        else:
            self.LSKH_result_data.loc[0, '不良贷款余额占比'] = sum(
                ye_0to50[ye_0to50["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)']) / sum(
                ye_0to50['贷款余额(元)'])
        # 50-500万
        if sum(ye_50to500['贷款余额(元)']) == 0:
            self.LSKH_result_data.loc[1, '不良贷款余额占比'] = 0
        else:
            self.LSKH_result_data.loc[1, '不良贷款余额占比'] = sum(
                ye_50to500[ye_50to500["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)']) / sum(
                ye_50to500['贷款余额(元)'])
        # 500万以上
        if sum(ye_500['贷款余额(元)']) == 0:
            self.LSKH_result_data.loc[2, '不良贷款余额占比'] = 0
        else:
            self.LSKH_result_data.loc[2, '不良贷款余额占比'] = sum(
                ye_500[ye_500["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)']) / sum(ye_500['贷款余额(元)'])

        # 不良率
        # 0-50万
        bl1 = sum(
            ye_0to50[ye_0to50["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)'])
        if bl1 == 0:
            self.LSKH_result_data.loc[0, '不良率'] = 0
        else:
            self.LSKH_result_data.loc[0, '不良率'] = bl1 / (self.LSKH_result_data.iloc[3]['贷款余额（万元）'] * 10000)
        # 50-500万
        bl2 = sum(
            ye_50to500[ye_50to500["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)'])
        if bl2 == 0:
            self.LSKH_result_data.loc[1, '不良率'] = 0
        else:
            self.LSKH_result_data.loc[1, '不良率'] = bl2 / (self.LSKH_result_data.iloc[3]['贷款余额（万元）'] * 10000)
        # 500万以上
        bl3 = sum(
            ye_500[ye_500["贷款形态"].isin(['次级', '可疑', '损失'])]['贷款余额(元)'])
        if bl3 == 0:
            self.LSKH_result_data.loc[2, '不良率'] = 0
        else:
            self.LSKH_result_data.loc[2, '不良率'] = bl3 / (self.LSKH_result_data.iloc[3]['贷款余额（万元）'] * 10000)

        # 收息率
        # 0-50万
        zz1 = sum(dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_0to50['贷款账号'])]['收回利息(元)'])
        # print(zz1)
        if zz1 == 0:
            self.LSKH_result_data.loc[0, '收息率'] = 0
        else:
            self.LSKH_result_data.loc[0, '收息率'] = zz1 / (sum(ye_0to50['表内应收利息（元）']) + zz1)
        # 50-500万
        zz2 = sum(dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_50to500['贷款账号'])]['收回利息(元)'])
        if zz2 == 0:
            self.LSKH_result_data.loc[1, '收息率'] = 0
        else:
            self.LSKH_result_data.loc[1, '收息率'] = zz2 / (sum(ye_50to500['表内应收利息（元）']) + zz2)

        # 500万以上
        zz3 = sum(dkhsdjb_data[dkhsdjb_data["贷款帐号"].isin(ye_500['贷款账号'])]['收回利息(元)'])
        if zz3 == 0:
            self.LSKH_result_data.loc[2, '收息率'] = 0
        else:
            self.LSKH_result_data.loc[2, '收息率'] = zz3 / (sum(ye_500['表内应收利息（元）']) + zz3)

        # 贷款加权利率
        sum0to50 = 0
        sum50to500 = 0
        sum500 = 0
        for i in range(len(ye_0to50)):
            sum0to50 = ye_0to50.iloc[i]['贷款余额(元)'] * ye_0to50.iloc[i]['利率(%)'] + sum0to50
        # print(sum0to50)
        for i in range(0, len(ye_50to500)):
            sum50to500 = ye_50to500.iloc[i]['贷款余额(元)'] * ye_50to500.iloc[i]['利率(%)'] + sum50to500
        for i in range(0, len(ye_500)):
            sum500 = ye_500.iloc[i]['贷款余额(元)'] * ye_500.iloc[i]['利率(%)'] + sum500
        # 0-50万
        if sum0to50 == 0:
            self.LSKH_result_data.loc[0, '贷款加权利率'] = 0
        else:
            self.LSKH_result_data.loc[0, '贷款加权利率'] = sum0to50 / sum(ye_0to50['贷款余额(元)'])
        # 50-500万
        if sum50to500 == 0:
            self.LSKH_result_data.loc[1, '贷款加权利率'] = 0
        else:
            self.LSKH_result_data.loc[1, '贷款加权利率'] = sum50to500 / sum(ye_50to500['贷款余额(元)'])
        # 500万以上
        if sum500 == 0:
            self.LSKH_result_data.loc[2, '贷款加权利率'] = 0
        else:
            self.LSKH_result_data.loc[2, '贷款加权利率'] = sum500 / sum(ye_500['贷款余额(元)'])

        # print(self.LSKH_result_data['贷款加权利率'])
        # 保留两位小数
        for i in range(0, 4):
            if self.LSKH_result_data.iloc[i]['贷款余额（万元）'] is not None:
                self.LSKH_result_data.loc[i, '贷款余额（万元）'] = '%.2f' % (
                    self.LSKH_result_data.iloc[i]['贷款余额（万元）'])
            if self.LSKH_result_data.iloc[i]['余额占比'] is not None:
                self.LSKH_result_data.loc[i, '余额占比'] = '%.2f' % (
                            self.LSKH_result_data.iloc[i]['余额占比'] * 100) + '%'
            if self.LSKH_result_data.iloc[i]['收息占比'] is not None:
                self.LSKH_result_data.loc[i, '收息占比'] = '%.2f' % (
                            self.LSKH_result_data.iloc[i]['收息占比'] * 100) + '%'
            if self.LSKH_result_data.iloc[i]['不良贷款余额占比'] is not None:
                self.LSKH_result_data.loc[i, '不良贷款余额占比'] = '%.2f' % (
                        self.LSKH_result_data.iloc[i]['不良贷款余额占比'] * 100) + '%'
            if self.LSKH_result_data.iloc[i]['不良率'] is not None:
                self.LSKH_result_data.loc[i, '不良率'] = '%.2f' % (self.LSKH_result_data.iloc[i]['不良率'] * 100) + '%'
            if self.LSKH_result_data.iloc[i]['贷款加权利率'] is not None:
                self.LSKH_result_data.loc[i, '贷款加权利率'] = '%.2f' % (
                    self.LSKH_result_data.iloc[i]['贷款加权利率']) + '%'
            if self.LSKH_result_data.iloc[i]['收息率'] is not None:
                self.LSKH_result_data.loc[i, '收息率'] = '%.2f' % (self.LSKH_result_data.iloc[i]['收息率'] * 100) + '%'

        # 补充备注
        self.LSKH_result_data.loc[7, '贷款金额'] = '注：'
        self.LSKH_result_data.loc[8, '贷款金额'] = '1、余额占比=区间贷款余额/各项贷款余额'
        self.LSKH_result_data.loc[9, '贷款金额'] = '2、收息占比=区间贷款实收利息/实收利息总额'
        self.LSKH_result_data.loc[10, '贷款金额'] = '3、不良贷款余额占比=区间不良贷款余额/区间贷款余额'
        self.LSKH_result_data.loc[11, '贷款金额'] = '4、不良率=区间贷款不良贷款余额/各项贷款余额'
        self.LSKH_result_data.loc[12, '贷款金额'] = '5、加权利率=∑（区间每笔贷款余额*对应贷款利率）/区间每笔贷款余额合计'
        self.LSKH_result_data.loc[13, '贷款金额'] = '6、收息率=当期实收利息/当期应收利息'
        self.LSKH_result_data.loc[14, '贷款金额'] = '各项数据来源：'
        self.LSKH_result_data.loc[15, '贷款金额'] = '1、当前实收利息：ODS贷款回收登记薄'
        self.LSKH_result_data.loc[16, '贷款金额'] = '2、收息率=当期实收利息/（当期实收利息+当期期末应收未收利息）'

        if self.LSKH_result_data is not None:
            messagebox.showinfo("Message", "数据处理完毕！点击步骤3，可以导出数据！")
        else:
            messagebox.showinfo("Message", "处理失败！请重试或咨询技术人员。")


# 数据导出,导出格式为Excel
def DKCS_export_file(self):
    export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if export_path == '':
        messagebox.showinfo("Message", "导出失败，请重新导出！")
        return None
    else:
        writer = pd.ExcelWriter(export_path, engine='openpyxl')
        # 导出表格
        self.LSKH_result_data.to_excel(writer, sheet_name='明细', index=False)
        writer.close()
        messagebox.showinfo("Message", "数据导出成功，位置在：" + export_path)
