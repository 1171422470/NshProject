import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import xlrd

# 贷款存款余额数据分析
def ZJGH_dk_ck_file_analyse(self):
    # 获取2023年12月底贷款数据，使用B-J列
    dk202312_df = pd.read_excel(self.dk202312, sheet_name=0, usecols='B:G, J, P', header=4,
                                dtype=str, keep_default_na=False)
    # 删除不需要的列
    dk202312_df.drop([len(dk202312_df) - 1], inplace=True)
    # 删除重复行
    # dk202312_1 = dk202312_df.drop_duplicates("证件号码", keep='first', inplace=False)
    # 获取2023年12月底存款数据，使用F-Q列
    ck202312_df = pd.read_excel(self.ck202312, sheet_name=0, usecols='F, Q', header=1,
                                dtype={'证件号码': str, '年平均': object})

    # 202312底存款按证件号得到一个透视表
    ck202312_1 = pd.pivot_table(ck202312_df, index=["证件号码"], values=["年平均"], aggfunc=sum)
    # 合并2023年底存款和贷款表
    dk202312_final_df = pd.merge(dk202312_df, ck202312_1, how='left', on="证件号码").rename \
        (columns={"年平均": "2023年底存款年平均"})
    # print(dk202312_final_df.keys())
    # dk_now_df = pd.read_excel(self.dk_now, sheet_name=0, usecols='B:D', header=4, dtype={'证件号码': str},
    #                           keep_default_na=False)
    # dk_now_1 = dk_now_df.drop_duplicates("证件号码", keep='first', inplace=False)
    # 获取当前存款数据，使用F-Q列
    ck_now_df = pd.read_excel(self.ck_now, sheet_name=0, usecols='F, Q', header=1,
                              dtype={'证件号码': str, '年平均': object})
    # 当前存款按证件号得到一个透视表
    ck_now_1 = pd.pivot_table(ck_now_df, index=["证件号码"], values=["年平均"], aggfunc=sum).rename \
        (columns={"年平均": "当前存款年平均"})
    # dk_now_final_df = pd.merge(dk_now_1, ck_now_1, how='left', on="证件号码")
    # dk_now_final_df.rename(columns={"年平均": "当前存款年平均"})

    # 合并当前存款数据和2023年底计算的存贷数据
    dk_2023_now_final_df = pd.merge(dk202312_final_df, ck_now_1, how='left', on="证件号码")
    # 计算存款平均差值
    dk_2023_now_final_df.loc[:, '差值'] = dk_2023_now_final_df['当前存款年平均'] - dk_2023_now_final_df[
        '2023年底存款年平均']
    # 空格填0
    dk_2023_now_final_df = dk_2023_now_final_df.fillna(0)
    # 明细表
    self.ck_dk_202312_result_mx = dk_2023_now_final_df
    self.ck_dk_202312_result_mx.sort_values(by=['开户机构', '证件号码'], inplace=True, ascending=True)  # 按照证件号排序
    # 汇总表
    data_tmp = dk_2023_now_final_df.drop_duplicates("证件号码", keep='first', inplace=False)  # 删除重复项
    self.ck_dk_202312_result_hz = pd.pivot_table(data_tmp, index=["开户机构"],
                                                 values=["2023年底存款年平均", "当前存款年平均", "差值"],
                                                 aggfunc=sum).reset_index()
    # print(self.ck_dk_202312_result_hz.keys())
    # print(self.ck_dk_202312_result_mx.keys())
    messagebox.showinfo("Message", "分析完成，请导出数据！")


# 贷款余额数据分析导出数据
def ZJGH_export(self):
    export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    writer = pd.ExcelWriter(export_path, engine='openpyxl')
    # 替换机构号为行名
    self.ck_dk_202312_result_hz = self.ck_dk_202312_result_hz.replace(
        {"开户机构": {"45500": "营业部", "45501": "辰阳支行", "45502": "沅江路支行", "45505": "城郊支行",
                      "45507": "田湾支行", "45508": "孝坪支行", "45509": "修溪支行", "45510": "伍家湾支行",
                      "45511": "谭家场支行", "45512": "潭湾支行", "45513": "桥头支行", "45514": "锦滨支行",
                      "45515": "安坪支行", "45516": "大水田支行", "45517": "龙泉岩支行", "45519": "火马冲支行",
                      "45520": "寺前支行", "45521": "小龙门支行", "45522": "锄头坪支行", "45523": "黄溪口支行",
                      "45524": "龙头庵支行", "45525": "后塘支行", "45526": "仙人湾支行", }})
    self.ck_dk_202312_result_mx = self.ck_dk_202312_result_mx.replace(
        {"开户机构": {"45500": "营业部", "45501": "辰阳支行", "45502": "沅江路支行", "45505": "城郊支行",
                      "45507": "田湾支行", "45508": "孝坪支行", "45509": "修溪支行", "45510": "伍家湾支行",
                      "45511": "谭家场支行", "45512": "潭湾支行", "45513": "桥头支行", "45514": "锦滨支行",
                      "45515": "安坪支行", "45516": "大水田支行", "45517": "龙泉岩支行", "45519": "火马冲支行",
                      "45520": "寺前支行", "45521": "小龙门支行", "45522": "锄头坪支行", "45523": "黄溪口支行",
                      "45524": "龙头庵支行", "45525": "后塘支行", "45526": "仙人湾支行", }})
    # 数据赋值
    self.ck_dk_202312_result_hz.set_index("开户机构", inplace=True)
    self.ck_dk_202312_result_hz.to_excel(writer, sheet_name="汇总", index=True)
    self.ck_dk_202312_result_mx.to_excel(writer, sheet_name="明细", index=False)
    writer.close()
    if writer is not None:
        messagebox.showinfo("Message", "数据导出成功！")
    else:
        messagebox.showinfo("Message", "请注意！数据导出失败！")


# 202312贷款余额导入文件路径字段
def ZJGH_import_dk202312_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    self.dk202312 = file_path

    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤2")


# 202312存款余额导入文件路径字段
def ZJGH_import_ck202312_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    self.ck202312 = file_path
    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤3")


# 当前贷款余额导入文件路径字段
def ZJGH_import_dk_now_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    self.dk_now = file_path
    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击步骤4")


# 当前存款余额导入文件路径字段
def ZJGH_import_ck_now_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    self.ck_now = file_path
    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击下一步")


# 贷款合同导入文件路径字段
def ZJGH_import_dkht_file(self):
    file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    self.dkht = file_path
    if file_path == '':
        messagebox.showinfo("Message", "未选中文件，请重新选择！")
        return False
    else:
        messagebox.showinfo("Message", "选中：" + file_path + "  请点击下一步")

