import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox

from function import selectExcel, dealData, excelOutput
from datetime import datetime
class ckAnalysis:
    #定义初始值
    def __init__(self,master):
        self.master = master
        self.master.title("辰溪农商银行数据分析小程序")
        self.master.geometry("800x500")  # 固定窗口大小
        self.expth = ''  # 初始化导入文件路径字段
        self.str = tk.StringVar() #初始化
        self.UI()
    #界面
    def UI(self):
        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)  # 用frame标签
        #文件导入按键
        manalyse_button = tk.Button(frame, text="1,选择定期存款余额表", command=self.import_file,
                                      relief=tk.RAISED,
                                      bd=1,
                                      bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                      highlightthickness=2)
        manalyse_button.pack(padx=10, pady=10, anchor='nw',side=tk.LEFT)


        # 创建一个文本标签，并设置其显示的文本
        label = tk.Label(frame, text="选择年龄段")
        # 创建下拉框
        combo_box = ttk.Combobox(frame, width=10, exportselection=False)
        options = ['20周岁以下','20到60周岁','60周岁以上']  # 下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.combox = combo_box
        label.pack(side=tk.LEFT, padx=5)
        combo_box.pack(side=tk.LEFT, padx=5)
        frame.pack(padx=1, pady=1, anchor='nw')

        #设置分析数据按键
        # 定期存款分析
        output_button = tk.Button(self.master, text="2,分析数据",
                                             command=self.DQCKAnalysis,
                                             relief=tk.RAISED,
                                             bd=1,
                                             bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                             highlightthickness=2)
        output_button.pack(padx=10, pady=10, anchor='nw')

        #导出数据
        export_button = tk.Button(self.master, text="3、导出表格", command=self.export_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        export_button.pack(padx=10, pady=10, anchor='nw')
    # 定期存款数据分析
    def DQCKAnalysis(self):
        age_date = self.combox.get()  # 获取下拉框值
        self.dataResult = self.DQCKdeal(self.expth, age_date)
        if self.dataResult is not None:
            self.str.set("恭喜!处理成功。点击步骤3，导出表格。")
        else:
            self.str.set("处理失败！请重试或咨询技术人员。")

    #定期存款数据处理
    def DQCKdeal(self,address,date):
        data = pd.read_excel(address, sheet_name=0, header=1)
        data['科目编码'].fillna(0, inplace=True)  # 假设0是一个合适的填充值
        data['科目编码'] = data['科目编码'].astype(int)

        #按年龄段进行筛选
        # 只筛选定期，根据科目编码筛选：一年定期：20040103，二年定期：20040104，三年定期：20040105，五年：20040106
        #金额大于五万元
        if date == '20周岁以下':
            data['Age'] = data['证件号码'].apply(self.calculate_age_from_id_card)
            self.result = data[(data['Age'] < 20) &((data['科目编码'] == 20040103)
            | (data['科目编码'] == 20040104) | (data['科目编码'] == 20040105) | (data['科目编码'] == 20040106))
                               & (data['账户余额(元)'] > 50000)]
        elif date == '20到60周岁':
            data['Age'] = data['证件号码'].apply(self.calculate_age_from_id_card)
            self.result = data[(data['Age'] >= 20) & (data['Age'] <= 60) &((data['科目编码'] == 20040103)
            | (data['科目编码'] == 20040104) | (data['科目编码'] == 20040105) | (data['科目编码'] == 20040106))
                               & (data['账户余额(元)'] > 50000)]

        else:
            data['Age'] = data['证件号码'].apply(self.calculate_age_from_id_card)
            self.result = data[data['Age'] > 60 &((data['科目编码'] == 20040103)
            | (data['科目编码'] == 20040104) | (data['科目编码'] == 20040105) | (data['科目编码'] == 20040106))
                               & (data['账户余额(元)'] > 50000)]



        # 透视表进行合并,按支行汇总
        self.counts = self.result.groupby('开户机构')['开户机构'].size()

        self.zh_result_hz = pd.pivot_table(self.result, index=["开户机构"],
                                                     values=["账户余额(元)"],
                                                     aggfunc=sum)
        # 透视表进行合并,按个人汇总
        self.GR_result_hz = pd.pivot_table(self.result, index=["证件号码"],
                                           values=["开户机构","户名","账户余额(元)"],
                                           aggfunc=sum)

        #判断文件是否导入
        if data is None:
            return None

    # 根据身份证计算年龄
    def calculate_age_from_id_card(self,id_card):

        try:
            birth_date_str = id_card[6:14]

            birth_date = datetime.strptime(birth_date_str, '%Y%m%d')
            today = datetime.now()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            return age
        except:
            print("非个人客户")



    # 导入数据源文件
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        # if file_path:
        #     self.data_frame = pd.read_excel(file_path)
        #     self.data_text.delete(1.0, tk.END)
        #     self.data_text.insert(tk.END, self.data_frame.to_string())
        self.expth = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + self.expth + "   点击步骤2，请耐心等待。。。")

    # 数据导出,导出格式为Excel
    def export_file(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        writer = pd.ExcelWriter(export_path, engine='openpyxl')
        # 数据赋值
        self.counts.to_excel(writer, sheet_name="汇总", index=True,header=['户数'])
        self.GR_result_hz.to_excel(writer, sheet_name="明细", index=True)
        writer.close()
        if writer is not None:
            messagebox.showinfo("Message", "数据导出成功！")
        else:
            messagebox.showinfo("Message", "请注意！数据导出失败！")
