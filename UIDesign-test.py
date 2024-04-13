import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
from tkinter import messagebox

from function import selectExcel, dealData, excelOutput


# 是否成功标识，0正在处理，1处理成功，-1处理失败
class ExcelGUI11:

    # 主界面
    def main_window(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()
            # 按揭贷款分析按揭
            ml_analyse_button = tk.Button(self.master, text="1、按揭贷款数据分析", command=self.ajdk_analyse,
                                          relief=tk.RAISED,
                                          bd=1,
                                          bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                          highlightthickness=2)
            ml_analyse_button.pack(padx=10, pady=10, anchor='nw')
            # 贷款客户资金归行情况分析
            dkzjgh_analyse_button = tk.Button(self.master, text="2、贷款客户资金归行情况分析", command=self.dkzjgh_excel,
                                              relief=tk.RAISED,
                                              bd=1,
                                              bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                              highlightthickness=2)
            dkzjgh_analyse_button.pack(padx=10, pady=10, anchor='nw')

    # 贷款客户资金归行情况分析
    def dkzjgh_excel(self):

        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()
            # 状态提示
            self.str.set("")
            self.status_lable = tk.Label(self.master, textvariable=self.str)
            self.status_lable.pack(side="bottom")
            # 202312贷款余额导入文件
            import_dk202312_button = tk.Button(self.master, text="1、导入2023年底贷款余额表",
                                               command=self.import_dk202312_file,
                                               relief=tk.RAISED,
                                               bd=1,
                                               bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                               highlightthickness=2)
            import_dk202312_button.pack(padx=10, pady=10, anchor='nw')
            # 202312存款余额导入文件
            import_ck202312_button = tk.Button(self.master, text="2、导入2023年底存款余额表",
                                               command=self.import_ck202312_file,
                                               relief=tk.RAISED,
                                               bd=1,
                                               bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                               highlightthickness=2)
            import_ck202312_button.pack(padx=10, pady=10, anchor='nw')

            # 现在贷款余额导入文件
            import_dk_now_button = tk.Button(self.master, text="3、导入当前贷款余额表",
                                             command=self.import_dk_now_file,
                                             relief=tk.RAISED,
                                             bd=1,
                                             bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                             highlightthickness=2)
            import_dk_now_button.pack(padx=10, pady=10, anchor='nw')

            # 现在存款余额导入文件
            import_ck_now_button = tk.Button(self.master, text="4、导入当前存款余额表",
                                             command=self.import_ck_now_file,
                                             relief=tk.RAISED,
                                             bd=1,
                                             bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                             highlightthickness=2)
            import_ck_now_button.pack(padx=10, pady=10, anchor='nw')

            # 2023年贷款客户资金归行情况
            output_ck_dk_2023_button = tk.Button(self.master, text="5、分析数据",
                                                 command=self.dk_ck_file_analyse,
                                                 relief=tk.RAISED,
                                                 bd=1,
                                                 bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                                 highlightthickness=2)
            output_ck_dk_2023_button.pack(padx=10, pady=10, anchor='nw')

            # 导出数据
            output_ck_dk_2023_button = tk.Button(self.master, text="6、导出数据",
                                                 command=self.dk_ck_export,
                                                 relief=tk.RAISED,
                                                 bd=1,
                                                 bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                                 highlightthickness=2)
            output_ck_dk_2023_button.pack(padx=10, pady=10, anchor='nw')

            # 返回主界面按钮键
            return_button = tk.Button(self.master, text="7、返回主界面", command=self.main_window, relief=tk.RAISED, bd=1,
                                      bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                      highlightthickness=2)
            return_button.pack(padx=10, pady=10, anchor='nw')

    # 贷款存款余额数据分析
    def dk_ck_file_analyse(self):
        messagebox.showinfo("Message", "点击确定，开始分析数据！")
        self.str.set("")
        #获取2023年12月底贷款数据，使用B-D列
        dk202312_df = pd.read_excel(self.dk202312, sheet_name=0, usecols='B:D', header=4, dtype={'证件号码': str},
                                    keep_default_na=False)
        #删除重复行
        dk202312_1 = dk202312_df.drop_duplicates("证件号码", keep='first', inplace=False)
        #获取2023年12月底存款数据，使用F-Q列
        ck202312_df = pd.read_excel(self.ck202312, sheet_name=0, usecols='F, Q', header=1,
                                    dtype={'证件号码': str, '年平均': object})
        #202312底存款按证件号得到一个透视表
        ck202312_1 = pd.pivot_table(ck202312_df, index=["证件号码"], values=["年平均"], aggfunc=sum)
        #合并2023年底存款和贷款表
        dk202312_final_df = pd.merge(dk202312_1, ck202312_1, how='left', on="证件号码").rename(columns={"年平均": "2023年底存款年平均"})

        # dk_now_df = pd.read_excel(self.dk_now, sheet_name=0, usecols='B:D', header=4, dtype={'证件号码': str},
        #                           keep_default_na=False)
        # dk_now_1 = dk_now_df.drop_duplicates("证件号码", keep='first', inplace=False)
        # 获取当前存款数据，使用F-Q列
        ck_now_df = pd.read_excel(self.ck_now, sheet_name=0, usecols='F, Q', header=1,
                                  dtype={'证件号码': str, '年平均': object})
        #当前存款按证件号得到一个透视表
        ck_now_1 = pd.pivot_table(ck_now_df, index=["证件号码"], values=["年平均"], aggfunc=sum).rename(columns={"年平均": "当前存款年平均"})
        # dk_now_final_df = pd.merge(dk_now_1, ck_now_1, how='left', on="证件号码")
        # dk_now_final_df.rename(columns={"年平均": "当前存款年平均"})

        #合并当前存款数据和2023年底计算的存贷数据
        dk_2023_now_final_df = pd.merge(dk202312_final_df, ck_now_1, how='left', on="证件号码")
        #计算存款平均差值
        dk_2023_now_final_df.loc[:, '差值'] = dk_2023_now_final_df['当前存款年平均'] - dk_2023_now_final_df['2023年底存款年平均']
        dk_2023_now_final_df.head()

        self.ck_dk_202312_result_mx = dk_2023_now_final_df
        #透视表进行合并
        self.ck_dk_202312_result_hz = pd.pivot_table(dk_2023_now_final_df, index=["开户机构"], values=["2023年底存款年平均", "当前存款年平均", "差值"], aggfunc=sum)

        messagebox.showinfo("Message", "分析完成，请导出数据！")

    # 贷款余额数据分析导出数据
    def dk_ck_export(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        writer = pd.ExcelWriter(export_path, engine='openpyxl')
        #数据赋值
        self.ck_dk_202312_result_hz.to_excel(writer, sheet_name="汇总", index=True)
        self.ck_dk_202312_result_mx.to_excel(writer, sheet_name="明细", index=False)
        writer.close()
        if writer is not None:
            messagebox.showinfo("Message", "数据导出成功！")
        else:
            messagebox.showinfo("Message", "请注意！数据导出失败！")

    # 202312贷款余额导入文件路径字段
    def import_dk202312_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.dk202312 = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + file_path + "   点击步骤2。")

    # 202312存款余额导入文件路径字段
    def import_ck202312_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.ck202312 = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + file_path + "   点击步骤3。")

    # 当前贷款余额导入文件路径字段
    def import_dk_now_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.dk_now = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + file_path + "   点击步骤4。")

    # 当前存款余额导入文件路径字段
    def import_ck_now_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.ck_now = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + file_path + "   点击步骤5。")

    # 贷款合同导入文件路径字段
    def import_dkht_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.dkht = file_path
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + file_path + "   点击步骤5。")



    # 按揭贷款分析界面
    def ajdk_analyse(self):
        # 清空界面
        for widget in self.master.winfo_children():
            widget.destroy()
        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)  # 用frame标签
        # 状态提示
        self.str.set("")
        self.status_lable = tk.Label(self.master, textvariable=self.str)
        self.status_lable.pack(side="bottom")
        # 设置导入文件按钮键
        import_button = tk.Button(frame, text="1、选择文件", command=self.import_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        import_button.pack(side=tk.LEFT, padx=10, pady=10, anchor='nw')

        # 创建一个文本标签，并设置其显示的文本
        label = tk.Label(frame, text="选择年限")
        # 创建下拉框
        combo_box = ttk.Combobox(frame, width=10, exportselection=False)
        options = ['一年', '两年', '三年', '四年', '五年']  # 下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.combox = combo_box
        label.pack(side=tk.LEFT, padx=5)
        combo_box.pack(side=tk.LEFT, padx=5)
        frame.pack(padx=1, pady=1, anchor='nw')

        # 设置分析数据按钮键
        analyse_button = tk.Button(self.master, text="2、分析数据", command=self.analyse_excel, relief=tk.RAISED,
                                   bd=1,
                                   bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                   highlightthickness=2)
        analyse_button.pack(padx=10, pady=10, anchor='nw')
        # 设置导出数据按钮键
        export_button = tk.Button(self.master, text="3、导出表格", command=self.export_file, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        export_button.pack(padx=10, pady=10, anchor='nw')
        # 返回主界面按钮键
        return_button = tk.Button(self.master, text="4、返回主界面", command=self.main_window, relief=tk.RAISED, bd=1,
                                  bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0',
                                  highlightthickness=2)
        return_button.pack(padx=10, pady=10, anchor='nw')

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
        if file_path == '':
            self.str.set("未选中文件，请重新选择！")
        else:
            self.str.set("选中：" + self.expth + "   点击步骤2，请耐心等待。。。")

    # 处理分析数据
    def analyse_excel(self):
        Y_date = self.combox.get()  #获取下拉框值
        self.dataResult = dealData(self.expth, Y_date)
        if self.dataResult is not None:
            self.str.set("恭喜!处理成功。点击步骤3，导出表格。")
        else:
            self.str.set("处理失败！请重试或咨询技术人员。")

    def __init__(self, master):
        self.master = master
        self.status_lable = tk.Label(self.master, text="")
        self.master.title("辰溪农商银行数据分析小程序")
        self.master.geometry("800x500")  # 固定窗口大小
        self.expth = ''  # 初始化按揭贷款分析导入文件路径字段
        self.dk202312 = ''  # 初始化202312贷款余额导入文件路径字段
        self.ck202312 = ''  # 初始化202312存款余额导入文件路径字段
        self.dk_now = ''  # 初始化当前贷款余额导入文件路径字段
        self.ck_now = ''  # 初始化当前存款余额导入文件路径字段
        self.dkht = ''  # 初始化贷款合同导入文件路径字段
        self.str = tk.StringVar()
        self.dataResult = 0
        self.ck_dk_202312_result_mx = 0
        self.ck_dk_202312_result_hz = 0
        # 启动主界面
        self.main_window()
