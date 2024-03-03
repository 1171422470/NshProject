import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk

from function import selectExcel


class ExcelGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("数据处理程序")
        self.master.geometry("800x500")   #固定窗口大小
        self.expth = ''#初始化路径字段
        SVBar = tk.Scrollbar(self.master) #设置纵向滚动条
        SVBar.pack(side=tk.RIGHT, fill="y")

        SHBar = tk.Scrollbar(self.master,orient = tk.HORIZONTAL) #设置横向滚动条
        SHBar.pack(side=tk.BOTTOM, fill="x")
        #设置按钮键
        self.import_button = tk.Button(self.master, text="选择文件", command=self.import_file, relief=tk.RAISED, bd=1, bg='lightblue', fg='black', padx=1, pady=1, borderwidth=1, border='0', highlightthickness=2)
        self.import_button.pack(padx=10, pady=10, anchor='nw')

        #添加筛选
        # 添加标签和输入框
        frame = tk.Frame(self.master, width=600)#用frame标签
        #创建下拉框
        combo_box = ttk.Combobox(frame,width=10,exportselection=False)
        options =['开户机构','户名','贷款账号']#下拉框赋值
        combo_box.set(options[0])
        combo_box['values'] = options
        self.combox = combo_box
        #创建输入框
        entry = tk.Entry(frame)
        self.entry = entry
        #创建按钮
        button = tk.Button(frame, text="查询",bg='lightblue',highlightthickness=2,command=self.dataGet)

        #创建按钮
        button1 = tk.Button(frame,text="导出数据",bg='lightblue',highlightthickness=2,command=self.export_file)
        #加入布局排列
        combo_box.pack(side=tk.LEFT,padx=5)
        entry.pack(side=tk.LEFT)
        button.pack(side=tk.LEFT)
        button1.pack(side=tk.RIGHT,padx=100)
        frame.pack(padx=10, pady=10, anchor='nw')

        #设置文本显示框
        self.data_text = tk.Text(self.master, height=40, width=120,xscrollcommand=SHBar.set,yscrollcommand=SVBar.set, wrap="none")
        SHBar.config(command=self.data_text.xview)#设置滚动条
        SVBar.config(command=self.data_text.yview)
        self.data_text.pack(padx=10, pady=10,anchor='nw')
        self.master.resizable(0, 0)#固定窗口大小


    def dataGet(self):
        if self.combox.get() == '开户机构':
            data = selectExcel(self.expth,self.combox.get(),int(self.entry.get()))
        else:
            data = selectExcel(self.expth, self.combox.get(), self.entry.get())
        self.data_text.delete(1.0, tk.END)
        self.data_text.insert(1.0, data)
        return data
    # 文件导入，打开一个Excel
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        if file_path:
            self.data_frame = pd.read_excel(file_path)
            self.data_text.delete(1.0, tk.END)
            self.data_text.insert(tk.END, self.data_frame.to_string())
            self.expth = file_path

    # 数据导出,导出格式为Excel
    def export_file(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        data = self.dataGet()
        df = pd.DataFrame(data)
        df.to_excel(export_path,index=False)