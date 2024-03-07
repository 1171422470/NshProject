import tkinter
from UIDesign import ExcelGUI
from function import getData, selectExcel

# Press the green button in the gutter to run the script.
# 运行程序
if __name__ == '__main__':
    root = tkinter.Tk()
    app = ExcelGUI(root)
    root.mainloop()
    # excelOutput("/home/xiaoyifan/0304数据源.xlsx", "/home/xiaoyifan/结果.xlsx")




