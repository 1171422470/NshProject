# 导入窗口控制器
from controller import Controller
# 导入布局界面
from ui import Win

if __name__ == '__main__':
    # 实例化一个窗口 将窗口控制器的实例传入
    app = Win(Controller())

    # 在这可对窗口操作 设置图标等.
    # 设置窗口大小、居中
    width = 800
    height = 530
    screenwidth = app.winfo_screenwidth()
    screenheight = app.winfo_screenheight()
    geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    app.geometry(geometry)
    # 运行程序
    app.mainloop()
