import random
from tkinter import *
from tkinter.ttk import *
from PIL import Image,ImageTk
from tkinter import messagebox
import tkinter as TK
from tkinter import filedialog
from UIDesign import ExcelGUI
from tkinter import filedialog
import pandas as pd
from CKAnalysis import ckAnalysis
class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        # 创建一个PhotoImage对象，并加载图片文件
        self.image = Image.open("image/nsh.png")
        self.image = self.image.resize((200, 40))
        self.image = ImageTk.PhotoImage(self.image)

        self.tk_frame_top_1 = self.__tk_frame_top_1(self)
        self.tk_button_top_button1 = self.__tk_button_top_button1(self.tk_frame_top_1)
        self.tk_button_top_button2 = self.__tk_button_top_button2(self.tk_frame_top_1)
        self.tk_button_top_button3 = self.__tk_button_top_button3(self.tk_frame_top_1)
        self.tk_button_top_button4 = self.__tk_button_top_button4(self.tk_frame_top_1)
        self.tk_label_top_label = self.__tk_label_top_label(self.tk_frame_top_1)
        self.tk_frame_left_1 = self.__tk_frame_left_1(self)
        self.tk_frame_left_top1 = self.__tk_frame_left_top1(self.tk_frame_left_1)
        self.tk_button_left_top_button1 = self.__tk_button_left_top_button1(self.tk_frame_left_top1)
        self.tk_button_left_top_button2 = self.__tk_button_left_top_button2(self.tk_frame_left_top1)
        self.tk_label_left_label1 = self.__tk_label_left_label1(self.tk_frame_left_1)
        self.tk_button_left_button1 = self.__tk_button_left_button1(self.tk_frame_left_1)
        self.tk_button_left_button2 = self.__tk_button_left_button2(self.tk_frame_left_1)
        self.tk_button_left_button3 = self.__tk_button_left_button3(self.tk_frame_left_1)
        self.tk_button_left_button4 = self.__tk_button_left_button4(self.tk_frame_left_1)
        self.tk_button_left_button5 = self.__tk_button_left_button5(self.tk_frame_left_1)
        self.tk_button_left_button6 = self.__tk_button_left_button6(self.tk_frame_left_1)
        self.tk_frame_botton = self.__tk_frame_botton(self)
        self.tk_label_botton_label1 = self.__tk_label_botton_label1(self.tk_frame_botton)
        self.tk_label_button_label2 = self.__tk_label_button_label2(self.tk_frame_botton)
        # self.tk_frame_main_frame = self.__tk_frame_main_frame(self)

    def __win(self):
        self.title("数据分析")
        # 设置窗口大小、居中
        width = 800
        height = 550
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)
        self.configure(background="#9DCAFF")
        self.resizable(width=False, height=False)

    def scrollbar_autohide(self, vbar, hbar, widget):
        """自动隐藏滚动条"""

        def show():
            if vbar: vbar.lift(widget)
            if hbar: hbar.lift(widget)

        def hide():
            if vbar: vbar.lower(widget)
            if hbar: hbar.lower(widget)

        hide()
        widget.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Leave>", lambda e: hide())
        if hbar: hbar.bind("<Enter>", lambda e: show())
        if hbar: hbar.bind("<Leave>", lambda e: hide())
        widget.bind("<Leave>", lambda e: hide())

    def v_scrollbar(self, vbar, widget, x, y, w, h, pw, ph):
        widget.configure(yscrollcommand=vbar.set)
        vbar.config(command=widget.yview)
        vbar.place(relx=(w + x) / pw, rely=y / ph, relheight=h / ph, anchor='ne')

    def h_scrollbar(self, hbar, widget, x, y, w, h, pw, ph):
        widget.configure(xscrollcommand=hbar.set)
        hbar.config(command=widget.xview)
        hbar.place(relx=x / pw, rely=(y + h) / ph, relwidth=w / pw, anchor='sw')

    def create_bar(self, master, widget, is_vbar, is_hbar, x, y, w, h, pw, ph):
        vbar, hbar = None, None
        if is_vbar:
            vbar = Scrollbar(master)
            self.v_scrollbar(vbar, widget, x, y, w, h, pw, ph)
        if is_hbar:
            hbar = Scrollbar(master, orient="horizontal")
            self.h_scrollbar(hbar, widget, x, y, w, h, pw, ph)
        self.scrollbar_autohide(vbar, hbar, widget)

    def __tk_frame_top_1(self, parent):
        frame = Frame(parent)
        frame.place(x=0, y=0, width=800, height=80)
        return frame

    def __tk_button_top_button1(self, parent):
        btn = Button(parent, text="工具", takefocus=False,)
        btn.place(x=340, y=18, width=92, height=40)
        return btn

    def __tk_button_top_button2(self, parent):
        btn = Button(parent, text="密码", takefocus=False, )
        btn.place(x=466, y=17, width=92, height=40)
        return btn

    def __tk_button_top_button3(self, parent):
        btn = Button(parent, text="设置", takefocus=False, )
        btn.place(x=587, y=17, width=92, height=40)
        return btn

    def __tk_button_top_button4(self, parent):
        btn = Button(parent, text="退出", takefocus=False, )
        btn.place(x=708, y=17, width=92, height=40)
        return btn

    def __tk_label_top_label(self, parent):
        label = Label(parent, text="", anchor="center", image=self.image)
        label.place(x=2, y=14, width=261, height=52)
        return label

    def __tk_frame_left_1(self, parent):
        frame = Frame(parent, )
        frame.place(x=2, y=86, width=201, height=386)
        return frame

    def __tk_frame_left_top1(self, parent):
        frame = Frame(parent, )
        frame.place(x=1, y=0, width=200, height=41)
        return frame

    def __tk_button_left_top_button1(self, parent):
        btn = Button(parent, text="菜单", takefocus=False, )
        btn.place(x=0, y=2, width=96, height=35)
        return btn

    def __tk_button_left_top_button2(self, parent):
        btn = Button(parent, text="常用菜单", takefocus=False, )
        btn.place(x=100, y=2, width=96, height=35)
        return btn

    def __tk_label_left_label1(self, parent):
        label = Label(parent, text="导航栏", anchor="center",background="#14409A" )
        label.place(x=2, y=46, width=196, height=38)
        return label

    def __tk_button_left_button1(self, parent):
        btn = Button(parent, text="贷款数据分析", takefocus=False, command=self.DKUI)
        btn.place(x=3, y=88, width=195, height=50)
        return btn

    def __tk_button_left_button2(self, parent):
        btn = Button(parent, text="存款数据分析", takefocus=False, command=self.CKUI)
        btn.place(x=3, y=137, width=195, height=50)
        return btn

    def __tk_button_left_button3(self, parent):
        btn = Button(parent, text="客户数据分析", takefocus=False, )
        btn.place(x=3, y=188, width=195, height=50)
        return btn

    def __tk_button_left_button4(self, parent):
        btn = Button(parent, text="客户流水分析", takefocus=False, )
        btn.place(x=3, y=236, width=195, height=50)
        return btn

    def __tk_button_left_button5(self, parent):
        btn = Button(parent, text="卡业务分析", takefocus=False, )
        btn.place(x=3, y=289, width=195, height=50)
        return btn

    def __tk_button_left_button6(self, parent):
        btn = Button(parent, text="其它业务", takefocus=False, )
        btn.place(x=3, y=334, width=195, height=50)
        return btn

    def __tk_frame_botton(self, parent):
        frame = Frame(parent, )
        frame.place(x=0, y=498, width=800, height=52)
        return frame

    def __tk_label_botton_label1(self, parent):
        label = Label(parent, text="日期：2024-04", anchor="center", )
        label.place(x=0, y=21, width=269, height=30)
        return label

    def __tk_label_button_label2(self, parent):
        label = Label(parent, text="辰溪农商行", anchor="center", )
        label.place(x=525, y=17, width=269, height=30)
        return label

    # def __tk_frame_main_frame(self,parent):
    #     frame = Frame(parent)
    #     frame.place(x=228, y=87, width=550, height=379)

    #跳转贷款分析界面
    def DKUI(self):
        root = TK.Tk()
        app = ExcelGUI(root)
        root.mainloop()

    #跳转存款分析界面
    def CKUI(self):
        root = TK.Tk()
        app = ckAnalysis(root)
        root.mainloop()

class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        super().__init__()
        self.__event_bind()
        self.__style_config()
        self.ctl.init(self)

    def __event_bind(self):
        pass

    def __style_config(self):
        pass


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
