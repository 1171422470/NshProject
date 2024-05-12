from tkinter.messagebox import showinfo
import tkinter as TK
# 导入UI 将 Controller 的属性 ui 类型设置成 Win
from ui import Win
class Controller:
    # 导入UI类后，替换以下的 object 类型，将获得 IDE 属性提示功能
    ui: Win
    def __init__(self):
        pass

    def init(self, ui):
        """
        得到UI实例，对组件进行初始化配置
        """
        self.ui = ui

    # 隐藏控件
    def hide(self):
        # 组件初始状态设置为隐藏
        self.ui.tk_label_frame_labelframe1.place_forget()
        self.ui.tk_button_ck_button1.place_forget()
        self.ui.tk_button_ck_button2.place_forget()
        self.ui.tk_button_ck_button3.place_forget()
        self.ui.tk_select_box_ck_select1.place_forget()
        self.ui.tk_label_ck_label1.place_forget()
        self.ui.tk_button_ck_button4.place_forget()
        self.ui.tk_button_ck_f_button1.place_forget()
        self.ui.tk_button_ck_f_button2.place_forget()

        self.ui.tk_label_frame_labelframe2.place_forget()
        self.ui.tk_button_dk_f_button1.place_forget()
        self.ui.tk_button_dk_f_button2.place_forget()
        self.ui.tk_button_dk_f_button3.place_forget()
        self.ui.tk_button_dk_f_button4.place_forget()
        self.ui.tk_button_dk_f_button5.place_forget()

        self.ui.tk_button_ajdk_button1.place_forget()
        self.ui.tk_button_ajdk_button2.place_forget()
        self.ui.tk_button_ajdk_button3.place_forget()
        self.ui.tk_button_ajdk_button4.place_forget()
        self.ui.tk_select_box_ajdk_select1.place_forget()
        self.ui.tk_label_ajdk_label1.place_forget()

        self.ui.tk_button_zjgh_button1.place_forget()
        self.ui.tk_button_zjgh_button2.place_forget()
        self.ui.tk_button_zjgh_button3.place_forget()
        self.ui.tk_button_zjgh_button4.place_forget()
        self.ui.tk_button_zjgh_button5.place_forget()
        self.ui.tk_button_zjgh_button6.place_forget()

        self.ui.tk_button_lskh_button1.place_forget()
        self.ui.tk_button_lskh_button2.place_forget()
        self.ui.tk_button_lskh_button3.place_forget()
        self.ui.tk_button_lskh_button4.place_forget()
        self.ui.tk_button_lskh_button5.place_forget()
        self.ui.tk_label_lskh_label1.place_forget()
        self.ui.tk_select_box_lskh_select1.place_forget()

        self.ui.tk_button_dkcs_button1.place_forget()
        self.ui.tk_button_dkcs_button2.place_forget()
        self.ui.tk_button_dkcs_button3.place_forget()
        self.ui.tk_button_dkcs_button4.place_forget()
        self.ui.tk_button_dkcs_button5.place_forget()
    # 绑定存款分析按钮
    def ckfx(self, evt):
        self.hide()
        self.ui.tk_label_frame_labelframe1.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_ck_f_button1.place(x=30, y=8, width=180, height=40)
        self.ui.tk_button_ck_f_button2.place(x=30, y=300, width=180, height=30)
    #绑定贷款分析按钮
    def dkfx(self, evt):
        self.hide()
        self.ui.tk_label_frame_labelframe2.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_dk_f_button1.place(x=30, y=8, width=250, height=40)
        self.ui.tk_button_dk_f_button2.place(x=30, y=58, width=250, height=40)
        self.ui.tk_button_dk_f_button3.place(x=30, y=108, width=250, height=40)
        self.ui.tk_button_dk_f_button4.place(x=30, y=158, width=250, height=40)
        self.ui.tk_button_dk_f_button5.place(x=30, y=300, width=250, height=30)

    #定期存款余额表分析
    def ck_dqkc(self,evt):
        self.hide()
        self.ui.tk_label_frame_labelframe1.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_ck_button1.place(x=0, y=8, width=158, height=30)
        self.ui.tk_button_ck_button2.place(x=0, y=58, width=158, height=30)
        self.ui.tk_button_ck_button3.place(x=0, y=108, width=158, height=30)
        self.ui.tk_select_box_ck_select1.place(x=312, y=14, width=150, height=30)
        self.ui.tk_label_ck_label1.place(x=212, y=13, width=92, height=30)
        self.ui.tk_button_ck_button4.place(x=0, y=300, width=158, height=30)
    #存款分析界面菜单退出
    def ck_exit(self,evt):
        self.ui.tk_label_frame_labelframe1.place_forget()
    #定期存款分析
    def ck_button2(self,evt):
        age_date = self.ui.tk_select_box_ck_select1.get() #获取下拉框值
        self.ui.ck.DQCKAnalysis(age_date)
    #定期存款导入文件
    def ck_button1(self,evt):
        self.ui.ck.import_file()
    #定期存款导出文件
    def ck_button3(self,evt):
        self.ui.ck.export_file()
    #贷款界面退出
    def dk_exit(self,evt):
        self.ui.tk_label_frame_labelframe2.place_forget()
    #1.正常类按揭贷款数据分析
    def dk_ajdk(self,evt):
        self.hide()
        self.ui.tk_label_frame_labelframe2.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_ajdk_button1.place(x=0, y=8, width=310, height=37)
        self.ui.tk_button_ajdk_button2.place(x=0, y=58, width=139, height=30)
        self.ui.tk_button_ajdk_button3.place(x=0, y=108, width=139, height=30)
        self.ui.tk_button_ajdk_button4.place(x=0, y=300, width=139, height=30)
        self.ui.tk_select_box_ajdk_select1.place(x=420, y=8, width=116, height=30)
        self.ui.tk_label_ajdk_label1.place(x=340, y=8, width=62, height=30)
    #按键贷款导入文件
    def DKAnalysis_AJDK_import_file(self,evt):
        self.ui.dk.DKAnalysis_AJDK_import_file()
    #按键贷款分析
    def DKAnalysis_AJDK_analyse_excel(self,evt):
        Y_date = self.ui.tk_select_box_ajdk_select1.get()  # 获取下拉框值
        self.ui.dk.DKAnalysis_AJDK_analyse_excel(Y_date)
    #按键贷款导出
    def DKAnalysis_AJDK_export_file(self,evt):
        self.ui.dk.DKAnalysis_AJDK_export_file()
    #资金归行
    def dk_zjgh(self,evt):
        self.hide()
        self.ui.tk_label_frame_labelframe2.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_zjgh_button1.place(x=30, y=8, width=264, height=30)
        self.ui.tk_button_zjgh_button2.place(x=30, y=58, width=264, height=30)
        self.ui.tk_button_zjgh_button3.place(x=30, y=108, width=264, height=30)
        self.ui.tk_button_zjgh_button4.place(x=30, y=158, width=264, height=30)
        self.ui.tk_button_zjgh_button5.place(x=30, y=208, width=264, height=30)
        self.ui.tk_button_zjgh_button6.place(x=30, y=300, width=264, height=30)
    def DKAnalysis_ZJGH_import_dk202312_file(self,evt):
        self.ui.dk.DKAnalysis_ZJGH_import_dk202312_file()
    def DKAnalysis_ZJGH_import_ck202312_file(self,evt):
        self.ui.dk.DKAnalysis_ZJGH_import_ck202312_file()
    def DKAnalysis_ZJGH_import_ck_now_file(self,evt):
        self.ui.dk.DKAnalysis_ZJGH_import_ck_now_file()
    def DKAnalysis_ZJGH_dk_ck_file_analyse(self,evt):
        self.ui.dk.DKAnalysis_ZJGH_dk_ck_file_analyse()
    def DKAnalysis_ZJGH_export(self,evt):
        self.ui.dk.DKAnalysis_ZJGH_export()
    #近5年有贷款往来，现在已经结清了的客户
    def dk_lskh(self,evt):
        self.hide()
        self.ui.tk_label_frame_labelframe2.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_lskh_button1.place(x=0, y=8, width=310, height=30)
        self.ui.tk_button_lskh_button2.place(x=0, y=58, width=310, height=30)
        self.ui.tk_button_lskh_button3.place(x=0, y=108, width=121, height=30)
        self.ui.tk_button_lskh_button4.place(x=0, y=158, width=121, height=30)
        self.ui.tk_button_lskh_button5.place(x=0, y=300, width=121, height=30)
        self.ui.tk_label_lskh_label1.place(x=313, y=7, width=123, height=30)
        self.ui.tk_select_box_lskh_select1.place(x=442, y=6, width=111, height=30)
    def DKAnalysis_LSKH_yjq_import_file(self,evt):
        self.ui.dk.DKAnalysis_LSKH_yjq_import_file()
    def DKAnalysis_LSKH_wjq_import_file(self,evt):
        self.ui.dk.DKAnalysis_LSKH_wjq_import_file()
    def DKAnalysis_LSKH_analysis(self,evt):
        Y_date = self.ui.tk_select_box_lskh_select1.get() # 获取下拉框值
        self.ui.dk.DKAnalysis_LSKH_analysis(Y_date)
    def DKAnalysis_LSKH_export(self,evt):
        self.ui.dk.DKAnalysis_LSKH_export()

    # 贷款测算
    def dk_dkcs(self, evt):
        self.hide()
        self.ui.tk_label_frame_labelframe2.place(x=225, y=85, width=561, height=386)
        self.ui.tk_button_dkcs_button1.place(x=0, y=8, width=310, height=30)
        self.ui.tk_button_dkcs_button2.place(x=0, y=58, width=310, height=30)
        self.ui.tk_button_dkcs_button3.place(x=0, y=108, width=121, height=30)
        self.ui.tk_button_dkcs_button4.place(x=0, y=158, width=121, height=30)
        self.ui.tk_button_dkcs_button5.place(x=0, y=300, width=121, height=30)

    def DKAnalysis_DKCS_DKYEB_import_file(self, evt):
        self.ui.dk.DKAnalysis_DKCS_DKYEB_import_file()

    def DKAnalysis_DKCS_DKHSDJB_import_file(self, evt):
        self.ui.dk.DKAnalysis_DKCS_DKHSDJB_import_file()

    def DKAnalysis_DKCS_analyse_excel(self, evt):
        self.ui.dk.DKAnalysis_DKCS_analyse_excel()

    def DKAnalysis_DKCS_export_file(self, evt):
        self.ui.dk.DKAnalysis_DKCS_export_file()
