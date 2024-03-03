from tkinter import filedialog
import tkinter as tk
import pandas as pd
import openpyxl

#读取指定的excel表数据，返回excel表内容
def getData(exAdress):
    data = pd.read_excel(exAdress,sheet_name=0)
    return  data

#通过指定的字段，从excel表中匹配相符的数据，返回excel表内容
def selectExcel(exAdress,name,value):
    data = getData(exAdress)
    data1 = data.loc[(data[name]) == value]
    return data1




