# from tkinter import filedialog
# import tkinter as tk
import pandas as pd
from tkinter import messagebox
# import openpyxl
#
# #读取指定的excel表数据，返回excel表内容
# def getData(exAdress):
#     data = pd.read_excel(exAdress,sheet_name=0)
#     return  data
#
# #通过指定的字段，从excel表中匹配相符的数据，返回excel表内容
# def selectExcel(exAdress,name,value):
#     data = getData(exAdress)
#     data1 = data.loc[(data[name]) == value]
#     return data1





# 读取指定的excel表数据，返回excel表内容
def getData(exAdress):
    data = pd.read_excel(exAdress, sheet_name=0, usecols='B:G, J, O, Y:Z', header=4, dtype={'贷款账号': str})
    if data is None:
        return None
    return data


# 通过指定的字段，从excel表中匹配相符的数据，返回excel表内容
def selectExcel(exAdress, name, value):
    data = getData(exAdress)
    data1 = data.loc[(data[name]) == value]
    return data1


# 处理数据
#输入值：文件地址 年份期限
def dealData(exAdress,Y_date):
    data = getData(exAdress)
    if data is None:
        return None
    # 贷款形态为正常
    data = data[data["贷款形态"] == "正常"]
    # 担保方式为抵押
    data = data[data["担保方式"] == "抵押"]
    # 贷款余额大于0
    data = data[data["贷款余额(元)"] > 0]
    # 贷款期限大于5年的，判断出是按揭贷款
    data[['借款日期', '到期日期', '数据日期']] = data[['借款日期', '到期日期', '数据日期']].astype('int')

    # data = data[(data["到期日期"] - data["借款日期"]) > 50000]
    # 已经还款3年的贷款
    # data = data[(data["数据日期"] - data["借款日期"]) >= 50000]
    if Y_date == '一年':
        data = data[(data["数据日期"] - data["借款日期"]) <= 10000]
    elif Y_date == '两年':
        data = data[((data["数据日期"] - data["借款日期"]) <= 20000) & ((data["数据日期"] - data["借款日期"]) > 10000)]
    elif Y_date == '三年':
        data = data[((data["数据日期"] - data["借款日期"]) <= 30000) & ((data["数据日期"] - data["借款日期"]) > 20000)]
    elif Y_date == '四年':
        data = data[((data["数据日期"] - data["借款日期"]) <= 40000) & ((data["数据日期"] - data["借款日期"]) > 30000)]
    else:
        data = data[(data["数据日期"] - data["借款日期"]) > 40000]

    # 替换机构号为行名
    data = data.replace({"开户机构": {45500: "营业部", 45501: "辰阳支行", 45502: "沅江路支行", 45505: "城郊支行",
                                      45507: "田湾支行", 45508: "孝坪支行", 45509: "修溪支行", 45510: "伍家湾支行",
                                      45511: "谭家场支行", 45512: "潭湾支行", 45513: "桥头支行", 45514: "锦滨支行",
                                      45515: "安坪支行", 45516: "大水田支行", 45517: "龙泉岩支行", 45519: "火马冲支行",
                                      45520: "寺前支行", 45521: "小龙门支行", 45522: "锄头坪支行", 45523: "黄溪口支行",
                                      45524: "龙头庵支行", 45525: "后塘支行", 45526: "仙人湾支行", }}).astype(str)
    messagebox.showinfo("Message", "数据处理完毕！可以导出数据！")
    return data


# 导出数据
def excelOutput(outAdress, data):
    writer = pd.ExcelWriter(outAdress, engine='openpyxl')
    # 导出明细表格
    df1 = pd.DataFrame(data)
    df1.to_excel(writer, sheet_name='明细', index=False)
    # 计算客户户数
    count = data["开户机构"].value_counts().to_frame().reset_index()
    # 导出汇总表格
    df2 = pd.DataFrame({'机构名称': count["开户机构"], '户数': count["count"]})
    print(df2)
    df2.to_excel(writer, sheet_name='汇总', index=False)
    writer.close()




