import comtypes.client as client
import os
excel_path = "E:/test2/1_实施人员签到表.xlsx"                         #这里是Excel文件的路径
pdf_path = "E:/test1/实施人员签到表2.pdf"                            #这里是输出PDF的保存路径

xlApp = client.CreateObject("Excel.Application")
books = xlApp.Workbooks.Open(excel_path)
books.ExportAsFixedFormat(0, pdf_path)
xlApp.Quit()

