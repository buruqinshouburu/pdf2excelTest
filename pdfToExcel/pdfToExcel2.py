import pdfplumber
import os
import yaml
import traceback
import logging
import pandas as pd

file = '../config.yml'

def transFormTable2(pdf_file_path, filename, index):
    data_name = excel_data_folder + str(index) + '_' + filename.replace(".pdf", ".xlsx")
    # 读取pdf文件,保存为pdf实例
    pdf = pdfplumber.open(pdf_file_path)
    # 访问第一页
    first_page = pdf.pages[0]
    # 自动读取表格信息,返回列表

    table = first_page.extract_table()
    # 将列表转化为dataframe
    table_data = pd.DataFrame(table[1:], columns=table[0])
    # 保存为excel
    table_data.to_excel(data_name, index=False)


def readConfig():
    global pdf_data_folder
    global excel_data_folder
    # 读取文件
    with open(file,'rb') as f:
        data = yaml.safe_load_all(f)
        dataDict = list(data).pop(0)
        pdf_data_folder = dataDict.get("pdf_data_folder")
        excel_data_folder = dataDict.get("excel_data_folder")



if __name__ == "__main__":
    try:
        readConfig()
        files = os.listdir(pdf_data_folder)
        index = 1
        for filename in files:
            if "pdf".__eq__(filename.split(".")[1]):
                #analysis_table(pdf_data_folder+filename, filename, index)
                transFormTable2(pdf_data_folder+filename, filename, index)
                index = index+1
    except:
        s = traceback.format_exc()
        logging.error(s)