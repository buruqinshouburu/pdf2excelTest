import pdfplumber
from openpyxl import Workbook
from tqdm import tqdm
import os
import yaml
import traceback
import logging

file = '../config.yml'

def analysis_table(pdf_file_path, filename, index):
    mergeCellList = list
    # 打开表格
    workbook = Workbook()
    sheet = workbook.active
    # 打开pdf
    with pdfplumber.open(pdf_file_path) as pdf:
        data_name = excel_data_folder+str(index)+'_'+filename.replace(".pdf", ".xlsx")
        # 遍历每页pdf
        for page in tqdm(pdf.pages):
            # 提取表格信息
            try:
                table = page.extract_table()
                print(type(table))
                # 格式化表格数据
                for i, row in enumerate(table):
                    sheet.append(row)
            except:
                break
    print(type(sheet))
    workbook.save(filename=data_name)



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
                analysis_table(pdf_data_folder+filename, filename, index)
                index = index+1
    except:
        s = traceback.format_exc()
        logging.error(s)