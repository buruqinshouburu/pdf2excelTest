import os
import yaml
import traceback
import logging
import camelot

file = '../config.yml'

def analysis_table(pdf_file_path, filename, index):
    data_name = excel_data_folder + str(index) + '_' + filename.replace(".pdf", ".csv")
    # 打开pdf
    tables = camelot.read_pdf(pdf_file_path)
    for table in tables:
        print(table)
        table.to_csv(data_name)
    #tables.export(data_name, f='xlsx', compress=True)



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