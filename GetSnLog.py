import os
import shutil
import pandas as pd


# 指定Excel檔案的路徑
excel_file_path =r'C:\Users\chicony\Desktop\test\SN.xlsx'

# 指定要搜尋的資料夾路徑
folder_path =r'D:\all\all'
destination_path=r'C:\Users\chicony\Desktop\test'


# 讀取Excel檔案
workbook = pd.read_excel(excel_file_path)
columTitle = workbook.columns[0]
columData=workbook.iloc[:,0].to_list()

for SN in columData:
    for filename in os.listdir(folder_path):
        if SN in filename:             
            folder_name = str(columTitle)
            newFolderPath = os.path.join(destination_path, folder_name)
            os.makedirs(newFolderPath, exist_ok=True)
            filePath=os.path.join(folder_path, filename)
            newdestination_path = os.path.join(newFolderPath, filename)
            print(newdestination_path)
            print(filePath)
            if os.path.isfile(filePath): 
                shutil.copy2(filePath,newdestination_path)
            elif os.path.isdir(filePath):   
                shutil.copytree(filePath,newdestination_path)


