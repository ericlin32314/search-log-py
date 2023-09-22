import os
import shutil
import xlwt
import openpyxl
import pandas as pd
import time
# 指定Excel檔案的路徑
excel_file_path =r'C:\Users\chicony\Desktop\test\SN.xlsx'
# 指定要搜尋的資料夾路徑
folder_path =r'D:\all\all'
#要放置找到檔案的資料夾
destination_path=r'C:\Users\chicony\Desktop\test'
# 讀取Excel檔案
workbook = pd.read_excel(excel_file_path)
columTitle = workbook.columns[0]
columData=workbook.iloc[:,0].to_list()
allfilename=os.listdir(folder_path)
textlenth=30
errlist=[]
folder_name = str('SnYouWant'+'_'+columTitle+'_'+str(int(time.time())))
newSnFolderPath = os.path.join(destination_path, folder_name)
os.makedirs(newSnFolderPath, exist_ok=True)
#取出需要的SN
for SN in columData:
    for filename in allfilename:
        if SN in filename:   
            #生成找到的檔案路徑
            filePath=os.path.join(folder_path, filename)
            #放置複製的檔案路徑
            newdestination_path = os.path.join(newSnFolderPath, filename)            
            if os.path.isfile(filePath): 
                shutil.copy2(filePath,newdestination_path)
            elif os.path.isdir(filePath):   
                shutil.copytree(filePath,newdestination_path)          
            if '0x' in columTitle:
                search_field=columTitle
                row=1
                if os.path.isfile(filePath):
                    file= open(filePath, 'r') 
                    content = file.read().lower()        
                # 判断是否存在所需字段
                    if search_field.lower() in content:
                        workbook = xlwt.Workbook()
                        sheet = workbook.add_sheet('Data')   
                # 设置Excel标题行
                        sheet.write(0, 0, 'File Name')
                        sheet.write(0, 1, 'Context')                       
                    # 写入Excel
                        sheet.write(row, 0, filename)
                        start_index = content.index(search_field.lower())
                        end_index = start_index + len(search_field.lower())
                        filenameSN=filename.split('_')
                        errlist.append(filenameSN[0])             
                        if end_index+textlenth< len(content):
                            end_index += textlenth                       
                        else:
                            end_index 
                        sheet.write(row, 1, content[start_index:end_index])
                        row += 1
                        excelName=columTitle+'.xls'
                        excelPath=os.path.join(destination_path,excelName)
                        workbook.save(excelPath) 
errlist=list(set(errlist))                       
#复制文件到目标文件夹 
errfolder_name = str(columTitle+'_'+str(int(time.time())))
newErrFolderPath = os.path.join(destination_path, errfolder_name)
os.makedirs(newErrFolderPath, exist_ok=True)         
for errFileName in allfilename:
    for ercode in errlist:                             
        if ercode in errFileName:
                newErrdestination_path = os.path.join(newErrFolderPath, errFileName)
                errfilepath=os.path.join(folder_path,errFileName)
                if os.path.isfile(errfilepath): 
                    shutil.copy2(errfilepath,newErrdestination_path)
                elif os.path.isdir(errfilepath):   
                    shutil.copytree(errfilepath,newErrdestination_path)
print('已經完成LOG尋找')
                        
