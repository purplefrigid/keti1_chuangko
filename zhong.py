import pandas as pd  
import os
# 读取 Excel 文件  
df = pd.read_excel('data1.xlsx', header=None)  

# pro_path = "C:\\Users\\zgnhz\\Desktop\\新建文件夹 (2)\\"
# data_path=pro_path+"kuandu_data"
data_path="./kuandu_data"
files = os.listdir(data_path)
for file in files:
# 遍历每一行
    bh= file.split(".")[0].split("_")[1]   
    file_path = data_path+"\\"+file
    fileHandler  =  open  (file_path,  "r")	
    listOfLines  =  fileHandler.readlines()
    for index, row in df.iterrows():  
        # b=18
        if row[0] ==int(bh):
            for  num ,line in  enumerate(listOfLines):
                a=line.strip().split()[1]
                pl=line.strip().split()[0]
                for ls, line in df.iteritems():  
                    if line[0] ==float(pl):
                        break
                df.at[index, ls] = float(a)
                # b=b+1
# 保存修改后的 DataFrame 到新的 Excel 文件  
df.to_excel('modified_file.xlsx', index=False, header=False)