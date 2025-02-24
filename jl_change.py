import pandas as pd 
import math
import numpy as np 
def zhanbi(k,thick):
    a=math.sqrt(3)*k
    b=math.sqrt(3)*k-2*thick
    return 1-np.square(b/a)

def calculate(ea, e0, g) :
    # return ea * ((2 - g) * e0 + g * ea) / (g * e0 + (2 - g) * ea) #径向
     return g*ea+(1-g)*e0#轴向

file_path1 = 'data1.xlsx'  # 替换为你的文件路径
file_path2 = 'jl.xlsx'  # 替换为你的文件路径  
data1 = pd.read_excel(file_path1)
data2 = pd.read_excel(file_path2) 
e0 = 1.0006
SEQ =(5,10,15)
for num in range(data1.shape[0]):
    for seq in SEQ:
        thick = data1.iloc[num, seq]
        kongjing =data1.iloc[num, seq-4]
        # thick = 0.05
        # kongjing=4
        g= zhanbi(kongjing,thick)
        target_file_path='./jl_change/'+str(data1.iloc[num, seq-1])+'_'+str(kongjing,)+'_'+str(data1.iloc[num, seq])+'.txt'
        for num1,row in enumerate(range(data2.shape[0])):
            ea=complex(data2.iloc[row, 1], data2.iloc[row, 2])
            epsilon_c=calculate(ea,e0,g)
            tangent=epsilon_c.imag/epsilon_c.real
            content = str(data2.iloc[row, 0])+' '+str(epsilon_c.real)+' '+str(tangent)+'\n'
            if num1==0:
                with open(target_file_path, 'w', encoding='utf-8') as target_file:  
                    target_file.write(content)  # 写入源文件的内容         
            else:       
                with open(target_file_path, 'a', encoding='utf-8') as target_file:  
                    target_file.write(content)  # 写入源文件的内容
        



