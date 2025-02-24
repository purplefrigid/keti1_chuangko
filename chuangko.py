import pygetwindow as gw  
import pyautogui  
import numpy as np  
import cv2  
from pynput import mouse  
import time  
import pandas as pd  
from random import randrange
import sys                                                                  
import signal
from pynput.mouse import  Button,Controller 
import random   
import os


import smtplib
from email.mime.text import MIMEText
from email.header import Header
# 创建 SMTP 对象
# smtp = smtplib.SMTP()
# # 发件人邮箱地址
# sendAddress = '2907402968@qq.com'
# # 发件人授权码
# password = 'tncftocstczedhba'
# 连接服务器
# server = smtplib.SMTP_SSL('smtp.qq.com', 465)
# # 构造MIMEText对象，参数为：正文，MIME的subtype，编码方式
# message = MIMEText('出问题了', 'plain', 'utf-8')
# message1 = MIMEText('开始运行', 'plain', 'utf-8')
# message['From'] = Header("me <2907402968@qq.com>")  # 发件人的昵称 用英文不报错
# message['To'] = Header("Anson <2907402968@qq.com>")  # 收件人的昵称  用英文不报错
# message['Subject'] = Header('出问题了', 'utf-8')  # 定义主题内容
# message1['From'] = Header("me <2907402968@qq.com>")  # 发件人的昵称 用英文不报错
# message1['To'] = Header("Anson <2907402968@qq.com>")  # 收件人的昵称  用英文不报错
# message1['Subject'] = Header('开始运行', 'utf-8')  # 定义主题内容

def quit(signum, frame):
    sys.exit()

# 全局变量存储窗口标题  
selected_window_title = None 
x=15.2
FREQ = np.arange(5.05, 18, 0.0875*2)  
file_path1 = 'data10.xlsx'  # 替换为你的文件路径
file_path2 = 'jl.xlsx'  # 替换为你的文件路径  
file_path3 = 'jl.xlsx'  # 替换为你的文件路径  
file_path4 = 'jl.xlsx'  # 替换为你的文件路径  
data1 = pd.read_excel(file_path1)
data2 = pd.read_excel(file_path2) 
data3 = pd.read_excel(file_path3)
data4 = pd.read_excel(file_path4)
source_file_path = '等效电磁参数.txt'  # 替换为你的源文件路径  
# 指定文本框的相对位置和要输入的文本  
relative_x = 174  # 浆料厚度  
relative_y = 234   
relative_x21 = 143  # 频率  
relative_y21 = 44    
relative_x22 = 200  # 介电实部  
relative_y22 = 139   
relative_x23 = 277  # 介电虚部1
relative_y23 = 144  
relative_xc = 324  # 计算
relative_yc = 236  
x_x=136
x_y=137
cfjc = 0
save_I=0
save_R=0

text1='BEGIN_INPUT: VERSION=TFD001'+'\n'+'LENGTH_UNIT mm'+'\n'+'NUM_MESH 1'+'\n'
text2='RCSCONFIG 3 1'+'\n'+'1 90.0 90.0'+'\n'+'1  0.0  0.0'+'\n'+'1 90.0 90.0'+'\n'+'1  0.0  0.0'+'\n'+'MATERIAL list'+'\n'+'3'+'\n'
text3='MOM_PARAM 1.0 1 1'+'\n'+'FMM_PARAM 0.0 -1'+'\n'+'SOLVER 4  400 0.0050  precond 1'+'\n'+'END_INPUT'

mouse_c = Controller()  

def bring_window_to_front(window):  
    window.activate()  # 激活窗口  
    window.always_on_top = True  # 设置窗口始终在最前面  
    #print(f"{window.title} is now always on top.") 
def on_click(x, y, button, pressed):  
    global selected_window_title  
    if pressed and button == mouse.Button.left:  
        # 获取鼠标点击位置的窗口  
        window = gw.getWindowsAt(x, y)  
        if window:  
            selected_window_title = window[0].title 
            selected_window_title= 'EquivalentEMparameters'
            print(f"已选择窗口: {selected_window_title}")  
            return False  # 停止监听  

def type_in_textbox(relative_x, relative_y, text):  
    global selected_window_title  
    if not selected_window_title:  
        print("请先选择一个窗口。")
        return  
    # 获取指定窗口  
    try:  
        window = gw.getWindowsWithTitle(selected_window_title)[0]  
    except IndexError:  
        print(f"未找到标题为 '{selected_window_title}' 的窗口。")  
        return  
    bring_window_to_front(window)
    # 获取窗口的绝对位置  
    window_x, window_y = window.topleft  
    #print(f"已绑定句柄位置: ({window_x}, {window_y})")  
    # 计算文本框的绝对位置  
    textbox_x = window_x + relative_x  
    textbox_y = window_y + relative_y  
    # 点击指定位置的文本框  
    #pyautogui.click(textbox_x, textbox_y)  
    #print(f"已点击文本框位置: ({textbox_x}, {textbox_y})")  

    # 输入文本  
    #pyautogui.hotkey('ctrl', 'a')  # Windows/Linux  
    #pyautogui.doubleClick(textbox_x, textbox_y)  # 双击以选中内容
    # current_position = mouse_c.position  
    mouse_c.position = (textbox_x, textbox_y)  
    mouse_c.click(Button.left, 2)  
    # mouse_c.position = current_position 
    # 删除选中的文本
      
    pyautogui.press('backspace')  # 或者使用 'delete' 键 
    pyautogui.press('backspace')
    time.sleep(0.1)
    if(mouse_c.position==(textbox_x, textbox_y)):
        pyautogui.typewrite(text, interval=0.1)  # interval 设置每个字符之间的间隔
    else:
        print("stop")
        time.sleep(5) 
        type_in_textbox(relative_x, relative_y, text)
    # time.sleep(0.1)

def mouse_click(relative_x, relative_y):  
    global selected_window_title  
    if not selected_window_title:  
        print("请先选择一个窗口。")
        return  
    # 获取指定窗口  
    try:  
        window = gw.getWindowsWithTitle(selected_window_title)[0]  
    except IndexError:  
        print(f"未找到标题为 '{selected_window_title}' 的窗口。")  
        return  
    bring_window_to_front(window)
    # 获取窗口的绝对位置  
    window_x, window_y = window.topleft  
    #print(f"已绑定句柄位置: ({window_x}, {window_y})")  
    # 计算文本框的绝对位置  
    textbox_x = window_x + relative_x  
    textbox_y = window_y + relative_y  
    mouse_c.position = (textbox_x, textbox_y)  

    mouse_c.click(Button.left, 1)  

def write_to_file(relative_x, relative_y, target_file_path,num):  
    global selected_window_title 
    global save_I
    global save_R 
    global cfjc
    if not selected_window_title:  
        print("请先选择一个窗口。")  
        return  
    # 获取指定窗口  
    try:  
        window = gw.getWindowsWithTitle(selected_window_title)[0]  
    except IndexError:  
        print(f"未找到标题为 '{selected_window_title}' 的窗口。")  
        return  
    bring_window_to_front(window)
    # 获取窗口的绝对位置  
    window_x, window_y = window.topleft  
    #print(f"已绑定句柄位置: ({window_x}, {window_y})")  
    # 计算文本框的绝对位置  
    textbox_x = window_x + relative_x  
    textbox_y = window_y + relative_y  
    # 点击指定位置的文本框  
    #pyautogui.click(textbox_x, textbox_y)  
    current_position = mouse_c.position  
    mouse_c.position = (textbox_x, textbox_y)  
    mouse_c.click(Button.left, 1)  
    mouse_c.position = current_position 
    time.sleep(0.2) 
    # 读取源文件的内容  
    with open(source_file_path, 'r', encoding='utf-8') as source_file:  
        line = source_file.readline().strip()  # 读取第一行并去除首尾空白字符  
        # 按照逗号分割数据  
        data = line.split(',')      
        # 提取前后的数据  
        value_before_comma = data[0].strip()  # 逗号前的数据  
        value_after_comma = data[1].strip()   # 逗号后的数据  
    # 将内容写入目标文件的末尾，并添加换行符  
    time.sleep(0.1) 
    with open(target_file_path, 'a', encoding='utf-8') as target_file:  
        content=str(num+1)+' 4 '+'('+value_before_comma+','+value_after_comma+')'+'\n'
        target_file.write(content)
        print(content)
    if save_R ==value_before_comma and save_I == value_after_comma:
        cfjc=cfjc+1
    else:
        cfjc = 0
    save_R=value_before_comma
    save_I=value_after_comma
    return cfjc


if __name__ == "__main__":  
    #等待用户选择窗口 
    jc=0 
    print("请点击要绑定的窗口...")  
    listener = mouse.Listener(on_click=on_click)  
    listener.start()  # 启动监听器  

    # 等待直到选择窗口 
    while selected_window_title is None:  
        time.sleep(0.1)  # 等待选择窗口  
    # 停止监听  
    listener.stop()  
    # 等待一段时间，以便用户准备  
    print("请在 5 秒内确保目标窗口处于前台...")  
    time.sleep(3)
    signal.signal(signal.SIGINT, quit)                                
    signal.signal(signal.SIGTERM, quit) 
    # loginResult = server.login(sendAddress, password)
    # print(loginResult)
    # server.sendmail('2907402968@qq.com', '2907402968@qq.com', message1.as_string())

    for i in range(len(FREQ)):  
        FREQ[i] = round(FREQ[i], 4)  # 将值保留三位小数   
    for num in range(data1.shape[0]): 
        # if jc >= 3:
        #     break 
        for fre in FREQ: 
            if jc >= 3:
                mouse_click(x_x,x_y)
                jc=0
                # break
            target_file_path='./save1/zjh_'+str(round(data1.iloc[num, 0]))+'_'+str(fre)+'Ghz.tfd_input' 
            text4='zjh_'+str(round(data1.iloc[num, 0]))+'_'+str(fre)+'Ghz.facet'+'\n'
            text5='FREQUENCY 1 '+str(fre)+' '+str(fre)+'\n' 

            material = [data1.iloc[num, 4],data1.iloc[num, 9],data1.iloc[num, 14]]; 
    
            thickness = [data1.iloc[num, 5],data1.iloc[num, 10],data1.iloc[num, 15]]  

            for index, materialn in enumerate(material): 

                if materialn == 1:
                    for row in range(data2.shape[0]):     
                        if  fre ==  data2.iloc[row, 0]:
                            break
                    value21 = data2.iloc[row, 0]
                    value22 = data2.iloc[row, 1]
                    value23 = data2.iloc[row, 3]
                    # text_to_type21 = str(value21) 
                    # text_to_type22 = str(value22) 
                    # text_to_type23 = str(value23) 
                elif materialn == 2:
                    for row in range(data3.shape[0]):     
                        if  fre ==  data3.iloc[row, 0]:
                            break
                    value21 = data3.iloc[row, 0]
                    value22 = data3.iloc[row, 1]
                    value23 = data3.iloc[row, 3]
                    # text_to_type21 = str(value21) 
                    # text_to_type22 = str(value22) 
                    # text_to_type23 = str(value23)
                elif materialn == 3:
                    for row in range(data4.shape[0]):     
                        if  fre ==  data4.iloc[row, 0]:
                            break
                    value21 = data4.iloc[row, 0]
                    value22 = data4.iloc[row, 1]
                    value23 = data4.iloc[row, 3]
                    # text_to_type21 = str(value21) 
                    # text_to_type22 = str(value22) 
                    # text_to_type23 = str(value23) 
                text_to_type21 = str(value21) 
                text_to_type22 = str(round(value22,2)) 
                text_to_type23 = str(round(value23,2)) 
                if index ==0:
                    with open(target_file_path, 'a', encoding='utf-8') as target_file:  
                        target_file.write(text1)  # 写入源文件的内容
                        target_file.write(text4)  # 写入源文件的内容
                        target_file.write(text5)  # 写入源文件的内容
                        target_file.write(text2)  # 写入源文件的内容
                    type_in_textbox(relative_x21, relative_y21, text_to_type21) 
                    type_in_textbox(relative_x22, relative_y22, text_to_type22) 
                    type_in_textbox(relative_x23, relative_y23, text_to_type23)
                time.sleep(0.1)

                type_in_textbox(relative_x, relative_y, str(thickness[index]))  
                jc = write_to_file(relative_xc,relative_yc,target_file_path,index)
                if jc == 3:
                    # server.sendmail('2907402968@qq.com', '2907402968@qq.com', message.as_string())
                    print(text4)
                    with open('./error.txt', 'a', encoding='utf-8') as target_file:  
                        target_file.write(text4)  # 写入源文件的内容
                    break
                #print(jc)
                if index==2:
                    with open(target_file_path, 'a', encoding='utf-8') as target_file:  
                        target_file.write(text3)  # 写入源文件的内容
