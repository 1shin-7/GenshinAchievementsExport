
from itertools import filterfalse
from re import I
import cv2
import keyboard
import numpy as np
from numpy.lib.shape_base import take_along_axis
import pyautogui
from paddleocr import PaddleOCR
import win32com.client
import time
import openpyxl

def cross(bbox,frameworks):
    for i,fw in enumerate(frameworks):
        center_fw=(fw[0]+fw[2]//2,fw[1]+fw[3]//2)
        center_bbox=(bbox[0]+bbox[2]//2,bbox[1]+bbox[3]//2)
        if abs(center_fw[0]-center_bbox[0])<(fw[2]+bbox[2])//2 and abs(center_fw[1]-center_bbox[1])<(fw[3]+bbox[3])//2:
            return i
    return -1

def get_rects(image,mode='left'):
    img=cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)
    thresh_min,thresh_max=(50,100)
    img=cv2.Canny(img,thresh_min,thresh_max)
    contours,hierachy=cv2.findContours(img,cv2.CHAIN_APPROX_SIMPLE,cv2.RETR_CCOMP)
    frameworks=[]
    if mode=='left': 
        for contour in contours:
            bbox=cv2.boundingRect(contour)
            if 100<=bbox[3]<=120 and (580<=bbox[2]<=700):
                if cross(bbox,frameworks)!=-1:
                    continue
                frameworks.append(bbox)
    else:
        
        for contour in contours:
            bbox=cv2.boundingRect(contour)
            if 110<=bbox[3]<=150 and (1000<=bbox[2]<=1500):
                if cross(bbox,frameworks)!=-1:
                    continue
                frameworks.append(bbox)
    frameworks.sort(key=lambda x:x[1])
    return frameworks

def get_page(name):
    speak.Speak('开始整理'+name)
    text_list=[]
    info_list=[]
    complete_list=[]
    if name in workbook.sheetnames:
        del workbook[name]
    worksheet=workbook.create_sheet(name)
    worksheet.cell(1,1).value='序号'
    worksheet.cell(1,2).value='名称'
    worksheet.cell(1,3).value='内容'
    worksheet.cell(1,4).value='完成日期或比例'
    worksheet.cell(1,5).value='是否达成'
    RUN=True
    while RUN:
        pyautogui.moveTo(10,10)
        image=pyautogui.screenshot()
        image=cv2.cvtColor(np.asarray(image),cv2.COLOR_RGB2BGR)
        frameworks=get_rects(image,'right')
        result=ocr.ocr(image)
        tmp_list=['']*len(frameworks)
        info=['']*len(frameworks)
        complete=['']*len(frameworks)
        for line in result:
            text=line[1][0]
            if text[0]=='大' or text[0]=='达':
                continue
            bbox=(line[0][0][0],line[0][0][1],line[0][2][0]-line[0][0][0],line[0][2][1]-line[0][0][1])
            index=cross(bbox,frameworks)
            if index!=-1:
                if tmp_list[index]=='':
                    tmp_list[index]=text
                elif len(text.split('/'))>1:
                    complete[index]=text
                else:
                    info[index]=text
        all_seen=True 
        for x,y,z in zip(tmp_list,info,complete):
            if x=='' or '%' in x:
                continue
            if not x in text_list:
                text_list.append(x)
                info_list.append(y)
                complete_list.append(z)
                all_seen=False
        if all_seen:
            #print(text_list)
            #print(info_list)
            #print(complete_list)
            speak.Speak(name+'整理完毕')
            for i,x in enumerate(text_list):
                worksheet.cell(i+2,1).value=i 
                worksheet.cell(i+2,2).value=x 
                worksheet.cell(i+2,3).value=info_list[i]
                worksheet.cell(i+2,4).value=complete_list[i]
                worksheet.cell(i+2,5).value='yes' if len(complete_list[i].split('/'))>2 else 'no'
            break                   
        lastbox=frameworks[-1]
        xpos=lastbox[0]+lastbox[2]//2+100
        ypos=lastbox[1]+lastbox[3]//2
        pyautogui.moveTo(xpos,ypos)
        pyautogui.dragTo(xpos,10,duration=1.5,tween=pyautogui.easeOutQuad,button='left')
        pyautogui.click(xpos,200)
            

def get_indexs():
    speak.Speak('查找列表')
    text_list=[]
    RUN=True
    while RUN:        
        pyautogui.moveTo(10,10)
        image=pyautogui.screenshot()
        image=cv2.cvtColor(np.asarray(image),cv2.COLOR_RGB2BGR)
        frameworks=get_rects(image,'left')
        tmp_list=['']*len(frameworks)
        up=80
        down=1080
        left=60
        right=600
        result=ocr.ocr(image[up:down,left:right])
        for line in result:
            text=line[1][0]
            bbox=(line[0][0][0]+left,line[0][0][1]+up,line[0][2][0]-line[0][0][0],line[0][2][1]-line[0][0][1])
            index=cross(bbox,frameworks)
            if index!=-1:
                if tmp_list[index]=='':
                    tmp_list[index]=text
        all_seen=True 
        for i,x in enumerate(tmp_list):
            fw=frameworks[i]
            if not x in text_list:
                text_list.append(x)
                pyautogui.doubleClick(fw[0]+fw[2]//2,fw[1]+fw[3]//2)
                time.sleep(1)
                all_seen=False 
                get_page(x)
                time.sleep(1)
        if all_seen:
            #print(text_list)
            speak.Speak('列表查找完毕') 
            break
        
        lastbox=frameworks[0]
        xpos=lastbox[0]+lastbox[2]//2+100
        ypos=lastbox[1]+lastbox[3]//2
        pyautogui.moveTo(xpos,ypos)
        pyautogui.dragTo(xpos,10,duration=1.5,tween=pyautogui.easeOutQuad,button='left')
        time.sleep(1)
if __name__=='__main__':
    speak = win32com.client.Dispatch('SAPI.SPVOICE')
    ocr= PaddleOCR(lang='ch')
    workbook=openpyxl.Workbook()
    keyboard.wait('r')
    speak.Speak('程序启动') 
    get_indexs()
    del workbook['Sheet']
    workbook.save('test.xlsx')

