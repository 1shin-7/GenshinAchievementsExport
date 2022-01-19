import cv2
import keyboard
import numpy as np
import pyautogui
import pytesseract
import win32com.client
import time
from PIL import Image
def isChinese(word):
    for ch in word:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False
def isEnglish(word):
    for ch in word:
        if 'a'<=ch<='z' or 'A'<=ch<='Z':
            return True
    return False

def format_str(content):
 content_str = ''
 for i in content:
    if isChinese(i) or isEnglish(i):
        content_str = content_str+i
 return content_str
def cross(bbox,frameworks):
    for fw in frameworks:
        center_fw=(fw[0]+fw[2]//2,fw[1]+fw[3]//2)
        center_bbox=(bbox[0]+bbox[2]//2,bbox[1]+bbox[3]//2)
        if abs(center_fw[0]-center_bbox[0])<(fw[2]+bbox[2])//2 and abs(center_fw[1]-center_bbox[1])<(fw[3]+bbox[3])//2:
            return True
    return False

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
                if cross(bbox,frameworks):
                    continue
                frameworks.append(bbox)
    else:
        for contour in contours:
            bbox=cv2.boundingRect(contour)
            if 120<=bbox[3]<=150 and (1000<=bbox[2]<=1500):
                if cross(bbox,frameworks):
                    continue
                frameworks.append(bbox)
    frameworks.sort(key=lambda x:-x[1])
    return frameworks

def get_page(name):
    speak.Speak('开始整理'+name)
    text_list=[]
    RUN=True
    while RUN:
        tmp_list=[]
        pyautogui.moveTo(10,10)
        image=pyautogui.screenshot()
        image=cv2.cvtColor(np.asarray(image),cv2.COLOR_RGB2BGR)
        frameworks=get_rects(image,'right')
        for fw in frameworks:
            image_text=image[fw[1]:fw[1]+fw[3],fw[0]:fw[0]+fw[2]]
            image_text= image_text[:,120:700]
            text=pytesseract.image_to_string(Image.fromarray(cv2.cvtColor(image_text,cv2.COLOR_BGR2RGB)),lang='chi_sim')
            text=text.strip().split('\n')[0]
            text=format_str(text)
            if not text in text_list:
                if text!='达成进度':
                    tmp_list.append(text)              
                    pyautogui.doubleClick(fw[0]+fw[2]//2,fw[1]+fw[3]//2)
            else:
                
                RUN=False
                break
        tmp_list.reverse()
        for x in tmp_list:
            text_list.append(x)
        if RUN:            
            lastbox=frameworks[0]
            xpos=lastbox[0]+lastbox[2]//2+100
            ypos=lastbox[1]+lastbox[3]//2
            pyautogui.moveTo(xpos,ypos)
            pyautogui.dragTo(xpos,10,duration=1.5,tween=pyautogui.easeOutQuad,button='left')
            pyautogui.click(xpos,200)
        else:
            print(text_list)
            speak.Speak(name+'整理完毕')
            with open('result/'+name+'.txt','w',encoding='utf-8') as f:
                for x in text_list:
                    f.write(x+'\n')

#def get_indexs():
#    speak.Speak('查找列表')
#    text_list=[]
#    RUN=True
#    while RUN:        
#        pyautogui.moveTo(10,10)
#        image=pyautogui.screenshot()
#        image=cv2.cvtColor(np.asarray(image),cv2.COLOR_RGB2BGR)
#        frameworks=get_rects(image,'left')
#        for fw in frameworks:                       
#            #cv2.imshow('img',image_text)
#            #cv2.waitKey()
#            pyautogui.doubleClick(fw[0]+fw[2]//2,fw[1]+fw[3]//2)
#            time.sleep(1)
#            image_text=image[fw[1]:fw[1]+fw[3],fw[0]+60:fw[0]+fw[2]]
#            text=pytesseract.image_to_string(Image.fromarray(cv2.cvtColor(image_text,cv2.COLOR_BGR2RGB)),lang='chi_sim')
#            text=text.strip().split('\n')[0]
#            text=format_str(text)
#            if not text in text_list:
#                text_list.append(text)
#                cv2.imencode('.png',image_text)[1].tofile('result/'+text+'.png')
#                #get_page(text)
#            else:
#                print(text_list)
#                speak.Speak('列表查找完毕')
#                RUN=False
#                break   
#        if RUN:
#            lastbox=frameworks[0]
#            xpos=lastbox[0]+lastbox[2]//2+100
#            ypos=lastbox[1]+lastbox[3]//2
#            pyautogui.moveTo(xpos,ypos)
#            pyautogui.dragTo(xpos,10,duration=1.5,tween=pyautogui.easeOutQuad,button='left')
#            time.sleep(1)

speak = win32com.client.Dispatch('SAPI.SPVOICE')
keyboard.wait('r')
speak.Speak('程序启动')
get_page('天地万象')

