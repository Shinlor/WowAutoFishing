#!/usr/bin/python
# -*- coding: UTF-8 -*-
import win32com.client
import win32file, win32api
import os
import time
import win32con
dm = win32com.client.Dispatch('dm.dmsoft')
#current version
print(dm.Ver())
X=0
Y=0
dy=0
s=0
#hwnd = dm.GetMousePointWindow()
dm_ret = dm.BindWindow(395948, "dx2", "normal", "normal", 1) #窗口句柄需每次手动找到
time.sleep(3)
win32api.keybd_event(73,0,0,0)  #游戏中钓鱼快捷键键位码
time.sleep(0.1)
win32api.keybd_event(73,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
for b in range(2000):

    for a in range(200):
     
        YS=dm.FindColor(400,200,1200,500,"FFFFFF",0.7,0,X,Y)  #调用找图函数找图，作为元组返回,默认当前显示界面中找
        
        X=YS[1] #提取元组中X坐标
        Y=YS[2] #提取元组中Y坐标
        time.sleep(0.1)
        s=s+1
        if (X>0) and (Y>0):
            print (X,Y,YS)
            dm.MoveTo (X,Y)  #移动鼠标
            time.sleep(1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,X,Y,0,0)
            time.sleep(0.2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,X,Y,0,0)
            time.sleep(0.2)
            dm.MoveTo (300,300)
            time.sleep(2)
            win32api.keybd_event(73,0,0,0)  #游戏中钓鱼快捷键键位码
            time.sleep(0.1)
            win32api.keybd_event(73,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
            dy=1
            s=0
        if (dy == 0)and (s>150):
            win32api.keybd_event(73,0,0,0)  #游戏中钓鱼快捷键键位码
            time.sleep(0.1)
            win32api.keybd_event(73,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
            s=0
            time.sleep(1)
        dy=0
        
print ("b=",b)     
dm_ret=dm.UnBindWindow()