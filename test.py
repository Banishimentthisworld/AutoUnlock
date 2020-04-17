import os
import win32con
import win32api
import time
# python批量更换后缀名
import os
import sys
import win32gui
import codecs
import pyperclip

def CtrlA():
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(65, win32api.MapVirtualKey(65, 0), 0, 0)
    time.sleep(0.2)
    win32api.keybd_event(65, win32api.MapVirtualKey(65, 0), win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), win32con.KEYEVENTF_KEYUP, 0)

def CtrlC():
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(67, win32api.MapVirtualKey(67, 0), 0, 0)
    time.sleep(0.2)
    win32api.keybd_event(67, win32api.MapVirtualKey(67, 0), win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), win32con.KEYEVENTF_KEYUP, 0)

def CtrlV():
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(86, win32api.MapVirtualKey(86, 0), 0, 0)
    time.sleep(0.2)
    win32api.keybd_event(86, win32api.MapVirtualKey(86, 0), win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), win32con.KEYEVENTF_KEYUP, 0)

def CtrlS():
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), 0, 0)
    time.sleep(0.1)
    win32api.keybd_event(83, win32api.MapVirtualKey(83, 0), 0, 0)
    time.sleep(0.2)
    win32api.keybd_event(83, win32api.MapVirtualKey(83, 0), win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    win32api.keybd_event(17, win32api.MapVirtualKey(17, 0), win32con.KEYEVENTF_KEYUP, 0)

def select(root, filename):
    os.startfile(os.path.join(root, filename))
    time.sleep(0.5)
    # 获取窗口信息
    titlename = filename +" - 记事本"
    # 根据titlename信息查找窗口
    hwnd = win32gui.FindWindow(0, titlename)
    win32gui.SetForegroundWindow(hwnd)

def close(filename):
    # 获取窗口信息
    titlename = filename + " - 记事本"
    # 根据titlename信息查找窗口
    hwnd = win32gui.FindWindow(0, titlename)
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

global fileslist
fileslist  = []

# 列出当前目录下所有的文件
for root, dirs, files in os.walk("E:\MM"):

    # root 表示当前正在访问的文件夹路径
    # dirs 表示该文件夹下的子目录名list
    # files 表示该文件夹下的文件list

    # 遍历文件
    for f in files:

        print(os.path.join(root, f))
        portion = os.path.splitext(f)
        # 如果后缀是.dat

        if portion[1] == ".cpp" or portion[1] == ".h" or portion[1] == ".cs":
            #把原文件后缀名改为 txt
            newName = portion[0] + ".txt"
            os.rename(os.path.join(root, f),os.path.join(root, newName))
            desktop_path = root  # 新创建的txt文件的存放路径
            full_path = desktop_path + "\Copy" + newName  # 也可以创建一个.doc的word文档
            print(full_path)
            file = open(full_path, 'w')
            file.close()

            # 复制原文件信息
            if int(os.path.getsize(os.path.os.path.join(root, newName))) > 0:
                select(root, newName)
                CtrlA()
                time.sleep(0.1)
                CtrlC()
                time.sleep(0.1)
                close(newName)

                fh = codecs.open(desktop_path + "\Copy" + newName, "w", "utf-16")
                text = pyperclip.paste()
                fh.write(text)
                fh.close()
            os.rename(os.path.join(root, "Copy" + newName), os.path.join(root, f))
            os.remove(os.path.join(root, newName))
            # select(root, "Copy" + newName)
            # CtrlV()
            # time.sleep(0.5)
            # try:
            #     CtrlS()
            #     close("Copy" + newName)
            #     os.rename(os.path.join(root, "Copy" + newName), os.path.join(root, f))
            #     os.remove(os.path.join(root, newName))
            # except:
            #     print("编码格式不对")
            #
            #     # 保存unicode格式
            #     fh = codecs.open("xxx.txt", "w", "utf-16")
            #     fh.close()


    # 遍历文件夹
    # for d in dirs:
    #     print(os.path.join(root, d))



