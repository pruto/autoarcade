import win32gui
import pyautogui
import numpy
import cv2
import win32api
import win32con
from time import sleep
from math import floor

def compimg(img1, img2):
    subimg1 = cv2.split(img1)
    subimg2 = cv2.split(img2)
    ret = 0
    for ch1, ch2 in zip(subimg1, subimg2):
        ret += compchannel(ch1, ch2)
    ret = ret / 3
    return ret

def compchannel(ch1, ch2):
    hist1 = cv2.calcHist([ch1], [0], None, [256], [0.0, 255.0])
    hist2 = cv2.calcHist([ch2], [0], None, [256], [0.0, 255.0])
    degree = 0
    for i in range(len(hist1)):
        if hist1[i] != hist2[i]:
            degree = degree + (1 - abs(hist1[i] - hist2[i]) / max(hist1[i], hist2[i]))
        else:
            degree = degree + 1
    degree = degree / len(hist1)
    return degree

def click(posx, posy):
    sleep(.05)
    pyautogui.moveTo(posx, posy, duration=.0)
    sleep(.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    sleep(.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def testabort(hnd):
    if hnd != win32gui.GetForegroundWindow():
        exit()

def main():
    hnd = win32gui.FindWindow(None, '地下城与勇士：创新世纪')
    win32gui.SetForegroundWindow(hnd)

    winleft, wintop, winright, winbottom = win32gui.GetWindowRect(hnd)
    winw = winright-winleft
    winh = winbottom-wintop
    wincx = winleft+.5*winw
    wincy = wintop+.5*winh
    piecew = 40
    pieceh = 30
    jigsawcx = wincx+5
    jigsawcy = wincy-74
    pendcx = int(wincx-259)
    pendcy = int(wincy+191)

    pyautogui.moveTo(x=jigsawcx, y=jigsawcy)

    # 生成192个拼图块参考图列表
    refli = []
    imgrefer = cv2.imread('refer.png')
    for i in range(192):
        refpce = imgrefer[floor(i/16)*30+13:floor(i/16)*30+17, i%16*40+18:i%16*40+22, :]
        refli.append(refpce)

    # 循环截图对比放置拼图
    while True:
        sleep(.15)
        simli = []
        cellx = -1
        celly = -1
        pendpce = pyautogui.screenshot(region=[int(pendcx-.5*piecew+18), int(pendcy-.5*pieceh+13), 4, 4])
        pendpce = cv2.cvtColor(numpy.asarray(pendpce), cv2.COLOR_RGB2BGR)
        for i in range(192):
            simval = compimg(pendpce, refli[i])
            if (simval == 1):
                cellx = i%16
                celly = floor(i/16)
                break

        testabort(hnd)
        if cellx == -1 or celly == -1:
            click(wincx, wincy+240)
            continue
        posx = int(jigsawcx+(cellx-7.5)*piecew)
        posy = int(jigsawcy+(celly-5.5)*pieceh)

        testabort(hnd)
        click(pendcx, pendcy)
        testabort(hnd)
        click(posx, posy)

if __name__ == '__main__':
    main()
