import ctypes
import pyautogui
import numpy as np
import time
import copy
import threading
import win32api
import win32con
import win32gui
from pymem import Pymem
from pymem.ptypes import RemotePointer
import cv2
import pymem
import multiprocessing as mp
from multiprocessing import Queue


regions = [[145, 444, 183, 464], [226, 418, 269, 438], [306, 406, 348, 428], [366, 402, 428, 426], [
    447, 401, 497, 422], [519, 406, 565, 429], [595, 415, 642, 443], [674, 440, 717, 463]]
colors = ['b3f2c9-000040,9|0|a2ffb1-000040',
          'ffc0ce-000040,12|1|ee738c-000040', 'a9aa7e-000040,-4|-1|f2f4af-000040', '7297fe-000030,15|-1|465eab-000030', 'ffc18a-000040,8|0|f8cc80-000040', 'fc8eff-000040,7|0|fd8cff-000040', '62d3dc-000040,8|0|63eae9-000040', 'ac71e7-000040,5|-1|df92ff-000040']
snowmans = [[82.16, 114.69], [-75.31, 120.89], [75.76, 45], [-77.7, 50.92]]


left, top, right, bottom = 0, 0, 0, 0
play_state = 0  # 是否在舞台上


def setWindowPos(hwnd):
    '''
    @@ 设定窗口位置
    '''
    win32gui.MoveWindow(hwnd, 0, 0, 936, 651, True)


def getWindowRect(hwnd):
    '''
    @@ 获取窗口坐标
    '''
    global left, top, right, bottom
    try:
        f = ctypes.windll.dwmapi.DwmGetWindowAttribute
    except WindowsError:
        f = None
    if f:
        rect = ctypes.wintypes.RECT()
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        f(ctypes.wintypes.HWND(hwnd),
          ctypes.wintypes.DWORD(DWMWA_EXTENDED_FRAME_BOUNDS),
          ctypes.byref(rect),
          ctypes.sizeof(rect)
          )
        left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom
        return rect.left, rect.top, rect.right, rect.bottom


def windowCapture(mode, x1=0, y1=0, x2=0, y2=0, hwnd=0):
    '''
    @@ 截图
    '''
    if mode == 1:
        x1, y1, x2, y2 = getWindowRect(hwnd)
        # width = x2 - x1
        # height = y2 - y1
        src = pyautogui.screenshot(region=[x1, y1, x2-x1, y2-y1])
        src = np.array(src)
        img = cv2.cvtColor(np.asarray(src), cv2.COLOR_RGB2BGR)
        cv2.imwrite('1.jpg', img)
        return src
    elif mode == 2:
        src = pyautogui.screenshot(region=[x1, y1, x2-x1, y2-y1])
        src = np.array(src)
        return src


def findMultColor(src, des):
    '''
    @@ 多点找色
    @  src: RBG(pyautogui)
    @  des: from Damo
    @  return: Find-(x,y) / NotFind- -1
    '''
    st = time.time()
    rgby = []
    ps = []
    a = 0
    firstXY = []
    res = np.empty([0, 2])
    for i in des.split(','):
        rgb_y = i[-13:]
        r = int(rgb_y[0:2], 16)
        g = int(rgb_y[2:4], 16)
        b = int(rgb_y[4:6], 16)
        y = int(rgb_y[-2:])
        rgby.append([r, g, b, y])
    for i in range(1, len(des.split(','))):
        ps.append([int(des.split(',')[i].split('|')[0]),
                  int(des.split(',')[i].split('|')[1])])
    for i in rgby:
        result = np.logical_and(
            abs(src[:, :, 0:1] - i[0]) < i[3], abs(src[:, :, 1:2] - i[1]) < i[3], abs(src[:, :, 2:3] - i[2]) < i[3])
        results = np.argwhere(np.all(result == True, axis=2)).tolist()
        if a == 0:
            firstXY = copy.deepcopy(results)
        else:
            nextnextXY = copy.deepcopy(results)
            for index in nextnextXY:
                index[0] = int(index[0]) - ps[a-1][1]
                index[1] = int(index[1]) - ps[a-1][0]
            q = set([tuple(t) for t in firstXY])
            w = set([tuple(t) for t in nextnextXY])
            matched = np.array(list(q.intersection(w)))
            if len(matched) == 0:
                return -1
            res = np.append(res, matched, axis=0)
        a += 1
    res = res.tolist()
    for i in res:
        if res.count(i) == len(des.split(',')) - 1:
            # print('time:', time.time() - st)
            return i[1], i[0]
    return -1


def matchBallColor(id, queimg, resultque, timecash):
    '''
    @@ 多线程找色调用
    @  resultque: 存储检测结果
    '''
    while(1):
        cap = queimg.get()
        cap = cap[regions[id][1]:regions[id][3], regions[id][0]:regions[id][2]]
        findpos = findMultColor(cap, colors[id])
        findtime = time.time()
        if findpos != -1 and findtime-timecash[id] > 0.7:
            print('Find:', str(id))
            timecash[id] = findtime
            resultque.append([id, findtime])


def clickMouse(hwnd, x, y):
    '''
    @@ 仅支持左键后台单击
    @
    '''
    x = int(x)
    y = int(y)
    position = win32api.MAKELONG(x, y)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN,
                         0, position)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, position)


def sendKeys(hwnd, keyid, method=0, times=0):
    '''
    @@ 后台窗口发送按键, 单击0, 按下1, 抬起2, 延时3
    @
    '''
    num1 = win32api.MapVirtualKey(keyid, 0)
    dparam = 1 | (num1 << 16)
    uparam = 1 | (num1 << 16) | (1 << 30) | (1 << 31)
    if method == 0:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, keyid, dparam)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, keyid, uparam)
    elif method == 1:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, keyid, dparam)
    elif method == 2:
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, keyid, uparam)
    elif method == 3:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, keyid, dparam)
        time.sleep(times)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, keyid, uparam)


def getPointerAddress(process_name, base, offsets):
    '''
    @@ 读取内存
    @  process_name: Gw2-64.exe
    @  base: Gw2_base_lp + 偏移
    @  offsets: 多级偏移
    @  return: 地址
    '''
    remote_pointer = RemotePointer(process_name.process_handle, base)
    for offset in offsets:
        if offset != offsets[-1]:
            remote_pointer = RemotePointer(
                process_name.process_handle, remote_pointer.value + offset)
        else:
            try:
                return remote_pointer.value + offset
            except:
                return -1
    if len(offsets) == 0:
        return RemotePointer(process_name.process_handle, remote_pointer.value)


def calAngle(x1, y1, x2, y2):
    '''
    @@ 计算角度
    @  return: angle (North-0°, 0~360°)
    '''
    if x1 == x2 and y2 >= y1:
        return 0
    if x1 == x2 and y2 < y1:
        return 180
    if y1 == y2 and x1 > x2:
        return 90
    if y1 == y2 and x1 < x2:
        return 270
    k = -(y2-y1)/(x2-x1)
    result = np.arctan(k)*57.29577
    if x1 > x2:
        result += 270
    else:
        result += 90
    return result


def getMoveAttribute(process_name, process_lp):
    '''
    @@ 获取坐标及面向
    @  return: x y z angle
    '''
    try:
        P_x = round(process_name.read_float(getPointerAddress(process_name,
                                                              process_lp + 0x2AAE168, [0x50, 0x30])), 3)
        P_y = round(process_name.read_float(getPointerAddress(process_name,
                                                              process_lp + 0x2AAE168, [0x50, 0x34])), 3)
        P_z = round(process_name.read_float(getPointerAddress(process_name,
                                                              process_lp + 0x2AAE168, [0x50, 0x38])), 3)
        Face_y = round(process_name.read_float(getPointerAddress(process_name,
                                                                 process_lp + 0x2AAE168, [0x50, 0x174])), 3)
        Face_x = round(process_name.read_float(getPointerAddress(process_name,
                                                                 process_lp + 0x2AAE168, [0x50, 0x170])), 3)
        P_a = calAngle(0, 0, Face_x, Face_y)
        return P_x, P_y, P_z, P_a
    except:
        return -1, -1, -1, -1


def getMapId():
    '''
    @@ 读取地图id
    @  过图成功后, oldid = newid
    '''
    Game = pymem.Pymem("Gw2-64.exe")
    modules = list(Game.list_modules())
    for module in modules:
        if module.name == 'Gw2-64.exe':
            Moduladdr = module.lpBaseOfDll
    oldid = Game.read_int(Moduladdr + 0x2481990)
    newid = Game.read_int(Moduladdr + 0x2A5FEAC)
    return oldid, newid


def positionCon(hwnd):
    '''
    @@ 人物位置检测, 用于多线程
    '''
    global play_state
    while(1):
        Gw2 = Pymem("Gw2-64.exe")
        Gw2_base_lp = Gw2.base_address
        oldid, newid = getMapId()
        if oldid == newid:
            P_x, P_y, _, _ = getMoveAttribute(Gw2, Gw2_base_lp)
            if oldid == 881 and -68.38 < P_x < 61.15 and 29.79 < P_y < 137.93:
                play_state = 1  # 舞台状态
                time.sleep(3)   # 3秒检测一次
            else:
                play_state = 0
                oldid, newid = getMapId()
                if oldid == 18:  # 在神佑地图上
                    sendKeys(hwnd, 70)  # F开启对话
                    time.sleep(0.8)
                    cap = windowCapture(mode=1, hwnd=hwnd)
                    findpos = findMultColor(
                        cap, 'f7c66e-000010,-2|-2|e8e9d2-000010,-5|-3|000000-000010,-4|-1|d8cab1-000010')
                    if findpos != -1:
                        clickMouse(hwnd, findpos[0], findpos[1]-30)    # 进图
                elif oldid == 881:  # 在游乐场里
                    # time.sleep(1)
                    findPress = 1
                    # 跑到雪人旁边
                    while(1):
                        P_x, P_y, _, _ = getMoveAttribute(Gw2, Gw2_base_lp)
                        for snowman in snowmans:
                            if snowman[0]-3 < P_x < snowman[0]+3 and snowman[1]-4 < P_y < snowman[1]+4:
                                sendKeys(hwnd, 87, method=2)
                                findPress = 0
                                break
                        # sendKeys(hwnd, 87, method=1)
                        if findPress == 0:
                            break
                        sendKeys(hwnd, 77)  # M
                        time.sleep(0.2)
                        sendKeys(hwnd, 77)  # M
                        time.sleep(0.2)
                        sendKeys(hwnd, 87, method=3, times=0.3)
                    # 和雪人对话 进场
                    while(1):
                        findPress = 1
                        # 等待时，防止被传送
                        P_x, P_y, _, _ = getMoveAttribute(Gw2, Gw2_base_lp)
                        for snowman in snowmans:
                            if snowman[0]-3 < P_x < snowman[0]+3 and snowman[1]-4 < P_y < snowman[1]+4:
                                sendKeys(hwnd, 87, method=2)
                                findPress = 0
                                break
                        if findPress == 1:
                            break   # 跳出，重新跑雪人
                        cap = windowCapture(mode=1, hwnd=hwnd)
                        findpos = findMultColor(
                            cap, 'c16503-000020,3|4|f5e895-000020')
                        if findpos != -1:
                            sendKeys(hwnd, 70)  # F
                            time.sleep(0.8)
                            cap = windowCapture(mode=1, hwnd=hwnd)
                            findpos = findMultColor(
                                cap, 'f7c76b-000010,-4|-1|000000-000010')
                            if findpos != -1:
                                clickMouse(
                                    hwnd, findpos[0], findpos[1]-30)    # 进场
                                time.sleep(0.5)
                                sendKeys(hwnd, 77)  # M
                                time.sleep(0.2)
                                sendKeys(hwnd, 77)  # M
                                time.sleep(0.2)
                            break


def closePopWindow(hwnd):
    while(1):
        time.sleep(1)
        # 结算弹窗
        # if play_state == 1:
        #     cap = windowCapture(mode=1, hwnd=hwnd)
        #     findpos = findMultColor(
        #         cap, '0d0504-000010,-2|-3|584b47-000010')
        #     if findpos != -1:
        #         clickMouse(hwnd, findpos[0], findpos[1]-30)    # 关闭结算界面

        # 说明 弹窗
        if play_state == 0:
            cap = windowCapture(mode=1, hwnd=hwnd)
            findpos = findMultColor(
                cap, '2a1e1d-000010,-3|-4|8f8478-000010,3|-4|b4aa9a-000010,-3|2|756c5d-000010')
            if findpos != -1:
                clickMouse(hwnd, findpos[0], findpos[1]-30)    # 关闭说明


if __name__ == "__main__":
    hwnd = win32gui.FindWindow(0, '激战2')
    setWindowPos(hwnd)
    windowCapture(hwnd)
    getWindowRect(hwnd)
    add_positionCon = threading.Thread(
        target=positionCon, name='positionCon', args=(hwnd,))
    add_positionCon.start()

    add_closePopWindow = threading.Thread(
        target=closePopWindow, name='closePopWindow', args=(hwnd,))
    add_closePopWindow.start()

    queimg = Queue()
    MG = mp.Manager()
    timecash = MG.list([0, 0, 0, 0, 0, 0, 0, 0])
    resultque = MG.list()

    for i in range(8):  # 创建8个线程, 弹琴
        process = mp.Process(  # threading.Thread(
            target=matchBallColor, name=str(i), args=(i, queimg, resultque, timecash))
        process.start()

    while(1):
        if queimg.empty():
            queimg.put(windowCapture(mode=1, hwnd=hwnd))

        if len(resultque) != 0:
            if resultque[0][0] == 0 and time.time() - resultque[0][1] >= 1.4:
                sendKeys(hwnd, 97)
                print('Press:', str(1))
                del resultque[0]
            elif resultque[0][0] == 1 and time.time() - resultque[0][1] >= 1.6:
                sendKeys(hwnd, 98)
                print('Press:', str(2))
                del resultque[0]
            elif resultque[0][0] == 2 and time.time() - resultque[0][1] >= 1.5:
                sendKeys(hwnd, 99)
                print('Press:', str(3))
                del resultque[0]
            elif resultque[0][0] == 3 and time.time() - resultque[0][1] >= 1.5:
                sendKeys(hwnd, 100)
                print('Press:', str(4))
                del resultque[0]
            elif resultque[0][0] == 4 and time.time() - resultque[0][1] >= 1.5:
                sendKeys(hwnd, 102)
                print('Press:', str(5))
                del resultque[0]
            elif resultque[0][0] == 5 and time.time() - resultque[0][1] >= 1.5:
                sendKeys(hwnd, 103)
                print('Press:', str(6))
                del resultque[0]
            elif resultque[0][0] == 6 and time.time() - resultque[0][1] >= 1.5:
                sendKeys(hwnd, 104)
                print('Press:', str(7))
                del resultque[0]
            elif resultque[0][0] == 7 and time.time() - resultque[0][1] >= 1.4:
                sendKeys(hwnd, 105)
                print('Press:', str(8))
                del resultque[0]
