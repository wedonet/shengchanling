import time
import win32gui,win32con

#import keyboardEmulation as ke
import pyautogui


def get_windows(windowsname,filename):
    # 获取窗口句柄
    hwnd = win32gui.FindWindow(None,windowsname)
    # 将窗口放在前台，并激活该窗口
    win32gui.SetForegroundWindow(hwnd)


    pyautogui.press('a')

    pyautogui.press(['p','y','space'], interval=0.1)

    # 输入helloworld

    #scancodes = [0x23, 0x12, 0x26, 0x26, 0x18, 0x11, 0x18, 0x13, 0x26, 0x20, 0x2a]

    #for code in  scancodes:
    #    ke.key_press(code)

    # 保存
    #ke.key_down(0x1d)
    #ke.key_down(0x1f)
    #ke.key_up(0x1d)
    #ke.key_up(0x1f)



    # 关闭窗口

    time.sleep(1);
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

get_windows("新建文本文档.txt - 记事本","截图.png")