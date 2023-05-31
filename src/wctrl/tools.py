import atexit
import ctypes
import math
import threading
import time

import cv2
import psutil
import win32api
import win32con
import win32gui
import win32process
import win32ui
from PIL import Image

# 桌面句柄
WINDOW_HWIN = win32gui.GetDesktopWindow()

# 屏幕缩放比率
SCALE_FACTOR = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0

# 正确的屏幕长和宽
SCREEN_WIDTH = int(win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN))
SCREEN_HEIGHT = int(win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN))
SCREEN_LEFT = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
SCREEN_TOP = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)


def position():
    return win32api.GetCursorPos()


def moveTo(x, y):
    # 获取鼠标速度
    # See https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event#remarks
    x = math.ceil(65535 / 1920 * x)
    y = math.ceil(65535 / 1080 * y)
    win32api.mouse_event(
        win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, x, y, 0, 0
    )


def move(x, y):
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, x, y, 0, 0)


def click(button: str = "left", duration=0.01):
    if button == "right":
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        time.sleep(duration)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        return
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(duration)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def getPids(name: str) -> list[int]:
    """获取一个进程的 pid, name 是程序的名称，例如 "java.exe" """
    result = []
    for proc in psutil.process_iter(["pid", "name"]):
        if proc.name() == name:
            result.append(proc.pid)
    return result


def getWindowsWithPid(pid: int) -> list[int]:
    """获取进程的所有可视窗口的 hwnd
    之后可以使用 win32ui.CreateWindowFromHandle 函数构建 App 所需的 window
    """
    windows: list[int] = []

    def callback(hwnd, hwnds):
        if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            if process_id == pid:
                hwnds.append(hwnd)

    win32gui.EnumWindows(callback, windows)
    return windows


def showCvMat(winname: str, mat: cv2.Mat) -> None:
    cv2.imshow(winname, mat)
    window = win32ui.FindWindow("", winname)
    window.CenterWindow()
    cv2.waitKey()


HWINDC = win32gui.GetWindowDC(WINDOW_HWIN)
SCREEN_DC = win32ui.CreateDCFromHandle(HWINDC)
SCREEN_MEMDC = SCREEN_DC.CreateCompatibleDC()


# 释放资源
def __free():
    SCREEN_DC.DeleteDC()
    win32gui.ReleaseDC(WINDOW_HWIN, HWINDC)
    SCREEN_MEMDC.DeleteDC()


atexit.register(__free)
__screencapLock = threading.Lock()


def screencap():
    __screencapLock.acquire()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(SCREEN_DC, SCREEN_WIDTH, SCREEN_HEIGHT)
    SCREEN_MEMDC.SelectObject(bmp)
    SCREEN_MEMDC.BitBlt(
        (0, 0),
        (SCREEN_WIDTH, SCREEN_HEIGHT),
        SCREEN_DC,
        (SCREEN_LEFT, SCREEN_TOP),
        win32con.SRCCOPY,
    )
    # 获取位图字节数据
    bits = bmp.GetBitmapBits(True)
    win32gui.DeleteObject(bmp.GetHandle())
    bmp_bytes = bytes(bits, "", "")
    pil_image = Image.frombytes(
        "RGBA",
        (SCREEN_WIDTH, SCREEN_HEIGHT),
        bmp_bytes,
        "raw",
        "BGRA",
        0,
        SCREEN_WIDTH * 4,
    )
    __screencapLock.release()
    return pil_image


def windowcap(window):
    """截取窗口截图, window 可以由 win32ui.FindWindow 获取"""
    r = window.GetClientRect()
    width, height = r[2], r[3]
    dc = window.GetDC()
    memdc = dc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(dc, width, height)
    memdc.SelectObject(bmp)
    memdc.BitBlt((0, 0), (width, height), dc, (0, 0), win32con.SRCCOPY)
    dc.DeleteDC()
    memdc.DeleteDC()
    # 获取位图字节数据
    bits = bmp.GetBitmapBits(True)
    win32gui.DeleteObject(bmp.GetHandle())
    bmp_bytes = bytes(bits, "", "")
    pil_image = Image.frombytes(
        "RGB", (width, height), bmp_bytes, "raw", "BGRX", 0, width * 4
    )
    return pil_image
