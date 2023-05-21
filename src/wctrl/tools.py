import atexit
import ctypes
import math
import threading
import time
from typing import List

import cv2
import psutil
import win32api
import win32con
import win32gui
import win32process
import win32ui
from img_tool import ImgTool
from keyboard import Keyboard
from mouse import Mouse
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


def getPids(name: str):
    """获取一个进程的 pid, name 是程序的名称，例如 "java.exe" """
    result = []
    for proc in psutil.process_iter(["pid", "name"]):
        if proc.name() == name:
            result.append(proc.pid)
    return result


def getWindowsWithPid(pid: int):
    """获取进程的所有可视窗口的 hwnd
    之后可以使用 win32ui.CreateWindowFromHandle 函数构建 App 所需的 window
    """
    windows: List[int] = []

    def callback(hwnd, hwnds):
        if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            if process_id == pid:
                hwnds.append(hwnd)

    win32gui.EnumWindows(callback, windows)
    return windows


def showCvMat(winname: str, mat: any):
    cv2.imshow(winname, mat)
    window = win32ui.FindWindow(None, winname)
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
    bmp_bytes = bytes(bits)
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
    bmp_bytes = bytes(bits)
    pil_image = Image.frombytes(
        "RGB", (width, height), bmp_bytes, "raw", "BGRX", 0, width * 4
    )
    return pil_image


class App:
    def __init__(
        self,
        windowName: str = None,
        className: str = None,
        window=None,
        threshold=0.9,
        is_set_center=False,
        method=cv2.TM_CCORR_NORMED,
        is_use_mask=False,
        is_show=False,
    ) -> None:
        """通过 windowName 和 className 来查找窗口，或者你也可以指定 window
        后面的参数与 ImgTool 的初始化参数相同
        """

        if window != None:
            self.__window = window
        elif windowName == None:
            self.__window = win32ui.GetForegroundWindow()
        else:
            self.__window = win32ui.FindWindow(className, windowName)

        # 键盘相关
        self.__keyboard = Keyboard(window=self.__window)
        # 鼠标相关
        self.__mouse = Mouse(window=window)
        # 截图相关
        self.__capLock = threading.Lock()
        self.__dc = self.__window.GetDC()
        self.__memdc = self.__dc.CreateCompatibleDC()
        atexit.register(self.__dc.DeleteDC)
        atexit.register(self.__memdc.DeleteDC)
        # cv相关
        self.__nowImg: Image = None
        self.__isWindowLocked = False
        self.__imgTool = ImgTool(threshold, is_set_center, method, is_use_mask, is_show)

    def keyDown(self, keyname: str, hold: float = 0, is_sync=False, key: int = 0):
        """与 Keyboard.down 相同"""
        return self.__keyboard.down(keyname, hold, is_sync, key)

    def keyUp(self, keyname="", key: int = 0):
        """与 Keyboard.up 相同"""
        return self.__keyboard.up(keyname, key)

    def MouseLDown(self, x: int, y: int):
        self.__mouse.leftDown(x, y)

    def MouseRDown(self, x: int, y: int):
        self.__mouse.rightDown(x, y)

    def MouseLUp(self, x: int, y: int):
        self.__mouse.leftUp(x, y)

    def MouseRUp(self, x: int, y: int):
        self.__mouse.rightUp(x, y)

    def MouseScoor(self, x: int, y: int, v: int):
        self.__mouse.scrool(x, y, v)

    def cap(self):
        self.__capLock.acquire()
        rect = self.__window.GetClientRect()
        width, height = rect[2], rect[3]
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(self.__dc, width, height)
        self.__memdc.SelectObject(bmp)
        self.__memdc.BitBlt(
            (0, 0), (width, height), self.__dc, (0, 0), win32con.SRCCOPY
        )
        # 获取位图字节数据
        bits = bmp.GetBitmapBits(True)
        win32gui.DeleteObject(bmp.GetHandle())
        self.__capLock.release()
        bmp_bytes = bytes(bits)
        pil_image = Image.frombytes(
            "RGB", (width, height), bmp_bytes, "raw", "BGRX", 0, width * 4
        )
        return pil_image

    def getNowImg(self):
        """获取当前 find 和 mfind 方法所处理的图像"""
        if self.__isWindowLocked:
            if self.__nowImg == None:
                self.__nowImg = self.cap()
            return self.__nowImg
        else:
            self.__nowImg = self.cap()
            return self.__nowImg

    def find(
        self,
        tmpl: str,
        threshold: float = None,
        is_set_center: bool = None,
        method: int = None,
        is_use_mask: bool = None,
        is_show: bool = None,
        img: Image = None,
    ):
        """参数同 ImgTool.find
        有所不同的是，此方法默认从窗口中截图来获取 img
        当输入参数中 img 不为 None 时便不会自动截图，而是使用传入的 img
        """
        if img == None:
            img = self.getNowImg()

        return self.__imgTool.find(
            img, tmpl, threshold, is_set_center, method, is_use_mask, is_show
        )

    def mfind(
        self,
        tmpl: str,
        threshold: float = None,
        is_set_center: bool = None,
        method: int = None,
        is_use_mask: bool = None,
        is_show: bool = None,
        img: Image = None,
    ):
        """参数同 ImgTool.mfind
        有所不同的是，此方法默认从窗口中截图来获取 img
        当输入参数中 img 不为 None 时便不会自动截图，而是使用传入的 img
        """
        if img == None:
            img = self.getNowImg()

        return self.__imgTool.mfind(
            img, tmpl, threshold, is_set_center, method, is_use_mask, is_show
        )

    def lockWindow(self):
        """锁定 find 和 mfind 方法使用的 img
        在调用 unlockWindow 之前, find 和 mfind 不再重新截图
        一般与 unlockWindow 成对使用
        """
        self.__isWindowLocked = True

    def unlockWindow(self):
        """解除 lockWindow"""
        self.__isWindowLocked = False
