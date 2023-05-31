import atexit
import ctypes
import threading

import win32con
import win32gui
import win32ui
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()


class CapTool:
    def __init__(self, handle: int) -> None:
        """为一个窗口创建一个截图工具, handle 是窗口的句柄可以使用 win32gui.FindWindow 获取"""
        self.__window = win32ui.CreateWindowFromHandle(handle)
        # 截图相关
        self.__capLock = threading.Lock()
        self.__dc = self.__window.GetDC()
        self.__memdc = self.__dc.CreateCompatibleDC()
        atexit.register(self.__dc.DeleteDC)
        atexit.register(self.__memdc.DeleteDC)

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
        bmp_bytes = bytes(bits)  # type: ignore
        pil_image = Image.frombytes(
            "RGB", (width, height), bmp_bytes, "raw", "BGRX", 0, width * 4
        )
        return pil_image
