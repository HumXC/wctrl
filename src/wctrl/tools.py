import atexit
import ctypes
import math
import string
import threading
import time
from typing import Dict, Tuple

import cv2
import numpy as np
import psutil
import pyautogui as pg
import win32api
import win32con
import win32gui
import win32process
import win32ui
from PIL import Image, ImageOps

# 桌面句柄
WINDOW_HWIN = win32gui.GetDesktopWindow()

# 屏幕缩放比率
SCALE_FACTOR = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0

# 正确的屏幕长和宽
SCREEN_WIDTH = int(win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN))
SCREEN_HEIGHT = int(win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN))
SCREEN_LEFT = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
SCREEN_TOP = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)


# 获取鼠标加速度值
def get_mouse_speed():
    # 查询 Windows 注册表
    key = win32api.RegOpenKeyEx(
        win32con.HKEY_CURRENT_USER, "Control Panel\\Mouse", 0, win32con.KEY_READ
    )
    # 查询鼠标加速度值
    val, _ = win32api.RegQueryValueEx(key, "MouseSensitivity")
    # 关闭注册表键
    win32api.RegCloseKey(key)
    return int(val)


MOUSE_SPEED = get_mouse_speed()


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
    windows = []

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
    pass


class ImgTool:
    def __init__(
        self,
        threshold=0.9,
        is_set_center=False,
        method=cv2.TM_CCORR_NORMED,
        is_use_mask=False,
        is_show=False,
    ) -> None:
        self.__threshold = threshold
        self.__is_set_center = is_set_center
        self.__method = method
        self.__is_show = is_show
        self.__is_use_mask = is_use_mask

    def __setArgs(
        self,
        threshold: float,
        is_set_center: bool,
        method: int,
        is_use_mask: bool,
        is_show: bool,
    ):
        if threshold == None:
            threshold = self.__threshold
        if is_set_center == None:
            is_set_center = self.__is_set_center
        if method == None:
            method = self.__method
        if is_use_mask == None:
            is_use_mask = self.__is_use_mask
        if is_show == None:
            is_show = self.__is_show
        return threshold, is_set_center, method, is_use_mask, is_show

    # 用于缓存模板图片
    __tmpl_cacha: Dict[str, cv2.Mat] = {}

    def __getTmplMat(self, name: str) -> cv2.Mat:
        # 添加进缓存
        tmpl_mat = self.__tmpl_cacha.get(name, None)
        if tmpl_mat is None:
            tmpl_mat = cv2.imread(name, cv2.IMREAD_UNCHANGED)
            self.__tmpl_cacha[name] = tmpl_mat
        return tmpl_mat

    def find(
        self,
        img: Image,
        tmpl: str,
        threshold: float = None,
        is_set_center: bool = None,
        method: int = None,
        is_use_mask: bool = None,
        is_show: bool = None,
    ) -> Tuple[int, int, bool, int]:
        """从 img 中查找 tmpl .
        tmpl 是模板图像的文件名称。
        threshold 是匹配的阈值，如果匹配结果低于 threshold, 第三个返回值是 False,
        否则视为匹配成功，第三个返回值是 True, 第四个返回值是匹配值
        isSetCenter 如果为 True, 将会使返回坐标偏移至模板中心
        is_use_mask 如果为 True, 模板中透明的部分会被用作遮罩，花费时间将增长为原来的 2 倍
        """
        threshold, is_set_center, method, is_use_mask, is_show = self.__setArgs(
            threshold, is_set_center, method, is_use_mask, is_show
        )
        tmpl_mat = self.__getTmplMat(tmpl)
        img_mat = cv2.cvtColor(np.array(img), cv2.COLOR_BGRA2RGBA)
        # 使用模板图像的透明通道作遮罩
        mask = None
        if is_use_mask:
            mask = tmpl_mat[:, :, 3]
        result = cv2.matchTemplate(img_mat, tmpl_mat, method, mask=mask)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val < threshold:
            return 0, 0, False, max_val
        x = max_loc[0]
        y = max_loc[1]
        # 偏移坐标至图片中心
        if is_set_center:
            x += tmpl_mat.shape[1] // 2
            y += tmpl_mat.shape[0] // 2
        if is_show:
            # 获取模板图的高度和宽度
            ht, wt, _ = tmpl_mat.shape
            # 在原图上用矩形标注匹配结果
            cv2.rectangle(
                img_mat, max_loc, (max_loc[0] + wt, max_loc[1] + ht), (0, 0, 255), 2
            )
            _img = cv2.cvtColor(img_mat, cv2.COLOR_BGR2RGB)
            img_pil = Image.fromarray(_img)
            img_pil.show()
            img_pil.close()
        return x, y, True, max_val

    def mfind(
        self,
        img: Image,
        tmpl: str,
        threshold: float = None,
        is_set_center: bool = None,
        method: int = None,
        is_use_mask: bool = None,
        is_show: bool = None,
    ) -> Tuple[Tuple[int, int]]:
        """从 img 中查找 tmpl 返回多个结果
        参数定义与 find 方法相同
        """
        threshold, is_set_center, method, is_show = self.__setArgs(
            threshold, is_set_center, method, is_show
        )
        tmpl_mat = self.__getTmplMat(tmpl)
        img_mat = cv2.cvtColor(np.array(img), cv2.COLOR_BGRA2RGBA)
        mask = None
        if is_use_mask:
            mask = tmpl_mat[:, :, 3]
        result = cv2.matchTemplate(img_mat, tmpl_mat, method, mask=mask)
        loc = np.where(result >= threshold)
        points = tuple(zip(*loc[::-1]))
        points = points[::-1]
        if is_set_center:
            offsetX = tmpl_mat.shape[1] // 2
            offsetY = tmpl_mat.shape[0] // 2
            for i in range(len(points)):
                points[i][0] += offsetX
                points[i][1] += offsetY

        if is_show:
            # 获取模板图的高度和宽度
            ht, wt, _ = tmpl_mat.shape
            # 在原图上用矩形标注匹配结果
            for i in range(len(points)):
                pt = points[i]
                cv2.rectangle(img_mat, pt, (pt[0] + wt, pt[1] + ht), (0, 0, 255), 2)
                cv2.putText(
                    img_mat, str(i), pt, cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2
                )
            _img = cv2.cvtColor(img_mat, cv2.COLOR_BGR2RGB)
            img_pil = Image.fromarray(_img)
            img_pil.show()
            img_pil.close()

        return points


def hold_key(key, hold_time):
    start = time.time()
    while time.time() - start < hold_time:
        pg.keyDown(key)


VkCode = {
    "ctrl": win32con.VK_CONTROL,
    "back": win32con.VK_BACK,
    "tab": win32con.VK_TAB,
    "return": win32con.VK_RETURN,
    "shift": win32con.VK_SHIFT,
    "control": win32con.VK_CONTROL,
    "menu": win32con.VK_MENU,
    "pause": win32con.VK_PAUSE,
    "capital": win32con.VK_CAPITAL,
    "escape": win32con.VK_ESCAPE,
    "space": win32con.VK_SPACE,
    "end": win32con.VK_END,
    "home": win32con.VK_HOME,
    "left": win32con.VK_LEFT,
    "up": win32con.VK_UP,
    "right": win32con.VK_RIGHT,
    "down": win32con.VK_DOWN,
    "print": win32con.VK_PRINT,
    "snapshot": win32con.VK_SNAPSHOT,
    "insert": win32con.VK_INSERT,
    "delete": win32con.VK_DELETE,
    "lwin": win32con.VK_LWIN,
    "rwin": win32con.VK_RWIN,
    "numpad0": win32con.VK_NUMPAD0,
    "numpad1": win32con.VK_NUMPAD1,
    "numpad2": win32con.VK_NUMPAD2,
    "numpad3": win32con.VK_NUMPAD3,
    "numpad4": win32con.VK_NUMPAD4,
    "numpad5": win32con.VK_NUMPAD5,
    "numpad6": win32con.VK_NUMPAD6,
    "numpad7": win32con.VK_NUMPAD7,
    "numpad8": win32con.VK_NUMPAD8,
    "numpad9": win32con.VK_NUMPAD9,
    "multiply": win32con.VK_MULTIPLY,
    "add": win32con.VK_ADD,
    "separator": win32con.VK_SEPARATOR,
    "subtract": win32con.VK_SUBTRACT,
    "decimal": win32con.VK_DECIMAL,
    "divide": win32con.VK_DIVIDE,
    "f1": win32con.VK_F1,
    "f2": win32con.VK_F2,
    "f3": win32con.VK_F3,
    "f4": win32con.VK_F4,
    "f5": win32con.VK_F5,
    "f6": win32con.VK_F6,
    "f7": win32con.VK_F7,
    "f8": win32con.VK_F8,
    "f9": win32con.VK_F9,
    "f10": win32con.VK_F10,
    "f11": win32con.VK_F11,
    "f12": win32con.VK_F12,
    "numlock": win32con.VK_NUMLOCK,
    "scroll": win32con.VK_SCROLL,
    "lshift": win32con.VK_LSHIFT,
    "rshift": win32con.VK_RSHIFT,
    "lcontrol": win32con.VK_LCONTROL,
    "rcontrol": win32con.VK_RCONTROL,
    "lmenu": win32con.VK_LMENU,
    "rmenu": win32con.VK_RMENU,
    "esc": win32con.VK_ESCAPE,
    "enter": win32con.VK_RETURN,
}


class Keyboard:
    def __init__(
        self, windowName: str = None, className: str = None, window=None
    ) -> None:
        """为一个窗口创建一个键盘，参数与 win32ui.FindWindow 函数相同
        当 windowName 为空时将选择当前活动窗口为目标窗口
        """
        if window != None:
            self.__window = window
        elif windowName == None:
            self.__window = win32ui.GetForegroundWindow()
        else:
            self.__window = win32ui.FindWindow(className, windowName)

    def __getKey(self, keyname: str, key: int):
        if keyname == "":
            return key

        if len(keyname) == 1 and keyname in string.printable:
            # https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-vkkeyscana
            return win32api.VkKeyScan(keyname) & 0xFF
        else:
            return VkCode[keyname]

    def down(self, keyname: str, hold: float = 0, is_sync=False, key: int = 0):
        """发送一个按键。keyname 是按键的字符串形式，例如 "a", " "
        hold 是按键的持续时间，在按下按键之后阻塞一段时间
        时间到了便自动调用 up 方法，如果 is_sync 为 False, 则不会阻塞当前线程
        key 是键码，可以在 win32con.VK_* 里找到
        key 不是必要的，如果 keyname 不为空，则会根据 keyname 推断 key 的值
        如果 keyname 无效，可以设置 key 的值
        """
        keycode = self.__getKey(keyname, key)
        scan_code = win32api.MapVirtualKey(keycode, 0)
        lparam = (scan_code << 16) | 1
        self.__window.PostMessage(win32con.WM_KEYDOWN, keycode, lparam)
        if hold != 0:

            def up():
                time.sleep(hold)
                self.up(key=keycode)

            if is_sync:
                up()
            else:
                threading.Thread(name="hold keyboard", target=up).start()

    def up(self, keyname="", key: int = 0):
        """松开按键，参数说明见 down 方法"""
        keycode = self.__getKey(keyname, key)
        scan_code = win32api.MapVirtualKey(keycode, 0)
        lparam = (scan_code << 16) | 0xC0000001
        self.__window.PostMessage(win32con.WM_KEYUP, keycode, lparam)


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
