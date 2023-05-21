import threading
from string import printable
from time import sleep

import win32api
import win32con
import win32ui

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

        if len(keyname) == 1 and keyname in printable:
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
                sleep(hold)
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
