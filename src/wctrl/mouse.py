import win32api
import win32con
import win32ui


# 获取鼠标速度值
def get_mouse_speed():
    # 查询 Windows 注册表
    key = win32api.RegOpenKey(
        win32con.HKEY_CURRENT_USER, "Control Panel\\Mouse", False, win32con.KEY_READ
    )
    # 查询鼠标加速度值
    val, _ = win32api.RegQueryValueEx(key, "MouseSensitivity")
    # 关闭注册表键
    win32api.RegCloseKey(key)
    return int(val)


MOUSE_SPEED = get_mouse_speed()


class Mouse:
    def __init__(self, handle: int) -> None:
        """为一个窗口创建一个鼠标, handle 是窗口的句柄可以使用 win32gui.FindWindow 获取"""
        self.__window = win32ui.CreateWindowFromHandle(handle)

    def moveTo(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_MOUSEMOVE, 0, lparam)

    def lDown(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_LBUTTONDOWN, 0, lparam)

    def lUp(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_LBUTTONUP, 0, lparam)

    def rDown(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_RBUTTONDOWN, 0, lparam)

    def rUp(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_RBUTTONUP, 0, lparam)

    def lClick(self, x: int, y: int):
        self.lDown(x, y)
        self.lUp(x, y)

    def rClick(self, x: int, y: int):
        self.rDown(x, y)
        self.rUp(x, y)

    def scrool(self, x: int, y: int, v: int):
        x_, y_, _, _, _ = self.__window.ClientToScreen((x, y), None)
        lparam = self.__getlParam(x_, y_)
        wparam = v << 16
        self.__window.PostMessage(win32con.WM_MOUSEWHEEL, wparam, lparam)

    def __getlParam(self, x, y):
        return y << 16 | x
