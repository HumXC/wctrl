import win32api
import win32con
import win32ui


# 获取鼠标速度值
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


class Mouse:
    def __init__(
        self, windowName: str = None, className: str = None, window=None
    ) -> None:
        """为一个窗口创建一个鼠标，参数与 win32ui.FindWindow 函数相同
        当 windowName 为空时将选择当前活动窗口为目标窗口
        """
        if window != None:
            self.__window = window
        elif windowName == None:
            self.__window = win32ui.GetForegroundWindow()
        else:
            self.__window = win32ui.FindWindow(className, windowName)

    def moveTo(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_MOUSEMOVE, 0, lparam)

    def leftDown(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_LBUTTONDOWN, 0, lparam)

    def leftUp(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_LBUTTONUP, 0, lparam)

    def rightDown(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_RBUTTONDOWN, 0, lparam)

    def rightUp(self, x: int, y: int):
        lparam = self.__getlParam(x, y)
        self.__window.PostMessage(win32con.WM_RBUTTONUP, 0, lparam)

    def scrool(self, x: int, y: int, v: int):
        x_, y_ = self.__window.ClientToScreen((x, y))
        lparam = self.__getlParam(x_, y_)
        wparam = v << 16
        self.__window.PostMessage(win32con.WM_MOUSEWHEEL, wparam, lparam)

    def __getlParam(self, x, y):
        return y << 16 | x
