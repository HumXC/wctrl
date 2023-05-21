from typing import Dict, Tuple

import cv2
import numpy as np
from PIL import Image


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
