# -*- coding: utf-8 -*-
"""
UI元素选择器 - 使用鼠标选择桌面UI元素
"""

import uiautomation as auto
import threading
import time
import ctypes
from ctypes import wintypes


class UISelector:
    def __init__(self, callback):
        """
        初始化UI选择器
        callback: 选择完成后的回调函数，参数为元素信息字典或None（取消）
        """
        self.callback = callback
        self.running = False
        self.current_element = None
        self.select_thread = None

        # 上一次高亮的矩形
        self.last_rect = None

    def start(self):
        """开始选择模式"""
        self.running = True

        # 在新线程中运行选择逻辑
        self.select_thread = threading.Thread(target=self._selection_loop, daemon=True)
        self.select_thread.start()

    def stop(self):
        """停止选择模式"""
        self.running = False
        # 清除高亮
        self._clear_highlight()

    def _selection_loop(self):
        """选择循环 - 跟踪鼠标下的元素"""
        # 在线程中初始化COM
        ctypes.windll.ole32.CoInitialize(None)

        try:
            self._do_selection_loop()
        finally:
            ctypes.windll.ole32.CoUninitialize()

    def _do_selection_loop(self):
        """实际的选择循环"""
        # 获取按键状态的常量
        VK_CONTROL = 0x11
        VK_LBUTTON = 0x01
        VK_ESCAPE = 0x1B

        user32 = ctypes.windll.user32

        ctrl_was_pressed = False

        while self.running:
            try:
                # 检查ESC键 - 取消选择
                if user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000:
                    self._cancel_selection()
                    break

                # 检查Ctrl+左键点击 - 确认选择
                ctrl_pressed = user32.GetAsyncKeyState(VK_CONTROL) & 0x8000
                lbutton_pressed = user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000

                if ctrl_pressed and lbutton_pressed and not ctrl_was_pressed:
                    # Ctrl+点击，确认选择
                    self._confirm_selection()
                    break

                ctrl_was_pressed = ctrl_pressed and lbutton_pressed

                # 获取鼠标位置
                point = auto.GetCursorPos()

                # 获取鼠标下的元素
                element = auto.ControlFromPoint(point[0], point[1])

                if element:
                    self.current_element = element
                    # 显示高亮框
                    self._show_highlight(element)

            except Exception as e:
                pass

            time.sleep(0.05)

    def _confirm_selection(self):
        """确认选择"""
        self.stop()

        if self.current_element:
            # 在当前线程获取元素信息（COM已初始化）
            element_info = self._get_element_info(self.current_element)
            # 使用after在主线程调用回调
            self._call_callback(element_info)
        else:
            self._call_callback(None)

    def _cancel_selection(self):
        """取消选择"""
        self.stop()
        self._call_callback(None)

    def _call_callback(self, result):
        """调用回调函数"""
        if self.callback:
            self.callback(result)

    def _get_element_info(self, element):
        """获取元素信息"""
        info = {
            "name": "",
            "automation_id": "",
            "class_name": "",
            "control_type": "",
            "value": "",
            "runtime_id": None,
            "bounding_rect": None,
            "process_id": 0,
            "locator": {}
        }

        try:
            info["name"] = element.Name or ""
        except:
            pass

        try:
            info["automation_id"] = element.AutomationId or ""
        except:
            pass

        try:
            info["class_name"] = element.ClassName or ""
        except:
            pass

        try:
            info["control_type"] = element.ControlTypeName or ""
        except:
            pass

        try:
            info["value"] = self._get_element_value(element)
        except:
            pass

        try:
            info["runtime_id"] = element.GetRuntimeId()
        except:
            pass

        try:
            rect = element.BoundingRectangle
            info["bounding_rect"] = {
                "left": rect.left,
                "top": rect.top,
                "right": rect.right,
                "bottom": rect.bottom
            }
        except:
            pass

        try:
            info["process_id"] = element.ProcessId
        except:
            pass

        info["locator"] = self._build_locator(element)

        return info

    def _get_element_value(self, element):
        """获取元素的值"""
        value = ""

        try:
            pattern = element.GetValuePattern()
            if pattern:
                value = pattern.Value
                if value:
                    return value
        except:
            pass

        try:
            pattern = element.GetTextPattern()
            if pattern:
                value = pattern.DocumentRange.GetText(-1)
                if value:
                    return value
        except:
            pass

        try:
            pattern = element.GetRangeValuePattern()
            if pattern:
                value = str(pattern.Value)
                if value:
                    return value
        except:
            pass

        try:
            if element.Name:
                return element.Name
        except:
            pass

        return value

    def _build_locator(self, element):
        """构建元素定位器"""
        locator = {
            "automation_id": "",
            "name": "",
            "class_name": "",
            "control_type": "",
            "process_id": 0,
            "path": []
        }

        try:
            locator["automation_id"] = element.AutomationId or ""
            locator["name"] = element.Name or ""
            locator["class_name"] = element.ClassName or ""
            locator["control_type"] = element.ControlTypeName or ""
            locator["process_id"] = element.ProcessId
        except:
            pass

        try:
            path = []
            current = element
            depth = 0
            max_depth = 20

            while current and depth < max_depth:
                try:
                    path_item = {
                        "name": current.Name or "",
                        "automation_id": current.AutomationId or "",
                        "class_name": current.ClassName or "",
                        "control_type": current.ControlTypeName or ""
                    }
                    path.insert(0, path_item)

                    parent = current.GetParentControl()
                    if parent is None or parent == current:
                        break
                    current = parent
                    depth += 1
                except:
                    break

            locator["path"] = path
        except:
            pass

        return locator

    def _show_highlight(self, element):
        """显示高亮框"""
        try:
            rect = element.BoundingRectangle
            if rect.width() > 0 and rect.height() > 0:
                new_rect = (rect.left, rect.top, rect.right, rect.bottom)

                # 只有矩形变化时才重绘
                if new_rect != self.last_rect:
                    self._clear_highlight()
                    self._draw_highlight_rect(*new_rect)
                    self.last_rect = new_rect
        except:
            pass

    def _draw_highlight_rect(self, left, top, right, bottom):
        """绘制高亮矩形"""
        try:
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32

            hdc = user32.GetDC(0)

            # 红色画笔，3像素宽
            pen = gdi32.CreatePen(0, 3, 0x0000FF)
            old_pen = gdi32.SelectObject(hdc, pen)
            old_brush = gdi32.SelectObject(hdc, gdi32.GetStockObject(5))

            gdi32.Rectangle(hdc, left, top, right, bottom)

            gdi32.SelectObject(hdc, old_pen)
            gdi32.SelectObject(hdc, old_brush)
            gdi32.DeleteObject(pen)

            user32.ReleaseDC(0, hdc)
        except:
            pass

    def _clear_highlight(self):
        """清除高亮框"""
        try:
            if self.last_rect:
                # 只刷新上次高亮的区域
                left, top, right, bottom = self.last_rect
                # 扩大一点范围确保完全清除
                rect = (ctypes.c_long * 4)(left - 5, top - 5, right + 5, bottom + 5)
                ctypes.windll.user32.InvalidateRect(0, rect, True)
                self.last_rect = None
        except:
            pass
