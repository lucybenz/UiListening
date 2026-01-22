# -*- coding: utf-8 -*-
"""
监控管理器 - 负责定位和获取UI元素的值
"""

import uiautomation as auto
import time


class MonitorManager:
    def __init__(self):
        self.element_cache = {}  # 缓存已定位的元素

    def get_element_value(self, element_info):
        """
        获取元素的当前值
        element_info: 元素信息字典（由UISelector生成）
        """
        # 尝试重新定位元素
        element = self._find_element(element_info)

        if element is None:
            return None

        # 获取值
        return self._get_value(element)

    def _find_element(self, element_info):
        """根据元素信息定位元素"""
        locator = element_info.get("locator", {})

        # 方法1: 通过AutomationId定位（最可靠）
        if locator.get("automation_id"):
            element = self._find_by_automation_id(locator)
            if element:
                return element

        # 方法2: 通过路径定位
        if locator.get("path"):
            element = self._find_by_path(locator)
            if element:
                return element

        # 方法3: 通过组合属性定位
        element = self._find_by_properties(locator)
        if element:
            return element

        return None

    def _find_by_automation_id(self, locator):
        """通过AutomationId定位"""
        try:
            automation_id = locator.get("automation_id", "")
            process_id = locator.get("process_id", 0)

            if not automation_id:
                return None

            # 先找到进程的窗口
            if process_id:
                windows = auto.GetRootControl().GetChildren()
                for win in windows:
                    try:
                        if win.ProcessId == process_id:
                            # 在窗口中搜索
                            element = win.Control(AutomationId=automation_id)
                            if element.Exists(0, 0):
                                return element
                    except:
                        continue

            # 全局搜索
            element = auto.Control(AutomationId=automation_id)
            if element.Exists(0, 0):
                return element

        except Exception as e:
            pass

        return None

    def _find_by_path(self, locator):
        """通过路径定位"""
        try:
            path = locator.get("path", [])
            if not path:
                return None

            # 从根开始
            current = auto.GetRootControl()

            # 跳过第一个（通常是Desktop）
            for i, path_item in enumerate(path[1:], 1):
                found = False
                children = current.GetChildren()

                for child in children:
                    try:
                        # 匹配属性
                        match = True

                        if path_item.get("automation_id"):
                            if child.AutomationId != path_item["automation_id"]:
                                match = False

                        if match and path_item.get("class_name"):
                            if child.ClassName != path_item["class_name"]:
                                match = False

                        if match and path_item.get("control_type"):
                            if child.ControlTypeName != path_item["control_type"]:
                                match = False

                        if match and path_item.get("name"):
                            if child.Name != path_item["name"]:
                                match = False

                        if match:
                            current = child
                            found = True
                            break
                    except:
                        continue

                if not found:
                    # 尝试只匹配control_type和class_name
                    for child in children:
                        try:
                            if (child.ControlTypeName == path_item.get("control_type", "") and
                                child.ClassName == path_item.get("class_name", "")):
                                current = child
                                found = True
                                break
                        except:
                            continue

                if not found:
                    return None

            return current

        except Exception as e:
            pass

        return None

    def _find_by_properties(self, locator):
        """通过属性组合定位"""
        try:
            process_id = locator.get("process_id", 0)
            control_type = locator.get("control_type", "")
            class_name = locator.get("class_name", "")
            name = locator.get("name", "")

            # 构建搜索条件
            search_props = {}

            if control_type:
                search_props["ControlType"] = getattr(auto.ControlType, control_type.replace("Control", ""), None)

            if class_name:
                search_props["ClassName"] = class_name

            if name:
                search_props["Name"] = name

            if not search_props:
                return None

            # 在特定进程中搜索
            if process_id:
                windows = auto.GetRootControl().GetChildren()
                for win in windows:
                    try:
                        if win.ProcessId == process_id:
                            element = win.Control(**search_props)
                            if element.Exists(0, 0):
                                return element
                    except:
                        continue

            # 全局搜索
            element = auto.Control(**search_props)
            if element.Exists(0, 0):
                return element

        except Exception as e:
            pass

        return None

    def _get_value(self, element):
        """获取元素的值"""
        value = ""

        # 尝试ValuePattern
        try:
            pattern = element.GetValuePattern()
            if pattern:
                value = pattern.Value
                if value:
                    return value
        except:
            pass

        # 尝试TextPattern
        try:
            pattern = element.GetTextPattern()
            if pattern:
                value = pattern.DocumentRange.GetText(-1)
                if value:
                    return value
        except:
            pass

        # 尝试RangeValuePattern
        try:
            pattern = element.GetRangeValuePattern()
            if pattern:
                value = str(pattern.Value)
                if value:
                    return value
        except:
            pass

        # 尝试SelectionPattern
        try:
            pattern = element.GetSelectionPattern()
            if pattern:
                selection = pattern.GetSelection()
                if selection:
                    names = [s.Name for s in selection if s.Name]
                    if names:
                        return ", ".join(names)
        except:
            pass

        # 尝试TogglePattern (复选框等)
        try:
            pattern = element.GetTogglePattern()
            if pattern:
                state = pattern.ToggleState
                return str(state)
        except:
            pass

        # 最后尝试Name
        try:
            if element.Name:
                return element.Name
        except:
            pass

        return value
