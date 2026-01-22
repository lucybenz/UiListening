# -*- coding: utf-8 -*-
"""
UI元素监控工具 - 主程序
功能：选择桌面UI元素，设置监控条件，触发时播放音效
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import json
import os
from ui_selector import UISelector
from monitor import MonitorManager
from sound_player import SoundPlayer


class MonitorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("UI元素监控工具")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # 组件
        self.monitor_manager = MonitorManager()
        self.sound_player = SoundPlayer()
        self.ui_selector = None

        # 监控项列表
        self.monitor_items = []

        # 配置文件路径
        self.config_file = "monitors.json"

        self.setup_ui()
        self.load_config()

        # 启动监控线程
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        """设置界面"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=10, pady=5)

        self.btn_add = ttk.Button(toolbar, text="添加监控 (鼠标选择)", command=self.start_selection)
        self.btn_add.pack(side=tk.LEFT, padx=5)

        self.btn_remove = ttk.Button(toolbar, text="删除选中", command=self.remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=5)

        self.btn_stop_sound = ttk.Button(toolbar, text="停止音效", command=self.stop_sound)
        self.btn_stop_sound.pack(side=tk.LEFT, padx=5)

        # 状态标签
        self.status_label = ttk.Label(toolbar, text="就绪")
        self.status_label.pack(side=tk.RIGHT, padx=5)

        # 监控列表
        list_frame = ttk.LabelFrame(self.root, text="监控列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 创建Treeview
        columns = ("name", "condition", "value", "sound", "status", "current")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("name", text="元素名称")
        self.tree.heading("condition", text="条件")
        self.tree.heading("value", text="目标值")
        self.tree.heading("sound", text="音效文件")
        self.tree.heading("status", text="状态")
        self.tree.heading("current", text="当前值")

        self.tree.column("name", width=200)
        self.tree.column("condition", width=60)
        self.tree.column("value", width=100)
        self.tree.column("sound", width=150)
        self.tree.column("status", width=80)
        self.tree.column("current", width=100)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部提示
        hint_frame = ttk.Frame(self.root)
        hint_frame.pack(fill=tk.X, padx=10, pady=5)

        hint_text = "提示：点击\"添加监控\"后，将鼠标移动到目标元素上，按 Ctrl+鼠标左键 确认选择，按 ESC 取消。"
        ttk.Label(hint_frame, text=hint_text, foreground="gray").pack(side=tk.LEFT)

    def start_selection(self):
        """开始选择UI元素"""
        self.root.iconify()  # 最小化主窗口
        self.status_label.config(text="选择中...")

        # 创建选择器
        self.ui_selector = UISelector(self.on_element_selected)
        self.ui_selector.start()

    def on_element_selected(self, element_info):
        """元素选择完成回调"""
        self.root.deiconify()  # 恢复主窗口

        if element_info is None:
            self.status_label.config(text="已取消")
            return

        self.status_label.config(text="就绪")

        # 打开配置对话框
        self.show_config_dialog(element_info)

    def show_config_dialog(self, element_info):
        """显示监控配置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("配置监控条件")
        dialog.transient(self.root)
        dialog.grab_set()

        # 元素信息
        info_frame = ttk.LabelFrame(dialog, text="选中的元素")
        info_frame.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(info_frame, text=f"名称: {element_info.get('name', 'N/A')}").pack(anchor=tk.W, padx=5, pady=3)
        ttk.Label(info_frame, text=f"类型: {element_info.get('control_type', 'N/A')}").pack(anchor=tk.W, padx=5, pady=3)
        ttk.Label(info_frame, text=f"当前值: {element_info.get('value', 'N/A')}").pack(anchor=tk.W, padx=5, pady=3)

        # 条件设置
        cond_frame = ttk.LabelFrame(dialog, text="监控条件")
        cond_frame.pack(fill=tk.X, padx=10, pady=8)

        # 值提取方式
        ttk.Label(cond_frame, text="提取方式:").grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)
        extract_var = tk.StringVar(value="原始值")
        extract_combo = ttk.Combobox(cond_frame, textvariable=extract_var,
                                      values=["原始值", "提取数字", "提取整数", "提取小数", "去除空格", "取长度"], width=12)
        extract_combo.grid(row=0, column=1, padx=5, pady=8, sticky=tk.W)

        # 提取方式说明
        extract_hint = ttk.Label(cond_frame, text="(用于处理混合文本)", foreground="gray")
        extract_hint.grid(row=0, column=2, padx=5, pady=8, sticky=tk.W)

        ttk.Label(cond_frame, text="条件:").grid(row=1, column=0, padx=5, pady=8, sticky=tk.W)
        condition_var = tk.StringVar(value="=")
        condition_combo = ttk.Combobox(cond_frame, textvariable=condition_var,
                                        values=[">", "<", "=", ">=", "<=", "!=", "包含", "不包含"], width=12)
        condition_combo.grid(row=1, column=1, padx=5, pady=8, sticky=tk.W)

        ttk.Label(cond_frame, text="目标值:").grid(row=2, column=0, padx=5, pady=8, sticky=tk.W)
        value_var = tk.StringVar(value=str(element_info.get('value', '')))
        value_entry = ttk.Entry(cond_frame, textvariable=value_var, width=30)
        value_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=8, sticky=tk.W)

        # 音效设置
        sound_frame = ttk.LabelFrame(dialog, text="提醒音效")
        sound_frame.pack(fill=tk.X, padx=10, pady=8)

        # 默认音效路径（程序同目录下的12788.wav）
        default_sound = os.path.join(os.path.dirname(os.path.abspath(__file__)), "12788.wav")
        sound_var = tk.StringVar(value=default_sound if os.path.exists(default_sound) else "")
        sound_entry = ttk.Entry(sound_frame, textvariable=sound_var, width=35)
        sound_entry.pack(side=tk.LEFT, padx=5, pady=8)

        def browse_sound():
            file_path = filedialog.askopenfilename(
                title="选择音效文件",
                filetypes=[("音频文件", "*.wav *.mp3"), ("所有文件", "*.*")]
            )
            if file_path:
                sound_var.set(file_path)

        ttk.Button(sound_frame, text="浏览...", command=browse_sound).pack(side=tk.LEFT, padx=5, pady=8)

        # 检测间隔
        interval_frame = ttk.LabelFrame(dialog, text="检测间隔")
        interval_frame.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(interval_frame, text="间隔(秒):").pack(side=tk.LEFT, padx=5, pady=8)
        interval_var = tk.StringVar(value="1")
        interval_entry = ttk.Entry(interval_frame, textvariable=interval_var, width=10)
        interval_entry.pack(side=tk.LEFT, padx=5, pady=8)

        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=15)

        def on_confirm():
            try:
                interval = float(interval_var.get())
            except ValueError:
                messagebox.showerror("错误", "检测间隔必须是数字")
                return

            if not sound_var.get():
                messagebox.showerror("错误", "请选择音效文件")
                return

            monitor_item = {
                "element_info": element_info,
                "condition": condition_var.get(),
                "target_value": value_var.get(),
                "extract_mode": extract_var.get(),
                "sound_file": sound_var.get(),
                "interval": interval,
                "enabled": True
            }

            self.add_monitor_item(monitor_item)
            dialog.destroy()

        ttk.Button(btn_frame, text="确定", command=on_confirm).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        # 自适应窗口大小
        dialog.update_idletasks()
        dialog.minsize(500, dialog.winfo_reqheight())
        dialog.geometry(f"500x{dialog.winfo_reqheight()}")

    def add_monitor_item(self, item):
        """添加监控项"""
        self.monitor_items.append(item)

        # 添加到列表
        element_info = item["element_info"]
        name = element_info.get("name", "") or element_info.get("automation_id", "") or "未命名"

        self.tree.insert("", tk.END, values=(
            name,
            item["condition"],
            item["target_value"],
            os.path.basename(item["sound_file"]),
            "监控中",
            element_info.get("value", "N/A")
        ))

        self.save_config()

    def remove_selected(self):
        """删除选中的监控项"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要删除的项")
            return

        index = self.tree.index(selected[0])
        self.tree.delete(selected[0])

        if 0 <= index < len(self.monitor_items):
            del self.monitor_items[index]

        self.save_config()

    def stop_sound(self):
        """停止音效"""
        self.sound_player.stop()

    def monitor_loop(self):
        """监控循环"""
        import time
        import ctypes

        # 在线程中初始化COM
        ctypes.windll.ole32.CoInitialize(None)

        try:
            self._do_monitor_loop()
        finally:
            ctypes.windll.ole32.CoUninitialize()

    def _do_monitor_loop(self):
        """实际的监控循环"""
        import time

        while self.monitoring:
            for i, item in enumerate(self.monitor_items):
                if not item.get("enabled", True):
                    continue

                try:
                    # 获取当前值
                    current_value = self.monitor_manager.get_element_value(item["element_info"])

                    # 应用提取方式
                    extract_mode = item.get("extract_mode", "原始值")
                    extracted_value = self.extract_value(current_value, extract_mode)

                    # 更新显示
                    self.update_tree_item(i, extracted_value)

                    # 检查条件
                    if self.check_condition(extracted_value, item["condition"], item["target_value"]):
                        # 触发音效
                        if not self.sound_player.is_playing():
                            self.sound_player.play(item["sound_file"])
                            self.update_tree_status(i, "已触发")
                    else:
                        self.update_tree_status(i, "监控中")

                except Exception as e:
                    self.update_tree_status(i, f"错误")

            time.sleep(0.5)

    def update_tree_item(self, index, current_value):
        """更新列表项的当前值"""
        try:
            items = self.tree.get_children()
            if 0 <= index < len(items):
                item_id = items[index]
                values = list(self.tree.item(item_id, "values"))
                values[5] = str(current_value)
                self.tree.item(item_id, values=values)
        except:
            pass

    def update_tree_status(self, index, status):
        """更新列表项的状态"""
        try:
            items = self.tree.get_children()
            if 0 <= index < len(items):
                item_id = items[index]
                values = list(self.tree.item(item_id, "values"))
                values[4] = status
                self.tree.item(item_id, values=values)
        except:
            pass

    def extract_value(self, value, mode):
        """根据提取方式处理值"""
        import re

        if value is None:
            return ""

        value_str = str(value)

        if mode == "原始值":
            return value_str

        elif mode == "提取数字":
            # 提取所有数字（包括小数点和负号）
            numbers = re.findall(r'-?\d+\.?\d*', value_str)
            if numbers:
                return numbers[0]  # 返回第一个匹配的数字
            return ""

        elif mode == "提取整数":
            # 只提取整数部分
            numbers = re.findall(r'-?\d+', value_str)
            if numbers:
                return numbers[0]
            return ""

        elif mode == "提取小数":
            # 提取小数
            numbers = re.findall(r'-?\d+\.\d+', value_str)
            if numbers:
                return numbers[0]
            # 如果没有小数，尝试提取整数
            numbers = re.findall(r'-?\d+', value_str)
            if numbers:
                return numbers[0]
            return ""

        elif mode == "去除空格":
            return value_str.replace(" ", "").replace("\t", "").replace("\n", "")

        elif mode == "取长度":
            return str(len(value_str))

        return value_str

    def check_condition(self, current, condition, target):
        """检查条件是否满足"""
        try:
            # 尝试数值比较
            current_num = float(current) if current else 0
            target_num = float(target) if target else 0

            if condition == ">":
                return current_num > target_num
            elif condition == "<":
                return current_num < target_num
            elif condition == "=":
                return current_num == target_num
            elif condition == ">=":
                return current_num >= target_num
            elif condition == "<=":
                return current_num <= target_num
            elif condition == "!=":
                return current_num != target_num
        except (ValueError, TypeError):
            pass

        # 字符串比较
        current_str = str(current) if current else ""
        target_str = str(target) if target else ""

        if condition == "=":
            return current_str == target_str
        elif condition == "!=":
            return current_str != target_str
        elif condition == "包含":
            return target_str in current_str
        elif condition == "不包含":
            return target_str not in current_str

        return False

    def save_config(self):
        """保存配置"""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.monitor_items, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    def load_config(self):
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    self.monitor_items = json.load(f)

                # 添加到列表
                for item in self.monitor_items:
                    element_info = item["element_info"]
                    name = element_info.get("name", "") or element_info.get("automation_id", "") or "未命名"

                    self.tree.insert("", tk.END, values=(
                        name,
                        item["condition"],
                        item["target_value"],
                        os.path.basename(item["sound_file"]),
                        "监控中",
                        "N/A"
                    ))
        except Exception as e:
            print(f"加载配置失败: {e}")

    def on_closing(self):
        """窗口关闭"""
        self.monitoring = False
        self.sound_player.stop()
        self.save_config()
        self.root.destroy()

    def run(self):
        """运行程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = MonitorApp()
    app.run()
