# -*- coding: utf-8 -*-
"""
音效播放器 - 循环播放音效，鼠标移动时停止
"""

import threading
import time
import os


class SoundPlayer:
    def __init__(self):
        self.playing = False
        self.play_thread = None
        self.mouse_listener = None
        self.last_mouse_pos = None
        self.sound_file = None

        # pygame初始化标志
        self._pygame_initialized = False
        self._mixer = None

    def _init_pygame(self):
        """延迟初始化pygame"""
        if not self._pygame_initialized:
            try:
                import pygame
                pygame.mixer.init()
                self._mixer = pygame.mixer
                self._pygame_initialized = True
            except Exception as e:
                print(f"pygame初始化失败: {e}")
                return False
        return True

    def play(self, sound_file):
        """开始播放音效（循环）"""
        if self.playing:
            return

        if not os.path.exists(sound_file):
            print(f"音效文件不存在: {sound_file}")
            return

        self.sound_file = sound_file
        self.playing = True

        # 记录初始鼠标位置
        self._record_mouse_position()

        # 启动播放线程
        self.play_thread = threading.Thread(target=self._play_loop, daemon=True)
        self.play_thread.start()

        # 启动鼠标监听
        self._start_mouse_listener()

    def _record_mouse_position(self):
        """记录当前鼠标位置"""
        try:
            from pynput import mouse
            # 获取当前鼠标位置
            import ctypes

            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

            pt = POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            self.last_mouse_pos = (pt.x, pt.y)
        except:
            self.last_mouse_pos = None

    def _play_loop(self):
        """播放循环"""
        if not self._init_pygame():
            # 如果pygame失败，使用winsound
            self._play_loop_winsound()
            return

        try:
            # 加载音效
            self._mixer.music.load(self.sound_file)
            self._mixer.music.play(-1)  # -1表示循环播放

            while self.playing:
                time.sleep(0.1)

            self._mixer.music.stop()
        except Exception as e:
            print(f"播放失败: {e}")
            # 尝试使用winsound
            self._play_loop_winsound()

    def _play_loop_winsound(self):
        """使用winsound播放（备用方案，仅支持wav）"""
        try:
            import winsound

            while self.playing:
                try:
                    # SND_ASYNC允许异步播放
                    winsound.PlaySound(self.sound_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                    # 等待一段时间再重复
                    for _ in range(30):  # 约3秒检查一次
                        if not self.playing:
                            break
                        time.sleep(0.1)
                except:
                    break

            # 停止播放
            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception as e:
            print(f"winsound播放失败: {e}")

    def _start_mouse_listener(self):
        """启动鼠标移动监听"""
        try:
            from pynput import mouse

            def on_move(x, y):
                if not self.playing:
                    return False  # 停止监听

                if self.last_mouse_pos:
                    # 计算移动距离
                    dx = abs(x - self.last_mouse_pos[0])
                    dy = abs(y - self.last_mouse_pos[1])

                    # 如果移动超过阈值，停止播放
                    if dx > 10 or dy > 10:
                        self.stop()
                        return False  # 停止监听

                return True

            self.mouse_listener = mouse.Listener(on_move=on_move)
            self.mouse_listener.start()
        except Exception as e:
            print(f"鼠标监听启动失败: {e}")

    def stop(self):
        """停止播放"""
        self.playing = False

        # 停止pygame
        try:
            if self._mixer:
                self._mixer.music.stop()
        except:
            pass

        # 停止winsound
        try:
            import winsound
            winsound.PlaySound(None, winsound.SND_PURGE)
        except:
            pass

        # 停止鼠标监听
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
            except:
                pass
            self.mouse_listener = None

    def is_playing(self):
        """是否正在播放"""
        return self.playing
