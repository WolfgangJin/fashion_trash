import tkinter as tk
from tkinter import ttk, messagebox
import keyboard
import time
import threading
import pythoncom
import uiautomation as auto
import os
from datetime import datetime
from pygame import mixer  # 用于播放MP3


class WeChatAutoReply:
    def __init__(self, root):
        self.root = root
        self.root.title("意识DT抢车位工具windows版 - Powered by 黑喵喵")

        # 初始化音频 mixer
        mixer.init()

        # 设置窗口图标
        ico_path = os.path.join(os.path.dirname(__file__), "dt.ico")
        if os.path.exists(ico_path):
            self.root.iconbitmap(ico_path)

        self.root.geometry("350x470")  # 增加高度以容纳新功能

        # 运行状态
        self.running = False
        self.keyword = "排的老板扣1"
        self.interval = 0.1
        self.reminder_time = "21:29"
        self.bgm_file = "bgm.mp3"  # 默认BGM文件

        # 说明文本
        ttk.Label(root, text="使用说明：", font=('Arial', 10, 'bold')).pack(pady=(5, 0))
        ttk.Label(root, text="1. 使用前需鼠标点击微信输入框\n2. 点击开始识别\n3. 识别到关键词后会自动输入1并回车",
                  wraplength=280, justify='left').pack()

        # 监控间隔设置
        ttk.Label(root, text="监控频率(0.1-1秒):").pack()
        self.interval_entry = ttk.Entry(root)
        self.interval_entry.insert(0, "0.1")
        self.interval_entry.pack()

        # 关键词设置
        ttk.Label(root, text="触发关键词:").pack()
        self.keyword_entry = ttk.Entry(root)
        self.keyword_entry.insert(0, self.keyword)
        self.keyword_entry.pack()

        # 提醒时间设置
        ttk.Label(root, text="提醒时间(HH:MM):").pack()
        self.reminder_entry = ttk.Entry(root)
        self.reminder_entry.insert(0, self.reminder_time)
        self.reminder_entry.pack()

        # BGM提醒开关
        self.bgm_var = tk.BooleanVar(value=True)
        self.bgm_cb = ttk.Checkbutton(root, text="提醒时播放BGM", variable=self.bgm_var)
        self.bgm_cb.pack(pady=5)

        # 控制按钮
        self.start_btn = ttk.Button(root, text="开始 (Ctrl+Alt+S)", command=self.start)
        self.start_btn.pack(pady=5)

        self.stop_btn = ttk.Button(root, text="停止 (Ctrl+Alt+S)", command=self.stop, state=tk.DISABLED)
        self.stop_btn.pack()

        # 日志区域
        self.log = tk.Text(root, height=8, font=('Arial', 9))
        self.log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        keyboard.add_hotkey('ctrl+alt+s', self.toggle_running)

        # 启动提醒检查线程
        self.reminder_thread = threading.Thread(target=self.check_reminder, daemon=True)
        self.reminder_thread.start()

    def validate_interval(self, value):
        """验证输入是否为0.1-1之间的数字"""
        try:
            num = float(value)
            if 0.1 <= num <= 1:
                return True
            return False
        except ValueError:
            return False

    def validate_time(self, time_str):
        """验证时间格式是否为HH:MM"""
        try:
            datetime.strptime(time_str, "%H:%M")
            return True
        except ValueError:
            return False

    def toggle_running(self):
        if self.running:
            self.stop()
        else:
            self.start()

    def start(self):
        if self.running:
            return

        interval_value = self.interval_entry.get()
        if not self.validate_interval(interval_value):
            messagebox.showerror("错误", "请输入0.1到1之间的数字！")
            return

        self.running = True
        self.keyword = self.keyword_entry.get()
        self.interval = float(interval_value)

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.write_log(f"监控已启动 (频率: {self.interval}秒)")

        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

    def stop(self):
        if not self.running:
            return
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.write_log("监控已停止")

    def play_bgm(self):
        """播放BGM音乐"""
        bgm_path = os.path.join(os.path.dirname(__file__), self.bgm_file)
        if os.path.exists(bgm_path):
            try:
                mixer.music.load(bgm_path)
                mixer.music.set_volume(0.1)  # 设置音量为50%
                mixer.music.play()
                self.write_log("正在播放提醒音乐...")
            except Exception as e:
                self.write_log(f"音乐播放失败: {str(e)}")
                self.root.bell()
        else:
            self.write_log(f"警告: 未找到BGM文件 {self.bgm_file}")
            self.root.bell()

    def check_reminder(self):
        """检查是否到达提醒时间"""
        last_reminder_time = None
        while True:
            now = datetime.now()
            current_time = now.strftime("%H:%M")
            reminder_time = self.reminder_entry.get()

            if self.validate_time(reminder_time):
                if current_time == reminder_time and reminder_time != last_reminder_time:
                    if self.bgm_var.get():  # 检查是否启用BGM
                        self.root.after(0, self.play_bgm)
                    self.root.after(0, self.show_reminder)
                    last_reminder_time = reminder_time
            time.sleep(10)  # 每10秒检查一次

    def show_reminder(self):
        """显示提醒弹窗"""
        reminder_time = self.reminder_entry.get()
        self.write_log(f"提醒时间到！当前时间: {reminder_time}")
        messagebox.showinfo("提醒", f"已到达设定的提醒时间 {reminder_time}！")

    def monitor_loop(self):
        pythoncom.CoInitialize()
        while self.running:
            try:
                wechat = auto.WindowControl(ClassName="WeChatMainWndForPC")
                if wechat.Exists():
                    last_msg = self.get_last_message(wechat)
                    if last_msg and self.keyword.lower() in last_msg.lower():
                        self.write_log(
                            "触发关键词: {}".format(last_msg[:20] + "..." if len(last_msg) > 20 else last_msg))
                        keyboard.write("1\n")  # 输入1并回车
            except Exception as e:
                self.write_log(f"错误: {str(e)}")
            time.sleep(self.interval)
        pythoncom.CoUninitialize()

    def get_last_message(self, window):
        chat_list = window.ListControl(Name="消息")
        if chat_list.Exists():
            items = chat_list.GetChildren()
            return items[-1].Name if items else None
        return None

    def write_log(self, message):
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log.see(tk.END)
        self.root.update()

    def on_closing(self):
        """窗口关闭时清理资源"""
        mixer.quit()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WeChatAutoReply(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()