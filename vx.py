import tkinter as tk
from tkinter import ttk
import keyboard
import time
import threading
import pythoncom
import uiautomation as auto


class WeChatAutoReply:
    def __init__(self, root):
        self.root = root
        self.root.title("意识DT上车抢位工具")
        self.root.geometry("300x300")

        # 运行状态
        self.running = False
        self.keyword = "test"
        self.interval = 0.2

        # 说明文本
        ttk.Label(root, text="使用说明：", font=('Arial', 10, 'bold')).pack(pady=(5, 0))
        ttk.Label(root, text="1. 使用前需鼠标点击输入框\n2. 开始识别后会自动输入1并回车",
                  wraplength=280, justify='left').pack()

        # 监控间隔设置
        ttk.Label(root, text="监控频率(秒):").pack()
        self.interval_spin = ttk.Spinbox(root, from_=0.1, to=1, increment=0.1, value=0.2)
        self.interval_spin.pack()

        # 关键词设置
        ttk.Label(root, text="触发关键词:").pack()
        self.keyword_entry = ttk.Entry(root)
        self.keyword_entry.insert(0, self.keyword)
        self.keyword_entry.pack()

        # 控制按钮
        self.start_btn = ttk.Button(root, text="开始 (Ctrl+Alt+S)", command=self.start)
        self.start_btn.pack(pady=10)

        self.stop_btn = ttk.Button(root, text="停止 (Ctrl+Alt+S)", command=self.stop, state=tk.DISABLED)
        self.stop_btn.pack()

        # 日志区域
        self.log = tk.Text(root, height=6, font=('Arial', 9))
        self.log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        keyboard.add_hotkey('ctrl+alt+s', self.toggle_running)

    def toggle_running(self):
        if self.running:
            self.stop()
        else:
            self.start()

    def start(self):
        if self.running: return
        self.running = True
        self.keyword = self.keyword_entry.get()
        self.interval = float(self.interval_spin.get())
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.write_log("监控已启动 (频率: {}秒)".format(self.interval))

        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

    def stop(self):
        if not self.running: return
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.write_log("监控已停止")

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
                self.write_log("错误: {}".format(str(e)))
            time.sleep(self.interval)
        pythoncom.CoUninitialize()

    def get_last_message(self, window):
        chat_list = window.ListControl(Name="消息")
        if chat_list.Exists():
            items = chat_list.GetChildren()
            return items[-1].Name if items else None
        return None

    def write_log(self, message):
        self.log.insert(tk.END, "[{}] {}\n".format(time.strftime("%H:%M:%S"), message))
        self.log.see(tk.END)
        self.root.update()


if __name__ == "__main__":
    root = tk.Tk()
    app = WeChatAutoReply(root)
    root.mainloop()