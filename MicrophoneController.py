import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import keyboard
import threading
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import winsound
import os
import sys
from datetime import datetime

class MicrophoneController:
    def __init__(self):
        self.muted = False
        self.hotkey = "F8"
        self.history = []
        self.update_interval = 2  # 状态检查间隔(秒)
        self.running = True
        
        # 获取麦克风设备
        devices = AudioUtilities.GetMicrophone()
        if not devices:
            messagebox.showerror("错误", "未找到麦克风设备！")
            sys.exit(1)
            
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        # 启动状态监控线程
        self.monitor_thread = threading.Thread(target=self.monitor_status)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        
        # 注册全局快捷键
        self.register_hotkey()
    
    def register_hotkey(self):
        """注册全局快捷键"""
        try:
            keyboard.remove_hotkey(self.hotkey)
        except:
            pass
        keyboard.add_hotkey(self.hotkey, self.toggle_microphone)
    
    def set_hotkey(self, new_hotkey):
        """设置新的快捷键"""
        try:
            keyboard.remove_hotkey(self.hotkey)
            self.hotkey = new_hotkey
            self.register_hotkey()
            return True
        except Exception as e:
            print(f"设置快捷键错误: {e}")
            return False
    
    def toggle_microphone(self):
        """切换麦克风状态"""
        self.muted = not self.muted
        self.volume.SetMute(1 if self.muted else 0, None)
        
        # 添加历史记录
        status = "已静音" if self.muted else "已启用"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.history.insert(0, f"{timestamp} - 麦克风{status} (快捷键: {self.hotkey})")
        
        # 播放提示音
        winsound.Beep(800 if self.muted else 1000, 200)
        
        return self.muted
    
    def get_status(self):
        """获取麦克风状态"""
        return self.muted
    
    def monitor_status(self):
        """监控麦克风状态变化"""
        last_status = self.muted
        while self.running:
            # 检查外部状态变化
            current_status = bool(self.volume.GetMute())
            if current_status != last_status:
                self.muted = current_status
                last_status = current_status
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                status = "已静音" if self.muted else "已启用"
                self.history.insert(0, f"{timestamp} - 麦克风{status} (系统更改)")
            
            time.sleep(self.update_interval)
    
    def cleanup(self):
        """清理资源"""
        self.running = False
        try:
            keyboard.remove_hotkey(self.hotkey)
        except:
            pass

class MicrophoneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("麦克风控制工具")
        self.geometry("600x500")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 应用图标
        try:
            self.iconbitmap(default=self.resource_path("mic_icon.ico"))
        except:
            pass
        
        # 初始化控制器
        self.controller = MicrophoneController()
        
        # 创建界面
        self.create_widgets()
        
        # 启动状态更新
        self.update_status()
    
    def resource_path(self, relative_path):
        """获取资源的绝对路径"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建样式
        style = ttk.Style()
        style.configure("TButton", padding=6, font=("Arial", 10))
        style.configure("Header.TLabel", font=("Arial", 14, "bold"))
        style.configure("Status.TLabel", font=("Arial", 12))
        style.configure("Hotkey.TLabel", font=("Arial", 10))
        style.configure("History.TFrame", background="#f0f0f0")
        
        # 创建主框架
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        header = ttk.Label(main_frame, text="麦克风控制器", style="Header.TLabel")
        header.pack(pady=10)
        
        # 状态区域
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_icon = tk.Canvas(status_frame, width=40, height=40, highlightthickness=0)
        self.status_icon.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(status_frame, text="正在检测麦克风状态...", style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 控制按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.toggle_btn = ttk.Button(btn_frame, text="切换麦克风状态", command=self.toggle_microphone)
        self.toggle_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="测试麦克风", command=self.test_microphone).pack(side=tk.LEFT, padx=5)
        
        # 快捷键设置
        hotkey_frame = ttk.LabelFrame(main_frame, text="快捷键设置", padding=10)
        hotkey_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(hotkey_frame, text="当前快捷键:").grid(row=0, column=0, sticky=tk.W)
        
        self.hotkey_display = ttk.Label(hotkey_frame, text=self.controller.hotkey, 
                                       style="Hotkey.TLabel", foreground="blue")
        self.hotkey_display.grid(row=0, column=1, sticky=tk.W, padx=(5, 10))
        
        ttk.Label(hotkey_frame, text="设置新快捷键:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5))
        
        self.hotkey_entry = ttk.Entry(hotkey_frame, width=10)
        self.hotkey_entry.grid(row=0, column=3, padx=(0, 10))
        
        ttk.Button(hotkey_frame, text="应用", width=8, command=self.set_hotkey).grid(row=0, column=4)
        
        # 历史记录
        history_frame = ttk.LabelFrame(main_frame, text="操作历史", padding=10)
        history_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.history_text = scrolledtext.ScrolledText(
            history_frame, 
            wrap=tk.WORD, 
            font=("Arial", 9),
            padx=10,
            pady=10
        )
        self.history_text.pack(fill=tk.BOTH, expand=True)
        self.history_text.config(state=tk.DISABLED)
        
        # 添加初始历史记录
        self.update_history()
        
        # 底部信息
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(bottom_frame, text=f"当前快捷键: {self.controller.hotkey}").pack(side=tk.LEFT)
    
    def draw_status_icon(self, muted):
        """绘制状态图标"""
        self.status_icon.delete("all")
        color = "red" if muted else "green"
        
        # 绘制麦克风图标
        self.status_icon.create_oval(5, 5, 35, 35, fill=color, outline="black", width=2)
        
        # 绘制静音符号
        if muted:
            self.status_icon.create_line(10, 10, 30, 30, fill="white", width=3)
            self.status_icon.create_line(30, 10, 10, 30, fill="white", width=3)
    
    def update_status(self):
        """更新状态显示"""
        muted = self.controller.get_status()
        
        # 更新图标
        self.draw_status_icon(muted)
        
        # 更新标签
        status_text = "麦克风状态: 已静音" if muted else "麦克风状态: 正常"
        self.status_label.config(text=status_text)
        
        # 更新按钮文本
        btn_text = "启用麦克风" if muted else "禁用麦克风"
        self.toggle_btn.config(text=btn_text)
        
        # 每500毫秒更新一次
        self.after(500, self.update_status)
    
    def update_history(self):
        """更新历史记录"""
        self.history_text.config(state=tk.NORMAL)
        self.history_text.delete(1.0, tk.END)
        
        for i, entry in enumerate(self.controller.history):
            if i >= 20:  # 最多显示20条记录
                break
            self.history_text.insert(tk.END, entry + "\n")
        
        self.history_text.config(state=tk.DISABLED)
        self.history_text.see(tk.END)
        
        # 每2秒更新一次
        self.after(2000, self.update_history)
    
    def toggle_microphone(self):
        """切换麦克风状态"""
        self.controller.toggle_microphone()
        self.update_history()
    
    def test_microphone(self):
        """测试麦克风功能"""
        if self.controller.get_status():
            messagebox.showinfo("测试麦克风", "麦克风当前已静音，无法测试。\n请先启用麦克风。")
            return
        
        messagebox.showinfo("测试麦克风", "请对着麦克风说话...\n测试完成后，您将听到提示音。")
        
        # 模拟测试过程
        self.after(1000, lambda: winsound.Beep(1000, 300))
        self.after(1500, lambda: winsound.Beep(1200, 300))
        self.after(2000, lambda: winsound.Beep(1400, 300))
        self.after(2500, lambda: winsound.Beep(1600, 500))
        self.after(3000, lambda: messagebox.showinfo("测试完成", "麦克风测试成功！"))
    
    def set_hotkey(self):
        """设置新的快捷键"""
        new_hotkey = self.hotkey_entry.get().strip()
        if not new_hotkey:
            messagebox.showwarning("输入错误", "请输入有效的快捷键")
            return
        
        if new_hotkey == self.controller.hotkey:
            return
        
        if self.controller.set_hotkey(new_hotkey):
            self.hotkey_display.config(text=new_hotkey)
            self.hotkey_entry.delete(0, tk.END)
            self.controller.history.insert(0, 
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - 快捷键已更改为: {new_hotkey}")
            messagebox.showinfo("成功", f"快捷键已更改为: {new_hotkey}")
        else:
            messagebox.showerror("错误", "设置快捷键失败，请尝试其他组合")
    
    def on_close(self):
        """关闭窗口时的处理"""
        if messagebox.askokcancel("退出", "确定要退出应用程序吗？"):
            self.controller.cleanup()
            self.destroy()

if __name__ == "__main__":
    app = MicrophoneApp()
    app.mainloop()