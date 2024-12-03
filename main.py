import os
import sys
import ctypes
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread, Event
import winreg
import json
import random
from PIL import Image, ImageDraw
import pystray
import win32gui
import win32con
from send2trash import send2trash

# 应用用户数据目录
# appdata_dir = os.getenv('APPDATA')
# # 配置文件夹
# config_dir = os.path.join(appdata_dir, 'WallpaperEngineShuiGeGe')
# # 配置文件路径
# config_path = os.path.join(config_dir, 'config.json')

# 应用名称
app_name = "【水哥哥】壁纸定时切换器"
# 包名称
app_package_name = "WallpaperSwitcher"
# 配置文件路径
config_path = "wsconfig.json"
# 默认时间间隔（分钟）
default_interval = 10

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))
# 获取图标路径
icon_path = os.path.join(current_dir, "icon.ico")

# 判断是否为开机启动
def is_startup():
    # 检查程序启动时是否带有 "--startup" 参数
    return '--startup' in sys.argv

# 获取图片文件列表
def get_image_files(folder):
    valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp']
    return [os.path.join(folder, f) for f in os.listdir(folder) if os.path.splitext(f)[1].lower() in valid_extensions]

# 更换桌面背景
def set_wallpaper(image_path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)

# 加载自定义托盘图标
def create_tray_icon(image_path):
    return Image.open(image_path)

# 创建一个简单的托盘图标
def create_image(width, height, color1, color2):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        [width // 2, 0, width, height // 2],
        fill=color2)
    dc.rectangle(
        [0, height // 2, width // 2, height],
        fill=color2)
    return image

# 查找指定标题的窗口
def find_window(text):
    # 寻找已有的窗口
    def enum_windows(hwnd, results):
        # 所有窗口进行排查
        window_text = win32gui.GetWindowText(hwnd)
        if text == window_text:
            results.append(hwnd)
        # 页面显示窗口进行排查
        # if win32gui.IsWindowVisible(hwnd):
        #     window_text = win32gui.GetWindowText(hwnd)
        #     if text == window_text:
        #         results.append(hwnd)
    hwnds = []
    win32gui.EnumWindows(enum_windows, hwnds)
    return hwnds

# 找到窗口在页面上的区域
def find_window_react (text):
    # 找到的窗口区域
    find_react = None
    # 枚举所有窗口
    def enum_windows(hwnd, results):
        nonlocal find_react
        # window_class = win32gui.GetClassName(hwnd)
        # if text in window_class:  # 假设微信窗口类名包含 WeChat
        #     rect = win32gui.GetWindowRect(hwnd)
        #     print(f"窗口位置: (Left: {rect[0]}, Top: {rect[1]})")
        #     print(f"窗口大小: (Width: {rect[2] - rect[0]}, Height: {rect[3] - rect[1]})")
        window_text = win32gui.GetWindowText(hwnd)
        if text in window_text:  # 假设微信窗口类名包含 WeChat
            rect = win32gui.GetWindowRect(hwnd)
            # print(f"窗口位置: (Left: {rect[0]}, Top: {rect[1]})")
            # print(f"窗口大小: (Width: {rect[2] - rect[0]}, Height: {rect[3] - rect[1]})")
            find_react = (rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1])
            # print(f"Region for pywin32: {found_region}")
    # 开始枚举所有窗口
    win32gui.EnumWindows(enum_windows, None)
    # 返回找到的窗口区域
    return find_react

# 激活窗口
def active_window(hwnd):
    # 激活窗口
    # ctypes.windll.user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
    # win32gui.SetForegroundWindow(hwnd)

    # 恢复窗口
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  # 恢复窗口
    time.sleep(0.1)  # 等待窗口恢复
    win32gui.SetForegroundWindow(hwnd)  # 将窗口置于前端
    
    # 如果窗口仍然没有刷新，尝试模拟鼠标或键盘事件
    # win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    # win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)

    # 恢复窗口并确保它正常显示
    # ctypes.windll.user32.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  # 恢复到正常大小
    # win32gui.SetForegroundWindow(hwnd)  # 将窗口置于前台
    # win32gui.SetFocus(hwnd)  # 获取焦点
    # time.sleep(0.5)  # 等待窗口完全恢复
    # # 确保消息循环继续执行
    # win32gui.PumpMessages()



# 创建Tkinter应用
class WallpaperChangerApp:
    def __init__(self, root):
        self.root = root
        self.folder_path = ""
        self.interval = default_interval
        self.startup_enabled = False  # 开机启动状态
        self.random_enabled = False  # 随机切换图片状态
        self.refresh_reset_enabled = False  # 随机切换图片状态
        self.close_minimize_enabled = True  # 最小化到托盘
        self.running = False
        self.thread = None
        self.image_files = []
        self.current_index = 0
        self.start_time = 0
        self.icon = None  # pystray 图标
        # self.tray_thread = None  # 托盘线程
        
        # 创建配置文件夹（如果不存在）
        # if not os.path.exists(config_dir):
        #     os.makedirs(config_dir)

        # 设置软件名称
        self.root.title(app_name)

        # 设置窗口图标
        self.root.iconbitmap(icon_path)

        # UI元素
        self.create_widgets()

        # 加载配置文件 (确保所有UI元素创建后调用)
        self.load_config()

        # 处理窗口关闭事件
        # self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 处理窗口关闭事件，最小化到托盘
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

    # 创建UI
    def create_widgets(self):

        # 显示文件夹路径
        self.folder_label = tk.Label(self.root, text="未选择文件夹")
        self.folder_label.pack(pady=10)

        # 创建一个框架来放置标签和输入框
        self.interval_frame = tk.Frame(self.root)
        self.interval_frame.pack(pady=5)

        # 时间间隔标签
        self.interval_label = tk.Label(self.interval_frame, text="间隔 (秒)")
        # self.interval_label = tk.Label(self.interval_frame, text="间隔 (分钟)")
        self.interval_label.pack(side=tk.LEFT)  # 将标签放在左侧

        # 时间间隔输入框
        self.interval_entry = tk.Entry(self.interval_frame, width=10)
        self.interval_entry.insert(0, self.interval)
        self.interval_entry.pack(side=tk.LEFT, padx=5)  # 将输入框放在标签右侧，增加一些间距
        
        # 选择和刷新按钮，水平布局
        self.select_folder_frame = tk.Frame(self.root)
        self.select_folder_frame.pack(pady=10)

        # 选择文件夹按钮
        self.select_folder_button = tk.Button(self.select_folder_frame, text="选择文件夹", command=self.select_folder)
        self.select_folder_button.grid(row=0, column=0, padx=10)

        # 刷新文件夹按钮
        self.refresh_button = tk.Button(self.select_folder_frame, text="刷新文件夹", command=self.refresh_images)
        self.refresh_button.grid(row=0, column=1, padx=10)
        # self.refresh_button.pack_forget()  # 初始隐藏

        # 上一张和下一张按钮，水平布局
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10)

        self.prev_button = tk.Button(self.button_frame, text="上一张", command=self.prev_wallpaper)
        self.prev_button.grid(row=0, column=0, padx=10)

        self.next_button = tk.Button(self.button_frame, text="下一张", command=self.next_wallpaper)
        self.next_button.grid(row=0, column=1, padx=10)

        # 开始/停止切换按钮
        self.toggle_button = tk.Button(self.root, text="开始切换", command=self.toggle_changing)
        self.toggle_button.pack(pady=10)
        
        # 删除壁纸按钮
        self.delete_button = tk.Button(self.root, text="删除壁纸", command=self.delete_wallpaper)
        self.delete_button.pack(pady=10)

        # 随机切换复选框
        self.random_var = tk.BooleanVar()
        self.random_frame = tk.Frame(self.root)
        self.random_frame.pack(pady=10)
        self.random_checkbox = tk.Checkbutton(self.random_frame, variable=self.random_var, command=self.toggle_random)
        self.random_checkbox.pack(side='left')
        self.random_label = tk.Label(self.random_frame, text="壁纸随机切换（下一张）", anchor='w', width=20)  # 设置固定宽度
        self.random_label.pack(side='left')

        # 刷新文件夹按钮会使图片顺序从0开始
        self.refresh_reset_var = tk.BooleanVar()
        self.refresh_reset_frame = tk.Frame(self.root)
        self.refresh_reset_frame.pack(pady=10)
        self.refresh_reset_checkbox = tk.Checkbutton(self.refresh_reset_frame, variable=self.refresh_reset_var, command=self.refresh_reset)
        self.refresh_reset_checkbox.pack(side='left')
        self.refresh_reset_label = tk.Label(self.refresh_reset_frame, text="刷新重置顺序（置 0）", anchor='w', width=20)  # 设置固定宽度
        self.refresh_reset_label.pack(side='left')

        # 右上角关闭按钮会最小化到托盘
        self.close_minimize_var = tk.BooleanVar()
        self.close_minimize_frame = tk.Frame(self.root)
        self.close_minimize_frame.pack(pady=10)
        self.close_minimize_checkbox = tk.Checkbutton(self.close_minimize_frame, variable=self.close_minimize_var, command=self.close_minimize)
        self.close_minimize_checkbox.pack(side='left')
        self.close_minimize_label = tk.Label(self.close_minimize_frame, text="最小化到托盘（右上X）", anchor='w', width=20)  # 设置固定宽度
        self.close_minimize_label.pack(side='left')

        # 开机启动复选框
        self.startup_var = tk.BooleanVar()
        self.startup_frame = tk.Frame(self.root)
        self.startup_frame.pack(pady=10)
        self.startup_checkbox = tk.Checkbutton(self.startup_frame, variable=self.startup_var, command=self.toggle_startup)
        self.startup_checkbox.pack(side='left')
        self.startup_label = tk.Label(self.startup_frame, text="开机启动（调试中...）", anchor='w', width=20)  # 设置固定宽度
        self.startup_label.pack(side='left')
        self.startup_checkbox.config(state='disabled')

        # 退出程序
        self.exit_button = tk.Button(self.root, text="退出应用", command=self.exit_app)
        self.exit_button.pack(pady=10)

    # 选择图片文件夹
    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path = folder_selected
            self.folder_label.config(text=f"已选择: {self.folder_path}")
            self.refresh_images()
        # else:
            # self.folder_label.config(text="未选择文件夹")

    # 刷新图片文件夹并重置当前索引
    def refresh_images(self):
        try:
            self.image_files = get_image_files(self.folder_path)
            if not self.image_files:
                messagebox.showerror("错误", "指定的文件夹中没有找到任何图片文件。")
                # self.folder_label.config(text="未选择文件夹")
                return False
            else:
                # 刷新后显示当前索引的图片，如果当前索引超出范围，显示最后一张
                if self.refresh_reset_enabled:
                    self.current_index = 0
                else:
                    if self.current_index >= len(self.image_files):
                        self.current_index = len(self.image_files) - 1
                self.set_wallpaper_in_background(self.image_files[self.current_index])
                return True
        except FileNotFoundError:
            messagebox.showerror("错误", "指定的文件夹不存在。")
            return False
        except PermissionError:
            messagebox.showerror("错误", "没有权限访问该文件夹。")
            return False
        except Exception as e:
            messagebox.showerror("错误", f"发生未知错误：{e}")
            return False
    
    # 删除壁纸
    def delete_wallpaper(self):
        try:
            # 直接移除，没法复原
            # os.remove(self.image_files[self.current_index])
            # 移除到回收站
            send2trash(os.path.normpath(self.image_files[self.current_index]))
            self.refresh_images()
        except FileNotFoundError:
            messagebox.showerror("错误", "指定的文件不存在。")
        except PermissionError:
            messagebox.showerror("错误", "没有权限删除该文件。")
        except Exception as e:
            messagebox.showerror("错误", f"发生未知错误：{e}")

    # 切换开始/停止按钮
    def toggle_changing(self):
        if self.running:
            self.stop_changing()
        else:
            self.start_changing()

    # 开始切换壁纸
    def start_changing(self):
        if not self.folder_path:
            messagebox.showerror("错误", "请先选择图片文件夹。")
            return
        try:
            self.interval = int(self.interval_entry.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的时间间隔。")
            return
        
        # 确保停止切换
        self.stop_changing()

        # 开始切换
        if self.refresh_images():
            self.running = True
            self.current_index -= 1  # 确保开始后，当前索引不变化
            self.update_tray_menu()  # 更新托盘菜单
            self.toggle_button.config(text="停止切换")
            self.thread = Thread(target=self.change_wallpaper)
            self.thread.start()
            
    # 停止切换壁纸
    def stop_changing(self):
        self.running = False
        self.update_tray_menu()  # 更新托盘菜单
        self.toggle_button.config(text="开始切换")
        # 等待线程结束
        if self.thread is not None:
            self.thread.join()  # 等待线程完成
            self.thread = None  # 清空线程引用

    # 切换壁纸的主函数
    def change_wallpaper(self):
        # 分钟转换为秒钟
        temp_interval = self.interval
        # temp_interval = self.interval * 60
        # 循环切换壁纸
        while self.running:
            if self.random_enabled:
                self.current_index = random.randint(0, len(self.image_files) - 1)
            else:
                self.current_index = (self.current_index + 1) % len(self.image_files)
            image_path = self.image_files[self.current_index]
            self.set_wallpaper_in_background(image_path)

            self.start_time = time.time()
            while time.time() - self.start_time < temp_interval:
                if not self.running:
                    return
                time.sleep(0.5)

    # 切换到上一张壁纸并重置计时器
    def prev_wallpaper(self):
        if self.image_files:
            self.current_index = (self.current_index - 1) % len(self.image_files)
            self.set_wallpaper_in_background(self.image_files[self.current_index])
            self.start_time = time.time()

    # 切换到下一张壁纸并重置计时器
    def next_wallpaper(self):
        if self.image_files:
            if self.random_enabled:
                self.current_index = random.randint(0, len(self.image_files) - 1)
            else:
                self.current_index = (self.current_index + 1) % len(self.image_files)
            self.set_wallpaper_in_background(self.image_files[self.current_index])
            self.start_time = time.time()
    
    # 随机切换状态切换
    def toggle_random(self):
        self.random_enabled = self.random_var.get()
        self.save_config()  # 保存随机状态

    # 刷新文件夹图片顺序从 0 开始
    def refresh_reset(self):
        self.refresh_reset_enabled = self.refresh_reset_var.get()
        self.save_config()  # 保存随机状态

    # 关闭窗口最小化
    def close_minimize(self):
        self.close_minimize_enabled = self.close_minimize_var.get()
        self.save_config()  # 保存随机状态

    # 在后台线程中更换壁纸，防止卡顿
    def set_wallpaper_in_background(self, image_path):
        # 保存配置
        self.save_config()
        # 切换背景任务
        def wallpaper_task():
            set_wallpaper(image_path)
            # print(f"已更换壁纸为：{image_path}")
        Thread(target=wallpaper_task).start()

    # 开机启动设置
    def toggle_startup(self):
        self.startup_enabled = self.startup_var.get()
        if self.startup_enabled:
            self.set_startup()
        else:
            self.remove_startup()

    # 设置开机启动
    def set_startup(self):
        self.startup_enabled = True
        self.startup_var.set(self.startup_enabled)
        key = winreg.HKEY_CURRENT_USER
        key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"
        exe_path = os.path.abspath(sys.argv[0])
        startup_command = f'"{exe_path}" --startup'  # 添加 --startup 参数
        # with winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS) as reg_key:
        #     winreg.SetValueEx(reg_key, app_package_name, 0, winreg.REG_SZ, exe_path)
        with winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS) as reg_key:
            winreg.SetValueEx(reg_key, app_package_name, 0, winreg.REG_SZ, startup_command)
        # messagebox.showinfo("提示", "已设置为开机启动。")
        self.save_config()  # 保存开机启动状态
        self.update_tray_menu()  # 更新托盘菜单

    # 取消开机启动
    def remove_startup(self):
        self.startup_enabled = False
        self.startup_var.set(self.startup_enabled)
        key = winreg.HKEY_CURRENT_USER
        key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(key, key_value, 0, winreg.KEY_ALL_ACCESS) as reg_key:
            try:
                winreg.DeleteValue(reg_key, app_package_name)
                # messagebox.showinfo("提示", "已取消开机启动。")
            except FileNotFoundError:
                messagebox.showinfo("提示", "未设置开机启动。")
        self.save_config()  # 保存开机启动状态
        self.update_tray_menu()  # 更新托盘菜单

    # 保存当前配置到本地
    def save_config(self):
        config = {
            "folder_path": self.folder_path,
            "interval": self.interval,
            "current_index": self.current_index,
            "running": self.running,
            "startup_enabled": self.startup_enabled,
            "random_enabled": self.random_enabled,
            "refresh_reset_enabled": self.refresh_reset_enabled,
            "close_minimize_enabled": self.close_minimize_enabled
        }
        with open(config_path, 'w') as f:
            json.dump(config, f)

    # 加载本地缓存的配置
    def load_config(self):
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                config = json.load(f)
                self.folder_path = config.get("folder_path", "")
                self.interval = config.get("interval", default_interval)
                self.current_index = config.get("current_index", 0)
                # self.running = config.get("running", False)
                self.startup_enabled = config.get("startup_enabled", False)
                self.random_enabled = config.get("random_enabled", False)
                self.refresh_reset_enabled = config.get("refresh_reset_enabled", False)
                self.close_minimize_enabled = config.get("close_minimize_enabled", True)
                # 更新UI
                self.interval_entry.delete(0, tk.END)
                self.interval_entry.insert(0, str(self.interval))
                self.startup_var.set(self.startup_enabled)
                self.random_var.set(self.random_enabled)
                self.refresh_reset_var.set(self.refresh_reset_enabled)
                self.close_minimize_var.set(self.close_minimize_enabled)
                if self.folder_path:
                    self.folder_label.config(text=f"已选择: {self.folder_path}")
                    self.refresh_images()
                    # if self.running:
                    #     self.start_changing()
                    # else:
                    #     self.refresh_images()
                else:
                    # self.running = False
                    self.folder_label.config(text="未选择文件夹")

    # 保存时间间隔（在输入框失去焦点时）
    def save_interval(self, event):
        try:
            self.interval = int(self.interval_entry.get())
            self.save_config()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的时间间隔。")

    # 启动托盘图标
    def start_tray_icon(self):
        self.icon = pystray.Icon(
            "壁纸定时更换",
            create_tray_icon(icon_path),
            menu=self.create_tray_menu()
        )

        # 1、阻塞运行，在它执行后则不能执行任何代码
        self.icon.run()

        # 2、非阻塞运行，后台运行，在它执行后还可以继续执行代码
        # self.icon.run_detached()
        # # 在这里可以执行其他操作
        # while True:
        #     print("后台任务在运行...")
        #     time.sleep(2)

    # 创建托盘菜单
    def create_tray_menu(self):
        # 切换按钮
        start_item = pystray.MenuItem('开始切换', self.start_changing, visible=not self.running)
        stop_item = pystray.MenuItem('停止切换', self.stop_changing, visible=self.running)

        # 添加开机启动勾选切换按钮
        # set_startup_item = pystray.MenuItem('开机启动', self.set_startup, visible=not self.startup_enabled)
        # remove_startup_item = pystray.MenuItem('取消开机启动', self.remove_startup, visible=self.startup_enabled)

        # 定义菜单
        menu_items = [
            pystray.MenuItem('选择文件夹', self.select_folder),
            pystray.MenuItem('刷新文件夹', self.refresh_images),
            pystray.MenuItem('上一张', self.prev_wallpaper),
            pystray.MenuItem('下一张', self.next_wallpaper),
            start_item,
            stop_item,
            pystray.Menu.SEPARATOR, # 分割线
            # set_startup_item,
            # remove_startup_item,
            pystray.MenuItem('显示', self.show_window, default=True), # 默认为鼠标单击事件
            # pystray.MenuItem('单击', self.touch1, default=True, enabled=False, visible=False),
            # pystray.MenuItem('退出', lambda x: print(2)), lambda 表达式是局部的，适合于在小范围内使用，不适合复杂逻辑
            pystray.MenuItem('退出', self.exit_app)
        ]

        return pystray.Menu(*menu_items)

    # 更新托盘菜单
    def update_tray_menu(self):
        if self.icon:
            self.icon.menu = self.create_tray_menu()

    # 最小化到系统托盘
    def minimize_to_tray(self):
        if self.close_minimize_enabled:
            print("最小化到托盘")
            # 隐藏窗口
            self.hide_window()
            # 启动托盘图标
            self.start_tray_icon()
        else:
            print("不允许最小化到托盘")
            # 退出应用程序
            self.exit_app()

    # 隐藏窗口
    def hide_window(self):
        self.root.withdraw()  # 隐藏Tkinter窗口

    # 恢复窗口
    def show_window(self, icon=None, item=None):
        self.root.deiconify()  # 恢复Tkinter窗口
        if self.icon:
            self.icon.stop()  # 停止托盘图标
            self.icon = None  # 清空图标引用

    # 退出应用程序
    def exit_app(self, icon=None, item=None):
        # 停止壁纸切换线程
        if self.running:
            self.stop_changing()  # 停止切换
        # 停止托盘图标
        if self.icon:
            self.icon.stop()
        # 停止Tkinter主循环
        self.root.quit()

    # 处理窗口关闭事件
    def on_closing(self):
        if self.running:
            self.stop_changing()
        self.root.destroy()  # 关闭应用程序

# 主函数
if __name__ == "__main__":
    
    # 是否开机启动
    # if is_startup():
    #     time.sleep(60)
    #     messagebox.showinfo("提示", "程序通过开机启动运行")
    # else:
    #     messagebox.showinfo("提示", "程序通过手动启动运行")
    # messagebox.showinfo("提示", "开始准备页面与配置了，程序即将开始运行")

    # # 是否存在已运行的应用程序
    # find_react = find_window_react(app_name)
    # if find_react is None:
    #     root = tk.Tk()
    #     app = WallpaperChangerApp(root)
    #     try:
    #         # 启动Tkinter主循环
    #         root.mainloop()
    #     except KeyboardInterrupt:
    #         # print("应用程序退出")
    #         sys.exit()  # 退出程序
    # else:
    #     sys.exit()  # 退出程序

    # 是否存在已运行的应用程序
    hwnds = find_window(app_name)
    if hwnds:
        # active_window(hwnds[0])  # 如果窗口存在，激活它
        sys.exit(0)  # 退出当前程序，防止新实例启动
    else:
        root = tk.Tk()
        app = WallpaperChangerApp(root)
        try:
            # 启动Tkinter主循环
            root.mainloop()
        except KeyboardInterrupt:
            # print("应用程序退出")
            sys.exit()  # 退出程序
