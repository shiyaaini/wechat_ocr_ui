import os
import json
import time
import threading
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk
from PIL import Image, ImageTk, ImageGrab
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import re

# 导入拖放支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_SUPPORTED = True
except ImportError:
    # 如果tkinterdnd2未安装，提供降级处理
    DRAG_DROP_SUPPORTED = False
    print("警告: tkinterdnd2库未安装，拖放功能不可用。请使用pip install tkinterdnd2安装。")

from wechat_ocr.ocr_manager import OcrManager, OCR_MAX_TASK_ID

# 兼容性支持
def make_window_draggable(window):
    """使窗口支持拖放功能"""
    if DRAG_DROP_SUPPORTED:
        TkinterDnD._require(window)
    return window


def parse_dnd_file_paths(data: str):
    """解析tkinterdnd2的拖放数据为文件路径列表。
    支持两种格式：
    1) {C:/a/b.png} {C:/c/d.jpg}
    2) C:/a/b.png C:/c/d.jpg
    同时去除可能的引号。
    """
    if not data:
        return []
    # 先匹配大括号包裹的多文件
    matches = re.findall(r"\{([^}]*)\}", data)
    if matches:
        raw_paths = matches
    else:
        # 按空白分割（适配未加大括号、且路径不含空格的情况）
        raw_paths = re.split(r"\s+", data.strip())
    paths = []
    for p in raw_paths:
        p = p.strip().strip('"').strip("'")
        if p:
            paths.append(p)
    return paths


class ImageZoomWindow:
    """图片放大预览窗口"""
    def __init__(self, image_path, parent_window=None):
        self.image_path = image_path
        self.parent_window = parent_window
        
        try:
            self.original_image = Image.open(image_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片: {str(e)}")
            return
        
        # 创建窗口
        self.root = tk.Toplevel()
        self.root.title(f"图片预览 - {os.path.basename(image_path)}")
        self.root.geometry("800x600")
        self.root.resizable(True, True)  # 允许调整窗口大小
        
        # 设置为模态窗口，确保能接收所有事件
        if parent_window:
            self.root.transient(parent_window)
            self.root.grab_set()  # 添加这行，确保窗口能接收事件
        
        # 确保窗口可以正常操作
        self.root.focus_force()  # 强制获取焦点
        
        # 窗口变量
        self.zoom_factor = 1.0
        self.pan_start_x = 0
        self.pan_start_y = 0
        self.is_panning = False
        self.photo = None  # 保存PhotoImage引用
        
        self.setup_ui()
        self.center_window()
        
        # 延迟显示图片，确保窗口完全初始化
        self.root.after(100, self.display_image)
        
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', self.on_window_configure)
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # 工具栏
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=X, pady=(0, 5))
        
        # 缩放按钮
        ttk.Label(toolbar, text="缩放:").pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="放大", command=self.zoom_in).pack(side=LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="缩小", command=self.zoom_out).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="适应窗口", command=self.fit_to_window).pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="原始大小", command=self.actual_size).pack(side=LEFT, padx=(0, 10))
        
        # 缩放比例显示
        self.zoom_label = ttk.Label(toolbar, text="100%")
        self.zoom_label.pack(side=LEFT, padx=(0, 10))
        
        # 关闭按钮
        ttk.Button(toolbar, text="关闭", command=self.close_window).pack(side=RIGHT)
        
        # 图片显示区域框架
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=BOTH, expand=True)
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        
        # 垂直滚动条
        self.v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)
        
        # 水平滚动条
        self.h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)
        
        # 使用grid布局确保滚动条正确显示
        self.canvas.grid(row=0, column=0, sticky='nsew')
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # 配置grid权重
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        # 绑定事件
        self.canvas.bind('<Button-1>', self.start_pan)
        self.canvas.bind('<B1-Motion>', self.pan_image)
        self.canvas.bind('<ButtonRelease-1>', self.end_pan)
        self.canvas.bind('<MouseWheel>', self.mouse_wheel)
        self.canvas.bind('<Button-4>', self.mouse_wheel)  # Linux滚轮支持
        self.canvas.bind('<Button-5>', self.mouse_wheel)  # Linux滚轮支持
        
        # 键盘事件
        self.root.bind('<Key>', self.key_press)
        self.root.focus_set()
        
        # 让画布可以获得焦点以接收键盘事件
        self.canvas.focus_set()
    
    def on_window_configure(self, event):
        """窗口大小变化事件"""
        # 只处理窗口本身的配置变化，不处理子组件的
        if event.widget == self.root:
            # 延迟更新，避免频繁刷新
            if hasattr(self, '_configure_after_id'):
                self.root.after_cancel(self._configure_after_id)
            self._configure_after_id = self.root.after(100, self.update_scrollbars)
    
    def update_scrollbars(self):
        """更新滚动条"""
        try:
            # 更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except:
            pass
    
    def center_window(self):
        """居中显示窗口"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def display_image(self):
        """显示图片"""
        try:
            # 计算显示尺寸
            img_width, img_height = self.original_image.size
            display_width = int(img_width * self.zoom_factor)
            display_height = int(img_height * self.zoom_factor)
            
            # 缩放图片
            if self.zoom_factor != 1.0:
                display_image = self.original_image.resize(
                    (display_width, display_height), 
                    Image.Resampling.LANCZOS
                )
            else:
                display_image = self.original_image.copy()
            
            # 转换为PhotoImage
            self.photo = ImageTk.PhotoImage(display_image)
            
            # 清空画布并显示图片
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            
            # 更新滚动区域
            self.canvas.configure(scrollregion=(0, 0, display_width, display_height))
            
            # 更新缩放比例显示
            self.zoom_label.config(text=f"{int(self.zoom_factor * 100)}%")
            
        except Exception as e:
            messagebox.showerror("错误", f"显示图片失败: {str(e)}")
    
    def zoom_in(self):
        """放大"""
        self.zoom_factor = min(self.zoom_factor * 1.25, 5.0)  # 最大5倍
        self.display_image()
    
    def zoom_out(self):
        """缩小"""
        self.zoom_factor = max(self.zoom_factor / 1.25, 0.1)  # 最小0.1倍
        self.display_image()
    
    def fit_to_window(self):
        """适应窗口大小"""
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = self.original_image.size
        
        if canvas_width > 1 and canvas_height > 1:  # 确保画布已初始化
            # 计算缩放比例
            width_ratio = canvas_width / img_width
            height_ratio = canvas_height / img_height
            self.zoom_factor = min(width_ratio, height_ratio, 1.0)  # 不放大
            self.display_image()
    
    def actual_size(self):
        """原始大小"""
        self.zoom_factor = 1.0
        self.display_image()
    
    def start_pan(self, event):
        """开始拖拽"""
        self.canvas.scan_mark(event.x, event.y)
        self.is_panning = True
        self.canvas.config(cursor="fleur")
    
    def pan_image(self, event):
        """拖拽图片"""
        if self.is_panning:
            # 使用canvas的scan功能进行平滑拖拽
            self.canvas.scan_dragto(event.x, event.y, gain=1)
    
    def end_pan(self, event):
        """结束拖拽"""
        self.is_panning = False
        self.canvas.config(cursor="")
    
    def mouse_wheel(self, event):
        """鼠标滚轮缩放"""
        # Windows和macOS
        if hasattr(event, 'delta'):
            if event.delta > 0:
                self.zoom_in()
            else:
                self.zoom_out()
        # Linux
        elif event.num == 4:
            self.zoom_in()
        elif event.num == 5:
            self.zoom_out()
    
    def key_press(self, event):
        """键盘快捷键"""
        if event.keysym == 'plus' or event.keysym == 'equal':
            self.zoom_in()
        elif event.keysym == 'minus':
            self.zoom_out()
        elif event.keysym == '0':
            self.actual_size()
        elif event.keysym == 'f':
            self.fit_to_window()
        elif event.keysym == 'Escape':
            self.close_window()
    
    def close_window(self):
        """关闭窗口"""
        self.root.destroy()


class ImageRotationWindow:
    """图片旋转框选窗口"""
    def __init__(self, image_path=None, callback=None):
        self.image_path = image_path
        self.callback = callback
        self.original_image = None
        self.current_rotation = 0
        self.selection_box = None  # 单个框选区域
        self.current_box = None
        self.start_x = None
        self.start_y = None
        
        # 框选调整状态
        self.is_resizing = False
        self.resize_edge = None  # 可以是 'n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'
        self.resize_box = None
        
        # 创建窗口
        self.root = tk.Toplevel()
        self.root.title("图片旋转和框选")
        self.root.geometry("1000x700")
        self.root.transient()  # 应添加父窗口参数
        self.root.grab_set()  # 确保窗口能接收事件
        
        # 拖放支持
        try:
            make_window_draggable(self.root)
            if DRAG_DROP_SUPPORTED:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.on_drop)
        except Exception as _e:
            pass
        
        self.setup_ui()
        # 如果初始有图片则加载
        if self.image_path:
            self.load_image(self.image_path)
        else:
            self.display_image()
        self.center_window()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # 顶部工具栏
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=X, pady=(0, 10))
        
        # 打开图片按钮
        ttk.Button(toolbar, text="打开图片", command=self.open_file, bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 10))
        
        # 图片旋转工具栏
        ttk.Label(toolbar, text="图片旋转:").pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="逆时针90°", command=self.rotate_left).pack(side=LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="顺时针90°", command=self.rotate_right).pack(side=LEFT, padx=(0, 5))
        
        # 微调按钮
        ttk.Button(toolbar, text="-1°", command=self.rotate_minus_one).pack(side=LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="+1°", command=self.rotate_plus_one).pack(side=LEFT, padx=(0, 5))
        
        # 自定义角度旋转
        ttk.Label(toolbar, text="角度:").pack(side=LEFT, padx=(5, 2))
        self.angle_var = tk.StringVar(value="0")
        self.angle_entry = ttk.Entry(toolbar, textvariable=self.angle_var, width=6)
        self.angle_entry.pack(side=LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="旋转", command=self.rotate_custom).pack(side=LEFT, padx=(0, 5))
        
        ttk.Button(toolbar, text="重置", command=self.reset_rotation).pack(side=LEFT, padx=(5, 10))
        
        # 框选操作
        ttk.Label(toolbar, text="框选:").pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="清除框选", command=self.clear_selections).pack(side=LEFT, padx=(0, 5))
        
        # 提示标签
        self.info_label = ttk.Label(toolbar, text="拖拽框选文字区域，或先打开/拖入图片")
        self.info_label.pack(side=LEFT, padx=(10, 0))
        
        # 图片显示区域
        image_frame = ttk.LabelFrame(main_frame, text="图片预览 (支持拖入图片)", padding=10)
        image_frame.pack(fill=BOTH, expand=True)
        
        # 创建滚动画布
        self.canvas_frame = ttk.Frame(image_frame)
        self.canvas_frame.pack(fill=BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg='white')
        h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        # 使用grid布局
        self.canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)
        
        # 绑定鼠标事件
        self.canvas.bind('<Button-1>', self.start_selection)
        self.canvas.bind('<B1-Motion>', self.update_selection)
        self.canvas.bind('<ButtonRelease-1>', self.end_selection)
        
        # 也为canvas支持拖放
        try:
            if DRAG_DROP_SUPPORTED:
                self.canvas.drop_target_register(DND_FILES)
                self.canvas.dnd_bind('<<Drop>>', self.on_drop)
        except Exception as _e:
            pass
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(10, 0))
        
        ttk.Button(button_frame, text="确认并OCR", command=self.confirm_and_ocr, bootstyle=PRIMARY).pack(side=LEFT)
        ttk.Button(button_frame, text="取消", command=self.cancel).pack(side=RIGHT)
    
    def open_file(self):
        file_path = filedialog.askopenfilename(
            title="选择需要旋转的图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.load_image(file_path)

    def on_drop(self, event):
        try:
            paths = parse_dnd_file_paths(event.data)
            if not paths:
                return
            if len(paths) > 1:
                messagebox.showinfo("提示", "一次仅支持打开一张图片，已自动选取第一张")
            self.load_image(paths[0])
        except Exception as e:
            messagebox.showerror("错误", f"拖放图片失败: {e}")

    def load_image(self, image_path):
        try:
            self.image_path = image_path
            self.original_image = Image.open(image_path)
            self.current_rotation = 0
            self.selection_box = None
            self.angle_var.set("0")
            self.display_image()
            self.info_label.config(text=f"已加载: {os.path.basename(image_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开图片: {str(e)}")

    def display_image(self):
        """显示图片"""
        try:
            if self.original_image is None:
                # 无图时显示提示
                self.canvas.delete("all")
                self.canvas.create_text(
                    400, 250,
                    text="请点击'打开图片'或拖放图片到此处",
                    fill='gray', font=('Arial', 16)
                )
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
                return
            
            # 应用旋转
            if self.current_rotation != 0:
                rotated_image = self.original_image.rotate(-self.current_rotation, expand=True)
            else:
                rotated_image = self.original_image.copy()
            
            # 计算显示尺寸（限制最大尺寸）
            max_width, max_height = 800, 500
            img_width, img_height = rotated_image.size
            
            if img_width > max_width or img_height > max_height:
                ratio = min(max_width / img_width, max_height / img_height)
                new_width = int(img_width * ratio)
                new_height = int(img_height * ratio)
                display_image = rotated_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.scale_factor = ratio
            else:
                display_image = rotated_image
                self.scale_factor = 1.0
            
            # 转换为PhotoImage并显示
            self.photo = ImageTk.PhotoImage(display_image)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            
            # 更新画布滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            
            # 绘制十字坐标
            display_width = display_image.width
            display_height = display_image.height
            self.canvas.create_line(0, display_height/2, display_width, display_height/2, fill="red", dash=(4, 4), width=1, tags="grid")
            self.canvas.create_line(display_width/2, 0, display_width/2, display_height, fill="red", dash=(4, 4), width=1, tags="grid")
            
            # 重新绘制框选
            self.redraw_selections()
            
        except Exception as e:
            messagebox.showerror("错误", f"图片显示失败: {str(e)}")

    # 旋转/框选等操作增加防护
    def _ensure_image(self):
        if self.original_image is None:
            messagebox.showwarning("提示", "请先打开或拖入图片")
            return False
        return True

    def rotate_left(self):
        if not self._ensure_image():
            return
        self.current_rotation = (self.current_rotation - 90) % 360
        self.selection_box = None
        self.display_image()
    
    def rotate_right(self):
        if not self._ensure_image():
            return
        self.current_rotation = (self.current_rotation + 90) % 360
        self.selection_box = None
        self.display_image()
    
    def rotate_minus_one(self):
        if not self._ensure_image():
            return
        self.current_rotation = (self.current_rotation - 1) % 360
        self.selection_box = None
        self.display_image()
        self.angle_var.set(str(self.current_rotation))
    
    def rotate_plus_one(self):
        if not self._ensure_image():
            return
        self.current_rotation = (self.current_rotation + 1) % 360
        self.selection_box = None
        self.display_image()
        self.angle_var.set(str(self.current_rotation))
    
    def rotate_custom(self):
        if not self._ensure_image():
            return
        try:
            angle = float(self.angle_var.get())
            self.current_rotation = angle % 360
            self.selection_box = None
            self.display_image()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的角度数值")
    
    def reset_rotation(self):
        if not self._ensure_image():
            return
        self.current_rotation = 0
        self.selection_box = None
        self.angle_var.set("0")
        self.display_image()

    def start_selection(self, event):
        if not self._ensure_image():
            return
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        
        # 检查是否点击了已有框选的边缘（用于调整大小）
        if self.selection_box:
            box_coords = self.canvas.coords(self.selection_box['canvas_id'])
            edge = self.check_edge(self.start_x, self.start_y, box_coords)
            if edge:
                self.is_resizing = True
                self.resize_edge = edge
                return
        
        # 如果已有框选并且不是在调整大小，则删除旧的框选
        if self.selection_box and not self.is_resizing:
            self.canvas.delete(self.selection_box['canvas_id'])
            if 'text_id' in self.selection_box:
                self.canvas.delete(self.selection_box['text_id'])
        
        # 创建新的框选矩形
        if not self.is_resizing:
            self.current_box = self.canvas.create_rectangle(
                self.start_x, self.start_y, self.start_x, self.start_y,
                outline='red', width=2, tags='selection'
            )
    
    def update_selection(self, event):
        if not self._ensure_image():
            return
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        
        if self.is_resizing and self.selection_box:
            box_coords = list(self.canvas.coords(self.selection_box['canvas_id']))
            if 'n' in self.resize_edge:
                box_coords[1] = current_y
            if 's' in self.resize_edge:
                box_coords[3] = current_y
            if 'w' in self.resize_edge:
                box_coords[0] = current_x
            if 'e' in self.resize_edge:
                box_coords[2] = current_x
            self.canvas.coords(self.selection_box['canvas_id'], *box_coords)
            if 'text_id' in self.selection_box:
                label_x = (box_coords[0] + box_coords[2]) / 2
                label_y = box_coords[1] - 10
                self.canvas.coords(self.selection_box['text_id'], label_x, label_y)
        elif self.current_box:
            self.canvas.coords(self.current_box, self.start_x, self.start_y, current_x, current_y)
    
    def end_selection(self, event):
        if not self._ensure_image():
            return
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        
        if self.is_resizing and self.selection_box:
            box_coords = self.canvas.coords(self.selection_box['canvas_id'])
            real_x1 = int(min(box_coords[0], box_coords[2]) / self.scale_factor)
            real_y1 = int(min(box_coords[1], box_coords[3]) / self.scale_factor)
            real_x2 = int(max(box_coords[0], box_coords[2]) / self.scale_factor)
            real_y2 = int(max(box_coords[1], box_coords[3]) / self.scale_factor)
            self.selection_box['box'] = (real_x1, real_y1, real_x2, real_y2)
            self.is_resizing = False
            self.resize_edge = None
        elif self.current_box:
            real_x1 = int(min(self.start_x, current_x) / self.scale_factor)
            real_y1 = int(min(self.start_y, current_y) / self.scale_factor)
            real_x2 = int(max(self.start_x, current_x) / self.scale_factor)
            real_y2 = int(max(self.start_y, current_y) / self.scale_factor)
            if (real_x2 - real_x1) * (real_y2 - real_y1) > 100:
                self.selection_box = {
                    'box': (real_x1, real_y1, real_x2, real_y2),
                    'canvas_id': self.current_box,
                }
                label_x = (self.start_x + current_x) / 2
                label_y = min(self.start_y, current_y) - 10
                text_id = self.canvas.create_text(
                    label_x, label_y, text='1',
                    fill='red', font=('Arial', 12, 'bold'), tags='selection'
                )
                self.selection_box['text_id'] = text_id
            else:
                self.canvas.delete(self.current_box)
                self.selection_box = None
            self.current_box = None

    def confirm_and_ocr(self):
        """确认并执行OCR"""
        if self.original_image is None:
            messagebox.showwarning("提示", "请先打开图片")
            return
        try:
            # 应用旋转并保存
            if self.current_rotation != 0:
                rotated_image = self.original_image.rotate(-self.current_rotation, expand=True)
            else:
                rotated_image = self.original_image.copy()
            
            # 保存旋转后的图片
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            temp_path = f"temp_rotated_{timestamp}.png"
            rotated_image.save(temp_path)
            
            # 传递框选信息
            result_data = {
                'image_path': temp_path,
                'rotation': self.current_rotation,
                'selections': [self.selection_box['box']] if self.selection_box else None
            }
            
            if self.callback:
                self.callback(result_data)
            self.root.destroy()
            
        except Exception as e:
            messagebox.showerror("错误", f"处理图片失败: {str(e)}")

    def cancel(self):
        """取消操作"""
        self.root.destroy()

    def check_edge(self, x, y, box_coords):
        """检查鼠标是否在框选边缘"""
        x1, y1, x2, y2 = box_coords
        edge_threshold = 10  # 边缘检测的阈值
        
        # 确保坐标从左上到右下
        if x1 > x2:
            x1, x2 = x2, x1
        if y1 > y2:
            y1, y2 = y2, y1
        
        # 检查各边缘
        edge = ''
        if abs(y - y1) < edge_threshold and x1 - edge_threshold <= x <= x2 + edge_threshold:
            edge += 'n'  # 北/上边缘
        if abs(y - y2) < edge_threshold and x1 - edge_threshold <= x <= x2 + edge_threshold:
            edge += 's'  # 南/下边缘
        if abs(x - x1) < edge_threshold and y1 - edge_threshold <= y <= y2 + edge_threshold:
            edge += 'w'  # 西/左边缘
        if abs(x - x2) < edge_threshold and y1 - edge_threshold <= y <= y2 + edge_threshold:
            edge += 'e'  # 东/右边缘
            
        return edge
    
    def redraw_selections(self):
        """重新绘制框选区域"""
        self.canvas.delete('selection')
        if self.selection_box:
            box = self.selection_box['box']
            # 应用缩放
            x1, y1, x2, y2 = [coord * self.scale_factor for coord in box]
            
            # 重新创建矩形
            rect_id = self.canvas.create_rectangle(
                x1, y1, x2, y2, outline='red', width=2, tags='selection'
            )
            
            # 重新创建序号标签
            label_x = (x1 + x2) / 2
            label_y = y1 - 10
            text_id = self.canvas.create_text(
                label_x, label_y, text='1',
                fill='red', font=('Arial', 12, 'bold'), tags='selection'
            )
            
            self.selection_box['canvas_id'] = rect_id
            self.selection_box['text_id'] = text_id
    
    def clear_selections(self):
        """清除框选"""
        self.selection_box = None
        self.canvas.delete('selection')

    def center_window(self):
        """居中显示窗口"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')


class ScreenshotWindow:
    """截屏选择窗口"""
    def __init__(self, callback):
        self.callback = callback
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        
        # 创建全屏透明窗口
        self.root = tk.Toplevel()
        self.root.attributes('-fullscreen', True)
        self.root.attributes('-alpha', 0.3)
        self.root.attributes('-topmost', True)
        self.root.configure(bg='black')
        
        # 确保窗口接收所有事件
        self.root.grab_set()
        
        # 创建画布
        self.canvas = tk.Canvas(self.root, highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # 绑定事件
        self.canvas.bind('<Button-1>', self.start_select)
        self.canvas.bind('<B1-Motion>', self.update_select)
        self.canvas.bind('<ButtonRelease-1>', self.end_select)
        self.root.bind('<Escape>', lambda e: self.cancel())
        self.root.bind('<F11>', lambda e: self.take_full_screenshot())  # F11截取全屏
        
        # 显示提示
        self.canvas.create_text(
            self.root.winfo_screenwidth()//2, 50,
            text="拖拽选择截屏区域，单击将截取全屏，按ESC取消，F11直接截取全屏",
            fill='white', font=('Arial', 16)
        )
    
    def start_select(self, event):
        self.start_x = event.x
        self.start_y = event.y
        
    def update_select(self, event):
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, event.x, event.y,
            outline='red', width=2
        )
    
    def take_full_screenshot(self):
        """直接截取全屏"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        self.root.destroy()
        
        try:
            # 将窗口隐藏后再截屏，避免截取到窗口本身
            time.sleep(0.1)  # 给窗口时间隐藏
            screenshot = ImageGrab.grab(bbox=(0, 0, screen_width, screen_height))
            self.callback(screenshot)
        except Exception as e:
            messagebox.showerror("错误", f"全屏截图失败: {str(e)}")
            print(f"全屏截图错误: {e}")
    
    def end_select(self, event):
        if self.start_x is not None and self.start_y is not None:
            # 计算截屏区域
            x1 = min(self.start_x, event.x)
            y1 = min(self.start_y, event.y)
            x2 = max(self.start_x, event.x)
            y2 = max(self.start_y, event.y)
            
            # 获取屏幕尺寸并限制截图区域不超出屏幕
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            x1 = max(0, min(x1, screen_width-1))
            y1 = max(0, min(y1, screen_height-1))
            x2 = max(0, min(x2, screen_width-1))
            y2 = max(0, min(y2, screen_height-1))
            
            # 如果是点击操作或区域太小，截取全屏
            if abs(x2 - x1) < 5 or abs(y2 - y1) < 5:
                x1, y1 = 0, 0
                x2, y2 = screen_width, screen_height
            
            # 先隐藏窗口以避免截取到窗口本身
            self.root.withdraw()
            self.root.update_idletasks()  # 确保窗口状态更新
            time.sleep(0.1)  # 给窗口时间隐藏
            
            try:
                # 截屏
                screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                self.root.destroy()  # 确保在回调前销毁窗口
                self.callback(screenshot)
            except Exception as e:
                self.root.destroy()
                messagebox.showerror("错误", f"截图失败: {str(e)}")
                print(f"截图错误: {e}")
        else:
            self.root.destroy()
    
    def cancel(self):
        self.root.destroy()


class HistoryWindow:
    """历史记录弹窗"""
    def __init__(self, parent, history, callback):
        self.parent = parent
        self.history = history
        self.callback = callback
        
        # 创建弹窗
        self.window = ttk.Toplevel(parent.root)
        self.window.title("OCR历史记录")
        self.window.geometry("900x600")
        self.window.transient(parent.root)
        self.window.grab_set()
        
        self.setup_ui()
        self.load_history_data()
        
        # 居中显示
        self.center_window()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # 顶部搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(search_frame, text="搜索:").pack(side=LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)
        
        # 刷新按钮
        refresh_btn = ttk.Button(search_frame, text="刷新", command=self.load_history_data)
        refresh_btn.pack(side=RIGHT)
        
        # 内容区域
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=BOTH, expand=True)
        
        # 左侧历史列表
        left_frame = ttk.LabelFrame(content_frame, text="历史记录列表", padding=10)
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))
        
        # 创建Treeview显示历史记录
        columns = ('时间', '预览')
        self.history_tree = ttk.Treeview(left_frame, columns=columns, show='tree headings', height=15)
        
        # 设置列
        self.history_tree.heading('#0', text='序号')
        self.history_tree.heading('时间', text='时间')
        self.history_tree.heading('预览', text='文本预览')
        
        self.history_tree.column('#0', width=50, minwidth=50)
        self.history_tree.column('时间', width=150, minwidth=150)
        self.history_tree.column('预览', width=300, minwidth=200)
        
        # 滚动条
        tree_scrollbar = ttk.Scrollbar(left_frame, orient=VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.history_tree.pack(side=LEFT, fill=BOTH, expand=True)
        tree_scrollbar.pack(side=RIGHT, fill=Y)
        
        # 绑定选择事件
        self.history_tree.bind('<<TreeviewSelect>>', self.on_item_select)
        
        # 右侧详情面板
        right_frame = ttk.LabelFrame(content_frame, text="详细信息", padding=10)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=False, padx=(5, 0))
        right_frame.configure(width=350)
        
        # 图片预览
        image_frame = ttk.LabelFrame(right_frame, text="图片预览", padding=5)
        image_frame.pack(fill=X, pady=(0, 10))
        
        self.detail_image_label = ttk.Label(image_frame, text="选择记录查看图片", anchor=CENTER, cursor="hand2")
        self.detail_image_label.pack(fill=BOTH, expand=True)
        
        # 绑定点击事件用于放大预览
        self.detail_image_label.bind('<Button-1>', self.on_image_click)
        self.current_image_path = None
        
        # 文本内容
        text_frame = ttk.LabelFrame(right_frame, text="OCR文本", padding=5)
        text_frame.pack(fill=BOTH, expand=True)
        
        self.detail_text = ttk.Text(text_frame, wrap=WORD, height=10)
        detail_scrollbar = ttk.Scrollbar(text_frame, orient=VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scrollbar.set)
        
        self.detail_text.pack(side=LEFT, fill=BOTH, expand=True)
        detail_scrollbar.pack(side=RIGHT, fill=Y)
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(10, 0))
        
        # 使用选中项按钮
        self.use_btn = ttk.Button(
            button_frame, text="使用此记录",
            command=self.use_selected,
            bootstyle=PRIMARY,
            state=DISABLED
        )
        self.use_btn.pack(side=LEFT, padx=(0, 10))
        
        # 复制文本按钮
        self.copy_detail_btn = ttk.Button(
            button_frame, text="复制文本",
            command=self.copy_detail_text,
            bootstyle=SUCCESS,
            state=DISABLED
        )
        self.copy_detail_btn.pack(side=LEFT, padx=(0, 10))
        
        # 删除选中项按钮
        self.delete_btn = ttk.Button(
            button_frame, text="删除选中",
            command=self.delete_selected,
            bootstyle=DANGER,
            state=DISABLED
        )
        self.delete_btn.pack(side=LEFT, padx=(0, 5))
        
        # 全选按钮
        self.select_all_btn = ttk.Button(
            button_frame, text="全选",
            command=self.select_all,
            bootstyle=INFO
        )
        self.select_all_btn.pack(side=LEFT, padx=(0, 5))
        
        # 批量删除按钮
        self.delete_all_selected_btn = ttk.Button(
            button_frame, text="删除全选",
            command=self.delete_all_selected,
            bootstyle=DANGER,
            state=DISABLED
        )
        self.delete_all_selected_btn.pack(side=LEFT, padx=(0, 10))
        
        # 关闭按钮
        close_btn = ttk.Button(button_frame, text="关闭", command=self.close_window)
        close_btn.pack(side=RIGHT)
    
    def center_window(self):
        """居中显示窗口"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_history_data(self):
        """加载历史记录数据"""
        # 清空现有数据
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # 加载历史记录（倒序显示，最新的在前面）
        for i, item in enumerate(reversed(self.history)):
            preview_text = item['text'][:50] + "..." if len(item['text']) > 50 else item['text']
            preview_text = preview_text.replace('\n', ' ')  # 替换换行符
            
            self.history_tree.insert('', 'end', 
                                   text=str(len(self.history) - i),
                                   values=(item['timestamp'], preview_text),
                                   tags=(str(len(self.history) - 1 - i),))  # 使用原始索引作为tag
    
    def on_search(self, event=None):
        """搜索功能"""
        search_text = self.search_var.get().lower()
        
        # 清空现有数据
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # 过滤并显示匹配的记录
        for i, item in enumerate(reversed(self.history)):
            if (search_text in item['text'].lower() or 
                search_text in item['timestamp'].lower()):
                
                preview_text = item['text'][:50] + "..." if len(item['text']) > 50 else item['text']
                preview_text = preview_text.replace('\n', ' ')
                
                self.history_tree.insert('', 'end',
                                       text=str(len(self.history) - i),
                                       values=(item['timestamp'], preview_text),
                                       tags=(str(len(self.history) - 1 - i),))
    
    def on_item_select(self, event):
        """选择项目事件"""
        selection = self.history_tree.selection()
        if selection:
            # 如果只选择了一个项目，显示详情
            if len(selection) == 1:
                item_id = selection[0]
                tags = self.history_tree.item(item_id, 'tags')
                if tags:
                    index = int(tags[0])
                    self.show_detail(index)
                    
                    # 启用单项操作按钮
                    self.use_btn.config(state=NORMAL)
                    self.copy_detail_btn.config(state=NORMAL)
                    self.delete_btn.config(state=NORMAL)
            else:
                # 多选时清空详情显示
                self.detail_text.delete(1.0, tk.END)
                self.detail_image_label.configure(image="", text="已选择多条记录")
                self.detail_image_label.image = None
                
                # 禁用单项操作按钮
                self.use_btn.config(state=DISABLED)
                self.copy_detail_btn.config(state=DISABLED)
                self.delete_btn.config(state=DISABLED)
            
            # 启用批量删除按钮（有选择项时）
            self.delete_all_selected_btn.config(state=NORMAL)
        else:
            # 没有选择时禁用所有按钮
            self.use_btn.config(state=DISABLED)
            self.copy_detail_btn.config(state=DISABLED)
            self.delete_btn.config(state=DISABLED)
            self.delete_all_selected_btn.config(state=DISABLED)
    
    def show_detail(self, index):
        """显示详细信息"""
        if 0 <= index < len(self.history):
            item = self.history[index]
            
            # 显示文本
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(1.0, item['text'])
            
            # 显示图片
            if 'image_path' in item and os.path.exists(item['image_path']):
                self.display_detail_image(item['image_path'])
                self.current_image_path = item['image_path']  # 保存当前图片路径
            else:
                self.detail_image_label.configure(image="", text="图片文件不存在")
                self.detail_image_label.image = None
                self.current_image_path = None
    
    def display_detail_image(self, image_path):
        """显示详情图片"""
        try:
            image = Image.open(image_path)
            
            # 计算缩放比例
            max_width = 300
            max_height = 150
            
            orig_width, orig_height = image.size
            width_ratio = max_width / orig_width
            height_ratio = max_height / orig_height
            scale_ratio = min(width_ratio, height_ratio, 1.0)
            
            new_width = int(orig_width * scale_ratio)
            new_height = int(orig_height * scale_ratio)
            
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            
            self.detail_image_label.configure(image=photo, text="")
            self.detail_image_label.image = photo
            
        except Exception as e:
            self.detail_image_label.configure(image="", text=f"图片加载失败: {str(e)}")
    
    def on_image_click(self, event):
        """图片点击事件 - 打开放大预览"""
        if self.current_image_path and os.path.exists(self.current_image_path):
            try:
                # 打开图片放大预览窗口
                ImageZoomWindow(self.current_image_path, self.window)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开图片预览: {str(e)}")
        else:
            messagebox.showwarning("提示", "没有可预览的图片")
    
    def display_detail_image(self, image_path):
        """显示详情图片"""
        try:
            image = Image.open(image_path)
            
            # 计算缩放比例
            max_width = 300
            max_height = 150
            
            orig_width, orig_height = image.size
            width_ratio = max_width / orig_width
            height_ratio = max_height / orig_height
            scale_ratio = min(width_ratio, height_ratio, 1.0)
            
            new_width = int(orig_width * scale_ratio)
            new_height = int(orig_height * scale_ratio)
            
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            
            self.detail_image_label.configure(image=photo, text="")
            self.detail_image_label.image = photo
            
        except Exception as e:
            self.detail_image_label.configure(image="", text=f"图片加载失败: {str(e)}")
    
    def use_selected(self):
        """使用选中的记录"""
        selection = self.history_tree.selection()
        if selection:
            item_id = selection[0]
            tags = self.history_tree.item(item_id, 'tags')
            if tags:
                index = int(tags[0])
                if 0 <= index < len(self.history):
                    item = self.history[index]
                    self.callback(item)
                    self.close_window()
    
    def copy_detail_text(self):
        """复制详情文本"""
        text = self.detail_text.get(1.0, tk.END).strip()
        if text:
            self.window.clipboard_clear()
            self.window.clipboard_append(text)
            messagebox.showinfo("提示", "文本已复制到剪贴板")
    
    def delete_selected(self):
        """删除选中的记录"""
        selection = self.history_tree.selection()
        if selection:
            if messagebox.askyesno("确认删除", "确定要删除这条记录吗？"):
                item_id = selection[0]
                tags = self.history_tree.item(item_id, 'tags')
                if tags:
                    index = int(tags[0])
                    if 0 <= index < len(self.history):
                        # 获取要删除的记录
                        record = self.history[index]
                        
                        # 删除对应的图片文件
                        if 'image_path' in record and record['image_path']:
                            try:
                                if os.path.exists(record['image_path']):
                                    os.remove(record['image_path'])
                                    print(f"已删除图片文件: {record['image_path']}")
                            except Exception as e:
                                print(f"删除图片文件失败: {e}")
                        
                        # 删除记录
                        del self.history[index]
                        self.parent.save_history()
                        
                        # 刷新显示
                        self.load_history_data()
                        
                        # 清空详情
                        self.detail_text.delete(1.0, tk.END)
                        self.detail_image_label.configure(image="", text="选择记录查看图片")
                        self.detail_image_label.image = None
                        
                        # 禁用按钮
                        self.use_btn.config(state=DISABLED)
                        self.copy_detail_btn.config(state=DISABLED)
                        self.delete_btn.config(state=DISABLED)
                        self.delete_all_selected_btn.config(state=DISABLED)
    
    def select_all(self):
        """全选所有记录"""
        # 获取所有项目
        all_items = self.history_tree.get_children()
        
        if all_items:
            # 选择所有项目
            self.history_tree.selection_set(all_items)
            
            # 启用批量删除按钮
            self.delete_all_selected_btn.config(state=NORMAL)
            
            # 显示选择数量
            messagebox.showinfo("全选", f"已选择 {len(all_items)} 条记录")
        else:
            messagebox.showinfo("提示", "没有记录可选择")
    
    def delete_all_selected(self):
        """删除所有选中的记录"""
        selected_items = self.history_tree.selection()
        
        if not selected_items:
            messagebox.showwarning("提示", "请先选择要删除的记录")
            return
        
        # 确认删除
        count = len(selected_items)
        if not messagebox.askyesno("确认批量删除", f"确定要删除选中的 {count} 条记录吗？\n同时会删除对应的图片文件。"):
            return
        
        # 收集要删除的索引（从大到小排序，避免删除时索引变化）
        indices_to_delete = []
        files_to_delete = []
        
        for item_id in selected_items:
            tags = self.history_tree.item(item_id, 'tags')
            if tags:
                index = int(tags[0])
                if 0 <= index < len(self.history):
                    indices_to_delete.append(index)
                    
                    # 收集要删除的文件
                    record = self.history[index]
                    if 'image_path' in record and record['image_path']:
                        files_to_delete.append(record['image_path'])
        
        # 按索引从大到小排序，避免删除时索引变化
        indices_to_delete.sort(reverse=True)
        
        # 删除图片文件
        deleted_files = 0
        for file_path in files_to_delete:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    deleted_files += 1
                    print(f"已删除图片文件: {file_path}")
            except Exception as e:
                print(f"删除图片文件失败: {file_path}, 错误: {e}")
        
        # 删除历史记录（从后往前删除）
        for index in indices_to_delete:
            del self.history[index]
        
        # 保存历史记录
        self.parent.save_history()
        
        # 刷新显示
        self.load_history_data()
        
        # 清空详情
        self.detail_text.delete(1.0, tk.END)
        self.detail_image_label.configure(image="", text="选择记录查看图片")
        self.detail_image_label.image = None
        
        # 禁用相关按钮
        self.use_btn.config(state=DISABLED)
        self.copy_detail_btn.config(state=DISABLED)
        self.delete_btn.config(state=DISABLED)
        self.delete_all_selected_btn.config(state=DISABLED)
        
        # 显示删除结果
        messagebox.showinfo("删除完成", f"已删除 {len(indices_to_delete)} 条记录和 {deleted_files} 个图片文件")
    
    def close_window(self):
        """关闭窗口"""
        try:
            # 如果还有任务在处理中，询问是否确认关闭
            if self.current_task or self.task_queue:
                if not messagebox.askyesno("确认", "还有任务正在处理，确定要关闭吗？"):
                    return
                
                # 如果确认关闭，先保存当前已完成的结果到历史记录
                for file_info in self.file_list:
                    if file_info['result'] and 'image_path' in file_info:
                        # 确保这个结果已经保存到历史记录
                        pass
            
            # 保存父窗口引用用于后续处理
            parent = self.parent
            parent_root = parent.root if parent else None
            
            # 重置父应用状态（直接调用，不等待窗口销毁后）
            if parent:
                parent.reset_app_state()
            
            # 清除对象引用，防止循环引用
            self.parent = None
            self.ocr_manager = None
            self.task_queue = []
            self.current_task = None
            
            # 先销毁窗口
            if hasattr(self, 'window') and self.window.winfo_exists():
                self.window.destroy()
                
            # 窗口销毁后，确保主窗口获得焦点
            if parent_root and parent_root.winfo_exists():
                # 等待事件循环处理完销毁事件
                parent_root.after(100, lambda: self._restore_main_window(parent_root))
        except Exception as e:
            print(f"关闭批量OCR窗口时出错: {e}")
            # 确保窗口被销毁
            if hasattr(self, 'window') and self.window.winfo_exists():
                self.window.destroy()
    
    def _restore_main_window(self, main_window):
        """恢复主窗口状态"""
        try:
            if main_window and main_window.winfo_exists():
                # 确保主窗口可见
                main_window.deiconify()
                # 强制获得焦点
                main_window.focus_force()
                # 提升到前面
                main_window.lift()
                # 更新窗口
                main_window.update()
        except Exception as e:
            print(f"恢复主窗口时出错: {e}")
    
    def update_progress_display(self):
        """更新进度显示"""
        # 更新总进度
        completed = sum(1 for f in self.file_list if f['status'] == '处理完成')
        total = len(self.file_list)
        percent = 0 if total == 0 else int(completed / total * 100)
        
        self.total_progress['value'] = percent
        self.progress_label.config(text=f"{completed}/{total} ({percent}%)")
        
        # 如果有完成的项，启用导出按钮
        if completed > 0:
            self.export_btn.config(state=NORMAL)


class OCRApp:
    def __init__(self):
        # 创建ttkbootstrap窗口
        self.root = ttk.Window(themename="cosmo")
        
        # 应用拖放功能
        if DRAG_DROP_SUPPORTED:
            make_window_draggable(self.root)
            
        self.root.title("WeChat OCR 工具")
        self.root.geometry("800x600")
        
        # OCR配置
        self.wechat_ocr_dir = f"{os.getcwd()}\\WeChatOCR\\WeChatOCR.exe"
        self.wechat_dir = f"{os.getcwd()}\\[3.9.9.35]"
        
        # 历史记录
        self.history_file = "ocr_history.json"
        self.history = self.load_history()
        
        # OCR管理器
        self.ocr_manager = None
        self.ocr_running = False
        
        # 批量OCR相关
        self.is_batch_ocr = False
        self.current_batch_file = None
        self.batch_window = None
        
        # 确保files目录存在
        self.ensure_files_directory()
        
        self.setup_ui()
        self.setup_ocr()
        
        # 设置拖放功能
        self.setup_drag_drop()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # 顶部按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(0, 10))
        
        # 截屏按钮
        self.screenshot_btn = ttk.Button(
            button_frame, text="截屏OCR", 
            command=self.screenshot_ocr,
            bootstyle=PRIMARY
        )
        self.screenshot_btn.pack(side=LEFT, padx=(0, 10))
        
        # 选择文件按钮
        self.file_btn = ttk.Button(
            button_frame, text="选择图片",
            command=self.select_file,
            bootstyle=SECONDARY
        )
        self.file_btn.pack(side=LEFT, padx=(0, 10))
        
        # 框选旋转按钮
        self.rotation_btn = ttk.Button(
            button_frame, text="框选旋转OCR",
            command=self.rotation_ocr,
            bootstyle=WARNING
        )
        self.rotation_btn.pack(side=LEFT, padx=(0, 10))
        
        # 截屏旋转按钮
        self.screenshot_rotation_btn = ttk.Button(
            button_frame, text="截屏旋转OCR",
            command=self.screenshot_rotation_ocr,
            bootstyle=WARNING
        )
        self.screenshot_rotation_btn.pack(side=LEFT, padx=(0, 10))
        
        # 批量OCR按钮
        self.batch_ocr_btn = ttk.Button(
            button_frame, text="批量OCR",
            command=self.batch_ocr,
            bootstyle=INFO
        )
        self.batch_ocr_btn.pack(side=LEFT, padx=(0, 10))
        
        # 历史记录按钮
        self.history_btn = ttk.Button(
            button_frame, text="历史记录",
            command=self.show_history_window,
            bootstyle=INFO
        )
        self.history_btn.pack(side=LEFT, padx=(0, 10))
        
        # 状态标签
        self.status_label = ttk.Label(button_frame, text="就绪")
        self.status_label.pack(side=RIGHT)
        
        # 中间内容区域 - 图片预览和OCR结果
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=BOTH, expand=True)
        
        # 图片预览区域
        image_frame = ttk.LabelFrame(content_frame, text="图片预览 (点击可放大)", padding=10)
        image_frame.pack(fill=X, pady=(0, 10))
        
        # 图片显示标签
        self.image_label = ttk.Label(image_frame, text="暂无图片", anchor=CENTER, cursor="hand2")
        self.image_label.pack(fill=BOTH, expand=True)
        
        # 绑定点击事件用于放大预览
        self.image_label.bind("<Button-1>", self.zoom_preview_image)
        
        # 当前图片路径
        self.current_image_path = None
        
        # OCR结果区域
        result_frame = ttk.LabelFrame(content_frame, text="OCR结果", padding=10)
        result_frame.pack(fill=BOTH, expand=True)
        
        # OCR结果文本框
        self.result_text = ttk.Text(result_frame, wrap=WORD)
        result_scrollbar = ttk.Scrollbar(result_frame, orient=VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        
        self.result_text.pack(side=LEFT, fill=BOTH, expand=True)
        result_scrollbar.pack(side=RIGHT, fill=Y)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=X, pady=(10, 0))
        
        # 复制按钮
        self.copy_btn = ttk.Button(
            bottom_frame, text="复制结果",
            command=self.copy_result,
            bootstyle=SUCCESS
        )
        self.copy_btn.pack(side=LEFT)
        
        # 清空历史按钮
        self.clear_btn = ttk.Button(
            bottom_frame, text="清空历史",
            command=self.clear_history,
            bootstyle=DANGER
        )
        self.clear_btn.pack(side=RIGHT)
    
    def show_history_window(self):
        """显示历史记录弹窗"""
        def on_history_select(item):
            # 在主窗口中显示选中的历史记录
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, item['text'])
            
            # 显示对应的图片
            if 'image_path' in item and os.path.exists(item['image_path']):
                self.display_image(item['image_path'])
            else:
                self.image_label.configure(image="", text="图片文件不存在")
                self.image_label.image = None
        
        # 创建历史记录弹窗
        HistoryWindow(self, self.history, on_history_select)
    
    def display_image(self, image_path):
        """显示图片预览"""
        try:
            # 打开图片
            image = Image.open(image_path)
            
            # 计算缩放比例，保持宽高比
            max_width = 400
            max_height = 200
            
            # 获取原始尺寸
            orig_width, orig_height = image.size
            
            # 计算缩放比例
            width_ratio = max_width / orig_width
            height_ratio = max_height / orig_height
            scale_ratio = min(width_ratio, height_ratio, 1.0)  # 不放大
            
            # 计算新尺寸
            new_width = int(orig_width * scale_ratio)
            new_height = int(orig_height * scale_ratio)
            
            # 缩放图片
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 转换为PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # 更新标签
            self.image_label.configure(image=photo, text="")
            self.image_label.image = photo  # 保持引用
            
            # 保存当前图片路径
            self.current_image_path = image_path
            
        except Exception as e:
            self.image_label.configure(image="", text=f"图片加载失败: {str(e)}")
            print(f"图片显示错误: {e}")
    
    def setup_ocr(self):
        """初始化OCR"""
        try:
            self.ocr_manager = OcrManager(self.wechat_dir)
            self.ocr_manager.SetExePath(self.wechat_ocr_dir)
            self.ocr_manager.SetUsrLibDir(self.wechat_dir)
            self.ocr_manager.SetOcrResultCallback(self.ocr_result_callback)
            self.ocr_manager.StartWeChatOCR()
            self.ocr_running = True
            self.status_label.config(text="OCR服务已启动")
            
            # 确保初始状态正常
            self.reset_app_state()
            
            # 如果拖放功能不可用，提示用户安装tkinterdnd2
            if not DRAG_DROP_SUPPORTED:
                self.root.after(1000, lambda: messagebox.showinfo("提示", 
                    "拖放功能不可用。如需使用拖放图片功能，请安装tkinterdnd2：\n"
                    "pip install tkinterdnd2\n"
                    "然后重启应用程序。"))
        except Exception as e:
            messagebox.showerror("错误", f"OCR初始化失败: {str(e)}")
            self.status_label.config(text="OCR服务启动失败")
    
    def screenshot_ocr(self):
        """截屏OCR"""
        if not self.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
            
        # 隐藏主窗口并确保它完全隐藏
        self.root.withdraw()
        self.root.update_idletasks()  # 确保窗口状态更新
        time.sleep(0.2)  # 给窗口足够的时间隐藏
        
        def on_screenshot(image):
            # 即使出错也确保主窗口最终会显示
            try:
                # 保存截屏到files文件夹
                screenshot_path = self.save_image_to_files(image, "screenshot")
                
                if screenshot_path:
                    # 显示图片预览
                    self.display_image(screenshot_path)
                    
                    # 执行OCR
                    self.status_label.config(text="正在识别...")
                    self.ocr_manager.DoOCRTask(screenshot_path)
                else:
                    messagebox.showerror("错误", "截屏保存失败")
            finally:
                # 无论成功与否，都显示主窗口
                self.root.deiconify()
        
        # 创建截屏选择窗口
        try:
            ScreenshotWindow(on_screenshot)
        except Exception as e:
            messagebox.showerror("截屏错误", f"创建截屏窗口失败: {str(e)}")
            # 发生错误时也要显示主窗口
            self.root.deiconify()
    
    def select_file(self):
        """选择图片文件"""
        if not self.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
            
        file_path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            # 将选择的文件复制到files文件夹
            copied_path = self.copy_file_to_files(file_path)
            
            # 显示图片预览
            self.display_image(copied_path)
            
            self.status_label.config(text="正在识别...")
            self.ocr_manager.DoOCRTask(copied_path)
    
    def rotation_ocr(self):
        """框选旋转OCR"""
        # 确保主窗口状态、焦点
        self.reset_app_state()
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except:
            pass
        
        if not self.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
        
        def on_rotation_complete(result_data):
            # 处理旋转和框选结果
            processed_path = result_data['image_path']
            selections = result_data.get('selections')
            
            # 将处理后的图片保存到files文件夹
            final_path = self.save_processed_image_to_files(processed_path)
            
            # 显示图片预览
            self.display_image(final_path)
            
            if selections:
                # 如果有框选区域，按顺序处理每个区域
                self.process_selections_ocr(final_path, selections)
            else:
                # 没有框选，直接OCR整个图片
                self.status_label.config(text="正在识别...")
                self.ocr_manager.DoOCRTask(final_path)
            
            # 清理临时文件
            try:
                if os.path.exists(processed_path) and processed_path.startswith("temp_"):
                    os.remove(processed_path)
            except:
                pass
            
            # 将主窗口置前显示
            try:
                self.root.after(100, lambda: (self.root.deiconify(), self.root.lift(), self.root.focus_force()))
            except:
                pass
        
        # 直接打开旋转/框选窗口，用户在窗口内选择或拖入图片
        ImageRotationWindow(None, on_rotation_complete)
    
    def save_processed_image_to_files(self, temp_path):
        """将处理后的图片保存到files文件夹"""
        try:
            # 获取日期文件夹路径
            date_folder = self.get_date_folder_path()
            
            # 生成文件名
            timestamp = datetime.now().strftime("%H%M%S")
            filename = f"rotated_{timestamp}.png"
            final_path = os.path.join(date_folder, filename)
            
            # 复制文件
            import shutil
            shutil.copy2(temp_path, final_path)
            
            return final_path
        except Exception as e:
            print(f"保存处理后图片失败: {e}")
            return temp_path
    
    def process_selections_ocr(self, image_path, selections):
        """处理框选区域的OCR"""
        try:
            from PIL import Image
            
            # 打开图片
            image = Image.open(image_path)
            
            # 只有一个框选区域
            if len(selections) == 1:
                x1, y1, x2, y2 = selections[0]
                # 裁剪框选区域
                cropped = image.crop((x1, y1, x2, y2))
                
                # 保存裁剪的图片
                crop_filename = f"crop_1_{datetime.now().strftime('%H%M%S')}.png"
                crop_path = os.path.join(self.get_date_folder_path(), crop_filename)
                cropped.save(crop_path)
                
                # 直接OCR这个区域
                self.status_label.config(text="正在识别框选区域...")
                self.ocr_manager.DoOCRTask(crop_path)
                
                # 保存路径用于清理
                self.crop_path = crop_path
                self.current_main_image_path = image_path
            else:
                # 无框选区域，直接OCR整个图片
                self.status_label.config(text="正在识别...")
                self.ocr_manager.DoOCRTask(image_path)
                
        except Exception as e:
            messagebox.showerror("错误", f"处理框选区域失败: {str(e)}")
            self.status_label.config(text="识别失败")
    
    def ocr_result_callback(self, img_path: str, results: dict):
        """OCR结果回调"""
        def update_ui():
            # 提取文本
            text_results = []
            if 'ocrResult' in results:
                # 按Y坐标排序，然后按X坐标排序（从上到下，从左到右）
                ocr_items = results['ocrResult']
                sorted_items = sorted(ocr_items, key=lambda x: (x.get('location', {}).get('y', 0), x.get('location', {}).get('x', 0)))
                for item in sorted_items:
                    if 'text' in item:
                        text_results.append(item['text'])
            
            ocr_text = '\n'.join(text_results)
            
            # 检查是否是批量处理模式
            if self.is_batch_ocr and self.current_batch_file:
                # 检查批量OCR窗口是否仍然存在
                if self.batch_window and hasattr(self.batch_window, 'window') and self.batch_window.window.winfo_exists():
                    try:
                        # 批量OCR结果处理
                        self.batch_window.on_file_processed(self.current_batch_file, ocr_text, img_path)
                    except Exception as e:
                        print(f"批量OCR回调处理错误: {e}")
                        # 如果出错，重置应用状态
                        self.reset_app_state()
                else:
                    # 窗口已关闭，但仍保存结果到历史记录
                    print("批量OCR窗口已关闭，仅保存结果到历史记录")
                    history_item = {
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'image_path': img_path,
                        'text': ocr_text,
                        'raw_result': {'type': 'batch_ocr', 'window_closed': True}
                    }
                    self.history.append(history_item)
                    self.save_history()
                    
                    # 窗口已关闭，确保状态重置
                    self.reset_app_state()
                
                return
            
            # 检查是否是框选区域的处理
            if hasattr(self, 'crop_path') and self.crop_path == img_path:
                # 框选OCR结果
                # 更新结果显示
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(1.0, ocr_text)
                
                # 添加到历史记录
                history_item = {
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'image_path': self.current_main_image_path,
                    'text': ocr_text,
                    'raw_result': {'type': 'selection_ocr', 'selection': True}
                }
                self.history.append(history_item)
                self.save_history()
                
                # 清理临时文件
                try:
                    if os.path.exists(self.crop_path):
                        os.remove(self.crop_path)
                except Exception as e:
                    print(f"删除临时文件失败: {e}")
                
                # 清理临时变量
                delattr(self, 'crop_path')
                if hasattr(self, 'current_main_image_path'):
                    delattr(self, 'current_main_image_path')
                
                self.status_label.config(text="框选识别完成")
            else:
                # 普通OCR结果
                # 更新结果显示
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(1.0, ocr_text)
                
                # 添加到历史记录
                history_item = {
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'image_path': img_path,
                    'text': ocr_text,
                    'raw_result': results
                }
                self.history.append(history_item)
                self.save_history()
                
                self.status_label.config(text="识别完成")
        
        # 在主线程中更新UI
        self.root.after(0, update_ui)
    
    def load_history(self):
        """加载历史记录"""
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []
    
    def save_history(self):
        """保存历史记录"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存历史记录失败: {e}")
    
    def ensure_files_directory(self):
        """确保files目录存在"""
        files_dir = "files"
        if not os.path.exists(files_dir):
            os.makedirs(files_dir)
    
    def get_date_folder_path(self):
        """获取当前日期的文件夹路径"""
        today = datetime.now().strftime("%Y-%m-%d")
        date_folder = os.path.join("files", today)
        
        # 确保日期文件夹存在
        if not os.path.exists(date_folder):
            os.makedirs(date_folder)
        
        return date_folder
    
    def save_image_to_files(self, image, image_type="screenshot"):
        """将图片保存到files文件夹中"""
        try:
            # 验证图像对象是否有效
            if image is None or not hasattr(image, 'size') or not hasattr(image, 'mode'):
                raise ValueError("无效的图像对象")
            
            # 验证图像尺寸是否合理
            if image.size[0] <= 0 or image.size[1] <= 0:
                raise ValueError(f"图像尺寸无效: {image.size}")
            
            # 获取日期文件夹路径
            date_folder = self.get_date_folder_path()
            
            # 生成文件名，增加毫秒避免重名
            timestamp = datetime.now().strftime("%H%M%S_%f")[:-3]  # 包含毫秒，截取前3位
            filename = f"{image_type}_{timestamp}.png"
            file_path = os.path.join(date_folder, filename)
            
            # 检查文件路径长度（Windows路径限制）
            if len(file_path) > 260:
                # 使用更短的文件名
                short_filename = f"{image_type}_{timestamp[:6]}.png"
                file_path = os.path.join(date_folder, short_filename)
            
            # 确保目录存在且可写
            os.makedirs(date_folder, exist_ok=True)
            
            # 安全复制图像对象以避免修改原始图像
            try:
                image_copy = image.copy()
            except Exception as e:
                print(f"复制图像失败，尝试使用原始图像: {e}")
                image_copy = image
            
            # 保存图片，使用更兼容的格式
            try:
                if image_copy.mode in ('RGBA', 'LA', 'P'):
                    # 创建RGB图像
                    rgb_image = Image.new('RGB', image_copy.size, (255, 255, 255))
                    # 适当处理调色板模式
                    if image_copy.mode == 'P':
                        try:
                            image_copy = image_copy.convert('RGBA')
                        except:
                            image_copy = image_copy.convert('RGB')
                    
                    # 使用安全的paste方法
                    try:
                        if image_copy.mode == 'RGBA' or image_copy.mode == 'LA':
                            mask = image_copy.split()[-1]
                            rgb_image.paste(image_copy, (0, 0), mask)
                        else:
                            rgb_image.paste(image_copy, (0, 0))
                    except Exception as e:
                        print(f"图像混合失败，尝试直接转换: {e}")
                        rgb_image = image_copy.convert('RGB')
                    
                    # 保存RGB图像
                    rgb_image.save(file_path, 'PNG', optimize=True)
                else:
                    image_copy.save(file_path, 'PNG', optimize=True)
            except Exception as e:
                print(f"保存图像失败，尝试其他方法: {e}")
                # 尝试直接转换为RGB再保存
                try:
                    image_copy.convert('RGB').save(file_path, 'PNG', optimize=True)
                except Exception as inner_e:
                    raise Exception(f"所有保存尝试都失败: {inner_e}")
            
            # 验证文件是否成功保存
            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                return file_path
            else:
                raise Exception("文件保存后验证失败")
                
        except Exception as e:
            print(f"保存图片到files文件夹失败: {e}")
            
            # 备用方案：保存到当前目录
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                backup_filename = f"{image_type}_{timestamp}.png"
                backup_path = os.path.join(os.getcwd(), backup_filename)
                
                # 直接尝试最简单的保存方法
                try:
                    image.convert('RGB').save(backup_path, 'PNG')
                    print(f"使用备用方案保存到: {backup_path}")
                    return backup_path
                except Exception as save_error:
                    print(f"备用保存方案也失败: {save_error}")
                    return None
                    
            except Exception as backup_error:
                print(f"备用保存方案完全失败: {backup_error}")
                return None
    
    def copy_file_to_files(self, source_path):
        """将选择的文件复制到files文件夹中"""
        try:
            # 获取日期文件夹路径
            date_folder = self.get_date_folder_path()
            
            # 获取原文件名和扩展名
            original_filename = os.path.basename(source_path)
            name, ext = os.path.splitext(original_filename)
            
            # 生成新文件名（添加时间戳避免重名）
            timestamp = datetime.now().strftime("%H%M%S")
            new_filename = f"selected_{timestamp}_{name}{ext}"
            new_path = os.path.join(date_folder, new_filename)
            
            # 复制文件
            import shutil
            shutil.copy2(source_path, new_path)
            
            return new_path
        except Exception as e:
            print(f"复制文件失败: {e}")
            return source_path  # 如果复制失败，返回原路径
    
    def copy_result(self):
        """复制结果到剪贴板"""
        text = self.result_text.get(1.0, tk.END).strip()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("提示", "结果已复制到剪贴板")
        else:
            messagebox.showwarning("提示", "没有可复制的内容")
    
    def clear_history(self):
        """清空历史记录"""
        if messagebox.askyesno("确认", "确定要清空所有历史记录吗？同时会删除对应的图片文件。"):
            # 删除所有对应的图片文件
            deleted_files = 0
            for record in self.history:
                if 'image_path' in record and record['image_path']:
                    try:
                        if os.path.exists(record['image_path']):
                            os.remove(record['image_path'])
                            deleted_files += 1
                    except Exception as e:
                        print(f"删除图片文件失败: {record['image_path']}, 错误: {e}")
            
            # 清空历史记录
            self.history.clear()
            self.save_history()
            self.result_text.delete(1.0, tk.END)
            
            # 清空图片预览
            self.image_label.configure(image="", text="暂无图片")
            self.image_label.image = None
            
            # 显示删除结果
            if deleted_files > 0:
                messagebox.showinfo("完成", f"已清空历史记录并删除了 {deleted_files} 个图片文件")
            else:
                messagebox.showinfo("完成", "已清空历史记录")
    
    def on_closing(self):
        """关闭应用"""
        if self.ocr_running and self.ocr_manager:
            self.ocr_manager.KillWeChatOCR()
        self.root.destroy()
    
    def run(self):
        """运行应用"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def zoom_preview_image(self, event=None):
        """点击放大预览图片"""
        if self.current_image_path and os.path.exists(self.current_image_path):
            try:
                # 使用ImageZoomWindow放大预览
                ImageZoomWindow(self.current_image_path, self.root)
            except Exception as e:
                messagebox.showerror("错误", f"无法预览图片: {str(e)}")
        else:
            # 如果没有图片，给出提示
            messagebox.showinfo("提示", "暂无图片可预览")

    def batch_ocr(self):
        """打开批量OCR窗口"""
        if not self.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
            
        # 创建批量OCR窗口前，确保主窗口状态正常
        self.reset_app_state()
        
        # 创建批量OCR窗口
        BatchOCRWindow(self)
    
    def reset_app_state(self):
        """重置应用状态，确保所有功能正常工作"""
        # 重置批量OCR相关状态
        self.is_batch_ocr = False
        self.current_batch_file = None
        if hasattr(self, 'batch_window') and self.batch_window is not None:
            # 尝试清理可能存在的批量窗口引用
            self.batch_window = None
        
        # 确保主窗口按钮状态正常
        if hasattr(self, 'screenshot_btn'):
            self.screenshot_btn.config(state=NORMAL)
        if hasattr(self, 'file_btn'):
            self.file_btn.config(state=NORMAL)
        if hasattr(self, 'rotation_btn'):
            self.rotation_btn.config(state=NORMAL)
        if hasattr(self, 'batch_ocr_btn'):
            self.batch_ocr_btn.config(state=NORMAL)
        if hasattr(self, 'history_btn'):
            self.history_btn.config(state=NORMAL)
        
        # 更新状态标签
        self.status_label.config(text="就绪")

    def screenshot_rotation_ocr(self):
        """截屏后打开旋转/框选窗口，再进行OCR"""
        # 确保主窗口状态、焦点
        self.reset_app_state()
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except:
            pass
        
        if not self.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
        
        # 隐藏主窗口进行截屏
        self.root.withdraw()
        self.root.update_idletasks()
        time.sleep(0.2)
        
        def on_screenshot(image):
            try:
                # 保存截屏到files
                screenshot_path = self.save_image_to_files(image, "screenshot")
                if not screenshot_path:
                    messagebox.showerror("错误", "截屏保存失败")
                    return
                
                # 定义旋转窗口完成回调
                def on_rotation_complete(result_data):
                    processed_path = result_data['image_path']
                    selections = result_data.get('selections')
                    
                    # 保存旋转后的图片到files
                    final_path = self.save_processed_image_to_files(processed_path)
                    
                    # 显示图片预览
                    self.display_image(final_path)
                    
                    if selections:
                        # 有框选时按区域识别
                        self.process_selections_ocr(final_path, selections)
                    else:
                        # 无框选识别整图
                        self.status_label.config(text="正在识别...")
                        self.ocr_manager.DoOCRTask(final_path)
                    
                    # 清理临时文件
                    try:
                        if os.path.exists(processed_path) and processed_path.startswith("temp_"):
                            os.remove(processed_path)
                    except:
                        pass
                
                # 打开旋转/框选窗口
                ImageRotationWindow(screenshot_path, on_rotation_complete)
            finally:
                # 恢复主窗口
                self.root.deiconify()
        
        # 打开截屏窗口
        try:
            ScreenshotWindow(on_screenshot)
        except Exception as e:
            messagebox.showerror("截屏错误", f"创建截屏窗口失败: {str(e)}")
            self.root.deiconify()

    def setup_drag_drop(self):
        """设置拖放功能"""
        if not DRAG_DROP_SUPPORTED:
            return
            
        try:
            # 为图片预览区域和文本结果区域添加拖放支持
            self.image_label.drop_target_register(DND_FILES)
            self.image_label.dnd_bind('<<Drop>>', self.handle_drop)
            
            self.result_text.drop_target_register(DND_FILES)
            self.result_text.dnd_bind('<<Drop>>', self.handle_drop)
            
            # 不在root上注册拖放，部分环境不支持
            # self.root.drop_target_register(DND_FILES)
            # self.root.dnd_bind('<<Drop>>', self.handle_drop)
            
            print("主窗口已启用控件级拖放功能")
        except Exception as e:
            print(f"设置主窗口拖放功能失败: {e}")
    
    def handle_drop(self, event):
        """处理拖放的文件"""
        try:
            file_paths = event.data
            print(f"主窗口收到拖放数据: {file_paths}")
            
            # 解析拖放的文件路径
            paths = parse_dnd_file_paths(file_paths)
            
            # 过滤有效的图片文件
            valid_paths = [p for p in paths if self.is_valid_image(p)]
            
            # 处理拖放的文件
            if len(valid_paths) == 1:
                self.process_dropped_image(valid_paths[0])
            elif len(valid_paths) > 1:
                self.handle_multiple_dropped_images(valid_paths)
            else:
                messagebox.showwarning("警告", "未找到有效的图片文件")
                print("未找到有效的图片文件")
        except Exception as e:
            print(f"处理拖放文件时出错: {e}")
            messagebox.showerror("错误", f"处理拖放文件时出错: {e}")
    
    def is_valid_image(self, file_path):
        """检查是否是有效的图片文件"""
        valid_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']
        _, ext = os.path.splitext(file_path.lower())
        return os.path.isfile(file_path) and ext in valid_extensions
    
    def process_dropped_image(self, image_path):
        """处理拖放的单个图片"""
        try:
            # 复制到files文件夹
            copied_path = self.copy_file_to_files(image_path)
            
            # 显示预览
            self.display_image(copied_path)
            
            # 执行OCR
            self.status_label.config(text="正在识别...")
            self.ocr_manager.DoOCRTask(copied_path)
            
        except Exception as e:
            messagebox.showerror("错误", f"处理拖放图片时出错: {str(e)}")
    
    def handle_multiple_dropped_images(self, image_paths):
        """处理拖放的多个图片"""
        options = ["单个处理", "批量处理", "取消"]
        
        choice = messagebox.askquestion(
            "多图片处理", 
            f"检测到{len(image_paths)}个图片文件，您想要如何处理？",
            type=messagebox.YESNOCANCEL,
            default=messagebox.YES
        )
        
        if choice == messagebox.YES:  # 单个处理
            self.process_dropped_image(image_paths[0])
        elif choice == messagebox.NO:  # 批量处理
            # 打开批量OCR窗口并添加这些图片
            self.batch_ocr()
            
            # 将图片添加到批量窗口
            if hasattr(self, 'batch_window') and self.batch_window:
                for path in image_paths:
                    self.batch_window.add_single_file(path)
        # CANCEL时不做任何处理


class BatchOCRWindow:
    """批量OCR窗口"""
    def __init__(self, parent):
        self.parent = parent
        self.ocr_manager = parent.ocr_manager
        
        # 创建窗口
        self.window = ttk.Toplevel(parent.root)
        
        # 应用拖放功能
        if DRAG_DROP_SUPPORTED:
            make_window_draggable(self.window)
            
        self.window.title("批量OCR处理")
        self.window.geometry("900x600")
        # 不使用transient，以避免最小化后无法激活的问题
        # self.window.transient(parent.root)
        # 使用更安全的窗口管理方式
        self.window.attributes('-topmost', True)  # 临时设置为顶层窗口
        self.window.update()  # 更新窗口状态
        self.window.attributes('-topmost', False)  # 取消顶层窗口设置
        
        # 不使用grab_set，避免窗口交互问题
        # self.window.grab_set()
        
        # 保存引用以在进程完成时更新窗口
        parent.batch_window = self
        
        # 绑定窗口图标化(最小化)和取消图标化事件
        self.window.bind("<Unmap>", self.on_window_minimize)
        self.window.bind("<Map>", self.on_window_restore)
        
        # 文件列表
        self.file_list = []
        
        # 结果字典
        self.results = {}
        
        # 任务队列
        self.task_queue = []
        
        # 当前正在处理的任务
        self.current_task = None
        
        # after调度ID（用于关闭时取消）
        self._after_id = None
        
        # 设置UI
        self.setup_ui()
        self.center_window()
        
        # 设置拖放支持
        self.setup_drag_drop()
    
    def on_window_minimize(self, event):
        """窗口最小化时的处理"""
        # 清除可能导致问题的状态
        if self.window.winfo_exists():
            try:
                # 确保窗口不会卡在grab状态
                self.window.grab_release()
            except:
                pass
    
    def on_window_restore(self, event):
        """窗口恢复时的处理"""
        if self.window.winfo_exists():
            # 确保窗口能获得焦点
            self.window.focus_force()
            self.window.lift()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # 顶部按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=(0, 10))
        
        # 添加文件按钮
        self.add_files_btn = ttk.Button(
            button_frame, text="添加文件",
            command=self.add_files,
            bootstyle=PRIMARY
        )
        self.add_files_btn.pack(side=LEFT, padx=(0, 10))
        
        # 添加文件夹按钮
        self.add_folder_btn = ttk.Button(
            button_frame, text="添加文件夹",
            command=self.add_folder,
            bootstyle=SECONDARY
        )
        self.add_folder_btn.pack(side=LEFT, padx=(0, 10))
        
        # 清空列表按钮
        self.clear_btn = ttk.Button(
            button_frame, text="清空列表",
            command=self.clear_files,
            bootstyle=DANGER
        )
        self.clear_btn.pack(side=LEFT, padx=(0, 10))
        
        # 开始处理按钮
        self.start_btn = ttk.Button(
            button_frame, text="开始批量OCR",
            command=self.start_batch_ocr,
            bootstyle=SUCCESS
        )
        self.start_btn.pack(side=RIGHT)
        
        # 创建中间区域分割窗格
        paned_window = ttk.PanedWindow(main_frame, orient=HORIZONTAL)
        paned_window.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # 文件列表区域
        list_frame = ttk.LabelFrame(paned_window, text="文件列表", padding=10)
        paned_window.add(list_frame, weight=60)
        
        # 图片预览和结果区域
        preview_frame = ttk.Frame(paned_window)
        paned_window.add(preview_frame, weight=40)
        
        # 创建Treeview显示文件列表
        columns = ('文件名', '状态', '处理进度')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        # 设置列
        self.file_tree.heading('文件名', text='文件名')
        self.file_tree.heading('状态', text='状态')
        self.file_tree.heading('处理进度', text='处理进度')
        
        self.file_tree.column('文件名', width=400, minwidth=200)
        self.file_tree.column('状态', width=100, minwidth=80)
        self.file_tree.column('处理进度', width=100, minwidth=80)
        
        # 滚动条
        tree_scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.file_tree.pack(side=LEFT, fill=BOTH, expand=True)
        tree_scrollbar.pack(side=RIGHT, fill=Y)
        
        # 绑定选择事件
        self.file_tree.bind('<Double-1>', self.view_result)
        self.file_tree.bind('<<TreeviewSelect>>', self.preview_selected_image)
        
        # 图片预览区域
        image_frame = ttk.LabelFrame(preview_frame, text="图片预览", padding=10)
        image_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # 图片预览标签
        self.preview_label = ttk.Label(image_frame, text="选择图片查看预览", anchor=CENTER, cursor="hand2")
        self.preview_label.pack(fill=BOTH, expand=True)
        
        # 绑定点击放大事件
        self.preview_label.bind("<Button-1>", self.zoom_preview_image)
        self.current_preview_path = None
        
        # 处理进度区域
        progress_frame = ttk.LabelFrame(main_frame, text="处理进度", padding=10)
        progress_frame.pack(fill=X, pady=(0, 10))
        
        # 总进度条
        ttk.Label(progress_frame, text="总进度:").pack(side=LEFT, padx=(0, 5))
        self.total_progress = ttk.Progressbar(progress_frame, length=500, mode='determinate')
        self.total_progress.pack(side=LEFT, fill=X, expand=True)
        
        # 进度信息
        self.progress_label = ttk.Label(progress_frame, text="0/0 (0%)")
        self.progress_label.pack(side=RIGHT, padx=(10, 0))
        
        # 结果展示区域
        result_frame = ttk.LabelFrame(main_frame, text="OCR结果预览", padding=10)
        result_frame.pack(fill=BOTH, expand=True)
        
        # 结果文本框
        self.result_text = ttk.Text(result_frame, wrap=WORD, height=8)
        result_scrollbar = ttk.Scrollbar(result_frame, orient=VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        
        self.result_text.pack(side=LEFT, fill=BOTH, expand=True)
        result_scrollbar.pack(side=RIGHT, fill=Y)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=X, pady=(10, 0))
        
        # 导出结果按钮
        self.export_btn = ttk.Button(
            bottom_frame, text="导出结果",
            command=self.export_results,
            bootstyle=SUCCESS,
            state=DISABLED
        )
        self.export_btn.pack(side=LEFT)
        
        # 关闭按钮
        self.close_btn = ttk.Button(
            bottom_frame, text="关闭",
            command=self.close_window
        )
        self.close_btn.pack(side=RIGHT)
        
    def preview_selected_image(self, event=None):
        """预览选中的图片"""
        selection = self.file_tree.selection()
        if selection:
            item_id = selection[0]
            item_index = self.file_tree.index(item_id)
            
            if item_index < len(self.file_list):
                file_info = self.file_list[item_index]
                self.display_preview_image(file_info['path'])
    
    def display_preview_image(self, image_path):
        """显示预览图片"""
        try:
            # 打开图片
            image = Image.open(image_path)
            
            # 计算缩放比例，保持宽高比
            max_width = 300
            max_height = 200
            
            # 获取原始尺寸
            orig_width, orig_height = image.size
            
            # 计算缩放比例
            width_ratio = max_width / orig_width
            height_ratio = max_height / orig_height
            scale_ratio = min(width_ratio, height_ratio, 1.0)  # 不放大
            
            # 计算新尺寸
            new_width = int(orig_width * scale_ratio)
            new_height = int(orig_height * scale_ratio)
            
            # 缩放图片
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # 转换为PhotoImage
            photo = ImageTk.PhotoImage(image)
            
            # 更新标签
            self.preview_label.configure(image=photo, text="")
            self.preview_label.image = photo  # 保持引用
            
            # 保存当前图片路径
            self.current_preview_path = image_path
            
        except Exception as e:
            self.preview_label.configure(image="", text=f"图片加载失败: {str(e)}")
            print(f"预览图片显示错误: {e}")
    
    def zoom_preview_image(self, event=None):
        """点击放大预览图片"""
        if self.current_preview_path and os.path.exists(self.current_preview_path):
            try:
                # 使用ImageZoomWindow放大预览
                ImageZoomWindow(self.current_preview_path, self.window)
            except Exception as e:
                messagebox.showerror("错误", f"无法预览图片: {str(e)}")
        else:
            # 如果没有图片，给出提示
            messagebox.showinfo("提示", "暂无图片可预览")
    
    def center_window(self):
        """居中显示窗口"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
    
    def add_files(self):
        """添加文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择图片文件",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_paths:
            added_count = 0
            for path in file_paths:
                if self.add_single_file(path):
                    added_count += 1
            
            # 提示添加结果
            if added_count > 0:
                messagebox.showinfo("添加完成", f"已添加{added_count}个图片文件")
            else:
                messagebox.showinfo("提示", "未添加任何新文件（可能已存在）")
    
    def add_folder(self):
        """添加文件夹中的所有图片"""
        folder_path = filedialog.askdirectory(title="选择包含图片的文件夹")
        
        if folder_path:
            # 支持的图片格式
            img_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
            
            # 遍历文件夹
            added_count = 0
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith(img_extensions):
                        full_path = os.path.join(root, file)
                        if self.add_single_file(full_path):
                            added_count += 1
            
            # 提示添加结果
            if added_count > 0:
                messagebox.showinfo("添加完成", f"已从文件夹添加{added_count}个图片文件")
            else:
                messagebox.showinfo("提示", "未添加任何新文件（可能已存在或文件夹中无图片）")
    
    def clear_files(self):
        """清空文件列表"""
        if messagebox.askyesno("确认", "确定要清空文件列表吗？"):
            self.file_list.clear()
            
            # 清空Treeview
            for item in self.file_tree.get_children():
                self.file_tree.delete(item)
            
            # 重置进度
            self.total_progress['value'] = 0
            self.progress_label.config(text="0/0 (0%)")
            self.result_text.delete(1.0, tk.END)
            
            # 禁用导出按钮
            self.export_btn.config(state=DISABLED)
    
    def start_batch_ocr(self):
        """开始批量OCR处理"""
        if not self.file_list:
            messagebox.showwarning("警告", "请先添加文件")
            return
        
        # 检查OCR服务是否启动
        if not self.parent.ocr_running:
            messagebox.showerror("错误", "OCR服务未启动")
            return
        
        # 禁用按钮，防止重复点击
        self.add_files_btn.config(state=DISABLED)
        self.add_folder_btn.config(state=DISABLED)
        self.clear_btn.config(state=DISABLED)
        self.start_btn.config(state=DISABLED)
        
        # 重置任务队列
        self.task_queue = [f for f in self.file_list if f['status'] != '处理完成']
        
        # 更新UI状态
        for i, file_info in enumerate(self.file_list):
            if file_info['status'] != '处理完成':
                file_info['status'] = '等待处理'
                file_info['progress'] = '0%'
                
                # 更新Treeview
                item_id = self.file_tree.get_children()[i]
                self.file_tree.item(item_id, values=(file_info['name'], file_info['status'], file_info['progress']))
        
        # 重置总进度
        self.total_progress['value'] = 0
        completed = len(self.file_list) - len(self.task_queue)
        total = len(self.file_list)
        percent = 0 if total == 0 else int(completed / total * 100)
        self.progress_label.config(text=f"{completed}/{total} ({percent}%)")
        
        # 开始处理第一个文件
        self.process_next_file()
    
    def process_next_file(self):
        """处理下一个文件"""
        if not self.task_queue:
            # 所有任务处理完成
            self.on_batch_complete()
            return
        
        # 获取下一个文件
        file_info = self.task_queue.pop(0)
        self.current_task = file_info
        
        # 更新状态
        file_info['status'] = '处理中'
        
        # 更新Treeview
        index = self.file_list.index(file_info)
        item_id = self.file_tree.get_children()[index]
        self.file_tree.item(item_id, values=(file_info['name'], file_info['status'], file_info['progress']))
        
        # 复制到工作目录
        copied_path = self.parent.copy_file_to_files(file_info['path'])
        
        # 设置回调标记
        self.parent.is_batch_ocr = True
        self.parent.current_batch_file = file_info
        
        # 执行OCR
        self.parent.ocr_manager.DoOCRTask(copied_path)
    
    def on_file_processed(self, file_info, ocr_result, image_path):
        """文件处理完成的回调"""
        # 检查窗口和组件是否仍然存在
        if not self.window.winfo_exists():
            print("批量OCR窗口已关闭，取消更新UI")
            return
            
        try:
            # 更新文件状态
            file_info['status'] = '处理完成'
            file_info['progress'] = '100%'
            file_info['result'] = ocr_result
            file_info['image_path'] = image_path
            
            # 更新Treeview
            index = self.file_list.index(file_info)
            # 检查Treeview是否仍然存在
            if hasattr(self, 'file_tree') and self.file_tree.winfo_exists():
                children = self.file_tree.get_children()
                if index < len(children):
                    item_id = children[index]
                    self.file_tree.item(item_id, values=(file_info['name'], file_info['status'], file_info['progress']))
                    
                    # 选中当前处理的项
                    self.file_tree.selection_set(item_id)
                    self.file_tree.see(item_id)
            
            # 更新总进度
            completed = sum(1 for f in self.file_list if f['status'] == '处理完成')
            total = len(self.file_list)
            percent = int(completed / total * 100)
            
            # 检查进度条是否仍然存在
            if hasattr(self, 'total_progress') and self.total_progress.winfo_exists():
                self.total_progress['value'] = percent
                
            # 检查标签是否仍然存在
            if hasattr(self, 'progress_label') and self.progress_label.winfo_exists():
                self.progress_label.config(text=f"{completed}/{total} ({percent}%)")
            
            # 显示结果预览
            if hasattr(self, 'result_text') and self.result_text.winfo_exists():
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(1.0, ocr_result)
            
            # 显示图片预览
            if hasattr(self, 'preview_label') and self.preview_label.winfo_exists():
                self.display_preview_image(file_info['path'])
            
            # 保存到历史记录
            history_item = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'image_path': image_path,
                'text': ocr_result,
                'raw_result': {'type': 'batch_ocr', 'original_path': file_info['path']}
            }
            self.parent.history.append(history_item)
            self.parent.save_history()
            
            # 处理下一个文件（使用受控after）
            self.current_task = None
            if self.window.winfo_exists():
                if self._after_id:
                    try:
                        self.window.after_cancel(self._after_id)
                    except:
                        pass
                self._after_id = self.window.after(200, self.process_next_file)
                
        except Exception as e:
            print(f"批量OCR处理文件时发生错误: {e}")
            # 尝试继续处理下一个文件
            self.current_task = None
            if self.window.winfo_exists():
                if self._after_id:
                    try:
                        self.window.after_cancel(self._after_id)
                    except:
                        pass
                self._after_id = self.window.after(200, self.process_next_file)
    
    def on_batch_complete(self):
        """批量处理完成"""
        # 启用按钮
        self.add_files_btn.config(state=NORMAL)
        self.add_folder_btn.config(state=NORMAL)
        self.clear_btn.config(state=NORMAL)
        self.start_btn.config(state=NORMAL)
        self.export_btn.config(state=NORMAL)
        
        # 清理父应用批量状态
        if self.parent:
            self.parent.is_batch_ocr = False
            self.parent.current_batch_file = None
        
        # 取消pending after
        if self._after_id:
            try:
                self.window.after_cancel(self._after_id)
            except:
                pass
            self._after_id = None
        
        # 显示完成消息
        messagebox.showinfo("完成", "批量OCR处理已完成")
    
    def view_result(self, event):
        """双击查看结果"""
        selection = self.file_tree.selection()
        if selection:
            item_id = selection[0]
            item_index = self.file_tree.index(item_id)
            
            if item_index < len(self.file_list):
                file_info = self.file_list[item_index]
                
                if file_info['result']:
                    # 显示结果
                    self.result_text.delete(1.0, tk.END)
                    self.result_text.insert(1.0, file_info['result'])
                    
                    # 显示图片
                    if 'image_path' in file_info and os.path.exists(file_info['image_path']):
                        self.parent.display_image(file_info['image_path'])
    
    def export_results(self):
        """导出OCR结果"""
        if not self.file_list or not any(f['result'] for f in self.file_list):
            messagebox.showwarning("警告", "没有可导出的结果")
            return
        
        # 选择保存位置
        save_path = filedialog.asksaveasfilename(
            title="导出OCR结果",
            defaultextension=".txt",
            filetypes=[
                ("文本文件", "*.txt"),
                ("CSV文件", "*.csv")
            ]
        )
        
        if save_path:
            try:
                with open(save_path, 'w', encoding='utf-8') as f:
                    if save_path.endswith('.csv'):
                        # CSV格式导出
                        import csv
                        writer = csv.writer(f)
                        writer.writerow(['文件名', 'OCR结果'])
                        
                        for file_info in self.file_list:
                            if file_info['result']:
                                writer.writerow([file_info['name'], file_info['result'].replace('\n', ' ')])
                    else:
                        # 文本格式导出
                        for file_info in self.file_list:
                            if file_info['result']:
                                f.write(f"===== {file_info['name']} =====\n")
                                f.write(file_info['result'])
                                f.write("\n\n")
                
                messagebox.showinfo("成功", f"结果已导出到 {save_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def close_window(self):
        """关闭窗口"""
        try:
            # 如果还有任务在处理中，询问是否确认关闭
            if self.current_task or self.task_queue:
                if not messagebox.askyesno("确认", "还有任务正在处理，确定要关闭吗？"):
                    return
                
                # 如果确认关闭，先保存当前已完成的结果到历史记录
                for file_info in self.file_list:
                    if file_info['result'] and 'image_path' in file_info:
                        # 确保这个结果已经保存到历史记录
                        pass
            
            # 保存父窗口引用用于后续处理
            parent = self.parent
            parent_root = parent.root if parent else None
            
            # 重置父应用状态（直接调用，不等待窗口销毁后）
            if parent:
                parent.reset_app_state()
                parent.is_batch_ocr = False
                parent.current_batch_file = None
                parent.batch_window = None
            
            # 取消pending after
            if self._after_id:
                try:
                    self.window.after_cancel(self._after_id)
                except:
                    pass
                self._after_id = None
            
            # 清除对象引用，防止循环引用
            self.parent = None
            self.ocr_manager = None
            self.task_queue = []
            self.current_task = None
            
            # 先销毁窗口
            if hasattr(self, 'window') and self.window.winfo_exists():
                try:
                    self.window.grab_release()
                except:
                    pass
                self.window.destroy()
                
            # 窗口销毁后，确保主窗口获得焦点
            if parent_root and parent_root.winfo_exists():
                # 等待事件循环处理完销毁事件
                parent_root.after(100, lambda: self._restore_main_window(parent_root))
        except Exception as e:
            print(f"关闭批量OCR窗口时出错: {e}")
            # 确保窗口被销毁
            if hasattr(self, 'window') and self.window.winfo_exists():
                try:
                    self.window.grab_release()
                except:
                    pass
                self.window.destroy()
    
    def _restore_main_window(self, main_window):
        """恢复主窗口状态"""
        try:
            if main_window and main_window.winfo_exists():
                # 确保主窗口可见
                main_window.deiconify()
                # 强制获得焦点
                main_window.focus_force()
                # 提升到前面
                main_window.lift()
                # 更新窗口
                main_window.update()
        except Exception as e:
            print(f"恢复主窗口时出错: {e}")
    
    def update_progress_display(self):
        """更新进度显示"""
        # 更新总进度
        completed = sum(1 for f in self.file_list if f['status'] == '处理完成')
        total = len(self.file_list)
        percent = 0 if total == 0 else int(completed / total * 100)
        
        self.total_progress['value'] = percent
        self.progress_label.config(text=f"{completed}/{total} ({percent}%)")
        
        # 如果有完成的项，启用导出按钮
        if completed > 0:
            self.export_btn.config(state=NORMAL)

    def add_single_file(self, path):
        """添加单个文件到列表"""
        # 检查是否已添加
        if path not in [f['path'] for f in self.file_list]:
            file_info = {
                'path': path,
                'name': os.path.basename(path),
                'status': '等待处理',
                'progress': '0%',
                'result': None
            }
            self.file_list.append(file_info)
            
            # 添加到Treeview
            self.file_tree.insert('', 'end', values=(file_info['name'], file_info['status'], file_info['progress']))
            
            # 更新UI
            self.update_progress_display()
            
            # 选择并预览新添加的项
            last_item = self.file_tree.get_children()[-1]
            self.file_tree.selection_set(last_item)
            self.file_tree.focus(last_item)
            self.file_tree.see(last_item)
            
            # 显示预览
            self.display_preview_image(path)
            
            return True
        return False

    def setup_drag_drop(self):
        """设置拖放功能"""
        if not DRAG_DROP_SUPPORTED:
            return
        
        try:
            # 为文件树、预览区域和结果区域添加拖放支持
            self.file_tree.drop_target_register(DND_FILES)
            self.file_tree.dnd_bind('<<Drop>>', self.handle_drop)
            
            self.preview_label.drop_target_register(DND_FILES)
            self.preview_label.dnd_bind('<<Drop>>', self.handle_drop)
            
            self.result_text.drop_target_register(DND_FILES)
            self.result_text.dnd_bind('<<Drop>>', self.handle_drop)
            
            # 为整个窗口添加拖放支持
            self.window.drop_target_register(DND_FILES)
            self.window.dnd_bind('<<Drop>>', self.handle_drop)
            
            print("批量OCR窗口已启用拖放功能")
        except Exception as e:
            print(f"设置批量OCR窗口拖放功能失败: {e}")
    
    def handle_drop(self, event):
        """处理拖放的文件"""
        try:
            file_paths = event.data
            print(f"批量窗口收到拖放数据: {file_paths}")
            
            # 解析拖放的文件路径
            paths = parse_dnd_file_paths(file_paths)
            
            # 过滤有效的图片文件
            valid_paths = [p for p in paths if self.is_valid_image(p)]
            
            # 添加到列表
            if valid_paths:
                added_count = 0
                for path in valid_paths:
                    if self.add_single_file(path):
                        added_count += 1
                
                if added_count > 0:
                    messagebox.showinfo("成功", f"已添加{added_count}个图片文件")
                    print(f"成功添加{added_count}个图片文件")
                else:
                    messagebox.showinfo("提示", "未添加任何新文件（可能已存在）")
            else:
                messagebox.showwarning("警告", "未找到有效的图片文件")
                print("未找到有效的图片文件")
        except Exception as e:
            print(f"处理拖放文件时出错: {e}")
            messagebox.showerror("错误", f"处理拖放文件时出错: {e}")

    def is_valid_image(self, file_path):
        """检查是否是有效的图片文件"""
        valid_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']
        try:
            # 确保路径有效
            if not file_path or not os.path.exists(file_path):
                return False
            
            # 检查文件扩展名
            _, ext = os.path.splitext(file_path.lower())
            return os.path.isfile(file_path) and ext in valid_extensions
        except Exception as e:
            print(f"检查图片文件有效性时出错: {e}")
            return False


if __name__ == "__main__":
    app = OCRApp()
    app.run()