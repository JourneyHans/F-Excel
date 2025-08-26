#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 基础模块抽象类

该模块定义了所有功能模块的通用接口和基础功能，
采用抽象基类设计模式，确保所有模块都遵循统一的接口规范。

设计模式:
- 抽象基类模式
- 模板方法模式
- 策略模式
"""

import tkinter as tk
from tkinter import ttk, messagebox
from abc import ABC, abstractmethod
import os
from typing import Optional, Dict, Any, List
from datetime import datetime


class BaseModule(ABC):
    """
    基础模块抽象类
    
    所有功能模块都应该继承此类，并实现必要的抽象方法。
    该类提供了通用的UI创建、文件处理、错误处理等功能。
    """
    
    def __init__(self):
        """初始化基础模块"""
        self.window: Optional[tk.Toplevel] = None
        self.input_text: Optional[tk.Text] = None
        self.output_text: Optional[tk.Text] = None
        self.uploaded_file_path: Optional[str] = None
        self.uploaded_file_name: Optional[str] = None
        self.status_bar: Optional[ttk.Label] = None
        self.progress_var: Optional[tk.DoubleVar] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.status_label: Optional[ttk.Label] = None
        
        # 模块配置
        self.module_config = self.get_module_config()
        
    @abstractmethod
    def get_module_config(self) -> Dict[str, Any]:
        """
        获取模块配置信息
        
        Returns:
            Dict[str, Any]: 模块配置字典，包含名称、描述、图标等信息
        """
        pass
    
    @abstractmethod
    def create_interface(self) -> None:
        """创建模块界面"""
        pass
    
    @abstractmethod
    def process_data(self, input_data: str) -> List[Dict[str, Any]]:
        """
        处理输入数据
        
        Args:
            input_data (str): 输入数据
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        pass
    
    def show(self) -> None:
        """
        显示模块窗口
        
        如果窗口不存在或已关闭，则创建新窗口；
        如果窗口已存在，则将其提升到前台。
        """
        if self.window is None or not self.window.winfo_exists():
            self.create_window()
        else:
            self.window.lift()
            self.window.focus()
    
    def create_window(self) -> None:
        """创建模块主窗口"""
        self.window = tk.Toplevel()
        self.window.title(f"{self.module_config['name']} - F-Excel")
        self.window.geometry(self.module_config.get('window_size', '900x700'))
        self.window.minsize(800, 600)
        
        # 创建界面
        self.create_interface()
        
        # 窗口居中
        self.center_window()
        
        # 绑定窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
    
    def center_window(self) -> None:
        """将窗口居中显示"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def on_window_close(self) -> None:
        """窗口关闭事件处理"""
        if hasattr(self, 'processing') and self.processing:
            if messagebox.askyesno("确认", "正在处理中，确定要关闭吗？"):
                self.processing = False
                self.window.destroy()
        else:
            self.window.destroy()
    
    def create_title_section(self, parent: tk.Widget) -> None:
        """
        创建标题区域
        
        Args:
            parent (tk.Widget): 父容器
        """
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill='x', padx=20, pady=10)
        
        title_label = ttk.Label(
            title_frame, 
            text=self.module_config['name'], 
            font=('Arial', 14, 'bold')
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            title_frame, 
            text=self.module_config['description'], 
            font=('Arial', 10)
        )
        subtitle_label.pack(pady=5)
    
    def create_file_upload_section(self, parent: tk.Widget, file_types: List[str]) -> None:
        """
        创建文件上传区域
        
        Args:
            parent (tk.Widget): 父容器
            file_types (List[str]): 支持的文件类型列表
        """
        file_frame = ttk.Frame(parent)
        file_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(file_frame, text="选择文件:").pack(side='left')
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(
            file_frame, 
            textvariable=self.file_path_var, 
            width=50, 
            state='readonly'
        ).pack(side='left', padx=10)
        
        ttk.Button(
            file_frame, 
            text="浏览", 
            command=self.browse_file
        ).pack(side='left', padx=5)
        
        ttk.Button(
            file_frame, 
            text="清空文件", 
            command=self.clear_file
        ).pack(side='left', padx=5)
    
    def create_progress_section(self, parent: tk.Widget) -> None:
        """
        创建进度条区域
        
        Args:
            parent (tk.Widget): 父容器
        """
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill='x', padx=20, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100
        )
        self.progress_bar.pack(fill='x', pady=5)
        
        self.status_label = ttk.Label(
            progress_frame, 
            text="就绪", 
            font=('Arial', 9)
        )
        self.status_label.pack()
    
    def create_status_bar(self, parent: tk.Widget) -> None:
        """
        创建状态栏
        
        Args:
            parent (tk.Widget): 父容器
        """
        self.status_bar = ttk.Label(
            parent, 
            text="就绪", 
            relief='sunken', 
            anchor='w'
        )
        self.status_bar.pack(side='bottom', fill='x')
    
    def update_status(self, message: str) -> None:
        """
        更新状态栏信息
        
        Args:
            message (str): 状态信息
        """
        if self.status_bar:
            self.status_bar.config(text=message)
    
    def update_progress(self, progress: float, status: str) -> None:
        """
        更新进度条和状态
        
        Args:
            progress (float): 进度值 (0-100)
            status (str): 状态信息
        """
        if self.progress_var:
            self.progress_var.set(progress)
        if self.status_label:
            self.status_label.config(text=status)
    
    def browse_file(self) -> None:
        """浏览并选择文件（子类需要重写）"""
        pass
    
    def clear_file(self) -> None:
        """清空已选择的文件（子类需要重写）"""
        pass
    
    def load_file_content(self, file_path: str) -> None:
        """
        加载文件内容到输入框
        
        Args:
            file_path (str): 文件路径
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if self.input_text:
                    self.input_text.delete('1.0', tk.END)
                    self.input_text.insert('1.0', content)
            self.update_status(f"已加载文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def generate_default_filename(self, prefix: str = "", suffix: str = "") -> str:
        """
        生成默认文件名
        
        Args:
            prefix (str): 文件名前缀
            suffix (str): 文件名后缀
            
        Returns:
            str: 生成的默认文件名
        """
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d")
        
        if self.uploaded_file_name:
            filename = f"{date_str}_{self.uploaded_file_name}"
        else:
            filename = f"{date_str}_{self.module_config['name']}"
        
        if prefix:
            filename = f"{prefix}_{filename}"
        if suffix:
            filename = f"{filename}_{suffix}"
            
        return filename
    
    def show_error(self, title: str, message: str) -> None:
        """
        显示错误消息
        
        Args:
            title (str): 错误标题
            message (str): 错误信息
        """
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str) -> None:
        """
        显示信息消息
        
        Args:
            title (str): 信息标题
            message (str): 信息内容
        """
        messagebox.showinfo(title, message)
    
    def show_warning(self, title: str, message: str) -> None:
        """
        显示警告消息
        
        Args:
            title (str): 警告标题
            message (str): 警告信息
        """
        messagebox.showwarning(title, message)
    
    def ask_confirmation(self, title: str, message: str) -> bool:
        """
        询问用户确认
        
        Args:
            title (str): 确认标题
            message (str): 确认信息
            
        Returns:
            bool: 用户是否确认
        """
        return messagebox.askyesno(title, message)
