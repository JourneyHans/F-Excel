#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 桌面应用程序主程序

该程序是F-Excel应用的主入口，采用模块化设计，
通过工厂模式管理功能模块，使用配置驱动的方式创建界面。

设计模式:
- 工厂模式（模块管理）
- 单例模式（应用实例）
- 观察者模式（状态更新）
- 配置驱动模式（界面创建）
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
from typing import Dict, Any

# 导入配置和模块
from config import APP_NAME, APP_VERSION, APP_DESCRIPTION, STYLES
from modules import get_all_modules, create_module


class FExcelApp:
    """
    F-Excel 主应用程序类
    
    负责创建主界面、管理模块、处理用户交互等核心功能。
    采用单例模式确保只有一个应用实例。
    """
    
    _instance = None
    
    def __new__(cls, *args, **kwargs):
        """单例模式实现"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, root: tk.Tk):
        """初始化主应用程序"""
        if hasattr(self, '_initialized'):
            return
        
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # 设置应用图标和样式
        self.setup_styles()
        
        # 创建主界面
        self.create_main_interface()
        
        # 初始化模块
        self.modules: Dict[str, Any] = {}
        self.init_modules()
        
        # 标记已初始化
        self._initialized = True
        
    def setup_styles(self) -> None:
        """设置应用样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置样式
        style.configure('Title.TLabel', font=STYLES['title_font'])
        style.configure('Subtitle.TLabel', font=STYLES['subtitle_font'])
        style.configure('Module.TButton', font=STYLES['button_font'], padding=10)
        style.configure('Header.TFrame', background=STYLES['header_background'])
        
    def create_main_interface(self) -> None:
        """创建主界面"""
        # 主标题
        self._create_title_section()
        
        # 模块选择区域
        self._create_modules_section()
        
        # 状态栏
        self._create_status_bar()
        
    def _create_title_section(self) -> None:
        """创建标题区域"""
        title_frame = ttk.Frame(self.root, style='Header.TFrame')
        title_frame.pack(fill='x', padx=20, pady=20)
        
        title_label = ttk.Label(
            title_frame, 
            text=APP_NAME, 
            style='Title.TLabel'
        )
        title_label.pack(pady=10)
        
        subtitle_label = ttk.Label(
            title_frame, 
            text=APP_DESCRIPTION, 
            style='Subtitle.TLabel'
        )
        subtitle_label.pack()
        
        # 版本信息
        version_label = ttk.Label(
            title_frame, 
            text=f"版本: {APP_VERSION}", 
            font=('Arial', 8)
        )
        version_label.pack(pady=5)
        
    def _create_modules_section(self) -> None:
        """创建模块选择区域"""
        self.modules_frame = ttk.Frame(self.root)
        self.modules_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 创建模块网格
        self._create_module_grid()
        
    def _create_module_grid(self) -> None:
        """创建模块网格布局"""
        # 获取所有可用模块信息
        available_modules = get_all_modules()
        
        # 创建网格
        for i, (module_key, module_info) in enumerate(available_modules.items()):
            row = i // 2
            col = i % 2
            
            module_frame = ttk.Frame(
                self.modules_frame, 
                relief='raised', 
                borderwidth=2
            )
            module_frame.grid(
                row=row, 
                column=col, 
                padx=10, 
                pady=10, 
                sticky='nsew'
            )
            
            # 模块图标
            icon_label = ttk.Label(
                module_frame, 
                text=module_info['icon'], 
                font=('Arial', 24)
            )
            icon_label.pack(pady=(20, 10))
            
            # 模块名称
            name_label = ttk.Label(
                module_frame, 
                text=module_info['name'], 
                style='Module.TButton'
            )
            name_label.pack(pady=5)
            
            # 模块描述
            desc_label = ttk.Label(
                module_frame, 
                text=module_info['description'], 
                wraplength=200, 
                justify='center'
            )
            desc_label.pack(pady=5, padx=10)
            
            # 启动按钮
            start_btn = ttk.Button(
                module_frame, 
                text="启动模块", 
                command=lambda m=module_key: self.start_module(m),
                style='Module.TButton'
            )
            start_btn.pack(pady=15)
            
        # 配置网格权重
        self.modules_frame.grid_columnconfigure(0, weight=1)
        self.modules_frame.grid_columnconfigure(1, weight=1)
        
    def _create_status_bar(self) -> None:
        """创建状态栏"""
        self.status_bar = ttk.Label(
            self.root, 
            text="就绪", 
            relief='sunken', 
            anchor='w'
        )
        self.status_bar.pack(side='bottom', fill='x')
        
    def init_modules(self) -> None:
        """初始化所有模块"""
        try:
            available_modules = get_all_modules()
            
            for module_key in available_modules.keys():
                module_instance = create_module(module_key)
                if module_instance:
                    self.modules[module_key] = module_instance
                    
            self.update_status(f"已加载 {len(self.modules)} 个模块")
            
        except Exception as e:
            self.show_error("模块初始化失败", f"初始化模块时发生错误: {str(e)}")
            self.update_status("模块初始化失败")
        
    def start_module(self, module_name: str) -> None:
        """
        启动指定模块
        
        Args:
            module_name (str): 模块名称
        """
        try:
            if module_name in self.modules:
                self.modules[module_name].show()
                self.update_status(f"已启动模块: {module_name}")
            else:
                self.show_error("错误", f"模块 {module_name} 不存在")
        except Exception as e:
            self.show_error("错误", f"启动模块时发生错误: {str(e)}")
            self.update_status("模块启动失败")
    
    def update_status(self, message: str) -> None:
        """
        更新状态栏信息
        
        Args:
            message (str): 状态信息
        """
        if self.status_bar:
            self.status_bar.config(text=message)
    
    def show_error(self, title: str, message: str) -> None:
        """
        显示错误消息
        
        Args:
            title (str): 错误标题
            message (str): 错误信息
        """
        messagebox.showerror(title, message)
    
    def center_window(self) -> None:
        """将主窗口居中显示"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")


def main() -> None:
    """主函数"""
    try:
        # 创建主窗口
        root = tk.Tk()
        
        # 创建应用实例
        app = FExcelApp(root)
        
        # 设置窗口居中
        app.center_window()
        
        # 启动主循环
        root.mainloop()
        
    except Exception as e:
        print(f"应用程序启动失败: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
