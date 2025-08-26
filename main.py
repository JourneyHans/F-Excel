#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 桌面应用程序
主程序入口
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os

# 导入模块
from modules.id_converter import IDConverterModule
from modules.excel_translator import ExcelTranslatorModule

class FExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("F-Excel 桌面工具集")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # 设置应用图标和样式
        self.setup_styles()
        
        # 创建主界面
        self.create_main_interface()
        
        # 初始化模块
        self.modules = {}
        self.init_modules()
        
    def setup_styles(self):
        """设置应用样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置样式
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Module.TButton', font=('Arial', 12), padding=10)
        style.configure('Header.TFrame', background='#f0f0f0')
        
    def create_main_interface(self):
        """创建主界面"""
        # 主标题
        title_frame = ttk.Frame(self.root, style='Header.TFrame')
        title_frame.pack(fill='x', padx=20, pady=20)
        
        title_label = ttk.Label(title_frame, text="F-Excel 桌面工具集", style='Title.TLabel')
        title_label.pack(pady=10)
        
        subtitle_label = ttk.Label(title_frame, text="模块化数据处理工具集合", font=('Arial', 10))
        subtitle_label.pack()
        
        # 模块选择区域
        self.modules_frame = ttk.Frame(self.root)
        self.modules_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 创建模块网格
        self.create_module_grid()
        
        # 状态栏
        self.status_bar = ttk.Label(self.root, text="就绪", relief='sunken', anchor='w')
        self.status_bar.pack(side='bottom', fill='x')
        
    def create_module_grid(self):
        """创建模块网格布局"""
        # 模块配置
        modules_config = [
            {
                'name': 'ID值转换器',
                'description': '将数字=值格式转换为Excel文件',
                'icon': '📊',
                'module_class': 'id_converter'
            },
            {
                'name': 'Excel翻译器',
                'description': '将Excel文件中的ID、中文、韩文转换为ID=韩文格式',
                'icon': '🌐',
                'module_class': 'excel_translator'
            }
        ]
        
        # 创建网格
        for i, config in enumerate(modules_config):
            row = i // 2
            col = i % 2
            
            module_frame = ttk.Frame(self.modules_frame, relief='raised', borderwidth=2)
            module_frame.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
            
            # 模块图标
            icon_label = ttk.Label(module_frame, text=config['icon'], font=('Arial', 24))
            icon_label.pack(pady=(20, 10))
            
            # 模块名称
            name_label = ttk.Label(module_frame, text=config['name'], font=('Arial', 12, 'bold'))
            name_label.pack(pady=5)
            
            # 模块描述
            desc_label = ttk.Label(module_frame, text=config['description'], 
                                 wraplength=200, justify='center')
            desc_label.pack(pady=5, padx=10)
            
            # 启动按钮
            start_btn = ttk.Button(module_frame, text="启动模块", 
                                 command=lambda m=config['module_class']: self.start_module(m),
                                 style='Module.TButton')
            start_btn.pack(pady=15)
            
        # 配置网格权重
        self.modules_frame.grid_columnconfigure(0, weight=1)
        self.modules_frame.grid_columnconfigure(1, weight=1)
        
    def init_modules(self):
        """初始化所有模块"""
        self.modules['id_converter'] = IDConverterModule()
        self.modules['excel_translator'] = ExcelTranslatorModule()
        
    def start_module(self, module_name):
        """启动指定模块"""
        try:
            if module_name in self.modules:
                self.modules[module_name].show()
                self.status_bar.config(text=f"已启动模块: {module_name}")
            else:
                messagebox.showerror("错误", f"模块 {module_name} 不存在")
        except Exception as e:
            messagebox.showerror("错误", f"启动模块时发生错误: {str(e)}")
            self.status_bar.config(text="模块启动失败")

def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = FExcelApp(root)
        
        # 设置窗口居中
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        root.mainloop()
        
    except Exception as e:
        print(f"应用程序启动失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
