#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ID值转换器模块
将【id=值】格式的文本转换为两列Excel文件
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import re
import os

class IDConverterModule:
    def __init__(self):
        self.window = None
        self.input_text = None
        self.output_text = None
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        
    def show(self):
        """显示模块窗口"""
        if self.window is None or not self.window.winfo_exists():
            self.create_window()
        else:
            self.window.lift()
            self.window.focus()
            
    def create_window(self):
        """创建模块窗口"""
        self.window = tk.Toplevel()
        self.window.title("ID值转换器 - F-Excel")
        self.window.geometry("900x700")
        self.window.minsize(800, 600)
        
        # 创建界面
        self.create_interface()
        
        # 窗口居中
        self.center_window()
        
    def create_interface(self):
        """创建模块界面"""
        # 主标题
        title_frame = ttk.Frame(self.window)
        title_frame.pack(fill='x', padx=20, pady=10)
        
        title_label = ttk.Label(title_frame, text="ID值转换器", font=('Arial', 14, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="将数字=值格式转换为Excel两列文件", font=('Arial', 10))
        subtitle_label.pack(pady=5)
        
        # 输入区域
        input_frame = ttk.LabelFrame(self.window, text="输入数据", padding=10)
        input_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 文件上传区域
        file_frame = ttk.Frame(input_frame)
        file_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(file_frame, text="选择文本文件:").pack(side='left')
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, state='readonly').pack(side='left', padx=10)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side='left', padx=5)
        ttk.Button(file_frame, text="清空文件", command=self.clear_file).pack(side='left', padx=5)
        
        ttk.Label(input_frame, text="或者直接输入数字=值格式的文本:").pack(anchor='w', pady=(10, 5))
        self.input_text = scrolledtext.ScrolledText(input_frame, height=15, width=80, wrap='word')
        self.input_text.pack(fill='both', expand=True, pady=(5, 0))
        
        # 示例数据
        example_text = """示例格式：
410325=提升{0}
410326=降低{0}
410327=无变化
410328=高于
410329=等于
410330=低于"""
        
        self.input_text.insert('1.0', example_text)
        
        # 操作按钮区域
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(button_frame, text="转换并预览", command=self.convert_and_preview).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="导出Excel", command=self.export_excel).pack(side='left', padx=10)
        ttk.Button(button_frame, text="清空数据", command=self.clear_data).pack(side='left', padx=10)
        
        # 输出预览区域
        output_frame = ttk.LabelFrame(self.window, text="转换结果预览", padding=10)
        output_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=10, width=80, wrap='word')
        self.output_text.pack(fill='both', expand=True)
        
        # 状态栏
        self.status_bar = ttk.Label(self.window, text="就绪", relief='sunken', anchor='w')
        self.status_bar.pack(side='bottom', fill='x')
        
    def center_window(self):
        """窗口居中"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
    def convert_and_preview(self):
        """转换并预览结果"""
        try:
            input_data = self.input_text.get('1.0', tk.END)
            
            if not input_data.strip():
                messagebox.showwarning("警告", "请输入数据")
                return
                
            # 转换数据
            converted_data = self.convert_id_values(input_data)
            
            if not converted_data:
                messagebox.showinfo("提示", "未找到有效的数字=值格式数据")
                return
                
            # 显示预览
            self.show_preview(converted_data)
            
            self.status_bar.config(text=f"转换完成，共处理 {len(converted_data)} 条数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            
    def convert_id_values(self, text):
        """转换数字=值格式的文本"""
        pattern = r'(\d+)=([^\s\n]+)'   # 数字=值格式
        
        matches = re.findall(pattern, text)
        
        data = []
        
        # 处理数字=值格式
        for code, value in matches:
            data.append({
                'ID': code,
                'Value': value.strip()
            })
            
        return data
        
    def browse_file(self):
        """浏览并选择文本文件"""
        file_path = filedialog.askopenfilename(
            title="选择文本文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            self.uploaded_file_path = file_path
            self.uploaded_file_name = os.path.splitext(os.path.basename(file_path))[0]
            self.file_path_var.set(file_path)
            self.load_file_content(file_path)
            
    def clear_file(self):
        """清空已选择的文件"""
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        self.file_path_var.set("")
        self.input_text.delete('1.0', tk.END)
        # 重新插入示例数据
        example_text = """示例格式：
410325=提升{0}
410326=降低{0}
410327=无变化
410328=高于
410329=等于
410330=低于"""
        self.input_text.insert('1.0', example_text)
        
    def load_file_content(self, file_path):
        """加载文件内容到输入框"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                self.input_text.delete('1.0', tk.END)
                self.input_text.insert('1.0', content)
            self.status_bar.config(text=f"已加载文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
        
    def show_preview(self, data):
        """显示转换结果预览"""
        self.output_text.delete('1.0', tk.END)
        
        preview_text = "转换结果预览:\n"
        preview_text += "=" * 50 + "\n"
        preview_text += f"{'序号':<8} {'ID':<8} {'值':<20}\n"
        preview_text += "-" * 50 + "\n"
        
        for i, row in enumerate(data, 1):
            preview_text += f"{i:<8} {row['ID']:<8} {row['Value']:<20}\n"
            
        preview_text += "=" * 50 + "\n"
        preview_text += f"总计: {len(data)} 条数据\n"
        
        self.output_text.insert('1.0', preview_text)
        
    def export_excel(self):
        """导出Excel文件"""
        try:
            output_content = self.output_text.get('1.0', tk.END)
            
            if "转换结果预览:" not in output_content:
                messagebox.showwarning("警告", "请先转换数据")
                return
                
            # 生成默认文件名
            default_filename = self.generate_default_filename()
            
            file_path = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                return
                
            input_data = self.input_text.get('1.0', tk.END)
            converted_data = self.convert_id_values(input_data)
            
            if not converted_data:
                messagebox.showerror("错误", "没有可导出的数据")
                return
                
            # 创建DataFrame并导出
            df = pd.DataFrame(converted_data)
            df.to_excel(file_path, sheet_name='ID值转换结果', index=False)
            
            messagebox.showinfo("成功", f"Excel文件已保存到:\n{file_path}")
            self.status_bar.config(text=f"Excel文件已导出: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            
    def generate_default_filename(self):
        """生成默认文件名：当前时间的年月日_文本文件名称"""
        from datetime import datetime
        
        # 获取当前时间
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d")
        
        # 如果有上传的文件，使用文件名；否则使用默认名称
        if self.uploaded_file_name:
            filename = f"{date_str}_{self.uploaded_file_name}.xlsx"
        else:
            filename = f"{date_str}_ID值转换结果.xlsx"
            
        return filename
            
    def clear_data(self):
        """清空数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.input_text.delete('1.0', tk.END)
            self.output_text.delete('1.0', tk.END)
            self.status_bar.config(text="数据已清空")
