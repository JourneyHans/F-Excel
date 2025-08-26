#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel ID值转换器模块

该模块继承自BaseModule基类，实现了ID值转换功能。
将【id=值】格式的文本转换为两列Excel文件，支持批量处理和文件导入导出。
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import os
import re
from typing import List, Dict, Any

from .base_module import BaseModule


class IDConverterModule(BaseModule):
    """
    ID值转换器模块
    
    继承自BaseModule，实现了ID值转换的特定功能。
    支持文件导入、数据转换和Excel导出。
    """
    
    def __init__(self):
        """初始化ID转换器模块"""
        # 先设置基本属性，避免在get_module_config中访问未初始化的属性
        self.converted_data: List[Dict[str, Any]] = []
        self.example_data = """示例格式：
410325=提升{0}
410326=降低{0}
410327=无变化
410328=高于
410329=等于
410330=低于"""
        
        # 调用父类初始化
        super().__init__()
    
    def get_module_config(self) -> Dict[str, Any]:
        """
        获取模块配置信息
        
        Returns:
            Dict[str, Any]: 模块配置字典
        """
        return {
            'name': 'ID值转换器',
            'description': '将数字=值格式转换为Excel两列文件',
            'icon': '📊',
            'window_size': '900x700',
            'supported_formats': ['id=值', '数字=值', 'ID=值']
        }
    
    def create_interface(self) -> None:
        """创建模块界面"""
        # 创建标题区域
        self.create_title_section(self.window)
        
        # 创建输入区域
        self._create_input_section()
        
        # 创建操作按钮区域
        self._create_button_section()
        
        # 创建输出预览区域
        self._create_output_section()
        
        # 创建状态栏
        self.create_status_bar(self.window)
    
    def _create_input_section(self) -> None:
        """创建输入数据区域"""
        input_frame = ttk.LabelFrame(self.window, text="输入数据", padding=10)
        input_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 创建文件上传区域
        self.create_file_upload_section(input_frame, ['.txt', '.csv'])
        
        # 创建文本输入区域
        ttk.Label(input_frame, text="或者直接输入数字=值格式的文本:").pack(anchor='w', pady=(10, 5))
        
        self.input_text = scrolledtext.ScrolledText(input_frame, height=15, width=80, wrap='word')
        self.input_text.pack(fill='both', expand=True, pady=(5, 0))
        
        # 插入示例数据
        self.input_text.insert('1.0', self.example_data)
    
    def _create_button_section(self) -> None:
        """创建操作按钮区域"""
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(
            button_frame, 
            text="转换并预览", 
            command=self.convert_and_preview
        ).pack(side='left', padx=(0, 10))
        
        ttk.Button(
            button_frame, 
            text="导出Excel", 
            command=self.export_excel
        ).pack(side='left', padx=10)
        
        ttk.Button(
            button_frame, 
            text="清空数据", 
            command=self.clear_data
        ).pack(side='left', padx=10)
    
    def _create_output_section(self) -> None:
        """创建输出预览区域"""
        output_frame = ttk.LabelFrame(self.window, text="转换结果预览", padding=10)
        output_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=10, width=80, wrap='word')
        self.output_text.pack(fill='both', expand=True)
    
    def browse_file(self) -> None:
        """浏览并选择文本文件"""
        file_path = filedialog.askopenfilename(
            title="选择文本文件",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.uploaded_file_path = file_path
            self.uploaded_file_name = os.path.splitext(os.path.basename(file_path))[0]
            self.file_path_var.set(file_path)
            self.load_file_content(file_path)
    
    def clear_file(self) -> None:
        """清空已选择的文件"""
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        self.file_path_var.set("")
        
        if self.input_text:
            self.input_text.delete('1.0', tk.END)
            self.input_text.insert('1.0', self.example_data)
    
    def convert_id_values(self, text: str, progress_callback=None) -> List[Dict[str, Any]]:
        """
        转换数字=值格式的文本
        
        Args:
            text (str): 输入文本
            progress_callback: 进度回调函数，用于实时更新进度
            
        Returns:
            List[Dict[str, Any]]: 转换后的数据列表
        """
        pattern = r'(\d+)=([^\s\n]+)'
        matches = re.findall(pattern, text)
        
        data = []
        total_matches = len(matches)
        
        for i, (code, value) in enumerate(matches):
            data.append({
                'ID': code,
                'Value': value.strip()
            })
            
            # 实时更新进度
            if progress_callback and total_matches > 0:
                progress = min(100, (i + 1) / total_matches * 100)
                progress_callback(progress, f"正在处理第 {i + 1}/{total_matches} 条数据...")
                
        return data
    
    def convert_and_preview(self) -> None:
        """转换并预览结果"""
        try:
            if not self.input_text:
                self.show_warning("警告", "输入框未初始化")
                return
            
            input_data = self.input_text.get('1.0', tk.END)
            
            if not input_data.strip():
                self.show_warning("警告", "请输入数据")
                return
            
            # 开始转换，显示进度
            self.update_progress(0, "正在转换数据...")
            
            # 使用带进度回调的转换方法
            self.converted_data = self.convert_id_values(input_data, self.update_progress)
            
            if not self.converted_data:
                self.show_info("提示", "未找到有效的数字=值格式数据")
                self.update_progress(0, "未找到有效数据")
                return
            
            # 显示预览
            self._show_preview()
            
            self.update_progress(100, f"转换完成，共处理 {len(self.converted_data)} 条数据")
            self.update_status(f"转换完成，共处理 {len(self.converted_data)} 条数据")
            
        except Exception as e:
            self.show_error("错误", f"转换失败: {str(e)}")
            self.update_progress(0, "转换失败")
    
    def _show_preview(self) -> None:
        """显示转换结果预览"""
        if not self.output_text:
            return
        
        self.output_text.delete('1.0', tk.END)
        
        preview_text = "转换结果预览:\n"
        preview_text += "=" * 50 + "\n"
        preview_text += f"{'序号':<8} {'ID':<8} {'值':<20}\n"
        preview_text += "-" * 50 + "\n"
        
        for i, row in enumerate(self.converted_data, 1):
            preview_text += f"{i:<8} {row['ID']:<8} {row['Value']:<20}\n"
        
        preview_text += "=" * 50 + "\n"
        preview_text += f"总计: {len(self.converted_data)} 条数据\n"
        
        self.output_text.insert('1.0', preview_text)
    
    def export_excel(self) -> None:
        """导出Excel文件"""
        try:
            if not self.converted_data:
                self.show_warning("警告", "请先转换数据")
                return
            
            # 生成默认文件名
            default_filename = self.generate_default_filename(suffix=".xlsx")
            
            file_path = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                return
            
            # 创建DataFrame并导出
            df = pd.DataFrame(self.converted_data)
            df.to_excel(file_path, sheet_name='ID值转换结果', index=False)
            
            self.show_info("成功", f"Excel文件已保存到:\n{file_path}")
            self.update_status(f"Excel文件已导出: {os.path.basename(file_path)}")
            
        except Exception as e:
            self.show_error("错误", f"导出失败: {str(e)}")
    
    def clear_data(self) -> None:
        """清空数据"""
        if self.ask_confirmation("确认", "确定要清空所有数据吗？"):
            if self.input_text:
                self.input_text.delete('1.0', tk.END)
                self.input_text.insert('1.0', self.example_data)
            
            if self.output_text:
                self.output_text.delete('1.0', tk.END)
            
            self.converted_data.clear()
            self.update_status("数据已清空")
    
    def process_data(self, input_data: str) -> List[Dict[str, Any]]:
        """
        处理输入数据（实现抽象方法）
        
        Args:
            input_data (str): 输入数据
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        return self.convert_id_values(input_data)
