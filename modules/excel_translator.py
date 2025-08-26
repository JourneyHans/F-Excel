#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel Excel翻译器模块

该模块继承自BaseModule基类，实现了Excel翻译功能。
将Excel文件中的ID、中文、韩文三列转换为ID=韩文格式，支持大文件处理和异步转换。
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import os
import threading
import time
from typing import List, Dict, Any, Optional

from .base_module import BaseModule


class ExcelTranslatorModule(BaseModule):
    """
    Excel翻译器模块
    
    继承自BaseModule，实现了Excel翻译的特定功能。
    支持大文件异步处理和进度更新。
    """
    
    def __init__(self):
        """初始化Excel翻译器模块"""
        # 先设置基本属性，避免在get_module_config中访问未初始化的属性
        self.converted_data: List[Dict[str, Any]] = []
        self.processing = False
        
        # 示例数据
        self.example_data = """示例格式（制表符分隔）：
999470001	反馈问题	피드백
999470002	画质:	해상도:
999470003	上传日志:	로그 업로드:
999470004	异常上报成功，感谢团长对阿克迈斯的关注！	업로드 성공!
999470005	确认	확인
999470006	实名认证	본인인증
999470007	切换账号	계정 변경
999470008	标题	제목"""
        
        # 调用父类初始化
        super().__init__()
    
    def get_module_config(self) -> Dict[str, Any]:
        """
        获取模块配置信息
        
        Returns:
            Dict[str, Any]: 模块配置字典
        """
        return {
            'name': 'Excel翻译器',
            'description': '将Excel文件中的ID、中文、韩文转换为ID=韩文格式',
            'icon': '🌐',
            'window_size': '1000x800',
            'supported_formats': ['制表符分隔的三列数据', 'ID\t中文\t韩文', 'Excel格式数据']
        }
    
    def create_interface(self) -> None:
        """创建模块界面"""
        # 创建标题区域
        self.create_title_section(self.window)
        
        # 创建输入区域
        self._create_input_section()
        
        # 创建操作按钮区域
        self._create_button_section()
        
        # 创建进度条区域
        self.create_progress_section(self.window)
        
        # 创建输出预览区域
        self._create_output_section()
        
        # 创建状态栏
        self.create_status_bar(self.window)
    
    def _create_input_section(self) -> None:
        """创建输入数据区域"""
        input_frame = ttk.LabelFrame(self.window, text="输入数据", padding=10)
        input_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 创建文件上传区域
        self.create_file_upload_section(input_frame, ['.xlsx', '.xls'])
        
        # 创建文本输入区域
        ttk.Label(input_frame, text="或者直接输入Excel格式的文本数据:").pack(anchor='w', pady=(10, 5))
        
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
            text="导出文本文件", 
            command=self.export_text
        ).pack(side='left', padx=10)
        
        ttk.Button(
            button_frame, 
            text="清空数据", 
            command=self.clear_data
        ).pack(side='left', padx=10)
        
        # 取消按钮（动态显示）
        self.cancel_button = ttk.Button(
            button_frame, 
            text="取消处理", 
            command=self.cancel_processing, 
            state='disabled'
        )
        self.cancel_button.pack(side='left', padx=10)
    
    def _create_output_section(self) -> None:
        """创建输出预览区域"""
        output_frame = ttk.LabelFrame(self.window, text="转换结果预览", padding=10)
        output_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=10, width=80, wrap='word')
        self.output_text.pack(fill='both', expand=True)
    
    def browse_file(self) -> None:
        """浏览并选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.uploaded_file_path = file_path
            self.uploaded_file_name = os.path.splitext(os.path.basename(file_path))[0]
            self.file_path_var.set(file_path)
            self.load_excel_file(file_path)
    
    def clear_file(self) -> None:
        """清空已选择的文件"""
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        self.file_path_var.set("")
        
        if self.input_text:
            self.input_text.delete('1.0', tk.END)
            self.input_text.insert('1.0', self.example_data)
    
    def convert_and_preview(self) -> None:
        """转换并预览结果"""
        if self.processing:
            self.show_warning("警告", "正在处理中，请稍候...")
            return
        
        if not self.input_text:
            self.show_warning("警告", "输入框未初始化")
            return
        
        input_data = self.input_text.get('1.0', tk.END)
        
        if not input_data.strip():
            self.show_warning("警告", "请输入数据")
            return
        
        # 检查数据量，大文件使用异步处理
        lines = [line.strip() for line in input_data.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式') and not line.startswith('=')]
        
        if len(lines) > 10000:  # 超过1万行使用异步处理
            self.start_async_conversion(input_data)
        else:
            self.convert_sync(input_data)
    
    def convert_excel_data(self, text: str, progress_callback=None) -> List[Dict[str, Any]]:
        """
        转换Excel格式的文本数据
        
        Args:
            text (str): 输入文本
            progress_callback: 进度回调函数，用于实时更新进度
            
        Returns:
            List[Dict[str, Any]]: 转换后的数据列表
        """
        lines = [line.strip() for line in text.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式') and not line.startswith('=')]
        
        if not lines:
            return []
        
        data = []
        total_lines = len(lines)
        
        for i, line in enumerate(lines):
            # 尝试解析制表符分隔的数据
            if '\t' in line:
                parts = line.split('\t')
                if len(parts) >= 3:
                    id_value = str(parts[0].strip())
                    chinese = str(parts[1].strip()) if parts[1].strip() else ""
                    korean = str(parts[2].strip()) if parts[2].strip() else ""
                    
                    # 验证ID是否为数字（字符串形式的数字）
                    if id_value.isdigit():
                        data.append({
                            'ID': id_value,
                            'Chinese': chinese,
                            'Korean': korean,
                            'Output': f"{id_value}={korean}"
                        })
            
            # 实时更新进度
            if progress_callback and i % max(1, total_lines // 100) == 0:  # 每1%更新一次进度
                progress = min(100, (i + 1) / total_lines * 100)
                progress_callback(progress, f"正在处理第 {i + 1}/{total_lines} 行...")
                
        return data
    
    def convert_sync(self, input_data: str) -> None:
        """同步转换数据"""
        try:
            self.update_progress(0, "正在转换数据...")
            
            # 使用带进度回调的转换逻辑
            self.converted_data = self.convert_excel_data(input_data, self.update_progress)
            
            if not self.converted_data:
                self.show_info("提示", "未找到有效的Excel格式数据")
                return
            
            # 显示预览
            self._show_preview()
            
            self.update_progress(100, f"转换完成，共处理 {len(self.converted_data)} 条数据")
            self.update_status(f"转换完成，共处理 {len(self.converted_data)} 条数据")
            
        except Exception as e:
            self.show_error("错误", f"转换失败: {str(e)}")
            self.update_progress(0, "转换失败")
    
    def start_async_conversion(self, input_data: str) -> None:
        """启动异步转换"""
        self.processing = True
        self.update_progress(0, "正在启动异步转换...")
        
        # 启用取消按钮
        self.cancel_button.config(state='normal')
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self.convert_async, args=(input_data,))
        thread.daemon = True
        thread.start()
    
    def convert_async(self, input_data: str) -> None:
        """异步转换数据"""
        try:
            # 使用带进度回调的转换逻辑
            self.converted_data = self.convert_excel_data(input_data, self.update_progress)
            
            if self.processing:  # 只有在未被取消时才完成
                self.window.after(0, self.conversion_completed)
            
        except Exception as e:
            self.window.after(0, self.conversion_failed, str(e))
    
    def conversion_completed(self) -> None:
        """转换完成处理"""
        self.processing = False
        self.update_progress(100, f"转换完成，共处理 {len(self.converted_data)} 条数据")
        
        # 禁用取消按钮
        self.cancel_button.config(state='disabled')
        
        # 显示预览
        self._show_preview()
        
        self.update_status(f"转换完成，共处理 {len(self.converted_data)} 条数据")
        
        # 显示完成消息
        self.show_info("完成", f"数据转换完成！\n共处理 {len(self.converted_data)} 条数据")
    
    def conversion_failed(self, error_msg: str) -> None:
        """转换失败处理"""
        self.processing = False
        self.update_progress(0, "转换失败")
        self.cancel_button.config(state='disabled')
        self.show_error("错误", f"转换失败: {error_msg}")
    
    def cancel_processing(self) -> None:
        """取消处理"""
        if self.processing:
            self.processing = False
            self.update_progress(0, "处理已取消")
            self.cancel_button.config(state='disabled')
            self.update_status("处理已取消")
    
    def export_text(self) -> None:
        """导出文本文件"""
        try:
            if not self.converted_data:
                self.show_warning("警告", "请先转换数据")
                return
            
            # 生成默认文件名
            default_filename = self.generate_default_filename(suffix="_translated.txt")
            
            file_path = filedialog.asksaveasfilename(
                title="保存文本文件",
                defaultextension=".txt",
                initialfile=default_filename,
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                return
            
            # 检查数据量，大文件使用分批写入
            if len(self.converted_data) > 10000:
                self.export_large_file(file_path)
            else:
                self.export_small_file(file_path)
                
        except Exception as e:
            self.show_error("错误", f"导出失败: {str(e)}")
    
    def export_small_file(self, file_path: str) -> None:
        """导出小型文件"""
        try:
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for row in self.converted_data:
                    f.write(f"{row['Output']}\n")
            
            self.show_info("成功", f"文本文件已保存到:\n{file_path}")
            self.update_status(f"文本文件已导出: {os.path.basename(file_path)}")
            
        except Exception as e:
            self.show_error("错误", f"导出失败: {str(e)}")
    
    def export_large_file(self, file_path: str) -> None:
        """导出大型文件（分批写入）"""
        try:
            self.update_progress(0, "正在导出大文件...")
            
            total_rows = len(self.converted_data)
            batch_size = 10000
            
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for i in range(0, total_rows, batch_size):
                    batch = self.converted_data[i:i + batch_size]
                    
                    for row in batch:
                        f.write(f"{row['Output']}\n")
                    
                    # 更新进度
                    progress = min(100, (i + batch_size) / total_rows * 100)
                    self.update_progress(progress, f"正在导出第 {i + 1}-{min(i + batch_size, total_rows)} 行...")
            
            self.update_progress(100, "导出完成")
            
            self.show_info("成功", f"大文件已保存到:\n{file_path}\n共导出 {total_rows} 行数据")
            self.update_status(f"大文件已导出: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            self.show_error("错误", f"导出失败: {str(e)}")
            self.update_progress(0, "导出失败")
    
    def load_excel_file(self, file_path: str) -> None:
        """加载Excel文件内容"""
        try:
            # 检查文件大小
            file_size = os.path.getsize(file_path)
            file_size_mb = file_size / (1024 * 1024)
            
            if file_size_mb > 50:  # 超过50MB使用分块读取
                self.load_large_excel_file(file_path)
            else:
                self.load_small_excel_file(file_path)
                
        except Exception as e:
            self.show_error("错误", f"读取Excel文件失败: {str(e)}")
    
    def load_small_excel_file(self, file_path: str) -> None:
        """加载小型Excel文件"""
        try:
            self.update_progress(25, "正在读取Excel文件...")
            
            # 读取Excel文件，确保第一列（ID列）为字符串类型
            df = pd.read_excel(file_path, dtype={0: str})
            
            # 检查列数
            if len(df.columns) < 3:
                self.show_warning("警告", "Excel文件至少需要3列（ID、中文、韩文）")
                return
            
            self.update_progress(50, "正在处理数据...")
            
            # 获取前3列数据
            df_subset = df.iloc[:, :3]
            
            # 转换为文本格式，确保ID列是字符串
            text_content = ""
            total_rows = len(df_subset)
            
            for i, (_, row) in enumerate(df_subset.iterrows()):
                # 确保ID列是字符串类型
                id_value = str(row.iloc[0])
                chinese_value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                korean_value = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                
                text_content += f"{id_value}\t{chinese_value}\t{korean_value}\n"
                
                # 更新进度
                if i % 1000 == 0:  # 每1000行更新一次进度
                    progress = 50 + (i / total_rows) * 50
                    self.update_progress(progress, f"正在处理第 {i + 1}/{total_rows} 行...")
            
            self.update_progress(100, "文件加载完成")
            
            if self.input_text:
                self.input_text.delete('1.0', tk.END)
                self.input_text.insert('1.0', text_content.strip())
            
            self.update_status(f"已加载Excel文件: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            self.show_error("错误", f"读取Excel文件失败: {str(e)}")
            self.update_progress(0, "文件加载失败")
    
    def load_large_excel_file(self, file_path: str) -> None:
        """加载大型Excel文件（分块读取）"""
        try:
            self.update_progress(10, "检测到大文件，正在分块读取...")
            
            # 对于大文件，使用更智能的读取策略
            # 先读取前几行确定列结构
            df_sample = pd.read_excel(file_path, dtype={0: str}, nrows=1000)
            
            if len(df_sample.columns) < 3:
                self.show_warning("警告", "Excel文件至少需要3列（ID、中文、韩文）")
                return
            
            self.update_progress(20, "正在读取文件结构...")
            
            # 获取总行数（通过读取所有数据）
            df_full = pd.read_excel(file_path, dtype={0: str})
            total_rows = len(df_full)
            
            if total_rows > 100000:  # 超过10万行时只保留最后的部分
                df_full = df_full.tail(100000)
                total_rows = 100000
                self.show_warning("警告", f"文件过大，只保留了最后 {total_rows} 行数据")
            
            self.update_progress(40, "正在处理数据...")
            
            # 获取前3列数据
            df_subset = df_full.iloc[:, :3]
            
            # 转换为文本格式，分批处理
            text_content = ""
            batch_size = 5000
            
            for i in range(0, total_rows, batch_size):
                batch_end = min(i + batch_size, total_rows)
                batch_df = df_subset.iloc[i:batch_end]
                
                for _, row in batch_df.iterrows():
                    id_value = str(row.iloc[0])
                    chinese_value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                    korean_value = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                    
                    text_content += f"{id_value}\t{chinese_value}\t{korean_value}\n"
                
                # 更新进度
                progress = 40 + (batch_end / total_rows) * 60
                self.update_progress(progress, f"正在处理第 {i + 1}-{batch_end}/{total_rows} 行...")
                
                # 短暂休息，避免界面冻结
                time.sleep(0.01)
            
            self.update_progress(100, "大文件加载完成")
            
            if self.input_text:
                self.input_text.delete('1.0', tk.END)
                self.input_text.insert('1.0', text_content.strip())
            
            self.update_status(f"已加载Excel文件: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            self.show_error("错误", f"读取大文件失败: {str(e)}")
            self.update_progress(0, "文件加载失败")
    
    def _show_preview(self) -> None:
        """显示转换结果预览"""
        if not self.output_text:
            return
        
        self.output_text.delete('1.0', tk.END)
        
        preview_text = "转换结果预览:\n"
        preview_text += "=" * 60 + "\n"
        preview_text += f"{'序号':<6} {'ID':<12} {'中文':<20} {'韩文':<20} {'输出格式':<30}\n"
        preview_text += "-" * 60 + "\n"
        
        for i, row in enumerate(self.converted_data, 1):
            preview_text += f"{i:<6} {row['ID']:<12} {row['Chinese']:<20} {row['Korean']:<20} {row['Output']:<30}\n"
        
        preview_text += "=" * 60 + "\n"
        preview_text += f"总计: {len(self.converted_data)} 条数据\n\n"
        
        # 添加纯输出格式预览
        preview_text += "纯输出格式预览:\n"
        preview_text += "-" * 30 + "\n"
        for row in self.converted_data:
            preview_text += f"{row['Output']}\n"
        
        self.output_text.insert('1.0', preview_text)
    
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
        return self.convert_excel_data(input_data)
