#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel翻译器模块
将Excel文件中的ID、中文、韩文三列转换为ID=韩文的文本格式
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import pandas as pd
import re
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

class ExcelTranslatorModule:
    def __init__(self):
        self.window = None
        self.input_text = None
        self.output_text = None
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        self.converted_data = []
        self.processing = False
        self.progress_var = None
        self.progress_bar = None
        self.status_label = None
        
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
        self.window.title("Excel翻译器 - F-Excel")
        self.window.geometry("1000x800")
        self.window.minsize(900, 700)
        
        # 创建界面
        self.create_interface()
        
        # 窗口居中
        self.center_window()
        
    def create_interface(self):
        """创建模块界面"""
        # 主标题
        title_frame = ttk.Frame(self.window)
        title_frame.pack(fill='x', padx=20, pady=10)
        
        title_label = ttk.Label(title_frame, text="Excel翻译器", font=('Arial', 14, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="将Excel文件中的ID、中文、韩文转换为ID=韩文格式", font=('Arial', 10))
        subtitle_label.pack(pady=5)
        
        # 输入区域
        input_frame = ttk.LabelFrame(self.window, text="输入数据", padding=10)
        input_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # 文件上传区域
        file_frame = ttk.Frame(input_frame)
        file_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(file_frame, text="选择Excel文件:").pack(side='left')
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, state='readonly').pack(side='left', padx=10)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side='left', padx=5)
        ttk.Button(file_frame, text="清空文件", command=self.clear_file).pack(side='left', padx=5)
        
        ttk.Label(input_frame, text="或者直接输入Excel格式的文本数据:").pack(anchor='w', pady=(10, 5))
        self.input_text = scrolledtext.ScrolledText(input_frame, height=15, width=80, wrap='word')
        self.input_text.pack(fill='both', expand=True, pady=(5, 0))
        
        # 示例数据
        example_text = """示例格式（制表符分隔）：
999470001	反馈问题	피드백
999470002	画质:	해상도:
999470003	上传日志:	로그 업로드:
999470004	异常上报成功，感谢团长对阿克迈斯的关注！	업로드 성공!
999470005	确认	확인
999470006	实名认证	본인인증
999470007	切换账号	계정 변경
999470008	标题	제목"""
        
        self.input_text.insert('1.0', example_text)
        
        # 操作按钮区域
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(button_frame, text="转换并预览", command=self.convert_and_preview).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="导出文本文件", command=self.export_text).pack(side='left', padx=10)
        ttk.Button(button_frame, text="清空数据", command=self.clear_data).pack(side='left', padx=10)
        
        # 取消按钮（动态显示）
        self.cancel_button = ttk.Button(button_frame, text="取消处理", command=self.cancel_processing, state='disabled')
        self.cancel_button.pack(side='left', padx=10)
        
        # 进度条区域
        progress_frame = ttk.Frame(self.window)
        progress_frame.pack(fill='x', padx=20, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', pady=5)
        
        self.status_label = ttk.Label(progress_frame, text="就绪", font=('Arial', 9))
        self.status_label.pack()
        
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
        if self.processing:
            messagebox.showwarning("警告", "正在处理中，请稍候...")
            return
            
        input_data = self.input_text.get('1.0', tk.END)
        
        if not input_data.strip():
            messagebox.showwarning("警告", "请输入数据")
            return
        
        # 检查数据量
        lines = [line.strip() for line in input_data.strip().split('\n') if line.strip() and not line.startswith('示例格式')]
        
        if len(lines) > 10000:  # 超过1万行使用异步处理
            self.start_async_conversion(input_data)
        else:
            self.convert_sync(input_data)
    
    def convert_sync(self, input_data):
        """同步转换数据"""
        try:
            self.status_label.config(text="正在转换数据...")
            self.progress_var.set(50)
            
            # 转换数据
            self.converted_data = self.convert_excel_data(input_data)
            
            if not self.converted_data:
                messagebox.showinfo("提示", "未找到有效的Excel格式数据")
                return
            
            # 显示预览
            self.show_preview()
            
            self.progress_var.set(100)
            self.status_label.config(text=f"转换完成，共处理 {len(self.converted_data)} 条数据")
            self.status_bar.config(text=f"转换完成，共处理 {len(self.converted_data)} 条数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            self.progress_var.set(0)
            self.status_label.config(text="转换失败")
            self.cancel_button.config(state='disabled')
    
    def start_async_conversion(self, input_data):
        """启动异步转换"""
        self.processing = True
        self.progress_var.set(0)
        self.status_label.config(text="正在启动异步转换...")
        
        # 启用取消按钮
        self.cancel_button.config(state='normal')
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self.convert_async, args=(input_data,))
        thread.daemon = True
        thread.start()
    
    def convert_async(self, input_data):
        """异步转换数据"""
        try:
            lines = [line.strip() for line in input_data.strip().split('\n') if line.strip() and not line.startswith('示例格式')]
            total_lines = len(lines)
            
            # 分批处理
            batch_size = 1000
            self.converted_data = []
            
            for i in range(0, total_lines, batch_size):
                if not self.processing:  # 检查是否被取消
                    break
                    
                batch_lines = lines[i:i + batch_size]
                batch_data = self.convert_batch(batch_lines)
                self.converted_data.extend(batch_data)
                
                # 更新进度
                progress = min(100, (i + batch_size) / total_lines * 100)
                self.window.after(0, self.update_progress, progress, f"正在处理第 {i + 1}-{min(i + batch_size, total_lines)} 行...")
                
                # 短暂休息，避免界面冻结
                time.sleep(0.01)
            
            if self.processing:  # 只有在未被取消时才完成
                self.window.after(0, self.conversion_completed)
            
        except Exception as e:
            self.window.after(0, self.conversion_failed, str(e))
    
    def convert_batch(self, lines):
        """分批转换数据"""
        data = []
        for line in lines:
            if '\t' in line:
                parts = line.split('\t')
                if len(parts) >= 3:
                    id_value = str(parts[0].strip())
                    chinese = str(parts[1].strip()) if parts[1].strip() else ""
                    korean = str(parts[2].strip()) if parts[2].strip() else ""
                    
                    if id_value.isdigit():
                        data.append({
                            'ID': id_value,
                            'Chinese': chinese,
                            'Korean': korean,
                            'Output': f"{id_value}={korean}"
                        })
        return data
    
    def update_progress(self, progress, status):
        """更新进度条和状态"""
        self.progress_var.set(progress)
        self.status_label.config(text=status)
    
    def conversion_completed(self):
        """转换完成处理"""
        self.processing = False
        self.progress_var.set(100)
        self.status_label.config(text=f"转换完成，共处理 {len(self.converted_data)} 条数据")
        
        # 禁用取消按钮
        self.cancel_button.config(state='disabled')
        
        # 显示预览
        self.show_preview()
        
        self.status_bar.config(text=f"转换完成，共处理 {len(self.converted_data)} 条数据")
        
        # 显示完成消息
        messagebox.showinfo("完成", f"数据转换完成！\n共处理 {len(self.converted_data)} 条数据")
    
    def conversion_failed(self, error_msg):
        """转换失败处理"""
        self.processing = False
        self.progress_var.set(0)
        self.status_label.config(text="转换失败")
        self.cancel_button.config(state='disabled')
        messagebox.showerror("错误", f"转换失败: {error_msg}")
    
    def cancel_processing(self):
        """取消处理"""
        if self.processing:
            self.processing = False
            self.progress_var.set(0)
            self.status_label.config(text="处理已取消")
            self.cancel_button.config(state='disabled')
            self.status_bar.config(text="处理已取消")
    
    def export_text(self):
        """导出文本文件"""
        try:
            if not self.converted_data:
                messagebox.showwarning("警告", "请先转换数据")
                return
            
            # 生成默认文件名
            default_filename = self.generate_default_filename()
            
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
            messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def export_small_file(self, file_path):
        """导出小型文件"""
        try:
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for row in self.converted_data:
                    f.write(f"{row['Output']}\n")
            
            messagebox.showinfo("成功", f"文本文件已保存到:\n{file_path}")
            self.status_bar.config(text=f"文本文件已导出: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def export_large_file(self, file_path):
        """导出大型文件（分批写入）"""
        try:
            self.status_label.config(text="正在导出大文件...")
            self.progress_var.set(0)
            
            total_rows = len(self.converted_data)
            batch_size = 10000
            
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for i in range(0, total_rows, batch_size):
                    batch = self.converted_data[i:i + batch_size]
                    
                    for row in batch:
                        f.write(f"{row['Output']}\n")
                    
                    # 更新进度
                    progress = min(100, (i + batch_size) / total_rows * 100)
                    self.progress_var.set(progress)
                    self.status_label.config(text=f"正在导出第 {i + 1}-{min(i + batch_size, total_rows)} 行...")
            
            self.progress_var.set(100)
            self.status_label.config(text="导出完成")
            
            messagebox.showinfo("成功", f"大文件已保存到:\n{file_path}\n共导出 {total_rows} 行数据")
            self.status_bar.config(text=f"大文件已导出: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            self.progress_var.set(0)
            self.status_label.config(text="导出失败")
            
    def convert_excel_data(self, text):
        """转换Excel格式的文本数据"""
        lines = text.strip().split('\n')
        data = []
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('示例格式') or line.startswith('='):
                continue
                
            # 尝试解析制表符分隔的数据
            if '\t' in line:
                parts = line.split('\t')
                if len(parts) >= 3:
                    id_value = str(parts[0].strip())  # 确保ID是字符串
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
                        
        return data
        
    def browse_file(self):
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
            
    def clear_file(self):
        """清空已选择的文件"""
        self.uploaded_file_path = None
        self.uploaded_file_name = None
        self.file_path_var.set("")
        self.input_text.delete('1.0', tk.END)
        # 重新插入示例数据
        example_text = """示例格式（制表符分隔）：
999470001	反馈问题	피드백
999470002	画质:	해상도:
999470003	上传日志:	로그 업로드:
999470004	异常上报成功，感谢团长对阿克迈斯的关注！	업로드 성공!
999470005	确认	확인
999470006	实名认证	본인인증
999470007	切换账号	계정 변경
999470008	标题	제목"""
        self.input_text.insert('1.0', example_text)
        
    def load_excel_file(self, file_path):
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
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")
    
    def load_small_excel_file(self, file_path):
        """加载小型Excel文件"""
        try:
            self.status_label.config(text="正在读取Excel文件...")
            self.progress_var.set(25)
            
            # 读取Excel文件，确保第一列（ID列）为字符串类型
            df = pd.read_excel(file_path, dtype={0: str})
            
            # 检查列数
            if len(df.columns) < 3:
                messagebox.showwarning("警告", "Excel文件至少需要3列（ID、中文、韩文）")
                return
            
            self.progress_var.set(50)
            
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
                    self.progress_var.set(progress)
                    self.status_label.config(text=f"正在处理第 {i + 1}/{total_rows} 行...")
            
            self.progress_var.set(100)
            self.status_label.config(text="文件加载完成")
            
            self.input_text.delete('1.0', tk.END)
            self.input_text.insert('1.0', text_content.strip())
            
            self.status_bar.config(text=f"已加载Excel文件: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")
            self.progress_var.set(0)
            self.status_label.config(text="文件加载失败")
    
    def load_large_excel_file(self, file_path):
        """加载大型Excel文件（分块读取）"""
        try:
            self.status_label.config(text="检测到大文件，正在分块读取...")
            self.progress_var.set(10)
            
            # 对于大文件，使用更智能的读取策略
            # 先读取前几行确定列结构
            df_sample = pd.read_excel(file_path, dtype={0: str}, nrows=1000)
            
            if len(df_sample.columns) < 3:
                messagebox.showwarning("警告", "Excel文件至少需要3列（ID、中文、韩文）")
                return
            
            self.progress_var.set(20)
            self.status_label.config(text="正在读取文件结构...")
            
            # 获取总行数（通过读取所有数据）
            df_full = pd.read_excel(file_path, dtype={0: str})
            total_rows = len(df_full)
            
            if total_rows > 100000:  # 超过10万行时只保留最后的部分
                df_full = df_full.tail(100000)
                total_rows = 100000
                messagebox.showwarning("警告", f"文件过大，只保留了最后 {total_rows} 行数据")
            
            self.progress_var.set(40)
            self.status_label.config(text="正在处理数据...")
            
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
                self.progress_var.set(progress)
                self.status_label.config(text=f"正在处理第 {i + 1}-{batch_end}/{total_rows} 行...")
                
                # 短暂休息，避免界面冻结
                time.sleep(0.01)
            
            self.progress_var.set(100)
            self.status_label.config(text="大文件加载完成")
            
            self.input_text.delete('1.0', tk.END)
            self.input_text.insert('1.0', text_content.strip())
            
            self.status_bar.config(text=f"已加载Excel文件: {os.path.basename(file_path)} ({total_rows} 行)")
            
        except Exception as e:
            messagebox.showerror("错误", f"读取大文件失败: {str(e)}")
            self.progress_var.set(0)
            self.status_label.config(text="文件加载失败")
        
    def show_preview(self):
        """显示转换结果预览"""
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
        
    def export_text(self):
        """导出文本文件"""
        try:
            if not self.converted_data:
                messagebox.showwarning("警告", "请先转换数据")
                return
                
            # 生成默认文件名
            default_filename = self.generate_default_filename()
            
            file_path = filedialog.asksaveasfilename(
                title="保存文本文件",
                defaultextension=".txt",
                initialfile=default_filename,
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                return
                
            # 写入文本文件
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                for row in self.converted_data:
                    f.write(f"{row['Output']}\n")
            
            messagebox.showinfo("成功", f"文本文件已保存到:\n{file_path}")
            self.status_bar.config(text=f"文本文件已导出: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            
    def generate_default_filename(self):
        """生成默认文件名"""
        from datetime import datetime
        
        # 获取当前时间
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d")
        
        # 如果有上传的文件，使用文件名；否则使用默认名称
        if self.uploaded_file_name:
            filename = f"{date_str}_{self.uploaded_file_name}_translated.txt"
        else:
            filename = f"{date_str}_Excel翻译结果.txt"
            
        return filename
            
    def clear_data(self):
        """清空数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.input_text.delete('1.0', tk.END)
            self.output_text.delete('1.0', tk.END)
            self.converted_data = []
            self.status_bar.config(text="数据已清空")
