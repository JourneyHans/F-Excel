#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 数据处理策略类

该模块实现了数据处理策略模式，为不同类型的数据转换提供统一的接口。
每种数据转换策略都实现了相同的接口，可以根据需要动态选择策略。

设计模式:
- 策略模式
- 工厂模式
- 单例模式
"""

import re
import pandas as pd
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time


class DataProcessor(ABC):
    """
    数据处理策略抽象基类
    
    定义了数据处理的标准接口，所有具体的数据处理策略都必须实现这些方法。
    """
    
    @abstractmethod
    def process(self, data: str) -> List[Dict[str, Any]]:
        """
        处理数据
        
        Args:
            data (str): 输入数据
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        pass
    
    @abstractmethod
    def validate(self, data: str) -> bool:
        """
        验证数据格式
        
        Args:
            data (str): 输入数据
            
        Returns:
            bool: 数据格式是否有效
        """
        pass
    
    @abstractmethod
    def get_supported_formats(self) -> List[str]:
        """
        获取支持的数据格式
        
        Returns:
            List[str]: 支持的数据格式列表
        """
        pass


class IDValueProcessor(DataProcessor):
    """
    ID值转换处理器
    
    处理【id=值】格式的文本，将其转换为结构化的数据。
    """
    
    def __init__(self):
        """初始化ID值处理器"""
        self.pattern = r'(\d+)=([^\s\n]+)'
        self.compiled_pattern = re.compile(self.pattern)
    
    def process(self, data: str) -> List[Dict[str, Any]]:
        """
        处理ID值格式的数据
        
        Args:
            data (str): 输入数据
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        if not self.validate(data):
            return []
        
        matches = self.compiled_pattern.findall(data)
        processed_data = []
        
        for code, value in matches:
            processed_data.append({
                'ID': code,
                'Value': value.strip(),
                'Type': 'ID_VALUE'
            })
        
        return processed_data
    
    def validate(self, data: str) -> bool:
        """
        验证ID值格式数据
        
        Args:
            data (str): 输入数据
            
        Returns:
            bool: 数据格式是否有效
        """
        if not data or not data.strip():
            return False
        
        # 检查是否包含至少一个有效的ID=值格式
        return bool(self.compiled_pattern.search(data))
    
    def get_supported_formats(self) -> List[str]:
        """
        获取支持的数据格式
        
        Returns:
            List[str]: 支持的数据格式列表
        """
        return ['id=值', '数字=值', 'ID=值']


class ExcelTranslationProcessor(DataProcessor):
    """
    Excel翻译数据处理器
    
    处理Excel格式的翻译数据，将其转换为ID=韩文格式。
    """
    
    def __init__(self):
        """初始化Excel翻译处理器"""
        self.separator = '\t'
        self.required_columns = 3
    
    def process(self, data: str) -> List[Dict[str, Any]]:
        """
        处理Excel翻译格式的数据
        
        Args:
            data (str): 输入数据
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        if not self.validate(data):
            return []
        
        lines = [line.strip() for line in data.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式')]
        
        processed_data = []
        
        for line in lines:
            if self.separator in line:
                parts = line.split(self.separator)
                if len(parts) >= self.required_columns:
                    id_value = str(parts[0].strip())
                    chinese = str(parts[1].strip()) if parts[1].strip() else ""
                    korean = str(parts[2].strip()) if parts[2].strip() else ""
                    
                    if id_value.isdigit():
                        processed_data.append({
                            'ID': id_value,
                            'Chinese': chinese,
                            'Korean': korean,
                            'Output': f"{id_value}={korean}",
                            'Type': 'EXCEL_TRANSLATION'
                        })
        
        return processed_data
    
    def validate(self, data: str) -> bool:
        """
        验证Excel翻译格式数据
        
        Args:
            data (str): 输入数据
            
        Returns:
            bool: 数据格式是否有效
        """
        if not data or not data.strip():
            return False
        
        lines = [line.strip() for line in data.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式')]
        
        if not lines:
            return False
        
        # 检查是否至少有一行包含制表符分隔的数据
        return any(self.separator in line for line in lines)
    
    def get_supported_formats(self) -> List[str]:
        """
        获取支持的数据格式
        
        Returns:
            List[str]: 支持的数据格式列表
        """
        return ['制表符分隔的三列数据', 'ID\t中文\t韩文', 'Excel格式数据']


class BatchDataProcessor:
    """
    批量数据处理器
    
    使用策略模式处理大量数据，支持异步处理和进度更新。
    """
    
    def __init__(self, processor: DataProcessor, batch_size: int = 1000):
        """
        初始化批量数据处理器
        
        Args:
            processor (DataProcessor): 数据处理器策略
            batch_size (int): 批处理大小
        """
        self.processor = processor
        self.batch_size = batch_size
        self.processing = False
        self.cancelled = False
    
    def process_batch(self, data: str, progress_callback=None) -> List[Dict[str, Any]]:
        """
        批量处理数据
        
        Args:
            data (str): 输入数据
            progress_callback: 进度回调函数
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        if not self.processor.validate(data):
            return []
        
        lines = [line.strip() for line in data.strip().split('\n') 
                if line.strip() and not line.startswith('示例格式')]
        
        total_lines = len(lines)
        if total_lines == 0:
            return []
        
        self.processing = True
        self.cancelled = False
        processed_data = []
        
        try:
            for i in range(0, total_lines, self.batch_size):
                if self.cancelled:
                    break
                
                batch_end = min(i + self.batch_size, total_lines)
                batch_lines = lines[i:batch_end]
                
                # 处理当前批次
                batch_data = self.process_batch_lines(batch_lines)
                processed_data.extend(batch_data)
                
                # 更新进度
                if progress_callback:
                    progress = min(100, (batch_end / total_lines) * 100)
                    progress_callback(progress, f"正在处理第 {i + 1}-{batch_end}/{total_lines} 行...")
                
                # 短暂休息，避免界面冻结
                time.sleep(0.01)
        
        finally:
            self.processing = False
        
        return processed_data
    
    def process_batch_lines(self, lines: List[str]) -> List[Dict[str, Any]]:
        """
        处理一批数据行
        
        Args:
            lines (List[str]): 数据行列表
            
        Returns:
            List[Dict[str, Any]]: 处理后的数据列表
        """
        batch_data = []
        
        for line in lines:
            if self.separator in line:
                parts = line.split(self.separator)
                if len(parts) >= self.required_columns:
                    id_value = str(parts[0].strip())
                    chinese = str(parts[1].strip()) if parts[1].strip() else ""
                    korean = str(parts[2].strip()) if parts[2].strip() else ""
                    
                    if id_value.isdigit():
                        batch_data.append({
                            'ID': id_value,
                            'Chinese': chinese,
                            'Korean': korean,
                            'Output': f"{id_value}={korean}",
                            'Type': 'EXCEL_TRANSLATION'
                        })
        
        return batch_data
    
    def cancel_processing(self) -> None:
        """取消处理"""
        self.cancelled = True
        self.processing = False


class DataProcessorFactory:
    """
    数据处理器工厂类
    
    使用工厂模式创建不同类型的数据处理器。
    """
    
    _processors = {
        'id_value': IDValueProcessor,
        'excel_translation': ExcelTranslationProcessor
    }
    
    @classmethod
    def create_processor(cls, processor_type: str) -> Optional[DataProcessor]:
        """
        创建数据处理器
        
        Args:
            processor_type (str): 处理器类型
            
        Returns:
            Optional[DataProcessor]: 数据处理器实例，如果类型不存在则返回None
        """
        processor_class = cls._processors.get(processor_type)
        if processor_class:
            return processor_class()
        return None
    
    @classmethod
    def get_available_processors(cls) -> List[str]:
        """
        获取可用的处理器类型
        
        Returns:
            List[str]: 可用的处理器类型列表
        """
        return list(cls._processors.keys())
    
    @classmethod
    def register_processor(cls, processor_type: str, processor_class: type) -> None:
        """
        注册新的数据处理器
        
        Args:
            processor_type (str): 处理器类型
            processor_class (type): 处理器类
        """
        if issubclass(processor_class, DataProcessor):
            cls._processors[processor_type] = processor_class
        else:
            raise ValueError(f"处理器类必须继承自 DataProcessor")


# 导出所有类
__all__ = [
    'DataProcessor',
    'IDValueProcessor', 
    'ExcelTranslationProcessor',
    'BatchDataProcessor',
    'DataProcessorFactory'
]
