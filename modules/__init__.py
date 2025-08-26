#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 核心模块包

该包包含所有功能模块的实现，采用模块化设计模式，
每个模块都是独立的类，通过统一的接口进行管理。

模块列表:
- IDConverterModule: ID值转换器模块
- ExcelTranslatorModule: Excel翻译器模块

设计模式:
- 模块化设计模式
- 工厂模式（模块管理）
- 策略模式（数据处理）
- 观察者模式（进度更新）
"""

__version__ = "1.0.0"
__author__ = "F-Excel Team"
__description__ = "F-Excel 核心功能模块包"

# 导入所有模块
from .id_converter import IDConverterModule
from .excel_translator import ExcelTranslatorModule

# 模块注册表
AVAILABLE_MODULES = {
    'id_converter': {
        'name': 'ID值转换器',
        'description': '将数字=值格式转换为Excel文件',
        'icon': '📊',
        'class': IDConverterModule,
        'version': '1.0.0'
    },
    'excel_translator': {
        'name': 'Excel翻译器',
        'description': '将Excel文件中的ID、中文、韩文转换为ID=韩文格式',
        'icon': '🌐',
        'class': ExcelTranslatorModule,
        'version': '1.0.0'
    }
}

def get_module_info(module_name: str) -> dict:
    """
    获取指定模块的信息
    
    Args:
        module_name (str): 模块名称
        
    Returns:
        dict: 模块信息字典，如果模块不存在则返回None
    """
    return AVAILABLE_MODULES.get(module_name)

def get_all_modules() -> dict:
    """
    获取所有可用模块的信息
    
    Returns:
        dict: 所有模块的信息字典
    """
    return AVAILABLE_MODULES.copy()

def create_module(module_name: str):
    """
    创建指定模块的实例
    
    Args:
        module_name (str): 模块名称
        
    Returns:
        模块实例，如果模块不存在则返回None
    """
    module_info = get_module_info(module_name)
    if module_info:
        return module_info['class']()
    return None

__all__ = [
    'IDConverterModule',
    'ExcelTranslatorModule',
    'AVAILABLE_MODULES',
    'get_module_info',
    'get_all_modules',
    'create_module'
]
