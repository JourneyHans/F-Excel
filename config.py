# -*- coding: utf-8 -*-
"""
F-Excel 配置文件
"""

# 应用程序信息
APP_NAME = "F-Excel"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = "模块化数据处理工具集合"

# 窗口配置
MAIN_WINDOW_WIDTH = 800
MAIN_WINDOW_HEIGHT = 600
MODULE_WINDOW_WIDTH = 900
MODULE_WINDOW_HEIGHT = 700

# 文件配置
DEFAULT_EXCEL_SHEET_NAME = "ID值转换结果"
SUPPORTED_FILE_TYPES = {
    'excel': ['.xlsx', '.xls'],
    'text': ['.txt', '.csv']
}

# 正则表达式模式
ID_VALUE_PATTERN = r'(\d+)=([^\s\n]+)'

# 样式配置
STYLES = {
    'title_font': ('Arial', 16, 'bold'),
    'subtitle_font': ('Arial', 10),
    'module_font': ('Arial', 12, 'bold'),
    'button_font': ('Arial', 10),
    'header_background': '#f0f0f0'
}
