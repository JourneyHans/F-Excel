#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 配置文件

该文件包含应用程序的所有配置信息，包括应用信息、窗口配置、
文件配置、样式配置等。采用集中配置管理，便于维护和修改。

配置分类:
- 应用程序信息
- 窗口配置
- 文件配置
- 样式配置
- 性能配置
- 正则表达式模式
"""

# =============================================================================
# 应用程序信息配置
# =============================================================================
APP_NAME = "F-Excel"
APP_VERSION = "2.0.0"
APP_DESCRIPTION = "模块化数据处理工具集合"
APP_AUTHOR = "F-Excel Team"
APP_COPYRIGHT = "Copyright © 2024 F-Excel Team"

# =============================================================================
# 窗口配置
# =============================================================================
# 主窗口配置
MAIN_WINDOW_WIDTH = 800
MAIN_WINDOW_HEIGHT = 600
MAIN_WINDOW_MIN_WIDTH = 800
MAIN_WINDOW_MIN_HEIGHT = 600

# 模块窗口配置
MODULE_WINDOW_WIDTH = 900
MODULE_WINDOW_HEIGHT = 700
MODULE_WINDOW_MIN_WIDTH = 800
MODULE_WINDOW_MIN_HEIGHT = 600

# 窗口位置配置
WINDOW_CENTER = True  # 是否自动居中
WINDOW_RESIZABLE = True  # 是否可调整大小

# =============================================================================
# 文件配置
# =============================================================================
# Excel文件配置
DEFAULT_EXCEL_SHEET_NAME = "ID值转换结果"
EXCEL_MAX_ROWS = 1000000  # Excel最大行数限制
EXCEL_MAX_COLUMNS = 1000  # Excel最大列数限制

# 支持的文件类型
SUPPORTED_FILE_TYPES = {
    'excel': ['.xlsx', '.xls'],
    'text': ['.txt', '.csv'],
    'csv': ['.csv'],
    'all': ['*.*']
}

# 文件大小限制（MB）
MAX_FILE_SIZE_MB = 100
LARGE_FILE_THRESHOLD_MB = 50

# 文件编码配置
DEFAULT_ENCODING = 'utf-8'
SUPPORTED_ENCODINGS = ['utf-8', 'gbk', 'gb2312', 'utf-16']

# =============================================================================
# 样式配置
# =============================================================================
STYLES = {
    # 字体配置
    'title_font': ('Arial', 16, 'bold'),
    'subtitle_font': ('Arial', 10),
    'module_font': ('Arial', 12, 'bold'),
    'button_font': ('Arial', 10),
    'status_font': ('Arial', 9),
    
    # 颜色配置
    'header_background': '#f0f0f0',
    'primary_color': '#007acc',
    'success_color': '#28a745',
    'warning_color': '#ffc107',
    'error_color': '#dc3545',
    
    # 间距配置
    'padding_small': 5,
    'padding_medium': 10,
    'padding_large': 20,
    
    # 边框配置
    'border_width': 2,
    'border_style': 'raised'
}

# =============================================================================
# 性能配置
# =============================================================================
# 批处理配置
BATCH_SIZE = 1000  # 批处理大小
LARGE_FILE_THRESHOLD = 10000  # 大文件阈值（行数）

# 异步处理配置
ASYNC_ENABLED = True  # 是否启用异步处理
ASYNC_THREAD_COUNT = 4  # 异步线程数
ASYNC_TIMEOUT = 300  # 异步超时时间（秒）

# 内存管理配置
MEMORY_LIMIT_MB = 512  # 内存使用限制（MB）
CACHE_SIZE = 1000  # 缓存大小

# =============================================================================
# 正则表达式模式配置
# =============================================================================
# ID值转换模式
ID_VALUE_PATTERN = r'(\d+)=([^\s\n]+)'
ID_VALUE_PATTERN_STRICT = r'^(\d+)=(.+)$'  # 严格模式

# Excel翻译模式
EXCEL_TRANSLATION_PATTERN = r'(\d+)\t(.+?)\t(.+?)$'
EXCEL_TRANSLATION_PATTERN_FLEXIBLE = r'(\d+)\s+([^\t]+)\s+([^\t]+)'

# 通用模式
NUMBER_PATTERN = r'\d+'
TEXT_PATTERN = r'[^\s\n]+'

# =============================================================================
# 日志配置
# =============================================================================
LOG_LEVEL = 'INFO'  # 日志级别
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_FILE = 'f_excel.log'  # 日志文件名
LOG_MAX_SIZE = 10 * 1024 * 1024  # 日志文件最大大小（10MB）
LOG_BACKUP_COUNT = 5  # 日志备份文件数量

# =============================================================================
# 错误处理配置
# =============================================================================
# 错误重试配置
MAX_RETRY_COUNT = 3  # 最大重试次数
RETRY_DELAY_SECONDS = 1  # 重试延迟时间（秒）

# 错误提示配置
SHOW_ERROR_DETAILS = True  # 是否显示详细错误信息
ERROR_LOG_ENABLED = True  # 是否启用错误日志

# =============================================================================
# 国际化配置
# =============================================================================
DEFAULT_LANGUAGE = 'zh_CN'  # 默认语言
SUPPORTED_LANGUAGES = ['zh_CN', 'en_US', 'ko_KR']  # 支持的语言

# 语言包配置
LANGUAGE_PACKS = {
    'zh_CN': {
        'app_name': 'F-Excel',
        'app_description': '模块化数据处理工具集合',
        'ready': '就绪',
        'error': '错误',
        'warning': '警告',
        'success': '成功',
        'confirm': '确认',
        'cancel': '取消'
    },
    'en_US': {
        'app_name': 'F-Excel',
        'app_description': 'Modular Data Processing Tool Collection',
        'ready': 'Ready',
        'error': 'Error',
        'warning': 'Warning',
        'success': 'Success',
        'confirm': 'Confirm',
        'cancel': 'Cancel'
    }
}

# =============================================================================
# 开发配置
# =============================================================================
# 调试模式配置
DEBUG_MODE = False  # 是否启用调试模式
VERBOSE_LOGGING = False  # 是否启用详细日志

# 性能监控配置
PERFORMANCE_MONITORING = False  # 是否启用性能监控
PROFILING_ENABLED = False  # 是否启用性能分析

# 测试配置
TEST_MODE = False  # 是否启用测试模式
MOCK_DATA_ENABLED = False  # 是否启用模拟数据
