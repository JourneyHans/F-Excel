# F-Excel 项目结构说明

## 目录结构

```
F-Excel/
├── main.py                 # 主程序入口
├── run.py                  # 启动脚本
├── config.py               # 配置文件
├── requirements.txt        # 依赖包列表
├── README.md              # 项目说明文档
├── PROJECT_STRUCTURE.md   # 项目结构说明
├── start.bat              # Windows启动脚本
├── start.sh               # Linux/Mac启动脚本
├── test_id_converter.py   # 测试文件
├── modules/               # 模块目录
│   ├── __init__.py        # 模块包初始化
│   ├── id_converter.py    # ID转换器模块
│   └── excel_translator.py # Excel翻译器模块
└── examples/              # 示例文件目录
    ├── sample_data.txt    # 示例数据文件
    └── sample_translation_data.txt # 翻译数据示例文件
```

## 文件说明

### 核心文件
- **main.py**: 主程序文件，包含应用程序的主窗口和模块管理
- **run.py**: 启动脚本，包含依赖检查和错误处理
- **config.py**: 配置文件，包含应用程序的各种配置参数

### 模块文件
- **modules/id_converter.py**: ID转换器模块，实现【id=值】到Excel的转换功能
- **modules/excel_translator.py**: Excel翻译器模块，实现Excel文件到ID=翻译文本的转换功能

### 启动脚本
- **start.bat**: Windows系统的批处理启动脚本
- **start.sh**: Linux/Mac系统的Shell启动脚本

### 文档和配置
- **README.md**: 项目说明和使用指南
- **requirements.txt**: Python依赖包列表
- **examples/sample_data.txt**: 示例数据文件
- **examples/sample_translation_data.txt**: 翻译数据示例文件

## 技术架构

### 前端界面
- 使用tkinter作为GUI框架
- 模块化设计，每个功能独立成模块
- 响应式布局，支持窗口大小调整

### 数据处理
- 使用pandas进行数据处理
- 使用openpyxl进行Excel文件操作
- 正则表达式进行文本解析

### 模块系统
- 插件式架构，易于扩展新功能
- 统一的模块接口
- 独立的模块窗口管理

## 扩展指南

### 添加新模块
1. 在`modules/`目录下创建新的模块文件
2. 实现模块类，继承自基础模块接口
3. 在`main.py`中注册新模块
4. 更新模块配置和界面

### 修改配置
- 编辑`config.py`文件中的配置参数
- 支持运行时配置修改
- 配置文件热重载

### 自定义样式
- 在`config.py`中定义样式参数
- 支持主题切换
- 可自定义字体、颜色等
