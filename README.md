# F-Excel 桌面应用程序

一个模块化的桌面工具集合，提供各种实用的数据处理功能。

## 功能模块

### 模块1：ID值转换器
- 将数字=值格式的文本转换为两列Excel文件
- 支持批量处理和文件导入导出

### 模块2：Excel翻译器
- 将Excel文件中的ID、中文、韩文三列转换为ID=韩文格式
- 支持Excel文件导入和文本文件导出
- 支持制表符分隔的文本数据输入

## 安装说明

1. 确保已安装Python 3.7+
2. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

运行主程序：
```bash
python main.py
```

## 技术栈

- Python 3.7+
- tkinter (GUI框架)
- pandas (数据处理)
- openpyxl (Excel文件操作)
- pillow (图像处理)
