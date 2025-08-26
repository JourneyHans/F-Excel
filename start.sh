#!/bin/bash

echo "========================================"
echo "           F-Excel 启动器"
echo "========================================"
echo

echo "正在检查Python环境..."
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3环境"
    echo "请先安装Python 3.7或更高版本"
    echo
    read -p "按回车键退出..."
    exit 1
fi

echo "Python环境检查通过"
echo

echo "正在检查依赖包..."
if ! python3 -c "import pandas, openpyxl" &> /dev/null; then
    echo "正在安装依赖包..."
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "依赖包安装失败"
        read -p "按回车键退出..."
        exit 1
    fi
    echo "依赖包安装完成"
    echo
fi

echo "启动F-Excel应用程序..."
echo
python3 run.py

if [ $? -ne 0 ]; then
    echo
    echo "程序运行出错，请检查错误信息"
    read -p "按回车键退出..."
fi

echo
echo "程序已退出"
read -p "按回车键退出..."
