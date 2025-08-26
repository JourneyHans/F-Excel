#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
F-Excel 启动脚本
"""

import sys
import os

def check_dependencies():
    """检查依赖包是否已安装"""
    required_packages = ['pandas', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("缺少以下依赖包:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\n请运行以下命令安装依赖:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    return True

def main():
    """主函数"""
    print("F-Excel 桌面应用程序启动中...")
    
    # 检查依赖
    if not check_dependencies():
        input("按回车键退出...")
        sys.exit(1)
    
    try:
        # 导入并运行主程序
        from main import main as run_app
        run_app()
    except ImportError as e:
        print(f"导入错误: {e}")
        print("请确保所有文件都在正确的位置")
        input("按回车键退出...")
        sys.exit(1)
    except Exception as e:
        print(f"程序运行错误: {e}")
        input("按回车键退出...")
        sys.exit(1)

if __name__ == "__main__":
    main()
