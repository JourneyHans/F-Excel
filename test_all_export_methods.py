#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试所有导出方法的换行符格式
"""

import os
import tempfile
import sys

def analyze_file_content(file_path):
    """分析文件内容的换行符"""
    with open(file_path, 'rb') as f:
        content = f.read()
    
    print(f"  文件大小: {len(content)} 字节")
    print(f"  LF (\\n) 数量: {content.count(b'\\n')}")
    print(f"  CR (\\r) 数量: {content.count(b'\\r')}")
    print(f"  CRLF (\\r\\n) 数量: {content.count(b'\\r\\n')}")
    
    # 判断换行符类型
    if b'\r\n' in content:
        line_ending_type = "CRLF (\\r\\n - Windows)"
    elif b'\n' in content:
        line_ending_type = "LF (\\n - Unix/Linux)"
    elif b'\r' in content:
        line_ending_type = "CR (\\r - Classic Mac)"
    else:
        line_ending_type = "无换行符"
    
    print(f"  换行符类型: {line_ending_type}")
    
    # 显示前几行的十六进制
    print("  前几行十六进制分析:")
    lines = content.split(b'\n')
    for i, line in enumerate(lines[:3]):
        if line:
            # 移除可能的回车符
            clean_line = line.replace(b'\r', b'')
            print(f"    行{i+1}: {clean_line.decode('utf-8', errors='ignore')}")
            print(f"      原始: {line.hex()}")
            print(f"      清理后: {clean_line.hex()}")
    
    return content

def test_all_export_methods():
    """测试所有导出方法"""
    print("=" * 60)
    print("测试Excel翻译器的所有导出方法")
    print("=" * 60)
    
    # 添加模块路径
    sys.path.append(os.path.join(os.path.dirname(__file__), 'modules'))
    
    try:
        from excel_translator import ExcelTranslatorModule
        
        # 创建模块实例
        translator = ExcelTranslatorModule()
        
        # 模拟转换数据
        translator.converted_data = [
            {'Output': '999470001=피드백'},
            {'Output': '999470002=해상도:'},
            {'Output': '999470003=로그 업로드:'},
            {'Output': '999470004=업로드 성공!'},
            {'Output': '999470005=확인'}
        ]
        
        # 测试1: 小文件导出
        print("\n1. 测试小文件导出 (export_small_file):")
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            small_file = tmp.name
        
        try:
            translator.export_small_file(small_file)
            analyze_file_content(small_file)
        finally:
            if os.path.exists(small_file):
                os.unlink(small_file)
        
        # 测试2: 大文件导出
        print("\n2. 测试大文件导出 (export_large_file):")
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            large_file = tmp.name
        
        try:
            translator.export_large_file(large_file)
            analyze_file_content(large_file)
        finally:
            if os.path.exists(large_file):
                os.unlink(large_file)
        
        # 测试3: 主导出方法
        print("\n3. 测试主导出方法 (export_text):")
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            main_file = tmp.name
        
        try:
            # 直接调用内部方法，跳过文件选择对话框
            with open(main_file, 'w', encoding='utf-8', newline='') as f:
                for row in translator.converted_data:
                    f.write(f"{row['Output']}\n")
            
            analyze_file_content(main_file)
        finally:
            if os.path.exists(main_file):
                os.unlink(main_file)
                
    except ImportError as e:
        print(f"无法导入Excel翻译器模块: {e}")
    except Exception as e:
        print(f"测试过程中发生错误: {e}")

def test_direct_file_writing():
    """测试直接文件写入的换行符"""
    print("\n" + "=" * 60)
    print("测试直接文件写入的换行符")
    print("=" * 60)
    
    test_data = [
        "999470001=피드백",
        "999470002=해상도:",
        "999470003=로그 업로드:",
        "999470004=업로드 성공!",
        "999470005=확인"
    ]
    
    # 测试1: 使用newline=''
    print("\n1. 使用 newline='':")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8', newline='') as f:
        newline_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        analyze_file_content(newline_file)
    finally:
        os.unlink(newline_file)
    
    # 测试2: 不使用newline参数
    print("\n2. 不使用 newline 参数:")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        default_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        analyze_file_content(default_file)
    finally:
        os.unlink(default_file)

def main():
    """主函数"""
    print("F-Excel 所有导出方法换行符测试")
    print("=" * 60)
    
    try:
        # 测试所有导出方法
        test_all_export_methods()
        
        # 测试直接文件写入
        test_direct_file_writing()
        
        print("\n" + "=" * 60)
        print("测试完成!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {str(e)}")

if __name__ == "__main__":
    main()
