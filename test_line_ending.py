#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试换行符格式的脚本
"""

import os
import tempfile

def test_line_endings():
    """测试换行符格式"""
    print("测试换行符格式...")
    
    # 测试数据
    test_data = [
        "999470001=피드백",
        "999470002=해상도:",
        "999470003=로그 업로드:",
        "999470004=업로드 성공!",
        "999470005=확인"
    ]
    
    # 测试1: 默认模式（可能产生\r\n）
    print("\n测试1: 默认模式")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        default_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        with open(default_file, 'rb') as f:
            content = f.read()
            print(f"  文件大小: {len(content)} 字节")
            print(f"  换行符数量: {content.count(b'\\n')}")
            print(f"  回车符数量: {content.count(b'\\r')}")
            print(f"  换行符类型: {'\\r\\n (Windows)' if b'\\r\\n' in content else '\\n (Unix)'}")
            
            # 显示前几行的十六进制
            print("  前几行十六进制:")
            for i, line in enumerate(content.split(b'\n')[:3]):
                if line:
                    print(f"    行{i+1}: {line.hex()}")
    finally:
        os.unlink(default_file)
    
    # 测试2: 指定newline=''模式（应该产生\n）
    print("\n测试2: 指定newline=''模式")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8', newline='') as f:
        newline_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        with open(newline_file, 'rb') as f:
            content = f.read()
            print(f"  文件大小: {len(content)} 字节")
            print(f"  换行符数量: {content.count(b'\\n')}")
            print(f"  回车符数量: {content.count(b'\\r')}")
            print(f"  换行符类型: {'\\r\\n (Windows)' if b'\\r\\n' in content else '\\n (Unix)'}")
            
            # 显示前几行的十六进制
            print("  前几行十六进制:")
            for i, line in enumerate(content.split(b'\n')[:3]):
                if line:
                    print(f"    行{i+1}: {line.hex()}")
    finally:
        os.unlink(newline_file)
    
    # 测试3: 手动写入\n
    print("\n测试3: 手动写入\\n")
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.txt', delete=False) as f:
        manual_file = f.name
        for line in test_data:
            f.write(f"{line}\n".encode('utf-8'))
    
    try:
        with open(manual_file, 'rb') as f:
            content = f.read()
            print(f"  文件大小: {len(content)} 字节")
            print(f"  换行符数量: {content.count(b'\\n')}")
            print(f"  回车符数量: {content.count(b'\\r')}")
            print(f"  换行符类型: {'\\r\\n (Windows)' if b'\\r\\n' in content else '\\n (Unix)'}")
            
            # 显示前几行的十六进制
            print("  前几行十六进制:")
            for i, line in enumerate(content.split(b'\n')[:3]):
                if line:
                    print(f"    行{i+1}: {line.hex()}")
    finally:
        os.unlink(manual_file)

def test_excel_translator_export():
    """测试Excel翻译器的导出功能"""
    print("\n" + "="*60)
    print("测试Excel翻译器的导出功能")
    print("="*60)
    
    import sys
    import os
    
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
        
        # 创建临时文件
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            test_file = tmp.name
        
        try:
            # 测试导出
            print(f"测试导出到: {test_file}")
            translator.export_small_file(test_file)
            
            # 检查文件内容
            with open(test_file, 'rb') as f:
                content = f.read()
                print(f"\n导出文件信息:")
                print(f"  文件大小: {len(content)} 字节")
                print(f"  换行符数量: {content.count(b'\\n')}")
                print(f"  回车符数量: {content.count(b'\\r')}")
                print(f"  换行符类型: {'\\r\\n (Windows)' if b'\\r\\n' in content else '\\n (Unix)'}")
                
                # 显示内容
                print(f"\n文件内容:")
                print(content.decode('utf-8'))
                
        finally:
            # 清理临时文件
            if os.path.exists(test_file):
                os.unlink(test_file)
                
    except ImportError as e:
        print(f"无法导入Excel翻译器模块: {e}")
    except Exception as e:
        print(f"测试过程中发生错误: {e}")

if __name__ == "__main__":
    test_line_endings()
    test_excel_translator_export()
