#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细测试换行符格式的脚本
"""

import os
import tempfile
import binascii

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

def test_different_write_modes():
    """测试不同的写入模式"""
    print("=" * 60)
    print("测试不同的文件写入模式")
    print("=" * 60)
    
    test_data = [
        "999470001=피드백",
        "999470002=해상도:",
        "999470003=로그 업로드:",
        "999470004=업로드 성공!",
        "999470005=확인"
    ]
    
    # 测试1: 默认文本模式
    print("\n1. 默认文本模式 (可能产生\\r\\n)")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        default_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        analyze_file_content(default_file)
    finally:
        os.unlink(default_file)
    
    # 测试2: 指定newline=''
    print("\n2. 指定newline='' (应该产生\\n)")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8', newline='') as f:
        newline_file = f.name
        for line in test_data:
            f.write(f"{line}\n")
    
    try:
        analyze_file_content(newline_file)
    finally:
        os.unlink(newline_file)
    
    # 测试3: 二进制模式手动写入
    print("\n3. 二进制模式手动写入\\n")
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.txt', delete=False) as f:
        binary_file = f.name
        for line in test_data:
            f.write(f"{line}\n".encode('utf-8'))
    
    try:
        analyze_file_content(binary_file)
    finally:
        os.unlink(binary_file)
    
    # 测试4: 使用os.linesep
    print("\n4. 使用os.linesep (系统默认)")
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        linesep_file = f.name
        for line in test_data:
            f.write(f"{line}{os.linesep}")
    
    try:
        analyze_file_content(linesep_file)
    finally:
        os.unlink(linesep_file)

def test_excel_translator_export():
    """测试Excel翻译器的导出功能"""
    print("\n" + "=" * 60)
    print("测试Excel翻译器的导出功能")
    print("=" * 60)
    
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
        
        # 测试小文件导出
        print("测试小文件导出:")
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            small_file = tmp.name
        
        try:
            translator.export_small_file(small_file)
            analyze_file_content(small_file)
        finally:
            if os.path.exists(small_file):
                os.unlink(small_file)
        
        # 测试大文件导出
        print("\n测试大文件导出:")
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            large_file = tmp.name
        
        try:
            translator.export_large_file(large_file)
            analyze_file_content(large_file)
        finally:
            if os.path.exists(large_file):
                os.unlink(large_file)
                
    except ImportError as e:
        print(f"无法导入Excel翻译器模块: {e}")
    except Exception as e:
        print(f"测试过程中发生错误: {e}")

def test_cross_platform_compatibility():
    """测试跨平台兼容性"""
    print("\n" + "=" * 60)
    print("测试跨平台兼容性")
    print("=" * 60)
    
    test_data = [
        "999470001=피드백",
        "999470002=해상도:",
        "999470003=로그 업로드:"
    ]
    
    # 测试在不同平台上的表现
    platforms = [
        ("Windows风格", "\r\n"),
        ("Unix风格", "\n"),
        ("Classic Mac风格", "\r")
    ]
    
    for platform_name, line_sep in platforms:
        print(f"\n{platform_name} (使用 {repr(line_sep)}):")
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8', newline='') as f:
            platform_file = f.name
            for line in test_data:
                f.write(f"{line}{line_sep}")
        
        try:
            analyze_file_content(platform_file)
        finally:
            os.unlink(platform_file)

def main():
    """主函数"""
    print("F-Excel 换行符格式详细测试")
    print("=" * 60)
    
    try:
        # 测试不同的写入模式
        test_different_write_modes()
        
        # 测试Excel翻译器导出
        test_excel_translator_export()
        
        # 测试跨平台兼容性
        test_cross_platform_compatibility()
        
        print("\n" + "=" * 60)
        print("测试完成!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {str(e)}")

if __name__ == "__main__":
    main()
