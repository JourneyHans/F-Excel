#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel翻译器模块测试文件
"""

import sys
import os

# 添加模块路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'modules'))

from excel_translator import ExcelTranslatorModule

def test_excel_translator():
    """测试Excel翻译器模块"""
    print("开始测试Excel翻译器模块...")
    
    # 创建模块实例
    translator = ExcelTranslatorModule()
    
    # 测试数据转换
    test_data = """999470001	反馈问题	피드백
999470002	画质:	해상도:
999470003	上传日志:	로그 업로드:
999470004	异常上报成功，感谢团长对阿克迈斯的关注！	업로드 성공!
999470005	确认	확인
999470006	实名认证	본인인증
999470007	切换账号	계정 변경
999470008	标题	제목"""
    
    print("测试数据:")
    print(test_data)
    print("\n" + "="*50 + "\n")
    
    # 测试转换功能
    converted_data = translator.convert_excel_data(test_data)
    
    print(f"转换结果 (共{len(converted_data)}条):")
    for i, row in enumerate(converted_data, 1):
        print(f"{i}. {row['Output']}")
        # 验证ID类型
        print(f"   ID类型: {type(row['ID']).__name__}, 值: {row['ID']}")
    
    print("\n" + "="*50 + "\n")
    
    # 测试输出格式
    print("纯输出格式:")
    for row in converted_data:
        print(row['Output'])
    
    print("\n" + "="*50 + "\n")
    
    # 测试ID类型处理
    print("ID类型验证:")
    for row in converted_data:
        id_type = type(row['ID']).__name__
        id_value = row['ID']
        is_string = isinstance(row['ID'], str)
        is_digit = row['ID'].isdigit() if isinstance(row['ID'], str) else False
        
        print(f"ID: {id_value}, 类型: {id_type}, 是字符串: {is_string}, 是数字: {is_digit}")
    
    print("\n测试完成!")

if __name__ == "__main__":
    test_excel_translator()
