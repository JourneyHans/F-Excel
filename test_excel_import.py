#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Excel文件导入时的ID类型处理
"""

import pandas as pd
import tempfile
import os

def test_excel_import():
    """测试Excel文件导入时的ID类型处理"""
    print("测试Excel文件导入时的ID类型处理...")
    
    # 创建测试数据
    test_data = {
        'ID': [999470001, 999470002, 999470003, 999470004, 999470005],
        'Chinese': ['反馈问题', '画质:', '上传日志:', '异常上报成功', '确认'],
        'Korean': ['피드백', '해상도:', '로그 업로드:', '업로드 성공!', '확인']
    }
    
    # 创建DataFrame
    df = pd.DataFrame(test_data)
    
    print("原始DataFrame:")
    print(df)
    print(f"ID列数据类型: {df['ID'].dtype}")
    print(f"ID列前3个值的类型: {[type(x) for x in df['ID'].head(3)]}")
    
    # 创建临时Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        excel_path = tmp.name
    
    try:
        # 保存为Excel文件
        df.to_excel(excel_path, index=False)
        print(f"\nExcel文件已保存到: {excel_path}")
        
        # 测试不同的读取方式
        print("\n" + "="*50)
        print("测试1: 默认读取方式")
        df_default = pd.read_excel(excel_path)
        print(f"ID列数据类型: {df_default['ID'].dtype}")
        print(f"ID列前3个值的类型: {[type(x) for x in df_default['ID'].head(3)]}")
        
        print("\n测试2: 指定ID列为字符串类型")
        df_string = pd.read_excel(excel_path, dtype={0: str})
        print(f"ID列数据类型: {df_string['ID'].dtype}")
        print(f"ID列前3个值的类型: {[type(x) for x in df_string['ID'].head(3)]}")
        
        print("\n测试3: 手动转换ID列为字符串")
        df_manual = pd.read_excel(excel_path)
        df_manual['ID'] = df_manual['ID'].astype(str)
        print(f"ID列数据类型: {df_manual['ID'].dtype}")
        print(f"ID列前3个值的类型: {[type(x) for x in df_manual['ID'].head(3)]}")
        
        # 测试字符串操作
        print("\n" + "="*50)
        print("测试字符串操作:")
        
        # 默认读取（数值类型）
        print("默认读取 - 测试isdigit():")
        for i, id_val in enumerate(df_default['ID'].head(3)):
            try:
                is_digit = str(id_val).isdigit()
                print(f"  ID: {id_val}, 类型: {type(id_val).__name__}, 转换为字符串后isdigit(): {is_digit}")
            except Exception as e:
                print(f"  ID: {id_val}, 类型: {type(id_val).__name__}, 错误: {e}")
        
        # 字符串类型读取
        print("\n字符串类型读取 - 测试isdigit():")
        for i, id_val in enumerate(df_string['ID'].head(3)):
            try:
                is_digit = id_val.isdigit()
                print(f"  ID: {id_val}, 类型: {type(id_val).__name__}, isdigit(): {is_digit}")
            except Exception as e:
                print(f"  ID: {id_val}, 类型: {type(id_val).__name__}, 错误: {e}")
        
    finally:
        # 清理临时文件
        if os.path.exists(excel_path):
            os.unlink(excel_path)
            print(f"\n临时文件已删除: {excel_path}")
    
    print("\n测试完成!")

if __name__ == "__main__":
    test_excel_import()
