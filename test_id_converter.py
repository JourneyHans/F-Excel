#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ID转换器模块测试
"""

import unittest
import re
import tempfile
import os
import pandas as pd
from modules.id_converter import IDConverterModule

class TestIDConverter(unittest.TestCase):
    def setUp(self):
        """测试前的准备工作"""
        self.converter = IDConverterModule()
        
    def test_convert_id_values(self):
        """测试ID值转换功能"""
        # 测试数据
        test_text = """410325=提升{0}
410326=降低{0}
410327=无变化
410328=高于"""
        
        # 执行转换
        result = self.converter.convert_id_values(test_text)
        
        # 验证结果
        self.assertEqual(len(result), 4)
        self.assertEqual(result[0]['ID'], '410325')
        self.assertEqual(result[0]['Value'], '提升{0}')
        self.assertEqual(result[2]['ID'], '410327')
        self.assertEqual(result[2]['Value'], '无变化')
        
    def test_convert_id_values_empty(self):
        """测试空输入"""
        result = self.converter.convert_id_values("")
        self.assertEqual(len(result), 0)
        
    def test_convert_id_values_no_match(self):
        """测试无匹配内容"""
        result = self.converter.convert_id_values("这是普通文本，没有id=值格式")
        self.assertEqual(len(result), 0)
        
    def test_convert_id_values_mixed(self):
        """测试混合内容"""
        test_text = """普通文本
410325=提升{0}
更多文本
410326=降低{0}
结尾"""
        
        result = self.converter.convert_id_values(test_text)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]['Value'], '提升{0}')
        self.assertEqual(result[1]['Value'], '降低{0}')
        
    def test_export_excel(self):
        """测试Excel导出功能"""
        # 创建测试数据
        test_data = [
            {'ID': '001', 'Value': '001'},
            {'ID': '002', 'Value': '张三'},
            {'ID': '003', 'Value': '李四'}
        ]
        
        # 创建临时文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            temp_path = tmp_file.name
            
        try:
            # 测试DataFrame创建
            df = pd.DataFrame(test_data)
            self.assertEqual(len(df), 3)
            self.assertEqual(df.columns.tolist(), ['ID', 'Value'])
            
            # 测试Excel导出
            df.to_excel(temp_path, sheet_name='测试', index=False)
            
            # 验证文件是否创建
            self.assertTrue(os.path.exists(temp_path))
            
            # 验证文件内容
            df_read = pd.read_excel(temp_path, sheet_name='测试')
            self.assertEqual(len(df_read), 3)
            self.assertEqual(int(df_read.iloc[0]['ID']), 1)
            self.assertEqual(df_read.iloc[1]['Value'], '张三')
            
        finally:
            # 清理临时文件
            if os.path.exists(temp_path):
                os.unlink(temp_path)

if __name__ == '__main__':
    unittest.main()
