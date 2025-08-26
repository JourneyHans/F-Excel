#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试大文件性能的脚本
"""

import time
import tempfile
import os
import pandas as pd

def generate_large_test_data(rows=50000):
    """生成大量测试数据"""
    print(f"正在生成 {rows} 行测试数据...")
    
    # 生成ID列表
    ids = [f"99947{str(i).zfill(4)}" for i in range(1, rows + 1)]
    
    # 生成中文数据
    chinese_samples = [
        "反馈问题", "画质:", "上传日志:", "异常上报成功", "确认", 
        "实名认证", "切换账号", "标题", "用户名", "密码",
        "登录", "注册", "设置", "帮助", "关于"
    ]
    chinese = [chinese_samples[i % len(chinese_samples)] for i in range(rows)]
    
    # 生成韩文数据
    korean_samples = [
        "피드백", "해상도:", "로그 업로드:", "업로드 성공!", "확인",
        "본인인증", "계정 변경", "제목", "사용자명", "비밀번호",
        "로그인", "회원가입", "설정", "도움말", "정보"
    ]
    korean = [korean_samples[i % len(korean_samples)] for i in range(rows)]
    
    # 创建DataFrame
    df = pd.DataFrame({
        'ID': ids,
        'Chinese': chinese,
        'Korean': korean
    })
    
    return df

def test_file_generation():
    """测试文件生成性能"""
    print("=" * 60)
    print("测试文件生成性能")
    print("=" * 60)
    
    # 测试不同大小的文件
    test_sizes = [1000, 5000, 10000, 50000]
    
    for size in test_sizes:
        print(f"\n生成 {size} 行数据...")
        start_time = time.time()
        
        df = generate_large_test_data(size)
        
        # 创建临时Excel文件
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            excel_path = tmp.name
        
        try:
            # 保存为Excel文件
            df.to_excel(excel_path, index=False)
            
            # 获取文件大小
            file_size = os.path.getsize(excel_path)
            file_size_mb = file_size / (1024 * 1024)
            
            generation_time = time.time() - start_time
            
            print(f"  - 生成时间: {generation_time:.2f} 秒")
            print(f"  - 文件大小: {file_size_mb:.2f} MB")
            print(f"  - 平均速度: {size / generation_time:.0f} 行/秒")
            
        finally:
            # 清理临时文件
            if os.path.exists(excel_path):
                os.unlink(excel_path)

def test_file_reading():
    """测试文件读取性能"""
    print("\n" + "=" * 60)
    print("测试文件读取性能")
    print("=" * 60)
    
    # 生成测试文件
    df = generate_large_test_data(10000)
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        excel_path = tmp.name
    
    try:
        # 保存为Excel文件
        df.to_excel(excel_path, index=False)
        
        # 测试不同读取方式
        print(f"\n测试读取 {len(df)} 行数据...")
        
        # 测试1: 默认读取
        start_time = time.time()
        df_default = pd.read_excel(excel_path)
        default_time = time.time() - start_time
        print(f"  - 默认读取: {default_time:.2f} 秒")
        
        # 测试2: 指定类型读取
        start_time = time.time()
        df_typed = pd.read_excel(excel_path, dtype={0: str})
        typed_time = time.time() - start_time
        print(f"  - 指定类型读取: {typed_time:.2f} 秒")
        
        # 测试3: 分批读取（模拟分块处理）
        start_time = time.time()
        df_full = pd.read_excel(excel_path, dtype={0: str})
        # 模拟分批处理
        batch_size = 1000
        processed_rows = 0
        for i in range(0, len(df_full), batch_size):
            batch = df_full.iloc[i:i + batch_size]
            processed_rows += len(batch)
        chunked_time = time.time() - start_time
        print(f"  - 分批处理: {chunked_time:.2f} 秒")
        
        # 性能对比
        print(f"\n性能对比:")
        print(f"  - 默认读取 vs 指定类型: {default_time/typed_time:.2f}x")
        print(f"  - 默认读取 vs 分批处理: {default_time/chunked_time:.2f}x")
        
    finally:
        # 清理临时文件
        if os.path.exists(excel_path):
            os.unlink(excel_path)

def test_data_processing():
    """测试数据处理性能"""
    print("\n" + "=" * 60)
    print("测试数据处理性能")
    print("=" * 60)
    
    # 生成测试数据
    df = generate_large_test_data(10000)
    
    print(f"\n测试处理 {len(df)} 行数据...")
    
    # 测试1: 逐行处理
    start_time = time.time()
    results = []
    for _, row in df.iterrows():
        id_value = str(row.iloc[0])
        chinese = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        korean = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        
        if id_value.isdigit():
            results.append(f"{id_value}={korean}")
    
    row_by_row_time = time.time() - start_time
    print(f"  - 逐行处理: {row_by_row_time:.2f} 秒, {len(results)} 条结果")
    
    # 测试2: 向量化处理
    start_time = time.time()
    df_clean = df.copy()
    df_clean['ID'] = df_clean['ID'].astype(str)
    df_clean['Chinese'] = df_clean['Chinese'].fillna('').astype(str)
    df_clean['Korean'] = df_clean['Korean'].fillna('').astype(str)
    
    # 过滤有效ID
    mask = df_clean['ID'].str.isdigit()
    df_filtered = df_clean[mask]
    
    # 生成输出格式
    vectorized_results = df_filtered['ID'] + '=' + df_filtered['Korean']
    vectorized_time = time.time() - start_time
    print(f"  - 向量化处理: {vectorized_time:.2f} 秒, {len(vectorized_results)} 条结果")
    
    # 性能对比
    print(f"\n性能对比:")
    print(f"  - 逐行处理 vs 向量化: {row_by_row_time/vectorized_time:.2f}x")
    
    return results, vectorized_results.tolist()

def main():
    """主函数"""
    print("F-Excel 大文件性能测试")
    print("=" * 60)
    
    try:
        # 测试文件生成性能
        test_file_generation()
        
        # 测试文件读取性能
        test_file_reading()
        
        # 测试数据处理性能
        results1, results2 = test_data_processing()
        
        # 验证结果一致性
        if len(results1) == len(results2):
            print(f"\n✅ 结果验证: 两种处理方式结果数量一致 ({len(results1)} 条)")
        else:
            print(f"\n❌ 结果验证: 两种处理方式结果数量不一致 ({len(results1)} vs {len(results2)})")
        
        print("\n" + "=" * 60)
        print("性能测试完成!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {str(e)}")

if __name__ == "__main__":
    main()
