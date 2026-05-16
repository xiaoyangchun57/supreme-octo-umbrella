#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
性能基准测试脚本
用于测试系统处理大规模数据的性能
"""

import os
import sys
import time
import psutil
import shutil
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from app.modules.auto_plot import AutoPlotter
from app.modules.plot_summary import PlotSummary
from app.config import OUTPUT_DIR, TEMPLATE_DIR, INPUT_DIR
from app.utils.performance_monitor import PerformanceMonitor

def generate_test_data(file_count=10):
    print(f"正在生成 {file_count} 个测试文件...")
    
    existing_files = []
    for f in os.listdir(OUTPUT_DIR):
        if '成果表' in f and f.endswith('.xlsx'):
            existing_files.append(os.path.join(OUTPUT_DIR, f))
    
    if not existing_files:
        print("错误：未找到成果表文件，请先运行CSV转成果表功能")
        return []
    
    test_dir = os.path.join(OUTPUT_DIR, 'performance_test')
    os.makedirs(test_dir, exist_ok=True)
    
    test_files = []
    source_file = existing_files[0]
    for i in range(file_count):
        base_name = os.path.basename(source_file)
        name_parts = base_name.split('_')
        if len(name_parts) > 1:
            new_name = f"{name_parts[0]}_{i:04d}_{name_parts[-1]}"
        else:
            new_name = f"测试成果表_{i:04d}.xlsx"
        
        dest_path = os.path.join(test_dir, new_name)
        shutil.copy(source_file, dest_path)
        test_files.append(dest_path)
    
    print(f"已生成 {len(test_files)} 个测试文件")
    return test_files

def test_auto_plot_performance(test_files):
    print("\n" + "="*60)
    print("测试自动成图性能")
    print("="*60)
    
    monitor = PerformanceMonitor()
    monitor.start()
    
    plotter = AutoPlotter()
    
    try:
        results = plotter.process_all(test_files, max_workers=4)
        monitor.update_results(len(results['success']), len(results['failed']), results['total'])
    finally:
        monitor.end()
    
    report = monitor.generate_report("自动成图")
    print(report)
    monitor.save_report("自动成图", OUTPUT_DIR)
    
    return monitor.get_duration(), monitor.peak_memory

def test_export_txt_performance(test_files):
    print("\n" + "="*60)
    print("测试导出TXT性能")
    print("="*60)
    
    monitor = PerformanceMonitor()
    monitor.start()
    
    exporter = PlotSummary()
    
    try:
        exporter.process_all(test_files)
        results = exporter.results
        monitor.update_results(len(results['success']), len(results['failed']), results['total'])
    finally:
        monitor.end()
    
    report = monitor.generate_report("导出TXT")
    print(report)
    monitor.save_report("导出TXT", OUTPUT_DIR)
    
    return monitor.get_duration(), monitor.peak_memory

def run_benchmark(file_counts=[10, 50, 100]):
    print("\n" + "="*60)
    print("性能基准测试开始")
    print("="*60)
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"CPU核心数: {psutil.cpu_count()}")
    print(f"可用内存: {psutil.virtual_memory().available / (1024**3):.2f} GB")
    
    max_files = max(file_counts)
    test_files = generate_test_data(max_files)
    
    if not test_files:
        return
    
    results = []
    for file_count in file_counts:
        print(f"\n" + "-"*60)
        print(f"测试规模: {file_count} 个文件")
        print("-"*60)
        
        files_to_test = test_files[:file_count]
        
        plot_time, plot_memory = test_auto_plot_performance(files_to_test)
        txt_time, txt_memory = test_export_txt_performance(files_to_test)
        
        results.append({
            'file_count': file_count,
            'plot_time': plot_time,
            'plot_memory': plot_memory,
            'txt_time': txt_time,
            'txt_memory': txt_memory
        })
    
    generate_benchmark_report(results)

def generate_benchmark_report(results):
    print("\n" + "="*60)
    print("性能基准测试报告")
    print("="*60)
    
    report = f"""
性能基准测试报告
================

测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
CPU核心数: {psutil.cpu_count()}
可用内存: {psutil.virtual_memory().available / (1024**3):.2f} GB

测试结果
--------

| 文件数量 | 自动成图时间 | 自动成图内存 | 导出TXT时间 | 导出TXT内存 |
|---------|-------------|-------------|-------------|------------|
"""
    
    for result in results:
        report += f"| {result['file_count']:>8} | {result['plot_time']:>11.2f}s | {result['plot_memory']:>11.2f}MB | {result['txt_time']:>11.2f}s | {result['txt_memory']:>10.2f}MB |\n"
    
    report += f"""

性能评估
--------
"""
    
    for result in results:
        if result['file_count'] == 100:
            plot_ok = "✓" if result['plot_time'] <= 120 else "✗"
            plot_mem_ok = "✓" if result['plot_memory'] <= 512 else "✗"
            txt_ok = "✓" if result['txt_time'] <= 60 else "✗"
            
            report += f"""
100个文件处理目标:
- 自动成图时间 ≤ 2分钟: {plot_ok} ({result['plot_time']:.2f}秒)
- 自动成图内存 ≤ 512MB: {plot_mem_ok} ({result['plot_memory']:.2f}MB)
- 导出TXT时间 ≤ 1分钟: {txt_ok} ({result['txt_time']:.2f}秒)
"""
    
    print(report)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(OUTPUT_DIR, f"性能基准测试报告_{timestamp}.txt")
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n报告已保存到: {report_path}")

if __name__ == "__main__":
    file_counts = [10, 50, 100]
    if len(sys.argv) > 1:
        try:
            file_counts = [int(x) for x in sys.argv[1:]]
        except ValueError:
            print("用法: python performance_test.py [文件数量1] [文件数量2] ...")
            print("示例: python performance_test.py 10 50 100")
            sys.exit(1)
    
    run_benchmark(file_counts)