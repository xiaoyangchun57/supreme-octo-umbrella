import time
import psutil
import os
from datetime import datetime

class PerformanceMonitor:
    def __init__(self):
        self.start_time = None
        self.end_time = None
        self.start_memory = None
        self.peak_memory = 0
        self.results = {
            'success': 0,
            'failed': 0,
            'total': 0
        }
    
    def start(self):
        """开始监控"""
        self.start_time = time.time()
        self.start_memory = self._get_memory_usage()
        self.peak_memory = self.start_memory
    
    def end(self):
        """结束监控"""
        self.end_time = time.time()
    
    def update_memory(self):
        """更新内存峰值"""
        current_memory = self._get_memory_usage()
        if current_memory > self.peak_memory:
            self.peak_memory = current_memory
    
    def update_results(self, success, failed, total):
        """更新处理结果"""
        self.results['success'] = success
        self.results['failed'] = failed
        self.results['total'] = total
    
    def _get_memory_usage(self):
        """获取当前内存使用量（MB）"""
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / (1024 * 1024)
    
    def get_duration(self):
        """获取处理时长（秒）"""
        if self.start_time and self.end_time:
            return self.end_time - self.start_time
        return 0
    
    def generate_report(self, task_name):
        """生成性能报告"""
        duration = self.get_duration()
        duration_str = self._format_duration(duration)
        
        report = f"""
╔══════════════════════════════════════════════════════════════╗
║              性能报告 - {task_name}                            ║
╠══════════════════════════════════════════════════════════════╣
║ 执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}      ║
╠══════════════════════════════════════════════════════════════╣
║ 处理时长: {duration_str}                                      ║
║ 内存峰值: {self.peak_memory:.2f} MB                          ║
║ 内存增量: {(self.peak_memory - self.start_memory):.2f} MB    ║
╠══════════════════════════════════════════════════════════════╣
║ 处理总数: {self.results['total']}                             ║
║ 成功数量: {self.results['success']}                           ║
║ 失败数量: {self.results['failed']}                            ║
║ 成功率:   {self._calculate_success_rate():.1f}%              ║
╚══════════════════════════════════════════════════════════════╝
"""
        return report
    
    def _format_duration(self, seconds):
        """格式化时长"""
        if seconds < 60:
            return f"{seconds:.2f} 秒"
        elif seconds < 3600:
            minutes = int(seconds // 60)
            secs = seconds % 60
            return f"{minutes} 分 {secs:.2f} 秒"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            secs = seconds % 60
            return f"{hours} 时 {minutes} 分 {secs:.2f} 秒"
    
    def _calculate_success_rate(self):
        """计算成功率"""
        if self.results['total'] == 0:
            return 0
        return (self.results['success'] / self.results['total']) * 100
    
    def save_report(self, task_name, output_dir='.'):
        """保存性能报告到文件"""
        report = self.generate_report(task_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"性能报告_{task_name}_{timestamp}.txt"
        filepath = os.path.join(output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(report)
        
        return filepath

# 性能监控装饰器
def monitor_performance(task_name="任务"):
    def decorator(func):
        def wrapper(*args, **kwargs):
            monitor = PerformanceMonitor()
            monitor.start()
            
            try:
                result = func(*args, **kwargs)
                
                if isinstance(result, dict) and 'success' in result and 'failed' in result and 'total' in result:
                    monitor.update_results(result['success'], result['failed'], result['total'])
                
                return result
            finally:
                monitor.end()
                report = monitor.generate_report(task_name)
                print(report)
                
                from ..config import OUTPUT_DIR
                monitor.save_report(task_name, OUTPUT_DIR)
        
        return wrapper
    return decorator