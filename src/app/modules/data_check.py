import os
import openpyxl
from ..utils import read_csv_file, log_message, log_error
from ..config import DEFAULT_OUTPUT_DIR as OUTPUT_DIR

class DataChecker:
    def __init__(self):
        self.results = {
            'success': [],
            'failed': [],
            'warnings': []
        }
    
    def check_section(self, csv_file, progress_callback=None):
        """检查断面数据"""
        try:
            csv_data = read_csv_file(csv_file)
            if not csv_data:
                raise ValueError("CSV文件为空")
            
            errors = []
            for row_idx, row in enumerate(csv_data, 1):
                if not row or all(cell.strip() == '' for cell in row):
                    errors.append(f"第{row_idx}行为空")
                
                for col_idx, cell in enumerate(row, 1):
                    if cell:
                        try:
                            float(cell)
                        except ValueError:
                            errors.append(f"第{row_idx}行第{col_idx}列数据格式不正确")
            
            file_name = os.path.basename(csv_file)
            
            if errors:
                self.results['warnings'].append((file_name, errors))
                log_message(f"断面检查警告: {file_name} - {'; '.join(errors)}")
            else:
                self.results['success'].append(file_name)
                log_message(f"断面检查通过: {file_name}")
            
            if progress_callback:
                progress_callback(f"已检查: {file_name}")
            
            return len(errors) == 0
        except Exception as e:
            file_name = os.path.basename(csv_file)
            self.results['failed'].append((file_name, str(e)))
            log_error(f"断面检查失败: {file_name}", e)
            return False
    
    def find_empty_sections(self, csv_files, progress_callback=None):
        """查找空白断面"""
        try:
            empty_files = []
            for csv_file in csv_files:
                csv_data = read_csv_file(csv_file)
                if not csv_data or all(not row or all(cell.strip() == '' for cell in row) for row in csv_data):
                    empty_files.append(os.path.basename(csv_file))
            
            if empty_files:
                self.results['warnings'].append(('空白断面', empty_files))
                log_message(f"发现空白断面: {', '.join(empty_files)}")
            
            if progress_callback:
                progress_callback(f"空白断面查找完成，发现{len(empty_files)}个空白断面")
            
            return empty_files
        except Exception as e:
            log_error("查找空白断面失败", e)
            return []
    
    def check_roughness(self, csv_file, progress_callback=None):
        """检查糙率数据"""
        try:
            csv_data = read_csv_file(csv_file)
            if not csv_data:
                raise ValueError("CSV文件为空")
            
            roughness_values = []
            for row in csv_data:
                for cell in row:
                    if cell:
                        try:
                            val = float(cell)
                            if 0 < val < 1:
                                roughness_values.append(val)
                        except ValueError:
                            pass
            
            file_name = os.path.basename(csv_file)
            
            if not roughness_values:
                self.results['warnings'].append((file_name, ['未找到有效的糙率值']))
                log_message(f"糙率检查警告: {file_name} - 未找到有效的糙率值")
            else:
                self.results['success'].append(file_name)
                log_message(f"糙率检查通过: {file_name}")
            
            if progress_callback:
                progress_callback(f"已检查糙率: {file_name}")
            
            return len(roughness_values) > 0
        except Exception as e:
            file_name = os.path.basename(csv_file)
            self.results['failed'].append((file_name, str(e)))
            log_error(f"糙率检查失败: {file_name}", e)
            return False
    
    def generate_report(self):
        """生成检查报告"""
        try:
            report_content = []
            report_content.append("=" * 50)
            report_content.append("数据检查报告")
            report_content.append("=" * 50)
            report_content.append(f"检查通过: {len(self.results['success'])}")
            report_content.append(f"检查失败: {len(self.results['failed'])}")
            report_content.append(f"警告项: {len(self.results['warnings'])}")
            report_content.append("")
            
            if self.results['success']:
                report_content.append("通过文件:")
                for item in self.results['success']:
                    report_content.append(f"  ✓ {item}")
                report_content.append("")
            
            if self.results['failed']:
                report_content.append("失败文件:")
                for item, error in self.results['failed']:
                    report_content.append(f"  ✗ {item}: {error}")
                report_content.append("")
            
            if self.results['warnings']:
                report_content.append("警告信息:")
                for item, warnings in self.results['warnings']:
                    report_content.append(f"  ⚠ {item}:")
                    if isinstance(warnings, list):
                        for warning in warnings:
                            report_content.append(f"    - {warning}")
                    else:
                        report_content.append(f"    - {warnings}")
            
            report_path = os.path.join(OUTPUT_DIR, '数据检查报告.txt')
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(report_content))
            
            log_message("成功生成数据检查报告")
            return report_path
        except Exception as e:
            log_error("生成检查报告失败", e)
            return None