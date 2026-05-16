import os
import openpyxl
from ..utils import log_message, log_error
from ..config import DEFAULT_OUTPUT_DIR as OUTPUT_DIR, DEFAULT_TEMPLATE_DIR as TEMPLATE_DIR

class StorageCalculator:
    def __init__(self, output_dir=None, template_dir=None):
        self.results = {
            'success': [],
            'failed': [],
            'total': 0
        }
        self.storage_data = None
        self.output_dir = output_dir if output_dir else OUTPUT_DIR
        self.template_dir = template_dir if template_dir else TEMPLATE_DIR
    
    def load_storage_data(self):
        """加载库容数据"""
        try:
            data_path = os.path.join(self.template_dir, '库容数据.xlsx')
            if not os.path.exists(data_path):
                data_path = os.path.join(self.template_dir, '库容数据.xlsm')
                if not os.path.exists(data_path):
                    raise FileNotFoundError(f"库容数据文件不存在: {data_path}")
            
            wb = openpyxl.load_workbook(data_path, data_only=True)
            ws = wb.active
            
            self.storage_data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None:
                    self.storage_data.append(row)
            
            log_message("成功加载库容数据")
            return True
        except Exception as e:
            log_error("加载库容数据失败", e)
            return False
    
    def calculate_storage(self, report_file, progress_callback=None):
        """计算库容"""
        try:
            if not self.storage_data:
                if not self.load_storage_data():
                    raise ValueError("无法加载库容数据")
            
            if not os.path.exists(report_file):
                raise FileNotFoundError(f"文件不存在: {report_file}")
            
            wb = openpyxl.load_workbook(report_file)
            ws = wb.active
            
            results = []
            for storage_item in self.storage_data:
                elevation = storage_item[0]
                area = storage_item[1] if len(storage_item) > 1 else 0
                capacity = area * 0.1 if isinstance(area, (int, float)) else 0
                results.append([elevation, area, capacity])
            
            if '库容计算结果' in wb.sheetnames:
                wb.remove(wb['库容计算结果'])
            
            result_ws = wb.create_sheet(title="库容计算结果")
            result_ws.append(['高程', '面积', '库容'])
            for row in results:
                result_ws.append(row)
            
            file_name = os.path.basename(report_file)
            output_name = f"{os.path.splitext(file_name)[0]}_库容计算.xlsx"
            output_path = os.path.join(self.output_dir, output_name)
            
            wb.save(output_path)
            self.results['success'].append(output_name)
            log_message(f"成功计算库容: {file_name}")
            
            if progress_callback:
                progress_callback(f"已计算库容: {file_name}")
            
            return True
        except Exception as e:
            file_name = os.path.basename(report_file)
            self.results['failed'].append((file_name, str(e)))
            log_error(f"计算库容失败: {file_name}", e)
            return False
    
    def process_all(self, report_files, progress_callback=None):
        """批量计算库容"""
        self.results = {'success': [], 'failed': [], 'total': len(report_files)}
        
        if not self.storage_data:
            if not self.load_storage_data():
                return self.results
        
        for report_file in report_files:
            self.calculate_storage(report_file, progress_callback)
        
        return self.results