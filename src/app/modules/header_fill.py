import os
import openpyxl
from ..utils import log_message, log_error
from ..config import DEFAULT_OUTPUT_DIR as OUTPUT_DIR, DEFAULT_TEMPLATE_DIR as TEMPLATE_DIR

class HeaderFiller:
    def __init__(self, output_dir=None, template_dir=None):
        self.results = {
            'success': [],
            'failed': [],
            'total': 0
        }
        self.dict1 = {}  # Sheet1数据
        self.dict2 = {}  # Sheet2数据
        self.output_dir = output_dir if output_dir else OUTPUT_DIR
        self.template_dir = template_dir if template_dir else TEMPLATE_DIR
    
    def load_mapping_table(self, sheet_type):
        """加载对应表数据（按照VB代码逻辑）"""
        try:
            table_path = os.path.join(self.template_dir, '对应表.xlsx')
            if not os.path.exists(table_path):
                table_path = os.path.join(self.template_dir, '对应表.xlsm')
                if not os.path.exists(table_path):
                    raise FileNotFoundError(f"对应表文件不存在: {table_path}")
            
            wb = openpyxl.load_workbook(table_path, data_only=True)
            if sheet_type == 1:
                ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.worksheets[0]
                current_dict = self.dict1
            else:
                ws = wb['Sheet2'] if 'Sheet2' in wb.sheetnames else wb.worksheets[1]
                current_dict = self.dict2
            
            current_dict.clear()
            
            last_row = ws.max_row
            for i in range(2, last_row + 1):
                key_cell = ws.cell(row=i, column=1)
                if key_cell.value is not None:
                    key = str(key_cell.value).strip()
                    arr = [''] * 11
                    for j in range(1, 12):
                        if j <= ws.max_column:
                            cell_val = ws.cell(row=i, column=j).value
                            arr[j-1] = str(cell_val).strip() if cell_val else ''
                    
                    if key not in current_dict:
                        current_dict[key] = []
                    current_dict[key].append(arr)
            
            wb.close()
            log_message(f"成功加载Sheet{sheet_type}数据，共 {len(current_dict)} 个关键字")
            return True
        except Exception as e:
            log_error(f"加载对应表Sheet{sheet_type}失败", e)
            return False
    
    def get_base_name(self, full_name):
        """获取文件名（不含扩展名）"""
        return os.path.splitext(full_name)[0]
    
    def is_valid_excel_file(self, file_name):
        """判断是否为有效Excel文件"""
        lower_name = file_name.lower()
        return (lower_name.endswith('.xlsx') or lower_name.endswith('.xlsm')) and \
               lower_name != '对应表.xlsx' and \
               lower_name != '对应表.xlsm' and \
               not lower_name.startswith('~$')
    
    def process_single_file(self, file_path, stage):
        """处理单个文件（按照VB代码逻辑）"""
        file_name = os.path.basename(file_path)
        
        if not self.is_valid_excel_file(file_name):
            return False, None
        
        try:
            if stage == 1:
                current_dict = self.dict1
            else:
                current_dict = self.dict2
            
            base_name = self.get_base_name(file_name)
            
            # 检查关键字是否存在
            if base_name in current_dict:
                data_list = current_dict[base_name]
                if len(data_list) > 0:
                    arr_data = data_list[0]
                    
                    wb = openpyxl.load_workbook(file_path)
                    ws = wb.active
                    
                    # 根据阶段更新单元格
                    if stage == 1:
                        self.update_cells_stage1(ws, arr_data)
                        new_name = arr_data[10] + '.xlsx' if arr_data[10] else file_name
                    else:
                        self.update_cells_stage2(ws, arr_data)
                        new_name = arr_data[7] + '.xlsx' if arr_data[7] else file_name
                    
                    wb.save(file_path)
                    wb.close()
                    
                    # 返回新文件名用于重命名
                    return True, new_name
                else:
                    log_message(f"数据列表为空: {file_name}")
            else:
                log_message(f"关键字不存在: {base_name}")
            
            return False, None
        except Exception as e:
            log_error(f"处理文件失败: {file_name}", e)
            return False, None
    
    def update_cells_stage1(self, ws, arr_data):
        """阶段1单元格更新（横断面）"""
        ws['B2'] = arr_data[1]  # B列 - B2
        ws['E2'] = arr_data[2]  # C列 - E2
        ws['B3'] = arr_data[3]  # D列 - B3
        ws['E3'] = arr_data[4]  # E列 - E3
        ws['B4'] = arr_data[5]  # F列 - B4
        ws['B5'] = arr_data[6]  # G列 - B5
        ws['E4'] = arr_data[7]  # H列 - E4
        ws['E5'] = arr_data[8]  # I列 - E5
        ws['E6'] = arr_data[9]  # J列 - E6
    
    def update_cells_stage2(self, ws, arr_data):
        """阶段2单元格更新（纵断面）"""
        ws['B2'] = arr_data[1]  # B列 - B2
        ws['D4'] = arr_data[2]  # C列 - D4
        ws['B3'] = arr_data[3]  # D列 - B3
        ws['D3'] = arr_data[4]  # E列 - D3
        ws['B4'] = arr_data[5]  # F列 - B4
    
    def process_all(self, report_files=None, progress_callback=None):
        """批量处理所有文件"""
        self.results = {'success': [], 'failed': [], 'total': 0}
        
        try:
            if report_files is None:
                report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            # 阶段1：处理Sheet1数据（横断面）
            if not self.load_mapping_table(1):
                log_error("加载Sheet1失败", "无法继续")
                return self.results
            
            renamed_files = {}
            
            for file_path in report_files:
                if os.path.isfile(file_path):
                    file_name = os.path.basename(file_path)
                    success, new_name = self.process_single_file(file_path, 1)
                    if success:
                        self.results['success'].append(file_name)
                        if new_name and new_name != file_name:
                            renamed_files[file_path] = os.path.join(self.output_dir, new_name)
                    
                    if progress_callback:
                        progress_callback(f"阶段1处理: {file_name}")
            
            # 阶段1重命名
            for old_path, new_path in renamed_files.items():
                try:
                    if os.path.exists(old_path) and not os.path.exists(new_path):
                        os.rename(old_path, new_path)
                        log_message(f"重命名: {os.path.basename(old_path)} -> {os.path.basename(new_path)}")
                except Exception as e:
                    log_error(f"重命名失败: {old_path}", e)
            
            # 阶段2：处理Sheet2数据（纵断面）
            if not self.load_mapping_table(2):
                log_error("加载Sheet2失败", "无法继续")
                return self.results
            
            # 获取更新后的文件列表
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            renamed_files = {}
            
            for file_path in report_files:
                if os.path.isfile(file_path):
                    file_name = os.path.basename(file_path)
                    success, new_name = self.process_single_file(file_path, 2)
                    if success and file_name not in self.results['success']:
                        self.results['success'].append(file_name)
                    if success and new_name and new_name != file_name:
                        renamed_files[file_path] = os.path.join(self.output_dir, new_name)
                    
                    if progress_callback:
                        progress_callback(f"阶段2处理: {file_name}")
            
            # 阶段2重命名
            for old_path, new_path in renamed_files.items():
                try:
                    if os.path.exists(old_path) and not os.path.exists(new_path):
                        os.rename(old_path, new_path)
                        log_message(f"重命名: {os.path.basename(old_path)} -> {os.path.basename(new_path)}")
                except Exception as e:
                    log_error(f"重命名失败: {old_path}", e)
            
            self.results['total'] = len(self.results['success'])
            log_message(f"表头填写完成，成功处理 {self.results['total']} 个文件")
            
            if progress_callback and self.results['failed']:
                progress_callback(f"\n===== 失败详情 =====")
                for item in self.results['failed']:
                    progress_callback(f"失败: {item[0]} - {item[1]}")
            
            return self.results
        except Exception as e:
            log_error("处理过程发生错误", e)
            return self.results
