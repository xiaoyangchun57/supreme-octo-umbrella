import os
import shutil
import csv
from datetime import datetime
from .config import CSV_ENCODING, OUTPUT_ENCODING, LOG_DIR, BACKUP_DIR

def read_csv_file(file_path):
    """读取CSV文件，返回数据列表"""
    try:
        with open(file_path, 'r', encoding=CSV_ENCODING) as f:
            reader = csv.reader(f)
            return list(reader)
    except Exception as e:
        try:
            with open(file_path, 'r', encoding='gbk') as f:
                reader = csv.reader(f)
                return list(reader)
        except:
            return []

def write_csv_file(file_path, data):
    """写入CSV文件"""
    with open(file_path, 'w', encoding=OUTPUT_ENCODING, newline='') as f:
        writer = csv.writer(f)
        writer.writerows(data)

def get_section_type(file_name):
    """根据文件名判断断面类型"""
    prefix = file_name[0].upper()
    if prefix in ['B', 'G', 'J']:
        return '横断面'
    elif prefix == 'Z':
        return '纵断面'
    elif prefix == 'Q':
        return '桥断面'
    elif prefix == 'K':
        return '库容断面'
    return None

def get_all_csv_files(data_dir):
    """获取目录下所有CSV文件"""
    csv_files = []
    for root, dirs, files in os.walk(data_dir):
        for file in files:
            if file.lower().endswith('.csv'):
                csv_files.append(os.path.join(root, file))
    return csv_files

def backup_file(file_path):
    """备份文件"""
    if os.path.exists(file_path):
        file_name = os.path.basename(file_path)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f"{os.path.splitext(file_name)[0]}_{timestamp}{os.path.splitext(file_name)[1]}"
        backup_path = os.path.join(BACKUP_DIR, backup_name)
        shutil.copy2(file_path, backup_path)
        return backup_path
    return None

def log_message(message, log_type='info'):
    """记录日志"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_line = f"[{timestamp}] [{log_type.upper()}] {message}\n"
    
    log_file = os.path.join(LOG_DIR, 'app.log')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(log_line)

def log_error(message, exception=None):
    """记录错误日志"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if exception:
        log_line = f"[{timestamp}] [ERROR] {message}\nException: {str(exception)}\n"
    else:
        log_line = f"[{timestamp}] [ERROR] {message}\n"
    
    log_file = os.path.join(LOG_DIR, 'error.log')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(log_line)