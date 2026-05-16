import os
import sys

def get_resource_path(relative_path):
    """获取资源文件路径（支持打包后运行）"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

try:
    BASE_PATH = sys._MEIPASS
except Exception:
    BASE_PATH = ROOT_DIR

DEFAULT_DATA_DIR = get_resource_path('断面')
DEFAULT_TEMPLATE_DIR = get_resource_path('模板')
DEFAULT_OUTPUT_DIR = os.path.join(ROOT_DIR, '成果')

LOG_DIR = os.path.join(ROOT_DIR, 'logs')
BACKUP_DIR = os.path.join(ROOT_DIR, 'backup')

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

SECTION_TYPES = {
    '横断面': ['B', 'G', 'J'],
    '纵断面': ['Z'],
    '桥断面': ['Q'],
    '库容断面': ['K']
}

CSV_ENCODING = 'utf-8-sig'
OUTPUT_ENCODING = 'utf-8'

ERROR_LOG_FILE = os.path.join(LOG_DIR, 'error_log.txt')
PROCESS_LOG_FILE = os.path.join(LOG_DIR, 'process_log.txt')