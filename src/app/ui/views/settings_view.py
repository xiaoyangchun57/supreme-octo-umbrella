from PySide6.QtWidgets import (QWidget, QLabel, QLineEdit, QPushButton, 
                               QHBoxLayout, QVBoxLayout, QGridLayout, 
                               QTextEdit, QFileDialog)
from PySide6.QtCore import Slot
from .base_view import BaseView
from ..widgets import CardFrame, TitleLabel, SubtitleLabel, ActionButton, SecondaryButton
from ...config import DEFAULT_DATA_DIR, DEFAULT_TEMPLATE_DIR, DEFAULT_OUTPUT_DIR
import os


class SettingsView(BaseView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data_dir = DEFAULT_DATA_DIR
        self.template_dir = DEFAULT_TEMPLATE_DIR
        self.output_dir = DEFAULT_OUTPUT_DIR
        self.root_dir = os.path.dirname(DEFAULT_DATA_DIR)
        
        self._setup_ui()
    
    def _setup_ui(self):
        card = self.add_card("路径设置")
        
        grid_layout = QGridLayout()
        grid_layout.setSpacing(12)
        
        row = 0
        
        grid_layout.addWidget(QLabel("根目录:"), row, 0)
        self.root_dir_entry = QLineEdit()
        self.root_dir_entry.setReadOnly(True)
        self.root_dir_entry.setText(self.root_dir)
        grid_layout.addWidget(self.root_dir_entry, row, 1)
        self.browse_root_button = QPushButton("浏览")
        self.browse_root_button.setObjectName("BrowseButton")
        self.browse_root_button.clicked.connect(self._browse_root_dir)
        grid_layout.addWidget(self.browse_root_button, row, 2)
        row += 1
        
        grid_layout.addWidget(QLabel("断面数据目录:"), row, 0)
        self.data_dir_entry = QLineEdit()
        self.data_dir_entry.setReadOnly(True)
        self.data_dir_entry.setText(self.data_dir)
        grid_layout.addWidget(self.data_dir_entry, row, 1)
        row += 1
        
        grid_layout.addWidget(QLabel("模板文件目录:"), row, 0)
        self.template_dir_entry = QLineEdit()
        self.template_dir_entry.setReadOnly(True)
        self.template_dir_entry.setText(self.template_dir)
        grid_layout.addWidget(self.template_dir_entry, row, 1)
        row += 1
        
        grid_layout.addWidget(QLabel("成果输出目录:"), row, 0)
        self.output_dir_entry = QLineEdit()
        self.output_dir_entry.setReadOnly(True)
        self.output_dir_entry.setText(self.output_dir)
        grid_layout.addWidget(self.output_dir_entry, row, 1)
        row += 1
        
        card.layout.addLayout(grid_layout)
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        
        self.apply_button = ActionButton("应用设置")
        self.apply_button.clicked.connect(self._apply_settings)
        button_layout.addWidget(self.apply_button)
        
        self.verify_button = SecondaryButton("验证路径")
        self.verify_button.clicked.connect(self._verify_paths)
        button_layout.addWidget(self.verify_button)
        
        card.layout.addLayout(button_layout)
        
        self.status_label = SubtitleLabel("")
        card.layout.addWidget(self.status_label)
        
        self.verify_result = QTextEdit()
        self.verify_result.setReadOnly(True)
        self.verify_result.setMaximumHeight(150)
        card.layout.addWidget(self.verify_result)
    
    @Slot()
    def _browse_root_dir(self):
        path = QFileDialog.getExistingDirectory(
            self, 
            "选择根目录（包含断面和模板文件夹）", 
            self.root_dir
        )
        if path:
            self.root_dir = path
            self.root_dir_entry.setText(path)
            
            self.data_dir = os.path.join(path, '断面')
            self.template_dir = os.path.join(path, '模板')
            self.output_dir = os.path.join(path, '成果')
            
            self.data_dir_entry.setText(self.data_dir)
            self.template_dir_entry.setText(self.template_dir)
            self.output_dir_entry.setText(self.output_dir)
    
    @Slot()
    def _apply_settings(self):
        self.root_dir = self.root_dir_entry.text()
        self.data_dir = os.path.join(self.root_dir, '断面')
        self.template_dir = os.path.join(self.root_dir, '模板')
        self.output_dir = os.path.join(self.root_dir, '成果')
        
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.template_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
        
        settings = f"""# 断面数据处理系统配置文件
ROOT_DIR={self.root_dir}
DATA_DIR={self.data_dir}
OUTPUT_DIR={self.output_dir}
TEMPLATE_DIR={self.template_dir}
"""
        config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'settings.ini')
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(settings)
        
        self.status_label.setText("设置已应用！")
        self.status_label.setStyleSheet("color: #10B981;")
    
    @Slot()
    def _verify_paths(self):
        self.verify_result.clear()
        
        results = []
        
        data_path = self.data_dir_entry.text()
        if os.path.exists(data_path):
            csv_count = len([f for f in os.listdir(data_path) if f.endswith('.csv')])
            results.append(f"✓ 数据源目录: {data_path} (包含 {csv_count} 个CSV文件)")
        else:
            results.append(f"✗ 数据源目录不存在: {data_path}")
        
        output_path = self.output_dir_entry.text()
        if os.path.exists(output_path):
            xlsx_count = len([f for f in os.listdir(output_path) if f.endswith('.xlsx')])
            results.append(f"✓ 输出目录: {output_path} (包含 {xlsx_count} 个Excel文件)")
        else:
            results.append(f"⚠ 输出目录不存在，将自动创建: {output_path}")
        
        template_path = self.template_dir_entry.text()
        if os.path.exists(template_path):
            templates = ['横断面成果表模板.xlsx', '纵断面模板.xlsx', '成图模板.xlsx', '对应表.xlsx']
            found = [t for t in templates if os.path.exists(os.path.join(template_path, t))]
            missing = [t for t in templates if not os.path.exists(os.path.join(template_path, t))]
            results.append(f"✓ 模板目录: {template_path}")
            if found:
                results.append(f"  已找到模板: {', '.join(found)}")
            if missing:
                results.append(f"  缺失模板: {', '.join(missing)}")
        else:
            results.append(f"✗ 模板目录不存在: {template_path}")
        
        corr_path = self.template_dir_entry.text()
        if corr_path and os.path.exists(os.path.join(corr_path, '对应表.xlsx')):
            results.append(f"✓ 对应表文件存在")
        elif corr_path:
            results.append(f"⚠ 对应表文件未找到: {os.path.join(corr_path, '对应表.xlsx')}")
        
        self.verify_result.setPlainText("\n".join(results))
    
    def get_paths(self):
        return {
            'root_dir': self.root_dir,
            'data_dir': self.data_dir,
            'template_dir': self.template_dir,
            'output_dir': self.output_dir
        }