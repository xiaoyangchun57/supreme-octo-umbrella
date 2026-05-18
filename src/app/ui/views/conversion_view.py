from PySide6.QtWidgets import (QWidget, QLabel, QPushButton, QHBoxLayout, 
                               QVBoxLayout, QTextEdit, QFileDialog)
from PySide6.QtCore import Slot, QThread, Signal
from .base_view import BaseView
from ..widgets import CardFrame, TitleLabel, SubtitleLabel, ActionButton, SecondaryButton, SectionStatsWidget, LogWidget, ProgressWidget
from ...utils import get_all_csv_files
import os


class ConversionWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    
    def __init__(self, csv_files, output_dir, template_dir):
        super().__init__()
        self.csv_files = csv_files
        self.output_dir = output_dir
        self.template_dir = template_dir
    
    def run(self):
        from ...modules.csv_to_report import CsvToReportConverter
        
        converter = CsvToReportConverter(output_dir=self.output_dir, template_dir=self.template_dir)
        
        def progress_callback(msg):
            self.progress.emit(msg)
        
        results = converter.process_all(self.csv_files, progress_callback)
        self.finished.emit(results)


class ConversionView(BaseView):
    def __init__(self, data_dir=None, output_dir=None, template_dir=None, parent=None):
        super().__init__(parent)
        self.data_dir = data_dir
        self.output_dir = output_dir
        self.template_dir = template_dir
        self.csv_files = []
        
        self._setup_ui()
    
    def _setup_ui(self):
        card = self.add_card("CSV转成果表")
        
        path_row = QHBoxLayout()
        self.path_label = SubtitleLabel("源目录: ")
        path_row.addWidget(self.path_label)
        
        self.browse_button = SecondaryButton("浏览")
        self.browse_button.clicked.connect(self._browse_data_dir)
        path_row.addWidget(self.browse_button)
        
        card.layout.addLayout(path_row)
        
        self.stats_widget = SectionStatsWidget()
        card.layout.addWidget(self.stats_widget)
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        
        self.load_button = SecondaryButton("加载CSV文件")
        self.load_button.clicked.connect(self._load_csv_files)
        button_layout.addWidget(self.load_button)
        
        self.convert_button = ActionButton("开始转换")
        self.convert_button.clicked.connect(self._start_conversion)
        button_layout.addWidget(self.convert_button)
        
        card.layout.addLayout(button_layout)
        
        self.progress_widget = ProgressWidget()
        card.layout.addWidget(self.progress_widget)
        
        self.log_widget = LogWidget()
        self.log_widget.setMaximumHeight(200)
        card.layout.addWidget(self.log_widget)
    
    @Slot()
    def _browse_data_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "选择CSV文件目录", self.data_dir or "")
        if folder:
            self.data_dir = folder
            self.path_label.setText(f"源目录: {folder}")
    
    @Slot()
    def _load_csv_files(self):
        if not self.data_dir or not os.path.exists(self.data_dir):
            self.log_widget.add_log("请先选择有效的数据目录", "error")
            return
        
        self.csv_files = get_all_csv_files(self.data_dir)
        
        zbc_count = 0
        zbm_count = 0
        
        for f in self.csv_files:
            filename = os.path.basename(f).lower()
            if filename.startswith('z'):
                zbm_count += 1
            else:
                zbc_count += 1
        
        self.stats_widget.update_stats(zbc_count, zbm_count)
        
        self.log_widget.add_log(f"检测到 {len(self.csv_files)} 个CSV文件")
        self.log_widget.add_log(f"  - 横断面: {zbc_count} 个")
        self.log_widget.add_log(f"  - 纵断面: {zbm_count} 个")
    
    @Slot()
    def _start_conversion(self):
        if not self.csv_files:
            self.log_widget.add_log("请先加载CSV文件", "warning")
            return
        
        if not self.output_dir:
            self.log_widget.add_log("请先设置输出目录", "error")
            return
        
        self.convert_button.setEnabled(False)
        self.progress_widget.reset()
        self.log_widget.clear_log()
        
        self.worker = ConversionWorker(
            csv_files=self.csv_files,
            output_dir=self.output_dir,
            template_dir=self.template_dir
        )
        
        self.worker.progress.connect(self._handle_progress)
        self.worker.finished.connect(self._handle_conversion_finished)
        
        self.worker.start()
    
    @Slot(str)
    def _handle_progress(self, message):
        self.log_widget.add_log(message)
        self.progress_widget.set_message(message)
        
        re = __import__('re')
        
        match = re.search(r'已完成 (\d+)/(\d+)', message)
        if match:
            completed = int(match.group(1))
            total = int(match.group(2))
            percentage = int((completed / total) * 100)
            self.progress_widget.set_progress(percentage)
            return
        
        match = re.search(r'已转换: (.+)', message)
        if match:
            self.progress_widget.set_progress(50)
            return
    
    @Slot(dict)
    def _handle_conversion_finished(self, results):
        self.log_widget.add_log(f"\n转换完成！", "success")
        self.log_widget.add_log(f"成功: {len(results['success'])}")
        self.log_widget.add_log(f"失败: {len(results['failed'])}")
        
        if results['failed']:
            self.log_widget.add_log("\n===== 失败详情 =====", "warning")
            for item in results['failed']:
                self.log_widget.add_log(f"失败: {item[0]} - {item[1]}", "error")
        
        self.convert_button.setEnabled(True)
        self.progress_widget.set_progress(100, "转换完成")
    
    def set_paths(self, data_dir, output_dir, template_dir):
        self.data_dir = data_dir
        self.output_dir = output_dir
        self.template_dir = template_dir
        self.path_label.setText(f"源目录: {data_dir}")