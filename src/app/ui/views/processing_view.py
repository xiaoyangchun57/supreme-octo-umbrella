from PySide6.QtWidgets import (QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout)
from PySide6.QtCore import Slot, QThread, Signal
from .base_view import BaseView
from ..widgets import CardFrame, TitleLabel, ActionButton, SecondaryButton, LogWidget, ProgressWidget
import os


class ProcessingWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    worker_type = ""
    
    def __init__(self, worker_type, output_dir, template_dir):
        super().__init__()
        self.worker_type = worker_type
        self.output_dir = output_dir
        self.template_dir = template_dir
    
    def run(self):
        results = {'success': [], 'failed': []}
        
        if self.worker_type == 'convert_85':
            from ...modules.report_to_85 import ReportTo85Converter
            converter = ReportTo85Converter(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx') and not f.endswith('_成图.xlsx')]
            
            def progress_callback(msg):
                self.progress.emit(msg)
            
            results = converter.process_all(report_files, progress_callback)
        
        elif self.worker_type == 'fill_header':
            from ...modules.header_fill import HeaderFiller
            filler = HeaderFiller(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.progress.emit(msg)
            
            results = filler.process_all(report_files, progress_callback)
        
        elif self.worker_type == 'auto_plot':
            from ...modules.auto_plot import AutoPlotter
            plotter = AutoPlotter(output_dir=self.output_dir, template_dir=self.template_dir)
            report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
            
            def progress_callback(msg):
                self.progress.emit(msg)
            
            results = plotter.process_all(report_files, progress_callback)
        
        elif self.worker_type == 'integrate_merge':
            from ...modules.folder_integration import FolderIntegrator
            integrator = FolderIntegrator(output_dir=self.output_dir, template_dir=self.template_dir)
            
            def progress_callback(msg):
                self.progress.emit(msg)
            
            results = integrator.process_all(self.output_dir, progress_callback)
        
        self.finished.emit(results)


class StorageCalculationWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    
    def __init__(self, output_dir, template_dir):
        super().__init__()
        self.output_dir = output_dir
        self.template_dir = template_dir
        self.calculator = None
    
    def stop(self):
        if self.calculator:
            self.calculator.stop()
    
    def run(self):
        from ...modules.storage_calculation import StorageCalculator
        self.calculator = StorageCalculator(output_dir=self.output_dir, template_dir=self.template_dir)
        
        report_files = [os.path.join(self.output_dir, f) for f in os.listdir(self.output_dir) if f.startswith('K') and f.endswith('.xlsx')]
        
        def progress_callback(msg):
            self.progress.emit(msg)
        
        results = self.calculator.process_all(report_files, progress_callback)
        self.finished.emit(results)


class ProcessingView(BaseView):
    def __init__(self, output_dir=None, template_dir=None, parent=None):
        super().__init__(parent)
        self.output_dir = output_dir
        self.template_dir = template_dir
        self.current_worker = None
        
        self._setup_ui()
    
    def _setup_ui(self):
        card = self.add_card("成果处理")
        
        grid_layout = QGridLayout()
        grid_layout.setSpacing(12)
        
        self.convert_85_button = ActionButton("转85高程")
        self.convert_85_button.clicked.connect(lambda: self._start_processing('convert_85'))
        grid_layout.addWidget(self.convert_85_button, 0, 0)
        
        self.fill_header_button = ActionButton("填写表头")
        self.fill_header_button.clicked.connect(lambda: self._start_processing('fill_header'))
        grid_layout.addWidget(self.fill_header_button, 0, 1)
        
        self.auto_plot_button = ActionButton("自动成图")
        self.auto_plot_button.clicked.connect(lambda: self._start_processing('auto_plot'))
        grid_layout.addWidget(self.auto_plot_button, 1, 0)
        
        self.integrate_button = ActionButton("整合合并")
        self.integrate_button.clicked.connect(lambda: self._start_processing('integrate_merge'))
        grid_layout.addWidget(self.integrate_button, 1, 1)
        
        self.storage_button = ActionButton("库容计算")
        self.storage_button.clicked.connect(self._start_storage_calculation)
        grid_layout.addWidget(self.storage_button, 2, 0)
        
        self.stop_button = SecondaryButton("停止计算")
        self.stop_button.clicked.connect(self._stop_processing)
        self.stop_button.setEnabled(False)
        grid_layout.addWidget(self.stop_button, 2, 1)
        
        card.layout.addLayout(grid_layout)
        
        self.progress_widget = ProgressWidget()
        card.layout.addWidget(self.progress_widget)
        
        self.log_widget = LogWidget()
        self.log_widget.setMaximumHeight(250)
        card.layout.addWidget(self.log_widget)
    
    @Slot()
    def _start_processing(self, worker_type):
        if not self.output_dir:
            self.log_widget.add_log("请先设置输出目录", "error")
            return
        
        self._disable_buttons()
        self.stop_button.setEnabled(True)
        self.progress_widget.reset()
        self.log_widget.clear_log()
        
        self.current_worker = ProcessingWorker(worker_type, self.output_dir, self.template_dir)
        self.current_worker.progress.connect(self._handle_progress)
        self.current_worker.finished.connect(self._handle_processing_finished)
        self.current_worker.start()
    
    @Slot()
    def _start_storage_calculation(self):
        if not self.output_dir:
            self.log_widget.add_log("请先设置输出目录", "error")
            return
        
        self._disable_buttons()
        self.stop_button.setEnabled(True)
        self.progress_widget.reset()
        self.log_widget.clear_log()
        
        self.current_worker = StorageCalculationWorker(self.output_dir, self.template_dir)
        self.current_worker.progress.connect(self._handle_progress)
        self.current_worker.finished.connect(self._handle_processing_finished)
        self.current_worker.start()
    
    @Slot()
    def _stop_processing(self):
        if self.current_worker:
            if hasattr(self.current_worker, 'stop'):
                self.current_worker.stop()
            self.log_widget.add_log("正在停止...", "warning")
    
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
        
        match = re.search(r'已转换85高程: (.+)', message)
        if match:
            self.progress_widget.set_progress(50)
            return
        
        match = re.search(r'已转换: (.+)', message)
        if match:
            self.progress_widget.set_progress(50)
            return
        
        match = re.search(r'处理: (.+)', message)
        if match:
            self.progress_widget.set_progress(50)
            return
    
    @Slot(dict)
    def _handle_processing_finished(self, results):
        if 'success' in results:
            self.log_widget.add_log(f"\n处理完成！", "success")
            self.log_widget.add_log(f"成功: {len(results['success'])}")
            self.log_widget.add_log(f"失败: {len(results['failed'])}")
            
            if results['failed']:
                self.log_widget.add_log("\n===== 失败详情 =====", "warning")
                for item in results['failed']:
                    if isinstance(item, tuple):
                        self.log_widget.add_log(f"失败: {item[0]} - {item[1]}", "error")
        
        self._enable_buttons()
        self.stop_button.setEnabled(False)
        self.progress_widget.set_progress(100, "处理完成")
    
    def _disable_buttons(self):
        self.convert_85_button.setEnabled(False)
        self.fill_header_button.setEnabled(False)
        self.auto_plot_button.setEnabled(False)
        self.integrate_button.setEnabled(False)
        self.storage_button.setEnabled(False)
    
    def _enable_buttons(self):
        self.convert_85_button.setEnabled(True)
        self.fill_header_button.setEnabled(True)
        self.auto_plot_button.setEnabled(True)
        self.integrate_button.setEnabled(True)
        self.storage_button.setEnabled(True)
    
    def set_paths(self, output_dir, template_dir):
        self.output_dir = output_dir
        self.template_dir = template_dir