from PySide6.QtWidgets import (QWidget, QPushButton, QHBoxLayout, QVBoxLayout, 
                               QTableWidget, QTableWidgetItem, QHeaderView, 
                               QAbstractItemView, QScrollArea, QFrame, QSplitter,
                               QSizePolicy, QLabel, QComboBox, QMessageBox)
from PySide6.QtCore import Slot, QThread, Signal, Qt
from PySide6.QtGui import QColor, QCursor
from .base_view import BaseView
from ..widgets import CardFrame, TitleLabel, ActionButton, LogWidget, ProgressWidget
from ...utils import get_all_xlsx_files
import os
import subprocess


class CheckWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    
    def __init__(self, output_dir):
        super().__init__()
        self.output_dir = output_dir
    
    def run(self):
        from ...modules.data_check import DataChecker
        
        checker = DataChecker(output_dir=self.output_dir)
        xlsx_files = get_all_xlsx_files(self.output_dir)
        
        def progress_callback(msg):
            self.progress.emit(msg)
        
        if not xlsx_files:
            self.progress.emit("未找到成果表文件")
            self.finished.emit({'success': [], 'failed': [], 'warnings': [], 'details': []})
            return
        
        details = []
        
        for xlsx_file in xlsx_files:
            result = checker.check_report_file(xlsx_file, progress_callback)
            
            file_name = os.path.basename(xlsx_file)
            if not result:
                for item in checker.results['failed']:
                    if item[0] == file_name:
                        details.append({
                            'file': file_name,
                            'status': 'failed',
                            'message': item[1]
                        })
            
            for item, warnings in checker.results['warnings']:
                if item == file_name:
                    if isinstance(warnings, list):
                        for warning in warnings:
                            details.append({
                                'file': file_name,
                                'status': 'warning',
                                'message': warning
                            })
                    else:
                        details.append({
                            'file': file_name,
                            'status': 'warning',
                            'message': str(warnings)
                        })
        
        report_path = checker.generate_report()
        if report_path:
            self.progress.emit(f"\n报告已生成: {report_path}")
        
        self.finished.emit({
            'success': checker.results['success'],
            'failed': checker.results['failed'],
            'warnings': checker.results['warnings'],
            'details': details
        })


class CheckView(BaseView):
    def __init__(self, output_dir=None, parent=None):
        super().__init__(parent)
        self.output_dir = output_dir
        self.details_data = []
        
        self._setup_ui()
    
    def _setup_ui(self):
        card = self.add_card("数据检查")
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        
        self.check_button = ActionButton("检查断面")
        self.check_button.clicked.connect(self._start_check)
        button_layout.addWidget(self.check_button)
        
        card.layout.addLayout(button_layout)
        
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(12)
        
        filter_label = QLabel("筛选:")
        filter_label.setStyleSheet("color: #9CA3AF;")
        
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["全部", "失败", "警告"])
        self.filter_combo.setStyleSheet("""
            QComboBox {
                background-color: #2D2D44;
                color: #FFFFFF;
                border: 1px solid #3F3F5A;
                border-radius: 8px;
                padding: 6px 12px;
                min-width: 100px;
            }
            QComboBox::drop-down {
                border-left: 1px solid #3F3F5A;
            }
        """)
        self.filter_combo.currentTextChanged.connect(self._apply_filter)
        
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.filter_combo)
        filter_layout.addStretch()
        
        self.open_button = ActionButton("打开文件")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self._open_selected_file)
        filter_layout.addWidget(self.open_button)
        
        card.layout.addLayout(filter_layout)
        
        self.progress_widget = ProgressWidget()
        card.layout.addWidget(self.progress_widget)
        
        splitter = QSplitter(Qt.Vertical)
        splitter.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.log_widget = LogWidget()
        self.log_widget.setMaximumHeight(150)
        splitter.addWidget(self.log_widget)
        
        self.error_table = QTableWidget()
        self.error_table.setColumnCount(4)
        self.error_table.setHorizontalHeaderLabels(['序号', '文件名', '状态', '问题描述'])
        self.error_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.error_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.error_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.error_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.error_table.setColumnWidth(0, 50)
        self.error_table.setColumnWidth(2, 80)
        self.error_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.error_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.error_table.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.error_table.itemSelectionChanged.connect(self._on_selection_changed)
        self.error_table.setStyleSheet("""
            QTableWidget {
                background-color: #1F1F2E;
                border: 1px solid #3F3F5A;
                border-radius: 8px;
                gridline-color: #3F3F5A;
            }
            QTableWidget::item {
                color: #FFFFFF;
                padding: 4px 8px;
            }
            QTableWidget::item:selected {
                background-color: #3B82F6;
            }
            QHeaderView::section {
                background-color: #2D2D44;
                color: #9CA3AF;
                padding: 8px;
                border: none;
                border-bottom: 1px solid #3F3F5A;
            }
            QScrollBar:vertical {
                background-color: #2D2D44;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background-color: #4B5563;
                border-radius: 6px;
            }
            QScrollBar:horizontal {
                background-color: #2D2D44;
                height: 12px;
            }
            QScrollBar::handle:horizontal {
                background-color: #4B5563;
                border-radius: 6px;
            }
        """)
        splitter.addWidget(self.error_table)
        
        card.layout.addWidget(splitter)
        
        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(20)
        
        self.success_label = QLabel("通过: 0")
        self.success_label.setStyleSheet("color: #10B981; font-weight: 500;")
        
        self.failed_label = QLabel("失败: 0")
        self.failed_label.setStyleSheet("color: #EF4444; font-weight: 500;")
        
        self.warning_label = QLabel("警告: 0")
        self.warning_label.setStyleSheet("color: #F59E0B; font-weight: 500;")
        
        stats_layout.addWidget(self.success_label)
        stats_layout.addWidget(self.failed_label)
        stats_layout.addWidget(self.warning_label)
        
        hint_label = QLabel("提示: 双击问题行可直接打开文件定位修改")
        hint_label.setStyleSheet("color: #6B7280; font-size: 12px;")
        stats_layout.addWidget(hint_label)
        
        stats_layout.addStretch()
        
        card.layout.addLayout(stats_layout)
        
        self.all_details = []
    
    @Slot()
    def _start_check(self):
        if not self.output_dir:
            self.log_widget.add_log("请先设置输出目录", "error")
            return
        
        self.check_button.setEnabled(False)
        self.progress_widget.reset()
        self.log_widget.clear_log()
        self.error_table.setRowCount(0)
        self.details_data = []
        self.all_details = []
        self.open_button.setEnabled(False)
        
        self.worker = CheckWorker(self.output_dir)
        self.worker.progress.connect(self._handle_progress)
        self.worker.finished.connect(self._handle_check_finished)
        self.worker.start()
    
    @Slot(str)
    def _handle_progress(self, message):
        self.log_widget.add_log(message)
        self.progress_widget.set_message(message)
    
    @Slot(dict)
    def _handle_check_finished(self, results):
        self.log_widget.add_log(f"\n检查完成！", "success")
        self.log_widget.add_log(f"通过: {len(results['success'])}")
        self.log_widget.add_log(f"失败: {len(results['failed'])}")
        self.log_widget.add_log(f"警告: {len(results['warnings'])}")
        
        self.success_label.setText(f"通过: {len(results['success'])}")
        self.failed_label.setText(f"失败: {len(results['failed'])}")
        self.warning_label.setText(f"警告: {len(results['warnings'])}")
        
        self.all_details = results['details']
        self._populate_error_table(self.all_details)
        
        if results['failed']:
            self.log_widget.add_log("\n===== 失败详情 =====", "warning")
            for item in results['failed']:
                self.log_widget.add_log(f"失败: {item[0]} - {item[1]}", "error")
        
        if results['warnings']:
            self.log_widget.add_log("\n===== 警告详情 =====", "warning")
            for item, warnings in results['warnings']:
                self.log_widget.add_log(f"警告: {item}")
                if isinstance(warnings, list):
                    for warning in warnings:
                        self.log_widget.add_log(f"  - {warning}", "warning")
        
        self.check_button.setEnabled(True)
        self.progress_widget.set_progress(100, "检查完成")
    
    def _populate_error_table(self, details):
        self.error_table.setRowCount(len(details))
        
        for row, detail in enumerate(details):
            index_item = QTableWidgetItem(str(row + 1))
            index_item.setFlags(index_item.flags() & ~Qt.ItemIsEditable)
            index_item.setTextAlignment(Qt.AlignCenter)
            
            file_item = QTableWidgetItem(detail['file'])
            file_item.setFlags(file_item.flags() & ~Qt.ItemIsEditable)
            
            status_item = QTableWidgetItem()
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            
            message_item = QTableWidgetItem(detail['message'])
            message_item.setFlags(message_item.flags() & ~Qt.ItemIsEditable)
            
            if detail['status'] == 'failed':
                status_item.setText("失败")
                status_item.setBackground(QColor("#EF4444"))
                status_item.setForeground(QColor("#FFFFFF"))
                message_item.setBackground(QColor("#374151"))
            elif detail['status'] == 'warning':
                status_item.setText("警告")
                status_item.setBackground(QColor("#F59E0B"))
                status_item.setForeground(QColor("#FFFFFF"))
                message_item.setBackground(QColor("#451A03"))
            
            self.error_table.setItem(row, 0, index_item)
            self.error_table.setItem(row, 1, file_item)
            self.error_table.setItem(row, 2, status_item)
            self.error_table.setItem(row, 3, message_item)
        
        self.error_table.resizeRowsToContents()
    
    @Slot(str)
    def _apply_filter(self, filter_text):
        if not self.all_details:
            return
        
        if filter_text == "全部":
            self._populate_error_table(self.all_details)
        elif filter_text == "失败":
            filtered = [d for d in self.all_details if d['status'] == 'failed']
            self._populate_error_table(filtered)
        elif filter_text == "警告":
            filtered = [d for d in self.all_details if d['status'] == 'warning']
            self._populate_error_table(filtered)
    
    @Slot(QTableWidgetItem)
    def _on_item_double_clicked(self, item):
        row = item.row()
        if row < len(self.all_details):
            detail = self.all_details[row]
            self._open_file(detail['file'])
    
    @Slot()
    def _on_selection_changed(self):
        selected_rows = self.error_table.selectedRows()
        self.open_button.setEnabled(len(selected_rows) > 0)
    
    @Slot()
    def _open_selected_file(self):
        selected_rows = self.error_table.selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
            if row < len(self.all_details):
                detail = self.all_details[row]
                self._open_file(detail['file'])
    
    def _open_file(self, filename):
        if not self.output_dir:
            QMessageBox.warning(self, "警告", "请先设置输出目录")
            return
        
        file_path = os.path.join(self.output_dir, filename)
        
        if os.path.exists(file_path):
            try:
                subprocess.Popen(['start', '', file_path], shell=True)
            except Exception as e:
                QMessageBox.warning(self, "打开失败", f"无法打开文件: {str(e)}")
        else:
            QMessageBox.warning(self, "文件不存在", f"文件不存在: {file_path}")
    
    def set_output_dir(self, output_dir):
        self.output_dir = output_dir