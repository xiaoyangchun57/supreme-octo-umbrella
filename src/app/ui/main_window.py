from PySide6.QtWidgets import (QMainWindow, QFrame, QVBoxLayout, QHBoxLayout, 
                               QStackedWidget, QLabel, QStatusBar)
from PySide6.QtCore import Qt, Slot
from .styles import StyleManager
from .widgets import NavButton
from .views.settings_view import SettingsView
from .views.conversion_view import ConversionView
from .views.processing_view import ProcessingView
from .views.check_view import CheckView
from ..config import DEFAULT_DATA_DIR, DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_DIR


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("断面数据处理系统")
        self.setGeometry(100, 100, 1000, 700)
        
        self.data_dir = DEFAULT_DATA_DIR
        self.output_dir = DEFAULT_OUTPUT_DIR
        self.template_dir = DEFAULT_TEMPLATE_DIR
        
        self._setup_ui()
    
    def _setup_ui(self):
        self.central_widget = QFrame()
        self.central_widget.setObjectName("MainContent")
        self.setCentralWidget(self.central_widget)
        
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        self._setup_sidebar()
        self._setup_content_area()
        self._setup_status_bar()
        
        self.setStyleSheet(StyleManager.get_style_sheet())
    
    def _setup_sidebar(self):
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(200)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        
        logo_frame = QFrame()
        logo_frame.setFixedHeight(80)
        logo_layout = QVBoxLayout(logo_frame)
        logo_layout.setContentsMargins(20, 20, 20, 0)
        
        title_label = QLabel("断面数据处理系统")
        title_label.setStyleSheet("font-size: 16px; font-weight: 600; color: #FFFFFF;")
        logo_layout.addWidget(title_label)
        
        subtitle_label = QLabel("Section Data Processor")
        subtitle_label.setStyleSheet("font-size: 11px; color: #71717A;")
        logo_layout.addWidget(subtitle_label)
        
        sidebar_layout.addWidget(logo_frame)
        
        nav_frame = QFrame()
        nav_layout = QVBoxLayout(nav_frame)
        nav_layout.setContentsMargins(8, 0, 8, 0)
        nav_layout.setSpacing(4)
        
        self.nav_buttons = []
        
        self.settings_button = NavButton("路径设置")
        self.settings_button.clicked.connect(lambda: self._switch_view(0))
        nav_layout.addWidget(self.settings_button)
        self.nav_buttons.append(self.settings_button)
        
        self.conversion_button = NavButton("CSV转成果表")
        self.conversion_button.clicked.connect(lambda: self._switch_view(1))
        nav_layout.addWidget(self.conversion_button)
        self.nav_buttons.append(self.conversion_button)
        
        self.processing_button = NavButton("成果处理")
        self.processing_button.clicked.connect(lambda: self._switch_view(2))
        nav_layout.addWidget(self.processing_button)
        self.nav_buttons.append(self.processing_button)
        
        self.check_button = NavButton("数据检查")
        self.check_button.clicked.connect(lambda: self._switch_view(3))
        nav_layout.addWidget(self.check_button)
        self.nav_buttons.append(self.check_button)
        
        sidebar_layout.addWidget(nav_frame)
        sidebar_layout.addStretch()
        
        self.main_layout.addWidget(self.sidebar)
    
    def _setup_content_area(self):
        self.content_stack = QStackedWidget()
        
        self.settings_view = SettingsView()
        self.settings_view.apply_button.clicked.connect(self._on_settings_applied)
        self.content_stack.addWidget(self.settings_view)
        
        self.conversion_view = ConversionView(
            data_dir=self.data_dir,
            output_dir=self.output_dir,
            template_dir=self.template_dir
        )
        self.content_stack.addWidget(self.conversion_view)
        
        self.processing_view = ProcessingView(
            output_dir=self.output_dir,
            template_dir=self.template_dir
        )
        self.content_stack.addWidget(self.processing_view)
        
        self.check_view = CheckView(output_dir=self.output_dir)
        self.content_stack.addWidget(self.check_view)
        
        self.main_layout.addWidget(self.content_stack)
    
    def _setup_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("准备就绪")
    
    @Slot()
    def _switch_view(self, index):
        self.content_stack.setCurrentIndex(index)
        
        for i, button in enumerate(self.nav_buttons):
            button.setChecked(i == index)
    
    @Slot()
    def _on_settings_applied(self):
        paths = self.settings_view.get_paths()
        
        self.data_dir = paths['data_dir']
        self.output_dir = paths['output_dir']
        self.template_dir = paths['template_dir']
        
        self.conversion_view.set_paths(
            data_dir=self.data_dir,
            output_dir=self.output_dir,
            template_dir=self.template_dir
        )
        
        self.processing_view.set_paths(
            output_dir=self.output_dir,
            template_dir=self.template_dir
        )
        
        self.check_view.set_output_dir(self.output_dir)
        
        self.status_bar.showMessage("设置已更新")
    
    def closeEvent(self, event):
        event.accept()