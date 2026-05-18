from PySide6.QtWidgets import (QWidget, QPushButton, QLabel, QLineEdit, 
                               QProgressBar, QTextEdit, QFrame, QVBoxLayout, 
                               QHBoxLayout, QGridLayout, QSizePolicy)
from PySide6.QtCore import Qt, Signal, Slot, QSize
from .styles import StyleManager


class NavButton(QPushButton):
    def __init__(self, text: str, icon_name: str = "", parent=None):
        super().__init__(text, parent)
        self.setObjectName("NavButton")
        self.setCheckable(True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    
    def set_active(self, active: bool):
        self.setChecked(active)


class ActionButton(QPushButton):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("ActionButton")
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)


class SecondaryButton(QPushButton):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("SecondaryButton")
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)


class CardFrame(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(16, 16, 16, 16)
        self.layout.setSpacing(12)


class TitleLabel(QLabel):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("TitleLabel")


class SubtitleLabel(QLabel):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("SubtitleLabel")


class StatsLabel(QLabel):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("StatsLabel")


class PathLineEdit(QLineEdit):
    browse_clicked = Signal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
    
    def mouseDoubleClickEvent(self, event):
        self.browse_clicked.emit()
        super().mouseDoubleClickEvent(event)


class ProgressWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(8)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("SubtitleLabel")
        
        self.layout.addWidget(self.status_label)
        self.layout.addWidget(self.progress_bar)
    
    def set_progress(self, percentage: int, message: str = ""):
        self.progress_bar.setValue(percentage)
        if message:
            self.status_label.setText(message)
    
    def set_message(self, message: str):
        self.status_label.setText(message)
    
    def reset(self):
        self.progress_bar.setValue(0)
        self.status_label.setText("准备就绪")


class LogWidget(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setLineWrapMode(QTextEdit.NoWrap)
        
        self.error_color = StyleManager.ERROR_COLOR
        self.warning_color = StyleManager.WARNING_COLOR
        self.success_color = StyleManager.SUCCESS_COLOR
        self.info_color = StyleManager.TEXT_PRIMARY
    
    def add_log(self, message: str, level: str = "info"):
        color_map = {
            "error": self.error_color,
            "warning": self.warning_color,
            "success": self.success_color,
            "info": self.info_color
        }
        
        color = color_map.get(level, self.info_color)
        timestamp = __import__('datetime').datetime.now().strftime("%H:%M:%S")
        log_line = f'<span style="color:{color}">[{timestamp}] {message}</span>'
        
        self.append(log_line)
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())
    
    def clear_log(self):
        self.clear()


class StatsCard(QWidget):
    def __init__(self, title: str, value: str, color: str = StyleManager.PRIMARY_COLOR, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(12, 12, 12, 12)
        self.layout.setSpacing(4)
        
        self.title_label = QLabel(title)
        self.title_label.setObjectName("StatsLabel")
        
        self.value_label = QLabel(value)
        self.value_label.setStyleSheet(f"font-size: 24px; font-weight: 600; color: {color};")
        
        self.layout.addWidget(self.title_label)
        self.layout.addWidget(self.value_label)
    
    def set_value(self, value: str):
        self.value_label.setText(value)


class SectionStatsWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(16)
        
        self.cross_section_card = StatsCard("横断面", "0", StyleManager.PRIMARY_COLOR)
        self.longitudinal_card = StatsCard("纵断面", "0", StyleManager.SUCCESS_COLOR)
        
        self.layout.addWidget(self.cross_section_card)
        self.layout.addWidget(self.longitudinal_card)
    
    def update_stats(self, cross_count: int, longitudinal_count: int):
        self.cross_section_card.set_value(str(cross_count))
        self.longitudinal_card.set_value(str(longitudinal_count))


class ToggleSwitch(QWidget):
    toggled = Signal(bool)
    
    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(8)
        
        self.label = QLabel(text)
        self.label.setObjectName("SubtitleLabel")
        
        self.switch = QPushButton()
        self.switch.setCheckable(True)
        self.switch.setFixedSize(40, 24)
        self.switch.setStyleSheet("""
            QPushButton {
                background-color: #3F3F5A;
                border-radius: 12px;
            }
            QPushButton::checked {
                background-color: #3B82F6;
            }
            QPushButton::indicator {
                width: 20px;
                height: 20px;
                background-color: white;
                border-radius: 10px;
                margin: 2px;
            }
            QPushButton::checked::indicator {
                margin-left: 18px;
            }
        """)
        
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.switch)
        
        self.switch.toggled.connect(self.toggled.emit)
    
    def is_checked(self) -> bool:
        return self.switch.isChecked()
    
    def set_checked(self, checked: bool):
        self.switch.setChecked(checked)