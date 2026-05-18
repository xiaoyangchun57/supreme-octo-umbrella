from PySide6.QtWidgets import QWidget, QVBoxLayout, QScrollArea
from ..widgets import CardFrame


class BaseView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(20)
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(16)
        
        self.scroll_area.setWidget(self.content_widget)
        self.layout.addWidget(self.scroll_area)
    
    def add_card(self, title: str = "") -> CardFrame:
        card = CardFrame()
        if title:
            from ..widgets import TitleLabel
            title_label = TitleLabel(title)
            card.layout.addWidget(title_label)
        self.content_layout.addWidget(card)
        return card
    
    def clear_content(self):
        while self.content_layout.count() > 0:
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()