class StyleManager:
    PRIMARY_COLOR = "#3B82F6"
    PRIMARY_HOVER = "#2563EB"
    SUCCESS_COLOR = "#10B981"
    WARNING_COLOR = "#F59E0B"
    ERROR_COLOR = "#EF4444"
    INFO_COLOR = "#06B6D4"
    
    BG_DARK = "#1E1E2E"
    BG_CARD = "#252536"
    BG_HOVER = "#2D2D44"
    BG_INPUT = "#2A2A3E"
    
    TEXT_PRIMARY = "#FFFFFF"
    TEXT_SECONDARY = "#A1A1AA"
    TEXT_MUTED = "#71717A"
    
    BORDER_COLOR = "#3F3F5A"
    BORDER_RADIUS = "8px"
    
    @staticmethod
    def get_style_sheet() -> str:
        return f"""
            QMainWindow {{
                background-color: {StyleManager.BG_DARK};
            }}
            
            QWidget {{
                color: {StyleManager.TEXT_PRIMARY};
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }}
            
            QFrame#Sidebar {{
                background-color: {StyleManager.BG_CARD};
                border-right: 1px solid {StyleManager.BORDER_COLOR};
            }}
            
            QFrame#MainContent {{
                background-color: {StyleManager.BG_DARK};
            }}
            
            QFrame#Card {{
                background-color: {StyleManager.BG_CARD};
                border-radius: {StyleManager.BORDER_RADIUS};
                border: 1px solid {StyleManager.BORDER_COLOR};
            }}
            
            QPushButton#NavButton {{
                background-color: transparent;
                color: {StyleManager.TEXT_SECONDARY};
                border: none;
                padding: 12px 16px;
                text-align: left;
                font-size: 14px;
                border-radius: {StyleManager.BORDER_RADIUS};
                min-height: 48px;
            }}
            
            QPushButton#NavButton:hover {{
                background-color: {StyleManager.BG_HOVER};
                color: {StyleManager.TEXT_PRIMARY};
            }}
            
            QPushButton#NavButton:checked {{
                background-color: {StyleManager.PRIMARY_COLOR};
                color: {StyleManager.TEXT_PRIMARY};
            }}
            
            QPushButton#BrowseButton {{
                background-color: {StyleManager.BG_HOVER};
                color: {StyleManager.TEXT_PRIMARY};
                border: 1px solid {StyleManager.BORDER_COLOR};
                padding: 8px 16px;
                border-radius: {StyleManager.BORDER_RADIUS};
                font-size: 13px;
            }}
            
            QPushButton#BrowseButton:hover {{
                background-color: {StyleManager.BORDER_COLOR};
            }}
            
            QPushButton#ActionButton {{
                background-color: {StyleManager.PRIMARY_COLOR};
                color: {StyleManager.TEXT_PRIMARY};
                border: none;
                padding: 10px 24px;
                border-radius: {StyleManager.BORDER_RADIUS};
                font-size: 14px;
                font-weight: 500;
            }}
            
            QPushButton#ActionButton:hover {{
                background-color: {StyleManager.PRIMARY_HOVER};
            }}
            
            QPushButton#ActionButton:pressed {{
                background-color: {StyleManager.PRIMARY_COLOR};
            }}
            
            QPushButton#ActionButton:disabled {{
                background-color: {StyleManager.BG_HOVER};
                color: {StyleManager.TEXT_MUTED};
            }}
            
            QPushButton#SecondaryButton {{
                background-color: {StyleManager.BG_HOVER};
                color: {StyleManager.TEXT_PRIMARY};
                border: 1px solid {StyleManager.BORDER_COLOR};
                padding: 10px 24px;
                border-radius: {StyleManager.BORDER_RADIUS};
                font-size: 14px;
            }}
            
            QPushButton#SecondaryButton:hover {{
                background-color: {StyleManager.BORDER_COLOR};
            }}
            
            QLineEdit {{
                background-color: {StyleManager.BG_INPUT};
                border: 1px solid {StyleManager.BORDER_COLOR};
                border-radius: {StyleManager.BORDER_RADIUS};
                padding: 10px 12px;
                color: {StyleManager.TEXT_PRIMARY};
                font-size: 14px;
            }}
            
            QLineEdit:focus {{
                border-color: {StyleManager.PRIMARY_COLOR};
                outline: none;
            }}
            
            QLineEdit:readonly {{
                background-color: {StyleManager.BG_CARD};
                color: {StyleManager.TEXT_MUTED};
            }}
            
            QLabel#TitleLabel {{
                font-size: 18px;
                font-weight: 600;
                color: {StyleManager.TEXT_PRIMARY};
            }}
            
            QLabel#SubtitleLabel {{
                font-size: 14px;
                color: {StyleManager.TEXT_SECONDARY};
            }}
            
            QLabel#StatsLabel {{
                font-size: 12px;
                color: {StyleManager.TEXT_MUTED};
            }}
            
            QProgressBar {{
                background-color: {StyleManager.BG_INPUT};
                border: none;
                border-radius: {StyleManager.BORDER_RADIUS};
                height: 8px;
                text-align: center;
            }}
            
            QProgressBar::chunk {{
                background-color: {StyleManager.PRIMARY_COLOR};
                border-radius: {StyleManager.BORDER_RADIUS};
            }}
            
            QTextEdit {{
                background-color: {StyleManager.BG_INPUT};
                border: 1px solid {StyleManager.BORDER_COLOR};
                border-radius: {StyleManager.BORDER_RADIUS};
                padding: 10px;
                color: {StyleManager.TEXT_PRIMARY};
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 13px;
            }}
            
            QTextEdit:readonly {{
                background-color: {StyleManager.BG_CARD};
            }}
            
            QTreeWidget {{
                background-color: {StyleManager.BG_INPUT};
                border: 1px solid {StyleManager.BORDER_RADIUS};
                border-radius: {StyleManager.BORDER_RADIUS};
                color: {StyleManager.TEXT_PRIMARY};
            }}
            
            QTreeWidget::item {{
                padding: 8px;
            }}
            
            QTreeWidget::item:hover {{
                background-color: {StyleManager.BG_HOVER};
            }}
            
            QTreeWidget::item:selected {{
                background-color: {StyleManager.PRIMARY_COLOR};
            }}
            
            QTabWidget {{
                background-color: {StyleManager.BG_DARK};
                border: none;
            }}
            
            QTabWidget::pane {{
                background-color: {StyleManager.BG_DARK};
                border: none;
            }}
            
            QTabWidget::tab-bar {{
                alignment: left;
            }}
            
            QTabBar::tab {{
                background-color: {StyleManager.BG_CARD};
                color: {StyleManager.TEXT_SECONDARY};
                padding: 12px 24px;
                margin-right: 4px;
                border-radius: {StyleManager.BORDER_RADIUS};
                font-size: 14px;
            }}
            
            QTabBar::tab:hover {{
                background-color: {StyleManager.BG_HOVER};
            }}
            
            QTabBar::tab:selected {{
                background-color: {StyleManager.PRIMARY_COLOR};
                color: {StyleManager.TEXT_PRIMARY};
            }}
            
            QScrollArea {{
                background-color: transparent;
                border: none;
            }}
            
            QScrollBar:vertical {{
                background-color: {StyleManager.BG_CARD};
                width: 8px;
                border-radius: 4px;
            }}
            
            QScrollBar::handle:vertical {{
                background-color: {StyleManager.BORDER_COLOR};
                border-radius: 4px;
            }}
            
            QScrollBar::handle:vertical:hover {{
                background-color: {StyleManager.TEXT_MUTED};
            }}
            
            QGroupBox {{
                background-color: {StyleManager.BG_CARD};
                border: 1px solid {StyleManager.BORDER_COLOR};
                border-radius: {StyleManager.BORDER_RADIUS};
                padding-top: 16px;
            }}
            
            QGroupBox::title {{
                color: {StyleManager.TEXT_SECONDARY};
                font-size: 13px;
                font-weight: 500;
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
            }}
            
            QStatusBar {{
                background-color: {StyleManager.BG_CARD};
                border-top: 1px solid {StyleManager.BORDER_COLOR};
                color: {StyleManager.TEXT_SECONDARY};
                font-size: 12px;
            }}
        """