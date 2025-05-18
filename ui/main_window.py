import os
from datetime import datetime
from PySide6.QtWidgets import (
    QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QMessageBox, QComboBox, QLineEdit
)
from PySide6.QtGui import QColor, QPalette, QFont, QIcon
from PySide6.QtCore import Qt

from ui.task_creation import TaskCreationWidget
from ui.task_merge import TaskMergeWidget
from ui.task_list import TaskListWidget

# Define color scheme
PRIMARY_COLOR = "#4CAF50"  # Green
SECONDARY_COLOR = "#FFFFFF"  # White
ACCENT_COLOR = "#2E7D32"  # Dark Green

class MainWindow(QMainWindow):
    """Main application window with tabs for different functionalities."""
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Quản Lý Nhiệm Vụ - Công An")
        self.setMinimumSize(1000, 700)
        self.setup_ui()
        self.apply_styles()
        
    def setup_ui(self):
        """Set up the user interface."""
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        
        # Create header
        header_layout = QHBoxLayout()
        logo_label = QLabel("CÔNG AN NHÂN DÂN")
        logo_label.setFont(QFont("Arial", 16, QFont.Bold))
        header_layout.addWidget(logo_label)
        
        # Add current date
        date_label = QLabel(datetime.now().strftime("%d/%m/%Y"))
        date_label.setAlignment(Qt.AlignRight)
        header_layout.addWidget(date_label)
        
        main_layout.addLayout(header_layout)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Create tabs
        self.task_creation_tab = TaskCreationWidget()
        self.task_merge_tab = TaskMergeWidget()
        self.task_list_tab = TaskListWidget()
        
        # Connect signals between tabs
        self.task_creation_tab.task_created.connect(self.on_task_created)
        
        # Add tabs to tab widget
        self.tab_widget.addTab(self.task_creation_tab, "Tạo Nhiệm Vụ")
        self.tab_widget.addTab(self.task_merge_tab, "Trộn File")
        self.tab_widget.addTab(self.task_list_tab, "Danh Sách Nhiệm Vụ")
        
        # Connect tab change signal to handle tab switching
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        
        main_layout.addWidget(self.tab_widget)
        
        # Create footer
        footer_layout = QHBoxLayout()
        footer_label = QLabel("© 2025 - Phần mềm Quản Lý Nhiệm Vụ")
        footer_layout.addWidget(footer_label)
        
        main_layout.addLayout(footer_layout)
    
    def on_tab_changed(self, index):
        """Handle tab change events."""
        # If switching to task list tab (index 2), refresh the data
        if index == 2:  # Task list tab
            self.task_list_tab.refresh_data()
            # Make sure we're showing the task list view, not the detail view
            if hasattr(self.task_list_tab, 'stacked_widget') and hasattr(self.task_list_tab, 'main_content'):
                self.task_list_tab.show_task_list()
        # If switching to merge tab (index 1), refresh the task list
        elif index == 1:  # Merge tab
            self.task_merge_tab.load_tasks()
    
    def on_task_created(self):
        """Handle task creation event."""
        # Refresh task list in merge tab
        self.task_merge_tab.load_tasks()
        # Also refresh task list tab
        self.task_list_tab.refresh_data()
    

    
    def apply_styles(self):
        """Apply styles to the application."""
        # Set application style
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {SECONDARY_COLOR};
            }}
            QTabWidget::pane {{
                border: 1px solid #cccccc;
                background-color: {SECONDARY_COLOR};
            }}
            QTabBar::tab {{
                background-color: #e0e0e0;
                padding: 8px 20px;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background-color: {PRIMARY_COLOR};
                color: {SECONDARY_COLOR};
                font-weight: bold;
            }}
            QPushButton {{
                background-color: {PRIMARY_COLOR};
                color: {SECONDARY_COLOR};
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {ACCENT_COLOR};
            }}
            QLabel {{
                color: #333333;
            }}
            QLineEdit, QComboBox {{
                padding: 6px;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }}
        """)
