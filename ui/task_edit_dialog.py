import os
from datetime import datetime
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QLineEdit, QSpinBox, QTextEdit, QFormLayout, QMessageBox
)
from PySide6.QtCore import Qt

from models.task import Task

class TaskEditDialog(QDialog):
    """Dialog for editing task information."""
    
    def __init__(self, parent=None, task=None):
        super().__init__(parent)
        self.task = task
        self.setup_ui()
        if task:
            self.load_task_data()
        
    def setup_ui(self):
        """Set up the user interface."""
        self.setWindowTitle("Sửa thông tin nhiệm vụ")
        self.setMinimumWidth(400)
        
        main_layout = QVBoxLayout(self)
        
        # Task information form
        form_layout = QFormLayout()
        
        self.task_name_edit = QLineEdit()
        self.task_year_spin = QSpinBox()
        self.task_year_spin.setRange(2000, 2100)
        self.task_year_spin.setValue(datetime.now().year)
        self.task_unit_edit = QLineEdit()
        self.task_description_edit = QTextEdit()
        
        form_layout.addRow("Tên nhiệm vụ:", self.task_name_edit)
        form_layout.addRow("Năm:", self.task_year_spin)
        form_layout.addRow("Đơn vị:", self.task_unit_edit)
        form_layout.addRow("Mô tả:", self.task_description_edit)
        
        main_layout.addLayout(form_layout)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("Lưu")
        save_button.clicked.connect(self.save_task)
        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(self.reject)
        
        buttons_layout.addWidget(save_button)
        buttons_layout.addWidget(cancel_button)
        
        main_layout.addLayout(buttons_layout)
    
    def load_task_data(self):
        """Load task data into the form."""
        if not self.task:
            return
        
        self.task_name_edit.setText(self.task.name)
        self.task_year_spin.setValue(self.task.year)
        self.task_unit_edit.setText(self.task.unit)
        if self.task.description:
            self.task_description_edit.setText(self.task.description)
    
    def save_task(self):
        """Save the task data."""
        # Validate inputs
        name = self.task_name_edit.text().strip()
        year = self.task_year_spin.value()
        unit = self.task_unit_edit.text().strip()
        description = self.task_description_edit.toPlainText().strip()
        
        if not name:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập tên nhiệm vụ")
            return
        
        if not unit:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập đơn vị")
            return
        
        # Store old name for folder renaming
        old_name = self.task.name
        
        # Update task object
        self.task.name = name
        self.task.year = year
        self.task.unit = unit
        self.task.description = description
        
        # Rename task folder if name has changed
        if old_name != name:
            try:
                self.task.rename_task_folder(old_name)
            except Exception as e:
                print(f"Error renaming task folder: {str(e)}")
        
        self.accept()
