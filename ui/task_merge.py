import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QFileDialog, QMessageBox, QListWidget, QListWidgetItem,
    QGroupBox, QComboBox
)
from PySide6.QtCore import Qt

from database.db_manager import get_session
from models.task import Task
from utils.excel_manager import merge_excel_files, import_excel_data

class TaskMergeWidget(QWidget):
    """Widget for merging Excel files and importing data."""
    
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.setup_ui()
        self.load_tasks()
        
    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QVBoxLayout(self)
        
        # Task selection section
        task_selection_group = QGroupBox("Chọn nhiệm vụ")
        task_selection_layout = QVBoxLayout()
        
        self.task_combo = QComboBox()
        self.task_combo.setPlaceholderText("Chọn nhiệm vụ")
        task_selection_layout.addWidget(self.task_combo)
        
        task_selection_group.setLayout(task_selection_layout)
        main_layout.addWidget(task_selection_group)
        
        # File selection section
        file_selection_group = QGroupBox("Chọn file Excel để trộn")
        file_selection_layout = QVBoxLayout()
        
        file_buttons_layout = QHBoxLayout()
        add_file_button = QPushButton("Thêm file")
        add_file_button.clicked.connect(self.add_files)
        remove_file_button = QPushButton("Xóa file")
        remove_file_button.clicked.connect(self.remove_file)
        
        file_buttons_layout.addWidget(add_file_button)
        file_buttons_layout.addWidget(remove_file_button)
        
        file_selection_layout.addLayout(file_buttons_layout)
        
        self.files_list = QListWidget()
        file_selection_layout.addWidget(self.files_list)
        
        file_selection_group.setLayout(file_selection_layout)
        main_layout.addWidget(file_selection_group)
        
        # Merge button
        buttons_layout = QHBoxLayout()
        
        merge_button = QPushButton("Tiến hành trộn file")
        merge_button.clicked.connect(self.merge_and_import_files)
        
        buttons_layout.addWidget(merge_button)
        
        main_layout.addLayout(buttons_layout)
    
    def load_tasks(self):
        """Load tasks from the database."""
        try:
            session = get_session()
            tasks = session.query(Task).all()
            session.close()
            
            self.task_combo.clear()
            for task in tasks:
                self.task_combo.addItem(f"{task.name} ({task.year}) - {task.unit}", task.id)
                
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tải danh sách nhiệm vụ: {str(e)}")
    
    def add_files(self):
        """Open a file dialog to select Excel files."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Chọn file Excel", "", "Excel Files (*.xlsx)"
        )
        
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.files_list.addItem(file)
    
    def remove_file(self):
        """Remove the selected file from the list."""
        selected_items = self.files_list.selectedItems()
        if selected_items:
            for item in selected_items:
                self.selected_files.remove(item.text())
                self.files_list.takeItem(self.files_list.row(item))
        else:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn file để xóa")
    
    def merge_and_import_files(self):
        """Merge selected Excel files and import the data."""
        if not self.selected_files:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn ít nhất một file Excel")
            return
        
        if self.task_combo.currentIndex() == -1:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn nhiệm vụ")
            return
        
        task_id = self.task_combo.currentData()
        task_name = self.task_combo.currentText()
        
        try:
            # Open a session for getting the task
            session = get_session()
            task = session.query(Task).filter(Task.id == task_id).first()
            
            if not task:
                session.close()
                QMessageBox.warning(self, "Lỗi", "Không tìm thấy nhiệm vụ")
                return
            
            # Get the original task Excel file
            original_file = task.excel_path
            if not os.path.exists(original_file):
                session.close()
                QMessageBox.warning(self, "Lỗi", f"Không tìm thấy file gốc của nhiệm vụ: {original_file}")
                return
            
            # Add the original file to the list of files to merge if not already included
            all_files = self.selected_files.copy()
            if original_file not in all_files:
                all_files.insert(0, original_file)  # Insert at the beginning to ensure it's processed first
            
            # Sử dụng file nguồn làm file đầu ra
            output_file = original_file
            
            # Merge files
            merge_excel_files(all_files, output_file)
            
            # Import data from merged file
            import_excel_data(output_file, task, session)
            
            # Commit changes and close session
            session.commit()
            session.close()
            
            # Clear selected files after successful merge and import
            num_files = len(self.selected_files)
            self.selected_files = []
            self.files_list.clear()
            
            QMessageBox.information(
                self, "Thành công", 
                f"Đã trộn {num_files} file Excel trực tiếp vào file gốc và import dữ liệu vào nhiệm vụ '{task_name}'\n"
                f"Dữ liệu đã được cập nhật trong file: {os.path.basename(output_file)}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể trộn file và import dữ liệu: {str(e)}")
