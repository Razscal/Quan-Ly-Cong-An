import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QLineEdit, QSpinBox, QTextEdit, QFileDialog, QFormLayout,
    QMessageBox, QGroupBox, QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt, Signal

from database.db_manager import get_session
from models.task import Task
from utils.excel_manager import create_excel_template

class TaskCreationWidget(QWidget):
    """Widget for creating new tasks with Excel templates."""
    
    # Signal emitted when a task is created
    task_created = Signal()
    
    def __init__(self):
        super().__init__()
        self.columns = []
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QVBoxLayout(self)
        
        # Task information section
        task_info_group = QGroupBox("Thông tin nhiệm vụ")
        task_info_layout = QFormLayout()
        
        self.task_name_edit = QLineEdit()
        self.task_year_spin = QSpinBox()
        self.task_year_spin.setRange(2000, 2100)
        self.task_year_spin.setValue(datetime.now().year)
        self.task_unit_edit = QLineEdit()
        self.task_description_edit = QTextEdit()
        
        task_info_layout.addRow("Tên nhiệm vụ:", self.task_name_edit)
        task_info_layout.addRow("Năm:", self.task_year_spin)
        task_info_layout.addRow("Đơn vị:", self.task_unit_edit)
        task_info_layout.addRow("Mô tả:", self.task_description_edit)
        
        task_info_group.setLayout(task_info_layout)
        main_layout.addWidget(task_info_group)
        
        # Excel columns section
        excel_columns_group = QGroupBox("Cột trong file Excel")
        excel_columns_layout = QVBoxLayout()
        
        column_input_layout = QHBoxLayout()
        self.column_name_edit = QLineEdit()
        self.column_name_edit.setPlaceholderText("Tên cột")
        add_column_button = QPushButton("Thêm cột")
        add_column_button.clicked.connect(self.add_column)
        
        column_input_layout.addWidget(self.column_name_edit)
        column_input_layout.addWidget(add_column_button)
        
        excel_columns_layout.addLayout(column_input_layout)
        
        # List of columns
        self.columns_list = QListWidget()
        excel_columns_layout.addWidget(self.columns_list)
        
        # Add default columns
        default_columns = ["Họ và tên", "Năm sinh", "Chức vụ"]
        for column in default_columns:
            self.columns.append(column)
            self.columns_list.addItem(column)
        
        # Remove column button
        remove_column_button = QPushButton("Xóa cột")
        remove_column_button.clicked.connect(self.remove_column)
        excel_columns_layout.addWidget(remove_column_button)
        
        excel_columns_group.setLayout(excel_columns_layout)
        main_layout.addWidget(excel_columns_group)
        
        # Excel file location section
        file_location_group = QGroupBox("Vị trí lưu file Excel")
        file_location_layout = QHBoxLayout()
        
        self.file_location_edit = QLineEdit()
        self.file_location_edit.setReadOnly(True)
        browse_button = QPushButton("Chọn vị trí")
        browse_button.clicked.connect(self.browse_location)
        
        file_location_layout.addWidget(self.file_location_edit)
        file_location_layout.addWidget(browse_button)
        
        file_location_group.setLayout(file_location_layout)
        main_layout.addWidget(file_location_group)
        
        # Create task button
        create_task_button = QPushButton("Tạo nhiệm vụ")
        create_task_button.clicked.connect(self.create_task)
        main_layout.addWidget(create_task_button)
        
    def add_column(self):
        """Add a new column to the list."""
        column_name = self.column_name_edit.text().strip()
        if column_name and column_name not in self.columns:
            self.columns.append(column_name)
            self.columns_list.addItem(column_name)
            self.column_name_edit.clear()
        elif not column_name:
            QMessageBox.warning(self, "Lỗi", "Tên cột không được để trống")
        else:
            QMessageBox.warning(self, "Lỗi", "Cột này đã tồn tại")
    
    def remove_column(self):
        """Remove the selected column from the list."""
        selected_items = self.columns_list.selectedItems()
        if selected_items:
            for item in selected_items:
                self.columns.remove(item.text())
                self.columns_list.takeItem(self.columns_list.row(item))
        else:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn cột để xóa")
    
    def browse_location(self):
        """Open a file dialog to select the Excel file location."""
        # Get directory only
        directory = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục lưu file Excel"
        )
        
        if directory:
            # Generate filename with current date
            from datetime import datetime
            current_date = datetime.now().strftime("%d%m%Y")
            task_name = self.task_name_edit.text().strip() or "task"
            # Remove special characters from task name for filename
            import re
            safe_task_name = re.sub(r'[^\w\s-]', '', task_name).strip().replace(' ', '_')
            
            # Create filename with format: task_name_ddMMyyyy.xlsx
            file_name = f"{safe_task_name}_{current_date}.xlsx"
            full_path = os.path.join(directory, file_name)
            
            self.file_location_edit.setText(full_path)
    
    def create_task(self):
        """Create a new task and generate the Excel template."""
        # Validate inputs
        task_name = self.task_name_edit.text().strip()
        task_year = self.task_year_spin.value()
        task_unit = self.task_unit_edit.text().strip()
        task_description = self.task_description_edit.toPlainText().strip()
        file_location = self.file_location_edit.text().strip()
        
        if not task_name:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập tên nhiệm vụ")
            return
        
        if not task_unit:
            QMessageBox.warning(self, "Lỗi", "Vui lòng nhập đơn vị")
            return
        
        if not file_location:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn vị trí lưu file Excel")
            return
        
        if not self.columns:
            QMessageBox.warning(self, "Lỗi", "Vui lòng thêm ít nhất một cột")
            return
        
        try:
            # Create task folder first
            import re
            safe_task_name = re.sub(r'[^\w\s-]', '', task_name).strip().replace(' ', '_')
            base_dir = os.path.dirname(file_location)
            task_folder = os.path.join(base_dir, safe_task_name)
            
            # Create the folder if it doesn't exist
            if not os.path.exists(task_folder):
                os.makedirs(task_folder)
            
            # Create the Excel file directly in the task folder
            filename = os.path.basename(file_location)
            excel_path = os.path.join(task_folder, filename)
            
            # Create Excel template in the task folder
            create_excel_template(excel_path, self.columns)
            
            # Save task to database with the correct path
            session = get_session()
            new_task = Task(
                name=task_name,
                year=task_year,
                unit=task_unit,
                description=task_description,
                excel_path=excel_path,
                created_at=datetime.now().date()
            )
            session.add(new_task)
            session.commit()
            
            session.close()
            
            QMessageBox.information(
                self, "Thành công", 
                f"Đã tạo nhiệm vụ '{task_name}' và file Excel tại:\n{file_location}"
            )
            
            # Emit signal to notify that a task has been created
            self.task_created.emit()
            
            # Clear form
            self.task_name_edit.clear()
            self.task_unit_edit.clear()
            self.task_description_edit.clear()
            self.file_location_edit.clear()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tạo nhiệm vụ: {str(e)}")
