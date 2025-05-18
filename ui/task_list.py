import os
import pandas as pd
import fnmatch
import re
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QComboBox, QLineEdit, QTableWidget, QTableWidgetItem,
    QGroupBox, QMessageBox, QHeaderView, QSplitter, QMenu, QDialog,
    QStackedWidget, QSizePolicy, QApplication, QStyle
)
from PySide6.QtCore import Qt, QPoint, Signal

from database.db_manager import get_session
from models.task import Task
from models.person import Person
from models.award import Award
from ui.task_detail_dialog import TaskDetailDialog

class TaskListWidget(QWidget):
    """Widget for listing tasks and viewing people with their awards."""
    
    # Signal to notify when a task is selected
    task_selected = Signal(int)
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.load_years()
        self.load_units()
        
    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QVBoxLayout(self)
        
        # Tạo stacked widget để chuyển đổi giữa các chế độ xem
        self.stacked_widget = QStackedWidget()
        
        # Main content widget (task list)
        self.main_content = QWidget()
        main_content_layout = QVBoxLayout(self.main_content)
        
        # Tasks table section
        tasks_group = QGroupBox("Danh sách nhiệm vụ")
        tasks_layout = QVBoxLayout()
        
        # Simple search bar directly above the task table
        search_layout = QHBoxLayout()
        
        # Search bar
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Tìm kiếm theo tên nhiệm vụ, đơn vị, mô tả hoặc năm...")
        self.search_edit.setClearButtonEnabled(True)  # Add clear button inside the search field
        self.search_edit.returnPressed.connect(self.filter_tasks)  # Allow pressing Enter to search
        self.search_edit.setMinimumHeight(30)  # Make search bar slightly taller
        search_layout.addWidget(self.search_edit, 1)  # Stretch factor 1
        
        # Search button
        search_button = QPushButton("Tìm kiếm")
        search_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        search_button.clicked.connect(self.filter_tasks)
        search_layout.addWidget(search_button)
        
        # Reset button
        reset_button = QPushButton("Xóa")
        reset_button.clicked.connect(self.reset_filters)
        search_layout.addWidget(reset_button)
        
        tasks_layout.addLayout(search_layout)
        
        # Hidden filters (not visible to user but still used in code)
        self.year_combo = QComboBox()
        self.year_combo.setVisible(False)
        self.unit_combo = QComboBox()
        self.unit_combo.setVisible(False)
        main_content_layout.addWidget(self.year_combo)
        main_content_layout.addWidget(self.unit_combo)
        
        # Create a splitter for tasks and people
        splitter = QSplitter(Qt.Vertical)
        
        self.tasks_table = QTableWidget()
        self.tasks_table.setColumnCount(4)
        self.tasks_table.setHorizontalHeaderLabels(["ID", "Tên nhiệm vụ", "Năm", "Đơn vị"])
        self.tasks_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tasks_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tasks_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tasks_table.clicked.connect(self.load_people)
        self.tasks_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tasks_table.customContextMenuRequested.connect(self.show_context_menu)
        self.tasks_table.setMinimumHeight(500)  # Kéo dài bảng danh sách nhiệm vụ
        # Đảm bảo bảng mở rộng tối đa theo chiều dọc
        size_policy = QSizePolicy()
        size_policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
        size_policy.setVerticalPolicy(QSizePolicy.Policy.Expanding)
        size_policy.setVerticalStretch(1)  # Đặt ưu tiên cao cho việc mở rộng theo chiều dọc
        self.tasks_table.setSizePolicy(size_policy)
        
        tasks_layout.addWidget(self.tasks_table)
        tasks_group.setLayout(tasks_layout)
        splitter.addWidget(tasks_group)
        
        # People table (initially hidden) - Deprecated but kept for compatibility
        self.people_group = QGroupBox("Danh sách người và danh hiệu")
        self.people_group.setVisible(False)  # Hide initially
        people_layout = QVBoxLayout()
        
        # Add a back button at the top
        back_button = QPushButton("Quay lại danh sách nhiệm vụ")
        back_button.clicked.connect(lambda: self.people_group.setVisible(False))  # Ẩn people_group khi bấm nút quay lại
        people_layout.addWidget(back_button)
        
        self.people_table = QTableWidget()
        self.people_table.setColumnCount(2)
        self.people_table.setHorizontalHeaderLabels(["Họ và tên", "Danh hiệu"])
        self.people_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.people_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        people_layout.addWidget(self.people_table)
        self.people_group.setLayout(people_layout)
        splitter.addWidget(self.people_group)
        
        main_layout.addWidget(splitter)
        
        # Không còn nút làm mới ở đây - Dữ liệu sẽ được tự động làm mới
        
        # Thêm main_content vào stacked widget
        self.stacked_widget.addWidget(self.main_content)
        
        # Thêm stacked widget vào main layout
        main_layout.addWidget(self.stacked_widget)
    
    def load_years(self):
        """Load years from the database."""
        try:
            session = get_session()
            years = session.query(Task.year).distinct().order_by(Task.year.desc()).all()
            session.close()
            
            self.year_combo.clear()
            self.year_combo.addItem("Tất cả", None)
            for year in years:
                self.year_combo.addItem(str(year[0]), year[0])
                
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tải danh sách năm: {str(e)}")
    
    def load_units(self):
        """Load units from the database."""
        try:
            session = get_session()
            units = session.query(Task.unit).distinct().order_by(Task.unit).all()
            session.close()
            
            self.unit_combo.clear()
            self.unit_combo.addItem("Tất cả", None)
            for unit in units:
                self.unit_combo.addItem(unit[0], unit[0])
                
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tải danh sách đơn vị: {str(e)}")
    
    def reset_filters(self):
        """Reset all filters to default values."""
        self.search_edit.clear()
        self.year_combo.setCurrentIndex(0)
        self.unit_combo.setCurrentIndex(0)
        self.filter_tasks()
    
    def filter_tasks(self):
        """Filter tasks based on selected criteria and search term."""
        try:
            session = get_session()
            query = session.query(Task)
            
            # Apply search filter (searches across name, unit, year, and description)
            search_term = self.search_edit.text().strip()
            if search_term:
                search_pattern = f"%{search_term}%"
                
                # Kiểm tra xem search_term có phải là năm không
                try:
                    year_search = int(search_term)
                    query = query.filter(
                        (Task.name.like(search_pattern)) | 
                        (Task.unit.like(search_pattern)) | 
                        (Task.description.like(search_pattern)) |
                        (Task.year == year_search)  # Tìm kiếm chính xác theo năm
                    )
                except ValueError:
                    # Nếu search_term không phải là số, chỉ tìm theo các trường văn bản
                    query = query.filter(
                        (Task.name.like(search_pattern)) | 
                        (Task.unit.like(search_pattern)) | 
                        (Task.description.like(search_pattern))
                    )
            
            # Apply year filter if advanced filters are visible
            if self.year_combo.currentIndex() > 0:
                year = int(self.year_combo.currentText())
                query = query.filter(Task.year == year)
            
            # Apply unit filter if advanced filters are visible
            if self.unit_combo.currentIndex() > 0:
                unit = self.unit_combo.currentText()
                query = query.filter(Task.unit == unit)
            
            # Get filtered tasks
            tasks = query.order_by(Task.year.desc(), Task.name).all()
            
            # Clear tasks table
            self.tasks_table.setRowCount(0)
            
            # Add tasks to table
            for task in tasks:
                row = self.tasks_table.rowCount()
                self.tasks_table.insertRow(row)
                
                # Hiển thị đúng thông tin vào các cột tương ứng
                self.tasks_table.setItem(row, 0, QTableWidgetItem(str(task.id)))
                self.tasks_table.setItem(row, 1, QTableWidgetItem(task.name))
                self.tasks_table.setItem(row, 2, QTableWidgetItem(str(task.year)))
                self.tasks_table.setItem(row, 3, QTableWidgetItem(task.unit))
            
            # Update status with result count
            result_count = self.tasks_table.rowCount()
            if search_term or self.year_combo.currentIndex() > 0 or self.unit_combo.currentIndex() > 0:
                self.setStatusTip(f"Tìm thấy {result_count} nhiệm vụ phù hợp với điều kiện tìm kiếm")
            else:
                self.setStatusTip(f"Hiển thị tất cả {result_count} nhiệm vụ")
            
            session.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lọc nhiệm vụ: {str(e)}")
    
    def load_people(self):
        """Show task detail in a popup dialog."""
        try:
            selected_items = self.tasks_table.selectedItems()
            if not selected_items:
                return
            
            # Get task ID from the first column of the selected row
            row = selected_items[0].row()
            task_id = int(self.tasks_table.item(row, 0).text())
            
            # Tạo và hiển thị dialog chi tiết nhiệm vụ
            detail_dialog = TaskDetailDialog(task_id, self)
            detail_dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tải dữ liệu: {str(e)}")
            print(f"Lỗi khi tải dữ liệu: {str(e)}")  # In lỗi ra console để dễ dàng debug
        

    
    def edit_task(self):
        """Edit the selected task."""
        selected_rows = self.tasks_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn nhiệm vụ để sửa")
            return
        
        task_id = int(self.tasks_table.item(selected_rows[0].row(), 0).text())
        
        try:
            session = get_session()
            task = session.query(Task).filter(Task.id == task_id).first()
            
            if not task:
                session.close()
                QMessageBox.warning(self, "Lỗi", "Không tìm thấy nhiệm vụ")
                return
            
            # Open edit dialog
            from ui.task_edit_dialog import TaskEditDialog
            dialog = TaskEditDialog(self, task)
            result = dialog.exec_()
            
            if result == QDialog.Accepted:
                # Save changes to database
                session.commit()
                QMessageBox.information(self, "Thành công", "Cập nhật nhiệm vụ thành công")
                self.refresh_data()
            
            session.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể sửa nhiệm vụ: {str(e)}")
    
    def delete_task(self):
        """Delete the selected task."""
        selected_rows = self.tasks_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn nhiệm vụ để xóa")
            return
        
        task_id = int(self.tasks_table.item(selected_rows[0].row(), 0).text())
        task_name = self.tasks_table.item(selected_rows[0].row(), 1).text()
        
        # Confirm deletion
        confirm = QMessageBox.question(
            self, "Xác nhận xóa", 
            f"Bạn có chắc chắn muốn xóa nhiệm vụ '{task_name}'?\n"
            "Tất cả dữ liệu liên quan đến nhiệm vụ này sẽ bị xóa.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if confirm != QMessageBox.Yes:
            return
        
        try:
            session = get_session()
            task = session.query(Task).filter(Task.id == task_id).first()
            
            if not task:
                session.close()
                QMessageBox.warning(self, "Lỗi", "Không tìm thấy nhiệm vụ")
                return
            
            # Delete task and all related data (cascade delete will handle relationships)
            session.delete(task)
            session.commit()
            session.close()
            
            QMessageBox.information(self, "Thành công", f"Đã xóa nhiệm vụ '{task_name}'")
            self.refresh_data()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể xóa nhiệm vụ: {str(e)}")
    
    def show_context_menu(self, position):
        """Show context menu for task list."""
        # Get selected row
        selected_rows = self.tasks_table.selectedItems()
        if not selected_rows:
            return
        
        # Create context menu
        context_menu = QMenu()
        view_action = context_menu.addAction("Xem chi tiết")
        edit_action = context_menu.addAction("Sửa nhiệm vụ")
        delete_action = context_menu.addAction("Xóa nhiệm vụ")
        
        # Show context menu at cursor position
        action = context_menu.exec_(self.tasks_table.mapToGlobal(position))
        
        # Handle selected action
        if action == view_action:
            self.load_people()  # This now loads the detail view
        elif action == edit_action:
            self.edit_task()
        elif action == delete_action:
            self.delete_task()
    
    def refresh_data(self):
        """Refresh all data."""
        self.load_years()
        self.load_units()
        self.filter_tasks()
        self.people_table.setRowCount(0)
