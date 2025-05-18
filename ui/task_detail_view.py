import os
import pandas as pd
import re
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QLineEdit, QTableWidget, QTableWidgetItem, QComboBox,
    QGroupBox, QMessageBox, QHeaderView, QCheckBox, QSplitter,
    QApplication, QFileDialog, QStyle, QMenu, QDialog, QFormLayout
)
from PySide6.QtCore import Qt, Signal, QSortFilterProxyModel, QRegularExpression
import subprocess
from PySide6.QtGui import QStandardItemModel, QStandardItem

from models.task import Task
from database.db_manager import get_session


class TaskDetailView(QWidget):
    """Widget for displaying and searching merged Excel data for a task."""
    
    # Signal to go back to task list
    back_signal = Signal()
    
    def __init__(self, task_id=None):
        super().__init__()
        self.task_id = task_id
        self.task = None
        self.df = None
        self.merged_file = None
        self.filter_columns = []
        self.setup_ui()
        
        if task_id:
            self.load_task_data(task_id)
    
    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)  # Thêm margin để giao diện đẹp hơn
        
        # Header with task info and buttons
        header_layout = QHBoxLayout()
        
        # Back button with icon and better styling
        back_button = QPushButton("Quay lại danh sách nhiệm vụ")
        back_button.setIcon(QApplication.style().standardIcon(QStyle.SP_ArrowBack))
        back_button.setStyleSheet("padding: 8px; font-weight: bold;")
        back_button.clicked.connect(self.go_back)
        header_layout.addWidget(back_button)
        

        
        # Task info label with better styling
        self.task_info_label = QLabel("Chi tiết nhiệm vụ")
        self.task_info_label.setStyleSheet("font-weight: bold; font-size: 16px; color: #2E7D32;")
        header_layout.addWidget(self.task_info_label, 1)  # Stretch factor 1
        
        # Export button with icon and styling
        export_button = QPushButton("Xuất Excel")
        export_button.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        export_button.setStyleSheet("padding: 8px; background-color: #4CAF50; color: white;")
        export_button.clicked.connect(self.export_to_excel)
        header_layout.addWidget(export_button)
        
        # Open source file button with better styling
        self.open_source_file_button = QPushButton("Mở file nguồn")
        self.open_source_file_button.setIcon(QApplication.style().standardIcon(QStyle.SP_FileIcon))
        self.open_source_file_button.setStyleSheet("padding: 8px;")
        self.open_source_file_button.clicked.connect(self.open_source_file)
        header_layout.addWidget(self.open_source_file_button)
        
        main_layout.addLayout(header_layout)
        
        # Search section with improved styling
        search_group = QGroupBox("Tìm kiếm")
        search_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 1px solid #4CAF50;
                border-radius: 5px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
                color: #2E7D32;
            }
        """)
        search_layout = QVBoxLayout()
        search_layout.setSpacing(10)  # Thêm khoảng cách giữa các phần tử
        
        # Search fields layout
        search_fields_layout = QHBoxLayout()
        search_fields_layout.setSpacing(15)  # Thêm khoảng cách giữa các trường tìm kiếm
        
        # Global search with better styling
        global_search_layout = QHBoxLayout()
        global_search_label = QLabel("Tìm kiếm tổng thể:")
        global_search_label.setStyleSheet("font-weight: bold;")
        self.global_search_input = QLineEdit()
        self.global_search_input.setPlaceholderText("Nhập từ khóa tìm kiếm...")
        self.global_search_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #4CAF50;
                border-radius: 4px;
            }
        """)
        self.global_search_input.setClearButtonEnabled(True)  # Thêm nút xóa
        self.global_search_input.textChanged.connect(self.apply_filters)
        global_search_layout.addWidget(global_search_label)
        global_search_layout.addWidget(self.global_search_input, 1)  # Stretch factor 1
        
        search_fields_layout.addLayout(global_search_layout)
        
        # Column selector with better styling
        column_layout = QHBoxLayout()
        column_label = QLabel("Tìm theo cột:")
        column_label.setStyleSheet("font-weight: bold;")
        self.column_combo = QComboBox()
        self.column_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #4CAF50;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                border: 0px;
            }
            QComboBox::down-arrow {
                width: 14px;
                height: 14px;
            }
        """)
        self.column_combo.currentIndexChanged.connect(self.update_column_filter)
        column_layout.addWidget(column_label)
        column_layout.addWidget(self.column_combo, 1)  # Stretch factor 1
        
        search_fields_layout.addLayout(column_layout)
        
        # Column value search with better styling
        column_value_layout = QHBoxLayout()
        self.column_value_input = QLineEdit()
        self.column_value_input.setPlaceholderText("Giá trị cột...")
        self.column_value_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #4CAF50;
                border-radius: 4px;
            }
        """)
        self.column_value_input.setClearButtonEnabled(True)  # Thêm nút xóa
        self.column_value_input.textChanged.connect(self.apply_filters)
        column_value_layout.addWidget(self.column_value_input, 1)  # Stretch factor 1
        
        search_fields_layout.addLayout(column_value_layout)
        
        search_layout.addLayout(search_fields_layout)
        
        # Advanced filter options with better styling
        advanced_layout = QHBoxLayout()
        advanced_layout.setSpacing(15)  # Thêm khoảng cách giữa các tùy chọn
        
        # Case sensitive option
        self.case_sensitive_check = QCheckBox("Phân biệt hoa/thường")
        self.case_sensitive_check.setStyleSheet("padding: 5px;")
        self.case_sensitive_check.stateChanged.connect(self.apply_filters)
        advanced_layout.addWidget(self.case_sensitive_check)
        
        # Exact match option
        self.exact_match_check = QCheckBox("Khớp chính xác")
        self.exact_match_check.setStyleSheet("padding: 5px;")
        self.exact_match_check.stateChanged.connect(self.apply_filters)
        advanced_layout.addWidget(self.exact_match_check)
        
        # Reset filters button with better styling
        reset_button = QPushButton("Xóa bộ lọc")
        reset_button.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogResetButton))
        reset_button.setStyleSheet("""
            QPushButton {
                padding: 6px 12px;
                background-color: #f44336;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        reset_button.clicked.connect(self.reset_filters)
        advanced_layout.addWidget(reset_button)
        
        search_layout.addLayout(advanced_layout)
        
        search_group.setLayout(search_layout)
        main_layout.addWidget(search_group)
        
        # Data table with improved styling
        self.data_table = QTableWidget()
        self.data_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.data_table.customContextMenuRequested.connect(self.show_context_menu)
        self.data_table.setSortingEnabled(True)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #4CAF50;
                border-radius: 5px;
                gridline-color: #E0E0E0;
                selection-background-color: #81C784;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 6px;
                border: none;
            }
            QTableWidget::item:selected {
                background-color: #C8E6C9;
                color: #2E7D32;
            }
        """)
        # Tăng kích thước của bảng dữ liệu
        self.data_table.setMinimumHeight(400)  # Đặt chiều cao tối thiểu
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_table.horizontalHeader().setStretchLastSection(True)
        self.data_table.verticalHeader().setDefaultSectionSize(30)  # Tăng chiều cao của mỗi dòng
        main_layout.addWidget(self.data_table, 1)  # Stretch factor 1
        
        # Status bar with better styling
        self.status_label = QLabel("Chưa có dữ liệu")
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #E8F5E9;
                color: #2E7D32;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
        """)
        main_layout.addWidget(self.status_label)
    
    def go_back(self):
        """Go back to task list."""
        # Phát tín hiệu để quay lại danh sách nhiệm vụ
        self.back_signal.emit()
        

    
    def load_task_data(self, task_id):
        """Load task data and find merged Excel file."""
        try:
            session = get_session()
            self.task = session.query(Task).filter_by(id=task_id).first()
            session.close()
            
            if not self.task:
                QMessageBox.warning(self, "Cảnh báo", "Không tìm thấy nhiệm vụ")
                return
            
            # Cập nhật tiêu đề
            self.task_info_label.setText(f"Chi tiết nhiệm vụ: {self.task.name} ({self.task.year}) - {self.task.unit}")
            
            # Sử dụng file nguồn làm nguồn dữ liệu duy nhất
            if os.path.exists(self.task.excel_path):
                self.merged_file = self.task.excel_path
                self.load_excel_data(self.task.excel_path)
                self.status_label.setText(f"Đã tải dữ liệu từ file nguồn: {os.path.basename(self.task.excel_path)}")
            else:
                QMessageBox.warning(self, "Cảnh báo", "Không tìm thấy file Excel nào cho nhiệm vụ này")
        
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể tải dữ liệu nhiệm vụ: {str(e)}")
    
    def load_excel_data(self, file_path):
        """Load data from Excel file into table."""
        try:
            # Đọc file Excel vào DataFrame
            self.df = pd.read_excel(file_path)
            
            # Hiển thị dữ liệu trong bảng
            self.populate_table(self.df)
            
            # Cập nhật các tùy chọn lọc theo cột
            self.update_column_filter_options()
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc file Excel: {str(e)}")
    
    def populate_table(self, dataframe):
        """Populate table with dataframe data."""
        if dataframe is None or dataframe.empty:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            self.status_label.setText("Không có dữ liệu")
            return
        
        # Thiết lập số dòng và cột
        self.data_table.setRowCount(len(dataframe))
        self.data_table.setColumnCount(len(dataframe.columns))
        
        # Thiết lập tiêu đề cột
        self.data_table.setHorizontalHeaderLabels(dataframe.columns)
        
        # Điền dữ liệu vào bảng
        for row in range(len(dataframe)):
            for col in range(len(dataframe.columns)):
                value = dataframe.iloc[row, col]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                self.data_table.setItem(row, col, item)
        
        # Cập nhật trạng thái
        self.status_label.setText(f"Hiển thị {len(dataframe)} dòng")
    
    def update_column_filter_options(self):
        """Update column filter dropdown with available columns."""
        if self.df is None:
            return
        
        self.column_combo.clear()
        self.column_combo.addItem("-- Chọn cột --")
        
        for column in self.df.columns:
            self.column_combo.addItem(str(column))
        
        self.filter_columns = list(self.df.columns)
    
    def update_column_filter(self):
        """Update column filter when selection changes."""
        self.apply_filters()
    
    def apply_filters(self):
        """Apply all filters to the data."""
        if self.df is None:
            return
        
        # Bắt đầu với toàn bộ dữ liệu
        filtered_df = self.df.copy()
        
        # Lấy các giá trị từ bộ lọc
        global_search = self.global_search_input.text().strip()
        column_index = self.column_combo.currentIndex()
        column_value = self.column_value_input.text().strip()
        case_sensitive = self.case_sensitive_check.isChecked()
        exact_match = self.exact_match_check.isChecked()
        
        # Áp dụng tìm kiếm toàn cục
        if global_search:
            # Tạo mask cho mỗi cột và kết hợp chúng với OR
            global_mask = pd.Series(False, index=filtered_df.index)
            
            for column in filtered_df.columns:
                if case_sensitive:
                    column_mask = filtered_df[column].astype(str).str.contains(global_search, regex=False, na=False)
                else:
                    column_mask = filtered_df[column].astype(str).str.contains(global_search, regex=False, case=False, na=False)
                
                global_mask = global_mask | column_mask
            
            filtered_df = filtered_df[global_mask]
        
        # Áp dụng lọc theo cột cụ thể
        if column_index > 0 and column_value:  # Chỉ áp dụng nếu đã chọn cột và có giá trị
            column_name = self.column_combo.itemText(column_index)
            
            if exact_match:
                if case_sensitive:
                    mask = filtered_df[column_name].astype(str) == column_value
                else:
                    mask = filtered_df[column_name].astype(str).str.lower() == column_value.lower()
            else:
                if case_sensitive:
                    mask = filtered_df[column_name].astype(str).str.contains(column_value, regex=False, na=False)
                else:
                    mask = filtered_df[column_name].astype(str).str.contains(column_value, regex=False, case=False, na=False)
            
            filtered_df = filtered_df[mask]
        
        # Cập nhật bảng với dữ liệu đã lọc
        self.populate_table(filtered_df)
        
        # Cập nhật trạng thái
        if global_search or (column_index > 0 and column_value):
            self.status_label.setText(f"Đã lọc: {len(filtered_df)} dòng từ {len(self.df)} dòng")
        else:
            self.status_label.setText(f"Hiển thị tất cả {len(filtered_df)} dòng.")
    
    def reset_filters(self):
        """Reset all filters."""
        self.global_search_input.clear()
        self.column_combo.setCurrentIndex(0)
        self.column_value_input.clear()
        self.case_sensitive_check.setChecked(False)
        self.exact_match_check.setChecked(False)
        
        # Tải lại dữ liệu gốc
        if self.df is not None:
            self.populate_table(self.df)
            self.status_label.setText(f"Đã xóa bộ lọc. Hiển thị tất cả {len(self.df)} dòng.")
    
    def open_source_file(self):
        """Open the source Excel file for the task."""
        if not self.task:
            QMessageBox.warning(self, "Cảnh báo", "Không có nhiệm vụ nào được chọn")
            return
        
        excel_path = self.task.excel_path
        
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "Cảnh báo", "File nguồn không tồn tại")
            return
        
        try:
            # Mở file Excel bằng ứng dụng mặc định
            if os.name == 'nt':  # Windows
                os.startfile(excel_path)
            elif os.name == 'posix':  # macOS or Linux
                subprocess.Popen(['open', excel_path])
            
            QMessageBox.information(self, "Thông báo", f"Đã mở file nguồn: {os.path.basename(excel_path)}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở file nguồn: {str(e)}")
    
    def show_context_menu(self, position):
        """Show context menu for data table."""
        # Kiểm tra xem có dòng nào được chọn không
        selected_items = self.data_table.selectedItems()
        if not selected_items:
            return
        
        # Lấy dòng được chọn
        row = selected_items[0].row()
        
        # Tạo menu chuột phải
        context_menu = QMenu()
        edit_action = context_menu.addAction("Sửa dòng này")
        delete_action = context_menu.addAction("Xóa dòng này")
        
        # Hiển thị menu tại vị trí con trỏ
        action = context_menu.exec_(self.data_table.mapToGlobal(position))
        
        # Xử lý hành động được chọn
        if action == edit_action:
            self.edit_record(row)
        elif action == delete_action:
            self.delete_record(row)
    
    def edit_record(self, row):
        """Edit a record in the data table and sync back to Excel."""
        if self.df is None or row >= len(self.df):
            return
        
        # Tạo dialog để chỉnh sửa
        dialog = QDialog(self)
        dialog.setWindowTitle("Sửa dữ liệu")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(dialog)
        form_layout = QFormLayout()
        
        # Tạo các trường nhập liệu cho mỗi cột
        fields = {}
        for header in self.df.columns:
            value = self.df.iloc[row][header]
            field = QLineEdit()
            field.setText(str(value) if pd.notna(value) else "")
            form_layout.addRow(str(header) + ":", field)
            fields[header] = field
        
        layout.addLayout(form_layout)
        
        # Thêm các nút OK và Cancel
        button_layout = QHBoxLayout()
        ok_button = QPushButton("Lưu")
        cancel_button = QPushButton("Hủy")
        
        ok_button.clicked.connect(dialog.accept)
        cancel_button.clicked.connect(dialog.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        # Hiển thị dialog
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            # Cập nhật dữ liệu trong dataframe
            for header, field in fields.items():
                self.df.at[row, header] = field.text()
            
            # Cập nhật bảng
            self.populate_table(self.df)
            
            # Đồng bộ với file Excel
            self.sync_to_excel()
            
            QMessageBox.information(self, "Thành công", "Dữ liệu đã được cập nhật và đồng bộ với file Excel")
    
    def delete_record(self, row):
        """Delete a record from the data table and sync back to Excel."""
        if self.df is None or row >= len(self.df):
            return
        
        # Xác nhận xóa
        confirm = QMessageBox.question(
            self, "Xác nhận xóa", 
            "Bạn có chắc chắn muốn xóa dòng này?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            # Xóa dòng khỏi dataframe
            self.df = self.df.drop(self.df.index[row]).reset_index(drop=True)
            
            # Cập nhật bảng
            self.populate_table(self.df)
            
            # Đồng bộ với file Excel
            self.sync_to_excel()
            
            QMessageBox.information(self, "Thành công", "Dòng đã được xóa và đồng bộ với file Excel")
    
    def sync_to_excel(self):
        """Sync the current dataframe back to the Excel file."""
        if self.df is None or not hasattr(self, 'merged_file') or not self.merged_file:
            return False
        
        try:
            # Lưu dataframe vào file Excel
            with pd.ExcelWriter(self.merged_file, engine='openpyxl') as writer:
                self.df.to_excel(writer, index=False, sheet_name="Nhiệm vụ")
            
            return True
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đồng bộ với file Excel: {str(e)}")
            return False
    
    def export_to_excel(self):
        """Export current filtered data to Excel."""
        if self.df is None:
            QMessageBox.warning(self, "Cảnh báo", "Không có dữ liệu để xuất")
            return
        
        # Get visible rows from table
        visible_rows = []
        for row in range(self.data_table.rowCount()):
            row_data = {}
            for col in range(self.data_table.columnCount()):
                header = self.data_table.horizontalHeaderItem(col).text()
                item = self.data_table.item(row, col)
                if item:
                    row_data[header] = item.text()
                else:
                    row_data[header] = ""
            visible_rows.append(row_data)
        
        if not visible_rows:
            QMessageBox.warning(self, "Cảnh báo", "Không có dữ liệu để xuất")
            return
        
        # Create dataframe from visible rows
        export_df = pd.DataFrame(visible_rows)
        
        # Get save location
        file_name = f"{self.task.name}_filtered_{pd.Timestamp.now().strftime('%d%m%Y')}.xlsx"
        safe_file_name = re.sub(r'[^\w\s-]', '', file_name).strip().replace(' ', '_')
        
        task_folder = os.path.dirname(self.task.excel_path)
        default_path = os.path.join(task_folder, safe_file_name)
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file Excel", default_path, "Excel Files (*.xlsx)"
        )
        
        if not file_path:
            return
        
        try:
            export_df.to_excel(file_path, index=False)
            QMessageBox.information(
                self, "Thành công", f"Đã xuất dữ liệu ra file:\n{file_path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể xuất file Excel: {str(e)}")
