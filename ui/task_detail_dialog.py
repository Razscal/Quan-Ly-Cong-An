import os
import pandas as pd
import re
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QLineEdit, QTableWidget, QTableWidgetItem, QComboBox,
    QGroupBox, QMessageBox, QHeaderView, QCheckBox, QSplitter,
    QApplication, QFileDialog, QStyle, QMenu, QFormLayout,
    QDialogButtonBox
)
from PySide6.QtCore import Qt, Signal, QSortFilterProxyModel, QRegularExpression
import subprocess
from PySide6.QtGui import QStandardItemModel, QStandardItem

from models.task import Task
from database.db_manager import get_session


class TaskDetailDialog(QDialog):
    """Dialog for displaying and searching merged Excel data for a task."""
    
    def __init__(self, task_id=None, parent=None):
        super().__init__(parent)
        self.task_id = task_id
        self.task = None
        self.df = None
        self.merged_file = None
        self.filter_columns = []
        
        # Thiết lập thuộc tính cửa sổ
        self.setWindowTitle("Chi tiết nhiệm vụ")
        self.resize(1400, 900)  # Kích thước lớn hơn để dễ nhìn
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint | Qt.WindowMinimizeButtonHint)
        # Đặt cửa sổ ở chế độ tối đa hóa theo mặc định
        self.setWindowState(Qt.WindowMaximized)
        
        self.setup_ui()
        
        if task_id:
            self.load_task_data(task_id)
    
    def setup_ui(self):
        """Set up the user interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)  # Thêm margin để giao diện đẹp hơn
        
        # Header with task info and buttons
        header_layout = QHBoxLayout()
        
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
        self.global_search_input.setPlaceholderText("Nhập từ khóa để tìm kiếm trong tất cả các cột...")
        self.global_search_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #BDBDBD;
                border-radius: 4px;
            }
            QLineEdit:focus {
                border: 1px solid #4CAF50;
            }
        """)
        self.global_search_input.setClearButtonEnabled(True)  # Thêm nút xóa
        self.global_search_input.textChanged.connect(self.apply_filters)
        global_search_layout.addWidget(global_search_label)
        global_search_layout.addWidget(self.global_search_input, 1)  # Stretch factor 1
        
        search_fields_layout.addLayout(global_search_layout)
        
        # Column filter with better styling
        column_filter_layout = QHBoxLayout()
        column_filter_label = QLabel("Lọc theo cột:")
        column_filter_label.setStyleSheet("font-weight: bold;")
        self.column_filter_combo = QComboBox()
        self.column_filter_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #BDBDBD;
                border-radius: 4px;
            }
            QComboBox:focus {
                border: 1px solid #4CAF50;
            }
        """)
        self.column_filter_combo.currentIndexChanged.connect(self.update_column_filter)
        column_filter_layout.addWidget(column_filter_label)
        column_filter_layout.addWidget(self.column_filter_combo, 1)  # Stretch factor 1
        
        search_fields_layout.addLayout(column_filter_layout)
        
        # Column value with better styling
        column_value_layout = QHBoxLayout()
        column_value_label = QLabel("Giá trị:")
        column_value_label.setStyleSheet("font-weight: bold;")
        self.column_value_input = QLineEdit()
        self.column_value_input.setPlaceholderText("Nhập giá trị để lọc theo cột đã chọn...")
        self.column_value_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #BDBDBD;
                border-radius: 4px;
            }
            QLineEdit:focus {
                border: 1px solid #4CAF50;
            }
        """)
        self.column_value_input.setClearButtonEnabled(True)  # Thêm nút xóa
        self.column_value_input.textChanged.connect(self.apply_filters)
        column_value_layout.addWidget(column_value_label)
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
        self.data_table.setMinimumHeight(600)  # Đặt chiều cao tối thiểu lớn hơn
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)  # Cho phép người dùng điều chỉnh kích thước cột
        self.data_table.horizontalHeader().setStretchLastSection(True)
        self.data_table.verticalHeader().setDefaultSectionSize(35)  # Tăng chiều cao của mỗi dòng
        # Tăng kích thước font chữ
        font = self.data_table.font()
        font.setPointSize(10)  # Tăng kích thước font
        self.data_table.setFont(font)
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
        
        # Thêm các nút ở dưới cùng
        button_layout = QHBoxLayout()
        
        # Nút làm mới dữ liệu
        refresh_button = QPushButton("Làm mới dữ liệu")
        refresh_button.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_button.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        refresh_button.clicked.connect(lambda: self.load_task_data(self.task_id))
        button_layout.addWidget(refresh_button)
        
        # Thêm khoảng trống để đẩy nút đóng sang phải
        button_layout.addStretch(1)
        
        # Nút đóng
        close_button = QPushButton("Đóng")
        close_button.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogCloseButton))
        close_button.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                background-color: #757575;
                color: white;
                font-weight: bold;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #616161;
            }
        """)
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        
        main_layout.addLayout(button_layout)
    
    def load_task_data(self, task_id):
        """Load task data and find merged Excel file."""
        try:
            # Get task from database
            session = get_session()
            self.task = session.query(Task).filter(Task.id == task_id).first()
            session.close()
            
            if not self.task:
                self.status_label.setText(f"Không tìm thấy nhiệm vụ với ID: {task_id}")
                return
            
            # Update task info label
            self.task_info_label.setText(f"Chi tiết nhiệm vụ: {self.task.name} - {self.task.unit} ({self.task.year})")
            
            # Set window title
            self.setWindowTitle(f"Chi tiết nhiệm vụ: {self.task.name}")
            
            # Find Excel file
            if self.task.excel_path and os.path.exists(self.task.excel_path):
                self.merged_file = self.task.excel_path
                self.load_excel_data(self.merged_file)
            else:
                self.status_label.setText(f"Không tìm thấy file Excel cho nhiệm vụ: {self.task.name}")
        
        except Exception as e:
            self.status_label.setText(f"Lỗi khi tải dữ liệu: {str(e)}")
    
    def load_excel_data(self, file_path):
        """Load data from Excel file into table."""
        try:
            # Load Excel file
            self.df = pd.read_excel(file_path)
            
            # Update status
            self.status_label.setText(f"Đã tải {len(self.df)} dòng dữ liệu từ {os.path.basename(file_path)}")
            
            # Populate table
            self.populate_table(self.df)
            
            # Update column filter options
            self.update_column_filter_options()
            
        except Exception as e:
            self.status_label.setText(f"Lỗi khi đọc file Excel: {str(e)}")
    
    def populate_table(self, dataframe):
        """Populate table with dataframe data."""
        if dataframe is None or dataframe.empty:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return
        
        # Clear existing data
        self.data_table.setRowCount(0)
        
        # Set column headers
        headers = dataframe.columns.tolist()
        self.data_table.setColumnCount(len(headers))
        self.data_table.setHorizontalHeaderLabels(headers)
        
        # Add data rows
        for i, row in dataframe.iterrows():
            row_position = self.data_table.rowCount()
            self.data_table.insertRow(row_position)
            
            for j, header in enumerate(headers):
                item = QTableWidgetItem(str(row[header]) if pd.notna(row[header]) else "")
                self.data_table.setItem(row_position, j, item)
        
        # Resize columns to content
        self.data_table.resizeColumnsToContents()
    
    def update_column_filter_options(self):
        """Update column filter dropdown with available columns."""
        if self.df is None or self.df.empty:
            return
        
        # Save current selection
        current_text = self.column_filter_combo.currentText()
        
        # Clear and update options
        self.column_filter_combo.clear()
        self.column_filter_combo.addItem("-- Chọn cột --")
        
        for column in self.df.columns:
            self.column_filter_combo.addItem(str(column))
        
        # Restore selection if possible
        index = self.column_filter_combo.findText(current_text)
        if index >= 0:
            self.column_filter_combo.setCurrentIndex(index)
    
    def update_column_filter(self):
        """Update column filter when selection changes."""
        self.apply_filters()
    
    def apply_filters(self):
        """Apply all filters to the data."""
        if self.df is None or self.df.empty:
            return
        
        try:
            # Make a copy of the original dataframe
            filtered_df = self.df.copy()
            
            # Get filter values
            global_search = self.global_search_input.text().strip()
            selected_column = self.column_filter_combo.currentText()
            column_value = self.column_value_input.text().strip()
            case_sensitive = self.case_sensitive_check.isChecked()
            exact_match = self.exact_match_check.isChecked()
            
            # Apply global search if provided
            if global_search:
                mask = pd.Series(False, index=filtered_df.index)
                
                for column in filtered_df.columns:
                    # Convert column to string for searching
                    col_str = filtered_df[column].astype(str)
                    
                    if exact_match:
                        if case_sensitive:
                            col_mask = col_str == global_search
                        else:
                            col_mask = col_str.str.lower() == global_search.lower()
                    else:
                        if case_sensitive:
                            col_mask = col_str.str.contains(global_search, regex=False, na=False)
                        else:
                            col_mask = col_str.str.contains(global_search, regex=False, case=False, na=False)
                    
                    mask = mask | col_mask
                
                filtered_df = filtered_df[mask]
            
            # Apply column filter if selected
            if selected_column != "-- Chọn cột --" and column_value:
                # Convert column to string for searching
                col_str = filtered_df[selected_column].astype(str)
                
                if exact_match:
                    if case_sensitive:
                        mask = col_str == column_value
                    else:
                        mask = col_str.str.lower() == column_value.lower()
                else:
                    if case_sensitive:
                        mask = col_str.str.contains(column_value, regex=False, na=False)
                    else:
                        mask = col_str.str.contains(column_value, regex=False, case=False, na=False)
                
                filtered_df = filtered_df[mask]
            
            # Update table with filtered data
            self.populate_table(filtered_df)
            
            # Update status
            self.status_label.setText(f"Hiển thị {len(filtered_df)} / {len(self.df)} dòng dữ liệu")
            
        except Exception as e:
            self.status_label.setText(f"Lỗi khi áp dụng bộ lọc: {str(e)}")
    
    def reset_filters(self):
        """Reset all filters."""
        # Clear filter inputs
        self.global_search_input.clear()
        self.column_filter_combo.setCurrentIndex(0)
        self.column_value_input.clear()
        self.case_sensitive_check.setChecked(False)
        self.exact_match_check.setChecked(False)
        
        # Reset table to show all data
        if self.df is not None:
            self.populate_table(self.df)
            self.status_label.setText(f"Hiển thị tất cả {len(self.df)} dòng dữ liệu")
    
    def open_source_file(self):
        """Open the source Excel file for the task."""
        if not self.task or not self.task.excel_path:
            QMessageBox.warning(self, "Cảnh báo", "Không có file Excel cho nhiệm vụ này")
            return
        
        if not os.path.exists(self.task.excel_path):
            QMessageBox.warning(
                self, "Cảnh báo", 
                f"Không tìm thấy file Excel tại đường dẫn:\n{self.task.excel_path}"
            )
            return
        
        try:
            # Open the file with the default application
            if os.name == 'nt':  # Windows
                os.startfile(self.task.excel_path)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.call(('open' if sys.platform == 'darwin' else 'xdg-open', self.task.excel_path))
            
            QMessageBox.information(
                self, "Thành công", 
                f"Đã mở file Excel:\n{os.path.basename(self.task.excel_path)}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở file Excel: {str(e)}")
    
    def show_context_menu(self, position):
        """Show context menu for data table."""
        if self.data_table.rowCount() == 0:
            return
        
        # Get the row under the cursor
        row = self.data_table.rowAt(position.y())
        if row < 0:
            return
        
        # Create context menu
        context_menu = QMenu(self)
        
        # Add actions
        edit_action = context_menu.addAction("Sửa dòng")
        edit_action.triggered.connect(lambda: self.edit_record(row))
        
        delete_action = context_menu.addAction("Xóa dòng")
        delete_action.triggered.connect(lambda: self.delete_record(row))
        
        # Show the menu
        context_menu.exec_(self.data_table.mapToGlobal(position))
    
    def edit_record(self, row):
        """Edit a record in the data table and sync back to Excel."""
        if self.df is None or row >= len(self.df):
            return
        
        # Get the row data
        row_data = self.df.iloc[row].to_dict()
        
        # Create edit dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Sửa dữ liệu")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(dialog)
        
        # Create form for each field
        form_layout = QFormLayout()
        fields = {}
        
        for header, value in row_data.items():
            field = QLineEdit()
            field.setText(str(value) if pd.notna(value) else "")
            form_layout.addRow(str(header) + ":", field)
            fields[header] = field
        
        layout.addLayout(form_layout)
        
        # Add buttons
        button_layout = QHBoxLayout()
        save_button = QPushButton("Lưu")
        save_button.clicked.connect(lambda: dialog.accept())
        cancel_button = QPushButton("Hủy")
        cancel_button.clicked.connect(lambda: dialog.reject())
        
        button_layout.addWidget(save_button)
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
