import os
from sqlalchemy import Column, Integer, String, Text, Date, ForeignKey
from sqlalchemy.orm import relationship
from database.db_manager import Base

class Task(Base):
    """Task model representing a mission/task with associated Excel file."""
    __tablename__ = 'tasks'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    year = Column(Integer, nullable=False)
    unit = Column(String(255), nullable=False)
    description = Column(Text, nullable=True)
    excel_path = Column(String(512), nullable=False)
    created_at = Column(Date, nullable=False)
    
    # Relationships
    people = relationship("Person", back_populates="task", cascade="all, delete-orphan")
    
    def __repr__(self):
        return f"<Task(id={self.id}, name='{self.name}', year={self.year}, unit='{self.unit}')>"
    
    def get_folder_path(self):
        """Get the folder path for this task."""
        if not self.excel_path:
            return None
        
        # Get the directory containing the Excel file
        base_dir = os.path.dirname(self.excel_path)
        
        # Create a safe folder name from task name
        import re
        safe_name = re.sub(r'[^\w\s-]', '', self.name).strip().replace(' ', '_')
        
        # Combine with base directory
        folder_path = os.path.join(base_dir, safe_name)
        return folder_path
    
    def create_task_folder(self):
        """Create a folder for this task based on its name."""
        folder_path = self.get_folder_path()
        if not folder_path:
            return None
        
        # Create the folder if it doesn't exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        
        return folder_path
    
    def rename_task_folder(self, old_name):
        """Rename the task folder and Excel file when the task name changes."""
        if not self.excel_path:
            return False
        
        # Get the directory containing the Excel file
        base_dir = os.path.dirname(os.path.dirname(self.excel_path))  # Lấy thư mục cha của thư mục chứa file Excel
        
        # Create safe folder names
        import re
        old_safe_name = re.sub(r'[^\w\s-]', '', old_name).strip().replace(' ', '_')
        new_safe_name = re.sub(r'[^\w\s-]', '', self.name).strip().replace(' ', '_')
        
        # Old and new folder paths
        old_folder_path = os.path.join(base_dir, old_safe_name)
        new_folder_path = os.path.join(base_dir, new_safe_name)
        
        # Generate new Excel file name based on new task name
        from datetime import datetime
        current_date = datetime.now().strftime("%d%m%Y")
        new_excel_name = f"{new_safe_name}_{current_date}.xlsx"
        
        # Tạo thư mục mới nếu chưa tồn tại
        if not os.path.exists(new_folder_path):
            try:
                os.makedirs(new_folder_path)
            except Exception as e:
                print(f"Error creating new folder: {str(e)}")
                return False
        
        # Đường dẫn mới cho file Excel
        new_excel_path = os.path.join(new_folder_path, new_excel_name)
        
        # Rename the Excel file and move to new folder
        if os.path.exists(self.excel_path):
            try:
                # Copy file Excel vào thư mục mới với tên mới
                import shutil
                shutil.copy2(self.excel_path, new_excel_path)
                
                # Xóa file cũ sau khi copy thành công
                if os.path.exists(new_excel_path):
                    os.remove(self.excel_path)
                
                # Cập nhật đường dẫn mới trong đối tượng Task
                self.excel_path = new_excel_path
                print(f"Excel file moved to: {new_excel_path}")
            except Exception as e:
                print(f"Error moving Excel file: {str(e)}")
                return False
        
        # Xóa thư mục cũ nếu rỗng
        if os.path.exists(old_folder_path) and old_folder_path != new_folder_path:
            try:
                # Kiểm tra xem thư mục cũ có rỗng không
                if not os.listdir(old_folder_path):
                    os.rmdir(old_folder_path)
                    print(f"Removed empty folder: {old_folder_path}")
            except Exception as e:
                print(f"Error removing old folder: {str(e)}")
        
        return True
