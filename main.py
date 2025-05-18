import sys
from PySide6.QtWidgets import QApplication
from ui.main_window import MainWindow
from database.db_manager import init_db

def main():
    """Main entry point for the application."""
    # Initialize the database
    init_db()
    
    # Create the application
    app = QApplication(sys.argv)
    
    # Create and show the main window
    window = MainWindow()
    window.show()
    
    # Run the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
