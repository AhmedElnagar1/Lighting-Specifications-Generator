import sys
import os
from typing import Optional
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
    QWidget, QPushButton, QLabel, QFileDialog, QComboBox,
    QProgressBar, QTextEdit, QMessageBox, QGroupBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
from final_excel_processor import process_excel_file


class ProcessingThread(QThread):
    """Thread for processing Excel files to prevent GUI freezing"""
    
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, str)  # success, message, pdf_path
    
    def __init__(self, excel_file_path: str, language: str, img_dir: str):
        """
        Initialize the processing thread
        
        Args:
            excel_file_path (str): Path to the Excel file to process
            language (str): Language code ('EN' or 'DE')
            img_dir (str): Path to the image directory
        """
        super().__init__()
        self.excel_file_path = excel_file_path
        self.language = language
        self.img_dir = img_dir
    
    def run(self) -> None:
        """Run the Excel processing in a separate thread"""
        try:
            self.progress_signal.emit("Starting Excel processing...")
            result = process_excel_file(self.excel_file_path, self.language, self.img_dir)
            
            if result and isinstance(result, str):
                # Success - result is the PDF path
                self.finished_signal.emit(True, "Processing completed successfully!", result)
            else:
                # Failure - result is False
                self.finished_signal.emit(False, "Processing failed. Check the console for details.", "")
                
        except Exception as e:
            self.finished_signal.emit(False, f"Error during processing: {str(e)}", "")


class ExcelProcessorApp(QMainWindow):
    """Main application window for Excel file processing"""
    
    def __init__(self):
        """Initialize the main application window"""
        super().__init__()
        self.selected_file_path: Optional[str] = None
        self.selected_img_dir: Optional[str] = None
        self.processing_thread: Optional[ProcessingThread] = None
        self.pdf_path: Optional[str] = None
        self.init_ui()
    
    def init_ui(self) -> None:
        """Initialize the user interface"""
        self.setWindowTitle("Lighting Specifications Generator")
        self.setGeometry(100, 100, 600, 750)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Title
        title_label = QLabel("Lighting Specifications Generator")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
        
        # File path display
        self.file_path_label = QLabel("No file selected")
        self.file_path_label.setStyleSheet("QLabel { padding: 8px; border: 1px solid #ccc; border-radius: 5px; }")
        file_layout.addWidget(self.file_path_label)
        
        # Browse button
        browse_button = QPushButton("Browse for Excel File")
        browse_button.clicked.connect(self.browse_file)
        browse_button.setStyleSheet("""
            QPushButton {
                padding: 10px;
                font-size: 14px;
                border-radius: 5px;
            }
        """)
        file_layout.addWidget(browse_button)
        
        main_layout.addWidget(file_group)
        
        # Image directory selection group
        img_group = QGroupBox("Image Directory")
        img_layout = QVBoxLayout(img_group)
        
        # Image directory path display
        self.img_path_label = QLabel("No image directory selected")
        self.img_path_label.setStyleSheet("QLabel { padding: 8px; border: 1px solid #ccc; border-radius: 5px; }")
        img_layout.addWidget(self.img_path_label)
        
        # Browse button for image directory
        browse_img_button = QPushButton("Browse for Image Directory")
        browse_img_button.clicked.connect(self.browse_image_directory)
        browse_img_button.setStyleSheet("""
            QPushButton {
                padding: 10px;
                font-size: 14px;
                border-radius: 5px;
            }
        """)
        img_layout.addWidget(browse_img_button)
        
        main_layout.addWidget(img_group)
        
        # Language selection group
        language_group = QGroupBox("Language Selection")
        language_layout = QVBoxLayout(language_group)
        
        self.language_combo = QComboBox()
        self.language_combo.addItem("English", "EN")
        self.language_combo.addItem("German", "DE")
        self.language_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        language_layout.addWidget(self.language_combo)
        
        main_layout.addWidget(language_group)
        
        # Process button
        self.process_button = QPushButton("Process Excel File")
        self.process_button.clicked.connect(self.process_file)
        self.process_button.setEnabled(False)
        self.process_button.setStyleSheet("""
            QPushButton {
                padding: 15px;
                font-size: 16px;
                font-weight: bold;
                border-radius: 5px;
            }
        """)
        main_layout.addWidget(self.process_button)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                border-radius: 5px;
            }
        """)
        main_layout.addWidget(self.progress_bar)
        
        # Status text area
        self.status_text = QTextEdit()
        self.status_text.setMaximumHeight(150)
        self.status_text.setReadOnly(True)
        self.status_text.setStyleSheet("""
            QTextEdit {
                border-radius: 5px;
                font-family: 'Courier New', monospace;
                font-size: 12px;
            }
        """)
        main_layout.addWidget(self.status_text)
        
        # PDF output group
        pdf_group = QGroupBox("PDF Output")
        pdf_layout = QVBoxLayout(pdf_group)
        
        # PDF path display
        self.pdf_path_label = QLabel("No PDF generated yet")
        self.pdf_path_label.setStyleSheet("""
            QLabel {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 5px;
            }
        """)
        self.pdf_path_label.setWordWrap(True)
        pdf_layout.addWidget(self.pdf_path_label)
        
        # Open PDF button
        self.open_pdf_button = QPushButton("Open PDF")
        self.open_pdf_button.clicked.connect(self.open_pdf)
        self.open_pdf_button.setEnabled(False)
        self.open_pdf_button.setStyleSheet("""
            QPushButton {
                padding: 10px;
                font-size: 14px;
                border-radius: 5px;
            }
        """)
        pdf_layout.addWidget(self.open_pdf_button)
        
        main_layout.addWidget(pdf_group)
        
        # Set window style
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
                color: black;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
                background-color: white;
                color: black;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                background-color: white;
                color: black;
            }
            QLabel {
                background-color: white;
                color: black;
            }
            QPushButton {
                background-color: white;
                color: black;
                border: 1px solid #cccccc;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
            }
            QPushButton:disabled {
                background-color: #f5f5f5;
                color: #666666;
            }
            QComboBox {
                background-color: white;
                color: black;
                border: 1px solid #cccccc;
            }
            QTextEdit {
                background-color: white;
                color: black;
                border: 1px solid #cccccc;
            }
            QProgressBar {
                background-color: white;
                color: black;
                border: 1px solid #cccccc;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
            }
        """)
    
    def browse_file(self) -> None:
        """Open file dialog to select Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*)"
        )
        
        if file_path:
            self.selected_file_path = file_path
            self.file_path_label.setText(f"Selected: {os.path.basename(file_path)}")
            self.update_process_button_state()
            self.log_message(f"File selected: {file_path}")
    
    def browse_image_directory(self) -> None:
        """Open directory dialog to select image directory"""
        img_dir = QFileDialog.getExistingDirectory(
            self,
            "Select Image Directory",
            ""
        )
        
        if img_dir:
            self.selected_img_dir = img_dir
            self.img_path_label.setText(f"Selected: {os.path.basename(img_dir)}")
            self.update_process_button_state()
            self.log_message(f"Image directory selected: {img_dir}")
    
    def update_process_button_state(self) -> None:
        """Update the process button enabled state based on selections"""
        self.process_button.setEnabled(
            self.selected_file_path is not None and 
            self.selected_img_dir is not None
        )
    
    def process_file(self) -> None:
        """Start processing the selected Excel file"""
        if not self.selected_file_path:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
            return
        
        if not os.path.exists(self.selected_file_path):
            QMessageBox.critical(self, "Error", "Selected file does not exist.")
            return
        
        if not self.selected_img_dir:
            QMessageBox.warning(self, "Warning", "Please select an image directory first.")
            return
        
        if not os.path.exists(self.selected_img_dir):
            QMessageBox.critical(self, "Error", "Selected image directory does not exist.")
            return
        
        # Get selected language
        language = self.language_combo.currentData()
        
        # Disable UI elements during processing
        self.process_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        
        # Clear status text
        self.status_text.clear()
        self.log_message(f"Starting processing with language: {language}")
        
        # Create and start processing thread
        self.processing_thread = ProcessingThread(self.selected_file_path, language, self.selected_img_dir)
        self.processing_thread.progress_signal.connect(self.log_message)
        self.processing_thread.finished_signal.connect(self.on_processing_finished)
        self.processing_thread.start()
    
    def on_processing_finished(self, success: bool, message: str, pdf_path: str) -> None:
        """Handle processing completion"""
        # Re-enable UI elements
        self.process_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # Log final message
        self.log_message(message)
        
        # Update PDF output section
        if success and pdf_path:
            self.pdf_path_label.setText(f"PDF generated: {pdf_path}")
            self.pdf_path_label.setStyleSheet("""
                QLabel {
                    background-color: #e8f5e8;
                    padding: 8px;
                    border: 1px solid #4CAF50;
                    border-radius: 5px;
                    color: #2e7d32;
                }
            """)
            self.open_pdf_button.setEnabled(True)
            self.pdf_path = pdf_path
        else:
            self.pdf_path_label.setText("No PDF generated")
            self.pdf_path_label.setStyleSheet("""
                QLabel {
                    padding: 8px;
                    border: 1px solid #ccc;
                    border-radius: 5px;
                }
            """)
            self.open_pdf_button.setEnabled(False)
            self.pdf_path = None
        
        # Show result dialog
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
    
    def log_message(self, message: str) -> None:
        """Add message to status text area"""
        self.status_text.append(f"[{self.get_current_time()}] {message}")
        # Auto-scroll to bottom
        self.status_text.verticalScrollBar().setValue(
            self.status_text.verticalScrollBar().maximum()
        )
    
    def get_current_time(self) -> str:
        """Get current time as formatted string"""
        from datetime import datetime
        return datetime.now().strftime("%H:%M:%S")
    
    def open_pdf(self) -> None:
        """Open the generated PDF file"""
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            QMessageBox.warning(self, "Warning", "PDF file not found or no PDF generated yet.")
            return
        
        try:
            # Use the default system application to open the PDF
            import subprocess
            import platform
            
            system = platform.system()
            if system == "Windows":
                os.startfile(self.pdf_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", self.pdf_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", self.pdf_path], check=True)
                
            self.log_message(f"Opened PDF: {self.pdf_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open PDF: {str(e)}")
            self.log_message(f"Error opening PDF: {str(e)}")
    
    def closeEvent(self, event) -> None:
        """Handle application close event"""
        if self.processing_thread and self.processing_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirm Exit",
                "Processing is still running. Are you sure you want to exit?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.processing_thread.terminate()
                self.processing_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main() -> None:
    """Main function to run the application"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Lighting Specifications Generator")
    app.setApplicationVersion("1.0.0")
    
    # Create and show main window
    window = ExcelProcessorApp()
    window.show()
    
    # Start event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
