import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QComboBox, QSpinBox, QPushButton, QFileDialog, QLabel)
from PySide6.QtCore import Qt
from datetime import datetime
import calendar
from common import generate_schedule, coworkers


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Shift Schedule Generator")
        self.setGeometry(100, 100, 400, 200)  # Position and size

        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(15)

        # Month selection
        self.month_label = QLabel("Select Month:")
        self.month_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.month_label)
        
        self.month_combo = QComboBox()
        for m in range(1, 13):
            self.month_combo.addItem(calendar.month_name[m], m)
        self.month_combo.setStyleSheet("padding: 5px; font-size: 14px;")
        layout.addWidget(self.month_combo)

        # Year selection
        self.year_label = QLabel("Select Year:")
        self.year_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.year_label)
        
        self.year_spin = QSpinBox()
        self.year_spin.setRange(1900, 2100)
        self.year_spin.setStyleSheet("padding: 5px; font-size: 14px;")
        layout.addWidget(self.year_spin)

        # Generate button
        self.generate_button = QPushButton("Generate Schedule")
        self.generate_button.setStyleSheet("""
            QPushButton {
                padding: 10px;
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.generate_button.clicked.connect(self.generate_schedule)
        layout.addWidget(self.generate_button)

        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # Set defaults to current date
        current_date = datetime.now()
        self.month_combo.setCurrentIndex(current_date.month - 1)
        self.year_spin.setValue(current_date.year)

    def generate_schedule(self):
        month = self.month_combo.currentData()
        year = self.year_spin.value()
        self.status_label.setText("Generating schedule...")
        QApplication.processEvents()  # Update GUI

        wb = generate_schedule(year, month, coworkers)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Schedule", f"schedule_{year}_{month}.xlsx", "Excel Files (*.xlsx)"
        )
        
        if file_path:
            wb.save(file_path)
            self.status_label.setText(f"Schedule saved to {file_path}")
        else:
            self.status_label.setText("Generation cancelled.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())