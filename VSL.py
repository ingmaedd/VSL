# -*- coding: utf-8 -*-
"""
@author: Eddie
"""

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget, QMessageBox
import pandas as pd
from datetime import datetime

class Vehicle:
    def __init__(self, name, model, year):
        self.name = name
        self.model = model
        self.year = year
        self.service_data = []

    def log_service(self, odometer, service_done, service_date):
        self.service_data.append({
            'Odometer': odometer,
            'Service Done': service_done,
            'Service Date': service_date
        })

    def save_to_excel(self, file_name):
        df = pd.DataFrame(self.service_data)
        df.to_excel(file_name, index=False)

class VehicleServiceLogger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Vehicle Service Logger")
        self.setGeometry(100, 100, 300, 250)

        self.vehicle = None

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.name_label = QLabel("Vehicle Brand:")
        self.name_input = QLineEdit()
        self.model_label = QLabel("Vehicle Model Name:")
        self.model_input = QLineEdit()
        self.year_label = QLabel("Vehicle Model Year:")
        self.year_input = QLineEdit()

        self.odometer_label = QLabel("Odometer Reading:")
        self.odometer_input = QLineEdit()
        self.service_label = QLabel("Service Done:")
        self.service_input = QLineEdit()
        self.date_label = QLabel("Date of Service (DD-MM-YYYY):")
        self.date_input = QLineEdit()

        self.save_button = QPushButton("Save Service Data")
        self.save_button.clicked.connect(self.save_service_data)

        self.export_button = QPushButton("Export to Excel")
        self.export_button.clicked.connect(self.export_to_excel)

        layout.addWidget(self.name_label)
        layout.addWidget(self.name_input)
        layout.addWidget(self.model_label)
        layout.addWidget(self.model_input)
        layout.addWidget(self.year_label)
        layout.addWidget(self.year_input)
        layout.addWidget(self.odometer_label)
        layout.addWidget(self.odometer_input)
        layout.addWidget(self.service_label)
        layout.addWidget(self.service_input)
        layout.addWidget(self.date_label)
        layout.addWidget(self.date_input)
        layout.addWidget(self.save_button)
        layout.addWidget(self.export_button)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def save_service_data(self):
        name = self.name_input.text()
        model = self.model_input.text()
        year = self.year_input.text()
        odometer = self.odometer_input.text()
        service_done = self.service_input.text()
        service_date = self.date_input.text()

        if not self.vehicle:
            self.vehicle = Vehicle(name, model, year)

        try:
            datetime.strptime(service_date, '%d-%m-%Y')
            self.vehicle.log_service(odometer, service_done, service_date)
            QMessageBox.information(self, "Success", "Service data saved.")
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid date format. Use DD-MM-YYYY.")

    def export_to_excel(self):
        if self.vehicle:
            file_name = f"{self.vehicle.name}_{self.vehicle.model}_{self.vehicle.year}_service_log.xlsx"
            self.vehicle.save_to_excel(file_name)
            QMessageBox.information(self, "Success", f"Data exported to {file_name}.")
        else:
            QMessageBox.warning(self, "Error", "No vehicle data to export.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VehicleServiceLogger()
    window.show()
    sys.exit(app.exec_())