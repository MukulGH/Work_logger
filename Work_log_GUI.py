import sys
import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QVBoxLayout, QLabel
from worklog_interaction import get_latest_excel_file, update_worklog
from worklog_interaction import get_start_time_from_sheet

import subprocess
import os

class MyWindow(QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()

        # Setup UI components
        self.button_create = QPushButton('Create Work Logbook for this week', self)
        self.button_read = QPushButton('Read Start Time', self)
        self.label_previous_start_time = QLabel('Suggested next Start Time:', self)
        self.output_previous_start_time = QLineEdit(self)
        self.textbox_start_time = QLineEdit(self)
        self.textbox_hours = QLineEdit(self)
        self.textbox_task = QLineEdit(self)
        self.button_submit = QPushButton('Submit', self)

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.button_create.clicked.connect(self.run_module)
        self.button_read.clicked.connect(self.read_start_time)
        self.button_submit.clicked.connect(self.submit_info)

        self.textbox_start_time.setPlaceholderText('Start time...')
        self.output_previous_start_time.setPlaceholderText('Previous start time...')
        self.textbox_hours.setPlaceholderText('Number of hours...')
        self.textbox_task.setPlaceholderText('Task...')
        self.output_previous_start_time.setReadOnly(True)

        self.setStyleSheet("QWidget { background-color: #799B9E; }"
                           "QPushButton { background-color: #9E6647; color: white; font-size: 27px; min-height: 45px; min-width: 150px;}"
                           "QLineEdit { background-color: lightgray; font-size: 27px; min-height: 30px; }"
                           "QLabel { color: white; font-size: 27px; }")

        layout.addWidget(self.button_create)
        layout.addWidget(self.button_read)
        layout.addWidget(self.label_previous_start_time)
        layout.addWidget(self.output_previous_start_time)
        layout.addWidget(self.textbox_start_time)
        layout.addWidget(self.textbox_hours)
        layout.addWidget(self.textbox_task)
        layout.addWidget(self.button_submit)

        self.setLayout(layout)
        self.setWindowTitle('My App')
        self.resize(600, 400)
        self.show()

    def run_module(self):
        subprocess.run(['python', 'worklog_creation.py'], check=True)

    def read_start_time(self):
        process = subprocess.Popen(['python', 'worklog_interaction.py'], stdout=subprocess.PIPE)
        stdout, stderr = process.communicate()
        previous_start_time = stdout.decode('utf-8').strip()
        self.output_previous_start_time.setText(previous_start_time)

    def submit_info(self):
        start_time = self.textbox_start_time.text()
        hours = self.textbox_hours.text()
        task = self.textbox_task.text()
        
        directory = os.getcwd()
        latest_excel_file = get_latest_excel_file(directory)
        
        if latest_excel_file:
            today = datetime.date.today()
            sheet_name = today.strftime("%m_%d_%Y")
            update_worklog(latest_excel_file, sheet_name, start_time, task, hours)
            
            # Get the new start time after submitting the information
            new_start_time = get_start_time_from_sheet(latest_excel_file, sheet_name)
            print (new_start_time)
            
            # Update the "Previous Start Time" field in the GUI
            self.output_previous_start_time.setText(new_start_time)
            
            print("Information submitted.")
        else:
            print("No Excel files found in the directory.")



# Create app
app = QApplication(sys.argv)
window = MyWindow()

# Enter application main loop
sys.exit(app.exec_())
