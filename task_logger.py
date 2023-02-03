import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit,QGridLayout, QPushButton, QMessageBox,QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem
from PyQt5.QtGui import QFont
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, pyqtSlot
from openpyxl import Workbook, load_workbook
import subprocess
import os


class TaskLogger(QWidget):
    def __init__(self):
        super().__init__()
        screen_resolution = app.desktop().screenGeometry()
        width, height = screen_resolution.width(), screen_resolution.height()
        wid_x = int(width/2)
        hei_y = int(600)
        x_cor = int(width/2) - int(wid_x/2)
        y_cor = int(height/2) - int(hei_y/2)
        self.setGeometry(x_cor,y_cor,wid_x,hei_y)
        self.setFixedSize(wid_x,hei_y)
        self.initUI()

    def initUI(self):

        # WINDOW TITLE AND INSTRUCTION
        self.title = QLabel("Task Logger")
        self.title.setObjectName("title")

        # LABEL DEFINITIONS
        self.label1 = QLabel("Scheme")
        self.label2 = QLabel("Task type")
        self.label3 = QLabel("Date assigned")
        self.label4 = QLabel("Checker")
        self.label5 = QLabel("Date Completed")

        # ENTRY BOX DEFINITIONS
        self.entry1 = QLineEdit()
        self.entry2 = QLineEdit()
        self.entry3 = QLineEdit()
        self.entry4 = QLineEdit()
        self.entry5 = QLineEdit()


        # BUTTON DEFINITIONS WITH FUNCTIONS ATTACHED  << ----------------- EDIT TO CONNECT FUNCTIONS TO BUTTONS
        self.button1 = QPushButton("Save")
        self.button1.setObjectName("button1")
        self.button1.clicked.connect(self.save_data)
        self.button2 = QPushButton("Open")
        self.button2.setObjectName("button2")
        self.button2.clicked.connect(self.open_excel_logger)


        grid = QGridLayout()
        grid.setSpacing(20)


        grid.addWidget(self.title, 0, 0, 2,2)

        grid.addWidget(self.label1,2,0)
        grid.addWidget(self.entry1,2,1)

        grid.addWidget(self.label2,3,0)
        grid.addWidget(self.entry2,3,1)

        grid.addWidget(self.label3,4,0)
        grid.addWidget(self.entry3,4,1)

        grid.addWidget(self.label4,5,0)
        grid.addWidget(self.entry4,5,1)

        grid.addWidget(self.label5,6,0)
        grid.addWidget(self.entry5,6,1)

        grid.addWidget(self.button1, 7, 0, 1, 1, Qt.AlignRight)
        grid.addWidget(self.button2, 7, 1, 1, 1, Qt.AlignLeft)


        self.setLayout(grid)
        self.setWindowTitle("Derrick's Task Logger")

         # Style the GUI using a stylesheet
        self.setStyleSheet("""
        QLabel {
            font-size: 20px;
            font-weight: bold;
            margin-bottom: 10px;
        }

        QLabel#title {
            font-size: 40px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        QLineEdit {
            font-size: 24px;
            padding: 5px;
            border: 1px solid gray;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        QPushButton {
            font-size: 24px;
            padding: 10px;
            background-color: green;
            color: white;
            border: none;
            border-radius: 5px;
        }
        """)

    def open_file(self):
        # Open the "Task Logger.xlsx" file using openpyxl
        wb = load_workbook("real project\Task Logger.xlsx")
        
        # Create a new widget to display the file contents
        file_widget = QtWidgets.QTextEdit()
        file_widget.setReadOnly(True)
        
        # Iterate through each sheet in the workbook and add its contents to the text edit
        for sheet in wb:
            file_widget.append(sheet.title)
            for row in sheet.rows:
                row_data = [cell.value for cell in row]
                file_widget.append(str(row_data))
        
        # Show the file widget as a new window
        file_window = QtWidgets.QMainWindow()
        file_window.setCentralWidget(file_widget)
        file_window.show()


    def save_data(self):
        scheme = self.entry1.text()
        task_type = self.entry2.text()
        date_assigned = self.entry3.text()
        checker = self.entry4.text()
        date_completed =  self.entry5.text()

        
        logger = "Task Logger.xlsx"

        if os.path.exists(logger):
            df = pd.read_excel(logger)
        else:
            df = pd.DataFrame()
            
        df = pd.concat([df,pd.DataFrame({'Scheme':[scheme], 'Task type': [task_type], 'Date assigned': [date_assigned], 'Checker': [checker], 'Date completed':[date_completed]})])
        df.to_excel(logger,index=False)

        self.entry1.clear()
        self.entry2.clear()
        self.entry3.clear()
        self.entry4.clear()
        self.entry5.clear()

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Tasked logged successfully!")
        msg.setWindowTitle("Success")
        msg.exec_()

    def open_excel_logger(file_path):
        logger = "Task Logger.xlsx"
        if os.path.exists(logger):
            subprocess.Popen(logger,shell=True)
        else:
            df = pd.DataFrame({'Scheme', 'Task type', 'Date assigned', 'Checker', 'Date completed'})
            df.to_excel(logger,index=False)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    run = TaskLogger()
    run.show()

    sys.exit(app.exec_())

