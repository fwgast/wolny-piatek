# -*- coding: utf-8 -*-
from config import SUBJECT, SENDER, BODY_SENSITIVE_DATA, FOOTER, PA_LINK
import csv
import sys, os, subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QGridLayout, QLabel, QPushButton, QWidget, QFileDialog, QMessageBox, QVBoxLayout
from PyQt5.QtGui import QIcon, QMovie
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from pandas import read_excel
import openpyxl

def get_application_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))
    
def set_environment_variable(name, value):
    command = f'setx {name} "{value}"'
    null_path = os.devnull
    subprocess.call(command, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loading...")
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowIcon(QIcon(resource_path('icon.ico')))
        self.loading_label = QLabel(self)
        self.movie = QMovie(resource_path('dance.gif'))

        self.loading_label.setMovie(self.movie)

        self.movie.start()

        layout = QVBoxLayout()
        layout.addWidget(self.loading_label)
        self.setLayout(layout)

        self.resize(self.movie.frameRect().size())


    def set_progress(self, value):
        self.setWindowTitle(f"Processing... ({value}%)")
        QApplication.processEvents()

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, file_path, checklist_path, medical_path, fitness_path):
        super().__init__()
        self.file_path = file_path
        self.checklist_path = checklist_path
        self.medical_path = medical_path
        self.fitness_path = fitness_path

    def run(self):
        def check_data_in_cell(data):
            return '1' if data.upper() in ['Y', 'T'] else '0'

        def compose_message(data, mail):
            columns_count = len(data)
            if data == '00111111111111' or data == '000111111111111':
                return None

            missing_documents = []
            attachments = []

            if data[columns_count - 12] == '0':
                missing_documents.append("wypelniona, podpisana i zeskanowana checklista (szablon w zalaczniku)")
                attachments.append({'path': self.checklist_path, 'name': 'CHECKLIST.pdf'})
            if data[columns_count - 11] == '0':
                missing_documents.append("Zaswiadczenie z ZUS o niezaleganiu (w przypadku braku - potwierdzenie zlozenia wniosku)")
            if data[columns_count - 10] == '0':
                missing_documents.append("Zaswiadczenie z US o niezaleganiu (w przypadku braku - potwierdzenie zlozenia wniosku)")
            if data[columns_count - 9] == '0':
                missing_documents.append("Zaswiadczenie o niekaralnosci")
            if data[columns_count - 8] == '0':
                missing_documents.append("Zdjecie (format paszportowy lub legitymacyjny)")
            if data[columns_count - 7] == '0':
                missing_documents.append("Dowod osobisty")
            if data[columns_count - 6] == '0':
                missing_documents.append("Prawo jazdy")
            if data[columns_count - 5] == '0':
                missing_documents.append("Medical Certificate (szablon w zalaczniku)")
                attachments.append({'path': self.medical_path, 'name': 'MEDICAL.pdf'})
            if data[columns_count - 4] == '0':
                missing_documents.append("Medical Fitness Certificate (szablon w zalaczniku)")
                attachments.append({'path': self.fitness_path, 'name': 'FITNESS.pdf'})
            if data[columns_count - 3] == '0':
                missing_documents.append("Polisa OC (musi posiadac bezwglednie wszystkie 6 kodow PKD*)")
            if data[columns_count - 2] == '0':
                missing_documents.append("Dyplom/swiadectwo ukonczenia szkoly")
            if data[columns_count - 1] == '0':
                missing_documents.append("Paszport - termin waznosci min. 6 miesiecy (w przypadku braku - potwierdzenie zlozenia wniosku)")

            body = f"""
                    <html>
                    <body>
                        <p>Hej,</p>
                        <p>W Twoim profilu brakuje nam nastepujacych dokumentow:</p>
                        <ul>
                            {''.join([f'<li>{doc}</li>' for doc in missing_documents])}
                        </ul>
                            {BODY_SENSITIVE_DATA}
                    </body>
                    </html>
                    """
            body+=FOOTER

            try:
                attachment1 = attachments[0]['path']
            except IndexError:
                attachment1 = None

            try:
                attachment2 = attachments[1]['path']
            except IndexError:
                attachment2 = None

            try:
                attachment3 = attachments[2]['path']
            except IndexError:
                attachment3 = None

            return body, SUBJECT, SENDER, mail, attachment1, attachment2, attachment3

        columns_in_csv = ['email', 'subject', 'sender', 'body', 'attachment1', 'attachment2', 'attachment3']

        with open('missing_documents.csv', mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file, delimiter='`')
            writer.writerow(columns_in_csv)

            df = read_excel(self.file_path, sheet_name=0, header=0)
            columns = list(df.columns.values)
            total_rows = len(df)

            for index, row in df.iterrows():
                data = ''.join([check_data_in_cell(str(row[f"{col}"])) for col in columns])
                mail = row['email']
                message = compose_message(data, mail)

                if message is not None:
                    body, subject, sender, email, attachment1, attachment2, attachment3 = message
                    writer.writerow([email, subject, sender, body, attachment1, attachment2, attachment3])

                progress_value = int((index + 1) / total_rows * 100)
                self.progress.emit(progress_value)
            
            self.progress.emit(100)
            self.finished.emit()

class EmailAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wolny Piątek")
        self.setWindowIcon(QIcon(resource_path('icon.ico')))
        self.setGeometry(100, 100, 600, 400)

        self.file_path = ""
        self.checklist_path = ""
        self.medical_path = ""
        self.fitness_path = ""

        self.initUI()

    def initUI(self):
        layout = QGridLayout()

        self.file_label = QLabel("Excel File: Not selected")
        self.file_button = QPushButton("Select Excel File")
        self.file_button.clicked.connect(self.select_file_path)
        layout.addWidget(self.file_label, 0, 0)
        layout.addWidget(self.file_button, 0, 1)

        self.checklist_label = QLabel("Checklist File: Not selected")
        self.checklist_button = QPushButton("Select Checklist File")
        self.checklist_button.clicked.connect(self.select_checklist_path)
        layout.addWidget(self.checklist_label, 1, 0)
        layout.addWidget(self.checklist_button, 1, 1)

        self.medical_label = QLabel("Medical File: Not selected")
        self.medical_button = QPushButton("Select Medical File")
        self.medical_button.clicked.connect(self.select_medical_path)
        layout.addWidget(self.medical_label, 2, 0)
        layout.addWidget(self.medical_button, 2, 1)

        self.fitness_label = QLabel("Fitness File: Not selected")
        self.fitness_button = QPushButton("Select Fitness File")
        self.fitness_button.clicked.connect(self.select_fitness_path)
        layout.addWidget(self.fitness_label, 3, 0)
        layout.addWidget(self.fitness_button, 3, 1)

        self.send_mail_button = QPushButton("Send mails")
        self.send_mail_button.clicked.connect(self.create_csv)
        layout.addWidget(self.send_mail_button, 4, 0, 1, 2)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file_path(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file_path:
            self.file_path = file_path
            self.file_label.setText(f"Excel File: {file_path}")

    def select_checklist_path(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Checklist File", "", "PDF Files (*.pdf)", options=options)
        if file_path:
            self.checklist_path = file_path
            self.checklist_label.setText(f"Checklist File: {file_path}")

    def select_medical_path(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Medical File", "", "PDF Files (*.pdf)", options=options)
        if file_path:
            self.medical_path = file_path
            self.medical_label.setText(f"Medical File: {file_path}")

    def select_fitness_path(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Fitness File", "", "PDF Files (*.pdf)", options=options)
        if file_path:
            self.fitness_path = file_path
            self.fitness_label.setText(f"Fitness File: {file_path}")

    def validate_paths(self):
        if not (self.file_path and self.checklist_path and self.medical_path and self.fitness_path):
            QMessageBox.warning(self, "Error", "Please select all required files and directories.")
            return False
        return True

    def create_csv(self):
        if not self.validate_paths():
            return
        
        self.loading_screen = LoadingScreen()
        self.loading_screen.show()
        self.worker = Worker(self.file_path, self.checklist_path, self.medical_path, self.fitness_path)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()


    def update_progress(self, value):
        value = max(0, min(100, value))
        self.loading_screen.set_progress(value)

    def on_finished(self):
        self.loading_screen.close()
        QMessageBox.information(self, "Completed", "Emails processed successfully!")
        set_environment_variable('wolny_piatek_csv_file', os.path.join(get_application_path(), 'missing_documents.csv'))
        #webbrowser.open(PA_LINK)           #only on premium -_-

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EmailAutomationApp()
    ex.show()
    sys.exit(app.exec_())
