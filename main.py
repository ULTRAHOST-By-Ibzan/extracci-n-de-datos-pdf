import sys
import os
import sqlite3
import fitz
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QProgressBar, QVBoxLayout, \
    QWidget
from PyQt6.QtCore import Qt, QThread, pyqtSignal


#Procesar archivos PDF
class PDFProcessor(QThread):
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal()

    def __init__(self, folder):
        super().__init__()
        self.folder = folder
        self.db_name = "data.db"

    def extract_data(self, pdf_path):
        try:
            with fitz.open(pdf_path) as doc:
                text = "\n".join([page.get_text() for page in doc])
                nit = self.find_value(text, "Nit:")
                name = self.find_value(text, "nombre:")
                total = self.find_value(text, "total:")
                return nit, name, total
        except Exception as e:
            print(f"Error en {pdf_path}: {e}")
            return None, None, None

    def find_value(self, text, keyword):
        for line in text.split("\n"):
            if line.lower().startswith(keyword.lower()):
                return line.split(":")[1].strip()
        return ""

    def run(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pdf_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nit TEXT UNIQUE,
                name TEXT,
                total TEXT
            )
        """)
        conn.commit()

        pdf_files = [f for f in os.listdir(self.folder) if f.endswith(".pdf")]
        total_files = len(pdf_files)

        for index, pdf_file in enumerate(pdf_files, 1):
            nit, name, total = self.extract_data(os.path.join(self.folder, pdf_file))
            if nit and name and total:
                cursor.execute("SELECT COUNT(*) FROM pdf_data WHERE nit = ?", (nit,))
                if cursor.fetchone()[0] == 0:  # Solo inserta si el NIT no existe
                    cursor.execute("INSERT INTO pdf_data (nit, name, total) VALUES (?, ?, ?)", (nit, name, total))
                    conn.commit()
            self.progress_signal.emit(index, total_files)

        conn.close()
        self.finished_signal.emit()


# Interfaz gráfica
class PDFAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Analizador de PDFs")
        self.setGeometry(100, 100, 400, 200)

        self.folder_path = ""
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Selecciona una carpeta")
        layout.addWidget(self.label)

        self.select_button = QPushButton("Seleccionar Carpeta")
        self.select_button.clicked.connect(self.select_folder)
        layout.addWidget(self.select_button)

        self.analyze_button = QPushButton("Analizar PDFs")
        self.analyze_button.setEnabled(False)
        self.analyze_button.clicked.connect(self.start_analysis)
        layout.addWidget(self.analyze_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("Progreso: 0/0")
        layout.addWidget(self.progress_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        if folder:
            self.folder_path = folder
            total_files = len([f for f in os.listdir(folder) if f.endswith(".pdf")])
            self.label.setText(f"Archivos encontrados: {total_files}")
            self.analyze_button.setEnabled(True)

    def start_analysis(self):
        self.progress_bar.setValue(0)
        self.progress_label.setText("Progreso: 0/0")

        self.processor = PDFProcessor(self.folder_path)
        self.processor.progress_signal.connect(self.update_progress)
        self.processor.finished_signal.connect(self.export_to_excel)
        self.processor.start()

    def update_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.progress_label.setText(f"Progreso: {current}/{total}")

    def export_to_excel(self):
        conn = sqlite3.connect("data.db")
        df = pd.read_sql_query("SELECT nit, name, total FROM pdf_data", conn)
        conn.close()
        df.to_excel("datos.xlsx", index=False)
        self.label.setText("Proceso finalizado. Archivo Excel generado.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PDFAnalyzerApp()
    window.show()
    sys.exit(app.exec())
