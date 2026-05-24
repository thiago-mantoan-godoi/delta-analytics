import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLineEdit, QLabel, 
                             QFileDialog, QMessageBox, QFrame)
from PySide6.QtCore import Qt
from pypdf import PdfReader, PdfWriter

from pathlib import Path
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
	
from utils.function import *
from utils.funcoes import *


class PDFToolApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Automação de Documentos PDF")
        self.setMinimumSize(800, 400)

        # Widget Central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout Principal (Horizontal para dividir os dois lados)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(30)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # --- LADO ESQUERDO: SPLIT ---
        split_container = QFrame()
        split_container.setFrameShape(QFrame.StyledPanel)
        split_layout = QVBoxLayout(split_container)
        
        lbl_split = QLabel("DIVIDIR PDF (Split)")
        lbl_split.setStyleSheet("font-weight: bold; font-size: 16px;")
        lbl_split.setAlignment(Qt.AlignCenter)
        
        self.input_split_file = QLineEdit()
        self.input_split_file.setPlaceholderText("Selecione o arquivo PDF...")
        btn_browse_split_file = QPushButton("Procurar Arquivo")
        btn_browse_split_file.clicked.connect(self.browse_split_file)

        self.input_split_dir = QLineEdit()
        self.input_split_dir.setPlaceholderText("Pasta para salvar as páginas...")
        btn_browse_split_dir = QPushButton("Selecionar Pasta")
        btn_browse_split_dir.clicked.connect(self.browse_split_dir)

        btn_run_split = QPushButton("EXECUTAR SPLIT")
        btn_run_split.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        btn_run_split.clicked.connect(self.execute_split)

        split_layout.addWidget(lbl_split)
        split_layout.addSpacing(20)
        split_layout.addWidget(QLabel("Arquivo Original:"))
        split_layout.addWidget(self.input_split_file)
        split_layout.addWidget(btn_browse_split_file)
        split_layout.addSpacing(10)
        split_layout.addWidget(QLabel("Destino:"))
        split_layout.addWidget(self.input_split_dir)
        split_layout.addWidget(btn_browse_split_dir)
        split_layout.addStretch()
        split_layout.addWidget(btn_run_split)

        # --- LADO DIREITO: MERGE ---
        merge_container = QFrame()
        merge_container.setFrameShape(QFrame.StyledPanel)
        merge_layout = QVBoxLayout(merge_container)

        lbl_merge = QLabel("UNIR PDFS (Merge)")
        lbl_merge.setStyleSheet("font-weight: bold; font-size: 16px;")
        lbl_merge.setAlignment(Qt.AlignCenter)

        self.input_merge_dir = QLineEdit()
        self.input_merge_dir.setPlaceholderText("Pasta com os arquivos PDF...")
        btn_browse_merge_dir = QPushButton("Selecionar Pasta")
        btn_browse_merge_dir.clicked.connect(self.browse_merge_dir)

        self.input_merge_output = QLineEdit()
        self.input_merge_output.setPlaceholderText("Onde salvar o arquivo final...")
        btn_browse_merge_save = QPushButton("Definir Nome/Local")
        btn_browse_merge_save.clicked.connect(self.browse_merge_save)

        btn_run_merge = QPushButton("EXECUTAR MERGE")
        btn_run_merge.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        btn_run_merge.clicked.connect(self.execute_merge)

        merge_layout.addWidget(lbl_merge)
        merge_layout.addSpacing(20)
        merge_layout.addWidget(QLabel("Pasta com os PDFs:"))
        merge_layout.addWidget(self.input_merge_dir)
        merge_layout.addWidget(btn_browse_merge_dir)
        merge_layout.addSpacing(10)
        merge_layout.addWidget(QLabel("Salvar Resultado em:"))
        merge_layout.addWidget(self.input_merge_output)
        merge_layout.addWidget(btn_browse_merge_save)
        merge_layout.addStretch()
        merge_layout.addWidget(btn_run_merge)

        # Adicionar os containers ao layout principal
        main_layout.addWidget(split_container)
        main_layout.addWidget(merge_container)

    # --- FUNÇÕES DE NAVEGAÇÃO ---
    @log_errors_record
    def browse_split_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "PDF Files (*.pdf)")
        if file:
            self.input_split_file.setText(file)

    @log_errors_record
    def browse_split_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Destino")
        if directory:
            self.input_split_dir.setText(directory)

    @log_errors_record
    def browse_merge_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Selecionar Pasta com PDFs")
        if directory:
            self.input_merge_dir.setText(directory)

    @log_errors_record
    def browse_merge_save(self):
        file, _ = QFileDialog.getSaveFileName(self, "Salvar PDF Unificado", "", "PDF Files (*.pdf)")
        if file:
            if not file.lower().endswith(".pdf"):
                file += ".pdf"
            self.input_merge_output.setText(file)

    # --- LÓGICA DE NEGÓCIO ---
    @log_errors_record
    def execute_split(self):
        input_pdf = self.input_split_file.text()
        output_dir = self.input_split_dir.text()

        if not input_pdf or not output_dir:
            QMessageBox.warning(self, "Erro", "Selecione o arquivo de entrada e a pasta de destino.")
            return

        try:
            reader = PdfReader(input_pdf)
            base_name = os.path.splitext(os.path.basename(input_pdf))[0]

            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                output_filename = os.path.join(output_dir, f"{base_name}_pagina_{i+1}.pdf")
                with open(output_filename, "wb") as f:
                    writer.write(f)

            QMessageBox.information(self, "Sucesso", f"PDF dividido em {len(reader.pages)} arquivos.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro: {str(e)}")

    @log_errors_record
    def execute_merge(self):
        input_dir = self.input_merge_dir.text()
        output_file = self.input_merge_output.text()

        if not input_dir or not output_file:
            QMessageBox.warning(self, "Erro", "Selecione a pasta de origem e o local de destino.")
            return

        try:
            files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
            files.sort() # Ordena alfabeticamente

            if not files:
                QMessageBox.warning(self, "Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
                return

            writer = PdfWriter()
            for pdf in files:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)

            with open(output_file, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "Sucesso", "Arquivos PDF unidos com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToolApp()
    window.show()
    sys.exit(app.exec())