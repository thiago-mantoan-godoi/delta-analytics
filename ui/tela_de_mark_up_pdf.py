import sys
import os
import pandas as pd
import re
import fitz  # PyMuPDF
import shutil
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QFileDialog, QProgressBar, QLabel, 
                             QMessageBox, QScrollArea)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QPixmap, QImage

from pathlib import Path
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
	
from utils.function import *
from utils.funcoes import *

class PDFWorker(QThread):
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, excel_path, pdf_path):
        super().__init__()
        self.excel_path = excel_path
        self.pdf_path = pdf_path

    def run(self):
        try:
            # 1. Carregar dados do Excel
            df = pd.read_excel(self.excel_path)
            if 'Item' not in df.columns or 'Descrição' not in df.columns:
                self.error.emit("O Excel deve conter exatamente as colunas 'Item' e 'Descrição'.")
                return

            # 2. Abrir o documento PDF
            doc = fitz.open(self.pdf_path)
            total_pages = len(doc)
            items_to_find = df.to_dict('records')
            red = (1, 0, 0)  # Cor vermelha para os desenhos

            for p_idx, page in enumerate(doc):
                # Extrair palavras com coordenadas
                words = page.get_text("words") 

                for row in items_to_find:
                    item_str = str(row['Item']).strip()
                    desc_str = str(row['Descrição']).strip()

                    # Criar padrão Regex para busca exata de palavra isolada
                    padrao_busca = rf"\b{re.escape(item_str)}\b"

                    for w in words:
                        texto_pdf = w[4].strip()
                        
                        if re.search(padrao_busca, texto_pdf):
                            rect_item = fitz.Rect(w[:4])
                            
                            # --- 1. Desenhar Retângulo no Item ---
                            page.draw_rect(rect_item, color=red, width=1.5)
                            
                            # --- 2. Criar Caixa de Descrição ---
                            distancia = 20  
                            largura_caixa = max(40, len(desc_str) * 6) 
                            altura_caixa = 15 

                            desc_box = fitz.Rect(
                                rect_item.x1 + distancia, 
                                rect_item.y0 - (altura_caixa / 2), 
                                rect_item.x1 + distancia + largura_caixa, 
                                rect_item.y0 + (altura_caixa / 2)
                            )
                            
                            page.draw_rect(desc_box, color=red, width=1)
                            
                            page.insert_textbox(desc_box, desc_str, 
                                              fontsize=8, 
                                              color=red, 
                                              align=1) 
                            
                            # --- 3. Desenhar a Seta ---
                            p_inicio = fitz.Point(desc_box.x0, (desc_box.y0 + desc_box.y1) / 2)
                            p_fim = fitz.Point(rect_item.x1, (rect_item.y0 + rect_item.y1) / 2)
                            
                            line_annot = page.add_line_annot(p_inicio, p_fim)
                            line_annot.set_colors(stroke=red)
                            line_annot.set_line_ends(0, 4) 
                            line_annot.update()

                self.progress.emit(int(((p_idx + 1) / total_pages) * 100))

            # Salvar em um arquivo temporário para preview
            temp_output = os.path.join(os.getcwd(), "temp_preview.pdf")
            doc.save(temp_output)
            doc.close()
            self.finished.emit(temp_output)

        except Exception as e:
            self.error.emit(f"Erro inesperado: {str(e)}")

class MarkUpPdf(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("App de Automação de PDF")
        self.setMinimumSize(1200, 720)

        self.excel_path = ""
        self.pdf_path = ""
        self.processed_pdf_path = ""

        # Layout da Interface
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        self.btn_excel = QPushButton("1. Selecionar Planilha Excel")
        self.btn_excel.setMinimumHeight(30)
        self.btn_excel.setFixedWidth(220)
        self.btn_excel.setStyleSheet("text-align: left; padding-left: 10px;")
        self.btn_excel.clicked.connect(self.select_excel)
        layout.addWidget(self.btn_excel)

        # NOVO BOTÃO PARA BAIXAR MODELO
        self.btn_modelo = QPushButton("2. 📥 Baixar modelo de arquivo (.xlsx)")
        self.btn_modelo.setMinimumHeight(30)
        self.btn_modelo.setFixedWidth(220)
        self.btn_modelo.setStyleSheet("background-color: #7f8c8d; color: white; text-align: left; padding-left: 10px;")
        self.btn_modelo.clicked.connect(self.download_template)
        layout.addWidget(self.btn_modelo)

        self.btn_pdf = QPushButton("3. Selecionar Arquivo PDF")
        self.btn_pdf.setMinimumHeight(30)
        self.btn_pdf.setFixedWidth(220)
        self.btn_pdf.setStyleSheet("text-align: left; padding-left: 10px;")
        self.btn_pdf.clicked.connect(self.select_pdf)
        layout.addWidget(self.btn_pdf)

        self.btn_processar = QPushButton("4. GERAR PDF COM SETAS")
        self.btn_processar.setMinimumHeight(30)
        self.btn_processar.setFixedWidth(220)
        self.btn_processar.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; text-align: left; padding-left: 10px;")
        self.btn_processar.clicked.connect(self.start_work)
        layout.addWidget(self.btn_processar)

        self.p_bar = QProgressBar()
        layout.addWidget(self.p_bar)

        self.status_lbl = QLabel("Aguardando comandos...")
        layout.addWidget(self.status_lbl)

        self.scroll = QScrollArea()
        self.preview_lbl = QLabel("Preview aparecerá aqui")
        self.preview_lbl.setAlignment(Qt.AlignCenter)
        self.scroll.setWidget(self.preview_lbl)
        self.scroll.setWidgetResizable(True)
        layout.addWidget(self.scroll)

        # BOTÃO SALVAR (Inicia desabilitado)
        self.btn_salvar = QPushButton("💾 SALVAR PDF PROCESSADO")
        self.btn_salvar.setMinimumHeight(40)
        self.btn_salvar.setEnabled(False)
        self.btn_salvar.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold;")
        self.btn_salvar.clicked.connect(self.save_final_pdf)
        layout.addWidget(self.btn_salvar)

    def select_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.btn_excel.setText(f"Excel: {os.path.basename(path)}")

    def select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir PDF", "", "PDF (*.pdf)")
        if path:
            self.pdf_path = path
            self.btn_pdf.setText(f"PDF: {os.path.basename(path)}")

    def start_work(self):
        if not self.excel_path or not self.pdf_path:
            QMessageBox.warning(self, "Aviso", "Selecione os dois arquivos primeiro!")
            return
        
        self.btn_processar.setEnabled(False)
        self.btn_salvar.setEnabled(False)
        self.status_lbl.setText("Processando...")
        self.worker = PDFWorker(self.excel_path, self.pdf_path)
        self.worker.progress.connect(self.p_bar.setValue)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(self.on_success)
        self.worker.start()

    def on_error(self, msg):
        QMessageBox.critical(self, "Erro", msg)
        self.btn_processar.setEnabled(True)

    def on_success(self, path):
        self.btn_processar.setEnabled(True)
        self.btn_salvar.setEnabled(True)
        self.processed_pdf_path = path
        self.status_lbl.setText("Processamento concluído! Clique em Salvar para escolher o local.")
        
        # Gerar a imagem de preview
        doc = fitz.open(path)
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        self.preview_lbl.setPixmap(QPixmap.fromImage(img))
        doc.close()

    def save_final_pdf(self):
        if not self.processed_pdf_path:
            return

        # Abre diálogo para escolher onde salvar
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo PDF", "PDF_Anotado_Final.pdf", "PDF Files (*.pdf)")
        
        if save_path:
            try:
                shutil.copy(self.processed_pdf_path, save_path)
                QMessageBox.information(self, "Sucesso", f"Arquivo salvo com sucesso em:\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo: {str(e)}")

    def download_template(self):
        """Gera e salva um arquivo Excel vazio com as colunas necessárias"""
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Modelo Excel", "modelo_automacao.xlsx", "Excel Files (*.xlsx)"
        )
        if save_path:
            try:
                if not save_path.endswith('.xlsx'):
                    save_path += '.xlsx'
                
                df_modelo = pd.DataFrame(columns=["Item", "Descrição"])
                df_modelo.to_excel(save_path, index=False)
                QMessageBox.information(self, "Sucesso", f"Modelo criado com sucesso em:\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao gerar modelo: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MarkUpPdf()
    window.show()
    sys.exit(app.exec())