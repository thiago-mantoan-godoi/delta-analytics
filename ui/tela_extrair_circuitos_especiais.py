import sys
import os
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog,
    QMessageBox, QProgressDialog, QStyle
)
from PySide6.QtGui import QFont, QGuiApplication, QCursor
from PySide6.QtCore import QThread, Signal

root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
from utils.function import *




# =========================================================
# WORKER (THREAD)
# =========================================================

class WorkerExtrairCktEspeciais(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(self, caminho_base, caminho_output):
        super().__init__()
        self.caminho_base = caminho_base
        self.caminho_output = caminho_output

    def run(self):

        try:
            self.processar()
            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))

    def processar(self):

        df = pd.read_csv(self.caminho_base, sep=";")
        df = df.dropna(how="all", axis=1)

        df = df[df["ProdVersion"] == 0].reset_index(drop=True)

        df = df.drop(
            columns=[
                "ProdVersion", "Description", "BatchSize",
                "PlanTimeBatch", "UserWSText5",
                "CVXVisionSystemCommand", "UserWSText7",
                "UserWSText8", "StrippingLengthD1",
                "StrippingLengthD2", "UserWSText1",
                "UserWSText2", "UserWSText6",
                "UserWSText4", "UserWSText3",
                "StrippingLength1", "PartStripLength1",
                "StrippingLength2", "PartStripLength2",
                "StrippingLengthB2", "StrippingLengthC1",
                "TwistWireLength", "PitchLength",
                "OpenEndLength1", "OpenEndLength2",
                "ReducedLeadLength", "ReducedWire",
                "StrippingLength4", "PartStripLength4",
                "PartStripLength3", "StrippingLength3",
                "Wire1CrossSection", "Wire1Length",
            ],
            errors="ignore"
        ).reset_index(drop=True)

        df = df[~df["Leadset"].str.startswith("A", na=False)]
        df = df.reset_index(drop=True)

        df["Status"] = None

        # =====================================================
        # CLASSIFICAÇÃO
        # =====================================================

        for i in df.index:

            term1 = df.at[i, "Terminal1Key"] or ""
            term2 = df.at[i, "Terminal2Key"] or ""
            term3 = df.at[i, "Terminal3Key"] or ""
            term4 = df.at[i, "Terminal4Key"] or ""

            seal1 = df.at[i, "Seal1Key"] or ""
            seal2 = df.at[i, "Seal2Key"] or ""
            seal3 = df.at[i, "Seal3Key"] or ""
            seal4 = df.at[i, "Seal4Key"] or ""

            texto1 = df.at[i, "UserText1"] or ""
            texto2 = df.at[i, "UserText2"] or ""
            texto3 = df.at[i, "UserText3"] or ""
            texto4 = df.at[i, "UserText4"] or ""

            if (
                (term1 == texto1 and term2 == texto2 and seal1 == texto3 and seal2 == texto4)
                or (term1 == texto2 and term2 == texto1 and seal1 == texto4 and seal2 == texto3)
                or (term2 == texto1 and term4 == texto2 and seal2 == texto3 and seal4 == texto4)
            ):
                df.at[i, "Status"] = "normal"
            else:
                df.at[i, "Status"] = "Especial"

        # =====================================================
        # COMPONENTES
        # =====================================================

        list_componentes = (
            df["Terminal1Key"].dropna().tolist() +
            df["Terminal2Key"].dropna().tolist() +
            df["Seal1Key"].dropna().tolist() +
            df["Seal2Key"].dropna().tolist()
        )

        # =====================================================
        # FILTRO FINAL
        # =====================================================

        df = df[df["Status"] == "Especial"].reset_index(drop=True)

        df = df[
            df["Terminal1Key"].isna() &
            df["Terminal2Key"].isna()
        ].reset_index(drop=True)

        dict_df = {
            "Circuitos Especiais": df[["Status", "Leadset"]]
        }

        # =====================================================
        # COMPONENTES LEAD PREP
        # =====================================================

        dados = [
            ("UserText1", "Terminal"),
            ("UserText2", "Terminal"),
            ("UserText3", "Selo"),
            ("UserText4", "Selo"),
        ]

        dfs = [
            pd.DataFrame({
                "Código": df[col].unique(),
                "Descrição": desc
            })
            for col, desc in dados
        ]

        df_new = (
            pd.concat(dfs, ignore_index=True)
            .drop_duplicates()
            .reset_index(drop=True)
        )

        df_new = df_new[~df_new["Código"].isin(list_componentes)]
        df_new = df_new.dropna(subset=["Código"]).reset_index(drop=True)

        df_new["Area"] = "Lead Prep"

        dict_df["Componentes do Lead Prep"] = df_new

        salvar_dict_df_em_excel(
            dict_df,
            os.path.join(self.caminho_output, "Circuitos_Especiais.xlsx")
        )


# =========================================================
# UI
# =========================================================

class TelaExtrairCktEspeciais(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TMGods 🧠 Industrial Engineering")
        self.resize(700, 150)
        self.centralizar_tela_atual()
        layout = QVBoxLayout()

        # =====================================================
        # ARQUIVO
        # =====================================================
        titulo = QLabel("Extrair circuitos especiais")
        titulo.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(titulo)
        layout.addWidget(self.linha())
        layout.addWidget(QLabel("Arquivo CSV"))

        file_layout = QHBoxLayout()

        self.input_file = QLineEdit()
        self.input_file.setPlaceholderText("Selecione o arquivo...")

        
        
        self.btn_file = QPushButton()
        self.btn_file.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        file_layout.addWidget(self.input_file)
        file_layout.addWidget(self.btn_file)

        layout.addLayout(file_layout)
        layout.addWidget(self.linha())
        # =====================================================
        # PASTA
        # =====================================================

        layout.addWidget(QLabel("Pasta de saída"))

        out_layout = QHBoxLayout()

        self.input_out = QLineEdit()
        self.input_out.setPlaceholderText("Selecione a pasta...")

        self.btn_out = QPushButton()
        self.btn_out.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))

        out_layout.addWidget(self.input_out)
        out_layout.addWidget(self.btn_out)

        layout.addLayout(out_layout)
        layout.addWidget(self.linha())
        # =====================================================
        # BOTÃO EXECUTAR
        # =====================================================

        btn_layout = QHBoxLayout()

        btn_layout.addStretch()  # empurra para direita

        self.btn_run = QPushButton("Executar")
        self.btn_run.setFixedWidth(120)

        self.btn_run.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                font-weight: bold;
                padding: 6px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #0090ff;
            }
        """)

        btn_layout.addWidget(self.btn_run)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # =====================================================
        # CONEXÕES
        # =====================================================

        self.btn_file.clicked.connect(self.selecionar_arquivo)
        self.btn_out.clicked.connect(self.selecionar_pasta)
        self.btn_run.clicked.connect(self.executar)

    # =====================================================
    # SELECT FILE
    # =====================================================
    
    @log_errors_record
    def linha(self):

        frame = QWidget()
        line = QVBoxLayout(frame)

        sep = QLabel("")
        sep.setStyleSheet("background-color:#444; max-height:1px;")

        return sep
    
    @log_errors_record
    def centralizar_tela_atual(self):

        # pega a tela onde o mouse está
        screen = QGuiApplication.screenAt(QCursor.pos())

        # fallback
        if screen is None:
            screen = QGuiApplication.primaryScreen()

        geo = screen.availableGeometry()

        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2

        self.move(x, y)
    
    @log_errors_record
    def selecionar_arquivo(self):

        file, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar CSV",
            "",
            "CSV (*.csv)"
        )

        if file:
            self.input_file.setText(file)

    # =====================================================
    # SELECT FOLDER
    # =====================================================
    @log_errors_record
    def selecionar_pasta(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Selecionar pasta"
        )

        if folder:
            self.input_out.setText(folder)

    # =====================================================
    # EXECUTAR
    # =====================================================
    @log_errors_record
    def executar(self):

        arquivo = self.input_file.text()
        pasta = self.input_out.text()

        self.progress = QProgressDialog(
            "Processando...",
            None,
            0,
            0,
            self
        )

        self.progress.setWindowTitle("Aguarde")
        self.progress.setCancelButton(None)
        self.progress.show()

        self.thread = WorkerExtrairCktEspeciais(
            arquivo,
            pasta
        )

        self.thread.finalizado.connect(self.sucesso)
        self.thread.erro.connect(self.erro)

        self.thread.start()

    # =====================================================
    # CALLBACKS
    # =====================================================
    @log_errors_record
    def sucesso(self):

        self.progress.close()

        QMessageBox.information(
            self,
            "Sucesso",
            "Arquivo gerado com sucesso!"
        )
    @log_errors_record
    def erro(self, msg):

        self.progress.close()

        QMessageBox.critical(
            self,
            "Erro",
            msg
        )


# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":

    app = QApplication(sys.argv)

    window = TelaExtrairCktEspeciais()
    window.show()

    sys.exit(app.exec())