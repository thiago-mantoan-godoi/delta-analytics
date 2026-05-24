import sys
import pandas as pd
import os
import warnings

warnings.simplefilter("ignore")

root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)

from PySide6.QtWidgets import (
    QApplication, QMessageBox, QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QLabel, QLineEdit, QFrame,
    QStyle, QProgressDialog
)

from PySide6.QtGui import QGuiApplication, QCursor

from PySide6.QtCore import Qt, QThread, Signal

from utils.function import *



def gerar_modelo_cao(caminho_arquivo_corte, caminho_output):

    df = pd.read_excel(caminho_arquivo_corte, dtype=str)

    df = df.rename(
        columns={
            "Internal Family": "Fam_UCS",
            "External Family": "Família",
            "CIRCUIT": "Wire Nb",
            "LENGTH": "Length",
            "Comp TW": "Comp_TW",
            "COLOR1": "C1",
            "COLOR2": "C2",
            "COLOR3": "C3",
            "WIRE_TUBE_SPLICE": "Int.PN",
            "SECTIONN": "CSA",
            "TERM_A": "Term. 1",
            "TERM_B": "Term. 2",
            "SEAL_A": "Seal 1",
            "SEAL_B": "Seal 2",
            "STRIP_A": "Strip 1",
            "STRIP_B": "Strip 2",
            "JOINT_TO_A": "Joint 1",
            "JOINT_TO_B": "Joint 2",
            "RMKS_A": "Note 1",
            "RMKS_B": "Note 2",
        }
    )

    df["Maco"] = None
    df["Rate"] = None
    df["P. Vision"] = None
    df["Severidade"] = None

    for i in df.index:

        if "LGK" in str(df.loc[i, "Note 1"]):
            df.loc[i, "Severidade"] = "LGKCC"

        if (
            str(df.loc[i, "Leadset"])[4:6] == "TW"
            or str(df.loc[i, "Leadset"])[3:5] == "TW"
        ):

            if str(df.loc[i, "Leadset"])[4:6] == "TW":

                df.loc[i, "Multicore"] = str(df.loc[i, "Leadset"])[4:]
                df.loc[i, "Maco"] = 50
                df.loc[i, "Rate"] = 800

            if str(df.loc[i, "Leadset"])[3:5] == "TW":

                df.loc[i, "Multicore"] = str(df.loc[i, "Leadset"])[3:]
                df.loc[i, "Maco"] = 50
                df.loc[i, "Rate"] = 800

    if "Volume_2" not in df.columns:
        df["Volume_2"] = None

    df["Quantidade"] = df["Volume_2"]
    df["T"] = None
    df["T_1"] = None
    df["T_2"] = None
    df["Node 1"] = None
    df["Node 2"] = None

    df["Multicore"] = df["Leadset"].str.extract(r"(TW.*)")

    df = adicionar_processos(df).rename(
        columns={
            "ProcessoA": "Processo_A",
            "ProcessoB": "Processo_B"
        }
    )

    df = df[
        [
            "Família",
            "Fam_UCS",
            "Wire Nb",
            "Leadset",
            "Maco",
            "Rate",
            "P. Vision",
            "Severidade",
            "Multicore",
            "T",
            "Length",
            "Comp_TW",
            "C1",
            "C2",
            "C3",
            "Int.PN",
            "CSA",
            "Term. 1",
            "Strip 1",
            "Seal 1",
            "Node 1",
            "Joint 1",
            "T_1",
            "Note 1",
            "Term. 2",
            "Strip 2",
            "Seal 2",
            "Node 2",
            "Joint 2",
            "T_2",
            "Note 2",
            "Processo_A",
            "Processo_B",
            "Quantidade",
        ]
    ]

    saida = os.path.join(caminho_output, "Cadastro_CAO.xlsx")

    df.to_excel(saida, index=False)

class WorkerListaModelo(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(self, arquivo_corte, pasta_output):
        super().__init__()

        self.arquivo_corte = arquivo_corte
        self.pasta_output = pasta_output

    def run(self):

        try:

            gerar_modelo_cao(
                caminho_arquivo_corte=self.arquivo_corte,
                caminho_output=self.pasta_output
            )

            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))


class ListaModeloCAO(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowMinimizeButtonHint |
            Qt.WindowType.WindowMaximizeButtonHint |
            Qt.WindowType.WindowCloseButtonHint
        )

        self.setWindowTitle("TMGods 🧠 Industrial Engineering")

        self.resize(700, 250)
        
        self.centralizar_tela_atual()

        self.setStyleSheet(self.estilo())

        layout = QVBoxLayout()

        titulo = QLabel("Gerar Modelo de Cadastro CAO")
        titulo.setObjectName("titulo")

        layout.addWidget(titulo)

        layout.addWidget(self.linha())

        # =========================================================
        # Arquivo Corte
        # =========================================================

        label_arquivo = QLabel("Lista do Corte")
        label_arquivo.setObjectName("labelCampo")

        layout.addWidget(label_arquivo)

        arquivo_layout = QHBoxLayout()

        self.input_arquivo = QLineEdit()
        self.input_arquivo.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo = QPushButton()
        self.btn_arquivo.setIcon(
            self.style().standardIcon(QStyle.SP_FileIcon)
        )

        arquivo_layout.addWidget(self.input_arquivo)
        arquivo_layout.addWidget(self.btn_arquivo)

        layout.addLayout(arquivo_layout)

        layout.addWidget(self.linha())

        # =========================================================
        # Pasta Output
        # =========================================================

        label_pasta = QLabel("Pasta de destino")
        label_pasta.setObjectName("labelCampo")

        layout.addWidget(label_pasta)

        pasta_layout = QHBoxLayout()

        self.input_pasta = QLineEdit()
        self.input_pasta.setPlaceholderText("Selecione a pasta...")

        self.btn_pasta = QPushButton()
        self.btn_pasta.setIcon(
            self.style().standardIcon(QStyle.SP_DirIcon)
        )

        pasta_layout.addWidget(self.input_pasta)
        pasta_layout.addWidget(self.btn_pasta)

        layout.addLayout(pasta_layout)

        layout.addWidget(self.linha())

        # =========================================================
        # Botão Executar
        # =========================================================

        self.btn_executar = QPushButton(" Executar")
        self.btn_executar .setObjectName("botaoExec")
        self.btn_executar.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))

        layout.addWidget(self.btn_executar)

        self.setLayout(layout)

        # Conexões

        self.btn_arquivo.clicked.connect(
            self.selecionar_arquivo
        )

        self.btn_pasta.clicked.connect(
            self.selecionar_pasta
        )

        self.btn_executar.clicked.connect(
            self.executar
        )

    # =========================================================
    # EXECUTAR
    # =========================================================

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


    def executar(self):

        arquivo = self.input_arquivo.text()
        pasta = self.input_pasta.text()

        self.progresso = QProgressDialog(
            "Processando...",
            None,
            0,
            0,
            self
        )

        self.progresso.setWindowTitle("Aguarde")
        self.progresso.setCancelButton(None)

        self.progresso.show()

        self.thread = WorkerListaModelo(
            arquivo,
            pasta
        )

        self.thread.finalizado.connect(
            self.processamento_finalizado
        )

        self.thread.erro.connect(
            self.processamento_erro
        )

        self.thread.start()

    def processamento_finalizado(self):

        self.progresso.close()

        QMessageBox.information(
            self,
            "Sucesso",
            "Arquivo gerado com sucesso!"
        )

    def processamento_erro(self, erro):

        self.progresso.close()

        QMessageBox.critical(
            self,
            "Erro",
            f"Ocorreu um erro:\n\n{erro}"
        )

    # =========================================================
    # VALIDAR ARQUIVO
    # =========================================================

    def validar_arquivo(self, caminho):

        COLUNAS_OBRIGATORIAS = [
            "Internal Family",
            "External Family",
            "CIRCUIT",
            "LENGTH",
            "Leadset"
        ]

        try:

            df = pd.read_excel(caminho)

            colunas = set(df.columns)

            obrigatorias = set(COLUNAS_OBRIGATORIAS)

            faltando = obrigatorias - colunas

            if faltando:

                QMessageBox.critical(
                    self,
                    "Arquivo inválido",
                    f"Faltam colunas obrigatórias:\n{', '.join(faltando)}"
                )

                return False

            return True

        except Exception as e:

            QMessageBox.critical(
                self,
                "Erro ao ler arquivo",
                str(e)
            )

            return False

    # =========================================================
    # SELECIONAR ARQUIVO
    # =========================================================

    def selecionar_arquivo(self):

        caminho, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo",
            "",
            "Excel (*.xlsx)"
        )

        if not caminho:
            return

        if self.validar_arquivo(caminho):

            self.input_arquivo.setText(caminho)

        else:

            self.input_arquivo.clear()

    # =========================================================
    # SELECIONAR PASTA
    # =========================================================

    def selecionar_pasta(self):

        pasta = QFileDialog.getExistingDirectory(
            self,
            "Selecionar pasta"
        )

        if pasta:
            self.input_pasta.setText(pasta)

    # =========================================================
    # LINHA
    # =========================================================

    def linha(self):

        linha = QFrame()

        linha.setFrameShape(QFrame.HLine)

        linha.setStyleSheet("""
            QFrame {
                border: none;
                background-color: #444;
                max-height: 1px;
            }
        """)

        return linha

    # =========================================================
    # ESTILO
    # =========================================================

    def estilo(self):

        return """
        QWidget {
            background-color: #1e1e1e;
            color: #ffffff;
            font-size: 13px;
        }

        #titulo {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
        }

        #labelCampo {
            margin-top: 6px;
            color: #cccccc;
        }

        QPushButton {
            background-color: #2d2d2d;
            border: 1px solid #444;
            padding: 6px;
            border-radius: 6px;
        }

        QPushButton:hover {
            background-color: #3a3a3a;
        }

        QPushButton:pressed {
            background-color: #0078d7;
        }

        #botaoExec {
            background-color: #0078d7;
            border: none;
            font-weight: bold;
        }

        #botaoExec:hover {
            background-color: #0090ff;
        }

        QLineEdit {
            background-color: #2d2d2d;
            border: 1px solid #444;
            padding: 5px;
            border-radius: 4px;
        }
        """


if __name__ == "__main__":

    app = QApplication(sys.argv)

    janela = ListaModeloCAO()

    janela.show()

    sys.exit(app.exec())