import sys
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMessageBox, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLabel, QLineEdit, QFrame, QStyle
)
from PySide6.QtCore import QDir, Qt
from PySide6.QtGui import QGuiApplication, QCursor





class CadastroCao(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowMinimizeButtonHint |
            Qt.WindowType.WindowMaximizeButtonHint |
            Qt.WindowType.WindowCloseButtonHint
        )

        self.setWindowTitle("Processador")
        self.resize(500, 320)
        self.centralizar_tela_atual()
        self.setStyleSheet(self.estilo())

        layout = QVBoxLayout()

        # 🔷 Título
        titulo = QLabel("Cadastro CAO")
        titulo.setObjectName("titulo")
        layout.addWidget(titulo)

        # 🔹 Botão baixar modelo
        self.btn_modelo = QPushButton(" Baixar modelo Excel")
        self.btn_modelo.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_modelo.clicked.connect(self.baixar_modelo)
        layout.addWidget(self.btn_modelo)

        layout.addWidget(self.linha())

        # 🔹 Arquivo
        label_arquivo = QLabel("Arquivo preenchido")
        label_arquivo.setObjectName("labelCampo")
        layout.addWidget(label_arquivo)

        arquivo_layout = QHBoxLayout()

        self.input_arquivo = QLineEdit()
        self.input_arquivo.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo = QPushButton()
        self.btn_arquivo.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_layout.addWidget(self.input_arquivo)
        arquivo_layout.addWidget(self.btn_arquivo)

        layout.addLayout(arquivo_layout)

        layout.addWidget(self.linha())

        # 🔹 Pasta
        label_pasta = QLabel("Pasta de destino")
        label_pasta.setObjectName("labelCampo")
        layout.addWidget(label_pasta)

        pasta_layout = QHBoxLayout()

        self.input_pasta = QLineEdit()
        self.input_pasta.setPlaceholderText("Selecione a pasta...")

        self.btn_pasta = QPushButton()
        self.btn_pasta.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))

        pasta_layout.addWidget(self.input_pasta)
        pasta_layout.addWidget(self.btn_pasta)

        layout.addLayout(pasta_layout)

        layout.addWidget(self.linha())

        # 🔹 Executar
        self.btn_executar = QPushButton(" Executar")
        self.btn_executar.setObjectName("botaoExec")
        self.btn_executar.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.btn_executar.clicked.connect(self.executar)
        layout.addWidget(self.btn_executar)

        self.setLayout(layout)

        # Conexões
        self.btn_arquivo.clicked.connect(self.selecionar_arquivo)
        self.btn_pasta.clicked.connect(self.selecionar_pasta)

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

    def validar_arquivo(self, caminho):
        COLUNAS_OBRIGATORIAS = ["Família", "Fam_UCS", "Leadset", "Wire Nb","Maco", "Rate", "P. Vision", "Severidade", "Multicore", 
                                "T", "Length", "Comp_TW", "C1", "C2","C3", "Int.PN", "CSA", "Term. 1", "Strip 1","Seal 1","Node 1", 
                                "Joint 1", "T_1", "Note 1", "Term. 2", "Strip 2", "Seal 2", "Node 2", "Joint 2", "T_2", "Note 2", 
                                "Processo_A", "Processo_B", "Quantidade"]
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
            QMessageBox.critical(self, "Erro ao ler arquivo", str(e))
            return False

    def executar(self):
        arquivo = self.input_arquivo.text()
        pasta = self.input_pasta.text()

        print("Arquivo carregado:", arquivo)
        print("Pasta selecionada:", pasta)

    def criar_modelo_excel(self, caminho):
        dados = {"Família":[], "Fam_UCS":[], "Leadset":[], "Wire Nb":[],"Maco":[], "Rate":[], "P. Vision":[], "Severidade":[], "Multicore":[], 
                 "T":[], "Length":[], "Comp_TW":[], "C1":[], "C2":[],"C3":[], "Int.PN":[], "CSA":[], "Term. 1":[], "Strip 1":[],"Seal 1":[],
                 "Node 1":[], "Joint 1":[], "T_1":[], "Note 1":[], "Term. 2":[], "Strip 2":[], "Seal 2":[], "Node 2":[], "Joint 2":[], "T_2":[], 
                 "Note 2":[], "Processo_A":[], "Processo_B":[], "Quantidade":[]
        }

        df =    pd.DataFrame(dados)
        df.to_excel(caminho, index=False)

    def baixar_modelo(self):
        caminho, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar modelo Excel",
            "Modelo_de_cadastro_CAO.xlsx",
            "Excel (*.xlsx)"
        )

        if caminho:
            try:
                self.criar_modelo_excel(caminho)
                QMessageBox.information(self, "Sucesso", "Modelo criado com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))
                
    # 📂 Selecionar arquivo
    def selecionar_arquivo(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "Excel (*.xlsx)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo(caminho):
            self.input_arquivo.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo.clear()

    # 📁 Selecionar pasta
    def selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar pasta")
        if pasta:
            self.input_pasta.setText(pasta)

    # ➖ Linha separadora
    def linha(self):
        linha = QFrame()
        linha.setFrameShape(QFrame.HLine)
        #linha.setFrameShadow(QFrame.Sunken)
        linha.setStyleSheet("""
                            QFrame {
                                border: none;
                                background-color: #444;
                                max-height: 1px;
                                }
                                """)
        return linha

    # 🎨 Estilo
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
    janela = CadastroCao()
    janela.show()
    sys.exit(app.exec())