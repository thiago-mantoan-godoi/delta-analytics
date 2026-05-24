import sys
import pandas as pd
import os
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
    
    
from PySide6.QtWidgets import (
    QApplication, QMessageBox, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLabel, QLineEdit, QFrame, QStyle, QProgressDialog
)
from PySide6.QtCore import QDir, Qt
from PySide6.QtGui import QIntValidator, QGuiApplication, QCursor
from PySide6.QtCore import QThread, Signal


from src.cadastro_cao import cadastro



class WorkerCadastro(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(
        self,
        arquivo_modelo,
        pasta,
        versao,
        arquivo_cabos,
        arquivo_terminais,
        arquivo_selos,
        arquivo_aplicador
    ):
        super().__init__()

        self.arquivo_modelo = arquivo_modelo
        self.pasta = pasta
        self.versao = versao
        self.arquivo_cabos = arquivo_cabos
        self.arquivo_terminais = arquivo_terminais
        self.arquivo_selos = arquivo_selos
        self.arquivo_aplicador = arquivo_aplicador

    def run(self):

        try:

            cadastro(
                caminho_arquivo=self.arquivo_modelo,
                caminho_output=self.pasta,
                nivel=self.versao,
                df_wire_cao=self.arquivo_cabos,
                df_term_cao=self.arquivo_terminais,
                df_selos_cao=self.arquivo_selos,
                df_app_cao=self.arquivo_aplicador
            )

            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))


class CadastroCao(QWidget):
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

        # 🔷 Título
        titulo = QLabel("Gerar arquivo de cadastro CAO")
        titulo.setObjectName("titulo")
        layout.addWidget(titulo)

        # 🔹 Botão baixar modelo
        self.btn_modelo = QPushButton(" Baixar modelo para cadastro")
        self.btn_modelo.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_modelo.clicked.connect(self.baixar_modelo)
        layout.addWidget(self.btn_modelo)

        layout.addWidget(self.linha())
        
        
        # =========================================================
        # 🔹 Campo numérico (0 a 10)
        # =========================================================
        label_numero = QLabel("Versão")
        label_numero.setObjectName("labelCampo")
        layout.addWidget(label_numero)

        self.input_numero = QLineEdit()
        self.input_numero.setPlaceholderText("Digite um número de 0 a 10")

        # Validação
        validator = QIntValidator(0, 10)
        self.input_numero.setValidator(validator)

        layout.addWidget(self.input_numero)

        layout.addWidget(self.linha())
        
        

        # 🔹 Arquivo Modelo --------------------------------------------------------------------
        label_arquivo_modelo = QLabel("Arquivo preenchido")
        label_arquivo_modelo.setObjectName("labelCampo")
        layout.addWidget(label_arquivo_modelo)

        arquivo_modelo_layout = QHBoxLayout()

        self.input_arquivo_modelo = QLineEdit()
        self.input_arquivo_modelo.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo_modelo = QPushButton()
        self.btn_arquivo_modelo.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_modelo_layout.addWidget(self.input_arquivo_modelo)
        arquivo_modelo_layout.addWidget(self.btn_arquivo_modelo)

        layout.addLayout(arquivo_modelo_layout)

        layout.addWidget(self.linha())
        
        
        
        # 🔹 Arquivo Cabos --------------------------------------------------------------------
        label_arquivo_cabos = QLabel("Extrato de Cabos do CAO")
        label_arquivo_cabos.setObjectName("labelCampo")
        layout.addWidget(label_arquivo_cabos)

        arquivo_cabos_layout = QHBoxLayout()

        self.input_arquivo_cabos = QLineEdit()
        self.input_arquivo_cabos.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo_cabos = QPushButton()
        self.btn_arquivo_cabos.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_cabos_layout.addWidget(self.input_arquivo_cabos)
        arquivo_cabos_layout.addWidget(self.btn_arquivo_cabos)

        layout.addLayout(arquivo_cabos_layout)

        layout.addWidget(self.linha())
        
        # 🔹 Arquivo Terminais --------------------------------------------------------------------
        label_arquivo_terminais = QLabel("Extrato de Terminais do CAO")
        label_arquivo_terminais.setObjectName("labelCampo")
        layout.addWidget(label_arquivo_terminais)

        arquivo_terminais_layout = QHBoxLayout()

        self.input_arquivo_terminais = QLineEdit()
        self.input_arquivo_terminais.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo_terminais = QPushButton()
        self.btn_arquivo_terminais.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_terminais_layout.addWidget(self.input_arquivo_terminais)
        arquivo_terminais_layout.addWidget(self.btn_arquivo_terminais)

        layout.addLayout(arquivo_terminais_layout)

        layout.addWidget(self.linha())
        
        # 🔹 Arquivo selos --------------------------------------------------------------------
        label_arquivo_selos = QLabel("Extrato de Selos do CAO")
        label_arquivo_selos.setObjectName("labelCampo")
        layout.addWidget(label_arquivo_selos)

        arquivo_selos_layout = QHBoxLayout()

        self.input_arquivo_selos = QLineEdit()
        self.input_arquivo_selos.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo_selos = QPushButton()
        self.btn_arquivo_selos.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_selos_layout.addWidget(self.input_arquivo_selos)
        arquivo_selos_layout.addWidget(self.btn_arquivo_selos)

        layout.addLayout(arquivo_selos_layout)

        layout.addWidget(self.linha())
        
        
        # 🔹 Botão baixar modelo_aplicador
        self.btn_modelo_aplicador = QPushButton(" Baixar modelo de lista de aplicadores")
        self.btn_modelo_aplicador.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.btn_modelo_aplicador.clicked.connect(self.baixar_modelo_aplicador)
        layout.addWidget(self.btn_modelo_aplicador)

        
        # 🔹 Arquivo aplicador --------------------------------------------------------------------
        label_arquivo_aplicador = QLabel("Lista de Aplicadores")
        label_arquivo_aplicador.setObjectName("labelCampo")
        layout.addWidget(label_arquivo_aplicador)

        arquivo_aplicador_layout = QHBoxLayout()

        self.input_arquivo_aplicador = QLineEdit()
        self.input_arquivo_aplicador.setPlaceholderText("Selecione o arquivo...")

        self.btn_arquivo_aplicador = QPushButton()
        self.btn_arquivo_aplicador.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        arquivo_aplicador_layout.addWidget(self.input_arquivo_aplicador)
        arquivo_aplicador_layout.addWidget(self.btn_arquivo_aplicador)

        layout.addLayout(arquivo_aplicador_layout)

        layout.addWidget(self.linha())

        # 🔹 Pasta -------------------------------------------------------------------
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
        self.btn_arquivo_modelo.clicked.connect(self.selecionar_arquivo_modelo)
        self.btn_arquivo_cabos.clicked.connect(self.selecionar_arquivo_cabos)
        self.btn_arquivo_terminais.clicked.connect(self.selecionar_arquivo_terminais)
        self.btn_arquivo_selos.clicked.connect(self.selecionar_arquivo_selos)
        self.btn_arquivo_aplicador.clicked.connect(self.selecionar_arquivo_aplicador)
        
        self.btn_pasta.clicked.connect(self.selecionar_pasta)

    # def executar(self):
    #     arquivo_modelo = self.input_arquivo_modelo.text()
    #     arquivo_cabos = self.input_arquivo_cabos.text()
    #     arquivo_terminais = self.input_arquivo_terminais.text()
    #     arquivo_selos = self.input_arquivo_selos.text()
    #     arquivo_aplicador = self.input_arquivo_aplicador.text()
    #     pasta = self.input_pasta.text()
    #     versao = self.input_numero.text()


    #     try:

    #         cadastro(
    #             caminho_arquivo=arquivo_modelo,
    #             caminho_output=pasta,
    #             nivel=versao,
    #             df_wire_cao=arquivo_cabos,
    #             df_term_cao=arquivo_terminais,
    #             df_selos_cao=arquivo_selos,
    #             df_app_cao=arquivo_aplicador
    #         )

    #         QMessageBox.information(
    #             self,
    #             "Sucesso",
    #             "Arquivo de Cadastro finalizado com sucesso!"
    #         )

    #     except Exception as erro:

    #         QMessageBox.critical(
    #             self,
    #             "Erro",
    #             f"Ocorreu um erro:\n\n{erro}"
    #         )


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


    def cadastro_finalizado(self):

        self.progresso.close()

        QMessageBox.information(
            self,
            "Sucesso",
            "Arquivo de Cadastro finalizado com sucesso!"
        )
        
    def cadastro_erro(self, erro):

        self.progresso.close()

        QMessageBox.critical(
            self,
            "Erro",
            f"Ocorreu um erro:\n\n{erro}"
        )
    def executar(self):

        arquivo_modelo = self.input_arquivo_modelo.text()
        arquivo_cabos = self.input_arquivo_cabos.text()
        arquivo_terminais = self.input_arquivo_terminais.text()
        arquivo_selos = self.input_arquivo_selos.text()
        arquivo_aplicador = self.input_arquivo_aplicador.text()
        pasta = self.input_pasta.text()
        versao = self.input_numero.text()

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

        self.thread = WorkerCadastro(
            arquivo_modelo,
            pasta,
            versao,
            arquivo_cabos,
            arquivo_terminais,
            arquivo_selos,
            arquivo_aplicador
        )

        self.thread.finalizado.connect(self.cadastro_finalizado)
        self.thread.erro.connect(self.cadastro_erro)

        self.thread.start()


    def validar_arquivo_modelo(self, caminho):
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

    def validar_arquivo_cabos(self, caminho):
        COLUNAS_OBRIGATORIAS = ['WireKey', 'Name', 'Barcode', 'Info', 'WireType', 'CrossSection',
                                'IsoDiameter', 'IsoMaterial', 'TwistDirection', 'Color1', 'Color2',
                                'Color3', 'Color4', 'NoOfStrands']
        try:
            df = pd.read_csv(caminho,sep=';')

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
        
    def validar_arquivo_terminais(self, caminho):
        COLUNAS_OBRIGATORIAS = ['TerminalKey', 'Name', 'Barcode', 'Info', 'TerminalType',
                                'DoubleCrimpHorizontal', 'FeedingType', 'TerminalLength',
                                'TerminalWidth', 'TerminalOverlength']
        try:
            df = pd.read_csv(caminho,sep=';')

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

    def validar_arquivo_selos(self, caminho):
        COLUNAS_OBRIGATORIAS = ['SealKey', 'Name', 'Barcode', 'Info', 'SealLength', 'SealWidth',
                                'SealPositionTol', 'Color']
        try:
            df = pd.read_csv(caminho,sep=';')

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
        
    def validar_arquivo_aplicador(self, caminho):
        COLUNAS_OBRIGATORIAS = ["Codigo SAP", "Terminal","Codigo SAP","Fabricante","Locação"]
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

    def criar_modelo_excel(self, caminho):
        dados = {"Família":[], "Fam_UCS":[], "Leadset":[], "Wire Nb":[],"Maco":[], "Rate":[], "P. Vision":[], "Severidade":[], "Multicore":[], 
                 "T":[], "Length":[], "Comp_TW":[], "C1":[], "C2":[],"C3":[], "Int.PN":[], "CSA":[], "Term. 1":[], "Strip 1":[],"Seal 1":[],
                 "Node 1":[], "Joint 1":[], "T_1":[], "Note 1":[], "Term. 2":[], "Strip 2":[], "Seal 2":[], "Node 2":[], "Joint 2":[], "T_2":[], 
                 "Note 2":[], "Processo_A":[], "Processo_B":[], "Quantidade":[]
        }

        df =    pd.DataFrame(dados)
        df.to_excel(caminho, index=False)
        
    def criar_modelo_aplicador_excel(self, caminho):
        dados = {"Codigo SAP":[], "Terminal":[],"Codigo SAP":[],"Fabricante":[],"Locação":[]}

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

    def baixar_modelo_aplicador(self):
        caminho, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar modelo Excel",
            "Lista_de_aplicador.xlsx",
            "Excel (*.xlsx)"
        )

        if caminho:
            try:
                self.criar_modelo_aplicador_excel(caminho)
                QMessageBox.information(self, "Sucesso", "Modelo criado com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", str(e))
                
    # 📂 Selecionar arquivo
    def selecionar_arquivo_modelo(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "Excel (*.xlsx)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo_modelo(caminho):
            self.input_arquivo_modelo.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo_modelo.clear()
            
    def selecionar_arquivo_cabos(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "csv (*.csv)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo_cabos(caminho):
            self.input_arquivo_cabos.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo_cabos.clear()
                      
    def selecionar_arquivo_terminais(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "csv (*.csv)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo_terminais(caminho):
            self.input_arquivo_terminais.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo_terminais.clear()

    def selecionar_arquivo_selos(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "csv (*.csv)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo_selos(caminho):
            self.input_arquivo_selos.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo_selos.clear()

    def selecionar_arquivo_aplicador(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", "Excel (*.xlsx)"
        )

        if not caminho:
            return

        # valida antes de aceitar
        if self.validar_arquivo_aplicador(caminho):
            self.input_arquivo_aplicador.setText(caminho)
        else:
            # rejeita o arquivo
            self.input_arquivo_aplicador.clear()


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