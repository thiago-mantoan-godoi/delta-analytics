import json
import os
import platform
import psutil
import socket
import uuid
from posixpath import sep
import sys
from datetime import datetime
import win32com.client
from PySide6.QtWidgets import ( 
    QApplication, QComboBox, QDialog, QLineEdit, QScrollArea, QSizePolicy, QSpinBox, QTextEdit, QWidget,
    QHBoxLayout, QVBoxLayout,
    QGroupBox, QTabWidget,
    QTableWidget, QTableWidgetItem, QFrame, 
    QLabel, QPushButton, QCheckBox
)

from PySide6.QtWidgets import QFileDialog, QMessageBox, QTableWidget, QHeaderView
from PySide6.QtCore import QDir, Qt

import pandas as pd
from utils.funcoes import (converte_arquivo_sap, adicionar_sequencia, definir_processos, 
                           add_volumes,top_processos_memoria, obter_info_colaborador, obter_info_maquina, testar_latencia)

from functools import wraps
import traceback
import logging

logger = logging.getLogger(__name__)

AUTOR = 'TMGods'
VERSAO = 1.0
ANO = 2026



def log_errors(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        self = args[0]

        try:
            return func(*args, **kwargs)

        except Exception as e:
            tb = sys.exc_info()[2]
            last_frame = traceback.extract_tb(tb)[-1]

            data_hora_atual = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

            # erro = (
            #     f"{'DATA/HORA':<25}: {data_hora_atual}\n"
            #     f"{'FUNÇÃO':<25}: {func.__name__}\n"
            #     f"{'ARQUIVO':<25}: {last_frame.filename}\n"
            #     f"{'LINHA':<25}: {last_frame.lineno}\n"
            #     f"{'CÓDIGO':<25}: {last_frame.line}\n"
            #     f"{'ERRO':<25}: {e}\n\n"
            #     f"{'TRACEBACK':<25}:\n{traceback.format_exc()}"
            # )
            
            erro = (
                f"{'='*80}\n"
                f"{'DATA/HORA':<25}: {data_hora_atual}\n"
                f"{'FUNÇÃO':<25}: {func.__name__}\n"
                f"{'TIPO ERRO':<25}: {type(e).__name__}\n"
                f"{'ARQUIVO':<25}: {last_frame.filename}\n"
                f"{'LINHA':<25}: {last_frame.lineno}\n"
                f"{'CÓDIGO':<25}: {last_frame.line}\n"
                f"{'ERRO':<25}: {e}\n"
                f"{'ARGS':<25}: {args[1:]}\n"
                f"{'KWARGS':<25}: {kwargs}\n\n"
                f"{'TRACEBACK':<25}:\n{traceback.format_exc()}\n"
            )

            if hasattr(self, "label_erros_report"):
                # self.label_erros_report.setText(f"<pre>{erro}</pre>")
                # QMessageBox.critical(self, "Erro", f"Erro:\n{erro}")
                
                # Acumula erros anteriores
                erros_anteriores = self.label_erros_report.toPlainText()
                if erros_anteriores:
                    self.label_erros_report.setPlainText(f"{erros_anteriores}\n{erro}")
                else:
                    self.label_erros_report.setPlainText(f"<pre>{erro}</pre>")

                # Mensagem resumida em popup
                QMessageBox.critical(
                    self,
                    "Erro",
                    f"{type(e).__name__}: {e}"
                )
            return None

    return wrapper

class MainWindow(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.WindowMinimizeButtonHint |
            Qt.WindowType.WindowMaximizeButtonHint |
            Qt.WindowType.WindowCloseButtonHint
        )

        self.setWindowTitle("Engenharia")
        self.resize(1000, 700)

        # Layout principal (horizontal)
        main_layout = QHBoxLayout(self)

        #==============================
        self.arquivo_sap = pd.DataFrame()
        self.lista_do_corte = pd.DataFrame()
        self.df_filtrado = pd.DataFrame()
        # Valores iniciais
        self.vision_komax = 17
        self.vision_schleuniger = 0
        #==============================

        # ==============================
        # LADO ESQUERDO - GROUPBOX
        # ==============================
        info_layout = QVBoxLayout()

        group_box = QGroupBox("Informações")
        group_layout = QVBoxLayout()

        autor = QLabel(f"{AUTOR} | {VERSAO} | {ANO}")
        autor.setAlignment(Qt.AlignCenter)

        user = QLabel(f"User: {os.environ.get('USERNAME')}")
        user.setStyleSheet("padding: 2px; color: #ddd;")
        
        versao_python = QLabel(f"Versão do Python: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
        versao_python.setStyleSheet("padding: 2px; color: #ddd;")
        
        latencia = QLabel(f"Latência da rede: {testar_latencia()}")
        latencia.setStyleSheet("padding: 2px; color: #ddd;")

        info = obter_info_colaborador()

        # mapeamento dos campos
        campos = {
            "Nome": info.get("Nome"),
            "E-mail": info.get("E-mail"),
            "Cargo": info.get("Cargo"),
            "Departamento": info.get("Departamento"),
            "Empresa": info.get("Empresa"),
            #"Escritorio": info.get("Escritorio"),
            "Telefone": info.get("Telefone"),
            #"Celular": info.get("Celular"),
            "Gestor Direto": info.get("Gestor_Direto"),
            "Cidade": info.get("Cidade"),
        }
        
        info_maq = obter_info_maquina()
        
        campos_maq = {
            "Nome_Computador": info_maq.get("Nome_Computador"),
            "Sistema_Operacional": info_maq.get("Sistema_Operacional"),
            #"Versao_SO": info_maq.get("Versao_SO"),
            "Arquitetura": info_maq.get("Arquitetura"),
            #"Processador": info_maq.get("Processador"),
            "Nucleos_Fisicos": info_maq.get("Nucleos_Fisicos"),
            "Memoria_RAM_GB": info_maq.get("Memoria_RAM_GB"),
            #"Usuario_Logado": info_maq.get("Usuario_Logado"),
            "Endereco_IP": info_maq.get("Endereco_IP"),
            #"MAC_Address": info_maq.get("MAC_Address")
        }
        

        # layout onde os labels serão colocados
        group_layout = QVBoxLayout()

        # cria labels automaticamente


        group_layout.addWidget(user)
        for chave, valor in campos.items():
            label = QLabel(f"{chave}: {valor}")
            label.setStyleSheet("padding: 2px; color: #ddd;")
            group_layout.addWidget(label)
        group_layout.addWidget(self.criar_linha())
        
        
        for chave, valor in campos_maq.items():
            label = QLabel(f"{chave}: {valor}")
            label.setStyleSheet("padding: 2px; color: #ddd;")
            group_layout.addWidget(label)
            
        group_layout.addWidget(self.criar_linha())

        for nome, mem in top_processos_memoria():
            label = QLabel(f"{nome}: {mem:>6.2f} GB")
            label.setStyleSheet("padding: 2px; color: #ddd;")
            group_layout.addWidget(label)
        
        group_layout.addWidget(self.criar_linha())
        group_layout.addWidget(versao_python)
        group_layout.addWidget(latencia)
        group_layout.addStretch()

        group_box.setLayout(group_layout)

        info_layout.addWidget(group_box)
        info_layout.addWidget(autor)

        # aplica no widget principal
        main_layout.addLayout(info_layout)
        
        # ==============================
        # LADO DIREITO - TABS
        # ==============================
        tabs = QTabWidget()

        # ----------------------------------------------------- Informação Base ----------------------------------------------------- 
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        tab1 = QWidget()

        # Layout principal da aba
        tab1_layout = QVBoxLayout()

        # Botões para atualizar lista de especificações de cabos
        self.btn_padrao_lista_de_cabos = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_cabos = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_cabos = QPushButton("Visualizar lista de cabos")
        
        self.btn_padrao_lista_de_cabos.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_cabos(),nome_padrao='Lista_de_cabos.csv'))
        self.btn_atualizar_lista_cabos.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_cabos.json"))
        caminho_cabos = os.path.join(os.getcwd(), "data", "Lista_de_cabos.json")
        self.btn_visualizar_lista_cabos.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_cabos))
        self.label_cabos = QLabel(f"{self.status_json(caminho_cabos)}")
        self.label_cabos.setFixedWidth(30)
        self.label_cabos.setAlignment(Qt.AlignCenter)
        
        # Botões para atualizar lista de especificações de terminais
        self.btn_padrao_lista_de_terminais = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_terminais = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_terminais = QPushButton("Visualizar lista de terminais")
        
        self.btn_padrao_lista_de_terminais.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_terminais(),nome_padrao='Lista_de_terminais.csv'))
        self.btn_atualizar_lista_terminais.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_terminais.json"))
        caminho_terminais = os.path.join(os.getcwd(), "data", "Lista_de_terminais.json")
        self.btn_visualizar_lista_terminais.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_terminais))
        self.label_terminais = QLabel(f"{self.status_json(caminho_terminais)}")
        self.label_terminais.setFixedWidth(30)
        self.label_terminais.setAlignment(Qt.AlignCenter)
        
        # Botões para atualizar lista de especificações de selos
        self.btn_padrao_lista_de_selos = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_selos = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_selos = QPushButton("Visualizar lista de selos")
        
        self.btn_padrao_lista_de_selos.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_selos(),nome_padrao='Lista_de_selos.csv'))
        self.btn_atualizar_lista_selos.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_selos.json"))
        caminho_selos = os.path.join(os.getcwd(), "data", "Lista_de_selos.json")
        self.btn_visualizar_lista_selos.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_selos))
        self.label_selos = QLabel(f"{self.status_json(caminho_selos)}")
        self.label_selos.setFixedWidth(30)
        self.label_selos.setAlignment(Qt.AlignCenter)        


        # Botões para atualizar lista de especificações de Máquina
        self.btn_padrao_lista_de_maquinas = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_maquinas = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_maquinas = QPushButton("Visualizar lista de Máquinas do Corte")
        
        self.btn_padrao_lista_de_maquinas.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_maquinas(),nome_padrao='Lista_de_maquinas.csv'))
        self.btn_atualizar_lista_maquinas.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_maquinas.json"))
        caminho_maquinas = os.path.join(os.getcwd(), "data", "Lista_de_maquinas.json")
        self.btn_visualizar_lista_maquinas.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_maquinas))
        self.label_maquinas = QLabel(f"{self.status_json(caminho_maquinas)}")
        self.label_maquinas.setFixedWidth(30)
        self.label_maquinas.setAlignment(Qt.AlignCenter)     
        

        # Botões para atualizar lista de especificações de Máquina
        self.btn_padrao_lista_de_maq_lead_prep = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_maq_lead_prep = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_maq_lead_prep = QPushButton("Visualizar lista de Máquinas do Lead Prep")
        
        self.btn_padrao_lista_de_maq_lead_prep.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_maq_lead_prep(),nome_padrao='Lista_de_maq_lead_prep.csv'))
        self.btn_atualizar_lista_maq_lead_prep.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_maq_lead_prep.json"))
        caminho_maq_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_maq_lead_prep.json")
        self.btn_visualizar_lista_maq_lead_prep.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_maq_lead_prep))
        self.label_maq_lead_prep = QLabel(f"{self.status_json(caminho_maq_lead_prep)}")
        self.label_maq_lead_prep.setFixedWidth(30)
        self.label_maq_lead_prep.setAlignment(Qt.AlignCenter)     
        
        
        # Botões para atualizar lista de setup Corte
        self.btn_padrao_lista_de_setup_corte = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_setup_corte = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_setup_corte = QPushButton("Visualizar tempos Setup")
        
        self.btn_padrao_lista_de_setup_corte.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_setup_corte(),nome_padrao='Lista_de_setup_corte.csv'))
        self.btn_atualizar_lista_setup_corte.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_setup_corte.json"))
        caminho_setup_corte = os.path.join(os.getcwd(), "data", "Lista_de_setup_corte.json")
        self.btn_visualizar_lista_setup_corte.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_setup_corte))
        self.label_setup_corte = QLabel(f"{self.status_json(caminho_setup_corte)}")
        self.label_setup_corte.setFixedWidth(30)
        self.label_setup_corte.setAlignment(Qt.AlignCenter)   
        
        
        # Botões para atualizar lista de setup Lead Prep
        self.btn_padrao_lista_de_setup_lead_prep = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_setup_lead_prep = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_setup_lead_prep = QPushButton("Visualizar tempos Setup")
        
        self.btn_padrao_lista_de_setup_lead_prep.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_setup_lead_prep(),nome_padrao='Lista_de_setup_lead_prep.csv'))
        self.btn_atualizar_lista_setup_lead_prep.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_setup_lead_prep.json"))
        caminho_setup_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_setup_lead_prep.json")
        self.btn_visualizar_lista_setup_lead_prep.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_setup_lead_prep))
        self.label_setup_lead_prep = QLabel(f"{self.status_json(caminho_setup_lead_prep)}")
        self.label_setup_lead_prep.setFixedWidth(30)
        self.label_setup_lead_prep.setAlignment(Qt.AlignCenter)  
        
        # Botões para atualizar lista de Rates Corte
        self.btn_padrao_lista_de_rates_corte = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_rates_corte = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_rates_corte = QPushButton("Visualizar Rates Corte")
        
        self.btn_padrao_lista_de_rates_corte.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_rates_corte(),nome_padrao='Lista_de_rates_corte.csv'))
        self.btn_atualizar_lista_rates_corte.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_rates_corte.json"))
        caminho_rates_corte = os.path.join(os.getcwd(), "data", "Lista_de_rates_corte.json")
        self.btn_visualizar_lista_rates_corte.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_rates_corte))
        self.label_rates_corte = QLabel(f"{self.status_json(caminho_rates_corte)}")
        self.label_rates_corte.setFixedWidth(30)
        self.label_rates_corte.setAlignment(Qt.AlignCenter)  
        
        # Botões para atualizar lista de Rates Lead Prep
        self.btn_padrao_lista_de_rates_lead_prep = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_rates_lead_prep = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_rates_lead_prep = QPushButton("Visualizar Rates Lead Prep")
        
        self.btn_padrao_lista_de_rates_lead_prep.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_rates_lead_prep(),nome_padrao='Lista_de_rates_lead_prep.csv'))
        self.btn_atualizar_lista_rates_lead_prep.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_rates_lead_prep.json"))
        caminho_rates_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_rates_lead_prep.json")
        self.btn_visualizar_lista_rates_lead_prep.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_rates_lead_prep))
        self.label_rates_lead_prep = QLabel(f"{self.status_json(caminho_rates_lead_prep)}")
        self.label_rates_lead_prep.setFixedWidth(30)
        self.label_rates_lead_prep.setAlignment(Qt.AlignCenter)  
        
        # Botões para atualizar lista ZMM247
        self.btn_padrao_lista_zmm247 = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_zmm247 = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_zmm247 = QPushButton("Visualizar ZMM247")
        
        self.btn_padrao_lista_zmm247.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_zmm247(),nome_padrao='Lista_de_zmm247.csv'))
        self.btn_atualizar_lista_zmm247.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_zmm247.json"))
        caminho_zmm247 = os.path.join(os.getcwd(), "data", "Lista_de_zmm247.json")
        self.btn_visualizar_lista_zmm247.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_zmm247))
        self.label_zmm247 = QLabel(f"{self.status_json(caminho_zmm247)}")
        self.label_zmm247.setFixedWidth(30)
        self.label_zmm247.setAlignment(Qt.AlignCenter)  
        
        # Botões para atualizar Master kanban
        self.btn_padrao_lista_master_kanban = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_master_kanban = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_master_kanban = QPushButton("Visualizar Master kanban")
        
        self.btn_padrao_lista_master_kanban.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_master_kanban(),nome_padrao='Lista_de_master_kanban.csv'))
        self.btn_atualizar_lista_master_kanban.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_master_kanban.json"))
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_master_kanban.json")
        self.btn_visualizar_lista_master_kanban.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_master_kanban))
        self.label_master_kanban = QLabel(f"{self.status_json(caminho_master_kanban)}")
        self.label_master_kanban.setFixedWidth(30)
        self.label_master_kanban.setAlignment(Qt.AlignCenter) 
        
        # Botões para atualizar Lista de aplicadores
        self.btn_padrao_lista_aplicadores = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_aplicadores = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_aplicadores = QPushButton("Visualizar Aplicadores")
        
        self.btn_padrao_lista_aplicadores.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_aplicadores(),nome_padrao='Lista_de_aplicadores.csv'))
        self.btn_atualizar_lista_aplicadores.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_aplicadores.json"))
        caminho_aplicadores = os.path.join(os.getcwd(), "data", "Lista_de_aplicadores.json")
        self.btn_visualizar_lista_aplicadores.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_aplicadores))
        self.label_aplicadores = QLabel(f"{self.status_json(caminho_aplicadores)}")
        self.label_aplicadores.setFixedWidth(30)
        self.label_aplicadores.setAlignment(Qt.AlignCenter) 
        
        # Botões para atualizar Lista de calhas
        self.btn_padrao_lista_calhas = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_calhas = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_calhas = QPushButton("Visualizar Calhas")
        
        self.btn_padrao_lista_calhas.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_calhas(),nome_padrao='Lista_de_calhas.csv'))
        self.btn_atualizar_lista_calhas.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_calhas.json"))
        caminho_calhas = os.path.join(os.getcwd(), "data", "Lista_de_calhas.json")
        self.btn_visualizar_lista_calhas.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_calhas))
        self.label_calhas = QLabel(f"{self.status_json(caminho_calhas)}")
        self.label_calhas.setFixedWidth(30)
        self.label_calhas.setAlignment(Qt.AlignCenter) 

        # Botões para atualizar ZPP260
        self.btn_padrao_lista_zpp260 = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_zpp260 = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_zpp260 = QPushButton("Visualizar ZPP260")
        
        self.btn_padrao_lista_zpp260.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_zpp260(),nome_padrao='Lista_de_zpp260.csv'))
        self.btn_atualizar_lista_zpp260.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_zpp260.json"))
        caminho_zpp260 = os.path.join(os.getcwd(), "data", "Lista_de_zpp260.json")
        self.btn_visualizar_lista_zpp260.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_zpp260))
        self.label_zpp260 = QLabel(f"{self.status_json(caminho_zpp260)}")
        self.label_zpp260.setFixedWidth(30)
        self.label_zpp260.setAlignment(Qt.AlignCenter) 
        

        # Botões para atualizar Mapa
        self.btn_padrao_lista_mapa_corte = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_mapa_corte = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_mapa_corte = QPushButton("Visualizar Mapa")
        
        self.btn_padrao_lista_mapa_corte.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_mapa_corte(),nome_padrao='Lista_de_mapa_corte.csv'))
        self.btn_atualizar_lista_mapa_corte.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_mapa_corte.json"))
        caminho_mapa_corte = os.path.join(os.getcwd(), "data", "Lista_de_mapa_corte.json")
        self.btn_visualizar_lista_mapa_corte.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_mapa_corte))
        self.label_mapa_corte = QLabel(f"{self.status_json(caminho_mapa_corte)}")
        self.label_mapa_corte.setFixedWidth(30)
        self.label_mapa_corte.setAlignment(Qt.AlignCenter) 
        
        # Botões para atualizar Criterios de Qualidade
        self.btn_padrao_lista_de_criterios = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_de_criterios = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_de_criterios = QPushButton("Visualizar Critérios")
        
        self.btn_padrao_lista_de_criterios.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_criterios_Qualidade(),nome_padrao='Lista_de_criterios_Qualidade.csv'))
        self.btn_atualizar_lista_de_criterios.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_criterios_Qualidade.json"))
        caminho_de_criterios = os.path.join(os.getcwd(), "data", "Lista_de_criterios_Qualidade.json")
        self.btn_visualizar_lista_de_criterios.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_de_criterios))
        self.label_de_criterios = QLabel(f"{self.status_json(caminho_de_criterios)}")
        self.label_de_criterios.setFixedWidth(30)
        self.label_de_criterios.setAlignment(Qt.AlignCenter) 
        
        # Botões para atualizar Cabos Legacy
        self.btn_padrao_lista_de_cabos_legacy = QPushButton("Baixar arquivo")
        self.btn_atualizar_lista_de_cabos_legacy = QPushButton("Atualizar Tabela")
        self.btn_visualizar_lista_de_cabos_legacy = QPushButton("Visualizar Códigos Legacy")
    
        self.btn_padrao_lista_de_cabos_legacy.clicked.connect(lambda: self.baixar_dataframe_csv(self.tabela_cabos_legacy(),nome_padrao='Lista_de_cabos_legacy.csv'))
        self.btn_atualizar_lista_de_cabos_legacy.clicked.connect(lambda: self.importar_csv_e_salvar_json(name_arquivo="Lista_de_cabos_legacy.json"))
        caminho_de_cabos_legacy = os.path.join(os.getcwd(), "data", "Lista_de_cabos_legacy.json")
        self.btn_visualizar_lista_de_cabos_legacy.clicked.connect(lambda: self.visualizar_json_como_tabela(caminho_de_cabos_legacy))
        self.label_de_cabos_legacy = QLabel(f"{self.status_json(caminho_de_cabos_legacy)}")
        self.label_de_cabos_legacy.setFixedWidth(30)
        self.label_de_cabos_legacy.setAlignment(Qt.AlignCenter) 
        
        # Layout cabos
        label_cabos = QLabel("Lista de Cabos - SEM")
        tab1_layout.addWidget(label_cabos)
        layout_interno_1 = QHBoxLayout()
        layout_interno_1.addWidget(self.label_cabos)
        layout_interno_1.addWidget(self.btn_padrao_lista_de_cabos)
        layout_interno_1.addWidget(self.btn_atualizar_lista_cabos)        
        layout_interno_1.addWidget(self.btn_visualizar_lista_cabos)
        tab1_layout.addLayout(layout_interno_1)
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Terminais
        label_terminais = QLabel("Lista de Terminais - SEM")
        tab1_layout.addWidget(label_terminais)
        layout_interno_2 = QHBoxLayout()
        layout_interno_2.addWidget(self.label_terminais)
        layout_interno_2.addWidget(self.btn_padrao_lista_de_terminais)
        layout_interno_2.addWidget(self.btn_atualizar_lista_terminais)
        layout_interno_2.addWidget(self.btn_visualizar_lista_terminais)
        tab1_layout.addLayout(layout_interno_2)
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout selos
        label_selos = QLabel("Lista de Selos - SEM")
        tab1_layout.addWidget(label_selos)
        layout_interno_2 = QHBoxLayout()
        layout_interno_2.addWidget(self.label_selos)
        layout_interno_2.addWidget(self.btn_padrao_lista_de_selos)
        layout_interno_2.addWidget(self.btn_atualizar_lista_selos)
        layout_interno_2.addWidget(self.btn_visualizar_lista_selos)
        tab1_layout.addLayout(layout_interno_2)
        tab1_layout.addWidget(self.criar_linha())

        # Layout Máquinas do Corte
        label_maquinas = QLabel("Lista de Máquinas Corte - OEM")
        tab1_layout.addWidget(label_maquinas)
        layout_interno_3 = QHBoxLayout()
        layout_interno_3.addWidget(self.label_maquinas)
        layout_interno_3.addWidget(self.btn_padrao_lista_de_maquinas)
        layout_interno_3.addWidget(self.btn_atualizar_lista_maquinas)
        layout_interno_3.addWidget(self.btn_visualizar_lista_maquinas)
        tab1_layout.addLayout(layout_interno_3)
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Máquinas de Lead Prep
        label_maquinas = QLabel("Lista de Máquinas Lead Prep - OEM")
        tab1_layout.addWidget(label_maquinas)
        layout_interno_4 = QHBoxLayout()
        layout_interno_4.addWidget(self.label_maq_lead_prep)
        layout_interno_4.addWidget(self.btn_padrao_lista_de_maq_lead_prep)
        layout_interno_4.addWidget(self.btn_atualizar_lista_maq_lead_prep)
        layout_interno_4.addWidget(self.btn_visualizar_lista_maq_lead_prep)
        tab1_layout.addLayout(layout_interno_4)
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Setup corte
        label_setup_corte = QLabel("Tempos Setup Corte - Cotação")
        tab1_layout.addWidget(label_setup_corte)
        layout_interno_5 = QHBoxLayout()
        layout_interno_5.addWidget(self.label_setup_corte)
        layout_interno_5.addWidget(self.btn_padrao_lista_de_setup_corte)
        layout_interno_5.addWidget(self.btn_atualizar_lista_setup_corte)
        layout_interno_5.addWidget(self.btn_visualizar_lista_setup_corte)
        tab1_layout.addLayout(layout_interno_5)
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Setup Lead Prep
        label_setup_lead_prep = QLabel("Tempos Setup Lead Prep - Cotação")
        tab1_layout.addWidget(label_setup_lead_prep)
        layout_interno_6 = QHBoxLayout()
        layout_interno_6.addWidget(self.label_setup_lead_prep)
        layout_interno_6.addWidget(self.btn_padrao_lista_de_setup_lead_prep)
        layout_interno_6.addWidget(self.btn_atualizar_lista_setup_lead_prep)
        layout_interno_6.addWidget(self.btn_visualizar_lista_setup_lead_prep)
        tab1_layout.addLayout(layout_interno_6) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Rates Lead Prep
        label_rates_corte = QLabel("Rates Corte - Cotação")
        tab1_layout.addWidget(label_rates_corte)
        layout_interno_7 = QHBoxLayout()
        layout_interno_7.addWidget(self.label_rates_corte)
        layout_interno_7.addWidget(self.btn_padrao_lista_de_rates_corte)
        layout_interno_7.addWidget(self.btn_atualizar_lista_rates_corte)
        layout_interno_7.addWidget(self.btn_visualizar_lista_rates_corte)
        tab1_layout.addLayout(layout_interno_7) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Rates Lead Prep
        label_rates_lead_prep = QLabel("Rates Lead Prep - Cotação")
        tab1_layout.addWidget(label_rates_lead_prep)
        layout_interno_8 = QHBoxLayout()
        layout_interno_8.addWidget(self.label_rates_lead_prep)
        layout_interno_8.addWidget(self.btn_padrao_lista_de_rates_lead_prep)
        layout_interno_8.addWidget(self.btn_atualizar_lista_rates_lead_prep)
        layout_interno_8.addWidget(self.btn_visualizar_lista_rates_lead_prep)
        tab1_layout.addLayout(layout_interno_8) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout ZMM247
        label_zmm247 = QLabel("ZMM247 - SAP")
        tab1_layout.addWidget(label_zmm247)
        layout_interno_9 = QHBoxLayout()
        layout_interno_9.addWidget(self.label_zmm247)
        layout_interno_9.addWidget(self.btn_padrao_lista_zmm247)
        layout_interno_9.addWidget(self.btn_atualizar_lista_zmm247)
        layout_interno_9.addWidget(self.btn_visualizar_lista_zmm247)
        tab1_layout.addLayout(layout_interno_9) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Master kanban
        label_master_kanban = QLabel("Master kanban - EDI")
        tab1_layout.addWidget(label_master_kanban)
        layout_interno_10 = QHBoxLayout()
        layout_interno_10.addWidget(self.label_master_kanban)
        layout_interno_10.addWidget(self.btn_padrao_lista_master_kanban)
        layout_interno_10.addWidget(self.btn_atualizar_lista_master_kanban)
        layout_interno_10.addWidget(self.btn_visualizar_lista_master_kanban)
        tab1_layout.addLayout(layout_interno_10) 
        tab1_layout.addWidget(self.criar_linha())
        
        
        # Layout Aplicadores
        label_aplicadores = QLabel("Lista de Aplicadores")
        tab1_layout.addWidget(label_aplicadores)
        layout_interno_11 = QHBoxLayout()
        layout_interno_11.addWidget(self.label_aplicadores)
        layout_interno_11.addWidget(self.btn_padrao_lista_aplicadores)
        layout_interno_11.addWidget(self.btn_atualizar_lista_aplicadores)
        layout_interno_11.addWidget(self.btn_visualizar_lista_aplicadores)
        tab1_layout.addLayout(layout_interno_11) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Calhas
        label_calhas = QLabel("Lista de Calhas de Selos")
        tab1_layout.addWidget(label_calhas)
        layout_interno_12 = QHBoxLayout()
        layout_interno_12.addWidget(self.label_calhas)
        layout_interno_12.addWidget(self.btn_padrao_lista_calhas)
        layout_interno_12.addWidget(self.btn_atualizar_lista_calhas)
        layout_interno_12.addWidget(self.btn_visualizar_lista_calhas)
        tab1_layout.addLayout(layout_interno_12) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout ZPP260
        label_zpp260 = QLabel("Lista de ZPP260")
        tab1_layout.addWidget(label_zpp260)
        layout_interno_12 = QHBoxLayout()
        layout_interno_12.addWidget(self.label_zpp260)
        layout_interno_12.addWidget(self.btn_padrao_lista_zpp260)
        layout_interno_12.addWidget(self.btn_atualizar_lista_zpp260)
        layout_interno_12.addWidget(self.btn_visualizar_lista_zpp260)
        tab1_layout.addLayout(layout_interno_12) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Mapa
        label_mapa_corte = QLabel("Mapa do Corte")
        tab1_layout.addWidget(label_mapa_corte)
        layout_interno_13 = QHBoxLayout()
        layout_interno_13.addWidget(self.label_mapa_corte)
        layout_interno_13.addWidget(self.btn_padrao_lista_mapa_corte)
        layout_interno_13.addWidget(self.btn_atualizar_lista_mapa_corte)
        layout_interno_13.addWidget(self.btn_visualizar_lista_mapa_corte)
        tab1_layout.addLayout(layout_interno_13) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Criterios
        label_criterios_de_qualidade = QLabel("Critérios de Qualidade")
        tab1_layout.addWidget(label_criterios_de_qualidade)
        layout_interno_14 = QHBoxLayout()
        layout_interno_14.addWidget(self.label_de_criterios)
        layout_interno_14.addWidget(self.btn_padrao_lista_de_criterios)
        layout_interno_14.addWidget(self.btn_atualizar_lista_de_criterios)
        layout_interno_14.addWidget(self.btn_visualizar_lista_de_criterios)
        tab1_layout.addLayout(layout_interno_14) 
        tab1_layout.addWidget(self.criar_linha())
        
        # Layout Cabom com legacy
        label_cabos_legacy = QLabel("Lista de Cabos com Legacy")
        tab1_layout.addWidget(label_cabos_legacy)
        layout_interno_15 = QHBoxLayout()
        layout_interno_15.addWidget(self.label_de_cabos_legacy)
        layout_interno_15.addWidget(self.btn_padrao_lista_de_cabos_legacy)
        layout_interno_15.addWidget(self.btn_atualizar_lista_de_cabos_legacy)
        layout_interno_15.addWidget(self.btn_visualizar_lista_de_cabos_legacy)
        tab1_layout.addLayout(layout_interno_15) 
        tab1_layout.addWidget(self.criar_linha())
        

        tab1_layout.addStretch()
        #tab1.setLayout(tab1_layout)
    
        container = QWidget()
        container.setLayout(tab1_layout)

        scroll.setWidget(container)

        layout_tab = QVBoxLayout()
        layout_tab.addWidget(scroll)

        tab1.setLayout(layout_tab)
        #-----------------------------------------------------------------------------------------------------------------------------
        # ---------- Aba toolkit ----------
        self.tab_toolkit = QWidget()

        layout_principal = QVBoxLayout()

        # Layout horizontal para os grupos
        layout_grupos = QHBoxLayout()

        # Criar os GroupBox
        group_box1 = QGroupBox("Informações1")
        group_box2 = QGroupBox("Informações2")
        group_box3 = QGroupBox("Informações3")

        # Layout interno de cada grupo
        layout_g1 = QVBoxLayout()
        layout_g2 = QVBoxLayout()
        layout_g3 = QVBoxLayout()

        # Botões
        layout_g1.addWidget(QPushButton("Botão 1.1"))
        layout_g1.addWidget(QPushButton("Botão 1.2"))
        layout_g1.addStretch() 

        layout_g2.addWidget(QPushButton("Botão 2.1"))
        layout_g2.addWidget(QPushButton("Botão 2.2"))
        layout_g2.addStretch() 

        layout_g3.addWidget(QPushButton("Botão 3.1"))
        layout_g3.addWidget(QPushButton("Botão 3.2"))
        layout_g3.addStretch() 

        # Aplicar layouts internos
        group_box1.setLayout(layout_g1)
        group_box2.setLayout(layout_g2)
        group_box3.setLayout(layout_g3)

        # Adicionar os grupos no layout horizontal
        layout_grupos.addWidget(group_box1)
        layout_grupos.addWidget(group_box2)
        layout_grupos.addWidget(group_box3)

        # Adicionar ao layout principal
        layout_principal.addLayout(layout_grupos)

        # Aplicar na aba
        self.tab_toolkit.setLayout(layout_principal)

        # ---------- Aba 2 ----------
        self.tab2 = QWidget()
        tab2_layout = QVBoxLayout()
        self.tab2.setLayout(tab2_layout)

        # Botão para carregar arquivo
        btn_carregar = QPushButton("Tool list")
        btn_carregar.setFixedWidth(150)
        btn_carregar.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        
        btn_converte_lista = QPushButton("Converte para lista")
        btn_converte_lista.setFixedWidth(150)
        btn_converte_lista.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        # Label com status

        # Layout horizontal para botão e label juntos
        layout_h_1 = QHBoxLayout()
        layout_h_1.addWidget(btn_carregar, alignment=Qt.AlignmentFlag.AlignLeft)
        
        
        layout_h_2 = QHBoxLayout()
        layout_h_2.addWidget(btn_converte_lista, alignment=Qt.AlignmentFlag.AlignLeft)
        
        # Checkbox
        self.checkbox_processo = QCheckBox("1. Definir Processos")
        self.checkbox_processo.stateChanged.connect(self.adicionar_processos)
        
        self.checkbox_seq = QCheckBox("2. Sequenciar")
        self.checkbox_seq.stateChanged.connect(self.adicionar_seq)
        
        self.checkbox_volume = QCheckBox("3. Volumes")
        self.checkbox_volume.stateChanged.connect(self.adicionar_volume)
        
    
        layout_h_3 = QHBoxLayout()
        layout_h_3.addWidget(self.checkbox_processo, alignment=Qt.AlignmentFlag.AlignLeft)

        layout_h_4 = QHBoxLayout()
        layout_h_4.addWidget(self.checkbox_seq, alignment=Qt.AlignmentFlag.AlignLeft)
        
        layout_h_5 = QHBoxLayout()
        layout_h_5.addWidget(self.checkbox_volume, alignment=Qt.AlignmentFlag.AlignLeft)
        
        # Tabela para mostrar dados
        self.tabela_corte = QTableWidget()
        
        
        # --- Filtro ---
        layout_filtro = QHBoxLayout()

        self.combo_coluna = QComboBox()
        self.input_valor = QLineEdit()
        self.input_valor.setPlaceholderText("Digite o valor...")
        self.input_valor.setFixedWidth(120)

        self.btn_filtrar = QPushButton("Filtrar")

        layout_filtro.addWidget(QLabel("Coluna:"))
        layout_filtro.addWidget(self.combo_coluna)
        layout_filtro.addWidget(QLabel("Valor:"))
        layout_filtro.addWidget(self.input_valor)
        layout_filtro.addWidget(self.btn_filtrar)
        layout_filtro.addStretch()

        self.btn_filtrar.clicked.connect(self.filtrar_dataframe)

        # Adiciona os layouts/widgets na ordem
        tab2_layout.addLayout(layout_h_1)
        tab2_layout.addWidget(self.criar_linha())
        tab2_layout.addLayout(layout_h_2)
        tab2_layout.addLayout(layout_h_3)
        tab2_layout.addLayout(layout_h_4)
        tab2_layout.addLayout(layout_h_5)
        tab2_layout.addLayout(layout_filtro)   
        tab2_layout.addWidget(self.criar_linha())
        tab2_layout.addWidget(self.tabela_corte)

        # Conecta o botão ao método de carregar arquivo
        btn_carregar.clicked.connect(self.carregar_arquivo)
        btn_converte_lista.clicked.connect(lambda: self.converter_sap())
        
        # ---------- Aba 10 ----------
        self.tab10 = QWidget()
        label_erros = QLabel("Logs de erros")
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)  # Define como linha horizontal
        line.setFrameShadow(QFrame.Shadow.Sunken) # Define a sombra (estilo 3D)
        line.setLineWidth(1) # Define a largura da linha


        self.label_erros_report = QTextEdit()
        self.label_erros_report.setReadOnly(True)
        self.label_erros_report.setLineWrapMode(QTextEdit.NoWrap)  # não quebra as linhas
        self.label_erros_report.setMinimumHeight(600)
        self.label_erros_report.setSizePolicy(QSizePolicy.Policy.Expanding,
                                            QSizePolicy.Policy.Expanding)
        self.label_erros_report.setFontFamily("Courier New")       # fonte monoespaçada
        tab10_layout = QVBoxLayout()
        tab10_layout.addWidget(label_erros)
        tab10_layout.addWidget(line)
        tab10_layout.addWidget(self.label_erros_report)
        tab10_layout.addStretch()
        self.tab10.setLayout(tab10_layout)
        
        #----------------------------

        tabs.addTab(tab1, "Basic Information")
        tabs.addTab(self.tab_toolkit, "Toolkit")
        tabs.addTab(self.tab2, "Lista de Circuitos - Corte")
        tabs.addTab(self.tab10, "Logs Erros")
        main_layout.addWidget(tabs)


    @log_errors
    def criar_linha(self):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setLineWidth(1)
        return line

    @log_errors
    def komax_changed(self, value):
        try:
            self.vision_komax = value
            print(f"Vision Komax atualizado para {self.vision_komax}%")
        except Exception as e:
            raise ValueError(f'Err: {e}')
        
    @log_errors
    def schleuniger_changed(self, value):
        try:    
            self.vision_schleuniger = value
            print(f"Vision Schleuniger atualizado para {self.vision_schleuniger}%")
        except Exception as e:
            raise ValueError(f'Err: {e}')
    
    @log_errors
    def load_json_to_table(self):
        try:
            caminho_de_criterios = os.path.join(os.getcwd(), "data", "Lista_de_criterios_Qualidade.json")
            with open(caminho_de_criterios, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            # Assumindo que data é uma lista de dicionários
            if isinstance(data, list) and data:
                headers = list(data[0].keys())
                self.table.setColumnCount(len(headers))
                self.table.setHorizontalHeaderLabels(headers)
                self.table.setRowCount(len(data))

                for row_idx, row_data in enumerate(data):
                    for col_idx, key in enumerate(headers):
                        self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(row_data[key])))
                        
                self.table.setSortingEnabled(True)
                self.table.setAlternatingRowColors(True)
                self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
                self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
                self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
                self.table.horizontalHeader().setStretchLastSection(True)
        except Exception as e:
            raise ValueError(f"Erro ao carregar JSON: {e}")

    @log_errors
    def atualizar_tabela(self, df):
        try:
            self.tabela_corte.setSortingEnabled(False)

            self.tabela_corte.clear()
            self.tabela_corte.setRowCount(len(df))
            self.tabela_corte.setColumnCount(len(df.columns))
            self.tabela_corte.setHorizontalHeaderLabels(df.columns.tolist())

            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    valor = df.iat[row_idx, col_idx]
                    self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

            self.tabela_corte.setSortingEnabled(True)
            self.tabela_corte.setAlternatingRowColors(True)
            self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
            self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            self.tabela_corte.horizontalHeader().setStretchLastSection(True)
        except Exception as e:
            raise ValueError(f'{e}')
            
    @log_errors
    def filtrar_dataframe(self):
        try:
            coluna = self.combo_coluna.currentText()
            valor = self.input_valor.text()

            if coluna and valor:
                self.df_filtrado = self.lista_do_corte[self.lista_do_corte[coluna].astype(str).str.contains(valor, case=False, na=False)]
                self.atualizar_tabela(self.df_filtrado)
        except Exception as e:
            raise ValueError(f'{e}')
    
    @log_errors
    def carregar_arquivo(self):
        filtro = "Arquivos Excel (*.xlsx *.XLSX)"
        caminho, _ = QFileDialog.getOpenFileName(self.tab2, "Selecionar arquivo Excel", "", filtro)
        if not caminho:
            return

        try:
            df = pd.read_excel(caminho,dtype=str)
            self.arquivo_sap = df.copy()
        except Exception as e:
            QMessageBox.critical(self.tab2, "Erro", f"Falha ao abrir arquivo Excel:\n{e}")
            raise ValueError(f'{e}')

        self.combo_coluna.clear()
        self.combo_coluna.addItems(self.arquivo_sap.columns.tolist())

        self.tabela_corte.setSortingEnabled(False)
        self.tabela_corte.clear()
        self.tabela_corte.setRowCount(len(self.arquivo_sap))
        self.tabela_corte.setColumnCount(len(self.arquivo_sap.columns))
        self.tabela_corte.setHorizontalHeaderLabels(self.arquivo_sap.columns.tolist())

        for row_idx in range(len(self.arquivo_sap)):
            for col_idx in range(len(self.arquivo_sap.columns)):
                valor = self.arquivo_sap.iat[row_idx, col_idx]
                self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

        self.tabela_corte.setSortingEnabled(True)
        self.tabela_corte.setAlternatingRowColors(True)
        self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tabela_corte.horizontalHeader().setStretchLastSection(True)
                                
    @log_errors
    def status_json(self, caminho):
        if not os.path.exists(caminho):
            return "❌"

        if os.path.getsize(caminho) == 0:
            return "❌"

        try:
            with open(caminho, "r", encoding="utf-8") as f:
                data = json.load(f)
                return "✔️" if data else "❌"
        except:
            return "❌"
    
    @log_errors
    def tabela_cabos(self):
        # Caminho do arquivo JSON
        caminho_terminais = os.path.join(os.getcwd(), "data", "Lista_de_cabos.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_terminais):
            # Cria DataFrame vazio com colunas padrão
            arquivo = pd.DataFrame(
                columns=['Part Number', 'Part Classification',
                         'CS Part', 'Wire Size', 'Temperature Rating', 
                         'Number of Strands', 'Nominal Insulation Thickness', 
                         'Primary Color','Secondary Color', 'Construction', 
                         'Outer Diameter']
            )
            return arquivo

        try:
            # Lê o JSON
            df = pd.read_json(caminho_terminais)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(columns=['Part Number', 'Part Classification',
                         'CS Part', 'Wire Size', 'Temperature Rating', 
                         'Number of Strands', 'Nominal Insulation Thickness', 
                         'Primary Color','Secondary Color', 'Construction', 
                         'Outer Diameter'])

            return df

        except Exception as e:
            # Caso dê algum erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(columns=['Part Number', 'Part Classification',
                         'CS Part', 'Wire Size', 'Temperature Rating', 
                         'Number of Strands', 'Nominal Insulation Thickness', 
                         'Primary Color','Secondary Color', 'Construction', 
                         'Outer Diameter'])

    @log_errors
    def tabela_terminais(self):
        # Caminho do arquivo JSON
        caminho_terminais = os.path.join(os.getcwd(), "data", "Lista_de_terminais.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_terminais):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Part Number', 'Connection Technology', 
                         'Terminal Size (Male or Female only)', 
                         'Min Wire Size (mm^2)', 'Max Wire Size (mm^2)', 
                         'Feed Type/Delivery Form', 'Accepts Seal?', 
                         'Terminal Style (Male or Female only)'
                         ]
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_terminais)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Part Number', 'Connection Technology', 
                         'Terminal Size (Male or Female only)', 
                         'Min Wire Size (mm^2)', 'Max Wire Size (mm^2)', 
                         'Feed Type/Delivery Form', 'Accepts Seal?', 
                         'Terminal Style (Male or Female only)'
                         ]
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Part Number', 'Connection Technology', 
                         'Terminal Size (Male or Female only)', 
                         'Min Wire Size (mm^2)', 'Max Wire Size (mm^2)', 
                         'Feed Type/Delivery Form', 'Accepts Seal?', 
                         'Terminal Style (Male or Female only)'
                         ]
            )
            
    @log_errors
    def tabela_selos(self):
        # Caminho do arquivo JSON
        caminho_selos = os.path.join(os.getcwd(), "data", "Lista_de_selos.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_selos):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=[
                    'Part Number',
                    'Connection Technology',
                    'Feed Type/Delivery Form'
                ]
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_selos)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=[
                        'Part Number',
                        'Connection Technology',
                        'Feed Type/Delivery Form'
                    ]
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=[
                    'Part Number',
                    'Connection Technology',
                    'Feed Type/Delivery Form'
                ]
            )
            
    @log_errors
    def tabela_maquinas(self):
        # Caminho do arquivo JSON
        caminho_maquinas_corte = os.path.join(os.getcwd(), "data", "Lista_de_maquinas.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_maquinas_corte):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=["Maqs", "OEM", "Model", "Process", "Length", "Time Batch(s)",'ConveyorLength','Vision System',
                         'Open Ends','Twisting','Length Opens Ends','Projeto','Min Sectionn','Max Sectionn','Min Length','Max Length', 'LGK-CC']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_maquinas_corte)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=["Maqs", "OEM", "Model", "Process", "Length", "Time Batch(s)",'ConveyorLength','Vision System',
                             'Open Ends','Twisting','Length Opens Ends','Projeto','Min Sectionn','Max Sectionn','Min Length','Max Length', 'LGK-CC']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=["Maqs", "OEM", "Model", "Process", "Length", "Time Batch(s)",'ConveyorLength','Vision System',
                         'Open Ends','Twisting','Length Opens Ends','Projeto','Min Sectionn','Max Sectionn','Min Length','Max Length', 'LGK-CC']
            )

    @log_errors
    def tabela_maq_lead_prep(self):
        # Caminho do arquivo JSON
        caminho_maquinas_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_maq_lead_prep.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_maquinas_lead_prep):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=[
                    'Part Number',
                    'Machine Type',
                    'Capacity'
                ]
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_maquinas_lead_prep)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=[
                        'Part Number',
                        'Machine Type',
                        'Capacity'
                    ]
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=[
                    'Part Number',
                    'Machine Type',
                    'Capacity'
                ]
            )

    @log_errors
    def tabela_setup_lead_prep(self):
        # Caminho do arquivo JSON
        caminho_setup_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_setup_lead_prep.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_setup_lead_prep):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_setup_lead_prep)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            )
            
    @log_errors
    def tabela_setup_corte(self):
        # Caminho do arquivo JSON
        caminho_setup_corte = os.path.join(os.getcwd(), "data", "Lista_de_setup_corte.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_setup_corte):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_setup_corte)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            )
            
    @log_errors
    def tabela_rates_corte(self):
        # Caminho do arquivo JSON
        caminho_rates_corte = os.path.join(os.getcwd(), "data", "Lista_de_rates_corte.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_rates_corte):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['ID','Model','Process','LClass','BatchTime (h)','StdTime OEM','Update','Rate_Global_Std']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_rates_corte)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['ID','Model','Process','LClass','BatchTime (h)','StdTime OEM','Update','Rate_Global_Std']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['ID','Model','Process','LClass','BatchTime (h)','StdTime OEM','Update','Rate_Global_Std']
            )
            
    @log_errors
    def tabela_rates_lead_prep(self):
        # Caminho do arquivo JSON
        caminho_rates_maquina_lead_prep = os.path.join(os.getcwd(), "data", "Lista_de_rates_lead_prep.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_rates_maquina_lead_prep):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Descrição', 'ID', 'Comprim.', 'Atributte', 'StdTime', 'StdTime AS']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_rates_maquina_lead_prep)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Descrição', 'ID', 'Comprim.', 'Atributte', 'StdTime', 'StdTime AS']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Descrição', 'ID', 'Comprim.', 'Atributte', 'StdTime', 'StdTime AS']
            )    
            
    @log_errors
    def tabela_zmm247(self):
        # Caminho do arquivo JSON
        caminho_zmm247 = os.path.join(os.getcwd(), "data", "Lista_de_zmm247.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_zmm247):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=[
                    'Plant',
                    'Internal Family',
                    'External Family'
                ]
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_zmm247)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=[
                        'Plant',
                        'Internal Family',
                        'External Family'
                    ]
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=[
                    'Part Number',
                    'Machine Type',
                    'Capacity'
                ]
            )
    
    @log_errors
    def tabela_master_kanban(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_master_kanban.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Derivativos CARGA','Famílias','Projeto','Código UCS','Total week']

            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Derivativos CARGA','Famílias','Projeto','Código UCS','Total week']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Derivativos CARGA','Famílias','Projeto','Código UCS','Total week']

            )
     
    @log_errors
    def tabela_aplicadores(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_aplicadores.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Terminal','Fornecedor','Código SAP','Projeto']

            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Terminal','Fornecedor','Código SAP','Projeto']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Terminal','Fornecedor','Código SAP','Projeto']

            )
                   
    @log_errors
    def tabela_calhas(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_calhas.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Selo','Fornecedor','Código Sap','Projeto']

            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Selo','Fornecedor','Código Sap','Projeto']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Selo','Fornecedor','Código Sap','Projeto']

            )
            
    @log_errors
    def tabela_zpp260(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_zpp260.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['WERKS','CIRC_MASTER','CIRC_COMUNS']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['WERKS','CIRC_MASTER','CIRC_COMUNS']
                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['WERKS','CIRC_MASTER','CIRC_COMUNS']
            )
            
    @log_errors
    def tabela_mapa_corte(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_mapa_corte.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Projeto','Leadset','Alocação']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Projeto','Leadset','Alocação']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Projeto','Leadset','Alocação']

            )
                    
    @log_errors
    def tabela_criterios_Qualidade(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_criterios_Qualidade.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Máquina',	'SmartDetect',	'WireCam (DECAPE)',	
                         'WireCam (SELO)', 'CFM', 'VisionSystem',	
                         'Double Cutting',	'Wire size', 'Other']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Máquina',	'SmartDetect',	'WireCam (DECAPE)',	
                         'WireCam (SELO)', 'CFM', 'VisionSystem',	
                         'Double Cutting',	'Wire size', 'Other']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Projeto','Leadset','Alocação']

            )
                    
    @log_errors
    def tabela_cabos_legacy(self):
        # Caminho do arquivo JSON
        caminho_master_kanban = os.path.join(os.getcwd(), "data", "Lista_de_cabos_legacy.json")

        # Verifica se o arquivo existe
        if not os.path.exists(caminho_master_kanban):
            # Cria DataFrame vazio com colunas padrão
            df = pd.DataFrame(
                columns=['Part Number','Legacy']
            )
            return df

        try:
            # Lê o JSON
            df = pd.read_json(caminho_master_kanban)

            # Se estiver vazio, retorna DataFrame com colunas padrão
            if df.empty:
                df = pd.DataFrame(
                    columns=['Part Number','Legacy']

                )

            return df

        except Exception as e:
            # Caso haja erro na leitura do JSON
            print(f"Erro ao ler o JSON: {e}")
            return pd.DataFrame(
                columns=['Part Number','Legacy']

            )
            
    @log_errors
    def baixar_dataframe_csv(self, df: pd.DataFrame, nome_padrao:str):

        caminho, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar arquivo CSV",
            QDir.homePath() + "/" + nome_padrao,  # 👈 nome padrão aqui
            "Arquivo CSV (*.csv)"
        )

        if caminho:
            try:
                if not caminho.lower().endswith(".csv"):
                    caminho += ".csv"

                df.to_csv(caminho, index=False, encoding="utf-8-sig", sep=';')

                QMessageBox.information(self, "Sucesso", "Arquivo salvo com sucesso!")

            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")
    
    @log_errors
    def importar_csv_e_salvar_json(self, name_arquivo: str):
        # 1️⃣ Selecionar CSV
        caminho_csv, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo CSV",
            "",
            "Arquivo CSV (*.csv)"
        )

        if not caminho_csv:
            raise ValueError("Seleção de arquivo cancelada pelo usuário")

        # 2️⃣ Ler CSV
        try:
            df = pd.read_csv(caminho_csv, sep=";",dtype=str)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler CSV:\n{e}")
            raise

        validar_import = self.validar_import(dados=df, name_arquivo=name_arquivo)
        if not validar_import:
            raise ValueError("Seleção de arquivo cancelada pelo usuário")

        # 3️⃣ Criar pasta de destino
        try:
            pasta_destino = os.path.join(os.getcwd(), "data")
            os.makedirs(pasta_destino, exist_ok=True)

            caminho_json = os.path.join(pasta_destino, name_arquivo)

            # 4️⃣ Salvar JSON
            df.to_json(
                caminho_json,
                orient="records",
                indent=4,
                force_ascii=False
            )

            QMessageBox.information(
                self,
                "Sucesso",
                f"Base salva automaticamente em:\n{caminho_json}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar JSON:\n{e}")
            raise ValueError(f"Erro ao salvar JSON:\n{e}")
    
    @log_errors
    def visualizar_json_como_tabela(self, caminho_json):
        if not os.path.exists(caminho_json):
            QMessageBox.critical(self, "Erro", "Arquivo não encontrado.")
            return

        try:
            # Lê JSON para DataFrame
            df = pd.read_json(caminho_json)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler arquivo:\n{e}")
            return

        # Cria tabela
        tabela = QTableWidget()
        tabela.setRowCount(len(df))
        tabela.setColumnCount(len(df.columns))
        tabela.setHorizontalHeaderLabels(df.columns.tolist())

        # Preenche a tabela
        for i, (_, row) in enumerate(df.iterrows()):
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(row[col]))
                tabela.setItem(i, j, item)

        # Configurações
        tabela.setSortingEnabled(True)
        tabela.setAlternatingRowColors(True)
        tabela.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tabela.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        tabela.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        # Ajuste de colunas
        tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # Janela
        tabela.setWindowTitle(os.path.basename(caminho_json).replace('.json','').replace('_',' '))
        tabela.resize(600, 600)
        tabela.show()

        # Mantém referência
        self._tabela_json = tabela
             
    @log_errors
    def validar_import(self, dados:pd.DataFrame=None, name_arquivo:str=None):
        
        if name_arquivo== "Lista_de_cabos.json":
            cols = ['Part Number', 'Part Classification',
                    'CS Part', 'Wire Size', 'Temperature Rating',
                    'Number of Strands', 'Nominal Insulation Thickness',
                    'Primary Color','Secondary Color', 'Construction', 
                    'Outer Diameter']
            if list(dados.columns) == cols:
                return True
            else:
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_terminais.json":
            cols = ['Part Number', 'Connection Technology', 
                    'Terminal Size (Male or Female only)', 
                    'Min Wire Size (mm^2)', 'Max Wire Size (mm^2)', 
                    'Feed Type/Delivery Form', 'Accepts Seal?', 
                    'Terminal Style (Male or Female only)']
            if list(dados.columns) == cols: return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_zmm247.json":
            cols = ['Plant', 'Internal Family', 'External Family']
            if list(dados.columns) == cols: return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_maquinas.json":
            cols = ["Maqs", "OEM", "Model", "Process", "Length", "Time Batch(s)",'ConveyorLength','Vision System',
                    'Open Ends','Twisting','Length Opens Ends','Projeto','Min Sectionn','Max Sectionn','Min Length','Max Length', 'LGK-CC']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_rates_corte.json":
            cols = ['ID','Model','Process','LClass','BatchTime (h)','StdTime OEM','Update','Rate_Global_Std']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_rates_lead_prep.json":
            cols = ['Descrição', 'ID', 'Comprim.', 'Atributte', 'StdTime', 'StdTime AS']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_setup_lead_prep.json":
            cols = ['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_setup_corte.json":
            cols = ['Setup','Tempo(H)','Tempo(m) (setup target NVT)']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_master_kanban.json":
            cols = ['Derivativos CARGA','Famílias','Projeto','Código UCS','Total week']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_calhas.json":
            cols = ['Part number','Fornecedor','Código Sap','Projeto']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_aplicadores.json":
            cols = ['Terminal','Fornecedor','Código SAP','Projeto']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_zpp260.json":
            cols = ['WERKS','CIRC_MASTER','CIRC_COMUNS']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_mapa_corte.json":
            cols = ['Projeto','Leadset','Alocação']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        
        elif name_arquivo == "Lista_de_criterios_Qualidade.json":
            cols = ['Máquina',	'SmartDetect',	'WireCam (DECAPE)',	
                         'WireCam (SELO)', 'CFM', 'VisionSystem',	
                         'Double Cutting',	'Wire size', 'Other']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
            
        elif name_arquivo == "Lista_de_cabos_legacy.json":
            cols = ['Part Number','Legacy']
            if list(dados.columns) == cols: 
                return True
            else: 
                faltando = set(cols) - set(dados.columns)
                extras = set(dados.columns) - set(cols)

                texto_erro = (
                    "JSON com estrutura inválida para esta tabela.\n"
                    f"{'Colunas esperadas':<25}: {cols}\n"
                    f"{'Colunas recebidas':<25}: {dados.columns.tolist()}\n"
                    f"{'Faltando':<25}: {list(faltando)}\n"
                    f"{'Extras':<25}: {list(extras)}"
                )
                QMessageBox.critical(self, "Erro", texto_erro)
                raise ValueError(texto_erro)
        else:
            return False

    #=============================================================================
    @log_errors
    def converter_sap(self):
        try:
            caminho_zpp260 = os.path.join(os.getcwd(), "data", "Lista_de_zpp260.json")
            df_cmz = pd.read_json(caminho_zpp260)
            self.lista_do_corte = converte_arquivo_sap(dados=self.arquivo_sap, dados_cmz=df_cmz)
            
            self.combo_coluna.clear()
            self.combo_coluna.addItems(self.lista_do_corte.columns.tolist())
            
            self.tabela_corte.setSortingEnabled(False)
            self.tabela_corte.clear()
            self.tabela_corte.setRowCount(len(self.lista_do_corte))
            self.tabela_corte.setColumnCount(len(self.lista_do_corte.columns))
            self.tabela_corte.setHorizontalHeaderLabels(self.lista_do_corte.columns.tolist())

            for row_idx in range(len(self.lista_do_corte)):
                for col_idx in range(len(self.lista_do_corte.columns)):
                    valor = self.lista_do_corte.iat[row_idx, col_idx]
                    self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

            self.tabela_corte.setSortingEnabled(True)
            self.tabela_corte.setAlternatingRowColors(True)
            self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
            self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            self.tabela_corte.horizontalHeader().setStretchLastSection(True)
            
            QApplication.processEvents()
            QMessageBox.information(self, "Sucesso", f"Análise feita!!!")
        except Exception as e:
            raise ValueError(f'Err: {e}')
        
    @log_errors  
    def adicionar_seq(self, checked_seq):
        try:
            if checked_seq:
                self.lista_do_corte = adicionar_sequencia(df=self.lista_do_corte)
                QApplication.processEvents()
                QMessageBox.information(self, "Sucesso", f"Análise feita!!!")
            else:
                self.lista_do_corte = self.lista_do_corte.drop(columns=['TermA_uso', 'TermB_uso', 'SEALA_uso','SEALB_uso',
                                                                        'Seq.','Processo','LClass','Alocações','Bundle size'], errors='ignore')

            self.combo_coluna.clear()
            self.combo_coluna.addItems(self.lista_do_corte.columns.tolist())

            self.tabela_corte.setSortingEnabled(False)

            self.tabela_corte.clear()
            self.tabela_corte.setRowCount(len(self.lista_do_corte))
            self.tabela_corte.setColumnCount(len(self.lista_do_corte.columns))
            self.tabela_corte.setHorizontalHeaderLabels(self.lista_do_corte.columns.tolist())

            for row_idx in range(len(self.lista_do_corte)):
                for col_idx in range(len(self.lista_do_corte.columns)):
                    valor = self.lista_do_corte.iat[row_idx, col_idx]
                    self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

            self.tabela_corte.setSortingEnabled(True)
            self.tabela_corte.setAlternatingRowColors(True)
            self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
            self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            self.tabela_corte.horizontalHeader().setStretchLastSection(True)
        except Exception as e:
            raise ValueError(f'Err: {e}')
        
    @log_errors  
    def adicionar_processos(self, checked_processos):
        try:
            if checked_processos:
                self.lista_do_corte = definir_processos(dados=self.lista_do_corte)
                QApplication.processEvents()
                QMessageBox.information(self, "Sucesso", f"Análise feita!!!")
            else:
                self.lista_do_corte = self.lista_do_corte.drop(columns=['Processo_A','Processo_B'], errors='ignore')

            self.combo_coluna.clear()
            self.combo_coluna.addItems(self.lista_do_corte.columns.tolist())

            self.tabela_corte.setSortingEnabled(False)

            self.tabela_corte.clear()
            self.tabela_corte.setRowCount(len(self.lista_do_corte))
            self.tabela_corte.setColumnCount(len(self.lista_do_corte.columns))
            self.tabela_corte.setHorizontalHeaderLabels(self.lista_do_corte.columns.tolist())

            for row_idx in range(len(self.lista_do_corte)):
                for col_idx in range(len(self.lista_do_corte.columns)):
                    valor = self.lista_do_corte.iat[row_idx, col_idx]
                    self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

            self.tabela_corte.setSortingEnabled(True)
            self.tabela_corte.setAlternatingRowColors(True)
            self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
            self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            self.tabela_corte.horizontalHeader().setStretchLastSection(True)
        except Exception as e:
            raise ValueError(f'Err: {e}')
        
    @log_errors  
    def adicionar_volume(self, checked_volume):
        try:
            if checked_volume:
                self.lista_do_corte = add_volumes(dados=self.lista_do_corte)
                QApplication.processEvents()
                QMessageBox.information(self, "Sucesso", f"Análise feita!!!")
            else:
                self.lista_do_corte = self.lista_do_corte.drop(columns=['Volumes','Comunizados','Vol/dia'], errors='ignore')

            self.combo_coluna.clear()
            self.combo_coluna.addItems(self.lista_do_corte.columns.tolist())
           
            self.tabela_corte.setSortingEnabled(False)
            self.tabela_corte.clear()
            self.tabela_corte.setRowCount(len(self.lista_do_corte))
            self.tabela_corte.setColumnCount(len(self.lista_do_corte.columns))
            self.tabela_corte.setHorizontalHeaderLabels(self.lista_do_corte.columns.tolist())

            for row_idx in range(len(self.lista_do_corte)):
                for col_idx in range(len(self.lista_do_corte.columns)):
                    valor = self.lista_do_corte.iat[row_idx, col_idx]
                    self.tabela_corte.setItem(row_idx, col_idx, QTableWidgetItem(str(valor)))

            self.tabela_corte.setSortingEnabled(True)
            self.tabela_corte.setAlternatingRowColors(True)
            self.tabela_corte.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self.tabela_corte.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
            self.tabela_corte.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            self.tabela_corte.horizontalHeader().setStretchLastSection(True)
            
        except Exception as e:
            raise ValueError(f'Err: {e}')
        
    @log_errors
    def salvar_csv(self, dados: pd.DataFrame, name_arquivo: str = "arquivo.csv"):
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar CSV",
                name_arquivo,
                "Arquivos CSV (*.csv)"
            )

            if not file_path:
                return

            data_formatada = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            if file_path.lower().endswith(".csv"):
                file_path = file_path[:-4] + f"_{data_formatada}.csv"
            else:
                file_path += f"_{data_formatada}.csv"

            dados.to_csv(
                file_path,
                index=False,
                sep=';',
                encoding='utf-8-sig',
                float_format='%.4f',
                decimal=','
            )

            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{file_path}")

        except Exception as e:
            raise RuntimeError(f"Erro ao salvar CSV: {e}") from e
             
    @log_errors
    def salvar_xlsx(self, dados: pd.DataFrame, name_arquivo: str = "arquivo.xlsx"):
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Excel",
                name_arquivo,
                "Arquivos Excel (*.xlsx)"
            )

            if not file_path:
                return

            #data_formatada = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            data_formatada = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            if file_path.lower().endswith(".xlsx"):
                file_path = file_path[:-5] + f"_{data_formatada}.xlsx"
            else:
                file_path += f"_{data_formatada}.xlsx"

            dados.to_excel(file_path, index=False)

            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{file_path}")

        except Exception as e:
            raise RuntimeError(f"Erro ao salvar arquivo: {e}") from e
        
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())