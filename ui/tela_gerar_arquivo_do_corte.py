import sys
import os
import pandas as pd
import numpy as np
import re
from collections import Counter
from datetime import date

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QLabel, QLineEdit,
    QFrame, QStyle, QProgressDialog, QListWidget,
    QMessageBox, QAbstractItemView
)

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QGuiApplication, QCursor

root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)

# Agora você pode importar normalmente
from utils.function import *
# =========================================================
# FUNCÕES
# =========================================================
def converter_sap(caminho_path: str = None):
    
    if caminho_path.lower().endswith(".xlsx"):
        df = pd.read_excel(caminho_path, dtype=str)

    cols = [
        "INNER_1",
        "INNER_2",
        "INNER_3",
        "INNER_4",
        "INNER_5",
        "INNER_6",
        "INNER_7",
        "INNER_8",
        "INNER_9",
    ]

    dict_processos = {}
    dict_comp = {}
    for c in cols:
        for j in df.index:
            ckt = df.loc[j, c]
            process = df.loc[j, "CIRCUIT"]
            if pd.notna(ckt):
                ucs = df.loc[j, "Internal Family"]
                dict_processos[f"{ucs}{ckt}"] = f"{ucs}{process}"
                dict_comp[f"{ucs}{process}"] = str(df.loc[j, "LENGTH"]).strip()

    for i in df.index:
        ckt = df.loc[i, "CIRCUIT"]
        ucs = df.loc[i, "Internal Family"]
        
        leadset = dict_processos.get(f"{ucs}{ckt}", f"{ucs}{ckt}")

        df.loc[i, "Leadset"] = leadset
        df.loc[i, "Comp TW"] = dict_comp.get(leadset)

    df_temp = df[df[cols].notna().any(axis=1)]
    df = df[df[cols].isna().all(axis=1)].reset_index(drop=True)

    cols = ["RMKS_A", "RMKS_B"]

    df[cols] = df[cols].applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df[cols] = df[cols].replace({
        'LGKCC': None,
        'BARE1': None,
        'BARE2': None,
        'BARE3': None,
        'Z': None,
        'BT':None,
        'PV':None
    })

    # mask = df["RMKS_A"].str.startswith("M", na=False)
    # df.loc[mask, "RMKS_A"] = pd.NA
    
    # mask = df["RMKS_B"].str.startswith("M", na=False)
    # df.loc[mask, "RMKS_B"] = pd.NA
    
    
    # mask = df["RMKS_A"].str.startswith("T", na=False)
    # df.loc[mask, "RMKS_A"] = pd.NA
    
    # mask = df["RMKS_B"].str.startswith("T", na=False)
    # df.loc[mask, "RMKS_B"] = pd.NA

    df = df.replace({pd.NA: np.nan, None: np.nan})

    cols = ['STRIP_A', 'STRIP_B']

    df[cols] = (
        df[cols]
        .convert_dtypes()
        .apply(lambda s: s.str.strip())
    )


    return df, df_temp

def gerar_comunizacao_arquivo_sap(dados: str = None, ordem: list[str] = None) -> pd.DataFrame:


    dtype_cat = pd.api.types.CategoricalDtype(categories=ordem, ordered=True)

    dados["Internal Family Categoria"] = dados["Internal Family"].astype(dtype_cat)

    # Ordena mantendo valores fora da lista no final
    dados = (
        dados.sort_values(["Internal Family Categoria", "Internal Family"])
        .drop(columns="Internal Family Categoria")
        .reset_index(drop=True)
    )

   
    colunas = ['WIRE_TUBE_SPLICE', 'LENGTH', 'Comp TW','SECTIONN','TERM_A', 'STRIP_A', 'SEAL_A','TERM_B', 'STRIP_B', 'SEAL_B','RMKS_A','RMKS_B']


    #--------------------------------------------------------------------------------------
    def gerar_chave_grupo(df):
        linhas = []
        
        for _, row in df.iterrows():
            # mesma lógica que você já usava na linha
            linha = tuple(sorted((str(k), v) for k, v in Counter(row).items()))
            linhas.append(linha)
        
        # ordena as linhas para garantir consistência
        return tuple(sorted(linhas))


    # aplicar por Leadset
    chaves = (
        dados
        .groupby("Leadset")[colunas]
        .apply(gerar_chave_grupo)
    )

    # jogar de volta no dataframe
    dados["linha_key"] = dados["Leadset"].map(chaves)
    #--------------------------------------------------------------------------------------


    dados["Status"] = None

    # Agrupar por essa chave
    grupos = dados.groupby("linha_key").indices
    id = 1
    # Mostrar grupos com mais de uma linha
    for key, indices in grupos.items():
        if len(indices) > 1:
            dados.loc[indices, "Status"] = id
            id += 1

    dados.drop(columns=["linha_key"], inplace=True)

    dados["CIRC_MASTER"] = None
    dados["CIRC_MASTER"] = None
    dados["CIRC_COMUNS"] = None


    valores_invalidos = ["0 0", "0 1", "0.14", "0.18"]

    id_list = dados["Status"].dropna().unique().tolist()

    for i in id_list:
        index_list = dados[dados["Status"] == i].index
        index_number = index_list[0]

        grupo = dados.loc[index_list]
        if len(grupo['Leadset'].unique().tolist())>1:
            # --- PRIMEIRA REGRA (Leadset) ---
            master_value = dados.loc[index_number, "Leadset"]

            dados.loc[index_list, "CIRC_MASTER"] = master_value
            dados.loc[index_number, "CIRC_COMUNS"] = master_value

            for j in index_list[1:]:
                dados.loc[j, "CIRC_COMUNS"] = dados.loc[j, "Leadset"]
                

            # --- VALIDAÇÃO ---
            item_master = grupo["CIRC_MASTER"].unique()
            item_comum = grupo["CIRC_COMUNS"].unique()

            # --- SEGUNDA REGRA (Internal Family + CIRCUIT) ---
            if len(item_master) == 1 and len(item_comum) == 1:

                master_value = (
                    dados.loc[index_number, "Internal Family"]
                    + dados.loc[index_number, "CIRCUIT"]
                )

                dados.loc[index_list, "CIRC_MASTER"] = master_value
                dados.loc[index_number, "CIRC_COMUNS"] = master_value

                for j in index_list[1:]:
                    dados.loc[j, "CIRC_COMUNS"] = (
                        dados.loc[j, "Internal Family"] + dados.loc[j, "CIRCUIT"]
                    )

            # --- SECTIONN (CORRIGIDO) ---
            sectionn = grupo["SECTIONN"].unique()

            if any(val in valores_invalidos for val in sectionn):
                dados.loc[index_list, ["CIRC_MASTER", "CIRC_COMUNS", "Status"]] = None
            
    dados.drop(columns=["Status"], inplace=True)

    # Remover comunização da cravação dupla
    try:
        mask = dados["JOINT_TO_A"].str.match(r"^\d", na=False) & ~dados["RMKS_A"].str.contains("MULT", na=False)

        dados.loc[mask, "CIRC_MASTER"] = None
        dados.loc[mask, "CIRC_COMUNS"] = None
        dados.loc[mask, "Status"] = None
        
    except: pass
    
    try:
        mask = dados["JOINT_TO_B"].str.match(r"^\d", na=False) & ~dados["RMKS_B"].str.contains("MULT", na=False)
    
    #-------------------------------------------------------------------
    
        dados.loc[mask, "CIRC_MASTER"] = None
        dados.loc[mask, "CIRC_COMUNS"] = None
        dados.loc[mask, "Status"] = None
    except: pass

    dados = dados.replace({pd.NA: np.nan, None: np.nan})

    return dados

def comparar_tabela_sap(df0: pd.DataFrame = None, path_dados: str = None) -> pd.DataFrame:

    with open(path_dados, "rb") as f:
        df2 = pd.read_csv(f, sep=";")

    df1 = df0.copy()
    df1 = df1.dropna(subset=["CIRC_MASTER"]).reset_index(drop=True)
    lista_master = df1["CIRC_MASTER"].unique().tolist()
    for i in lista_master:
        df_prov1 = sorted(df1[df1["CIRC_MASTER"] == i]["CIRC_COMUNS"].unique().tolist())
        df_prov2 = sorted(
            df2[df2["CIRC_COMUNS"].isin(df_prov1)]["CIRC_COMUNS"].unique().tolist()
        )

        if df_prov1 != df_prov2:
            if len(df_prov1) > 0 and len(df_prov2) == 0:
                lista_index1 = df0[df0["CIRC_COMUNS"].isin(df_prov1)].index
                df0.loc[lista_index1, "STATUS"] = "Adicionar"
            elif len(df_prov1) == 0 and len(df_prov2) > 0:
                lista_index2 = df0[df0["CIRC_COMUNS"].isin(df_prov2)].index
                df0.loc[lista_index2, "STATUS"] = "Remover"
        else:
            if len(df_prov1) == len(df_prov2):
                lista_index = df0[df0["CIRC_COMUNS"].isin(df_prov1)].index
                df0.loc[lista_index, "STATUS"] = "Já esta comunizado."

    colunas = [
        "Leadset",
        "TYPE",
        "WERKS",
        "External Family",
        "FILE_LINE",
        "STATUS_REGISTRO",
        "Internal Family",
        "CIRC_MASTER",
        "CIRC_COMUNS",
        "STATUS",
        "CIRCUIT",
        "WIRE_TUBE_SPLICE",
        "LENGTH",
        "Comp TW",
    ]
    df0 = reordenar_colunas(df0, colunas)

    return df0

def arrumar_leadset(dados: pd.DataFrame = None) -> pd.DataFrame:

    def extrair_codigo(texto):
        return re.sub(r"^[A-Z]+\d+_?", "", texto)

    def extrair_codigo_SHIELDE(texto):
        return re.sub(r"^[A-Z]+\d+", "", texto)

    for i in dados.index:
        leadset = dados.loc[i, "Leadset"]

        RMKS_A = str(dados.loc[i, "RMKS_A"]).split(",")
        RMKS_A = [x for x in RMKS_A if x]

        if len(RMKS_A) == 1 and str(leadset).startswith("B") and "W" in leadset:
            dados.loc[i, "MULTICORE"] = leadset[3:]

        elif len(RMKS_A) > 1 and str(leadset).startswith("B") and "W" in leadset:
            dados.loc[i, "MULTICORE"] = leadset[3:]

        # Criar multicore
        if ("TW" in leadset or "MC" in leadset) and (
            str(leadset).startswith("G") or str(leadset).startswith("S")
        ):
            dados.loc[i, "MULTICORE"] = extrair_codigo(leadset)

        if ("MC" in leadset or "TW" in leadset) and str(leadset).startswith("V"):
            dados.loc[i, "MULTICORE"] = extrair_codigo(leadset)

        if "SHIELDE" in leadset and str(leadset).startswith("X"):
            dados.loc[i, "MULTICORE"] = extrair_codigo_SHIELDE(leadset)

        # Ajustar
        if str(leadset).startswith("V") and "TW" in leadset:
            dados.loc[i, "Leadset"] = (
                f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            )

        if (str(leadset).startswith("G") or str(leadset).startswith("S")) and float(
            dados.loc[i, "SECTIONN"].replace(" ", "")
        ) >= 1.5:
            dados.loc[i, "Leadset"] = (
                f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            )

        if str(leadset).startswith("X") and "TW" in leadset:
            dados.loc[i, "Leadset"] = (
                f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            )

        if len(RMKS_A) == 1 and str(leadset).startswith("B") and "W" in leadset:
            dados.loc[i, "Leadset"] = (
                f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            )

    return dados

def add_tW(dados: pd.DataFrame = None) -> pd.DataFrame:
    
    dados['Leadset_2'] = dados['Internal Family']+dados['CIRCUIT']

    dict_circuitos = dict(zip(dados['Leadset_2'], dados['Leadset']))
    
    dados['CIRC_COMUNS_PROCESSO'] = dados['CIRC_COMUNS'].map(dict_circuitos)
    dados['CIRC_MASTER_PROCESSO'] = dados['CIRC_MASTER'].map(dict_circuitos)
    
    dados.drop(columns=['Leadset_2'],inplace=True)
    
    return dados

def arrumar_processo(dados: pd.DataFrame = None, processo:str=None):

    ckt_TW = dados[dados['CIRC_MASTER_PROCESSO'].str.contains(processo,na=False)]['CIRC_MASTER_PROCESSO'].unique().tolist()

    for i in ckt_TW:
        df_temp = dados[dados['CIRC_MASTER_PROCESSO']==i]
        
        
        for j in df_temp['WIRE_TUBE_SPLICE'].unique():
            df_temp2 = df_temp[df_temp['WIRE_TUBE_SPLICE']==j]
            leadset = df_temp2.iloc[0]['CIRC_MASTER']
            circuito = df_temp2.iloc[0]['CIRCUIT']
            
            if leadset[3]=='_':
                leadset = leadset[:4]
            else:
                leadset = leadset[:3]
            leadset_novo = f'{leadset}{circuito}'
            mask = df_temp2.index
            dados.loc[mask,'CIRC_MASTER'] = leadset_novo
            
    return dados

# =========================================================
# THREAD
# =========================================================
class WorkerListaCorte(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(self, sap, cmz, pasta, familias, grupos):
        super().__init__()

        self.sap = sap
        self.cmz = cmz
        self.pasta = pasta
        self.familias = familias
        self.grupos = grupos

    def run(self):

        try:

            dados_base, _ = converter_sap(self.sap)

            # remove famílias selecionadas
            dados_base = dados_base[
                ~dados_base["Internal Family"].isin(self.familias)
            ].reset_index(drop=True)

            dados_base = arrumar_leadset(dados=dados_base)

            resultados = []

            for grupo in self.grupos:

                df = dados_base[
                    dados_base["Internal Family"].isin(grupo)
                ]

                if df.empty:
                    continue

                parcial = gerar_comunizacao_arquivo_sap(
                    dados=df,
                    ordem=grupo
                )

                if parcial is not None and not parcial.empty:
                    resultados.append(parcial)

            if not resultados:
                self.erro.emit("Nenhum grupo gerou resultado.")
                return

            final = pd.concat(resultados, ignore_index=True)

            # =====================================================
            # PROCESSOS
            # =====================================================
            final = add_tW(dados=final)

            final = comparar_tabela_sap(
                df0=final,
                path_dados=self.cmz
            )

            final = arrumar_processo(
                dados=final,
                processo='TW'
            )

            colunas = [
                'Leadset',
                'TYPE',
                'WERKS',
                'External Family',
                'FILE_LINE',
                'STATUS_REGISTRO',
                'Internal Family',
                'CIRC_MASTER',
                'CIRC_COMUNS',
                'STATUS',
                'CIRC_MASTER_PROCESSO',
                'CIRC_COMUNS_PROCESSO',
                'MULTICORE'
            ]

            final = reordenar_colunas(
                df=final,
                colunas_prioritarias=colunas
            )

            # =====================================================
            # CMZ
            # =====================================================
            dados_cmz = final[
                final['CIRC_MASTER'].notna()
            ][['CIRC_MASTER', 'CIRC_COMUNS']].reset_index(drop=True)

            dados_cmz['WERKS'] = 1600

            dados_cmz = dados_cmz[
                ['WERKS', 'CIRC_MASTER', 'CIRC_COMUNS']
            ]

            dados_cmz = dados_cmz.drop_duplicates().reset_index(drop=True)

            # =====================================================
            # EXPORTAÇÃO
            # =====================================================
            data_atual = datetime.now().strftime("%d-%m-%Y")
            data_atual_novo = datetime.now().strftime("%d.%m.%Y")

            final.to_excel(
                rf"{self.pasta}\LISTA_DE_CORTE_{data_atual}.xlsx",
                index=False
            )

            dados_cmz.to_csv(
                rf"{self.pasta}\1600.Comuniza_novo_{data_atual_novo}.csv",
                sep=';',
                index=False
            )

            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))


# =========================================================
# UI
# =========================================================
class ListaDeCorte(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TMGods 🧠 Industrial Engineering")
        self.resize(700, 620)
        self.centralizar_tela_atual()
        self.grupos = []

        self.setStyleSheet(self.estilo())

        layout = QVBoxLayout()

        # =====================================================
        # SAP
        # =====================================================
        layout.addWidget(QLabel("Arquivo SAP"))

        sap_layout = QHBoxLayout()

        self.input_sap = QLineEdit()
        self.btn_sap = QPushButton("...")

        sap_layout.addWidget(self.input_sap)
        sap_layout.addWidget(self.btn_sap)

        layout.addLayout(sap_layout)

        # =====================================================
        # CMZ
        # =====================================================
        layout.addWidget(QLabel("Arquivo Comunização"))

        cmz_layout = QHBoxLayout()

        self.input_cmz = QLineEdit()
        self.btn_cmz = QPushButton("...")

        cmz_layout.addWidget(self.input_cmz)
        cmz_layout.addWidget(self.btn_cmz)

        layout.addLayout(cmz_layout)

        # =====================================================
        # PASTA
        # =====================================================
        layout.addWidget(QLabel("Pasta saída"))

        pasta_layout = QHBoxLayout()

        self.input_pasta = QLineEdit()
        self.btn_pasta = QPushButton("...")

        pasta_layout.addWidget(self.input_pasta)
        pasta_layout.addWidget(self.btn_pasta)

        layout.addLayout(pasta_layout)

        layout.addWidget(self.linha())

        # =====================================================
        # REMOÇÃO
        # =====================================================
        layout.addWidget(QLabel("Código UCS"))

        dual_layout = QHBoxLayout()

        self.list_disponiveis = QListWidget()
        self.list_selecionadas = QListWidget()

        right_layout = QVBoxLayout()

        self.label_remocao = QLabel(
            "Nenhum item selecionado para remoção"
        )

        right_layout.addWidget(self.label_remocao)
        right_layout.addWidget(self.list_selecionadas)

        btns = QVBoxLayout()

        self.btn_add = QPushButton("→")
        self.btn_rm = QPushButton("←")

        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_rm)

        dual_layout.addWidget(self.list_disponiveis)
        dual_layout.addLayout(btns)
        dual_layout.addLayout(right_layout)

        layout.addLayout(dual_layout)

        layout.addWidget(self.linha())

        # =====================================================
        # GRUPOS
        # =====================================================
        layout.addWidget(
            QLabel("Montar Grupos")
        )

        grupo_layout = QHBoxLayout()

        self.list_grupo_disponiveis = QListWidget()
        self.list_grupo_disponiveis.setSelectionMode(
            QAbstractItemView.MultiSelection
        )

        self.list_grupo_montado = QListWidget()

        grupo_btns = QVBoxLayout()

        self.btn_add_g = QPushButton("→")
        self.btn_rm_g = QPushButton("←")
        self.btn_fechar_g = QPushButton("Fechar grupo")

        grupo_btns.addWidget(self.btn_add_g)
        grupo_btns.addWidget(self.btn_rm_g)
        grupo_btns.addWidget(self.btn_fechar_g)

        grupo_layout.addWidget(self.list_grupo_disponiveis)
        grupo_layout.addLayout(grupo_btns)
        grupo_layout.addWidget(self.list_grupo_montado)

        layout.addLayout(grupo_layout)

        self.lista_grupos = QListWidget()

        layout.addWidget(QLabel("Grupos criados"))
        layout.addWidget(self.lista_grupos)
        layout.addWidget(self.linha())
        # =====================================================
        # EXECUTAR
        # =====================================================
        self.btn_exec = QPushButton("Executar")
        self.btn_exec.setObjectName("botaoExec")
        self.btn_exec.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        
        layout.addWidget(self.btn_exec)

        self.setLayout(layout)

        # =====================================================
        # CONEXÕES
        # =====================================================
        self.btn_add.clicked.connect(self.add_item)
        self.btn_rm.clicked.connect(self.remove_item)

        self.btn_add_g.clicked.connect(self.add_grupo_item)
        self.btn_rm_g.clicked.connect(self.remove_grupo_item)
        self.btn_fechar_g.clicked.connect(self.fechar_grupo)

        self.btn_exec.clicked.connect(self.executar)
        self.btn_exec.setObjectName("botaoExec")

        self.btn_sap.clicked.connect(self.load_sap)
        self.btn_cmz.clicked.connect(self.load_cmz)
        self.btn_pasta.clicked.connect(self.load_pasta)

    # =====================================================
    # LABEL
    # =====================================================
    
    
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
    
    def atualizar_label(self):

        total = self.list_selecionadas.count()

        if total == 0:
            self.label_remocao.setText(
                "Nenhum item selecionado para remoção"
            )
        else:
            self.label_remocao.setText(
                f"{total} itens serão removidos"
            )

    # =====================================================
    # LIMPEZA TOTAL DOS GRUPOS
    # =====================================================
    def sync_grupos(self):

        validos = {
            self.list_disponiveis.item(i).text()
            for i in range(self.list_disponiveis.count())
        }

        # =====================================================
        # LIMPA MONTAGEM
        # =====================================================
        self.list_grupo_montado.blockSignals(True)
        self.list_grupo_montado.clear()
        self.list_grupo_montado.blockSignals(False)

        # =====================================================
        # LIMPA FECHADOS
        # =====================================================
        self.grupos.clear()

        self.lista_grupos.blockSignals(True)
        self.lista_grupos.clear()
        self.lista_grupos.blockSignals(False)

        # =====================================================
        # RECONSTRUIR DISPONÍVEIS
        # =====================================================
        self.list_grupo_disponiveis.blockSignals(True)
        self.list_grupo_disponiveis.clear()
        self.list_grupo_disponiveis.addItems(
            sorted(validos)
        )
        self.list_grupo_disponiveis.blockSignals(False)

    # =====================================================
    # REMOÇÃO
    # =====================================================
    def add_item(self):

        item = self.list_disponiveis.currentItem()

        if item:

            self.list_disponiveis.takeItem(
                self.list_disponiveis.row(item)
            )

            self.list_selecionadas.addItem(
                item.text()
            )

            self.atualizar_label()

            # IMPORTANTE:
            # qualquer mudança limpa grupos
            self.sync_grupos()

    def remove_item(self):

        item = self.list_selecionadas.currentItem()

        if item:

            self.list_selecionadas.takeItem(
                self.list_selecionadas.row(item)
            )

            self.list_disponiveis.addItem(
                item.text()
            )

            self.atualizar_label()

            # IMPORTANTE:
            # qualquer mudança limpa grupos
            self.sync_grupos()

    # =====================================================
    # GRUPOS
    # =====================================================
    def add_grupo_item(self):

        items = self.list_grupo_disponiveis.selectedItems()

        if not items:
            return

        for item in items:

            self.list_grupo_montado.addItem(
                item.text()
            )

        for item in items:

            self.list_grupo_disponiveis.takeItem(
                self.list_grupo_disponiveis.row(item)
            )

    def remove_grupo_item(self):

        item = self.list_grupo_montado.currentItem()

        if item:

            self.list_grupo_montado.takeItem(
                self.list_grupo_montado.row(item)
            )

            self.list_grupo_disponiveis.addItem(
                item.text()
            )

    def fechar_grupo(self):

        if self.list_grupo_montado.count() == 0:

            QMessageBox.warning(
                self,
                "Erro",
                "Grupo vazio"
            )

            return

        grupo = [
            self.list_grupo_montado.item(i).text()
            for i in range(
                self.list_grupo_montado.count()
            )
        ]

        self.grupos.append(grupo)

        self.lista_grupos.addItem(
            " + ".join(grupo)
        )

        self.list_grupo_montado.clear()

    # =====================================================
    # EXECUTAR
    # =====================================================
    def executar(self):

        sap = self.input_sap.text().strip()
        cmz = self.input_cmz.text().strip()
        pasta = self.input_pasta.text().strip()

        familias = [
            self.list_selecionadas.item(i).text()
            for i in range(
                self.list_selecionadas.count()
            )
        ]

        # =====================================================
        # VALIDAÇÕES
        # =====================================================
        if not sap:
            QMessageBox.warning(
                self,
                "Erro",
                "Selecione o arquivo SAP"
            )
            return

        if not cmz:
            QMessageBox.warning(
                self,
                "Erro",
                "Selecione o arquivo Comunização"
            )
            return

        if not pasta:
            QMessageBox.warning(
                self,
                "Erro",
                "Selecione a pasta de saída"
            )
            return

        # if not familias:
        #     QMessageBox.warning(
        #         self,
        #         "Erro",
        #         "Selecione itens para remoção"
        #     )
        #     return

        if self.list_grupo_montado.count() > 0:
            QMessageBox.warning(
                self,
                "Erro",
                "Existe grupo em montagem aberto"
            )
            return

        if self.list_grupo_disponiveis.count() > 0:
            QMessageBox.warning(
                self,
                "Erro",
                "Ainda existem itens sem grupo"
            )
            return

        if not self.grupos:
            QMessageBox.warning(
                self,
                "Erro",
                "Nenhum grupo foi criado"
            )
            return

        # =====================================================
        # PROCESSO
        # =====================================================
        self.progresso = QProgressDialog(
            "Processando...",
            None,
            0,
            0,
            self
        )

        self.progresso.setCancelButton(None)
        self.progresso.show()

        self.thread = WorkerListaCorte(
            sap=sap,
            cmz=cmz,
            pasta=pasta,
            familias=familias,
            grupos=self.grupos
        )

        self.thread.finalizado.connect(self.ok)
        self.thread.erro.connect(self.err)

        self.thread.start()

    # =====================================================
    # THREAD CALLBACKS
    # =====================================================
    def ok(self):

        self.progresso.close()

        QMessageBox.information(
            self,
            "OK",
            "Finalizado com sucesso"
        )

    def err(self, e):

        self.progresso.close()

        QMessageBox.critical(
            self,
            "Erro",
            str(e)
        )

    # =====================================================
    # LOADS
    # =====================================================
    def load_sap(self):

        file, _ = QFileDialog.getOpenFileName(
            self,
            "SAP",
            "",
            "Excel (*.xlsx *.csv)"
        )

        if file:

            self.input_sap.setText(file)

            dados, _ = converter_sap(file)

            itens = sorted(
                dados["Internal Family"]
                .dropna()
                .unique()
            )

            self.list_disponiveis.clear()
            self.list_disponiveis.addItems(itens)

            self.list_selecionadas.clear()

            self.sync_grupos()

            self.atualizar_label()

    def load_cmz(self):

        file, _ = QFileDialog.getOpenFileName(
            self,
            "CMZ",
            "",
            "Excel (*.xlsx *.csv)"
        )

        if file:
            self.input_cmz.setText(file)

    def load_pasta(self):

        pasta = QFileDialog.getExistingDirectory(
            self,
            "Pasta"
        )

        if pasta:
            self.input_pasta.setText(pasta)

    # =====================================================
    # UI HELPERS
    # =====================================================
    def linha(self):

        f = QFrame()
        f.setFrameShape(QFrame.HLine)

        return f


    def estilo(self):

        return """
        QWidget {
            background:#1e1e1e;
            color:white;
            font-size:13px;
        }

        QPushButton {
            background:#2d2d2d;
            border:1px solid #444;
        }

        QPushButton:hover {
            background:#3a3a3a;
        }

        QListWidget {
            background:#2d2d2d;
        }

        QPushButton#botaoExec {
            background-color: #0078d7;
            color: white;
            font-weight: bold;
            border: none;
            border-radius: 6px;
            padding: 8px 14px;
        }

        QPushButton#botaoExec:hover {
            background-color: #2893ff;
        }

        QPushButton#botaoExec:pressed {
            background-color: #005ea6;
        }
        """

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    app = QApplication(sys.argv)

    janela = ListaDeCorte()
    janela.show()

    sys.exit(app.exec())