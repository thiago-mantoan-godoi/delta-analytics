import sys
import pandas as pd
import numpy as np
import os
from typing import Iterable, Iterator, Union, List, Optional, Dict
import tempfile
from collections import defaultdict
import re
from openpyxl.utils import get_column_letter
from collections import Counter

import warnings
warnings.filterwarnings('ignore')




from PySide6.QtWidgets import (
    QWidget, QApplication, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog,
    QProgressDialog, QMessageBox, QStyle
)
from PySide6.QtCore import QThread, Signal, Qt


from pathlib import Path
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
	
from utils.function import *
from utils.funcoes import *

# =========================================================
# FUNCTION
# =========================================================
@log_errors_record
def procurar_indices_linhas(palavras: List[str], df: pd.DataFrame) -> List[int]:
    """
    Retorna os índices das linhas cujo valor na primeira coluna
    está presente na lista `palavras`.

    Args:
        palavras (List[str]): Palavras a procurar na primeira coluna.
        df (pd.DataFrame): DataFrame a ser analisado.

    Returns:
        List[int]: Lista de índices encontrados, com o último índice
                   igual ao número de linhas do DataFrame.
    """
    primeira_coluna = df.columns[0]
    indices = df[df[primeira_coluna].isin(palavras)].index.tolist()
    indices.append(df.shape[0])  # adiciona índice final para conveniência
    #logger.debug("Índices encontrados para palavras %s: %s", palavras, indices)
    return indices

@log_errors_record
def renomear_colunas_duplicadas(df: pd.DataFrame, sep: str = "_", copiar: bool = True) -> pd.DataFrame:
    """
    Renomeia colunas duplicadas adicionando um sufixo incremental.

    Exemplo:
        ['A', 'B', 'A', 'A'] → ['A', 'B', 'A_1', 'A_2']

    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame de entrada.
    sep : str, opcional
        Separador entre o nome original e o sufixo numérico (default "_").
    copiar : bool, opcional
        Se True, retorna uma cópia do DataFrame (default True).

    Retorno
    -------
    pd.DataFrame
        DataFrame com colunas renomeadas.
    """
    if copiar:
        df = df.copy()

    contadores = {}
    novas_colunas = []

    for coluna in df.columns:
        if coluna not in contadores:
            contadores[coluna] = 0
            novas_colunas.append(coluna)
        else:
            contadores[coluna] += 1
            novas_colunas.append(f"{coluna}{sep}{contadores[coluna]}")

    df.columns = novas_colunas
    return df

@log_errors_record
def reordenar_colunas(df: pd.DataFrame, colunas_prioritarias: List[str]) -> pd.DataFrame:
    """
    Reordena as colunas do DataFrame, colocando as colunas prioritárias
    no início, na ordem fornecida, e mantendo as demais após.

    Args:
        df (pd.DataFrame): DataFrame a ser reorganizado.
        colunas_prioritarias (List[str]): Lista de colunas que devem aparecer primeiro.

    Returns:
        pd.DataFrame: DataFrame com colunas reordenadas.
    """
    colunas_prioritarias_existentes = [col for col in colunas_prioritarias if col in df.columns]
    colunas_restantes = [col for col in df.columns if col not in colunas_prioritarias_existentes]

    #logger.debug("Colunas reordenadas. Prioritárias: %s, Restantes: %s",
    #             colunas_prioritarias_existentes, colunas_restantes)
    return df[colunas_prioritarias_existentes + colunas_restantes]

@log_errors_record
def excluir_colunas(df: pd.DataFrame, colunas_para_excluir: List[str]) -> pd.DataFrame:
    """
    Exclui colunas do DataFrame, considerando apenas as colunas existentes.

    Args:
        df (pd.DataFrame): DataFrame original.
        colunas_para_excluir (List[str]): Lista de colunas a remover.

    Returns:
        pd.DataFrame: DataFrame com colunas removidas.
    """
    colunas_existentes = [col for col in colunas_para_excluir if col in df.columns]
    #logger.debug("Colunas a excluir: %s", colunas_existentes)
    return df.drop(columns=colunas_existentes)

@log_errors_record
def format_cut(path_name=str):
    with open(path_name, "rb") as f:
        df = pd.read_excel(f,dtype=str)

    palavras = ["Wire Nb", "Mult.Wire Nb", "Splice Nb", "Sleeve Nb"]
    indices_linhas = procurar_indices_linhas(palavras, df)

    dict_df = {}

    for i in range(len(indices_linhas) - 1):
        index_1 = indices_linhas[i]
        index_2 = indices_linhas[i + 1]

        df_filtrado = df.loc[index_1 : index_2 - 1].reset_index(drop=True)

        novo_header = df_filtrado.iloc[0]
        df_novo = df_filtrado[1:].copy()
        df_novo.columns = novo_header
        df_novo.reset_index(drop=True, inplace=True)

        name = df_novo.columns[0].replace(" ", "_").replace(".", "")

        df_novo = (df_novo.dropna(how="all", axis=0).dropna(how="all", axis=1).reset_index(drop=True))

        df_novo = renomear_colunas_duplicadas(df_novo)

        df_novo["Nome do arquivo"] = os.path.basename(path_name).replace(".xlsx", "")
        
        if name == "Wire_Nb":
            try:
                list_add = ['Sub-Assembly', "Multicore", "Term. 1", "Term. 2", "Seal 1", "Seal 2", "Joint 1", "Joint 2"]
                for l in list_add:
                    if l not in df_novo.columns:
                        df_novo[l] = None
                        
                colunas_prioritarias = ["Nome do arquivo", 'Wire Nb', "Multicore",'Sub-Assembly']
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
                
            except Exception as err:
                print(f"Erro reordenar colunas no {name}: {err}")
            try:
                colunas_para_excluir = ["Wire Spec", "Pack", "Options", "Cav. 1", "Cav. 2", "Term.Mat 1", "Term.Mat 2"]
                df_novo = excluir_colunas(df_novo, colunas_para_excluir)
            except Exception as err:
                print(f"Erro excluir colunas no {name}: {err}")

        if name == "MultWire_Nb":
            try:
                colunas_prioritarias = ["Nome do arquivo","Mult.Wire Nb","T","CSA","Length","Wire Spec","Pack",
                                        "CutBack1","CutBack2","W1", "W2", "W3", "W4", "Int.PN"]
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
            except Exception as err:
                print(f"Erro reordenar colunas no {name}: {err}")

        if name == "Sleeve_Nb":
            try:
                df_novo["Length"] = df_novo["Length"].astype(float).astype(int)
                df_novo["ID"] = (
                        df_novo["Int.PN"].astype(str)
                        + "-"
                        + df_novo["Length"].astype(int).astype(str)
                    )
                colunas_prioritarias = ["Nome do arquivo",'ID']
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
            except Exception as err:
                print(f"Erro reordenar colunas no {name}: {err}")

        if name == "Splice_Nb":
            try:
                colunas = ["L1","L2","L3","L4","L5","L6","L7","L8","L9","L10","R1","R2","R3","R4","R5","R6","R7","R8","R9","R10"]

                for c in colunas:
                    if c not in df_novo.columns:
                        df_novo[c] = None

                colunas_prioritarias = ["Nome do arquivo","Splice Nb","Int.PN",
                                        "Extra Component PN","Pack","CSA Left","CSA Right","CSA Total","Node", 
                                        "L1","L2","L3","L4","L5","L6","L7","L8","L9","L10",
                                        "R1","R2","R3","R4","R5","R6","R7","R8","R9","R10"]
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
                
            except Exception as err:
                print(f"Erro reordenar colunas no {name}: {err}")

        dict_df[name] = df_novo
        
    df_sheets = pd.ExcelFile(path_name).sheet_names
    
    for sheets in df_sheets:
        if "BARE" in str(sheets).upper():
            dict_df['Bare'] = pd.read_excel(path_name, sheet_name=sheets, dtype=str).dropna(how="all", axis=0)
        if "MAP" in str(sheets).upper() or "MULT" in str(sheets).upper():
            dict_df['Multicrimp'] = pd.read_excel(path_name, sheet_name=sheets, dtype=str).dropna(how="all", axis=0)
        
    return dict_df

@log_errors_record
def comparador_de_bases(path_atual: str, path_novo: str, colunas_tag: list[str], coluna_ref = str, nome_base=str):
    dict_atual = format_cut(path_atual)
    dict_novo = format_cut(path_novo)
    
    
    wire_atual = dict_atual[nome_base]
    wire_novo = dict_novo[nome_base]
    
    if len(wire_atual)>0 and len(wire_novo)>0:

        wire_atual_index = wire_atual.set_index(coluna_ref)
        wire_novo_index = wire_novo.set_index(coluna_ref)

        todas_keys = set(wire_atual_index.index) | set(wire_novo_index.index)

        todas_keys = {
            k for k in todas_keys
            if pd.notna(k)
        }

        linhas = []

        for key in sorted(todas_keys, key=str):

            existe_atual = key in wire_atual_index.index
            existe_novo = key in wire_novo_index.index

            # =====================================================
            # CIRCUITO NOVO
            # =====================================================
            if not existe_atual and existe_novo:

                linhas.append({
                    "Atual": "-",
                    "Novo": key,
                    "Status": "Novo",
                    "Observação": ""
                })
                continue

            # =====================================================
            # CIRCUITO EXCLUIDO
            # =====================================================
            if existe_atual and not existe_novo:

                linhas.append({
                    "Atual": key,
                    "Novo": "-",
                    "Status": "Excluido",
                    "Observação": ""
                })
                continue

            # =====================================================
            # COMPARAÇÃO
            # =====================================================
            row_a = wire_atual_index.loc[key, colunas_tag]
            row_b = wire_novo_index.loc[key, colunas_tag]

            valores_a = row_a.astype(str).fillna("nan").tolist()
            valores_b = row_b.astype(str).fillna("nan").tolist()

            if valores_a == valores_b:

                linhas.append({
                    "Atual": key,
                    "Novo": key,
                    "Status": "Sem alteração",
                    "Observação": ""
                })

            else:

                diffs = []

                for col, a, b in zip(colunas_tag, valores_a, valores_b):

                    if str(a) != str(b):

                        diffs.append(f"{col}: {a} → {b}")

                observacao = "\n".join(diffs)

                linhas.append({
                    "Atual": key,
                    "Novo": key,
                    "Status": "Alteração",
                    "Observação": observacao
                })

        return pd.DataFrame(linhas)
    elif len(wire_atual)>0:
        resultado = pd.DataFrame({
            "Atual": wire_atual[coluna_ref].reset_index(drop=True),
            "Novo": "-",
            "Status": "Excluido",
            "Observação": ""
        })
        
        return resultado
    elif len(wire_novo)>0:
        try:
            resultado = pd.DataFrame({
                "Atual": "Novo",
                "Novo": wire_novo[coluna_ref].reset_index(drop=True),
                "Status": "-",
                "Observação": ""
            })
        except:
            display(wire_atual)
        return resultado

@log_errors_record
def revisar_splice(path_splice: str = "", dict_alt: Optional[dict] = None, df_diff: pd.DataFrame=None):
    if dict_alt is None:
        dict_alt = {}
    df_sp = format_cut(path_name=path_splice)['Splice_Nb']
    colunas_tag = [
        'L1','L2','L3','L4','L5','L6','L7','L8','L9','L10',
        'R1','R2','R3','R4','R5','R6','R7','R8','R9','R10'
    ]
    dict_new = defaultdict(list)
    
    for i in df_sp.index:

        sp = df_sp.loc[i, 'Splice Nb']

        valores = (
            df_sp.loc[i, colunas_tag]
            .dropna()
            .astype(str)
            .tolist()
        )

        for k, v in dict_alt.items():

            if k in valores:
                dict_new[sp].append(f"Circuito: {k} - {v}")

    resultado = {
        sp: ", \n".join(obs)
        for sp, obs in dict_new.items()
    }


    df_diff['Observação'] = df_diff['Novo'].apply(lambda x: resultado.get(x, ""))

    df_diff.loc[df_diff['Observação'].fillna("").str.strip().ne(""), 'Status'] = "Alteração"

    return df_diff

@log_errors_record
def gerar_resumos_das_comp(path_atual: str, path_novo: str):
    dict_final = {}
    # WIRE
    Wire_colunas = ['Multicore', 'Sub-Assembly', 'T', 'CSA', 'Length', 'C1', 'C2', 'Int.PN', 
            'Term. 1', 'Strip 1', 'Seal 1', 'Node 1', 'Joint 1', 'T_1', 'Note 1', 
            'Term. 2', 'Strip 2', 'Seal 2','Node 2', 'Joint 2', 'T_2', 'Note 2']

    Wire_diff = comparador_de_bases(path_atual=path_atual, 
                                path_novo=path_novo,
                                colunas_tag=Wire_colunas,
                                coluna_ref="Wire Nb",
                                nome_base="Wire_Nb")

    dict_alt_wire = {}
    dict_alt_wire = {row['Novo']:row['Observação'] for i, row in Wire_diff[Wire_diff['Status']=='Alteração'].iterrows()}

    dict_final["Wire_Nb"] = Wire_diff
    
    # MultWire_Nb
    MultWire_colunas = ['T', 'CSA', 'Length', 'Wire Spec', 'Pack', 'CutBack1', 'CutBack2', 'W1', 'W2']

    MultWire_diff = comparador_de_bases(path_atual=path_atual, 
                                path_novo=path_novo,
                                colunas_tag=MultWire_colunas,
                                coluna_ref='Mult.Wire Nb',
                                nome_base="MultWire_Nb")

    dict_final["MultWire_Nb"] = MultWire_diff
    # Sleeve Nb
    Sleeve_colunas = ['Int.PN', 'Length', 'Type', 'Pack','Description', 'Mater.', 'Type_1', 'Width', 
            'Group', 'Color', 'Slit','Conv', 'Wall Th.', 'Layer', 'Node 1', 'Node 2', 'Item']

    Sleeve_diff = comparador_de_bases(path_atual=path_atual, 
                                path_novo=path_novo,
                                colunas_tag=Sleeve_colunas,
                                coluna_ref='Sleeve Nb',
                                nome_base="Sleeve_Nb")
    
    dict_final["Sleeve_Nb"] = Sleeve_diff

    # Splice Nb
    Splice_colunas = ['Int.PN', 'Extra Component PN', 'Pack','CSA Left', 'CSA Right', 'CSA Total', 'Node', 
            'L1', 'L2', 'L3', 'L4','L5', 'L6', 'L7', 'L8', 'L9', 'L10', 
            'R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'R7', 'R8', 'R9', 'R10',]

    Splice_diff = comparador_de_bases(path_atual=path_atual, 
                                path_novo=path_novo,
                                colunas_tag=Splice_colunas,
                                coluna_ref='Splice Nb',
                                nome_base="Splice_Nb")
    Splice_diff = revisar_splice(path_splice=path_novo, dict_alt=dict_alt_wire, df_diff=Splice_diff)
    
    dict_final["Splice_Nb"] = Splice_diff
    
    return dict_final

@log_errors_record
def salvar_dict_df_em_excel(dfs_dict, caminho_arquivo):

    # filtra somente dfs com conteúdo
    dfs_dict = {
        nome: df
        for nome, df in dfs_dict.items()
        if df is not None
        and not df.empty
        and not df.dropna(how="all").empty
    }

    # fallback
    if not dfs_dict:
        dfs_dict = {
            "SemDados": pd.DataFrame(
                {"Mensagem": ["Nenhum dado disponível"]}
            )
        }

    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:

        for nome_aba, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)

        # Ajusta largura das colunas
        sheets = writer.sheets

        for nome_aba, df in dfs_dict.items():

            worksheet = sheets[nome_aba]

            for idx, col in enumerate(df.columns, 1):

                max_len = (
                    max(
                        df[col]
                        .astype(str)
                        .apply(lambda x: len(str(x)))
                        .max(),
                        len(str(col)),
                    )
                    + 2
                )

                col_letter = get_column_letter(idx)

                worksheet.column_dimensions[col_letter].width = max_len
                
            


# =========================================================
# WORKER
# =========================================================

class WorkerResumoComp(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(self, cut_atual, cut_nova, caminho_saida):
        super().__init__()

        self.cut_atual = cut_atual
        self.cut_nova = cut_nova
        self.caminho_saida = caminho_saida

    def run(self):

        try:

            dict_final = gerar_resumos_das_comp(
                path_atual=self.cut_atual,
                path_novo=self.cut_nova
            )

            caminho_arquivo = os.path.join(
                self.caminho_saida,
                "Delta_Cut.xlsx"
            )

            salvar_dict_df_em_excel(
                dict_final,
                caminho_arquivo=caminho_arquivo
            )

            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))



class TelaComparacaoCUT(QWidget):

    def __init__(self):
        super().__init__()

        
        self.setWindowTitle("TMGods 🧠 Industrial Engineering")
        self.resize(700, 250)
        self.setStyleSheet(self.estilo())
        layout = QVBoxLayout()

        # =====================================================
        # TÍTULO
        # =====================================================

        titulo = QLabel("Gerar Comparação CUT")
        titulo.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(titulo)

        layout.addWidget(self.linha())

        # =====================================================
        # ARQUIVO ATUAL
        # =====================================================

        layout.addWidget(QLabel("Arquivo Atual"))

        h1 = QHBoxLayout()
        self.input_atual = QLineEdit()
        self.input_atual.setPlaceholderText("Selecione o arquivo atual...")

        btn_atual = QPushButton()
        btn_atual.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        btn_atual.clicked.connect(self.selecionar_atual)

        h1.addWidget(self.input_atual)
        h1.addWidget(btn_atual)

        layout.addLayout(h1)

        layout.addWidget(self.linha())

        # =====================================================
        # ARQUIVO NOVO
        # =====================================================

        layout.addWidget(QLabel("Arquivo Novo"))

        h2 = QHBoxLayout()
        self.input_novo = QLineEdit()
        self.input_novo.setPlaceholderText("Selecione o arquivo novo...")

        btn_novo = QPushButton()
        btn_novo.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))

        btn_novo.clicked.connect(self.selecionar_novo)

        h2.addWidget(self.input_novo)
        h2.addWidget(btn_novo)

        layout.addLayout(h2)

        layout.addWidget(self.linha())

        # =====================================================
        # PASTA SAÍDA
        # =====================================================

        layout.addWidget(QLabel("Salvar em"))

        h3 = QHBoxLayout()
        self.input_saida = QLineEdit()
        self.input_saida.setPlaceholderText("Selecione onde salvar...")

        btn_saida = QPushButton()
        btn_saida.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))

        btn_saida.clicked.connect(self.selecionar_saida)

        h3.addWidget(self.input_saida)
        h3.addWidget(btn_saida)

        layout.addLayout(h3)

        layout.addWidget(self.linha())

        # =====================================================
        # BOTÃO EXECUTAR
        # =====================================================

        layout.addStretch()

        self.btn_exec = QPushButton("Executar")
        self.btn_exec.setObjectName("botaoExec")
        self.btn_exec.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.btn_exec.setFixedWidth(120)
        self.btn_exec.clicked.connect(self.executar)

        layout.addWidget(self.btn_exec, alignment=Qt.AlignRight)

        self.setLayout(layout)

    # =========================================================
    # EXECUÇÃO
    # =========================================================

    def executar(self):

        cut_atual = self.input_atual.text().strip()
        cut_novo = self.input_novo.text().strip()
        saida = self.input_saida.text().strip()

        if not cut_atual or not cut_novo or not saida:
            QMessageBox.warning(
                self,
                "Atenção",
                "Preencha todos os campos antes de executar."
            )
            return

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

        self.thread = WorkerResumoComp(
            cut_atual=cut_atual,
            cut_nova=cut_novo,
            caminho_saida=saida
        )

        self.thread.finalizado.connect(self.finalizado)
        self.thread.erro.connect(self.erro)

        self.thread.start()

    def finalizado(self):

        self.progresso.close()

        QMessageBox.information(
            self,
            "Sucesso",
            "Arquivo gerado com sucesso!"
        )

    def erro(self, msg):

        self.progresso.close()

        QMessageBox.critical(
            self,
            "Erro",
            f"Ocorreu um erro:\n\n{msg}"
        )

    # =========================================================
    # SELETORES
    # =========================================================

    def selecionar_atual(self):

        file, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo atual",
            "",
            "Excel (*.xlsx)"
        )

        if file:
            self.input_atual.setText(file)

    def selecionar_novo(self):

        file, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo novo",
            "",
            "Excel (*.xlsx)"
        )

        if file:
            self.input_novo.setText(file)

    def selecionar_saida(self):

        pasta = QFileDialog.getExistingDirectory(
            self,
            "Selecionar pasta de saída"
        )

        if pasta:
            self.input_saida.setText(pasta)

    # =========================================================
    # LINHA
    # =========================================================

    def linha(self):

        frame = QWidget()
        line = QVBoxLayout(frame)

        sep = QLabel("")
        sep.setStyleSheet("background-color:#444; max-height:1px;")

        return sep


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

    janela = TelaComparacaoCUT()
    
    janela.show()

    sys.exit(app.exec())