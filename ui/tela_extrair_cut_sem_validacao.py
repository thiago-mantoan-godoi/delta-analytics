
from PySide6.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QLineEdit,
    QScrollArea, QMessageBox, QListWidget,
    QProgressBar, QApplication, QStyle
)
from PySide6.QtGui import QFont, QGuiApplication, QCursor
from PySide6.QtCore import Qt, Signal, QObject, QThread


import os
import tempfile
import shutil
import sys

from pathlib import Path
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).resolve().parent.parent


"""
Setup para notebooks:
- Ajusta sys.path para reconhecer o projeto
- Configura logging centralizado
- Importa decoradores de log
"""
#-----------------------------------------------------------------
import sys
from datetime import datetime
from pathlib import Path
import logging

from pathlib import Path
from typing import Iterable, Iterator, Union, List, Optional, Dict

import pandas as pd
from tabulate import tabulate
import re
from openpyxl.utils import get_column_letter
from collections import defaultdict

import warnings
warnings.simplefilter("ignore", UserWarning)


# Ajusta root do projeto para que imports funcionem
ROOT = Path.cwd().resolve()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


from pathlib import Path
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
	
from utils.function import *
from utils.funcoes import *

CONFIG_PATH = r"C:\Projetos\app\engineering\infra\files\config.xlsx"
#-------------------------------------------------------------------


class CopyWorker(QObject):
    started = Signal(int, str)         # total (unused), destino
    progressed = Signal(int)           # contador
    item_copied = Signal(int, str)     # índice, nome exibido
    error = Signal(str)
    finished = Signal(int, str)        # total copiado, destino

    def __init__(self, rows_data, palavras_chave, palavras_proibidas):
        super().__init__()
        self.rows_data = rows_data
        self.palavras_chave = [p.upper() for p in palavras_chave]
        self.palavras_proibidas = [p.upper() for p in palavras_proibidas]
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            # nome_pasta = "PastaTemp"
            # destino_root = os.path.join(tempfile.gettempdir(), nome_pasta)
            # os.makedirs(destino_root, exist_ok=True)
            
            nome_pasta = "PastaTemp"
            destino_root = os.path.join(tempfile.gettempdir(), nome_pasta)

            # Se existir, apaga tudo
            if os.path.exists(destino_root):
                shutil.rmtree(destino_root)

            # Cria novamente vazia
            os.makedirs(destino_root)

            count = 0
            self.started.emit(0, destino_root)

            for row in self.rows_data:
                prefixo = f"{row['project'].upper()}_{row['phase'].upper()}"
                origem = row["path"]

                for raiz, _, arquivos in os.walk(origem):
                    for arquivo in arquivos:
                        if self._cancel:
                            self.finished.emit(count, destino_root)
                            return

                        nome_up = arquivo.upper()
                        #if (any(k in nome_up for k in self.palavras_chave) and not any(b in nome_up for b in self.palavras_proibidas)):
                        if any(k in nome_up for k in self.palavras_chave) and all(b not in nome_up for b in self.palavras_proibidas):

                            src = os.path.join(raiz, arquivo)
                            nome, ext = os.path.splitext(arquivo)
                            novo_nome = f"{prefixo}_{nome}{ext}"
                            dst = os.path.join(destino_root, novo_nome)

                            try:
                                os.makedirs(os.path.dirname(dst), exist_ok=True)
                                shutil.copyfile(src, dst)  # mais rápido que copy2
                                count += 1
                                self.item_copied.emit(count, novo_nome)
                                self.progressed.emit(count)
                            except Exception as e:
                                self.error.emit(str(e))

            self.finished.emit(count, destino_root)

        except Exception as e:
            self.error.emit(str(e))
            self.finished.emit(0, "")

class MasterDataWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TMGods 🧠 Industrial Engineering")
        self.resize(700, 520)
        self.centralizar_tela_atual()
        self.setStyleSheet(self.estilo())
        self.items = []

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        
        self.destino_cut_consolidada = None

        title = QLabel("CUT´s")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        main_layout.addWidget(title)

        top_bar = QHBoxLayout()
        self.btn_add = QPushButton("Adicionar")
        self.btn_add.setFixedWidth(120)
        self.btn_add.clicked.connect(self.add_item)
        top_bar.addWidget(self.btn_add)
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)

        self.container = QWidget()
        self.items_layout = QVBoxLayout(self.container)
        self.items_layout.addStretch()
        self.scroll.setWidget(self.container)
        main_layout.addWidget(self.scroll)
        
        #----------------------------------------------------------------------
        dest_layout = QHBoxLayout()

        self.dest_path = QLineEdit()
        self.dest_path.setReadOnly(True)
        self.dest_path.setPlaceholderText("Selecione a pasta de destino")

        btn_dest = QPushButton("📁")
        btn_dest.clicked.connect(self.select_dest_folder)

        dest_layout.addWidget(QLabel("Destino:"))
        dest_layout.addWidget(self.dest_path)
        dest_layout.addWidget(btn_dest)

        main_layout.addLayout(dest_layout)
        #----------------------------------------------------------------------

        label = QLabel("Progresso")
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        main_layout.addWidget(label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)               # <--- texto visível
        self.progress.setFormat("Aguardando...")         # <--- texto inicial
        main_layout.addWidget(self.progress)

        self.log_list = QListWidget()
        self.log_list.setMinimumHeight(200)
        main_layout.addWidget(self.log_list)

        bottom = QHBoxLayout()
        bottom.addStretch()

        self.btn_execute = QPushButton("Executar")
        self.btn_execute.setObjectName("botaoExec")
        self.btn_execute.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))        
        self.btn_execute.setFixedWidth(120)
        self.btn_execute.clicked.connect(self.execute)
        bottom.addWidget(self.btn_execute)

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setFixedWidth(120)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_copy)
        bottom.addWidget(self.btn_cancel)

        main_layout.addLayout(bottom)

        self._thread = None
        self._worker = None


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

    def select_dest_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecione a pasta de destino")
        if folder:
            self.dest_path.setText(folder)
            
    def add_item(self):
        row = {}

        layout = QHBoxLayout()

        row["project"] = QLineEdit()
        row["project"].setPlaceholderText("Projeto")

        row["phase"] = QLineEdit()
        row["phase"].setPlaceholderText("Fase")

        row["path"] = QLineEdit()
        row["path"].setReadOnly(True)
        row["path"].setPlaceholderText("Selecione a pasta")

        btn_browse = QPushButton("📂")
        btn_browse.clicked.connect(lambda _, r=row: self.select_folder(r))

        btn_remove = QPushButton("🗑️")
        btn_remove.clicked.connect(lambda _, r=row, l=layout: self.remove_item(r, l))

        layout.addWidget(row["project"])
        layout.addWidget(row["phase"])
        layout.addWidget(row["path"])
        layout.addWidget(btn_browse)
        layout.addWidget(btn_remove)

        self.items_layout.insertLayout(self.items_layout.count() - 1, layout)
        self.items.append(row)

    def select_folder(self, row):
        folder = QFileDialog.getExistingDirectory(self, "Selecione uma pasta")
        if folder:
            row["path"].setText(folder)

    def remove_item(self, row, layout):
        if row in self.items:
            self.items.remove(row)
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def execute(self):
        self.destino_cut_consolidada = self.dest_path.text().strip()

        if not self.destino_cut_consolidada:
            QMessageBox.warning(self, "Erro", "Selecione a pasta de destino.")
            return
                
        if not self.items:
            QMessageBox.warning(self, "Aviso", "Nenhum item adicionado.")
            return

        rows_data = []
        for i, row in enumerate(self.items, start=1):
            project = row["project"].text().strip()
            phase = row["phase"].text().strip()
            path = row["path"].text().strip()

            if not project or not phase or not path:
                QMessageBox.warning(self, "Erro", f"Item {i} está incompleto.")
                return

            rows_data.append({"project": project, "phase": phase, "path": path})

        palavras_chave = ["CUT"]
        palavras_proibidas = ["CHARTED", "BOM", "LABOR", "COMUNIZADO", "CONSOLIDADO", "DELTA"]

        self.log_list.clear()
        self.progress.setRange(0, 0)                       # indeterminado
        self.progress.setFormat("Preparando...")           # texto inicial

        self.toggle_ui(False)
        self.btn_cancel.setEnabled(True)

        self._thread = QThread()
        self._worker = CopyWorker(rows_data, palavras_chave, palavras_proibidas)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.started.connect(self.on_copy_started)
        self._worker.item_copied.connect(self.on_copy_item)
        self._worker.error.connect(self.on_copy_error)
        self._worker.finished.connect(self.on_copy_finished)

        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()
        
    def cancel_copy(self):
        if self._worker:
            self._worker.cancel()
            self.btn_cancel.setEnabled(False)

    def on_copy_started(self, total, destino):
        self.progress.setRange(0, 0)                        # indeterminado
        self.progress.setFormat("Copiando... 0 arquivos")   # <-- mostra 0
        self.log_list.addItem(f"Destino: {destino}")
        self.log_list.addItem("Processando arquivos...")
        self.log_list.addItem("-" * 40)

    def on_copy_item(self, count, name):
        self.log_list.addItem(f"{count}. {name}")
        self.log_list.scrollToBottom()
        self.progress.setFormat(f"{count} arquivos copiados")  # <-- atualiza

    def on_copy_error(self, msg):
        self.log_list.addItem("[ERRO] " + msg)
        
    def consolidar_cuts_status(self):
        try:
            consolidar_cut(path_destino=self.destino_cut_consolidada)

            QMessageBox.information(
                self,
                "Concluído",
                f"Cuts Consolidadas:\n{self.destino_cut_consolidada}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def on_thread_finished(self):
        # self._worker = None
        # self._thread = None

        # self.toggle_ui(True)

        # if getattr(self, "_pending_consolidation", False):
        #     self._pending_consolidation = False
        #     self.consolidar_cuts_status()
        
        self.btn_cancel.setEnabled(False)

        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        # self.progress.setFormat(f"Concluído: {count} arquivos")

        # self.destino_cut_consolidada = destino

        # marca que precisa consolidar depois que thread morrer
        self._pending_consolidation = True

    def on_copy_finished(self, count, destino):
        self.btn_cancel.setEnabled(False)
        self.toggle_ui(True)

        # fixa o total no final
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.progress.setFormat(f"Concluído: {count} arquivos")  # <-- final


        self.consolidar_cuts_status()
        self._worker = None
        self._thread = None
          
    def toggle_ui(self, enabled):
        self.btn_add.setEnabled(enabled)
        self.btn_execute.setEnabled(enabled)
        for row in self.items:
            row["project"].setEnabled(enabled)
            row["phase"].setEnabled(enabled)
            row["path"].setEnabled(enabled)

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



#-------------------------------------------------------------------
@log_errors_record
def encontrar_arquivos_excel(diretorio: Union[str, Path]) -> Iterable[Path]:
    diretorio = Path(diretorio).resolve()

    if not diretorio.exists():
        raise FileNotFoundError(f"O diretório '{diretorio}' não existe.")
    if not diretorio.is_dir():
        raise NotADirectoryError(f"'{diretorio}' não é um diretório válido.")

    extensoes_validas = {".xls", ".xlsx"}

    for arquivo in diretorio.rglob("*"):
        if arquivo.is_file() and arquivo.suffix.lower() in extensoes_validas and not arquivo.name.startswith("~$"):
            ##logger.debug("Arquivo Excel encontrado: %s", arquivo)
            yield arquivo

@log_errors_record            
def renomear_colunas_duplicadas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renomeia colunas duplicadas adicionando um sufixo '_n' 
    para torná-las únicas.

    Args:
        df (pd.DataFrame): DataFrame com possíveis colunas duplicadas.

    Returns:
        pd.DataFrame: DataFrame com colunas renomeadas.
    """
    colunas_originais = df.columns.tolist()
    colunas_novas = []
    contador = {}

    for col in colunas_originais:
        if col not in contador:
            contador[col] = 0
            colunas_novas.append(col)
        else:
            contador[col] += 1
            novo_nome = f"{col}_{contador[col]}"
            colunas_novas.append(novo_nome)

    df.columns = colunas_novas
    #logger.debug("Colunas renomeadas: %s", df.columns.tolist())
    return df

@log_errors_record            
def identificar_cut_sheets(caminho_arquivo: Union[str, Path]) -> Iterator[Path]:
    """
    Identifica se um arquivo Excel contém Cut Sheets normais ou Charted-Cut.

    Args:
        caminho_arquivo (Path | str): Caminho para o arquivo Excel.

    Returns:
        int: 0 se for Zero Cut normal, 1 se for Charted-Cut.
    
    Raises:
        FileNotFoundError: Se o arquivo não existir.
        ValueError: Se não conseguir ler o Excel.
    """
    caminho_arquivo = Path(caminho_arquivo)

    if not caminho_arquivo.exists():
        raise FileNotFoundError(f"O arquivo '{caminho_arquivo}' não existe.")

    try:
        with caminho_arquivo.open("rb") as f:
            xls = pd.ExcelFile(f)
            lista1 = xls.sheet_names
    except Exception as e:
        #logger.exception("Erro ao ler o arquivo Excel: %s", caminho_arquivo)
        raise ValueError(f"Não foi possível ler o Excel '{caminho_arquivo}'") from e

    # Lista padrão de Cut Sheets
    lista2 = [
        'FG Summary', 'Wire Sheet', 'Multicore Sheet', 
        'Tube Sheet', 'Splice Sheet', 'Multicrimp Sheet'
    ]

    # Contagem de sheets comuns
    num_comum = len(set(lista1) & set(lista2))
    #logger.debug(
    #    "Arquivo '%s': %d sheets comuns encontrados", caminho_arquivo, num_comum
    #)

    # Regra: >=3 → Zero Charted-Cut, caso contrário → Cut normal
    return 0 if num_comum >= 3 else 1

@log_errors_record
def salvar_dict_df_em_excel(dfs_dict: Dict[str, pd.DataFrame],caminho_arquivo: Union[str, Path]) -> None:
    caminho_arquivo = Path(caminho_arquivo).resolve()

    if not isinstance(dfs_dict, dict):
        raise ValueError("dfs_dict deve ser um dicionário de DataFrames.")

    if not dfs_dict or all(df.empty for df in dfs_dict.values()):
        #logger.warning("Nenhum dado disponível. Criando aba padrão 'SemDados'.")
        dfs_dict = {
            "SemDados": pd.DataFrame(
                {"Mensagem": ["Nenhum dado disponível"]}
            )
        }

    #logger.debug(
    #     "Salvando %d abas no arquivo Excel: %s",
    #     len(dfs_dict),
    #     caminho_arquivo
    # )

    try:
        with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
            for nome_aba, df in dfs_dict.items():
                df.to_excel(writer, sheet_name=nome_aba, index=False)

            for nome_aba, df in dfs_dict.items():
                worksheet = writer.sheets[nome_aba]
                for idx, col in enumerate(df.columns, start=1):
                    largura = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.column_dimensions[
                        get_column_letter(idx)
                    ].width = largura

        #logger.info("Arquivo Excel salvo com sucesso: %s", caminho_arquivo)

    except Exception:
        #logger.exception("Erro ao salvar o arquivo Excel: %s", caminho_arquivo)
        raise
    
@log_errors_record
def add_leadset_wire(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona a coluna 'Leadset' ao DataFrame com base nas regras
    específicas de UCS, CSA, Wire Nb e Multicore.

    Args:
        df (pd.DataFrame): DataFrame de entrada.

    Returns:
        pd.DataFrame: DataFrame consolidado com a coluna 'Leadset'.
    """
    colunas_obrigatorias = {"UCS", "Wire Nb", "CSA", "Multicore"}

    if not colunas_obrigatorias.issubset(df.columns):
        faltantes = colunas_obrigatorias - set(df.columns)
        raise KeyError(f"Colunas obrigatórias ausentes: {faltantes}")

    df = df.copy()
    resultado = []

    # =========================
    # GEM / SPIN
    # =========================
    df_gm = df[df["UCS"].str.contains("G|S", case=False, na=False)].copy()

    if not df_gm.empty:
        df_gm["CSA"] = df_gm["CSA"].astype(float)

        diretos = df_gm[df_gm["Multicore"].isna()].copy()
        tw_menor = df_gm[(df_gm["CSA"] < 1.5) & (df_gm["Multicore"].notna())].copy()
        tw_maior = df_gm[(df_gm["CSA"] >= 1.5) & (df_gm["Multicore"].notna())].copy()

        diretos["Leadset"] = diretos["UCS"].astype(str) + diretos["Wire Nb"].astype(str)
        tw_menor["Leadset"] = tw_menor["UCS"].astype(str) + tw_menor["Multicore"].astype(str)
        tw_maior["Leadset"] = tw_maior["UCS"].astype(str) + tw_maior["Wire Nb"].astype(str)

        resultado.extend([diretos, tw_menor, tw_maior])

        #logger.debug("Leadset GEM/SPIN processado (%d linhas)", len(df_gm))

    # =========================
    # VS30
    # =========================
    df_vs30 = df[df["UCS"].str.contains("V", case=False, na=False)].copy()

    if not df_vs30.empty:
        diretos = df_vs30[df_vs30["Multicore"].isna()].copy()
        tw = df_vs30[
            (df_vs30["Multicore"].notna()) &
            (~df_vs30["Multicore"].str.contains("MC", na=False))
        ].copy()
        mc = df_vs30[
            df_vs30["Multicore"].str.contains("MC", na=False)
        ].copy()

        diretos["Leadset"] = diretos["UCS"].astype(str) + diretos["Wire Nb"].astype(str)
        tw["Leadset"] = tw["UCS"].astype(str) + tw["Wire Nb"].astype(str)
        mc["Leadset"] = mc["UCS"].astype(str) + mc["Multicore"].astype(str)

        resultado.extend([diretos, tw, mc])

        #logger.debug("Leadset VS30 processado (%d linhas)", len(df_vs30))

    # =========================
    # U11
    # =========================
    df_u11 = df[df["UCS"].str.contains("B", case=False, na=False)].copy()

    if not df_u11.empty:
        diretos = df_u11[df_u11["Multicore"].isna()].copy()
        tw = df_u11[df_u11["Multicore"].notna()].copy()

        tw["dup_count"] = tw["Multicore"].map(tw["Multicore"].value_counts())

        mc = tw[tw["dup_count"] >= 3].drop(columns="dup_count").copy()
        tw = tw[tw["dup_count"] < 3].drop(columns="dup_count").copy()

        diretos["Leadset"] = diretos["UCS"].astype(str) + diretos["Wire Nb"].astype(str)
        mc["Leadset"] = mc["UCS"].astype(str) + mc["Multicore"].astype(str)
        tw["Leadset"] = tw["UCS"].astype(str) + tw["Wire Nb"].astype(str)

        resultado.extend([diretos, mc, tw])

        #logger.debug("Leadset U11 processado (%d linhas)", len(df_u11))

    # =========================
    # XFD (mantido como está)
    # =========================
    df_xfd = df[df["UCS"].str.contains("X", case=False, na=False)].copy()
    if not df_xfd.empty:
        resultado.append(df_xfd)
        #logger.debug("Leadset XFD incluído (%d linhas)", len(df_xfd))

    if not resultado:
        #logger.warning("Nenhum Leadset foi gerado.")
        return pd.DataFrame()

   #df_final = pd.concat(resultado, ignore_index=True)

    return pd.concat(resultado, ignore_index=True)

@log_errors_record
def adicionar_processos(dados_wire: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona as colunas 'ProcessoA' e 'ProcessoB' ao DataFrame
    com base no mapeamento dos códigos T_1 e T_2.

    Args:
        dados_wire (pd.DataFrame): DataFrame contendo as colunas T_1 e T_2.

    Returns:
        pd.DataFrame: DataFrame atualizado com as colunas de processos.

    Raises:
        ValueError: Se o DataFrame for None ou vazio.
        KeyError: Se as colunas obrigatórias não existirem.
    """
    if dados_wire is None or dados_wire.empty:
        raise ValueError("O DataFrame 'dados_wire' não pode ser None ou vazio.")

    colunas_obrigatorias = {"T_1", "T_2"}
    if not colunas_obrigatorias.issubset(dados_wire.columns):
        faltantes = colunas_obrigatorias - set(dados_wire.columns)
        raise KeyError(f"Colunas obrigatórias ausentes: {faltantes}")

    dict_processos = {
        "N": "Corte",
        "W": "Prensa - Duplas",
        "D": "Prensa - Duplas",
        "M": "Prensa - Multicrimp",
        "S": "Splice",
    }

    dados_wire = dados_wire.copy()

    dados_wire["ProcessoA"] = dados_wire["T_1"].map(dict_processos)
    dados_wire["ProcessoB"] = dados_wire["T_2"].map(dict_processos)

    #logger.debug(
    #     "Processos adicionados com sucesso (%d linhas)",
    #     len(dados_wire)
    # )

    return dados_wire

@log_errors_record
def adicionar_bitolas_em_splices(df_wire: pd.DataFrame, df_splices: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona as bitolas (CSA) nos splices e gera combinações de bitolas
    lado esquerdo / direito, com IDs e combinações ajustadas.

    Args:
        df_wire (pd.DataFrame): DataFrame com colunas 'UCS', 'Wire Nb' e 'CSA'.
        df_splices (pd.DataFrame): DataFrame com colunas 'UCS' e 'L1'-'L10', 'R1'-'R10'.

    Returns:
        pd.DataFrame: df_splices atualizado com bitolas, combinações e IDs.
    """
    # ---------------------------------------------------------
    # 1️⃣ Cria dicionário de mapeamento UCS + Wire Nb -> CSA
    # ---------------------------------------------------------
    dict_wire = {f"{row['UCS']}{row['Wire Nb']}": row['CSA'] for _, row in df_wire.iterrows()}

    colunas = [f"L{i}" for i in range(1, 11)] + [f"R{i}" for i in range(1, 11)]

    for col in colunas:
        df_splices[f"{col}_CSA"] = df_splices.apply(
            lambda x: dict_wire.get(f"{x['UCS']}{x[col]}"), axis=1
        )

    # ---------------------------------------------------------
    # 2️⃣ Gera combinações lado esquerdo / direito
    # ---------------------------------------------------------
    colunas_esq = [f"L{i}_CSA" for i in range(1, 11)]
    colunas_dir = [f"R{i}_CSA" for i in range(1, 11)]

    for idx in df_splices.index:
        # Lados esquerdo/direito: ordena e concatena como string
        esq = sorted(df_splices.loc[idx, colunas_esq].dropna().astype(float).tolist())
        dir_ = sorted(df_splices.loc[idx, colunas_dir].dropna().astype(float).tolist())

        df_splices.loc[idx, 'Combinação - Esq'] = ' + '.join(f"{v:g}" for v in esq)
        df_splices.loc[idx, 'Combinação - Dir'] = ' + '.join(f"{v:g}" for v in dir_)

    # ---------------------------------------------------------
    # 3️⃣ Cria IDs únicos para cada combinação
    # ---------------------------------------------------------
    lista_keys = sorted(
        set(df_splices['Combinação - Esq'].dropna().unique().tolist() +
            df_splices['Combinação - Dir'].dropna().unique().tolist()),
        reverse=True
    )
    df_keys = pd.DataFrame({'Key': lista_keys})
    df_keys['ID'] = range(1, len(df_keys) + 1)

    dict_key = dict(zip(df_keys['Key'], df_keys['ID']))
    dict_key_inv = dict(zip(df_keys['ID'], df_keys['Key']))

    df_splices['Comb_Esq_ID'] = df_splices['Combinação - Esq'].map(dict_key)
    df_splices['Comb_Dir_ID'] = df_splices['Combinação - Dir'].map(dict_key)

    # Garantir que Comb_Esq_ID <= Comb_Dir_ID
    mask_invert = df_splices['Comb_Esq_ID'] > df_splices['Comb_Dir_ID']
    df_splices.loc[mask_invert, ['Comb_Esq_ID', 'Comb_Dir_ID']] = \
        df_splices.loc[mask_invert, ['Comb_Dir_ID', 'Comb_Esq_ID']].values

    # ---------------------------------------------------------
    # 4️⃣ Mapear combinações invertidas
    # ---------------------------------------------------------
    df_splices['Comb_Esq_Inv'] = df_splices['Comb_Esq_ID'].map(dict_key_inv)
    df_splices['Comb_Dir_Inv'] = df_splices['Comb_Dir_ID'].map(dict_key_inv)

    # ---------------------------------------------------------
    # 5️⃣ Função para combinar lados com pipe de forma segura
    # ---------------------------------------------------------
    def combinar_lados(esq, dir_):
        esq_valido = pd.notna(esq) and str(esq).strip() != ''
        dir_valido = pd.notna(dir_) and str(dir_).strip() != ''

        if esq_valido and dir_valido:
            return f"{esq} | {dir_}"
        if esq_valido:
            return str(esq)
        if dir_valido:
            return str(dir_)
        return ''

    df_splices['Combinação - Original'] = df_splices.apply(
        lambda r: combinar_lados(r['Combinação - Esq'], r['Combinação - Dir']),
        axis=1
    )

    df_splices['Combinação - Ajustada'] = df_splices.apply(
        lambda r: combinar_lados(r['Comb_Esq_Inv'], r['Comb_Dir_Inv']),
        axis=1
    )

    # ---------------------------------------------------------
    # 6️⃣ Limpeza de colunas temporárias
    # ---------------------------------------------------------
    df_splices.drop(columns=['Comb_Esq_Inv', 'Comb_Dir_Inv', 'Comb_Esq_ID', 'Comb_Dir_ID'],
                     errors='ignore', inplace=True)

    # ---------------------------------------------------------
    # 7️⃣ Reordenar colunas (opcional)
    # ---------------------------------------------------------
    colunas_ord = ['Nome do arquivo','Projeto','Fase','UCS','Splice Nb', 'Int.PN','Extra Component PN','Pack',
                   'CSA Left', 'CSA Right','CSA Total','Node'] + colunas + [f"{c}_CSA" for c in colunas] + \
                  ['Combinação - Esq','Combinação - Dir','Combinação - Original','Combinação - Ajustada']

    df_splices = reordenar_colunas(df_splices, colunas_prioritarias=colunas_ord)
    
    df_splices.sort_values(by=['Combinação - Ajustada'],inplace=True)

    return df_splices

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
def criar_multicrimp(dados) -> pd.DataFrame:
    def sort_key(s):
        # Remove MULTI-
        s = s.replace("MULTI-", "")

        # Regex para separar partes: letras, número, sufixo
        match = re.match(r"([A-Z]+)(\d*)(?:_(\w))?", s)
        if match:
            prefix = match.group(1)
            number = int(match.group(2)) if match.group(2).isdigit() else float("inf")
            suffix = match.group(3) or ""
            return (prefix, number, suffix)
        else:
            # Se não casar, coloca no final
            return (s, float("inf"), "")

    for col in ["Joint 1", "Joint 2", "Note 1", "Note 2"]:
        if col not in dados.columns:
            dados[col] = None

    filtro = (
        dados["Note 1"].str.startswith("M", na=False)
        | dados["Note 2"].str.startswith("M", na=False)
    ) & (
        dados["Joint 1"].str.startswith("S", na=False)
        | dados["Joint 2"].str.startswith("S", na=False)
    )

    mc = dados[filtro][["Joint 1", "Joint 2", "Note 1", "Note 2"]]

    filtro = (
        dados["Note 1"].str.startswith("M", na=False)
        | dados["Note 2"].str.startswith("M", na=False)
    ) & (
        ~dados["Joint 1"].str.startswith("S", na=False)
        & ~dados["Joint 2"].str.startswith("S", na=False)
    )

    mult = dados[filtro][["Joint 1", "Joint 2", "Note 1", "Note 2"]]
    # list_mult = mult['Note 2'].unique().tolist()

    list_completa = pd.concat([mc, mult], ignore_index=True)
    list_completa = list_completa.drop_duplicates().reset_index(drop=True)

    cols_validas = [
        c
        for c in list_completa.columns
        if list_completa[c].astype(str).str.startswith(("S", "M"), na=False).any()
    ]

    list_completa = list_completa[cols_validas].reset_index(drop=True)

    # Encontra os valores únicos de "Note 2" onde "Joint 1" NÃO é NaN
    if "Joint 1" not in list_completa:
        list_completa["Joint 1"] = None

    if "Note 2" not in list_completa:
        list_completa["Note 2"] = None

    valid_notes = set(
        list_completa.loc[list_completa["Joint 1"].notna(), "Note 2"].unique()
    )

    # Remove linhas onde "Joint 1" é NaN E "Note 2" já existe em valid_notes
    list_completa = list_completa[
        ~(list_completa["Joint 1"].isna() & list_completa["Note 2"].isin(valid_notes))
    ].reset_index(drop=True)

    if "Joint 1" not in list_completa.columns:
        list_completa["Joint 1"] = None
    if "Joint 2" not in list_completa.columns:
        list_completa["Joint 2"] = None
    if "Note 1" not in list_completa.columns:
        list_completa["Note 1"] = None
    if "Note 2" not in list_completa.columns:
        list_completa["Note 2"] = None

    df_new = pd.DataFrame()

    list_completa["Note 1"] = None
    mcs = pd.concat([list_completa["Note 1"], list_completa["Note 2"]])
    mcs = mcs.dropna().astype(str).unique().tolist()
    mcs = [v for v in mcs if v.startswith("MU")]

    mcs = sorted(mcs, key=sort_key)

    s = None

    number_mc = 1
    for m in mcs:
        df_m = list_completa[
            (list_completa["Note 1"] == m) | (list_completa["Note 2"] == m)
        ]

        valores = pd.concat(
            [df_m["Joint 1"], df_m["Joint 2"], df_m["Note 1"], df_m["Note 2"]]
        )

        valores = valores.dropna().astype(str).unique().tolist()

        valores = [v for v in valores if v.startswith("SP") or v.startswith("MU")]

        valores_validos = [v for v in valores if pd.notna(v) and v not in [None, ""]]
        idx = dados[dados.isin(valores_validos).any(axis=1)].index
        df_prov = dados.loc[idx]  # .reset_index(drop=True)
        coluna_tag = df_prov.columns.get_loc("Note 2") + 1

        dict_list = defaultdict(list)

        colunas_list = df_prov.iloc[:, :coluna_tag].columns.to_list()

        contador_mult1 = 1
        contador_mult2 = 1
        total_hist = 0

        for i in df_prov.columns[coluna_tag:]:
            df_prov2 = df_prov[df_prov[i].notna()]  # .reset_index(drop=True)
            number = "-".join(map(str, df_prov2.index.to_list()))
            if number != "":
                dict_list[number].append(i)

        for d in dict_list.keys():
            colunas_list = df_prov.iloc[:, :coluna_tag].columns.to_list()
            colunas_list = colunas_list + dict_list[d]

            df_prov3 = df_prov[df_prov[dict_list[d][0]].notna()].reset_index(drop=True)

            df_prov3[dict_list[d]] = "X"

            try:
                splices = pd.concat([df_prov3["Joint 1"], df_prov3["Joint 2"]])
                splices = splices.dropna().astype(str).unique().tolist()
                splices = [v for v in splices if v.startswith("SP")]
                s = "+".join(splices)
            except:
                s = ""

            if s != "":
                df_prov3["Nome Processo"] = f"MC{number_mc}"
                df_prov3["Número"] = number_mc
                number_mc += 1

                total = df_prov3[df_prov3["Note 2"] == m].shape[0]
                if total_hist == 0:
                    df_prov3["M-Crimp"] = f"{m} ({contador_mult2:02d})"
                    total_hist = total
                    df_prov3["Nome AV"] = f"{m} ({contador_mult2:02d})+{s}"
                    colunas_list = (
                        colunas_list
                        + ["Nome Processo"]
                        + ["Número"]
                        + ["M-Crimp"]
                        + ["Nome AV"]
                    )

                elif total_hist != total:
                    contador_mult2 += 1
                    df_prov3["M-Crimp"] = f"{m} ({contador_mult2:02d})"
                    total_hist = total
                    df_prov3["Nome AV"] = f"{m} ({contador_mult2:02d})+{s}"
                    colunas_list = (
                        colunas_list
                        + ["Nome Processo"]
                        + ["Número"]
                        + ["M-Crimp"]
                        + ["Nome AV"]
                    )

                else:
                    df_prov3["M-Crimp"] = f"{m} ({contador_mult2:02d})"
                    df_prov3["Nome AV"] = f"{m} ({contador_mult2:02d})+{s}"
                    colunas_list = (
                        colunas_list
                        + ["Nome Processo"]
                        + ["Número"]
                        + ["M-Crimp"]
                        + ["Nome AV"]
                    )

            else:
                df_prov3["Nome Processo"] = f"{m} ({contador_mult1:02d})"
                df_prov3["M-Crimp"] = df_prov3["Nome Processo"]
                colunas_list = colunas_list + ["Nome Processo"] + ["M-Crimp"]
                if len(df_prov3) > 0:
                    contador_mult1 += 1

            # if 'Joint 1' not in df_prov3[colunas_list].columns:
            #     df_prov3[colunas_list]['Joint 1'] = None
            # df_prov3[colunas_list] = df_prov3[colunas_list].sort_values(by=['Term. 2','Joint 2','Joint 1'],).reset_index(drop=True)

            # Cria as colunas se não existirem
            for col in ["Joint 1", "Joint 2"]:
                if col not in df_prov3.columns:
                    df_prov3[col] = None

            #df_prov3["Joint"] = df_prov3["Joint 1"].fillna(df_prov3["Joint 2"])
            df_prov3["Joint"] = (df_prov3["Joint 1"].combine_first(df_prov3["Joint 2"]).infer_objects(copy=False))

            
            # Agora ordena com segurança
            df_prov3 = df_prov3.sort_values(by=["Joint", "Term. 2"]).reset_index(
                drop=True
            )

            df_prov3.drop(columns=["Joint"], inplace=True)

            # Mantém apenas as colunas desejadas
            df_prov3 = df_prov3[colunas_list]

            df_new = pd.concat([df_new, df_prov3[colunas_list]], ignore_index=True)

    df_new["Combinação"] = None
    df_new["Circuitos Amarração"] = None
    df_new["Terminal"] = None
    df_new["Ckts M Crimp"] = None

    if "Joint 1" not in df_new:
        df_new["Joint 1"] = None

    for i in df_new.index:
        if str(df_new.loc[i, "Joint 1"])[:2] == "SP":
            df_new.loc[i, "Combinação"] = str(df_new.loc[i, "Joint 1"])
            df_new.loc[i, "Circuitos Amarração"] = str(df_new.loc[i, "Leadset"])
            df_new.loc[i, "Terminal"] = str(df_new.loc[i, "Term. 2"])
        if str(df_new.loc[i, "Joint 2"])[:2] == "SP":
            df_new.loc[i, "Combinação"] = str(df_new.loc[i, "Joint 2"])
            df_new.loc[i, "Circuitos Amarração"] = str(df_new.loc[i, "Leadset"])
            if pd.isna(df_new.loc[i, "Term. 2"]):
                term = ""
            else:
                term = df_new.loc[i, "Term. 2"]
            df_new.loc[i, "Terminal"] = term

        if str(df_new.loc[i, "Note 2"])[:2] == "MU":
            df_new.loc[i, "Ckts M Crimp"] = str(df_new.loc[i, "Leadset"])

    colunas_ord = [
        "Nome Processo",
        "Nome AV",
        "M-Crimp",
        "Terminal",
        "Ckts M Crimp",
        "Combinação",
        "Circuitos Amarração",
    ]
    df_new = reordenar_colunas(df_new, colunas_prioritarias=colunas_ord)
    df_new = df_new.drop_duplicates().reset_index(drop=True)
    try:
        df_new.drop(columns=["Número"], inplace=True)
    except:
        pass
    colunas_ord = ["Nome Processo","Nome AV","M-Crimp","Terminal","Ckts M Crimp","Combinação","Circuitos Amarração"
    ]
    df_new = reordenar_colunas(df_new, colunas_prioritarias=colunas_ord)
    return df_new

@log_errors_record
def adicionar_mult_estudo(dados: pd.DataFrame) -> pd.DataFrame:

    def create_position_tracker(lista_dados: list, dict_df: dict):
        new_list = []

        lista_dados.sort()
        for i in lista_dados:
            # se ainda não existe no dict
            if i not in dict_df.keys():
                # se o dict está vazio, começa em 1
                if not dict_df:
                    dict_df[i] = 1
                else:
                    dict_df[i] = max(dict_df.values()) + 1

        for i in dict_df.keys():
            if i not in lista_dados:
                new_list.append(None)
            else:
                new_list.append(i)

        return new_list

    def sort_key(s):
        # Remove MULTI-
        s = s.replace("MULTI-", "")

        # Regex para separar partes: letras, número, sufixo
        match = re.match(r"([A-Z]+)(\d*)(?:_(\w))?", s)
        if match:
            prefix = match.group(1)
            number = int(match.group(2)) if match.group(2).isdigit() else float("inf")
            suffix = match.group(3) or ""
            return (prefix, number, suffix)
        else:
            # Se não casar, coloca no final
            return (s, float("inf"), "")

    for col in ["Joint 1", "Joint 2", "Note 1", "Note 2"]:
        if col not in dados.columns:
            dados[col] = None

    filtro = (
        dados["Note 1"].str.startswith("M", na=False)
        | dados["Note 2"].str.startswith("M", na=False)
    ) & (
        dados["Joint 1"].str.startswith("S", na=False)
        | dados["Joint 2"].str.startswith("S", na=False)
    )

    mc = dados[filtro][["Joint 1", "Joint 2", "Note 1", "Note 2"]]

    filtro = (
        dados["Note 1"].str.startswith("M", na=False)
        | dados["Note 2"].str.startswith("M", na=False)
    ) & (
        ~dados["Joint 1"].str.startswith("S", na=False)
        & ~dados["Joint 2"].str.startswith("S", na=False)
    )

    mult = dados[filtro][["Joint 1", "Joint 2", "Note 1", "Note 2"]]
    # list_mult = mult['Note 2'].unique().tolist()

    list_completa = pd.concat([mc, mult], ignore_index=True)
    list_completa = list_completa.drop_duplicates().reset_index(drop=True)

    # cols_validas = [
    #     c
    #     for c in list_completa.columns
    #     if list_completa[c].astype(str).str.startswith(("S", "M"), na=False).any()
    # ]

    # list_completa = list_completa[cols_validas].reset_index(drop=True)
    
    # Colunas que nunca devem ser removidas
    colunas_essenciais = ["Joint 1", "Joint 2", "Note 1", "Note 2"]

    # Filtra apenas colunas que não são essenciais
    cols_validas = [
        c for c in list_completa.columns 
        if c in colunas_essenciais or list_completa[c].astype(str).str.startswith(("S","M"), na=False).any()
    ]

    list_completa = list_completa[cols_validas].reset_index(drop=True)


    # Encontra os valores únicos de "Note 2" onde "Joint 1" NÃO é NaN
    if "Joint 1" not in list_completa:
        list_completa["Joint 1"] = None

    if "Note 2" not in list_completa:
        list_completa["Note 2"] = None

    valid_notes = set(
        list_completa.loc[list_completa["Joint 1"].notna(), "Note 2"].unique()
    )

    # Remove linhas onde "Joint 1" é NaN E "Note 2" já existe em valid_notes
    list_completa = list_completa[
        ~(list_completa["Joint 1"].isna() & list_completa["Note 2"].isin(valid_notes))
    ].reset_index(drop=True)

    if "Joint 1" not in list_completa.columns:
        list_completa["Joint 1"] = None
    if "Joint 2" not in list_completa.columns:
        list_completa["Joint 2"] = None
    if "Note 1" not in list_completa.columns:
        list_completa["Note 1"] = None
    if "Note 2" not in list_completa.columns:
        list_completa["Note 2"] = None

    list_completa["Note 1"] = None
    mcs = pd.concat([list_completa["Note 1"], list_completa["Note 2"]])
    mcs = mcs.dropna().astype(str).unique().tolist()
    mcs = [v for v in mcs if v.startswith("MU")]

    mcs = sorted(mcs, key=sort_key)

    lista_combinacoes = []

    # for c in list_completa.columns:
    #     for j in list_completa.index:
    #         if (
    #             str(list_completa.loc[j, c])[:1] != "S"
    #             and str(list_completa.loc[j, c])[:1] != "M"
    #         ):
    #             list_completa.loc[j, c] = None
                
    # mask = ~list_completa.apply(lambda col: col.astype(str).str.startswith(("S","M")))
    # list_completa[mask] = None
    
    # mask = ~list_completa.applymap(lambda x: str(x).startswith(("S","M")))
    # list_completa = list_completa.mask(mask)


    mask = ~list_completa.apply(
        lambda col: col.astype(str).str.startswith(("S", "M"))
    )

    list_completa = list_completa.mask(mask)


    for m in mcs:
        mc = list(
            set(
                list_completa[
                    (list_completa["Note 1"] == m) | (list_completa["Note 2"] == m)
                ]["Note 1"]
                .dropna()
                .tolist()
                + list_completa[
                    (list_completa["Note 1"] == m) | (list_completa["Note 2"] == m)
                ]["Note 2"]
                .dropna()
                .tolist()
            )
        )
        splices = list(
            set(
                list_completa[
                    (list_completa["Note 1"] == m) | (list_completa["Note 2"] == m)
                ]["Joint 1"]
                .dropna()
                .tolist()
                + list_completa[
                    (list_completa["Note 1"] == m) | (list_completa["Note 2"] == m)
                ]["Joint 2"]
                .dropna()
                .tolist()
            )
        )
        if len(splices) > 0:
            mc = mc + splices

        lista_combinacoes.append(mc)


    df_multicrimp = pd.DataFrame()

    for m in lista_combinacoes:
        df_prov = dados[
            (dados["Note 1"].isin(m))
            | (dados["Note 2"].isin(m))
            | (dados["Joint 1"].isin(m))
            | (dados["Joint 2"].isin(m))
        ]

        coluna_tag = df_prov.columns.get_loc("Note 2") + 1

        number_derivativo = 1
        for c in df_prov.columns[coluna_tag:]:
            colunas_base = df_prov.columns[:coluna_tag].to_list()
            colunas_sem_pn = df_prov.columns[:coluna_tag].to_list()
            ", ".join(df_prov[df_prov[c].notna()].index.astype(str).tolist())

            lista_splices = list(
                set(
                    pd.concat(
                        [
                            df_prov[df_prov[c].notna()]["Joint 1"]
                            .dropna()
                            .loc[lambda s: s.str.startswith("S", na=False)],
                            df_prov[df_prov[c].notna()]["Joint 2"]
                            .dropna()
                            .loc[lambda s: s.str.startswith("S", na=False)],
                        ]
                    )
                    .unique()
                    .tolist()
                )
            )

            mcs = [x for x in m if str(x).startswith("M")]

            lista_de_circuitos = list(
                set(
                    dados[(dados[c].notna()) & (dados["Joint 1"].isin(lista_splices))][
                        "Wire Nb"
                    ].to_list()
                    + dados[
                        (dados[c].notna()) & (dados["Joint 2"].isin(lista_splices))
                    ]["Wire Nb"].to_list()
                )
            )

            lista_de_circuitos_mc = list(
                set(
                    dados[(dados[c].notna()) & (dados["Note 2"].isin(mcs))][
                        "Wire Nb"
                    ].to_list()
                    + dados[(dados[c].notna()) & (dados["Note 2"].isin(mcs))][
                        "Wire Nb"
                    ].to_list()
                )
            )

            lista_de_circuitos = list(set(lista_de_circuitos + lista_de_circuitos_mc))

            lista_de_circuitos = list({str(i) for i in lista_de_circuitos})

            contador = 1
            texto = "MC.Splice"
            df_prov = df_prov.copy()
            for a in lista_splices:
                #df_prov[texto + str(contador)] = a
                df_prov.loc[:, f"MC.Splice{contador}"] = a

                colunas_base = colunas_base + [texto + str(contador)]
                colunas_sem_pn = colunas_sem_pn + [texto + str(contador)]
                contador += 1

            # Criar circuitos splices
            contador = 1
            texto = "MC.W"
            df_prov = df_prov.copy()
            for a in lista_de_circuitos:
                #df_prov[texto + str(contador)] = a
                df_prov.loc[:, f"MC.W{contador}"] = a

                colunas_base = colunas_base + [texto + str(contador)]
                colunas_sem_pn = colunas_sem_pn + [texto + str(contador)]
                contador += 1

            colunas_base = colunas_base + [c]
            colunas_dell = [
                "ProcessoA",
                "ProcessoB",
                "Leadset",
                "Wire Nb",
                "Multicore",
                "T",
                "CSA",
                "C1",
                "C2",
                "Length",
                "Int.PN",
                "Strip 1",
                "Strip 2",
                "T_1",
                "T_2",
                "Seal 1",
                "Seal 2",
                "Joint 1",
                "Joint 2",
                "Note 1",
                "Node 1",
                "Node 2",
            ]
            df_linha = (
                df_prov[(df_prov[c].notna()) & (dados["Term. 2"].notna())][colunas_base]
                .drop(columns=colunas_dell, errors="ignore")
                .iloc[:1]
                .reset_index(drop=True)
            )

            colunas_sem_pn = list(set(colunas_sem_pn) - set(colunas_dell))

            if len(df_multicrimp) > 0:
                try:
                    mask = (
                        df_multicrimp[colunas_sem_pn]
                        .apply(tuple, axis=1)
                        .isin(df_linha[colunas_sem_pn].apply(tuple, axis=1))
                    )
                    indices_existentes = df_multicrimp.index[mask].tolist()
                    if len(indices_existentes) > 0:
                        df_multicrimp.loc[indices_existentes, c] = number_derivativo
                    else:
                        df_linha = df_linha.dropna(axis=1, how="all")
                        df_multicrimp = pd.concat([df_multicrimp, df_linha], ignore_index=True)
                except:
                    df_linha = df_linha.dropna(axis=1, how="all")
                    df_multicrimp = pd.concat([df_multicrimp, df_linha], ignore_index=True)
            else:
                df_multicrimp = pd.concat([df_multicrimp, df_linha], ignore_index=True)

            number_derivativo += 1

    df_multicrimp = df_multicrimp.drop(columns=["Term. 1"], errors="ignore").rename(columns={"Term. 2": "MC.Terminal"})
    colunas_ord = ["Nome do arquivo","Projeto","Fase","UCS","Term. 1","Term. 2","MC.Terminal","Note 2",
                   "MC.Splice1","MC.Splice2","MC.Splice3","MC.Splice4","MC.Splice5","MC.Splice6","MC.Splice7","MC.Splice8","MC.Splice9","MC.Splice10",
                   "MC.W1","MC.W2","MC.W3","MC.W4","MC.W5","MC.W6","MC.W7","MC.W8","MC.W9","MC.W10",
                   "MC.W11","MC.W12","MC.W13","MC.W14","MC.W15","MC.W16","MC.W17","MC.W18","MC.W19","MC.W20"]
    
    df_multicrimp = reordenar_colunas(df_multicrimp, colunas_prioritarias=colunas_ord)

   

    return df_multicrimp

@log_errors_record
def ajustar_multicrimp_estudo(df_estudo:pd.DataFrame=None,df_normal:pd.DataFrame=None):

    colunas_alvo =["MC.W1","MC.W2","MC.W3","MC.W4","MC.W5","MC.W6","MC.W7","MC.W8","MC.W9","MC.W10",
                "MC.W11","MC.W12","MC.W13","MC.W14","MC.W15","MC.W16","MC.W17","MC.W18","MC.W19","MC.W20"]

    colunas_existentes = [c for c in colunas_alvo if c in df_estudo.columns]

    df_estudo = df_estudo.copy()  # evita SettingWithCopyWarning

    for i in df_estudo.index:
        item = df_estudo.loc[i, 'Note 2']
        lista_ckt = df_estudo.loc[i, colunas_existentes].dropna().tolist()
        ucs = df_estudo.loc[i, 'UCS']
        lista_ckt = list(dict.fromkeys([f"{ucs}{x}" for x in lista_ckt]))  # mantém ordem

        nome = df_normal[df_normal['M-Crimp'].fillna('').str.contains(item)]['Nome Processo'].unique().tolist()

        for j in nome:
            df_temp = df_normal[df_normal['Nome Processo']==j]['Leadset'].tolist()
            if set(lista_ckt) == set(df_temp):  # comparação sem se preocupar com ordem
                df_estudo.loc[i, 'Nome Processo'] = j
                break
            
    df_estudo.rename(columns={'Note 2':'M-Crimp'},inplace=True)
    colunas_ord = ['Nome do arquivo','Projeto','Fase','UCS','MC.Terminal','Nome Processo','M-Crimp']
    df_estudo = reordenar_colunas(df_estudo,colunas_prioritarias=colunas_ord)
    return df_estudo

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
def consultar_arquivos_base(id_name: str, show_display: bool = False) -> str:
    """
    Consulta o arquivo de configuração para obter o caminho (network ou local)
    de um determinado ID.

    Args:
        id_name (str): Nome do ID a ser consultado.
        show_display (bool, optional): Se True, exibe a tabela completa de configuração.
                                       Defaults to False.

    Returns:
        str: Caminho válido (network ou local) correspondente ao ID.

    Raises:
        FileNotFoundError: Se o arquivo de configuração não existir.
        ValueError: Se o ID não for encontrado ou não houver caminho válido.
        Exception: Para erros inesperados ao ler o Excel.
    """
    config_file = Path(CONFIG_PATH)

    if not config_file.exists():
        #logger.error("Arquivo de configuração não encontrado: %s", config_file)
        raise FileNotFoundError(f"Arquivo de configuração não encontrado: {config_file}")

    try:
        df = pd.read_excel(config_file)
        #logger.debug("Arquivo de configuração lido com sucesso: %s", config_file)
    except Exception as e:
        #logger.exception("Erro ao ler o arquivo de configuração: %s", config_file)
        raise Exception(f"Erro ao ler o arquivo de configuração: {e}") from e

    if show_display:
        table = tabulate(df, headers="keys", tablefmt="grid")
        print(table)

    # Filtra as linhas pelo ID
    df_filtrado = df[df["Id name"] == id_name]

    if df_filtrado.empty:
        #logger.warning("ID '%s' não encontrado no arquivo de configuração.", id_name)
        raise ValueError(f"ID '{id_name}' não encontrado no arquivo de configuração.")

    # Recupera paths
    path_network = df_filtrado.iloc[0].get("Network [Path]")
    path_local = df_filtrado.iloc[0].get("Local [Path]")

    # Verifica paths válidos
    if pd.notna(path_network) and os.path.exists(path_network):
        #logger.debug("Caminho network válido encontrado para '%s': %s", id_name, path_network)
        return path_network
    elif pd.notna(path_local) and os.path.exists(path_local):
        #logger.debug("Caminho local válido encontrado para '%s': %s", id_name, path_local)
        return path_local
    else:
        #logger.error("O ID '%s' não possui caminho válido.", id_name)
        raise ValueError(f"O ID '{id_name}' não possui caminho válido.")

@log_errors_record
def adicionar_codigo_ucs(dados: pd.DataFrame,coluna_tag: Optional[str] = None) -> pd.DataFrame:
    """
    Adiciona a coluna 'UCS' ao DataFrame, mapeando arquivos para códigos UCS
    com base em um arquivo de referência.

    Args:
        dados (pd.DataFrame): DataFrame contendo os dados dos arquivos.
        coluna_tag (str, optional): Coluna opcional para identificar os arquivos.
                                     Se fornecida, a função usará esta coluna
                                     em vez de 'Nome do arquivo'.

    Returns:
        pd.DataFrame: DataFrame atualizado com a coluna 'UCS'.

    Raises:
        KeyError: Se a coluna necessária não existir no DataFrame.
        Exception: Se houver problema ao ler o arquivo de UCS.
    """
    # if dados is None or dados.empty:
    #     #logger.error("O DataFrame fornecido está vazio ou é None.")
    #     raise ValueError("O DataFrame fornecido não pode ser vazio ou None.")

    # Lê o arquivo de UCS
    try:
        caminho_ucs = BASE_DIR / "data" / "Lista_de_zmm247.json"
        df_ucs = pd.read_json(caminho_ucs,dtype=str)
        
        # caminho_ucs = consultar_arquivos_base("codigo_ucs")
        # df_ucs = pd.read_excel(caminho_ucs)
        #logger.debug("Arquivo de UCS carregado com sucesso: %s", caminho_ucs)
    except Exception as e:
        raise Exception(f"Erro ao carregar o arquivo de UCS: {e}") from e

    # Cria o dicionário para mapear "Familia Externa" -> "Fam Int"
    dict_ucs = {
        str(row["External Family"]).replace('.', ''): row["Internal Family"]
        for _, row in df_ucs.iterrows()
    }

    # Verifica se a coluna necessária existe
    coluna_referencia = coluna_tag if coluna_tag else "Nome do arquivo"
    if coluna_referencia not in dados.columns:
        #logger.error("Coluna necessária '%s' não encontrada no DataFrame.", coluna_referencia)
        raise KeyError(f"A coluna '{coluna_referencia}' não existe no DataFrame fornecido.")

    # Função interna que encontra UCS para cada arquivo
    def encontrar_ucs(nome_arquivo: str) -> str:
        for chave, valor in dict_ucs.items():
            if pd.notna(chave) and str(chave) in nome_arquivo:
                return valor
        return "Verificar"

    # Remove espaços e aplica o mapeamento
    dados["UCS"] = (
        dados[coluna_referencia]
        .astype(str)
        .str.replace(" ", "", regex=False)
        .apply(encontrar_ucs)
    )

    #logger.debug("Coluna 'UCS' adicionada com sucesso.")
    return dados

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
def arrumar_colunas(df_dados: pd.DataFrame, coluna_alvo: str) -> pd.DataFrame:
    """
    Reordena colunas do DataFrame conforme o tipo informado em coluna_alvo.
    """

    if df_dados is None or df_dados.empty:
        return df_dados

    if not coluna_alvo:
        return df_dados
    
    if coluna_alvo == "Wire_Nb":
        colunas_ord = ['Nome do arquivo', 'Projeto', 'Fase', 'ProcessoA', 'ProcessoB', 'UCS',
                      'Leadset', 'Wire Nb', 'Multicore','Sub-Assembly', 'T', 'CSA', 'Length', 'C1', 'C2',
                      'Int.PN', 'Term. 1', 'Strip 1', 'Seal 1', 'Node 1', 'Joint 1', 'T_1',
                      'Note 1', 'Term. 2', 'Strip 2', 'Seal 2', 'Node 2', 'Joint 2', 'T_2','Note 2']

    elif coluna_alvo == "Splice_Nb":
        colunas_lr = [f"L{i}" for i in range(1, 11)] + [f"R{i}" for i in range(1, 11)]

        colunas_ord = (
            ['Nome do arquivo','Projeto','Fase','UCS','Splice Nb','Int.PN',
             'Extra Component PN','Pack','CSA Left','CSA Right','CSA Total','Node']
            + colunas_lr
            + [f"{c}_CSA" for c in colunas_lr]
            + ['Combinação - Esq','Combinação - Dir',
               'Combinação - Original','Combinação - Ajustada']
        )

    elif coluna_alvo == "Multcrimp_Nb - Estudo":
        colunas_ord = [
            "Nome do arquivo","Projeto","Fase","UCS","Term. 1","Term. 2",
            "Nome Processo","M-Crimp","MC.Terminal",
            *[f"MC.Splice{i}" for i in range(1, 11)],
            *[f"MC.W{i}" for i in range(1, 21)],
        ]

    elif coluna_alvo == "MultWire_Nb":
        # colunas_ord = ["Nome do arquivo","Projeto","Fase","UCS","Mult.Wire Nb","T","CSA",
        #                "Length","Wire Spec","Pack","CutBack1","CutBack2","W1","W2","W3","W4","Int.PN"]
        colunas_ord = ['Nome do arquivo', 'Projeto', 'Fase','UCS', 'Mult.Wire Nb','T',
                       'Length', 'Int.PN','CSA','CutBack1', 'CutBack2','Wire Spec','Pack',
                       'Inc.BOM','Inc.Chart', 'W1', 'W2','W3','W4']

    else:
        # Caso desconhecido → retorna sem alterar
        return df_dados

    return reordenar_colunas(df_dados, colunas_prioritarias=colunas_ord)

@log_errors_record
def format_cut(path):
    with open(path, "rb") as f:
        df = pd.read_excel(f)

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

        df_novo = (
            df_novo.dropna(how="all", axis=0)
            .dropna(how="all", axis=1)
            .reset_index(drop=True)
        )

        df_novo = renomear_colunas_duplicadas(df_novo)

        df_novo["Nome do arquivo"] = os.path.basename(path).replace(".xlsx", "")
        
        match = re.match(r'^([^_]+)_([^_]+)_', os.path.basename(path).replace(".xlsx", ""))
        
        if match:
            df_novo["Projeto"] = match.group(1)
            df_novo["Fase"] = match.group(2)
        else:
            df_novo["Projeto"] = None
            df_novo["Fase"] = None
        if name == "Wire_Nb":
            try:
                list_add = [
                    'Sub-Assembly',
                    "Multicore",
                    "Term. 1",
                    "Term. 2",
                    "Seal 1",
                    "Seal 2",
                    "Joint 1",
                    "Joint 2",
                ]
                for l in list_add:
                    if l not in df_novo.columns:
                        df_novo[l] = None

                # Processos
                if len(df_novo)>0:
                    df_novo = adicionar_codigo_ucs(df_novo)
                    df_novo = add_leadset_wire(df_novo)
                    df_novo = adicionar_processos(df_novo)
                

                colunas_prioritarias = [
                    "Nome do arquivo",
                    'Projeto',
                    "Fase",
                    "ProcessoA",
                    "ProcessoB",
                    "UCS",
                    "Leadset",
                    'Sub-Assembly'
                ]
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
            except Exception as err:
                pass
                #logger.error(f"Erro reordenar colunas no {name}: {err}")


            try:
                colunas_para_excluir = [
                    "Wire Spec",
                    "Pack",
                    "Options",
                    "Cav. 1",
                    "Cav. 2",
                    "Term.Mat 1",
                    "Term.Mat 2",
                ]
                df_novo = excluir_colunas(df_novo, colunas_para_excluir)
            except Exception as err:
                pass
                #logger.error(f"Erro excluir colunas no {name}: {err}")

        if name == "MultWire_Nb":
            try:
                if len(df_novo)>0:
                    df_novo = adicionar_codigo_ucs(df_novo)
                colunas_prioritarias = [
                    "Nome do arquivo",
                    'Projeto',
                    "Fase",
                    "UCS",
                    "Mult.Wire Nb",
                    "T",
                    "CSA",
                    "Length",
                    "Wire Spec",
                    "Pack",
                    "CutBack1",
                    "CutBack2",
                    "W1",
                    "W2",
                    "W3",
                    "W4",
                    "Int.PN",
                ]
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
            except Exception as err:
                pass
                #logger.error(f"Erro reordenar colunas no {name}: {err}")


        if name == "Sleeve_Nb":
            try:
                if len(df_novo)>0:
                    df_novo = adicionar_codigo_ucs(df_novo)
                colunas_prioritarias = ["Nome do arquivo", "Fase", "UCS",'Projeto']
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
            except Exception as err:
                pass
                #logger.error(f"Erro reordenar colunas no {name}: {err}")

        if name == "Splice_Nb":
            try:
                df_novo = adicionar_codigo_ucs(df_novo)

                colunas = [
                    "L1",
                    "L2",
                    "L3",
                    "L4",
                    "L5",
                    "L6",
                    "L7",
                    "L8",
                    "L9",
                    "L10",
                    "R1",
                    "R2",
                    "R3",
                    "R4",
                    "R5",
                    "R6",
                    "R7",
                    "R8",
                    "R9",
                    "R10",
                ]

                for c in colunas:
                    if c not in df_novo.columns:
                        df_novo[c] = None

                colunas_prioritarias = [
                    "Nome do arquivo",
                    'Projeto',
                    "Fase",
                    "UCS",
                    "Splice Nb",
                    "Int.PN",
                    "Extra Component PN",
                    "Pack",
                    "CSA Left",
                    "CSA Right",
                    "CSA Total",
                    "Node",
                    "L1",
                    "L2",
                    "L3",
                    "L4",
                    "L5",
                    "L6",
                    "L7",
                    "L8",
                    "L9",
                    "L10",
                    "R1",
                    "R2",
                    "R3",
                    "R4",
                    "R5",
                    "R6",
                    "R7",
                    "R8",
                    "R9",
                    "R10",
                ]
                df_novo = reordenar_colunas(df_novo, colunas_prioritarias)
                
            except Exception as err:
                pass
                #logger.error(f"Erro reordenar colunas no {name}: {err}")

        # if name.upper() == "MULTWIRE_NB" or ('Mult' in name and 'Wire' in name):
        #     colunas_prioritarias =['Mult.Wire Nb', 'T', 'CSA', 'Length', 'Wire Spec', 'Pack', 'CutBack1','CutBack2', 'W1', 'W2','Int.PN']
        #     df_novo = reordenar_colunas(df_novo, colunas_prioritarias)

        dict_df[name] = df_novo
    
    if len(dict_df['Splice_Nb'])>0:
        dict_df['Splice_Nb'] = adicionar_bitolas_em_splices(dict_df['Wire_Nb'],dict_df['Splice_Nb'])
    
    dict_df["Multcrimp_Nb"] = criar_multicrimp(dict_df['Wire_Nb'])
    dict_df["Multcrimp_Nb - Estudo"] = adicionar_mult_estudo(dict_df['Wire_Nb'])
    dict_df["Multcrimp_Nb - Estudo"] = ajustar_multicrimp_estudo(dict_df["Multcrimp_Nb - Estudo"],dict_df["Multcrimp_Nb"])
    return dict_df

@log_errors_record
def consolidar_cut(path_destino:str= None):
    
    nome_pasta = "PastaTemp"
    destino_root = os.path.join(tempfile.gettempdir(), nome_pasta)
    
    arquivos_excel = list(encontrar_arquivos_excel(destino_root))

    dict_consolidado = {
        "Wire_Nb": pd.DataFrame(), 
        "MultWire_Nb": pd.DataFrame(),
        "Sleeve_Nb": pd.DataFrame(), 
        "Splice_Nb": pd.DataFrame(), 
        "Multcrimp_Nb - Estudo": pd.DataFrame(),
        "Multcrimp_Nb": pd.DataFrame()
    }

    for arquivo in arquivos_excel:
        tipo = identificar_cut_sheets(arquivo)

        if tipo == 1:
            #print(f'Tipo: {tipo} | Cut´s')
            dict_temp = format_cut(arquivo)

            for chave in dict_consolidado:
                df_novo = dict_temp.get(chave)
                df_novo = df_novo.dropna(axis=1, how="all")
                if not isinstance(df_novo, pd.DataFrame) or df_novo.empty:
                    continue
                
                if dict_consolidado[chave].empty:
                    # primeiro DataFrame → atribuição direta
                    dict_consolidado[chave] = df_novo.copy()
                else:
                    dict_consolidado[chave] = pd.concat([dict_consolidado[chave], df_novo],ignore_index=True)
        else:
            print(f'Tipo: {tipo} | CHARTED-CUT')

    dict_consolidado["Splice_Nb"] = arrumar_colunas(dict_consolidado["Splice_Nb"], coluna_alvo="Splice_Nb")
    dict_consolidado["Multcrimp_Nb - Estudo"] = arrumar_colunas(dict_consolidado["Multcrimp_Nb - Estudo"], coluna_alvo="Multcrimp_Nb - Estudo")
    dict_consolidado["MultWire_Nb"] = arrumar_colunas(dict_consolidado["MultWire_Nb"], coluna_alvo="MultWire_Nb")
    

    hoje = datetime.now().strftime("%d-%m-%Y")
    
    nome_arquivo = f"Cuts_Consolidadas_{hoje}.xlsx"

    caminho_arquivo = os.path.join(path_destino, nome_arquivo)

    salvar_dict_df_em_excel(dict_consolidado, caminho_arquivo=caminho_arquivo)
    
    nome_pasta = "PastaTemp"
    destino_root = os.path.join(tempfile.gettempdir(), nome_pasta)

    # Se existir, apaga tudo
    if os.path.exists(destino_root):
        shutil.rmtree(destino_root)

if __name__ == "__main__":

    app = QApplication(sys.argv)

    janela = MasterDataWindow()

    janela.show()

    sys.exit(app.exec())