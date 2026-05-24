from art import *
from datetime import datetime
from tabulate import tabulate
import time
import re
import sys
import math
import os
import shutil
import zipfile
import pandas as pd
from openpyxl.utils import get_column_letter
from collections import defaultdict
from tqdm import tqdm
from PySide6.QtWidgets import QMessageBox
import logging
import functools
from pathlib import Path
from PySide6.QtWidgets import QMessageBox




import traceback
from functools import wraps

def logo(texto):
    # Texto em ASCII art
    titulo = text2art(texto, font="small")

    # Rodapé customizado
    rodape = "\n" + "-" * 80

    # Exibir no terminal
    print(titulo + rodape)


def excluir_colunas(df, colunas_para_excluir):
    # Filtra apenas as colunas que realmente existem no DataFrame
    colunas_existentes = [col for col in colunas_para_excluir if col in df.columns]

    # Retorna o DataFrame com as colunas removidas
    return df.drop(columns=colunas_existentes)


def reordenar_colunas(df, colunas_prioritarias):
    # Garante que as colunas prioritárias existem no DataFrame
    colunas_prioritarias = [col for col in colunas_prioritarias if col in df.columns]

    # Pega as colunas restantes (que não estão nas prioritárias)
    colunas_restantes = [col for col in df.columns if col not in colunas_prioritarias]

    # Reordena o DataFrame
    return df[colunas_prioritarias + colunas_restantes]




def salvar_dict_df_em_excel(dfs_dict, caminho_arquivo):
    # Se não houver dados ou todos os DataFrames forem vazios, cria uma aba padrão
    if not dfs_dict or all(df.empty for df in dfs_dict.values()):
        dfs_dict = {"SemDados": pd.DataFrame({"Mensagem": ["Nenhum dado disponível"]})}

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

def consultar_arquivos_base(id_name: str, show_display: bool = False) -> str:
    # Lê a planilha de configuração
    config_path = "C:/Projetos/Tool/files/config.xlsx"
    try:
        df = pd.read_excel(config_path)
    except Exception as e:
        return f"Erro ao ler o arquivo de configuração: {e}"

    if show_display:
        table = tabulate(df, headers="keys", tablefmt="grid")
        print(table)

    # Filtra as linhas com o id_name desejado
    df_filtrado = df[df["Id name"] == id_name]

    # Verifica se encontrou alguma linha correspondente
    if df_filtrado.empty:
        return f"ID '{id_name}' não encontrado no arquivo de configuração."

    # Recupera os caminhos da primeira ocorrência
    path_network = df_filtrado.iloc[0].get("Network [Path]")
    path_local = df_filtrado.iloc[0].get("Local [Path]")

    if pd.notna(path_network) and os.path.exists(path_network):
        return path_network
    elif pd.notna(path_local) and os.path.exists(path_local):
        return path_local
    else:
        return f"O ID {id_name} não possui caminho válido."

def adicionar_codigo_ucs(dados: pd.DataFrame=None, coluna_tag:str=None) -> pd.DataFrame:
    # Lê o arquivo de UCS
    df_ucs = pd.read_excel(consultar_arquivos_base("codigo_ucs"))

    # Cria o dicionário para mapear "Familia Externa" -> "Fam Int"
    dict_ucs = {
        str(row["Familia Externa"]).replace('.',''): row["Fam Int"] for _, row in df_ucs.iterrows()
    }

    # Verifica se a coluna necessária existe
    if "Nome do arquivo" not in dados.columns:
        raise KeyError("A coluna 'Nome do arquivo' não existe no DataFrame fornecido.")

    # Função que verifica se alguma chave do dicionário está contida no nome do arquivo
    def encontrar_ucs(nome_arquivo: str):
        for chave, valor in dict_ucs.items():
            if pd.notna(chave) and str(chave) in str(nome_arquivo):
                return valor
        return "Verificar"  # Retorno padrão se nenhuma chave for encontrada

    # Aplica a função para cada linha

    if coluna_tag and coluna_tag in dados.columns:
        dados["UCS"] = (
            dados[coluna_tag]
            .astype(str)
            .str.replace(" ", "")   # remove TODOS os espaços
            .apply(encontrar_ucs)
        )
    else:
        dados["UCS"] = dados["Nome do arquivo"].apply(encontrar_ucs)

    return dados

def add_leadset_wire(df):
    df_completo = pd.DataFrame()

    df_gm = df[df["UCS"].str.contains("G|S", case=False, na=False)]
    df_VS30 = df[df["UCS"].str.contains("V", case=False, na=False)]
    df_U11 = df[df["UCS"].str.contains("B", case=False, na=False)]
    df_XFD = df[df["UCS"].str.contains("X", case=False, na=False)]

    if len(df_gm) > 0:
        if "Multicore" not in df_gm.columns:
            df_gm["Multicore"] = None

        df_completo_GM = pd.DataFrame()
        df_gm_diretos = df_gm[df_gm["Multicore"].isna()].reset_index(drop=True)
        df_gm_tw_menor_1_5 = df_gm[
            (df_gm["CSA"].astype(float) < 1.5) & (df_gm["Multicore"].notna())
        ].reset_index(drop=True)
        df_gm_tw_maior_1_5 = df_gm[
            (df_gm["CSA"].astype(float) >= 1.5) & (df_gm["Multicore"].notna())
        ].reset_index(drop=True)

        df_gm_diretos["Leadset"] = df_gm_diretos["UCS"].astype(str) + df_gm_diretos[
            "Wire Nb"
        ].astype(str)
        df_gm_tw_menor_1_5["Leadset"] = df_gm_tw_menor_1_5["UCS"].astype(
            str
        ) + df_gm_tw_menor_1_5["Multicore"].astype(str)
        df_gm_tw_maior_1_5["Leadset"] = df_gm_tw_maior_1_5["UCS"].astype(
            str
        ) + df_gm_tw_maior_1_5["Wire Nb"].astype(str)

        df_completo_GM = pd.concat([df_completo_GM, df_gm_diretos], ignore_index=False)
        df_completo_GM = pd.concat(
            [df_completo_GM, df_gm_tw_menor_1_5], ignore_index=False
        )
        df_completo_GM = pd.concat(
            [df_completo_GM, df_gm_tw_maior_1_5], ignore_index=False
        )
    else:
        df_completo_GM = pd.DataFrame()
        # print(f'Dataframe: GEM/ SPIN esta Vazio')

    if len(df_VS30) > 0:
        df_completo_VS30 = pd.DataFrame()

        df_VS30_diretos = df_VS30[df_VS30["Multicore"].isna()].reset_index(drop=True)
        df_VS30_TW = df_VS30[
            (df_VS30["Multicore"].notna())
            & (~df_VS30["Multicore"].str.contains("MC", na=False))
        ].reset_index(drop=True)
        df_VS30_MC = df_VS30[
            (df_VS30["Multicore"].notna())
            & (df_VS30["Multicore"].str.contains("MC", na=False))
        ].reset_index(drop=True)

        df_VS30_diretos["Leadset"] = df_VS30_diretos["UCS"].astype(
            str
        ) + df_VS30_diretos["Wire Nb"].astype(str)
        df_VS30_TW["Leadset"] = df_VS30_TW["UCS"].astype(str) + df_VS30_TW[
            "Wire Nb"
        ].astype(str)
        df_VS30_MC["Leadset"] = df_VS30_MC["UCS"].astype(str) + df_VS30_MC[
            "Multicore"
        ].astype(str)

        df_completo_VS30 = pd.concat(
            [df_completo_VS30, df_VS30_diretos], ignore_index=False
        )
        df_completo_VS30 = pd.concat([df_completo_VS30, df_VS30_TW], ignore_index=False)
        df_completo_VS30 = pd.concat([df_completo_VS30, df_VS30_MC], ignore_index=False)

    else:
        df_completo_VS30 = pd.DataFrame()
        # print(f'Dataframe: VS30 esta Vazio')

    if len(df_U11) > 0:
        df_completo_U11 = pd.DataFrame()

        df_U11_diretos = df_U11[df_U11["Multicore"].isna()].reset_index(drop=True)
        df_U11_TW = df_U11[df_U11["Multicore"].notna()].reset_index(drop=True)

        df_U11_TW["Total_duplicates"] = df_U11_TW["Multicore"].map(
            df_U11_TW["Multicore"].value_counts()
        )

        df_U11_MC = (
            df_U11_TW[df_U11_TW["Total_duplicates"] >= 3]
            .drop(columns=["Total_duplicates"])
            .reset_index(drop=True)
        )
        df_U11_TW = (
            df_U11_TW[df_U11_TW["Total_duplicates"] < 3]
            .drop(columns=["Total_duplicates"])
            .reset_index(drop=True)
        )

        df_U11_diretos["Leadset"] = df_U11_diretos["UCS"].astype(str) + df_U11_diretos[
            "Wire Nb"
        ].astype(str)
        df_U11_MC["Leadset"] = df_U11_MC["UCS"].astype(str) + df_U11_MC[
            "Multicore"
        ].astype(str)
        df_U11_TW["Leadset"] = df_U11_TW["UCS"].astype(str) + df_U11_TW[
            "Wire Nb"
        ].astype(str)

        df_completo_U11 = pd.concat(
            [df_completo_U11, df_U11_diretos], ignore_index=False
        )
        df_completo_U11 = pd.concat([df_completo_U11, df_U11_MC], ignore_index=False)
        df_completo_U11 = pd.concat([df_completo_U11, df_U11_TW], ignore_index=False)

    else:
        df_completo_U11 = pd.DataFrame()
        # print(f'Dataframe: U11 esta Vazio')

    if len(df_XFD) > 0:
        pd.DataFrame()
    else:
        pd.DataFrame()
        # print(f'Dataframe: XFD esta Vazio')

    df_completo = pd.concat([df_completo, df_completo_GM], ignore_index=False)
    df_completo = pd.concat([df_completo, df_completo_VS30], ignore_index=False)
    df_completo = pd.concat([df_completo, df_completo_U11], ignore_index=False)
    df_completo = pd.concat([df_completo, df_XFD], ignore_index=False)

    return df_completo

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
    for m in tqdm(mcs, desc="[+] Gerando Multicrimp", colour="blue"):
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

            df_prov3["Joint"] = df_prov3["Joint 1"].fillna(df_prov3["Joint 2"])
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

    return df_new

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

    list_completa["Note 1"] = None
    mcs = pd.concat([list_completa["Note 1"], list_completa["Note 2"]])
    mcs = mcs.dropna().astype(str).unique().tolist()
    mcs = [v for v in mcs if v.startswith("MU")]

    mcs = sorted(mcs, key=sort_key)

    lista_combinacoes = []

    for c in list_completa.columns:
        for j in list_completa.index:
            if (
                str(list_completa.loc[j, c])[:1] != "S"
                and str(list_completa.loc[j, c])[:1] != "M"
            ):
                list_completa.loc[j, c] = None

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

    for m in tqdm(lista_combinacoes, desc="[+] Gerando Multicrimp", colour="blue"):
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
            for a in lista_splices:
                df_prov[texto + str(contador)] = a
                colunas_base = colunas_base + [texto + str(contador)]
                colunas_sem_pn = colunas_sem_pn + [texto + str(contador)]
                contador += 1

            # Criar circuitos splices
            contador = 1
            texto = "MC.W"
            for a in lista_de_circuitos:
                df_prov[texto + str(contador)] = a
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
                        df_multicrimp = pd.concat(
                            [df_multicrimp, df_linha], ignore_index=True
                        )
                except:
                    df_multicrimp = pd.concat(
                        [df_multicrimp, df_linha], ignore_index=True
                    )
            else:
                df_multicrimp = pd.concat([df_multicrimp, df_linha], ignore_index=True)

            number_derivativo += 1

    colunas_ord = [
        "Nome do arquivo",
        "Fase",
        "UCS",
        "Term. 1",
        "Term. 2",
        "Note 2",
        "MC.Splice1",
        "MC.Splice2",
        "MC.Splice3",
        "MC.Splice5",
        "MC.Splice5",
        "MC.Splice6",
        "MC.Splice7",
        "MC.Splice8",
        "MC.Splice9",
        "MC.Splice10",
        "MC.W1",
        "MC.W2",
        "MC.W3",
        "MC.W4",
        "MC.W5",
        "MC.W6",
        "MC.W7",
        "MC.W8",
        "MC.W9",
        "MC.W10",
        "MC.W11",
        "MC.W12",
        "MC.W13",
        "MC.W14",
        "MC.W15",
        "MC.W16",
        "MC.W17",
        "MC.W18",
        "MC.W19",
        "MC.W20",
    ]

    df_multicrimp = reordenar_colunas(df_multicrimp, colunas_prioritarias=colunas_ord)

    df_multicrimp = df_multicrimp.drop(columns=["Term. 1"], errors="ignore").rename(
        columns={"Term. 2": "MC.Terminal"}
    )

    return df_multicrimp

def encontrar_arquivos_zip(diretorio):
    arquivos_zip = []  # Lista para armazenar os caminhos dos arquivos .zip

    # Caminha por todas as pastas e subpastas no diretório
    for root, dirs, files in os.walk(diretorio):
        for file in files:
            # Verifica se o arquivo tem a extensão .zip
            if file.endswith(".zip"):
                # Adiciona o caminho completo do arquivo à lista
                arquivos_zip.append(os.path.join(root, file))

    return arquivos_zip

def resource_path(relative_path):
    """Retorna o caminho absoluto do recurso, compatível com PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        # Quando empacotado pelo PyInstaller
        base_path = sys._MEIPASS
    else:
        # Quando rodando direto do código fonte
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def extrair_zip(caminhos_zip):
    """
    Extrai arquivos .zip para o mesmo diretório onde cada .zip está localizado.

    :param caminhos_zip: Lista com os caminhos dos arquivos .zip
    """
    for caminho_zip in caminhos_zip:
        if not os.path.isfile(caminho_zip):
            print(f"Arquivo não encontrado: {caminho_zip}")
            continue

        # Define o diretório onde o zip está localizado
        diretorio_destino = os.path.dirname(caminho_zip)

        try:
            with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                zip_ref.extractall(diretorio_destino)
                print(
                    f"Arquivo {caminho_zip} extraído com sucesso para {diretorio_destino}"
                )
        except zipfile.BadZipFile:
            print(f"Arquivo corrompido ou inválido: {caminho_zip}")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar extrair {caminho_zip}: {e}")

def banner(msg: str, align: str = "center", char: str = "-", width: int = None):
    """
    Gera uma linha de banner formatada.

    :param msg: Mensagem a ser exibida.
    :param align: Alinhamento ('left', 'center', 'right').
    :param char: Caractere de preenchimento.
    :param width: Largura total (se None, usa largura do terminal).
    :return: String formatada.
    """
    if width is None:
        width = shutil.get_terminal_size(
            (80, 20)
        ).columns  # largura do terminal (padrão=80)

    # Garante que msg tenha espaços antes/depois
    text = f" {msg} "

    if align == "left":
        return text.ljust(width, char)
    elif align == "right":
        return text.rjust(width, char)
    else:  # default = center
        return text.center(width, char)

def encontrar_arquivos_excel(diretorio):
    arquivos_excel = []
    for raiz, pastas, arquivos in os.walk(diretorio):
        for arquivo in arquivos:
            if arquivo.lower().endswith((".xlsx", ".xls")):
                caminho_completo = os.path.join(raiz, arquivo)
                arquivos_excel.append(caminho_completo)
    return arquivos_excel

def listar_abas_excel(caminho_arquivo):
    with open(caminho_arquivo, "rb") as f:
        xls = pd.ExcelFile(f)
        return xls.sheet_names

def procurar_indices_linhas(palavras, df):
    indices = df[df[df.columns[0]].isin(palavras)].index.tolist()
    indices.extend([df.shape[0]])
    return indices

def renomear_colunas_duplicadas(df):
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
    return df

def adicionar_processos(dados: pd.DataFrame) -> pd.DataFrame:
    # Lê o arquivo de UCS
    df_processo = pd.read_excel(consultar_arquivos_base("circuitos_especiais"))

    dict_processo = {row["Código"]: "PRENSA" for i, row in df_processo.iterrows()}

    dados["ProcessoA"] = None
    dados["ProcessoB"] = None

    if "Joint 1" not in dados.columns:
        dados["Joint 1"] = None
    if "Joint 2" not in dados.columns:
        dados["Joint 2"] = None

    if "Term. 1" not in dados.columns:
        dados["Term. 1"] = None
    if "Term. 1" not in dados.columns:
        dados["Term. 1"] = None

    for i in dados.index:
        if "SP" in str(dados.at[i, "Node 1"]) or "SP" in str(dados.at[i, "Joint 1"]):
            dados.at[i, "ProcessoA"] = "SPLICE"
        if "SP" in str(dados.at[i, "Node 2"]) or "SP" in str(dados.at[i, "Joint 2"]):
            dados.at[i, "ProcessoB"] = "SPLICE"

        if pd.notna(dados.at[i, "Term. 1"]) and pd.isna(dados.at[i, "ProcessoA"]):
            dados.at[i, "ProcessoA"] = dict_processo.get(
                dados.at[i, "Term. 1"], "CORTE"
            )
        if pd.notna(dados.at[i, "Term. 2"]) and pd.isna(dados.at[i, "ProcessoB"]):
            dados.at[i, "ProcessoB"] = dict_processo.get(
                dados.at[i, "Term. 2"], "CORTE"
            )
        if pd.isna(dados.at[i, "Term. 1"]) and pd.isna(dados.at[i, "ProcessoA"]):
            dados.at[i, "ProcessoA"] = "CORTE"
        if pd.isna(dados.at[i, "Term. 2"]) and pd.isna(dados.at[i, "ProcessoB"]):
            dados.at[i, "ProcessoB"] = "CORTE"

    return dados

def add_BOM(df_base: pd.DataFrame):
    with open(consultar_arquivos_base("terminais_sem"), "rb") as f:
        df_term = pd.read_excel(f)

    dict_Delivery = {
        row["Part Number"]: row["Feed Type/Delivery Form"]
        for i, row in df_term.iterrows()
    }
    dict_Technology = {
        row["Part Number"]: row["Connection Technology"]
        for i, row in df_term.iterrows()
    }

    df_base["Connection Technology_A"] = df_base["TERM_A"].map(dict_Technology)
    df_base["Connection Technology_B"] = df_base["TERM_B"].map(dict_Technology)

    df_base["Feed Type/Delivery Form_A"] = df_base["TERM_A"].map(dict_Delivery)
    df_base["Feed Type/Delivery Form_B"] = df_base["TERM_B"].map(dict_Delivery)

    return df_base

def add_ALOC(df_base: pd.DataFrame):
    with open(consultar_arquivos_base("mapa_aloc_corte"), "rb") as f:
        df_aloc = pd.read_excel(f)

    dict_aloc = {}

    for i in df_aloc.index:
        lista_ckt = ["G15", "G16", "G17", "G26"]
        leadset = str(df_aloc.loc[i, "FAMÍLIA_CIRCUITO"])
        alocac = str(df_aloc.loc[i, "ALOCAÇÃO NOVA"])

        if any(item in leadset for item in lista_ckt):
            if leadset[3] != "_":
                leadset = leadset.replace(leadset[:3], leadset[:3] + "_")

        dict_aloc[leadset] = alocac
    df_base["Alocações"] = df_base["Leadset"].map(dict_aloc)

    return df_base

def salvar_arquivo(
    df: pd.DataFrame, caminho_pasta: str, nome_arquivo: str, formato: str = "csv"
):
    """
    Salva um DataFrame em um arquivo com data e hora no nome.

    Parâmetros:
    - df (pd.DataFrame): DataFrame a ser salvo.
    - caminho_pasta (str): Caminho da pasta onde o arquivo será salvo.
    - nome_arquivo (str): Nome base do arquivo (sem extensão).
    - formato (str): Formato do arquivo ('csv' ou 'excel'). Padrão: 'csv'.

    Retorno:
    - str: Caminho completo do arquivo salvo.
    """
    # Garantir que a pasta existe
    os.makedirs(caminho_pasta, exist_ok=True)

    # Gerar timestamp no formato dd-mm-aaaa_HHMM
    timestamp = datetime.now().strftime("%d-%m-%Y_%H%M")

    # Nome final do arquivo
    nome_final = f"{nome_arquivo}_{timestamp}"

    # Definir caminho completo
    if formato == "csv":
        caminho_completo = os.path.join(caminho_pasta, f"{nome_final}.csv")
        df.to_csv(caminho_completo, index=False,sep=';')
    elif formato == "excel":
        caminho_completo = os.path.join(caminho_pasta, f"{nome_final}.xlsx")
        df.to_excel(caminho_completo, index=False)
    else:
        raise ValueError("Formato inválido. Use 'csv' ou 'excel'.")

    return caminho_completo


def limpar_colunas_redundantes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove colunas de um DataFrame que:
    - Têm todos os valores NaN (dropna)
    - Têm todos os valores iguais a itens indesejados (ex: "-", "0", "X", etc.)
    - Contêm apenas espaços em branco ou strings vazias

    Parâmetros:
        df (pd.DataFrame): DataFrame de entrada

    Retorna:
        pd.DataFrame: DataFrame limpo
    """
    df_limpo = df.copy()

    # 1. Remover colunas inteiras com todos os valores NaN
    df_limpo.dropna(how="all", axis=1, inplace=True)

    # 2. Remover colunas com apenas um valor indesejado
    valores_a_remover = ["-", "0", "0.0", 0, "X", "false", "true", " "]
    for valor in valores_a_remover:
        df_limpo = df_limpo.loc[:, ~(df_limpo.eq(valor).all())]

    # 3. Remover colunas com apenas None ou NaN (redundante com dropna, mas garante)
    df_limpo = df_limpo.loc[:, ~df_limpo.isna().all()]

    # 4. Remover colunas com apenas espaços em branco ou strings vazias
    for col in df_limpo.columns:
        if df_limpo[col].apply(lambda x: isinstance(x, str) and x.strip() == "").all():
            df_limpo.drop(columns=col, inplace=True)

    return df_limpo


def is_empty(value):
    # None ou ""
    if value is None or value == "":
        return True
    
    # NaN
    if isinstance(value, float) and math.isnan(value):
        return True
    
    # listas, dicts ou sets vazios
    if isinstance(value, (list, dict, set)) and len(value) == 0:
        return True
    
    return False


def limpar_dict(d):
    return {k: v for k, v in d.items() if not is_empty(v)}

def convert_legacy(dados):
    dict_tipos = {
        # 'ISXA': '65762',
        # 'ITXA': '65763',
        # 'IUXA': '65764',
        # 'ISXB': '65768',
        # 'ITXB': '65769',
        # 'IUXB': '65770',
        # 'ISXC': '65864',
        # 'ITXC': '64750',
        # 'ITSA': '65791',
        # 'ITXA8': '65988',
        # 'IXEA':'65826',  # verificar
        # 'ITXS':'65763',  # verificar
        # 'STXA': '65860',  # verificar
        "FLRYA": "65763",
        "FLRY-A": "65763",
        "FRLYB": "65769",
        "FRLY-B": "65769",
        "FLRY-B": "65769",  # verificar
        "FLR2X-A": "65927",
        "FLR21X-A": "65860",
        "FLR2X-B": "65784",
        "FLR91X-B": "65855",
        "FLR9Y-A": "65778",
        "FLR91X-A": "65856",
        "CU-R2PVC-A": "65763",
        "CU-R2PVC-B": "65769",
        "CU-R3XLPE-A": "65766",
        "CU-R3XLPE-B": "65784",
        "CU-R4XLPE-B": "65855",
        "CU-R4XLPO-B": "65855",
        "CU-R2PP-E1": "65778",
        "CU-R4XLPE-A": "65856",
    }

    def map_wire_code(wire_code):
        if isinstance(wire_code, str):  # Verifica se wire_code é string
            for key in dict_tipos:
                if key in wire_code:
                    return dict_tipos[key]
        return None

    # 1ª tentativa: pela coluna "Wire Code"
    dados["Legacy"] = dados["Wire Code"].apply(map_wire_code)

    # 2ª tentativa: somente onde Legacy está None -> usar "Part Description"
    mask = dados["Legacy"].isna()
    dados.loc[mask, "Legacy"] = dados.loc[mask, "Part Description"].apply(map_wire_code)

    return dados

def converter_cores(dados):
    color_codes = {
        "Red": "RD",
        "Black": "BK",
        "Yellow": "YE",
        "White": "WH",
        "Violet": "V",
        "Brown": "BR",
        "Blue": "BL",
        "Pink": "PK",
        "Orange": "OG",
        "Green": "GN",
        "Dark Green": "DG",
        "Grey": "GY",
        "Dark Blue": "DB",
        "Tan/Beige": "BG",
        "Sky Blue": "SB",
        "Light Green": "LG",
        "Lavender": "LV",
        "Purple": "PU",
        "Tan": "TN",
        "Turquoise": "TQ",
        "Light Blue": "LB",
    }
    dados["Cor 1"] = dados["Primary Color"].map(color_codes)
    dados["Cor 2"] = dados["Secondary Color"].map(color_codes)

    return dados


def corrigir_valor(val):
    dict_correcao = {
        "3,20E+08": "3202980E2",
        "3,20E+09": "3202518E3",
        "3,20E+14": "3202570E8",
    }
    # Se for string, tenta substituir direto
    if isinstance(val, str) and val in dict_correcao:
        return dict_correcao[val]

    # Se for float ou int, tenta formatar como string com vírgula
    elif isinstance(val, (float, int)):
        val_str = f"{val:.2E}".replace(".", ",")
        if val_str in dict_correcao:
            return dict_correcao[val_str]

    # Retorna o original se não encontrou
    return val


class LoggerTerminal:
    CORES = {
        "padrao": "\033[0m",
        "vermelho": "\033[31m",
        "verde": "\033[32m",
        "amarelo": "\033[33m",
        "ciano": "\033[36m",
        "magenta": "\033[35m",
        "branco": "\033[37m",
    }

    def __init__(
        self, salvar_log: bool = False, log_path: str = "log.txt", typing: bool = False
    ):
        self.salvar_log = salvar_log
        self.log_path = log_path
        self.typing = typing

        if salvar_log and not os.path.exists(log_path):
            with open(log_path, "w") as f:
                f.write("==== LOG INICIAL ====\n")

    def _log(self, simbolo: str, cor: str, mensagem: str):
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cor_ansi = self.CORES.get(cor, self.CORES["padrao"])
        reset = self.CORES["padrao"]
        texto_formatado = f"[{simbolo}] {agora} {mensagem}"
        saida_terminal = f"{cor_ansi}{texto_formatado}{reset}"

        # Exibe no terminal
        if self.typing:
            for c in saida_terminal:
                print(c, end="", flush=True)
                time.sleep(0.01)
            print()
        else:
            print(saida_terminal)

        # Salva em arquivo
        if self.salvar_log:
            with open(self.log_path, "a") as f:
                f.write(texto_formatado + "\n")

    def sucesso(self, mensagem: str):
        self._log("+", "verde", mensagem)

    def error(self, mensagem: str):
        self._log("x", "vermelho", mensagem)

    def atencao(self, mensagem: str):
        self._log("!", "amarelo", mensagem)

    def duvida(self, mensagem: str):
        self._log("?", "ciano", mensagem)

    def verifique(self, mensagem: str):
        self._log("-", "magenta", mensagem)

    def info(self, mensagem: str):
        self._log("*", "branco", mensagem)


# logger = LoggerTerminal(salvar_log=False, typing=True)



# --- DEFINIÇÃO DINÂMICA DE CAMINHOS ---

if getattr(sys, 'frozen', False):
    # Se o app for um executável (.exe), BASE_DIR é a pasta onde o .exe está
    # Usamos sys.executable para garantir que o log fique na pasta do programa e não na temp
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # Se for script .py, sobe dois níveis para definir a raiz do projeto
    BASE_DIR = Path(__file__).resolve().parent.parent

# Define a pasta de logs: C:\...\Delta_analytics\logs
LOG_DIR = BASE_DIR / "logs"

# Cria a pasta de logs automaticamente se ela não existir
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Define o nome do arquivo de log com a data atual
LOG_FILE = LOG_DIR / f"debug_{datetime.now().strftime('%Y-%m-%d')}.log"

# --- CONFIGURAÇÃO DO LOGGING ---

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.ERROR,
    format='%(asctime)s | %(levelname)s | %(name)s\n%(message)s\n' + '-'*50,
    encoding='utf-8'
)

# --- DECORADOR DE ERROS ---

def log_errors_record(func):
    """
    Decorador que captura erros, identifica o arquivo e a linha,
    salva no log e avisa o usuário sem travar o programa.
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Captura o rastro completo do erro
            tb = traceback.extract_tb(sys.exc_info()[2])
            file_name, line_number, func_name, text = tb[-1]
            
            error_details = traceback.format_exc()
            
            # Mensagem formatada para o arquivo de log
            msg_log = (
                f"FUNÇÃO: {func.__name__}\n"
                f"ARQUIVO: {file_name}\n"
                f"LINHA: {line_number}\n"
                f"ERRO: {str(e)}\n"
                f"TRACEBACK COMPLETO:\n{error_details}"
            )
            
            logging.error(msg_log)
            
            # Alerta visual para o usuário
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("Erro de Execução")
            msg_box.setText(f"Erro na função: {func.__name__}")
            msg_box.setInformativeText(f"Linha: {line_number} | Arquivo: {os.path.basename(file_name)}")
            msg_box.setDetailedText(error_details)
            msg_box.exec()
            
            return None
    return wrapper