import os
from datetime import datetime
import pandas as pd
import sys
import traceback
import warnings

warnings.simplefilter("ignore")

root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)


from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QProgressDialog,
    QApplication, QStyle
)

from PySide6.QtGui import QGuiApplication, QCursor

from pathlib import Path
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
	
from utils.function import *
from utils.funcoes import *

@log_errors_record
def converter_nan(x):
    if pd.isna(x):  # cobre np.nan, None, pd.NaT, etc.
        return 0
    if isinstance(x, str) and x.strip().lower() in {"", "nan", "none", "null"}:
        return 0
    return str(x)


@log_errors_record
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


@log_errors_record
def Comp_SAP_vs_CAO(
    path_sap,
    path_cao,
    retirar_undeline_sap=False,
    retirar_undeline_cao=False,
    atualizar_dimensional=True,
):
    
    try:
        df_comp_completo = pd.DataFrame()
        ckt_add_completo = pd.DataFrame()
        new_df_cao_completo = pd.DataFrame()

        with open(path_sap, "rb") as f:
            df1_base = pd.read_excel(f,dtype=str)

        with open(path_cao, "rb") as f:
            
            df2_base = pd.read_csv(f, sep=";")
            df2_base = df2_base.applymap(corrigir_valor)

        df1_base = df1_base.replace("D_MULTICR", None)
        df2_base = df2_base.replace("D_MULTICRIMP", None)

        leadset_check = df1_base["Internal Family"].unique().tolist()

        for led in leadset_check:
            df1 = df1_base[
                (df1_base["Internal Family"] == led)
                & (pd.to_numeric(df1_base["SECTIONN"], errors="coerce") >= 0.35)
            ].reset_index(drop=True)

            df2 = df2_base[
                (df2_base["Leadset"].str.startswith(led, na=False))
                & (df2_base["ProdVersion"] == 0)
                & (~df2_base["Description"].str.contains("OBSOLETO", na=False))
            ].reset_index(drop=True)

            df1 = df1.sort_values(by=["WIRE_TUBE_SPLICE"]).reset_index(drop=True)
            df2 = df2.reset_index(drop=True).sort_values(by=["Wire1Key"])

            lista_ckt_diretos_add = []
            lista_ckt_tw_add = []

            df_comp = pd.DataFrame()

            linha_controle = 0

            leadset_SAP = df1["Leadset"].dropna().unique().tolist()

            for ckt in leadset_SAP:
                # for ckt in tqdm(leadset_SAP, desc='Processando 1 [{led} ]',colour='blue'):

                ckt1 = ckt

                if retirar_undeline_sap:
                    ckt1 = f"{ckt[:3]}{ckt[4:]}"

                if retirar_undeline_cao and ckt[3] != "_":
                    ckt1 = f"{ckt[:3]}_{ckt[3:]}"

                df_prov_sap = df1[df1["Leadset"] == ckt].reset_index(drop=True)
                df_prov_cao = df2[df2["Leadset"] == ckt1].reset_index(drop=True)

                # CABOS DIRETOS
                if len(df_prov_sap) == 1 and len(df_prov_cao) == 1:
                    leadset_sap = df_prov_sap.loc[0, "Leadset"]
                    cabo_sap = converter_nan(df_prov_sap.loc[0, "WIRE_TUBE_SPLICE"])
                    comp_sap = int(df_prov_sap.loc[0, "LENGTH"])
                    termA_sap = converter_nan(df_prov_sap.loc[0, "TERM_A"])
                    termB_sap = converter_nan(df_prov_sap.loc[0, "TERM_B"])
                    seloA_sap = converter_nan(df_prov_sap.loc[0, "SEAL_A"])
                    seloB_sap = converter_nan(df_prov_sap.loc[0, "SEAL_B"])

                    leadset_cao = df_prov_cao.loc[0, "Leadset"]
                    cabo_cao = str(df_prov_cao.loc[0, "Wire1Key"])
                    comp_cao = int(df_prov_cao.loc[0, "Wire1Length"])

                    # Leitura
                    termA_cao = converter_nan(df_prov_cao.loc[0, "Terminal1Key"])
                    termB_cao = converter_nan(df_prov_cao.loc[0, "Terminal2Key"])
                    seloA_cao = converter_nan(df_prov_cao.loc[0, "Seal1Key"])
                    seloB_cao = converter_nan(df_prov_cao.loc[0, "Seal2Key"])

                    # Boleta
                    termA_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText1"])
                    termB_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText2"])
                    seloA_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText3"])
                    seloB_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText4"])

                    if cabo_sap == cabo_cao:
                        pass
                    else:
                        df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                        df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                        df_comp.loc[linha_controle, "Wire Atual"] = cabo_cao
                        df_comp.loc[linha_controle, "Wire Novo"] = cabo_sap

                    if comp_sap == comp_cao:
                        pass
                    else:
                        df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                        df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                        df_comp.loc[linha_controle, "Length Atual"] = comp_cao
                        df_comp.loc[linha_controle, "Length Novo"] = comp_sap

                    # APLICADO
                    if (termA_sap == termA_cao and termB_sap == termB_cao) or (
                        termA_sap == termB_cao and termB_sap == termA_cao
                    ):
                        pass
                    else:
                        if termA_sap != termA_cao and termB_sap == termB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term A Atual"] = termA_cao
                            df_comp.loc[linha_controle, "Term A Novo"] = termA_sap

                        elif termA_sap == termA_cao and termB_sap != termB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term B Atual"] = termB_cao
                            df_comp.loc[linha_controle, "Term B Novo"] = termB_sap

                    if (seloA_sap == seloA_cao and seloB_sap == seloB_cao) or (
                        seloA_sap == seloB_cao and seloB_sap == seloA_cao
                    ):
                        pass
                    else:
                        if seloA_sap != seloA_cao and seloB_sap == seloB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo A Atual"] = seloA_cao
                            df_comp.loc[linha_controle, "Selo A Novo"] = seloA_sap

                        elif seloA_sap == seloA_cao and seloB_sap != seloB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo B Atual"] = seloB_cao
                            df_comp.loc[linha_controle, "Selo B Novo"] = seloB_sap

                    # BOLETA
                    if (termA_sap == termA_cao_bol and termB_sap == termB_cao_bol) or (
                        termA_sap == termB_cao_bol and termB_sap == termA_cao_bol
                    ):
                        # Nova parte
                        df_comp.loc[linha_controle, "Term A Atual"] = None
                        df_comp.loc[linha_controle, "Term A Novo"] = None
                        df_comp.loc[linha_controle, "Term B Atual"] = None
                        df_comp.loc[linha_controle, "Term B Novo"] = None
                    else:
                        if termA_sap != termA_cao_bol and termB_sap == termB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term A Atual [Boleta]"] = (
                                termA_cao_bol
                            )
                            df_comp.loc[linha_controle, "Term A Novo [Boleta]"] = termA_sap

                        elif termA_sap == termA_cao_bol and termB_sap != termB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term B Atual [Boleta]"] = (
                                termB_cao_bol
                            )
                            df_comp.loc[linha_controle, "Term B Novo [Boleta]"] = termB_sap

                    if (seloA_sap == seloA_cao_bol and seloB_sap == seloB_cao_bol) or (
                        seloA_sap == seloB_cao_bol and seloB_sap == seloA_cao_bol
                    ):
                        # Nova parte
                        df_comp.loc[linha_controle, "Selo A Atual"] = None
                        df_comp.loc[linha_controle, "Selo A Novo"] = None
                        df_comp.loc[linha_controle, "Selo B Atual"] = None
                        df_comp.loc[linha_controle, "Selo B Novo"] = None
                    else:
                        if seloA_sap != seloA_cao_bol and seloB_sap == seloB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo A Atual [Boleta]"] = (
                                seloA_cao_bol
                            )
                            df_comp.loc[linha_controle, "Selo A Novo [Boleta]"] = seloA_sap

                        elif seloA_sap == seloA_cao_bol and seloB_sap != seloB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo B Atual [Boleta]"] = (
                                seloB_cao_bol
                            )
                            df_comp.loc[linha_controle, "Selo B Novo [Boleta]"] = seloB_sap

                    linha_controle += 1

                # CABOS TWISTERS
                elif len(df_prov_sap) == 2 and len(df_prov_cao) == 1:
                    leadset_sap = df_prov_sap.loc[0, "Leadset"]
                    cabo_sap1 = str(df_prov_sap.loc[0, "WIRE_TUBE_SPLICE"])
                    cabo_sap2 = str(df_prov_sap.loc[1, "WIRE_TUBE_SPLICE"])

                    # comp_sap = int(df_prov_sap.loc[0,'Comp TW'])
                    try:
                        comp_sap = int(df_prov_sap.loc[0, "Comp TW"])
                    except ValueError:
                        comp_sap = 0  # ou trate de outra forma
                    if "MC" in leadset_sap and pd.isna(df_prov_sap.loc[0, "Comp TW"]):
                        comp_sap = int(df_prov_sap.loc[0, "LENGTH"])

                    termA_sap = converter_nan(df_prov_sap.loc[0, "TERM_A"])
                    termB_sap = converter_nan(df_prov_sap.loc[0, "TERM_B"])
                    seloA_sap = converter_nan(df_prov_sap.loc[0, "SEAL_A"])
                    seloB_sap = converter_nan(df_prov_sap.loc[0, "SEAL_B"])

                    cabo_cao1 = converter_nan(df_prov_cao.loc[0, "Wire1Key"])
                    cabo_cao2 = converter_nan(df_prov_cao.loc[0, "Wire2Key"])

                    leadset_cao = df_prov_cao.loc[0, "Leadset"]

                    if ("MC" in leadset_cao or "TW" in leadset_cao) and pd.isna(
                        df_prov_cao.loc[0, "Wire2Key"]
                    ):
                        cabo_cao2 = cabo_cao1

                    comp_cao = int(df_prov_cao.loc[0, "Wire1Length"])

                    # Leitura

                    termA_cao = converter_nan(df_prov_cao.loc[0, "Terminal1Key"])
                    termB_cao = converter_nan(df_prov_cao.loc[0, "Terminal2Key"])
                    seloA_cao = converter_nan(df_prov_cao.loc[0, "Seal1Key"])
                    seloB_cao = converter_nan(df_prov_cao.loc[0, "Seal2Key"])

                    # Boleta
                    termA_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText1"])
                    termB_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText2"])
                    seloA_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText3"])
                    seloB_cao_bol = converter_nan(df_prov_cao.loc[0, "UserText4"])

                    if (cabo_sap1 == cabo_cao1 and cabo_sap2 == cabo_cao2) or (
                        cabo_sap1 == cabo_cao2 and cabo_sap2 == cabo_cao1
                    ):
                        pass
                    else:
                        if cabo_sap1 != cabo_cao1 and cabo_sap2 == cabo_cao2:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Wire Atual"] = cabo_cao1
                            df_comp.loc[linha_controle, "Wire Novo"] = cabo_sap1
                        elif cabo_sap1 == cabo_cao1 and cabo_sap2 != cabo_cao2:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Wire2 Atual"] = cabo_cao2
                            df_comp.loc[linha_controle, "Wire2 Novo"] = cabo_sap2
                        else:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Wire Atual"] = cabo_cao1
                            df_comp.loc[linha_controle, "Wire Novo"] = cabo_sap1
                            df_comp.loc[linha_controle, "Wire2 Atual"] = cabo_cao2
                            df_comp.loc[linha_controle, "Wire2 Novo"] = cabo_sap2

                    if comp_sap == comp_cao:
                        pass
                    else:
                        df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                        df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                        df_comp.loc[linha_controle, "Length TW Atual"] = comp_cao
                        df_comp.loc[linha_controle, "Length TW Novo"] = comp_sap

                    # APLICADO
                    if (termA_sap == termA_cao and termB_sap == termB_cao) or (
                        termA_sap == termB_cao and termB_sap == termA_cao
                    ):
                        pass
                    else:
                        if termA_sap != termA_cao and termB_sap == termB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term A Atual"] = termA_cao
                            df_comp.loc[linha_controle, "Term A Novo"] = termA_sap

                        elif termA_sap == termA_cao and termB_sap != termB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term B Atual"] = termB_cao
                            df_comp.loc[linha_controle, "Term B Novo"] = termB_sap

                    if (seloA_sap == seloA_cao and seloB_sap == seloB_cao) or (
                        seloA_sap == seloB_cao and seloB_sap == seloA_cao
                    ):
                        pass
                    else:
                        if seloA_sap != seloA_cao and seloB_sap == seloB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo A Atual"] = seloA_cao
                            df_comp.loc[linha_controle, "Selo A Novo"] = seloA_sap

                        elif seloA_sap == seloA_cao and seloB_sap != seloB_cao:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo B Atual"] = seloB_cao
                            df_comp.loc[linha_controle, "Selo B Novo"] = seloB_sap

                    # BOLETA
                    if (termA_sap == termA_cao_bol and termB_sap == termB_cao_bol) or (
                        termA_sap == termB_cao_bol and termB_sap == termA_cao_bol
                    ):
                        # Nova parte
                        df_comp.loc[linha_controle, "Term A Atual"] = None
                        df_comp.loc[linha_controle, "Term A Novo"] = None
                        df_comp.loc[linha_controle, "Term B Atual"] = None
                        df_comp.loc[linha_controle, "Term B Novo"] = None
                    else:
                        if termA_sap != termA_cao_bol and termB_sap == termB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term A Atual [Boleta]"] = (
                                termA_cao_bol
                            )
                            df_comp.loc[linha_controle, "Term A Novo [Boleta]"] = termA_sap

                        elif termA_sap == termA_cao_bol and termB_sap != termB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Term B Atual [Boleta]"] = (
                                termB_cao_bol
                            )
                            df_comp.loc[linha_controle, "Term B Novo [Boleta]"] = termB_sap

                    if (seloA_sap == seloA_cao_bol and seloB_sap == seloB_cao_bol) or (
                        seloA_sap == seloB_cao_bol and seloB_sap == seloA_cao_bol
                    ):
                        # Nova parte
                        df_comp.loc[linha_controle, "Selo A Atual"] = None
                        df_comp.loc[linha_controle, "Selo A Novo"] = None
                        df_comp.loc[linha_controle, "Selo B Atual"] = None
                        df_comp.loc[linha_controle, "Selo B Novo"] = None
                    else:
                        if seloA_sap != seloA_cao_bol and seloB_sap == seloB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo A Atual [Boleta]"] = (
                                seloA_cao_bol
                            )
                            df_comp.loc[linha_controle, "Selo A Novo [Boleta]"] = seloA_sap

                        elif seloA_sap == seloA_cao_bol and seloB_sap != seloB_cao_bol:
                            df_comp.loc[linha_controle, "Leadset Atual"] = leadset_cao
                            df_comp.loc[linha_controle, "Leadset Novo"] = leadset_sap
                            df_comp.loc[linha_controle, "Selo B Atual [Boleta]"] = (
                                seloB_cao_bol
                            )
                            df_comp.loc[linha_controle, "Selo B Novo [Boleta]"] = seloB_sap

                    linha_controle += 1

                elif len(df_prov_sap) == 1 and len(df_prov_cao) == 0:
                    lista_ckt_diretos_add.append(ckt)

                elif len(df_prov_sap) == 2 and len(df_prov_cao) == 0:
                    lista_ckt_tw_add.append(ckt)

            colunas_ord = [
                "Leadset Atual",
                "Leadset Novo",
                "Wire Atual",
                "Wire Novo",
                "Wire2 Atual",
                "Wire2 Novo",
                "Length Atual",
                "Length Novo",
                "Length TW Atual",
                "Length TW Novo",
                "Term A Atual",
                "Term A Novo",
                "Term B Atual",
                "Term B Novo",
                "Selo A Atual",
                "Selo A Novo",
                "Selo B Atual",
                "Selo B Novo",
                "Term A Atual [Boleta]",
                "Term A Novo [Boleta]",
                "Term B Atual [Boleta]",
                "Term B Novo [Boleta]",
                "Selo A Atual [Boleta]",
                "Selo A Novo [Boleta]",
                "Selo B Atual [Boleta]",
                "Selo B Novo [Boleta]",
            ]
            df_comp = reordenar_colunas(df_comp, colunas_ord)

            df_comp = df_comp.reset_index(drop=True)

            dict_comp = {}
            dict_comp1 = {}
            if "Length Novo" in df_comp.columns:
                dict_comp = {
                    row["Leadset Novo"]: int(row["Length Novo"])
                    for i, row in df_comp.dropna(subset=["Length Novo"]).iterrows()
                }

            if "Length TW Novo" in df_comp.columns:
                dict_comp1 = {
                    row["Leadset Novo"]: int(row["Length TW Novo"])
                    for i, row in df_comp.dropna(subset=["Length TW Novo"]).iterrows()
                }

            if dict_comp and dict_comp1:
                dict_comp.update(dict_comp1)

            elif dict_comp1:
                dict_comp = dict_comp1

            # ATUALIZANDO CAO
            new_df_cao = df2[df2["Leadset"].str.startswith("G", na=False)].copy()

            for i in new_df_cao.index:
                leadset = new_df_cao.loc[i, "Leadset"]

                if retirar_undeline_sap and leadset[3] != "_":
                    leadset = f"{leadset[:3]}_{leadset[3:]}"

                if retirar_undeline_cao:
                    leadset = f"{leadset[:3]}{leadset[4:]}"
                if leadset[:1] == led[:1]:
                    new_df_cao.loc[i, "Leadset"] = leadset

                    if len(leadset) >= 4 and leadset[3] == "_" and leadset[4] == "_":
                        new_df_cao.loc[i, "UserWSText7"] = f"{leadset[:3]}-{leadset[5:]}"
                    elif len(leadset) >= 4 and leadset[3] == "_":
                        new_df_cao.loc[i, "UserWSText7"] = f"{leadset[:3]}-{leadset[4:]}"
                    else:
                        new_df_cao.loc[i, "UserWSText7"] = f"{leadset[:3]}-{leadset[4:]}"

                    if dict_comp and atualizar_dimensional:
                        new_df_cao.loc[i, "Wire1Length"] = dict_comp.get(
                            new_df_cao.loc[i, "Leadset"], new_df_cao.loc[i, "Wire1Length"]
                        )

                        new_df_cao.loc[i, "Wire2Length"] = dict_comp.get(
                            new_df_cao.loc[i, "Leadset"], new_df_cao.loc[i, "Wire2Length"]
                        )

                        new_df_cao.loc[i, "TwistWireLength"] = dict_comp.get(
                            new_df_cao.loc[i, "Leadset"],
                            new_df_cao.loc[i, "TwistWireLength"],
                        )

            ckt_add = pd.DataFrame(
                lista_ckt_diretos_add + lista_ckt_tw_add, columns=["Leadset"]
            )

            ckt_add["Status"] = "Adicionar"

            df_comp_completo = df_comp_completo.dropna(how="all").reset_index(drop=True)
            df_comp_completo = df_comp_completo.dropna(how="all", axis=1)

            colunas_ord = [
                "Leadset Atual",
                "Leadset Novo",
                "Wire Atual",
                "Wire Novo",
                "Wire2 Atual",
                "Wire2 Novo",
                "Length Atual",
                "Length Novo",
                "Length TW Atual",
                "Length TW Novo",
                "Term A Atual",
                "Term A Novo",
                "Term B Atual",
                "Term B Novo",
                "Term A Atual [Boleta]",
                "Term A Novo [Boleta]",
                "Term B Atual [Boleta]",
                "Term B Novo [Boleta]",
                "Selo A Atual",
                "Selo A Novo",
                "Selo B Atual",
                "Selo B Novo",
                "Selo A Atual [Boleta]",
                "Selo A Novo [Boleta]",
                "Selo B Atual [Boleta]",
                "Selo B Novo [Boleta]",
            ]

            df_comp_completo = reordenar_colunas(
                df_comp_completo, colunas_prioritarias=colunas_ord
            )

            df_comp_completo = pd.concat([df_comp_completo, df_comp], ignore_index=True)
            ckt_add_completo = pd.concat([ckt_add_completo, ckt_add], ignore_index=True)
            new_df_cao_completo = pd.concat(
                [new_df_cao_completo, new_df_cao], ignore_index=True
            )

        colunas_limpar = df_comp_completo.columns[2:]
        df_comp_completo = df_comp_completo.dropna(
            how="all", subset=colunas_limpar, axis=0
        ).reset_index(drop=True)
        df_comp_completo = df_comp_completo.dropna(how="all", axis=1).reset_index(drop=True)

        return df_comp_completo, ckt_add_completo, new_df_cao_completo
    except Exception as e:
        traceback.print_exc()


# =========================================================
# WORKER
# =========================================================

class WorkerComparacaoSAPCAO(QThread):

    finalizado = Signal()
    erro = Signal(str)

    def __init__(self, caminho_sap, caminho_cao, caminho_output, atualizar_dimensional):
        super().__init__()

        self.caminho_sap = caminho_sap
        self.caminho_cao = caminho_cao
        self.caminho_output = caminho_output
        self.atualizar_dimensional = True
    @log_errors_record
    def run(self):

        try:
            dict_consolidado = {}

            # ============================
            # PROCESSAMENTO PRINCIPAL
            # ============================
            
            df_prov, add_ckt, df_new_cao = Comp_SAP_vs_CAO(
                self.caminho_sap,
                self.caminho_cao,
                retirar_undeline_sap=False,
                retirar_undeline_cao=False,
                atualizar_dimensional=True,
            )

            dict_consolidado["Análise"] = df_prov
            dict_consolidado["Add Circuitos"] = add_ckt
            dict_consolidado["Cadastro CAO"] = df_new_cao
            

            data_atual = datetime.now().strftime("%Y_%m_%d")

            # ============================
            # SALVAR EXCEL PRINCIPAL
            # ============================
            arquivo_excel = os.path.join(
                self.caminho_output,
                f"Análise_CAO_Completo_{data_atual}.xlsx"
            )

            salvar_dict_df_em_excel(dict_consolidado, arquivo_excel)

            # ============================
            # SALVAR CSV
            # ============================
            # df_new_cao.to_csv(
            #     os.path.join(
            #         self.caminho_output,
            #         f"Cadastro_CAO_{data_atual}.csv"
            #     ),
            #     sep=";",
            #     index=False
            # )

            # ============================
            # SALVAR EXCEL CADASTRO
            # ============================
            # df_new_cao.dropna(how="all", axis=1).to_excel(
            #     os.path.join(
            #         self.caminho_output,
            #         f"Cadastro_CAO_{data_atual}.xlsx"
            #     ),
            #     index=False
            # )

            self.finalizado.emit()

        except Exception as e:
            self.erro.emit(str(e))


# =========================================================
# UI
# =========================================================

class TelaComparacaoSAPCAO(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TMGods 🧠 Industrial Engineering")
        self.resize(700, 250)
        self.centralizar_tela_atual()
        self.setStyleSheet(self.estilo())

        layout = QVBoxLayout()

        # =========================
        # SAP
        # =========================
        layout.addWidget(QLabel("Arquivo SAP"))

        h1 = QHBoxLayout()
        self.input_sap = QLineEdit()
        self.input_sap.setPlaceholderText("Selecione o arquivo SAP...")

        btn_sap = QPushButton()
        btn_sap.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        btn_sap.clicked.connect(self.selecionar_sap)

        h1.addWidget(self.input_sap)
        h1.addWidget(btn_sap)

        layout.addLayout(h1)

        layout.addWidget(self.linha())

        # =========================
        # CAO
        # =========================
        layout.addWidget(QLabel("Arquivo Master Data"))

        h2 = QHBoxLayout()
        self.input_cao = QLineEdit()
        self.input_cao.setPlaceholderText("Selecione o arquivo Master Data...")

        btn_cao = QPushButton()
        btn_cao.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        btn_cao.clicked.connect(self.selecionar_cao)

        h2.addWidget(self.input_cao)
        h2.addWidget(btn_cao)

        layout.addLayout(h2)

        layout.addWidget(self.linha())

        # =========================
        # OUTPUT
        # =========================
        layout.addWidget(QLabel("Salvar em"))

        h3 = QHBoxLayout()
        self.input_output = QLineEdit()
        self.input_output.setPlaceholderText("Selecione a pasta de saída...")

        btn_output = QPushButton()
        btn_output.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))
        btn_output.clicked.connect(self.selecionar_output)

        h3.addWidget(self.input_output)
        h3.addWidget(btn_output)

        layout.addLayout(h3)

        layout.addWidget(self.linha())

        # =========================
        # BOTÃO EXECUTAR
        # =========================
        layout.addStretch()

        self.btn_exec = QPushButton("Executar")
        self.btn_exec.setObjectName("botaoExec")
        self.btn_exec.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.btn_exec.setFixedWidth(140)
        self.btn_exec.clicked.connect(self.executar)

        layout.addWidget(self.btn_exec, alignment=Qt.AlignRight)

        self.setLayout(layout)

    # =========================================================
    # EXECUÇÃO
    # =========================================================
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
    def executar(self):

        sap = self.input_sap.text().strip()
        cao = self.input_cao.text().strip()
        output = self.input_output.text().strip()

        if not sap or not cao or not output:
            QMessageBox.warning(self, "Atenção", "Preencha todos os campos.")
            return

        # resposta = QMessageBox.question(
        #     self,
        #     "Atualizar dimensional",
        #     "Deseja atualizar o dimensional do CAO?",
        #     QMessageBox.Yes | QMessageBox.No
        # )

        # atualizar_dimensional = (resposta == QMessageBox.Yes)

        self.progresso = QProgressDialog("Processando...", None, 0, 0, self)
        self.progresso.setCancelButton(None)
        self.progresso.setWindowTitle("Aguarde")
        self.progresso.show()

        self.thread = WorkerComparacaoSAPCAO(
            caminho_sap=sap,
            caminho_cao=cao,
            caminho_output=output,
            atualizar_dimensional=True
        )

        self.thread.finalizado.connect(self.finalizado)
        self.thread.erro.connect(self.erro)

        self.thread.start()

    # =========================================================
    # CALLBACKS
    # =========================================================

    @log_errors_record
    def finalizado(self):
        self.progresso.close()
        QMessageBox.information(self, "Sucesso", "Arquivos gerados com sucesso!")
    @log_errors_record
    def erro(self, msg):
        self.progresso.close()
        QMessageBox.critical(self, "Erro", msg)

    # =========================================================
    # SELETORES
    # =========================================================
    @log_errors_record
    def selecionar_sap(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Selecionar SAP", "", "Excel (*.xlsx)"
        )
        if file:
            self.input_sap.setText(file)


    @log_errors_record
    def selecionar_cao(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Selecionar CAO", "", "CSV (*.csv)"
        )
        if file:
            self.input_cao.setText(file)


    @log_errors_record
    def selecionar_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Selecionar pasta de saída"
        )
        if folder:
            self.input_output.setText(folder)

    # =========================================================
    # UI HELPERS
    # =========================================================

    @log_errors_record
    def linha(self):
        sep = QLabel("")
        sep.setStyleSheet("background-color:#444; max-height:1px;")
        return sep
    
    
    @log_errors_record
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

    import sys

    app = QApplication(sys.argv)

    janela = TelaComparacaoSAPCAO()
    janela.show()

    sys.exit(app.exec())