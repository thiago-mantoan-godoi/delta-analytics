import pandas as pd
import os, sys
from datetime import datetime, timedelta
root_path = os.path.abspath(os.path.join(os.getcwd(), ".."))
if root_path not in sys.path:
    sys.path.append(root_path)
    
    
    

from utils.function import *
from art import *
import warnings

warnings.simplefilter("ignore")


logger = LoggerTerminal(salvar_log=False, typing=True)

from pathlib import Path

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).resolve().parent.parent


colunas = [
    "Leadset",
    "ProdVersion",
    "Description",
    "CableClass",
    "BatchSize",
    "PlanTimeBatch",
    "Font1",
    "Font2",
    "FontDescription1",
    "FontDescription2",
    "Wire1Key",
    "Wire1Name",
    "Wire1CrossSection",
    "Wire1Length",
    "Wire2Key",
    "Wire2Name",
    "Wire2CrossSection",
    "Wire2Length",
    "Terminal1Key",
    "Terminal1Name",
    "StrippingLength1",
    "PartStripLength1",
    "Terminal2Key",
    "Terminal2Name",
    "StrippingLength2",
    "PartStripLength2",
    "Terminal3Key",
    "Terminal3Name",
    "StrippingLength3",
    "PartStripLength3",
    "Terminal4Key",
    "Terminal4Name",
    "StrippingLength4",
    "PartStripLength4",
    "Seal1Key",
    "Seal1Name",
    "Seal2Key",
    "Seal2Name",
    "Seal3Key",
    "Seal3Name",
    "Seal4Key",
    "Seal4Name",
    "TwistWireLength",
    "PitchLength",
    "OpenEndLength1",
    "OpenEndLength2",
    "ReducedLeadLength",
    "ReducedWire",
    "Print11BeginText",
    "Print11BeginOffset",
    "Print11BeginTurnText",
    "Print12BeginText",
    "Print12BeginOffset",
    "Print12BeginTurnText",
    "Print13BeginText",
    "Print13BeginOffset",
    "Print13BeginTurnText",
    "Print14BeginText",
    "Print14BeginOffset",
    "Print14BeginTurnText",
    "Print15BeginText",
    "Print15BeginOffset",
    "Print15BeginTurnText",
    "Print16BeginText",
    "Print16BeginOffset",
    "Print16BeginTurnText",
    "Print17BeginText",
    "Print17BeginOffset",
    "Print17BeginTurnText",
    "Print18BeginText",
    "Print18BeginOffset",
    "Print18BeginTurnText",
    "Print19BeginText",
    "Print19BeginOffset",
    "Print19BeginTurnText",
    "Print11EndText",
    "Print11EndOffset",
    "Print11EndTurnText",
    "Print12EndText",
    "Print12EndOffset",
    "Print12EndTurnText",
    "Print13EndText",
    "Print13EndOffset",
    "Print13EndTurnText",
    "Print14EndText",
    "Print14EndOffset",
    "Print14EndTurnText",
    "Print15EndText",
    "Print15EndOffset",
    "Print15EndTurnText",
    "Print16EndText",
    "Print16EndOffset",
    "Print16EndTurnText",
    "Print17EndText",
    "Print17EndOffset",
    "Print17EndTurnText",
    "Print18EndText",
    "Print18EndOffset",
    "Print18EndTurnText",
    "Print19EndText",
    "Print19EndOffset",
    "Print19EndTurnText",
    "Print1ContText",
    "Print1ContOffset",
    "Print1ContTurnText",
    "Print1ContAlternateText",
    "Print21BeginText",
    "Print21BeginOffset",
    "Print21BeginTurnText",
    "Print22BeginText",
    "Print22BeginOffset",
    "Print22BeginTurnText",
    "Print23BeginText",
    "Print23BeginOffset",
    "Print23BeginTurnText",
    "Print24BeginText",
    "Print24BeginOffset",
    "Print24BeginTurnText",
    "Print25BeginText",
    "Print25BeginOffset",
    "Print25BeginTurnText",
    "Print26BeginText",
    "Print26BeginOffset",
    "Print26BeginTurnText",
    "Print27BeginText",
    "Print27BeginOffset",
    "Print27BeginTurnText",
    "Print28BeginText",
    "Print28BeginOffset",
    "Print28BeginTurnText",
    "Print29BeginText",
    "Print29BeginOffset",
    "Print29BeginTurnText",
    "Print21EndText",
    "Print21EndOffset",
    "Print21EndTurnText",
    "Print22EndText",
    "Print22EndOffset",
    "Print22EndTurnText",
    "Print23EndText",
    "Print23EndOffset",
    "Print23EndTurnText",
    "Print24EndText",
    "Print24EndOffset",
    "Print24EndTurnText",
    "Print25EndText",
    "Print25EndOffset",
    "Print25EndTurnText",
    "Print26EndText",
    "Print26EndOffset",
    "Print26EndTurnText",
    "Print27EndText",
    "Print27EndOffset",
    "Print27EndTurnText",
    "Print28EndText",
    "Print28EndOffset",
    "Print28EndTurnText",
    "Print29EndText",
    "Print29EndOffset",
    "Print29EndTurnText",
    "Print2ContText",
    "Print2ContOffset",
    "Print2ContTurnText",
    "Print2ContAlternateText",
    "Hotstamp1BeginOffset",
    "Hotstamp1EndOffset",
    "Hotstamp1BeginAndEnd",
    "Hotstamp2BeginOffset",
    "Hotstamp2EndOffset",
    "Hotstamp2BeginAndEnd",
    "UserText1",
    "UserText2",
    "UserText3",
    "UserText4",
    "UserText5",
    "UserText6",
    "UserText7",
    "UserText8",
    "UserText9",
    "UserText10",
    "UserWSText1",
    "UserWSText2",
    "UserWSText3",
    "UserWSText4",
    "UserWSText5",
    "UserWSText6",
    "UserWSText7",
    "UserWSText8",
    "UserWSText9",
    "UserWSText10",
    "StrippingLengthB1",
    "StrippingLengthB2",
    "StrippingLengthB3",
    "StrippingLengthB4",
    "StrippingLengthC1",
    "StrippingLengthC2",
    "StrippingLengthC3",
    "StrippingLengthC4",
    "StrippingLengthD1",
    "StrippingLengthD2",
    "StrippingLengthD3",
    "StrippingLengthD4",
]


def criar_base(caminho_arquivo, number=3):
    
    dados = pd.read_excel(caminho_arquivo,dtype=str)
    dados = dados[
        [
            "Família",
            "Fam_UCS",
            "Leadset",
            "Wire Nb",
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
            "Quantidade"
        ]
    ]

    df_diretos = dados[dados["Multicore"].isna()].reset_index(drop=True)
    df_TW = dados[dados["Multicore"].notna()].reset_index(drop=True)

    df_diretos["Leadset"] = df_diretos["Fam_UCS"].astype(str) + df_diretos[
        "Wire Nb"
    ].astype(str)

    df_new_diretos = pd.DataFrame(columns=colunas)

    ckts = df_diretos["Leadset"].dropna().unique().tolist()

    for ckt in ckts:
        df_prov = df_diretos[df_diretos["Leadset"] == ckt].reset_index(drop=True)
        df_prov["Rate"] = pd.to_numeric(df_prov["Rate"], errors="coerce")
        df_prov["Maco"] = pd.to_numeric(df_prov["Maco"], errors="coerce")

        df_new = {}

        df_new["Leadset"] = str(df_prov.loc[0, "Fam_UCS"]) + str(
            df_prov.loc[0, "Wire Nb"]
        )
        df_new["ProdVersion"] = number
        if pd.notna(df_prov.loc[0, "P. Vision"]):
            ucs = df_prov.loc[0, "Fam_UCS"]
            if ucs[3] == "_":
                ucs = ucs[:3]
            df_new["Description"] = (
                str(ucs)
                + "-"
                + str(df_prov.loc[0, "Wire Nb"])
                + " RATE: "
                + str(df_prov.loc[0, "Rate"])
                + " P: "
                + str(df_prov.loc[0, "P. Vision"])
            )
        else:
            ucs = df_prov.loc[0, "Fam_UCS"]
            if len(ucs) > 3 and ucs[3] == "_":
                    ucs = ucs[:3]
            df_new["Description"] = (
                str(ucs)
                + "-"
                + str(df_prov.loc[0, "Wire Nb"])
                + " RATE: "
                + str(df_prov.loc[0, "Rate"])
            )
        df_new["CableClass"] = "S"
        df_new["BatchSize"] = int(df_prov.loc[0, "Maco"])
        df_new["PlanTimeBatch"] = round(
            int(df_prov.loc[0, "Maco"]) / (int(df_prov.loc[0, "Rate"]) / 60), 3
        )
        df_new["Wire1Key"] = df_prov.loc[0, "Int.PN"]
        df_new["Wire1CrossSection"] = df_prov.loc[0, "CSA"]
        df_new["Wire1Length"] = int(df_prov.loc[0, "Length"])

        if "CORTE" in str(df_prov.loc[0, "Processo_A"]).upper():
            df_new["Terminal1Key"] = df_prov.loc[0, "Term. 1"]
            df_new["Seal1Key"] = df_prov.loc[0, "Seal 1"]

        if "CORTE" in str(df_prov.loc[0, "Processo_B"]).upper():
            df_new["Terminal2Key"] = df_prov.loc[0, "Term. 2"]
            df_new["Seal2Key"] = df_prov.loc[0, "Seal 2"]

        df_new["StrippingLength1"] = df_prov.loc[0, "Strip 1"]
        df_new["PartStripLength1"] = 3
        df_new["StrippingLength2"] = df_prov.loc[0, "Strip 2"]
        df_new["PartStripLength2"] = 3

        df_new["UserText1"] = df_prov.loc[0, "Term. 1"]
        df_new["UserText2"] = df_prov.loc[0, "Term. 2"]
        df_new["UserText3"] = df_prov.loc[0, "Seal 1"]
        df_new["UserText4"] = df_prov.loc[0, "Seal 2"]

        df_new["UserWSText1"] = df_prov.loc[0, "Processo_A"]
        df_new["UserWSText2"] = df_prov.loc[0, "Processo_B"]

        df_new["UserWSText5"] = "LT: 0 C BR: 0 C AM: 0 Abst:0"

        ucs = df_prov.loc[0, "Fam_UCS"]
        if len(ucs) > 3 and ucs[3] == "_":
                ucs = ucs[:3]
        df_new["UserWSText7"] = str(ucs)[:3] + "-" + str(df_prov.loc[0, "Wire Nb"])

        if pd.notna(df_prov.loc[0, "P. Vision"]):
            df_new["CVXVisionSystemCommand"] = df_prov.loc[0, "P. Vision"]

        df_new_diretos = pd.concat(
            [df_new_diretos, pd.DataFrame([df_new])], ignore_index=True
        )

    colunas_prioritarias = colunas
    df_new = reordenar_colunas(df_new_diretos, colunas_prioritarias)

    df_new_diretos = df_new_diretos.drop_duplicates().reset_index(drop=True)

    # CADASTRO TWISTER----------------------------------------------------------------------------------------------------------------------------------


    df_new_TW = pd.DataFrame(columns=colunas)
    if len(df_TW) > 0:
        if len(df_TW) % 2 == 0:
            ckt_tw = df_TW["Multicore"].dropna().unique().tolist()

            for ckt in ckt_tw:
                df_prov = df_TW[df_TW["Multicore"] == ckt].reset_index(drop=True)
                df_prov["Rate"] = pd.to_numeric(df_prov["Rate"], errors="coerce")
                df_prov["Maco"] = pd.to_numeric(df_prov["Maco"], errors="coerce")

                dict_TW = {}

                dict_TW["Leadset"] = str(df_prov.loc[0, "Fam_UCS"]) + str(
                    df_prov.loc[0, "Multicore"]
                )

                dict_TW["ProdVersion"] = number

                if pd.isna(df_prov.loc[0, "P. Vision"]):
                    ucs = df_prov.loc[0, "Fam_UCS"]
                    if len(ucs) > 3 and ucs[3] == "_":
                            ucs = ucs[:3]
                    dict_TW["Description"] = (
                        str(ucs)
                        + "-"
                        + str(df_prov.loc[0, "Multicore"])
                        + " RATE: "
                        + str(df_prov.loc[0, "Rate"])
                    )
                else:
                    ucs = df_prov.loc[0, "Fam_UCS"]
                    if len(ucs) > 3 and ucs[3] == "_":
                            ucs = ucs[:3]
                    dict_TW["Description"] = (
                        str(ucs)
                        + "-"
                        + str(df_prov.loc[0, "Multicore"])
                        + " RATE: "
                        + str(df_prov.loc[0, "Rate"])
                        + " P: "
                        + str(df_prov.loc[0, "P. Vision"])
                    )

                dict_TW["CableClass"] = "T"

                dict_TW["BatchSize"] = int(df_prov.loc[0, "Maco"])

                dict_TW["PlanTimeBatch"] = float(
                    round(df_prov.loc[0, "Maco"] / (df_prov.loc[0, "Rate"] / 60), 3)
                )

                dict_TW["Wire1Key"] = df_prov.loc[0, "Int.PN"]
                dict_TW["Wire2Key"] = df_prov.loc[1, "Int.PN"]

                dict_TW["Wire1CrossSection"] = float(df_prov.loc[0, "CSA"])
                dict_TW["Wire2CrossSection"] = float(df_prov.loc[1, "CSA"])

                dict_TW["Wire1Length"] = int(df_prov.loc[0, "Comp_TW"])
                dict_TW["Wire2Length"] = int(df_prov.loc[0, "Comp_TW"])

                # Terminais
                dict_TW["Terminal1Key"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Term. 1"])
                    else str(df_prov.loc[0, "Term. 1"])
                )
                dict_TW["Terminal3Key"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Term. 1"])
                    else str(df_prov.loc[1, "Term. 1"])
                )
                dict_TW["Terminal2Key"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Term. 2"])
                    else str(df_prov.loc[0, "Term. 2"])
                )
                dict_TW["Terminal4Key"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Term. 2"])
                    else str(df_prov.loc[1, "Term. 2"])
                )

                # Strip
                dict_TW["StrippingLength1"] = float(df_prov.loc[0, "Strip 1"])
                dict_TW["StrippingLength3"] = float(df_prov.loc[1, "Strip 1"])
                dict_TW["StrippingLength2"] = float(df_prov.loc[0, "Strip 2"])
                dict_TW["StrippingLength4"] = float(df_prov.loc[1, "Strip 2"])

                # Part Strip
                dict_TW["PartStripLength1"] = (
                    None if pd.isna(df_prov.loc[0, "Strip 1"]) else 3
                )
                dict_TW["PartStripLength3"] = (
                    None if pd.isna(df_prov.loc[1, "Strip 1"]) else 3
                )
                dict_TW["PartStripLength2"] = (
                    None if pd.isna(df_prov.loc[1, "Strip 2"]) else 3
                )
                dict_TW["PartStripLength4"] = (
                    None if pd.isna(df_prov.loc[1, "Strip 2"]) else 3
                )

                # Selos
                dict_TW["Seal1Key"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Seal 1"])
                    else str(df_prov.loc[0, "Seal 1"])
                )
                dict_TW["Seal3Key"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Seal 1"])
                    else str(df_prov.loc[1, "Seal 1"])
                )
                dict_TW["Seal2Key"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Seal 2"])
                    else str(df_prov.loc[0, "Seal 2"])
                )
                dict_TW["Seal4Key"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Seal 2"])
                    else str(df_prov.loc[1, "Seal 2"])
                )

                # Crítico
                dict_TW["TwistWireLength"] = int(df_prov.loc[0, "Comp_TW"])
                dict_TW["PitchLength"] = 25

                if pd.isna(df_prov.loc[0, "Term. 1"]):
                    print("aqui 1")

                    dict_TW["OpenEndLength2"] = 113
                    dict_TW["OpenEndLength1"] = 50
                    
                    dict_TW["ReducedLeadLength"] = 15
                    if float(df_prov.loc[0, "Length"]) > float(
                        df_prov.loc[1, "Length"]
                    ) or float(df_prov.loc[1, "Length"]) > float(
                        df_prov.loc[0, "Length"]
                    ):
                        if float(df_prov.loc[1, "Length"]) > float(
                            df_prov.loc[0, "Length"]
                        ):
                            dict_TW["ReducedWire"] = 1
                        else:
                            dict_TW["ReducedWire"] = 2
                        dict_TW["Terminal1Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Term. 2"])
                            else str(df_prov.loc[0, "Term. 2"])
                        )
                        dict_TW["Terminal3Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Term. 2"])
                            else str(df_prov.loc[1, "Term. 2"])
                        )
                        dict_TW["Terminal2Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Term. 1"])
                            else str(df_prov.loc[0, "Term. 1"])
                        )
                        dict_TW["Terminal4Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Term. 1"])
                            else str(df_prov.loc[1, "Term. 1"])
                        )

                        # Strip
                        dict_TW["StrippingLength1"] = float(df_prov.loc[0, "Strip 2"])
                        dict_TW["StrippingLength3"] = float(df_prov.loc[1, "Strip 2"])
                        dict_TW["StrippingLength2"] = float(df_prov.loc[0, "Strip 1"])
                        dict_TW["StrippingLength4"] = float(df_prov.loc[1, "Strip 1"])

                        # Part Strip
                        dict_TW["PartStripLength1"] = (
                            None if pd.isna(df_prov.loc[0, "Strip 2"]) else 3
                        )
                        dict_TW["PartStripLength3"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 2"]) else 3
                        )
                        dict_TW["PartStripLength2"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 1"]) else 3
                        )
                        dict_TW["PartStripLength4"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 1"]) else 3
                        )

                        # Selos
                        dict_TW["Seal1Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Seal 2"])
                            else str(df_prov.loc[0, "Seal 2"])
                        )
                        dict_TW["Seal3Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Seal 2"])
                            else str(df_prov.loc[1, "Seal 2"])
                        )
                        dict_TW["Seal2Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Seal 1"])
                            else str(df_prov.loc[0, "Seal 1"])
                        )
                        dict_TW["Seal4Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Seal 1"])
                            else str(df_prov.loc[1, "Seal 1"])
                        )
                    else:
                        dict_TW["ReducedWire"] = 2
                elif pd.isna(df_prov.loc[0, "Term. 2"]):
                    dict_TW["OpenEndLength2"] = 50
                    dict_TW["OpenEndLength1"] = 113
                    dict_TW["ReducedLeadLength"] = 15
                    if float(df_prov.loc[0, "Length"]) > float(
                        df_prov.loc[1, "Length"]
                    ) or float(df_prov.loc[1, "Length"]) > float(
                        df_prov.loc[0, "Length"]
                    ):
                        if float(df_prov.loc[1, "Length"]) > float(
                            df_prov.loc[0, "Length"]
                        ):
                            dict_TW["ReducedWire"] = 1
                        else:
                            dict_TW["ReducedWire"] = 2
                        dict_TW["Terminal1Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Term. 2"])
                            else str(df_prov.loc[0, "Term. 2"])
                        )
                        dict_TW["Terminal3Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Term. 2"])
                            else str(df_prov.loc[1, "Term. 2"])
                        )
                        dict_TW["Terminal2Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Term. 1"])
                            else str(df_prov.loc[0, "Term. 1"])
                        )
                        dict_TW["Terminal4Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Term. 1"])
                            else str(df_prov.loc[1, "Term. 1"])
                        )

                        # Strip
                        dict_TW["StrippingLength1"] = float(df_prov.loc[0, "Strip 2"])
                        dict_TW["StrippingLength3"] = float(df_prov.loc[1, "Strip 2"])
                        dict_TW["StrippingLength2"] = float(df_prov.loc[0, "Strip 1"])
                        dict_TW["StrippingLength4"] = float(df_prov.loc[1, "Strip 1"])

                        # Part Strip
                        dict_TW["PartStripLength1"] = (
                            None if pd.isna(df_prov.loc[0, "Strip 2"]) else 3
                        )
                        dict_TW["PartStripLength3"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 2"]) else 3
                        )
                        dict_TW["PartStripLength2"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 1"]) else 3
                        )
                        dict_TW["PartStripLength4"] = (
                            None if pd.isna(df_prov.loc[1, "Strip 1"]) else 3
                        )

                        # Selos
                        dict_TW["Seal1Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Seal 2"])
                            else str(df_prov.loc[0, "Seal 2"])
                        )
                        dict_TW["Seal3Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Seal 2"])
                            else str(df_prov.loc[1, "Seal 2"])
                        )
                        dict_TW["Seal2Key"] = (
                            None
                            if pd.isna(df_prov.loc[0, "Seal 1"])
                            else str(df_prov.loc[0, "Seal 1"])
                        )
                        dict_TW["Seal4Key"] = (
                            None
                            if pd.isna(df_prov.loc[1, "Seal 1"])
                            else str(df_prov.loc[1, "Seal 1"])
                        )
                    else:
                        dict_TW["ReducedWire"] = 1
                else:
                    dict_TW["OpenEndLength1"] = 50
                    dict_TW["OpenEndLength2"] = 50

                dict_TW["UserText1"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Term. 1"])
                    else str(df_prov.loc[0, "Term. 1"])
                )
                dict_TW["UserText2"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Term. 2"])
                    else str(df_prov.loc[1, "Term. 2"])
                )
                dict_TW["UserText3"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Seal 1"])
                    else str(df_prov.loc[0, "Seal 1"])
                )
                dict_TW["UserText4"] = (
                    None
                    if pd.isna(df_prov.loc[1, "Seal 2"])
                    else str(df_prov.loc[1, "Seal 2"])
                )

                dict_TW["UserWSText1"] = str(df_prov.loc[0, "Processo_A"])
                dict_TW["UserWSText2"] = str(df_prov.loc[0, "Processo_B"])

                dict_TW["UserWSText3"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Note 1"])
                    else str(df_prov.loc[0, "Note 1"])
                )
                dict_TW["UserWSText4"] = (
                    None
                    if pd.isna(df_prov.loc[0, "Note 2"])
                    else str(df_prov.loc[0, "Note 2"])
                )

                dict_TW["UserWSText5"] = "Lt: 4 C BR: 1 C AM: 1 Abst:2"

                dict_TW["UserWSText6"] = (
                    ""
                    if pd.isna(df_prov.loc[0, "Severidade"])
                    else str(df_prov.loc[0, "Severidade"])
                )

                ucs = df_prov.loc[0, "Fam_UCS"]
                if len(ucs) > 3 and ucs[3] == "_":
                        ucs = ucs[:3]

                dict_TW["UserWSText7"] = (
                    str(ucs) + "-" + str(df_prov.loc[0, "Multicore"])
                )

                if pd.notna(df_prov.loc[0, "P. Vision"]):
                    dict_TW["CVXVisionSystemCommand"] = df_prov.loc[0, "P. Vision"]

                df_new_TW = pd.concat(
                    [df_new_TW, pd.DataFrame([dict_TW])], ignore_index=True
                )

        else:
            logger.error(f"Verificar, falta cabos na Twister: {len(df_TW) % 2}")

    else:
        logger.info("Sem cadastro de Twister.")

    df_consolidado = pd.concat([df_new_diretos, df_new_TW], ignore_index=True)

    return df_consolidado


def gerar_ordens(path_cadastro, df, number:int=3):
    if number > 0:
        df_cadastro = pd.read_excel(path_cadastro,dtype=str)

        dict_volume = {
            row["Leadset"]: row["Quantidade"] for i, row in df_cadastro.iterrows()
        }

        colunas = [
            "OrderNo",
            "Leadset",
            "ProdVersion",
            "Description",
            "Quantity",
            "TargetDate",
            "FlagPrototype",
            "BatchSize",
            "TargetLocation",
        ]

        arquivo_ordens = pd.DataFrame(columns=colunas)

        arquivo_ordens["Leadset"] = df["Leadset"]
        arquivo_ordens["OrderNo"] = range(1, len(arquivo_ordens) + 1)
        arquivo_ordens["ProdVersion"] = number
        arquivo_ordens["Quantity"] = df["Leadset"].map(dict_volume)
        arquivo_ordens["BatchSize"] = df["Leadset"].map(dict_volume)
        arquivo_ordens["FlagPrototype"] = 0

        if arquivo_ordens["ProdVersion"].loc[0] == 3:
            arquivo_ordens["TargetLocation"] = "Enviar Prot."
            arquivo_ordens["Description"] = "Corte para Protótipo"
        elif arquivo_ordens["ProdVersion"].loc[0] == 2:
            arquivo_ordens["TargetLocation"] = "Enviar P&A."
            arquivo_ordens["Description"] = "Corte para P&A"

        agora = datetime.now()

        # Adiciona um dia
        amanha = agora + timedelta(days=1)

        # Formata como desejado: YYYYMMDD HHMMSS
        formato_personalizado = amanha.strftime("%Y%m%d %H%M%S")

        arquivo_ordens["TargetDate"] = formato_personalizado

        return arquivo_ordens
    else:
        return pd.DataFrame()


def cadastro_componentes(dados):
    def get_unique_keys(cols):
        """Obtém valores únicos, removendo NaN e strings vazias, convertendo tudo para string."""
        valores = dados[cols].values.ravel()
        unicos = set(valores)
        return sorted(
            str(x).strip() for x in unicos if pd.notna(x) and str(x).strip() != ""
        )

    wire_cols = ["Wire1Key", "Wire2Key"]
    terminal_cols = ["Terminal1Key", "Terminal2Key", "Terminal3Key", "Terminal4Key"]
    seal_cols = ["Seal1Key", "Seal2Key", "Seal3Key", "Seal4Key"]

    wires = get_unique_keys(wire_cols)
    terminais = get_unique_keys(terminal_cols)
    selos = get_unique_keys(seal_cols)

    colunas_wire = [
        "WireKey",
        "Name",
        "Barcode",
        "Info",
        "WireType",
        "CrossSection",
        "IsoDiameter",
        "IsoMaterial",
        "TwistDirection",
        "Color1",
        "Color2",
        "Color3",
        "Color4",
        "NoOfStrands",
    ]

    colunas_terminais = [
        "TerminalKey",
        "Name",
        "Barcode",
        "Info",
        "TerminalType",
        "DoubleCrimpHorizontal",
        "FeedingType",
        "TerminalLength",
        "TerminalWidth",
        "TerminalOverlength",
    ]

    colunas_selos = [
        "SealKey",
        "Name",
        "Barcode",
        "Info",
        "SealLength",
        "SealWidth",
        "SealPositionTol",
        "Color",
    ]

    # Carregamento de Dados
    
    lista_de_cabos_path = BASE_DIR / "data" / "Lista_de_cabos.json"
    book_cabos = pd.read_json(lista_de_cabos_path,dtype=str)

    lista_legacy_path = BASE_DIR / "data" / "Lista_de_cabos_legacy.json"
    book_legacy = pd.read_json(lista_legacy_path,dtype=str)

    dict_legacy = {
        row["Part Number"]: row["Legacy"] for i, row in book_legacy.iterrows()
    }

    book_cabos = converter_cores(book_cabos)

    dict_cor1 = {row["Part Number"]: row["Cor 1"] for i, row in book_cabos.iterrows()}

    dict_cor2 = {row["Part Number"]: row["Cor 2"] for i, row in book_cabos.iterrows()}

    dict_IsoDiameter = {
        row["Part Number"]: row["Outer Diameter"] for i, row in book_cabos.iterrows()
    }
    dict_NoOfStrands = {
        row["Part Number"]: row["Number of Strands"] for i, row in book_cabos.iterrows()
    }

    dict_CrossSection = {}
    for i, row in book_cabos.iterrows():
        try:
            valor = str(row["Wire Size"]).split(" ")[1]
            dict_CrossSection[row["Part Number"]] = valor
        except (IndexError, TypeError):
            # Pode registrar ou ignorar o erro
            valor = str(row["Wire Size"]).split(" ")[0]
            dict_CrossSection[row["Part Number"]] = valor

    # CARREGAMENTO 2
    lista_de_terminais_path = BASE_DIR / "data" / "Lista_de_terminais.json"
    book_term = pd.read_json(lista_de_terminais_path,dtype=str)

    dict_TerminalType = {}

    for i, row in book_term.iterrows():
        part_number = row["Part Number"]
        classification = str(row["Part Classification"]).upper()
        description = str(row["Part Description"]).upper()

        if "FLAT" in classification or "FLAT" in description:
            if "GOLD" in classification or "GOLD" in description:
                dict_TerminalType[part_number] = "Flat Plug - Gold"
            elif "SLEEVE" in classification or "SLEEVE" in description:
                dict_TerminalType[part_number] = "Flat Plug Sleeve"
            else:
                dict_TerminalType[part_number] = "Flat Plug"

        if "ROUND" in classification or "ROUND" in description:
            if "GOLD" in classification or "GOLD" in description:
                dict_TerminalType[part_number] = "Round Connector - Gold"
            else:
                dict_TerminalType[part_number] = "Round Connector"

        if "BRACKET" in classification or "BRACKET" in description:
            if "OPEN" in classification or "OPEN" in description:
                dict_TerminalType[part_number] = "Wire Bracket (open)"
            else:
                dict_TerminalType[part_number] = "Wire Bracket (closed)"

        if ("FASTON" in classification or "FASTON" in description) and ("SLEEVE" in classification or "SLEEVE" in description):
            dict_TerminalType[part_number] = "Faston Plug with Sleeve"

    # CADASTRO CABOS
    df_wire = pd.DataFrame(columns=colunas_wire)
    df_wire["WireKey"] = wires
    df_wire["Barcode"] = wires
    df_wire["WireType"] = df_wire["WireKey"].map(dict_legacy)
    df_wire["IsoDiameter"] = df_wire["WireKey"].map(dict_IsoDiameter)
    df_wire["CrossSection"] = df_wire["WireKey"].map(dict_CrossSection)
    df_wire["IsoMaterial"] = "Undefined"
    df_wire["TwistDirection"] = "Counterclockwise (S)"
    df_wire["Color1"] = df_wire["WireKey"].map(dict_cor1)
    df_wire["Color2"] = df_wire["WireKey"].map(dict_cor2)
    df_wire["NoOfStrands"] = df_wire["WireKey"].map(dict_NoOfStrands)

    # CADASTRO TERMINAIS
    df_terminais = pd.DataFrame(columns=colunas_terminais)
    df_terminais["TerminalKey"] = terminais
    df_terminais["Barcode"] = terminais
    df_terminais["TerminalType"] = df_terminais["TerminalKey"].map(dict_TerminalType)
    df_terminais["DoubleCrimpHorizontal"] = "No"
    df_terminais["FeedingType"] = "Transverse left"

    df_terminais["TerminalLength"] = 22
    df_terminais["TerminalWidth"] = 3
    df_terminais["TerminalOverlength"] = 7

    # CADASTRO SELOS
    df_selos = pd.DataFrame(columns=colunas_selos)
    df_selos["SealKey"] = selos
    df_selos["Barcode"] = selos

    return df_wire, df_terminais, df_selos


def compara_material(
    df_wire, df_wire_cao, df_terminais, df_term_cao, df_selos, df_selos_cao
):
    if len(df_wire_cao) > 0:
        df_wire_exclusivo = df_wire[
            ~df_wire["WireKey"].isin(df_wire_cao["WireKey"])
        ].reset_index(drop=True)
    else:
        df_wire_exclusivo = df_wire

    if len(df_term_cao) > 0:
        df_term_exclusivo = df_terminais[
            ~df_terminais["TerminalKey"].isin(df_term_cao["TerminalKey"])
        ].reset_index(drop=True)
    else:
        df_term_exclusivo = df_terminais

    if len(df_selos_cao) > 0:
        df_selo_exclusivo = df_selos[
            ~df_selos["SealKey"].isin(df_selos_cao["SealKey"])
        ].reset_index(drop=True)
    else:
        df_selo_exclusivo = df_selos

    return df_wire_exclusivo, df_term_exclusivo, df_selo_exclusivo


def resposta_positiva(resposta):
    return resposta.strip().lower() in {"s", "sim", "y", "yes"}


def ler_csv_validado(nome_arquivo, coluna_esperada):
    while True:
        logger.info(f"Digite o caminho do arquivo de {nome_arquivo}: ")
        caminho = input("[>] ").strip().replace('"', "")

        if not os.path.exists(caminho):
            logger.error(f"O caminho '{caminho}' não existe. Tente novamente.")
            continue

        try:
            df = pd.read_csv(caminho, sep=";",dtype=str)
            if coluna_esperada not in df.columns:
                logger.error(
                    f"O arquivo '{nome_arquivo}' não contém a coluna obrigatória '{coluna_esperada}'."
                )
                continue
            logger.info(f"{nome_arquivo} carregado com sucesso.")
            return df
        except Exception as e:
            logger.error(f"Erro ao carregar '{nome_arquivo}': {e}")


def ler_excel_validado(nome_arquivo, coluna_esperada):
    while True:
        logger.info(f"Digite o caminho do arquivo de {nome_arquivo}: ")
        caminho = input("[>] ").strip().replace('"', "")

        if not os.path.exists(caminho):
            logger.error(f"O caminho '{caminho}' não existe. Tente novamente.")
            continue

        try:
            df = pd.read_excel(caminho,dtype=str)
            if coluna_esperada not in df.columns:
                logger.error(
                    f"O arquivo '{nome_arquivo}' não contém a coluna obrigatória '{coluna_esperada}'."
                )
                continue
            logger.info(f"{nome_arquivo} carregado com sucesso.")
            return df
        except Exception as e:
            logger.error(f"Erro ao carregar '{nome_arquivo}': {e}")


def gerar_sport_tape(dados, number=3):
    colunas = [
        "Leadset",
        "ProdVersion",
        "SpotTapingBeginKind",
        "SpotTapingBeginPosition",
        "SpotTapingBeginLength",
        "SpotTapingEndKind",
        "SpotTapingEndPosition",
        "SpotTapingEndLength",
    ]

    df_new = pd.DataFrame(columns=colunas)

    df_prov = dados[dados["Leadset"].str.contains("TW")]["Leadset"].reset_index(
        drop=True
    )

    if len(df_prov) > 0:
        df_new["Leadset"] = df_prov
        df_new["ProdVersion"] = number

        df_new["SpotTapingBeginKind"] = "Coroplast 301"
        df_new["SpotTapingBeginPosition"] = 0
        df_new["SpotTapingBeginLength"] = 32
        df_new["SpotTapingEndKind"] = "Coroplast 301"
        df_new["SpotTapingEndPosition"] = 0
        df_new["SpotTapingEndLength"] = 32

        # COMPLETAR-----------------------------------

        return df_new

    else:
        return pd.DataFrame()


def gerar_Tools_Terminals(dados, lista_de_aplicadores=None):
    if lista_de_aplicadores is None:
        lista_de_aplicadores = pd.DataFrame()
    else:
        dict_codigo_sap = {
            row["Terminal"]: row["Codigo SAP"]
            for i, row in lista_de_aplicadores.iterrows()
        }

    colunas = ["TerminalKey", "InventoryNo", "SealKey", "WireType1", "WireType2"]

    df_new = pd.DataFrame(columns=colunas)

    df_term = pd.DataFrame()

    df_term1 = (
        dados[["Terminal1Key", "Seal1Key"]]
        .dropna(how="all", axis=0)
        .drop_duplicates()
        .reset_index(drop=True)
        .rename(columns={"Terminal1Key": "Terminal", "Seal1Key": "Seal"})
    )
    df_term2 = (
        dados[["Terminal2Key", "Seal2Key"]]
        .dropna(how="all", axis=0)
        .drop_duplicates()
        .reset_index(drop=True)
        .rename(columns={"Terminal2Key": "Terminal", "Seal2Key": "Seal"})
    )
    df_term3 = (
        dados[["Terminal3Key", "Seal3Key"]]
        .dropna(how="all", axis=0)
        .drop_duplicates()
        .reset_index(drop=True)
        .rename(columns={"Terminal3Key": "Terminal", "Seal3Key": "Seal"})
    )
    df_term4 = (
        dados[["Terminal4Key", "Seal4Key"]]
        .dropna(how="all", axis=0)
        .drop_duplicates()
        .reset_index(drop=True)
        .rename(columns={"Terminal4Key": "Terminal", "Seal4Key": "Seal"})
    )

    df_term = pd.concat([df_term, df_term1], ignore_index=True)
    df_term = pd.concat([df_term, df_term2], ignore_index=True)
    df_term = pd.concat([df_term, df_term3], ignore_index=True)
    df_term = pd.concat([df_term, df_term4], ignore_index=True)

    df_term = df_term.drop_duplicates().reset_index(drop=True)

    df_new["TerminalKey"] = df_term["Terminal"]
    df_new["SealKey"] = df_term["Seal"]
    df_new["WireType1"] = "*"
    df_new["WireType2"] = "*"

    if len(lista_de_aplicadores) > 0:
        df_new["InventoryNo"] = df_new["TerminalKey"].map(dict_codigo_sap)
        df_new["InventoryNo"] = (
            df_new["TerminalKey"].map(dict_codigo_sap).fillna(0).astype(int)
        )
    else:
        df_new["InventoryNo"] = range(len(df_new))

    return df_new


def gerar_Terminal_Applicators(dados=None, lista_de_aplicadores=None):
    if dados is None:
        dados = pd.DataFrame()

    if lista_de_aplicadores is None:
        lista_de_aplicadores = pd.DataFrame()
    else:
        lista_de_aplicadores = lista_de_aplicadores[
            lista_de_aplicadores["Codigo SAP"].notna()
        ].reset_index(drop=True)

    if len(dados) > 0 and len(lista_de_aplicadores):
        lista_de_aplicadores = lista_de_aplicadores[
            ~lista_de_aplicadores["Terminal"].isin(dados["TerminalKey"])
        ].reset_index(drop=True)

        colunas = [
            "InventoryNo",
            "Manufacturer",
            "Description",
            "MaxCounter",
            "Intervall_VT2",
            "Intervall_VT3",
            "InspectionIntervall",
            "Pre_Intervall_VT",
            "MasterLocation",
            "Location",
            "LocationType",
            "RequalificationInterval",
            "RequalificationWarning",
        ]

        df_new = pd.DataFrame(columns=colunas)

        df_new["InventoryNo"] = lista_de_aplicadores["Codigo SAP"].astype(int)
        df_new["Manufacturer"] = lista_de_aplicadores["Fabricante"]
        df_new["MasterLocation"] = lista_de_aplicadores["Locação"]

        df_new["LocationType"] = "Machine"
        df_new["MaxCounter"] = 999999
        df_new["Intervall_VT2"] = 999999
        df_new["Intervall_VT3"] = 999999
        df_new["InspectionIntervall"] = 999999
        df_new["Pre_Intervall_VT"] = 999999
        df_new["RequalificationInterval"] = 999999
        df_new["RequalificationWarning"] = 999999

        return df_new
    else:
        return pd.DataFrame()


def gerar_CrimpSTD(dados:str=None): 


    lista_de_Qualidade_path = BASE_DIR / "data" / "Lista_de_criterios_Qualidade.json"
    df_spec = pd.read_json(lista_de_Qualidade_path,dtype=str)

    # Duplicando --------------------------------------------------------------------------------------------------------------------------
    iso = ["65766", "65860", "65927"]
    duplicated_rows = []  # Para armazenar as duplicatas

    # Loop para percorrer cada linha do DataFrame
    for i in df_spec.index:
        isolacao_value = str(df_spec.at[i, "Isolação"])

        # Se o valor de 'Isolação' estiver na lista 'iso'
        if isolacao_value in iso:
            # Para cada valor em 'iso', exceto o valor encontrado
            for item in iso:
                if isolacao_value != str(item):
                    # Faz uma cópia da linha original
                    duplicated_row = df_spec.loc[[i]].copy()
                    # Altera o valor de 'Isolação' para o valor atual da lista
                    duplicated_row["Isolação"] = item
                    # Adiciona a linha duplicada à lista
                    duplicated_rows.append(duplicated_row)

    # Agora adiciona as duplicatas ao DataFrame original
    df_spec = pd.concat([df_spec] + duplicated_rows, ignore_index=True)

    df_spec = df_spec.reset_index(drop=True)

    df_spec = df_spec.drop_duplicates().reset_index(drop=True)

    # Duplicando --------------------------------------------------------------------------------------------------------------------------

    # df_spec = df_spec[(df_spec['Relatório'].str.contains('USCAR|BR', na=False)) & (~df_spec['Relatório'].str.contains('PERM|perm|PROV|desvio|DESVIO|Desvio', na=False))].reset_index(drop=True)

    df_spec["Bitola"] = df_spec["Bitola"].replace(",", ".", regex=True)

    for (
        i,
        row,
    ) in df_spec.iterrows():  # Usando iterrows para obter as linhas e seus índices
        bitola = str(row["Bitola"])

        # Separa os valores por '+'
        valores = bitola.split("+")

        # Definindo valor1, valor2 e valor3
        valor1 = valores[0]
        valor2 = valores[1] if len(valores) > 1 else None
        valor3 = valores[2] if len(valores) > 2 else None

        # Formata valor1
        if len(valor1) < 4:
            valor1 = f"{float(valor1):.2f}"

        # Formata valor2 (se existir)
        if valor2 and len(valor2) < 4:
            valor2 = f"{float(valor2):.2f}"

        # Formata valor3 (se existir)
        if valor3 and len(valor3) < 4:
            valor3 = f"{float(valor3):.2f}"

        # Monta a nova bitola considerando os valores
        if valor3:
            df_spec.at[i, "Bitola"] = f"{valor1}+{valor2}+{valor3}"
        elif valor2:
            df_spec.at[i, "Bitola"] = f"{valor1}+{valor2}"
        else:
            df_spec.at[i, "Bitola"] = valor1

    df_prov_3 = pd.DataFrame()
    df_prov_4 = pd.DataFrame()
    df_prov_5 = pd.DataFrame()
    df_prov_6 = pd.DataFrame()

    
    lista_legacy_path = BASE_DIR / "data" / "Lista_de_cabos_legacy.json"
    df_spec_cabos = pd.read_json(lista_legacy_path,dtype=str)
    
    dict_tipos = {
        row["Part Number"]: row["Legacy"] for i, row in df_spec_cabos.iterrows()
    }

    colunas = [
        "TerminalKey",
        "SealKey",
        "WireType1",
        "CrossSection1",
        "WireType2",
        "CrossSection2",
        "InventoryNo",
        "StrippingLength",
        "StrippingLengthMinus",
        "StrippingLengthPlus",
        "StrandOverlapMin",
        "StrandOverlapMax",
        "CrimpHeight",
        "CrimpHeightTol",
        "CrimpHeightTolMinus",
        "CrimpWidth",
        "CrimpWidthTol",
        "CrimpWidthTolMinus",
        "IsoCrimpHeight",
        "IsoCrimpHeightTol",
        "IsoCrimpHeightTolMinus",
        "IsoCrimpWidth",
        "IsoCrimpWidthTol",
        "IsoCrimpWidthTolMinus",
        "PullOffForce",
        "CrimpHeightSecondTol",
        "CrimpHeightSecondTolMinus",
    ]

    df_completo = pd.DataFrame()

    df_diretos = dados[dados["CableClass"] == "S"].reset_index(drop=True)
    df_TW = dados[dados["CableClass"] == "T"].reset_index(drop=True)

    if len(df_diretos) > 0:
        df_prov_1 = (
            df_diretos[["Wire1Key", "Wire1CrossSection", "Terminal1Key", "Seal1Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire1Key": "Wire",
                    "Terminal1Key": "Terminal",
                    "Seal1Key": "Seal",
                    "Wire1CrossSection": "CrossSection",
                }
            )
        )
        df_prov_2 = (
            df_diretos[["Wire1Key", "Wire1CrossSection", "Terminal2Key", "Seal2Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire1Key": "Wire",
                    "Terminal2Key": "Terminal",
                    "Seal2Key": "Seal",
                    "Wire1CrossSection": "CrossSection",
                }
            )
        )
    else:
        df_prov_1 = pd.DataFrame()
        df_prov_2 = pd.DataFrame()
    if len(df_TW) > 0:
        df_prov_3 = (
            df_TW[["Wire1Key", "Wire1CrossSection", "Terminal1Key", "Seal1Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire1Key": "Wire",
                    "Terminal1Key": "Terminal",
                    "Seal1Key": "Seal",
                    "Wire1CrossSection": "CrossSection",
                }
            )
        )
        df_prov_4 = (
            df_TW[["Wire1Key", "Wire1CrossSection", "Terminal2Key", "Seal2Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire1Key": "Wire",
                    "Terminal2Key": "Terminal",
                    "Seal2Key": "Seal",
                    "Wire2CrossSection": "CrossSection",
                }
            )
        )

        df_prov_5 = (
            df_TW[["Wire2Key", "Wire2CrossSection", "Terminal3Key", "Seal3Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire2Key": "Wire",
                    "Terminal3Key": "Terminal",
                    "Seal3Key": "Seal",
                    "Wire2CrossSection": "CrossSection",
                }
            )
        )
        df_prov_6 = (
            df_TW[["Wire2Key", "Wire2CrossSection", "Terminal4Key", "Seal4Key"]]
            .dropna(how="all", axis=0)
            .drop_duplicates()
            .rename(
                columns={
                    "Wire2Key": "Wire",
                    "Terminal4Key": "Terminal",
                    "Seal4Key": "Seal",
                    "Wire2CrossSection": "CrossSection",
                }
            )
        )

    if len(df_prov_1) > 0:
        df_prov_1 = df_prov_1.dropna(subset=["Wire"]).reset_index(drop=True)
    if len(df_prov_2) > 0:
        df_prov_2 = df_prov_2.dropna(subset=["Wire"]).reset_index(drop=True)
    if len(df_prov_3) > 0:
        df_prov_3 = df_prov_3.dropna(subset=["Wire"]).reset_index(drop=True)
    if len(df_prov_4) > 0:
        df_prov_4 = df_prov_4.dropna(subset=["Wire"]).reset_index(drop=True)
    if len(df_prov_5) > 0:
        df_prov_5 = df_prov_5.dropna(subset=["Wire"]).reset_index(drop=True)
    if len(df_prov_6) > 0:
        df_prov_6 = df_prov_6.dropna(subset=["Wire"]).reset_index(drop=True)

    df_completo = pd.concat([df_completo, df_prov_1], ignore_index=False)
    df_completo = pd.concat([df_completo, df_prov_2], ignore_index=False)
    df_completo = pd.concat([df_completo, df_prov_3], ignore_index=False)
    df_completo = pd.concat([df_completo, df_prov_4], ignore_index=False)
    df_completo = pd.concat([df_completo, df_prov_5], ignore_index=False)
    df_completo = pd.concat([df_completo, df_prov_6], ignore_index=False)

    df_completo = (
        df_completo.drop_duplicates()
        .reset_index(drop=True)
        .rename(
            columns={
                "Wire": "WireType1",
                "Terminal": "TerminalKey",
                "Seal": "SealKey",
                "CrossSection": "CrossSection1",
            }
        )
    )

    df_completo["WireType1"] = df_completo["WireType1"].apply(
        lambda x: dict_tipos[x] if x in dict_tipos and dict_tipos[x] is not None else x
    )

    df_completo = df_completo.drop_duplicates().reset_index(drop=True)

    df_completo = df_completo.dropna(subset=["TerminalKey"]).reset_index(drop=True)

    for i in colunas:
        if i not in df_completo.columns:
            df_completo[i] = None
    df_completo = df_completo[colunas]

    for i in df_completo.index:
        term = df_completo.loc[i, "TerminalKey"]
        selo = df_completo.loc[i, "SealKey"]
        wire = df_completo.loc[i, "WireType1"]
        try:
            bit = float(df_completo.loc[i, "CrossSection1"])
        except:
            bit = str(df_completo.loc[i, "CrossSection1"])

        if pd.isna(selo) or selo == "" or selo == "nan" or selo == "None" or selo is None:
            selo = 0

        for j in df_spec.index:
            term_spec = df_spec.loc[j, "Terminal"]
            selo_spec = df_spec.loc[j, "Selo"]
            wire_spec = df_spec.loc[j, "Isolação"]
            bit_spec = df_spec.loc[j, "Bitola"]

            try:
                bit_spec = float(df_spec.loc[j, "Bitola"])
            except:
                bit_spec = str(df_spec.loc[j, "Bitola"])

            if pd.isna(selo_spec) or selo_spec == "" or selo_spec == "nan" or selo_spec == "None" or selo_spec is None:
                selo_spec = 0
            if (
                str(term).strip() == str(term_spec).strip()
                and str(selo).strip() == str(selo_spec).strip()
                and str(wire).strip() == str(wire_spec).strip()
                and str(bit).strip() == str(bit_spec).strip()
            ):
                

                # Altura
                df_completo.loc[i, "CrimpHeight"] = df_spec.loc[j, "CCH"]
                df_completo.loc[i, "CrimpHeightTolMinus"] = df_spec.loc[j, "CCH-"]
                df_completo.loc[i, "CrimpHeightTol"] = df_spec.loc[j, "CCH+"]

                # Largura
                df_completo.loc[i, "CrimpWidth"] = df_spec.loc[j, "CCW"]
                df_completo.loc[i, "CrimpWidthTolMinus"] = df_spec.loc[j, "CCW-"]
                df_completo.loc[i, "CrimpWidthTol"] = df_spec.loc[j, "CCW+"]

                # Altura Iso
                df_completo.loc[i, "IsoCrimpHeight"] = df_spec.loc[j, "ICH"]
                df_completo.loc[i, "IsoCrimpHeightTolMinus"] = df_spec.loc[j, "ICH-"]
                df_completo.loc[i, "IsoCrimpHeightTol"] = df_spec.loc[j, "ICH+"]

                # Largura iso
                df_completo.loc[i, "IsoCrimpWidth"] = df_spec.loc[j, "ICW"]
                df_completo.loc[i, "IsoCrimpWidthTolMinus"] = df_spec.loc[j, "ICW-"]
                df_completo.loc[i, "IsoCrimpWidthTol"] = df_spec.loc[j, "ICW+"]

    return df_completo


def cadastro(caminho_arquivo:str=None,
             caminho_output:str=None, 
             nivel:int=None,
             df_wire_cao:str=None,
             df_term_cao:str=None,
             df_selos_cao:str=None,
             df_app_cao:str=None):
    """Gera e exporta os arquivos de cadastro e apoio para integração com o CAO a partir de um arquivo de entrada.
    Organiza os dados em múltiplos arquivos CSV e um arquivo de conferência em Excel, considerando materiais e aplicadores existentes.

    Args:
        caminho_arquivo: Caminho do arquivo de cadastro de entrada utilizado como base para geração dos dados.
        caminho_output: Diretório de saída onde todos os arquivos gerados serão salvos.
        nivel: Nível ou versão de processamento utilizado para parametrizar a geração dos arquivos.
        df_wire_cao: DataFrame opcional contendo cadastros de cabos já existentes no CAO.
        df_term_cao: DataFrame opcional contendo cadastros de terminais já existentes no CAO.
        df_selos_cao: DataFrame opcional contendo cadastros de selos já existentes no CAO.
        df_app_cao: DataFrame opcional contendo cadastros de mini aplicadores já existentes.
    """
    
    if df_wire_cao is None or not df_wire_cao:
        df_wire_cao = pd.DataFrame()
    else:
        df_wire_cao = pd.read_csv(df_wire_cao, sep=";",dtype=str)
    if df_term_cao is None or not df_term_cao:
        df_term_cao = pd.DataFrame()
    else:
        df_term_cao = pd.read_csv(df_term_cao, sep=";",dtype=str)
    if df_selos_cao is None or not df_selos_cao:
        df_selos_cao = pd.DataFrame()
    else:
        df_selos_cao = pd.read_csv(df_selos_cao, sep=";",dtype=str)
        
    if df_app_cao is None or not df_app_cao:
        df_app_cao = pd.DataFrame()
    else:
        df_app_cao = pd.read_excel(df_app_cao,dtype=str)
    
    
    nivel = int(nivel)

    # caminho_arquivo = r"C:\Users\tgodoi01\Downloads\teste\1_CAO_CADASTRO_Renault_Alterar - OK.xlsx"
    # caminho_output = r"C:\Users\tgodoi01\Downloads\teste"

    # os.system("cls")

    # logo("Cadastro CAO")

    # logger.info("Entre com o caminho do arquivo de cadastro")
    # caminho_arquivo = input("[>] ").replace("/", "\\").replace('"', "")

    # logger.info("Entre com o caminho do diretório, onde quer salvar os arquivos")
    # caminho_output = input("[>] ")

    # while True:
    #     try:
    #         logger.info("Entre com o nível")
    #         nivel = int(input("[>] "))
    #         break  # Sai do loop se a conversão funcionar
    #     except ValueError:
    #         logger.error("Entrada inválida. Por favor, digite um número inteiro.")

    # while True:
    #     logger.info("Deseja carregar os arquivos de CAO? (S/N): ")
    #     resposta = input("[>] ")
    #     if resposta_positiva(resposta):
    #         df_wire_cao = ler_csv_validado("Cabos", "WireKey")
    #         df_term_cao = ler_csv_validado("Terminais", "TerminalKey")
    #         df_selos_cao = ler_csv_validado("Selos", "SealKey")
    #         break

    #     elif resposta.strip().lower() in {"n", "nao", "não", "no"}:
    #         logger.info("Leitura de arquivos CAO ignorada pelo usuário.")
    #         df_wire_cao = pd.DataFrame()
    #         df_term_cao = pd.DataFrame()
    #         df_selos_cao = pd.DataFrame()
    #         break
    #     else:
    #         logger.error("Resposta inválida. Digite S ou N (sim/não).")

    # while True:
    #     logger.info("Deseja carregar os arquivos de mini aplicadores? (S/N): ")
    #     resposta = input("[>] ")
    #     if resposta_positiva(resposta):
    #         df_app_cao = ler_excel_validado("Mini Aplicadores", "Terminal")
    #         break

    #     elif resposta.strip().lower() in {"n", "nao", "não", "no"}:
    #         logger.info("Leitura de arquivos ignorada pelo usuário.")
    #         df_app_cao = pd.DataFrame()
    #         break
    #     else:
    #         logger.error("Resposta inválida. Digite S ou N (sim/não).")

    # logger.info("Processando...")

    df_cadastro = criar_base(caminho_arquivo, number=nivel)

    df_completo = {
        "1. Cadastro Conferência": df_cadastro.dropna(how="all", axis=1)
    }
    df_completo["2. Cadastro Completo"] = df_cadastro

    df_cadastro.to_csv(
        f"{caminho_output}/2. Arquivo_de_cadastro_no_master_data.csv",
        sep=";",
        index=False,
    )

    agora = datetime.now()
    formato_personalizado = agora.strftime("%Y%m%d_%H%M%S")

    arq_ordens = gerar_ordens(caminho_arquivo, df_cadastro, number=nivel)

    if len(arq_ordens) > 0:
        arq_ordens.to_csv(
            f"{caminho_output}/10. Ordens_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )

    df_wire, df_terminais, df_selos = cadastro_componentes(df_cadastro)

    df_wire, df_terminais, df_selos = compara_material(
        df_wire, df_wire_cao, df_terminais, df_term_cao, df_selos, df_selos_cao
    )

    if len(df_wire) > 0:
        df_wire.to_csv(
            f"{caminho_output}/3. Cadastro_Cabos_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )
    if len(df_terminais) > 0:
        df_terminais.to_csv(
            f"{caminho_output}/4. Cadastro_Terminais_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )
    if len(df_selos) > 0:
        df_selos.to_csv(
            f"{caminho_output}/5. Cadastro_Selos_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )

    df_spot = gerar_sport_tape(df_cadastro, number=nivel)
    if len(df_spot) > 0:
        df_spot.to_csv(
            f"{caminho_output}/6. LeadSets_SpotTape_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )

    df_Tools_Terminals = gerar_Tools_Terminals(
        df_cadastro, lista_de_aplicadores=df_app_cao
    )
    if len(df_Tools_Terminals) > 0:
        df_Tools_Terminals.to_csv(
            f"{caminho_output}/7. ToolsXTerminals_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )

    if len(df_app_cao) > 0:
        df_Terminal_Applicators = gerar_Terminal_Applicators(
            dados=df_Tools_Terminals, lista_de_aplicadores=df_app_cao
        )
        if len(df_Terminal_Applicators) > 0:
            df_Terminal_Applicators.to_csv(
                f"{caminho_output}/8. Terminal_Applicators_{formato_personalizado}.csv",
                sep=";",
                index=False,
            )

    df_CrimpSTD = gerar_CrimpSTD(df_cadastro)
    if len(df_CrimpSTD) > 0:
        df_CrimpSTD.to_csv(
            f"{caminho_output}/9. CrimpSTD_{formato_personalizado}.csv",
            sep=";",
            index=False,
        )

    while True:
        caminho_arquivo = f"{caminho_output}/Arquivo de Conferencia.xlsx"
        try:
            salvar_dict_df_em_excel(df_completo, caminho_arquivo)
            break  # Sai do loop se a função for bem-sucedida
        except PermissionError:
            logger.error(
                f"Permissão negada ao tentar salvar o arquivo:\n  >> '{caminho_arquivo}'."
                f"\n  **Verifique se o arquivo está aberto ou se você tem permissão de escrita.**"
            )
            input("  Feche o arquivo e aperte [ENTER] para tentar novamente...")
        except Exception as e:
            logger.error(f"Ocorreu um erro inesperado ao salvar o arquivo: {e}")
            break  # Sai do loop se for um erro que não seja PermissionError (opcional)

    logger.info("Arquivo salvo com sucesso.")

