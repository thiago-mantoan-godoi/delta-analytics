import os

import pandas as pd
import warnings
import re
import math
import numpy as np
from collections import Counter
import platform
import psutil
import socket
import uuid
import win32com.client
import subprocess
import time

warnings.simplefilter("ignore")

def reordenar_colunas(df, colunas_prioritarias):
    # Garante que as colunas prioritárias existem no DataFrame
    colunas_prioritarias = [col for col in colunas_prioritarias if col in df.columns]

    # Pega as colunas restantes (que não estão nas prioritárias)
    colunas_restantes = [col for col in df.columns if col not in colunas_prioritarias]

    # Reordena o DataFrame
    return df[colunas_prioritarias + colunas_restantes]

def converter_sap(df:pd.DataFrame=None) -> pd.DataFrame:
    # if caminho_path.lower().endswith(".xlsx"):
    #     df = pd.read_excel(caminho_path,dtype=str)

    cols = ['INNER_1', 'INNER_2', 'INNER_3', 'INNER_4', 'INNER_5', 'INNER_6', 'INNER_7', 'INNER_8', 'INNER_9']

    dict_processos = {}
    dict_comp = {}
    for c in cols:
        for j in df.index:
            ckt = df.loc[j, c]
            process = df.loc[j, 'CIRCUIT']
            if pd.notna(ckt):
                ucs = df.loc[j,'Internal Family']
                dict_processos[f"{ucs}{ckt}"] = f"{ucs}{process}"
                dict_comp[f"{ucs}{process}"] = df.loc[j, 'LENGTH']
                
    for i in df.index:
        ckt = df.loc[i, 'CIRCUIT']
        ucs = df.loc[i,'Internal Family']
        
        leadset = dict_processos.get(f"{ucs}{ckt}", f"{ucs}{ckt}")
        
        df.loc[i,'Leadset']= leadset
        df.loc[i,'LENGTH_TW']= dict_comp.get(leadset, None)

    df_temp = df[df[cols].notna().any(axis=1)]
    df = df[df[cols].isna().all(axis=1)].reset_index(drop=True)
    
    return df, df_temp

def gerar_comunizacao_arquivo_sap(dados:str = None, ordem: list[str]=None) -> pd.DataFrame:
    # with open(path_dados, 'rb') as f:
    #     dados = pd.read_excel(f)
   
    dtype_cat = pd.api.types.CategoricalDtype(categories=ordem, ordered=True)

    dados["Internal Family Categoria"] = (
        dados["Internal Family"]
        .astype(dtype_cat)
    )

    # Ordena mantendo valores fora da lista no final
    dados = dados.sort_values(["Internal Family Categoria", "Internal Family"])\
        .drop(columns="Internal Family Categoria").reset_index(drop=True)
    
    

    colunas = ['WIRE_TUBE_SPLICE', 'LENGTH', 'LENGTH_TW','SECTIONN',
            'TERM_A', 'STRIP_A', 'SEAL_A',
            'TERM_B', 'STRIP_B', 'SEAL_B','RMKS_A','RMKS_B']

    # Converter tudo para string antes de usar Counter
    dados['linha_key'] = dados[colunas].apply(
        lambda row: tuple(sorted((str(k), v) for k, v in Counter(row).items())),
        axis=1
    )

    dados['Status'] = None

    # Agrupar por essa chave
    grupos = dados.groupby('linha_key').indices
    id = 1
    # Mostrar grupos com mais de uma linha
    for key, indices in grupos.items():
        if len(indices) > 1:
            dados.loc[indices,'Status'] = id
            id+=1

    dados.drop(columns=['linha_key'],inplace=True)

    dados['CIRC_MASTER'] = None
    dados['CIRC_MASTER'] = None
    dados['CIRC_COMUNS'] = None


    id_list = dados['Status'].dropna().unique().tolist()
    for i in id_list:
        index_number = dados[dados['Status']==i].index[0]
        index_list = dados[dados['Status']==i].index

        dados.loc[index_list,'CIRC_MASTER'] = dados.loc[index_number,'Internal Family']+dados.loc[index_number,'CIRCUIT']
        dados.loc[index_number,'CIRC_COMUNS'] = dados.loc[index_number,'Internal Family']+dados.loc[index_number,'CIRCUIT']
        for j in index_list[1:]:
            dados.loc[j,'CIRC_COMUNS'] = dados.loc[j,'Internal Family']+dados.loc[j,'CIRCUIT']

        item_master = dados[dados['Status']==i]['CIRC_MASTER'].unique()
        item_comum = dados[dados['Status']==i]['CIRC_COMUNS'].unique()

        if len(item_master)==1  and len(item_comum)==1:
            
            dados.loc[index_list,'CIRC_MASTER'] = dados.loc[index_number,'Internal Family']+dados.loc[index_number,'CIRCUIT']
            dados.loc[index_number,'CIRC_COMUNS'] = dados.loc[index_number,'Internal Family']+dados.loc[index_number,'CIRCUIT']
            for j in index_list[1:]:
                dados.loc[j,'CIRC_COMUNS'] = dados.loc[j,'Internal Family']+dados.loc[j,'CIRCUIT']

        sectionn = dados[dados['Status']==i]['SECTIONN'].unique()

        if sectionn == '0 0' or sectionn == '0 1' or sectionn == '0.14'  or sectionn == '0.18':
            dados.loc[index_list,'CIRC_MASTER'] = None
            dados.loc[index_list,'CIRC_COMUNS'] = None
            dados.loc[index_list,'Status'] = None

    dados.drop(columns=['Status'], inplace=True)

    return dados

def comparar_tabela_sap(df0:pd.DataFrame=None, df2:pd.DataFrame=None) -> pd.DataFrame:

    # with open(path_dados, "rb") as f:
    #     df2 = pd.read_csv(f, sep=";")

    df1 = df0.copy()
    df1 = df1.dropna(subset=['CIRC_MASTER']).reset_index(drop=True)
    lista_master = df1['CIRC_MASTER'].unique().tolist()
    for i in lista_master:
        df_prov1 = sorted(df1[df1['CIRC_MASTER']==i]['CIRC_COMUNS'].unique().tolist())
        df_prov2 = sorted(df2[df2['CIRC_COMUNS'].isin(df_prov1)]['CIRC_COMUNS'].unique().tolist())

        if df_prov1!=df_prov2:
            if len(df_prov1) > 0 and len(df_prov2) == 0 :
                lista_index1 = df0[df0['CIRC_COMUNS'].isin(df_prov1)].index 
                df0.loc[lista_index1,'STATUS']='Adicionar'
            elif len(df_prov1) == 0 and len(df_prov2) > 0 :
                lista_index2 = df0[df0['CIRC_COMUNS'].isin(df_prov2)].index 
                df0.loc[lista_index2,'STATUS']='Remover'
        else:
            if len(df_prov1)==len(df_prov2):
                lista_index = df0[df0['CIRC_COMUNS'].isin(df_prov1)].index
                df0.loc[lista_index,'STATUS']='Já esta comunizado.'

    colunas = ['Leadset','TYPE','WERKS','External Family','FILE_LINE','STATUS_REGISTRO',
               'Internal Family','CIRC_MASTER','CIRC_COMUNS','STATUS','MULTICORE','CIRCUIT','WIRE_TUBE_SPLICE',
               'LENGTH','LENGTH_TW']
    df0 = reordenar_colunas(df0, colunas)

    return df0

def arrumar_leadset(dados:pd.DataFrame=None) -> pd.DataFrame:

    def extrair_codigo(texto):
        return re.sub(r'^[A-Z]+\d+_?', '', texto)

    def extrair_codigo_SHIELDE(texto):
        return re.sub(r'^[A-Z]+\d+', '', texto)


    for i in dados.index:
        leadset = dados.loc[i, 'Leadset']
        
        RMKS_A = str(dados.loc[i, 'RMKS_A']).split(',')
        RMKS_A = [x for x in RMKS_A if x]
        
        if len(RMKS_A)==1 and str(leadset).startswith("B") and 'W' in leadset:
            dados.loc[i, 'MULTICORE'] = leadset[3:]
            
        elif len(RMKS_A)>1  and str(leadset).startswith("B") and 'W' in leadset:
            dados.loc[i, 'MULTICORE'] = leadset[3:]
            
            
            
        # Criar multicore
        if ('TW' in leadset or 'MC' in leadset) and (str(leadset).startswith("G") or str(leadset).startswith("S")):
            dados.loc[i, 'MULTICORE'] =  extrair_codigo(leadset)
            
        if ('MC' in leadset or 'TW' in leadset) and str(leadset).startswith("V"):
            dados.loc[i, 'MULTICORE'] =  extrair_codigo(leadset)
            
        if 'SHIELDE' in leadset and str(leadset).startswith("X"):
            dados.loc[i, 'MULTICORE'] =  extrair_codigo_SHIELDE(leadset)
        
        # Ajustar
        if str(leadset).startswith("V") and "TW" in leadset:
            dados.loc[i, 'Leadset'] = f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            
        if (str(leadset).startswith("G") or str(leadset).startswith("S")) and float(dados.loc[i, 'SECTIONN'].replace(' ',''))>=1.5:
            dados.loc[i, 'Leadset'] = f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            
        if str(leadset).startswith("X") and "TW" in leadset:
            dados.loc[i, 'Leadset'] = f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            
        if len(RMKS_A)==1 and str(leadset).startswith("B") and 'W' in leadset:
            dados.loc[i, 'Leadset'] = f"{dados.loc[i, 'Internal Family']}{dados.loc[i, 'CIRCUIT']}"
            
    colunas = ['Leadset','TYPE','WERKS','External Family','FILE_LINE','STATUS_REGISTRO',
               'Internal Family','CIRC_MASTER','CIRC_COMUNS','STATUS','MULTICORE','CIRCUIT','WIRE_TUBE_SPLICE',
               'LENGTH','LENGTH_TW']
    dados = reordenar_colunas(dados, colunas)
    return dados

def converte_arquivo_sap(dados:pd.DataFrame=None, dados_cmz: pd.DataFrame=None) -> pd.DataFrame:
    
    df, _ = converter_sap(df=dados)

    df_prov_1 = arrumar_leadset(df)

    df_prov_2 = gerar_comunizacao_arquivo_sap(df_prov_1,ordem=sorted(df_prov_1['Internal Family'].unique().tolist()))

    dados = comparar_tabela_sap(df_prov_2, dados_cmz) 
    
    return dados

def adicionar_sequencia(df:pd.DataFrame=None)-> pd.DataFrame:
    
    caminho_cabos = os.path.join(os.getcwd(), "data", 'Lista_de_cabos.json')
    df_cabos = pd.read_json(caminho_cabos)
    dict_cabos = dict(zip(df_cabos['Part Number'],df_cabos['Part Classification']))

    caminho_mapa = os.path.join(os.getcwd(), "data", 'Lista_de_mapa_corte.json')
    df_mapa = pd.read_json(caminho_mapa).drop_duplicates().reset_index(drop=True)
    dict_mapa = dict(zip(df_mapa['Leadset'],df_mapa['Alocação']))


    # Terminais
    Terminais = (
        df[['TERM_A','TERM_B']]
        .stack()
        .value_counts()
        .index
        .to_series()
        .reset_index(drop=True)
        .reset_index()
    )

    Terminais.columns = ['ID','Terminais']

    # Selos
    Selos = (
        df[['SEAL_A','SEAL_B']]
        .stack()
        .value_counts()
        .index
        .to_series()
        .reset_index(drop=True)
        .reset_index()
    )

    Selos.columns = ['ID','Selos']

    # Adiciona categorias especiais
    Terminais.loc[len(Terminais)] = [1000, 'Sem terminal']
    Selos.loc[len(Selos)] = [1000, 'Sem selo']

    # Dicionários
    dict_term = dict(zip(Terminais['Terminais'], Terminais['ID']))
    dict_selos = dict(zip(Selos['Selos'], Selos['ID']))

    dict_id_term = dict(zip(Terminais['ID'], Terminais['Terminais']))
    dict_id_selos = dict(zip(Selos['ID'], Selos['Selos']))

    df['Term_A_Temp'] = df['TERM_A'].where(df['Processo_A'].eq('Corte')).fillna('Sem terminal')
    df['Selo_A_Temp'] = df['SEAL_A'].where(df['Processo_A'].eq('Corte')).fillna('Sem selo')

    # B
    df['Term_B_Temp'] = df['TERM_B'].where(df['Processo_B'].eq('Corte')).fillna('Sem terminal')
    df['Selo_B_Temp'] = df['SEAL_B'].where(df['Processo_B'].eq('Corte')).fillna('Sem selo')

    # IDs
    df['TermA_ID'] = df['Term_A_Temp'].map(dict_term)
    df['TermB_ID'] = df['Term_B_Temp'].map(dict_term)
    df['SEALA_ID'] = df['Selo_A_Temp'].map(dict_selos)
    df['SEALB_ID'] = df['Selo_B_Temp'].map(dict_selos)

    for i in df.index:
        termA_ID = df.loc[i, 'TermA_ID']
        termB_ID = df.loc[i, 'TermB_ID']
        seloA_ID = df.loc[i, 'SEALA_ID']
        seloB_ID = df.loc[i, 'SEALB_ID']
        if termA_ID < termB_ID:
            df.loc[i, 'TermA_ID'] = termB_ID
            df.loc[i, 'TermB_ID'] = termA_ID
            df.loc[i, 'SEALA_ID'] = seloB_ID
            df.loc[i, 'SEALB_ID'] = seloA_ID
            
    df['TermA_uso'] = df['TermA_ID'].map(dict_id_term)
    df['TermB_uso'] = df['TermB_ID'].map(dict_id_term)
    df['SEALA_uso'] = df['SEALA_ID'].map(dict_id_selos)
    df['SEALB_uso'] = df['SEALB_ID'].map(dict_id_selos)
        
    df = df.sort_values(by=['SEALA_ID', 'SEALB_ID','TermA_ID', 'TermB_ID']).reset_index(drop=True)

    df.drop(columns=['TermA_ID', 'TermB_ID', 'SEALA_ID', 'SEALB_ID','Term_A_Temp','Term_B_Temp','Selo_A_Temp','Selo_B_Temp'], inplace=True)

    df.loc[(df['SEALA_uso']!='Sem selo') & (df['SEALB_uso']!='Sem selo'), 'Processo'] = 'SS'
    df.loc[(df['SEALA_uso']=='Sem selo') & (df['SEALB_uso']=='Sem selo'), 'Processo'] = 'TT'
    df.loc[(df['SEALA_uso']=='Sem selo') ^ (df['SEALB_uso']=='Sem selo'), 'Processo'] = 'TS'
    
    df['Seq.'] = range(len(df))
    
    
    df['Tipo de cabo'] = df['WIRE_TUBE_SPLICE'].map(dict_cabos)
    mask1 = df['Tipo de cabo'].str.contains('Shielded|Multicore', na=False)
    df.loc[mask1, 'Processo'] = 'MS'
    
    
    s = pd.to_numeric(df['SECTIONN'].str.replace(' ', ''), errors='coerce')
    df.loc[s >= 10, 'Processo'] = 'MS'
    df.loc[df['COLOR1'].isin(['BA', 'SC', 'NT']), 'Processo'] = 'MS'

    mask2 = df['Tipo de cabo'].str.contains('Subassembly|Coaxial', na=False)
    df.loc[mask2, 'Processo'] = 'Subassembly'
    
    
    df.drop(columns=['Tipo de cabo'], inplace=True)
    
    
    # Substitui vírgula por ponto
    df['LENGTH_clean'] = df['LENGTH'].astype(str).str.replace(',', '.')

    # Converte para float, valores inválidos viram NaN
    df['LENGTH_float'] = pd.to_numeric(df['LENGTH_clean'], errors='coerce')

    # Calcula LClass usando ceil
    df['LClass'] = np.ceil(df['LENGTH_float'] / 500)

    # Opcional: remover colunas temporárias
    df.drop(columns=['LENGTH_clean','LENGTH_float'], inplace=True)
     
    df['Alocações'] = df['Leadset'].map(dict_mapa)

    df['Bundle size'] = np.select(
        [
            pd.to_numeric(df['SECTIONN'].astype(str).str.strip().str.replace(' ', '.').str.replace(',', '.'), errors='coerce') < 0.75,
            (pd.to_numeric(df['SECTIONN'].astype(str).str.strip().str.replace(' ', '.').str.replace(',', '.'), errors='coerce') > 2) |
            (pd.to_numeric(df['LENGTH'].astype(str).str.replace(',', '.'), errors='coerce') > 3500)
        ],
        [100, 25],
        default=50
    )
    
    mask3 = df[(df['Leadset'].str.contains('TW', na=False)) & (df['Processo'] != 'MS')].index
    df.loc[mask3, 'Bundle size'] = 40
    
    
    mask4 = df[df['Processo']== 'MS'].index
    df.loc[mask4, 'Bundle size'] = 25
    
    cols = ['Leadset','Processo_A','Processo_B','TermA_uso', 'TermB_uso', 'SEALA_uso','SEALB_uso','Seq.','Alocações','Processo','LClass','Bundle size']
    df = reordenar_colunas(df,colunas_prioritarias=cols)
    
    return df 

def definir_processos(dados:pd.DataFrame=None)-> pd.DataFrame:
    
    # CABOS
    caminho_cabos = os.path.join(os.getcwd(), "data", 'Lista_de_cabos.json')
    df_cabos = pd.read_json(caminho_cabos)
    dict_cabos = dict(zip(df_cabos['Part Number'],df_cabos['Part Classification']))
    

    
    # TERMINAIS
    caminho_terminais = os.path.join(os.getcwd(), "data", 'Lista_de_terminais.json')
    df_terminais = pd.read_json(caminho_terminais)
    dict_Technology = dict(zip(df_terminais['Part Number'],df_terminais['Connection Technology']))
    dict_Delivery = dict(zip(df_terminais['Part Number'],df_terminais['Feed Type/Delivery Form']))
    dict_Style = dict(zip(df_terminais['Part Number'],df_terminais['Terminal Style (Male or Female only)']))
    lista_temp = df_terminais[pd.to_numeric(df_terminais['Min Wire Size (mm^2)'], errors='coerce') >= 6]['Part Number']
    
    lista_resticoes = ['Clamp', 'Solder', 'Tube Crimp', 'Weld','Loose Piece','Ultrasonic Weld']
    
    
    dados['Processo_A'] = None
    dados['Processo_B'] = None
    
    for i in dados.index:
        termA = None if pd.isna(dados.loc[i, 'TERM_A']) else dados.loc[i, 'TERM_A']
        termB = None if pd.isna(dados.loc[i, 'TERM_B']) else dados.loc[i, 'TERM_B']
        cabo = None if pd.isna(dados.loc[i, 'WIRE_TUBE_SPLICE']) else dados.loc[i, 'WIRE_TUBE_SPLICE']
        
        lista_rest = []
        if termA:
            lista_rest.append(dict_Technology.get(termA))
            lista_rest.append(dict_Delivery.get(termA))
            lista_rest.append(dict_Style.get(termA))
            lista_rest = [i for i in lista_rest if pd.notna(i) and i != '']
            if len(lista_rest)>0:
                if any(i in lista_rest for i in lista_resticoes):
                    if 'Weld' in lista_rest:
                        dados.loc[i, 'Processo_A'] = 'Weld'
                    else:
                        dados.loc[i, 'Processo_A'] = 'Prensa'
                        

        lista_rest = []
        if termB:
            lista_rest.append(dict_Technology.get(termB))
            lista_rest.append(dict_Delivery.get(termB))
            lista_rest.append(dict_Style.get(termB))
            lista_rest = [i for i in lista_rest if pd.notna(i) and i != '']
            if len(lista_rest)>0:
                if any(i in lista_rest for i in lista_resticoes):
                    if 'Weld' in lista_rest:
                        dados.loc[i, 'Processo_B'] = 'Weld'
                    else:
                        dados.loc[i, 'Processo_B'] = 'Prensa'
                        
                        
        if cabo in dict_cabos and ("Shielded" in str(dict_cabos[cabo]) or "Multicore" in str(dict_cabos[cabo]) or "Subassembly" in str(dict_cabos[cabo]) or "Coaxial" in str(dict_cabos[cabo])):
            if pd.isna(dados.loc[i, 'Processo_A']):
                dados.loc[i, 'Processo_A']='Prensa'
            if pd.isna(dados.loc[i, 'Processo_B']):
                dados.loc[i, 'Processo_B']='Prensa'
                
        if termA in lista_temp:
            if pd.isna(dados.loc[i, 'Processo_A']):
                dados.loc[i, 'Processo_A']='Prensa'
                
        if termB in lista_temp:
            if pd.isna(dados.loc[i, 'Processo_B']):
                dados.loc[i, 'Processo_B']='Prensa'
    
    dados[['Processo_A','Processo_B']] = dados[['Processo_A','Processo_B']].fillna('Corte')
    
    cols = ['Leadset','Processo_A','Processo_B']
    
    dados = reordenar_colunas(dados, cols)

    return dados 

def add_volumes(dados: pd.DataFrame = None) -> pd.DataFrame:
    
    caminho_volumes = os.path.join(os.getcwd(), "data", "Lista_de_master_kanban.json")
    df_volumes = pd.read_json(caminho_volumes)

    df_volumes["Volumes"] = (
        pd.to_numeric(
            df_volumes["Total week"].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        ) / 5
    )
    dict_volumes = df_volumes.set_index("Derivativos CARGA")["Volumes"].to_dict()

    num = dados.columns.get_loc("PN_DERIVATIVO_1")
    cols = dados.columns[num:].tolist()

    totais = []

    for i in dados.index:
        derivados = dados.loc[i, cols].dropna().values.tolist()
        derivados = [str(d).replace(' ','') for d in derivados]
        total = sum(dict_volumes.get(d, 0) for d in derivados)
        totais.append(total)
        
    dados["Volumes"] = totais
    
    cols = ['Leadset', 'Processo_A', 'Processo_B', 'TermA_uso', 'TermB_uso', 'SEALA_uso', 'SEALB_uso', 'Seq.', 'Processo','Volumes']
    
    dados = reordenar_colunas(df=dados, colunas_prioritarias=cols)
    
    dados.loc[dados.duplicated('Leadset'), 'Volumes'] = 0


    def normalize_side(row, cols):
        vals = row[cols].dropna().values
        return tuple(sorted(vals))

    def build_key(row):
        ladoA = normalize_side(row, ['TERM_A','STRIP_A','SEAL_A'])
        ladoB = normalize_side(row, ['TERM_B','STRIP_B','SEAL_B'])
        
        # tornar A-B equivalente a B-A
        lados = tuple(sorted([ladoA, ladoB]))
        
        return (
            row['WIRE_TUBE_SPLICE'],
            row['LENGTH'],
            row['LENGTH_TW'],
            lados
        )

    dados['group_key'] = dados.apply(build_key, axis=1)

    dados['Comunizados'] = dados.groupby('group_key')['Leadset'].transform('first')


    cols = ['Leadset','CIRC_MASTER', 'CIRC_COMUNS','Volumes']

    dados["Vol/dia"] = dados.groupby('Comunizados')["Volumes"].transform("sum")
    dados.loc[dados.duplicated("Comunizados"), "Vol/dia"] = 0

    dados.drop(columns=['group_key'],inplace=True)

    cols = ['Leadset', 'Processo_A', 'Processo_B', 'TermA_uso', 'TermB_uso', 'SEALA_uso', 'SEALB_uso', 'Seq.', 'Processo','Comunizados','Volumes',"Vol/dia"]

    dados = reordenar_colunas(df=dados, colunas_prioritarias=cols)

    return dados

def top_processos_memoria():
    processos = []

    for proc in psutil.process_iter(['name', 'memory_info']):
        try:
            #mem = proc.info['memory_info'].rss / (1024 * 1024)  # MB
            mem = proc.info['memory_info'].rss / (1024 * 1024 * 1024)  # GB
            nome = proc.info['name'] or "Desconhecido"
            processos.append((nome, mem))

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    # ordena e pega top 5
    top5 = sorted(processos, key=lambda x: x[1], reverse=True)[:5]

    return top5

def testar_latencia():
    try:
        # Tenta pingar o DNS do Google 1 vez
        saida = subprocess.check_output("ping -n 1 8.8.8.8", shell=True).decode('cp850')
        
        # Usa Expressão Regular para buscar qualquer número seguido de 'ms'
        # Isso funciona tanto em Windows PT-BR quanto EN-US
        ms = re.findall(r"(\d+)ms", saida)
        
        if ms:
            return f"{ms[-1]}ms" # Retorna o último valor de ms encontrado (geralmente a média)
        return "Ping realizado, mas tempo não identificado."
    except:
        return "Falha na conexão (Host inacessível)"

def obter_info_colaborador():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        curr_user = ns.CurrentUser.AddressEntry
        ex_user = curr_user.GetExchangeUser()

        if not ex_user:
            return _fallback("Conta não é Exchange")

        # gestor pode não existir
        try:
            gestor = ex_user.GetManager().Name
        except:
            gestor = "Não atribuído"

        return {
            "Nome": ex_user.Name or "-",
            "E-mail": ex_user.PrimarySmtpAddress or "-",
            "Cargo": ex_user.JobTitle or "-",
            "Departamento": ex_user.Department or "-",
            "Empresa": ex_user.CompanyName or "-",
            "Escritorio": ex_user.OfficeLocation or "-",
            "Telefone": ex_user.BusinessTelephoneNumber or "-",
            "Celular": ex_user.MobileTelephoneNumber or "-",
            "Gestor_Direto": gestor,
            "Cidade": ex_user.City or "-"
        }

    except Exception:
        return _fallback("Outlook não disponível")

def _fallback(motivo):
    # retorno padrão amigável
    return {
        "Nome": "Não disponível",
        "E-mail": "Não disponível",
        "Cargo": "Não disponível",
        "Departamento": "Não disponível",
        "Empresa": "Não disponível",
        "Escritorio": "Não disponível",
        "Telefone": "Não disponível",
        "Celular": "Não disponível",
        "Gestor_Direto": motivo,
        "Cidade": "Não disponível"
    }

def obter_info_maquina():
    try:
        # Informações básicas do SO
        info = {
            "Nome_Computador": socket.gethostname(),
            "Sistema_Operacional": f"{platform.system()} {platform.release()}",
            "Versao_SO": platform.version(),
            "Arquitetura": platform.machine(),
            "Processador": platform.processor(),
            "Nucleos_Fisicos": psutil.cpu_count(logical=False),
            "Memoria_RAM_GB": round(psutil.virtual_memory().total / (1024**3), 2),
            "Usuario_Logado": os.getlogin(),
            "Endereco_IP": socket.gethostbyname(socket.gethostname()),
            "MAC_Address": ':'.join(['{:02x}'.format((uuid.getnode() >> i) & 0xff) for i in range(0,8*6,8)][::-1])
        }
        return info
    except Exception as e:
        return f"Erro ao obter dados da máquina: {e}"
    
