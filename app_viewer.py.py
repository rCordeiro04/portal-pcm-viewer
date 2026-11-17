import calendar
import io
import os
import shutil
from datetime import date, datetime

import altair as alt
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURAÇÃO DA PÁGINA & TEMA GLOBAL
# ==========================================
st.set_page_config(
    page_title="Portal PCM - Manutenção",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

def tema_altair():
    return {
        "config": {
            "view": {"strokeWidth": 0},
            "background": "transparent",
            "axis": {
                "domainColor": "#cbd5e1",
                "gridColor": "#f8fafc",
                "labelColor": "#64748b",
                "labelFontWeight": 600,
                "titleColor": "#0f172a",
                "titleFontWeight": 800,
                "tickColor": "#cbd5e1"
            },
            "legend": {
                "labelColor": "#475569",
                "titleColor": "#0f172a",
                "titleFontWeight": 800
            }
        }
    }
alt.themes.register("tema_pcm", tema_altair)
alt.themes.enable("tema_pcm")

CSS_TELA_LIMPA = "<style>header[data-testid='stHeader'] {display: none !important;} .block-container {padding-top: 1rem !important; max-width: 100% !important;}</style>"

# ==========================================
# CONSTANTES, ARQUIVOS & PALETAS DE CORES
# ==========================================
PASTA_BD = "banco_dados"
os.makedirs(PASTA_BD, exist_ok=True)

ARQUIVO_CORREIAS = os.path.join(PASTA_BD, "CORREIAS.xlsx")
ARQUIVO_CORRETIVAS = os.path.join(PASTA_BD, "CORRETIVAS.xlsx")
ARQUIVO_TRABALHOS_FUSOS = os.path.join(PASTA_BD, "TRABALHOS FUSOS.xlsx")
ARQUIVO_PREVENTIVAS = os.path.join(PASTA_BD, "PREVENTIVAS.xlsx")
ARQUIVO_ESTOQUE = os.path.join(PASTA_BD, "ESTOQUE.xlsx")
ARQUIVO_TERCEIROS = os.path.join(PASTA_BD, "TERCEIROS.xlsx")
ARQUIVO_CHECKLIST_ELETRICA = os.path.join(PASTA_BD, "CHECKLIST ELETRICA.xlsx")
ARQUIVO_CHECKLIST_MECANICA = os.path.join(PASTA_BD, "CHECKLIST MECANICA.xlsx")

COLUNAS_FUSOS = ["Ano", "Mes", "Dia", "Setor", "Maquina_TAG", "Quantidade_Quebras", "Tipo_Fuso"]
COLUNAS_CORREIAS = [
    "Setor", "Maquina_TAG", "Tipo_Correia_1", "Data_Instalacao_1", "Tipo_Correia_2", "Data_Instalacao_2", 
    "Ultima_Preventiva", "Proxima_Preventiva", "Prev_Tipo", "Prev_Validade_Meses", 
    "Prev_Participantes", "Prev_Atividades", "Prev_Pecas"
]
COLUNAS_ESTOQUE = ["NOME", "DESCRIÇÃO", "QUANTIDADE", "ESTADO", "CATEGORIAS", "ESTOQUE MINIMO"]
COLUNAS_MANUTENCAO = ["NOME", "OCORRENCIA", "QUANTIDADE", "DATA"]

OPCOES_TIPO_FUSO = ["FAG", "TEP", "M4BA", "MENEGATTO", "M4ZD", "USL"]
LISTA_MESES_PUROS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MAPA_MES_ABREV = {"Janeiro": "JAN", "Fevereiro": "FEV", "Março": "MAR", "Abril": "ABR", "Maio": "MAI", "Junho": "JUN", "Julho": "JUL", "Agosto": "AGO", "Setembro": "SET", "Outubro": "OUT", "Novembro": "NOV", "Dezembro": "DEZ"}
ORDEM_MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

DICIONARIO_SETORES = {
    "Setor A": [f"L-{i:02d}" for i in range(1, 29)],
    "Setor B": ["L-29", "L-30", "L-31", "L-32", "L-33", "L-34", "L-35", "L-36", "L-37", "L-38", "L-39", "L-40", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46", "L-50", "L-51", "L-52", "L-53"],
    "Setor Látex": ["B-71", "B-72", "B-73", "B-74", "B-75", "B-76", "B-77", "B-78", "B-79", "B-80", "B-83", "B-84", "B-85", "B-86", "B-87", "B-88", "B-102", "B-103", "B-104"],
    "Setor Menegatto": ["B-47", "B-48", "B-49", "B-81", "B-82", "B-93", "B-94", "B-95", "B-96", "B-97", "B-98", "B-99", "B-100", "B-101", "B-107", "B-108"],
}

CORES_SETORES = {
    "Setor A": "#3b82f6", "Setor B": "#f59e0b", "Setor Látex": "#10b981", "Setor Menegatto": "#ef4444",
    "A": "#3b82f6", "B": "#f59e0b", "Látex": "#10b981", "Latex": "#10b981", "Menegatto": "#ef4444", "Men": "#ef4444"
}
CORES_FUSOS = {"FAG": "#3b82f6", "TEP": "#f59e0b", "M4BA": "#10b981", "MENEGATTO": "#ef4444", "M4ZD": "#8b5cf6", "USL": "#ec4899"}
CORES_TIPO = {"Corretiva": "#ef4444", "Preventiva": "#10b981", "Setup": "#f59e0b", "Planejada": "#3b82f6"}
PALETA_DINAMICA = ['#3b82f6', '#10b981', '#ef4444', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4', '#84cc16', '#64748b']

# ==========================================
# FUNÇÕES DE SUPORTE & HTML
# ==========================================
def formatar_modelo(val):
    if val is None or pd.isna(val): return ""
    v_str = str(val).strip()
    if v_str.lower() in ["", "nan", "none", "nat"]: return ""
    if "." in v_str:
        partes = v_str.split(".")
        p_dec = partes[1][:3].ljust(3, "0")
        return f"{partes[0]}.{p_dec}"
    return v_str

def html_lista_ponto(texto, cor_texto="#475569", font_weight="600", font_size="0.85rem"):
    if not texto or pd.isna(texto) or str(texto).lower() in ["nan", "none", ""]: return ""
    txt = str(texto).strip()
    if ";" in txt:
        itens = [x.strip() for x in txt.split(";") if x.strip()]
        return f"<ul style='margin:4px 0 4px 22px; padding:0; font-size:{font_size}; color:{cor_texto}; font-weight:{font_weight}; line-height:1.4;'>" + "".join([f"<li style='margin-bottom:3px;'>{i}</li>" for i in itens]) + "</ul>"
    return f"<div style='font-size:{font_size}; color:{cor_texto}; font-weight:{font_weight}; line-height:1.4; margin:4px 0;'>{txt}</div>"

def formatar_pills(texto):
    if not texto or pd.isna(texto) or str(texto).lower() in ["nan", "none", ""]: 
        return "<span style='background:#f8fafc; padding:3px 8px; border:1px solid #e2e8f0; border-radius:6px; color:#334155; font-size: 0.75rem;'>N/I</span>"
    txt = str(texto).strip()
    if ";" in txt:
        itens = [x.strip() for x in txt.split(";") if x.strip()]
        return "".join([f"<span style='background:#f8fafc; padding:3px 8px; border:1px solid #e2e8f0; border-radius:6px; color:#334155; margin-right:4px; display:inline-block; margin-top:4px; font-size: 0.75rem; font-weight: 700; box-shadow: 0 1px 2px rgba(0,0,0,0.02);'>🏷️ {i}</span>" for i in itens])
    return f"<span style='background:#f8fafc; padding:3px 8px; border:1px solid #e2e8f0; border-radius:6px; color:#334155; display:inline-block; margin-top:4px; font-size: 0.75rem; font-weight: 700; box-shadow: 0 1px 2px rgba(0,0,0,0.02);'>🏷️ {txt}</span>"

def obter_fuso_padrao(maq_tag, setor_nome):
    maq_str = str(maq_tag).strip().upper()
    MAQUINAS_FAG = [f"L-{i:02d}" for i in range(1, 29)] + ["L-52", "L-53", "B-47", "B-48", "B-49"]
    MAQUINAS_MENEGATTO = ["L-29", "L-30", "L-31", "L-35", "L-38", "L-50", "L-51", "L-41", "L-42", "L-43", "L-44", "L-45", "L-46"]
    MAQUINAS_USL = [f"B-{i}" for i in range(70, 80)]
    if maq_str in MAQUINAS_FAG: return "FAG"
    if maq_str in MAQUINAS_MENEGATTO: return "MENEGATTO"
    if maq_str in MAQUINAS_USL: return "USL"
    if setor_nome == "Setor A": return "FAG"
    elif setor_nome == "Setor B": return "TEP"
    elif setor_nome == "Setor Látex": return "M4BA"
    return "MENEGATTO"

def obter_maquinas_setor(setor_nome, df_f=None):
    base = set(DICIONARIO_SETORES.get(setor_nome, []))
    if df_f is not None and not df_f.empty: 
        base.update(df_f[df_f["Setor"] == setor_nome]["Maquina_TAG"].dropna().unique())
    return sorted(list(base))

def gerar_backup_seguro(caminho_arquivo):
    if os.path.exists(caminho_arquivo):
        os.makedirs("backups", exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arq = os.path.basename(caminho_arquivo)
        try: shutil.copy2(caminho_arquivo, os.path.join("backups", f"{ts}_{nome_arq}"))
        except: pass

@st.cache_data(show_spinner=False, ttl=5)
def carregar_mapa_tipos_fuso():
    mapa = {}
    if not os.path.exists(PASTA_BD): return mapa
    arquivos_na_pasta = os.listdir(PASTA_BD)
    arquivos_fusos = [f for f in arquivos_na_pasta if "FUSOS" in f.upper() and f.endswith(".xlsx")]
    
    for arq in arquivos_fusos:
        caminho_completo = os.path.join(PASTA_BD, arq)
        try:
            xls = pd.ExcelFile(caminho_completo)
            for nome_aba in xls.sheet_names:
                df_aba = pd.read_excel(xls, sheet_name=nome_aba)
                col_maq = None
                col_tipo = None
                for c in df_aba.columns:
                    c_str = str(c).upper().strip()
                    if "MAQ" in c_str or "TAG" in c_str or "MÁQ" in c_str: col_maq = c
                    elif "FUSO" in c_str or "TIPO" in c_str or "MODELO" in c_str: col_tipo = c
                if col_maq and col_tipo:
                    for _, linha in df_aba.iterrows():
                        tag_maquina = str(linha[col_maq]).strip().upper()
                        if not tag_maquina or tag_maquina.lower() in ["nan", "none", ""]: continue
                        if pd.notna(linha[col_tipo]):
                            val_fuso = str(linha[col_tipo]).strip().upper()
                            if val_fuso not in ["NAN", "NONE", ""]:
                                mapa[tag_maquina] = val_fuso
        except Exception:
            pass
    return mapa

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_corretivas():
    if os.path.exists(ARQUIVO_CORRETIVAS):
        try:
            df = pd.read_excel(ARQUIVO_CORRETIVAS)
            mapa_colunas = {"SETOR": "Setor", "MAQUINA": "Maquina", "DATA": "Data", "DESCRIÇÃO": "Descricao", "ELE/MEC": "Especialidade", "TIPO": "Tipo", "DURAÇÃO (Horas)": "Duracao"}
            df = df.rename(columns={k: v for k, v in mapa_colunas.items() if k in df.columns})
            if "Duracao" in df.columns: df["Duracao"] = pd.to_numeric(df["Duracao"], errors="coerce").fillna(0)
            if "Data" in df.columns: df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
            return df
        except Exception as e: pass
    return pd.DataFrame(columns=["Setor", "Maquina", "Data", "Descricao", "Especialidade", "Tipo", "Duracao"])

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_correias():
    if os.path.exists(ARQUIVO_CORREIAS):
        try: df_c = pd.read_excel(ARQUIVO_CORREIAS)
        except: df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
    else:
        df_c = pd.DataFrame(columns=COLUNAS_CORREIAS)
        df_c.to_excel(ARQUIVO_CORREIAS, index=False)

    for col in COLUNAS_CORREIAS:
        if col not in df_c.columns: df_c[col] = ""
        df_c[col] = df_c[col].astype(object)
    df_c["Tipo_Correia_1"] = df_c["Tipo_Correia_1"].apply(formatar_modelo)
    df_c["Data_Instalacao_1"] = df_c["Data_Instalacao_1"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    df_c["Tipo_Correia_2"] = df_c["Tipo_Correia_2"].apply(formatar_modelo)
    df_c["Data_Instalacao_2"] = df_c["Data_Instalacao_2"].astype(str).replace({"nan": "", "NaT": "", "None": ""})
    return df_c

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_fusos():
    arquivos_na_pasta = os.listdir(PASTA_BD)
    arquivos_fusos = [f for f in arquivos_na_pasta if "FUSOS" in f.upper() and f.endswith(".xlsx")]
    lista_registros = []
    for arq in arquivos_fusos:
        caminho_completo = os.path.join(PASTA_BD, arq)
        try:
            setor_inferido = "Setor A"
            if "B" in arq.upper(): setor_inferido = "Setor B"
            if "LATEX" in arq.upper() or "LÁTEX" in arq.upper(): setor_inferido = "Setor Látex"
            if "MEN" in arq.upper() or "MENEGATTO" in arq.upper(): setor_inferido = "Setor Menegatto"
            xls = pd.ExcelFile(caminho_completo)
            for nome_aba in xls.sheet_names:
                df_aba = pd.read_excel(xls, sheet_name=nome_aba)
                col_maq = None
                col_tipo = None
                for c in df_aba.columns:
                    c_str = str(c).upper().strip()
                    if "MAQ" in c_str or "TAG" in c_str or "MÁQ" in c_str: col_maq = c
                    elif "FUSO" in c_str or "TIPO" in c_str or "MODELO" in c_str: col_tipo = c
                if not col_maq: continue 
                for _, linha in df_aba.iterrows():
                    tag_maquina = str(linha[col_maq]).strip().upper()
                    if not tag_maquina or tag_maquina.lower() in ["nan", "none", ""]: continue
                    tipo_fuso_lido = ""
                    if col_tipo and pd.notna(linha[col_tipo]):
                        val_fuso = str(linha[col_tipo]).strip().upper()
                        if val_fuso not in ["NAN", "NONE", ""]: tipo_fuso_lido = val_fuso
                    if not tipo_fuso_lido: tipo_fuso_lido = obter_fuso_padrao(tag_maquina, setor_inferido)
                    for c in df_aba.columns:
                        if c == col_maq or c == col_tipo: continue
                        if str(c).strip().isdigit():
                            dia = int(str(c).strip())
                            qtd = 0
                            try:
                                val = float(linha[c])
                                if pd.notna(val) and val > 0: qtd = int(val)
                            except: pass
                            if qtd > 0:
                                lista_registros.append({"Ano": 2026, "Mes": nome_aba.strip().capitalize(), "Dia": dia, "Setor": setor_inferido, "Maquina_TAG": tag_maquina, "Quantidade_Quebras": qtd, "Tipo_Fuso": tipo_fuso_lido})
        except Exception as e: pass
    if lista_registros: df_f = pd.DataFrame(lista_registros)
    else: df_f = pd.DataFrame(columns=COLUNAS_FUSOS)
    for col in COLUNAS_FUSOS:
        if col not in df_f.columns:
            if col == "Ano": df_f[col] = date.today().year
            elif col == "Mes": df_f[col] = "Janeiro"
            elif col == "Dia": df_f[col] = 1
            elif col == "Quantidade_Quebras": df_f[col] = 0
            else: df_f[col] = ""
    return df_f

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_trabalhos_fusos():
    if os.path.exists(ARQUIVO_TRABALHOS_FUSOS):
        try:
            df = pd.read_excel(ARQUIVO_TRABALHOS_FUSOS)
            mapa_colunas = {"SETOR": "Setor", "STATUS": "Status", "DATA": "Data", "MAQUINA": "Maquina", "QUANTIDADE": "Quantidade", "DESCRIÇÃO": "Descricao"}
            df = df.rename(columns={k: v for k, v in mapa_colunas.items() if k in df.columns})
            if "Data" in df.columns: df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
            return df
        except Exception as e: pass
    return pd.DataFrame(columns=["Setor", "Status", "Data", "Maquina", "Quantidade", "Descricao"])

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_observacoes_fusos():
    if os.path.exists(ARQUIVO_TRABALHOS_FUSOS):
        try:
            xls = pd.ExcelFile(ARQUIVO_TRABALHOS_FUSOS)
            aba_alvo = None
            
            # Pega obrigatoriamente a segunda planilha (índice 1) se existir
            if len(xls.sheet_names) > 1:
                aba_alvo = xls.sheet_names[1]
            else:
                for aba in xls.sheet_names:
                    if "OBSERV" in aba.upper():
                        aba_alvo = aba
                        break
            
            if aba_alvo:
                df = pd.read_excel(xls, sheet_name=aba_alvo)
                mapa_colunas = {"SETOR": "Setor", "TÍTULO": "Titulo", "TITULO": "Titulo", "DESCRIÇÃO": "Descricao", "DESCRICAO": "Descricao", "OBSERVAÇÃO": "Descricao", "OBSERVACAO": "Descricao"}
                df = df.rename(columns={k: v for k, v in mapa_colunas.items() if k in df.columns})
                
                for col in ["Setor", "Titulo", "Descricao"]:
                    if col not in df.columns:
                        df[col] = ""
                        
                return df
        except Exception:
            pass
    return pd.DataFrame(columns=["Setor", "Titulo", "Descricao"])

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_preventivas():
    if os.path.exists(ARQUIVO_PREVENTIVAS):
        try:
            df = pd.read_excel(ARQUIVO_PREVENTIVAS)
            colunas_padrao = {}
            for i, c in enumerate(df.columns):
                c_limpo = str(c).upper().strip()
                if "SETOR" in c_limpo or i == 0: colunas_padrao[c] = "SETOR"
                elif "MAQ" in c_limpo or i == 1: colunas_padrao[c] = "MAQUINA"
                elif "INIC" in c_limpo or "INÍC" in c_limpo or i == 2: colunas_padrao[c] = "INICIO"
                elif "FIM" in c_limpo or i == 3: colunas_padrao[c] = "FIM"
                elif "SERV" in c_limpo or i == 4: colunas_padrao[c] = "SERVIÇOS REALIZADOS"
                elif "PEÇ" in c_limpo or "PEC" in c_limpo or i == 5: colunas_padrao[c] = "PEÇAS TROCADAS"
                elif "EXEC" in c_limpo or i == 6: colunas_padrao[c] = "EXECUTANTES"
            
            df = df.rename(columns=colunas_padrao)
            if "INICIO" in df.columns: df["INICIO"] = pd.to_datetime(df["INICIO"], errors="coerce")
            if "FIM" in df.columns: df["FIM"] = pd.to_datetime(df["FIM"], errors="coerce")
            return df
        except Exception: pass
    df_vazio = pd.DataFrame(columns=["SETOR", "MAQUINA", "INICIO", "FIM", "SERVIÇOS REALIZADOS", "PEÇAS TROCADAS", "EXECUTANTES"])
    df_vazio.to_excel(ARQUIVO_PREVENTIVAS, index=False)
    return df_vazio

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_estoque():
    if os.path.exists(ARQUIVO_ESTOQUE):
        try: 
            xls = pd.ExcelFile(ARQUIVO_ESTOQUE)
            sheet = "ESTOQUE" if "ESTOQUE" in xls.sheet_names else 0
            df = pd.read_excel(xls, sheet_name=sheet)
            for col in COLUNAS_ESTOQUE:
                if col not in df.columns: df[col] = ""
            return df
        except Exception: 
            pass
    df_vazio = pd.DataFrame(columns=COLUNAS_ESTOQUE)
    with pd.ExcelWriter(ARQUIVO_ESTOQUE, engine='openpyxl') as writer:
        df_vazio.to_excel(writer, sheet_name="ESTOQUE", index=False)
        pd.DataFrame(columns=COLUNAS_MANUTENCAO).to_excel(writer, sheet_name="MANUTENÇÃO", index=False)
    return df_vazio

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_pecas_manutencao():
    if os.path.exists(ARQUIVO_ESTOQUE):
        try:
            xls = pd.ExcelFile(ARQUIVO_ESTOQUE)
            if "MANUTENÇÃO" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name="MANUTENÇÃO")
                if "DATA" in df.columns:
                    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
                for col in COLUNAS_MANUTENCAO:
                    if col not in df.columns: df[col] = ""
                return df
        except Exception:
            pass
    return pd.DataFrame(columns=COLUNAS_MANUTENCAO)

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_terceiros():
    if os.path.exists(ARQUIVO_TERCEIROS):
        try:
            xls = pd.ExcelFile(ARQUIVO_TERCEIROS)
            dfs = []
            for aba in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=aba)
                
                # Mapeia as colunas independente de maiúsculas/minúsculas
                df = df.rename(columns=lambda x: str(x).strip().upper())
                mapa_col = {"DATA": "Data", "SETOR": "Setor", "SERVICO": "Servico", "SERVIÇO": "Servico"}
                df = df.rename(columns=mapa_col)
                
                if "Data" in df.columns:
                    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
                
                for col in ["Data", "Setor", "Servico"]:
                    if col not in df.columns:
                        df[col] = None
                        
                df["Empresa"] = aba
                dfs.append(df)
            
            if dfs:
                return pd.concat(dfs, ignore_index=True)
        except Exception: 
            pass
            
    # Cria estrutura padrão caso o arquivo não exista ou dê erro
    df_vazio = pd.DataFrame(columns=["Data", "Setor", "Servico", "Empresa"])
    with pd.ExcelWriter(ARQUIVO_TERCEIROS, engine='openpyxl') as writer:
        df_vazio.to_excel(writer, sheet_name="PRIMETECH", index=False)
    return df_vazio

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_checklist_eletrica():
    if os.path.exists(ARQUIVO_CHECKLIST_ELETRICA):
        try:
            xls = pd.ExcelFile(ARQUIVO_CHECKLIST_ELETRICA)
            dfs = []
            for aba in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=aba)
                
                # Padroniza as colunas do arquivo
                mapa_colunas = {
                    "MAQ": "Maquina", 
                    "ITENS": "Item", 
                    "FREQUÊNCIA": "Frequencia", 
                    "FREQUENCIA": "Frequencia", 
                    "DATA ULTIMA": "Data_Ultima", 
                    "DATA ÚLTIMA": "Data_Ultima"
                }
                df = df.rename(columns=lambda x: str(x).strip().upper())
                df = df.rename(columns=mapa_colunas)
                
                # Garante que as colunas existam mesmo se a planilha estiver vazia
                for col in ["Maquina", "Item", "Frequencia", "Data_Ultima"]:
                    if col not in df.columns:
                        df[col] = None
                        
                if "Data_Ultima" in df.columns:
                    df["Data_Ultima"] = pd.to_datetime(df["Data_Ultima"], errors="coerce")
                    
                # A aba do arquivo vira o "Setor_Planilha"
                df["Setor_Planilha"] = aba
                dfs.append(df)
                
            if dfs:
                return pd.concat(dfs, ignore_index=True)
        except Exception: 
            pass
            
    # Cria estrutura padrão caso o arquivo não exista ou dê erro
    df_vazio = pd.DataFrame(columns=["Maquina", "Item", "Frequencia", "Data_Ultima", "Setor_Planilha"])
    with pd.ExcelWriter(ARQUIVO_CHECKLIST_ELETRICA, engine='openpyxl') as writer:
        df_vazio.to_excel(writer, sheet_name="Setor A", index=False)
    return df_vazio

@st.cache_data(show_spinner=False, ttl=5)
def carregar_dados_checklist_mecanica():
    if os.path.exists(ARQUIVO_CHECKLIST_MECANICA):
        try:
            xls = pd.ExcelFile(ARQUIVO_CHECKLIST_MECANICA)
            dfs = []
            for aba in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=aba)
                
                # Padroniza as colunas do arquivo
                mapa_colunas = {
                    "MAQ": "Maquina", 
                    "ITENS": "Item", 
                    "FREQUÊNCIA": "Frequencia", 
                    "FREQUENCIA": "Frequencia", 
                    "DATA ULTIMA": "Data_Ultima", 
                    "DATA ÚLTIMA": "Data_Ultima"
                }
                df = df.rename(columns=lambda x: str(x).strip().upper())
                df = df.rename(columns=mapa_colunas)
                
                # Garante que as colunas existam mesmo se a planilha estiver vazia
                for col in ["Maquina", "Item", "Frequencia", "Data_Ultima"]:
                    if col not in df.columns:
                        df[col] = None
                        
                if "Data_Ultima" in df.columns:
                    df["Data_Ultima"] = pd.to_datetime(df["Data_Ultima"], errors="coerce")
                    
                # A aba do arquivo vira o "Setor_Planilha"
                df["Setor_Planilha"] = aba
                dfs.append(df)
                
            if dfs:
                return pd.concat(dfs, ignore_index=True)
        except Exception: 
            pass
            
    # Cria estrutura padrão caso o arquivo não exista ou dê erro
    df_vazio = pd.DataFrame(columns=["Maquina", "Item", "Frequencia", "Data_Ultima", "Setor_Planilha"])
    with pd.ExcelWriter(ARQUIVO_CHECKLIST_MECANICA, engine='openpyxl') as writer:
        df_vazio.to_excel(writer, sheet_name="Setor A", index=False)
    return df_vazio

df_fusos = carregar_dados_fusos()
df_correias = carregar_dados_correias()
df_corretivas = carregar_dados_corretivas()
df_trabalhos_fusos = carregar_dados_trabalhos_fusos()
df_observacoes_fusos = carregar_dados_observacoes_fusos()
df_preventivas = carregar_dados_preventivas()
df_estoque = carregar_dados_estoque()
df_pecas_manutencao = carregar_dados_pecas_manutencao()
df_terceiros = carregar_dados_terceiros()
df_checklist_eletrica = carregar_dados_checklist_eletrica()
df_checklist_mecanica = carregar_dados_checklist_mecanica()

mapa_setor_maquina = {}
for s_nome, lista_m in DICIONARIO_SETORES.items():
    for m in lista_m: mapa_setor_maquina[m] = s_nome
for _, r in df_correias.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")): mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()
for _, r in df_fusos.iterrows():
    if pd.notna(r.get("Maquina_TAG")) and pd.notna(r.get("Setor")): mapa_setor_maquina[str(r["Maquina_TAG"]).strip()] = str(r["Setor"]).strip()

# ==========================================
# ESTADOS DA SESSÃO
# ==========================================
if "pagina_atual" not in st.session_state: st.session_state.pagina_atual = "Setores"
if "aba_setor_fuso" not in st.session_state: st.session_state.aba_setor_fuso = "Geral"
if "maq_clicada_cor" not in st.session_state: st.session_state.maq_clicada_cor = None
if "maq_clicada_prev" not in st.session_state: st.session_state.maq_clicada_prev = None
if "maq_clicada_check" not in st.session_state: st.session_state.maq_clicada_check = None
if "aba_estoque" not in st.session_state: st.session_state.aba_estoque = "Inventário"
if "empresa_clicada_terc" not in st.session_state: st.session_state.empresa_clicada_terc = None

# ==========================================
# CSS CUSTOMIZADO SUPERIOR
# ==========================================
CSS_STRING = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
html, body, [class*="css"]  { font-family: 'Inter', sans-serif !important; }
.block-container { padding: 3rem 2rem 1.5rem 2rem !important; max-width: 1400px; }
[data-testid="stSidebar"] { background-color: #0b1120 !important; border-right: 1px solid #1e293b !important; }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color: #f8fafc; }
[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
    background-color: transparent !important; color: #94a3b8 !important; border: 1px solid transparent !important;
    border-radius: 8px !important; font-weight: 600 !important; height: 42px !important; justify-content: flex-start; 
    padding-left: 10px !important; font-size: 0.85rem !important; transition: all 0.2s;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: #1e293b !important; color: #f8fafc !important; }
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563eb, #1d4ed8) !important; color: #ffffff !important;
    border: 1px solid transparent !important; border-radius: 8px !important; font-weight: 800 !important; 
    height: 42px !important; justify-content: flex-start; padding-left: 10px !important; font-size: 0.85rem !important; 
    box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
}
div.stButton > button {
    background: #ffffff !important; color: #0f172a !important; border: 1px solid #e2e8f0 !important;
    padding: 0px 4px !important; font-size: 0.85rem !important; font-weight: 800 !important;
    height: 32px !important; min-height: 32px !important; line-height: 28px !important;
    border-radius: 6px !important; box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important; transition: all 0.15s ease-in-out !important;
}
div.stButton > button:hover { border-color: #3b82f6 !important; background: #f8fafc !important; color: #1d4ed8 !important; transform: translateY(-1px); box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important; }
.card-kpi-bonito {
    background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 10px 16px;
    display: flex; align-items: center; justify-content: space-between; height: 72px; box-sizing: border-box;
    position: relative; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.04); transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.card-kpi-bonito:hover { transform: translateY(-3px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.08); }
.card-kpi-bonito::after { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 6px; }
.card-kpi-bonito.c-total::after { background: #3b82f6; }
.card-kpi-bonito.c-ok::after { background: #10b981; }
.card-kpi-bonito.c-warn::after { background: #f59e0b; }
.card-kpi-bonito.c-crit::after { background: #ef4444; }
.kpi-val { font-size: 1.35rem; font-weight: 900; line-height: 1.1; font-family: 'Inter', sans-serif; color: #0f172a;}
.kpi-lbl { font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }
.alerta-manutencao {
    background: #fef2f2; border: 1px solid #fecaca; border-left: 6px solid #ef4444; border-radius: 8px;
    padding: 8px 14px; margin: 6px 0 12px 0; font-size: 0.85rem; font-weight: 600; color: #991b1b; box-shadow: 0 2px 4px rgba(239, 68, 68, 0.05);
}
.badge-status { padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; }
.badge-verde { background: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }
.badge-amarelo { background: #fef9c3; color: #854d0e; border: 1px solid #fef08a; }
.badge-vermelho { background: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }
.badge-cinza { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }
.tag-pill { background: #f1f5f9; color: #334155; padding: 4px 10px; border-radius: 6px; font-size: 0.8rem; font-weight: 600; border: 1px solid #cbd5e1; }
.hud-detalhe { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 6px solid #3b82f6; }
.hud-detalhe.status-verde { border-left-color: #10b981; }
.hud-detalhe.status-amarelo { border-left-color: #f59e0b; }
.hud-detalhe.status-vermelho { border-left-color: #ef4444; }
.header-setor-dash { display:flex; justify-content:space-between; align-items:flex-end; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; margin-top: 15px; margin-bottom: 15px; font-weight: 900; font-size: 1.1rem; color: #0f172a; }
/* Ajustes para o expander do Streamlit */
[data-testid="stExpander"] { border: 1px solid #e2e8f0 !important; border-radius: 12px !important; box-shadow: 0 2px 4px rgba(0,0,0,0.02) !important; margin-bottom: 12px !important; overflow: hidden !important; }
[data-testid="stExpander"] summary { background: #f8fafc !important; padding: 15px 20px !important; border-bottom: 1px solid transparent !important; }
[data-testid="stExpander"] summary:hover { background: #f1f5f9 !important; }
[data-testid="stExpander"] [data-testid="stExpanderDetails"] { padding: 20px !important; background: #ffffff !important; border-top: 1px solid #e2e8f0 !important; }
details > summary { list-style: none; }
details > summary::-webkit-details-marker { display: none; }
</style>
"""
st.markdown(CSS_STRING, unsafe_allow_html=True)

# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown("<h2 style='font-size:1.45rem; font-weight:900; margin:0 0 18px 0; color:#ffffff; text-align:center;'>⚙️ Portal PCM</h2>", unsafe_allow_html=True)
    
    # ------------------------------------------
    # MÓDULO: PREVENTIVA
    # ------------------------------------------
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:8px 0 8px 0; letter-spacing:1px; text-align:center;'>MÓDULO: PREVENTIVA</p>", unsafe_allow_html=True)
    
    if st.button("📋 Painel Preventiva", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Registro Checklist" else "secondary"):
        st.session_state.pagina_atual = "Registro Checklist"
        st.rerun()

    if st.button("⚡ Preventiva Elétrica", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Checklist Eletrica" else "secondary"):
        st.session_state.pagina_atual = "Checklist Eletrica"
        st.rerun()
        
    if st.button("⚙️ Preventiva Mecânica", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Checklist Mecanica" else "secondary"):
        st.session_state.pagina_atual = "Checklist Mecanica"
        st.rerun()

    # ------------------------------------------
    # MÓDULO: FUSOS
    # ------------------------------------------
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:20px 0 8px 0; letter-spacing:1px; text-align:center;'>MÓDULO: FUSOS</p>", unsafe_allow_html=True)
    if st.button("🔩 Painel de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Fusos" else "secondary"):
        st.session_state.pagina_atual = "Fusos"
        st.rerun()
        
    if st.button("🛠️ Controle Sala de Fusos", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Sala Fusos" else "secondary"):
        st.session_state.pagina_atual = "Sala Fusos"
        st.rerun()

    # ------------------------------------------
    # MÓDULO: PAINÉIS
    # ------------------------------------------
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:20px 0 8px 0; letter-spacing:1px; text-align:center;'>MÓDULO: PAINÉIS</p>", unsafe_allow_html=True)
    if st.button("🏭 Painel de Setores", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Setores" else "secondary"):
        st.session_state.pagina_atual = "Setores"
        st.rerun()
        
    if st.button("🔄 Painel de Correias", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Correias" else "secondary"):
        st.session_state.pagina_atual = "Correias"
        st.rerun()

    # ------------------------------------------
    # MÓDULO: REGISTROS
    # ------------------------------------------
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:20px 0 8px 0; letter-spacing:1px; text-align:center;'>MÓDULO: REGISTROS</p>", unsafe_allow_html=True)
    if st.button("📖 Diário de Máquinas", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Diario" else "secondary"):
        st.session_state.pagina_atual = "Diario"
        st.rerun()

    if st.button("🤝 Serviços Terceiros", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Terceiros" else "secondary"):
        st.session_state.pagina_atual = "Terceiros"
        st.rerun()
        
    # ------------------------------------------
    # MÓDULO: MATERIAIS
    # ------------------------------------------
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:20px 0 8px 0; letter-spacing:1px; text-align:center;'>MÓDULO: MATERIAIS</p>", unsafe_allow_html=True)
    if st.button("📦 Estoque de Peças", use_container_width=True, type="primary" if st.session_state.pagina_atual == "Estoque" else "secondary"):
        st.session_state.pagina_atual = "Estoque"
        st.rerun()
        
   # ------------------------------------------
    # IMPORTAÇÃO DE DADOS (VISUALIZADOR)
    # ------------------------------------------
    import zipfile
    
    st.markdown("<hr style='border-color:#1e293b; margin:15px 0;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.75rem; font-weight:800; color:#94a3b8; margin:0 0 8px 0; text-align:center; letter-spacing:1px;'>📂 IMPORTAR DADOS</p>", unsafe_allow_html=True)
    
    arquivo_zip = st.file_uploader("Suba o arquivo .zip gerado no Portal Master", type="zip", label_visibility="collapsed")
    
    if arquivo_zip is not None:
        # Cria uma "etiqueta" para o código reconhecer se já descompactou este ficheiro
        id_arquivo = arquivo_zip.name + str(arquivo_zip.size)
        
        if st.session_state.get("ultimo_zip") != id_arquivo:
            os.makedirs(PASTA_BD, exist_ok=True)
            
            with zipfile.ZipFile(arquivo_zip, "r") as zip_ref:
                zip_ref.extractall(PASTA_BD)
            
            st.session_state["ultimo_zip"] = id_arquivo
            st.cache_data.clear()
            st.rerun() 
        else:
            st.success("✅ Base de dados atualizada e pronta!")
            # Botão de segurança para forçar a leitura dos ficheiros
            if st.button("🔄 Forçar Atualização dos Gráficos", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
            
    st.markdown("<br><div style='text-align:center; font-size:0.7rem; color:#475569; font-weight:600;'>Portal PCM • Visualizador Executivo</div>", unsafe_allow_html=True)


# ==========================================
# PAINEL DE SETORES
# ==========================================
if st.session_state.pagina_atual == "Setores":
    c_ts1, c_ts2, c_ts3, c_prt = st.columns([30, 15, 15, 12])
    with c_ts2:
        setores_disp = ["Todos os Setores"] + list(DICIONARIO_SETORES.keys())
        setor_filtro_painel = st.selectbox("Filtrar Setor:", setores_disp, label_visibility="collapsed")
    with c_ts3:
        meses_disp = ["Acumulado do Ano"] + LISTA_MESES_PUROS
        mes_filtro_painel = st.selectbox("Mês:", meses_disp, label_visibility="collapsed")
    with c_prt:
        modo_limpo = st.toggle("🖨️ Tela Limpa", key="tgl_set")
        if modo_limpo: st.markdown(CSS_TELA_LIMPA, unsafe_allow_html=True)

    with c_ts1: 
        titulo_setor = f"Painel {setor_filtro_painel}" if setor_filtro_painel != "Todos os Setores" else "Painel Todos os Setores"
        st.markdown(f"<h2 style='margin:0; font-weight:900;'>🏭 {titulo_setor}</h2>", unsafe_allow_html=True)

    st.caption("Visão executiva de confiabilidade: Distribuição de Intervenções, Ocorrências, MTBF, MTTR e Checklists.")
    
    # Card de Aviso sobre o início das medições (1 de Setembro)
    st.markdown("<div style='background:#eff6ff; border:1px solid #bfdbfe; border-left:5px solid #3b82f6; padding:10px 16px; border-radius:8px; margin-top:8px; margin-bottom:15px; font-size:0.85rem; color:#1e3a8a; font-weight:600;'>ℹ️ <b>Aviso de Medição:</b> O cálculo dos indicadores de Confiabilidade (<b>MTTR</b> e <b>MTBF</b>) considera exclusivamente os dados a partir de <b>1º de Setembro de 2026</b> (início oficial dos registros operacionais).</div>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # CÁLCULOS DE ENGENHARIA (MTBF, MTTR E COBERTURA)
    # ---------------------------------------------------------
    df_filtrado_setor = df_corretivas.copy()
    if setor_filtro_painel != "Todos os Setores":
        df_filtrado_setor = df_filtrado_setor[df_filtrado_setor["Setor"] == setor_filtro_painel]

    ano_atual = date.today().year
    data_corte_mtbf = pd.to_datetime('2026-09-01')

    if mes_filtro_painel != "Acumulado do Ano":
        num_mes = LISTA_MESES_PUROS.index(mes_filtro_painel) + 1
        if not df_filtrado_setor.empty and "Data" in df_filtrado_setor.columns:
            df_filtrado_setor = df_filtrado_setor[df_filtrado_setor["Data"].dt.month == num_mes]
        dias_periodo = calendar.monthrange(ano_atual, num_mes)[1]
        
        # Filtro especial de MTBF/MTTR (Corte de 1 de Setembro)
        df_mtbf = df_filtrado_setor.copy()
        if not df_mtbf.empty and "Data" in df_mtbf.columns:
            df_mtbf = df_mtbf[df_mtbf["Data"] >= data_corte_mtbf]
        dias_periodo_mtbf = dias_periodo if num_mes >= 9 else 0

    else:
        dias_periodo = max(1, (date.today() - date(ano_atual, 1, 1)).days)
        
        # Filtro especial de MTBF/MTTR para o Acumulado
        df_mtbf = df_filtrado_setor.copy()
        if not df_mtbf.empty and "Data" in df_mtbf.columns:
            df_mtbf = df_mtbf[df_mtbf["Data"] >= data_corte_mtbf]
        dias_periodo_mtbf = max(1, (date.today() - data_corte_mtbf.date()).days)

    total_horas_paradas = df_filtrado_setor["Duracao"].sum() if not df_filtrado_setor.empty else 0
    total_intervencoes = len(df_filtrado_setor) if not df_filtrado_setor.empty else 0
    
    preventivas_qtd = len(df_filtrado_setor[df_filtrado_setor["Tipo"].astype(str).str.contains("Prev", case=False, na=False)]) if not df_filtrado_setor.empty else 0
    corretivas_qtd = len(df_filtrado_setor[df_filtrado_setor["Tipo"].astype(str).str.contains("Corretiva", case=False, na=False)]) if not df_filtrado_setor.empty else 0
    
    # Define universo de máquinas do filtro
    if setor_filtro_painel == "Todos os Setores":
        maqs_escopo = list(mapa_setor_maquina.keys())
        if not maqs_escopo: maqs_escopo = [f"M-{i}" for i in range(1, 100)] # Fallback
    else:
        maqs_escopo = obter_maquinas_setor(setor_filtro_painel, df_fusos)
    
    qtd_maqs_escopo = max(1, len(maqs_escopo))

    # MTTR e MTBF (Convertido para Dias) com base PÓS 01/09
    corretivas_qtd_mtbf = len(df_mtbf[df_mtbf["Tipo"].astype(str).str.contains("Corretiva", case=False, na=False)]) if not df_mtbf.empty else 0
    horas_paradas_mtbf = df_mtbf["Duracao"].sum() if not df_mtbf.empty else 0

    mttr = horas_paradas_mtbf / corretivas_qtd_mtbf if corretivas_qtd_mtbf > 0 else 0.0
    tempo_disp_total = qtd_maqs_escopo * dias_periodo_mtbf * 24
    mtbf_horas = (tempo_disp_total - horas_paradas_mtbf) / corretivas_qtd_mtbf if corretivas_qtd_mtbf > 0 else tempo_disp_total
    mtbf_dias = mtbf_horas / 24

    # Coberturas Checklists (Cálculo Ajustado: Por Máquina)
    df_chk_mec_escopo = df_checklist_mecanica[df_checklist_mecanica["Maquina"].isin(maqs_escopo)]
    df_chk_ele_escopo = df_checklist_eletrica[df_checklist_eletrica["Maquina"].isin(maqs_escopo)]
    
    def calcular_cobertura_maquinas(df_chk, lista_maquinas):
        feitos = 0
        pendentes = 0
        for m in lista_maquinas:
            df_m = df_chk[df_chk["Maquina"] == m]
            if df_m.empty or df_m["Data_Ultima"].isna().any():
                pendentes += 1
            else:
                feitos += 1
        
        tot = len(lista_maquinas)
        cob = (feitos / tot * 100) if tot > 0 else 0.0
        return feitos, pendentes, cob

    feitos_mec, pend_mec, cob_mec = calcular_cobertura_maquinas(df_chk_mec_escopo, maqs_escopo)
    feitos_ele, pend_ele, cob_ele = calcular_cobertura_maquinas(df_chk_ele_escopo, maqs_escopo)

    # RENDERIZAÇÃO DOS KPIs SUPERIORES (APENAS COBERTURAS)
    k1, k2 = st.columns(2)
    
    cor_cob_mec = "#10b981" if cob_mec == 100 else "#f59e0b" if cob_mec >= 50 else "#ef4444"
    k1.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_cob_mec};'><div><div class='kpi-lbl'>Cobertura do Checklist (MECÂNICA)</div><div class='kpi-val' style='color:{cor_cob_mec};'>{cob_mec:.0f}% <span style='font-size:0.8rem; color:#64748b; font-weight:700;'>({feitos_mec} máqs em dia | {pend_mec} pendentes)</span></div></div><div style='font-size:1.6rem;'>⚙️</div></div>", unsafe_allow_html=True)

    cor_cob_ele = "#10b981" if cob_ele == 100 else "#f59e0b" if cob_ele >= 50 else "#ef4444"
    k2.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_cob_ele};'><div><div class='kpi-lbl'>Cobertura do Checklist (ELÉTRICA)</div><div class='kpi-val' style='color:{cor_cob_ele};'>{cob_ele:.0f}% <span style='font-size:0.8rem; color:#64748b; font-weight:700;'>({feitos_ele} máqs em dia | {pend_ele} pendentes)</span></div></div><div style='font-size:1.6rem;'>⚡</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:15px;'></div>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # LINHA 1 DE GRÁFICOS: DISTRIBUIÇÃO | MTTR | MTBF
    # ---------------------------------------------------------
    c_graf1, c_graf2, c_graf3 = st.columns([1.35, 1, 1])
    
    with c_graf1:
        with st.container(border=True):
            st.markdown("<div style='text-align:center; font-size:1.1rem; font-weight:900; color:#1e293b; margin-bottom:10px;'>🛠️ Distribuição de Intervenções</div>", unsafe_allow_html=True)
            if not df_filtrado_setor.empty and total_intervencoes > 0:
                df_tipo = df_filtrado_setor.copy()
                df_tipo["Tipo"] = df_tipo["Tipo"].astype(str).fillna("Não Informado")
                df_tipo_count = df_tipo.groupby("Tipo").size().reset_index(name="Ocorrencias")
                tipos_lidos = df_tipo_count["Tipo"].tolist()
                cores_lidas = [CORES_TIPO.get(t, PALETA_DINAMICA[i % len(PALETA_DINAMICA)]) for i, t in enumerate(tipos_lidos)]
                
                # Monta a Rosca
                rosca_tipo = alt.Chart(df_tipo_count).mark_arc(innerRadius=70, outerRadius=110, stroke="#ffffff", strokeWidth=2).encode(
                    theta=alt.Theta("Ocorrencias:Q"),
                    color=alt.Color("Tipo:N", scale=alt.Scale(domain=tipos_lidos, range=cores_lidas), legend=None),
                    tooltip=["Tipo", "Ocorrencias"]
                )
                
                # Textos centralizados com o Total
                texto_num = alt.Chart(pd.DataFrame({'T': [f"{total_intervencoes}"]})).mark_text(size=40, fontWeight=900, color="#0f172a", dy=-5).encode(text='T:N')
                texto_lbl = alt.Chart(pd.DataFrame({'L': ["OCORRÊNCIAS"]})).mark_text(size=10, fontWeight=800, color="#64748b", dy=20).encode(text='L:N')
                
                # Combina tudo (Altura exata de 240px para alinhar perfeitamente com os quadros laterais)
                grafico_final = alt.layer(rosca_tipo, texto_num, texto_lbl).properties(height=240)
                st.altair_chart(grafico_final, use_container_width=True)
                
                html_b3 = "<div style='display:flex; flex-wrap:wrap; gap:10px; justify-content:center; margin-top:15px;'>"
                for _, r in df_tipo_count.iterrows():
                    cor = CORES_TIPO.get(r['Tipo'], '#3b82f6')
                    pct_tipo = (r['Ocorrencias'] / total_intervencoes) * 100
                    html_b3 += f"<div style='background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:8px 12px; text-align:center; flex:1; min-width:80px;'><div style='font-size:0.75rem; color:#64748b; font-weight:700; text-transform:uppercase;'>{r['Tipo']}</div><div style='font-size:1.15rem; color:{cor}; font-weight:900;'>{pct_tipo:.1f}%</div></div>"
                html_b3 += "</div>"
                st.markdown(html_b3, unsafe_allow_html=True)
            else: 
                st.markdown("<div style='height:340px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Sem dados de intervenção no período.</div>", unsafe_allow_html=True)

    with c_graf2:
        cor_mttr = "#10b981" if mttr < 2 else "#f59e0b" if mttr < 4 else "#ef4444"
        html_mttr = f"""
        <div style='background:#ffffff; border:1px solid #e2e8f0; border-top:6px solid {cor_mttr}; border-radius:12px; padding:24px; box-shadow:0 4px 10px -2px rgba(0,0,0,0.05); height: 385px; display:flex; flex-direction:column;'>
            <div style='display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:15px;'>
                <div>
                    <div style='font-size:1.3rem; font-weight:900; color:#0f172a;'>MTTR</div>
                    <div style='font-size:0.72rem; font-weight:800; color:#64748b; text-transform:uppercase; letter-spacing: 0.5px;'>Tempo Médio de Reparo</div>
                </div>
                <div style='background:#f8fafc; padding:10px; border-radius:10px; font-size:1.4rem; border:1px solid #e2e8f0;'>⏱️</div>
            </div>
            <div style='flex-grow: 1; display:flex; align-items:center; justify-content:center;'>
                <div style='font-size:3.8rem; font-weight:900; color:{cor_mttr}; line-height:1; letter-spacing:-1px;'>{mttr:.1f}<span style='font-size:1.5rem; color:#94a3b8; margin-left:4px; font-weight:800;'>h</span></div>
            </div>
            <div style='font-size:0.85rem; color:#475569; font-weight:600; line-height:1.45; background:#f8fafc; padding:16px; border-radius:8px; border:1px dashed #cbd5e1; margin-top:15px;'>
                O <b>MTTR</b> indica o tempo médio gasto para reparar os ativos após uma falha. Quanto <b style='color:#10b981;'>menor</b> esse valor, mais ágil e eficiente é a equipe na resolução de problemas.
            </div>
        </div>
        """
        st.markdown(html_mttr, unsafe_allow_html=True)

    with c_graf3:
        # Régua ajustada para dias: > 12.5 dias (Verde), > 6.2 dias (Amarelo), senão (Vermelho)
        cor_mtbf = "#10b981" if mtbf_dias > 12.5 else "#f59e0b" if mtbf_dias > 6.2 else "#ef4444"
        html_mtbf = f"""
        <div style='background:#ffffff; border:1px solid #e2e8f0; border-top:6px solid {cor_mtbf}; border-radius:12px; padding:24px; box-shadow:0 4px 10px -2px rgba(0,0,0,0.05); height: 385px; display:flex; flex-direction:column;'>
            <div style='display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:15px;'>
                <div>
                    <div style='font-size:1.3rem; font-weight:900; color:#0f172a;'>MTBF</div>
                    <div style='font-size:0.72rem; font-weight:800; color:#64748b; text-transform:uppercase; letter-spacing: 0.5px;'>Tempo Médio Entre Falhas</div>
                </div>
                <div style='background:#f8fafc; padding:10px; border-radius:10px; font-size:1.4rem; border:1px solid #e2e8f0;'>📈</div>
            </div>
            <div style='flex-grow: 1; display:flex; align-items:center; justify-content:center;'>
                <div style='font-size:3.8rem; font-weight:900; color:{cor_mtbf}; line-height:1; letter-spacing:-1px;'>{mtbf_dias:.1f}<span style='font-size:1.3rem; color:#94a3b8; margin-left:6px; font-weight:800;'>dias</span></div>
            </div>
            <div style='font-size:0.85rem; color:#475569; font-weight:600; line-height:1.45; background:#f8fafc; padding:16px; border-radius:8px; border:1px dashed #cbd5e1; margin-top:15px;'>
                O <b>MTBF</b> mede a confiabilidade da planta, apontando o tempo que os ativos operam sem falhar. Quanto <b style='color:#10b981;'>maior</b> esse número, mais estável é a operação.
            </div>
        </div>
        """
        st.markdown(html_mtbf, unsafe_allow_html=True)

    st.markdown("<div style='height:15px;'></div>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # LINHA 2 DE GRÁFICOS: OCORRÊNCIAS POR MÊS E TOP MÁQUINAS
    # ---------------------------------------------------------
    c_graf4, c_graf5 = st.columns(2)
    
    with c_graf4:
        with st.container(border=True):
            st.markdown("<div style='text-align:center; font-size:1.1rem; font-weight:900; color:#1e293b; margin-bottom:10px;'>📊 Histórico de Ocorrências por Mês</div>", unsafe_allow_html=True)
            
            df_hist_ocorrencias = df_corretivas.copy()
            if setor_filtro_painel != "Todos os Setores":
                df_hist_ocorrencias = df_hist_ocorrencias[df_hist_ocorrencias["Setor"] == setor_filtro_painel]
            
            if not df_hist_ocorrencias.empty and "Data" in df_hist_ocorrencias.columns:
                df_hist_ocorrencias["Mes_Abrev"] = df_hist_ocorrencias["Data"].dt.month.map(lambda x: LISTA_MESES_PUROS[x-1][:3].upper() if pd.notna(x) else "")
                df_tendencia_oc = df_hist_ocorrencias.groupby("Mes_Abrev").size().reindex(ORDEM_MESES_ABREV, fill_value=0).reset_index(name="Ocorrencias")
                
                barras_oc = alt.Chart(df_tendencia_oc).mark_bar(color="#3b82f6", cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0, grid=False, labelFontWeight=800)),
                    y=alt.Y("Ocorrencias:Q", title="Qtd Ocorrências", axis=alt.Axis(grid=True)),
                    tooltip=["Mes_Abrev", "Ocorrencias"]
                )
                
                textos_oc = barras_oc.mark_text(dy=-8, fontSize=11, fontWeight=900).encode(
                    text=alt.condition("datum.Ocorrencias > 0", alt.Text("Ocorrencias:Q"), alt.value(""))
                )
                
                st.altair_chart((barras_oc + textos_oc).properties(height=240), use_container_width=True)
                
                # Caixas de Resumo Espelhadas (3 caixas para equilibrar a altura com o gráfico vizinho)
                tot_oc_hist = df_tendencia_oc["Ocorrencias"].sum()
                media_oc = df_tendencia_oc[df_tendencia_oc["Ocorrencias"] > 0]["Ocorrencias"].mean()
                if pd.isna(media_oc): media_oc = 0
                idx_pico = df_tendencia_oc["Ocorrencias"].idxmax()
                pico_mes = df_tendencia_oc.loc[idx_pico, "Mes_Abrev"] if tot_oc_hist > 0 else "-"
                pico_val = df_tendencia_oc.loc[idx_pico, "Ocorrencias"] if tot_oc_hist > 0 else 0

                html_resumo_oc = f"""
                <div style='display:flex; flex-wrap:wrap; gap:10px; justify-content:center; margin-top:15px;'>
                    <div style='background:#eff6ff; border:1px solid #bfdbfe; border-radius:8px; padding:8px; text-align:center; flex:1;'>
                        <div style='font-size:0.75rem; color:#1d4ed8; font-weight:700; text-transform:uppercase;'>Total Ano</div>
                        <div style='font-size:1.15rem; color:#2563eb; font-weight:900;'>{tot_oc_hist}</div>
                    </div>
                    <div style='background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:8px; text-align:center; flex:1;'>
                        <div style='font-size:0.75rem; color:#166534; font-weight:700; text-transform:uppercase;'>Média/Mês</div>
                        <div style='font-size:1.15rem; color:#10b981; font-weight:900;'>{media_oc:.1f}</div>
                    </div>
                    <div style='background:#fef2f2; border:1px solid #fecaca; border-radius:8px; padding:8px; text-align:center; flex:1;'>
                        <div style='font-size:0.75rem; color:#991b1b; font-weight:700; text-transform:uppercase;'>Pico ({pico_mes})</div>
                        <div style='font-size:1.15rem; color:#ef4444; font-weight:900;'>{pico_val}</div>
                    </div>
                </div>
                """
                st.markdown(html_resumo_oc, unsafe_allow_html=True)
            else: 
                st.markdown("<div style='height:340px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Sem histórico suficiente.</div>", unsafe_allow_html=True)

    with c_graf5:
        with st.container(border=True):
            st.markdown("<div style='text-align:center; font-size:1.1rem; font-weight:900; color:#1e293b; margin-bottom:10px;'>📉 Top Máquinas (Ocorrências)</div>", unsafe_allow_html=True)
            
            if not df_filtrado_setor.empty and total_intervencoes > 0:
                df_top_maqs_oc = df_filtrado_setor.groupby("Maquina").size().reset_index(name="Ocorrencias").sort_values("Ocorrencias", ascending=False).head(7)
                
                barras_maq_oc = alt.Chart(df_top_maqs_oc).mark_bar(color="#f59e0b", cornerRadiusTopRight=5, cornerRadiusBottomRight=5).encode(
                    x=alt.X("Ocorrencias:Q", title="Qtd Ocorrências", axis=alt.Axis(grid=True)), 
                    y=alt.Y("Maquina:N", sort="-x", title=None, axis=alt.Axis(labelFontWeight=800)), 
                    tooltip=["Maquina", "Ocorrencias"]
                )
                
                textos_maq_oc = barras_maq_oc.mark_text(align='left', baseline='middle', dx=3, fontWeight=900).encode(
                    text='Ocorrencias:Q'
                )
                
                st.altair_chart((barras_maq_oc + textos_maq_oc).properties(height=240), use_container_width=True)
                
                # Caixas de Resumo Top 3
                html_top_oc = "<div style='display:flex; flex-wrap:wrap; gap:10px; justify-content:center; margin-top:15px;'>"
                for _, r in df_top_maqs_oc.head(3).iterrows():
                    html_top_oc += f"<div style='background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:8px; text-align:center; flex:1;'><div style='font-size:0.75rem; color:#b45309; font-weight:700; text-transform:uppercase;'>{r['Maquina']}</div><div style='font-size:1.15rem; color:#f59e0b; font-weight:900;'>{r['Ocorrencias']} un.</div></div>"
                html_top_oc += "</div>"
                st.markdown(html_top_oc, unsafe_allow_html=True)
            else: 
                st.markdown("<div style='height:340px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Nenhuma ocorrência registrada para o período.</div>", unsafe_allow_html=True)


# ==========================================
# PAINEL DE CORREIAS (DESIGN PREMIUM & INTERATIVO)
# ==========================================
elif st.session_state.pagina_atual == "Correias":
    data_hoje = date.today()
    dados_maquinas = {}
    lista_correias_todas, lista_correias_novas, lista_correias_meia, lista_correias_criticas = [], [], [], []

    def avaliar_correia(dt_val):
        if not dt_val or str(dt_val).strip() in ["", "nan", "NaT", "None"]:
            return None, "Sem registro", "Sem histórico", 0
        try:
            dt_inst = pd.to_datetime(dt_val).date()
            dias = max(0, (data_hoje - dt_inst).days)
            meses = round(dias / 30.4, 1)
            score = 1 if dias <= 365 else 2 if dias <= 547 else 3
            return score, dt_inst.strftime("%d/%m/%Y"), f"{meses} meses ({dias}d)", dias
        except:
            return None, str(dt_val), "Data inválida", 0

    todas_maquinas_totais = []
    for s_nome in DICIONARIO_SETORES.keys():
        for m in obter_maquinas_setor(s_nome, df_fusos):
            if m not in todas_maquinas_totais:
                todas_maquinas_totais.append(m)

    ult_correias_dict = {}
    if not df_correias.empty:
        for (s_idx, m_idx), sub_c in df_correias.groupby(["Setor", "Maquina_TAG"]):
            ult_correias_dict[(s_idx, m_idx)] = sub_c.iloc[-1]

    for maq_tag in todas_maquinas_totais:
        setor_m = mapa_setor_maquina.get(maq_tag, "Setor A")
        ult = ult_correias_dict.get((setor_m, maq_tag))
        
        t1, d1_str, t1_uso, c1_score, dias1 = "Não informada", "Sem registro", "Sem histórico", None, 0
        t2, d2_str, t2_uso, c2_score, dias2 = "Não informada", "Sem registro", "Sem histórico", None, 0

        if ult is not None:
            v1, dt1_raw = formatar_modelo(ult.get("Tipo_Correia_1", "")), str(ult.get("Data_Instalacao_1", "")).strip()
            if v1 or (dt1_raw not in ["", "nan", "NaT", "None"]):
                t1 = v1 or "Não informada"
                c1_score, d1_str, t1_uso, dias1 = avaliar_correia(dt1_raw)
                r1 = {"tag": maq_tag, "pos": "Superior", "modelo": t1, "data": d1_str, "uso": t1_uso, "setor": setor_m, "score": c1_score}
                lista_correias_todas.append(r1)
                if c1_score == 1: lista_correias_novas.append(r1)
                elif c1_score == 2: lista_correias_meia.append(r1)
                elif c1_score == 3: lista_correias_criticas.append(r1)

            v2, dt2_raw = formatar_modelo(ult.get("Tipo_Correia_2", "")), str(ult.get("Data_Instalacao_2", "")).strip()
            if v2 or (dt2_raw not in ["", "nan", "NaT", "None"]):
                t2 = v2 or "Não informada"
                c2_score, d2_str, t2_uso, dias2 = avaliar_correia(dt2_raw)
                r2 = {"tag": maq_tag, "pos": "Inferior", "modelo": t2, "data": d2_str, "uso": t2_uso, "setor": setor_m, "score": c2_score}
                lista_correias_todas.append(r2)
                if c2_score == 1: lista_correias_novas.append(r2)
                elif c2_score == 2: lista_correias_meia.append(r2)
                elif c2_score == 3: lista_correias_criticas.append(r2)

        scores = [s for s in [c1_score, c2_score] if s is not None]
        if scores:
            pior = max(scores)
            classe_card = "status-verde" if pior == 1 else "status-amarelo" if pior == 2 else "status-vermelho"
            status_label = "Nova (Em Dia)" if pior == 1 else "Meia-Vida" if pior == 2 else "Troca Necessária"
            cor_key = "verde" if pior == 1 else "amarelo" if pior == 2 else "vermelho"
        else:
            classe_card, status_label, cor_key = "status-cinza", "Sem Dados", "cinza"

        dados_maquinas[maq_tag] = {
            "setor": setor_m,
            "t1": t1, "d1": d1_str, "uso1": t1_uso, "score1": c1_score, "dias1": dias1,
            "t2": t2, "d2": d2_str, "uso2": t2_uso, "score2": c2_score, "dias2": dias2,
            "status_label": status_label, "classe_card": classe_card, "cor_key": cor_key
        }

    # CABEÇALHO E FILTROS
    c_t, c_f1, c_f2 = st.columns([5.5, 2.25, 2.25])
    with c_t:
        st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>🔄 Gestão de Correias e Transmissão</h2>", unsafe_allow_html=True)
        st.markdown("<div style='color:#64748b; font-size:0.9rem; font-weight:500;'>Monitoramento de vida útil dos eixos Superior e Inferior por ativo.</div>", unsafe_allow_html=True)
    with c_f1:
        mods_un = sorted(list({r["modelo"] for r in lista_correias_todas if r["modelo"] and r["modelo"] != "Não informada"}))
        filtro_modelo = st.selectbox("🏷️ Modelo:", ["Todos os Modelos"] + mods_un, label_visibility="collapsed")
    with c_f2:
        filtro_status = st.selectbox("🚦 Status:", ["Todos os Status", "🟢 Novas (≤ 1 ano)", "🟡 Meia-Vida (1 a 1.5 ano)", "🔴 Críticas (> 1.5 ano)", "⚪ Sem Dados"], label_visibility="collapsed")

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    # KPIs COM PERCENTUAL DE SAÚDE
    tot_c = len(lista_correias_todas)
    qtd_ok = len(lista_correias_novas)
    qtd_mv = len(lista_correias_meia)
    qtd_cr = len(lista_correias_criticas)
    pct_saude = int(round(((qtd_ok + qtd_mv) / tot_c) * 100)) if tot_c > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Correias Mapeadas</div><div class='kpi-val'>{tot_c} <span style='font-size:0.8rem; color:#64748b; font-weight:700;'>({pct_saude}% saudáveis)</span></div></div><div style='font-size:1.8rem;'>⚙️</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Vida Útil OK (≤ 1 Ano)</div><div class='kpi-val' style='color:#059669;'>{qtd_ok} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🟢</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Meia-Vida (1 a 1.5 Ano)</div><div class='kpi-val' style='color:#d97706;'>{qtd_mv} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🟡</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Críticas / Troca (> 1.5 Ano)</div><div class='kpi-val' style='color:#dc2626;'>{qtd_cr} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🔴</div></div>", unsafe_allow_html=True)

    # ALERTA DE COMPRAS / REPOSIÇÃO POR MODELO CRÍTICO
    if lista_correias_criticas:
        contagem_crit = pd.DataFrame(lista_correias_criticas)["modelo"].value_counts()
        chips = "".join([
            f"<span style='background:#ffffff; color:#991b1b; border:1px solid #fca5a5; border-left:4px solid #ef4444; padding:4px 12px; border-radius:8px; font-size:0.82rem; font-weight:800; margin-right:8px; display:inline-block; margin-top:4px; box-shadow:0 1px 2px rgba(0,0,0,0.04);'>🏷️ Modelo <b>{k}</b>: <span style='background:#fee2e2; padding:1px 6px; border-radius:4px; color:#dc2626;'>{v} un.</span></span>"
            for k, v in contagem_crit.items()
        ])
        html_alerta = (
            "<div style='background:linear-gradient(90deg, #fef2f2 0%, #fff5f5 100%); border:1px solid #fecaca; border-left:6px solid #ef4444; border-radius:12px; padding:12px 18px; margin:14px 0 10px 0; box-shadow:0 4px 6px -1px rgba(239, 68, 68, 0.06);'>"
            "<div style='font-size:0.85rem; font-weight:900; color:#991b1b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;'>🚨 Resumo de Correias Vencidas por Especificação (Para Separação/Troca):</div>"
            f"<div>{chips}</div></div>"
        )
        st.markdown(html_alerta, unsafe_allow_html=True)

    # =========================================================================
    # HUD DE DETALHES DA MÁQUINA SELECIONADA (DOSSIÊ RÁPIDO DOS EIXOS)
    # =========================================================================
    if st.session_state.maq_clicada_cor:
        sel_tag = st.session_state.maq_clicada_cor.get("tag")
        sel = dados_maquinas.get(sel_tag, st.session_state.maq_clicada_cor)
        
        cor_hero = "#10b981" if sel["cor_key"] == "verde" else "#f59e0b" if sel["cor_key"] == "amarelo" else "#ef4444" if sel["cor_key"] == "vermelho" else "#94a3b8"
        
        def montar_card_eixo(titulo_eixo, sub_eixo, icone_eixo, modelo, data_inst, tempo_uso, score_eixo, dias_eixo):
            if score_eixo == 1:
                c_borda, c_bg, c_txt, lbl_st = "#10b981", "#dcfce7", "#166534", "🟢 NOVA / OK"
            elif score_eixo == 2:
                c_borda, c_bg, c_txt, lbl_st = "#f59e0b", "#fef9c3", "#854d0e", "🟡 MEIA-VIDA"
            elif score_eixo == 3:
                c_borda, c_bg, c_txt, lbl_st = "#ef4444", "#fee2e2", "#991b1b", "🔴 TROCA URGENTE"
            else:
                c_borda, c_bg, c_txt, lbl_st = "#64748b", "#f1f5f9", "#475569", "⚪ SEM DADOS"
            
            # Calcula barra de desgaste (base 547 dias = 1.5 ano como 100% do limite)
            pct_desgaste = min(100, int(round((dias_eixo / 547) * 100))) if score_eixo else 0
            
            return (
                f"<div style='flex:1; min-width:280px; background:rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.12); border-left:5px solid {c_borda}; border-radius:12px; padding:16px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;'>"
                f"<div style='display:flex; align-items:center; gap:8px;'><span style='font-size:1.25rem;'>{icone_eixo}</span><div><div style='color:#ffffff; font-weight:900; font-size:1rem;'>{titulo_eixo}</div><div style='color:#94a3b8; font-size:0.7rem; font-weight:700; text-transform:uppercase;'>{sub_eixo}</div></div></div>"
                f"<span style='background:{c_bg}; color:{c_txt}; padding:3px 10px; border-radius:20px; font-size:0.72rem; font-weight:900;'>{lbl_st}</span>"
                "</div>"
                "<div style='display:flex; justify-content:space-between; align-items:center; background:rgba(15,23,42,0.5); padding:10px 14px; border-radius:8px; margin-bottom:10px;'>"
                f"<div><div style='color:#94a3b8; font-size:0.68rem; font-weight:700; text-transform:uppercase;'>Especificação</div><div style='color:#60a5fa; font-size:1.1rem; font-weight:900;'>🏷️ {modelo}</div></div>"
                f"<div style='text-align:right;'><div style='color:#94a3b8; font-size:0.68rem; font-weight:700; text-transform:uppercase;'>Instalação</div><div style='color:#f8fafc; font-size:0.9rem; font-weight:800;'>📅 {data_inst}</div></div>"
                "</div>"
                "<div style='display:flex; justify-content:space-between; font-size:0.75rem; color:#cbd5e1; font-weight:700; margin-bottom:4px;'>"
                f"<span>⏱️ Uso: <b style='color:#ffffff;'>{tempo_uso}</b></span>"
                f"<span>Desgaste: <b style='color:{c_borda};'>{pct_desgaste}%</b></span>"
                "</div>"
                "<div style='width:100%; height:7px; background:rgba(255,255,255,0.12); border-radius:99px; overflow:hidden;'>"
                f"<div style='width:{pct_desgaste}%; height:100%; background:{c_borda}; border-radius:99px;'></div>"
                "</div></div>"
            )

        card_sup = montar_card_eixo("Correia Superior", "Transmissão Principal", "🔼", sel["t1"], sel["d1"], sel["uso1"], sel.get("score1"), sel.get("dias1", 0))
        card_inf = montar_card_eixo("Correia Inferior", "Acionamento Secundário", "🔽", sel["t2"], sel["d2"], sel["uso2"], sel.get("score2"), sel.get("dias2", 0))

        c_hud, c_cls = st.columns([8.6, 1.4])
        with c_hud:
            html_hud = (
                f"<div style='background:linear-gradient(115deg, #0f172a 0%, #1e293b 100%); border-radius:16px; padding:20px 24px; margin:10px 0 15px 0; border-left:8px solid {cor_hero}; box-shadow:0 12px 20px -5px rgba(15, 23, 42, 0.25);'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:10px;'>"
                f"<div style='display:flex; align-items:center; gap:12px;'><span style='background:{cor_hero}; color:#ffffff; padding:5px 14px; border-radius:8px; font-size:1.15rem; font-weight:900;'>⚙️ {sel_tag}</span><span style='color:#cbd5e1; font-size:0.9rem; font-weight:700;'>📍 {sel['setor']}</span></div>"
                f"<span style='color:{cor_hero}; font-size:0.85rem; font-weight:900; text-transform:uppercase; letter-spacing:0.5px;'>Diagnóstico: {sel['status_label']}</span>"
                "</div>"
                f"<div style='display:flex; gap:16px; flex-wrap:wrap;'>{card_sup}{card_inf}</div>"
                "</div>"
            )
            st.markdown(html_hud, unsafe_allow_html=True)
        with c_cls:
            st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
            if st.button("✖ Fechar Card", key="btn_cls_hud", use_container_width=True):
                st.session_state.maq_clicada_cor = None
                st.rerun()

    # BARRA DE LEGENDA PREMIUM
    html_legenda = (
        "<div style='background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:12px 20px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.03); margin-top:12px; margin-bottom:18px; display:flex; gap:12px; flex-wrap:wrap; align-items:center; justify-content:space-between;'>"
        "<div style='font-size:0.8rem; font-weight:900; color:#475569; letter-spacing:0.5px;'>🎨 GUIA DE VIDA ÚTIL (CLIQUE NA MÁQUINA PARA INSPECIONAR OS EIXOS):</div>"
        "<div style='display:flex; gap:10px; flex-wrap:wrap;'>"
        "<span style='background:linear-gradient(135deg, #f0fdf4, #dcfce7); color:#166534; border:1px solid #86efac; border-left:5px solid #10b981; padding:5px 12px; border-radius:8px; font-size:0.78rem; font-weight:800;'>🟢 Nova (&le; 1 ano)</span>"
        "<span style='background:linear-gradient(135deg, #fefce8, #fef08a); color:#854d0e; border:1px solid #fde047; border-left:5px solid #eab308; padding:5px 12px; border-radius:8px; font-size:0.78rem; font-weight:800;'>🟡 Meia-Vida (1 a 1.5 ano)</span>"
        "<span style='background:linear-gradient(135deg, #fef2f2, #fee2e2); color:#991b1b; border:1px solid #fca5a5; border-left:5px solid #ef4444; padding:5px 12px; border-radius:8px; font-size:0.78rem; font-weight:800;'>🔴 Troca Urgente (&gt; 1.5 ano)</span>"
        "<span style='background:linear-gradient(135deg, #f8fafc, #f1f5f9); color:#475569; border:1px solid #cbd5e1; border-left:5px solid #94a3b8; padding:5px 12px; border-radius:8px; font-size:0.78rem; font-weight:800;'>⚪ Sem Histórico</span>"
        "</div></div>"
    )
    st.markdown(html_legenda, unsafe_allow_html=True)

    # MAPA DE MÁQUINAS POR SETOR COM CARDS COLORIDOS
    mapa_filtro_st = {
        "🟢 Novas (≤ 1 ano)": "verde",
        "🟡 Meia-Vida (1 a 1.5 ano)": "amarelo",
        "🔴 Críticas (> 1.5 ano)": "vermelho",
        "⚪ Sem Dados": "cinza"
    }

    for s_nome in DICIONARIO_SETORES.keys():
        maquinas_do_setor = obter_maquinas_setor(s_nome, df_fusos)
        
        # Filtra por Modelo
        if filtro_modelo != "Todos os Modelos":
            maquinas_do_setor = [
                m for m in maquinas_do_setor
                if dados_maquinas.get(m, {}).get("t1") == filtro_modelo or dados_maquinas.get(m, {}).get("t2") == filtro_modelo
            ]
            
        # Filtra por Status
        if filtro_status != "Todos os Status":
            alvo_cor = mapa_filtro_st.get(filtro_status)
            maquinas_do_setor = [
                m for m in maquinas_do_setor
                if dados_maquinas.get(m, {}).get("cor_key") == alvo_cor
            ]

        if not maquinas_do_setor:
            continue

        # Resumo do setor no cabeçalho da linha
        c_ok_s = sum(1 for m in maquinas_do_setor if dados_maquinas.get(m, {}).get("cor_key") == "verde")
        c_mv_s = sum(1 for m in maquinas_do_setor if dados_maquinas.get(m, {}).get("cor_key") == "amarelo")
        c_cr_s = sum(1 for m in maquinas_do_setor if dados_maquinas.get(m, {}).get("cor_key") == "vermelho")
        cor_setor_b = CORES_SETORES.get(s_nome, "#3b82f6")

        html_hdr_s = (
            f"<div style='display:flex; justify-content:space-between; align-items:center; background:#ffffff; border:1px solid #e2e8f0; border-left:6px solid {cor_setor_b}; border-radius:10px; padding:10px 16px; margin-top:18px; margin-bottom:12px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>"
            f"<div style='font-weight:900; font-size:1.05rem; color:#0f172a;'>🏭 {s_nome} <span style='font-size:0.78rem; color:#64748b; font-weight:700;'>({len(maquinas_do_setor)} ativos)</span></div>"
            "<div style='display:flex; gap:8px;'>"
            f"<span style='background:#dcfce7; color:#166534; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:800;'>🟢 {c_ok_s} OK</span>"
            f"<span style='background:#fef9c3; color:#854d0e; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:800;'>🟡 {c_mv_s} Meia-Vida</span>"
            f"<span style='background:#fee2e2; color:#991b1b; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:800;'>🔴 {c_cr_s} Críticas</span>"
            "</div></div>"
        )
        st.markdown(html_hdr_s, unsafe_allow_html=True)
        
        cols_g = 10
        for chunk in [maquinas_do_setor[i:i + cols_g] for i in range(0, len(maquinas_do_setor), cols_g)]:
            cols = st.columns(cols_g)
            for i, m in enumerate(chunk):
                info_m = dados_maquinas.get(m, {})
                cor_k = info_m.get("cor_key", "cinza")
                
                if cor_k == "verde":
                    icone = "🟢"
                    bg_btn = "linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%)"
                    bg_hover = "linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%)"
                    brd_btn, brd_left, txt_color, shd_btn = "#86efac", "#10b981", "#14532d", "rgba(16, 185, 129, 0.18)"
                elif cor_k == "amarelo":
                    icone = "🟡"
                    bg_btn = "linear-gradient(135deg, #fefce8 0%, #fef9c3 100%)"
                    bg_hover = "linear-gradient(135deg, #fef9c3 0%, #fef08a 100%)"
                    brd_btn, brd_left, txt_color, shd_btn = "#fde047", "#eab308", "#713f12", "rgba(234, 179, 8, 0.22)"
                elif cor_k == "vermelho":
                    icone = "🔴"
                    bg_btn = "linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%)"
                    bg_hover = "linear-gradient(135deg, #fee2e2 0%, #fecaca 100%)"
                    brd_btn, brd_left, txt_color, shd_btn = "#fca5a5", "#ef4444", "#7f1d1d", "rgba(239, 68, 68, 0.22)"
                else:
                    icone = "⚪"
                    bg_btn = "linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)"
                    bg_hover = "linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%)"
                    brd_btn, brd_left, txt_color, shd_btn = "#cbd5e1", "#94a3b8", "#334155", "rgba(100, 116, 139, 0.12)"

                css_anchor = f"ancora_cor_{m.replace('-', '_')}"
                css_html = (
                    f"<div id='{css_anchor}'></div>"
                    "<style>"
                    f"div.element-container:has(#{css_anchor}) {{ display: none !important; }}"
                    f"div.element-container:has(#{css_anchor}) + div.element-container {{ margin-bottom: -6px !important; }}"
                    f"div.element-container:has(#{css_anchor}) + div.element-container button,"
                    f"div[data-testid='column']:has(#{css_anchor}) button,"
                    f"div[data-testid='stColumn']:has(#{css_anchor}) button {{"
                    f"background: {bg_btn} !important; border: 1px solid {brd_btn} !important; border-left: 5px solid {brd_left} !important;"
                    f"color: {txt_color} !important; font-weight: 900 !important; border-radius: 10px !important;"
                    f"box-shadow: 0 3px 6px -1px {shd_btn}, 0 2px 4px -2px rgba(0,0,0,0.05) !important;"
                    "height: 42px !important; min-height: 42px !important; width: 100% !important; padding: 0 4px !important;"
                    "transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;"
                    "}"
                    f"div.element-container:has(#{css_anchor}) + div.element-container button p,"
                    f"div[data-testid='column']:has(#{css_anchor}) button p,"
                    f"div[data-testid='stColumn']:has(#{css_anchor}) button p {{"
                    f"font-size: 0.85rem !important; font-weight: 900 !important; color: {txt_color} !important; margin: 0 !important;"
                    "}"
                    f"div[data-testid='column']:has(#{css_anchor}) button:hover,"
                    f"div[data-testid='stColumn']:has(#{css_anchor}) button:hover {{"
                    f"background: {bg_hover} !important; border-color: {brd_left} !important;"
                    "transform: translateY(-3px) scale(1.03) !important;"
                    f"box-shadow: 0 10px 15px -3px {shd_btn}, 0 4px 6px -2px rgba(0,0,0,0.08) !important; z-index: 10 !important;"
                    "}"
                    "</style>"
                )
                cols[i].markdown(css_html, unsafe_allow_html=True)
                
                help_txt = f"Sup: {info_m.get('t1')} ({info_m.get('d1')}) | Inf: {info_m.get('t2')} ({info_m.get('d2')})"
                if cols[i].button(f"{icone} {m}", key=f"btn_c_{m.replace('-', '_')}", help=help_txt, use_container_width=True):
                    st.session_state.maq_clicada_cor = {"tag": m, **info_m}
                    st.rerun()


# ==========================================
# PAINEL DE FUSOS (DESIGN PREMIUM & KANBAN EXECUTIVO)
# ==========================================
elif st.session_state.pagina_atual == "Fusos":

    cf_t, cf_a, cf_prt = st.columns([6.0, 2.0, 2.0])
    with cf_t:
        st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>🔩 Inteligência & Controle de Fusos</h2>", unsafe_allow_html=True)
        st.markdown("<div style='color:#64748b; font-size:0.9rem; font-weight:500;'>Monitoramento de quebras, metas setoriais, mapa de calor e plano de ação.</div>", unsafe_allow_html=True)
    with cf_a:
        ano_f = st.selectbox("📅 Ano Base", [2024, 2025, 2026, 2027], index=2, label_visibility="collapsed")
    with cf_prt:
        modo_limpo = st.toggle("🖨️ Tela Limpa", key="tgl_fus")
        if modo_limpo:
            st.markdown(CSS_TELA_LIMPA, unsafe_allow_html=True)

    data_hoje_ref = date.today()
    ano_atual_ref = data_hoje_ref.year
    mes_atual_num_ref = data_hoje_ref.month
    div_meses = 12 if int(ano_f) < ano_atual_ref else max(1, mes_atual_num_ref) if int(ano_f) == ano_atual_ref else 1
    desc_divisor = "12 meses" if int(ano_f) < ano_atual_ref else f"Jan a {ORDEM_MESES_ABREV[div_meses - 1]}" if int(ano_f) == ano_atual_ref else "Previsto"

    METAS_SETOR_FUSO = {"Setor A": 90, "Setor B": 90, "Setor Látex": 45, "Setor Menegatto": 45}

    # BARRA DE NAVEGAÇÃO DE SETORES
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
    b_geral, b_sa, b_sb, b_lat, b_men = st.columns(5)
    abas_fuso_info = [
        (b_geral, "Geral", "🌐 Visão Geral"),
        (b_sa, "Setor A", "🏭 Setor A"),
        (b_sb, "Setor B", "🏭 Setor B"),
        (b_lat, "Setor Látex", "🌿 Setor Látex"),
        (b_men, "Setor Menegatto", "⚙️ Setor Menegatto")
    ]
    for b_col, nome_aba, rotulo_btn in abas_fuso_info:
        if b_col.button(rotulo_btn, use_container_width=True, type="primary" if st.session_state.aba_setor_fuso == nome_aba else "secondary"):
            st.session_state.aba_setor_fuso = nome_aba
            st.rerun()
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    df_ano_f = df_fusos[df_fusos["Ano"] == int(ano_f)].copy() if not df_fusos.empty else pd.DataFrame(columns=COLUNAS_FUSOS)

    # =========================================================================
    # ABA 1: VISÃO GERAL DA FÁBRICA
    # =========================================================================
    if st.session_state.aba_setor_fuso == "Geral":
        tot_fabrica = int(df_ano_f["Quantidade_Quebras"].sum()) if not df_ano_f.empty else 0
        med_fabrica = round(tot_fabrica / div_meses, 1)
        ult_mes_fab, tot_ult_mes, setor_ofensor, qtd_setor_ofensor = "Nenhum", 0, "Nenhum", 0

        if not df_ano_f.empty and tot_fabrica > 0:
            df_reais = df_ano_f[df_ano_f["Quantidade_Quebras"] > 0]
            for m_teste in reversed(LISTA_MESES_PUROS):
                sub_m = df_reais[df_reais["Mes"] == m_teste]
                if not sub_m.empty and sub_m["Quantidade_Quebras"].sum() > 0:
                    ult_mes_fab = m_teste
                    tot_ult_mes = int(sub_m["Quantidade_Quebras"].sum())
                    agrup_s = sub_m.groupby("Setor")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                    setor_ofensor = agrup_s.index[0]
                    qtd_setor_ofensor = int(agrup_s.iloc[0])
                    break

        projecao_mes = "-"
        lbl_projecao = "Projeção Mês"
        if int(ano_f) == ano_atual_ref:
            nome_mes_atual = LISTA_MESES_PUROS[mes_atual_num_ref - 1]
            df_mes_atual = df_ano_f[df_ano_f["Mes"] == nome_mes_atual]
            quebras_mes_ate_hoje = int(df_mes_atual["Quantidade_Quebras"].sum()) if not df_mes_atual.empty else 0
            dias_passados = data_hoje_ref.day
            dias_totais_mes = calendar.monthrange(ano_atual_ref, mes_atual_num_ref)[1]
            if dias_passados > 0:
                projecao_mes = int(round((quebras_mes_ate_hoje / dias_passados) * dias_totais_mes, 0))
            lbl_projecao = f"Projeção ({MAPA_MES_ABREV[nome_mes_atual]})"

        kf1, kf2, kf3, kf4, kf5 = st.columns(5)
        kf1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Fábrica ({ano_f})</div><div class='kpi-val'>{tot_fabrica} <span style='font-size:0.8rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        kf2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_fabrica}</div></div><div style='font-size:1.8rem;'>📈</div></div>", unsafe_allow_html=True)
        kf3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Setor Crítico ({ult_mes_fab})</div><div class='kpi-val' style='color:#d97706; font-size:1.15rem;'>{setor_ofensor} ({qtd_setor_ofensor})</div></div><div style='font-size:1.8rem;'>🏭</div></div>", unsafe_allow_html=True)
        kf4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Quebras no Mês ({ult_mes_fab})</div><div class='kpi-val' style='color:#dc2626;'>{tot_ult_mes}</div></div><div style='font-size:1.8rem;'>🚨</div></div>", unsafe_allow_html=True)
        kf5.markdown(f"<div class='card-kpi-bonito' style='border-left:6px solid #6366f1;'><div><div class='kpi-lbl'>{lbl_projecao}</div><div class='kpi-val' style='color:#6366f1;'>{projecao_mes}</div></div><div style='font-size:1.8rem;'>🔮</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

        def renderizar_box_grafico_setor(nome_s, icone_s, meta_valor):
            sub_s = df_ano_f[df_ano_f["Setor"] == nome_s].groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
            sub_s["Mes_Abrev"] = sub_s["Mes"].map(MAPA_MES_ABREV)
            sub_s["Meta"] = meta_valor
            tot_s = int(sub_s["Quantidade_Quebras"].sum())
            med_s_card = round(tot_s / div_meses, 1)
            cor_s = CORES_SETORES.get(nome_s, "#3b82f6")
            
            dentro_meta = med_s_card <= meta_valor
            bg_badge = "#dcfce7" if dentro_meta else "#fee2e2"
            cor_badge = "#166534" if dentro_meta else "#991b1b"
            txt_badge = "🟢 MÉDIA OK" if dentro_meta else "🔴 ACIMA DA META"

            html_top_graf = (
                f"<div style='display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid #f1f5f9; padding-bottom:10px; margin-bottom:10px;'>"
                f"<div><div style='font-weight:900; font-size:1.05rem; color:#0f172a;'>{icone_s} {nome_s}</div>"
                f"<div style='font-size:0.75rem; color:#64748b; font-weight:700;'>Meta Mensal: <b>{meta_valor} un.</b> &nbsp;•&nbsp; Média Atual: <b style='color:{cor_s};'>{med_s_card}/mês</b></div></div>"
                f"<div style='text-align:right;'><span style='background:{bg_badge}; color:{cor_badge}; padding:3px 10px; border-radius:20px; font-size:0.7rem; font-weight:900;'>{txt_badge}</span>"
                f"<div style='font-size:0.95rem; font-weight:900; color:{cor_s}; margin-top:3px;'>Total: {tot_s} fusos</div></div></div>"
            )
            st.markdown(html_top_graf, unsafe_allow_html=True)

            bars = alt.Chart(sub_s).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, size=24).encode(
                x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                color=alt.condition(f"datum.Quantidade_Quebras > {meta_valor}", alt.value("#ef4444"), alt.value(cor_s)),
                tooltip=["Mes", "Quantidade_Quebras", "Meta"]
            )
            txt = alt.Chart(sub_s).mark_text(dy=-8, fontSize=11, fontWeight=800).encode(
                x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                y=alt.Y("Quantidade_Quebras:Q"),
                text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")),
                color=alt.condition(f"datum.Quantidade_Quebras > {meta_valor}", alt.value("#dc2626"), alt.value("#0f172a"))
            )
            linha_meta = alt.Chart(pd.DataFrame({'Meta': [meta_valor]})).mark_rule(color='#ef4444', strokeWidth=2, strokeDash=[5, 4]).encode(y='Meta:Q')
            st.altair_chart((bars + txt + linha_meta).properties(height=230), use_container_width=True)

        cg1, cg2 = st.columns(2)
        with cg1:
            with st.container(border=True):
                renderizar_box_grafico_setor("Setor A", "🏭", 90)
            with st.container(border=True):
                renderizar_box_grafico_setor("Setor Látex", "🌿", 45)
        with cg2:
            with st.container(border=True):
                renderizar_box_grafico_setor("Setor B", "🏭", 90)
            with st.container(border=True):
                renderizar_box_grafico_setor("Setor Menegatto", "⚙️", 45)

        # KANBAN GERAL DO PLANO DE AÇÃO
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        st.markdown("<div style='background:linear-gradient(90deg, #0f172a, #1e293b); color:#ffffff; padding:12px 20px; border-radius:12px; font-size:1.05rem; font-weight:900; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;'><span>📋 Plano de Ação Operacional (Visão Geral da Fábrica)</span><span style='font-size:0.78rem; color:#94a3b8; font-weight:700;'>Acompanhamento de Trocas e Reparos em Fusos</span></div>", unsafe_allow_html=True)
        
        df_p_g = df_trabalhos_fusos[df_trabalhos_fusos["Status"].astype(str).str.upper() == "PENDENTE"] if not df_trabalhos_fusos.empty else pd.DataFrame()
        df_f_g = df_trabalhos_fusos[df_trabalhos_fusos["Status"].astype(str).str.upper() == "FINALIZADO"] if not df_trabalhos_fusos.empty else pd.DataFrame()

        c_pend_g, c_fin_g = st.columns(2)
        with c_pend_g:
            html_p_full_g = (
                "<div style='background:#ffffff; border:1px solid #fde68a; border-top:6px solid #f59e0b; padding:18px; border-radius:14px; box-shadow:0 6px 15px -3px rgba(245, 158, 11, 0.08); min-height:280px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:12px; margin-bottom:14px; border-bottom:2px dashed #fef3c7;'>"
                "<div style='font-size:1.1rem; font-weight:900; color:#b45309;'>⏳ Trabalhos Pendentes</div>"
                f"<span style='background:#f59e0b; color:#ffffff; padding:3px 12px; border-radius:20px; font-size:0.78rem; font-weight:900;'>{len(df_p_g)} ordens</span></div>"
            )
            if not df_p_g.empty:
                html_p_full_g += "<div style='max-height:380px; overflow-y:auto; padding-right:6px;'>"
                for _, r in df_p_g.iterrows():
                    maq, desc_raw, qtd, setor = r.get("Maquina", "N/I"), r.get("Descricao", ""), r.get("Quantidade", 0), r.get("Setor", "N/I")
                    desc_html = html_lista_ponto(desc_raw, cor_texto="#334155", font_weight="600")
                    cor_s_pill = CORES_SETORES.get(setor, "#475569")
                    html_p_full_g += (
                        "<div style='background:linear-gradient(to right, #fffbeb, #ffffff); padding:14px 16px; border:1px solid #fde68a; border-left:5px solid #f59e0b; border-radius:12px; margin-bottom:10px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>"
                        "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'>"
                        f"<div style='display:flex; align-items:center; gap:8px;'><span style='background:#0f172a; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:900; font-size:0.95rem;'>⚙️ {maq}</span><span style='font-size:0.72rem; background:#f8fafc; border:1px solid #cbd5e1; padding:2px 8px; border-radius:6px; color:{cor_s_pill}; font-weight:800;'>📍 {setor}</span></div>"
                        f"<span style='background:#fef3c7; border:1px solid #fde68a; color:#b45309; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:900;'>🔩 {qtd} UN.</span></div>"
                        f"<div>{desc_html}</div></div>"
                    )
                html_p_full_g += "</div>"
            else:
                html_p_full_g += "<div style='text-align:center; padding:35px 15px; color:#b45309; font-weight:700; background:#fffbeb; border-radius:10px;'>🎉 Nenhuma atividade pendente na fábrica!</div>"
            html_p_full_g += "</div>"
            st.markdown(html_p_full_g, unsafe_allow_html=True)
            
        with c_fin_g:
            html_f_full_g = (
                "<div style='background:#ffffff; border:1px solid #bbf7d0; border-top:6px solid #10b981; padding:18px; border-radius:14px; box-shadow:0 6px 15px -3px rgba(16, 185, 129, 0.08); min-height:280px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:12px; margin-bottom:14px; border-bottom:2px dashed #dcfce7;'>"
                "<div style='font-size:1.1rem; font-weight:900; color:#047857;'>✅ Últimos Finalizados</div>"
                f"<span style='background:#10b981; color:#ffffff; padding:3px 12px; border-radius:20px; font-size:0.78rem; font-weight:900;'>{len(df_f_g)} concluídos</span></div>"
            )
            if not df_f_g.empty:
                df_f_g_show = df_f_g.sort_values(by="Data", ascending=False).head(15) if "Data" in df_f_g.columns else df_f_g.head(15)
                html_f_full_g += "<div style='max-height:380px; overflow-y:auto; padding-right:6px;'>"
                for _, r in df_f_g_show.iterrows():
                    maq, desc_raw, qtd, setor = r.get("Maquina", "N/I"), r.get("Descricao", ""), r.get("Quantidade", 0), r.get("Setor", "N/I")
                    dt_val = r.get("Data")
                    dt_str = dt_val.strftime("%d/%m/%Y") if pd.notna(dt_val) else "Sem Data"
                    desc_html = html_lista_ponto(desc_raw, cor_texto="#334155", font_weight="600")
                    html_f_full_g += (
                        "<div style='background:linear-gradient(to right, #f0fdf4, #ffffff); padding:14px 16px; border:1px solid #bbf7d0; border-left:5px solid #10b981; border-radius:12px; margin-bottom:10px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>"
                        "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'>"
                        f"<div style='display:flex; align-items:center; gap:8px;'><span style='background:#065f46; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:900; font-size:0.95rem;'>✓ {maq}</span><span style='font-size:0.72rem; background:#ffffff; border:1px solid #cbd5e1; padding:2px 8px; border-radius:6px; color:#475569; font-weight:800;'>📍 {setor}</span></div>"
                        f"<div style='display:flex; gap:6px;'><span style='background:#dcfce7; color:#166534; padding:3px 9px; border-radius:20px; font-size:0.73rem; font-weight:900;'>🔩 {qtd} un.</span><span style='background:#f1f5f9; border:1px solid #e2e8f0; color:#475569; padding:3px 9px; border-radius:20px; font-size:0.73rem; font-weight:800;'>📅 {dt_str}</span></div></div>"
                        f"<div>{desc_html}</div></div>"
                    )
                html_f_full_g += "</div>"
            else:
                html_f_full_g += "<div style='text-align:center; padding:35px 15px; color:#047857; font-weight:700; background:#f0fdf4; border-radius:10px;'>Nenhuma atividade finalizada registrada.</div>"
            html_f_full_g += "</div>"
            st.markdown(html_f_full_g, unsafe_allow_html=True)

        # OBSERVAÇÕES GERAIS
        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:1.1rem; font-weight:900; color:#1e293b; margin-bottom:10px;'>📌 Quadro de Avisos e Observações Técnicas (Geral)</div>", unsafe_allow_html=True)
        
        if not df_observacoes_fusos.empty:
            html_info = "<div style='display:flex; gap:15px; flex-wrap:wrap; margin-bottom:15px;'>"
            for _, r_info in df_observacoes_fusos.iterrows():
                titulo_info = r_info.get("Titulo")
                if pd.isna(titulo_info) or str(titulo_info).strip() == "": titulo_info = "Observação Técnica"
                setor_info = r_info.get("Setor", "Geral")
                if pd.isna(setor_info) or str(setor_info).strip() == "": setor_info = "Geral"
                desc_info_raw = r_info.get("Descricao", "")
                desc_info_html = html_lista_ponto(desc_info_raw, cor_texto="#334155", font_weight="600", font_size="0.92rem")
                cor_borda_obs = CORES_SETORES.get(setor_info, "#3b82f6")
                
                html_info += (
                    f"<div style='flex:1; min-width:300px; background:#ffffff; border:1px solid #e2e8f0; border-left:6px solid {cor_borda_obs}; border-radius:12px; padding:16px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.04);'>"
                    "<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'>"
                    f"<div style='color:#0f172a; font-weight:900; font-size:1.02rem;'>📌 {titulo_info}</div>"
                    f"<span style='background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe; padding:2px 8px; border-radius:6px; font-size:0.7rem; font-weight:800; text-transform:uppercase;'>📍 {setor_info}</span></div>"
                    f"<div>{desc_info_html}</div></div>"
                )
            html_info += "</div>"
            st.markdown(html_info, unsafe_allow_html=True)
        else:
            st.markdown("<div style='background:#f8fafc; border:1px solid #e2e8f0; color:#64748b; padding:15px; border-radius:10px; font-weight:600;'>Nenhuma observação registrada na segunda aba do arquivo Trabalhos Fusos.</div>", unsafe_allow_html=True)

    # =========================================================================
    # ABA 2: VISÃO INDIVIDUAL POR SETOR (COM MAPA DE CALOR PREMIUM)
    # =========================================================================
    else:
        s_ativo = st.session_state.aba_setor_fuso
        meta_s_ativo = METAS_SETOR_FUSO.get(s_ativo, 90)
        cor_setor_atual = CORES_SETORES.get(s_ativo, "#3b82f6")
        df_sa = df_ano_f[df_ano_f["Setor"] == s_ativo].copy()

        tot_s = int(df_sa["Quantidade_Quebras"].sum()) if not df_sa.empty else 0
        maqs_setor = obter_maquinas_setor(s_ativo, df_fusos)
        qtd_maqs_s = len(maqs_setor)
        med_s = round(tot_s / div_meses, 1)
        ult_mes_s, top_maq_s, qtd_top_s, quebras_ult_mes = "Nenhum", "Nenhuma", 0, 0

        if not df_sa.empty and tot_s > 0:
            df_reais_s = df_sa[df_sa["Quantidade_Quebras"] > 0]
            for m_teste in reversed(LISTA_MESES_PUROS):
                sub_m = df_reais_s[df_reais_s["Mes"] == m_teste]
                if not sub_m.empty and sub_m["Quantidade_Quebras"].sum() > 0:
                    ult_mes_s = m_teste
                    quebras_ult_mes = int(sub_m["Quantidade_Quebras"].sum())
                    agrup_m = sub_m.groupby("Maquina_TAG")["Quantidade_Quebras"].sum().sort_values(ascending=False)
                    top_maq_s = agrup_m.index[0]
                    qtd_top_s = int(agrup_m.iloc[0])
                    break

        quebras_por_maq = round(quebras_ult_mes / qtd_maqs_s, 1) if (qtd_maqs_s > 0 and quebras_ult_mes > 0) else 0.0

        projecao_mes_s = "-"
        lbl_projecao_s = "Projeção Mês"
        if int(ano_f) == ano_atual_ref:
            nome_mes_atual = LISTA_MESES_PUROS[mes_atual_num_ref - 1]
            df_mes_atual_s = df_sa[df_sa["Mes"] == nome_mes_atual]
            quebras_mes_ate_hoje_s = int(df_mes_atual_s["Quantidade_Quebras"].sum()) if not df_mes_atual_s.empty else 0
            dias_passados = data_hoje_ref.day
            dias_totais_mes = calendar.monthrange(ano_atual_ref, mes_atual_num_ref)[1]
            if dias_passados > 0:
                projecao_mes_s = int(round((quebras_mes_ate_hoje_s / dias_passados) * dias_totais_mes, 0))
            lbl_projecao_s = f"Projeção ({MAPA_MES_ABREV[nome_mes_atual]})"

        ks1, ks2, ks3, ks4, ks5 = st.columns(5)
        ks1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total ({s_ativo})</div><div class='kpi-val'>{tot_s} <span style='font-size:0.75rem; color:#94a3b8;'>Meta: {meta_s_ativo}/m</span></div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        ks2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Média Mensal ({desc_divisor})</div><div class='kpi-val' style='color:#059669;'>{med_s}</div></div><div style='font-size:1.8rem;'>📅</div></div>", unsafe_allow_html=True)
        ks3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Tx Falha/Máq ({ult_mes_s})</div><div class='kpi-val' style='color:#d97706;'>{quebras_por_maq}</div></div><div style='font-size:1.8rem;'>⚙️</div></div>", unsafe_allow_html=True)
        ks4.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Maior Ofensor ({ult_mes_s})</div><div class='kpi-val' style='color:#dc2626; font-size:1.15rem;'>{top_maq_s} ({qtd_top_s})</div></div><div style='font-size:1.8rem;'>⚠️</div></div>", unsafe_allow_html=True)
        ks5.markdown(f"<div class='card-kpi-bonito' style='border-left:6px solid #6366f1;'><div><div class='kpi-lbl'>{lbl_projecao_s}</div><div class='kpi-val' style='color:#6366f1;'>{projecao_mes_s}</div></div><div style='font-size:1.8rem;'>🔮</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

        c_evol, c_rosca = st.columns([1.6, 1.4])
        with c_evol:
            with st.container(border=True):
                st.markdown(f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:2px solid #f1f5f9; padding-bottom:8px;'><span style='font-size:1.05rem; font-weight:900; color:#1e293b;'>📊 Evolução Mensal vs Meta ({ano_f})</span><span style='background:#fef2f2; color:#dc2626; border:1px solid #fecaca; padding:2px 10px; border-radius:20px; font-size:0.75rem; font-weight:800;'>--- Meta: {meta_s_ativo} fusos</span></div>", unsafe_allow_html=True)
                df_evol = df_sa.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                df_evol["Mes_Abrev"] = df_evol["Mes"].map(MAPA_MES_ABREV)
                df_evol["Meta"] = meta_s_ativo

                barras_s = alt.Chart(df_evol).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, size=26).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                    color=alt.condition(f"datum.Quantidade_Quebras > {meta_s_ativo}", alt.value("#ef4444"), alt.value(cor_setor_atual)),
                    tooltip=["Mes", "Quantidade_Quebras", "Meta"]
                )
                rotulos_s = alt.Chart(df_evol).mark_text(dy=-8, fontSize=12, fontWeight=800).encode(
                    x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV),
                    y=alt.Y("Quantidade_Quebras:Q"),
                    text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value("")),
                    color=alt.condition(f"datum.Quantidade_Quebras > {meta_s_ativo}", alt.value("#dc2626"), alt.value("#0f172a"))
                )
                regra_meta = alt.Chart(pd.DataFrame({'Meta': [meta_s_ativo]})).mark_rule(color='#ef4444', strokeWidth=2, strokeDash=[5, 4]).encode(y='Meta:Q')
                st.altair_chart((barras_s + rotulos_s + regra_meta).properties(height=330), use_container_width=True)

        with c_rosca:
            with st.container(border=True):
                cr_col1, cr_col2 = st.columns([2.0, 1.4])
                with cr_col1:
                    st.markdown("<div style='font-size:1.05rem; font-weight:900; color:#1e293b; margin-top:4px;'>🍩 Quebras por Modelo</div>", unsafe_allow_html=True)
                with cr_col2:
                    mes_filtro_rosca_fuso = st.selectbox("Mês:", ["Todos os Meses"] + LISTA_MESES_PUROS, key=f"sel_mes_rosca_fuso_{s_ativo}", label_visibility="collapsed")

                df_sa_rosca = df_sa.copy()
                if mes_filtro_rosca_fuso != "Todos os Meses":
                    df_sa_rosca = df_sa_rosca[df_sa_rosca["Mes"] == mes_filtro_rosca_fuso]
                if not df_sa_rosca.empty and df_sa_rosca["Quantidade_Quebras"].sum() > 0:
                    df_tipos = df_sa_rosca[df_sa_rosca["Quantidade_Quebras"] > 0].groupby("Tipo_Fuso")["Quantidade_Quebras"].sum().reset_index()
                    tipos_lista_fuso = df_tipos["Tipo_Fuso"].tolist()
                    cores_lidas_fuso = [CORES_FUSOS.get(t, PALETA_DINAMICA[i % len(PALETA_DINAMICA)]) for i, t in enumerate(tipos_lista_fuso)]
                    
                    arc_fuso = alt.Chart(df_tipos).mark_arc(innerRadius=58, outerRadius=100, stroke="#ffffff", strokeWidth=2).encode(
                        theta=alt.Theta("Quantidade_Quebras:Q", stack=True),
                        color=alt.Color("Tipo_Fuso:N", scale=alt.Scale(domain=tipos_lista_fuso, range=cores_lidas_fuso), legend=None),
                        tooltip=["Tipo_Fuso", "Quantidade_Quebras"]
                    )
                    st.altair_chart(arc_fuso.properties(height=215), use_container_width=True)
                    
                    html_bf = "<div style='display:flex; flex-wrap:wrap; gap:8px; justify-content:center; margin-top:8px;'>"
                    for _, r in df_tipos.iterrows():
                        cor = CORES_FUSOS.get(r['Tipo_Fuso'], '#0f172a')
                        html_bf += f"<div style='background:#f8fafc; border:1px solid #e2e8f0; border-top:3px solid {cor}; border-radius:8px; padding:6px 12px; text-align:center; flex:1; min-width:75px;'><div style='font-size:0.72rem; color:#64748b; font-weight:800;'>{r['Tipo_Fuso']}</div><div style='font-size:1.1rem; color:{cor}; font-weight:900;'>{r['Quantidade_Quebras']}</div></div>"
                    html_bf += "</div>"
                    st.markdown(html_bf, unsafe_allow_html=True)
                else:
                    st.markdown("<div style='height:285px; display:flex; align-items:center; justify-content:center; color:#64748b; font-weight:600;'>Sem quebras registradas para o filtro.</div>", unsafe_allow_html=True)

        # KANBAN DO SETOR ATIVO
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        df_trab_setor = df_trabalhos_fusos[df_trabalhos_fusos["Setor"] == s_ativo].copy() if not df_trabalhos_fusos.empty else pd.DataFrame()
        df_p_s = df_trab_setor[df_trab_setor["Status"].astype(str).str.upper() == "PENDENTE"] if not df_trab_setor.empty else pd.DataFrame()
        df_f_s = df_trab_setor[df_trab_setor["Status"].astype(str).str.upper() == "FINALIZADO"] if not df_trab_setor.empty else pd.DataFrame()

        c_pend, c_fin = st.columns(2)
        with c_pend:
            html_p_full = (
                "<div style='background:#ffffff; border:1px solid #fde68a; border-top:6px solid #f59e0b; padding:18px; border-radius:14px; box-shadow:0 6px 15px -3px rgba(245, 158, 11, 0.08); min-height:250px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:12px; border-bottom:2px dashed #fef3c7;'>"
                f"<div style='font-size:1.05rem; font-weight:900; color:#b45309;'>⏳ Plano de Ação: Pendentes ({s_ativo})</div>"
                f"<span style='background:#f59e0b; color:#ffffff; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:900;'>{len(df_p_s)}</span></div>"
            )
            if not df_p_s.empty:
                html_p_full += "<div style='max-height:340px; overflow-y:auto; padding-right:5px;'>"
                for _, r in df_p_s.iterrows():
                    maq, desc_raw, qtd = r.get("Maquina", "N/I"), r.get("Descricao", ""), r.get("Quantidade", 0)
                    desc_html = html_lista_ponto(desc_raw, cor_texto="#334155", font_weight="600")
                    html_p_full += (
                        "<div style='background:linear-gradient(to right, #fffbeb, #ffffff); padding:14px; border:1px solid #fde68a; border-left:5px solid #f59e0b; border-radius:12px; margin-bottom:10px;'>"
                        f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;'><span style='background:#0f172a; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:900; font-size:0.95rem;'>⚙️ {maq}</span><span style='background:#fef3c7; border:1px solid #fde68a; color:#d97706; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:900;'>🔩 {qtd} UN.</span></div>"
                        f"<div>{desc_html}</div></div>"
                    )
                html_p_full += "</div>"
            else:
                html_p_full += "<div style='color:#b45309; font-size:0.9rem; font-weight:700; text-align:center; padding:30px 10px; background:#fffbeb; border-radius:10px;'>🎉 Nenhuma pendência no setor!</div>"
            html_p_full += "</div>"
            st.markdown(html_p_full, unsafe_allow_html=True)
            
        with c_fin:
            html_f_full = (
                "<div style='background:#ffffff; border:1px solid #bbf7d0; border-top:6px solid #10b981; padding:18px; border-radius:14px; box-shadow:0 6px 15px -3px rgba(16, 185, 129, 0.08); min-height:250px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:12px; border-bottom:2px dashed #dcfce7;'>"
                f"<div style='font-size:1.05rem; font-weight:900; color:#047857;'>✅ Plano de Ação: Finalizados ({s_ativo})</div>"
                f"<span style='background:#10b981; color:#ffffff; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:900;'>{len(df_f_s)}</span></div>"
            )
            if not df_f_s.empty:
                df_f_s_show = df_f_s.sort_values(by="Data", ascending=False) if "Data" in df_f_s.columns else df_f_s
                html_f_full += "<div style='max-height:340px; overflow-y:auto; padding-right:5px;'>"
                for _, r in df_f_s_show.iterrows():
                    maq, desc_raw, qtd = r.get("Maquina", "N/I"), r.get("Descricao", ""), r.get("Quantidade", 0)
                    dt_val = r.get("Data")
                    dt_str = dt_val.strftime("%d/%m/%Y") if pd.notna(dt_val) else "Sem Data"
                    desc_html = html_lista_ponto(desc_raw, cor_texto="#334155", font_weight="600")
                    html_f_full += (
                        "<div style='background:linear-gradient(to right, #f0fdf4, #ffffff); padding:14px; border:1px solid #bbf7d0; border-left:5px solid #10b981; border-radius:12px; margin-bottom:10px;'>"
                        f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;'><span style='background:#065f46; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:900; font-size:0.95rem;'>✓ {maq}</span><div style='display:flex; gap:6px;'><span style='background:#dcfce7; color:#166534; padding:3px 9px; border-radius:20px; font-size:0.73rem; font-weight:900;'>🔩 {qtd} un.</span><span style='background:#f1f5f9; border:1px solid #e2e8f0; color:#475569; padding:3px 9px; border-radius:20px; font-size:0.73rem; font-weight:800;'>📅 {dt_str}</span></div></div>"
                        f"<div>{desc_html}</div></div>"
                    )
                html_f_full += "</div>"
            else:
                html_f_full += "<div style='color:#047857; font-size:0.9rem; font-weight:700; text-align:center; padding:30px 10px; background:#f0fdf4; border-radius:10px;'>Nenhuma atividade finalizada.</div>"
            html_f_full += "</div>"
            st.markdown(html_f_full, unsafe_allow_html=True)

        # OBSERVAÇÕES DO SETOR
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        df_obs_setor = df_observacoes_fusos[df_observacoes_fusos["Setor"] == s_ativo].copy() if not df_observacoes_fusos.empty else pd.DataFrame()
        
        if not df_obs_setor.empty:
            st.markdown(f"<div style='font-size:1.05rem; font-weight:900; color:#1e293b; margin-bottom:10px;'>📌 Observações Técnicas ({s_ativo})</div>", unsafe_allow_html=True)
            html_info = "<div style='display:flex; gap:15px; flex-wrap:wrap; margin-bottom:14px;'>"
            for _, r_info in df_obs_setor.iterrows():
                titulo_info = r_info.get("Titulo")
                if pd.isna(titulo_info) or str(titulo_info).strip() == "": titulo_info = "Observação"
                desc_info_raw = r_info.get("Descricao", "")
                desc_info_html = html_lista_ponto(desc_info_raw, cor_texto="#334155", font_weight="600", font_size="0.92rem")
                html_info += (
                    f"<div style='flex:1; min-width:300px; background:#ffffff; border:1px solid #e2e8f0; border-left:5px solid {cor_setor_atual}; border-radius:12px; padding:15px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>"
                    f"<div style='margin-bottom:8px; color:#0f172a; font-weight:900; font-size:1rem;'>📌 {titulo_info}</div>"
                    f"<div>{desc_info_html}</div></div>"
                )
            html_info += "</div>"
            st.markdown(html_info, unsafe_allow_html=True)

        # MAPA DE CALOR OPERACIONAL (HEATMAP)
        with st.container(border=True):
            html_hdr_heat = (
                "<div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:12px; padding-bottom:10px; border-bottom:2px solid #f1f5f9;'>"
                f"<div><span style='font-size:1.1rem; font-weight:900; color:#0f172a;'>🔥 Mapa de Calor de Quebras por Máquina ({s_ativo} • {ano_f})</span></div>"
                "<div style='display:flex; gap:8px; align-items:center; font-size:0.75rem; font-weight:800;'>"
                "<span style='background:#dcfce7; color:#166534; padding:3px 10px; border-radius:6px; border:1px solid #bbf7d0;'>0 a 2: Normal</span>"
                "<span style='background:#fef08a; color:#854d0e; padding:3px 10px; border-radius:6px; border:1px solid #fde047;'>3 a 7: Atenção</span>"
                "<span style='background:#fed7aa; color:#9a3412; padding:3px 10px; border-radius:6px; border:1px solid #fdba74;'>8 a 14: Alerta</span>"
                "<span style='background:#ef4444; color:#ffffff; padding:3px 10px; border-radius:6px;'>≥ 15: Crítico</span>"
                "</div></div>"
            )
            st.markdown(html_hdr_heat, unsafe_allow_html=True)
            
            idx_grid = pd.MultiIndex.from_product([maqs_setor, LISTA_MESES_PUROS], names=["MAQ", "Mes"]).to_frame().reset_index(drop=True)
            idx_grid["MES"] = idx_grid["Mes"].map(MAPA_MES_ABREV)
            
            mapa_tipos_arquivo = carregar_mapa_tipos_fuso()
            fuso_por_maq = {}
            for m in maqs_setor:
                if m in mapa_tipos_arquivo:
                    fuso_por_maq[m] = mapa_tipos_arquivo[m]
                else:
                    df_maq = df_sa[df_sa["Maquina_TAG"] == m]
                    if not df_maq.empty: fuso_por_maq[m] = df_maq["Tipo_Fuso"].mode()[0]
                    else: fuso_por_maq[m] = obter_fuso_padrao(m, s_ativo)
            
            agrup_c = df_sa.groupby(["Maquina_TAG", "Mes"]).agg(
                Quantidade_Quebras=("Quantidade_Quebras", "sum"),
                Tipo_Fuso=("Tipo_Fuso", lambda x: " / ".join(x.dropna().unique()))
            ).reset_index()

            m_calor = pd.merge(idx_grid, agrup_c, left_on=["MAQ", "Mes"], right_on=["Maquina_TAG", "Mes"], how="left")
            m_calor["Quantidade_Quebras"] = m_calor["Quantidade_Quebras"].fillna(0)
            m_calor["Tipo_Fuso"] = m_calor["Tipo_Fuso"].fillna("")

            def formatar_rotulo(row):
                tipo = row["Tipo_Fuso"] if row["Tipo_Fuso"] else fuso_por_maq.get(row["MAQ"], "N/I")
                return f"{int(row['Quantidade_Quebras'])} ({tipo})"

            m_calor["Tipo_Fuso"] = m_calor.apply(lambda r: r["Tipo_Fuso"] if r["Tipo_Fuso"] else fuso_por_maq.get(r["MAQ"], "N/I"), axis=1)
            m_calor["Rotulo"] = m_calor.apply(formatar_rotulo, axis=1)

            mes_atual_abrev = MAPA_MES_ABREV[LISTA_MESES_PUROS[mes_atual_num_ref - 1]]
            m_calor["MES_DISPLAY"] = m_calor["MES"].apply(lambda x: f"► {x} ◄" if x == mes_atual_abrev else x)
            ordem_meses_display = [f"► {x} ◄" if x == mes_atual_abrev else x for x in ORDEM_MESES_ABREV]

            rect = alt.Chart(m_calor).mark_rect(cornerRadius=4).encode(
                x=alt.X("MES_DISPLAY:N", sort=ordem_meses_display, title=None, axis=alt.Axis(orient="top", labelAngle=0, labelFontSize=12, labelFontWeight=800)),
                y=alt.Y("MAQ:N", sort=maqs_setor, title=None, axis=alt.Axis(labelFontSize=12, labelFontWeight=800)),
                color=alt.Color("Quantidade_Quebras:Q", scale=alt.Scale(domain=[0, 3, 8, 15], range=["#dcfce7", "#fef08a", "#f97316", "#ef4444"]), legend=None),
                stroke=alt.value("#ffffff"),
                strokeWidth=alt.value(2),
                tooltip=["MAQ", "MES", "Quantidade_Quebras", "Tipo_Fuso"]
            )
            
            txt = alt.Chart(m_calor).mark_text(baseline="middle", fontSize=11, fontWeight=800).encode(
                x=alt.X("MES_DISPLAY:N", sort=ordem_meses_display), 
                y=alt.Y("MAQ:N", sort=maqs_setor),
                text=alt.Text("Rotulo:N"),
                color=alt.condition("datum.Quantidade_Quebras >= 10", alt.value("#ffffff"), alt.value("#0f172a"))
            )
            
            st.altair_chart((rect + txt).properties(height=max(280, len(maqs_setor) * 30)), use_container_width=True)





# ==========================================
# REGISTRO DE CHECKLIST (DESIGN PREMIUM)
# ==========================================
elif st.session_state.pagina_atual == "Registro Checklist":
    # Inicializa estado da especialidade se não existir
    if "esp_check_sel" not in st.session_state:
        st.session_state.esp_check_sel = "⚙️ Mecânica"

    # Cabeçalho e Seletor Estilo Abas Modernas
    c_tit, c_esp1, c_esp2 = st.columns([6, 2, 2])
    with c_tit:
        st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>📋 Registros de Checklist</h2>", unsafe_allow_html=True)
        st.markdown("<div style='color:#64748b; margin-bottom:15px; font-weight:500;'>Mapa visual de conformidade e dossiê detalhado de inspeções por equipamento.</div>", unsafe_allow_html=True)
    
    with c_esp1:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        if st.button("⚙️ Mecânica", key="btn_aba_mec", use_container_width=True, type="primary" if st.session_state.esp_check_sel == "⚙️ Mecânica" else "secondary"):
            st.session_state.esp_check_sel = "⚙️ Mecânica"
            st.rerun()
    with c_esp2:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        if st.button("⚡ Elétrica", key="btn_aba_ele", use_container_width=True, type="primary" if st.session_state.esp_check_sel == "⚡ Elétrica" else "secondary"):
            st.session_state.esp_check_sel = "⚡ Elétrica"
            st.rerun()

    tipo_checklist = st.session_state.esp_check_sel

    # Define qual DataFrame e paleta de cores usar
    if tipo_checklist == "⚙️ Mecânica":
        df_check_view = df_checklist_mecanica.copy()
        cor_tema = "#f59e0b"
        grad_tema = "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)"
        icone_tema = "⚙️"
    else:
        df_check_view = df_checklist_eletrica.copy()
        cor_tema = "#3b82f6"
        grad_tema = "linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%)"
        icone_tema = "⚡"

    todas_maquinas = []
    for s_nome in DICIONARIO_SETORES.keys():
        todas_maquinas.extend(obter_maquinas_setor(s_nome, df_fusos))

    # =========================================================================
    # MODO 1: DOSSIÊ INDIVIDUAL DA MÁQUINA (QUANDO CLICADA)
    # =========================================================================
    if st.session_state.get("maq_clicada_check") and st.session_state.maq_clicada_check in todas_maquinas:
        maq_sel = st.session_state.maq_clicada_check
        setor_maq_sel = mapa_setor_maquina.get(maq_sel, "Setor Industrial")

        # Cabeçalho com o botão Voltar e as opções de Impressão
        c_voltar, c_print_tog, c_print_btn, _ = st.columns([2.5, 3.0, 3.0, 3.5])
        with c_voltar:
            if st.button("⬅️ Voltar ao Mapa Geral", key="btn_voltar_grid_check", use_container_width=True):
                st.session_state.maq_clicada_check = None
                st.rerun()
        with c_print_tog:
            st.markdown("<div style='margin-top: 5px;'></div>", unsafe_allow_html=True)
            modo_imp_maq = st.toggle("🖨️ Modo Impressão", key="tgl_imp_maq", help="Oculta os concluídos e prepara para impressão.")
        
        with c_print_btn:
            if modo_imp_maq:
                st.markdown(
                    '<a href="javascript:window.print()" style="text-decoration:none;">'
                    '<div style="background-color: #0f172a; color: white; padding: 6px 12px; border-radius: 8px; text-align: center; font-weight: 800; font-size: 0.85rem; border: 1px solid #1e293b; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: 4px; transition: 0.2s;">'
                    '🖨️ Imprimir Pendências</div></a>', 
                    unsafe_allow_html=True
                )
                # CSS que limpa a tela SOMENTE na hora da impressão do navegador
                st.markdown("""
                <style>
                @media print {
                    header[data-testid='stHeader'], [data-testid='stSidebar'], .stToggle, button { display: none !important; }
                    .block-container { padding: 1rem !important; max-width: 100% !important; }
                    * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
                    a { text-decoration: none !important; }
                }
                </style>
                """, unsafe_allow_html=True)

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

        df_maq_check = df_check_view[df_check_view["Maquina"] == maq_sel].copy()

        if not df_maq_check.empty:
            df_maq_check = df_maq_check.sort_values(by="Data_Ultima", ascending=False, na_position='first')
            df_pend = df_maq_check[df_maq_check["Data_Ultima"].isna()]
            df_ok = df_maq_check[df_maq_check["Data_Ultima"].notna()]
            
            tot_m = len(df_maq_check)
            ok_m = len(df_ok)
            pend_m = len(df_pend)
            pct_concluido = int(round((ok_m / tot_m) * 100)) if tot_m > 0 else 0

            cor_status_hero = "#10b981" if pend_m == 0 else "#ef4444"
            lbl_status_hero = "CONFORME (EM DIA)" if pend_m == 0 else f"{pend_m} PENDÊNCIA(S) ATIVA(S)"

            # HERO BANNER COM BARRA DE PROGRESSO
            html_hero_chk = (
                f"<div style='background: linear-gradient(110deg, #0f172a 0%, #1e293b 100%); padding: 26px 32px; border-radius: 18px; margin-bottom: 24px; box-shadow: 0 12px 25px -5px rgba(15, 23, 42, 0.25); border-left: 8px solid {cor_status_hero}; position: relative; overflow: hidden;'>"
                "<div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 20px;'>"
                "<div>"
                "<div style='display: flex; align-items: center; gap: 10px; margin-bottom: 6px;'>"
                f"<span style='background: {grad_tema}; color: #ffffff; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; box-shadow: 0 2px 6px rgba(0,0,0,0.2);'>{tipo_checklist}</span>"
                f"<span style='color: #94a3b8; font-size: 0.85rem; font-weight: 700;'>📍 {setor_maq_sel}</span>"
                "</div>"
                f"<div style='color: #ffffff; font-size: 2.5rem; font-weight: 900; line-height: 1.1; letter-spacing: -0.5px;'>Ativo {maq_sel}</div>"
                "</div>"
                "<div style='background: rgba(255,255,255,0.06); backdrop-filter: blur(8px); padding: 16px 24px; border-radius: 14px; border: 1px solid rgba(255,255,255,0.1); min-width: 240px; text-align: right;'>"
                f"<div style='color: {cor_status_hero}; font-size: 1.05rem; font-weight: 900; letter-spacing: 0.5px;'>{lbl_status_hero}</div>"
                "<div style='display: flex; justify-content: space-between; font-size: 0.8rem; color: #cbd5e1; font-weight: 700; margin-top: 10px; margin-bottom: 5px;'>"
                "<span>Progresso do Checklist</span>"
                f"<span style='color: #ffffff; font-weight: 900;'>{pct_concluido}% ({ok_m}/{tot_m})</span>"
                "</div>"
                "<div style='width: 100%; background: rgba(255,255,255,0.15); height: 8px; border-radius: 99px; overflow: hidden;'>"
                f"<div style='width: {pct_concluido}%; background: {cor_status_hero}; height: 100%; border-radius: 99px; box-shadow: 0 0 10px {cor_status_hero};'></div>"
                "</div>"
                "</div></div></div>"
            )
            st.markdown(html_hero_chk, unsafe_allow_html=True)

            # ===========================================================
            # LÓGICA DE COLUNAS: 
            # Se for Impressão, exibe apenas a coluna de Pendências 100% larga
            # ===========================================================
            if modo_imp_maq:
                c_pend = st.container()
            else:
                c_pend, c_ok = st.columns(2)

            with c_pend:
                html_p = (
                    "<div style='background: #ffffff; border: 1px solid #fecaca; border-top: 6px solid #ef4444; padding: 22px; border-radius: 16px; box-shadow: 0 10px 20px -5px rgba(239, 68, 68, 0.08); min-height: 340px;'>"
                    "<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; padding-bottom: 12px; border-bottom: 2px dashed #fee2e2;'>"
                    "<div style='display: flex; align-items: center; gap: 10px;'>"
                    "<span style='background: #fee2e2; color: #dc2626; width: 36px; height: 36px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.1rem; font-weight: 900;'>🚨</span>"
                    "<h4 style='color: #991b1b; margin: 0; font-weight: 900; font-size: 1.15rem;'>Itens Pendentes</h4>"
                    "</div>"
                    f"<span style='background: #ef4444; color: #ffffff; padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 900; box-shadow: 0 2px 4px rgba(239, 68, 68, 0.3);'>{pend_m} itens</span>"
                    "</div>"
                )
                if not df_pend.empty:
                    # Remove rolagem (overflow) no modo impressão para evitar corte na folha
                    scroll_style = "height: auto;" if modo_imp_maq else "max-height: 450px; overflow-y: auto; padding-right: 6px;"
                    html_p += f"<div style='{scroll_style}'>"
                    
                    for _, r in df_pend.iterrows():
                        item = r.get("Item", "Sem descrição")
                        freq = r.get("Frequencia", "N/I")
                        html_p += (
                            "<div style='background: linear-gradient(to right, #fff5f5, #ffffff); padding: 15px 18px; border: 1px solid #fecaca; border-left: 5px solid #ef4444; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.03); display: flex; justify-content: space-between; align-items: center; gap: 12px;'>"
                            "<div>"
                            f"<div style='font-weight: 800; color: #0f172a; font-size: 0.98rem; margin-bottom: 6px;'>{item}</div>"
                            f"<div style='display: inline-flex; align-items: center; gap: 5px; background: #f1f5f9; color: #475569; padding: 3px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; border: 1px solid #e2e8f0;'>🔄 Freq: {freq}</div>"
                            "</div>"
                            "<span style='background: #fee2e2; color: #b91c1c; border: 1px solid #fca5a5; padding: 5px 10px; border-radius: 8px; font-size: 0.72rem; font-weight: 900; white-space: nowrap; text-transform: uppercase;'>⏳ Aguardando</span>"
                            "</div>"
                        )
                    html_p += "</div>"
                else:
                    html_p += (
                        "<div style='background: #f0fdf4; border: 1px dashed #86efac; padding: 35px 20px; border-radius: 12px; text-align: center; margin-top: 20px;'>"
                        "<div style='font-size: 2.5rem; margin-bottom: 8px;'>🎉</div>"
                        "<div style='color: #166534; font-weight: 900; font-size: 1.05rem;'>Tudo em dia por aqui!</div>"
                        "<div style='color: #15803d; font-size: 0.85rem; font-weight: 600; margin-top: 4px;'>Nenhuma inspeção pendente para este equipamento.</div>"
                        "</div>"
                    )
                html_p += "</div>"
                st.markdown(html_p, unsafe_allow_html=True)

            # Apenas exibe itens concluídos se NÃO estiver no modo impressão
            if not modo_imp_maq:
                with c_ok:
                    html_k = (
                        "<div style='background: #ffffff; border: 1px solid #bbf7d0; border-top: 6px solid #10b981; padding: 22px; border-radius: 16px; box-shadow: 0 10px 20px -5px rgba(16, 185, 129, 0.08); min-height: 340px;'>"
                        "<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; padding-bottom: 12px; border-bottom: 2px dashed #dcfce7;'>"
                        "<div style='display: flex; align-items: center; gap: 10px;'>"
                        "<span style='background: #dcfce7; color: #059669; width: 36px; height: 36px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.1rem; font-weight: 900;'>✅</span>"
                        "<h4 style='color: #065f46; margin: 0; font-weight: 900; font-size: 1.15rem;'>Itens Concluídos</h4>"
                        "</div>"
                        f"<span style='background: #10b981; color: #ffffff; padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 900; box-shadow: 0 2px 4px rgba(16, 185, 129, 0.3);'>{ok_m} itens</span>"
                        "</div>"
                    )
                    if not df_ok.empty:
                        html_k += "<div style='max-height: 450px; overflow-y: auto; padding-right: 6px;'>"
                        for _, r in df_ok.iterrows():
                            item = r.get("Item", "Sem descrição")
                            freq = r.get("Frequencia", "N/I")
                            dt_ult = r.get("Data_Ultima").strftime("%d/%m/%Y")
                            html_k += (
                                "<div style='background: linear-gradient(to right, #f0fdf4, #ffffff); padding: 15px 18px; border: 1px solid #bbf7d0; border-left: 5px solid #10b981; border-radius: 12px; margin-bottom: 12px; box-shadow: 0 2px 5px rgba(0,0,0,0.03); display: flex; justify-content: space-between; align-items: center; gap: 12px;'>"
                                "<div>"
                                f"<div style='font-weight: 800; color: #064e3b; font-size: 0.98rem; margin-bottom: 6px;'>{item}</div>"
                                f"<div style='display: inline-flex; align-items: center; gap: 5px; background: #f8fafc; color: #64748b; padding: 3px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; border: 1px solid #e2e8f0;'>🔄 Freq: {freq}</div>"
                                "</div>"
                                f"<span style='background: #dcfce7; color: #166534; border: 1px solid #86efac; padding: 5px 12px; border-radius: 8px; font-size: 0.78rem; font-weight: 900; white-space: nowrap;'>✓ {dt_ult}</span>"
                                "</div>"
                            )
                        html_k += "</div>"
                    else:
                        html_k += (
                            "<div style='background: #f8fafc; border: 1px dashed #cbd5e1; padding: 35px 20px; border-radius: 12px; text-align: center; margin-top: 20px;'>"
                            "<div style='font-size: 2.2rem; margin-bottom: 8px;'>📋</div>"
                            "<div style='color: #475569; font-weight: 800; font-size: 1rem;'>Nenhum item concluído registrado</div>"
                            "<div style='color: #94a3b8; font-size: 0.85rem; margin-top: 4px;'>Preencha a coluna de data na planilha para atualizar.</div>"
                            "</div>"
                        )
                    html_k += "</div>"
                    st.markdown(html_k, unsafe_allow_html=True)

        else:
            st.info(f"Nenhum registro de {tipo_checklist} encontrado para a máquina {maq_sel}.")

    # =========================================================================
    # MODO 2: MAPA GERAL DE MÁQUINAS COM CARDS COLORIDOS
    # =========================================================================
    else:
        # Calcula estatísticas rápidas para os KPIs do topo
        qtd_maqs_ok, qtd_maqs_pend, qtd_maqs_sem = 0, 0, 0
        status_por_maquina = {}

        for m in todas_maquinas:
            df_m = df_check_view[df_check_view["Maquina"] == m]
            if df_m.empty:
                qtd_maqs_sem += 1
                status_por_maquina[m] = ("cinza", "SEM REGISTRO", "⚪", 0, 0)
            else:
                pend_count = int(df_m["Data_Ultima"].isna().sum())
                ok_count = int(df_m["Data_Ultima"].notna().sum())
                if pend_count > 0:
                    qtd_maqs_pend += 1
                    status_por_maquina[m] = ("vermelho", f"{pend_count} PENDENTE(S)", "🔴", pend_count, ok_count)
                else:
                    qtd_maqs_ok += 1
                    status_por_maquina[m] = ("verde", "100% EM DIA", "🟢", 0, ok_count)

        # KPIs DE RESUMO DO MAPA
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Monitorado</div><div class='kpi-val'>{len(todas_maquinas)} <span style='font-size:0.85rem; color:#94a3b8;'>máqs</span></div></div><div style='font-size:1.8rem;'>{icone_tema}</div></div>", unsafe_allow_html=True)
        kc2.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>100% Concluídas</div><div class='kpi-val' style='color:#059669;'>{qtd_maqs_ok}</div></div><div style='font-size:1.8rem;'>🟢</div></div>", unsafe_allow_html=True)
        kc3.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Com Pendências</div><div class='kpi-val' style='color:#dc2626;'>{qtd_maqs_pend}</div></div><div style='font-size:1.8rem;'>🔴</div></div>", unsafe_allow_html=True)
        kc4.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid #94a3b8;'><div><div class='kpi-lbl'>Sem Planilha/Dados</div><div class='kpi-val' style='color:#64748b;'>{qtd_maqs_sem}</div></div><div style='font-size:1.8rem;'>⚪</div></div>", unsafe_allow_html=True)

        # BARRA DE LEGENDA ESTILIZADA
        st.markdown("""
        <div style='background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:12px 20px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.03); margin-top:16px; margin-bottom:20px; display:flex; gap:15px; flex-wrap:wrap; align-items:center; justify-content:space-between;'>
            <div style='font-size:0.8rem; font-weight:900; color:#475569; letter-spacing:0.5px; display:flex; align-items:center; gap:8px;'>
                <span>🎨 LEGENDA DOS CARDS (CLIQUE NO ATIVO PARA ABRIR O DOSSIÊ):</span>
            </div>
            <div style='display:flex; gap:12px; flex-wrap:wrap;'>
                <span style='background:linear-gradient(135deg, #f0fdf4, #dcfce7); color:#166534; border:1px solid #86efac; border-left:5px solid #10b981; padding:5px 14px; border-radius:8px; font-size:0.8rem; font-weight:800; box-shadow:0 2px 4px rgba(16,185,129,0.1);'>🟢 Em Dia (100% Concluído)</span>
                <span style='background:linear-gradient(135deg, #fef2f2, #fee2e2); color:#991b1b; border:1px solid #fca5a5; border-left:5px solid #ef4444; padding:5px 14px; border-radius:8px; font-size:0.8rem; font-weight:800; box-shadow:0 2px 4px rgba(239,68,68,0.1);'>🔴 Pendente (Inspeção Atrasada)</span>
                <span style='background:linear-gradient(135deg, #f8fafc, #f1f5f9); color:#475569; border:1px solid #cbd5e1; border-left:5px solid #94a3b8; padding:5px 14px; border-radius:8px; font-size:0.8rem; font-weight:800;'>⚪ Sem Registro</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # RENDERIZA OS SETORES E OS CARDS COLORIDOS
        for s_nome in DICIONARIO_SETORES.keys():
            maquinas_do_setor = obter_maquinas_setor(s_nome, df_fusos)
            if not maquinas_do_setor: continue
            
            ok_setor = sum(1 for m in maquinas_do_setor if status_por_maquina.get(m, ("cinza",))[0] == "verde")
            pend_setor = sum(1 for m in maquinas_do_setor if status_por_maquina.get(m, ("cinza",))[0] == "vermelho")
            cor_setor_badge = CORES_SETORES.get(s_nome, "#3b82f6")

            st.markdown(f"""
            <div style='display:flex; justify-content:space-between; align-items:center; background:#ffffff; border:1px solid #e2e8f0; border-left:6px solid {cor_setor_badge}; border-radius:10px; padding:10px 16px; margin-top:18px; margin-bottom:12px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>
                <div style='font-weight:900; font-size:1.05rem; color:#0f172a; display:flex; align-items:center; gap:8px;'>
                    <span>🏭 {s_nome}</span>
                    <span style='font-size:0.75rem; color:#64748b; font-weight:700;'>({len(maquinas_do_setor)} máquinas)</span>
                </div>
                <div style='display:flex; gap:8px;'>
                    <span style='background:#dcfce7; color:#166534; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:800;'>✓ {ok_setor} OK</span>
                    <span style='background:#fee2e2; color:#991b1b; padding:3px 10px; border-radius:20px; font-size:0.75rem; font-weight:800;'>⚠️ {pend_setor} Pendentes</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            cols_g = 10 # 10 colunas para deixar os cards mais largos, legíveis e imponentes
            for chunk in [maquinas_do_setor[i:i + cols_g] for i in range(0, len(maquinas_do_setor), cols_g)]:
                cols = st.columns(cols_g)
                for i, m in enumerate(chunk):
                    cor_st, texto_status, icone, p_cnt, ok_cnt = status_por_maquina.get(m, ("cinza", "SEM REGISTRO", "⚪", 0, 0))
                    
                    # Define a paleta de cores exata do Card (Fundo, Borda, Texto e Sombra)
                    if cor_st == "verde":
                        bg_btn = "linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%)"
                        bg_hover = "linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%)"
                        brd_btn = "#86efac"
                        brd_left = "#10b981"
                        txt_color = "#14532d"
                        shd_btn = "rgba(16, 185, 129, 0.18)"
                    elif cor_st == "vermelho":
                        bg_btn = "linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%)"
                        bg_hover = "linear-gradient(135deg, #fee2e2 0%, #fecaca 100%)"
                        brd_btn = "#fca5a5"
                        brd_left = "#ef4444"
                        txt_color = "#7f1d1d"
                        shd_btn = "rgba(239, 68, 68, 0.22)"
                    else:
                        bg_btn = "linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)"
                        bg_hover = "linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%)"
                        brd_btn = "#cbd5e1"
                        brd_left = "#94a3b8"
                        txt_color = "#334155"
                        shd_btn = "rgba(100, 116, 139, 0.12)"

                    css_anchor = f"ancora_check_{m.replace('-', '_')}"
                    css_html = f"<div id='{css_anchor}'></div>"
                    
                    # Injeta estilo pintando cada botão individualmente com sua respectiva cor e sombra
                    css_html += f"""<style>
                    div.element-container:has(#{css_anchor}) {{ display: none !important; }} 
                    div.element-container:has(#{css_anchor}) + div.element-container {{ margin-bottom: -6px !important; }}
                    
                    div.element-container:has(#{css_anchor}) + div.element-container button,
                    div[data-testid="column"]:has(#{css_anchor}) button,
                    div[data-testid="stColumn"]:has(#{css_anchor}) button {{
                        background: {bg_btn} !important; 
                        border: 1px solid {brd_btn} !important;
                        border-left: 5px solid {brd_left} !important;
                        color: {txt_color} !important; 
                        font-weight: 900 !important; 
                        border-radius: 10px !important; 
                        box-shadow: 0 3px 6px -1px {shd_btn}, 0 2px 4px -2px rgba(0,0,0,0.05) !important; 
                        height: 42px !important; 
                        min-height: 42px !important; 
                        width: 100% !important; 
                        padding: 0 4px !important;
                        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
                    }}
                    
                    div.element-container:has(#{css_anchor}) + div.element-container button p,
                    div[data-testid="column"]:has(#{css_anchor}) button p,
                    div[data-testid="stColumn"]:has(#{css_anchor}) button p {{
                        font-size: 0.85rem !important; 
                        font-weight: 900 !important;
                        color: {txt_color} !important;
                        margin: 0 !important; 
                        letter-spacing: 0.2px !important;
                    }}
                    
                    div[data-testid="column"]:has(#{css_anchor}) button:hover,
                    div[data-testid="stColumn"]:has(#{css_anchor}) button:hover {{ 
                        background: {bg_hover} !important;
                        border-color: {brd_left} !important; 
                        transform: translateY(-3px) scale(1.03) !important; 
                        box-shadow: 0 10px 15px -3px {shd_btn}, 0 4px 6px -2px rgba(0,0,0,0.08) !important;
                        z-index: 10 !important;
                    }}
                    </style>"""
                    
                    cols[i].markdown(css_html, unsafe_allow_html=True)
                    
                    if cols[i].button(f"{icone} {m}", key=f"btn_chk_{m.replace('-', '_')}", help=f"Ativo {m} • Status: {texto_status}", use_container_width=True):
                        st.session_state.maq_clicada_check = m
                        st.rerun()


# ==========================================
# PAINEL SERVIÇOS TERCEIROS
# ==========================================
elif st.session_state.pagina_atual == "Terceiros":
    st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>🤝 Controle de Serviços de Terceiros</h2>", unsafe_allow_html=True)
    st.markdown("<div style='color:#64748b; margin-bottom:20px; font-weight:500;'>Acompanhamento de manutenções realizadas por empresas contratadas.</div>", unsafe_allow_html=True)
    
    df_terc = df_terceiros.copy()
    
    if st.session_state.empresa_clicada_terc:
        emp_sel = st.session_state.empresa_clicada_terc
        if st.button("⬅️ Voltar para Empresas", key="btn_voltar_terc"):
            st.session_state.empresa_clicada_terc = None
            st.rerun()
        
        st.markdown(f"<h3 style='color:#1e293b; font-weight:900; margin-top:15px; margin-bottom:20px; display:flex; align-items:center; gap:10px;'><span style='font-size:2rem;'>🏢</span> {emp_sel}</h3>", unsafe_allow_html=True)
        
        df_emp = df_terc[df_terc["Empresa"] == emp_sel].copy()
        
        if not df_emp.empty:
            df_emp = df_emp.dropna(subset=["Data"]).sort_values(by="Data", ascending=False)
            datas_unicas = df_emp["Data"].dt.date.unique()
            
            for dt in datas_unicas:
                df_dt = df_emp[df_emp["Data"].dt.date == dt].copy()
                dt_str = dt.strftime("%d/%m/%Y")
                
                df_dt["Setor"] = df_dt["Setor"].fillna("N/I")
                df_dt["Servico"] = df_dt["Servico"].fillna("Sem descrição").astype(str)
                df_agrupado = df_dt.groupby("Setor")["Servico"].apply(lambda x: " &nbsp;•&nbsp; ".join(x)).reset_index()
                
                html_dt = f"<div style='background:#f8fafc; border:1px solid #e2e8f0; border-left:5px solid #3b82f6; border-radius:12px; padding:16px; margin-bottom:15px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>"
                html_dt += f"<div style='font-size:1.1rem; font-weight:900; color:#0f172a; margin-bottom:12px; border-bottom:1px solid #e2e8f0; padding-bottom:8px;'>📅 {dt_str}</div>"
                
                for _, r in df_agrupado.iterrows():
                    setor = r["Setor"]
                    serv = r["Servico"]
                    
                    html_dt += f"<div style='display:flex; align-items:flex-start; gap:10px; margin-bottom:8px;'>"
                    html_dt += f"<span style='background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe; padding:2px 8px; border-radius:6px; font-size:0.75rem; font-weight:800; min-width:110px; text-align:center;'>📍 {setor}</span>"
                    html_dt += f"<span style='color:#475569; font-size:0.95rem; font-weight:600;'>{serv}</span>"
                    html_dt += "</div>"
                    
                html_dt += "</div>"
                st.markdown(html_dt, unsafe_allow_html=True)
        else:
            st.info("Nenhum serviço registrado ou data não informada para esta empresa.")
    
    else:
        if df_terc.empty:
            st.info("Nenhum serviço registrado. Crie planilhas com o nome da empresa no arquivo TERCEIROS.xlsx na pasta banco_dados.")
        else:
            empresas = sorted(list(df_terc["Empresa"].dropna().unique()))
            
            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
            
            cols_g = 3
            for i in range(0, len(empresas), cols_g):
                cols = st.columns(cols_g)
                for j, emp in enumerate(empresas[i:i+cols_g]):
                    df_e = df_terc[df_terc["Empresa"] == emp]
                    qtd_serv = len(df_e)
                    ult_data = df_e["Data"].max()
                    ult_dt_str = ult_data.strftime("%d/%m/%Y") if pd.notna(ult_data) else "S/ Data"
                    
                    css_anchor = f"ancora_emp_{emp.replace(' ', '_')}"
                    
                    css_html = f"<div id='{css_anchor}'></div>"
                    css_html += f"""<style>
                    div.element-container:has(#{css_anchor}) {{ display: none !important; }} 
                    div.element-container:has(#{css_anchor}) + div.element-container {{ margin-bottom: -15px !important; }}
                    div[data-testid="column"]:has(#{css_anchor}) button,
                    div[data-testid="stColumn"]:has(#{css_anchor}) button {{
                        background-color: #f8fafc !important; 
                        border: 1px solid #cbd5e1 !important; 
                        color: #1d4ed8 !important; 
                        font-weight: 800 !important; 
                        border-radius: 8px !important; 
                        box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important; 
                        width: 100% !important; 
                        height: 38px !important;
                    }}
                    div[data-testid="column"]:has(#{css_anchor}) button:hover,
                    div[data-testid="stColumn"]:has(#{css_anchor}) button:hover {{ 
                        border-color: #3b82f6 !important; 
                        background-color: #eff6ff !important; 
                    }}
                    </style>"""
                    
                    cols[j].markdown(css_html, unsafe_allow_html=True)
                    
                    card_html = f"""
                    <div style='background:#ffffff; border:1px solid #e2e8f0; border-left:5px solid #10b981; border-radius:12px; padding:20px; margin-bottom:10px; box-shadow:0 2px 4px rgba(0,0,0,0.02);'>
                        <div style='font-size:1.2rem; font-weight:900; color:#0f172a; margin-bottom:5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;'>🏢 {emp}</div>
                        <div style='display:flex; justify-content:space-between; align-items:center; margin-top:15px; border-top:1px dashed #e2e8f0; padding-top:12px;'>
                            <span style='font-size:0.75rem; color:#64748b; font-weight:700;'>Último: {ult_dt_str}</span>
                            <span style='background:#f0fdf4; color:#166534; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:800; border:1px solid #bbf7d0;'>{qtd_serv} serviços</span>
                        </div>
                    </div>
                    """
                    cols[j].markdown(card_html, unsafe_allow_html=True)
                    
                    if cols[j].button(f"🔍 Abrir Histórico", key=f"btn_emp_{emp}", use_container_width=True):
                        st.session_state.empresa_clicada_terc = emp
                        st.rerun()


# ==========================================
# DIÁRIO DE MÁQUINAS (DOSSIÊ PREMIUM)
# ==========================================
elif st.session_state.pagina_atual == "Diario":
    st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>📖 Diário de Máquinas</h2>", unsafe_allow_html=True)
    st.markdown("<div style='color:#64748b; margin-bottom:20px; font-weight:500;'>Consulta rápida ao dossiê completo de histórico operacional, quebras, correias e pendências do ativo.</div>", unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("<div style='font-size:0.9rem; font-weight:800; color:#475569; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;'>🔍 Selecione o Ativo</div>", unsafe_allow_html=True)
        c_f1, c_f2, _ = st.columns([3, 3, 6])
        with c_f1: setor_sel = st.selectbox("📍 Setor", list(DICIONARIO_SETORES.keys()), label_visibility="collapsed")
        with c_f2:
            maquinas_disp = obter_maquinas_setor(setor_sel, df_fusos)
            maq_sel = st.selectbox("⚙️ Máquina", maquinas_disp, label_visibility="collapsed")

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    if maq_sel:
        df_maq_manut = df_corretivas[(df_corretivas["Setor"] == setor_sel) & (df_corretivas["Maquina"] == maq_sel)].copy()
        df_maq_prev = df_preventivas[(df_preventivas["SETOR"] == setor_sel) & (df_preventivas["MAQUINA"] == maq_sel)].copy()
        
        lista_timeline = []
        if not df_maq_manut.empty:
            for _, r in df_maq_manut.iterrows():
                desc_raw = r.get("Descricao", "")
                desc_html = html_lista_ponto(desc_raw, cor_texto="#1e293b", font_weight="700", font_size="1.05rem")
                
                exe_raw = r.get("Especialidade", "")
                if pd.isna(exe_raw): exe_raw = ""
                
                lista_timeline.append({"Data": r.get("Data"), "Tipo": r.get("Tipo", "Manutenção"), "Especialidade": str(exe_raw), "Duracao": r.get("Duracao", 0), "Desc_HTML": desc_html})
        
        if not df_maq_prev.empty:
            for _, r in df_maq_prev.iterrows():
                d_hrs = 0
                ini = r.get("INICIO")
                fim = r.get("FIM")
                if pd.notna(ini) and pd.notna(fim):
                    try: d_hrs = round((fim - ini).total_seconds() / 3600, 1)
                    except: pass
                
                serv = str(r.get("SERVIÇOS REALIZADOS", ""))
                pecas = str(r.get("PEÇAS TROCADAS", ""))
                
                html_d = ""
                if serv and serv.lower() not in ["nan", "none", ""]:
                    html_d += f"<div style='margin-top:8px; font-weight:800; color:#0f172a; font-size:0.85rem; text-transform:uppercase;'>🛠️ Serviços Realizados:</div>"
                    html_d += html_lista_ponto(serv, cor_texto="#475569", font_weight="600", font_size="0.95rem")
                
                if pecas and pecas.lower() not in ["nan", "none", ""]:
                    html_d += f"<div style='margin-top:8px; font-weight:800; color:#0f172a; font-size:0.85rem; text-transform:uppercase;'>📦 Peças Trocadas:</div>"
                    html_d += html_lista_ponto(pecas, cor_texto="#475569", font_weight="600", font_size="0.95rem")
                    
                if not html_d: 
                    html_d = "<div style='font-size: 1.05rem; color: #1e293b; font-weight: 700; margin-top: 6px;'>Preventiva Geral (Sem detalhes)</div>"
                
                exe_raw = r.get("EXECUTANTES", "")
                if pd.isna(exe_raw): exe_raw = ""
                
                lista_timeline.append({"Data": ini, "Tipo": "Preventiva", "Especialidade": str(exe_raw), "Duracao": d_hrs, "Desc_HTML": html_d})
        
        df_timeline = pd.DataFrame(lista_timeline)
        if not df_timeline.empty:
            df_timeline["Data"] = pd.to_datetime(df_timeline["Data"], errors="coerce")
            df_timeline = df_timeline.sort_values(by="Data", ascending=False)
            qtd_intervencoes = len(df_timeline)
            total_horas_maq = df_timeline["Duracao"].sum()
        else:
            qtd_intervencoes, total_horas_maq = 0, 0
            
        df_maq_fusos = df_fusos[(df_fusos["Setor"] == setor_sel) & (df_fusos["Maquina_TAG"] == maq_sel)].copy() if not df_fusos.empty else pd.DataFrame()
        total_quebras_maq = int(df_maq_fusos["Quantidade_Quebras"].sum()) if not df_maq_fusos.empty else 0
        
        df_maq_trab = df_trabalhos_fusos[(df_trabalhos_fusos["Setor"] == setor_sel) & (df_trabalhos_fusos["Maquina"] == maq_sel)].copy()
        if not df_maq_trab.empty and "Data" in df_maq_trab.columns: df_maq_trab = df_maq_trab.sort_values(by="Data", ascending=False)

        df_maq_corr = df_correias[(df_correias["Setor"] == setor_sel) & (df_correias["Maquina_TAG"] == maq_sel)].copy()
        ult_corr = df_maq_corr.iloc[-1] if not df_maq_corr.empty else None

        status_correia_lbl, cor_corr = "Sem Registro", "#94a3b8"
        if ult_corr is not None:
            dt1 = str(ult_corr.get("Data_Instalacao_1", "")).strip()
            dt2 = str(ult_corr.get("Data_Instalacao_2", "")).strip()
            tem_dado, pior_dias = False, 0
            if dt1 and dt1 not in ["", "nan", "NaT", "None"]:
                try: pior_dias = max(pior_dias, (date.today() - pd.to_datetime(dt1).date()).days); tem_dado = True
                except: pass
            if dt2 and dt2 not in ["", "nan", "NaT", "None"]:
                try: pior_dias = max(pior_dias, (date.today() - pd.to_datetime(dt2).date()).days); tem_dado = True
                except: pass
            if tem_dado:
                if pior_dias <= 365: status_correia_lbl, cor_corr = "OK (Nova)", "#10b981"
                elif pior_dias <= 547: status_correia_lbl, cor_corr = "Meia-Vida", "#f59e0b"
                else: status_correia_lbl, cor_corr = "Troca Urgente", "#ef4444"

        html_hero = "<div style='background: linear-gradient(90deg, #0f172a 0%, #1e293b 100%); padding: 25px 30px; border-radius: 16px; margin-bottom: 25px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border-left: 8px solid #3b82f6; display: flex; justify-content: space-between; align-items: center;'>"
        html_hero += "<div><div style='color: #94a3b8; font-size: 0.85rem; font-weight: 800; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 5px;'>Dossiê do Ativo</div>"
        html_hero += f"<div style='color: #ffffff; font-size: 2.4rem; font-weight: 900; line-height: 1;'>{maq_sel}</div>"
        html_hero += f"<div style='color: #cbd5e1; font-size: 1rem; font-weight: 600; margin-top: 8px; display: flex; align-items: center; gap: 6px;'><span style='background: rgba(255,255,255,0.1); padding: 4px 10px; border-radius: 6px;'>📍 {setor_sel}</span></div></div>"
        html_hero += "<div style='text-align: right; background: rgba(0,0,0,0.2); padding: 15px 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.05);'>"
        html_hero += f"<div style='color: {cor_corr}; font-size: 1.35rem; font-weight: 900;'>{status_correia_lbl}</div>"
        html_hero += "<div style='color: #94a3b8; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 3px;'>Saúde das Correias</div></div></div>"
        st.markdown(html_hero, unsafe_allow_html=True)
        
        k1, k2, k3 = st.columns(3)
        k1.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Total Intervenções</div><div class='kpi-val'>{qtd_intervencoes} <span style='font-size:0.9rem; color:#94a3b8;'>({total_horas_maq}h)</span></div></div><div style='font-size:1.8rem;'>🛠️</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='card-kpi-bonito c-crit'><div><div class='kpi-lbl'>Total Quebra de Fusos</div><div class='kpi-val' style='color:#dc2626;'>{total_quebras_maq} <span style='font-size:0.9rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        
        pendencias_ativas = len(df_maq_trab[df_maq_trab["Status"].astype(str).str.upper() == "PENDENTE"]) if not df_maq_trab.empty else 0
        cor_pend = "#ef4444" if pendencias_ativas > 0 else "#10b981"
        k3.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_pend};'><div><div class='kpi-lbl'>Pendências Abertas</div><div class='kpi-val' style='color:{cor_pend};'>{pendencias_ativas} <span style='font-size:0.9rem; color:#94a3b8;'>tarefas</span></div></div><div style='font-size:1.8rem;'>📋</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)
        
        c_det1, c_det2 = st.columns([1.3, 1])
        
        with c_det1:
            html_timeline = "<div style='background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02);'>"
            html_timeline += "<div style='display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #f1f5f9; padding-bottom: 12px; margin-bottom: 20px;'>"
            html_timeline += "<h4 style='color:#0f172a; font-weight:900; margin:0;'>🛠️ Linha do Tempo de Intervenções</h4>"
            html_timeline += "<span style='background:#f1f5f9; color:#475569; padding:4px 10px; border-radius:20px; font-size:0.8rem; font-weight:700;'>Últimos registros combinados</span></div>"
            
            if not df_timeline.empty:
                for _, r in df_timeline.head(8).iterrows():
                    dt = r['Data'].strftime('%d/%m/%Y') if pd.notna(r['Data']) else 'S/ Data'
                    tipo = str(r['Tipo'])
                    cor_tipo = "#10b981" if "prev" in tipo.lower() else "#ef4444" if "corret" in tipo.lower() else "#3b82f6"
                    label_exe = "Equipe:" if "prev" in tipo.lower() else "Técnico/Especialidade:"
                    pills_exe = formatar_pills(r['Especialidade'])
                    
                    html_timeline += f"<div style='border-left: 3px solid {cor_tipo}; padding-left: 20px; margin-bottom: 25px; position: relative;'>"
                    html_timeline += f"<div style='position: absolute; left: -9px; top: 0; width: 15px; height: 15px; border-radius: 50%; background: {cor_tipo}; border: 3px solid #ffffff; box-shadow: 0 0 0 1px #e2e8f0;'></div>"
                    html_timeline += f"<div style='font-size: 0.8rem; color: #64748b; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;'><span style='color:#0f172a;'>{dt}</span> &nbsp;•&nbsp; {tipo} &nbsp;•&nbsp; ⏱️ {r['Duracao']}h</div>"
                    html_timeline += f"<div style='margin-top: 6px;'>{r['Desc_HTML']}</div>"
                    html_timeline += f"<div style='font-size: 0.8rem; color: #64748b; margin-top: 10px; font-weight: 600; display:flex; align-items:center; flex-wrap:wrap; gap:6px;'><span>{label_exe}</span> {pills_exe}</div>"
                    html_timeline += "</div>"
            else:
                html_timeline += "<div style='color:#94a3b8; font-size:0.9rem; font-weight:600; text-align:center; padding: 20px 0;'>Nenhuma intervenção (Corretiva ou Preventiva) registrada no log.</div>"
            html_timeline += "</div>"
            st.markdown(html_timeline, unsafe_allow_html=True)
            
            st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
            
            st.markdown("<h4 style='color:#1e293b; font-weight:900; margin-bottom:15px;'>📋 Pendências Abertas (Trabalhos/Fusos)</h4>", unsafe_allow_html=True)
            if pendencias_ativas > 0:
                for _, r in df_maq_trab[df_maq_trab["Status"].astype(str).str.upper() == "PENDENTE"].iterrows():
                    desc_raw = r.get("Descricao", "Sem descrição")
                    desc_html = html_lista_ponto(desc_raw, cor_texto="#92400e", font_weight="700", font_size="1.05rem")
                    qtd = r.get("Quantidade", 0)
                    dt_val = r.get("Data")
                    dt_str = dt_val.strftime("%d/%m/%Y") if pd.notna(dt_val) else "S/ Data"
                    
                    html_pend = "<div style='background:#fffbeb; border-left:5px solid #f59e0b; padding:18px; border-radius:12px; margin-bottom:15px; border-top:1px solid #e2e8f0; border-right:1px solid #e2e8f0; border-bottom:1px solid #e2e8f0; box-shadow:0 4px 6px -1px rgba(0,0,0,0.02);'>"
                    html_pend += f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;'><span style='background:#fef3c7; border:1px solid #fde68a; color:#d97706; padding:4px 12px; border-radius:20px; font-size:0.75rem; font-weight:900; letter-spacing:0.5px;'><span style='margin-right:4px;'>⏳</span> PENDENTE</span><span style='background:#f8fafc; border:1px solid #e2e8f0; padding:4px 10px; border-radius:20px; font-size:0.75rem; color:#64748b; font-weight:700;'>📅 {dt_str}</span></div>"
                    html_pend += desc_html
                    html_pend += f"<div style='margin-top:12px; font-size:0.85rem; font-weight:800; color:#b45309; background:#fffbeb; display:inline-block; padding:4px 12px; border-radius:8px;'>Quantidade afetada: {qtd} un.</div></div>"
                    st.markdown(html_pend, unsafe_allow_html=True)
            else:
                st.markdown("<div style='background:#f0fdf4; border:1px solid #bbf7d0; color:#166534; padding:15px; border-radius:8px; font-weight:700;'>🎉 Nenhuma pendência aberta para esta máquina!</div>", unsafe_allow_html=True)

        with c_det2:
            st.markdown("<h4 style='color:#1e293b; font-weight:900; margin-bottom:15px;'>🔄 Dados de Correias</h4>", unsafe_allow_html=True)
            
            t1 = formatar_modelo(ult_corr.get("Tipo_Correia_1")) if ult_corr is not None else ""
            d1 = str(ult_corr.get("Data_Instalacao_1")).strip() if ult_corr is not None else ""
            t2 = formatar_modelo(ult_corr.get("Tipo_Correia_2")) if ult_corr is not None else ""
            d2 = str(ult_corr.get("Data_Instalacao_2")).strip() if ult_corr is not None else ""
            
            if d1 in ["nan", "NaT", "None"]: d1 = ""
            elif d1: d1 = pd.to_datetime(d1).strftime("%d/%m/%Y")
            if d2 in ["nan", "NaT", "None"]: d2 = ""
            elif d2: d2 = pd.to_datetime(d2).strftime("%d/%m/%Y")

            t1_disp = t1 if t1 else 'Não Informado'
            d1_disp = d1 if d1 else 'Sem Registro'
            t2_disp = t2 if t2 else 'Não Informado'
            d2_disp = d2 if d2 else 'Sem Registro'

            html_corr = "<div style='background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 24px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05);'>"
            html_corr += "<div style='display:flex; justify-content:space-between; align-items:center; border-bottom: 2px dashed #e2e8f0; padding-bottom: 12px; margin-bottom: 20px;'>"
            html_corr += "<div style='font-size: 0.85rem; font-weight: 800; color: #64748b; text-transform:uppercase; letter-spacing: 1px;'>Eixo Operacional</div>"
            html_corr += "<div style='font-size: 0.85rem; font-weight: 800; color: #64748b; text-transform:uppercase; letter-spacing: 1px;'>Especificações</div></div>"
            
            html_corr += "<div style='background: linear-gradient(to right, #f8fafc, #ffffff); border: 1px solid #e2e8f0; border-left: 4px solid #3b82f6; border-radius: 12px; padding: 16px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.02);'>"
            html_corr += "<div style='display: flex; align-items: center; gap: 12px;'>"
            html_corr += "<div style='background: #eff6ff; width: 42px; height: 42px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);'>🔼</div>"
            html_corr += "<div><div style='font-size: 1.15rem; font-weight: 900; color: #0f172a;'>Superior</div><div style='font-size: 0.7rem; font-weight: 800; color: #94a3b8; text-transform: uppercase; letter-spacing:0.5px;'>Transmissão Principal</div></div></div>"
            html_corr += "<div style='display: flex; flex-direction: column; align-items: flex-end; gap: 8px;'>"
            html_corr += f"<div style='background: #eff6ff; border: 1px solid #bfdbfe; color: #1d4ed8; padding: 4px 14px; border-radius: 20px; font-weight: 900; font-size: 1rem; box-shadow: 0 1px 2px rgba(29, 78, 216, 0.1);'>{t1_disp}</div>"
            html_corr += f"<div style='background: #f8fafc; border: 1px solid #e2e8f0; color: #64748b; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 0.75rem;'>📅 {d1_disp}</div>"
            html_corr += "</div></div>"

            html_corr += "<div style='background: linear-gradient(to right, #f8fafc, #ffffff); border: 1px solid #e2e8f0; border-left: 4px solid #8b5cf6; border-radius: 12px; padding: 16px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.02);'>"
            html_corr += "<div style='display: flex; align-items: center; gap: 12px;'>"
            html_corr += "<div style='background: #f3e8ff; width: 42px; height: 42px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);'>🔽</div>"
            html_corr += "<div><div style='font-size: 1.15rem; font-weight: 900; color: #0f172a;'>Inferior</div><div style='font-size: 0.7rem; font-weight: 800; color: #94a3b8; text-transform: uppercase; letter-spacing:0.5px;'>Acionamento Secund.</div></div></div>"
            html_corr += "<div style='display: flex; flex-direction: column; align-items: flex-end; gap: 8px;'>"
            html_corr += f"<div style='background: #faf5ff; border: 1px solid #e9d5ff; color: #7e22ce; padding: 4px 14px; border-radius: 20px; font-weight: 900; font-size: 1rem; box-shadow: 0 1px 2px rgba(126, 34, 206, 0.1);'>{t2_disp}</div>"
            html_corr += f"<div style='background: #f8fafc; border: 1px solid #e2e8f0; color: #64748b; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 0.75rem;'>📅 {d2_disp}</div>"
            html_corr += "</div></div></div>"
            
            st.markdown(html_corr, unsafe_allow_html=True)

            st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)
            
            st.markdown("<h4 style='color:#1e293b; font-weight:900; margin-bottom:15px;'>📉 Evolução de Quebras (Fusos)</h4>", unsafe_allow_html=True)
            with st.container(border=True):
                if not df_maq_fusos.empty and total_quebras_maq > 0:
                    df_graf_maq_fusos = df_maq_fusos[df_maq_fusos["Quantidade_Quebras"] > 0].copy()
                    df_graf_maq_fusos = df_graf_maq_fusos.groupby("Mes")["Quantidade_Quebras"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                    df_graf_maq_fusos["Mes_Abrev"] = df_graf_maq_fusos["Mes"].map(MAPA_MES_ABREV)
                    
                    b_maq = alt.Chart(df_graf_maq_fusos).mark_bar(color="#f59e0b", cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0)), 
                        y=alt.Y("Quantidade_Quebras:Q", title="Quebras"),
                        tooltip=["Mes", "Quantidade_Quebras"]
                    )
                    t_maq = alt.Chart(df_graf_maq_fusos).mark_text(dy=-8, fontSize=11, fontWeight=800).encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV), 
                        y=alt.Y("Quantidade_Quebras:Q"), 
                        text=alt.condition("datum.Quantidade_Quebras > 0", alt.Text("Quantidade_Quebras:Q"), alt.value(""))
                    )
                    st.altair_chart((b_maq + t_maq).properties(height=260), use_container_width=True)
                else:
                    st.markdown("<div style='height:260px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:600; text-align:center;'>Nenhuma quebra de fuso registrada para esta máquina no período.</div>", unsafe_allow_html=True)


# ==========================================
# ESTOQUE DE PEÇAS & MANUTENÇÃO
# ==========================================
elif st.session_state.pagina_atual == "Estoque":
    
    c_tit, c_tabs1, c_tabs2 = st.columns([5, 2.5, 2.5])
    with c_tit:
        st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>📦 Gestão de Materiais</h2>", unsafe_allow_html=True)
        st.markdown("<div style='color:#64748b; margin-bottom:20px; font-weight:500;'>Controle de estoque de peças, componentes de reposição e itens em reparo.</div>", unsafe_allow_html=True)
    
    with c_tabs1:
        if st.button("📦 Inventário Físico", use_container_width=True, type="primary" if st.session_state.aba_estoque == "Inventário" else "secondary"):
            st.session_state.aba_estoque = "Inventário"
            st.rerun()
    with c_tabs2:
        if st.button("🔧 Peças em Manutenção", use_container_width=True, type="primary" if st.session_state.aba_estoque == "Manutenção" else "secondary"):
            st.session_state.aba_estoque = "Manutenção"
            st.rerun()
            
    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # ABA: INVENTÁRIO FÍSICO
    # ---------------------------------------------------------
    if st.session_state.aba_estoque == "Inventário":
        df_calc = df_estoque.copy()
        if not df_calc.empty:
            df_calc["QUANTIDADE"] = pd.to_numeric(df_calc["QUANTIDADE"], errors="coerce").fillna(0)
            df_calc["ESTOQUE MINIMO"] = pd.to_numeric(df_calc["ESTOQUE MINIMO"], errors="coerce").fillna(0)
        
        total_itens = len(df_calc)
        baixo_estoque = len(df_calc[df_calc["QUANTIDADE"] <= df_calc["ESTOQUE MINIMO"]]) if total_itens > 0 else 0
        total_pecas_fisicas = int(df_calc["QUANTIDADE"].sum()) if total_itens > 0 else 0
            
        ke1, ke2, ke3 = st.columns(3)
        ke1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total de Itens Cadastrados</div><div class='kpi-val'>{total_itens}</div></div><div style='font-size:1.8rem;'>📋</div></div>", unsafe_allow_html=True)
        
        cor_alerta = "#ef4444" if baixo_estoque > 0 else "#10b981"
        ke2.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_alerta};'><div><div class='kpi-lbl'>Alertas de Estoque Mínimo</div><div class='kpi-val' style='color:{cor_alerta};'>{baixo_estoque}</div></div><div style='font-size:1.8rem;'>🚨</div></div>", unsafe_allow_html=True)
        ke3.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Total de Peças Físicas</div><div class='kpi-val' style='color:#059669;'>{total_pecas_fisicas} un.</div></div><div style='font-size:1.8rem;'>🔧</div></div>", unsafe_allow_html=True)
        
        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        
        todas_categorias = set()
        if total_itens > 0:
            for cat_str in df_calc["CATEGORIAS"].dropna():
                for c in str(cat_str).split(";"):
                    if c.strip(): todas_categorias.add(c.strip())
        
        with st.container(border=True):
            c_busca, c_cat = st.columns([7, 3])
            busca_texto = c_busca.text_input("🔍 Buscar peça por nome ou descrição...", placeholder="Digite para pesquisar...")
            cat_selecionada = c_cat.selectbox("🏷️ Filtrar por Categoria", ["Todas"] + sorted(list(todas_categorias)))
        
        df_show = df_calc.copy()
        if busca_texto and not df_show.empty:
            df_show = df_show[
                df_show["NOME"].astype(str).str.contains(busca_texto, case=False, na=False) | 
                df_show["DESCRIÇÃO"].astype(str).str.contains(busca_texto, case=False, na=False)
            ]
        if cat_selecionada != "Todas" and not df_show.empty:
            df_show = df_show[df_show["CATEGORIAS"].astype(str).str.contains(cat_selecionada, case=False, na=False)]

        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
        
        if not df_show.empty:
            cols_por_linha = 3
            for i in range(0, len(df_show), cols_por_linha):
                cols = st.columns(cols_por_linha)
                for j, (_, r) in enumerate(df_show.iloc[i:i+cols_por_linha].iterrows()):
                    nome = r.get("NOME", "N/I")
                    desc = r.get("DESCRIÇÃO", "Sem descrição")
                    qtd = int(r.get("QUANTIDADE", 0))
                    minimo = int(r.get("ESTOQUE MINIMO", 0))
                    estado = r.get("ESTADO", "N/I")
                    categorias_str = str(r.get("CATEGORIAS", ""))
                    
                    cor_qtd = "#10b981" if qtd > minimo else "#ef4444"
                    pills_html = formatar_pills(categorias_str) if categorias_str and categorias_str.lower() != "nan" else "<span style='color:#94a3b8; font-size:0.8rem;'>Sem categoria</span>"
                    
                    card_html = f"""
                    <div style='background:#ffffff; border:1px solid #e2e8f0; border-left:5px solid {cor_qtd}; border-radius:12px; padding:16px; margin-bottom:15px; box-shadow:0 2px 4px rgba(0,0,0,0.02); height: 100%; display: flex; flex-direction: column;'>
                        <div style='display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 5px;'>
                            <div style='font-weight:900; font-size:1.1rem; color:#0f172a; line-height: 1.2;'>{nome}</div>
                            <div style='background:#eff6ff; color:#1d4ed8; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:800; border: 1px solid #bfdbfe; white-space: nowrap;'>{estado}</div>
                        </div>
                        <div style='color:#64748b; font-size:0.85rem; flex-grow: 1;'>{desc}</div>
                        <div style='margin-top:10px; margin-bottom:10px;'>
                            {pills_html}
                        </div>
                        <div style='padding-top:10px; border-top:1px dashed #e2e8f0; display:flex; justify-content:space-between; align-items:center;'>
                            <div style='font-size:0.8rem; font-weight:700; color:#475569;'>Mínimo: {minimo}</div>
                            <div style='font-size:1.2rem; font-weight:900; color:{cor_qtd};'>{qtd} un.</div>
                        </div>
                    </div>
                    """
                    cols[j].markdown(card_html, unsafe_allow_html=True)
        else:
            if total_itens == 0:
                st.info("Nenhuma peça cadastrada no momento. Adicione os registros na aba ESTOQUE do arquivo ESTOQUE.xlsx.")
            else:
                st.warning("Nenhuma peça encontrada com os filtros atuais.")

    # ---------------------------------------------------------
    # ABA: PEÇAS EM MANUTENÇÃO
    # ---------------------------------------------------------
    else:
        df_manut = df_pecas_manutencao.copy()
        total_registros_manut = len(df_manut)
        total_pecas_manut = int(pd.to_numeric(df_manut["QUANTIDADE"], errors="coerce").sum()) if not df_manut.empty else 0

        ke1, ke2, ke3 = st.columns(3)
        ke1.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Registros de Manutenção</div><div class='kpi-val'>{total_registros_manut}</div></div><div style='font-size:1.8rem;'>📝</div></div>", unsafe_allow_html=True)
        ke2.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid #f59e0b;'><div><div class='kpi-lbl'>Total de Peças Fora (Em Reparo)</div><div class='kpi-val' style='color:#d97706;'>{total_pecas_manut} un.</div></div><div style='font-size:1.8rem;'>🔧</div></div>", unsafe_allow_html=True)
        ke3.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid #8b5cf6;'><div><div class='kpi-lbl'>Status Geral</div><div class='kpi-val' style='color:#7c3aed; font-size:1.1rem;'>Aguardando Retorno</div></div><div style='font-size:1.8rem;'>⏳</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)
        
        if not df_manut.empty:
            cols_por_linha = 3
            for i in range(0, len(df_manut), cols_por_linha):
                cols = st.columns(cols_por_linha)
                for j, (_, r) in enumerate(df_manut.iloc[i:i+cols_por_linha].iterrows()):
                    nome = r.get("NOME", "N/I")
                    ocorrencia = r.get("OCORRENCIA", "Sem descrição")
                    qtd = int(r.get("QUANTIDADE", 0)) if pd.notna(r.get("QUANTIDADE")) else 0
                    
                    dt_val = r.get("DATA")
                    dt_str = dt_val.strftime("%d/%m/%Y") if pd.notna(dt_val) else "Sem Data"

                    card_html = f"""
                    <div style='background:#ffffff; border:1px solid #e2e8f0; border-left:5px solid #f59e0b; border-radius:12px; padding:16px; margin-bottom:15px; box-shadow:0 2px 4px rgba(0,0,0,0.02); height: 100%; display: flex; flex-direction: column;'>
                        <div style='display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 8px;'>
                            <div style='font-weight:900; font-size:1.1rem; color:#0f172a; line-height: 1.2;'>{nome}</div>
                            <div style='background:#fef3c7; color:#d97706; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:800; border: 1px solid #fde68a; white-space: nowrap;'>EM MANUTENÇÃO</div>
                        </div>
                        <div style='color:#64748b; font-size:0.9rem; flex-grow: 1; font-weight: 600;'><span style='color:#475569;'>Ação:</span> {ocorrencia}</div>
                        <div style='padding-top:12px; margin-top:12px; border-top:1px dashed #e2e8f0; display:flex; justify-content:space-between; align-items:center;'>
                            <div style='font-size:0.8rem; font-weight:700; color:#475569; background:#f8fafc; padding:3px 8px; border-radius:6px; border:1px solid #e2e8f0;'>📅 Data: {dt_str}</div>
                            <div style='font-size:1.1rem; font-weight:900; color:#d97706;'>{qtd} un.</div>
                        </div>
                    </div>
                    """
                    cols[j].markdown(card_html, unsafe_allow_html=True)
        else:
            st.info("Não há registros de peças em manutenção. Adicione itens na aba MANUTENÇÃO do arquivo ESTOQUE.xlsx.")


# ==========================================
# FUNÇÃO UNIFICADA: PAINEL DE CHECKLIST (ELÉTRICA & MECÂNICA) - SEM BUG DE HTML
# ==========================================
def renderizar_painel_checklist_operacional(df_origem, titulo, subtitulo, icone_mod, cor_primaria, bg_suave, brd_suave, chave_mod):
    css_chk = (
        "<style>"
        "@media print {"
        "[data-testid='stSidebar'], header[data-testid='stHeader'], .stDeployButton, [data-testid='stToolbar'] { display: none !important; }"
        ".block-container { padding: 0.5rem 1rem !important; max-width: 100% !important; }"
        ".no-print, div[data-testid='stSelectbox'], div[data-testid='stToggle'] { display: none !important; }"
        ".chk-print-card { break-inside: avoid !important; page-break-inside: avoid !important; box-shadow: none !important; margin-bottom: 12px !important; }"
        "* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }"
        "}"
        ".chk-print-card {"
        "background: #ffffff; border-radius: 14px; padding: 18px; margin-bottom: 16px;"
        "box-shadow: 0 4px 12px -2px rgba(15, 23, 42, 0.06), 0 2px 4px -2px rgba(15, 23, 42, 0.03);"
        "transition: transform 0.15s ease, box-shadow 0.15s ease; break-inside: avoid; page-break-inside: avoid;"
        "}"
        ".chk-print-card:hover { transform: translateY(-2px); box-shadow: 0 10px 20px -4px rgba(15, 23, 42, 0.1); }"
        ".chk-item-row {"
        "display: flex; align-items: flex-start; justify-content: space-between; gap: 10px;"
        "padding: 8px 10px; border-radius: 8px; margin-bottom: 6px; font-size: 0.86rem; line-height: 1.35;"
        "}"
        ".chk-box-square {"
        "display: inline-block; width: 15px; height: 15px; min-width: 15px; border: 2px solid #dc2626;"
        "border-radius: 4px; margin-right: 8px; margin-top: 1px; background: #ffffff;"
        "}"
        ".chk-box-checked {"
        "display: inline-flex; align-items: center; justify-content: center; width: 16px; height: 16px; min-width: 16px;"
        "background: #10b981; color: #ffffff; border-radius: 4px; margin-right: 8px; margin-top: 1px; font-size: 0.7rem; font-weight: 900;"
        "}"
        "</style>"
    )
    st.markdown(css_chk, unsafe_allow_html=True)

    c_tit, c_flt1, c_flt2, c_opt1, c_opt2 = st.columns([3.8, 1.8, 1.8, 1.3, 1.3])
    
    df_trab = df_origem.copy()
    if not df_trab.empty and "Setor_Planilha" in df_trab.columns:
        lista_setores = ["Todos os Setores"] + sorted([s for s in df_trab["Setor_Planilha"].dropna().unique()])
    else:
        lista_setores = ["Todos os Setores"]

    with c_flt1:
        setor_filtro = st.selectbox("📍 Setor", lista_setores, label_visibility="collapsed", key=f"setor_{chave_mod}")
    with c_flt2:
        meses_disp = ["Todos os Meses"] + LISTA_MESES_PUROS
        mes_filtro = st.selectbox("📅 Mês", meses_disp, label_visibility="collapsed", key=f"mes_{chave_mod}")
    with c_opt1:
        so_pendentes = st.toggle("🚨 Só Pendentes", value=False, key=f"pend_{chave_mod}", help="Oculta as máquinas concluídas para economizar folha na impressão.")
    with c_opt2:
        modo_imp = st.toggle("🖨️ Modo Impressão", value=False, key=f"imp_{chave_mod}", help="Ativa tela limpa pronta para impressão.")
        if modo_imp:
            st.markdown(CSS_TELA_LIMPA, unsafe_allow_html=True)

    with c_tit:
        st.markdown(f"<h2 style='margin:0; font-weight:900; color:#0f172a; display:flex; align-items:center; gap:8px;'><span>{icone_mod}</span> {titulo}</h2>", unsafe_allow_html=True)
        st.markdown(f"<div style='color:#64748b; font-size:0.88rem; font-weight:500;'>{subtitulo}</div>", unsafe_allow_html=True)

    if setor_filtro != "Todos os Setores":
        df_trab = df_trab[df_trab["Setor_Planilha"] == setor_filtro]

    if mes_filtro != "Todos os Meses":
        num_mes = LISTA_MESES_PUROS.index(mes_filtro) + 1
        df_trab = df_trab[(df_trab["Data_Ultima"].dt.month == num_mes) | (df_trab["Data_Ultima"].isna())]

    df_pendentes = df_trab[df_trab["Data_Ultima"].isna()].copy()
    df_concluidos = df_trab[df_trab["Data_Ultima"].notna()].copy()

    tot_itens = len(df_trab)
    tot_pendentes = len(df_pendentes)
    tot_concluidos = len(df_concluidos)
    maqs_com_pendencia = df_pendentes["Maquina"].nunique() if not df_pendentes.empty else 0
    pct_geral = int(round((tot_concluidos / tot_itens) * 100)) if tot_itens > 0 else 0
    data_emissao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    cor_barra = "#10b981" if pct_geral == 100 else cor_primaria if pct_geral >= 50 else "#ef4444"
    
    html_banner = (
        f"<div style='background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #ffffff; border-radius: 14px; padding: 18px 24px; margin: 14px 0 18px 0; border-left: 8px solid {cor_primaria}; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.12);'>"
        "<div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;'>"
        "<div>"
        "<div style='font-size: 0.72rem; font-weight: 800; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px;'>RELATÓRIO OPERACIONAL DE INSPEÇÃO • PORTAL PCM</div>"
        f"<div style='font-size: 1.35rem; font-weight: 900; color: #ffffff; margin-top: 2px;'>{titulo} &nbsp;<span style='font-size:0.9rem; font-weight:700; color:#cbd5e1;'>| &nbsp;📍 {setor_filtro} &nbsp;•&nbsp; 📅 {mes_filtro}</span></div>"
        "</div>"
        "<div style='text-align: right; min-width: 220px;'>"
        "<div style='display: flex; justify-content: space-between; font-size: 0.78rem; font-weight: 800; color: #cbd5e1; margin-bottom: 5px;'>"
        "<span>CONFORMIDADE DO SETOR</span>"
        f"<span style='color: #ffffff; font-weight: 900;'>{pct_geral}%</span>"
        "</div>"
        "<div style='width: 100%; height: 8px; background: rgba(255,255,255,0.15); border-radius: 99px; overflow: hidden;'>"
        f"<div style='width: {pct_geral}%; height: 100%; background: {cor_barra}; border-radius: 99px;'></div>"
        "</div>"
        f"<div style='font-size: 0.68rem; color: #94a3b8; margin-top: 5px; font-weight: 600;'>Emitido em: {data_emissao}</div>"
        "</div></div></div>"
    )
    st.markdown(html_banner, unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    cor_pend_kpi = "#ef4444" if tot_pendentes > 0 else "#10b981"
    k1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total de Inspeções</div><div class='kpi-val'>{tot_itens}</div></div><div style='font-size:1.7rem;'>📋</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_pend_kpi};'><div><div class='kpi-lbl'>Itens Pendentes</div><div class='kpi-val' style='color:{cor_pend_kpi};'>{tot_pendentes}</div></div><div style='font-size:1.7rem;'>🚨</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Itens Concluídos</div><div class='kpi-val' style='color:#059669;'>{tot_concluidos}</div></div><div style='font-size:1.7rem;'>✅</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='card-kpi-bonito' style='border-left: 6px solid {cor_pend_kpi};'><div><div class='kpi-lbl'>Máquinas c/ Pendência</div><div class='kpi-val' style='color:{cor_pend_kpi};'>{maqs_com_pendencia}</div></div><div style='font-size:1.7rem;'>⚙️</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    # SEÇÃO 1: MÁQUINAS COM PENDÊNCIAS
    html_hdr_pend = (
        "<div style='display:flex; justify-content:space-between; align-items:center; background:#fef2f2; border:1px solid #fecaca; border-left:6px solid #ef4444; padding:10px 16px; border-radius:10px; margin-bottom:15px;'>"
        "<div style='font-size:1.05rem; font-weight:900; color:#991b1b; display:flex; align-items:center; gap:8px;'>"
        "<span>🚨 ORDEM DE INSPEÇÃO: MÁQUINAS COM PENDÊNCIAS</span>"
        "</div>"
        f"<span style='background:#ef4444; color:#ffffff; padding:3px 12px; border-radius:20px; font-size:0.78rem; font-weight:900;'>{maqs_com_pendencia} Ativos Pendentes</span>"
        "</div>"
    )
    st.markdown(html_hdr_pend, unsafe_allow_html=True)

    if not df_pendentes.empty:
        maquinas_pend = sorted(df_pendentes["Maquina"].dropna().unique())
        cols_g = 3
        for i in range(0, len(maquinas_pend), cols_g):
            cols = st.columns(cols_g)
            for j, maq in enumerate(maquinas_pend[i:i+cols_g]):
                itens_maq = df_pendentes[df_pendentes["Maquina"] == maq]
                setor_maq = itens_maq.iloc[0]["Setor_Planilha"]
                qtd_p_maq = len(itens_maq)

                html_card = (
                    "<div class='chk-print-card' style='border: 1px solid #fca5a5; border-top: 6px solid #ef4444; background: linear-gradient(180deg, #fff5f5 0%, #ffffff 45%);'>"
                    "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:12px; border-bottom:2px dashed #fecaca;'>"
                    "<div style='display:flex; align-items:center; gap:10px;'>"
                    "<div style='background:#ef4444; color:#ffffff; width:36px; height:36px; border-radius:9px; display:flex; align-items:center; justify-content:center; font-weight:900; font-size:1.05rem; box-shadow: 0 2px 6px rgba(239, 68, 68, 0.3);'>⚙️</div>"
                    "<div>"
                    f"<div style='font-size:1.2rem; font-weight:900; color:#0f172a; line-height:1.1;'>{maq}</div>"
                    f"<div style='font-size:0.72rem; font-weight:800; color:#b91c1c; text-transform:uppercase; letter-spacing:0.4px; margin-top:2px;'>📍 {setor_maq}</div>"
                    "</div></div>"
                    "<div style='text-align:right;'>"
                    f"<span style='background:#fee2e2; color:#991b1b; border:1px solid #fca5a5; padding:4px 10px; border-radius:20px; font-size:0.74rem; font-weight:900; text-transform:uppercase;'>⚠️ {qtd_p_maq} Pendente(s)</span>"
                    "</div></div>"
                )

                for _, r_item in itens_maq.iterrows():
                    nome_item = str(r_item.get("Item", "Sem descrição")).strip()
                    freq_str = str(r_item.get("Frequencia", "")).strip()
                    freq_badge = f"<span style='background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; padding:2px 7px; border-radius:6px; font-size:0.68rem; font-weight:800; white-space:nowrap;'>🔄 {freq_str}</span>" if freq_str and freq_str.lower() not in ["nan", "none", ""] else ""
                    
                    html_card += (
                        "<div class='chk-item-row' style='background:#ffffff; border:1px solid #fee2e2; border-left:3px solid #ef4444;'>"
                        "<div style='display:flex; align-items:flex-start; color:#1e293b; font-weight:700;'>"
                        "<span class='chk-box-square'></span>"
                        f"<span>{nome_item}</span>"
                        "</div>"
                        f"{freq_badge}"
                        "</div>"
                    )

                html_card += "</div>"
                cols[j].markdown(html_card, unsafe_allow_html=True)
    else:
        st.markdown("<div style='background:#f0fdf4; border:1px solid #86efac; border-left:6px solid #10b981; color:#166534; padding:18px; border-radius:12px; font-weight:800; font-size:0.95rem;'>🎉 Excelente! Nenhuma máquina possui pendências para os filtros selecionados.</div>", unsafe_allow_html=True)

    # SEÇÃO 2: MÁQUINAS EM DIA (CONCLUÍDAS)
    if not so_pendentes:
        st.markdown("<div style='height:15px;'></div>", unsafe_allow_html=True)
        maqs_ok_count = df_concluidos["Maquina"].nunique() if not df_concluidos.empty else 0
        
        html_hdr_ok = (
            "<div style='display:flex; justify-content:space-between; align-items:center; background:#f0fdf4; border:1px solid #bbf7d0; border-left:6px solid #10b981; padding:10px 16px; border-radius:10px; margin-bottom:15px;'>"
            "<div style='font-size:1.05rem; font-weight:900; color:#065f46; display:flex; align-items:center; gap:8px;'>"
            "<span>✅ HISTÓRICO DE CONFORMIDADE: INSPEÇÕES CONCLUÍDAS</span>"
            "</div>"
            f"<span style='background:#10b981; color:#ffffff; padding:3px 12px; border-radius:20px; font-size:0.78rem; font-weight:900;'>{maqs_ok_count} Ativos c/ Baixa</span>"
            "</div>"
        )
        st.markdown(html_hdr_ok, unsafe_allow_html=True)

        if not df_concluidos.empty:
            maquinas_ok = sorted(df_concluidos["Maquina"].dropna().unique())
            cols_g = 3
            for i in range(0, len(maquinas_ok), cols_g):
                cols = st.columns(cols_g)
                for j, maq in enumerate(maquinas_ok[i:i+cols_g]):
                    itens_maq = df_concluidos[df_concluidos["Maquina"] == maq]
                    setor_maq = itens_maq.iloc[0]["Setor_Planilha"]
                    qtd_ok_maq = len(itens_maq)

                    html_card = (
                        "<div class='chk-print-card' style='border: 1px solid #86efac; border-top: 6px solid #10b981; background: linear-gradient(180deg, #f0fdf4 0%, #ffffff 45%);'>"
                        "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:12px; border-bottom:2px dashed #bbf7d0;'>"
                        "<div style='display:flex; align-items:center; gap:10px;'>"
                        "<div style='background:#10b981; color:#ffffff; width:36px; height:36px; border-radius:9px; display:flex; align-items:center; justify-content:center; font-weight:900; font-size:1.05rem; box-shadow: 0 2px 6px rgba(16, 185, 129, 0.25);'>✓</div>"
                        "<div>"
                        f"<div style='font-size:1.2rem; font-weight:900; color:#0f172a; line-height:1.1;'>{maq}</div>"
                        f"<div style='font-size:0.72rem; font-weight:800; color:#047857; text-transform:uppercase; letter-spacing:0.4px; margin-top:2px;'>📍 {setor_maq}</div>"
                        "</div></div>"
                        f"<span style='background:#dcfce7; color:#166534; border:1px solid #86efac; padding:4px 10px; border-radius:20px; font-size:0.74rem; font-weight:900; text-transform:uppercase;'>✓ {qtd_ok_maq} Feito(s)</span>"
                        "</div>"
                    )

                    for _, r_item in itens_maq.iterrows():
                        nome_item = str(r_item.get("Item", "Sem descrição")).strip()
                        freq_str = str(r_item.get("Frequencia", "")).strip()
                        dt_val = r_item.get("Data_Ultima")
                        dt_str = dt_val.strftime("%d/%m/%Y") if pd.notna(dt_val) else "OK"
                        freq_txt = f" • {freq_str}" if freq_str and freq_str.lower() not in ["nan", "none", ""] else ""

                        html_card += (
                            "<div class='chk-item-row' style='background:#ffffff; border:1px solid #dcfce7; border-left:3px solid #10b981;'>"
                            "<div style='display:flex; align-items:flex-start; color:#064e3b; font-weight:700;'>"
                            "<span class='chk-box-checked'>✓</span>"
                            f"<span>{nome_item} <span style='color:#64748b; font-size:0.74rem; font-weight:600;'>{freq_txt}</span></span>"
                            "</div>"
                            f"<span style='background:#dcfce7; color:#15803d; border:1px solid #bbf7d0; padding:2px 8px; border-radius:6px; font-size:0.72rem; font-weight:900; white-space:nowrap;'>📅 {dt_str}</span>"
                            "</div>"
                        )

                    html_card += "</div>"
                    cols[j].markdown(html_card, unsafe_allow_html=True)
        else:
            st.markdown("<div style='background:#f8fafc; border:1px solid #e2e8f0; color:#64748b; padding:16px; border-radius:10px; font-weight:700;'>Nenhuma inspeção concluída encontrada com os filtros atuais.</div>", unsafe_allow_html=True)


# ==========================================
# PAINEL DE CHECKLIST ELÉTRICA
# ==========================================
if st.session_state.pagina_atual == "Checklist Eletrica":
    renderizar_painel_checklist_operacional(
        df_origem=df_checklist_eletrica,
        titulo="Checklist Elétrica",
        subtitulo="Controle de inspeções preventivas, painéis, motores e componentes elétricos por ativo.",
        icone_mod="⚡",
        cor_primaria="#3b82f6",
        bg_suave="#eff6ff",
        brd_suave="#bfdbfe",
        chave_mod="ele"
    )

# ==========================================
# PAINEL DE CHECKLIST MECÂNICA
# ==========================================
elif st.session_state.pagina_atual == "Checklist Mecanica":
    renderizar_painel_checklist_operacional(
        df_origem=df_checklist_mecanica,
        titulo="Checklist Mecânica",
        subtitulo="Controle de inspeções mecânicas, lubrificação, transmissões e alinhamento por ativo.",
        icone_mod="⚙️",
        cor_primaria="#f59e0b",
        bg_suave="#fffbeb",
        brd_suave="#fde68a",
        chave_mod="mec"
    )
# ==========================================
# CONTROLE SALA DE FUSOS (COM ABAS: PRODUÇÃO DE FUSOS & GIRO DE PEÇAS EM 3 COLUNAS)
# ==========================================
if st.session_state.pagina_atual == "Sala Fusos":
    if "aba_sala_fusos" not in st.session_state:
        st.session_state.aba_sala_fusos = "Producao"

    caminho_sala_fusos = os.path.join(PASTA_BD, "SALA DOS FUSOS.xlsx")
    if not os.path.exists(caminho_sala_fusos) and os.path.exists(os.path.join(PASTA_BD, "SALA FUSOS.xlsx")):
        caminho_sala_fusos = os.path.join(PASTA_BD, "SALA FUSOS.xlsx")

    PECAS_BORRACHA = ["2820", "2475", "2476", "3303", "2515", "3755", "2805", "1196", "6072"]
    PECAS_ROLAMENTO = ["822", "63000", "6001"]
    TODAS_PECAS_SF = PECAS_BORRACHA + PECAS_ROLAMENTO

    def normalizar_modelo_fuso(val):
        s = str(val).strip().upper()
        if s in ["MEN", "MENEGATTO"]:
            return "MENEGATTO"
        return s

    def limpar_codigo_peca(val):
        s = str(val).strip().upper()
        if s.endswith(".0"):
            s = s[:-2]
        return s

    df_manut_sf = pd.DataFrame({
        "DATA": pd.Series(dtype="datetime64[ns]"),
        "MODELO": pd.Series(dtype="str"),
        "MANUTENCAO": pd.Series(dtype="str"),
        "QUANTIDADE": pd.Series(dtype="int")
    })
    df_ret_sf = pd.DataFrame({
        "DATA": pd.Series(dtype="datetime64[ns]"),
        "TIPO": pd.Series(dtype="str"),
        "QUANTIDADE": pd.Series(dtype="int")
    })
    mapa_receita_uso = {}

    if os.path.exists(caminho_sala_fusos):
        try:
            xls_sf = pd.ExcelFile(caminho_sala_fusos)
            
            # 1. Aba MANUTENÇÃO
            aba_m = next((a for a in xls_sf.sheet_names if "MANUT" in a.upper()), xls_sf.sheet_names[0])
            df_m_tmp = pd.read_excel(xls_sf, sheet_name=aba_m)
            df_m_tmp = df_m_tmp.rename(columns=lambda x: str(x).strip().upper())
            df_m_tmp = df_m_tmp.rename(columns={"MANUTENÇÃO": "MANUTENCAO"})
            for col_ob in ["DATA", "MODELO", "MANUTENCAO", "QUANTIDADE"]:
                if col_ob not in df_m_tmp.columns:
                    df_m_tmp[col_ob] = None
            df_m_tmp["DATA"] = pd.to_datetime(df_m_tmp["DATA"], dayfirst=True, errors="coerce")
            df_m_tmp["QUANTIDADE"] = pd.to_numeric(df_m_tmp["QUANTIDADE"], errors="coerce").fillna(0).astype(int)
            df_m_tmp["MODELO"] = df_m_tmp["MODELO"].apply(normalizar_modelo_fuso)
            df_m_tmp["MANUTENCAO"] = df_m_tmp["MANUTENCAO"].astype(str).str.strip().str.upper()
            df_manut_sf = df_m_tmp.dropna(subset=["DATA"]).copy()
            df_manut_sf["DATA"] = pd.to_datetime(df_manut_sf["DATA"])

            # 2. Aba RETIRADA
            aba_r = next((a for a in xls_sf.sheet_names if "RETIR" in a.upper()), None)
            if not aba_r and len(xls_sf.sheet_names) > 1:
                aba_r = xls_sf.sheet_names[1]
            if aba_r:
                df_r_tmp = pd.read_excel(xls_sf, sheet_name=aba_r)
                df_r_tmp = df_r_tmp.rename(columns=lambda x: str(x).strip().upper())
                for col_ob in ["DATA", "TIPO", "QUANTIDADE"]:
                    if col_ob not in df_r_tmp.columns:
                        df_r_tmp[col_ob] = None
                df_r_tmp["DATA"] = pd.to_datetime(df_r_tmp["DATA"], dayfirst=True, errors="coerce")
                df_r_tmp["QUANTIDADE"] = pd.to_numeric(df_r_tmp["QUANTIDADE"], errors="coerce").fillna(0).astype(int)
                df_r_tmp["TIPO"] = df_r_tmp["TIPO"].apply(limpar_codigo_peca)
                df_ret_sf = df_r_tmp.dropna(subset=["DATA"]).copy()
                df_ret_sf["DATA"] = pd.to_datetime(df_ret_sf["DATA"])

            # 3. Aba USO
            aba_u = next((a for a in xls_sf.sheet_names if "USO" in a.upper()), None)
            if not aba_u and len(xls_sf.sheet_names) > 2:
                aba_u = xls_sf.sheet_names[2]
            if aba_u:
                df_u_raw = pd.read_excel(xls_sf, sheet_name=aba_u, header=None)
                if len(df_u_raw) >= 3:
                    idx_codigos = 1
                    for idx_r in range(min(4, len(df_u_raw))):
                        vals_linha = [limpar_codigo_peca(x) for x in df_u_raw.iloc[idx_r].values]
                        if "2820" in vals_linha or "3755" in vals_linha:
                            idx_codigos = idx_r
                            break
                    
                    codigos_col = {}
                    for col_idx in range(2, df_u_raw.shape[1]):
                        cod = limpar_codigo_peca(df_u_raw.iloc[idx_codigos, col_idx])
                        if cod and cod not in ["NAN", "NONE", ""]:
                            codigos_col[col_idx] = cod

                    for row_idx in range(idx_codigos + 1, len(df_u_raw)):
                        mod_u = normalizar_modelo_fuso(df_u_raw.iloc[row_idx, 0])
                        man_u = str(df_u_raw.iloc[row_idx, 1]).strip().upper()
                        if not mod_u or mod_u in ["NAN", "NONE", ""]:
                            continue
                        mapa_receita_uso[(mod_u, man_u)] = {}
                        for c_idx, cod_peca in codigos_col.items():
                            val_cel = str(df_u_raw.iloc[row_idx, c_idx]).strip().upper()
                            if val_cel in ["X", "", "NAN", "NONE"]:
                                qtd_por_fuso = 0
                            else:
                                try:
                                    qtd_por_fuso = int(float(val_cel))
                                except Exception:
                                    qtd_por_fuso = 0
                            mapa_receita_uso[(mod_u, man_u)][cod_peca] = qtd_por_fuso
        except Exception as e:
            st.error(f"Erro ao ler o arquivo SALA DOS FUSOS.xlsx: {e}")
    else:
        st.warning(f"Arquivo 'SALA DOS FUSOS.xlsx' não encontrado dentro da pasta '{PASTA_BD}'.")

    # =========================================================================
    # CABEÇALHO E BOTÕES DE NAVEGAÇÃO ENTRE AS 2 PÁGINAS DA SALA DE FUSOS
    # =========================================================================
    c_tit, c_aba1, c_aba2 = st.columns([5.2, 2.4, 2.4])
    with c_tit:
        st.markdown("<h2 style='margin:0; font-weight:900; color:#0f172a;'>🛠️ Controle Sala de Fusos</h2>", unsafe_allow_html=True)
        st.markdown("<div style='color:#64748b; font-size:0.88rem; font-weight:500;'>Gestão de manutenções realizadas na bancada e giro de peças do almoxarifado.</div>", unsafe_allow_html=True)
    with c_aba1:
        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
        if st.button("🔩 Produção de Fusos", use_container_width=True, type="primary" if st.session_state.aba_sala_fusos == "Producao" else "secondary"):
            st.session_state.aba_sala_fusos = "Producao"
            st.rerun()
    with c_aba2:
        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Giro de Peças", use_container_width=True, type="primary" if st.session_state.aba_sala_fusos == "Giro" else "secondary"):
            st.session_state.aba_sala_fusos = "Giro"
            st.rerun()

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # FILTROS GLOBAIS DA PÁGINA (ANO E MÊS)
    anos_m = df_manut_sf["DATA"].dt.year.dropna().astype(int).tolist() if not df_manut_sf.empty else []
    anos_r = df_ret_sf["DATA"].dt.year.dropna().astype(int).tolist() if not df_ret_sf.empty else []
    anos_sf = sorted(list(set(anos_m) | set(anos_r) | {2026}))

    with st.container(border=True):
        cf1, cf2, cf3 = st.columns([3, 3, 4])
        with cf1:
            ano_sel_sf = st.selectbox("📅 Filtrar Ano:", ["Todos"] + anos_sf, index=len(anos_sf), key="sf_ano_unif")
        with cf2:
            mes_sel_sf = st.selectbox("🗓️ Filtrar Mês:", ["Todos os Meses"] + LISTA_MESES_PUROS, index=0, key="sf_mes_unif")
        with cf3:
            mods_plan = [m for m in df_manut_sf["MODELO"].dropna().unique().tolist() if m and m not in ["NAN", "NONE"]] if not df_manut_sf.empty else []
            modelos_existentes = sorted(list(set(mods_plan) | set(OPCOES_TIPO_FUSO)))
            mod_sel_sf = st.selectbox("🔩 Filtrar Modelo de Fuso:", ["Todos Modelos"] + modelos_existentes, index=0, key="sf_mod_unif")

    df_m_filt = df_manut_sf.copy()
    df_r_filt = df_ret_sf.copy()

    if ano_sel_sf != "Todos":
        if not df_m_filt.empty:
            df_m_filt = df_m_filt[df_m_filt["DATA"].dt.year == int(ano_sel_sf)]
        if not df_r_filt.empty:
            df_r_filt = df_r_filt[df_r_filt["DATA"].dt.year == int(ano_sel_sf)]

    if mes_sel_sf != "Todos os Meses":
        num_m_sf = LISTA_MESES_PUROS.index(mes_sel_sf) + 1
        if not df_m_filt.empty:
            df_m_filt = df_m_filt[df_m_filt["DATA"].dt.month == num_m_sf]
        if not df_r_filt.empty:
            df_r_filt = df_r_filt[df_r_filt["DATA"].dt.month == num_m_sf]

    if mod_sel_sf != "Todos Modelos" and not df_m_filt.empty:
        df_m_filt = df_m_filt[df_m_filt["MODELO"] == mod_sel_sf]

    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

    # =========================================================================
    # PÁGINA 1: PRODUÇÃO DE FUSOS (OS 3 GRÁFICOS LIMPOS)
    # =========================================================================
    if st.session_state.aba_sala_fusos == "Producao":
        tot_fusos_manut = int(df_m_filt["QUANTIDADE"].sum()) if not df_m_filt.empty else 0
        tot_completa = int(df_m_filt[df_m_filt["MANUTENCAO"].str.contains("COMPLET", na=False)]["QUANTIDADE"].sum()) if not df_m_filt.empty else 0
        tot_amortecedor = int(df_m_filt[df_m_filt["MANUTENCAO"].str.contains("AMORTEC", na=False)]["QUANTIDADE"].sum()) if not df_m_filt.empty else 0

        kp1, kp2, kp3 = st.columns(3)
        kp1.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total de Fusos Feitos</div><div class='kpi-val'>{tot_fusos_manut} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🔩</div></div>", unsafe_allow_html=True)
        kp2.markdown(f"<div class='card-kpi-bonito' style='border-left:6px solid #8b5cf6;'><div><div class='kpi-lbl'>Manutenção Completa</div><div class='kpi-val' style='color:#7c3aed;'>{tot_completa} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>⚙️</div></div>", unsafe_allow_html=True)
        kp3.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Troca de Amortecedor</div><div class='kpi-val' style='color:#d97706;'>{tot_amortecedor} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🛡️</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

        c_g1, c_g2 = st.columns(2)

        with c_g1:
            with st.container(border=True):
                st.markdown("<div style='font-size:1.05rem; font-weight:900; color:#0f172a; margin-bottom:12px;'>📊 Fusos Feitos por Modelo</div>", unsafe_allow_html=True)
                if not df_m_filt.empty:
                    df_mod = df_m_filt.groupby("MODELO")["QUANTIDADE"].sum().reset_index()
                    graf_mod = alt.Chart(df_mod).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, color="#3b82f6", size=32).encode(
                        x=alt.X("MODELO:N", title=None, axis=alt.Axis(labelAngle=0, labelFontWeight=800)),
                        y=alt.Y("QUANTIDADE:Q", title="Quantidade de Fusos"),
                        tooltip=["MODELO", "QUANTIDADE"]
                    )
                    txt_mod = graf_mod.mark_text(dy=-10, fontSize=12, fontWeight=800).encode(text="QUANTIDADE:Q")
                    st.altair_chart((graf_mod + txt_mod).properties(height=260), use_container_width=True)
                else:
                    st.markdown("<div style='height:260px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Sem dados para exibir.</div>", unsafe_allow_html=True)

        with c_g2:
            with st.container(border=True):
                st.markdown("<div style='font-size:1.05rem; font-weight:900; color:#0f172a; margin-bottom:12px;'>📈 Fusos Feitos por Mês</div>", unsafe_allow_html=True)
                if not df_m_filt.empty:
                    df_m_filt["MES_NOME"] = df_m_filt["DATA"].dt.month.map(lambda x: LISTA_MESES_PUROS[x-1] if pd.notna(x) and 1 <= x <= 12 else "")
                    df_mes = df_m_filt.groupby("MES_NOME")["QUANTIDADE"].sum().reindex(LISTA_MESES_PUROS, fill_value=0).reset_index()
                    df_mes["Mes_Abrev"] = df_mes["MES_NOME"].map(MAPA_MES_ABREV)

                    graf_mes = alt.Chart(df_mes).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, color="#10b981", size=24).encode(
                        x=alt.X("Mes_Abrev:N", sort=ORDEM_MESES_ABREV, title=None, axis=alt.Axis(labelAngle=0, labelFontWeight=800)),
                        y=alt.Y("QUANTIDADE:Q", title="Quantidade de Fusos"),
                        tooltip=["MES_NOME", "QUANTIDADE"]
                    )
                    txt_mes = graf_mes.mark_text(dy=-10, fontSize=11, fontWeight=800).encode(
                        text=alt.condition("datum.QUANTIDADE > 0", alt.Text("QUANTIDADE:Q"), alt.value(""))
                    )
                    st.altair_chart((graf_mes + txt_mes).properties(height=260), use_container_width=True)
                else:
                    st.markdown("<div style='height:260px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Sem dados para exibir.</div>", unsafe_allow_html=True)

        st.markdown("<div style='height:15px;'></div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("<div style='font-size:1.05rem; font-weight:900; color:#0f172a; margin-bottom:12px;'>📅 Fusos Feitos por Dia</div>", unsafe_allow_html=True)
            if not df_m_filt.empty:
                df_dia = df_m_filt.groupby(df_m_filt["DATA"].dt.date)["QUANTIDADE"].sum().reset_index()
                df_dia.columns = ["DATA", "QUANTIDADE"]
                df_dia["DATA_STR"] = pd.to_datetime(df_dia["DATA"]).dt.strftime("%d/%m/%Y")

                graf_dia = alt.Chart(df_dia).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5, color="#f59e0b", size=24).encode(
                    x=alt.X("DATA_STR:N", title=None, axis=alt.Axis(labelAngle=0, labelFontWeight=800)),
                    y=alt.Y("QUANTIDADE:Q", title="Quantidade de Fusos"),
                    tooltip=["DATA_STR", "QUANTIDADE"]
                )
                txt_dia = graf_dia.mark_text(dy=-10, fontSize=11, fontWeight=800).encode(text="QUANTIDADE:Q")
                st.altair_chart((graf_dia + txt_dia).properties(height=240), use_container_width=True)
            else:
                st.markdown("<div style='height:240px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-weight:700;'>Sem dados para exibir no período selecionado.</div>", unsafe_allow_html=True)

    # =========================================================================
    # PÁGINA 2: GIRO DE PEÇAS (DIVIDIDA EM 3 PARTES: RETIRADA | AMORTECEDORES | ROLAMENTOS)
    # =========================================================================
    else:
        consumo_por_peca = {p: 0 for p in TODAS_PECAS_SF}
        for _, row_m in df_m_filt.iterrows():
            mod_m = normalizar_modelo_fuso(row_m.get("MODELO", ""))
            tipo_man = str(row_m.get("MANUTENCAO", "")).strip().upper()
            qtd_fusos_feitos = int(row_m.get("QUANTIDADE", 0))
            receita = mapa_receita_uso.get((mod_m, tipo_man), {})
            for cod_p, qtd_unit in receita.items():
                consumo_por_peca[cod_p] = consumo_por_peca.get(cod_p, 0) + (qtd_fusos_feitos * qtd_unit)

        retirada_por_peca = {p: 0 for p in TODAS_PECAS_SF}
        for _, row_r in df_r_filt.iterrows():
            cod_r = limpar_codigo_peca(row_r.get("TIPO", ""))
            qtd_r = int(row_r.get("QUANTIDADE", 0))
            retirada_por_peca[cod_r] = retirada_por_peca.get(cod_r, 0) + qtd_r

        tot_retirado = sum(retirada_por_peca.values())
        tot_amort_usados = sum(consumo_por_peca.get(p, 0) for p in PECAS_BORRACHA)
        tot_rol_usados = sum(consumo_por_peca.get(p, 0) for p in PECAS_ROLAMENTO)

        # 3 KPIs alinhados com as 3 colunas abaixo
        kg1, kg2, kg3 = st.columns(3)
        kg1.markdown(f"<div class='card-kpi-bonito c-ok'><div><div class='kpi-lbl'>Total Retirado (Almoxarifado)</div><div class='kpi-val' style='color:#059669;'>{tot_retirado} <span style='font-size:0.85rem; color:#94a3b8;'>pçs</span></div></div><div style='font-size:1.8rem;'>📥</div></div>", unsafe_allow_html=True)
        kg2.markdown(f"<div class='card-kpi-bonito c-warn'><div><div class='kpi-lbl'>Total Amortecedores Usados</div><div class='kpi-val' style='color:#d97706;'>{tot_amort_usados} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>🛡️</div></div>", unsafe_allow_html=True)
        kg3.markdown(f"<div class='card-kpi-bonito c-total'><div><div class='kpi-lbl'>Total Rolamentos Usados</div><div class='kpi-val' style='color:#2563eb;'>{tot_rol_usados} <span style='font-size:0.85rem; color:#94a3b8;'>un.</span></div></div><div style='font-size:1.8rem;'>⚙️</div></div>", unsafe_allow_html=True)

        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

        # DIVISÃO EM 3 COLUNAS PRINCIPAIS
        col_retirada, col_amortecedores, col_rolamentos = st.columns(3)

        # -----------------------------------------------------------------
        # COLUNA 1: REGISTRO DE RETIRADA
        # -----------------------------------------------------------------
        with col_retirada:
            html_hr = (
                "<div style='background:#ffffff; border:1px solid #bbf7d0; border-top:6px solid #10b981; border-radius:14px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.03); min-height:480px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:14px; border-bottom:2px dashed #dcfce7;'>"
                "<div style='font-size:1.02rem; font-weight:900; color:#065f46;'>📥 Registro de Retirada</div>"
                f"<span style='background:#10b981; color:#ffffff; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:900;'>{len(df_r_filt)} registros</span></div>"
            )
            if not df_r_filt.empty:
                df_r_ord = df_r_filt.sort_values(by="DATA", ascending=False)
                html_hr += "<div style='max-height:420px; overflow-y:auto; padding-right:4px;'>"
                for _, rr in df_r_ord.iterrows():
                    dt_r = rr["DATA"].strftime("%d/%m/%Y") if pd.notna(rr["DATA"]) else "S/ Data"
                    cod_t = limpar_codigo_peca(rr.get("TIPO", ""))
                    qtd_t = int(rr.get("QUANTIDADE", 0))
                    cat_peca = "🛡️ Borracha" if cod_t in PECAS_BORRACHA else "⚙️ Rolamento" if cod_t in PECAS_ROLAMENTO else "📦 Peça"

                    html_hr += (
                        "<div style='background:linear-gradient(to right, #f0fdf4, #ffffff); border:1px solid #bbf7d0; border-left:5px solid #10b981; border-radius:10px; padding:12px 14px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;'>"
                        f"<div><div style='font-weight:900; font-size:0.95rem; color:#064e3b;'>🏷️ Cód. {cod_t}</div><div style='font-size:0.72rem; color:#047857; font-weight:700; margin-top:2px;'>{cat_peca}</div></div>"
                        f"<div style='text-align:right;'><div style='background:#dcfce7; border:1px solid #86efac; color:#166534; padding:2px 10px; border-radius:8px; font-size:0.88rem; font-weight:900; display:inline-block;'>+{qtd_t} un.</div><div style='font-size:0.7rem; color:#64748b; font-weight:700; margin-top:3px;'>📅 {dt_r}</div></div>"
                        "</div>"
                    )
                html_hr += "</div>"
            else:
                html_hr += "<div style='text-align:center; padding:50px 15px; color:#64748b; font-weight:700;'>Nenhuma retirada registrada no período.</div>"
            html_hr += "</div>"
            st.markdown(html_hr, unsafe_allow_html=True)

        # -----------------------------------------------------------------
        # COLUNA 2: AMORTECEDORES USADOS (BORRACHAS)
        # -----------------------------------------------------------------
        with col_amortecedores:
            html_am = (
                "<div style='background:#ffffff; border:1px solid #fde68a; border-top:6px solid #f59e0b; border-radius:14px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.03); min-height:480px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:14px; border-bottom:2px dashed #fef3c7;'>"
                "<div style='font-size:1.02rem; font-weight:900; color:#b45309;'>🛡️ Amortecedores Usados</div>"
                f"<span style='background:#f59e0b; color:#ffffff; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:900;'>{tot_amort_usados} usados</span></div>"
                "<div style='max-height:420px; overflow-y:auto; padding-right:4px;'>"
            )
            for cod_b in PECAS_BORRACHA:
                uso_b = consumo_por_peca.get(cod_b, 0)
                ret_b = retirada_por_peca.get(cod_b, 0)
                saldo_b = ret_b - uso_b
                cor_sb = "#059669" if saldo_b > 0 else "#dc2626" if saldo_b < 0 else "#64748b"
                bg_row_b = "linear-gradient(to right, #fffbeb, #ffffff)" if uso_b > 0 or ret_b > 0 else "#f8fafc"
                borda_l_b = "#f59e0b" if uso_b > 0 else "#cbd5e1"

                html_am += (
                    f"<div style='background:{bg_row_b}; border:1px solid #fde68a; border-left:5px solid {borda_l_b}; border-radius:10px; padding:10px 14px; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;'>"
                    f"<div><span style='background:#0f172a; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:900; font-size:0.88rem;'>🏷️ {cod_b}</span>"
                    f"<div style='font-size:0.72rem; color:#64748b; font-weight:700; margin-top:4px;'>📥 Pego: <b>{ret_b}</b> &nbsp;|&nbsp; <span style='color:{cor_sb}; font-weight:800;'>Saldo: {saldo_b:+d}</span></div></div>"
                    f"<div style='text-align:right;'><div style='font-size:0.68rem; color:#94a3b8; font-weight:800; text-transform:uppercase;'>Usado</div><div style='font-size:1.15rem; font-weight:900; color:#d97706;'>{uso_b} un.</div></div>"
                    "</div>"
                )
            html_am += "</div></div>"
            st.markdown(html_am, unsafe_allow_html=True)

        # -----------------------------------------------------------------
        # COLUNA 3: ROLAMENTOS USADOS
        # -----------------------------------------------------------------
        with col_rolamentos:
            html_rl = (
                "<div style='background:#ffffff; border:1px solid #bfdbfe; border-top:6px solid #3b82f6; border-radius:14px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.03); min-height:480px;'>"
                "<div style='display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; margin-bottom:14px; border-bottom:2px dashed #dbeafe;'>"
                "<div style='font-size:1.02rem; font-weight:900; color:#1d4ed8;'>⚙️ Rolamentos Usados</div>"
                f"<span style='background:#3b82f6; color:#ffffff; padding:3px 10px; border-radius:20px; font-size:0.74rem; font-weight:900;'>{tot_rol_usados} usados</span></div>"
                "<div>"
            )
            for cod_r in PECAS_ROLAMENTO:
                uso_r = consumo_por_peca.get(cod_r, 0)
                ret_r = retirada_por_peca.get(cod_r, 0)
                saldo_r = ret_r - uso_r
                cor_sr = "#059669" if saldo_r > 0 else "#dc2626" if saldo_r < 0 else "#64748b"
                bg_row_r = "linear-gradient(to right, #eff6ff, #ffffff)" if uso_r > 0 or ret_r > 0 else "#f8fafc"
                borda_l_r = "#3b82f6" if uso_r > 0 else "#cbd5e1"

                html_rl += (
                    f"<div style='background:{bg_row_r}; border:1px solid #bfdbfe; border-left:5px solid {borda_l_r}; border-radius:10px; padding:14px 16px; margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;'>"
                    f"<div><span style='background:#0f172a; color:#ffffff; padding:4px 12px; border-radius:6px; font-weight:900; font-size:0.92rem;'>🏷️ {cod_r}</span>"
                    f"<div style='font-size:0.75rem; color:#64748b; font-weight:700; margin-top:6px;'>📥 Pego: <b>{ret_r}</b> &nbsp;|&nbsp; <span style='color:{cor_sr}; font-weight:800;'>Saldo: {saldo_r:+d}</span></div></div>"
                    f"<div style='text-align:right;'><div style='font-size:0.7rem; color:#94a3b8; font-weight:800; text-transform:uppercase;'>Usado</div><div style='font-size:1.25rem; font-weight:900; color:#2563eb;'>{uso_r} un.</div></div>"
                    "</div>"
                )
            html_rl += "</div></div>"
            st.markdown(html_rl, unsafe_allow_html=True)
