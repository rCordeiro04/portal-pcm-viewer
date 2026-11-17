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
        # Cria uma "etiqueta" para o código reconhecer se já descompactou este ficheiro hoje
        id_arquivo = arquivo_zip.name + str(arquivo_zip.size)
        
        if st.session_state.get("ultimo_zip") != id_arquivo:
            os.makedirs(PASTA_BD, exist_ok=True)
            
            with zipfile.ZipFile(arquivo_zip, "r") as zip_ref:
                zip_ref.extractall(PASTA_BD)
            
            st.session_state["ultimo_zip"] = id_arquivo
            st.cache_data.clear()
            st.rerun() # Agora ele reinicia só 1 vez, vê a etiqueta e para o ciclo!
        else:
            st.success("✅ Base de dados atualizada e pronta!")
            st.markdown("<br><div style='text-align:center; font-size:0.7rem; color:#475569; font-weight:600;'>Portal PCM • Visualizador Executivo</div>", unsafe_allow_html=True)
    
