import streamlit as st
import pandas as pd
import os
import re
import unicodedata
from datetime import datetime
import io
import matplotlib.pyplot as plt

PASTA_DADOS = "dados"
st.set_page_config(layout="wide")
st.title("🚚 Torre de Controle Logística")

# =============================
# FUNÇÕES
# =============================
def normalizar(txt):
    return unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode().lower()

def achar_coluna(df, termos):
    if isinstance(termos, str):
        termos = [termos]
    termos = [normalizar(t) for t in termos]

    for col in df.columns:
        col_norm = normalizar(col)
        if all(t in col_norm for t in termos):
            return col
    return None

def coluna_excel_para_indice(letra):
    resultado = 0
    for char in letra.upper():
        resultado = resultado * 26 + (ord(char) - ord('A') + 1)
    return resultado - 1

def limpar_nome(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto)
    texto = re.sub(r'\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}', '', texto)
    return texto.strip().upper()

def arquivo_recente():
    if not os.path.exists(PASTA_DADOS):
        return None
    arquivos = [
        os.path.join(PASTA_DADOS, f)
        for f in os.listdir(PASTA_DADOS)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]
    return max(arquivos, key=os.path.getmtime) if arquivos else None

@st.cache_data
def ler_excel(arquivo):
    for i in range(5):
        df = pd.read_excel(arquivo, header=i)
        cols = [normalizar(c) for c in df.columns]
        if any("nota" in c or "previs" in c for c in cols):
            return df
    return pd.read_excel(arquivo, header=1)

def tratar_dados(df):
    df = df.copy()
    df.columns = df.columns.map(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("unnamed", case=False)]
    df = df.dropna(how="all")
    return df

def gerar_excel(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for nome, df_aba in df_dict.items():
            df_aba.to_excel(writer, index=False, sheet_name=nome)
    output.seek(0)
    return output

# =============================
# INPUT
# =============================
upload = st.sidebar.file_uploader("Upload", type=["xlsx"])
caminho = upload if upload else arquivo_recente()

if not caminho:
    st.error("Nenhum arquivo encontrado")
    st.stop()

df = tratar_dados(ler_excel(caminho))

# =============================
# MAPEAMENTO
# =============================
col_nf = achar_coluna(df, "nota")
col_doc = achar_coluna(df, "documento")
col_praca = achar_coluna(df, ["praca", "destino"]) or achar_coluna(df, ["cidade", "destino"])
col_atraso = achar_coluna(df, "dias")
col_ocorr = achar_coluna(df, "ocorr")
col_loc = achar_coluna(df, ["localizacao", "atual"])
col_desc_ultima = achar_coluna(df, "descricao da ultima ocorrencia")
col_data_ultima = achar_coluna(df, "data da ultima ocorrencia")

def pegar_coluna_segura(df, letra):
    idx = coluna_excel_para_indice(letra)
    if idx >= len(df.columns):
        return None
    col = df.columns[idx]
    if "unnamed" in str(col).lower():
        return None
    if df[col].isna().all():
        return None
    return col

col_rem = pegar_coluna_segura(df, "P")
col_dest = pegar_coluna_segura(df, "AJ")

# =============================
# TRATAMENTO BASE
# =============================
df["NF"] = df[col_nf]
df["Documento"] = df[col_doc] if col_doc else ""
df["Documento_str"] = df["Documento"].fillna("").astype(str).str.lower()

df["Tipo_Documento"] = df.apply(lambda row: (
    "DEVOLUÇÃO"
    if (
        "devolucao" in str(row["Documento_str"]) or
        any(p in str(row.get("Desc_Ocorrencia", "")).lower() for p in [
            "devolucao", "devolvida", "retorno", "avaria", "recusa", "insucesso"
        ])
    )
    else "NORMAL"
), axis=1)

df["Tipo de Nota"] = df["Tipo_Documento"]

df["Remetente"] = df[col_rem].fillna("").astype(str).apply(limpar_nome)
df["Destinatario"] = df[col_dest].fillna("").astype(str).apply(limpar_nome)
df["Praca"] = df[col_praca] if col_praca else "SEM PRAÇA"

df["Dias_Atraso"] = pd.to_numeric(df[col_atraso], errors="coerce").fillna(0)
df["Ocorrencia"] = pd.to_numeric(df[col_ocorr], errors="coerce")

df["Descricao da Ultima Ocorrencia"] = (
    df[col_desc_ultima].astype(str) if col_desc_ultima else ""
)

df["Desc_Ocorrencia"] = df["Descricao da Ultima Ocorrencia"].astype(str).str.lower()

df["Data_Ultima_Ocorrencia"] = (
    pd.to_datetime(df[col_data_ultima], errors="coerce")
    if col_data_ultima else pd.NaT
)

df["Dias_Saida_Entrega"] = (datetime.now() - df["Data_Ultima_Ocorrencia"]).dt.days

# =============================
# REGRAS
# =============================
df = df[~df["Ocorrencia"].isin([36, 87, 94, 99])]
df = df.sort_values(by=["NF"]).drop_duplicates(subset=["NF"], keep="last")

df["Status"] = "PENDENTE"

df.loc[df["Ocorrencia"].isin([50,51]), "Status"] = "SOBRA/FALTA"
df.loc[df["Ocorrencia"].isin([5,7,9,10,11,13,26,31,32,33,34,37]), "Status"] = "INSUCESSO"
df.loc[df["Ocorrencia"] == 85, "Status"] = "SAÍDA PARA ENTREGA"
df.loc[df["Tipo_Documento"] == "DEVOLUÇÃO", "Status"] = "DEVOLUÇÃO"

df["Entregue"] = (
    df["Ocorrencia"].isin([1]) |
    (
        df["Desc_Ocorrencia"].str.contains(
            r"\b(entregue|entrega realizada|baixado definitivo)\b",
            regex=True,
            na=False
        )
        &
        ~df["Desc_Ocorrencia"].str.contains(
            r"saida|rota|transferencia|em transito",
            na=False
        )
    )
)

df.loc[
    (df["Entregue"]) &
    (df["Dias_Atraso"] == 0) &
    (~df["Ocorrencia"].eq(85)) &
    (df["Status"] != "DEVOLUÇÃO"),
    "Status"
] = "ENTREGUE NO PRAZO"

df.loc[
    (df["Entregue"]) &
    (df["Dias_Atraso"] > 0) &
    (~df["Ocorrencia"].eq(85)) &
    (df["Status"] != "DEVOLUÇÃO"),
    "Status"
] = "ENTREGUE EM ATRASO"

# =============================
# PRIORIDADE
# =============================
df["Prioridade"] = "BAIXA"
df.loc[df["Ocorrencia"].isin([12,16,41]), "Prioridade"] = "CRITICA"
df.loc[df["Ocorrencia"].isin([50,51]), "Prioridade"] = "CRITICA"
df.loc[df["Ocorrencia"] == 85, "Prioridade"] = "MEDIA"
df.loc[(df["Status"] == "INSUCESSO") | (df["Dias_Atraso"] >= 3), "Prioridade"] = "ALTA"

# =============================
# FILTROS (VISUAL SOMENTE)
# =============================
st.sidebar.header("Filtros")

filtro_remetente = st.sidebar.multiselect("Remetente", df["Remetente"].dropna().unique())
filtro_praca = st.sidebar.multiselect("Praça", df["Praca"].dropna().unique())
filtro_status = st.sidebar.multiselect("Status", df["Status"].dropna().unique())

incluir_devolucao = st.sidebar.checkbox("Incluir devoluções", value=False)

df_base = df.copy()

if not incluir_devolucao:
    df_base = df_base[df_base["Tipo_Documento"] != "DEVOLUÇÃO"]

df_filtro = df_base.copy()

if filtro_remetente:
    df_filtro = df_filtro[df_filtro["Remetente"].isin(filtro_remetente)]

if filtro_praca:
    df_filtro = df_filtro[df_filtro["Praca"].isin(filtro_praca)]

if filtro_status:
    df_filtro = df_filtro[df_filtro["Status"].isin(filtro_status)]

df_filtro = df_filtro.sort_values(by="Data_Ultima_Ocorrencia", ascending=True)

# =============================
# 🔵 BASE ANALÍTICA (NÃO FILTRADA)
# =============================
df_analise = df_base.copy()

df_insucesso = df_analise[df_analise["Status"] == "INSUCESSO"]
df_sac = df_analise[df_analise["Ocorrencia"].isin([12,16,41])]
df_sobras = df_analise[df_analise["Ocorrencia"].isin([50,51])]
df_saida_entrega = df_analise[df_analise["Ocorrencia"] == 85]

# =============================
# COLUNAS
# =============================
colunas = [
    "NF","Remetente","Destinatario","Praca",
    "Tipo de Nota","Status","Prioridade",
    "Dias_Atraso","Descricao da Ultima Ocorrencia"
]

# =============================
# TABS
# =============================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📊 Dashboard","📋 Operacional","📊 Consolidado",
    "⚠️ Insucesso","📞 SAC","📦 Sobras/Faltas",
    "🚚 Saída para Entrega"
])

with tab1:
    st.metric("Total", len(df_filtro))
    st.metric("Pendentes", (df_filtro["Status"] == "PENDENTE").sum())

    fig, ax = plt.subplots()
    df_filtro["Status"].value_counts().plot(kind="bar", ax=ax)
    st.pyplot(fig)

with tab2:
    st.dataframe(df_filtro[colunas], use_container_width=True)

with tab3:
    df_group = df_filtro.groupby(["Praca","Status"]).size().unstack(fill_value=0)
    df_group["Total"] = df_group.sum(axis=1)
    st.dataframe(df_group.reset_index(), use_container_width=True)

with tab4:
    st.dataframe(df_insucesso[colunas], use_container_width=True)

with tab5:
    st.dataframe(df_sac[colunas], use_container_width=True)

with tab6:
    st.dataframe(df_sobras[colunas], use_container_width=True)

with tab7:
    st.dataframe(df_saida_entrega[colunas + ["Dias_Saida_Entrega"]], use_container_width=True)

# =============================
# EXPORT
# =============================
excel_bytes = gerar_excel({
    "Operacional": df_filtro[colunas],
    "Consolidado": df_group.reset_index(),
    "Insucesso": df_insucesso[colunas],
    "SAC": df_sac[colunas],
    "Sobras_Faltas": df_sobras[colunas],
    "Saida_Entrega": df_saida_entrega[colunas + ["Dias_Saida_Entrega"]]
})

st.download_button(
    "📥 Baixar Excel",
    data=excel_bytes,
    file_name="torre_controle.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")