import io
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

PASTA_DADOS = "DADOS"

st.set_page_config(
    layout="wide", page_title="Torre de Controle Logística", page_icon="🚚"
)
st.title("🚚 Torre de Controle Logística")


# =============================
# DICIONÁRIO DE CÓDIGOS SSW (BACKUP)
# =============================

DESCRICOES_SSW = {
    1: "ENTREGA REALIZADA",
    5: "REMETENTE RETEVE A MERCADORIA",
    7: "ENDEREÇO DO DESTINATÁRIO NÃO LOCALIZADO",
    9: "MERCADORIA RECUSADA - DEVOLUÇÃO",
    10: "DESTINATÁRIO MUDOU-SE",
    11: "VEÍCULO AVARIADO / QUEBRADO",
    12: "EXTRAVIO DE MERCADORIA",
    13: "DESTINATÁRIO AUSENTE OU FECHADO",
    15: "ENTREGA AGENDADA",
    16: "AVARIA TOTAL",
    26: "ESTABELECIMENTO FECHADO",
    31: "GREVE / PARALISAÇÃO NA REGIÃO",
    32: "ROUBO DE CARGA / CORTE",
    33: "APREENSÃO FISCAL / SEFAZ",
    34: "DADOS DO DESTINATÁRIO INCORRETOS",
    35: "AGUARDANDO AGENDAMENTO",
    37: "FALTA DE ESPAÇO NO DESTINATÁRIO",
    41: "AVARIA PARCIAL",
    50: "SOBRA DE MERCADORIA",
    51: "FALTA DE MERCADORIA",
    80: "SEM MOVIMENTAÇÃO / RETENÇÃO",
    84: "EM TRANSITO / NA UNIDADE",
    85: "SAÍDA PARA ENTREGA",
}


# =============================
# FUNÇÕES DE TRATAMENTO
# =============================


def normalizar(txt):
    if pd.isna(txt):
        return ""
    return (
        unicodedata.normalize("NFKD", str(txt))
        .encode("ASCII", "ignore")
        .decode()
        .lower()
        .strip()
    )


def achar_coluna_texto_pura(df, termos_busca):
    if isinstance(termos_busca, str):
        termos_busca = [termos_busca]

    for col in df.columns:
        col_norm = normalizar(col)

        if any(
            c in col_norm
            for c in ["cod", "codigo", "num", "numero", "dt", "data", "hora"]
        ):
            continue

        for termo in termos_busca:
            termo_norm = normalizar(termo)
            partes = termo_norm.split()

            if all(p in col_norm for p in partes):
                amostra = df[col].dropna()
                if amostra.empty:
                    continue

                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    continue

                amostra_str = amostra.astype(str)
                is_num = amostra_str.str.isnumeric().all()
                is_date = amostra_str.str.contains(
                    r"\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}", regex=True
                ).any()

                if not is_num and not is_date:
                    return col

    return None


def achar_coluna(df, termos):
    if isinstance(termos, str):
        termos = [termos]

    for col in df.columns:
        col_norm = normalizar(col)
        for termo in termos:
            termo_norm = normalizar(termo)
            partes = termo_norm.split()
            if all(p in col_norm for p in partes):
                return col

    return None


def coluna_excel_para_indice(letra):
    resultado = 0
    for char in letra.upper():
        resultado = resultado * 26 + (ord(char) - ord("A") + 1)
    return resultado - 1


def limpar_nome(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto)
    texto = re.sub(r"\d{2}.?\d{3}.?\d{3}/?\d{4}-?\d{2}", "", texto)
    return texto.strip().upper()


def gerar_excel(df_dict):
    output = io.BytesIO()

    try:
        engine_excel = "xlsxwriter"
    except ImportError:
        engine_excel = "openpyxl"

    with pd.ExcelWriter(output, engine=engine_excel) as writer:
        for nome, df_aba in df_dict.items():
            df_aba.to_excel(writer, index=False, sheet_name=nome)

    output.seek(0)
    return output


# =============================
# INPUT (UPLOAD OU PASTA)
# =============================

upload = st.sidebar.file_uploader("Upload", type=["xlsx"])


def arquivo_recente():
    if not os.path.exists(PASTA_DADOS):
        return None

    arquivos = [
        os.path.join(PASTA_DADOS, f)
        for f in os.listdir(PASTA_DADOS)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not arquivos:
        return None

    return max(arquivos, key=os.path.getmtime)


if upload is not None:
    caminho = upload
else:
    caminho = arquivo_recente()


if not caminho:
    st.error("Nenhum arquivo encontrado (upload ou pasta dados).")
    st.stop()


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


df = tratar_dados(ler_excel(caminho))


# =============================
# MAPEAMENTO RIGOROSO AJUSTADO
# =============================

col_nf = achar_coluna(df, ["nota fiscal", "num nota", "nota", "nf"])
col_doc = achar_coluna(df, ["documento", "tipo doc", "doc"])

col_setor = achar_coluna(
    df, ["setor de destino", "setor destino", "setor de entrega", "setor"]
)

col_atraso = achar_coluna(df, ["dias atraso", "atraso", "dias"])
col_ocorr = achar_coluna(df, ["cod ocorr", "codigo ocorrencia", "cod oc"])

col_data_ultima = achar_coluna(
    df,
    [
        "data da ultima ocorrencia",
        "data ultima ocorrencia",
        "dt ultima ocorrencia",
    ],
) or achar_coluna(df, ["data ocorrencia", "dt ocorrencia"])

col_desc_ultima = achar_coluna_texto_pura(
    df,
    [
        "descricao da ultima ocorrencia",
        "desc ultima ocorrencia",
        "descricao ocorrencia",
        "desc ocorrencia",
        "nome ocorrencia",
        "historico",
        "observacao",
    ],
)


def pegar_coluna_segura(df, letra):
    idx = coluna_excel_para_indice(letra)
    if idx >= len(df.columns):
        return None
    col = df.columns[idx]
    if df[col].isna().all():
        return None
    return col


# Colunas Fixas por Letra do Excel:
col_rem = pegar_coluna_segura(df, "P")         # Remetente
col_dest = pegar_coluna_segura(df, "AJ")      # Destinatário
col_unidade_receptora = pegar_coluna_segura(df, "BA")  # Unidade Receptora / Destino (Fixa na Coluna BA)
col_unidade = pegar_coluna_segura(df, "EB")   # Unidade da Ultima Ocorrência (Fixada na Coluna EB)


# =============================
# TRATAMENTO BASE
# =============================

df["NF"] = df[col_nf] if col_nf else "SEM NF"
df["Documento"] = df[col_doc] if col_doc else ""
df["Documento_str"] = df["Documento"].fillna("").astype(str).str.lower()
df["Ocorrencia"] = (
    pd.to_numeric(df[col_ocorr], errors="coerce").fillna(0) if col_ocorr else 0
)

if col_desc_ultima and col_desc_ultima in df.columns:
    df["Descricao da Ultima Ocorrencia"] = (
        df[col_desc_ultima].fillna("").astype(str)
    )
else:
    df["Descricao da Ultima Ocorrencia"] = df["Ocorrencia"].map(
        lambda cod: DESCRICOES_SSW.get(
            int(cod), f"OCORRÊNCIA CÓDIGO {int(cod)}"
        )
        if cod > 0
        else "SEM REGISTRO"
    )

df["Desc_Ocorrencia"] = df["Descricao da Ultima Ocorrencia"].str.lower()

df["Tipo_Documento"] = df.apply(
    lambda row: (
        "DEVOLUÇÃO"
        if (
            "devolucao" in str(row["Documento_str"])
            or any(
                p in str(row.get("Desc_Ocorrencia", "")).lower()
                for p in [
                    "devolucao",
                    "devolvida",
                    "retorno",
                    "avaria",
                    "recusa",
                    "insucesso",
                ]
            )
        )
        else "NORMAL"
    ),
    axis=1,
)

df["Tipo de Nota"] = df["Tipo_Documento"]
df["Remetente"] = (
    df[col_rem].fillna("").astype(str).apply(limpar_nome) if col_rem else "N/A"
)
df["Destinatario"] = (
    df[col_dest].fillna("").astype(str).apply(limpar_nome)
    if col_dest
    else "N/A"
)

df["Setor_Destino"] = df[col_setor] if col_setor else "NÃO INFORMADO"

# Puxando estritamente da Coluna BA (Unidade Receptora / Destino)
if col_unidade_receptora and col_unidade_receptora in df.columns:
    df["Unidade_Receptora"] = (
        df[col_unidade_receptora]
        .fillna("NÃO INFORMADA")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .str.upper()
    )
else:
    df["Unidade_Receptora"] = "NÃO INFORMADA"

# Puxando estritamente da Coluna EB (Unidade Atual / Última Ocorrência)
if col_unidade and col_unidade in df.columns:
    df["Unidade_Atual"] = (
        df[col_unidade]
        .fillna("NÃO INFORMADA")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .str.upper()
    )
else:
    df["Unidade_Atual"] = "NÃO INFORMADA"

df["Dias_Atraso"] = (
    pd.to_numeric(df[col_atraso], errors="coerce").fillna(0) if col_atraso else 0
)
df["Data_Ultima_Ocorrencia"] = (
    pd.to_datetime(df[col_data_ultima], errors="coerce")
    if col_data_ultima
    else pd.NaT
)
df["Dias_Sem_Movimento"] = (
    (datetime.now() - df["Data_Ultima_Ocorrencia"]).dt.days
    if col_data_ultima
    else 0
)


# =============================
# FILTRAGEM E DEDUPLICAÇÃO
# =============================

OCORRENCIAS_IGNORAR = [36, 87, 93, 94, 99]

df = (
    df[~df["Ocorrencia"].isin(OCORRENCIAS_IGNORAR)]
    .sort_values(by=["NF"])
    .drop_duplicates(subset=["NF"], keep="last")
    .copy()
)


# =============================
# REGRAS DE STATUS
# =============================

df["Status"] = "PENDENTE"
df.loc[df["Ocorrencia"].isin([50, 51]), "Status"] = "SOBRA/FALTA"
df.loc[
    df["Ocorrencia"].isin([5, 7, 9, 10, 11, 13, 26, 31, 32, 33, 34, 37]),
    "Status",
] = "INSUCESSO"
df.loc[df["Ocorrencia"] == 85, "Status"] = "SAÍDA PARA ENTREGA"
df.loc[df["Ocorrencia"] == 80, "Status"] = "SEM MOVIMENTAÇÃO"
df.loc[df["Ocorrencia"] == 84, "Status"] = "EM TRÂNSITO UNIDADE"
df.loc[df["Ocorrencia"].isin([15, 35]), "Status"] = "AGENDAMENTO"
df.loc[df["Tipo_Documento"] == "DEVOLUÇÃO", "Status"] = "DEVOLUÇÃO"

df["Status_Agendamento"] = "OUTROS"
df.loc[df["Ocorrencia"] == 35, "Status_Agendamento"] = "AGUARDANDO AGENDAMENTO"
df.loc[df["Ocorrencia"] == 15, "Status_Agendamento"] = "AGENDADO"

df["Entregue"] = df["Ocorrencia"].isin([1]) | (
    df["Desc_Ocorrencia"].str.contains(
        r"\b(?:entregue|entrega realizada|baixado definitivo)\b",
        regex=True,
        na=False,
    )
    & ~df["Desc_Ocorrencia"].str.contains(
        r"saida|rota|transferencia|em transito", regex=True, na=False
    )
)

df.loc[
    (df["Entregue"])
    & (df["Dias_Atraso"] == 0)
    & (~df["Ocorrencia"].isin([85, 80, 84]))
    & (df["Status"] != "DEVOLUÇÃO"),
    "Status",
] = "ENTREGUE NO PRAZO"

df.loc[
    (df["Entregue"])
    & (df["Dias_Atraso"] > 0)
    & (~df["Ocorrencia"].isin([85, 80, 84]))
    & (df["Status"] != "DEVOLUÇÃO"),
    "Status",
] = "ENTREGUE EM ATRASO"


# =============================
# PRIORIDADE
# =============================

df["Prioridade"] = "BAIXA"
df.loc[df["Ocorrencia"].isin([12, 16, 41]), "Prioridade"] = "CRITICA"
df.loc[df["Ocorrencia"].isin([50, 51]), "Prioridade"] = "CRITICA"
df.loc[df["Ocorrencia"] == 80, "Prioridade"] = "CRITICA"
df.loc[df["Ocorrencia"] == 35, "Prioridade"] = "ALTA"
df.loc[df["Ocorrencia"] == 15, "Prioridade"] = "MEDIA"
df.loc[df["Ocorrencia"] == 85, "Prioridade"] = "MEDIA"
df.loc[
    (df["Status"] == "INSUCESSO") | (df["Dias_Atraso"] >= 3), "Prioridade"
] = "ALTA"


# =============================
# FILTROS SOLICITADOS (SIDEBAR)
# =============================

st.sidebar.header("🔍 Filtros de Busca")

# 1. Pesquisa por Nota Fiscal
busca_nf = st.sidebar.text_input(
    "Buscar por Nota Fiscal (NF)", placeholder="Ex: 12345 ou 12345, 67890"
)

# Opções tratadas com astype(str) para evitar erro de ordenação (str vs int)
unidades_receptora_opts = sorted(df["Unidade_Receptora"].dropna().astype(str).unique())
unidades_opts = sorted(df["Unidade_Atual"].dropna().astype(str).unique())
remetentes_opts = sorted(df["Remetente"].dropna().astype(str).unique())
setores_opts = sorted(df["Setor_Destino"].dropna().astype(str).unique())
status_opts = sorted(df["Status"].dropna().astype(str).unique())

# 2. Filtro de Unidade Receptora / Destino (Coluna BA)
filtro_unidade_receptora = st.sidebar.multiselect(
    "🎯 Unidade Destino / Receptora", unidades_receptora_opts
)

# 3. Filtro de Unidade Atual (Coluna EB)
filtro_unidade = st.sidebar.multiselect(
    "🏢 Unidade Atual (Última Ocorrência)", unidades_opts
)

# 4. Filtro de Remetente
filtro_remetente = st.sidebar.multiselect("👤 Remetente", remetentes_opts)

# 5. Filtro de Setor de Destino
filtro_setor = st.sidebar.multiselect("📍 Setor de Destino", setores_opts)

filtro_status = st.sidebar.multiselect("Status", status_opts)
incluir_devolucao = st.sidebar.checkbox("Incluir devoluções", value=False)

df_base = df.copy()

if not incluir_devolucao:
    df_base = df_base[df_base["Tipo_Documento"] != "DEVOLUÇÃO"]

df_filtro = df_base.copy()

# Aplicação dos filtros dinâmicos
if busca_nf.strip():
    lista_nfs = [nf.strip() for nf in busca_nf.split(",") if nf.strip()]
    regex_nfs = "|".join(re.escape(nf) for nf in lista_nfs)
    df_filtro = df_filtro[
        df_filtro["NF"]
        .astype(str)
        .str.contains(regex_nfs, case=False, na=False)
    ]

if filtro_unidade_receptora:
    df_filtro = df_filtro[df_filtro["Unidade_Receptora"].astype(str).isin(filtro_unidade_receptora)]

if filtro_unidade:
    df_filtro = df_filtro[df_filtro["Unidade_Atual"].astype(str).isin(filtro_unidade)]

if filtro_remetente:
    df_filtro = df_filtro[df_filtro["Remetente"].astype(str).isin(filtro_remetente)]

if filtro_setor:
    df_filtro = df_filtro[df_filtro["Setor_Destino"].astype(str).isin(filtro_setor)]

if filtro_status:
    df_filtro = df_filtro[df_filtro["Status"].astype(str).isin(filtro_status)]

if "Data_Ultima_Ocorrencia" in df_filtro.columns:
    df_filtro = df_filtro.sort_values(
        by="Data_Ultima_Ocorrencia", ascending=True
    )


# =============================
# EXIBIÇÃO DE DADOS
# =============================

df_filtro_exibicao = df_filtro.copy()
if "Data_Ultima_Ocorrencia" in df_filtro_exibicao.columns:
    df_filtro_exibicao["Data da Ocorrencia"] = (
        df_filtro_exibicao["Data_Ultima_Ocorrencia"]
        .dt.strftime("%d/%m/%Y %H:%M")
        .fillna("-")
    )
else:
    df_filtro_exibicao["Data da Ocorrencia"] = "-"

colunas = [
    "NF",
    "Unidade_Receptora",
    "Unidade_Atual",
    "Remetente",
    "Destinatario",
    "Setor_Destino",
    "Status",
    "Prioridade",
    "Dias_Atraso",
    "Data da Ocorrencia",
    "Descricao da Ultima Ocorrencia",
]


# =============================
# TABS DE NAVEGAÇÃO
# =============================

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs(
    [
        "📊 Dashboard",
        "📋 Operacional",
        "📊 Consolidado",
        "🛑 Sem Movimento (80)",
        "📍 Unidade Atual (84)",
        "⚠️ Insucesso",
        "📞 SAC",
        "📦 Sobras/Faltas",
        "🚚 Saída para Entrega",
        "📅 Agendamento",
    ]
)


# =============================
# 1. DASHBOARD
# =============================

with tab1:
    total = len(df_filtro)

    entregues_prazo = (df_filtro["Status"] == "ENTREGUE NO PRAZO").sum()
    entregues_atraso = (df_filtro["Status"] == "ENTREGUE EM ATRASO").sum()
    pendentes = (df_filtro["Status"] == "PENDENTE").sum()
    devolucoes = (df_filtro["Status"] == "DEVOLUÇÃO").sum()
    insucessos = (df_filtro["Status"] == "INSUCESSO").sum()
    agendamentos = (df_filtro["Status"] == "AGENDAMENTO").sum()

    total_entregue = entregues_prazo + entregues_atraso
    pct_otif = (
        (entregues_prazo / total_entregue * 100) if total_entregue > 0 else 0
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("📦 Total de Notas", f"{total:,}")
    c2.metric("🎯 Nível de Serviço (OTIF)", f"{pct_otif:.1f}%")
    c3.metric("⏳ Pendentes em Trânsito", f"{pendentes:,}")
    c4.metric("📅 Agendamentos", f"{agendamentos:,}")
    c5.metric("⚠️ Insucessos", f"{insucessos:,}")
    c6.metric("↩️ Devoluções", f"{devolucoes:,}")

    st.divider()

    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.markdown("### 📊 Status das Entregas")
        if total > 0:
            status_counts = df_filtro["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Quantidade"]

            fig_donut = px.pie(
                status_counts,
                names="Status",
                values="Quantidade",
                hole=0.5,
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig_donut.update_traces(
                textposition="inside", textinfo="percent+label"
            )
            fig_donut.update_layout(
                showlegend=False, margin=dict(t=20, b=20, l=10, r=10)
            )
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info("Nenhum dado encontrado.")

    with g_col2:
        st.markdown("### 🚨 Prioridade de Atendimento")
        if total > 0:
            prio_counts = (
                df_filtro["Prioridade"].value_counts().reset_index()
            )
            prio_counts.columns = ["Prioridade", "Quantidade"]

            cores_prioridade = {
                "CRITICA": "#FF4B4B",
                "ALTA": "#FFA500",
                "MEDIA": "#FACA2B",
                "BAIXA": "#29B6F6",
            }

            fig_prio = px.bar(
                prio_counts,
                x="Prioridade",
                y="Quantidade",
                color="Prioridade",
                color_discrete_map=cores_prioridade,
                text="Quantidade",
            )
            fig_prio.update_traces(textposition="outside")
            fig_prio.update_layout(
                showlegend=False,
                xaxis_title="",
                yaxis_title="Notas",
                margin=dict(t=20, b=20, l=10, r=10),
            )
            st.plotly_chart(fig_prio, use_container_width=True)

    st.divider()

    st.markdown("### 🏢 Top Unidades com Maior Volume de Pendências/Atrasos")
    df_gargalo_unidade = df_filtro[
        df_filtro["Status"].isin(
            ["INSUCESSO", "ENTREGUE EM ATRASO", "PENDENTE", "SEM MOVIMENTAÇÃO"]
        )
    ]

    if not df_gargalo_unidade.empty:
        top_unidades = (
            df_gargalo_unidade.groupby(["Unidade_Atual", "Status"])
            .size()
            .reset_index(name="Quantidade")
            .sort_values(by="Quantidade", ascending=False)
            .head(15)
        )

        fig_unidade = px.bar(
            top_unidades,
            x="Quantidade",
            y="Unidade_Atual",
            color="Status",
            orientation="h",
            barmode="stack",
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig_unidade.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="Quantidade de NFs",
            yaxis_title="Unidade da Última Ocorrência (Coluna EB)",
        )
        st.plotly_chart(fig_unidade, use_container_width=True)
    else:
        st.success(
            "Nenhum gargalo identificado nas unidades para os filtros aplicados!"
        )


# =============================
# 2. OPERACIONAL
# =============================

with tab2:
    st.subheader("📋 Gestão Operacional")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 Total Registros", len(df_filtro))
    m2.metric(
        "✅ Entregues (Prazo)",
        (df_filtro["Status"] == "ENTREGUE NO PRAZO").sum(),
    )
    m3.metric(
        "🚨 Entregues (Atraso)",
        (df_filtro["Status"] == "ENTREGUE EM ATRASO").sum(),
    )
    m4.metric(
        "⏳ Em Trânsito (Pendente)", (df_filtro["Status"] == "PENDENTE").sum()
    )

    st.divider()
    st.dataframe(df_filtro_exibicao[colunas], use_container_width=True)


# =============================
# 3. CONSOLIDADO
# =============================

with tab3:
    st.subheader("📊 Gestão Consolidada por Unidade")

    df_group_unidade = (
        df_filtro.groupby(["Unidade_Receptora", "Unidade_Atual", "Status"])
        .size()
        .unstack(fill_value=0)
    )
    df_group_unidade["Total"] = df_group_unidade.sum(axis=1)

    m1, m2 = st.columns(2)
    m1.metric("🏢 Total de Agrupamentos", len(df_group_unidade))
    m2.metric("📦 Volume Total Processado", df_group_unidade["Total"].sum())

    st.divider()
    st.dataframe(df_group_unidade.reset_index(), use_container_width=True)


# =============================
# 4. SEM MOVIMENTAÇÃO (OCORRÊNCIA 80)
# =============================

with tab4:
    st.subheader("🛑 Gestão de Notas Sem Movimentação (Cód. 80)")

    df_sem_mov = df_filtro[df_filtro["Ocorrencia"] == 80]
    df_sem_mov_exib = df_filtro_exibicao[df_filtro_exibicao["Ocorrencia"] == 80]

    m1, m2, m3 = st.columns(3)
    m1.metric("🛑 Total Paradas (Cód. 80)", len(df_sem_mov))
    m2.metric(
        "⚠️ Paradas há > 3 dias", (df_sem_mov["Dias_Sem_Movimento"] > 3).sum()
    )
    m3.metric(
        "🚨 Paradas há > 7 dias", (df_sem_mov["Dias_Sem_Movimento"] > 7).sum()
    )

    st.divider()
    st.dataframe(
        df_sem_mov_exib[colunas + ["Dias_Sem_Movimento"]],
        use_container_width=True,
    )


# =============================
# 5. UNIDADE ATUAL (OCORRÊNCIA 84)
# =============================

with tab5:
    st.subheader("📍 Gestão de Localização por Unidade (Cód. 84)")

    df_unid = df_filtro[df_filtro["Ocorrencia"] == 84]
    df_unid_exib = df_filtro_exibicao[df_filtro_exibicao["Ocorrencia"] == 84]

    m1, m2 = st.columns(2)
    m1.metric("📍 Total de Cargas na Unidade (Cód. 84)", len(df_unid))
    m2.metric(
        "🏢 Total de Unidades Mapeadas", df_unid["Unidade_Atual"].nunique()
    )

    st.divider()
    st.dataframe(df_unid_exib[colunas], use_container_width=True)


# =============================
# 6. INSUCESSO
# =============================

with tab6:
    st.subheader("⚠️ Gestão de Insucessos")

    df_insucesso = df_filtro[df_filtro["Status"] == "INSUCESSO"]
    df_insuc_exib = df_filtro_exibicao[
        df_filtro_exibicao["Status"] == "INSUCESSO"
    ]

    m1, m2, m3 = st.columns(3)
    m1.metric("⚠️ Total de Insucessos", len(df_insucesso))
    m2.metric(
        "👤 Ausente/Fechado (Cód. 13/26)",
        df_insucesso["Ocorrencia"].isin([13, 26]).sum(),
    )
    m3.metric(
        "📍 Endereço/Dados (Cód. 7/34)",
        df_insucesso["Ocorrencia"].isin([7, 34]).sum(),
    )

    st.divider()
    st.dataframe(df_insuc_exib[colunas], use_container_width=True)


# =============================
# 7. SAC
# =============================

with tab7:
    st.subheader("📞 Gestão de Ocorrências SAC")

    df_sac = df_filtro[df_filtro["Ocorrencia"].isin([12, 16, 41])]
    df_sac_exib = df_filtro_exibicao[
        df_filtro_exibicao["Ocorrencia"].isin([12, 16, 41])
    ]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📞 Total de Casos SAC", len(df_sac))
    m2.metric("📦 Extravio (Cód. 12)", (df_sac["Ocorrencia"] == 12).sum())
    m3.metric("💥 Avaria Total (Cód. 16)", (df_sac["Ocorrencia"] == 16).sum())
    m4.metric("⚡ Avaria Parcial (Cód. 41)", (df_sac["Ocorrencia"] == 41).sum())

    st.divider()
    st.dataframe(df_sac_exib[colunas], use_container_width=True)


# =============================
# 8. SOBRAS/FALTAS
# =============================

with tab8:
    st.subheader("📦 Gestão de Sobras e Faltas")

    df_sobras = df_filtro[df_filtro["Ocorrencia"].isin([50, 51])]
    df_sobras_exib = df_filtro_exibicao[
        df_filtro_exibicao["Ocorrencia"].isin([50, 51])
    ]

    m1, m2, m3 = st.columns(3)
    m1.metric("📦 Total de Divergências", len(df_sobras))
    m2.metric("➕ Sobra de Carga (Cód. 50)", (df_sobras["Ocorrencia"] == 50).sum())
    m3.metric("➖ Falta de Carga (Cód. 51)", (df_sobras["Ocorrencia"] == 51).sum())

    st.divider()
    st.dataframe(df_sobras_exib[colunas], use_container_width=True)


# =============================
# 9. SAÍDA PARA ENTREGA
# =============================

with tab9:
    st.subheader("🚚 Gestão de Saída para Entrega")

    df_saida_entrega = df_filtro[df_filtro["Ocorrencia"] == 85]
    df_saida_exib = df_filtro_exibicao[df_filtro_exibicao["Ocorrencia"] == 85]

    m1, m2, m3 = st.columns(3)
    m1.metric("🚚 Total em Rota (Cód. 85)", len(df_saida_entrega))
    m2.metric(
        "🟢 Em rota hoje (< 1 dia)",
        (df_saida_entrega["Dias_Sem_Movimento"] <= 1).sum(),
    )
    m3.metric(
        "🔴 Em rota há > 2 dias",
        (df_saida_entrega["Dias_Sem_Movimento"] > 2).sum(),
    )

    st.divider()
    st.dataframe(
        df_saida_exib[colunas + ["Dias_Sem_Movimento"]],
        use_container_width=True,
    )


# =============================
# 10. AGENDAMENTO
# =============================

with tab10:
    st.subheader("📅 Gestão de Agendamentos")

    df_agendamento = df_filtro[df_filtro["Ocorrencia"].isin([15, 35])]
    df_agend_exib = df_filtro_exibicao[
        df_filtro_exibicao["Ocorrencia"].isin([15, 35])
    ]

    m1, m2, m3 = st.columns(3)
    qtd_aguardando = (df_agendamento["Ocorrencia"] == 35).sum()
    qtd_agendado = (df_agendamento["Ocorrencia"] == 15).sum()

    m1.metric("📦 Total de Agendamentos", len(df_agendamento))
    m2.metric("⏳ Aguardando Agendamento (Cód. 35)", qtd_aguardando)
    m3.metric("✅ Agendado (Cód. 15)", qtd_agendado)

    st.divider()
    st.dataframe(
        df_agend_exib[colunas + ["Status_Agendamento"]],
        use_container_width=True,
    )


# =============================
# EXPORT
# =============================

excel_bytes = gerar_excel(
    {
        "Operacional": df_filtro_exibicao[colunas],
        "Consolidado_Unidades": df_group_unidade.reset_index(),
        "Sem_Movimento_80": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"] == 80
        ][colunas + ["Dias_Sem_Movimento"]],
        "Unidade_Atual_84": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"] == 84
        ][colunas],
        "Insucesso": df_filtro_exibicao[
            df_filtro_exibicao["Status"] == "INSUCESSO"
        ][colunas],
        "SAC": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"].isin([12, 16, 41])
        ][colunas],
        "Sobras_Faltas": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"].isin([50, 51])
        ][colunas],
        "Saida_Entrega": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"] == 85
        ][colunas + ["Dias_Sem_Movimento"]],
        "Agendamento": df_filtro_exibicao[
            df_filtro_exibicao["Ocorrencia"].isin([15, 35])
        ][colunas + ["Status_Agendamento"]],
    }
)

st.download_button(
    "📥 Baixar Excel",
    data=excel_bytes,
    file_name="torre_controle_unidades.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
