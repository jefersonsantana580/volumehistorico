import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder


st.set_page_config(
    page_title="Visão de Volumes",
    layout="wide"
)

st.title("📊 Visão de Volumes por Site e Product DR")

# Ordem visual das colunas conforme o layout desejado
ORDEM_CICLOS = [
    "0+0 Bgt", "0+12", "01+11", "02+10", "03+9", "04+8",
    "05+7", "06+6", "07+5", "08+4", "09+3", "10+2", "11+1", "12+0"
]

ARQUIVO_EXCEL = "dados/base_volume_sites.xlsx"
ABA = "base"


@st.cache_data
def carregar_dados():
    df = pd.read_excel(
        ARQUIVO_EXCEL,
        sheet_name=ABA,
        engine="openpyxl"
    )

    # Limpeza dos nomes das colunas
    df.columns = df.columns.astype(str).str.strip()

    # Colunas de texto que vamos tratar
    colunas_texto = [
        "Tipo Base",
        "Nº CICLO",
        "SITE",
        "Product DR",
        "BRAND",
        "PRODUCT MARKET"
    ]

    for col in colunas_texto:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Normalização da coluna Tipo Base
    if "Tipo Base" in df.columns:
        df["Tipo Base"] = df["Tipo Base"].str.upper()

    # Normalização leve da coluna Nº CICLO
    if "Nº CICLO" in df.columns:
        df["Nº CICLO"] = df["Nº CICLO"].replace({
            "0+0 BGT": "0+0 Bgt",
            "0+0 bgt": "0+0 Bgt",
            "0+0 Bgt ": "0+0 Bgt"
        })

    # Garantir coluna numérica
    if "Total" in df.columns:
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)

    return df


def aplicar_filtro_opcional(df, coluna, valor):
    if valor == "Todos":
        return df
    return df[df[coluna] == valor]


try:
    df = carregar_dados()
except Exception as e:
    st.error(f"Erro ao carregar o arquivo Excel: {e}")
    st.stop()

# =========================
# Validação mínima
# =========================
colunas_necessarias = [
    "Tipo Base", "ANO", "BRAND", "PRODUCT MARKET",
    "SITE", "Product DR", "Nº CICLO", "Total"
]

colunas_faltantes = [c for c in colunas_necessarias if c not in df.columns]
if colunas_faltantes:
    st.error(f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltantes)}")
    st.stop()

# =========================
# Filtro fixo
# =========================
df = df[df["Tipo Base"] == "F_RESPONSE"].copy()

# =========================
# Filtros da tela
# =========================
st.subheader("Filtros")

col1, col2, col3 = st.columns(3)

anos = ["Todos"] + sorted(df["ANO"].dropna().unique().tolist())
brands = ["Todos"] + sorted(df["BRAND"].dropna().unique().tolist())
markets = ["Todos"] + sorted(df["PRODUCT MARKET"].dropna().unique().tolist())

with col1:
    ano_sel = st.selectbox("ANO", anos)

with col2:
    brand_sel = st.selectbox("BRAND", brands)

with col3:
    market_sel = st.selectbox("PRODUCT MARKET", markets)

df_filtrado = df.copy()
df_filtrado = aplicar_filtro_opcional(df_filtrado, "ANO", ano_sel)
df_filtrado = aplicar_filtro_opcional(df_filtrado, "BRAND", brand_sel)
df_filtrado = aplicar_filtro_opcional(df_filtrado, "PRODUCT MARKET", market_sel)


df_filtrado = df.copy()
df_filtrado = aplicar_filtro_opcional(df_filtrado, "ANO", ano_sel)
df_filtrado = aplicar_filtro_opcional(df_filtrado, "BRAND", brand_sel)
df_filtrado = aplicar_filtro_opcional(df_filtrado, "PRODUCT MARKET", market_sel)

# Remover Product DR = PC e CO para o site GENERAL RODRIGUEZ
df_filtrado = df_filtrado[
    ~(
        (df_filtrado["SITE"] == "GENERAL RODRIGUEZ") &
        (df_filtrado["Product DR"].isin(["PC", "CO"]))
    )
]

# =========================
# Montagem da tabela
# Linhas: SITE + Product DR
# Colunas: Nº CICLO
# Valor: Total
# =========================
if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

tabela = (
    df_filtrado
    .pivot_table(
        values="Total",
        index=["SITE", "Product DR"],
        columns="Nº CICLO",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

# Adiciona colunas faltantes para manter o layout fixo
for ciclo in ORDEM_CICLOS:
    if ciclo not in tabela.columns:
        tabela[ciclo] = 0

# Reordena as colunas no padrão visual da imagem
tabela = tabela[["SITE", "Product DR"] + ORDEM_CICLOS]

# Ordenação das linhas
tabela = tabela.sort_values(["SITE", "Product DR"]).reset_index(drop=True)




st.subheader("Tabela consolidada")
st.caption("Valor exibido: soma da coluna Total")
st.dataframe(tabela, use_container_width=True, hide_index=True)




# Mini gráfico abaixo da tabela
st.subheader("Mini gráfico de linha")
st.caption("Resumo do total por ciclo com base na tabela exibida acima")

serie_total = tabela[ORDEM_CICLOS].sum(axis=0)
chart_df = pd.DataFrame({
    "Ciclo": ORDEM_CICLOS,
    "Total": [serie_total.get(c, 0) for c in ORDEM_CICLOS]
}).set_index("Ciclo")

st.line_chart(chart_df, height=220, use_container_width=True)

with st.expander("Exibir linhas específicas no gráfico"):
    tabela_plot = tabela.copy()
    tabela_plot["Série"] = (
        tabela_plot["SITE"].astype(str) + " | " + tabela_plot["Product DR"].astype(str)
    )

    opcoes = tabela_plot["Série"].tolist()
    default = opcoes[:5]

    selecionadas = st.multiselect(
        "Escolha até 5 linhas para comparar",
        options=opcoes,
        default=default[: min(5, len(default))]
    )

    if selecionadas:
        dados_grafico = tabela_plot[tabela_plot["Série"].isin(selecionadas)].copy()
        dados_grafico = dados_grafico.set_index("Série")[ORDEM_CICLOS].T
        dados_grafico.index.name = "Ciclo"
        st.line_chart(dados_grafico, height=260, use_container_width=True)

# =========================
# Resumo rápido
# =========================
col_a, col_b = st.columns(2)

with col_a:
    st.metric("Linhas exibidas", len(tabela))

with col_b:
    soma_geral = tabela[ORDEM_CICLOS].sum().sum()
    st.metric("Soma geral", f"{soma_geral:,.0f}".replace(",", "."))
