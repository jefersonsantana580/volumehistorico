
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder

import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet



# =========================
# Configuração da página
# =========================
st.set_page_config(
    page_title="Visão de Volumes",
    layout="wide"
)

st.title("📊 Visão de Volumes por Site e Product DR")

# Ordem visual das colunas
ORDEM_CICLOS = [
    "0+0 Bgt", "0+12", "01+11", "02+10", "03+09", "04+08",
    "05+07", "06+06", "07+05", "08+04", "09+03", "10+02", "11+01", "12+0"
]

ARQUIVO_EXCEL = "dados/base_volume_sites.xlsx"
ABA = "base"


# =========================
# Funções
# =========================
@st.cache_data
def carregar_dados():
    df = pd.read_excel(
        ARQUIVO_EXCEL,
        sheet_name=ABA,
        engine="openpyxl"
    )

    df.columns = df.columns.astype(str).str.strip()

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

    if "Tipo Base" in df.columns:
        df["Tipo Base"] = df["Tipo Base"].str.upper()

    if "Nº CICLO" in df.columns:
        df["Nº CICLO"] = df["Nº CICLO"].replace({
            "0+0 BGT": "0+0 Bgt",
            "0+0 bgt": "0+0 Bgt",
            "0+0 Bgt ": "0+0 Bgt"
        })

    if "Total" in df.columns:
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)

    return df


def aplicar_filtro_opcional(df, coluna, valor):
    if valor == "Todos":
        return df
    return df[df[coluna] == valor]


# ===== Exportação =====
def gerar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tabela")
    output.seek(0)
    return output


def gerar_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    c.setFont("Helvetica", 8)
    x_start = 30
    y = height - 40

    # Cabeçalho
    x = x_start
    for col in df.columns:
        c.drawString(x, y, str(col))
        x += 50

    y -= 15

    # Linhas
    for _, row in df.iterrows():
        x = x_start
        for value in row:
            c.drawString(x, y, str(value))
            x += 50
        y -= 12

        if y < 40:
            c.showPage()
            c.setFont("Helvetica", 8)
            y = height - 40

    c.save()
    buffer.seek(0)
    return buffer


# =========================
# Carga de dados
# =========================
try:
    df = carregar_dados()
except Exception as e:
    st.error(f"Erro ao carregar o arquivo Excel: {e}")
    st.stop()


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

df_filtrado["SITE"] = df_filtrado["SITE"].astype(str).str.strip().str.upper()
df_filtrado["Product DR"] = df_filtrado["Product DR"].astype(str).str.strip().str.upper()

df_filtrado = df_filtrado[
    ~(
        (df_filtrado["Product DR"] == "PC") |
        (
            (df_filtrado["SITE"] == "GENERAL RODRIGUEZ") &
            (df_filtrado["Product DR"] == "CO")
        )
    )
]


# =========================
# Pivot da tabela
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

for ciclo in ORDEM_CICLOS:
    if ciclo not in tabela.columns:
        tabela[ciclo] = 0

tabela = tabela[["SITE", "Product DR"] + ORDEM_CICLOS]
tabela = tabela.sort_values(["SITE", "Product DR"]).reset_index(drop=True)


# =========================
# Exibição
# =========================
st.subheader("Tabela consolidada")
st.caption("Valor exibido: soma da coluna Total")
st.dataframe(tabela, use_container_width=True, hide_index=True)


# =========================
# Download
# =========================

st.divider()
st.subheader("Download da tabela")

with st.popover("📥 Baixar dados"):
    st.write("Escolha o formato:")

    # Excel (ok gerar antes)
    excel_file = gerar_excel(tabela)
    st.download_button(
        label="📗 Excel",
        data=excel_file,
        file_name="visao_volumes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # PDF (geração lazy – só quando clicar)
    st.download_button(
        label="📕 PDF",
        data=gerar_pdf(tabela),
        file_name="visao_volumes.pdf",
        mime="application/pdf",
        use_container_width=True
    )
   



# =========================
# Resumo
# =========================
col_a, col_b = st.columns(2)

with col_a:
    st.metric("Linhas exibidas", len(tabela))

with col_b:
    soma_geral = tabela[ORDEM_CICLOS].sum().sum()
    st.metric("Soma geral", f"{soma_geral:,.0f}".replace(",", "."))
