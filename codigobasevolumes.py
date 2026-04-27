
import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder

import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import matplotlib.pyplot as plt
import plotly.graph_objects as go


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


def aplicar_filtro_multiplos(df, coluna, valores):
    if not valores:
        return df
    return df[df[coluna].isin(valores)]


def render_checkbox_filter(label, options, key_prefix):
    selecionados = []

    with st.popover(label, use_container_width=True):
        st.caption(f"Selecione um ou mais itens de {label}")

        for opt in options:
            opt_str = str(opt)
            if st.checkbox(opt_str, key=f"{key_prefix}_{opt_str}"):
                selecionados.append(opt)

    return selecionados


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

col1, col2, col3, col4 = st.columns(4)

anos = sorted(df["ANO"].dropna().unique().tolist())
brands = sorted(df["BRAND"].dropna().astype(str).str.strip().unique().tolist())
markets = sorted(df["PRODUCT MARKET"].dropna().astype(str).str.strip().unique().tolist())
product_drs = sorted(
    df["Product DR"]
    .dropna()
    .astype(str)
    .str.strip()
    .str.upper()
    .unique()
    .tolist()
)

with col1:
    ano_sel = render_checkbox_filter("ANO", anos, "ano")

with col2:
    brand_sel = render_checkbox_filter("BRAND", brands, "brand")

with col3:
    market_sel = render_checkbox_filter("PRODUCT MARKET", markets, "market")

with col4:
    product_dr_sel = render_checkbox_filter("PRODUCT DR", product_drs, "productdr")


df_filtrado = df.copy()

# padronização antes dos filtros textuais
df_filtrado["SITE"] = df_filtrado["SITE"].astype(str).str.strip().str.upper()
df_filtrado["Product DR"] = df_filtrado["Product DR"].astype(str).str.strip().str.upper()
df_filtrado["BRAND"] = df_filtrado["BRAND"].astype(str).str.strip()
df_filtrado["PRODUCT MARKET"] = df_filtrado["PRODUCT MARKET"].astype(str).str.strip()

# aplica filtros múltiplos
df_filtrado = aplicar_filtro_multiplos(df_filtrado, "ANO", ano_sel)
df_filtrado = aplicar_filtro_multiplos(df_filtrado, "BRAND", brand_sel)
df_filtrado = aplicar_filtro_multiplos(df_filtrado, "PRODUCT MARKET", market_sel)
df_filtrado = aplicar_filtro_multiplos(df_filtrado, "Product DR", product_dr_sel)

# regras de exclusão
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
# Gráficos por filial
# =========================
st.divider()
st.subheader("📈 Gráficos por filial")
st.caption("Cada gráfico mostra os volumes por ciclo e produto.")

# mesma cor para o mesmo Product DR em qualquer filial
cores_produtos = {
    "DF": "#1F77B4",      # azul forte
    "MOM": "#6DC8A0",     # verde água
    "RIG": "#6F4C9B",     # roxo
    "TA": "#D62728",      # vermelho
    "PU": "#2CA02C",      # verde
    "CO": "#FF7F0E",      # laranja
    "CO PKD": "#8C564B"   # marrom
}

sites_unicos = sorted(tabela["SITE"].dropna().unique().tolist())

if not sites_unicos:
    st.info("Nenhuma filial encontrada para exibir gráficos.")
else:
    sites_graficos = sites_unicos[:5]

    for i in range(0, len(sites_graficos), 2):
        cols = st.columns(2)

        for j, site in enumerate(sites_graficos[i:i+2]):
            with cols[j]:
                df_site = tabela[tabela["SITE"] == site].copy()

                # linhas = ciclos / colunas = Product DR
                df_plot = df_site.set_index("Product DR")[ORDEM_CICLOS].T

                if df_plot.empty:
                    st.info(f"{site}: sem volume para exibir.")
                    continue

                # ordem fixa dos produtos
                ordem_fixa_produtos = ["DF", "MOM", "RIG", "TA", "PU", "CO", "CO PKD"]
                produtos_existentes = [p for p in ordem_fixa_produtos if p in df_plot.columns]
                outros_produtos = [p for p in df_plot.columns if p not in produtos_existentes]
                df_plot = df_plot[produtos_existentes + outros_produtos]

                fig = go.Figure()

                # barras empilhadas por produto
                for produto in df_plot.columns:
                    valores = df_plot[produto].fillna(0)

                    fig.add_trace(
                        go.Bar(
                            x=df_plot.index.tolist(),
                            y=valores.tolist(),
                            name=produto,
                            marker_color=cores_produtos.get(produto, "#666666"),
                            text=[f"{int(v):,}".replace(",", ".") if v > 0 else "" for v in valores],
                            textposition="inside",
                            insidetextanchor="middle",
                            textfont=dict(color="white", size=10),
                            hovertemplate=(
                                f"<b>{site}</b><br>"
                                "Ciclo: %{x}<br>"
                                f"Produto: {produto}<br>"
                                "Volume: %{y:,.0f}<extra></extra>"
                            )
                        )
                    )

                fig.update_layout(
                    title=dict(
                        text=f"{site}",
                        x=0.02,
                        xanchor="left",
                        font=dict(color="black", size=18)
                    ),
                    barmode="stack",
                    height=300,
                    margin=dict(l=20, r=20, t=45, b=30),
                    xaxis=dict(
                        title=dict(text="Nº CICLO", font=dict(color="black", size=12)),
                        tickfont=dict(color="black", size=11),
                        showgrid=False,
                        tickangle=0
                    ),
                    yaxis=dict(
                        title=dict(text="", font=dict(color="black", size=12)),
                        tickfont=dict(color="black", size=11),
                        showticklabels=False,
                        showgrid=False,
                        zeroline=False
                    ),
                    legend=dict(
                        title=dict(text="Product DR", font=dict(color="black", size=11)),
                        font=dict(color="black", size=10),
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    font=dict(color="black"),
                    plot_bgcolor="white",
                    paper_bgcolor="white"
                )

                st.plotly_chart(fig, use_container_width=True)


# =========================
# Download
# =========================
st.divider()
st.subheader("Download da tabela")

with st.popover("📥 Baixar dados"):
    st.write("Escolha o formato:")

    # Excel
    excel_file = gerar_excel(tabela)
    st.download_button(
        label="📗 Excel",
        data=excel_file,
        file_name="visao_volumes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # PDF
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

