
import streamlit as st
import pandas as pd
import io
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import plotly.graph_objects as go


# =========================
# Configuração da página
# =========================
st.set_page_config(
    page_title="Painel Gerencial de Volumes",
    layout="wide"
)

st.title("📊 Painel Gerencial de Volumes por Filial e Product DR")

# Ordem visual das colunas
ORDEM_CICLOS = [
    "0+0 Bgt", "0+12", "01+11", "02+10", "03+09", "04+08",
    "05+07", "06+06", "07+05", "08+04", "09+03", "10+02", "11+01", "12+0"
]

ARQUIVO_EXCEL = "dados/base_volume_sites.xlsx"
ABA = "base"

# mesma cor para o mesmo Product DR em qualquer lugar
cores_produtos = {
    "DF": "#1F77B4",      # azul
    "MOM": "#6DC8A0",     # verde água
    "RIG": "#6F4C9B",     # roxo
    "TA": "#D62728",      # vermelho
    "PU": "#2CA02C",      # verde
    "CO": "#FF7F0E",      # laranja
    "CO PKD": "#8C564B"   # marrom
}


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


def aplicar_filtro_multiplos(df, coluna, valores):
    if not valores:
        return df
    return df[df[coluna].isin(valores)]


def render_checkbox_filter(label, options, key_prefix):
    qtd_sel = st.session_state.get(f"{key_prefix}_count", 0)
    titulo = f"{label} ({qtd_sel})" if qtd_sel > 0 else label

    selecionados = []

    with st.popover(titulo, use_container_width=True):
        st.caption(f"Selecione um ou mais itens de {label}")
        for opt in options:
            opt_str = str(opt)
            if st.checkbox(opt_str, key=f"{key_prefix}_{opt_str}"):
                selecionados.append(opt)

    st.session_state[f"{key_prefix}_count"] = len(selecionados)
    return selecionados


def resumo_filtro(valores):
    if not valores:
        return "Todos"
    valores_str = [str(v) for v in valores]
    if len(valores_str) <= 3:
        return ", ".join(valores_str)
    return ", ".join(valores_str[:3]) + f" +{len(valores_str) - 3}"


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
    x_start = 25
    y = height - 35

    # Cabeçalho
    x = x_start
    for col in df.columns:
        c.drawString(x, y, str(col))
        x += 45

    y -= 15

    # Linhas
    for _, row in df.iterrows():
        x = x_start
        for value in row:
            c.drawString(x, y, str(value))
            x += 45
        y -= 12

        if y < 40:
            c.showPage()
            c.setFont("Helvetica", 8)
            y = height - 35

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


# =========================
# Aplicação dos filtros
# =========================
df_filtrado = df.copy()

# padronização
df_filtrado["SITE"] = df_filtrado["SITE"].astype(str).str.strip().str.upper()
df_filtrado["Product DR"] = df_filtrado["Product DR"].astype(str).str.strip().str.upper()
df_filtrado["BRAND"] = df_filtrado["BRAND"].astype(str).str.strip()
df_filtrado["PRODUCT MARKET"] = df_filtrado["PRODUCT MARKET"].astype(str).str.strip()

# filtros múltiplos
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

if df_filtrado.empty:
    st.warning("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()


# =========================
# Contexto do painel
# =========================
st.caption(
    f"**Filtros aplicados** | "
    f"ANO: {resumo_filtro(ano_sel)} | "
    f"BRAND: {resumo_filtro(brand_sel)} | "
    f"PRODUCT MARKET: {resumo_filtro(market_sel)} | "
    f"PRODUCT DR: {resumo_filtro(product_dr_sel)} | "
    f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
)


# =========================
# Pivot base
# =========================
tabela_base = (
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
    if ciclo not in tabela_base.columns:
        tabela_base[ciclo] = 0

tabela_base = tabela_base[["SITE", "Product DR"] + ORDEM_CICLOS]
tabela_base = tabela_base.sort_values(["SITE", "Product DR"]).reset_index(drop=True)


# =========================
# Controles gerenciais
# =========================
st.divider()

col_ctrl1, col_ctrl2 = st.columns([2, 1])

with col_ctrl1:
    mostrar_apenas_ciclos_com_volume = st.checkbox(
        "Mostrar apenas ciclos com volume",
        value=True
    )

with col_ctrl2:
    st.markdown("")


# =========================
# Cálculos de apoio
# =========================
ciclos_com_volume = [c for c in ORDEM_CICLOS if tabela_base[c].sum() > 0]

if mostrar_apenas_ciclos_com_volume:
    ciclos_exibicao = ciclos_com_volume.copy()
else:
    ciclos_exibicao = ORDEM_CICLOS.copy()

if not ciclos_exibicao:
    ciclos_exibicao = ORDEM_CICLOS.copy()

tabela = tabela_base[["SITE", "Product DR"] + ciclos_exibicao].copy()

volume_total = df_filtrado["Total"].sum()
filiais_ativas = df_filtrado["SITE"].nunique()
produtos_ativos = df_filtrado["Product DR"].nunique()
ultimo_ciclo_com_volume = ciclos_com_volume[-1] if ciclos_com_volume else "-"


# =========================
# KPIs
# =========================
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    st.metric("Volume total", f"{volume_total:,.0f}".replace(",", "."))

with kpi2:
    st.metric("Filiais ativas", filiais_ativas)

with kpi3:
    st.metric("Produtos ativos", produtos_ativos)

with kpi4:
    st.metric("Último ciclo com volume", ultimo_ciclo_com_volume)


# =========================
# Resumo executivo
# =========================
st.divider()
st.subheader("📌 Resumo executivo")

res1, res2 = st.columns(2)

# Volume total por filial
with res1:
    vol_filial = (
        df_filtrado
        .groupby("SITE", as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=True)
    )

    fig_filial = go.Figure()

    fig_filial.add_trace(
        go.Bar(
            x=vol_filial["Total"],
            y=vol_filial["SITE"],
            orientation="h",
            marker_color="#1F77B4",
            text=[f"{int(v):,}".replace(",", ".") for v in vol_filial["Total"]],
            textposition="outside",
            hovertemplate="Filial: %{y}<br>Volume: %{x:,.0f}<extra></extra>"
        )
    )

    fig_filial.update_layout(
        title=dict(
            text="Volume total por filial",
            x=0.02,
            xanchor="left",
            font=dict(color="black", size=18)
        ),
        height=350,
        margin=dict(l=20, r=20, t=50, b=20),
        xaxis=dict(
            title="Volume",
            showgrid=False,
            showticklabels=False
        ),
        yaxis=dict(
            title="",
            tickfont=dict(color="black", size=11),
            showgrid=False
        ),
        font=dict(color="black"),
        plot_bgcolor="white",
        paper_bgcolor="white"
    )

    st.plotly_chart(fig_filial, use_container_width=True)

# Participação por Product DR
with res2:
    vol_produto = (
        df_filtrado
        .groupby("Product DR", as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
    )

    fig_produto = go.Figure()

    fig_produto.add_trace(
        go.Pie(
            labels=vol_produto["Product DR"],
            values=vol_produto["Total"],
            hole=0.55,
            textinfo="label+percent",
            marker=dict(
                colors=[cores_produtos.get(p, "#999999") for p in vol_produto["Product DR"]]
            ),
            hovertemplate="Produto: %{label}<br>Volume: %{value:,.0f}<br>Participação: %{percent}<extra></extra>"
        )
    )

    fig_produto.update_layout(
        title=dict(
            text="Participação por Product DR",
            x=0.02,
            xanchor="left",
            font=dict(color="black", size=18)
        ),
        height=350,
        margin=dict(l=20, r=20, t=50, b=20),
        font=dict(color="black"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(
            title="",
            orientation="h",
            yanchor="bottom",
            y=-0.10,
            xanchor="center",
            x=0.5,
            font=dict(color="black", size=10)
        )
    )

    st.plotly_chart(fig_produto, use_container_width=True)


# =========================
# Gráficos por filial
# =========================
st.divider()
st.subheader("📈 Gráficos por filial")
st.caption("Visual comparativo por ciclo e produto.")

sites_unicos = sorted(tabela["SITE"].dropna().unique().tolist())

if not sites_unicos:
    st.info("Nenhuma filial encontrada para exibir gráficos.")
else:
    sites_graficos = sites_unicos[:5]

    for i in range(0, len(sites_graficos), 2):
        cols = st.columns(2)

        for j, site in enumerate(sites_graficos[i:i+2]):
            with cols[j]:
                df_site = tabela_base[tabela_base["SITE"] == site].copy()

                # define os ciclos do gráfico
                if mostrar_apenas_ciclos_com_volume:
                    ciclos_site = [c for c in ciclos_exibicao if df_site[c].sum() > 0]
                else:
                    ciclos_site = ciclos_exibicao.copy()

                if not ciclos_site:
                    st.info(f"{site}: sem volume para exibir.")
                    continue

                df_plot = df_site.set_index("Product DR")[ciclos_site].T

                if df_plot.empty:
                    st.info(f"{site}: sem volume para exibir.")
                    continue

                ordem_fixa_produtos = ["DF", "MOM", "RIG", "TA", "PU", "CO", "CO PKD"]
                produtos_existentes = [p for p in ordem_fixa_produtos if p in df_plot.columns]
                outros_produtos = [p for p in df_plot.columns if p not in produtos_existentes]
                df_plot = df_plot[produtos_existentes + outros_produtos]

                fig = go.Figure()

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
# Tabela detalhada
# =========================
st.divider()
st.subheader("📋 Tabela detalhada")
st.caption(
    f"Valor exibido: soma da coluna Total | Ciclos exibidos: {', '.join(ciclos_exibicao)}"
)
st.dataframe(tabela, use_container_width=True, hide_index=True)


# =========================
# Download
# =========================
st.divider()
st.subheader("Download da tabela")

with st.popover("📥 Baixar dados"):
    st.write("Escolha o formato:")

    excel_file = gerar_excel(tabela)
    st.download_button(
        label="📗 Excel",
        data=excel_file,
        file_name="painel_gerencial_volumes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.download_button(
        label="📕 PDF",
        data=gerar_pdf(tabela),
        file_name="painel_gerencial_volumes.pdf",
        mime="application/pdf",
        use_container_width=True
    )
