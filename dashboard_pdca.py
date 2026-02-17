import streamlit as st
import pandas as pd
import plotly.express as px

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(
    page_title="Summary of Rejected Lot",
    layout="wide"
)

# =========================
# ESTILO GLOBAL (PRINT FRIENDLY)
# =========================
st.markdown("""
<style>
html, body, [class*="css"]  {
    font-size: 16px;
}
h1 {
    font-size: 34px !important;
}
h2 {
    font-size: 22px !important;
}
h3 {
    font-size: 20px !important;
}
[data-testid="stMetricValue"] {
    font-size: 30px;
}
[data-testid="stMetricLabel"] {
    font-size: 16px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# CABEÇALHO
# =========================
st.markdown(
    "<h1 style='text-align:center;'>SUMMARY OF REJECTED LOT</h1>"
    "<h2 style='text-align:center; letter-spacing:4px;'>GRUPO MK</h2>",
    unsafe_allow_html=True
)

st.divider()

# =========================
# LEITURA DO EXCEL (READ ONLY)
# =========================
arquivo = r"C:\Users\eebarreto\OneDrive - MK BR SA (1)\Automacao_Liberacao\ANALISES GERAIS.xlsx"

df = pd.read_excel(
    arquivo,
    sheet_name="PDCA",
    header=2
)

# =========================
# NORMALIZAÇÃO DAS COLUNAS
# =========================
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.lower()
    .str.replace("\n", " ")
    .str.replace("  ", " ")
)

mapa = {
    "sent date": "date",
    "inspetor china": "inspector",
    "who": "engineer",
    "part code": "code",
    "part name": "part",
    "qty of lot": "qty_lot",
    "what": "what"
}

df = df[list(mapa.keys())].rename(columns=mapa)

# =========================
# DATAS
# =========================
df["date"] = pd.to_datetime(df["date"], errors="coerce")
df = df.dropna(subset=["date"])

df["year"] = df["date"].dt.year
df["month"] = df["date"].dt.month
df["month_name"] = df["date"].dt.strftime("%B")

# =========================
# FILTROS
# =========================
st.sidebar.header("Filtros")

ano = st.sidebar.selectbox(
    "Ano",
    sorted(df["year"].unique(), reverse=True)
)

meses_disp = (
    df[df["year"] == ano][["month", "month_name"]]
    .drop_duplicates()
    .sort_values("month")
)

meses_selecionados = st.sidebar.multiselect(
    "Meses",
    meses_disp["month_name"].tolist(),
    default=meses_disp["month_name"].tolist()
)

df_filtro = df[
    (df["year"] == ano) &
    (df["month_name"].isin(meses_selecionados))
]

# =========================
# KPIs (COMPACTOS)
# =========================
c1, c2, c3 = st.columns(3)

c1.metric("Rejected Lots", len(df_filtro))
c2.metric("Inspectors", df_filtro["inspector"].nunique())
c3.metric("Engineers", df_filtro["engineer"].nunique())

# =========================
# GRÁFICOS (SUBIR NA TELA)
# =========================
col1, col2 = st.columns(2)

insp = (
    df_filtro["inspector"]
    .value_counts(normalize=True)
    .mul(100)
    .reset_index()
)
insp.columns = ["Inspector", "Percent"]

fig_insp = px.bar(
    insp,
    x="Percent",
    y="Inspector",
    orientation="h",
    text="Percent",
    title="% per Inspector",
    template="plotly_dark",
    height=300
)
fig_insp.update_traces(
    texttemplate="%{text:.1f}%",
    textposition="inside",
    textfont_size=16
)

col1.plotly_chart(fig_insp, use_container_width=True)

eng = (
    df_filtro["engineer"]
    .value_counts(normalize=True)
    .mul(100)
    .reset_index()
)
eng.columns = ["Engineer", "Percent"]

fig_eng = px.bar(
    eng,
    x="Percent",
    y="Engineer",
    orientation="h",
    text="Percent",
    title="% per Engineer",
    template="plotly_dark",
    height=300
)
fig_eng.update_traces(
    texttemplate="%{text:.1f}%",
    textposition="inside",
    textfont_size=16
)

col2.plotly_chart(fig_eng, use_container_width=True)

# =========================
# TABELA (SEM ESTOURAR TELA)
# =========================
st.subheader("Detailed Rejected Lots")

st.dataframe(
    df_filtro[
        ["date", "inspector", "engineer", "code", "part", "qty_lot", "what"]
    ].sort_values("date"),
    use_container_width=True,
    height=260
)
