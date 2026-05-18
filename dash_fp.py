"""
Dashboard Financeiro Pessoal
Análise das guias DB_DESPESAS e BD_BudgetPessoal

Requisitos:
    pip install streamlit plotly pandas openpyxl

Execução:
    streamlit run dashboard.py
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from pathlib import Path

# ─── Configuração da Página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Financeiro Pessoal",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS Customizado ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f1117; }
    .stMetric { background: #1e2130; border-radius: 10px; padding: 12px; }
    .stMetric label { color: #a0aab4 !important; font-size: 13px; }
    .stMetric [data-testid="metric-container"] { background: #1e2130; border-radius: 10px; }
    h1, h2, h3 { color: #e8eaf0 !important; }
    .section-title {
        font-size: 18px; font-weight: 600; color: #7c83ff;
        margin: 20px 0 10px; border-bottom: 1px solid #2d3250; padding-bottom: 6px;
    }
    [data-testid="stSidebar"] { background-color: #141726; }
</style>
""", unsafe_allow_html=True)

# ─── Paleta de Cores ───────────────────────────────────────────────────────────
COLORS = {
    "primary":    "#7c83ff",
    "success":    "#4caf7d",
    "danger":     "#f05454",
    "warning":    "#f5a623",
    "secondary":  "#a0aab4",
    "bg":         "#1e2130",
    "grid":       "#2d3250",
}

CAT_PALETTE = [
    "#7c83ff","#4caf7d","#f05454","#f5a623","#56cfe1",
    "#ff7096","#c77dff","#06d6a0","#ff9f1c","#2ec4b6",
    "#e76f51","#457b9d","#a8dadc","#ffd166","#ef476f",
    "#118ab2","#06d6a0","#ffd166","#ef476f","#073b4c",
]

GROUP_LABELS = {
    "D.P.": "Despesas Pessoais",
    "D.T.": "Despesas de Transporte/Fixas",
    "D.F.": "Despesas Financeiras",
    "PGT.": "Pagamentos",
    "Vend": "Vendas/Receitas",
}

# ─── Carregamento de Dados ─────────────────────────────────────────────────────
@st.cache_data
def load_data(filepath: str):
    xl = pd.ExcelFile(filepath)

    # ── DB_DESPESAS ──
    sheet_name = "DB_DESPESAS" if "DB_DESPESAS" in xl.sheet_names else "DB_DESPESAS"
    df = pd.read_excel(xl, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    df["Data Lançamento"] = pd.to_datetime(df["Data Lançamento"], errors="coerce")
    df["Data Base"]       = pd.to_datetime(df["Data Base"],       errors="coerce")
    df["Saída(R$)"]       = pd.to_numeric(df["Saída(R$)"],  errors="coerce").abs().fillna(0)
    df["Entrada(R$)"]     = pd.to_numeric(df["Entrada(R$)"],errors="coerce").abs().fillna(0)
    df["AnoMes"]          = df["Data Base"].dt.to_period("M").astype(str)
    df["Mês"]             = df["Data Base"].dt.strftime("%b/%y")
    df["GRUPO LABEL"]     = df["GRUPO REAL"].map(GROUP_LABELS).fillna(df["GRUPO REAL"])

    # ── BD_BudgetPessoal ──
    bdf = pd.read_excel(xl, sheet_name="BD_BudgetPessoal", header=0)
    bdf.columns = bdf.columns.str.strip()
    bdf["Data Contábil"]     = pd.to_datetime(bdf["Data Contábil"],     errors="coerce")
    bdf["Entrada Real"]      = pd.to_numeric(bdf["Entrada Real"],      errors="coerce").fillna(0)
    bdf["Entrada Esperada"]  = pd.to_numeric(bdf["Entrada Esperada"],  errors="coerce").fillna(0)

    # Converter Data Base p/ período
    def parse_period(val):
        try:
            parts = str(val).split("/")
            m, y = int(parts[0]), int(parts[1])
            return pd.Period(year=y, month=m, freq="M")
        except Exception:
            return pd.NaT
    bdf["Período"] = bdf["Data Base"].apply(parse_period)
    bdf["AnoMes"]  = bdf["Período"].astype(str)
    bdf["Mês"]     = bdf["Data Contábil"].dt.strftime("%b/%y")

    return df, bdf


# ─── Caminho do arquivo ────────────────────────────────────────────────────────
DEFAULT_FILE = Path(__file__).parent / "CC - PT v3.xlsx"

with st.sidebar:
    st.markdown("## 📂 Fonte de Dados")
    uploaded = st.file_uploader("Upload do arquivo Excel", type=["xlsx"])
    if uploaded:
        import tempfile, shutil
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        shutil.copyfileobj(uploaded, tmp)
        tmp.flush()
        filepath = tmp.name
    elif DEFAULT_FILE.exists():
        filepath = str(DEFAULT_FILE)
    else:
        st.error("Nenhum arquivo encontrado. Faça o upload do Excel.")
        st.stop()

df, bdf = load_data(filepath)

# ─── Sidebar – Filtros ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("## 🔍 Filtros")

    all_months = sorted(df["AnoMes"].dropna().unique())
    # Default: últimos 12 meses com dados
    default_start = all_months[-12] if len(all_months) >= 12 else all_months[0]
    default_end   = all_months[-1]

    col1, col2 = st.columns(2)
    with col1:
        start_m = st.selectbox("De", all_months, index=all_months.index(default_start))
    with col2:
        end_m = st.selectbox("Até", all_months, index=all_months.index(default_end))

    valid_range = [m for m in all_months if start_m <= m <= end_m]

    all_groups = sorted(df["GRUPO REAL"].dropna().unique())
    sel_groups = st.multiselect("Grupos", all_groups, default=all_groups)

    all_cats = sorted(df["CATEGORIA"].dropna().unique())
    sel_cats = st.multiselect("Categorias", all_cats, default=all_cats)

    status_opts = sorted(df["STATUS"].dropna().unique())
    sel_status = st.multiselect("Status", status_opts, default=status_opts)

# ─── Filtro Principal ──────────────────────────────────────────────────────────
mask = (
    df["AnoMes"].isin(valid_range) &
    df["GRUPO REAL"].isin(sel_groups) &
    df["CATEGORIA"].isin(sel_cats) &
    df["STATUS"].isin(sel_status)
)
fdf = df[mask].copy()

bdf_range = bdf[bdf["AnoMes"].isin(valid_range)].copy()

# ─── Header ───────────────────────────────────────────────────────────────────
st.title("💰 Dashboard Financeiro Pessoal")
st.caption(f"Período selecionado: **{start_m}** → **{end_m}** · {len(valid_range)} meses · {len(fdf):,} transações")

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 1 — KPIs
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📊 Resumo do Período</div>', unsafe_allow_html=True)

total_saida   = fdf["Saída(R$)"].sum()
total_entrada = fdf["Entrada(R$)"].sum()
saldo         = total_entrada - total_saida
pago          = fdf[fdf["STATUS"] == "PAGO"]["Saída(R$)"].sum()
pendente      = fdf[fdf["STATUS"] == "PENDENTE"]["Saída(R$)"].sum()
media_mensal  = fdf.groupby("AnoMes")["Saída(R$)"].sum().mean() if valid_range else 0

# Budget KPIs
total_real     = bdf_range["Entrada Real"].sum()
total_esperado = bdf_range["Entrada Esperada"].sum()
budget_gap     = total_real - total_esperado

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("💸 Total Saídas",    f"R$ {total_saida:,.0f}")
k2.metric("💵 Total Entradas",  f"R$ {total_entrada:,.0f}")
k3.metric("📈 Saldo Líquido",   f"R$ {saldo:,.0f}",    delta=f"R$ {saldo:,.0f}", delta_color="normal")
k4.metric("✅ Pago",            f"R$ {pago:,.0f}")
k5.metric("⏳ Pendente",        f"R$ {pendente:,.0f}", delta_color="inverse")
k6.metric("📅 Média Mensal",    f"R$ {media_mensal:,.0f}")

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 2 — Evolução Mensal
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📅 Evolução Mensal de Despesas</div>', unsafe_allow_html=True)

monthly_grp = (
    fdf.groupby(["AnoMes", "GRUPO REAL", "GRUPO LABEL"])["Saída(R$)"]
    .sum().reset_index()
    .sort_values("AnoMes")
)

fig_evo = px.bar(
    monthly_grp, x="AnoMes", y="Saída(R$)", color="GRUPO LABEL",
    color_discrete_sequence=CAT_PALETTE,
    labels={"AnoMes": "Mês", "Saída(R$)": "R$", "GRUPO LABEL": "Grupo"},
    template="plotly_dark",
)
fig_evo.update_layout(
    plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
    legend=dict(orientation="h", y=-0.2),
    margin=dict(t=20, b=60),
    xaxis=dict(tickangle=-45),
    barmode="stack",
)
st.plotly_chart(fig_evo, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 3 — Por Categoria + Grupos
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🏷️ Distribuição por Categoria e Grupo</div>', unsafe_allow_html=True)

col_a, col_b = st.columns([1.3, 1])

with col_a:
    cat_sum = (
        fdf.groupby("CATEGORIA")["Saída(R$)"]
        .sum().sort_values(ascending=True).reset_index()
    )
    fig_cat = px.bar(
        cat_sum, x="Saída(R$)", y="CATEGORIA", orientation="h",
        color="Saída(R$)", color_continuous_scale="Blues",
        labels={"Saída(R$)": "R$", "CATEGORIA": ""},
        template="plotly_dark",
    )
    fig_cat.update_layout(
        plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
        coloraxis_showscale=False, margin=dict(t=10, b=10, l=10, r=10),
        height=460,
    )
    st.plotly_chart(fig_cat, use_container_width=True)

with col_b:
    grp_sum = fdf.groupby("GRUPO LABEL")["Saída(R$)"].sum().reset_index()
    fig_pie = px.pie(
        grp_sum, values="Saída(R$)", names="GRUPO LABEL",
        color_discrete_sequence=CAT_PALETTE,
        template="plotly_dark", hole=0.45,
    )
    fig_pie.update_traces(textposition="outside", textinfo="percent+label")
    fig_pie.update_layout(
        plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
        showlegend=False, margin=dict(t=10, b=10, l=10, r=10), height=460,
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 4 — Heatmap Mensal por Categoria
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🌡️ Heatmap: Categoria × Mês</div>', unsafe_allow_html=True)

heat = (
    fdf.groupby(["AnoMes", "CATEGORIA"])["Saída(R$)"]
    .sum().unstack(fill_value=0)
    .sort_index()
)
heat = heat[sorted(heat.columns)]

fig_heat = go.Figure(data=go.Heatmap(
    z=heat.values,
    x=heat.columns.tolist(),
    y=heat.index.tolist(),
    colorscale="Blues",
    text=[[f"R$ {v:,.0f}" for v in row] for row in heat.values],
    texttemplate="%{text}",
    textfont={"size": 9},
    hoverongaps=False,
))
fig_heat.update_layout(
    plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
    margin=dict(t=10, b=10), height=420,
    xaxis=dict(tickangle=-45),
    font=dict(color="#e8eaf0"),
)
st.plotly_chart(fig_heat, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 5 — Budget Pessoal
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">🎯 Budget Pessoal — Entradas Real vs Esperado</div>', unsafe_allow_html=True)

bk1, bk2, bk3 = st.columns(3)
bk1.metric("💰 Entrada Real",     f"R$ {total_real:,.0f}")
bk2.metric("🎯 Entrada Esperada", f"R$ {total_esperado:,.0f}")
bk3.metric(
    "📊 Gap Real vs Esperado",
    f"R$ {budget_gap:,.0f}",
    delta=f"R$ {budget_gap:,.0f}",
    delta_color="normal",
)

# Agrupado por mês
bud_monthly = (
    bdf_range.groupby("AnoMes")[["Entrada Real", "Entrada Esperada"]]
    .sum().reset_index().sort_values("AnoMes")
)

fig_bud = go.Figure()
fig_bud.add_trace(go.Bar(
    x=bud_monthly["AnoMes"], y=bud_monthly["Entrada Esperada"],
    name="Esperado", marker_color=COLORS["secondary"], opacity=0.6,
))
fig_bud.add_trace(go.Bar(
    x=bud_monthly["AnoMes"], y=bud_monthly["Entrada Real"],
    name="Real", marker_color=COLORS["success"],
))
fig_bud.update_layout(
    barmode="group",
    plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
    legend=dict(orientation="h", y=-0.2),
    xaxis=dict(tickangle=-45),
    yaxis_title="R$",
    margin=dict(t=10, b=60),
    font=dict(color="#e8eaf0"),
    template="plotly_dark",
)
st.plotly_chart(fig_bud, use_container_width=True)

# Por Título
col_c, col_d = st.columns([1, 1])

with col_c:
    bud_titulo_real = (
        bdf_range.groupby("Título")[["Entrada Real", "Entrada Esperada"]]
        .sum().sort_values("Entrada Real", ascending=False).reset_index()
    )
    fig_titulo = go.Figure()
    fig_titulo.add_trace(go.Bar(
        x=bud_titulo_real["Título"], y=bud_titulo_real["Entrada Esperada"],
        name="Esperado", marker_color=COLORS["secondary"], opacity=0.6,
    ))
    fig_titulo.add_trace(go.Bar(
        x=bud_titulo_real["Título"], y=bud_titulo_real["Entrada Real"],
        name="Real", marker_color=COLORS["success"],
    ))
    fig_titulo.update_layout(
        barmode="group", title="Entradas por Título",
        plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
        xaxis=dict(tickangle=-35),
        legend=dict(orientation="h", y=-0.3),
        margin=dict(t=40, b=80), height=380,
        font=dict(color="#e8eaf0"), template="plotly_dark",
    )
    st.plotly_chart(fig_titulo, use_container_width=True)

with col_d:
    # Gauge: % realizado do budget
    pct = (total_real / total_esperado * 100) if total_esperado > 0 else 0
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=pct,
        number={"suffix": "%", "font": {"color": "#e8eaf0"}},
        delta={"reference": 100, "suffix": "%"},
        title={"text": "% do Budget Realizado", "font": {"color": "#e8eaf0", "size": 16}},
        gauge={
            "axis": {"range": [0, 120], "tickcolor": "#a0aab4"},
            "bar": {"color": COLORS["primary"]},
            "steps": [
                {"range": [0, 70],   "color": "#f05454"},
                {"range": [70, 90],  "color": "#f5a623"},
                {"range": [90, 105], "color": "#4caf7d"},
                {"range": [105, 120],"color": "#7c83ff"},
            ],
            "threshold": {
                "line": {"color": "white", "width": 3},
                "thickness": 0.75, "value": 100,
            },
            "bgcolor": COLORS["bg"],
        },
    ))
    fig_gauge.update_layout(
        paper_bgcolor=COLORS["bg"],
        font=dict(color="#e8eaf0"),
        margin=dict(t=40, b=20, l=20, r=20),
        height=340,
    )
    st.plotly_chart(fig_gauge, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 6 — Despesas vs Receitas ao longo do tempo
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📉 Fluxo: Despesas × Receitas Mensais</div>', unsafe_allow_html=True)

desp_m = fdf.groupby("AnoMes")["Saída(R$)"].sum().reset_index().rename(columns={"Saída(R$)": "Despesas"})
entr_m = fdf.groupby("AnoMes")["Entrada(R$)"].sum().reset_index().rename(columns={"Entrada(R$)": "Receitas"})
flux   = desp_m.merge(entr_m, on="AnoMes", how="outer").sort_values("AnoMes").fillna(0)
flux["Saldo"] = flux["Receitas"] - flux["Despesas"]

fig_flux = go.Figure()
fig_flux.add_trace(go.Scatter(
    x=flux["AnoMes"], y=flux["Despesas"],
    name="Despesas", mode="lines+markers",
    line=dict(color=COLORS["danger"], width=2.5),
))
fig_flux.add_trace(go.Scatter(
    x=flux["AnoMes"], y=flux["Receitas"],
    name="Receitas", mode="lines+markers",
    line=dict(color=COLORS["success"], width=2.5),
))
fig_flux.add_trace(go.Bar(
    x=flux["AnoMes"], y=flux["Saldo"],
    name="Saldo", marker_color=[
        COLORS["success"] if v >= 0 else COLORS["danger"] for v in flux["Saldo"]
    ], opacity=0.4,
))
fig_flux.update_layout(
    plot_bgcolor=COLORS["bg"], paper_bgcolor=COLORS["bg"],
    legend=dict(orientation="h", y=-0.2),
    xaxis=dict(tickangle=-45),
    yaxis_title="R$",
    margin=dict(t=10, b=60),
    font=dict(color="#e8eaf0"),
    template="plotly_dark",
)
st.plotly_chart(fig_flux, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 7 — Tabela de Detalhes
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📋 Detalhamento de Transações</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🔴 Despesas", "🟢 Budget Pessoal"])

with tab1:
    cols_show = ["Data Lançamento", "DESCRIÇÃO", "Saída(R$)", "Entrada(R$)", "CC",
                 "CATEGORIA", "GRUPO LABEL", "STATUS", "AnoMes"]
    show_df = fdf[cols_show].copy()
    show_df["Data Lançamento"] = show_df["Data Lançamento"].dt.strftime("%d/%m/%Y")
    show_df["Saída(R$)"]  = show_df["Saída(R$)"].apply(lambda x: f"R$ {x:,.2f}" if x else "-")
    show_df["Entrada(R$)"]= show_df["Entrada(R$)"].apply(lambda x: f"R$ {x:,.2f}" if x else "-")
    st.dataframe(show_df.rename(columns={"GRUPO LABEL": "Grupo"}), use_container_width=True, height=360)

with tab2:
    bdf_show = bdf_range.copy()
    bdf_show["Data Contábil"]    = bdf_show["Data Contábil"].dt.strftime("%d/%m/%Y")
    bdf_show["Entrada Real"]     = bdf_show["Entrada Real"].apply(lambda x: f"R$ {x:,.2f}" if x else "-")
    bdf_show["Entrada Esperada"] = bdf_show["Entrada Esperada"].apply(lambda x: f"R$ {x:,.2f}" if x else "-")
    st.dataframe(
        bdf_show[["Data Contábil", "AnoMes", "Título", "Entrada Real", "Entrada Esperada"]],
        use_container_width=True, height=360,
    )

# ── Footer ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Dashboard Financeiro Pessoal · Dados: CC_-_PT_v3_1.xlsx")