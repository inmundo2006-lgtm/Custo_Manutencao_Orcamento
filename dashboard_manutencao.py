import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import requests
import io
from datetime import date
import msal

# ─────────────────────────────────────────────
#  CREDENCIAIS — lidas do Streamlit Secrets
#  (configuradas em share.streamlit.io, nunca no código)
# ─────────────────────────────────────────────
CLIENT_ID     = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID     = st.secrets["azure"]["tenant_id"]

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES DO ARQUIVO — altere se necessário
# ─────────────────────────────────────────────
DRIVE_ID    = "b!pmfpHxnsMES0o-2sMMGk49LUqLmtEzNLs4Ii54SlwfwHm6POvL6ySpNF80576tJs"
ITEM_ID     = "01SHXZOAWVYXPFRVNXY5FJV5UQJCVGCC6I"
SHEET_MANUT = "fManutencao"
SHEET_TON   = "COLHEDORAS"
TAXA_TON    = 2.60
FIM_SAFRA   = date(2026, 11, 30)

# ─────────────────────────────────────────────
#  AUTENTICAÇÃO E DOWNLOAD DO SHAREPOINT
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def carregar_dados():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in token:
        st.error(f"Erro na autenticação: {token.get('error_description')}")
        st.stop()

    headers = {"Authorization": f"Bearer {token['access_token']}"}
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
        f"/items/{ITEM_ID}/content"
    )
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        st.error(f"Erro ao baixar arquivo ({resp.status_code})")
        st.stop()

    excel_bytes = io.BytesIO(resp.content)
    df_manut = pd.read_excel(excel_bytes, sheet_name=SHEET_MANUT)
    excel_bytes.seek(0)
    df_ton   = pd.read_excel(excel_bytes, sheet_name=SHEET_TON)
    return df_manut, df_ton


# ─────────────────────────────────────────────
#  PRÉ-PROCESSAMENTO
# ─────────────────────────────────────────────
def preparar_dados(df_manut_raw, df_ton_raw):
    # ── Manutenção
    df_m = df_manut_raw.copy()
    df_m.columns = df_m.columns.str.strip()
    df_m["Data"]         = pd.to_datetime(df_m["Data"], dayfirst=True, errors="coerce")
    df_m["Centro_Custo"] = df_m["Centro_Custo"].astype(str).str.zfill(3)
    df_m["Valor"]        = pd.to_numeric(df_m["Valor"], errors="coerce").fillna(0)

    # ── Toneladas (detecta colunas automaticamente)
    df_t = df_ton_raw.copy()
    df_t.columns = df_t.columns.str.strip()
    col_data = next((c for c in df_t.columns if "data" in c.lower()), df_t.columns[0])
    col_cc   = next((c for c in df_t.columns if "centro" in c.lower() or "cc" in c.lower()), df_t.columns[1])
    col_ton  = next((c for c in df_t.columns if "ton" in c.lower()), df_t.columns[2])

    df_t = df_t.rename(columns={col_data: "Data", col_cc: "Centro_Custo", col_ton: "Toneladas"})
    df_t["Data"]         = pd.to_datetime(df_t["Data"], dayfirst=True, errors="coerce")
    df_t["Centro_Custo"] = df_t["Centro_Custo"].astype(str).str.zfill(3)
    df_t["Toneladas"]    = pd.to_numeric(df_t["Toneladas"], errors="coerce").fillna(0)
    df_t["Orcamento"]    = df_t["Toneladas"] * TAXA_TON

    # ── Filtrar manutenção apenas pelos CCs das colhedoras
    ccs_validos = df_t["Centro_Custo"].unique()
    df_m = df_m[df_m["Centro_Custo"].isin(ccs_validos)].copy()

    return df_m, df_t


# ─────────────────────────────────────────────
#  PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(page_title="Dashboard Manutenção", layout="wide", page_icon="🌾")

st.markdown("""
<style>
    [data-testid="stSidebar"] { background: #0e1e14; }
</style>
""", unsafe_allow_html=True)

st.title("🌾 Orçamento x Manutenção — Colhedoras")
st.caption("Safra 2026/2027 · Dados atualizados a cada 5 minutos")

# ── Carregar
with st.spinner("Carregando dados do SharePoint..."):
    try:
        df_manut_raw, df_ton_raw = carregar_dados()
        df_m, df_t = preparar_dados(df_manut_raw, df_ton_raw)
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        st.stop()

# ─────────────────────────────────────────────
#  SIDEBAR — FILTROS
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("🔎 Filtros")

    cc_opcoes = sorted(df_t["Centro_Custo"].unique())
    cc_sel = st.multiselect(
        "Centro de Custo",
        options=cc_opcoes,
        default=cc_opcoes,
    )

    data_min = df_t["Data"].min().date()
    data_max = df_t["Data"].max().date()
    periodo  = st.date_input(
        "Período",
        value=(data_min, data_max),
        min_value=data_min,
        max_value=date.today(),
    )
    d_ini, d_fim = periodo if len(periodo) == 2 else (data_min, data_max)

    st.divider()
    st.caption(f"Manutenção: {df_m['Data'].min().date()} → {df_m['Data'].max().date()}")

# ─────────────────────────────────────────────
#  APLICAR FILTROS
# ─────────────────────────────────────────────
df_t_f = df_t[
    df_t["Centro_Custo"].isin(cc_sel) &
    (df_t["Data"].dt.date >= d_ini) &
    (df_t["Data"].dt.date <= d_fim)
]
df_m_f = df_m[
    df_m["Centro_Custo"].isin(cc_sel) &
    (df_m["Data"].dt.date >= d_ini) &
    (df_m["Data"].dt.date <= d_fim)
]

# ─────────────────────────────────────────────
#  CÁLCULOS
# ─────────────────────────────────────────────
dias_colhidos  = max((d_fim - d_ini).days, 1)
dias_restantes = max((FIM_SAFRA - d_fim).days, 0)

ton_atual   = df_t_f["Toneladas"].sum()
orc_atual   = df_t_f["Orcamento"].sum()
gasto_atual = df_m_f["Valor"].sum()
saldo_atual = orc_atual - gasto_atual
pct_uso     = (gasto_atual / orc_atual * 100) if orc_atual > 0 else 0

taxa_ton          = ton_atual / dias_colhidos
ton_projetada     = ton_atual + taxa_ton * dias_restantes
orc_projetado     = ton_projetada * TAXA_TON
gasto_projetado   = gasto_atual + (gasto_atual / dias_colhidos) * dias_restantes
saldo_projetado   = orc_projetado - gasto_projetado

# ─────────────────────────────────────────────
#  KPIs — REALIZADO
# ─────────────────────────────────────────────
st.subheader("📊 Realizado até hoje")
k1, k2, k3, k4 = st.columns(4)
k1.metric("🌾 Toneladas Colhidas", f"{ton_atual:,.0f} t")
k2.metric("💰 Orçamento Gerado",   f"R$ {orc_atual:,.2f}")
k3.metric("🔧 Gasto Manutenção",   f"R$ {gasto_atual:,.2f}")
k4.metric("📈 Saldo",              f"R$ {saldo_atual:,.2f}",
          delta=f"{pct_uso:.1f}% utilizado", delta_color="inverse")

st.divider()

# ─────────────────────────────────────────────
#  KPIs — PROJEÇÃO
# ─────────────────────────────────────────────
st.subheader("🔮 Projeção da Safra")
p1, p2, p3, p4 = st.columns(4)
p1.metric("🌾 Ton. Projetadas",     f"{ton_projetada:,.0f} t",
          delta=f"+{ton_projetada - ton_atual:,.0f} restantes")
p2.metric("💰 Orçamento Projetado", f"R$ {orc_projetado:,.2f}")
p3.metric("🔧 Gasto Projetado",     f"R$ {gasto_projetado:,.2f}")
p4.metric("📈 Saldo Projetado",     f"R$ {saldo_projetado:,.2f}",
          delta_color="inverse")

st.divider()

# ─────────────────────────────────────────────
#  GRÁFICOS
# ─────────────────────────────────────────────
col_esq, col_dir = st.columns(2)

# ── Orçamento x Gasto por CC
with col_esq:
    st.subheader("Orçamento x Gasto por CC")

    res_ton  = df_t_f.groupby("Centro_Custo")["Orcamento"].sum().reset_index()
    res_mant = df_m_f.groupby("Centro_Custo")["Valor"].sum().reset_index()
    resumo   = res_ton.merge(res_mant, on="Centro_Custo", how="left").fillna(0)
    resumo.columns = ["CC", "Orçamento", "Gasto"]
    resumo["Saldo"] = resumo["Orçamento"] - resumo["Gasto"]

    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        name="Orçamento", x=resumo["CC"], y=resumo["Orçamento"],
        marker_color="#2ecc71", opacity=0.85
    ))
    fig_bar.add_trace(go.Bar(
        name="Gasto", x=resumo["CC"], y=resumo["Gasto"],
        marker_color="#e74c3c", opacity=0.85
    ))
    fig_bar.update_layout(
        barmode="group", height=350,
        plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=20, b=20)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# ── Evolução acumulada
with col_dir:
    st.subheader("Evolução Acumulada")

    orc_m   = df_t_f.groupby(df_t_f["Data"].dt.to_period("M"))["Orcamento"].sum().cumsum()
    gasto_m = df_m_f.groupby(df_m_f["Data"].dt.to_period("M"))["Valor"].sum().cumsum()
    idx     = orc_m.index.union(gasto_m.index).sort_values()

    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=[str(i) for i in idx], y=orc_m.reindex(idx).ffill().fillna(0).values,
        name="Orçamento Acum.", line=dict(color="#2ecc71", width=2),
        fill="tozeroy", fillcolor="rgba(46,204,113,0.1)"
    ))
    fig_line.add_trace(go.Scatter(
        x=[str(i) for i in idx], y=gasto_m.reindex(idx).ffill().fillna(0).values,
        name="Gasto Acum.", line=dict(color="#e74c3c", width=2)
    ))
    fig_line.update_layout(
        height=350,
        plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=20, b=20)
    )
    st.plotly_chart(fig_line, use_container_width=True)

# ─────────────────────────────────────────────
#  TABELA COM SEMÁFORO
# ─────────────────────────────────────────────
st.subheader("📋 Detalhe por Centro de Custo")

def semaforo(row):
    pct = (row["Gasto"] / row["Orçamento"] * 100) if row["Orçamento"] > 0 else 0
    emoji = "🟢" if pct <= 80 else "🟡" if pct <= 100 else "🔴"
    return f"{emoji} {pct:.1f}%"

resumo["% Uso"]        = resumo.apply(semaforo, axis=1)
resumo["Orçamento R$"] = resumo["Orçamento"].apply(lambda x: f"R$ {x:,.2f}")
resumo["Gasto R$"]     = resumo["Gasto"].apply(lambda x: f"R$ {x:,.2f}")
resumo["Saldo R$"]     = resumo["Saldo"].apply(lambda x: f"R$ {x:,.2f}")

st.dataframe(
    resumo[["CC", "Orçamento R$", "Gasto R$", "Saldo R$", "% Uso"]],
    use_container_width=True,
    hide_index=True
)
