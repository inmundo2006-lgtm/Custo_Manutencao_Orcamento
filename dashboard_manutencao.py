import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import requests
import io
from datetime import date
import msal

# ─────────────────────────────────────────────
#  CREDENCIAIS
# ─────────────────────────────────────────────
CLIENT_ID     = st.secrets["azure"]["client_id"]
CLIENT_SECRET = st.secrets["azure"]["client_secret"]
TENANT_ID     = st.secrets["azure"]["tenant_id"]

DRIVE_ID    = "b!pmfpHxnsMES0o-2sMMGk49LUqLmtEzNLs4Ii54SlwfwHm6POvL6ySpNF80576tJs"
ITEM_ID     = "01SHXZOAWVYXPFRVNXY5FJV5UQJCVGCC6I"
SHEET_MANUT = "fManutencao"
SHEET_TON   = "COLHEDORAS"
SHEET_ORC   = "fOrcamento"

TAXA_TON  = 2.60
FIM_SAFRA = date(2026, 11, 30)
INI_SAFRA = date(2026, 4, 1)

CCS_COLHEITA = ["003", "005", "029", "041", "044", "050", "051"]
CCS_AGRO     = ["028", "037", "038", "046", "047", "049", "052", "054", "056"]

# Descrições dos CCs (código -> descrição completa)
DESC_CC = {
    "003": "003 - VALE DO IVAI",
    "005": "005 - NOVA PRODUTIVA",
    "029": "029 - RIO AMAMBAI",
    "041": "041 - COCAL",
    "044": "044 - SOL NASCENTE",
    "050": "050 - LOBO GUARA",
    "051": "051 - COGO",
    "028": "028 - AGRO VALE DO IVAI",
    "037": "037 - AGRO NAVIRAI",
    "038": "038 - AGRO ASTORGA PLANTIO/CUL",
    "046": "046 - PREPARO DE SOLO",
    "047": "047 - PLANTIO DE GRAOS",
    "049": "049 - AGRO SANTA CANDIDA",
    "052": "052 - PULVERIZACAO AEREA",
    "054": "054 - AGRO RORAIMA",
    "056": "056 - PLANTIO DE GRAOS RORAIMA",
}

USUARIOS = {
    "colheita": ("col2026",  "colheita"),
    "agro":     ("agro2026", "agro"),
}

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def fix_cc(valor):
    try:
        return str(int(float(str(valor).strip()))).zfill(3)
    except:
        return str(valor).strip().zfill(3)

def fmt_brl(valor):
    sinal = "-" if valor < 0 else ""
    return f"{sinal}R$ {abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def semaforo(pct):
    if pct <= 80:  return "🟢"
    if pct <= 100: return "🟡"
    return "🔴"

def fracao_periodo(d_ini, d_fim):
    safra_ini  = date(2026, 4, 1)
    safra_fim  = date(2027, 3, 31)
    total_dias = (safra_fim - safra_ini).days
    dias_sel   = (d_fim - d_ini).days
    return min(dias_sel / total_dias, 1.0) if total_dias > 0 else 0

# ─────────────────────────────────────────────
#  DOWNLOAD SHAREPOINT
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def carregar_dados():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        st.error(f"Erro na autenticação: {token.get('error_description')}")
        st.stop()

    headers = {"Authorization": f"Bearer {token['access_token']}"}
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        st.error(f"Erro ao baixar arquivo ({resp.status_code})")
        st.stop()

    excel_bytes = io.BytesIO(resp.content)
    df_manut = pd.read_excel(excel_bytes, sheet_name=SHEET_MANUT)
    excel_bytes.seek(0)
    df_ton = pd.read_excel(excel_bytes, sheet_name=SHEET_TON)
    excel_bytes.seek(0)
    df_orc = pd.read_excel(excel_bytes, sheet_name=SHEET_ORC)
    return df_manut, df_ton, df_orc

# ─────────────────────────────────────────────
#  PRÉ-PROCESSAMENTO
# ─────────────────────────────────────────────
def preparar_manutencao(df_raw):
    df = df_raw.copy()
    df.columns = df.columns.str.strip()
    df["Data"]         = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["Centro_Custo"] = df["Centro_Custo"].apply(fix_cc)
    df["Valor"]        = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    return df

def preparar_colhedoras(df_raw):
    df = df_raw.copy()
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "DATA":            "Data",
        "CENTRO DE CUSTO": "Centro_Custo",
        "TONELADAS":       "Toneladas",
        "ORÇAMENTO":       "Orcamento",
    })
    df["Data"]         = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["Centro_Custo"] = df["Centro_Custo"].apply(fix_cc)
    df["Toneladas"]    = pd.to_numeric(df["Toneladas"], errors="coerce").fillna(0)
    df["Orcamento"]    = pd.to_numeric(df["Orcamento"], errors="coerce").fillna(0)
    return df

def preparar_orcamento_agro(df_raw):
    df = df_raw.copy()
    df.columns = df.columns.str.strip()
    df["Centro_Custo"]    = df["Centro_Custo"].apply(fix_cc)
    df["Orcamento_Total"] = pd.to_numeric(
        df["Orcamento_Total"].astype(str)
        .str.replace("R$", "").str.replace(".", "").str.replace(",", ".").str.strip(),
        errors="coerce"
    ).fillna(0)
    return df

# ─────────────────────────────────────────────
#  PÁGINA CONFIG
# ─────────────────────────────────────────────
st.set_page_config(page_title="Custo Manutenção", layout="wide", page_icon="🌾")

st.markdown("""
<style>
    [data-testid="stSidebar"] { background: #0e1e14; }
    [data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {
        max-height: none !important;
        height: auto !important;
    }
    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {
        white-space: normal !important;
        overflow: visible !important;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  LOGIN
# ─────────────────────────────────────────────
if "logado" not in st.session_state:
    st.session_state.logado  = False
    st.session_state.modulo  = None
    st.session_state.usuario = None

def fazer_login(usuario, senha):
    u = usuario.strip().lower()
    if u in USUARIOS and USUARIOS[u][0] == senha:
        st.session_state.logado  = True
        st.session_state.modulo  = USUARIOS[u][1]
        st.session_state.usuario = u
        return True
    return False

if not st.session_state.logado:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image("assets/logo.png", width=220)
        st.markdown("""
        <div style='text-align:center; margin-bottom:24px'>
            <span style='color:#2ecc71; font-size:20px; font-weight:bold'>
                Custo x Manutenção
            </span><br>
            <span style='color:#888; font-size:13px'>MS Colheitas e Serviços</span>
        </div>
        """, unsafe_allow_html=True)
        usuario = st.text_input("Usuário", placeholder="colheita / agro")
        senha   = st.text_input("Senha", type="password")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if fazer_login(usuario, senha):
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos")
    st.stop()

# ─────────────────────────────────────────────
#  CARREGAR DADOS
# ─────────────────────────────────────────────
with st.spinner("Carregando dados do SharePoint..."):
    try:
        df_manut_raw, df_ton_raw, df_orc_raw = carregar_dados()
        df_m   = preparar_manutencao(df_manut_raw)
        df_t   = preparar_colhedoras(df_ton_raw)
        df_orc = preparar_orcamento_agro(df_orc_raw)
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        st.stop()

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
modulo = st.session_state.modulo

with st.sidebar:
    st.markdown(f"**👤 {st.session_state.usuario.upper()}**")
    st.divider()
    st.header("🔎 Filtros")

    if modulo == "colheita":
        cc_opcoes = CCS_COLHEITA
    else:
        cc_opcoes = CCS_AGRO

    cc_sel_desc = st.multiselect(
        "Centro de Custo",
        options=cc_opcoes,
        default=cc_opcoes,
        format_func=lambda c: DESC_CC.get(c, c)
    )
    cc_sel = cc_sel_desc

    periodo = st.date_input(
        "Período",
        value=(INI_SAFRA, date.today()),
        min_value=INI_SAFRA,
        max_value=date.today(),
        format="DD/MM/YYYY",
    )
    d_ini, d_fim = periodo if len(periodo) == 2 else (INI_SAFRA, date.today())

    st.divider()
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state.logado  = False
        st.session_state.modulo  = None
        st.session_state.usuario = None
        st.rerun()
    st.caption(f"{d_ini.strftime('%d/%m/%Y')} → {d_fim.strftime('%d/%m/%Y')}")

# ═══════════════════════════════════════════════
#  MÓDULO COLHEITA
# ═══════════════════════════════════════════════
if modulo == "colheita":

    st.title("🌾 Orçamento x Manutenção")
    st.caption("Safra 2026/2027 · Dados atualizados a cada 5 minutos")

    df_t_f = df_t[
        df_t["Centro_Custo"].isin(cc_sel) &
        (df_t["Data"].dt.date >= d_ini) &
        (df_t["Data"].dt.date <= d_fim)
    ]
    df_m_f = df_m[
        df_m["Centro_Custo"].isin(CCS_COLHEITA) &
        df_m["Centro_Custo"].isin(cc_sel) &
        (df_m["Data"].dt.date >= d_ini) &
        (df_m["Data"].dt.date <= d_fim)
    ]

    dias_colhidos  = max((d_fim - d_ini).days, 1)
    dias_restantes = max((FIM_SAFRA - d_fim).days, 0)
    ton_atual      = df_t_f["Toneladas"].sum()
    orc_atual      = df_t_f["Orcamento"].sum()
    gasto_atual    = df_m_f["Valor"].sum()
    saldo_atual    = orc_atual - gasto_atual
    pct_uso        = (gasto_atual / orc_atual * 100) if orc_atual > 0 else 0
    taxa_ton       = ton_atual / dias_colhidos
    ton_proj       = ton_atual + taxa_ton * dias_restantes
    orc_proj       = ton_proj * TAXA_TON
    gasto_proj     = gasto_atual + (gasto_atual / dias_colhidos) * dias_restantes
    saldo_proj     = orc_proj - gasto_proj

    # KPIs realizado
    st.subheader("📊 Realizado até hoje")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🌾 Toneladas Colhidas", f"{ton_atual:,.0f} t".replace(",", "."))
    k2.metric("💰 Orçamento Gerado",   fmt_brl(orc_atual))
    k3.metric("🔧 Gasto Manutenção",   fmt_brl(gasto_atual))
    k4.metric("📈 Saldo",              fmt_brl(saldo_atual),
              delta=f"{pct_uso:.1f}% utilizado", delta_color="inverse")
    st.divider()

    # KPIs projeção
    st.subheader("🔮 Projeção da Safra")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("🌾 Ton. Projetadas",     f"{ton_proj:,.0f} t".replace(",", "."))
    p2.metric("💰 Orçamento Projetado", fmt_brl(orc_proj))
    p3.metric("🔧 Gasto Projetado",     fmt_brl(gasto_proj))
    p4.metric("📈 Saldo Projetado",     fmt_brl(saldo_proj), delta_color="inverse")
    st.divider()

    # Gráficos
    col_esq, col_dir = st.columns(2)

    with col_esq:
        st.subheader("Orçamento x Gasto por CC")
        res = df_t_f.groupby("Centro_Custo")["Orcamento"].sum().reset_index()
        res = res.merge(
            df_m_f.groupby("Centro_Custo")["Valor"].sum().reset_index(),
            on="Centro_Custo", how="left"
        ).fillna(0)
        res.columns = ["CC", "Orçamento", "Gasto"]
        res["Saldo"] = res["Orçamento"] - res["Gasto"]

        fig = go.Figure()
        fig.add_trace(go.Bar(name="Orçamento", x=res["CC"], y=res["Orçamento"], marker_color="#2ecc71", opacity=0.85))
        fig.add_trace(go.Bar(name="Gasto",     x=res["CC"], y=res["Gasto"],     marker_color="#e74c3c", opacity=0.85))
        fig.update_layout(barmode="group", height=350,
            plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02), margin=dict(t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    with col_dir:
        st.subheader("Evolução Acumulada")
        orc_m   = df_t_f.groupby(df_t_f["Data"].dt.to_period("M"))["Orcamento"].sum().cumsum()
        gasto_m = df_m_f.groupby(df_m_f["Data"].dt.to_period("M"))["Valor"].sum().cumsum()
        idx     = orc_m.index.union(gasto_m.index).sort_values()

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=[str(i) for i in idx], y=orc_m.reindex(idx).ffill().fillna(0).values,
            name="Orçamento Acum.", line=dict(color="#2ecc71", width=2),
            fill="tozeroy", fillcolor="rgba(46,204,113,0.1)"
        ))
        fig2.add_trace(go.Scatter(
            x=[str(i) for i in idx], y=gasto_m.reindex(idx).ffill().fillna(0).values,
            name="Gasto Acum.", line=dict(color="#e74c3c", width=2)
        ))
        fig2.update_layout(height=350,
            plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02), margin=dict(t=20, b=20))
        st.plotly_chart(fig2, use_container_width=True)

    # Tabela
    st.subheader("📋 Detalhe por Centro de Custo")
    res["% Uso"]        = res.apply(lambda r: f"{semaforo(r['Gasto']/r['Orçamento']*100 if r['Orçamento']>0 else 0)} {r['Gasto']/r['Orçamento']*100 if r['Orçamento']>0 else 0:.1f}%", axis=1)
    res["Orçamento R$"] = res["Orçamento"].apply(fmt_brl)
    res["Gasto R$"]     = res["Gasto"].apply(fmt_brl)
    res["Saldo R$"]     = res["Saldo"].apply(fmt_brl)
    st.dataframe(res[["CC", "Orçamento R$", "Gasto R$", "Saldo R$", "% Uso"]],
                 use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════
#  MÓDULO AGRO
# ═══════════════════════════════════════════════
elif modulo == "agro":

    st.title("🚜 Orçamento x Manutenção — Agropecuárias")
    st.caption("Safra 2026/2027 · Dados atualizados a cada 5 minutos")

    frac           = fracao_periodo(d_ini, d_fim)
    df_orc_sel     = df_orc[df_orc["Centro_Custo"].isin(cc_sel)].copy()
    df_orc_sel["Orcamento_Periodo"] = df_orc_sel["Orcamento_Total"] * frac

    df_m_f = df_m[
        df_m["Centro_Custo"].isin(CCS_AGRO) &
        df_m["Centro_Custo"].isin(cc_sel) &
        (df_m["Data"].dt.date >= d_ini) &
        (df_m["Data"].dt.date <= d_fim)
    ]

    orc_periodo = df_orc_sel["Orcamento_Periodo"].sum()
    orc_safra   = df_orc_sel["Orcamento_Total"].sum()
    gasto_atual = df_m_f["Valor"].sum()
    saldo_atual = orc_periodo - gasto_atual
    pct_uso     = (gasto_atual / orc_periodo * 100) if orc_periodo > 0 else 0

    # KPIs período
    st.subheader("📊 Realizado no período")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📅 Fração da Safra",   f"{frac*100:.1f}%")
    k2.metric("💰 Orçamento Período", fmt_brl(orc_periodo))
    k3.metric("🔧 Gasto Manutenção",  fmt_brl(gasto_atual))
    k4.metric("📈 Saldo",             fmt_brl(saldo_atual),
              delta=f"{pct_uso:.1f}% utilizado", delta_color="inverse")
    st.divider()

    # KPIs safra total
    st.subheader("📋 Orçamento Total da Safra")
    s1, s2, s3 = st.columns(3)
    s1.metric("💰 Orçamento Safra Inteira", fmt_brl(orc_safra))
    s2.metric("🔧 Gasto Acumulado",         fmt_brl(gasto_atual))
    s3.metric("📈 Saldo da Safra",          fmt_brl(orc_safra - gasto_atual))
    st.divider()

    # Gráficos
    col_esq, col_dir = st.columns(2)

    with col_esq:
        st.subheader("Orçamento x Gasto por CC")
        res = df_orc_sel[["Centro_Custo", "Orcamento_Periodo"]].copy()
        res.columns = ["CC", "Orçamento"]
        res = res.merge(
            df_m_f.groupby("Centro_Custo")["Valor"].sum().reset_index().rename(
                columns={"Centro_Custo": "CC", "Valor": "Gasto"}),
            on="CC", how="left"
        ).fillna(0)
        res["Saldo"] = res["Orçamento"] - res["Gasto"]

        fig = go.Figure()
        fig.add_trace(go.Bar(name="Orçamento Período", x=res["CC"], y=res["Orçamento"], marker_color="#3498db", opacity=0.85))
        fig.add_trace(go.Bar(name="Gasto",             x=res["CC"], y=res["Gasto"],     marker_color="#e74c3c", opacity=0.85))
        fig.update_layout(barmode="group", height=350,
            plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02), margin=dict(t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    with col_dir:
        st.subheader("Evolução Acumulada")
        gasto_m = df_m_f.groupby(df_m_f["Data"].dt.to_period("M"))["Valor"].sum().cumsum()

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=[str(i) for i in gasto_m.index], y=gasto_m.values,
            name="Gasto Acum.", line=dict(color="#e74c3c", width=2),
            fill="tozeroy", fillcolor="rgba(231,76,60,0.1)"
        ))
        fig2.add_hline(y=orc_periodo, line_dash="dash", line_color="#3498db",
                       annotation_text="Orçamento Período", annotation_position="top left")
        fig2.update_layout(height=350,
            plot_bgcolor="#0c1711", paper_bgcolor="#0c1711", font_color="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02), margin=dict(t=20, b=20))
        st.plotly_chart(fig2, use_container_width=True)

    # Tabela
    st.subheader("📋 Detalhe por Centro de Custo")
    res2 = res.merge(
        df_orc_sel[["Centro_Custo", "Orcamento_Total", "Descricao"]].rename(
            columns={"Centro_Custo": "CC"}),
        on="CC", how="left"
    )
    res2["% Uso"]             = res2.apply(lambda r: f"{semaforo(r['Gasto']/r['Orçamento']*100 if r['Orçamento']>0 else 0)} {r['Gasto']/r['Orçamento']*100 if r['Orçamento']>0 else 0:.1f}%", axis=1)
    res2["Orçamento Safra"]   = res2["Orcamento_Total"].apply(fmt_brl)
    res2["Orçamento Período"] = res2["Orçamento"].apply(fmt_brl)
    res2["Gasto R$"]          = res2["Gasto"].apply(fmt_brl)
    res2["Saldo R$"]          = res2["Saldo"].apply(fmt_brl)
    st.dataframe(
        res2[["CC", "Descricao", "Orçamento Safra", "Orçamento Período", "Gasto R$", "Saldo R$", "% Uso"]],
        use_container_width=True, hide_index=True
    )