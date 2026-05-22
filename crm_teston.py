import streamlit as st
import pandas as pd
import requests
import msal
import os
import base64
import folium
from folium.plugins import Fullscreen, MiniMap
from streamlit_folium import st_folium
import plotly.express as px
from datetime import date, datetime

# ============================================================
# 0. LOGIN
# ============================================================
USUARIOS = {
    "fernando":  {"senha": "teston@2026", "perfil": "vendedor",  "nome": "Fernando"},
    "wanderson": {"senha": "teston@2026", "perfil": "vendedor",  "nome": "Wanderson"},
    "admin":     {"senha": "adm@teston",  "perfil": "admin",     "nome": "Administrador"},
}

def img_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def tela_login():
    st.set_page_config(
        page_title="CRM Teston · Login",
        layout="centered",
        page_icon="\U0001f3ed"
    )

    logo_path = "teston_logo.png"
    logo_html = ""
    if os.path.exists(logo_path):
        b64 = img_base64(logo_path)
        logo_html = f'<img src="data:image/png;base64,{b64}" style="width:220px;margin-bottom:8px;">'

    st.markdown(f"""
    <style>
    html, body, [data-testid="stAppViewContainer"] {{
        background: #0c1711 !important;
    }}
    [data-testid="stAppViewContainer"] > .main {{
        background: #0c1711;
    }}
    .login-box {{
        background: #111d14;
        border: 1px solid #00FF8833;
        border-radius: 16px;
        padding: 40px 44px 36px 44px;
        max-width: 420px;
        margin: 60px auto 0 auto;
        box-shadow: 0 0 40px #00FF8811;
        text-align: center;
    }}
    .login-box h2 {{
        color: #00FF88;
        font-size: 22px;
        font-weight: 800;
        margin: 16px 0 4px 0;
        letter-spacing: 1px;
    }}
    .login-box p {{
        color: #4a7a5a;
        font-size: 13px;
        margin-bottom: 28px;
    }}
    div[data-testid="stTextInput"] input {{
        background: #0c1711 !important;
        border: 1px solid #00FF8844 !important;
        border-radius: 8px !important;
        color: #e0ffe0 !important;
        font-size: 15px !important;
    }}
    div[data-testid="stTextInput"] input:focus {{
        border-color: #00FF88 !important;
        box-shadow: 0 0 0 2px #00FF8822 !important;
    }}
    div[data-testid="stTextInput"] label {{
        color: #7abf8a !important;
        font-size: 13px !important;
        font-weight: 600 !important;
    }}
    div[data-testid="stForm"] {{
        background: transparent !important;
        border: none !important;
    }}
    </style>
    <div class="login-box">
        {logo_html}
        <h2>INTELIGÊNCIA COMERCIAL</h2>
        <p>Sistema de Gestão de Relacionamento</p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        usuario = st.text_input("\U0001f464 Usuário").strip().lower()
        senha   = st.text_input("\U0001f512 Senha", type="password")
        entrar  = st.form_submit_button("Entrar →", use_container_width=True, type="primary")

    if entrar:
        if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
            st.session_state.logado       = True
            st.session_state.usuario      = usuario
            st.session_state.perfil       = USUARIOS[usuario]["perfil"]
            st.session_state.nome_usuario = USUARIOS[usuario]["nome"]
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")

if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    tela_login()
    st.stop()


# ============================================================
# 1. SHAREPOINT / GRAPH API
# ============================================================
SCOPE = ["https://graph.microsoft.com/.default"]

@st.cache_data(ttl=3500)
def _get_token() -> str:
    cfg = st.secrets["sharepoint"]
    app = msal.ConfidentialClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        client_credential=cfg["client_secret"],
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in result:
        erro = result.get("error_description") or result.get("error") or str(result)
        st.error(f"🔐 Falha MSAL: {erro}")
        st.stop()
    return result["access_token"]

def _headers() -> dict:
    return {"Authorization": f"Bearer {_get_token()}", "Content-Type": "application/json"}

def _base_url() -> str:
    cfg = st.secrets["sharepoint"]
    return (
        f"https://graph.microsoft.com/v1.0"
        f"/sites/{cfg['site_id']}/lists/{cfg['list_id']}/items"
    )

# ── Campos SharePoint ↔ colunas SQLite originais ─────────────
# Todos os campos ficam em item["fields"]. O campo "Title" é obrigatório
# no SP; usamos ele como "usina". O id numérico do SP substitui o rowid.
SP_FIELDS = [
    "id", "Title", "cidade_estado", "marca_equipamento", "vendedor",
    "moagem_anual", "frota_atual", "e_cliente",
    "grau_relacionamento", "ultima_visita", "proxima_visita",
    "participou_evento", "testou_maquina", "maquinas_testadas",
    "nome_contato", "cargo_contato", "telefone_contato",
    "nome_contato2", "cargo_contato2", "telefone_contato2",
    "pretende_investir", "quantidade_prevista",
    "tipo_equipamento", "concorrentes_cotados", "dor_operacional",
    "maquinas_instaladas", "historico_problemas", "grau_satisfacao",
    "quem_decide", "quem_bloqueia", "fluxo_aprovacao",
    "janela_compra", "relacao_concorrentes", "situacao_financeira",
    "data_registro",
]

def _item_to_row(item: dict) -> dict:
    """Converte item Graph → dict com mesmas chaves do SQLite antigo."""
    f = item.get("fields", {})
    row = {k: f.get(k, "") for k in SP_FIELDS}
    row["id"] = item.get("id", f.get("id", ""))
    row["usina"] = f.get("Title", "")
    row["regiao"] = f.get("marca_equipamento", "")  # alias
    row["data_registro"] = f.get("data_registro") or item.get("createdDateTime", "")
    return row

def fmt_data(valor, hora=False):
    if not valor or str(valor).strip() in ("", "None", "nan", "NaT"):
        return "—"
    try:
        dt = datetime.strptime(str(valor).strip()[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        try:
            dt = datetime.strptime(str(valor).strip()[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            try:
                dt = datetime.strptime(str(valor).strip()[:10], "%Y-%m-%d")
            except Exception:
                return str(valor)[:10]
    return dt.strftime("%d/%m/%Y %H:%M" if hora else "%d/%m/%Y")


def fmt_telefone(valor: str) -> str:
    """Aplica máscara (xx) xxxxx-xxxx enquanto o usuário digita."""
    digits = ''.join(c for c in str(valor) if c.isdigit())[:11]
    if len(digits) == 0:
        return ""
    elif len(digits) <= 2:
        return f"({digits}"
    elif len(digits) <= 7:
        return f"({digits[:2]}) {digits[2:]}"
    elif len(digits) <= 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    return valor

# ── CRUD ─────────────────────────────────────────────────────

def salvar_crm(**kwargs):
    """Cria novo item na SharePoint List."""
    payload = {k: v for k, v in kwargs.items()}
    payload["Title"] = kwargs.get("usina", "")          # SP exige Title
    payload.pop("usina", None)
    payload["data_registro"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    r = requests.post(_base_url(), headers=_headers(), json={"fields": payload})
    r.raise_for_status()

def atualizar_crm(record_id, **kwargs):
    """Atualiza item existente (PATCH nos fields)."""
    payload = {k: v for k, v in kwargs.items()}
    payload["Title"] = kwargs.get("usina", "")
    payload.pop("usina", None)
    url = f"{_base_url()}/{record_id}/fields"
    r = requests.patch(url, headers=_headers(), json=payload)
    r.raise_for_status()

def deletar_registro(record_id):
    """Exclui item da lista."""
    url = f"{_base_url()}/{record_id}"
    r = requests.delete(url, headers=_headers())
    r.raise_for_status()

@st.cache_data(ttl=30)
def carregar_registros() -> pd.DataFrame:
    """Carrega todos os itens da lista (com paginação)."""
    select = ",".join([f for f in SP_FIELDS if f != "id"])
    url = f"{_base_url()}?expand=fields(select={select})&$top=999"
    rows = []
    while url:
        r = requests.get(url, headers=_headers())
        r.raise_for_status()
        data = r.json()
        rows.extend([_item_to_row(i) for i in data.get("value", [])])
        url = data.get("@odata.nextLink")
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=SP_FIELDS + ["usina"])

@st.cache_data(ttl=30)
def ultima_visita_por_usina() -> dict:
    df = carregar_registros()
    if df.empty:
        return {}
    df["data_registro"] = pd.to_datetime(df["data_registro"], errors="coerce")
    idx = df.groupby("usina")["data_registro"].idxmax()
    sub = df.loc[idx, ["usina","data_registro","grau_relacionamento","proxima_visita","pretende_investir","id"]]
    result = {}
    for _, row in sub.iterrows():
        result[row["usina"]] = {
            "ultima":               row["data_registro"],
            "grau_relacionamento":  row["grau_relacionamento"],
            "proxima_visita":       row["proxima_visita"],
            "pretende_investir":    row["pretende_investir"],
            "id":                   row["id"],
        }
    return result

@st.cache_data(ttl=30)
def historico_usina(nome_usina: str) -> pd.DataFrame:
    df = carregar_registros()
    if df.empty:
        return pd.DataFrame()
    df = df[df["usina"] == nome_usina].copy()
    df["data_registro"] = pd.to_datetime(df["data_registro"], errors="coerce")
    df = df.sort_values("data_registro", ascending=False).head(10)
    return df[["id","data_registro","ultima_visita","grau_relacionamento","pretende_investir","nome_contato","situacao_financeira"]]

@st.cache_data(ttl=30)
def buscar_registro_por_id(record_id) -> dict:
    url = f"{_base_url()}/{record_id}?expand=fields"
    r = requests.get(url, headers=_headers())
    r.raise_for_status()
    return _item_to_row(r.json())

# ============================================================
# 2. USINAS
# ============================================================
@st.cache_data
def carregar_usinas():
    nome_arquivo = "USINAS_GEOCODIFICADAS.xlsx"
    if os.path.exists(nome_arquivo):
        try:
            return pd.read_excel(nome_arquivo, sheet_name="USINAS_COORDS")
        except ValueError:
            return pd.read_excel(nome_arquivo)
    else:
        st.error(f"Arquivo '{nome_arquivo}' não encontrado.")
        return pd.DataFrame()

df_usinas = carregar_usinas()

def detectar_col(df, candidatos):
    for c in candidatos:
        if c in df.columns:
            return c
    return None

COL_LAT = detectar_col(df_usinas, ["LAT","LATITUDE","Latitude","lat","latitude"])
COL_LON = detectar_col(df_usinas, ["LON","LONG","LONGITUDE","Longitude","lon","longitude","lng"])

# ============================================================
# 3. SESSION STATE
# ============================================================
def _ss(k, v):
    if k not in st.session_state:
        st.session_state[k] = v

_ss("modo", "grid")
_ss("usina_row", None)
_ss("registro_id", None)

# ============================================================
# 4. PÁGINA E ESTILOS
# ============================================================
st.set_page_config(
    page_title="CRM Teston", layout="wide",
    initial_sidebar_state=("collapsed" if st.session_state.modo == "formulario" else "expanded"),
    page_icon="🏭"
)

st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background: rgba(0,255,136,0.06);
        border: 1px solid rgba(0,255,136,0.22);
        border-radius: 10px; padding: 12px 14px;
    }
    div[data-testid="stMetricValue"] { color: #00FF88 !important; }
    button[data-baseweb="tab"] { font-size: 15px !important; font-weight: 600 !important; }
    .badge-ok   { background:#00FF88; color:#0e1117; font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; display:inline-block; }
    .badge-none { background:#444; color:#aaa; font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; display:inline-block; }
    .form-header { background: rgba(0,255,136,0.08); border-left: 4px solid #00FF88; padding: 16px 20px; border-radius: 0 8px 8px 0; margin-bottom: 20px; }
    .status-tag  { display:inline-block; font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; letter-spacing:0.4px; }
    .tag-visitada   { background:#00FF8820; color:#00FF88; border:1px solid #00FF8860; }
    .tag-programada { background:#FFD70020; color:#FFD700; border:1px solid #FFD70060; }
    .tag-pendente   { background:#FF4B4B20; color:#FF4B4B; border:1px solid #FF4B4B60; }
    .card-dates { font-size:12px; color:#aaa; margin:4px 0 10px 0; display:flex; gap:16px; flex-wrap:wrap; }
    .card-dates b { color:#ddd; }
    div[data-testid="stVerticalBlockBorderWrapper"] { min-height:210px; display:flex; flex-direction:column; justify-content:space-between; transition:border-color 0.2s; }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.card-visitada)   { border-left:5px solid #00FF88 !important; }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.card-programada) { border-left:5px solid #FFD700 !important; }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.card-pendente)   { border-left:5px solid #FF4B4B !important; }
</style>
""", unsafe_allow_html=True)

GRAU_EMOJI = {"Quente": "🔴", "Morno": "🟡", "Frio": "🔵"}
GRAU_COLOR = {"Quente": "#FF4B4B", "Morno": "#FFD700", "Frio": "#4B9CFF"}
INV_EMOJI  = {"Sim": "✅", "Não": "❌", "Indefinido": "❓"}

# ============================================================
# 5. HELPERS DE NAVEGAÇÃO
# ============================================================
def abrir_formulario(row_dict, registro_id=None):
    st.session_state.modo = "formulario"
    st.session_state.usina_row = row_dict
    st.session_state.registro_id = registro_id
    st.rerun()

def voltar_grid():
    st.session_state.modo = "grid"
    st.session_state.usina_row = None
    st.session_state.registro_id = None
    st.cache_data.clear()
    st.rerun()

# ============================================================
# 6. SIDEBAR
# ============================================================
if st.session_state.modo == "grid":
    with st.sidebar:
        st.markdown("## 🏭 Teston CRM")
        st.divider()
        st.markdown("### 🔍 Filtros")
        if not df_usinas.empty:
            lista_vendedores = ["Todos"] + sorted(df_usinas["VENDEDOR"].dropna().unique().tolist())
            vendedor_sel = st.selectbox("Vendedor", lista_vendedores)
            if "REGIAO" in df_usinas.columns:
                lista_regioes = ["Todas"] + sorted(df_usinas["REGIAO"].dropna().unique().tolist())
                regiao_sel = st.selectbox("Região", lista_regioes)
            else:
                regiao_sel = "Todas"
            if "ESTADO" in df_usinas.columns:
                lista_estados = ["Todos"] + sorted(df_usinas["ESTADO"].dropna().unique().tolist())
                estado_sel = st.selectbox("Estado", lista_estados)
            else:
                estado_sel = "Todos"
            busca_nome = st.text_input("🔎 Buscar usina", "")
            st.divider()
            apenas_sem_visita = st.toggle("🚨 Apenas sem visita registrada", value=False)
        else:
            vendedor_sel = busca_nome = "Todos"
            regiao_sel = estado_sel = "Todas"
            apenas_sem_visita = False
        st.divider()
        st.caption("Inteligência Comercial Teston · v5.0 · SP")

# ============================================================
# MODO FORMULÁRIO FULL-SCREEN
# ============================================================
if st.session_state.modo == "formulario":
    row    = st.session_state.usina_row
    reg_id = st.session_state.registro_id
    dados  = buscar_registro_por_id(reg_id) if reg_id else {}
    editando = bool(reg_id and dados)

    def v(campo, default=""):
        return dados.get(campo, default) or default

    def vi(campo, opcoes, default=0):
        val = dados.get(campo, "")
        return opcoes.index(val) if val in opcoes else default

    btn_col, titulo_col = st.columns([1, 8])
    with btn_col:
        if st.button("← Voltar", type="secondary"):
            voltar_grid()
    with titulo_col:
        modo_label = "✏️ Editando Registro" if editando else "📝 Nova Visita"
        st.markdown(
            f'<div class="form-header">'
            f'<span style="font-size:13px;color:#00FF88;font-weight:700;">{modo_label}</span><br>'
            f'<span style="font-size:22px;font-weight:800;">{row["USINA"]}</span> &nbsp;'
            f'<span style="font-size:15px;color:#888;">{row.get("CIDADE","")} – {row.get("ESTADO","")} | 👤 {row["VENDEDOR"]}</span>'
            f'</div>', unsafe_allow_html=True)

    if editando:
        st.info(f"📌 Editando registro **#{reg_id}** salvo em {fmt_data(dados.get('data_registro',''), hora=True)}")

    st.divider()

    with st.form("form_fullscreen", border=False):
        st.markdown("#### 🏭 Dados da Usina")
        c1, c2 = st.columns(2)
        moagem   = c1.text_input("Moagem anual (ton)", value=v("moagem_anual"))
        OPT_MARCA = ["","TESTON","TMA","MEGATEC","SANTA ISABEL","ANTONIOSI","GRUNNER","BMB"]
        marca_salva = v("marca_equipamento", "")
        marca_idx = OPT_MARCA.index(marca_salva) if marca_salva in OPT_MARCA else 0
        marca_eq = c2.selectbox("Marca do Equipamento", OPT_MARCA, index=marca_idx)

        st.markdown("##### 🚜 Frota Atual por Marca")
        st.caption("Informe a quantidade de máquinas por marca (0 = não possui)")

        MARCAS_FROTA = ["TESTON","TMA","MEGATEC","SANTA ISABEL","ANTONIOSI","GRUNNER","BMB"]

        def parse_frota_atual(txt):
            qtds = {}
            if txt:
                for parte in str(txt).split("|"):
                    parte = parte.strip()
                    if ":" in parte:
                        marca, q = parte.split(":", 1)
                        try:
                            qtds[marca.strip()] = int(q.strip())
                        except ValueError:
                            qtds[marca.strip()] = 0
            return qtds

        frota_atual_salva = parse_frota_atual(v("frota_atual"))
        qtds_frota = {}
        for marca in MARCAS_FROTA:
            col_marca, col_qtd = st.columns([5, 2])
            col_marca.markdown(
                f'<div style="padding:6px 0;color:#cccccc;font-weight:600;">{marca}</div>',
                unsafe_allow_html=True)
            qtds_frota[marca] = col_qtd.number_input(
                "Qtd", min_value=0,
                value=frota_atual_salva.get(marca, 0), step=1,
                key=f"frota_atual_{marca.replace(' ','_')}"
            )

        frota = " | ".join(
            f"{m}: {q}" for m, q in qtds_frota.items() if q > 0
        ) or ""
        st.markdown("---")

        t1, t2, t3, t4 = st.tabs([
            "🤝 Relacionamento & Contato",
            "💰 Potencial de Venda",
            "🔧 Pós-Venda",
            "🧠 Inteligência Comercial"
        ])

        with t1:
            st.markdown("##### Classificação")
            OPT_CLI  = ["Sim","Não","Inativo"]
            OPT_GRAU = ["🔵 Frio","🟡 Morno","🔴 Quente"]
            grau_salvo = v("grau_relacionamento")
            grau_idx   = next((i for i,o in enumerate(OPT_GRAU) if grau_salvo in o), 0)
            c1,c2,c3,c4 = st.columns(4)
            cliente       = c1.selectbox("Já é cliente?",           OPT_CLI,  index=vi("e_cliente", OPT_CLI))
            grau          = c2.selectbox("Temperatura",              OPT_GRAU, index=grau_idx)
            participou_ev = c3.selectbox("Participou de algum evento?", ["Não","Sim"], index=vi("participou_evento",["Não","Sim"]))
            # c4 reservado — testou_maquina derivado do multiselect abaixo
            c4.markdown("")  # espaço visual

            # Multiselect de máquinas testadas (vazio = não testou)
            OPT_MAQUINAS = ["RAPTOR 13.000","GIGANTE M 22.000","TRACTOR 22.000","GIGANTE M 30.000"]
            maq_salvas = [m.strip() for m in v("maquinas_testadas","").split(",") if m.strip() in OPT_MAQUINAS]
            maquinas_testadas = st.multiselect(
                "🚜 Testou alguma máquina? (selecione os modelos testados ou deixe vazio)",
                OPT_MAQUINAS, default=maq_salvas
            )
            testou_maq = "Sim" if maquinas_testadas else "Não"

            st.markdown("##### Dados da Visita")
            def parse_date(s):
                try:    return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
                except: return date.today()
            c1,c2 = st.columns(2)
            ult_visita  = c1.date_input("Última visita",  value=parse_date(v("ultima_visita")), format="DD/MM/YYYY")
            prox_visita = c2.date_input("Próxima visita", value=parse_date(v("proxima_visita")), format="DD/MM/YYYY")
            resultado_visita = st.text_area(
                "📝 O que foi tratado nesta visita",
                value=v("resultado_teste"),
                height=140,
                placeholder="Descreva os assuntos abordados, andamento da negociação, próximos passos..."
            )

            st.markdown("##### Contato 1")
            OPT_CARGO = ["","GERENTE AGRICOLA","GERENTE AUTOMOTIVO","DIRETOR",
                         "COORDENADOR AGRICOLA","COORDENADOR AUTOMOTIVO","SUPRIMENTOS"]
            c1,c2,c3 = st.columns(3)
            nome_ctt  = c1.text_input("Nome do contato", value=v("nome_contato"))
            cargo_idx = OPT_CARGO.index(v("cargo_contato")) if v("cargo_contato") in OPT_CARGO else 0
            cargo_ctt = c2.selectbox("Cargo / função", OPT_CARGO, index=cargo_idx)
            tel_ctt   = c3.text_input("Telefone / WhatsApp",
                                       value=fmt_telefone(v("telefone_contato")),
                                       placeholder="(xx) xxxxx-xxxx",
                                       help="Digite apenas os números — a máscara é aplicada ao salvar")

            st.markdown("##### Contato 2")
            c1,c2,c3 = st.columns(3)
            nome_ctt2  = c1.text_input("Nome do contato 2", value=v("nome_contato2"))
            cargo_idx2 = OPT_CARGO.index(v("cargo_contato2")) if v("cargo_contato2") in OPT_CARGO else 0
            cargo_ctt2 = c2.selectbox("Cargo / função 2", OPT_CARGO, index=cargo_idx2)
            tel_ctt2   = c3.text_input("Telefone / WhatsApp 2",
                                        value=fmt_telefone(v("telefone_contato2")),
                                        placeholder="(xx) xxxxx-xxxx",
                                        help="Digite apenas os números — a máscara é aplicada ao salvar")

        with t2:
            OPT_INV = ["Indefinido","Sim","Não"]
            investe = st.selectbox("Pretende investir?", OPT_INV, index=vi("pretende_investir", OPT_INV))

            st.markdown("##### 🚜 Equipamentos de Interesse e Quantidade Prevista")
            st.caption("Informe a quantidade prevista para cada modelo de interesse (0 = não tem interesse)")

            MODELOS_VENDA = {
                "RAPTOR 13.000":    None,
                "GIGANTE M 22.000": None,
                "TRACTOR 22.000":   None,
                "GIGANTE M 30.000": None,
            }

            # Parser do campo tipo_equipamento: "RAPTOR 13.000: 3 | GIGANTE M 22.000: 2"
            def parse_tipo_eq(txt):
                qtds = {}
                if txt:
                    for parte in str(txt).split("|"):
                        parte = parte.strip()
                        if ":" in parte:
                            modelo, q = parte.split(":", 1)
                            try:
                                qtds[modelo.strip()] = int(q.strip())
                            except ValueError:
                                qtds[modelo.strip()] = 0
                return qtds

            tipo_eq_salvo = parse_tipo_eq(v("tipo_equipamento"))
            qtds_venda = {}
            for modelo in MODELOS_VENDA:
                col_mod, col_qtd = st.columns([5, 2])
                col_mod.markdown(
                    f'<div style="padding:8px 0;color:#cccccc;font-weight:600;">{modelo}</div>',
                    unsafe_allow_html=True)
                qtds_venda[modelo] = col_qtd.number_input(
                    "Qtd prevista", min_value=0, value=tipo_eq_salvo.get(modelo, 0), step=1,
                    key=f"venda_{modelo.replace(' ','_').replace('.','')}"
                )

            tipo_eq = " | ".join(
                f"{m}: {q}" for m, q in qtds_venda.items() if q > 0
            ) or ""
            qtd = str(sum(qtds_venda.values())) if any(q > 0 for q in qtds_venda.values()) else ""

            st.markdown("---")
            OPT_MESES = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                         "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
            mes_salvo = v("janela_compra","")
            mes_idx   = OPT_MESES.index(mes_salvo) if mes_salvo in OPT_MESES else 0
            capex     = st.selectbox("📅 CAPEX — Mês de abertura", OPT_MESES, index=mes_idx)
            c1,c2 = st.columns(2)
            concorrentes = c1.text_area("Concorrentes cotados",     value=v("concorrentes_cotados"), height=110)
            dor          = c2.text_area("Dor operacional relatada", value=v("dor_operacional"),      height=110)

        with t3:
            if cliente != "Sim":
                st.info("⚠️ Esta aba está disponível apenas para **clientes ativos** (Já é cliente = Sim).")
                maq_inst  = v("maquinas_instaladas")
                hist_prob = v("historico_problemas")
                satisfacao = v("grau_satisfacao", "N/A")
            else:
                # ── Frota instalada por modelo ─────────────────────
                FROTA_MODELOS = {
                    "RAPTOR 13.000":    None,
                    "GIGANTE M 22.000": None,
                    "TRACTOR 22.000":   None,
                    "GIGANTE M 30.000": None,
                }

                # Parser: "RAPTOR 13.000: 3 | GIGANTE M 22.000: 0 | ..."
                def parse_frota(txt):
                    qtds = {}
                    if txt:
                        for parte in txt.split("|"):
                            parte = parte.strip()
                            if ":" in parte:
                                modelo, qtd = parte.split(":", 1)
                                try:
                                    qtds[modelo.strip()] = int(qtd.strip())
                                except ValueError:
                                    qtds[modelo.strip()] = 0
                    return qtds

                frota_salva = parse_frota(v("maquinas_instaladas"))

                st.markdown("##### 🚜 Máquinas Teston Instaladas")
                st.caption("Preencha a quantidade instalada em cada modelo (0 = não possui)")

                qtds_input = {}
                for modelo, _ in FROTA_MODELOS.items():
                    col_mod, col_qtd = st.columns([5, 2])
                    col_mod.markdown(
                        f'<div style="padding:8px 0;color:#cccccc;font-weight:600;">'
                        f'{modelo}</div>', unsafe_allow_html=True)
                    qtd_val = frota_salva.get(modelo, 0)
                    qtds_input[modelo] = col_qtd.number_input(
                        "Qtd instalada", min_value=0,
                        value=qtd_val, step=1,
                        key=f"frota_{modelo.replace(' ','_').replace('.','')}"
                    )

                # Serializa como string para salvar no SP
                maq_inst = " | ".join(
                    f"{modelo}: {qtd}"
                    for modelo, qtd in qtds_input.items()
                    if qtd > 0
                ) or ""

                st.markdown("---")
                hist_prob = st.text_area("📋 Histórico de problemas / chamados",
                                          value=v("historico_problemas"), height=90)
                OPT_SAT   = ["N/A","Baixo","Médio","Alto"]
                satisfacao = st.select_slider("⭐ Grau de satisfação", options=OPT_SAT,
                    value=v("grau_satisfacao","N/A") if v("grau_satisfacao","N/A") in OPT_SAT else "N/A")

        with t4:
            c1,c2 = st.columns(2)
            quem_dec  = c1.text_input("Quem decide a compra", value=v("quem_decide"))
            quem_bloq = c2.text_input("Quem pode bloquear",   value=v("quem_bloqueia"))
            fluxo = st.text_area("Como funciona a aprovação interna", value=v("fluxo_aprovacao"), height=90)
            OPT_FIN  = ["Desconhecida","Boa","Atenção","Risco"]
            OPT_REL  = ["","RUIM","MÉDIO","BOM"]
            c1,c2 = st.columns(2)
            sit_fin   = c1.selectbox("Situação financeira", OPT_FIN, index=vi("situacao_financeira", OPT_FIN))
            rel_salva = v("relacao_concorrentes","")
            rel_idx   = OPT_REL.index(rel_salva) if rel_salva in OPT_REL else 0
            rel_conc  = c2.selectbox("Relação com concorrentes", OPT_REL, index=rel_idx)
            capex     = v("janela_compra","")  # mantém valor existente

        st.divider()
        cidade_uf = f"{row.get('CIDADE','')} - {row.get('ESTADO','')}"

        def montar_payload():
            return dict(
                usina=row["USINA"], cidade_estado=cidade_uf,
                marca_equipamento=marca_eq,
                vendedor=row["VENDEDOR"], moagem_anual=moagem, frota_atual=frota,
                e_cliente=cliente, grau_relacionamento=grau.split(" ",1)[-1],
                ultima_visita=str(ult_visita), proxima_visita=str(prox_visita),
                participou_evento=participou_ev,
                testou_maquina=testou_maq,
                maquinas_testadas=", ".join(maquinas_testadas),
                nome_contato=nome_ctt, cargo_contato=cargo_ctt, telefone_contato=fmt_telefone(tel_ctt),
                nome_contato2=nome_ctt2, cargo_contato2=cargo_ctt2, telefone_contato2=tel_ctt2,
                pretende_investir=investe, quantidade_prevista=qtd,
                tipo_equipamento=tipo_eq, concorrentes_cotados=concorrentes, dor_operacional=dor,
                maquinas_instaladas=maq_inst, historico_problemas=hist_prob, grau_satisfacao=satisfacao,
                quem_decide=quem_dec, quem_bloqueia=quem_bloq, fluxo_aprovacao=fluxo,
                janela_compra=capex, relacao_concorrentes=rel_conc, situacao_financeira=sit_fin,
            )

        if editando:
            bc1,bc2 = st.columns(2)
            if bc1.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True):
                try:
                    atualizar_crm(reg_id, **montar_payload())
                    st.success("✅ Registro atualizado!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao atualizar: {e}")
            if bc2.form_submit_button("➕ Salvar como Nova Visita", use_container_width=True):
                try:
                    salvar_crm(**montar_payload())
                    st.success("✅ Nova visita registrada!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")
        else:
            if st.form_submit_button("💾 Salvar Registro Comercial", type="primary", use_container_width=True):
                try:
                    salvar_crm(**montar_payload())
                    st.success(f"✅ Visita de **{row['USINA']}** salva!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

    if editando:
        st.divider()
        with st.expander("🗑️ Zona de Perigo"):
            st.warning("Esta ação é irreversível.")
            if st.button("Excluir este registro", type="secondary"):
                try:
                    deletar_registro(reg_id)
                    st.cache_data.clear()
                    voltar_grid()
                except Exception as e:
                    st.error(f"Erro ao excluir: {e}")
    st.stop()

# ============================================================
# MODO GRID
# ============================================================
st.title("🏭 Inteligência Comercial Teston")

aba_forms, aba_mapa, aba_dash, aba_agenda = st.tabs([
    "📋  Formulários de Visita",
    "🗺️  Mapa de Cobertura",
    "📊  Dashboard",
    "📅  Agenda"
])

# ============================================================
# ABA 1 — FORMULÁRIOS
# ============================================================
with aba_forms:
    if df_usinas.empty:
        st.info("Nenhuma usina carregada.")
        st.stop()

    df_filtrado = df_usinas.copy()
    if vendedor_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["VENDEDOR"] == vendedor_sel]
    if busca_nome:
        df_filtrado = df_filtrado[df_filtrado["USINA"].str.contains(busca_nome, case=False, na=False)]
    if estado_sel != "Todos" and "ESTADO" in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado["ESTADO"] == estado_sel]
    if regiao_sel != "Todas" and "REGIAO" in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado["REGIAO"] == regiao_sel]

    ultima_map = ultima_visita_por_usina()
    if apenas_sem_visita:
        df_filtrado = df_filtrado[~df_filtrado["USINA"].isin(ultima_map.keys())]

    total_u   = len(df_filtrado)
    visitadas = sum(1 for r in df_filtrado.itertuples() if r.USINA in ultima_map)
    mc1,mc2,mc3 = st.columns(3)
    mc1.metric("Usinas listadas", total_u)
    mc2.metric("Com visita", visitadas)
    mc3.metric("Sem visita", total_u - visitadas)
    st.divider()

    st.markdown("""
    <div style="display:flex;gap:18px;align-items:center;margin-bottom:12px;font-size:12px;">
        <span style="font-weight:700;color:#888;">Status:</span>
        <span class="status-tag tag-visitada">🟢 Visitada</span>
        <span class="status-tag tag-programada">🟡 Visita Programada</span>
        <span class="status-tag tag-pendente">🔴 Pendente</span>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    for idx, row in enumerate(df_filtrado.itertuples()):
        col_atual    = col1 if idx % 2 == 0 else col2
        cidade_uf    = f"{row.CIDADE} - {row.ESTADO}" if hasattr(row, "ESTADO") else str(row.CIDADE)
        regiao_usina = getattr(row, "REGIAO", "") or ""
        info_visita  = ultima_map.get(row.USINA)
        row_dict     = row._asdict()

        prox_str = info_visita.get("proxima_visita","") if info_visita else ""
        try:
            prox_futura = datetime.strptime(str(prox_str)[:10], "%Y-%m-%d").date() > date.today()
        except Exception:
            prox_futura = False

        if not info_visita:
            css_class, tag_class, tag_label = "card-pendente",   "tag-pendente",   "🔴 Pendente"
        elif prox_futura:
            css_class, tag_class, tag_label = "card-programada", "tag-programada", "🟡 Programada"
        else:
            css_class, tag_class, tag_label = "card-visitada",   "tag-visitada",   "🟢 Visitada"

        with col_atual:
            with st.container(border=True):
                st.markdown(f'<div class="{css_class}" style="display:none;"></div>', unsafe_allow_html=True)
                h1, h2 = st.columns([3, 1])
                with h1:
                    st.markdown(f"**{row.USINA}**")
                    loc = f"📍 {cidade_uf}"
                    if regiao_usina: loc += f" · 🗺️ {regiao_usina}"
                    loc += f" · 👤 {row.VENDEDOR}"
                    st.caption(loc)
                with h2:
                    st.markdown(
                        f'<div style="text-align:right;margin-top:4px;">'
                        f'<span class="status-tag {tag_class}">{tag_label}</span></div>',
                        unsafe_allow_html=True)

                if info_visita:
                    ultima_fmt = fmt_data(info_visita.get("ultima",""))
                    prox_fmt   = fmt_data(info_visita.get("proxima_visita",""))
                    grau_e     = GRAU_EMOJI.get(info_visita.get("grau_relacionamento",""),"")
                    st.markdown(
                        f'<div class="card-dates">'
                        f'<span>📅 Última: <b>{ultima_fmt}</b> {grau_e}</span>'
                        f'<span>📆 Próxima: <b>{prox_fmt}</b></span>'
                        f'</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="card-dates" style="opacity:0;">─</div>', unsafe_allow_html=True)

                b1, b2 = st.columns(2)
                if b1.button("📝 Nova Visita", key=f"nova_{row.Index}", use_container_width=True, type="primary"):
                    abrir_formulario(row_dict)
                if info_visita:
                    if b2.button("✏️ Editar Última", key=f"edit_{row.Index}", use_container_width=True):
                        abrir_formulario(row_dict, registro_id=info_visita["id"])
                else:
                    b2.button("✏️ Editar Última", key=f"edit_dis_{row.Index}", use_container_width=True, disabled=True)

                with st.expander("📜 Histórico de visitas"):
                    if info_visita:
                        df_hist = historico_usina(row.USINA)
                        if not df_hist.empty:
                            for _, hr in df_hist.iterrows():
                                hc1,hc2,hc3,hc4 = st.columns([2,2,2,1])
                                hc1.caption(fmt_data(hr["data_registro"]))
                                hc2.caption(GRAU_EMOJI.get(hr["grau_relacionamento"],"") + " " + str(hr["grau_relacionamento"]))
                                hc3.caption(str(hr["pretende_investir"]))
                                if hc4.button("✏️", key=f"hedit_{hr['id']}", help="Editar este registro"):
                                    abrir_formulario(row_dict, registro_id=str(hr["id"]))
                    else:
                        st.caption("Nenhuma visita registrada ainda para esta usina.")

# ============================================================
# ABA 2 — MAPA
# ============================================================
with aba_mapa:
    st.markdown("""
    <style>
    .neon-title { text-align:center; font-size:22px; font-weight:900; letter-spacing:2px; color:#00f0ff;
        text-shadow:0 0 8px #00f0ff,0 0 20px #00aaff; padding:14px 0 10px 0;
        border-bottom:1px solid #00f0ff44; margin-bottom:14px; }
    .neon-card { background:#0d1526; border:1px solid #00f0ff66; border-radius:6px; padding:10px 16px;
        text-align:center; box-shadow:0 0 10px #00f0ff22,inset 0 0 10px #00f0ff08; }
    .neon-card .kpi-label { font-size:12px; color:#7ecfdf; font-weight:600; letter-spacing:1px; text-transform:uppercase; }
    .neon-card .kpi-value { font-size:32px; font-weight:900; color:#00f0ff; text-shadow:0 0 8px #00f0ff; line-height:1.1; }
    .neon-card .kpi-sub   { font-size:11px; color:#4a7a88; margin-top:2px; }
    .filtro-panel  { background:#0d1526; border:1px solid #00f0ff44; border-radius:8px; padding:14px; }
    .filtro-titulo { font-size:11px; font-weight:700; letter-spacing:1.5px; color:#00f0ff; text-transform:uppercase;
        margin:12px 0 4px 0; border-bottom:1px solid #00f0ff33; padding-bottom:3px; }
    .legenda-linha { display:flex; align-items:center; gap:8px; font-size:12px; color:#aad4dd; margin:5px 0; }
    .legenda-dot   { width:12px; height:12px; border-radius:50%; display:inline-block; flex-shrink:0; }
    .selected-panel { background:#0d1526; border:1px solid #00FF8866; border-left:4px solid #00FF88;
        border-radius:6px; padding:14px 18px; margin-top:12px; }
    .selected-panel h4 { color:#00FF88; margin:0 0 4px 0; font-size:16px; }
    .selected-panel p  { color:#aaddcc; font-size:13px; margin:2px 0; }
    </style>
    """, unsafe_allow_html=True)

    if df_usinas.empty or not COL_LAT or not COL_LON:
        st.info("Dados de usinas ou coordenadas não disponíveis.")
        st.stop()

    ultima_map_m = ultima_visita_por_usina()
    CORES_VENDEDOR = {"Wanderson": "#e74c3c", "Fernando": "#2777c4"}

    def cor_vendedor(nome):
        for k, cv in CORES_VENDEDOR.items():
            if k.lower() in str(nome).lower():
                return cv
        return "#7f8c8d"

    safra = "2026/2027"
    st.markdown(f'<div class="neon-title">🏭 MAPA DE USINAS — SAFRA {safra}</div>', unsafe_allow_html=True)

    df_base  = df_usinas.dropna(subset=[COL_LAT, COL_LON]).copy()
    total_u  = len(df_base)
    kpi_cols = st.columns(4)
    estados_top = df_base["ESTADO"].value_counts().head(3).index.tolist() if "ESTADO" in df_base.columns else []
    kpi_dados   = [("Total Usinas", total_u, "carteira completa")]
    for est in estados_top:
        kpi_dados.append((f"Usinas {est}", (df_base["ESTADO"] == est).sum(), f"estado {est}"))
    while len(kpi_dados) < 4:
        kpi_dados.append(("—","—",""))
    for col, (label, valor, sub) in zip(kpi_cols, kpi_dados[:4]):
        col.markdown(f'<div class="neon-card"><div class="kpi-label">{label}</div>'
                     f'<div class="kpi-value">{valor}</div><div class="kpi-sub">{sub}</div></div>',
                     unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_mapa, col_filtros = st.columns([3, 1])
    with col_filtros:
        st.markdown('<div class="filtro-panel">', unsafe_allow_html=True)
        st.markdown('<div class="filtro-titulo">Estado</div>', unsafe_allow_html=True)
        estados_disp = ["Todos"] + (sorted(df_base["ESTADO"].dropna().unique().tolist()) if "ESTADO" in df_base.columns else [])
        estado_sel_m = st.radio("Estado", estados_disp, horizontal=True, key="mapa_estado", label_visibility="collapsed")

        df_f = df_base.copy()
        if estado_sel_m != "Todos" and "ESTADO" in df_f.columns:
            df_f = df_f[df_f["ESTADO"] == estado_sel_m]
        st.markdown('<div class="filtro-titulo">Cidade</div>', unsafe_allow_html=True)
        cidades = ["Todos"] + (sorted(df_f["CIDADE"].dropna().unique().tolist()) if "CIDADE" in df_f.columns else [])
        cidade_sel_m = st.selectbox("Cidade", cidades, key="mapa_cidade", label_visibility="collapsed")

        if cidade_sel_m != "Todos" and "CIDADE" in df_f.columns:
            df_f = df_f[df_f["CIDADE"] == cidade_sel_m]
        st.markdown('<div class="filtro-titulo">Usina</div>', unsafe_allow_html=True)
        usinas_lista = ["Todos"] + sorted(df_f["USINA"].dropna().unique().tolist())
        usina_sel_m = st.selectbox("Usina", usinas_lista, key="mapa_usina", label_visibility="collapsed")

        st.markdown('<div class="filtro-titulo">Vendedor</div>', unsafe_allow_html=True)
        vends = ["Todos"] + (sorted(df_base["VENDEDOR"].dropna().unique().tolist()) if "VENDEDOR" in df_base.columns else [])
        vendedor_sel_m = st.selectbox("Vendedor", vends, key="mapa_vend", label_visibility="collapsed")

        st.markdown('<div class="filtro-titulo">Status de Visita</div>', unsafe_allow_html=True)
        status_sel_m = st.radio("Status", ["Todas","✅ Visitadas","⬜ Sem visita"], key="mapa_status", label_visibility="collapsed")

        st.markdown('<div class="filtro-titulo">Legenda</div>', unsafe_allow_html=True)
        for nome, cor in CORES_VENDEDOR.items():
            st.markdown(f'<div class="legenda-linha"><span class="legenda-dot" style="background:{cor};"></span>{nome}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col_mapa:
        df_map = df_base.copy()
        if estado_sel_m    != "Todos"  and "ESTADO"   in df_map.columns: df_map = df_map[df_map["ESTADO"]   == estado_sel_m]
        if cidade_sel_m    != "Todos"  and "CIDADE"   in df_map.columns: df_map = df_map[df_map["CIDADE"]   == cidade_sel_m]
        if usina_sel_m     != "Todos":                                    df_map = df_map[df_map["USINA"]    == usina_sel_m]
        if vendedor_sel_m  != "Todos"  and "VENDEDOR" in df_map.columns: df_map = df_map[df_map["VENDEDOR"] == vendedor_sel_m]
        if status_sel_m == "✅ Visitadas":  df_map = df_map[ df_map["USINA"].isin(ultima_map_m.keys())]
        if status_sel_m == "⬜ Sem visita": df_map = df_map[~df_map["USINA"].isin(ultima_map_m.keys())]

        lat_c = df_map[COL_LAT].mean() if not df_map.empty else -20.0
        lon_c = df_map[COL_LON].mean() if not df_map.empty else -48.0
        zoom  = 7 if usina_sel_m != "Todos" else 6

        m = folium.Map(location=[lat_c, lon_c], zoom_start=zoom, tiles="CartoDB positron", control_scale=True)
        Fullscreen(position="topright").add_to(m)
        MiniMap(toggle_display=True, position="bottomright", zoom_level_offset=-6, width=130, height=100).add_to(m)

        def popup_html(ru, info):
            nome  = ru.get("USINA","")
            cid   = f"{ru.get('CIDADE','')} - {ru.get('ESTADO','')}"
            vend  = ru.get("VENDEDOR","")
            cor_v = cor_vendedor(vend)
            if info:
                grau  = info.get("grau_relacionamento","")
                inv   = info.get("pretende_investir","")
                ult   = fmt_data(info.get("ultima",""))
                prox  = fmt_data(info.get("proxima_visita",""))
                gc    = GRAU_COLOR.get(grau,"#aaa")
                body  = (f'<span style="background:{gc};color:#fff;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px;">'
                         f'{GRAU_EMOJI.get(grau,"")} {grau}</span><br>'
                         f'<span style="font-size:12px;">📅 Última: <b>{ult}</b></span><br>'
                         f'<span style="font-size:12px;">📅 Próxima: <b>{prox}</b></span><br>'
                         f'<span style="font-size:12px;">💰 Vai investir: <b>{INV_EMOJI.get(inv,"")} {inv}</b></span>')
            else:
                body = '<span style="background:#e74c3c;color:#fff;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px;">⬜ Sem visita</span>'
            return (f'<div style="font-family:sans-serif;padding:4px;min-width:220px;">'
                    f'<div style="border-left:4px solid {cor_v};padding-left:8px;margin-bottom:6px;">'
                    f'<b style="font-size:14px;color:#111;">{nome}</b><br>'
                    f'<span style="font-size:11px;color:#666;">📍 {cid}</span><br>'
                    f'<span style="font-size:11px;color:{cor_v};font-weight:700;">👤 {vend}</span></div>{body}</div>')

        for _, ru in df_map.iterrows():
            info_u   = ultima_map_m.get(ru["USINA"])
            cor_v    = cor_vendedor(ru.get("VENDEDOR",""))
            visitada = info_u is not None
            folium.CircleMarker(
                location=[ru[COL_LAT], ru[COL_LON]], radius=7,
                color=cor_v, weight=2, fill=True,
                fill_color=cor_v if visitada else "#ffffff",
                fill_opacity=0.85 if visitada else 0.0,
                tooltip=folium.Tooltip(
                    f"<b>{ru['USINA']}</b><br>{'✅ Visitada' if visitada else '⬜ Sem visita'} · {ru.get('VENDEDOR','')}",
                    sticky=True),
                popup=folium.Popup(folium.IFrame(popup_html(ru.to_dict(), info_u), width=265, height=200), max_width=275)
            ).add_to(m)

        map_data = st_folium(m, use_container_width=True, height=560, returned_objects=["last_object_clicked_tooltip"])

        clicked_name = map_data.get("last_object_clicked_tooltip")
        if clicked_name and not df_usinas.empty:
            match = df_usinas[df_usinas["USINA"] == clicked_name]
            if not match.empty:
                row_c  = match.iloc[0].to_dict()
                info_c = ultima_map_m.get(clicked_name)
                cor_v  = cor_vendedor(row_c.get("VENDEDOR",""))
                grau_e = GRAU_EMOJI.get(info_c.get("grau_relacionamento",""),"") if info_c else ""
                status_txt = (f"Visitada em {fmt_data(info_c.get('ultima',''))} {grau_e}" if info_c else "⬜ Sem visita registrada")
                st.markdown(f"""
                <div class="selected-panel">
                    <h4>📍 {clicked_name}</h4>
                    <p>🏙️ {row_c.get('CIDADE','')} – {row_c.get('ESTADO','')}
                       &nbsp;|&nbsp; 👤 <span style="color:{cor_v};font-weight:700;">{row_c.get('VENDEDOR','')}</span>
                       &nbsp;|&nbsp; {status_txt}</p>
                </div>""", unsafe_allow_html=True)
                btn1, btn2, _ = st.columns([1,1,2])
                if btn1.button("📝 Nova Visita",   type="primary", use_container_width=True, key="mn"):
                    abrir_formulario(row_c)
                if info_c and btn2.button("✏️ Editar Última", use_container_width=True, key="me"):
                    abrir_formulario(row_c, registro_id=info_c["id"])

    total_f = len(df_map)
    vis_f   = sum(1 for u in df_map["USINA"] if u in ultima_map_m)
    pct_f   = round(vis_f / total_f * 100) if total_f else 0
    st.markdown(f"""
    <div style="text-align:center;color:#4a7a88;font-size:12px;margin-top:6px;">
        Exibindo <b style="color:#00f0ff">{total_f}</b> usinas &nbsp;|&nbsp;
        <b style="color:#00FF88">{vis_f}</b> visitadas &nbsp;|&nbsp;
        <b style="color:#e74c3c">{total_f - vis_f}</b> sem visita &nbsp;|&nbsp;
        Cobertura: <b style="color:#00f0ff">{pct_f}%</b>
    </div>""", unsafe_allow_html=True)

    with st.expander("📋 Ver lista"):
        cols_r = [c for c in ["USINA","CIDADE","ESTADO","VENDEDOR"] if c in df_map.columns]
        df_r   = df_map[cols_r].copy()
        df_r["Status"]        = df_r["USINA"].apply(lambda u: "✅ Visitada" if u in ultima_map_m else "⬜ Sem visita")
        df_r["Última visita"] = df_r["USINA"].apply(lambda u: fmt_data(ultima_map_m[u]["ultima"]) if u in ultima_map_m else "—")
        st.dataframe(df_r, use_container_width=True, hide_index=True)

# ============================================================
# ABA 3 — DASHBOARD
# ============================================================
with aba_dash:
    df_reg = carregar_registros()
    if df_reg.empty:
        st.info("📭 Nenhum registro ainda.")
        st.stop()

    st.markdown("### 📊 Dashboard Comercial")
    fc1,fc2,fc3 = st.columns([2,2,3])
    with fc1:
        vendedor_dash = st.selectbox("Filtrar vendedor",
            ["Todos"] + sorted(df_reg["vendedor"].dropna().unique().tolist()), key="dash_vend")
    with fc2:
        if "regiao" in df_reg.columns and df_reg["regiao"].notna().any():
            regiao_dash = st.selectbox("Filtrar região",
                ["Todas"] + sorted(df_reg["regiao"].dropna().unique().tolist()), key="dash_reg")
        else:
            regiao_dash = "Todas"
    with fc3:
        periodo = st.date_input("Período", value=(date(2024,1,1), date.today()), key="dash_periodo", format="DD/MM/YYYY")

    df_d = df_reg.copy()
    df_d["data_registro"] = pd.to_datetime(df_d["data_registro"], errors="coerce", utc=True).dt.tz_localize(None)
    if vendedor_dash != "Todos": df_d = df_d[df_d["vendedor"] == vendedor_dash]
    if regiao_dash   != "Todas": df_d = df_d[df_d["regiao"]   == regiao_dash]
    if len(periodo)  == 2:
        df_d = df_d[(df_d["data_registro"].dt.date >= periodo[0]) &
                    (df_d["data_registro"].dt.date <= periodo[1])]

    if df_d.empty:
        st.warning("Nenhum registro no filtro.")
        st.stop()

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Total Visitas",      len(df_d))
    k2.metric("Usinas Únicas",      df_d["usina"].nunique())
    k3.metric("Clientes Ativos",    (df_d["e_cliente"]=="Sim").sum())
    k4.metric("Leads Quentes 🔴",   (df_d["grau_relacionamento"]=="Quente").sum())
    k5.metric("Pretendem Investir", (df_d["pretende_investir"]=="Sim").sum())
    k6.metric("Participaram Evento", (df_d["participou_evento"]=="Sim").sum())
    st.divider()

    r1c1,r1c2 = st.columns(2)
    with r1c1:
        st.markdown("##### 🌡️ Temperatura do Relacionamento")
        gc = df_d["grau_relacionamento"].value_counts().reset_index(); gc.columns = ["Grau","Qtd"]
        fig = px.pie(gc, names="Grau", values="Qtd", hole=0.45,
                     color="Grau", color_discrete_map={"Quente":"#FF4B4B","Morno":"#FFD700","Frio":"#4B9CFF"})
        fig.update_traces(textinfo="label+percent")
        fig.update_layout(showlegend=False, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    with r1c2:
        st.markdown("##### 💰 Intenção de Investimento")
        ic = df_d["pretende_investir"].value_counts().reset_index(); ic.columns = ["Resp","Qtd"]
        fig = px.bar(ic, x="Resp", y="Qtd", color="Resp", text="Qtd",
                     color_discrete_map={"Sim":"#00FF88","Não":"#FF4B4B","Indefinido":"#FFD700"})
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=10,b=10), xaxis_title="", yaxis_title="Qtd")
        st.plotly_chart(fig, use_container_width=True)

    r2c1,r2c2 = st.columns(2)
    with r2c1:
        if "regiao" in df_d.columns and df_d["regiao"].notna().any():
            st.markdown("##### 🗺️ Visitas por Região")
            rc = df_d["regiao"].value_counts().reset_index(); rc.columns = ["Região","Qtd"]
            fig = px.bar(rc, x="Qtd", y="Região", orientation="h",
                         color="Qtd", color_continuous_scale="teal", text="Qtd")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False, coloraxis_showscale=False,
                               margin=dict(t=10,b=10), yaxis_title="", xaxis_title="Visitas")
            st.plotly_chart(fig, use_container_width=True)
    with r2c2:
        st.markdown("##### 👤 Visitas por Vendedor")
        vc = df_d["vendedor"].value_counts().reset_index(); vc.columns = ["Vendedor","Qtd"]
        fig = px.bar(vc, x="Qtd", y="Vendedor", orientation="h",
                     color="Qtd", color_continuous_scale="teal", text="Qtd")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, coloraxis_showscale=False,
                           margin=dict(t=10,b=10), yaxis_title="", xaxis_title="Registros")
        st.plotly_chart(fig, use_container_width=True)

    r3c1,r3c2 = st.columns(2)
    with r3c1:
        st.markdown("##### ⭐ Satisfação (Clientes Ativos)")
        df_cli = df_d[df_d["e_cliente"]=="Sim"]
        if not df_cli.empty:
            sc = df_cli["grau_satisfacao"].value_counts().reset_index(); sc.columns = ["Sat","Qtd"]
            fig = px.pie(sc, names="Sat", values="Qtd", hole=0.45,
                         color="Sat", color_discrete_map={"Alto":"#00FF88","Médio":"#FFD700","Baixo":"#FF4B4B","N/A":"#555"})
            fig.update_traces(textinfo="label+percent")
            fig.update_layout(showlegend=False, margin=dict(t=10,b=10))
            st.plotly_chart(fig, use_container_width=True)
    with r3c2:
        st.markdown("##### 💳 Situação Financeira")
        fc2d = df_d["situacao_financeira"].value_counts().reset_index(); fc2d.columns = ["Situação","Qtd"]
        fig = px.bar(fc2d, x="Situação", y="Qtd", color="Situação", text="Qtd",
                     color_discrete_map={"Boa":"#00FF88","Atenção":"#FFD700","Risco":"#FF4B4B","Desconhecida":"#888"})
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=10,b=10), xaxis_title="", yaxis_title="Qtd")
        st.plotly_chart(fig, use_container_width=True)

    r4c1,r4c2 = st.columns(2)
    with r4c1:
        st.markdown("##### 🎪 Participaram de Evento?")
        fab = df_d["participou_evento"].value_counts().reset_index(); fab.columns = ["V","Qtd"]
        fig = px.pie(fab, names="V", values="Qtd", hole=0.45,
                     color="V", color_discrete_map={"Sim":"#00FF88","Não":"#444"})
        fig.update_traces(textinfo="label+percent")
        fig.update_layout(showlegend=False, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    with r4c2:
        st.markdown("##### 🚜 Testaram a Máquina?")
        tst = df_d["testou_maquina"].value_counts().reset_index(); tst.columns = ["T","Qtd"]
        fig = px.pie(tst, names="T", values="Qtd", hole=0.45,
                     color="T", color_discrete_map={"Sim":"#00CFFF","Não":"#444"})
        fig.update_traces(textinfo="label+percent")
        fig.update_layout(showlegend=False, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### 📅 Registros ao Longo do Tempo")
    tl = df_d.groupby(df_d["data_registro"].dt.date).size().reset_index()
    tl.columns = ["Data","Registros"]
    fig = px.bar(tl, x="Data", y="Registros", color_discrete_sequence=["#00FF88"])
    fig.update_layout(margin=dict(t=10,b=10), xaxis_title="", bargap=0.3)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### 📆 Próximas Visitas Agendadas")
    df_prox = df_d[["usina","vendedor","proxima_visita","grau_relacionamento"]].copy()
    df_prox["proxima_visita"] = pd.to_datetime(df_prox["proxima_visita"], errors="coerce")
    df_prox = df_prox[df_prox["proxima_visita"] >= pd.Timestamp.today()]
    df_prox = df_prox.sort_values("proxima_visita").drop_duplicates("usina")
    if not df_prox.empty:
        df_prox["proxima_visita"] = df_prox["proxima_visita"].dt.strftime("%d/%m/%Y")
        st.dataframe(df_prox.rename(columns={
            "usina":"Usina","vendedor":"Vendedor",
            "proxima_visita":"Próxima Visita","grau_relacionamento":"Temperatura"
        }), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma visita futura registrada.")

    st.divider()
    cols_show = [c for c in ["usina","cidade_estado","regiao","vendedor",
                              "grau_relacionamento","pretende_investir",
                              "situacao_financeira","data_registro"] if c in df_d.columns]
    df_show = df_d[cols_show].copy()
    if "data_registro" in df_show.columns:
        df_show["data_registro"] = df_show["data_registro"].apply(lambda x: fmt_data(x, hora=True))
    st.dataframe(
        df_show.rename(columns={
            "usina":"Usina","cidade_estado":"Cidade/UF","regiao":"Região",
            "vendedor":"Vendedor","grau_relacionamento":"Temperatura",
            "pretende_investir":"Vai Investir?",
            "situacao_financeira":"Situação Fin.","data_registro":"Registrado em"
        }), use_container_width=True, hide_index=True)

    csv = df_d.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Exportar (.csv)", data=csv,
        file_name=f"crm_teston_{datetime.today().strftime('%Y%m%d')}.csv", mime="text/csv")

# ============================================================
# ABA 4 — AGENDA (CALENDÁRIO DE VISITAS)
# ============================================================
with aba_agenda:
    import calendar as cal_mod
    import plotly.graph_objects as go

    df_reg_ag = carregar_registros()

    st.markdown("### 📅 Agenda de Visitas")

    # ── Filtros ──────────────────────────────────────────────
    ag1, ag2, ag3 = st.columns([2, 2, 3])
    with ag1:
        anos_disp = list(range(date.today().year, date.today().year + 3))
        ano_sel   = ag1.selectbox("Ano", anos_disp, index=0, key="ag_ano")
    with ag2:
        MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
        mes_sel  = ag2.selectbox("Mês", MESES_PT, index=date.today().month - 1, key="ag_mes")
        mes_num  = MESES_PT.index(mes_sel) + 1
    with ag3:
        if not df_reg_ag.empty and "vendedor" in df_reg_ag.columns:
            vends_ag = ["Todos"] + sorted(df_reg_ag["vendedor"].dropna().unique().tolist())
        else:
            vends_ag = ["Todos"]
        vend_ag = ag3.selectbox("Vendedor", vends_ag, key="ag_vend")

    # ── Montar df de próximas visitas ────────────────────────
    if df_reg_ag.empty:
        st.info("Nenhum registro encontrado.")
        st.stop()

    df_ag = df_reg_ag.copy()
    df_ag["proxima_visita"] = pd.to_datetime(df_ag["proxima_visita"], errors="coerce")
    df_ag = df_ag.dropna(subset=["proxima_visita"])
    df_ag = df_ag.sort_values("proxima_visita").drop_duplicates("usina")

    if vend_ag != "Todos":
        df_ag = df_ag[df_ag["vendedor"] == vend_ag]

    # Filtrar pelo mês/ano selecionado
    df_mes = df_ag[
        (df_ag["proxima_visita"].dt.year  == ano_sel) &
        (df_ag["proxima_visita"].dt.month == mes_num)
    ].copy()

    # ── Calendário gráfico ───────────────────────────────────
    primeiro_dia, total_dias = cal_mod.monthrange(ano_sel, mes_num)
    # primeiro_dia: 0=seg ... 6=dom (padrão python) → ajustar para domingo=0
    primeiro_dia_dom = (primeiro_dia + 1) % 7  # converter para domingo=0

    DIAS_SEMANA = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"]
    CORES_VEND  = {"Fernando": "#2777c4", "Wanderson": "#e74c3c"}
    COR_DEFAULT = "#00FF88"

    # Mapear dia → visitas
    visitas_por_dia = {}
    for _, r in df_mes.iterrows():
        d = r["proxima_visita"].day
        if d not in visitas_por_dia:
            visitas_por_dia[d] = []
        visitas_por_dia[d].append(r)

    # Grid de semanas
    semanas = []
    semana_atual = [None] * primeiro_dia_dom
    for dia in range(1, total_dias + 1):
        semana_atual.append(dia)
        if len(semana_atual) == 7:
            semanas.append(semana_atual)
            semana_atual = []
    if semana_atual:
        while len(semana_atual) < 7:
            semana_atual.append(None)
        semanas.append(semana_atual)

    n_semanas = len(semanas)

    fig = go.Figure()
    fig.update_layout(
        height=120 + n_semanas * 110,
        margin=dict(t=40, b=10, l=10, r=10),
        paper_bgcolor="#0c1711",
        plot_bgcolor="#0c1711",
        showlegend=False,
        xaxis=dict(visible=False, range=[-0.5, 6.5]),
        yaxis=dict(visible=False, range=[-n_semanas - 0.2, 0.8]),
    )

    # Cabeçalho dos dias da semana
    for col, nome in enumerate(DIAS_SEMANA):
        fig.add_annotation(
            x=col, y=0.6,
            text=f"<b>{nome}</b>",
            showarrow=False,
            font=dict(color="#00FF88", size=13),
            xanchor="center"
        )

    hoje = date.today()

    for s_idx, semana in enumerate(semanas):
        y_row = -(s_idx + 0.5)
        for col, dia in enumerate(semana):
            if dia is None:
                continue

            # Célula de fundo
            is_hoje = (dia == hoje.day and mes_num == hoje.month and ano_sel == hoje.year)
            cor_fundo = "#1a3a1a" if is_hoje else "#111d14"
            cor_borda = "#00FF88" if is_hoje else "rgba(0,255,136,0.13)"

            fig.add_shape(
                type="rect",
                x0=col - 0.45, x1=col + 0.45,
                y0=y_row - 0.44, y1=y_row + 0.44,
                fillcolor=cor_fundo,
                line=dict(
                    color="rgba(0,255,136,0.9)" if is_hoje else "rgba(0,255,136,0.13)",
                    width=1.5 if is_hoje else 0.5
                ),
                layer="below"
            )

            # Número do dia
            fig.add_annotation(
                x=col, y=y_row + 0.28,
                text=f"<b>{dia}</b>",
                showarrow=False,
                font=dict(color="#00FF88" if is_hoje else "#888888", size=12),
                xanchor="center"
            )

            # Alfinetes das visitas
            if dia in visitas_por_dia:
                visitas = visitas_por_dia[dia]
                for v_idx, visita in enumerate(visitas[:3]):  # max 3 por célula
                    vend_nome = str(visita.get("vendedor",""))
                    cor_pin   = CORES_VEND.get(vend_nome, COR_DEFAULT)
                    usina_nm  = str(visita.get("usina",""))[:18]
                    grau_e    = {"Quente":"🔴","Morno":"🟡","Frio":"🔵"}.get(
                        str(visita.get("grau_relacionamento","")), "⚪")

                    y_pin = y_row + 0.05 - v_idx * 0.18

                    # Pin (marcador)
                    fig.add_trace(go.Scatter(
                        x=[col], y=[y_pin],
                        mode="markers+text",
                        marker=dict(symbol="circle", size=10, color=cor_pin,
                                    line=dict(color="white", width=1)),
                        text=f" {grau_e} {usina_nm}",
                        textposition="middle right",
                        textfont=dict(color=cor_pin, size=9),
                        hovertemplate=(
                            f"<b>{visita.get('usina','')}</b><br>"
                            f"👤 {vend_nome}<br>"
                            f"📅 {visita['proxima_visita'].strftime('%d/%m/%Y')}<br>"
                            f"🌡️ {visita.get('grau_relacionamento','')}<br>"
                            f"💰 Investe: {visita.get('pretende_investir','')}"
                            "<extra></extra>"
                        ),
                        showlegend=False,
                    ))

                # Indicador se tem mais de 3
                if len(visitas) > 3:
                    fig.add_annotation(
                        x=col, y=y_row - 0.38,
                        text=f"+{len(visitas)-3}",
                        showarrow=False,
                        font=dict(color="#888", size=8),
                        xanchor="center"
                    )

    fig.add_annotation(
        x=3, y=0.85,
        text=f"<b>{mes_sel.upper()} {ano_sel}</b>",
        showarrow=False,
        font=dict(color="#ffffff", size=16),
        xanchor="center"
    )

    st.plotly_chart(fig, use_container_width=True)

    # ── Lista de eventos do mês ───────────────────────────────
    st.divider()
    st.markdown(f"#### 📋 Visitas em {mes_sel} {ano_sel}")

    if df_mes.empty:
        st.info(f"Nenhuma visita agendada para {mes_sel} {ano_sel}.")
    else:
        df_show_ag = df_mes[["usina","vendedor","proxima_visita",
                              "grau_relacionamento","pretende_investir",
                              "cidade_estado"]].copy()
        df_show_ag["proxima_visita"] = df_show_ag["proxima_visita"].dt.strftime("%d/%m/%Y")
        df_show_ag = df_show_ag.sort_values("proxima_visita")
        st.dataframe(
            df_show_ag.rename(columns={
                "usina":"Usina", "vendedor":"Vendedor",
                "proxima_visita":"Data da Visita",
                "grau_relacionamento":"Temperatura",
                "pretende_investir":"Vai Investir?",
                "cidade_estado":"Cidade/UF"
            }),
            use_container_width=True, hide_index=True
        )

    # ── Legenda ───────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex;gap:20px;align-items:center;font-size:12px;color:#888;margin-top:8px;">
        <span>Legenda:</span>
        <span><span style="color:#2777c4;font-size:16px;">●</span> Fernando</span>
        <span><span style="color:#e74c3c;font-size:16px;">●</span> Wanderson</span>
        <span><span style="color:#00FF88;font-size:16px;">●</span> Outros</span>
        <span>🔴 Quente &nbsp; 🟡 Morno &nbsp; 🔵 Frio</span>
        <span style="border:1px solid #00FF88;padding:1px 6px;border-radius:4px;color:#00FF88;">Hoje</span>
    </div>
    """, unsafe_allow_html=True)