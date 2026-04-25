import pandas as pd
import streamlit as st
import time

st.set_page_config(page_title="Painel de Faturamento", page_icon="📊", layout="wide")

# ==============================
# CONFIG GITHUB
# ==============================
OWNER = "irensegabriel"
REPO = "painel-faturamento"
BRANCH = "main"

BASE_URL = f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{BRANCH}/dashboard"

ARQUIVOS = {
    "notas": "notas_dashboard.csv",
    "contratos": "faturamento_contratos_dashboard.csv",
    "dias": "faturamento_dias_dashboard.csv",
    "carro": "faturamento_carro_estimado_dashboard.csv",
    "carro_dias": "faturamento_carro_dias_dashboard.csv",
}

ORDEM_DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# ==============================
# AUTO REFRESH (A CADA 60s)
# ==============================
st.markdown(
    """
    <script>
    setTimeout(function(){
        window.location.reload();
    }, 60000);
    </script>
    """,
    unsafe_allow_html=True
)

# ==============================
# LEITURA CSV (SEM CACHE PESADO)
# ==============================
@st.cache_data(ttl=60)
def ler_csv_github(nome_arquivo):
    url = f"{BASE_URL}/{nome_arquivo}?t={int(time.time())}"  # evita cache do navegador
    df = pd.read_csv(url, sep=";", encoding="utf-8-sig")

    for col in df.columns:
        if "FATURAMENTO" in col:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        if col in ["QTD_NOTAS", "QTD_EXECUTORES", "DIA_SEMANA_NUM"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df

# ==============================
# CARREGAR BASES
# ==============================
def carregar_bases():
    bases = {}
    faltando = []

    for chave, nome in ARQUIVOS.items():
        try:
            bases[chave] = ler_csv_github(nome)
        except:
            faltando.append(nome)

    return bases, faltando

# ==============================
# FORMATADORES
# ==============================
def dinheiro(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

def numero(valor):
    try:
        return f"{int(valor):,}".replace(",", ".")
    except:
        return valor

def formatar_tabela(df):
    df2 = df.copy()
    for col in df2.columns:
        if "FATURAMENTO" in col or col in ["CORTE", "RELIGUE", "TOTAL", "MÍNIMO", "MÁXIMO"]:
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["QTD_NOTAS", "NOTAS"]:
            df2[col] = df2[col].apply(numero)
    return df2

# ==============================
# APP
# ==============================
bases, faltando = carregar_bases()

st.title("📊 Painel de Faturamento")
st.caption("Atualiza automaticamente a cada 60 segundos 🚀")

if faltando:
    st.warning("Arquivos não encontrados: " + ", ".join(faltando))

if not bases:
    st.error("Nenhum CSV encontrado no GitHub.")
    st.stop()

contratos_original = bases.get("contratos", pd.DataFrame())
carro_original = bases.get("carro", pd.DataFrame())
dias_original = bases.get("dias", pd.DataFrame())
carro_dias_original = bases.get("carro_dias", pd.DataFrame())
notas = bases.get("notas", pd.DataFrame())

# ==============================
# FILTROS
# ==============================
st.sidebar.header("Filtros")

contratos_lista = []
for base in [contratos_original, dias_original, carro_original, carro_dias_original]:
    if not base.empty and "CONTRATO" in base.columns:
        contratos_lista += base["CONTRATO"].dropna().unique().tolist()

contratos_lista = sorted(set(contratos_lista))

contratos_selecionados = st.sidebar.multiselect(
    "Contratos", contratos_lista, default=contratos_lista
)

contratos = contratos_original.copy()
carro = carro_original.copy()
dias = dias_original.copy()
carro_dias = carro_dias_original.copy()

if contratos_selecionados:
    if "CONTRATO" in contratos.columns:
        contratos = contratos[contratos["CONTRATO"].isin(contratos_selecionados)]
    if "CONTRATO" in dias.columns:
        dias = dias[dias["CONTRATO"].isin(contratos_selecionados)]
    if "CONTRATO" in carro.columns:
        carro = carro[carro["CONTRATO"].isin(contratos_selecionados)]
    if "CONTRATO" in carro_dias.columns:
        carro_dias = carro_dias[carro_dias["CONTRATO"].isin(contratos_selecionados)]

mostrar_carro = not carro.empty

# ==============================
# TABS
# ==============================
aba_resumo, aba_dias, aba_carro, aba_notas = st.tabs([
    "Resumo", "Dias", "Carro", "Notas"
])

# ==============================
# RESUMO
# ==============================
with aba_resumo:
    total = contratos["FATURAMENTO"].sum() if "FATURAMENTO" in contratos.columns else 0
    st.metric("Faturamento total", dinheiro(total))

    if not contratos.empty:
        resumo = contratos.groupby("CONTRATO", as_index=False)["FATURAMENTO"].sum()
        st.bar_chart(resumo, x="CONTRATO", y="FATURAMENTO")

# ==============================
# DIAS
# ==============================
with aba_dias:
    if not dias.empty:
        por_dia = dias.groupby("DIA_SEMANA", as_index=False)["FATURAMENTO"].sum()
        st.bar_chart(por_dia, x="DIA_SEMANA", y="FATURAMENTO")

# ==============================
# CARRO
# ==============================
with aba_carro:
    if not carro.empty:
        st.dataframe(formatar_tabela(carro), use_container_width=True)

# ==============================
# NOTAS
# ==============================
with aba_notas:
    if not notas.empty:
        st.dataframe(notas.head(2000), use_container_width=True)
