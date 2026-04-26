from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Painel de Faturamento", page_icon="📊", layout="wide")

SENHA_CORRETA = "SCS@2026"

# ==============================
# IDENTIDADE VISUAL / CORES
# ==============================

st.markdown(
    """
    <style>
    .main .block-container { padding-top: 1.6rem; }

    div[data-testid="stMetric"] {
        background: rgba(125, 125, 125, 0.08);
        border: 1px solid rgba(125, 125, 125, 0.22);
        border-radius: 16px;
        padding: 16px 18px;
        box-shadow: 0 3px 12px rgba(0, 0, 0, 0.08);
    }
    div[data-testid="stMetric"] label { font-weight: 700 !important; }
    div[data-testid="stMetricValue"] { font-weight: 800 !important; }

    .executive-card {
        background: rgba(31, 120, 180, 0.14);
        color: inherit;
        border: 1px solid rgba(31, 120, 180, 0.35);
        border-left: 7px solid #1f78b4;
        border-radius: 18px;
        padding: 18px 22px;
        margin-bottom: 14px;
        box-shadow: 0 8px 22px rgba(0, 0, 0, 0.10);
    }
    .executive-card h3 { margin-bottom: 6px; }

    .ranking-podium {
        border-radius: 14px;
        padding: 10px 12px;
        margin: 4px 0 10px 0;
        border: 1px solid rgba(125, 125, 125, 0.22);
        background: rgba(125, 125, 125, 0.08);
    }
    .gold { border-left: 7px solid #d4af37; }
    .silver { border-left: 7px solid #a8a8a8; }
    .bronze { border-left: 7px solid #cd7f32; }

    .soft-note {
        border-radius: 14px;
        padding: 10px 14px;
        background: rgba(125, 125, 125, 0.07);
        border: 1px solid rgba(125, 125, 125, 0.16);
        margin: 8px 0 14px 0;
        font-size: 0.94rem;
    }
    .section-title {
        font-weight: 800;
        margin: 18px 0 6px 0;
    }

    @media (prefers-color-scheme: dark) {
        div[data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.06);
            border-color: rgba(255, 255, 255, 0.14);
        }
        .executive-card {
            background: rgba(59, 130, 246, 0.12);
            border-color: rgba(147, 197, 253, 0.35);
            border-left-color: #60a5fa;
        }
        .ranking-podium {
            background: rgba(255, 255, 255, 0.05);
            border-color: rgba(255, 255, 255, 0.13);
        }
        .soft-note {
            background: rgba(255, 255, 255, 0.045);
            border-color: rgba(255, 255, 255, 0.12);
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acesso restrito")

    senha = st.text_input("Digite a senha para acessar o painel:", type="password")

    if st.button("Entrar"):
        if senha == SENHA_CORRETA:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Senha incorreta")

    st.stop()

PASTA_DASHBOARD = Path("dashboard")
PASTA_ATUAL = Path(".")

ARQUIVOS = {
    "notas": "notas_dashboard.csv",
    "contratos": "faturamento_contratos_dashboard.csv",
    "dias": "faturamento_dias_dashboard.csv",
    "carro": "faturamento_carro_estimado_dashboard.csv",
    "carro_dias": "faturamento_carro_dias_dashboard.csv",
}

ORDEM_DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# Como dias anteriores não mudam, mantemos o carregamento geral em 15 minutos
# e deixamos cálculos históricos/ranqueamentos em cache mais longo.
CACHE_TTL_SEGUNDOS = 900
CACHE_TTL_RANKING_SEGUNDOS = 86400

# Atualiza a página automaticamente a cada 15 minutos.
st.markdown(
    """
    <script>
    setTimeout(function(){
        window.location.reload();
    }, 900000);
    </script>
    """,
    unsafe_allow_html=True
)


def caminho_arquivo(nome):
    candidatos = [
        PASTA_DASHBOARD / nome,
        PASTA_ATUAL / nome,
        PASTA_ATUAL / nome.replace(".csv", "(1).csv"),
    ]

    for caminho in candidatos:
        if caminho.exists():
            return caminho

    achados = list(PASTA_ATUAL.glob(nome.replace(".csv", "*.csv"))) + list(PASTA_DASHBOARD.glob(nome.replace(".csv", "*.csv")))
    return achados[0] if achados else None


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def ler_csv(caminho):
    df = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")

    for col in df.columns:
        if "FATURAMENTO" in col:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        if col in ["QTD_NOTAS", "QTD_EXECUTORES", "DIA_SEMANA_NUM"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


def dinheiro(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return valor


def numero(valor):
    try:
        return f"{int(valor):,}".replace(",", ".")
    except Exception:
        return valor


def formatar_tabela(df):
    df2 = df.copy()

    colunas_moeda = ORDEM_DIAS + [
        "CORTE", "RELIGUE", "TOTAL", "MÍNIMO", "MÁXIMO", "VALOR",
        "Total semana", "FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
    ]

    for col in df2.columns:
        if "FATURAMENTO" in col or col in colunas_moeda:
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["QTD_NOTAS", "NOTAS", "CORTES", "RELIGUES", "TOTAL_NOTAS"]:
            df2[col] = df2[col].apply(numero)

    return df2




def preparar_tabela_ranking(df, colunas_moeda=None):
    """Formata ranking sem depender de Styler/Jinja2, evitando erro no Streamlit Cloud."""
    if df.empty:
        return df

    df2 = df.copy()
    colunas_moeda = colunas_moeda or []

    for col in df2.columns:
        if col in colunas_moeda or "FATURAMENTO" in col or col == "TICKET_MÉDIO":
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["POSIÇÃO", "NOTAS", "CORTES", "RELIGUES", "DIAS_ATIVOS", "QTD_EQUIPES", "QTD_RECURSOS"]:
            df2[col] = df2[col].apply(numero)
        elif col in ["MÉDIA_NOTAS_DIA"]:
            df2[col] = df2[col].apply(lambda v: f"{float(v):.2f}".replace(".", ","))

    return df2


def mostrar_podio_ranking(ranking, nome_coluna="RECURSO"):
    """Mostra os três primeiros colocados com destaque compatível com tema claro/escuro."""
    if ranking.empty:
        return

    classes = ["gold", "silver", "bronze"]
    medalhas = ["🥇", "🥈", "🥉"]
    for i, (_, row) in enumerate(ranking.head(3).iterrows()):
        st.markdown(
            f"""
            <div class="ranking-podium {classes[i]}">
                <b>{medalhas[i]} {numero(row.get('POSIÇÃO', i + 1))}º — {row.get(nome_coluna, '')}</b><br>
                {numero(row.get('NOTAS', 0))} notas • {dinheiro(row.get('FATURAMENTO_ATRIBUÍDO', 0))} em faturamento
            </div>
            """,
            unsafe_allow_html=True,
        )


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def montar_base_executores(notas):
    """Monta base de ranking por RECURSO/equipe, não por código de eletricista."""
    parcial = preparar_parcial_do_dia(notas)

    if parcial.empty:
        return pd.DataFrame()

    base = parcial.copy()
    if "RECURSO" not in base.columns:
        base["RECURSO"] = ""

    base["RECURSO"] = base["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    base = base[base["RECURSO"] != ""].copy()

    if base.empty:
        return pd.DataFrame()

    base["FATURAMENTO_ATRIBUÍDO"] = pd.to_numeric(base.get("FATURAMENTO", 0), errors="coerce").fillna(0)
    base["FATURAMENTO_MIN_ATRIBUÍDO"] = pd.to_numeric(base.get("FATURAMENTO_MIN", 0), errors="coerce").fillna(0)
    base["FATURAMENTO_MAX_ATRIBUÍDO"] = pd.to_numeric(base.get("FATURAMENTO_MAX", 0), errors="coerce").fillna(0)

    base["MES"] = base["DATA_DT"].dt.strftime("%m/%Y")
    base["SEMANA_INICIO_DT"] = base["DATA_DT"] - pd.to_timedelta(base["DATA_DT"].dt.weekday, unit="D")
    base["SEMANA"] = base["SEMANA_INICIO_DT"].dt.strftime("%d/%m/%Y")
    return base


def filtrar_base_executores(base, contrato, tipo_periodo, valor_periodo):
    df = base.copy()

    if contrato != "Todos" and "CONTRATO" in df.columns:
        df = df[df["CONTRATO"] == contrato]

    if tipo_periodo == "Dia" and valor_periodo:
        df = df[df["DATA"] == valor_periodo]
    elif tipo_periodo == "Semana" and valor_periodo:
        df = df[df["SEMANA"] == valor_periodo]
    elif tipo_periodo == "Mês" and valor_periodo:
        df = df[df["MES"] == valor_periodo]

    return df


def calcular_ranking_executores(base_filtrada, criterio="Notas"):
    if base_filtrada.empty:
        return pd.DataFrame()

    ranking = (
        base_filtrada.groupby("RECURSO", dropna=False)
        .agg(
            NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            DIAS_ATIVOS=("DATA", "nunique"),
            QTD_EQUIPES=("RECURSO", "nunique"),
            FATURAMENTO_ATRIBUÍDO=("FATURAMENTO_ATRIBUÍDO", "sum"),
            FATURAMENTO_MIN_ATRIBUÍDO=("FATURAMENTO_MIN_ATRIBUÍDO", "sum"),
            FATURAMENTO_MAX_ATRIBUÍDO=("FATURAMENTO_MAX_ATRIBUÍDO", "sum"),
            FATURAMENTO_EQUIPE=("FATURAMENTO", "sum"),
        )
        .reset_index()
    )

    ranking["MÉDIA_NOTAS_DIA"] = ranking.apply(
        lambda r: (r["NOTAS"] / r["DIAS_ATIVOS"]) if r["DIAS_ATIVOS"] else 0,
        axis=1,
    )
    ranking["TICKET_MÉDIO"] = ranking.apply(
        lambda r: (r["FATURAMENTO_ATRIBUÍDO"] / r["NOTAS"]) if r["NOTAS"] else 0,
        axis=1,
    )

    coluna_ordem = "NOTAS" if criterio == "Notas" else "FATURAMENTO_ATRIBUÍDO"
    ranking = ranking.sort_values([coluna_ordem, "NOTAS"], ascending=False).reset_index(drop=True)
    ranking.insert(0, "POSIÇÃO", range(1, len(ranking) + 1))

    return ranking


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def opcoes_periodo_ranking(base):
    """Pré-calcula listas de datas/semanas/meses para o ranking."""
    if base.empty:
        return [], [], []

    dias = (
        base[["DATA", "DATA_DT"]]
        .drop_duplicates()
        .sort_values("DATA_DT", ascending=False)["DATA"]
        .tolist()
    )

    semanas = (
        base[["SEMANA", "SEMANA_INICIO_DT"]]
        .drop_duplicates()
        .sort_values("SEMANA_INICIO_DT", ascending=False)["SEMANA"]
        .tolist()
    )

    meses_df = base[["MES", "DATA_DT"]].drop_duplicates().copy()
    meses_df["PERIODO"] = pd.to_datetime(meses_df["DATA_DT"]).dt.to_period("M")
    meses = (
        meses_df[["MES", "PERIODO"]]
        .drop_duplicates()
        .sort_values("PERIODO", ascending=False)["MES"]
        .tolist()
    )

    return dias, semanas, meses


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def ranking_recursos_cacheado(base, contrato, tipo_periodo, valor_periodo, criterio):
    """Filtra e calcula o ranking em cache para acelerar trocas de filtro."""
    base_filtrada = filtrar_base_executores(base, contrato, tipo_periodo, valor_periodo)
    ranking = calcular_ranking_executores(base_filtrada, criterio)
    return base_filtrada, ranking


def carregar_bases():
    bases = {}
    faltando = []

    for chave, nome in ARQUIVOS.items():
        caminho = caminho_arquivo(nome)

        if caminho:
            bases[chave] = ler_csv(str(caminho))
        else:
            faltando.append(nome)

    return bases, faltando


# ==============================
# REGRAS DE CONTRATO / FATURAMENTO
# Usadas na aba "Parcial do dia"
# ==============================

def eh_disjuntor_jundiai(recurso):
    recurso_norm = str(recurso).strip().upper()
    return recurso_norm.startswith("JUN55") or recurso_norm.startswith("JUN59") or recurso_norm.startswith("SAL55")


def eh_disjuntor_santa_cruz(recurso):
    import re
    recurso_norm = str(recurso).strip().upper()
    m = re.search(r"(\d+)", recurso_norm)
    if not m:
        return False
    primeiros_numeros = m.group(1)
    return primeiros_numeros.startswith("89") or primeiros_numeros.startswith("20")


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def preparar_parcial_do_dia(notas):
    if notas.empty:
        return pd.DataFrame()

    df = notas.copy()

    for col in ["ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "RECUSA", "ELETRICISTA1", "ELETRICISTA2", "DATA"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "QTD_EXECUTORES" not in df.columns:
        df["QTD_EXECUTORES"] = ((df["ELETRICISTA1"] != "").astype(int) + (df["ELETRICISTA2"] != "").astype(int))
    else:
        df["QTD_EXECUTORES"] = pd.to_numeric(df["QTD_EXECUTORES"], errors="coerce").fillna(0).astype(int)

    df["GRUPO_NOTA"] = df["GRUPO_NOTA"].str.upper()
    df["RECURSO"] = df["RECURSO"].str.upper()
    df["RECUSA"] = df["RECUSA"].fillna("").astype(str).str.strip()

    # Parcial considera apenas notas pagáveis, ou seja, sem recusa.
    df = df[df["RECUSA"] == ""].copy()

    linhas = []

    for _, row in df.iterrows():
        recurso = row.get("RECURSO", "")
        grupo = row.get("GRUPO_NOTA", "")
        qtd_exec = int(row.get("QTD_EXECUTORES", 0) or 0)

        contrato = ""
        faturamento = 0.0
        faturamento_min = 0.0
        faturamento_max = 0.0

        if eh_disjuntor_jundiai(recurso):
            contrato = "Disjuntor Jundiaí"
            faturamento = {"CORTE": 13.72, "RELIGUE": 27.43}.get(grupo, 0.0)
            faturamento_min = faturamento
            faturamento_max = faturamento

        elif eh_disjuntor_santa_cruz(recurso):
            contrato = "Disjuntor Santa Cruz"
            faturamento = {"CORTE": 11.98, "RELIGUE": 23.97}.get(grupo, 0.0)
            faturamento_min = faturamento
            faturamento_max = faturamento

        elif str(recurso).startswith("JUN58") and qtd_exec >= 2:
            contrato = "Contrato Carro STC estimado"
            faturamento_min = {"CORTE": 38.18, "RELIGUE": 36.36}.get(grupo, 0.0)
            faturamento_max = {"CORTE": 45.45, "RELIGUE": 50.91}.get(grupo, 0.0)
            faturamento = faturamento_min

        if contrato:
            item = row.to_dict()
            item["CONTRATO"] = contrato
            item["FATURAMENTO"] = faturamento
            item["FATURAMENTO_MIN"] = faturamento_min
            item["FATURAMENTO_MAX"] = faturamento_max
            item["EH_CORTE"] = 1 if grupo == "CORTE" else 0
            item["EH_RELIGUE"] = 1 if grupo == "RELIGUE" else 0
            linhas.append(item)

    if not linhas:
        return pd.DataFrame()

    parcial = pd.DataFrame(linhas)
    parcial["DATA_DT"] = pd.to_datetime(parcial["DATA"], dayfirst=True, errors="coerce")
    parcial = parcial.dropna(subset=["DATA_DT"])
    parcial["DATA"] = parcial["DATA_DT"].dt.strftime("%d/%m/%Y")

    return parcial




def meses_disponiveis_da_base(notas):
    """Retorna os meses disponíveis na base acumulada, do mais recente para o mais antigo."""
    if notas.empty:
        return pd.DataFrame(columns=["MES", "PERIODO"])

    df = notas.copy()
    coluna_data = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in df.columns else "DATA"

    if coluna_data not in df.columns:
        return pd.DataFrame(columns=["MES", "PERIODO"])

    datas = pd.to_datetime(df[coluna_data], dayfirst=True, errors="coerce")
    periodos = datas.dt.to_period("M")

    meses = (
        pd.DataFrame({"PERIODO": periodos})
        .dropna()
        .drop_duplicates()
        .sort_values("PERIODO", ascending=False)
    )

    if meses.empty:
        return pd.DataFrame(columns=["MES", "PERIODO"])

    meses["MES"] = meses["PERIODO"].dt.strftime("%m/%Y")
    return meses[["MES", "PERIODO"]].reset_index(drop=True)



def data_maxima_do_mes(notas, mes):
    """Retorna a maior data encontrada dentro de um mês no formato MM/AAAA."""
    if notas.empty:
        return None

    df = notas.copy()
    coluna_data = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in df.columns else "DATA"

    if coluna_data not in df.columns:
        return None

    datas = pd.to_datetime(df[coluna_data], dayfirst=True, errors="coerce")
    periodo_mes = pd.Period(f"{mes[3:7]}-{mes[0:2]}", freq="M")
    datas_mes = datas[datas.dt.to_period("M") == periodo_mes]

    if datas_mes.dropna().empty:
        return None

    return datas_mes.max()


def texto_mes_com_parcial(notas, mes):
    """Mostra o mês e, se ainda não fechou, informa até qual data há dados."""
    data_max = data_maxima_do_mes(notas, mes)

    if data_max is None:
        return mes

    periodo_mes = data_max.to_period("M")
    ultimo_dia_mes = periodo_mes.end_time.normalize()

    if data_max.normalize() < ultimo_dia_mes:
        return f"{mes} (parcial até {data_max.strftime('%d/%m/%Y')})"

    return f"{mes} (mês fechado)"

@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def resumo_por_periodo(notas, meses_escolhidos, contrato_escolhido="Todos"):
    """Monta resumo financeiro por contrato e por grupo para os meses escolhidos."""
    parcial = preparar_parcial_do_dia(notas)

    if parcial.empty:
        return pd.DataFrame(), pd.DataFrame()

    if not meses_escolhidos:
        meses_base = meses_disponiveis_da_base(notas)
        if meses_base.empty:
            return pd.DataFrame(), pd.DataFrame()
        meses_escolhidos = [meses_base.iloc[0]["MES"]]

    parcial["MES"] = parcial["DATA_DT"].dt.strftime("%m/%Y")
    parcial = parcial[parcial["MES"].isin(meses_escolhidos)].copy()

    if contrato_escolhido != "Todos" and "CONTRATO" in parcial.columns:
        parcial = parcial[parcial["CONTRATO"] == contrato_escolhido]

    if parcial.empty:
        return pd.DataFrame(), pd.DataFrame()

    resumo_contrato = (
        parcial.groupby("CONTRATO", dropna=False)
        .agg(
            TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            FATURAMENTO=("FATURAMENTO", "sum"),
            FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
            FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
        )
        .reset_index()
        .sort_values("FATURAMENTO", ascending=False)
    )

    resumo_grupo = (
        parcial.groupby(["CONTRATO", "GRUPO_NOTA"], dropna=False)
        .agg(
            TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            FATURAMENTO=("FATURAMENTO", "sum"),
            FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
            FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
        )
        .reset_index()
    )

    return resumo_contrato, resumo_grupo


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def calcular_resumo_mensal(notas, mes, contrato_escolhido="Todos"):
    resumo_contrato, _ = resumo_por_periodo(notas, [mes], contrato_escolhido)

    if resumo_contrato.empty:
        return {
            "FATURAMENTO": 0.0,
            "TOTAL_NOTAS": 0,
            "CORTES": 0,
            "RELIGUES": 0,
            "FATURAMENTO_MIN": 0.0,
            "FATURAMENTO_MAX": 0.0,
        }

    return {
        "FATURAMENTO": float(resumo_contrato["FATURAMENTO"].sum()),
        "TOTAL_NOTAS": int(resumo_contrato["TOTAL_NOTAS"].sum()),
        "CORTES": int(resumo_contrato["CORTES"].sum()),
        "RELIGUES": int(resumo_contrato["RELIGUES"].sum()),
        "FATURAMENTO_MIN": float(resumo_contrato["FATURAMENTO_MIN"].sum()),
        "FATURAMENTO_MAX": float(resumo_contrato["FATURAMENTO_MAX"].sum()),
    }


def variacao_percentual(atual, anterior):
    if anterior == 0:
        if atual == 0:
            return "0,0%"
        return "novo"
    valor = ((atual - anterior) / anterior) * 100
    sinal = "+" if valor >= 0 else ""
    return f"{sinal}{valor:.1f}%".replace(".", ",")


# ==============================
# CARREGAMENTO
# ==============================

bases, faltando = carregar_bases()

st.title("📊 Painel de Faturamento")
st.caption("Painel atualizado automaticamente a cada 15 minutos. Cálculos em cache para trocar filtros mais rápido.")

if faltando:
    st.warning("Arquivos não encontrados: " + ", ".join(faltando))

if not bases:
    st.error("Nenhum CSV foi encontrado. Verifique se os arquivos estão na pasta dashboard.")
    st.stop()

contratos_original = bases.get("contratos", pd.DataFrame())
carro_original = bases.get("carro", pd.DataFrame())
dias_original = bases.get("dias", pd.DataFrame())
carro_dias_original = bases.get("carro_dias", pd.DataFrame())
notas = bases.get("notas", pd.DataFrame())

# ==============================
# FILTROS EM BOTÕES
# ==============================

st.sidebar.header("Filtros")

contratos_lista = []

for base in [contratos_original, dias_original, carro_original, carro_dias_original]:
    if not base.empty and "CONTRATO" in base.columns:
        contratos_lista += base["CONTRATO"].dropna().unique().tolist()

contratos_lista = sorted(set(contratos_lista))

if "contrato_escolhido" not in st.session_state:
    st.session_state.contrato_escolhido = "Todos"

st.sidebar.markdown("### Contratos")

if st.sidebar.button("📊 Todos", use_container_width=True):
    st.session_state.contrato_escolhido = "Todos"

for contrato_nome in contratos_lista:
    if st.sidebar.button(f"🔹 {contrato_nome}", use_container_width=True):
        st.session_state.contrato_escolhido = contrato_nome

contrato_escolhido = st.session_state.contrato_escolhido

st.sidebar.markdown("---")
st.sidebar.markdown("**Contrato selecionado:**")
st.sidebar.info(contrato_escolhido)

# Este período vale para a tela inicial "Resumo".
# Por padrão, fica só no mês mais recente da base, para não somar março + abril sem querer.
meses_base = meses_disponiveis_da_base(notas)
meses_escolhidos_resumo = []

if not meses_base.empty:
    opcoes_meses_resumo = meses_base["MES"].tolist()
    mes_mais_recente = opcoes_meses_resumo[0]

    if "meses_resumo" not in st.session_state:
        st.session_state.meses_resumo = [mes_mais_recente]

    st.sidebar.markdown("### Período do resumo")

    if st.sidebar.button("📅 Usar mês mais recente", use_container_width=True):
        st.session_state.meses_resumo = [mes_mais_recente]

    if st.sidebar.button("🧮 Somar todos os meses", use_container_width=True):
        st.session_state.meses_resumo = opcoes_meses_resumo.copy()

    meses_escolhidos_resumo = st.sidebar.multiselect(
        "Meses que entram na tela inicial",
        opcoes_meses_resumo,
        default=st.session_state.meses_resumo,
    )

    if not meses_escolhidos_resumo:
        meses_escolhidos_resumo = [mes_mais_recente]

    st.session_state.meses_resumo = meses_escolhidos_resumo

contratos = contratos_original.copy()
carro = carro_original.copy()
dias = dias_original.copy()
carro_dias = carro_dias_original.copy()

if contrato_escolhido != "Todos":
    if not contratos.empty and "CONTRATO" in contratos.columns:
        contratos = contratos[contratos["CONTRATO"] == contrato_escolhido]

    if not dias.empty and "CONTRATO" in dias.columns:
        dias = dias[dias["CONTRATO"] == contrato_escolhido]

    if not carro.empty and "CONTRATO" in carro.columns:
        carro = carro[carro["CONTRATO"] == contrato_escolhido]

    if not carro_dias.empty and "CONTRATO" in carro_dias.columns:
        carro_dias = carro_dias[carro_dias["CONTRATO"] == contrato_escolhido]

mostrar_carro = not carro.empty

mostrar_aba_carro = contrato_escolhido in ["Todos", "Contrato Carro STC estimado"]

nomes_abas = ["Resumo", "Parcial do dia", "Ranking de recursos", "Comparativo mensal", "Dias da semana"]
if mostrar_aba_carro:
    nomes_abas.append("Carro estimado")
nomes_abas += ["Notas", "Downloads"]

abas = st.tabs(nomes_abas)
aba_resumo = abas[0]
aba_parcial = abas[1]
aba_ranking = abas[2]
aba_comparativo = abas[3]
aba_dias = abas[4]

if mostrar_aba_carro:
    aba_carro = abas[5]
    aba_notas = abas[6]
    aba_download = abas[7]
else:
    aba_notas = abas[5]
    aba_download = abas[6]

# ==============================
# ABA RESUMO
# ==============================

with aba_resumo:
    resumo_contrato_periodo, resumo_grupo_periodo = resumo_por_periodo(
        notas,
        meses_escolhidos_resumo,
        contrato_escolhido,
    )

    periodo_texto = ", ".join(meses_escolhidos_resumo) if meses_escolhidos_resumo else "mês mais recente"
    st.caption(f"Resumo considerando: {periodo_texto}")

    if resumo_contrato_periodo.empty:
        st.info("Não há dados para o período selecionado.")
    else:
        total_contratos = resumo_contrato_periodo["FATURAMENTO"].sum()
        qtd_notas = int(resumo_contrato_periodo["TOTAL_NOTAS"].sum())

        carro_periodo = resumo_contrato_periodo[
            resumo_contrato_periodo["CONTRATO"] == "Contrato Carro STC estimado"
        ].copy()

        mostrar_carro_periodo = not carro_periodo.empty

        if mostrar_carro_periodo:
            total_carro_min = carro_periodo["FATURAMENTO_MIN"].sum()
            total_carro_max = carro_periodo["FATURAMENTO_MAX"].sum()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Faturamento contratos", dinheiro(total_contratos))
            c2.metric("Carro estimado mínimo", dinheiro(total_carro_min))
            c3.metric("Carro estimado máximo", dinheiro(total_carro_max))
            c4.metric("Notas únicas", numero(qtd_notas))
        else:
            c1, c2 = st.columns(2)
            c1.metric("Faturamento contratos", dinheiro(total_contratos))
            c2.metric("Notas únicas", numero(qtd_notas))

        st.subheader("Faturamento por contrato")

        grafico_resumo = resumo_contrato_periodo.copy()
        st.bar_chart(grafico_resumo, x="CONTRATO", y="FATURAMENTO")

        st.markdown("**Resumo com corte + religue**")

        if not resumo_grupo_periodo.empty:
            tabela_resumo = resumo_grupo_periodo.pivot_table(
                index="CONTRATO",
                columns="GRUPO_NOTA",
                values="FATURAMENTO",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            for col in ["CORTE", "RELIGUE"]:
                if col not in tabela_resumo.columns:
                    tabela_resumo[col] = 0

            tabela_resumo["TOTAL"] = tabela_resumo[["CORTE", "RELIGUE"]].sum(axis=1)
            tabela_resumo = tabela_resumo[["CONTRATO", "CORTE", "RELIGUE", "TOTAL"]]

            st.dataframe(formatar_tabela(tabela_resumo), use_container_width=True, hide_index=True)

        st.markdown("**Detalhamento por contrato no período**")
        colunas_detalhe_resumo = [
            "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES",
            "FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
        ]
        colunas_detalhe_resumo = [c for c in colunas_detalhe_resumo if c in resumo_contrato_periodo.columns]
        st.dataframe(
            formatar_tabela(resumo_contrato_periodo[colunas_detalhe_resumo]),
            use_container_width=True,
            hide_index=True,
        )

# ==============================
# ABA PARCIAL DO DIA
# ==============================

with aba_parcial:
    st.subheader("Parcial do dia por equipe")

    parcial = preparar_parcial_do_dia(notas)

    if parcial.empty:
        st.info("Ainda não há dados suficientes para montar a parcial do dia.")
    else:
        if contrato_escolhido != "Todos" and "CONTRATO" in parcial.columns:
            parcial = parcial[parcial["CONTRATO"] == contrato_escolhido]

        datas_disponiveis = (
            parcial[["DATA", "DATA_DT"]]
            .drop_duplicates()
            .sort_values("DATA_DT", ascending=False)
        )

        if datas_disponiveis.empty:
            st.info("Nenhuma data encontrada na base de notas.")
        else:
            opcoes_datas = datas_disponiveis["DATA"].tolist()
            data_escolhida = st.selectbox("Escolha o dia", opcoes_datas, index=0)

            parcial_dia = parcial[parcial["DATA"] == data_escolhida].copy()

            if parcial_dia.empty:
                st.info("Nenhuma nota encontrada para esse dia.")
            else:
                total_notas = parcial_dia["ORDEM_DE_SERVICO"].nunique()
                total_cortes = int(parcial_dia["EH_CORTE"].sum())
                total_religues = int(parcial_dia["EH_RELIGUE"].sum())
                total_faturamento = parcial_dia["FATURAMENTO"].sum()
                total_faturamento_min = parcial_dia["FATURAMENTO_MIN"].sum()
                total_faturamento_max = parcial_dia["FATURAMENTO_MAX"].sum()

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Notas no dia", numero(total_notas))
                c2.metric("Cortes", numero(total_cortes))
                c3.metric("Religues", numero(total_religues))

                tem_carro_no_dia = "CONTRATO" in parcial_dia.columns and (
                    parcial_dia["CONTRATO"] == "Contrato Carro STC estimado"
                ).any()

                if tem_carro_no_dia:
                    c4.metric("Faturamento estimado", f"{dinheiro(total_faturamento_min)} a {dinheiro(total_faturamento_max)}")
                else:
                    c4.metric("Faturamento", dinheiro(total_faturamento))

                st.markdown("**Resumo por equipe**")

                resumo_equipe = (
                    parcial_dia.groupby(["RECURSO", "CONTRATO"], dropna=False)
                    .agg(
                        TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
                        CORTES=("EH_CORTE", "sum"),
                        RELIGUES=("EH_RELIGUE", "sum"),
                        FATURAMENTO=("FATURAMENTO", "sum"),
                        FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
                        FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
                    )
                    .reset_index()
                    .sort_values(["CONTRATO", "RECURSO"])
                )

                def faturamento_linha_equipe(row):
                    if row.get("CONTRATO") == "Contrato Carro STC estimado":
                        return f"{dinheiro(row.get('FATURAMENTO_MIN', 0))} a {dinheiro(row.get('FATURAMENTO_MAX', 0))}"
                    return dinheiro(row.get("FATURAMENTO", 0))

                tabela_equipe = resumo_equipe.copy()
                tabela_equipe["FATURAMENTO"] = tabela_equipe.apply(faturamento_linha_equipe, axis=1)
                tabela_equipe = tabela_equipe[[
                    "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO"
                ]]

                st.dataframe(formatar_tabela(tabela_equipe), use_container_width=True, hide_index=True)

                st.markdown("**Detalhamento das notas do dia**")
                colunas_detalhe = [
                    "ORDEM_DE_SERVICO", "RECURSO", "CONTRATO", "GRUPO_NOTA", "DATA", "ELETRICISTA1", "ELETRICISTA2"
                ]
                colunas_detalhe = [c for c in colunas_detalhe if c in parcial_dia.columns]
                st.dataframe(parcial_dia[colunas_detalhe], use_container_width=True, hide_index=True)



# ==============================
# ABA RANKING DE RECURSOS
# ==============================

with aba_ranking:
    st.subheader("🏆 Ranking de recursos")
    st.caption("Ranking por RECURSO/equipe, usando o código operacional da equipe, como SAL5539-EMP.")
    st.markdown(
        '<div class="soft-note">⚡ Otimizado com cache: dias anteriores ficam reaproveitados, então alternar filtros tende a ficar mais rápido após o primeiro carregamento.</div>',
        unsafe_allow_html=True,
    )

    base_exec = montar_base_executores(notas)

    if base_exec.empty:
        st.info("Ainda não há dados suficientes de eletricistas/executores para montar o ranking.")
    else:
        col_f1, col_f2, col_f3, col_f4 = st.columns([1.2, 1.1, 1.2, 1.1])

        dias_ranking, semanas_ranking, meses_ranking = opcoes_periodo_ranking(base_exec)
        contratos_exec = ["Todos"] + sorted(base_exec["CONTRATO"].dropna().unique().tolist())
        contrato_ranking = col_f1.selectbox(
            "Contrato",
            contratos_exec,
            index=contratos_exec.index(contrato_escolhido) if contrato_escolhido in contratos_exec else 0,
            key="ranking_contrato",
        )

        tipo_periodo = col_f2.selectbox(
            "Período",
            ["Total", "Dia", "Semana", "Mês"],
            index=3,
            key="ranking_tipo_periodo",
        )

        valor_periodo = None
        if tipo_periodo == "Dia":
            valor_periodo = col_f3.selectbox("Dia", dias_ranking, key="ranking_dia")
        elif tipo_periodo == "Semana":
            valor_periodo = col_f3.selectbox("Semana iniciada em", semanas_ranking, key="ranking_semana")
        elif tipo_periodo == "Mês":
            valor_periodo = col_f3.selectbox("Mês", meses_ranking, key="ranking_mes")
        else:
            col_f3.info("Considerando toda a base")

        criterio = col_f4.selectbox("Ordenar por", ["Notas", "Faturamento"], key="ranking_criterio")

        base_filtrada_exec, ranking_exec = ranking_recursos_cacheado(
            base_exec, contrato_ranking, tipo_periodo, valor_periodo, criterio
        )

        if ranking_exec.empty:
            st.info("Nenhum recurso encontrado para os filtros selecionados.")
        else:
            total_notas_exec = int(base_filtrada_exec["ORDEM_DE_SERVICO"].nunique())
            total_executores = int(ranking_exec["RECURSO"].nunique())
            total_fat_atribuido = float(ranking_exec["FATURAMENTO_ATRIBUÍDO"].sum())
            media_notas_executor = total_notas_exec / total_executores if total_executores else 0

            lider = ranking_exec.iloc[0]

            st.markdown(
                f"""
                <div class="executive-card">
                    <h3>Resumo executivo do ranking</h3>
                    <div>🥇 Líder: <b>{lider['RECURSO']}</b> • {numero(lider['NOTAS'])} notas • {dinheiro(lider['FATURAMENTO_ATRIBUÍDO'])} em faturamento atribuído</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Recursos ativos", numero(total_executores))
            m2.metric("Notas únicas", numero(total_notas_exec))
            m3.metric("Faturamento atribuído", dinheiro(total_fat_atribuido))
            m4.metric("Média notas/recurso", f"{media_notas_executor:.1f}".replace(".", ","))

            st.markdown('<div class="section-title">Top 10 recursos</div>', unsafe_allow_html=True)
            top10 = ranking_exec.head(10).copy()
            st.bar_chart(top10, x="RECURSO", y="NOTAS" if criterio == "Notas" else "FATURAMENTO_ATRIBUÍDO")

            st.markdown('<div class="section-title">Pódio</div>', unsafe_allow_html=True)
            mostrar_podio_ranking(ranking_exec, nome_coluna="RECURSO")

            st.markdown('<div class="section-title">Ranking detalhado</div>', unsafe_allow_html=True)
            colunas_ranking = [
                "POSIÇÃO", "RECURSO", "NOTAS", "CORTES", "RELIGUES", "DIAS_ATIVOS",
                "MÉDIA_NOTAS_DIA", "TICKET_MÉDIO", "FATURAMENTO_ATRIBUÍDO",
                "FATURAMENTO_MIN_ATRIBUÍDO", "FATURAMENTO_MAX_ATRIBUÍDO", "FATURAMENTO_EQUIPE", "QTD_EQUIPES"
            ]
            colunas_ranking = [c for c in colunas_ranking if c in ranking_exec.columns]
            st.dataframe(
                preparar_tabela_ranking(ranking_exec[colunas_ranking]),
                use_container_width=True,
                hide_index=True,
            )

            with st.expander("Ver notas consideradas no ranking"):
                detalhe_cols = [
                    "DATA", "RECURSO", "CONTRATO", "ORDEM_DE_SERVICO",
                    "GRUPO_NOTA", "FATURAMENTO", "FATURAMENTO_ATRIBUÍDO"
                ]
                detalhe_cols = [c for c in detalhe_cols if c in base_filtrada_exec.columns]
                detalhe = base_filtrada_exec[detalhe_cols].sort_values(["DATA", "RECURSO"], ascending=[False, True])
                st.dataframe(
                    preparar_tabela_ranking(detalhe, colunas_moeda=["FATURAMENTO", "FATURAMENTO_ATRIBUÍDO"]),
                    use_container_width=True,
                    hide_index=True,
                )

            csv_ranking = ranking_exec.to_csv(index=False, sep=";", encoding="utf-8-sig")
            st.download_button(
                "Baixar ranking de recursos em CSV",
                csv_ranking,
                file_name="ranking_recursos.csv",
                mime="text/csv",
                use_container_width=True,
            )

# ==============================
# ABA COMPARATIVO MENSAL
# ==============================

with aba_comparativo:
    st.subheader("Comparativo mensal")
    st.caption("Compara o mês escolhido com o mês anterior, usando a base acumulada de notas.")

    meses_base_comp = meses_disponiveis_da_base(notas)

    if meses_base_comp.empty:
        st.info("Ainda não encontrei meses disponíveis na base de notas.")
    else:
        opcoes_meses = meses_base_comp["MES"].tolist()
        mes_escolhido = st.selectbox("Escolha o mês para comparar", opcoes_meses, index=0)

        periodo_escolhido = meses_base_comp.loc[meses_base_comp["MES"] == mes_escolhido, "PERIODO"].iloc[0]
        periodo_anterior = periodo_escolhido - 1
        mes_anterior = periodo_anterior.strftime("%m/%Y")

        atual = calcular_resumo_mensal(notas, mes_escolhido, contrato_escolhido)
        anterior = calcular_resumo_mensal(notas, mes_anterior, contrato_escolhido)

        st.markdown(f"**Comparando: {mes_escolhido} x {mes_anterior}**")

        data_max_comp = data_maxima_do_mes(notas, mes_escolhido)
        if data_max_comp is not None:
            ultimo_dia_comp = data_max_comp.to_period("M").end_time.normalize()
            if data_max_comp.normalize() < ultimo_dia_comp:
                st.info(
                    f"Atenção: {mes_escolhido} ainda é parcial. "
                    f"Os dados vão até {data_max_comp.strftime('%d/%m/%Y')}."
                )

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Faturamento", dinheiro(atual["FATURAMENTO"]), variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"]))
        c2.metric("Notas", numero(atual["TOTAL_NOTAS"]), variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"]))
        c3.metric("Cortes", numero(atual["CORTES"]), variacao_percentual(atual["CORTES"], anterior["CORTES"]))
        c4.metric("Religues", numero(atual["RELIGUES"]), variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"]))

        tabela_comparativo = pd.DataFrame([
            {"Indicador": "Faturamento", mes_escolhido: dinheiro(atual["FATURAMENTO"]), mes_anterior: dinheiro(anterior["FATURAMENTO"]), "Variação": variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"])},
            {"Indicador": "Notas", mes_escolhido: numero(atual["TOTAL_NOTAS"]), mes_anterior: numero(anterior["TOTAL_NOTAS"]), "Variação": variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"])},
            {"Indicador": "Cortes", mes_escolhido: numero(atual["CORTES"]), mes_anterior: numero(anterior["CORTES"]), "Variação": variacao_percentual(atual["CORTES"], anterior["CORTES"])},
            {"Indicador": "Religues", mes_escolhido: numero(atual["RELIGUES"]), mes_anterior: numero(anterior["RELIGUES"]), "Variação": variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"])},
        ])
        st.dataframe(tabela_comparativo, use_container_width=True, hide_index=True)

        st.markdown("**Evolução mês a mês**")
        linhas_evolucao = []
        for mes in reversed(opcoes_meses):
            r = calcular_resumo_mensal(notas, mes, contrato_escolhido)
            linhas_evolucao.append({
                "MES": mes,
                "FATURAMENTO": r["FATURAMENTO"],
                "NOTAS": r["TOTAL_NOTAS"],
                "CORTES": r["CORTES"],
                "RELIGUES": r["RELIGUES"],
            })
        evolucao = pd.DataFrame(linhas_evolucao)

        if not evolucao.empty:
            st.bar_chart(evolucao, x="MES", y="FATURAMENTO")
            st.dataframe(formatar_tabela(evolucao), use_container_width=True, hide_index=True)

        st.markdown("**Resumo por contrato no mês escolhido**")
        resumo_mes, _ = resumo_por_periodo(notas, [mes_escolhido], contrato_escolhido)
        if resumo_mes.empty:
            st.info("Nenhum dado encontrado para esse mês.")
        else:
            colunas = ["CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]
            colunas = [c for c in colunas if c in resumo_mes.columns]
            st.dataframe(formatar_tabela(resumo_mes[colunas]), use_container_width=True, hide_index=True)

# ==============================
# ABA DIAS
# ==============================

with aba_dias:
    st.subheader("Faturamento por dia da semana")

    if not dias.empty:
        tabela = dias.pivot_table(
            index=["CONTRATO", "SEMANA_INICIO"],
            columns="DIA_SEMANA",
            values="FATURAMENTO",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        colunas_dias = [c for c in ORDEM_DIAS if c in tabela.columns]

        tabela["Total semana"] = tabela[colunas_dias].sum(axis=1)

        # Ordena a tabela pela data real do início da semana.
        # Sem isso, datas como 06/04 podem aparecer antes de 09/03,
        # porque o Streamlit/Pandas pode tratar o campo como texto.
        tabela["SEMANA_INICIO_DT"] = pd.to_datetime(
            tabela["SEMANA_INICIO"],
            dayfirst=True,
            errors="coerce",
        )
        tabela = tabela.sort_values(["SEMANA_INICIO_DT", "CONTRATO"])

        tabela = tabela[["CONTRATO", "SEMANA_INICIO"] + colunas_dias + ["Total semana"]]

        st.dataframe(formatar_tabela(tabela), use_container_width=True, hide_index=True)

        por_dia = dias.groupby("DIA_SEMANA", as_index=False)["FATURAMENTO"].sum()
        por_dia["ordem"] = por_dia["DIA_SEMANA"].map({d: i for i, d in enumerate(ORDEM_DIAS)})
        por_dia = por_dia.sort_values("ordem")

        st.bar_chart(por_dia, x="DIA_SEMANA", y="FATURAMENTO")
    else:
        st.info("Nenhum dado para o contrato selecionado.")

# ==============================
# ABA CARRO
# ==============================

if mostrar_aba_carro:
    with aba_carro:
        st.subheader("Contrato do carro — estimativa")

        if not carro.empty:
            c1, c2 = st.columns(2)
            c1.metric("Mínimo estimado", dinheiro(carro["FATURAMENTO_MIN"].sum()))
            c2.metric("Máximo estimado", dinheiro(carro["FATURAMENTO_MAX"].sum()))

            st.dataframe(formatar_tabela(carro), use_container_width=True, hide_index=True)
        else:
            st.info("Selecione o contrato do carro no menu lateral para visualizar.")

        st.subheader("Carro estimado por dia")

        if not carro_dias.empty:
            tabela_carro = carro_dias.pivot_table(
                index=["CONTRATO", "SEMANA_INICIO"],
                columns="DIA_SEMANA",
                values=["FATURAMENTO_MIN", "FATURAMENTO_MAX"],
                aggfunc="sum",
                fill_value=0,
            )

            st.dataframe(tabela_carro.style.format(dinheiro), use_container_width=True)
        else:
            st.info("Nenhum dado diário do carro para o contrato selecionado.")

# ==============================
# ABA NOTAS
# ==============================

with aba_notas:
    st.subheader("Consulta de notas")

    if not notas.empty:
        df_notas = notas.copy()

        # A base de notas acumulada não tem contrato salvo. Por isso, para filtrar por contrato,
        # reaproveitamos a classificação feita na parcial.
        parcial_para_filtro = preparar_parcial_do_dia(notas)
        if contrato_escolhido != "Todos" and not parcial_para_filtro.empty:
            ordens_do_contrato = parcial_para_filtro.loc[
                parcial_para_filtro["CONTRATO"] == contrato_escolhido,
                "ORDEM_DE_SERVICO"
            ].astype(str).unique().tolist()
            if "ORDEM_DE_SERVICO" in df_notas.columns:
                df_notas["ORDEM_DE_SERVICO"] = df_notas["ORDEM_DE_SERVICO"].astype(str)
                df_notas = df_notas[df_notas["ORDEM_DE_SERVICO"].isin(ordens_do_contrato)]

        grupo = st.selectbox(
            "Grupo de nota",
            ["Todos"] + sorted(df_notas.get("GRUPO_NOTA", pd.Series(dtype=str)).dropna().unique().tolist())
        )

        qtd_exec = st.selectbox("Quantidade de executores", ["Todos", 1, 2])

        if grupo != "Todos" and "GRUPO_NOTA" in df_notas.columns:
            df_notas = df_notas[df_notas["GRUPO_NOTA"] == grupo]

        if qtd_exec != "Todos" and "QTD_EXECUTORES" in df_notas.columns:
            df_notas = df_notas[df_notas["QTD_EXECUTORES"] == qtd_exec]

        total_filtrado = len(df_notas)
        st.info(f"{numero(total_filtrado)} notas encontradas com os filtros atuais.")

        if st.button("Carregar notas", use_container_width=True):
            st.dataframe(df_notas.head(2000), use_container_width=True, hide_index=True)
            st.caption("Mostrando até 2000 linhas para não deixar o painel pesado.")
        else:
            st.caption("As notas não são carregadas automaticamente para deixar o painel mais rápido.")
    else:
        st.info("Base de notas não encontrada.")

# ==============================
# ABA DOWNLOAD
# ==============================

with aba_download:
    st.subheader("Arquivos carregados")

    for chave, nome in ARQUIVOS.items():
        caminho = caminho_arquivo(nome)

        if caminho:
            with open(caminho, "rb") as f:
                st.download_button(f"Baixar {caminho.name}", f, file_name=caminho.name)
