from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import json
import tempfile
import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(page_title="Painel de Faturamento", page_icon="📊", layout="wide")

SENHA_CORRETA = "SCS@2026"

# ==============================
# IDENTIDADE VISUAL / CORES
# ==============================

st.markdown(
    """
    <style>
    /* Layout geral */
    .stApp {
        background: linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
    }

    .main .block-container {
        padding-top: 1.4rem;
        padding-bottom: 3rem;
        max-width: 1280px;
    }

    h1, h2, h3 {
        letter-spacing: -0.03em;
        color: #0f172a;
    }

    .stCaption, div[data-testid="stCaptionContainer"] {
        color: #64748b;
    }

    /* Métricas */
    div[data-testid="stMetric"] {
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(15, 23, 42, 0.08);
        border-radius: 22px;
        padding: 20px 22px;
        box-shadow: 0 14px 35px rgba(15, 23, 42, 0.08);
    }

    div[data-testid="stMetric"] label {
        color: #64748b !important;
        font-weight: 800 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-weight: 900 !important;
        letter-spacing: -0.04em;
    }

    /* Cards */
    .executive-card {
        background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
        color: white;
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 24px;
        padding: 24px 28px;
        margin: 12px 0 20px 0;
        box-shadow: 0 18px 45px rgba(29, 78, 216, 0.25);
    }

    .executive-card h3 {
        margin: 0 0 8px 0;
        color: white;
    }

    .ranking-podium {
        border-radius: 18px;
        padding: 14px 16px;
        margin: 6px 0 12px 0;
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(15, 23, 42, 0.08);
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.07);
    }

    .gold { border-left: 8px solid #f59e0b; }
    .silver { border-left: 8px solid #94a3b8; }
    .bronze { border-left: 8px solid #b45309; }

    .soft-note {
        border-radius: 18px;
        padding: 13px 16px;
        background: rgba(37, 99, 235, 0.08);
        border: 1px solid rgba(37, 99, 235, 0.16);
        margin: 10px 0 18px 0;
        color: #1e3a8a;
        font-size: 0.95rem;
    }

    .status-card {
        background: rgba(255, 255, 255, 0.88);
        border: 1px solid rgba(15, 23, 42, 0.08);
        border-radius: 22px;
        padding: 16px 18px;
        margin: 10px 0 20px 0;
        box-shadow: 0 14px 35px rgba(15, 23, 42, 0.08);
    }

    .status-card b {
        color: #0f172a;
    }

    .zero-card {
        border-radius: 18px;
        padding: 14px 16px;
        background: rgba(245, 158, 11, 0.12);
        border: 1px solid rgba(245, 158, 11, 0.28);
        color: #92400e;
        margin: 10px 0 18px 0;
    }

    .section-title {
        font-size: 1.05rem;
        font-weight: 900;
        margin: 26px 0 10px 0;
        color: #0f172a;
    }

    /* Abas */
    div[data-baseweb="tab-list"] {
        gap: 8px;
        flex-wrap: wrap;
    }

    button[data-baseweb="tab"] {
        border-radius: 999px;
        padding: 8px 18px;
        background: rgba(255, 255, 255, 0.68);
        border: 1px solid rgba(15, 23, 42, 0.08);
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        background: #1d4ed8;
        color: white;
    }

    /* Botões e inputs */
    button[kind="secondary"] {
        border-radius: 14px !important;
        border: 1px solid rgba(15, 23, 42, 0.10) !important;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stDataFrame"] {
        border-radius: 18px;
        overflow: hidden;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06);
    }

    section[data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.74);
        border-right: 1px solid rgba(15, 23, 42, 0.08);
    }

    @media (prefers-color-scheme: dark) {
        .stApp {
            background: linear-gradient(180deg, #020617 0%, #0f172a 100%);
        }

        h1, h2, h3, .section-title {
            color: #e5e7eb;
        }

        div[data-testid="stMetric"],
        .ranking-podium {
            background: rgba(15, 23, 42, 0.86);
            border-color: rgba(255, 255, 255, 0.10);
        }

        div[data-testid="stMetricValue"] {
            color: #f8fafc !important;
        }

        .soft-note {
            color: #bfdbfe;
            background: rgba(59, 130, 246, 0.12);
            border-color: rgba(147, 197, 253, 0.22);
        }

        .status-card {
            background: rgba(15, 23, 42, 0.86);
            border-color: rgba(255, 255, 255, 0.10);
        }

        .status-card b {
            color: #f8fafc;
        }

        .zero-card {
            color: #fde68a;
            background: rgba(245, 158, 11, 0.13);
            border-color: rgba(245, 158, 11, 0.32);
        }

        section[data-testid="stSidebar"] {
            background: rgba(15, 23, 42, 0.82);
            border-right: 1px solid rgba(255, 255, 255, 0.08);
        }

        button[data-baseweb="tab"] {
            background: rgba(15, 23, 42, 0.75);
            border-color: rgba(255, 255, 255, 0.10);
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            background: #2563eb;
            color: white;
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
# Streamlit Cloud pode não manter/gravar bem no diretório do app.
# /tmp é o local mais seguro para guardar o snapshot temporário entre atualizações.
STATUS_SNAPSHOT_PATH = Path(tempfile.gettempdir()) / "status_dashboard_snapshot.json"

ARQUIVOS = {
    "notas": "notas_dashboard.csv",
    "contratos": "faturamento_contratos_dashboard.csv",
    "dias": "faturamento_dias_dashboard.csv",
    "carro": "faturamento_carro_estimado_dashboard.csv",
    "carro_dias": "faturamento_carro_dias_dashboard.csv",
}

# ==============================
# CONTRATO LEITURA (em testes)
# ==============================
# No PC da operação, o extrator de leitura grava aqui.
# No Streamlit Cloud, o app só consegue ler se os arquivos forem enviados ao GitHub,
# preferencialmente em dashboard/leitura/.
PASTAS_LEITURA = [
    Path(r"C:\Users\user\Desktop\LEITURA\saida"),
    PASTA_DASHBOARD / "leitura",
    PASTA_DASHBOARD,
    PASTA_ATUAL / "leitura",
    PASTA_ATUAL,
]

ARQUIVOS_LEITURA = {
    "Americana": ["Parcial_Americana.xlsx", "Parcial_Americana*.xlsx"],
    "Piracicaba": ["Parcial_Piracicaba.xlsx", "Parcial_Piracicaba*.xlsx"],
}

ORDEM_DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# Como dias anteriores não mudam, mantemos o carregamento geral em 15 minutos
# e deixamos cálculos históricos/ranqueamentos em cache mais longo.
CACHE_TTL_SEGUNDOS = 900
CACHE_TTL_RANKING_SEGUNDOS = 900

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


def caminho_leitura(base_nome):
    """Localiza a parcial de leitura mais recente para Americana ou Piracicaba."""
    padroes = ARQUIVOS_LEITURA.get(base_nome, [])
    candidatos = []

    for pasta in PASTAS_LEITURA:
        try:
            if not pasta.exists():
                continue
        except Exception:
            continue

        for padrao in padroes:
            if "*" in padrao:
                candidatos.extend(list(pasta.glob(padrao)))
            else:
                caminho = pasta / padrao
                if caminho.exists():
                    candidatos.append(caminho)

    candidatos = [c for c in candidatos if c.is_file() and c.suffix.lower() in [".xlsx", ".xls"]]
    if not candidatos:
        return None

    return max(candidatos, key=lambda p: p.stat().st_mtime)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def ler_parcial_leitura(caminho):
    """Lê a parcial de leitura gerada pelo extrator CWSI."""
    df = pd.read_excel(caminho, engine="openpyxl")
    df.columns = [str(c).strip().upper() for c in df.columns]

    mapa_colunas = {}
    for col in df.columns:
        col_norm = str(col).strip().upper()
        if col_norm in ["AGENTE COMERCIAL", "AGENTE"]:
            mapa_colunas[col] = "AGENTE COMERCIAL"
        elif col_norm in ["T. INSTALA", "T INSTALA", "TOTAL INSTALA", "INSTALA"]:
            mapa_colunas[col] = "T. INSTALA"
        elif col_norm in ["T. VISITADA", "T VISITADA", "TOTAL VISITADA", "VISITADA"]:
            mapa_colunas[col] = "T. VISITADA"
        elif col_norm in ["FALTAM", "FALTAM "]:
            mapa_colunas[col] = "FALTAM"
        elif "%" in col_norm and "EXEC" in col_norm:
            mapa_colunas[col] = "% EXECUTADO"

    df = df.rename(columns=mapa_colunas)

    for col in ["AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA"]:
        if col not in df.columns:
            df[col] = "" if col == "AGENTE COMERCIAL" else 0

    df["AGENTE COMERCIAL"] = df["AGENTE COMERCIAL"].fillna("").astype(str).str.strip()
    df = df[df["AGENTE COMERCIAL"] != ""].copy()

    df["T. INSTALA"] = pd.to_numeric(df["T. INSTALA"], errors="coerce").fillna(0).astype(int)
    df["T. VISITADA"] = pd.to_numeric(df["T. VISITADA"], errors="coerce").fillna(0).astype(int)

    df["FALTAM"] = (df["T. INSTALA"] - df["T. VISITADA"]).clip(lower=0).astype(int)
    df["% EXECUTADO"] = 0.0
    mask = df["T. INSTALA"] > 0
    df.loc[mask, "% EXECUTADO"] = ((df.loc[mask, "T. VISITADA"] / df.loc[mask, "T. INSTALA"]) * 100).round(2)

    df = df[["AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA", "FALTAM", "% EXECUTADO"]]
    return df.sort_values(["% EXECUTADO", "FALTAM"], ascending=[True, False]).reset_index(drop=True)


def mostrar_base_leitura(base_nome):
    caminho = caminho_leitura(base_nome)

    st.markdown(f"### {base_nome}")
    if not caminho:
        st.warning(f"Parcial de {base_nome} não encontrada. Procure por Parcial_{base_nome}.xlsx em C:\\Users\\user\\Desktop\\LEITURA\\saida ou em dashboard/leitura no GitHub.")
        return

    try:
        df = ler_parcial_leitura(str(caminho))
    except Exception as e:
        st.error(f"Não foi possível ler a parcial de {base_nome}: {e}")
        st.caption(f"Arquivo encontrado: {caminho}")
        return

    if df.empty:
        st.info(f"A parcial de {base_nome} foi encontrada, mas está vazia.")
        st.caption(f"Arquivo encontrado: {caminho}")
        return

    total_instala = int(df["T. INSTALA"].sum())
    total_visitada = int(df["T. VISITADA"].sum())
    total_faltam = int(df["FALTAM"].sum())
    percentual = (total_visitada / total_instala * 100) if total_instala else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("T. Instala", numero(total_instala))
    c2.metric("T. Visitada", numero(total_visitada))
    c3.metric("Faltam", numero(total_faltam))
    c4.metric("% Executado", f"{percentual:.1f}%".replace(".", ","))

    mtime = arquivo_mtime_datetime(caminho)
    if mtime:
        st.caption(f"Arquivo: {caminho} • atualizado em {mtime.strftime('%d/%m/%Y %H:%M:%S')}")
    else:
        st.caption(f"Arquivo: {caminho}")

    tabela = df.copy()
    tabela["% EXECUTADO"] = tabela["% EXECUTADO"].apply(lambda v: f"{float(v):.0f}%")
    st.dataframe(tabela, use_container_width=True, hide_index=True)



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
        elif col in ["QTD_NOTAS", "NOTAS", "CORTES", "RELIGUES", "RECUSAS", "EXPRESS", "TOTAL_NOTAS"]:
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
        elif col in ["POSIÇÃO", "NOTAS", "CORTES", "RELIGUES", "RECUSAS", "EXPRESS", "DIAS_ATIVOS", "QTD_EQUIPES", "QTD_RECURSOS"]:
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
def preparar_parcial_do_dia(notas, incluir_recusas=False):
    """
    Monta a base da parcial do dia.

    Por padrão, mantém o comportamento antigo: considera apenas notas pagáveis,
    ou seja, sem recusa.

    Quando incluir_recusas=True, mantém também as recusas para exibição na aba
    "Parcial do dia". As recusas entram com faturamento zerado e NÃO contam como
    notas feitas nos indicadores/ranking.
    """
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
    df["EH_RECUSA"] = (df["RECUSA"] != "").astype(int)

    # No modo padrão, mantém apenas notas pagáveis.
    # No modo incluir_recusas=True, as recusas permanecem apenas para exibição.
    if not incluir_recusas:
        df = df[df["RECUSA"] == ""].copy()

    linhas = []

    for _, row in df.iterrows():
        recurso = row.get("RECURSO", "")
        grupo = row.get("GRUPO_NOTA", "")
        qtd_exec = int(row.get("QTD_EXECUTORES", 0) or 0)
        eh_recusa = str(row.get("RECUSA", "")).strip() != ""

        contrato = ""
        faturamento = 0.0
        faturamento_min = 0.0
        faturamento_max = 0.0

        if eh_disjuntor_jundiai(recurso):
            contrato = "Disjuntor Jundiaí"
            if not eh_recusa:
                faturamento = {"CORTE": 13.72, "RELIGUE": 27.43}.get(grupo, 0.0)
                faturamento_min = faturamento
                faturamento_max = faturamento

        elif eh_disjuntor_santa_cruz(recurso):
            contrato = "Disjuntor Santa Cruz"
            if not eh_recusa:
                faturamento = {"CORTE": 11.98, "RELIGUE": 23.97}.get(grupo, 0.0)
                faturamento_min = faturamento
                faturamento_max = faturamento

        elif str(recurso).startswith("JUN58") and qtd_exec >= 2:
            contrato = "Contrato Carro STC estimado"
            if not eh_recusa:
                faturamento_min = {"CORTE": 38.18, "RELIGUE": 36.36}.get(grupo, 0.0)
                faturamento_max = {"CORTE": 45.45, "RELIGUE": 50.91}.get(grupo, 0.0)
                faturamento = faturamento_min

        if contrato:
            item = row.to_dict()
            item["CONTRATO"] = contrato
            item["FATURAMENTO"] = faturamento
            item["FATURAMENTO_MIN"] = faturamento_min
            item["FATURAMENTO_MAX"] = faturamento_max
            item["EH_CORTE"] = 1 if (grupo == "CORTE" and not eh_recusa) else 0
            item["EH_RELIGUE"] = 1 if (grupo == "RELIGUE" and not eh_recusa) else 0
            item["EH_RECUSA"] = 1 if eh_recusa else 0
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

    resumo_contrato = aplicar_express_no_resumo_contrato(
        resumo_contrato,
        notas,
        meses_escolhidos,
        contrato_escolhido,
    )

    if not resumo_contrato.empty and "FATURAMENTO" in resumo_contrato.columns:
        resumo_contrato = resumo_contrato.sort_values("FATURAMENTO", ascending=False)

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
            "EXPRESS": 0,
            "FATURAMENTO_EXPRESS": 0.0,
            "FATURAMENTO_MIN": 0.0,
            "FATURAMENTO_MAX": 0.0,
        }

    return {
        "FATURAMENTO": float(resumo_contrato["FATURAMENTO"].sum()),
        "TOTAL_NOTAS": int(resumo_contrato["TOTAL_NOTAS"].sum()),
        "CORTES": int(resumo_contrato["CORTES"].sum()),
        "RELIGUES": int(resumo_contrato["RELIGUES"].sum()),
        "EXPRESS": int(resumo_contrato["EXPRESS"].sum()) if "EXPRESS" in resumo_contrato.columns else 0,
        "FATURAMENTO_EXPRESS": float(resumo_contrato["FATURAMENTO_EXPRESS"].sum()) if "FATURAMENTO_EXPRESS" in resumo_contrato.columns else 0.0,
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


def arquivo_mtime_datetime(caminho):
    """Retorna a data/hora de São Paulo da última modificação do arquivo."""
    try:
        return datetime.fromtimestamp(
            Path(caminho).stat().st_mtime,
            tz=ZoneInfo("America/Sao_Paulo")
        )
    except Exception:
        return None


def contar_notas_por_contrato(notas):
    """
    Conta notas feitas por contrato, sem contar recusas.
    Também retorna o total geral.
    """
    parcial = preparar_parcial_do_dia(notas)

    contagens = {"Todos": 0}

    if parcial.empty:
        return contagens

    contagens["Todos"] = int(parcial["ORDEM_DE_SERVICO"].nunique())

    por_contrato = (
        parcial.groupby("CONTRATO", dropna=False)["ORDEM_DE_SERVICO"]
        .nunique()
        .to_dict()
    )

    for contrato, qtd in por_contrato.items():
        contagens[str(contrato)] = int(qtd)

    return contagens



def normalizar_executor(valor):
    """Normaliza código de executor vindo da base de notas ou da planilha de express."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def normalizar_ordem_servico(valor):
    """Normaliza Ordem de Serviço/NOTA para permitir o cruzamento entre Excel Express e CSV de notas."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    # Mantém só dígitos quando a nota veio como número/texto com pontuação.
    apenas_digitos = "".join(ch for ch in texto if ch.isdigit())
    return apenas_digitos or texto.upper()


def normalizar_nome_pessoa(valor):
    """Normaliza nome para fazer o DE/PARA do Pagamento Express sem depender de acentos/caixa/espaços."""
    import re
    import unicodedata

    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


DEPARA_NOME_RECURSO_EXPRESS = {
    normalizar_nome_pessoa("FERNANDO LOPES TEIXEIRA"): "JUN5537-EMP",
    normalizar_nome_pessoa("MATHEUS RIBEIRO SILVA"): "JUN5994-EMP",
    normalizar_nome_pessoa("IVANILSON ANTONIO DA SILVA"): "SAL5507-EMP",
    normalizar_nome_pessoa("JESSICA TAYANE DE SOUZA PEREIRA"): "JUN5983-EMP",
    normalizar_nome_pessoa("NATAN GONCALVES GARCIA"): "JUN5539-EMP",
    normalizar_nome_pessoa("WESLEY BRUNO MESSIAS PEREIRA"): "SAL5505-EMP",
    normalizar_nome_pessoa("RENATO DE LIMA"): "JUN5972-EMP",
    normalizar_nome_pessoa("GUILHERME VINICIUS DOS SANTOS"): "JUN5973-EMP",
    normalizar_nome_pessoa("JEANDERSON LUIZ NEVES DE JESUS"): "JUN5974-EMP",
    normalizar_nome_pessoa("JOSE DANILO SANTOS DA SILVA"): "JUN5975-EMP",
    normalizar_nome_pessoa("FRANCISCO JOSE DE SOUSA FILHO"): "JUN5976-EMP",
    normalizar_nome_pessoa("FRANCISCO JOSE DE SOUZA FILHO"): "JUN5976-EMP",
    normalizar_nome_pessoa("MISAEL DE SOUZA ESTRELA"): "JUN5977-EMP",
    normalizar_nome_pessoa("JENIFFER RODRIGUES"): "JUN5981-EMP",
    normalizar_nome_pessoa("KEMERSON JAQUES DA CRUZ"): "JUN5982-EMP",
    normalizar_nome_pessoa("CHARMONE DIONATAS PINHEIRO RODRIGUES"): "JUN5993-EMP",
    normalizar_nome_pessoa("JEFERSON DE SOUZA DA SILVA"): "JUN5990-EMP",
    normalizar_nome_pessoa("THAINA MORAIS DIAS"): "JUN5991-EMP",
    normalizar_nome_pessoa("THAINA DE MORAIS DIAS"): "JUN5991-EMP",
    normalizar_nome_pessoa("IVANA LAIS DE PAULA OLIVEIRA"): "JUN5992-EMP",
    normalizar_nome_pessoa("SIDNEI DE MORAIS SOARES"): "SAL5500-EMP",
    normalizar_nome_pessoa("RAFAEL DELFINO DA SILVA"): "SAL5504-EMP",
    normalizar_nome_pessoa("ELIO CALDEIRA RIBEIRO JUNIOR"): "JUN5929",
    normalizar_nome_pessoa("EVANDRO GOMES ANASTACIO"): "SAL5506-EMP",
    normalizar_nome_pessoa("RONALDO CESAR JACINTO FRANCISCO"): "JUN5995-EMP",
    normalizar_nome_pessoa("ANDERSON ALEXANDRE NORONHA CAMARGO"): "SAL5508-EMP",
    normalizar_nome_pessoa("HENRIQUE SOTTO MARTINS"): "SAL5509-EMP",
    normalizar_nome_pessoa("BRUNO LEONARDO CAETANO DE OLIVEIRA"): "JUN5925",
    normalizar_nome_pessoa("JOSE DANILO DA SILVA SANTOS"): "JUN5927",
    normalizar_nome_pessoa("ELSON ALESSANDRO RODRIGUES"): "SAL5515-EMP",
    normalizar_nome_pessoa("RENE DA SILVA PEREIRA"): "SAL5520-EMP",
    normalizar_nome_pessoa("CLEITON MEDRADO SILVA"): "SAL5521-EMP",
    normalizar_nome_pessoa("LUCAS CHAGAS"): "JUN5970-EMP",
    normalizar_nome_pessoa("JARBAS PEREIRA DOS SANTOS"): "JUN5978-EMP",
    normalizar_nome_pessoa("LEANDRO DOS SANTOS"): "JUN5980-EMP",
    normalizar_nome_pessoa("ADRIANO ROSA DA SILVA"): "JUN5971-EMP",
    normalizar_nome_pessoa("TIAGO CALANDRINI DE OLIVEIRA"): "JUN5979-EMP",


    # Santa Cruz - DE/PARA por número da equipe. O app resolve automaticamente
    # o prefixo correto da base (ex.: 8913 -> MOC8913-EMP).
    normalizar_nome_pessoa("ELDER FABIANO DOS SANTOS"): "8901",
    normalizar_nome_pessoa("REGINALDO APARECIDO RODRIGUES"): "8903",
    normalizar_nome_pessoa("JORGE HUGO DE SOUZA"): "8904",
    normalizar_nome_pessoa("ITN8905"): "8905",
    normalizar_nome_pessoa("LUCAS ALVES CARRIEL"): "8906",
    normalizar_nome_pessoa("LEANDRO APARECIDO"): "8907",
    normalizar_nome_pessoa("PAMELA SALLES AFONSO"): "8933",
    normalizar_nome_pessoa("RAUL VINICIUS BUZZO BARBOSA"): "8902",
    normalizar_nome_pessoa("RAUL VINICIUS BUZZO BARBOSA."): "8902",
    normalizar_nome_pessoa("FERNANDO HERNANDES JESUINO"): "8913",
    normalizar_nome_pessoa("CRISTIAN KAUAN SILVA DE ALMEIDA"): "8909",
    normalizar_nome_pessoa("YURI DA COSTA PAULA"): "2002",
    normalizar_nome_pessoa("ANTONIO CESAR"): "8910",
    normalizar_nome_pessoa("LUIS GUSTAVO SALES ALVES"): "8911",
    normalizar_nome_pessoa("WAGNER BALDUINO DA ROCHA"): "8914",
    normalizar_nome_pessoa("PAULO HENRIQUE DA SILVA"): "8999",
    normalizar_nome_pessoa("JOAO VINICIUS DA SILVA"): "8916",
    normalizar_nome_pessoa("THIAGO CARLOS DA SILVA DE MELLO"): "8932",
    normalizar_nome_pessoa("HENRY DE ALMEIDA"): "8922",
    normalizar_nome_pessoa("LUCAS CAMARGO"): "8921",
}


# Carro STC estimado - só conta quando a dupla completa bate.
# O valor mapeado é só o número; o app transforma em JUN58xx-EMP.
DEPARA_DUPLA_CARRO_EXPRESS = {
    frozenset([normalizar_nome_pessoa("ANDERSON GOMES DA SILVA"), normalizar_nome_pessoa("IGOR TORRES BEZERRA")]): "5808",
    frozenset([normalizar_nome_pessoa("EDUARDO DOS SANTOS"), normalizar_nome_pessoa("CARLOS LEANDRO LOPES DE SOUZA")]): "5812",
    frozenset([normalizar_nome_pessoa("JULIANO CESAR DOS SANTOS"), normalizar_nome_pessoa("JOAO PAULO ROMANHA DOS SANTOS")]): "5803",
    frozenset([normalizar_nome_pessoa("ALEXANDRE LUIZ SANTANA PRAJELAS"), normalizar_nome_pessoa("FELIPE HENRIQUE DE SOUZA FERREIRA")]): "5802",
    frozenset([normalizar_nome_pessoa("LUIS FELIPE SA DOS SANTOS"), normalizar_nome_pessoa("WESLEY APARECIDO DE SA SOUZA")]): "5817",
    frozenset([normalizar_nome_pessoa("LUCIANO HENRIQUE DE SOUZA"), normalizar_nome_pessoa("ORLANDO MANOEL DE SOUZA JANSEN")]): "5814",
    frozenset([normalizar_nome_pessoa("ANDERSON PEREIRA SANTOS"), normalizar_nome_pessoa("DIOLENO CONCEICAO NOGUEIRA")]): "5801",
    frozenset([normalizar_nome_pessoa("ANDERSON SILVA LEONARDO"), normalizar_nome_pessoa("JEFERSON E SILVA GOMES")]): "5820",
    frozenset([normalizar_nome_pessoa("LUCAS EDUARDO LOPES DE SOUZA"), normalizar_nome_pessoa("JHONATAN MATHEUS LOPES DE SOUZA")]): "5810",
    frozenset([normalizar_nome_pessoa("ADILSON APARECIDO VASCO"), normalizar_nome_pessoa("RILDO AUGUSTO DOS SANTOS")]): "5806",
    frozenset([normalizar_nome_pessoa("MARISTON OLIVEIRA NASCIMENTO"), normalizar_nome_pessoa("VALDECI SARDELA")]): "5807",
}


def recurso_carro_por_dupla(nome_1_norm, nome_2_norm):
    """Retorna o recurso do Carro quando os DOIS nomes da dupla batem."""
    nome_1_norm = str(nome_1_norm).strip()
    nome_2_norm = str(nome_2_norm).strip()
    if not nome_1_norm or not nome_2_norm:
        return ""
    return DEPARA_DUPLA_CARRO_EXPRESS.get(frozenset([nome_1_norm, nome_2_norm]), "")


def contrato_por_recurso_express(recurso):
    recurso = str(recurso).strip().upper()
    if eh_disjuntor_jundiai(recurso):
        return "Disjuntor Jundiaí"
    if eh_disjuntor_santa_cruz(recurso):
        return "Disjuntor Santa Cruz"
    if recurso.startswith("JUN58"):
        return "Contrato Carro STC estimado"
    return ""


def codigo_numerico_recurso(valor):
    """Extrai o código numérico de um recurso/equipe, ex.: MOC8913-EMP -> 8913."""
    import re
    texto = str(valor).strip().upper()
    m = re.search(r"(\d+)", texto)
    return m.group(1) if m else ""


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def mapa_codigo_para_recurso_real(notas):
    """
    Monta um mapa código numérico -> RECURSO real encontrado na base.
    Isso permite cadastrar Santa Cruz só como 8913 e converter para MOC8913-EMP,
    ITN8922-EMP etc., conforme aparece no ranking.
    """
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)
    if parcial.empty or "RECURSO" not in parcial.columns:
        return {}

    tmp = parcial[["RECURSO"]].copy()
    tmp["RECURSO"] = tmp["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    tmp = tmp[tmp["RECURSO"] != ""].copy()
    if tmp.empty:
        return {}

    tmp["CODIGO_RECURSO"] = tmp["RECURSO"].apply(codigo_numerico_recurso)
    tmp = tmp[tmp["CODIGO_RECURSO"] != ""].copy()
    if tmp.empty:
        return {}

    contagem = (
        tmp.groupby(["CODIGO_RECURSO", "RECURSO"])
        .size()
        .reset_index(name="QTD")
        .sort_values(["CODIGO_RECURSO", "QTD"], ascending=[True, False])
    )
    contagem = contagem.drop_duplicates(subset=["CODIGO_RECURSO"], keep="first")
    return dict(zip(contagem["CODIGO_RECURSO"], contagem["RECURSO"]))


def resolver_recurso_depara(valor, mapa_codigo_recurso):
    """
    Resolve o valor do DE/PARA para o recurso real do ranking.
    - Se vier completo (SAL5508-EMP), mantém.
    - Se vier só número (8913), busca na base e vira MOC8913-EMP/ITN8913-EMP etc.
    """
    recurso = str(valor).strip().upper()
    if recurso == "" or recurso in ["NAN", "NONE"]:
        return ""
    if recurso.isdigit():
        # Santa Cruz normalmente é resolvido pelo mapa da base (MOC/ITN/etc.).
        # Para Carro, se a base não tiver o recurso no mapa, força o padrão JUN58xx-EMP.
        if recurso.startswith("58"):
            return mapa_codigo_recurso.get(recurso, f"JUN{recurso}-EMP")
        return mapa_codigo_recurso.get(recurso, recurso)
    return recurso


def caminho_pagamento_express():
    """
    Procura a planilha manual de Pagamento Express.

    Aceita nomes como:
    - pagamento_express.xlsx
    - pagamento_express.xlsx.xlsx
    - pagamento_express.csv
    - express.xlsx
    - express.csv
    """
    nomes = [
        "pagamento_express.xlsx",
        "pagamento_express.xlsx.xlsx",
        "pagamento_express.csv",
        "express.xlsx",
        "express.csv",
    ]

    for nome in nomes:
        for pasta in [PASTA_DASHBOARD, PASTA_ATUAL]:
            caminho = pasta / nome
            if caminho.exists():
                return caminho

    for pasta in [PASTA_DASHBOARD, PASTA_ATUAL]:
        achados = (
            list(pasta.glob("pagamento_express*.xlsx"))
            + list(pasta.glob("pagamento_express*.csv"))
            + list(pasta.glob("express*.xlsx"))
            + list(pasta.glob("express*.csv"))
        )
        if achados:
            return achados[0]

    return None


def ler_pagamento_express(caminho):
    """
    Lê a planilha manual do Pagamento Express.

    Versão robusta:
    - não depende de acento no nome das colunas;
    - aplica filtro de VALIDAÇÃO de forma não destrutiva;
    - se o filtro de validação zerar o arquivo, mantém as linhas e mostra auditoria;
    - lê DT_REFERENCIA mesmo quando a coluna vem como data real do Excel;
    - usa NOME_EXECUTOR_01/02 para o DE/PARA Nome -> Recurso.
    """
    if not caminho:
        return pd.DataFrame()

    caminho = str(caminho)

    try:
        if caminho.lower().endswith(".xlsx"):
            df_original = pd.read_excel(caminho, engine="openpyxl")
        else:
            try:
                df_original = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
            except Exception:
                df_original = pd.read_csv(caminho, sep=",", encoding="utf-8-sig")
    except Exception as e:
        df_erro = pd.DataFrame()
        df_erro.attrs["ERRO_LEITURA_EXPRESS"] = str(e)
        return df_erro

    if df_original.empty:
        return pd.DataFrame()

    df = df_original.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Mapa de colunas normalizadas, tolerando acento/espaço.
    colunas_norm = {normalizar_nome_pessoa(c): c for c in df.columns}

    def achar_coluna(*nomes):
        for nome in nomes:
            chave = normalizar_nome_pessoa(nome)
            if chave in colunas_norm:
                return colunas_norm[chave]
        return None

    # Filtro de validação NÃO destrutivo.
    # Seu arquivo já é um arquivo de pagamento express; o filtro só é usado se encontrar linhas.
    col_validacao = achar_coluna("VALIDAÇÃO", "VALIDACAO")
    linhas_brutas = len(df)
    linhas_pos_validacao = None
    if col_validacao:
        validacao_txt = df[col_validacao].fillna("").astype(str).apply(normalizar_nome_pessoa)
        mascara_express = validacao_txt.str.contains("PAGAMENTO", na=False) & validacao_txt.str.contains("EXPRESS", na=False)
        linhas_pos_validacao = int(mascara_express.sum())
        if linhas_pos_validacao > 0:
            df = df[mascara_express].copy()
        else:
            # Se o filtro não encontrou nada, não zera o arquivo.
            df = df.copy()

    df.attrs["EXPRESS_LINHAS_BRUTAS"] = linhas_brutas
    df.attrs["EXPRESS_LINHAS_POS_VALIDACAO"] = linhas_pos_validacao

    # NOTA/OS fica disponível para auditoria, mas a contabilização usa o nome.
    col_nota = achar_coluna("NOTA", "ORDEM_DE_SERVICO", "ORDEM DE SERVICO", "OS")
    if col_nota:
        df["NOTA_NORM"] = df[col_nota].apply(normalizar_ordem_servico)
    else:
        df["NOTA_NORM"] = ""

    # Nome principal da pessoa no arquivo Express.
    col_nome_1 = achar_coluna("NOME_EXECUTOR_01", "NOME EXECUTOR 01", "NOME_EXECUTOR", "NOME EXECUTOR")
    col_nome_2 = achar_coluna("NOME_EXECUTOR_02", "NOME EXECUTOR 02")
    col_executor = achar_coluna("EXECUTOR")

    if col_nome_1:
        df["NOME_EXPRESS"] = df[col_nome_1].fillna("").astype(str).str.strip()
    elif col_executor:
        df["NOME_EXPRESS"] = df[col_executor].fillna("").astype(str).str.strip()
    else:
        df["NOME_EXPRESS"] = ""

    # Mantém o segundo nome separado para o contrato Carro.
    # Importante: não misturar NOME_EXECUTOR_02 dentro do NOME_EXPRESS principal,
    # porque Jundiaí/Santa Cruz continuam usando o nome 01 individualmente.
    if col_nome_2:
        df["NOME_EXPRESS_02"] = df[col_nome_2].fillna("").astype(str).str.strip()
    else:
        df["NOME_EXPRESS_02"] = ""

    # Se o nome 01 vier vazio, aí sim usa o nome 02 como fallback para não perder
    # linhas antigas que tinham só uma coluna de executor.
    if col_nome_2:
        mascara_vazia = df["NOME_EXPRESS"].eq("") | df["NOME_EXPRESS"].str.upper().eq("NAN")
        df.loc[mascara_vazia, "NOME_EXPRESS"] = (
            df.loc[mascara_vazia, col_nome_2].fillna("").astype(str).str.strip()
        )

    df["NOME_EXPRESS"] = df["NOME_EXPRESS"].replace({"nan": "", "NaN": "", "None": ""})
    df["NOME_EXPRESS_02"] = df["NOME_EXPRESS_02"].replace({"nan": "", "NaN": "", "None": ""})
    df["NOME_EXPRESS_NORM"] = df["NOME_EXPRESS"].apply(normalizar_nome_pessoa)
    df["NOME_EXPRESS_02_NORM"] = df["NOME_EXPRESS_02"].apply(normalizar_nome_pessoa)
    df = df[(df["NOME_EXPRESS_NORM"] != "") | (df["NOME_EXPRESS_02_NORM"] != "")].copy()

    # Data de referência do Express.
    col_data = achar_coluna(
        "DT_REFERENCIA", "DT REFERENCIA", "DT_REFERÊNCIA", "DT REFERÊNCIA",
        "DATA_REFERENCIA", "DATA REFERENCIA", "DATA_REFERÊNCIA", "DATA REFERÊNCIA", "DATA"
    )

    if col_data is None:
        for col in df.columns:
            col_norm = normalizar_nome_pessoa(col)
            if ("REFERENCIA" in col_norm or "REF" in col_norm) and ("DATA" in col_norm or "DT" in col_norm):
                col_data = col
                break

    if col_data is not None:
        serie_data = df[col_data]
        df["DATA_EXPRESS_DT"] = pd.to_datetime(serie_data, dayfirst=True, errors="coerce")

        # Fallback para datas numéricas do Excel.
        if df["DATA_EXPRESS_DT"].isna().all():
            serie_num = pd.to_numeric(serie_data, errors="coerce")
            df["DATA_EXPRESS_DT"] = pd.to_datetime(serie_num, unit="D", origin="1899-12-30", errors="coerce")
    else:
        df["DATA_EXPRESS_DT"] = pd.NaT

    df.attrs["EXPRESS_COLUNAS"] = list(df_original.columns)
    df.attrs["EXPRESS_COL_VALIDACAO"] = col_validacao or ""
    df.attrs["EXPRESS_COL_DATA"] = col_data or ""
    df.attrs["EXPRESS_COL_NOME_1"] = col_nome_1 or ""

    return df



@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def mapa_executor_recurso(notas):
    """
    Cria o de/para EXECUTOR -> RECURSO/CONTRATO usando a própria base de notas.
    Considera que cada executor é único e pertence a um único recurso.
    """
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)

    if parcial.empty:
        return pd.DataFrame(columns=["EXECUTOR_NORM", "RECURSO", "CONTRATO"])

    linhas = []

    for col in ["ELETRICISTA1", "ELETRICISTA2"]:
        if col in parcial.columns:
            tmp = parcial[[col, "RECURSO", "CONTRATO"]].copy()
            tmp = tmp.rename(columns={col: "EXECUTOR"})
            tmp["EXECUTOR_NORM"] = tmp["EXECUTOR"].apply(normalizar_executor)
            tmp = tmp[tmp["EXECUTOR_NORM"] != ""].copy()
            linhas.append(tmp[["EXECUTOR_NORM", "RECURSO", "CONTRATO"]])

    if not linhas:
        return pd.DataFrame(columns=["EXECUTOR_NORM", "RECURSO", "CONTRATO"])

    mapa = pd.concat(linhas, ignore_index=True).drop_duplicates(subset=["EXECUTOR_NORM"])
    mapa["RECURSO"] = mapa["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    mapa["CONTRATO"] = mapa["CONTRATO"].fillna("").astype(str).str.strip()

    return mapa


def valor_express_por_contrato(contrato):
    """
    Express faturado como religue para equipes de disjuntor.
    Carro é diferente; por enquanto fica zerado no faturamento express.
    """
    contrato = str(contrato)
    if contrato == "Disjuntor Jundiaí":
        return 27.43
    if contrato == "Disjuntor Santa Cruz":
        return 23.97
    if contrato == "Contrato Carro STC estimado":
        return 38.18
    return 0.0


def calcular_express_mensal(notas, mes):
    """
    Calcula Pagamento Express por RECURSO para o mês escolhido.

    Regra definitiva: usa o DE/PARA manual Nome -> Recurso.
    Isso evita depender de executor, recurso vindo no Excel ou casamento por OS
    quando a nota não bate exatamente com a base atual.
    """
    caminho = caminho_pagamento_express()

    if not caminho:
        return pd.DataFrame(), "", pd.DataFrame(), ""

    express = ler_pagamento_express(str(caminho))
    if express.empty:
        return pd.DataFrame(), "", pd.DataFrame(), str(caminho)

    data_max_txt = ""
    if "DATA_EXPRESS_DT" in express.columns and express["DATA_EXPRESS_DT"].notna().any():
        express["MES_EXPRESS"] = express["DATA_EXPRESS_DT"].dt.strftime("%m/%Y")
        express = express[express["MES_EXPRESS"] == mes].copy()
        data_max = express["DATA_EXPRESS_DT"].max()
        data_max_txt = data_max.strftime("%d/%m/%Y") if pd.notna(data_max) else ""

    # Se a planilha não tiver data válida, não joga tudo fora: deixa a auditoria mostrar
    # o que foi lido. Para este arquivo específico, DT_REFERENCIA deve filtrar 03/2026.
    if express.empty:
        return pd.DataFrame(), data_max_txt, pd.DataFrame(), str(caminho)

    mapa_codigo_recurso = mapa_codigo_para_recurso_real(notas)

    # Jundiaí/Santa Cruz: mantêm exatamente a lógica que já funcionava, por nome individual.
    express["RECURSO_DEPARA"] = express["NOME_EXPRESS_NORM"].map(DEPARA_NOME_RECURSO_EXPRESS).fillna("")

    # Carro: regra adicional, sem interferir nos outros contratos.
    # Só conta se houver NOME_EXECUTOR_01 e NOME_EXECUTOR_02 e a dupla completa bater.
    if "NOME_EXPRESS_02_NORM" not in express.columns:
        express["NOME_EXPRESS_02_NORM"] = ""

    express["RECURSO_CARRO"] = express.apply(
        lambda r: recurso_carro_por_dupla(
            r.get("NOME_EXPRESS_NORM", ""),
            r.get("NOME_EXPRESS_02_NORM", ""),
        ),
        axis=1,
    )

    # O Carro tem prioridade apenas quando a dupla bate; caso contrário, permanece o DE/PARA antigo.
    mascara_carro = express["RECURSO_CARRO"].fillna("").astype(str).str.strip() != ""
    express.loc[mascara_carro, "RECURSO_DEPARA"] = express.loc[mascara_carro, "RECURSO_CARRO"]

    express["RECURSO"] = express["RECURSO_DEPARA"].apply(lambda v: resolver_recurso_depara(v, mapa_codigo_recurso))
    express["RECURSO"] = express["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    express["CONTRATO"] = express["RECURSO"].apply(contrato_por_recurso_express)

    sem_vinculo = express[(express["RECURSO"] == "") | (express["CONTRATO"] == "")].copy()
    express_ok = express[(express["RECURSO"] != "") & (express["CONTRATO"] != "")].copy()

    if express_ok.empty:
        return pd.DataFrame(), data_max_txt, sem_vinculo, str(caminho)

    resumo = (
        express_ok.groupby(["RECURSO", "CONTRATO"], dropna=False)
        .size()
        .reset_index(name="EXPRESS")
    )

    resumo["EXPRESS"] = pd.to_numeric(resumo["EXPRESS"], errors="coerce").fillna(0).astype(int)
    resumo["FATURAMENTO_EXPRESS"] = resumo.apply(
        lambda r: r["EXPRESS"] * valor_express_por_contrato(r.get("CONTRATO", "")),
        axis=1,
    )

    return resumo, data_max_txt, sem_vinculo, str(caminho)


def resumo_express_periodo(notas, meses, contrato_escolhido="Todos"):
    """
    Resume o Pagamento Express por contrato para um ou mais meses.

    O Express usa a coluna DT_REFERENCIA da planilha para cair no mês certo
    e o DE/PARA Nome -> Recurso para descobrir o contrato.
    """
    if not meses:
        return pd.DataFrame(columns=["CONTRATO", "EXPRESS", "FATURAMENTO_EXPRESS"])

    partes = []
    for mes in meses:
        express_resumo, _, _, _ = calcular_express_mensal(notas, mes)
        if express_resumo.empty:
            continue
        tmp = express_resumo.copy()
        if contrato_escolhido != "Todos" and "CONTRATO" in tmp.columns:
            tmp = tmp[tmp["CONTRATO"] == contrato_escolhido].copy()
        if not tmp.empty:
            partes.append(tmp)

    if not partes:
        return pd.DataFrame(columns=["CONTRATO", "EXPRESS", "FATURAMENTO_EXPRESS"])

    express = pd.concat(partes, ignore_index=True)
    resumo = (
        express.groupby("CONTRATO", dropna=False)
        .agg(
            EXPRESS=("EXPRESS", "sum"),
            FATURAMENTO_EXPRESS=("FATURAMENTO_EXPRESS", "sum"),
        )
        .reset_index()
    )
    resumo["EXPRESS"] = pd.to_numeric(resumo["EXPRESS"], errors="coerce").fillna(0).astype(int)
    resumo["FATURAMENTO_EXPRESS"] = pd.to_numeric(resumo["FATURAMENTO_EXPRESS"], errors="coerce").fillna(0.0)
    return resumo


def aplicar_express_no_resumo_contrato(resumo_contrato, notas, meses, contrato_escolhido="Todos"):
    """Soma Express ao resumo financeiro mensal/por período."""
    resumo = resumo_contrato.copy()

    if "EXPRESS" not in resumo.columns:
        resumo["EXPRESS"] = 0
    if "FATURAMENTO_EXPRESS" not in resumo.columns:
        resumo["FATURAMENTO_EXPRESS"] = 0.0

    express = resumo_express_periodo(notas, meses, contrato_escolhido)
    if express.empty:
        return resumo

    resumo = resumo.merge(express, on="CONTRATO", how="outer", suffixes=("", "_NOVO"))

    for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES"]:
        if col not in resumo.columns:
            resumo[col] = 0
        resumo[col] = pd.to_numeric(resumo[col], errors="coerce").fillna(0).astype(int)

    for col in ["FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
        if col not in resumo.columns:
            resumo[col] = 0.0
        resumo[col] = pd.to_numeric(resumo[col], errors="coerce").fillna(0.0)

    resumo["EXPRESS"] = pd.to_numeric(resumo.get("EXPRESS", 0), errors="coerce").fillna(0).astype(int)
    resumo["EXPRESS_NOVO"] = pd.to_numeric(resumo.get("EXPRESS_NOVO", 0), errors="coerce").fillna(0).astype(int)
    resumo["FATURAMENTO_EXPRESS"] = pd.to_numeric(resumo.get("FATURAMENTO_EXPRESS", 0), errors="coerce").fillna(0.0)
    resumo["FATURAMENTO_EXPRESS_NOVO"] = pd.to_numeric(resumo.get("FATURAMENTO_EXPRESS_NOVO", 0), errors="coerce").fillna(0.0)

    resumo["EXPRESS"] = resumo["EXPRESS"] + resumo["EXPRESS_NOVO"]
    resumo["FATURAMENTO_EXPRESS"] = resumo["FATURAMENTO_EXPRESS"] + resumo["FATURAMENTO_EXPRESS_NOVO"]

    # Express entra no total de notas e no faturamento mensal do contrato.
    resumo["TOTAL_NOTAS"] = resumo["TOTAL_NOTAS"] + resumo["EXPRESS"]
    resumo["FATURAMENTO"] = resumo["FATURAMENTO"] + resumo["FATURAMENTO_EXPRESS"]
    resumo["FATURAMENTO_MIN"] = resumo["FATURAMENTO_MIN"] + resumo["FATURAMENTO_EXPRESS"]
    resumo["FATURAMENTO_MAX"] = resumo["FATURAMENTO_MAX"] + resumo["FATURAMENTO_EXPRESS"]

    resumo = resumo.drop(columns=[c for c in ["EXPRESS_NOVO", "FATURAMENTO_EXPRESS_NOVO"] if c in resumo.columns])

    if "CONTRATO" in resumo.columns:
        resumo["CONTRATO"] = resumo["CONTRATO"].fillna(contrato_escolhido if contrato_escolhido != "Todos" else "")

    return resumo


def aplicar_express_no_ranking_mensal(ranking, notas, mes, contrato_ranking):
    """
    Soma Pagamento Express ao ranking mensal por RECURSO.

    Express entra em:
    - NOTAS;
    - EXPRESS;
    - FATURAMENTO_ATRIBUÍDO;
    - FATURAMENTO_EQUIPE.
    """
    ranking = ranking.copy()
    express_resumo, data_max_txt, sem_vinculo, caminho = calcular_express_mensal(notas, mes)

    if "EXPRESS" not in ranking.columns:
        ranking["EXPRESS"] = 0
    if "FATURAMENTO_EXPRESS" not in ranking.columns:
        ranking["FATURAMENTO_EXPRESS"] = 0.0

    if express_resumo.empty:
        total_express = 0
        fat_express = 0.0
        return ranking, express_resumo, data_max_txt, sem_vinculo, caminho, total_express, fat_express

    if contrato_ranking != "Todos" and "CONTRATO" in express_resumo.columns:
        express_resumo = express_resumo[express_resumo["CONTRATO"] == contrato_ranking].copy()

    if express_resumo.empty:
        total_express = 0
        fat_express = 0.0
        return ranking, express_resumo, data_max_txt, sem_vinculo, caminho, total_express, fat_express

    total_express = int(express_resumo["EXPRESS"].sum())
    fat_express = float(express_resumo["FATURAMENTO_EXPRESS"].sum())

    ranking = ranking.merge(
        express_resumo[["RECURSO", "EXPRESS", "FATURAMENTO_EXPRESS"]],
        on="RECURSO",
        how="outer",
        suffixes=("", "_NOVO"),
    )

    for col in ["NOTAS", "CORTES", "RELIGUES", "DIAS_ATIVOS", "QTD_EQUIPES"]:
        if col not in ranking.columns:
            ranking[col] = 0
        ranking[col] = pd.to_numeric(ranking[col], errors="coerce").fillna(0)

    for col in [
        "FATURAMENTO_ATRIBUÍDO", "FATURAMENTO_MIN_ATRIBUÍDO",
        "FATURAMENTO_MAX_ATRIBUÍDO", "FATURAMENTO_EQUIPE"
    ]:
        if col not in ranking.columns:
            ranking[col] = 0.0
        ranking[col] = pd.to_numeric(ranking[col], errors="coerce").fillna(0.0)

    ranking["EXPRESS"] = pd.to_numeric(ranking.get("EXPRESS", 0), errors="coerce").fillna(0)
    ranking["EXPRESS_NOVO"] = pd.to_numeric(ranking.get("EXPRESS_NOVO", 0), errors="coerce").fillna(0)
    ranking["FATURAMENTO_EXPRESS"] = pd.to_numeric(ranking.get("FATURAMENTO_EXPRESS", 0), errors="coerce").fillna(0.0)
    ranking["FATURAMENTO_EXPRESS_NOVO"] = pd.to_numeric(ranking.get("FATURAMENTO_EXPRESS_NOVO", 0), errors="coerce").fillna(0.0)

    ranking["EXPRESS"] = (ranking["EXPRESS"] + ranking["EXPRESS_NOVO"]).astype(int)
    ranking["FATURAMENTO_EXPRESS"] = ranking["FATURAMENTO_EXPRESS"] + ranking["FATURAMENTO_EXPRESS_NOVO"]

    ranking["NOTAS"] = (ranking["NOTAS"] + ranking["EXPRESS"]).astype(int)
    ranking["FATURAMENTO_ATRIBUÍDO"] = ranking["FATURAMENTO_ATRIBUÍDO"] + ranking["FATURAMENTO_EXPRESS"]
    ranking["FATURAMENTO_EQUIPE"] = ranking["FATURAMENTO_EQUIPE"] + ranking["FATURAMENTO_EXPRESS"]

    ranking = ranking.drop(columns=[
        c for c in ["EXPRESS_NOVO", "FATURAMENTO_EXPRESS_NOVO", "POSIÇÃO"]
        if c in ranking.columns
    ])

    ranking["DIAS_ATIVOS"] = pd.to_numeric(ranking["DIAS_ATIVOS"], errors="coerce").fillna(0).astype(int)
    ranking["MÉDIA_NOTAS_DIA"] = ranking.apply(
        lambda r: (r["NOTAS"] / r["DIAS_ATIVOS"]) if r["DIAS_ATIVOS"] else 0,
        axis=1,
    )
    ranking["TICKET_MÉDIO"] = ranking.apply(
        lambda r: (r["FATURAMENTO_ATRIBUÍDO"] / r["NOTAS"]) if r["NOTAS"] else 0,
        axis=1,
    )

    ranking = ranking.sort_values(["NOTAS", "FATURAMENTO_ATRIBUÍDO"], ascending=False).reset_index(drop=True)
    ranking.insert(0, "POSIÇÃO", range(1, len(ranking) + 1))

    return ranking, express_resumo, data_max_txt, sem_vinculo, caminho, total_express, fat_express


def resumo_parcial_mais_recente(notas, contrato_escolhido="Todos"):
    """
    Calcula a produção da data mais recente da base.
    Usa apenas notas feitas, sem recusas.
    """
    parcial = preparar_parcial_do_dia(notas)

    resumo = {
        "data": "",
        "notas": 0,
        "cortes": 0,
        "religues": 0,
        "por_contrato": {},
    }

    if parcial.empty:
        return resumo

    ultima_data_dt = parcial["DATA_DT"].max()
    parcial_dia = parcial[parcial["DATA_DT"] == ultima_data_dt].copy()

    resumo["data"] = ultima_data_dt.strftime("%d/%m/%Y")
    resumo["notas"] = int(parcial_dia["ORDEM_DE_SERVICO"].nunique())
    resumo["cortes"] = int(parcial_dia["EH_CORTE"].sum())
    resumo["religues"] = int(parcial_dia["EH_RELIGUE"].sum())

    for contrato, df_contrato in parcial_dia.groupby("CONTRATO", dropna=False):
        contrato = str(contrato)
        resumo["por_contrato"][contrato] = {
            "notas": int(df_contrato["ORDEM_DE_SERVICO"].nunique()),
            "cortes": int(df_contrato["EH_CORTE"].sum()),
            "religues": int(df_contrato["EH_RELIGUE"].sum()),
        }

    return resumo


def atualizar_status_dashboard(notas, caminho_notas, contrato_escolhido):
    """
    Mantém um snapshot local da última atualização do CSV.

    Agora o comparativo principal é da produção do dia mais recente:
    notas, cortes e religues desde a atualização anterior.
    """
    caminho_status = STATUS_SNAPSHOT_PATH
    agora = datetime.now(ZoneInfo("America/Sao_Paulo"))
    mtime_dt = arquivo_mtime_datetime(caminho_notas) if caminho_notas else None
    mtime = mtime_dt.isoformat() if mtime_dt else ""

    try:
        tamanho_arquivo = Path(caminho_notas).stat().st_size if caminho_notas else 0
    except Exception:
        tamanho_arquivo = 0

    arquivo_id = f"{mtime}|{tamanho_arquivo}"

    contagens = contar_notas_por_contrato(notas)
    total_atual = int(contagens.get("Todos", 0))

    parcial_atual = resumo_parcial_mais_recente(notas, contrato_escolhido)

    status_antigo = {}
    snapshot_erro = ""
    if caminho_status.exists():
        try:
            status_antigo = json.loads(caminho_status.read_text(encoding="utf-8"))
        except Exception:
            status_antigo = {}

    arquivo_id_antigo = status_antigo.get("arquivo_id", "")
    parcial_anterior = status_antigo.get("parcial_atual", {})
    contagens_anteriores = status_antigo.get("contagens", {})

    primeira_execucao = not bool(parcial_anterior)

    if arquivo_id != arquivo_id_antigo:
        if primeira_execucao:
            delta_hoje = 0
            delta_cortes = 0
            delta_religues = 0
            deltas_por_contrato = {
                contrato: {"notas": 0, "cortes": 0, "religues": 0}
                for contrato in parcial_atual.get("por_contrato", {}).keys()
            }
            delta_geral_base = 0
        else:
            # Se mudou a data, começa um novo baseline para o novo dia.
            mesma_data = parcial_atual.get("data") == parcial_anterior.get("data")

            if mesma_data:
                delta_hoje = max(0, int(parcial_atual.get("notas", 0)) - int(parcial_anterior.get("notas", 0)))
                delta_cortes = max(0, int(parcial_atual.get("cortes", 0)) - int(parcial_anterior.get("cortes", 0)))
                delta_religues = max(0, int(parcial_atual.get("religues", 0)) - int(parcial_anterior.get("religues", 0)))
            else:
                delta_hoje = 0
                delta_cortes = 0
                delta_religues = 0

            deltas_por_contrato = {}
            parcial_ant_por_contrato = parcial_anterior.get("por_contrato", {}) if mesma_data else {}

            for contrato, valores in parcial_atual.get("por_contrato", {}).items():
                anterior = parcial_ant_por_contrato.get(contrato, {})
                deltas_por_contrato[contrato] = {
                    "notas": max(0, int(valores.get("notas", 0)) - int(anterior.get("notas", valores.get("notas", 0)))),
                    "cortes": max(0, int(valores.get("cortes", 0)) - int(anterior.get("cortes", valores.get("cortes", 0)))),
                    "religues": max(0, int(valores.get("religues", 0)) - int(anterior.get("religues", valores.get("religues", 0)))),
                }

            delta_geral_base = max(0, total_atual - int(contagens_anteriores.get("Todos", total_atual)))

        status = {
            "arquivo_id": arquivo_id,
            "mtime": mtime,
            "ultima_verificacao": agora.isoformat(),
            "contagens": contagens,
            "parcial_atual": parcial_atual,
            "ultimo_delta_geral_base": int(delta_geral_base),
            "ultimo_delta_hoje": int(delta_hoje),
            "ultimo_delta_cortes": int(delta_cortes),
            "ultimo_delta_religues": int(delta_religues),
            "ultimo_delta_por_contrato": deltas_por_contrato,
        }

        try:
            caminho_status.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            snapshot_erro = str(e)
    else:
        delta_geral_base = int(status_antigo.get("ultimo_delta_geral_base", 0))
        delta_hoje = int(status_antigo.get("ultimo_delta_hoje", 0))
        delta_cortes = int(status_antigo.get("ultimo_delta_cortes", 0))
        delta_religues = int(status_antigo.get("ultimo_delta_religues", 0))
        deltas_por_contrato = status_antigo.get("ultimo_delta_por_contrato", {})

    delta_contrato_info = deltas_por_contrato.get(contrato_escolhido, {"notas": 0, "cortes": 0, "religues": 0})
    if contrato_escolhido == "Todos":
        delta_contrato_info = {
            "notas": int(delta_hoje),
            "cortes": int(delta_cortes),
            "religues": int(delta_religues),
        }

    return {
        "ultima_atualizacao": mtime_dt,
        "total_atual": total_atual,
        "delta_geral_base": int(delta_geral_base),
        "data_parcial": parcial_atual.get("data", ""),
        "notas_hoje": int(parcial_atual.get("notas", 0)),
        "cortes_hoje": int(parcial_atual.get("cortes", 0)),
        "religues_hoje": int(parcial_atual.get("religues", 0)),
        "delta_hoje": int(delta_hoje),
        "delta_cortes": int(delta_cortes),
        "delta_religues": int(delta_religues),
        "delta_contrato": int(delta_contrato_info.get("notas", 0)),
        "delta_contrato_cortes": int(delta_contrato_info.get("cortes", 0)),
        "delta_contrato_religues": int(delta_contrato_info.get("religues", 0)),
        "snapshot_caminho": str(caminho_status),
        "snapshot_erro": snapshot_erro,
    }


def mostrar_status_atualizacao(notas, contrato_escolhido):
    caminho_notas = caminho_arquivo(ARQUIVOS["notas"])
    status = atualizar_status_dashboard(notas, caminho_notas, contrato_escolhido)

    ultima = status.get("ultima_atualizacao")
    ultima_txt = ultima.strftime("%d/%m/%Y %H:%M:%S") if ultima else "não identificado"

    delta_hoje = status.get("delta_hoje", 0)
    delta_cortes = status.get("delta_cortes", 0)
    delta_religues = status.get("delta_religues", 0)

    delta_hoje_txt = f"+{numero(delta_hoje)}" if delta_hoje >= 0 else numero(delta_hoje)
    delta_cortes_txt = f"+{numero(delta_cortes)}" if delta_cortes >= 0 else numero(delta_cortes)
    delta_religues_txt = f"+{numero(delta_religues)}" if delta_religues >= 0 else numero(delta_religues)

    delta_contrato = status.get("delta_contrato", 0)
    delta_contrato_cortes = status.get("delta_contrato_cortes", 0)
    delta_contrato_religues = status.get("delta_contrato_religues", 0)

    delta_contrato_txt = f"+{numero(delta_contrato)}" if delta_contrato >= 0 else numero(delta_contrato)
    delta_contrato_cortes_txt = f"+{numero(delta_contrato_cortes)}" if delta_contrato_cortes >= 0 else numero(delta_contrato_cortes)
    delta_contrato_religues_txt = f"+{numero(delta_contrato_religues)}" if delta_contrato_religues >= 0 else numero(delta_contrato_religues)

    data_parcial = status.get("data_parcial", "")
    texto_data = f" em {data_parcial}" if data_parcial else ""
    snapshot_erro = status.get("snapshot_erro", "")
    aviso_snapshot = f"<br><b>⚠️ Snapshot:</b> erro ao salvar comparativo ({snapshot_erro})" if snapshot_erro else ""

    if contrato_escolhido == "Todos":
        texto_contrato = "Todos os contratos"
        detalhe = (
            f"{delta_hoje_txt} notas na última atualização "
            f"(Cortes: {delta_cortes_txt} • Religues: {delta_religues_txt})"
        )
    else:
        texto_contrato = contrato_escolhido
        detalhe = (
            f"{delta_contrato_txt} notas na última atualização "
            f"(Cortes: {delta_contrato_cortes_txt} • Religues: {delta_contrato_religues_txt})"
        )

    st.markdown(
        f"""
        <div class="status-card">
            <b>🕒 Última atualização dos dados:</b> {ultima_txt}<br>
            <b>📈 Parcial do dia{texto_data}:</b> {delta_hoje_txt} notas na última atualização
            (Cortes: {delta_cortes_txt} • Religues: {delta_religues_txt})<br>
            <b>📊 Total atual do dia:</b> {numero(status.get("notas_hoje", 0))} notas
            (Cortes: {numero(status.get("cortes_hoje", 0))} • Religues: {numero(status.get("religues_hoje", 0))})<br>
            <b>📦 Base geral:</b> {numero(status.get("total_atual", 0))} notas acumuladas<br>
            <b>📌 {texto_contrato}:</b> {detalhe}{aviso_snapshot}
        </div>
        """,
        unsafe_allow_html=True,
    )


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

if st.sidebar.button("🔄 Atualizar dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()


contratos_lista = []

for base in [contratos_original, dias_original, carro_original, carro_dias_original]:
    if not base.empty and "CONTRATO" in base.columns:
        contratos_lista += base["CONTRATO"].dropna().unique().tolist()

contratos_lista = sorted(set(contratos_lista))

if "contrato_escolhido" not in st.session_state:
    st.session_state.contrato_escolhido = "Todos"

if "modo_painel" not in st.session_state:
    st.session_state.modo_painel = "corte"

st.sidebar.markdown("### Contratos - Corte")

if st.sidebar.button("📊 Todos", use_container_width=True):
    st.session_state.modo_painel = "corte"
    st.session_state.contrato_escolhido = "Todos"

for contrato_nome in contratos_lista:
    if st.sidebar.button(f"🔹 {contrato_nome}", use_container_width=True):
        st.session_state.modo_painel = "corte"
        st.session_state.contrato_escolhido = contrato_nome

st.sidebar.markdown("---")
st.sidebar.markdown("### Contratos - Leitura (em testes)")

if st.sidebar.button("📖 Americana e Piracicaba", use_container_width=True):
    st.session_state.modo_painel = "leitura"

modo_painel = st.session_state.modo_painel
contrato_escolhido = st.session_state.contrato_escolhido

st.sidebar.markdown("---")
st.sidebar.markdown("**Tela selecionada:**")
if modo_painel == "leitura":
    st.sidebar.info("Leitura (em testes)")
else:
    st.sidebar.info(contrato_escolhido)

if modo_painel == "corte" and not notas.empty:
    mostrar_status_atualizacao(notas, contrato_escolhido)

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

if modo_painel == "leitura":
    st.subheader("📖 Contrato Leitura (em testes)")
    st.caption("Parciais separadas de Americana e Piracicaba, lidas dos Excels gerados pelo extrator CWSI.")

    mostrar_base_leitura("Americana")
    st.markdown("---")
    mostrar_base_leitura("Piracicaba")
    st.stop()

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
            "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "EXPRESS",
            "FATURAMENTO", "FATURAMENTO_EXPRESS", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
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
    st.subheader("Parcial do dia por recurso")

    # Base com recusas para mostrar na parcial.
    parcial_com_recusas = preparar_parcial_do_dia(notas, incluir_recusas=True)

    if parcial_com_recusas.empty:
        st.info("Ainda não há dados suficientes para montar a parcial do dia.")
    else:
        if contrato_escolhido != "Todos" and "CONTRATO" in parcial_com_recusas.columns:
            parcial_com_recusas = parcial_com_recusas[parcial_com_recusas["CONTRATO"] == contrato_escolhido]

        datas_disponiveis = (
            parcial_com_recusas[["DATA", "DATA_DT"]]
            .drop_duplicates()
            .sort_values("DATA_DT", ascending=False)
        )

        if datas_disponiveis.empty:
            st.info("Nenhuma data encontrada na base de notas.")
        else:
            opcoes_datas = datas_disponiveis["DATA"].tolist()
            data_escolhida = st.selectbox("Escolha o dia", opcoes_datas, index=0)

            parcial_dia_tudo = parcial_com_recusas[parcial_com_recusas["DATA"] == data_escolhida].copy()
            parcial_dia = parcial_dia_tudo[parcial_dia_tudo["EH_RECUSA"] == 0].copy()
            recusas_dia = parcial_dia_tudo[parcial_dia_tudo["EH_RECUSA"] == 1].copy()

            if parcial_dia_tudo.empty:
                st.info("Nenhuma nota encontrada para esse dia.")
            else:
                total_notas = parcial_dia["ORDEM_DE_SERVICO"].nunique() if not parcial_dia.empty else 0
                total_recursos_ativos = parcial_dia["RECURSO"].nunique() if not parcial_dia.empty else 0
                total_cortes = int(parcial_dia["EH_CORTE"].sum()) if not parcial_dia.empty else 0
                total_religues = int(parcial_dia["EH_RELIGUE"].sum()) if not parcial_dia.empty else 0
                total_recusas = recusas_dia["ORDEM_DE_SERVICO"].nunique() if not recusas_dia.empty else 0
                total_faturamento = parcial_dia["FATURAMENTO"].sum() if not parcial_dia.empty else 0
                total_faturamento_min = parcial_dia["FATURAMENTO_MIN"].sum() if not parcial_dia.empty else 0
                total_faturamento_max = parcial_dia["FATURAMENTO_MAX"].sum() if not parcial_dia.empty else 0

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Recursos ativos", numero(total_recursos_ativos))
                c2.metric("Notas feitas", numero(total_notas))
                c3.metric("Cortes", numero(total_cortes))
                c4.metric("Religues", numero(total_religues))
                c5.metric("Recusas", numero(total_recusas))

                tem_carro_no_dia = "CONTRATO" in parcial_dia.columns and (
                    parcial_dia["CONTRATO"] == "Contrato Carro STC estimado"
                ).any()

                if tem_carro_no_dia:
                    st.metric("Faturamento estimado", f"{dinheiro(total_faturamento_min)} a {dinheiro(total_faturamento_max)}")
                else:
                    st.metric("Faturamento", dinheiro(total_faturamento))

                st.markdown('<div class="section-title">Ranking do dia por produção</div>', unsafe_allow_html=True)

                resumo_producao = (
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
                    if not parcial_dia.empty
                    else pd.DataFrame(columns=[
                        "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES",
                        "FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
                    ])
                )

                resumo_recusas_por_recurso = (
                    recusas_dia.groupby(["RECURSO", "CONTRATO"], dropna=False)
                    .agg(RECUSAS=("ORDEM_DE_SERVICO", "nunique"))
                    .reset_index()
                    if not recusas_dia.empty
                    else pd.DataFrame(columns=["RECURSO", "CONTRATO", "RECUSAS"])
                )

                resumo_equipe = resumo_producao.merge(
                    resumo_recusas_por_recurso,
                    on=["RECURSO", "CONTRATO"],
                    how="outer",
                ).fillna({
                    "TOTAL_NOTAS": 0,
                    "CORTES": 0,
                    "RELIGUES": 0,
                    "FATURAMENTO": 0,
                    "FATURAMENTO_MIN": 0,
                    "FATURAMENTO_MAX": 0,
                    "RECUSAS": 0,
                })

                for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES", "RECUSAS"]:
                    if col in resumo_equipe.columns:
                        resumo_equipe[col] = pd.to_numeric(resumo_equipe[col], errors="coerce").fillna(0).astype(int)

                for col in ["FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
                    if col in resumo_equipe.columns:
                        resumo_equipe[col] = pd.to_numeric(resumo_equipe[col], errors="coerce").fillna(0)

                resumo_equipe = (
                    resumo_equipe
                    .sort_values(["TOTAL_NOTAS", "FATURAMENTO", "RECUSAS"], ascending=[False, False, False])
                    .reset_index(drop=True)
                )

                if resumo_equipe.empty:
                    st.info("Nenhuma nota ou recusa encontrada para esse dia.")
                else:
                    resumo_equipe.insert(0, "POSIÇÃO", range(1, len(resumo_equipe) + 1))

                    recursos_zero = resumo_equipe[
                        (resumo_equipe["TOTAL_NOTAS"] == 0) & (resumo_equipe["RECUSAS"] > 0)
                    ].copy()

                    if not recursos_zero.empty:
                        lista_zero = ", ".join(recursos_zero["RECURSO"].astype(str).head(12).tolist())
                        if len(recursos_zero) > 12:
                            lista_zero += f" e mais {len(recursos_zero) - 12}"
                        st.markdown(
                            f"""
                            <div class="zero-card">
                                ⚠️ <b>Recursos com zero produção no dia:</b> {numero(len(recursos_zero))}<br>
                                {lista_zero}
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

                    top10_dia = resumo_equipe.head(10).copy()

                    grafico_parcial = (
                        alt.Chart(top10_dia)
                        .mark_bar(
                            cornerRadiusTopLeft=8,
                            cornerRadiusTopRight=8,
                        )
                        .encode(
                            x=alt.X(
                                "RECURSO:N",
                                sort=alt.SortField(field="TOTAL_NOTAS", order="descending"),
                                title="Recurso",
                                axis=alt.Axis(labelAngle=-90),
                            ),
                            y=alt.Y("TOTAL_NOTAS:Q", title="Notas feitas"),
                            tooltip=[
                                alt.Tooltip("POSIÇÃO:Q", title="Posição"),
                                alt.Tooltip("RECURSO:N", title="Recurso"),
                                alt.Tooltip("TOTAL_NOTAS:Q", title="Notas feitas"),
                                alt.Tooltip("CORTES:Q", title="Cortes"),
                                alt.Tooltip("RELIGUES:Q", title="Religues"),
                                alt.Tooltip("RECUSAS:Q", title="Recusas"),
                                alt.Tooltip("FATURAMENTO:Q", title="Faturamento", format=",.2f"),
                            ],
                        )
                        .properties(height=330)
                    )

                    st.altair_chart(grafico_parcial, use_container_width=True)

                    def faturamento_linha_equipe(row):
                        if row.get("CONTRATO") == "Contrato Carro STC estimado":
                            return f"{dinheiro(row.get('FATURAMENTO_MIN', 0))} a {dinheiro(row.get('FATURAMENTO_MAX', 0))}"
                        return dinheiro(row.get("FATURAMENTO", 0))

                    tabela_equipe = resumo_equipe.copy()
                    tabela_equipe["FATURAMENTO"] = tabela_equipe.apply(faturamento_linha_equipe, axis=1)
                    tabela_equipe = tabela_equipe[[
                        "POSIÇÃO", "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "RECUSAS", "FATURAMENTO"
                    ]]

                    st.dataframe(formatar_tabela(tabela_equipe), use_container_width=True, hide_index=True)

                st.markdown('<div class="section-title">Recusas do dia</div>', unsafe_allow_html=True)

                if recusas_dia.empty:
                    st.success("Nenhuma recusa encontrada para esse dia.")
                else:
                    with st.expander("Ver detalhes das recusas", expanded=True):
                        colunas_recusa = [
                            "ORDEM_DE_SERVICO", "RECURSO", "CONTRATO", "GRUPO_NOTA",
                            "RECUSA", "DATA", "ELETRICISTA1", "ELETRICISTA2"
                        ]
                        colunas_recusa = [c for c in colunas_recusa if c in recusas_dia.columns]
                        st.dataframe(
                            recusas_dia[colunas_recusa].sort_values(["RECURSO", "ORDEM_DE_SERVICO"]),
                            use_container_width=True,
                            hide_index=True,
                        )

                st.markdown('<div class="section-title">Detalhamento das notas feitas no dia</div>', unsafe_allow_html=True)
                colunas_detalhe = [
                    "ORDEM_DE_SERVICO", "RECURSO", "CONTRATO", "GRUPO_NOTA", "DATA", "ELETRICISTA1", "ELETRICISTA2"
                ]
                colunas_detalhe = [c for c in colunas_detalhe if c in parcial_dia.columns]
                if parcial_dia.empty:
                    st.info("Nenhuma nota feita para detalhar.")
                else:
                    st.dataframe(
                        parcial_dia[colunas_detalhe].sort_values(["RECURSO", "ORDEM_DE_SERVICO"]),
                        use_container_width=True,
                        hide_index=True,
                    )


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

        express_data_max = ""
        express_resumo_recurso = pd.DataFrame()
        express_sem_vinculo = pd.DataFrame()
        express_caminho = ""
        total_express_mensal = 0
        fat_express_mensal = 0.0

        if tipo_periodo == "Mês" and valor_periodo:
            (
                ranking_exec,
                express_resumo_recurso,
                express_data_max,
                express_sem_vinculo,
                express_caminho,
                total_express_mensal,
                fat_express_mensal,
            ) = aplicar_express_no_ranking_mensal(
                ranking_exec,
                notas,
                valor_periodo,
                contrato_ranking,
            )

        if ranking_exec.empty:
            st.info("Nenhum recurso encontrado para os filtros selecionados.")
        else:
            total_notas_exec = int(ranking_exec["NOTAS"].sum()) if "NOTAS" in ranking_exec.columns else int(base_filtrada_exec["ORDEM_DE_SERVICO"].nunique())
            total_executores = int(ranking_exec["RECURSO"].nunique())
            total_fat_atribuido = float(ranking_exec["FATURAMENTO_ATRIBUÍDO"].sum())

            if tipo_periodo == "Mês" and valor_periodo:
                if express_caminho:
                    if express_data_max:
                        st.info(f"Pagamento Express conciliado por DE/PARA Nome → Recurso até {express_data_max}.")
                    else:
                        st.info("Pagamento Express conciliado por DE/PARA Nome → Recurso. A planilha não trouxe data válida para exibir o limite.")
                else:
                    st.caption("Pagamento Express: arquivo não localizado.")

                if not express_sem_vinculo.empty:
                    st.warning(f"Pagamento Express: {numero(len(express_sem_vinculo))} linha(s) não encontraram nome no DE/PARA Nome → Recurso.")

                # Auditoria sempre visível no período mensal: se zerar, a tela mostra exatamente o motivo.
                with st.expander("Auditoria do Pagamento Express", expanded=(total_express_mensal == 0)):
                    if express_caminho:
                        st.caption(f"Arquivo lido: {express_caminho}")
                    if not express_resumo_recurso.empty:
                        st.success(f"Express conciliado: {numero(total_express_mensal)} nota(s) no mês {valor_periodo}.")
                        st.dataframe(
                            formatar_tabela(express_resumo_recurso.sort_values(["EXPRESS", "RECURSO"], ascending=[False, True])),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.info("Nenhum Express entrou no ranking para este filtro. Verifique abaixo se o arquivo foi lido, se a data bate com o mês e se os nomes estão no DE/PARA.")
                        caminho_debug = caminho_pagamento_express()
                        if caminho_debug:
                            express_debug = ler_pagamento_express(str(caminho_debug))
                            if express_debug.empty:
                                st.warning("O arquivo foi encontrado, mas ficou vazio após o filtro de VALIDAÇÃO = PAGAMENTO EXPRESS ou sem nome de executor.")
                            else:
                                total_linhas_debug = len(express_debug)
                                datas_validas_debug = int(express_debug.get("DATA_EXPRESS_DT", pd.Series(dtype=object)).notna().sum()) if "DATA_EXPRESS_DT" in express_debug.columns else 0
                                st.write({
                                    "linhas_lidas": total_linhas_debug,
                                    "datas_validas": datas_validas_debug,
                                    "meses_no_excel": express_debug["DATA_EXPRESS_DT"].dt.strftime("%m/%Y").value_counts(dropna=False).to_dict() if "DATA_EXPRESS_DT" in express_debug.columns and express_debug["DATA_EXPRESS_DT"].notna().any() else {},
                                    "nomes_mapeados": int(express_debug.get("NOME_EXPRESS_NORM", pd.Series(dtype=object)).map(DEPARA_NOME_RECURSO_EXPRESS).fillna("").ne("").sum()) if "NOME_EXPRESS_NORM" in express_debug.columns else 0,
                                })
                                cols_debug = [
                                    "NOME_EXPRESS", "NOME_EXPRESS_NORM", "DATA_EXPRESS_DT", "NOTA_NORM", "VALIDAÇÃO", "VALIDACAO"
                                ]
                                cols_debug = [c for c in cols_debug if c in express_debug.columns]
                                amostra_debug = express_debug.copy()
                                if "DATA_EXPRESS_DT" in amostra_debug.columns and amostra_debug["DATA_EXPRESS_DT"].notna().any():
                                    amostra_debug = amostra_debug[amostra_debug["DATA_EXPRESS_DT"].dt.strftime("%m/%Y") == valor_periodo].copy()
                                amostra_debug["RECURSO_DEPARA"] = amostra_debug.get("NOME_EXPRESS_NORM", pd.Series(dtype=object)).map(DEPARA_NOME_RECURSO_EXPRESS).fillna("") if "NOME_EXPRESS_NORM" in amostra_debug.columns else ""
                                cols_debug = cols_debug + ["RECURSO_DEPARA"]
                                st.dataframe(amostra_debug[cols_debug].head(80), use_container_width=True, hide_index=True)
                        else:
                            st.error("Arquivo pagamento_express.xlsx não localizado na pasta dashboard nem na raiz do app.")

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
            if tipo_periodo == "Mês" and valor_periodo:
                m4.metric("Express", numero(total_express_mensal))
            else:
                m4.metric("Média notas/recurso", f"{media_notas_executor:.1f}".replace(".", ","))

            if tipo_periodo == "Mês" and valor_periodo and not express_resumo_recurso.empty:
                st.markdown('<div class="section-title">Pagamento Express conciliado por recurso</div>', unsafe_allow_html=True)

                grafico_express_recurso = (
                    alt.Chart(express_resumo_recurso.head(15))
                    .mark_bar(
                        cornerRadiusTopLeft=8,
                        cornerRadiusTopRight=8,
                    )
                    .encode(
                        x=alt.X(
                            "RECURSO:N",
                            sort=alt.SortField(field="EXPRESS", order="descending"),
                            title="Recurso",
                            axis=alt.Axis(labelAngle=-90),
                        ),
                        y=alt.Y("EXPRESS:Q", title="Express"),
                        tooltip=[
                            alt.Tooltip("RECURSO:N", title="Recurso"),
                            alt.Tooltip("CONTRATO:N", title="Contrato"),
                            alt.Tooltip("EXPRESS:Q", title="Notas Express"),
                            alt.Tooltip("FATURAMENTO_EXPRESS:Q", title="Faturamento Express", format=",.2f"),
                        ],
                    )
                    .properties(height=360)
                )

                st.altair_chart(grafico_express_recurso, use_container_width=True)

                tabela_express_recurso = express_resumo_recurso.copy().sort_values(
                    ["EXPRESS", "FATURAMENTO_EXPRESS"], ascending=False
                )
                st.dataframe(
                    formatar_tabela(tabela_express_recurso[["RECURSO", "CONTRATO", "EXPRESS", "FATURAMENTO_EXPRESS"]]),
                    use_container_width=True,
                    hide_index=True,
                )

            if tipo_periodo == "Mês" and valor_periodo and not express_sem_vinculo.empty:
                with st.expander("Ver Express sem vínculo de Ordem de Serviço"):
                    cols_sem_vinculo = [
                        "NOTA", "NOTA_NORM", "DATA_EXPRESS_DT", "VALIDAÇÃO", "VALIDACAO",
                        "NOME_EXECUTOR_01", "NOME_EXECUTOR_02", "NOME_EXECUTOR", "EXECUTOR"
                    ]
                    cols_sem_vinculo = [c for c in cols_sem_vinculo if c in express_sem_vinculo.columns]
                    st.dataframe(
                        express_sem_vinculo[cols_sem_vinculo],
                        use_container_width=True,
                        hide_index=True,
                    )

            st.markdown('<div class="section-title">Top 10 recursos</div>', unsafe_allow_html=True)
            top10 = ranking_exec.head(10).copy()
            coluna_grafico = "NOTAS" if criterio == "Notas" else "FATURAMENTO_ATRIBUÍDO"
            titulo_eixo_y = "Notas" if criterio == "Notas" else "Faturamento atribuído"

            grafico_top10 = (
                alt.Chart(top10)
                .mark_bar(
                    cornerRadiusTopLeft=8,
                    cornerRadiusTopRight=8,
                )
                .encode(
                    x=alt.X(
                        "RECURSO:N",
                        sort=alt.SortField(field=coluna_grafico, order="descending"),
                        title="Recurso",
                        axis=alt.Axis(labelAngle=-90),
                    ),
                    y=alt.Y(
                        f"{coluna_grafico}:Q",
                        title=titulo_eixo_y,
                    ),
                    tooltip=[
                        alt.Tooltip("POSIÇÃO:Q", title="Posição"),
                        alt.Tooltip("RECURSO:N", title="Recurso"),
                        alt.Tooltip("NOTAS:Q", title="Notas"),
                        alt.Tooltip("FATURAMENTO_ATRIBUÍDO:Q", title="Faturamento atribuído", format=",.2f"),
                    ],
                )
                .properties(height=360)
            )

            st.altair_chart(grafico_top10, use_container_width=True)

            st.markdown('<div class="section-title">Pódio</div>', unsafe_allow_html=True)
            mostrar_podio_ranking(ranking_exec, nome_coluna="RECURSO")

            st.markdown('<div class="section-title">Ranking detalhado</div>', unsafe_allow_html=True)
            colunas_ranking = [
                "POSIÇÃO", "RECURSO", "NOTAS", "CORTES", "RELIGUES", "EXPRESS", "DIAS_ATIVOS",
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
    st.caption("Compara o mês escolhido com o mês anterior, somando Pagamento Express pelo mês de referência.")

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

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Faturamento", dinheiro(atual["FATURAMENTO"]), variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"]))
        c2.metric("Notas", numero(atual["TOTAL_NOTAS"]), variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"]))
        c3.metric("Cortes", numero(atual["CORTES"]), variacao_percentual(atual["CORTES"], anterior["CORTES"]))
        c4.metric("Religues", numero(atual["RELIGUES"]), variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"]))
        c5.metric("Express", numero(atual.get("EXPRESS", 0)), variacao_percentual(atual.get("EXPRESS", 0), anterior.get("EXPRESS", 0)))

        tabela_comparativo = pd.DataFrame([
            {"Indicador": "Faturamento", mes_escolhido: dinheiro(atual["FATURAMENTO"]), mes_anterior: dinheiro(anterior["FATURAMENTO"]), "Variação": variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"])},
            {"Indicador": "Notas", mes_escolhido: numero(atual["TOTAL_NOTAS"]), mes_anterior: numero(anterior["TOTAL_NOTAS"]), "Variação": variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"])},
            {"Indicador": "Cortes", mes_escolhido: numero(atual["CORTES"]), mes_anterior: numero(anterior["CORTES"]), "Variação": variacao_percentual(atual["CORTES"], anterior["CORTES"])},
            {"Indicador": "Religues", mes_escolhido: numero(atual["RELIGUES"]), mes_anterior: numero(anterior["RELIGUES"]), "Variação": variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"])},
            {"Indicador": "Express", mes_escolhido: numero(atual.get("EXPRESS", 0)), mes_anterior: numero(anterior.get("EXPRESS", 0)), "Variação": variacao_percentual(atual.get("EXPRESS", 0), anterior.get("EXPRESS", 0))},
            {"Indicador": "Faturamento Express", mes_escolhido: dinheiro(atual.get("FATURAMENTO_EXPRESS", 0)), mes_anterior: dinheiro(anterior.get("FATURAMENTO_EXPRESS", 0)), "Variação": variacao_percentual(atual.get("FATURAMENTO_EXPRESS", 0), anterior.get("FATURAMENTO_EXPRESS", 0))},
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
                "EXPRESS": r.get("EXPRESS", 0),
                "FATURAMENTO_EXPRESS": r.get("FATURAMENTO_EXPRESS", 0),
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
            colunas = ["CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "EXPRESS", "FATURAMENTO", "FATURAMENTO_EXPRESS", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]
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
