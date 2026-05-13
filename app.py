from pathlib import Path
import os
from collections.abc import Mapping
from datetime import datetime
from zoneinfo import ZoneInfo
import json
import html
import tempfile
import sqlite3
import unicodedata
import pandas as pd
import streamlit as st
import altair as alt
import streamlit.components.v1 as components

st.set_page_config(page_title="G.Z.U.S. | Gestão Inteligente de Serviços", page_icon="🤖", layout="wide")

# ==============================
# SEGREDOS / CONFIGURAÇÕES PRIVADAS
# ==============================
# Mantém o repositório público sem expor senhas, caminhos locais, tarifas
# e mapas manuais. Configure estes valores em .streamlit/secrets.toml
# no ambiente local ou em App settings > Secrets no Streamlit Cloud.

def secret_value(nome, padrao=None):
    try:
        if nome in st.secrets:
            return st.secrets[nome]
    except Exception:
        pass
    return os.getenv(nome, padrao)


def secret_list(nome, padrao=None):
    valor = secret_value(nome, padrao if padrao is not None else [])
    if valor is None:
        return []
    if isinstance(valor, (list, tuple)):
        return list(valor)
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return []
        try:
            parsed = json.loads(texto)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            pass
        return [v.strip() for v in texto.split(',') if v.strip()]
    return list(valor) if hasattr(valor, '__iter__') else []


def secret_dict(nome, padrao=None):
    valor = secret_value(nome, padrao if padrao is not None else {})
    if valor is None:
        return {}
    if isinstance(valor, Mapping):
        return dict(valor)
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return {}
        try:
            parsed = json.loads(texto)
            return parsed if isinstance(parsed, dict) else {}
        except Exception:
            return {}
    return {}


def secret_float(nome, padrao=0.0):
    valor = secret_value(nome, padrao)
    try:
        return float(valor)
    except Exception:
        return float(padrao)


SENHA_CORRETA = secret_value("SENHA_GERENTE", "")

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

# ==============================
# CONTROLE DE ACESSO POR PERFIL
# ==============================

USUARIOS_ACESSO = {
    "gerente": {
        "senha": secret_value("SENHA_GERENTE", ""),
        "perfil": "gerente",
        "nome": secret_value("NOME_GERENTE", "Gerente"),
    },
    "supervisor_stc": {
        "senha": secret_value("SENHA_STC", ""),
        "perfil": "supervisor_stc",
        "nome": secret_value("NOME_SUPERVISOR_STC", "Supervisor STC"),
    },
    "supervisor_leitura": {
        "senha": secret_value("SENHA_LEITURA", ""),
        "perfil": "supervisor_leitura",
        "nome": secret_value("NOME_SUPERVISOR_LEITURA", "Supervisor Leitura"),
    },
}

CONTRATOS_SUPERVISOR_STC = secret_list("CONTRATOS_SUPERVISOR_STC", ["STC Jundiai", "Disjuntor Santa Cruz"])


# ==============================
# AJUSTE MOBILE / RESPONSIVIDADE
# ==============================
st.markdown(
    """
    <style>
    @media (max-width: 768px) {
        html, body, .stApp { background: #0b1220 !important; color: #e5e7eb !important; }
        .main .block-container {
            padding-top: 0.75rem !important;
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
            padding-bottom: 7.5rem !important;
            max-width: 100% !important;
        }
        h1 {
            font-size: 1.55rem !important;
            line-height: 1.15 !important;
            margin-bottom: 0.35rem !important;
            color: #f8fafc !important;
        }
        h2 { font-size: 1.28rem !important; color: #f8fafc !important; }
        h3 { font-size: 1.08rem !important; color: #f8fafc !important; }
        .stCaption, div[data-testid="stCaptionContainer"] {
            color: #cbd5e1 !important;
            font-size: 0.88rem !important;
        }
        div[data-testid="stMetric"], .ranking-podium, .status-card, .zero-card {
            background: rgba(15, 23, 42, 0.96) !important;
            border-color: rgba(148, 163, 184, 0.28) !important;
            box-shadow: 0 10px 28px rgba(0,0,0,0.28) !important;
        }
        div[data-testid="stMetric"] label, div[data-testid="stMetricValue"] {
            color: #f8fafc !important;
        }
        .executive-card {
            background: linear-gradient(135deg, #111827 0%, #1d4ed8 100%) !important;
            border-radius: 20px !important;
            padding: 16px 18px !important;
            margin: 8px 0 14px 0 !important;
            box-shadow: 0 12px 30px rgba(29, 78, 216, 0.22) !important;
        }
        .executive-card h3, .executive-card p, .executive-card div, .executive-card span {
            color: #ffffff !important;
        }
        section[data-testid="stSidebar"] {
            background: #0f172a !important;
        }
        div[data-testid="stDataFrame"] {
            max-width: 100% !important;
            overflow-x: auto !important;
        }
        button[kind="secondary"], button[data-baseweb="tab"] {
            min-height: 40px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "perfil_acesso" not in st.session_state:
    st.session_state.perfil_acesso = ""

if "nome_acesso" not in st.session_state:
    st.session_state.nome_acesso = ""

if not st.session_state.autenticado:
    # Login em uma única passagem: antes usava st.rerun() depois de validar a senha.
    # Isso fazia o Streamlit executar o app duas vezes no clique de Entrar e aumentava
    # bastante a sensação de demora pós-login. Agora, ao acertar a senha, o código
    # continua no mesmo ciclo e já monta o painel.
    st.markdown("""
    <style>
    section[data-testid="stSidebar"] {display: none !important;}
    .main .block-container {max-width: 760px; padding-top: 7rem;}
    </style>
    """, unsafe_allow_html=True)

    login_box = st.empty()
    with login_box.container():
        st.title("🔒 Acesso restrito")
        st.caption("Entre com o perfil autorizado para acessar o painel.")

        usuario = st.selectbox(
            "Perfil",
            options=list(USUARIOS_ACESSO.keys()),
            format_func=lambda u: USUARIOS_ACESSO[u]["nome"],
        )
        senha = st.text_input("Digite a senha para acessar o painel:", type="password")

        entrar = st.button("Entrar")

    if entrar:
        dados_usuario = USUARIOS_ACESSO.get(usuario, {})
        if senha == dados_usuario.get("senha"):
            st.session_state.autenticado = True
            st.session_state.perfil_acesso = dados_usuario.get("perfil", "")
            st.session_state.nome_acesso = dados_usuario.get("nome", "")
            st.session_state.login_recente_sem_rerun = True
            login_box.empty()
        else:
            st.error("Senha incorreta")
            st.stop()
    else:
        st.stop()

PERFIL_ACESSO = st.session_state.get("perfil_acesso", "gerente")
NOME_ACESSO = st.session_state.get("nome_acesso", "Gerente")
PODE_VER_FINANCEIRO = PERFIL_ACESSO == "gerente"

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
# SQLITE / BANCO LOCAL OPCIONAL
# ==============================
# O painel continua funcionando com CSV/Excel.
# Prioridade: banco ultraleve do Streamlit (gzus_dashboard.db).
# Se algo falhar ou o banco não tiver a tabela esperada, volta automaticamente para os CSVs.
BANCO_GZUS_CANDIDATOS = [
    PASTA_DASHBOARD / "gzus_dashboard.db",
    PASTA_ATUAL / "gzus_dashboard.db",
    PASTA_DASHBOARD / "gzus.db",
    PASTA_ATUAL / "gzus.db",
]

TABELAS_SQLITE_DASHBOARD = {
    "notas": "notas",
    "contratos": "faturamento_contratos",
    "dias": "faturamento_dias",
    "carro": "faturamento_carro_estimado",
    "carro_dias": "faturamento_carro_dias",
}


# ==============================
# CONTRATO LEITURA (em testes)
# ==============================
# No PC da operação, o extrator de leitura grava aqui.
# No Streamlit Cloud, o app só consegue ler se os arquivos forem enviados ao GitHub,
# preferencialmente em dashboard/leitura/.
PASTAS_LEITURA = [Path(p) for p in secret_list("PASTAS_LEITURA_PRIVADAS", [])] + [
    PASTA_DASHBOARD / "leitura",
    PASTA_DASHBOARD,
    PASTA_ATUAL / "leitura",
    PASTA_ATUAL,
]

ARQUIVOS_LEITURA = {
    # Formato antigo: Parcial_Americana*.xlsx / Parcial_Piracicaba*.xlsx
    # Formato novo do extrator de tarefas: Tarefas_Americana*.xlsx / Tarefas_Piracicaba*.xlsx
    # Consolidado novo: Resumo_D_por_base_municipio*.xlsx
    "Americana": [
        "Tarefas_Americana*.xlsx",
        "Parcial_Americana.xlsx",
        "Parcial_Americana*.xlsx",
        "Resumo_D_por_base_municipio*.xlsx",
    ],
    "Piracicaba": [
        "Tarefas_Piracicaba*.xlsx",
        "Parcial_Piracicaba.xlsx",
        "Parcial_Piracicaba*.xlsx",
        "Resumo_D_por_base_municipio*.xlsx",
    ],
}

ORDEM_DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# Atualização mais curta para evitar que o Streamlit fique preso em dados antigos.
# O cache continua existindo para não pesar o app, mas agora é revalidado com frequência.
CACHE_TTL_SEGUNDOS = int(secret_float("CACHE_TTL_SEGUNDOS", 180))
CACHE_TTL_RANKING_SEGUNDOS = int(secret_float("CACHE_TTL_RANKING_SEGUNDOS", 600))
# Antes estava em 60s. Isso deixava o pós-login e cada rerun mais sujeito a
# esperar git fetch. Para operação, 5 minutos costuma ser um equilíbrio melhor:
# rápido para atualizar, mas sem travar a navegação o tempo todo.
GITHUB_SYNC_INTERVALO_SEGUNDOS = int(secret_float("GITHUB_SYNC_INTERVALO_SEGUNDOS", 300))
STATUS_GITHUB_SYNC_PATH = Path(tempfile.gettempdir()) / "gzus_github_sync_status.json"
LEITURA_HABILITADA = False  # Removido temporariamente para deixar o painel principal mais leve.

# Recarregamento automático por JavaScript removido para evitar travamentos e tela piscando.
# Use o botão "Atualizar dados" quando quiser puxar o GitHub manualmente.


def _github_sync_habilitado():
    valor = str(secret_value("SINCRONIZAR_GITHUB_AUTO", "true") or "true").strip().lower()
    return valor not in ["0", "false", "nao", "não", "no", "off"]


def _ler_status_github_sync():
    try:
        if STATUS_GITHUB_SYNC_PATH.exists():
            return json.loads(STATUS_GITHUB_SYNC_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _salvar_status_github_sync(status):
    try:
        STATUS_GITHUB_SYNC_PATH.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _executar_git(args):
    import subprocess
    return subprocess.run(
        ["git", "-C", str(PASTA_ATUAL), *args],
        capture_output=True,
        text=True,
        timeout=30,
    )


def _branch_git_atual():
    branch_secret = str(secret_value("GITHUB_SYNC_BRANCH", "") or "").strip()
    if branch_secret:
        return branch_secret

    for env_nome in ["STREAMLIT_GIT_BRANCH", "GITHUB_BRANCH", "GIT_BRANCH"]:
        branch_env = str(os.getenv(env_nome, "") or "").strip()
        if branch_env:
            return branch_env.replace("origin/", "")

    try:
        r = _executar_git(["rev-parse", "--abbrev-ref", "HEAD"])
        branch = (r.stdout or "").strip()
        if branch and branch != "HEAD":
            return branch
    except Exception:
        pass

    return "main"


def sincronizar_github_se_preciso(forcar=False):
    """Puxa atualizações do GitHub sem precisar rebootar o app no Streamlit Cloud.

    Por que isso existe:
    - O extrator sobe CSV/Excel no GitHub.
    - Em alguns deploys do Streamlit Cloud, o app continua lendo o checkout antigo
      até reboot/redeploy.
    - Esta rotina faz git fetch + reset para a branch atual e limpa o cache quando
      detecta commit novo.

    Para desligar: coloque SINCRONIZAR_GITHUB_AUTO = "false" nos Secrets.
    Para escolher branch: coloque GITHUB_SYNC_BRANCH = "main" ou o nome usado.
    """
    if not _github_sync_habilitado() and not forcar:
        return {"ok": False, "changed": False, "skipped": True, "message": "Sincronização automática desligada"}

    agora = datetime.now(ZoneInfo("America/Sao_Paulo"))
    status_antigo = _ler_status_github_sync()
    ultimo_ts = float(status_antigo.get("timestamp", 0) or 0)

    if not forcar and ultimo_ts and (agora.timestamp() - ultimo_ts) < GITHUB_SYNC_INTERVALO_SEGUNDOS:
        status_antigo["skipped"] = True
        return status_antigo

    status = {
        "ok": False,
        "changed": False,
        "skipped": False,
        "timestamp": agora.timestamp(),
        "quando": agora.strftime("%d/%m/%Y %H:%M:%S"),
    }

    try:
        dentro_repo = _executar_git(["rev-parse", "--is-inside-work-tree"])
        if dentro_repo.returncode != 0:
            status["message"] = "Este ambiente não parece ser um repositório Git."
            _salvar_status_github_sync(status)
            return status

        branch = _branch_git_atual()
        status["branch"] = branch

        antes = _executar_git(["rev-parse", "HEAD"])
        commit_antes = (antes.stdout or "").strip()
        status["commit_antes"] = commit_antes[:12]

        fetch = _executar_git(["fetch", "origin", branch])
        if fetch.returncode != 0:
            # Fallback para ambientes onde o fetch por branch falha, mas fetch geral funciona.
            fetch = _executar_git(["fetch", "origin"])
        if fetch.returncode != 0:
            status["message"] = (fetch.stderr or fetch.stdout or "Erro ao executar git fetch.").strip()[-500:]
            _salvar_status_github_sync(status)
            return status

        remoto = _executar_git(["rev-parse", f"origin/{branch}"])
        if remoto.returncode != 0:
            status["message"] = (remoto.stderr or remoto.stdout or f"Não encontrei origin/{branch}.").strip()[-500:]
            _salvar_status_github_sync(status)
            return status

        commit_remoto = (remoto.stdout or "").strip()
        status["commit_remoto"] = commit_remoto[:12]

        if commit_remoto and commit_remoto != commit_antes:
            reset = _executar_git(["reset", "--hard", f"origin/{branch}"])
            if reset.returncode != 0:
                status["message"] = (reset.stderr or reset.stdout or "Erro ao executar git reset.").strip()[-500:]
                _salvar_status_github_sync(status)
                return status

            st.cache_data.clear()
            status["changed"] = True
            status["ok"] = True
            status["message"] = "Atualização puxada do GitHub e cache limpo."
            _salvar_status_github_sync(status)
            return status

        status["ok"] = True
        status["message"] = "Sem commit novo no GitHub."
        _salvar_status_github_sync(status)
        return status

    except Exception as e:
        status["message"] = f"Erro na sincronização GitHub: {e}"
        _salvar_status_github_sync(status)
        return status


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


def _data_chave_arquivo_leitura(caminho):
    """Extrai a data operacional do nome do arquivo, quando existir.

    Exemplos aceitos:
    - Tarefas_Americana_2026-04-27_20260501_232238.xlsx -> 2026-04-27
    - Tarefas_Piracicaba_2026-04-29.xlsx -> 2026-04-29
    - Parcial_Americana_20260429_193337.xlsx -> 20260429
    """
    import re
    nome = Path(caminho).stem
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", nome)
    if m:
        return m.group(1)
    m = re.search(r"_(20\d{6})(?:_|$)", nome)
    if m:
        return m.group(1)
    return nome


def _tipo_arquivo_leitura(caminho):
    nome = Path(caminho).name.upper()
    if nome.startswith("TAREFAS_"):
        return "TAREFAS"
    if nome.startswith("RESUMO_D"):
        return "RESUMO_CONSOLIDADO"
    if nome.startswith("PARCIAL_"):
        return "PARCIAL"
    return "OUTRO"


def caminhos_leitura(base_nome):
    """Localiza TODOS os arquivos de leitura úteis para a base.

    Antes o painel pegava apenas o arquivo mais recente por base. Isso fazia sumir D1/D3
    e também podia esconder Piracicaba. Agora ele carrega todos os dias encontrados,
    mantendo só a versão mais recente quando houver mais de um arquivo do mesmo dia.
    """
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

    # Remove duplicados e mantém apenas Excel.
    unicos = {}
    for c in candidatos:
        try:
            c = Path(c).resolve()
            if c.is_file() and c.suffix.lower() in [".xlsx", ".xls"]:
                unicos[str(c)] = c
        except Exception:
            pass

    candidatos = list(unicos.values())
    if not candidatos:
        return []

    # Se houver vários arquivos do mesmo tipo/base/data, usa o mais recente.
    melhores = {}
    for caminho in candidatos:
        tipo = _tipo_arquivo_leitura(caminho)
        data_chave = _data_chave_arquivo_leitura(caminho)
        base_chave = _nome_base_por_arquivo(caminho) or str(base_nome).upper()
        chave = (base_chave, tipo, data_chave)
        atual = melhores.get(chave)
        if atual is None or caminho.stat().st_mtime > atual.stat().st_mtime:
            melhores[chave] = caminho

    # Prioriza arquivos novos de TAREFAS. Mantém Parcial apenas para aba Parcial do dia.
    return sorted(melhores.values(), key=lambda p: (p.stat().st_mtime, p.name), reverse=True)


def caminho_leitura(base_nome):
    """Compatibilidade: retorna o arquivo mais recente da base."""
    caminhos = caminhos_leitura(base_nome)
    return caminhos[0] if caminhos else None


def caminhos_leitura_disponiveis():
    """Retorna todos os arquivos por base."""
    return {
        "Americana": caminhos_leitura("Americana"),
        "Piracicaba": caminhos_leitura("Piracicaba"),
    }


MAPA_MUNICIPIOS_LEITURA = {
    "AME": "Americana",
    "COS": "Cosmópolis",
    "ELF": "Elias Fausto",
    "HOR": "Hortolândia",
    "MTM": "Monte Mor",
    "NOO": "Nova Odessa",
    "PAU": "Paulínia",
    "SBO": "Santa Bárbara do Oeste",
    "SUM": "Sumaré",
}


def _norm_col_leitura(col):
    import unicodedata
    texto = str(col).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = texto.replace(".", " ").replace("_", " ").replace("-", " ").replace("/", " ")
    return " ".join(texto.split())


def _ordenar_d(valor):
    txt = str(valor).upper().replace("D", "").strip()
    try:
        return int(txt)
    except Exception:
        return 9999


def _mes_arquivo_leitura(caminho):
    """Retorna competência MM/AAAA a partir do nome do arquivo, quando possível."""
    try:
        data_chave = _data_chave_arquivo_leitura(caminho)
        dt = pd.to_datetime(data_chave, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%m/%Y")
    except Exception:
        pass
    return ""


def _garantir_mes_operacional(df, caminho=None):
    """Garante coluna MÊS OPERACIONAL para separar viradas de mês.

    A leitura zera na virada do mês, então D0/D1/D2 só deve ser analisado
    dentro da competência correspondente à DT PREVISTA.
    """
    df = df.copy()
    if "MÊS OPERACIONAL" in df.columns and not (df["MÊS OPERACIONAL"].fillna("").astype(str).str.strip() == "").all():
        df["MÊS OPERACIONAL"] = df["MÊS OPERACIONAL"].fillna("").astype(str).str.strip()
        return df

    if "DT PREVISTA" in df.columns:
        dt = pd.to_datetime(df["DT PREVISTA"], dayfirst=True, errors="coerce")
        df["MÊS OPERACIONAL"] = dt.dt.strftime("%m/%Y").fillna("")
    else:
        df["MÊS OPERACIONAL"] = ""

    mes_arquivo = _mes_arquivo_leitura(caminho) if caminho else ""
    if mes_arquivo:
        vazio = df["MÊS OPERACIONAL"].fillna("").astype(str).str.strip() == ""
        df.loc[vazio, "MÊS OPERACIONAL"] = mes_arquivo
    return df




def _lote_maio_2026(data_valor):
    """Classificação manual de lote para Maio/2026.

    Regra definida para esta competência:
    - 04/05/2026 = Lote 1
    - cada dia útil seguinte incrementa +1
    - sábados e domingos não geram lote próprio

    Depois essa regra pode ser substituída por uma tabela/calendário automático.
    """
    dt = pd.to_datetime(data_valor, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    dt = dt.date()
    inicio = pd.Timestamp("2026-05-04").date()
    if dt < inicio or dt.month != 5 or dt.year != 2026:
        return ""
    if dt.weekday() >= 5:
        return ""
    contador = 0
    atual = inicio
    while atual <= dt:
        if atual.weekday() < 5:
            contador += 1
        atual = atual + pd.Timedelta(days=1).date() - pd.Timestamp("1970-01-01").date() if False else atual
        break
    # loop simples sem truque de pandas para manter compatível
    from datetime import timedelta as _timedelta
    contador = 0
    atual = inicio
    while atual <= dt:
        if atual.weekday() < 5:
            contador += 1
        atual += _timedelta(days=1)
    return f"Lote {contador}" if contador > 0 else ""


def _ordenar_lote(valor):
    import re
    m = re.search(r"(\d+)", str(valor or ""))
    return int(m.group(1)) if m else 9999


def _garantir_lote_operacional(df):
    """Garante coluna LOTE OPERACIONAL.

    Para Maio/2026 usa a regra manual solicitada: 04/05/2026 começa como Lote 1.
    Se a data não for de Maio/2026, deixa vazio por enquanto para não inventar lote.
    """
    df = df.copy()
    if "LOTE OPERACIONAL" in df.columns and not (df["LOTE OPERACIONAL"].fillna("").astype(str).str.strip() == "").all():
        df["LOTE OPERACIONAL"] = df["LOTE OPERACIONAL"].fillna("").astype(str).str.strip()
        return df
    if "DT PREVISTA" in df.columns:
        df["LOTE OPERACIONAL"] = pd.to_datetime(df["DT PREVISTA"], dayfirst=True, errors="coerce").apply(_lote_maio_2026)
    else:
        df["LOTE OPERACIONAL"] = ""
    return df

def _nome_base_por_arquivo(caminho):
    nome = Path(caminho).name.upper()
    if "PIRACICABA" in nome:
        return "PIRACICABA"
    if "AMERICANA" in nome:
        return "AMERICANA"
    return ""


def _normalizar_base_leitura(valor, caminho=None):
    txt = str(valor or "").strip().upper()
    if "PIRACICABA" in txt:
        return "PIRACICABA"
    if "AMERICANA" in txt:
        return "AMERICANA"
    if caminho:
        return _nome_base_por_arquivo(caminho)
    return txt


def _padronizar_colunas_leitura(df):
    """Padroniza colunas dos formatos antigos e novos da leitura.

    Importante: os Excels gerados pelo extrator podem sair com cabeçalhos abreviados
    pela tabela do Excel/print, por exemplo:
    - D OPERAÇÃO / D OPERAÇÃ / D OPERACIONA
    - FEIT, PARCIA, PENDENT, SEM TOTA
    Esta função aceita nomes completos e truncados.
    """
    mapa = {}

    for col in df.columns:
        n = _norm_col_leitura(col)

        if n == "BASE":
            mapa[col] = "BASE"

        elif n in ["TAREFA", "NUMERO TAREFA", "N TAREFA", "ID TAREFA"]:
            mapa[col] = "TAREFA"

        elif n in ["STATUS", "STATUS TAREFA"]:
            mapa[col] = "STATUS"

        elif n in ["STATUS OPERACIONAL", "SITUACAO LEITURA", "SITUACAO"]:
            mapa[col] = "STATUS OPERACIONAL"

        elif n in ["MUNICIPIO", "MUNICIPIO CODIGO", "MUN"]:
            mapa[col] = "MUNICÍPIO"

        elif n in ["MUNICIPIO NOME", "NOME MUNICIPIO", "CIDADE", "CIDADE NOME"]:
            mapa[col] = "MUNICÍPIO NOME"

        elif (
            n in ["D", "D OPER", "DIA OPERACIONAL", "D OPERACIONAL", "D OPERACIONA", "D OPERACAO"]
            or n.startswith("D OPER")
            or n.startswith("D OPERA")
        ):
            mapa[col] = "D OPERACIONAL"

        elif n in ["DT PREVISTA", "DATA PREVISTA"]:
            mapa[col] = "DT PREVISTA"

        elif n in ["DT LIMITE", "DATA LIMITE"]:
            mapa[col] = "DT LIMITE"

        elif n in ["DT PLANEJA", "DT PLANEJADA", "DATA PLANEJADA"]:
            mapa[col] = "DT PLANEJA"

        elif n in ["MES", "COMPETENCIA", "MES OPERACIONAL", "MÊS OPERACIONAL"]:
            mapa[col] = "MÊS OPERACIONAL"

        elif n in ["LOTE", "LOTE OPERACIONAL", "LOTE DA LEITURA"] or n.startswith("LOTE"):
            mapa[col] = "LOTE OPERACIONAL"

        elif n in ["AGENTE COMERCIAL", "AGENTE"]:
            mapa[col] = "AGENTE COMERCIAL"

        elif n in ["T INSTALA", "T INSTALADA", "TOTAL INSTALA", "TOTAL INSTALADA", "INSTALA", "INSTALADA"]:
            mapa[col] = "T. INSTALA"

        elif n in ["T VISITADA", "TOTAL VISITADA", "VISITADA"]:
            mapa[col] = "T. VISITADA"

        elif n in ["FALTAM", "FALTA"] or n.startswith("FALT"):
            mapa[col] = "FALTAM"

        elif n in ["DESCRICAO", "DESCRICAO TAREFA"]:
            mapa[col] = "DESCRIÇÃO"

        elif n == "TIPO":
            mapa[col] = "TIPO"

        elif (
            n in ["TOTAL TAREFAS", "TOTAL TAREFA", "TOTAL TARE", "TOTAL"]
            or n.startswith("TOTAL TARE")
        ):
            mapa[col] = "TOTAL TAREFA"

        elif n in ["FEITA", "FEIT", "FINALIZADA"] or n.startswith("FEIT"):
            mapa[col] = "FEITA"

        elif n in ["PARCIAL", "PARCIA"] or n.startswith("PARCIA"):
            mapa[col] = "PARCIAL"

        elif n in ["PENDENTE", "PENDENT"] or n.startswith("PENDENT"):
            mapa[col] = "PENDENTE"

        elif n in ["SEM TOTAL", "SEM TOTA"] or n.startswith("SEM TOTA"):
            mapa[col] = "SEM TOTAL"

    return df.rename(columns=mapa)

def _classificar_status_operacional(row):
    status = str(row.get("STATUS OPERACIONAL", "") or "").strip().upper()
    # No contrato de leitura, toda tarefa deve ter leitura prevista.
    # Se vier SEM TOTAL do arquivo antigo/grade, tratamos como PENDENTE
    # para não criar uma categoria operacional separada.
    if status == "SEM TOTAL":
        return "PENDENTE"
    if status in ["FEITA", "PARCIAL", "PENDENTE"]:
        return status

    status_tela = str(row.get("STATUS", "") or "").strip().upper()
    instala = pd.to_numeric(row.get("T. INSTALA", 0), errors="coerce")
    visitada = pd.to_numeric(row.get("T. VISITADA", 0), errors="coerce")
    instala = 0 if pd.isna(instala) else int(instala)
    visitada = 0 if pd.isna(visitada) else int(visitada)

    if instala <= 0:
        return "PENDENTE"
    if visitada >= instala or "FINAL" in status_tela:
        return "FEITA"
    if visitada > 0:
        return "PARCIAL"
    return "PENDENTE"


def _preparar_tarefas_leitura(df, caminho=None):
    """Prepara aba de tarefas real, quando existe coluna TAREFA."""
    df = _padronizar_colunas_leitura(df.copy())

    if "TAREFA" not in df.columns:
        return pd.DataFrame()

    for col in ["BASE", "MUNICÍPIO", "MUNICÍPIO NOME", "D OPERACIONAL", "STATUS", "STATUS OPERACIONAL", "AGENTE COMERCIAL", "DESCRIÇÃO", "TIPO"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    for col in ["T. INSTALA", "T. VISITADA", "FALTAM"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    for col in ["DT PREVISTA", "DT LIMITE", "DT PLANEJA"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    df = _garantir_mes_operacional(df, caminho)
    df = _garantir_lote_operacional(df)

    df["TAREFA"] = df["TAREFA"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df = df[df["TAREFA"] != ""].copy()
    if df.empty:
        return pd.DataFrame()

    if "BASE" not in df.columns or (df["BASE"].fillna("").astype(str).str.strip() == "").all():
        df["BASE"] = _nome_base_por_arquivo(caminho)
    else:
        df["BASE"] = df["BASE"].apply(lambda v: _normalizar_base_leitura(v, caminho))

    df["MUNICÍPIO"] = df["MUNICÍPIO"].str.upper().replace("", "SEM MUNICÍPIO")
    if "MUNICÍPIO NOME" not in df.columns or (df["MUNICÍPIO NOME"].fillna("").astype(str).str.strip() == "").all():
        df["MUNICÍPIO NOME"] = df["MUNICÍPIO"].map(MAPA_MUNICIPIOS_LEITURA).fillna(df["MUNICÍPIO"])
    else:
        sem_nome = df["MUNICÍPIO NOME"].fillna("").astype(str).str.strip() == ""
        df.loc[sem_nome, "MUNICÍPIO NOME"] = df.loc[sem_nome, "MUNICÍPIO"].map(MAPA_MUNICIPIOS_LEITURA).fillna(df.loc[sem_nome, "MUNICÍPIO"])

    if (df["D OPERACIONAL"].fillna("").astype(str).str.strip() == "").all():
        df["D OPERACIONAL"] = "D?"
    df["D OPERACIONAL"] = df["D OPERACIONAL"].astype(str).str.upper().str.strip()

    df["STATUS OPERACIONAL"] = df.apply(_classificar_status_operacional, axis=1)
    df["FALTAM"] = (df["T. INSTALA"] - df["T. VISITADA"]).clip(lower=0).astype(int)

    # Garante tarefa única apenas dentro do mesmo arquivo/dia.
    # A mesma tarefa pode aparecer em mais de uma DT PREVISTA e, nesse caso,
    # deve contar como ocorrência do dia/lote para bater com a Parcial do Dia.
    df["_completude"] = df.notna().sum(axis=1) + (df["T. INSTALA"] > 0).astype(int) + (df["T. VISITADA"] > 0).astype(int)
    if "DT PREVISTA" in df.columns:
        df["_DT_PREVISTA_DEDUP"] = pd.to_datetime(df["DT PREVISTA"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    else:
        df["_DT_PREVISTA_DEDUP"] = ""
    chaves_dedup_arquivo = ["BASE", "TAREFA", "_DT_PREVISTA_DEDUP"]
    df = (
        df.sort_values("_completude", ascending=False)
        .drop_duplicates(subset=chaves_dedup_arquivo, keep="first")
        .drop(columns=["_completude", "_DT_PREVISTA_DEDUP"], errors="ignore")
    )
    df["FORMATO_ORIGEM"] = "TAREFAS"
    return df.reset_index(drop=True)


def _preparar_agentes_leitura_antigo(df, caminho=None):
    """Fallback para o formato antigo: Sheet1 com AGENTE COMERCIAL/T. INSTALA/T. VISITADA.

    Esse formato não traz tarefa, município nem D. Para o painel não quebrar, criamos
    linhas sintéticas por agente e marcamos D?/SEM MUNICÍPIO. Quando o extrator novo
    gerar TAREFAS_*.xlsx, o painel passa a usar D e município reais automaticamente.
    """
    df = _padronizar_colunas_leitura(df.copy())
    obrig = {"AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA"}
    if not obrig.issubset(set(df.columns)):
        return pd.DataFrame()

    base_nome = _nome_base_por_arquivo(caminho)
    for col in ["AGENTE COMERCIAL"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
    for col in ["T. INSTALA", "T. VISITADA"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df = df[(df["AGENTE COMERCIAL"] != "") | (df["T. INSTALA"] > 0) | (df["T. VISITADA"] > 0)].copy()
    if df.empty:
        return pd.DataFrame()

    df.loc[df["AGENTE COMERCIAL"] == "", "AGENTE COMERCIAL"] = "SEM AGENTE"
    df["BASE"] = base_nome
    df["MUNICÍPIO"] = "SEM MUNICÍPIO"
    df["MUNICÍPIO NOME"] = "Sem município no arquivo"
    df["D OPERACIONAL"] = "D?"
    df["STATUS"] = ""
    df["STATUS OPERACIONAL"] = df.apply(_classificar_status_operacional, axis=1)
    df["FALTAM"] = (df["T. INSTALA"] - df["T. VISITADA"]).clip(lower=0).astype(int)
    df["TAREFA"] = [f"LEGADO-{base_nome}-{i+1:05d}" for i in range(len(df))]
    df["FORMATO_ORIGEM"] = "PARCIAL_AGENTE"
    return df.reset_index(drop=True)


def _preparar_resumo_leitura(df, caminho=None):
    df = _padronizar_colunas_leitura(df.copy())
    if "BASE" not in df.columns:
        df["BASE"] = _nome_base_por_arquivo(caminho)
    df["BASE"] = df["BASE"].apply(lambda v: _normalizar_base_leitura(v, caminho))

    if "D OPERACIONAL" not in df.columns:
        return pd.DataFrame()

    # Em alguns arquivos de resumo por município existe apenas MUNICÍPIO NOME
    # (sem o código AME/COS/etc). Nesse caso usamos o próprio nome como identificador,
    # em vez de cair em TOTAL/SEM MUNICÍPIO.
    if "MUNICÍPIO NOME" not in df.columns and "MUNICÍPIO" in df.columns:
        df["MUNICÍPIO NOME"] = df["MUNICÍPIO"]
    if "MUNICÍPIO" not in df.columns and "MUNICÍPIO NOME" in df.columns:
        df["MUNICÍPIO"] = df["MUNICÍPIO NOME"]

    for col in ["MUNICÍPIO", "MUNICÍPIO NOME"]:
        if col not in df.columns:
            df[col] = "TOTAL" if col == "MUNICÍPIO" else "TOTAL DA BASE"
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["MUNICÍPIO"] = df["MUNICÍPIO"].replace("", "TOTAL")
    df["MUNICÍPIO NOME"] = df["MUNICÍPIO NOME"].replace("", "TOTAL DA BASE")

    for col in ["TOTAL TAREFA", "FEITA", "PARCIAL", "PENDENTE", "SEM TOTAL", "FALTAM", "T. INSTALA", "T. VISITADA"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df = _garantir_mes_operacional(df, caminho)
    df = _garantir_lote_operacional(df)

    saida = df[[c for c in ["BASE", "MÊS OPERACIONAL", "LOTE OPERACIONAL", "MUNICÍPIO", "MUNICÍPIO NOME", "D OPERACIONAL", "TOTAL TAREFA", "FEITA", "PARCIAL", "PENDENTE", "SEM TOTAL", "FALTAM", "T. INSTALA", "T. VISITADA"] if c in df.columns]].copy()
    saida["FORMATO_ORIGEM"] = "RESUMO"
    return saida


def _prioridade_aba_tarefas(nome_aba):
    n = _norm_col_leitura(nome_aba)
    if "TAREFAS CONSOLIDADAS" in n:
        return 100
    if n == "TAREFAS":
        return 90
    return 10


def _prioridade_aba_resumo(nome_aba):
    n = _norm_col_leitura(nome_aba)
    if "DETALHE MUNICIPIO" in n:
        return 100
    if "RESUMO GERAL" in n:
        return 90
    if "RESUMO MUNICIPIO" in n:
        return 80
    if "RESUMO" in n:
        return 50
    return 10


def _ler_excel_leitura_robusto(caminho):
    """Lê formatos antigo e novo do extrator.

    Novo individual: RESUMO, RESUMO_MUNICIPIO, TAREFAS, abas por município.
    Novo consolidado: RESUMO_GERAL, DETALHE_MUNICIPIO, TAREFAS_CONSOLIDADAS.
    Antigo: Sheet1 com AGENTE COMERCIAL/T. INSTALA/T. VISITADA.
    """
    caminho = str(caminho)
    xls = pd.ExcelFile(caminho, engine="openpyxl")
    candidatos_tarefas = []
    candidatos_resumo = []
    candidatos_legado = []
    diagnostico = []

    for aba in xls.sheet_names:
        try:
            bruto = pd.read_excel(caminho, sheet_name=aba, engine="openpyxl")
        except Exception as e:
            diagnostico.append(f"{aba}: erro ao ler ({e})")
            continue
        if bruto.empty:
            diagnostico.append(f"{aba}: vazia")
            continue

        cols_norm = {_norm_col_leitura(c) for c in bruto.columns}
        tem_tarefa = "TAREFA" in cols_norm
        tem_resumo = ("D OPERACIONAL" in cols_norm or "D" in cols_norm) and ("TOTAL TAREFAS" in cols_norm or "TOTAL TAREFA" in cols_norm or "FEITA" in cols_norm)
        tem_legado = {"AGENTE COMERCIAL", "T INSTALA", "T VISITADA"}.issubset(cols_norm)

        if tem_tarefa:
            df_t = _preparar_tarefas_leitura(bruto, caminho)
            if not df_t.empty:
                prioridade = _prioridade_aba_tarefas(aba)
                df_t["ARQUIVO"] = Path(caminho).name
                df_t["ABA"] = aba
                candidatos_tarefas.append((prioridade, aba, df_t))
                diagnostico.append(f"{aba}: tarefas reconhecidas ({len(df_t)}) | prioridade {prioridade}")
                continue

        if tem_resumo:
            df_r = _preparar_resumo_leitura(bruto, caminho)
            if not df_r.empty:
                prioridade = _prioridade_aba_resumo(aba)
                df_r["ARQUIVO"] = Path(caminho).name
                df_r["ABA"] = aba
                candidatos_resumo.append((prioridade, aba, df_r))
                diagnostico.append(f"{aba}: resumo reconhecido ({len(df_r)}) | prioridade {prioridade}")
                continue

        if tem_legado:
            df_l = _preparar_agentes_leitura_antigo(bruto, caminho)
            if not df_l.empty:
                df_l["ARQUIVO"] = Path(caminho).name
                df_l["ABA"] = aba
                candidatos_legado.append((1, aba, df_l))
                diagnostico.append(f"{aba}: formato antigo por agente reconhecido ({len(df_l)})")
                continue

        diagnostico.append(f"{aba}: colunas não reconhecidas: {list(bruto.columns)[:10]}")

    # Evita duplicar tarefas: quando existe TAREFAS/TAREFAS_CONSOLIDADAS, ignora abas por município.
    tarefas = []
    if candidatos_tarefas:
        max_prio = max(p for p, _, _ in candidatos_tarefas)
        escolhidos = [(aba, df) for p, aba, df in candidatos_tarefas if p == max_prio]
        tarefas = [df for _, df in escolhidos]
        diagnostico.append("Usando aba(s) de tarefas: " + ", ".join(aba for aba, _ in escolhidos))
    elif candidatos_legado:
        tarefas = [df for _, _, df in candidatos_legado]
        diagnostico.append("Usando fallback legado por agente: " + ", ".join(aba for _, aba, _ in candidatos_legado))

    resumos = []
    if candidatos_resumo:
        max_prio = max(p for p, _, _ in candidatos_resumo)
        escolhidos_r = [(aba, df) for p, aba, df in candidatos_resumo if p == max_prio]
        resumos = [df for _, df in escolhidos_r]
        diagnostico.append("Usando aba(s) de resumo: " + ", ".join(aba for aba, _ in escolhidos_r))

    df_tarefas = pd.concat(tarefas, ignore_index=True) if tarefas else pd.DataFrame()
    if not df_tarefas.empty:
        df_tarefas["_ord"] = df_tarefas["D OPERACIONAL"].apply(_ordenar_d)
        # dedup final entre abas/arquivos, mantendo a primeira linha mais completa.
        df_tarefas["_completude"] = df_tarefas.notna().sum(axis=1) + (df_tarefas.get("T. INSTALA", 0) > 0).astype(int)
        if "DT PREVISTA" in df_tarefas.columns:
            df_tarefas["_DT_PREVISTA_DEDUP"] = pd.to_datetime(df_tarefas["DT PREVISTA"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        else:
            df_tarefas["_DT_PREVISTA_DEDUP"] = ""
        df_tarefas = (
            df_tarefas.sort_values(["_completude", "_ord"], ascending=[False, True])
            .drop_duplicates(subset=["BASE", "TAREFA", "_DT_PREVISTA_DEDUP"], keep="first")
            .drop(columns=["_ord", "_completude", "_DT_PREVISTA_DEDUP"], errors="ignore")
            .reset_index(drop=True)
        )

    df_resumo = pd.concat(resumos, ignore_index=True) if resumos else pd.DataFrame()
    return df_tarefas, df_resumo, diagnostico


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def carregar_leitura_completa_cache(chaves_arquivos):
    # chaves_arquivos = tupla de (base, caminho, mtime) para invalidar cache quando atualizar.
    tarefas = []
    resumos = []
    diag = []
    for base_nome, caminho_str, _mtime in chaves_arquivos:
        if not caminho_str:
            continue
        df_t, df_r, d = _ler_excel_leitura_robusto(caminho_str)
        if not df_t.empty:
            tarefas.append(df_t)
        if not df_r.empty:
            resumos.append(df_r)
        diag.extend([f"{base_nome} / {linha}" for linha in d])

    df_tarefas = pd.concat(tarefas, ignore_index=True) if tarefas else pd.DataFrame()
    df_resumo = pd.concat(resumos, ignore_index=True) if resumos else pd.DataFrame()
    return df_tarefas, df_resumo, diag


def carregar_leitura_completa():
    arquivos_por_base = {base: lista for base, lista in caminhos_leitura_disponiveis().items() if lista}
    chaves_lista = []
    for base, caminhos in arquivos_por_base.items():
        for caminho in caminhos:
            chaves_lista.append((base, str(caminho), caminho.stat().st_mtime))
    chaves = tuple(chaves_lista)
    df_tarefas, df_resumo, diag = carregar_leitura_completa_cache(chaves)
    return df_tarefas, df_resumo, arquivos_por_base, diag


def _deduplicar_tarefas_leitura(df):
    """Remove duplicidades globais de tarefa antes dos agrupamentos.

    O painel pode carregar várias execuções do mesmo dia/arquivo no GitHub/local.
    Sem esta etapa, TOTAL TAREFA fica correto por usar nunique, mas FEITA/PARCIAL/
    PENDENTE e leituras acabam somando a mesma tarefa várias vezes.
    """
    if df.empty or "TAREFA" not in df.columns:
        return df

    base = df.copy()
    base["TAREFA"] = base["TAREFA"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    base = base[base["TAREFA"] != ""].copy()

    if "DT PREVISTA" in base.columns:
        base["_DT_PREVISTA_DEDUP"] = pd.to_datetime(base["DT PREVISTA"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    else:
        base["_DT_PREVISTA_DEDUP"] = ""

    for col in ["BASE", "MÊS OPERACIONAL", "LOTE OPERACIONAL"]:
        if col not in base.columns:
            base[col] = ""
        base[col] = base[col].fillna("").astype(str).str.strip()

    # Preferimos a última versão carregada. ARQUIVO normalmente contém timestamp
    # da execução; quando não contém, a ordenação ainda é estável e segura.
    ordenacao = [c for c in ["ARQUIVO", "ABA"] if c in base.columns]
    if ordenacao:
        base = base.sort_values(ordenacao)

    # Na visão mensal/lote, usamos a MESMA unidade da Parcial do Dia:
    # tarefa por DT PREVISTA. Assim, a soma dos dias bate com a visão mensal/lote.
    # Se a mesma tarefa aparece em dias diferentes, ela conta uma vez em cada dia.
    # Se aparece várias vezes no mesmo dia por múltiplas execuções, mantém só a última.
    chaves = ["BASE", "TAREFA", "_DT_PREVISTA_DEDUP"]
    for col in chaves:
        if col not in base.columns:
            base[col] = ""
        base[col] = base[col].fillna("").astype(str).str.strip()

    base = base.drop_duplicates(subset=chaves, keep="last").drop(columns=["_DT_PREVISTA_DEDUP"], errors="ignore")
    return base.reset_index(drop=True)


def _resumo_leitura_from_tarefas(base):
    """Resumo mensal/lote consolidado por TAREFA.

    Regra operacional definida:
    - Tarefa conta 1 vez dentro do recorte mensal/lote/município.
    - A mesma tarefa pode reaparecer em outro dia porque sobrou leitura,
      mas isso não pode inflar FEITA/PARCIAL/PENDENTE.
    - O status final da tarefa é consolidado assim:
        FEITA vence PARCIAL, e PARCIAL vence PENDENTE.
    - Leituras são consolidadas por tarefa para não deixar VISITADA maior
      que INSTALA no agregado mensal/lote.
    """
    if base.empty:
        return pd.DataFrame()

    df = _garantir_mes_operacional(base.copy())
    df = _garantir_lote_operacional(df)
    df = _deduplicar_tarefas_leitura(df)

    if "TAREFA" not in df.columns:
        return pd.DataFrame()

    df["TAREFA"] = df["TAREFA"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df = df[df["TAREFA"] != ""].copy()
    if df.empty:
        return pd.DataFrame()

    df["STATUS OPERACIONAL"] = (
        df.get("STATUS OPERACIONAL", "")
        .fillna("")
        .astype(str)
        .str.upper()
        .replace({"SEM TOTAL": "PENDENTE"})
    )

    for col in ["T. INSTALA", "T. VISITADA", "FALTAM"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    agrupadores = ["BASE", "MÊS OPERACIONAL", "LOTE OPERACIONAL", "MUNICÍPIO", "MUNICÍPIO NOME"]
    for col in agrupadores:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Consolidação do status por tarefa.
    # Isso é o ponto que impede FEITA > TOTAL TAREFA.
    prioridade_status = {"PENDENTE": 0, "PARCIAL": 1, "FEITA": 2}
    status_por_prioridade = {0: "PENDENTE", 1: "PARCIAL", 2: "FEITA"}
    df["_STATUS_PRIORIDADE"] = df["STATUS OPERACIONAL"].map(prioridade_status).fillna(0).astype(int)

    tarefa_grupo = agrupadores + ["TAREFA"]
    tarefas_consolidadas = (
        df.groupby(tarefa_grupo, dropna=False)
        .agg(
            _STATUS_PRIORIDADE=("_STATUS_PRIORIDADE", "max"),
            **{
                "T. INSTALA": ("T. INSTALA", "max"),
                "T. VISITADA": ("T. VISITADA", "max"),
            },
        )
        .reset_index()
    )

    tarefas_consolidadas["STATUS OPERACIONAL"] = tarefas_consolidadas["_STATUS_PRIORIDADE"].map(status_por_prioridade).fillna("PENDENTE")

    # Se o status consolidado é FEITA, a leitura final considerada é o total previsto.
    # Caso contrário, usa o maior progresso visto, limitado ao total previsto.
    tarefas_consolidadas["T. VISITADA"] = tarefas_consolidadas[["T. VISITADA", "T. INSTALA"]].min(axis=1).astype(int)
    mask_feita = tarefas_consolidadas["STATUS OPERACIONAL"] == "FEITA"
    tarefas_consolidadas.loc[mask_feita, "T. VISITADA"] = tarefas_consolidadas.loc[mask_feita, "T. INSTALA"]
    tarefas_consolidadas["FALTAM"] = (tarefas_consolidadas["T. INSTALA"] - tarefas_consolidadas["T. VISITADA"]).clip(lower=0).astype(int)

    for col in ["FEITA", "PARCIAL", "PENDENTE"]:
        tarefas_consolidadas[col] = (tarefas_consolidadas["STATUS OPERACIONAL"] == col).astype(int)
    tarefas_consolidadas["SEM TOTAL"] = 0

    resumo = (
        tarefas_consolidadas.groupby(agrupadores, dropna=False)
        .agg(
            **{
                "TOTAL TAREFA": ("TAREFA", "count"),
                "FEITA": ("FEITA", "sum"),
                "PARCIAL": ("PARCIAL", "sum"),
                "PENDENTE": ("PENDENTE", "sum"),
                "SEM TOTAL": ("SEM TOTAL", "sum"),
                "FALTAM": ("FALTAM", "sum"),
                "T. INSTALA": ("T. INSTALA", "sum"),
                "T. VISITADA": ("T. VISITADA", "sum"),
            }
        )
        .reset_index()
    )

    resumo["ORDEM_LOTE"] = resumo["LOTE OPERACIONAL"].apply(_ordenar_lote)
    return resumo.sort_values(["BASE", "MÊS OPERACIONAL", "ORDEM_LOTE", "MUNICÍPIO NOME"]).drop(columns=["ORDEM_LOTE"])

def _formatar_df_parcial_agente(df):
    if df.empty:
        return df
    cols = [c for c in ["BASE", "AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA", "FALTAM", "STATUS OPERACIONAL"] if c in df.columns]
    out = df[cols].copy()
    if "STATUS OPERACIONAL" in out.columns:
        out = out.rename(columns={"STATUS OPERACIONAL": "STATUS"})
    return out.sort_values([c for c in ["BASE", "FALTAM", "AGENTE COMERCIAL"] if c in out.columns], ascending=[True, False, True][:len([c for c in ["BASE", "FALTAM", "AGENTE COMERCIAL"] if c in out.columns])])


def _resumo_parcial_agente(df):
    if df.empty:
        return pd.DataFrame()
    base = df.copy()
    for col in ["T. INSTALA", "T. VISITADA", "FALTAM"]:
        if col not in base.columns:
            base[col] = 0
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0).astype(int)
    if "AGENTE COMERCIAL" not in base.columns:
        base["AGENTE COMERCIAL"] = ""
    return (
        base.groupby(["BASE", "AGENTE COMERCIAL"], dropna=False)
        .agg({"T. INSTALA": "sum", "T. VISITADA": "sum", "FALTAM": "sum"})
        .reset_index()
        .sort_values(["BASE", "FALTAM", "AGENTE COMERCIAL"], ascending=[True, False, True])
    )


def mostrar_painel_leitura():
    st.subheader("📖 Contrato Leitura")
    st.caption("Acompanhamento CWSI por base, município, lote operacional e parcial do dia.")

    df_tarefas, df_resumo_arquivo, arquivos, diagnostico = carregar_leitura_completa()

    if not arquivos:
        st.warning("Nenhum arquivo de leitura encontrado. Envie os Excels para dashboard/leitura ou mantenha em C:\\Users\\user\\Desktop\\LEITURA\\saida.")
        return

    with st.expander("Arquivos carregados", expanded=False):
        for base_nome, caminhos in arquivos.items():
            st.markdown(f"**{base_nome}**")
            for caminho in caminhos:
                mtime = arquivo_mtime_datetime(caminho)
                if mtime:
                    st.caption(f"{caminho} • atualizado em {mtime.strftime('%d/%m/%Y %H:%M:%S')}")
                else:
                    st.caption(f"{caminho}")
        if diagnostico:
            st.markdown("**Diagnóstico das abas:**")
            st.code("\n".join(diagnostico[:120]))

    if df_tarefas.empty and df_resumo_arquivo.empty:
        st.error("Os arquivos foram encontrados, mas nenhuma aba reconhecida foi carregada. Abra o expander acima e confira o diagnóstico das abas.")
        return

    # Se o arquivo tem resumo por município/D, ele é a fonte principal da visão operacional.
    # O fallback legado AGENTE COMERCIAL/T. INSTALA/T. VISITADA fica somente na aba Parcial do dia.
    df_parcial = pd.DataFrame()
    df_tarefas_reais = pd.DataFrame()
    if not df_tarefas.empty:
        origem = df_tarefas.get("FORMATO_ORIGEM", pd.Series([""] * len(df_tarefas), index=df_tarefas.index)).astype(str)
        df_parcial = df_tarefas[origem == "PARCIAL_AGENTE"].copy()
        df_tarefas_reais = df_tarefas[origem != "PARCIAL_AGENTE"].copy()

    # Para respeitar a virada de mês, a visão operacional prioriza as tarefas reais,
    # pois elas trazem DT PREVISTA e permitem separar por competência.
    if not df_tarefas_reais.empty:
        df_operacional = _resumo_leitura_from_tarefas(df_tarefas_reais)
    elif not df_resumo_arquivo.empty:
        df_operacional = _garantir_mes_operacional(df_resumo_arquivo.copy())
    else:
        df_operacional = pd.DataFrame()

    if df_operacional.empty and df_parcial.empty:
        st.error("Encontrei arquivos, mas não há dados operacionais nem parcial por agente para mostrar.")
        return

    aba_operacional, aba_parcial = st.tabs(["📦 Mês / Lote", "📋 Parcial do dia"])

    with aba_operacional:
        if df_operacional.empty:
            st.warning("Não encontrei dados mensais/lotes de leitura. Verifique se o extrator subiu arquivos Tarefas_*.xlsx com DT PREVISTA.")
        else:
            st.markdown("#### Filtros")
            f0, f1, f2, f3 = st.columns([1.0, 1.1, 1.2, 1.8])
            if "MÊS OPERACIONAL" not in df_operacional.columns:
                df_operacional = _garantir_mes_operacional(df_operacional)
            if "LOTE OPERACIONAL" not in df_operacional.columns:
                df_operacional = _garantir_lote_operacional(df_operacional)

            meses_disp = df_operacional.get("MÊS OPERACIONAL", pd.Series(dtype=str)).dropna().astype(str).unique().tolist()
            meses_disp = [m for m in meses_disp if m]
            meses_disp = sorted(
                meses_disp,
                key=lambda m: pd.to_datetime("01/" + str(m), dayfirst=True, errors="coerce"),
                reverse=True,
            )
            mes_atual = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%m/%Y")
            # Para a leitura, a competência atual não deve puxar mês anterior.
            # Se ainda não houver arquivo de Maio, a visão mensal/lote fica vazia
            # em vez de mostrar Abril como se ainda fosse backlog válido.
            meses_default = [mes_atual] if mes_atual in meses_disp else []

            bases_disp = sorted(df_operacional.get("BASE", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            lotes_disp = sorted(
                [l for l in df_operacional.get("LOTE OPERACIONAL", pd.Series(dtype=str)).dropna().astype(str).unique().tolist() if l],
                key=_ordenar_lote,
            )
            municipios_disp = sorted(df_operacional.get("MUNICÍPIO NOME", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            municipios_disp = [m for m in municipios_disp if m and str(m).upper() not in ["TOTAL DA BASE", "TOTAL"]]

            meses_sel = f0.multiselect("Mês", meses_disp, default=meses_default, key="leit_mes_op")
            bases_sel = f1.multiselect("Base", bases_disp, default=bases_disp, key="leit_base_op")
            lotes_sel = f2.multiselect("Lote", lotes_disp, default=lotes_disp, key="leit_lote_op")
            municipios_sel = f3.multiselect("Município", municipios_disp, default=municipios_disp, key="leit_mun_op")

            resumo = df_operacional.copy()
            if "MÊS OPERACIONAL" in resumo.columns:
                if meses_sel:
                    resumo = resumo[resumo["MÊS OPERACIONAL"].isin(meses_sel)]
                else:
                    # Sem seleção de mês = não mostrar meses antigos por engano.
                    resumo = resumo.iloc[0:0]
            if bases_sel and "BASE" in resumo.columns:
                resumo = resumo[resumo["BASE"].isin(bases_sel)]
            if lotes_sel and "LOTE OPERACIONAL" in resumo.columns:
                resumo = resumo[resumo["LOTE OPERACIONAL"].isin(lotes_sel)]
            if municipios_sel and "MUNICÍPIO NOME" in resumo.columns:
                resumo = resumo[resumo["MUNICÍPIO NOME"].isin(municipios_sel)]

            for col in ["TOTAL TAREFA", "FEITA", "PARCIAL", "PENDENTE", "SEM TOTAL", "FALTAM", "T. INSTALA", "T. VISITADA"]:
                if col not in resumo.columns:
                    resumo[col] = 0
                resumo[col] = pd.to_numeric(resumo[col], errors="coerce").fillna(0).astype(int)

            total_tarefas = int(resumo["TOTAL TAREFA"].sum())
            feitas = int(resumo["FEITA"].sum())
            parciais = int(resumo["PARCIAL"].sum())
            pendentes = int(resumo["PENDENTE"].sum()) + int(resumo["SEM TOTAL"].sum())
            leituras_total = int(resumo["T. INSTALA"].sum())
            leituras_feitas = int(resumo["T. VISITADA"].sum())
            faltam = max(0, leituras_total - leituras_feitas)
            perc_leitura = (leituras_feitas / leituras_total * 100) if leituras_total else 0
            perc_faltante = (faltam / leituras_total * 100) if leituras_total else 0

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Tarefas", numero(total_tarefas))
            c2.metric("Feitas", numero(feitas))
            c3.metric("Parciais", numero(parciais))
            c4.metric("Pendentes", numero(pendentes))
            c5.metric("Leituras faltantes", numero(faltam), f"{perc_faltante:.1f}%".replace(".", ","))
            c6.metric("Execução leitura", f"{perc_leitura:.1f}%".replace(".", ","))

            if faltam:
                st.markdown(f"<div class='zero-card'><b>Leituras faltantes no mês/lote:</b> {numero(faltam)} ({perc_faltante:.1f}%)</div>".replace(".", ","), unsafe_allow_html=True)

            st.markdown("#### Resumo por mês, base e lote")
            resumo_base_lote = (
                resumo.groupby(["MÊS OPERACIONAL", "BASE", "LOTE OPERACIONAL"], dropna=False)
                .agg({"TOTAL TAREFA": "sum", "FEITA": "sum", "PARCIAL": "sum", "PENDENTE": "sum", "SEM TOTAL": "sum", "FALTAM": "sum", "T. INSTALA": "sum", "T. VISITADA": "sum"})
                .reset_index()
            )
            if not resumo_base_lote.empty:
                resumo_base_lote["ORDEM_LOTE"] = resumo_base_lote["LOTE OPERACIONAL"].apply(_ordenar_lote)
                resumo_base_lote = resumo_base_lote.sort_values(["MÊS OPERACIONAL", "BASE", "ORDEM_LOTE"]).drop(columns=["ORDEM_LOTE"])
            tabela_base_lote = resumo_base_lote.rename(columns={"T. INSTALA": "LEITURAS TOTAL", "T. VISITADA": "LEITURAS FEITAS", "FALTAM": "LEITURAS FALTANTES"})
            st.dataframe(tabela_base_lote, use_container_width=True, hide_index=True)

            st.markdown("#### Detalhe por município")
            detalhe = resumo.copy()
            if "ORDEM_LOTE" not in detalhe.columns:
                detalhe["ORDEM_LOTE"] = detalhe["LOTE OPERACIONAL"].apply(_ordenar_lote)
            detalhe = detalhe.sort_values([c for c in ["BASE", "MUNICÍPIO NOME", "ORDEM_LOTE"] if c in detalhe.columns]).drop(columns=["ORDEM_LOTE"], errors="ignore")
            colunas = [c for c in ["MÊS OPERACIONAL", "BASE", "LOTE OPERACIONAL", "MUNICÍPIO", "MUNICÍPIO NOME", "TOTAL TAREFA", "FEITA", "PARCIAL", "PENDENTE", "FALTAM", "T. INSTALA", "T. VISITADA"] if c in detalhe.columns]
            tabela_detalhe = detalhe[colunas].rename(columns={"T. INSTALA": "LEITURAS TOTAL", "T. VISITADA": "LEITURAS FEITAS", "FALTAM": "LEITURAS FALTANTES"})
            st.dataframe(tabela_detalhe, use_container_width=True, hide_index=True)

            if not resumo_base_lote.empty:
                chart_df = resumo_base_lote.melt(
                    id_vars=["BASE", "LOTE OPERACIONAL"],
                    value_vars=["FEITA", "PARCIAL", "PENDENTE"],
                    var_name="STATUS",
                    value_name="QTD",
                )
                grafico = (
                    alt.Chart(chart_df)
                    .mark_bar()
                    .encode(
                        x=alt.X("LOTE OPERACIONAL:N", sort=sorted(chart_df["LOTE OPERACIONAL"].unique().tolist(), key=_ordenar_lote)),
                        y="QTD:Q",
                        color="STATUS:N",
                        column="BASE:N",
                        tooltip=["BASE", "LOTE OPERACIONAL", "STATUS", "QTD"],
                    )
                    .properties(height=260)
                )
                st.altair_chart(grafico, use_container_width=True)

    with aba_parcial:
        st.caption("Parcial do dia da leitura: escolha a data e veja produção por agente e serviços que ainda não estão prontos.")

        # Preferimos o formato novo do extrator, porque ele traz DT PREVISTA, tarefa,
        # município, D operacional e status real da tarefa. O formato antigo por agente
        # fica apenas como fallback quando ainda não houver arquivos Tarefas_*.xlsx.
        if not df_tarefas_reais.empty:
            parcial_dia = df_tarefas_reais.copy()

            # Normaliza datas para permitir seleção igual ao painel de corte.
            if "DT PREVISTA" in parcial_dia.columns:
                parcial_dia["DT_PREVISTA_DT"] = pd.to_datetime(parcial_dia["DT PREVISTA"], dayfirst=True, errors="coerce")
            else:
                parcial_dia["DT_PREVISTA_DT"] = pd.NaT
            parcial_dia = _garantir_lote_operacional(parcial_dia)

            datas_disponiveis_df = (
                parcial_dia[["DT_PREVISTA_DT"]]
                .dropna()
                .drop_duplicates()
                .sort_values("DT_PREVISTA_DT", ascending=False)
            )

            if datas_disponiveis_df.empty:
                st.warning("Os arquivos novos foram encontrados, mas não trouxeram DT PREVISTA. Exibindo sem filtro de data.")
                data_escolhida = None
                parcial_filtrada = parcial_dia.copy()
            else:
                datas_opcoes = datas_disponiveis_df["DT_PREVISTA_DT"].dt.strftime("%d/%m/%Y").tolist()
                data_escolhida = st.selectbox("Escolha o dia", datas_opcoes, index=0, key="leit_parcial_dia_select")
                data_escolhida_dt = pd.to_datetime(data_escolhida, dayfirst=True, errors="coerce")
                parcial_filtrada = parcial_dia[parcial_dia["DT_PREVISTA_DT"] == data_escolhida_dt].copy()

            if "STATUS OPERACIONAL" in parcial_filtrada.columns:
                parcial_filtrada["STATUS OPERACIONAL"] = parcial_filtrada["STATUS OPERACIONAL"].fillna("").astype(str).str.upper().replace({"SEM TOTAL": "PENDENTE"})

            fbase, fmun, fstatus = st.columns([1.1, 1.7, 1.3])
            bases_disp_p = sorted(parcial_filtrada.get("BASE", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            municipios_disp_p = sorted(parcial_filtrada.get("MUNICÍPIO NOME", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            status_disp_p = sorted(parcial_filtrada.get("STATUS OPERACIONAL", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

            base_sel_p = fbase.multiselect("Base", bases_disp_p, default=bases_disp_p, key="leit_parcial_base_novo")
            municipio_sel_p = fmun.multiselect("Município", municipios_disp_p, default=municipios_disp_p, key="leit_parcial_mun_novo")
            status_sel_p = fstatus.multiselect("Status", status_disp_p, default=status_disp_p, key="leit_parcial_status_novo")

            if base_sel_p:
                parcial_filtrada = parcial_filtrada[parcial_filtrada["BASE"].isin(base_sel_p)]
            if municipio_sel_p and "MUNICÍPIO NOME" in parcial_filtrada.columns:
                parcial_filtrada = parcial_filtrada[parcial_filtrada["MUNICÍPIO NOME"].isin(municipio_sel_p)]
            if status_sel_p and "STATUS OPERACIONAL" in parcial_filtrada.columns:
                parcial_filtrada = parcial_filtrada[parcial_filtrada["STATUS OPERACIONAL"].isin(status_sel_p)]

            for col in ["T. INSTALA", "T. VISITADA", "FALTAM"]:
                if col not in parcial_filtrada.columns:
                    parcial_filtrada[col] = 0
                parcial_filtrada[col] = pd.to_numeric(parcial_filtrada[col], errors="coerce").fillna(0).astype(int)

            # Pode haver vários arquivos do mesmo dia no GitHub/local, porque o extrator
            # salva com timestamp a cada execução. Antes o painel somava as linhas
            # duplicadas para FEITA/PARCIAL/PENDENTE, mas contava TAREFA como única.
            # Isso gerava situações impossíveis, como 212 tarefas e 632 feitas.
            # Aqui consolidamos novamente por tarefa única no recorte escolhido.
            if "TAREFA" in parcial_filtrada.columns and not parcial_filtrada.empty:
                chaves_dedup = [c for c in ["BASE", "TAREFA", "DT_PREVISTA_DT"] if c in parcial_filtrada.columns]
                if not chaves_dedup:
                    chaves_dedup = ["TAREFA"]
                parcial_filtrada = (
                    parcial_filtrada
                    .sort_values([c for c in ["ARQUIVO", "ABA"] if c in parcial_filtrada.columns])
                    .drop_duplicates(subset=chaves_dedup, keep="last")
                    .reset_index(drop=True)
                )

            total_tarefas = int(parcial_filtrada["TAREFA"].nunique()) if "TAREFA" in parcial_filtrada.columns else len(parcial_filtrada)
            status_series = parcial_filtrada.get("STATUS OPERACIONAL", pd.Series(dtype=str)).fillna("").astype(str).str.upper() if not parcial_filtrada.empty else pd.Series(dtype=str)
            status_series = status_series.replace({"SEM TOTAL": "PENDENTE"})
            feitas = int((status_series == "FEITA").sum())
            parciais = int((status_series == "PARCIAL").sum())
            pendentes = int((status_series == "PENDENTE").sum())

            leituras_total = int(parcial_filtrada["T. INSTALA"].sum()) if not parcial_filtrada.empty else 0
            leituras_feitas = int(parcial_filtrada["T. VISITADA"].sum()) if not parcial_filtrada.empty else 0
            leituras_faltantes = max(0, leituras_total - leituras_feitas)
            perc_leituras = (leituras_feitas / leituras_total * 100) if leituras_total else 0
            perc_faltantes = (leituras_faltantes / leituras_total * 100) if leituras_total else 0

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Tarefas", numero(total_tarefas))
            m2.metric("Feitas", numero(feitas))
            m3.metric("Parciais", numero(parciais))
            m4.metric("Pendentes", numero(pendentes))
            m5.metric("Leituras faltantes", numero(leituras_faltantes), f"{perc_faltantes:.1f}%".replace(".", ","))
            m6.metric("Execução leitura", f"{perc_leituras:.1f}%".replace(".", ","))

            l1, l2, l3 = st.columns(3)
            l1.metric("Leituras totais", numero(leituras_total))
            l2.metric("Leituras feitas", numero(leituras_feitas))
            l3.metric("Leituras faltantes", numero(leituras_faltantes), f"{perc_faltantes:.1f}%".replace(".", ","))

            if leituras_faltantes:
                texto_faltantes = f"{perc_faltantes:.1f}%".replace(".", ",")
                st.markdown(f"<div class='zero-card'><b>Leituras faltantes no dia:</b> {numero(leituras_faltantes)} ({texto_faltantes})</div>", unsafe_allow_html=True)

            # Alertas operacionais do dia
            alerta_base = parcial_filtrada.copy()
            if "AGENTE COMERCIAL" not in alerta_base.columns:
                alerta_base["AGENTE COMERCIAL"] = ""
            alerta_base["AGENTE COMERCIAL"] = alerta_base["AGENTE COMERCIAL"].fillna("").astype(str).str.strip()

            status_alerta = alerta_base.get("STATUS OPERACIONAL", pd.Series([""] * len(alerta_base), index=alerta_base.index)).fillna("").astype(str).str.upper()
            faltam_alerta = pd.to_numeric(alerta_base.get("FALTAM", 0), errors="coerce").fillna(0)
            sem_agente_df = alerta_base[(alerta_base["AGENTE COMERCIAL"] == "") & (status_alerta != "FEITA") & (faltam_alerta > 0)].copy()
            if not sem_agente_df.empty:
                qtd_sem_agente = int(sem_agente_df["TAREFA"].nunique()) if "TAREFA" in sem_agente_df.columns else len(sem_agente_df)
                exemplos_sem_agente = sem_agente_df.get("TAREFA", pd.Series(dtype=str)).astype(str).drop_duplicates().head(12).tolist()
                texto_exemplos = ", ".join(exemplos_sem_agente) if exemplos_sem_agente else "ver tabela de tarefas"
                st.markdown(
                    f"<div class='zero-card'><b>⚠️ Tarefas sem agente atribuído:</b> {numero(qtd_sem_agente)}<br>{texto_exemplos}</div>",
                    unsafe_allow_html=True,
                )

            agentes_inicio = alerta_base[alerta_base["AGENTE COMERCIAL"] != ""].copy()
            if not agentes_inicio.empty:
                resumo_inicio = (
                    agentes_inicio.groupby(["BASE", "AGENTE COMERCIAL"], dropna=False)
                    .agg({"T. INSTALA": "sum", "T. VISITADA": "sum"})
                    .reset_index()
                )
                sem_inicio = resumo_inicio[(resumo_inicio["T. INSTALA"] > 0) & (resumo_inicio["T. VISITADA"] <= 0)].copy()
                if not sem_inicio.empty:
                    nomes_sem_inicio = sem_inicio["AGENTE COMERCIAL"].astype(str).tolist()
                    st.markdown(
                        f"<div class='zero-card'><b>⚠️ Agentes sem iniciar leitura no dia:</b> {numero(len(nomes_sem_inicio))}<br>{', '.join(nomes_sem_inicio[:20])}</div>",
                        unsafe_allow_html=True,
                    )

            st.markdown("#### Produção por agente comercial")
            agentes = parcial_filtrada.copy()
            if "AGENTE COMERCIAL" not in agentes.columns:
                agentes["AGENTE COMERCIAL"] = ""
            agentes["AGENTE COMERCIAL"] = agentes["AGENTE COMERCIAL"].fillna("").astype(str).str.strip()
            agentes.loc[agentes["AGENTE COMERCIAL"] == "", "AGENTE COMERCIAL"] = "SEM AGENTE"

            agentes["STATUS OPERACIONAL"] = agentes.get("STATUS OPERACIONAL", "").replace({"SEM TOTAL": "PENDENTE"})
            for status_nome in ["FEITA", "PARCIAL", "PENDENTE"]:
                agentes[status_nome] = (agentes.get("STATUS OPERACIONAL", "") == status_nome).astype(int)

            if agentes.empty:
                st.info("Nenhuma tarefa encontrada para os filtros selecionados.")
            else:
                resumo_agente = (
                    agentes.groupby(["BASE", "AGENTE COMERCIAL"], dropna=False)
                    .agg(
                        TAREFAS=("TAREFA", "nunique") if "TAREFA" in agentes.columns else ("AGENTE COMERCIAL", "size"),
                        FEITAS=("FEITA", "sum"),
                        PARCIAIS=("PARCIAL", "sum"),
                        PENDENTES=("PENDENTE", "sum"),
                        **{
                            "LEITURAS TOTAL": ("T. INSTALA", "sum"),
                            "LEITURAS FEITAS": ("T. VISITADA", "sum"),
                            "LEITURAS FALTANTES": ("FALTAM", "sum"),
                        },
                    )
                    .reset_index()
                )
                resumo_agente["% LEITURA EXECUTADA"] = 0.0
                mask = resumo_agente["LEITURAS TOTAL"] > 0
                resumo_agente.loc[mask, "% LEITURA EXECUTADA"] = ((resumo_agente.loc[mask, "LEITURAS FEITAS"] / resumo_agente.loc[mask, "LEITURAS TOTAL"]) * 100).round(1)
                resumo_agente["% LEITURA FALTANTE"] = 0.0
                resumo_agente.loc[mask, "% LEITURA FALTANTE"] = ((resumo_agente.loc[mask, "LEITURAS FALTANTES"] / resumo_agente.loc[mask, "LEITURAS TOTAL"]) * 100).round(1)
                resumo_agente = resumo_agente.sort_values(["LEITURAS FALTANTES", "PENDENTES", "PARCIAIS", "TAREFAS"], ascending=[False, False, False, False])
                tabela_agente = resumo_agente.copy()
                tabela_agente["% LEITURA EXECUTADA"] = tabela_agente["% LEITURA EXECUTADA"].apply(lambda v: f"{float(v):.1f}%".replace(".", ","))
                tabela_agente["% LEITURA FALTANTE"] = tabela_agente["% LEITURA FALTANTE"].apply(lambda v: f"{float(v):.1f}%".replace(".", ","))
                st.dataframe(tabela_agente, use_container_width=True, hide_index=True)

            st.markdown("#### Serviços que não estão prontos")
            nao_prontos = parcial_filtrada[
                (parcial_filtrada.get("STATUS OPERACIONAL", "") != "FEITA") |
                (pd.to_numeric(parcial_filtrada.get("FALTAM", 0), errors="coerce").fillna(0) > 0)
            ].copy()

            if nao_prontos.empty:
                st.success("Nenhum serviço pendente/parcial para os filtros selecionados.")
            else:
                for col in ["DT PREVISTA", "DT LIMITE", "DT PLANEJA"]:
                    if col in nao_prontos.columns:
                        nao_prontos[col] = pd.to_datetime(nao_prontos[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                cols_pend = [c for c in ["BASE", "MUNICÍPIO", "MUNICÍPIO NOME", "LOTE OPERACIONAL", "TAREFA", "STATUS OPERACIONAL", "STATUS", "DT PREVISTA", "AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA", "FALTAM", "DESCRIÇÃO", "TIPO"] if c in nao_prontos.columns]
                nao_prontos = nao_prontos.sort_values([c for c in ["BASE", "MUNICÍPIO NOME", "FALTAM", "TAREFA"] if c in nao_prontos.columns], ascending=[True, True, False, True][:len([c for c in ["BASE", "MUNICÍPIO NOME", "FALTAM", "TAREFA"] if c in nao_prontos.columns])])
                tabela_nao_prontos = nao_prontos[cols_pend].rename(columns={"T. INSTALA": "LEITURAS TOTAL", "T. VISITADA": "LEITURAS FEITAS", "FALTAM": "LEITURAS FALTANTES"})
                st.dataframe(tabela_nao_prontos, use_container_width=True, hide_index=True)

            with st.expander("Ver todas as tarefas do dia", expanded=False):
                todas = parcial_filtrada.copy()
                for col in ["DT PREVISTA", "DT LIMITE", "DT PLANEJA"]:
                    if col in todas.columns:
                        todas[col] = pd.to_datetime(todas[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                cols_todas = [c for c in ["BASE", "MUNICÍPIO", "MUNICÍPIO NOME", "LOTE OPERACIONAL", "TAREFA", "STATUS OPERACIONAL", "STATUS", "DT PREVISTA", "AGENTE COMERCIAL", "T. INSTALA", "T. VISITADA", "FALTAM", "DESCRIÇÃO", "TIPO"] if c in todas.columns]
                tabela_todas = todas[cols_todas].rename(columns={"T. INSTALA": "LEITURAS TOTAL", "T. VISITADA": "LEITURAS FEITAS", "FALTAM": "LEITURAS FALTANTES"})
                st.dataframe(tabela_todas, use_container_width=True, hide_index=True)

        elif not df_parcial.empty:
            st.warning("Só encontrei o formato antigo por agente, sem data da tarefa. Para escolher o dia, rode o extrator novo Tarefas_*.xlsx.")
            bases_disp_p = sorted(df_parcial.get("BASE", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            base_sel_p = st.multiselect("Base da parcial", bases_disp_p, default=bases_disp_p, key="leit_base_parcial_legado")
            parcial_filtrada = df_parcial.copy()
            if base_sel_p:
                parcial_filtrada = parcial_filtrada[parcial_filtrada["BASE"].isin(base_sel_p)]

            total_instala = int(pd.to_numeric(parcial_filtrada.get("T. INSTALA", 0), errors="coerce").fillna(0).sum())
            total_visitada = int(pd.to_numeric(parcial_filtrada.get("T. VISITADA", 0), errors="coerce").fillna(0).sum())
            total_faltam = int(pd.to_numeric(parcial_filtrada.get("FALTAM", 0), errors="coerce").fillna(0).sum())
            percentual = (total_visitada / total_instala * 100) if total_instala else 0

            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Leituras totais", numero(total_instala))
            p2.metric("Leituras feitas", numero(total_visitada))
            p3.metric("Leituras faltantes", numero(total_faltam))
            p4.metric("% Leitura executada", f"{percentual:.1f}%".replace(".", ","))

            resumo_agente = _resumo_parcial_agente(parcial_filtrada)
            tabela = resumo_agente.copy()
            if not tabela.empty:
                tabela["% EXECUTADO"] = 0.0
                mask = tabela["T. INSTALA"] > 0
                tabela.loc[mask, "% EXECUTADO"] = ((tabela.loc[mask, "T. VISITADA"] / tabela.loc[mask, "T. INSTALA"]) * 100).round(1)
                tabela["% EXECUTADO"] = tabela["% EXECUTADO"].apply(lambda v: f"{float(v):.1f}%".replace(".", ","))
            st.dataframe(tabela, use_container_width=True, hide_index=True)
        else:
            st.warning("Não encontrei dados para a parcial do dia.")


def mostrar_base_leitura(base_nome):
    """Mantido por compatibilidade; agora mostra o painel completo integrado."""
    mostrar_painel_leitura()




def sqlite_ativado():
    """Permite desligar o SQLite por Secret caso precise voltar 100% para CSV."""
    valor = str(secret_value("USAR_SQLITE_GZUS", "true") or "true").strip().lower()
    return valor not in ["0", "false", "nao", "não", "no", "off"]


def caminho_banco_gzus():
    for caminho in BANCO_GZUS_CANDIDATOS:
        try:
            if caminho.exists() and caminho.is_file():
                return caminho
        except Exception:
            pass
    return None


def _sqlite_tabela_existe(caminho_banco, tabela):
    try:
        with sqlite3.connect(str(caminho_banco)) as conn:
            cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tabela,))
            return cur.fetchone() is not None
    except Exception:
        return False


def _sqlite_tabela_parece_valida(df):
    """Evita usar tabela importada com separador errado ou vazia."""
    if df is None or df.empty:
        return False
    if len(df.columns) <= 1:
        return False
    # Se o CSV foi importado com separador errado, normalmente vira uma coluna gigante com ';' no nome.
    if any(";" in str(c) for c in df.columns):
        return False
    return True


def _normalizar_df_dashboard(df):
    """Aplica conversões numéricas iguais às usadas nos CSVs."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Metadados criados pelo banco_gzus.py não atrapalham, mas também não são usados no painel.
    # Mantemos por segurança; se quiser ocultar depois, basta filtrar aqui.
    for col in df.columns:
        if "FATURAMENTO" in str(col):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        if col in ["QTD_NOTAS", "QTD_EXECUTORES", "DIA_SEMANA_NUM"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def ler_sqlite_dashboard_cache(caminho_banco_str, tabela, mtime_banco):
    # mtime_banco entra só para invalidar cache quando gzus.db for atualizado.
    del mtime_banco
    with sqlite3.connect(caminho_banco_str) as conn:
        return pd.read_sql_query(f'SELECT * FROM "{tabela}"', conn)


def ler_sqlite_dashboard(chave):
    """Tenta carregar uma base do dashboard pelo SQLite. Retorna DataFrame vazio se não der."""
    if not sqlite_ativado():
        return pd.DataFrame()

    caminho_banco = caminho_banco_gzus()
    if not caminho_banco:
        return pd.DataFrame()

    tabela = TABELAS_SQLITE_DASHBOARD.get(chave)
    if not tabela or not _sqlite_tabela_existe(caminho_banco, tabela):
        return pd.DataFrame()

    try:
        mtime = caminho_banco.stat().st_mtime
        df = ler_sqlite_dashboard_cache(str(caminho_banco), tabela, mtime)
        if not _sqlite_tabela_parece_valida(df):
            return pd.DataFrame()
        return _normalizar_df_dashboard(df)
    except Exception:
        return pd.DataFrame()


def ler_base_dashboard(chave, nome_arquivo):
    """Fonte única de leitura.

    Regra importante:
    - Para tabelas pequenas/resumidas, usa SQLite primeiro e CSV como plano B.
    - Para NOTAS, força CSV primeiro.

    Motivo: o gzus_dashboard.db é um banco leve do painel e pode não conter o
    histórico completo de notas. Ranking, parcial do dia e meses anteriores
    dependem do dashboard/notas_dashboard.csv acumulado.
    """
    if chave == "notas":
        caminho = caminho_arquivo(nome_arquivo)
        if caminho:
            return ler_csv(str(caminho)), "csv_notas_historico"

        # Fallback local: só usa SQLite se o CSV realmente não existir.
        df_sqlite = ler_sqlite_dashboard(chave)
        if not df_sqlite.empty:
            return df_sqlite, "sqlite_fallback_notas"

        return pd.DataFrame(), "faltando"

    df_sqlite = ler_sqlite_dashboard(chave)
    if not df_sqlite.empty:
        return df_sqlite, "sqlite"

    caminho = caminho_arquivo(nome_arquivo)
    if caminho:
        return ler_csv(str(caminho)), "csv"

    return pd.DataFrame(), "faltando"


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
        "CORTE", "RELIGUE", "VERIFICACAO", "VERIFICAÇÃO", "TOTAL", "MÍNIMO", "MÁXIMO", "VALOR",
        "Total semana", "FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
    ]

    for col in df2.columns:
        if "FATURAMENTO" in col or col in colunas_moeda:
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["QTD_NOTAS", "NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS", "EXPRESS", "TOTAL_NOTAS"]:
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
        elif col in ["POSIÇÃO", "NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS", "EXPRESS", "DIAS_ATIVOS", "QTD_EQUIPES", "QTD_RECURSOS"]:
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
    """Monta base de ranking por RECURSO/equipe, incluindo recusas para auditoria."""
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)

    if parcial.empty:
        return pd.DataFrame()

    base = parcial.copy()
    if "RECURSO" not in base.columns:
        base["RECURSO"] = ""

    base["RECURSO"] = base["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    base = base[base["RECURSO"] != ""].copy()

    if base.empty:
        return pd.DataFrame()

    if "EH_RECUSA" not in base.columns:
        base["EH_RECUSA"] = 0
    base["EH_RECUSA"] = pd.to_numeric(base["EH_RECUSA"], errors="coerce").fillna(0).astype(int)

    base["ORDEM_SERVICO_PAGAVEL"] = base["ORDEM_DE_SERVICO"].where(base["EH_RECUSA"] == 0, pd.NA)
    base["ORDEM_SERVICO_RECUSA"] = base["ORDEM_DE_SERVICO"].where(base["EH_RECUSA"] == 1, pd.NA)
    base["DATA_PAGAVEL"] = base["DATA"].where(base["EH_RECUSA"] == 0, pd.NA)

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

    base_calc = base_filtrada.copy()
    for col in ["ORDEM_SERVICO_PAGAVEL", "ORDEM_SERVICO_RECUSA", "DATA_PAGAVEL"]:
        if col not in base_calc.columns:
            base_calc[col] = pd.NA

    ranking = (
        base_calc.groupby("RECURSO", dropna=False)
        .agg(
            NOTAS=("ORDEM_SERVICO_PAGAVEL", "nunique"),
            RECUSAS=("ORDEM_SERVICO_RECUSA", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            DIAS_ATIVOS=("DATA_PAGAVEL", "nunique"),
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
    ranking = ranking.sort_values([coluna_ordem, "NOTAS", "RECUSAS"], ascending=[False, False, False]).reset_index(drop=True)
    ranking.insert(0, "POSIÇÃO", range(1, len(ranking) + 1))

    return ranking


def calcular_recusas_por_tipo(base_filtrada):
    """Resume as recusas por equipe e por motivo no período filtrado."""
    if base_filtrada.empty or "EH_RECUSA" not in base_filtrada.columns:
        return pd.DataFrame(columns=["RECURSO", "CONTRATO", "RECUSA", "QTD_RECUSAS"])

    recusas = base_filtrada.copy()
    recusas["EH_RECUSA"] = pd.to_numeric(recusas.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
    recusas = recusas[recusas["EH_RECUSA"] == 1].copy()

    if recusas.empty:
        return pd.DataFrame(columns=["RECURSO", "CONTRATO", "RECUSA", "QTD_RECUSAS"])

    recusas["RECUSA"] = recusas.get("RECUSA", "").fillna("").astype(str).str.strip()
    recusas.loc[recusas["RECUSA"] == "", "RECUSA"] = "Não informado"

    resumo = (
        recusas.groupby(["RECURSO", "CONTRATO", "RECUSA"], dropna=False)
        .agg(QTD_RECUSAS=("ORDEM_DE_SERVICO", "nunique"))
        .reset_index()
        .sort_values(["RECURSO", "QTD_RECUSAS", "RECUSA"], ascending=[True, False, True])
    )
    return resumo


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
    fontes = {}

    for chave, nome in ARQUIVOS.items():
        df, fonte = ler_base_dashboard(chave, nome)
        if not df.empty:
            bases[chave] = df
            fontes[chave] = fonte
        else:
            faltando.append(nome)
            fontes[chave] = fonte

    st.session_state["fontes_dados_dashboard"] = fontes
    return bases, faltando


# ==============================
# CARREGAMENTO RÁPIDO POR SQL
# ==============================
# Esta parte evita carregar a tabela grande de notas no pós-login.
# O app carrega primeiro só as tabelas pequenas e, depois, busca notas do período escolhido.

COLUNAS_NOTAS_MINIMAS = [
    "ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "RECUSA",
    "ELETRICISTA1", "ELETRICISTA2", "DATA", "DATA_ENCERRAMENTO",
    "QTD_EXECUTORES", "CONTRATO", "MUNICIPIO", "BAIRRO", "STATUS", "EH_VERIFICACAO",
]


def carregar_bases_leves():
    bases = {}
    faltando = []
    fontes = {}
    for chave, nome in ARQUIVOS.items():
        if chave == "notas":
            continue
        df, fonte = ler_base_dashboard(chave, nome)
        if not df.empty:
            bases[chave] = df
            fontes[chave] = fonte
        else:
            faltando.append(nome)
            fontes[chave] = fonte
    st.session_state["fontes_dados_dashboard"] = fontes
    return bases, faltando


def _sqlite_colunas(caminho_banco_str, tabela):
    try:
        with sqlite3.connect(caminho_banco_str) as conn:
            return [r[1] for r in conn.execute(f'PRAGMA table_info("{tabela}")').fetchall()]
    except Exception:
        return []


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def meses_disponiveis_sql_cache(caminho_banco_str, tabela, mtime_banco):
    del mtime_banco
    cols = _sqlite_colunas(caminho_banco_str, tabela)
    col_data = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in cols else ("DATA" if "DATA" in cols else "")
    if not col_data:
        return pd.DataFrame(columns=["MES", "PERIODO"])
    with sqlite3.connect(caminho_banco_str) as conn:
        df = pd.read_sql_query(f'SELECT DISTINCT "{col_data}" AS DATA_REF FROM "{tabela}" WHERE "{col_data}" IS NOT NULL', conn)
    if df.empty:
        return pd.DataFrame(columns=["MES", "PERIODO"])
    datas = pd.to_datetime(df["DATA_REF"], dayfirst=True, errors="coerce")
    meses = pd.DataFrame({"PERIODO": datas.dt.to_period("M")}).dropna().drop_duplicates().sort_values("PERIODO", ascending=False)
    if meses.empty:
        return pd.DataFrame(columns=["MES", "PERIODO"])
    meses["MES"] = meses["PERIODO"].dt.strftime("%m/%Y")
    return meses[["MES", "PERIODO"]].reset_index(drop=True)


def meses_disponiveis_rapido():
    """Lista meses disponíveis usando o CSV histórico de notas.

    Não usa o SQLite leve para notas, porque esse banco pode não carregar o
    histórico completo necessário para ranking/parciais antigas.
    """
    df, _ = ler_base_dashboard("notas", ARQUIVOS["notas"])
    return meses_disponiveis_da_base(df)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def ler_notas_sql_periodo_cache(caminho_banco_str, tabela, meses, mtime_banco):
    del mtime_banco
    cols = _sqlite_colunas(caminho_banco_str, tabela)
    if not cols:
        return pd.DataFrame()
    colunas = [c for c in COLUNAS_NOTAS_MINIMAS if c in cols]
    if not colunas:
        colunas = cols
    select_cols = ", ".join([f'"{c}"' for c in colunas])
    where = []
    params = []
    meses = list(meses or [])
    col_data = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in cols else ("DATA" if "DATA" in cols else "")
    if meses and col_data:
        partes = []
        for mes in meses:
            try:
                mm, aa = str(mes).split("/")
                partes.append(f'"{col_data}" LIKE ?')
                params.append(f"%/{mm}/{aa}")
                partes.append(f'"{col_data}" LIKE ?')
                params.append(f"{aa}-{mm}%")
            except Exception:
                pass
        if partes:
            where.append("(" + " OR ".join(partes) + ")")
    sql = f'SELECT {select_cols} FROM "{tabela}"'
    if where:
        sql += " WHERE " + " AND ".join(where)
    with sqlite3.connect(caminho_banco_str) as conn:
        df = pd.read_sql_query(sql, conn, params=params)
    return _normalizar_df_dashboard(df)


def carregar_notas_rapido(meses=None):
    """Carrega notas para ranking/parciais sempre a partir do CSV histórico.

    O SQLite continua sendo usado para bases pequenas de faturamento, mas as
    notas completas ficam no dashboard/notas_dashboard.csv. Isso preserva os
    meses anteriores no ranking e na parcial do dia.
    """
    df, fonte = ler_base_dashboard("notas", ARQUIVOS["notas"])
    if meses and not df.empty:
        col_data = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in df.columns else "DATA"
        if col_data in df.columns:
            datas = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
            df = df[datas.dt.strftime("%m/%Y").isin(list(meses))].copy()
    fontes = st.session_state.get("fontes_dados_dashboard", {})
    fontes["notas"] = fonte
    st.session_state["fontes_dados_dashboard"] = fontes
    return df


# ==============================
# HOME ULTRARRÁPIDA SEM NOTAS BRUTAS
# ==============================
# A tela inicial não deve depender da tabela grande de notas. Ela usa as tabelas
# já pré-processadas pelo extrator: faturamento_dias e faturamento_carro_dias.
# Isso faz o primeiro painel aparecer rápido e deixa a tabela grande só para as
# telas que realmente precisam dela.

def meses_disponiveis_leves(dias_df, carro_dias_df=None):
    partes = []
    for df in [dias_df, carro_dias_df if carro_dias_df is not None else pd.DataFrame()]:
        if df is None or df.empty or "DATA" not in df.columns:
            continue
        datas = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
        parte = pd.DataFrame({"PERIODO": datas.dt.to_period("M")}).dropna()
        if not parte.empty:
            partes.append(parte)
    if not partes:
        return meses_disponiveis_rapido()
    meses = pd.concat(partes, ignore_index=True).drop_duplicates().sort_values("PERIODO", ascending=False)
    meses["MES"] = meses["PERIODO"].dt.strftime("%m/%Y")
    return meses[["MES", "PERIODO"]].reset_index(drop=True)


def _filtrar_df_por_meses_coluna_data(df, meses):
    if df is None or df.empty or not meses or "DATA" not in df.columns:
        return df.copy() if df is not None else pd.DataFrame()
    out = df.copy()
    datas = pd.to_datetime(out["DATA"], dayfirst=True, errors="coerce")
    return out[datas.dt.strftime("%m/%Y").isin(list(meses))].copy()


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def resumo_operacional_dia_cache(caminho_banco_str, meses_escolhidos, mtime_banco):
    """Lê resumo_dia do SQLite para recuperar CORTES/RELIGUES/RECUSAS da Home.

    A tabela faturamento_dias é leve e rápida, mas não guarda cortes/religues.
    Esses números ficam em resumo_dia, que é pequena e já existe no banco leve.
    """
    del mtime_banco
    if not caminho_banco_str:
        return pd.DataFrame(columns=["CONTRATO", "TOTAL_NOTAS_OP", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"])
    try:
        with sqlite3.connect(caminho_banco_str) as conn:
            if not _sqlite_tabela_existe(caminho_banco_str, "resumo_dia"):
                return pd.DataFrame(columns=["CONTRATO", "TOTAL_NOTAS_OP", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"])
            df = pd.read_sql_query('SELECT DATA, CONTRATO, TOTAL_NOTAS, CORTES, RELIGUES, COALESCE(VERIFICACOES, 0) AS VERIFICACOES, RECUSAS FROM resumo_dia', conn)
    except Exception:
        return pd.DataFrame(columns=["CONTRATO", "TOTAL_NOTAS_OP", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"])

    if df.empty:
        return pd.DataFrame(columns=["CONTRATO", "TOTAL_NOTAS_OP", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"])

    datas = pd.to_datetime(df["DATA"], dayfirst=True, errors="coerce")
    meses = list(meses_escolhidos or [])
    if meses:
        df = df[datas.dt.strftime("%m/%Y").isin(meses)].copy()
    if df.empty:
        return pd.DataFrame(columns=["CONTRATO", "TOTAL_NOTAS_OP", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"])

    for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return (
        df.groupby("CONTRATO", dropna=False)
        .agg(
            TOTAL_NOTAS_OP=("TOTAL_NOTAS", "sum"),
            CORTES=("CORTES", "sum"),
            RELIGUES=("RELIGUES", "sum"),
            VERIFICACOES=("VERIFICACOES", "sum"),
            RECUSAS=("RECUSAS", "sum"),
        )
        .reset_index()
    )


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def resumo_home_leve_cache(dias_df, carro_dias_df, meses_escolhidos, contrato_escolhido):
    dias_base = _filtrar_df_por_meses_coluna_data(dias_df, tuple(meses_escolhidos or []))
    carro_base = _filtrar_df_por_meses_coluna_data(carro_dias_df, tuple(meses_escolhidos or []))

    if contrato_escolhido != "Todos" and not dias_base.empty and "CONTRATO" in dias_base.columns:
        dias_base = dias_base[dias_base["CONTRATO"] == contrato_escolhido]
    if contrato_escolhido != "Todos" and not carro_base.empty and "CONTRATO" in carro_base.columns:
        # Quando o usuário escolhe STC Jundiai, o carro estimado continua sendo útil.
        if contrato_escolhido in ["STC Jundiai", "Contrato Carro STC estimado"]:
            pass
        else:
            carro_base = carro_base.iloc[0:0]

    resumo_contrato = pd.DataFrame()
    if not dias_base.empty:
        for col in ["QTD_NOTAS", "FATURAMENTO"]:
            if col not in dias_base.columns:
                dias_base[col] = 0
            dias_base[col] = pd.to_numeric(dias_base[col], errors="coerce").fillna(0)
        resumo_contrato = (
            dias_base.groupby("CONTRATO", dropna=False)
            .agg(TOTAL_NOTAS=("QTD_NOTAS", "sum"), FATURAMENTO=("FATURAMENTO", "sum"))
            .reset_index()
        )
        resumo_contrato["CORTES"] = 0
        resumo_contrato["RELIGUES"] = 0
        resumo_contrato["FATURAMENTO_MIN"] = 0.0
        resumo_contrato["FATURAMENTO_MAX"] = 0.0

    if not carro_base.empty:
        for col in ["QTD_NOTAS", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
            if col not in carro_base.columns:
                carro_base[col] = 0
            carro_base[col] = pd.to_numeric(carro_base[col], errors="coerce").fillna(0)
        carro_resumo = (
            carro_base.groupby("CONTRATO", dropna=False)
            .agg(
                TOTAL_NOTAS=("QTD_NOTAS", "sum"),
                FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
                FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
            )
            .reset_index()
        )
        carro_resumo["FATURAMENTO"] = 0.0
        carro_resumo["CORTES"] = 0
        carro_resumo["RELIGUES"] = 0
        if resumo_contrato.empty:
            resumo_contrato = carro_resumo
        else:
            resumo_contrato = pd.concat([resumo_contrato, carro_resumo], ignore_index=True, sort=False)

    if resumo_contrato.empty:
        return pd.DataFrame(), pd.DataFrame()

    for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES"]:
        if col not in resumo_contrato.columns:
            resumo_contrato[col] = 0
        resumo_contrato[col] = pd.to_numeric(resumo_contrato[col], errors="coerce").fillna(0).astype(int)
    for col in ["FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
        if col not in resumo_contrato.columns:
            resumo_contrato[col] = 0.0
        resumo_contrato[col] = pd.to_numeric(resumo_contrato[col], errors="coerce").fillna(0.0)

    # Recupera cortes/religues/recusas do resumo_dia para a Home.
    # Isso corrige o detalhamento por contrato no Resumo, que usa faturamento_dias
    # para ser rápido, mas precisa do resumo_dia para os contadores operacionais.
    try:
        caminho = caminho_banco_gzus()
        if sqlite_ativado() and caminho and _sqlite_tabela_existe(caminho, "resumo_dia"):
            resumo_ops = resumo_operacional_dia_cache(str(caminho), tuple(meses_escolhidos or []), caminho.stat().st_mtime)
            if not resumo_ops.empty:
                resumo_contrato = resumo_contrato.merge(resumo_ops, on="CONTRATO", how="left", suffixes=("", "_OP"))
                for col in ["CORTES", "RELIGUES", "RECUSAS"]:
                    op_col = f"{col}_OP"
                    if op_col in resumo_contrato.columns:
                        resumo_contrato[col] = pd.to_numeric(resumo_contrato[op_col], errors="coerce").fillna(resumo_contrato[col]).fillna(0).astype(int)
                        resumo_contrato = resumo_contrato.drop(columns=[op_col])
                if "TOTAL_NOTAS_OP" in resumo_contrato.columns:
                    resumo_contrato = resumo_contrato.drop(columns=["TOTAL_NOTAS_OP"])
    except Exception:
        pass

    resumo_contrato = resumo_contrato.sort_values(["FATURAMENTO", "FATURAMENTO_MAX"], ascending=False).reset_index(drop=True)
    resumo_grupo = pd.DataFrame()
    return resumo_contrato, resumo_grupo


def resumo_home_leve(dias_df, carro_dias_df, meses_escolhidos, contrato_escolhido):
    return resumo_home_leve_cache(dias_df, carro_dias_df, tuple(meses_escolhidos or []), contrato_escolhido)


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


def contrato_para_base_notas(contrato):
    """Mapeia contratos de visão/estimativa para o contrato operacional da base de notas.

    Alguns botões da lateral representam visões estimadas (ex.: contrato do carro),
    mas a aba Parcial/Ranking/Comparativo usa a base de notas, onde essas notas
    entram como STC Jundiai. Sem esse mapeamento, a tela fica sem datas mesmo
    quando existe produção do carro/STC no dia.
    """
    nome = str(contrato or "").strip()
    nome_upper = nome.upper()
    if "CARRO" in nome_upper and "STC" in nome_upper:
        return "STC Jundiai"
    return nome



def normalizar_grupo_nota(valor):
    texto = str(valor or '').strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    if 'VERIFIC' in texto:
        return 'VERIFICACAO'
    if 'RELIG' in texto:
        return 'RELIGUE'
    if 'CORTE' in texto:
        return 'CORTE'
    return texto

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

    df["GRUPO_NOTA"] = df["GRUPO_NOTA"].apply(normalizar_grupo_nota)
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
                faturamento = {"CORTE": secret_float("TARIFA_DISJUNTOR_JUNDIAI_CORTE", 13.72), "VERIFICACAO": secret_float("TARIFA_DISJUNTOR_JUNDIAI_CORTE", 13.72), "RELIGUE": secret_float("TARIFA_DISJUNTOR_JUNDIAI_RELIGUE", 27.43)}.get(grupo, 0.0)
                faturamento_min = faturamento
                faturamento_max = faturamento

        elif eh_disjuntor_santa_cruz(recurso):
            contrato = "Disjuntor Santa Cruz"
            if not eh_recusa:
                faturamento = {"CORTE": secret_float("TARIFA_DISJUNTOR_SANTA_CRUZ_CORTE", 11.98), "VERIFICACAO": secret_float("TARIFA_DISJUNTOR_SANTA_CRUZ_VERIFICACAO", 23.97), "RELIGUE": secret_float("TARIFA_DISJUNTOR_SANTA_CRUZ_RELIGUE", 23.97)}.get(grupo, 0.0)
                faturamento_min = faturamento
                faturamento_max = faturamento

        elif str(recurso).startswith("JUN58") and qtd_exec >= 2:
            contrato = "STC Jundiai"
            if not eh_recusa:
                faturamento_min = {"CORTE": secret_float("TARIFA_STC_JUNDIAI_CORTE_MIN", 38.18), "RELIGUE": secret_float("TARIFA_STC_JUNDIAI_RELIGUE_MIN", 36.36)}.get(grupo, 0.0)
                faturamento_max = {"CORTE": secret_float("TARIFA_STC_JUNDIAI_CORTE_MAX", 45.45), "RELIGUE": secret_float("TARIFA_STC_JUNDIAI_RELIGUE_MAX", 50.91)}.get(grupo, 0.0)
                faturamento = faturamento_min

        if contrato:
            item = row.to_dict()
            item["CONTRATO"] = contrato
            item["FATURAMENTO"] = faturamento
            item["FATURAMENTO_MIN"] = faturamento_min
            item["FATURAMENTO_MAX"] = faturamento_max
            item["EH_CORTE"] = 1 if (((grupo == "CORTE") or (grupo == "VERIFICACAO" and contrato == "Disjuntor Jundiaí")) and not eh_recusa) else 0
            item["EH_RELIGUE"] = 1 if (grupo == "RELIGUE" and not eh_recusa) else 0
            item["EH_VERIFICACAO"] = 1 if (grupo == "VERIFICACAO" and contrato == "Disjuntor Santa Cruz" and not eh_recusa) else 0
            item["EH_RECUSA"] = 1 if eh_recusa else 0
            linhas.append(item)

    if not linhas:
        return pd.DataFrame()

    parcial = pd.DataFrame(linhas)
    if "EH_VERIFICACAO" not in parcial.columns:
        parcial["EH_VERIFICACAO"] = 0
    parcial["DATA_DT"] = pd.to_datetime(parcial["DATA"], dayfirst=True, errors="coerce")
    parcial = parcial.dropna(subset=["DATA_DT"])
    parcial["DATA"] = parcial["DATA_DT"].dt.strftime("%d/%m/%Y")

    return parcial




def calcular_recursos_sem_movimento_no_dia(parcial_com_recusas, data_escolhida):
    """Identifica equipes esperadas que não tiveram movimento algum no dia.

    Regra para contratos de corte:
    - A base esperada vem das equipes que aparecem no mesmo mês e contrato
      até a data escolhida.
    - Movimento do dia considera qualquer ocorrência: nota feita OU recusa.
    - Portanto, equipe com somente recusa NÃO entra no alerta de zero movimento.
    """
    colunas_saida = ["RECURSO", "CONTRATO"]
    if parcial_com_recusas is None or parcial_com_recusas.empty:
        return pd.DataFrame(columns=colunas_saida)

    base = parcial_com_recusas.copy()
    for col in ["RECURSO", "CONTRATO", "DATA", "DATA_DT"]:
        if col not in base.columns:
            return pd.DataFrame(columns=colunas_saida)

    base["RECURSO"] = base["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    base["CONTRATO"] = base["CONTRATO"].fillna("").astype(str).str.strip()
    base["DATA_DT"] = pd.to_datetime(base["DATA_DT"], dayfirst=True, errors="coerce")
    data_ref = pd.to_datetime(data_escolhida, dayfirst=True, errors="coerce")

    if pd.isna(data_ref):
        return pd.DataFrame(columns=colunas_saida)

    base = base[(base["RECURSO"] != "") & (base["CONTRATO"] != "") & base["DATA_DT"].notna()].copy()
    if base.empty:
        return pd.DataFrame(columns=colunas_saida)

    mesmo_mes_ate_dia = (
        (base["DATA_DT"].dt.year == data_ref.year)
        & (base["DATA_DT"].dt.month == data_ref.month)
        & (base["DATA_DT"] <= data_ref)
    )

    recursos_esperados = (
        base.loc[mesmo_mes_ate_dia, colunas_saida]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    if recursos_esperados.empty:
        return pd.DataFrame(columns=colunas_saida)

    movimento_dia = (
        base.loc[base["DATA_DT"] == data_ref, colunas_saida]
        .drop_duplicates()
        .assign(_MOVIMENTOU=1)
    )

    sem_movimento = recursos_esperados.merge(
        movimento_dia,
        on=colunas_saida,
        how="left",
    )
    sem_movimento = sem_movimento[sem_movimento["_MOVIMENTOU"].isna()].drop(columns=["_MOVIMENTOU"], errors="ignore")
    return sem_movimento.sort_values(["CONTRATO", "RECURSO"]).reset_index(drop=True)



@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def calcular_parcial_dia_processada_cache(parcial_com_recusas, data_escolhida):
    """Pré-processa a parcial de um dia e mantém o resultado em cache.

    Trocar o selectbox de dia fazia o Streamlit recalcular filtros, agrupamentos,
    recusas e alerta de equipes sem movimento em cima da base inteira. Como dia
    anterior quase não muda, este cache guarda o resultado por data/contrato já
    filtrado. Quando os arquivos mudam ou o cache é limpo pelo sync do GitHub, o
    cálculo é refeito automaticamente.
    """
    resultado_vazio = {
        "parcial_dia_tudo": pd.DataFrame(),
        "parcial_dia": pd.DataFrame(),
        "recusas_dia": pd.DataFrame(),
        "resumo_equipe": pd.DataFrame(),
        "recursos_sem_movimento": pd.DataFrame(),
        "totais": {
            "total_notas": 0,
            "total_recursos_ativos": 0,
            "total_cortes": 0,
            "total_religues": 0,
            "total_recusas": 0,
            "total_faturamento": 0.0,
            "total_faturamento_min": 0.0,
            "total_faturamento_max": 0.0,
        },
    }

    if parcial_com_recusas is None or parcial_com_recusas.empty or not data_escolhida:
        return resultado_vazio

    base = parcial_com_recusas.copy()
    if "DATA" not in base.columns:
        return resultado_vazio

    parcial_dia_tudo = base[base["DATA"] == data_escolhida].copy()
    if parcial_dia_tudo.empty:
        return resultado_vazio

    if "EH_RECUSA" not in parcial_dia_tudo.columns:
        parcial_dia_tudo["EH_RECUSA"] = 0
    parcial_dia_tudo["EH_RECUSA"] = pd.to_numeric(parcial_dia_tudo["EH_RECUSA"], errors="coerce").fillna(0).astype(int)

    parcial_dia = parcial_dia_tudo[parcial_dia_tudo["EH_RECUSA"] == 0].copy()
    recusas_dia = parcial_dia_tudo[parcial_dia_tudo["EH_RECUSA"] == 1].copy()

    for df_tmp in [parcial_dia, recusas_dia]:
        for col in ["ORDEM_DE_SERVICO", "RECURSO", "CONTRATO"]:
            if col not in df_tmp.columns:
                df_tmp[col] = ""
        for col in ["EH_CORTE", "EH_RELIGUE", "EH_VERIFICACAO"]:
            if col not in df_tmp.columns:
                df_tmp[col] = 0
            df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0).astype(int)
        for col in ["FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
            if col not in df_tmp.columns:
                df_tmp[col] = 0.0
            df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0.0)

    totais = {
        "total_notas": parcial_dia["ORDEM_DE_SERVICO"].nunique() if not parcial_dia.empty else 0,
        "total_recursos_ativos": parcial_dia["RECURSO"].nunique() if not parcial_dia.empty else 0,
        "total_cortes": int(parcial_dia["EH_CORTE"].sum()) if not parcial_dia.empty else 0,
        "total_religues": int(parcial_dia["EH_RELIGUE"].sum()) if not parcial_dia.empty else 0,
        "total_verificacoes": int(parcial_dia["EH_VERIFICACAO"].sum()) if not parcial_dia.empty and "EH_VERIFICACAO" in parcial_dia.columns else 0,
        "total_recusas": recusas_dia["ORDEM_DE_SERVICO"].nunique() if not recusas_dia.empty else 0,
        "total_faturamento": float(parcial_dia["FATURAMENTO"].sum()) if not parcial_dia.empty else 0.0,
        "total_faturamento_min": float(parcial_dia["FATURAMENTO_MIN"].sum()) if not parcial_dia.empty else 0.0,
        "total_faturamento_max": float(parcial_dia["FATURAMENTO_MAX"].sum()) if not parcial_dia.empty else 0.0,
    }

    resumo_producao = (
        parcial_dia.groupby(["RECURSO", "CONTRATO"], dropna=False)
        .agg(
            TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            FATURAMENTO=("FATURAMENTO", "sum"),
            FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
            FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
        )
        .reset_index()
        if not parcial_dia.empty
        else pd.DataFrame(columns=[
            "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES",
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
        "VERIFICACOES": 0,
        "FATURAMENTO": 0,
        "FATURAMENTO_MIN": 0,
        "FATURAMENTO_MAX": 0,
        "RECUSAS": 0,
    })

    for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS"]:
        if col in resumo_equipe.columns:
            resumo_equipe[col] = pd.to_numeric(resumo_equipe[col], errors="coerce").fillna(0).astype(int)

    for col in ["FATURAMENTO", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]:
        if col in resumo_equipe.columns:
            resumo_equipe[col] = pd.to_numeric(resumo_equipe[col], errors="coerce").fillna(0.0)

    resumo_equipe = (
        resumo_equipe
        .sort_values(["TOTAL_NOTAS", "FATURAMENTO", "RECUSAS"], ascending=[False, False, False])
        .reset_index(drop=True)
    )

    recursos_sem_movimento = calcular_recursos_sem_movimento_no_dia(base, data_escolhida)

    return {
        "parcial_dia_tudo": parcial_dia_tudo,
        "parcial_dia": parcial_dia,
        "recusas_dia": recusas_dia,
        "resumo_equipe": resumo_equipe,
        "recursos_sem_movimento": recursos_sem_movimento,
        "totais": totais,
    }


def render_alerta_recursos_sem_movimento(recursos_sem_movimento, contrato_unico=False):
    """Mostra o alerta de equipes sem nenhuma nota ou recusa no dia."""
    if recursos_sem_movimento is None or recursos_sem_movimento.empty:
        return

    base = recursos_sem_movimento.copy()
    base["RECURSO"] = base["RECURSO"].fillna("").astype(str).str.strip()
    base["CONTRATO"] = base["CONTRATO"].fillna("").astype(str).str.strip()

    if contrato_unico:
        itens = base["RECURSO"].drop_duplicates().astype(str).head(12).tolist()
    else:
        itens = (base["RECURSO"] + " (" + base["CONTRATO"] + ")").drop_duplicates().astype(str).head(12).tolist()

    lista = ", ".join(itens)
    total = len(base.drop_duplicates(subset=["RECURSO", "CONTRATO"]))
    if total > 12:
        lista += f" e mais {total - 12}"

    st.markdown(
        f"""
        <div class="zero-card">
            ⚠️ <b>Recursos sem nenhum movimento no dia:</b> {numero(total)}<br>
            <span>Sem nota feita e sem recusa registrada. Equipe com somente recusa não entra neste alerta.</span><br>
            {lista}
        </div>
        """,
        unsafe_allow_html=True,
    )


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
        meses_base = meses_disponiveis_rapido()
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
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
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
            "VERIFICACOES": 0,
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
        "VERIFICACOES": int(resumo_contrato["VERIFICACOES"].sum()) if "VERIFICACOES" in resumo_contrato.columns else 0,
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


def caminho_depara_express():
    """Procura uma tabela opcional de DE/PARA do Pagamento Express.

    Isso evita depender exclusivamente do st.secrets. Se o arquivo existir no
    repositório, o painel consegue mapear Nome -> Recurso automaticamente.

    Nomes aceitos:
    - depara_nome_recurso_express.csv / .xlsx
    - depara_pagamento_express.csv / .xlsx
    - depara_express.csv / .xlsx

    Colunas aceitas para nome: NOME, NOME_EXPRESS, NOME_EXECUTOR_01, DE.
    Colunas aceitas para recurso: RECURSO, EQUIPE, PARA, CODIGO_RECURSO.
    """
    nomes = [
        "depara_nome_recurso_express.csv",
        "depara_nome_recurso_express.xlsx",
        "depara_pagamento_express.csv",
        "depara_pagamento_express.xlsx",
        "depara_express.csv",
        "depara_express.xlsx",
    ]
    for nome in nomes:
        for pasta in [PASTA_DASHBOARD, PASTA_ATUAL]:
            caminho = pasta / nome
            if caminho.exists():
                return caminho
    return None


def carregar_depara_nome_recurso_express():
    """Carrega DE/PARA Nome -> Recurso via Secrets e/ou arquivo no dashboard.

    A ordem é proposital:
    1. Secrets continuam funcionando para quem já configurou.
    2. Arquivo CSV/XLSX no dashboard complementa ou sobrescreve entradas.

    Assim o Express não zera quando o código está público e o mapa deixou de
    estar hardcoded no .py.
    """
    mapa = {
        normalizar_nome_pessoa(nome): str(recurso).strip().upper()
        for nome, recurso in secret_dict("DEPARA_NOME_RECURSO_EXPRESS", {}).items()
        if str(nome).strip() and str(recurso).strip()
    }

    caminho = caminho_depara_express()
    if not caminho:
        return mapa

    try:
        if str(caminho).lower().endswith(".xlsx"):
            df = pd.read_excel(caminho, engine="openpyxl")
        else:
            try:
                df = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
            except Exception:
                df = pd.read_csv(caminho, sep=",", encoding="utf-8-sig")
    except Exception:
        return mapa

    if df.empty:
        return mapa

    df.columns = [str(c).strip().upper() for c in df.columns]
    colunas_norm = {normalizar_nome_pessoa(c): c for c in df.columns}

    def achar_coluna(*nomes):
        for nome in nomes:
            chave = normalizar_nome_pessoa(nome)
            if chave in colunas_norm:
                return colunas_norm[chave]
        return None

    col_nome = achar_coluna(
        "NOME", "NOME_EXPRESS", "NOME EXECUTOR", "NOME_EXECUTOR",
        "NOME_EXECUTOR_01", "NOME EXECUTOR 01", "EXECUTOR", "DE"
    )
    col_recurso = achar_coluna(
        "RECURSO", "EQUIPE", "PREFIXO", "PARA", "DEPARA", "DE/PARA",
        "CODIGO_RECURSO", "COD RECURSO", "CÓDIGO RECURSO", "CODIGO", "CÓDIGO"
    )

    if not col_nome or not col_recurso:
        return mapa

    for _, row in df.iterrows():
        nome_norm = normalizar_nome_pessoa(row.get(col_nome, ""))
        recurso = str(row.get(col_recurso, "")).strip().upper()
        if nome_norm and recurso and recurso not in ["NAN", "NONE"]:
            mapa[nome_norm] = recurso

    return mapa


DEPARA_NOME_RECURSO_EXPRESS = carregar_depara_nome_recurso_express()


# Carro STC estimado - só conta quando a dupla completa bate.
# O valor mapeado é só o número; o app transforma em JUN58xx-EMP.
DEPARA_DUPLA_CARRO_EXPRESS = {
    frozenset([normalizar_nome_pessoa(parte) for parte in str(dupla).split("|") if str(parte).strip()]): str(recurso).strip()
    for dupla, recurso in secret_dict("DEPARA_DUPLA_CARRO_EXPRESS", {}).items()
    if "|" in str(dupla) and str(recurso).strip()
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
        return "STC Jundiai"
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



@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def mapa_nota_para_recurso_real(notas):
    """Monta mapa OS/nota -> recurso/contrato usando a própria base de notas.

    Este é o fallback mais seguro para o Pagamento Express quando o arquivo
    Express não traz RECURSO e o DE/PARA de nome não está configurado.
    """
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)
    colunas = ["ORDEM_DE_SERVICO", "RECURSO", "CONTRATO"]
    if parcial.empty or not set(colunas).issubset(set(parcial.columns)):
        return {}

    tmp = parcial[colunas].copy()
    tmp["NOTA_NORM"] = tmp["ORDEM_DE_SERVICO"].apply(normalizar_ordem_servico)
    tmp["RECURSO"] = tmp["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    tmp["CONTRATO"] = tmp["CONTRATO"].fillna("").astype(str).str.strip()
    tmp = tmp[(tmp["NOTA_NORM"] != "") & (tmp["RECURSO"] != "")].copy()
    if tmp.empty:
        return {}

    contagem = (
        tmp.groupby(["NOTA_NORM", "RECURSO", "CONTRATO"], dropna=False)
        .size()
        .reset_index(name="QTD")
        .sort_values(["NOTA_NORM", "QTD"], ascending=[True, False])
        .drop_duplicates(subset=["NOTA_NORM"], keep="first")
    )
    return {
        str(r["NOTA_NORM"]): (str(r["RECURSO"]).strip().upper(), str(r["CONTRATO"]).strip())
        for _, r in contagem.iterrows()
    }


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def mapa_nome_para_recurso_real(notas):
    """Monta mapa Nome/Executor -> recurso/contrato usando a base de notas.

    Se ELETRICISTA1/2 vierem como nome, este fallback resolve o Express sem
    precisar de DE/PARA manual. Se vierem como código, não atrapalha.
    """
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)
    if parcial.empty or "RECURSO" not in parcial.columns:
        return {}

    linhas = []
    for col in ["ELETRICISTA1", "ELETRICISTA2"]:
        if col in parcial.columns:
            cols = [col, "RECURSO"] + (["CONTRATO"] if "CONTRATO" in parcial.columns else [])
            tmp = parcial[cols].copy()
            tmp = tmp.rename(columns={col: "NOME_BASE"})
            tmp["NOME_NORM"] = tmp["NOME_BASE"].apply(normalizar_nome_pessoa)
            tmp["RECURSO"] = tmp["RECURSO"].fillna("").astype(str).str.strip().str.upper()
            if "CONTRATO" not in tmp.columns:
                tmp["CONTRATO"] = tmp["RECURSO"].apply(contrato_por_recurso_express)
            else:
                tmp["CONTRATO"] = tmp["CONTRATO"].fillna("").astype(str).str.strip()
            tmp = tmp[(tmp["NOME_NORM"] != "") & (tmp["RECURSO"] != "")].copy()
            if not tmp.empty:
                linhas.append(tmp[["NOME_NORM", "RECURSO", "CONTRATO"]])

    if not linhas:
        return {}

    base = pd.concat(linhas, ignore_index=True)
    contagem = (
        base.groupby(["NOME_NORM", "RECURSO", "CONTRATO"], dropna=False)
        .size()
        .reset_index(name="QTD")
        .sort_values(["NOME_NORM", "QTD"], ascending=[True, False])
        .drop_duplicates(subset=["NOME_NORM"], keep="first")
    )
    return {
        str(r["NOME_NORM"]): (str(r["RECURSO"]).strip().upper(), str(r["CONTRATO"]).strip())
        for _, r in contagem.iterrows()
    }


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

    # Algumas versões da planilha já trazem o recurso/equipe direto.
    # Quando existir, usamos como fallback/prioridade e não dependemos do DE/PARA.
    col_recurso_direto = achar_coluna(
        "RECURSO", "EQUIPE", "PREFIXO", "RECURSO_EQUIPE", "RECURSO EQUIPE",
        "CODIGO_RECURSO", "COD RECURSO", "CÓDIGO RECURSO",
        "CODIGO", "CÓDIGO", "PARA", "DEPARA", "DE/PARA", "RECURSO_DEPARA"
    )
    if col_recurso_direto:
        df["RECURSO_EXPRESS_DIRETO"] = df[col_recurso_direto].fillna("").astype(str).str.strip().str.upper()
    else:
        df["RECURSO_EXPRESS_DIRETO"] = ""

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
    df.attrs["EXPRESS_COL_RECURSO_DIRETO"] = col_recurso_direto or ""

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
    Express faturado conforme tarifas configuradas em Secrets.
    Carro é diferente; por enquanto fica zerado no faturamento express.
    """
    contrato = str(contrato)
    if contrato == "Disjuntor Jundiaí":
        return secret_float("TARIFA_EXPRESS_DISJUNTOR_JUNDIAI", 27.43)
    if contrato == "Disjuntor Santa Cruz":
        return secret_float("TARIFA_EXPRESS_DISJUNTOR_SANTA_CRUZ", 23.97)
    if contrato == "STC Jundiai":
        return secret_float("TARIFA_EXPRESS_STC_JUNDIAI", 38.18)
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
    mapa_nota_recurso = mapa_nota_para_recurso_real(notas)
    mapa_nome_recurso = mapa_nome_para_recurso_real(notas)

    # Jundiaí/Santa Cruz: primeiro tenta o DE/PARA manual por nome.
    express["RECURSO_DEPARA"] = express["NOME_EXPRESS_NORM"].map(DEPARA_NOME_RECURSO_EXPRESS).fillna("")

    # Fallback/prioridade 1: se a própria planilha Express trouxer RECURSO/EQUIPE/PARA,
    # usa esse valor.
    if "RECURSO_EXPRESS_DIRETO" in express.columns:
        recurso_direto = express["RECURSO_EXPRESS_DIRETO"].fillna("").astype(str).str.strip().str.upper()
        mascara_recurso_direto = (recurso_direto != "") & (~recurso_direto.isin(["NAN", "NONE"]))
        express.loc[mascara_recurso_direto, "RECURSO_DEPARA"] = recurso_direto[mascara_recurso_direto]

    # Fallback 2: cruza pela OS/NOTA com a base de notas.
    # Este é o principal conserto para quando o DE/PARA saiu do código/secrets.
    if "NOTA_NORM" in express.columns and mapa_nota_recurso:
        mascara_sem_recurso = express["RECURSO_DEPARA"].fillna("").astype(str).str.strip().eq("")
        recursos_por_nota = express.loc[mascara_sem_recurso, "NOTA_NORM"].map(
            lambda n: mapa_nota_recurso.get(str(n), ("", ""))[0]
        )
        express.loc[mascara_sem_recurso, "RECURSO_DEPARA"] = recursos_por_nota.fillna("")

    # Fallback 3: se ELETRICISTA1/2 da base tiverem os mesmos nomes do Express,
    # resolve por nome sem precisar de DE/PARA manual.
    if mapa_nome_recurso:
        mascara_sem_recurso = express["RECURSO_DEPARA"].fillna("").astype(str).str.strip().eq("")
        recursos_por_nome = express.loc[mascara_sem_recurso, "NOME_EXPRESS_NORM"].map(
            lambda n: mapa_nome_recurso.get(str(n), ("", ""))[0]
        )
        express.loc[mascara_sem_recurso, "RECURSO_DEPARA"] = recursos_por_nome.fillna("")

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



# ==============================
# META CPFL - CONTRATO CARRO / STC
# ==============================
# Primeira versão aplicada somente para o contrato operacional STC Jundiai.
# Regra: Meta CPFL considera CORTES + Pagamento Express de corte.

FERIADOS_CPFL = {
    "2026-04-03",  # Sexta-feira Santa
    "2026-04-21",  # Tiradentes
    "2026-05-01",  # Dia do Trabalho
}

METAS_CPFL_STC = {
    "03/2026": {"util": 224, "sexta_vespera": 137, "sabado": 46},
    "04/2026": {"util": 262, "sexta_vespera": 149, "sabado": 50},
    "05/2026": {"util": 247, "sexta_vespera": 143, "sabado": 48},
}


def _data_para_timestamp(valor):
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def _eh_feriado_cpfl(data_ts):
    if pd.isna(data_ts):
        return False
    return data_ts.strftime("%Y-%m-%d") in FERIADOS_CPFL


def _eh_vespera_feriado_cpfl(data_ts):
    if pd.isna(data_ts):
        return False
    proximo = data_ts + pd.Timedelta(days=1)
    return proximo.strftime("%Y-%m-%d") in FERIADOS_CPFL


def meta_cpfl_stc_dia(data_valor):
    """Retorna a meta diária CPFL para STC Jundiai conforme mês/dia da semana."""
    data_ts = _data_para_timestamp(data_valor)
    if pd.isna(data_ts):
        return 0

    mes = data_ts.strftime("%m/%Y")
    regra = METAS_CPFL_STC.get(mes)
    if not regra:
        return 0

    if _eh_feriado_cpfl(data_ts):
        return 0

    dia_semana = int(data_ts.weekday())  # segunda=0 ... domingo=6
    if dia_semana == 6:
        return 0
    if dia_semana == 5:
        return int(regra.get("sabado", 0))
    if dia_semana == 4 or _eh_vespera_feriado_cpfl(data_ts):
        return int(regra.get("sexta_vespera", 0))
    return int(regra.get("util", 0))


def meta_cpfl_stc_periodo(inicio, fim):
    inicio_ts = _data_para_timestamp(inicio)
    fim_ts = _data_para_timestamp(fim)
    if pd.isna(inicio_ts) or pd.isna(fim_ts):
        return 0
    dias = pd.date_range(inicio_ts.normalize(), fim_ts.normalize(), freq="D")
    return int(sum(meta_cpfl_stc_dia(d) for d in dias))


def _periodo_datas_cpfl(tipo_periodo, valor_periodo):
    """Converte seleção Dia/Semana/Mês em intervalo de datas."""
    if tipo_periodo == "Dia" and valor_periodo:
        dt = _data_para_timestamp(valor_periodo)
        return dt, dt
    if tipo_periodo == "Semana" and valor_periodo:
        inicio = _data_para_timestamp(valor_periodo)
        return inicio, inicio + pd.Timedelta(days=6)
    if tipo_periodo == "Mês" and valor_periodo:
        inicio = pd.to_datetime("01/" + str(valor_periodo), dayfirst=True, errors="coerce")
        if pd.isna(inicio):
            return pd.NaT, pd.NaT
        fim = inicio + pd.offsets.MonthEnd(0)
        return inicio, fim
    return pd.NaT, pd.NaT


@st.cache_data(ttl=CACHE_TTL_RANKING_SEGUNDOS, show_spinner=False)
def express_detalhado_cpfl_cache(notas):
    """Retorna as linhas de Pagamento Express já conciliadas com RECURSO e CONTRATO."""
    caminho = caminho_pagamento_express()
    if not caminho:
        return pd.DataFrame()

    express = ler_pagamento_express(str(caminho))
    if express.empty:
        return pd.DataFrame()

    mapa_codigo_recurso = mapa_codigo_para_recurso_real(notas)
    mapa_nota_recurso = mapa_nota_para_recurso_real(notas)
    mapa_nome_recurso = mapa_nome_para_recurso_real(notas)

    express = express.copy()
    express["RECURSO_DEPARA"] = express.get("NOME_EXPRESS_NORM", pd.Series(dtype=object)).map(DEPARA_NOME_RECURSO_EXPRESS).fillna("")

    if "RECURSO_EXPRESS_DIRETO" in express.columns:
        recurso_direto = express["RECURSO_EXPRESS_DIRETO"].fillna("").astype(str).str.strip().str.upper()
        mascara_recurso_direto = (recurso_direto != "") & (~recurso_direto.isin(["NAN", "NONE"]))
        express.loc[mascara_recurso_direto, "RECURSO_DEPARA"] = recurso_direto[mascara_recurso_direto]

    if "NOTA_NORM" in express.columns and mapa_nota_recurso:
        mascara_sem_recurso = express["RECURSO_DEPARA"].fillna("").astype(str).str.strip().eq("")
        recursos_por_nota = express.loc[mascara_sem_recurso, "NOTA_NORM"].map(
            lambda n: mapa_nota_recurso.get(str(n), ("", ""))[0]
        )
        express.loc[mascara_sem_recurso, "RECURSO_DEPARA"] = recursos_por_nota.fillna("")

    if mapa_nome_recurso:
        mascara_sem_recurso = express["RECURSO_DEPARA"].fillna("").astype(str).str.strip().eq("")
        recursos_por_nome = express.loc[mascara_sem_recurso, "NOME_EXPRESS_NORM"].map(
            lambda n: mapa_nome_recurso.get(str(n), ("", ""))[0]
        )
        express.loc[mascara_sem_recurso, "RECURSO_DEPARA"] = recursos_por_nome.fillna("")

    if "NOME_EXPRESS_02_NORM" not in express.columns:
        express["NOME_EXPRESS_02_NORM"] = ""

    express["RECURSO_CARRO"] = express.apply(
        lambda r: recurso_carro_por_dupla(
            r.get("NOME_EXPRESS_NORM", ""),
            r.get("NOME_EXPRESS_02_NORM", ""),
        ),
        axis=1,
    )
    mascara_carro = express["RECURSO_CARRO"].fillna("").astype(str).str.strip() != ""
    express.loc[mascara_carro, "RECURSO_DEPARA"] = express.loc[mascara_carro, "RECURSO_CARRO"]

    express["RECURSO"] = express["RECURSO_DEPARA"].apply(lambda v: resolver_recurso_depara(v, mapa_codigo_recurso))
    express["RECURSO"] = express["RECURSO"].fillna("").astype(str).str.strip().str.upper()
    express["CONTRATO"] = express["RECURSO"].apply(contrato_por_recurso_express)

    express_ok = express[(express["RECURSO"] != "") & (express["CONTRATO"] != "")].copy()
    return express_ok


def contar_express_cpfl_periodo(notas, contrato, inicio, fim):
    express = express_detalhado_cpfl_cache(notas)
    if express.empty or "DATA_EXPRESS_DT" not in express.columns:
        return 0

    inicio_ts = _data_para_timestamp(inicio)
    fim_ts = _data_para_timestamp(fim)
    if pd.isna(inicio_ts) or pd.isna(fim_ts):
        return 0

    base = express.copy()
    base["DATA_EXPRESS_DT"] = pd.to_datetime(base["DATA_EXPRESS_DT"], errors="coerce")
    base = base[base["DATA_EXPRESS_DT"].notna()].copy()

    if contrato != "Todos" and "CONTRATO" in base.columns:
        base = base[base["CONTRATO"] == contrato].copy()

    base = base[
        (base["DATA_EXPRESS_DT"].dt.normalize() >= inicio_ts.normalize())
        & (base["DATA_EXPRESS_DT"].dt.normalize() <= fim_ts.normalize())
    ].copy()
    return int(len(base))


def render_meta_cpfl_stc(titulo, meta, cortes_feitos, express_feitos):
    total_feito = int(cortes_feitos) + int(express_feitos)
    saldo = total_feito - int(meta)
    execucao = (total_feito / meta * 100) if meta else 0

    st.markdown(f'<div class="section-title">{titulo}</div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Meta CPFL", numero(meta))
    c2.metric("Cortes feitos", numero(cortes_feitos))
    c3.metric("Express feitos", numero(express_feitos))
    c4.metric("Total CPFL", numero(total_feito))
    c5.metric("Execução", f"{execucao:.1f}%".replace(".", ","), numero(saldo))



def _periodos_meses_cpfl_ate_momento(meses):
    """Retorna períodos mensais para meta CPFL acumulada.

    Para o mês atual, considera somente até hoje.
    Para meses anteriores, considera o mês inteiro.
    Meses futuros ficam com período vazio para não projetar execução antes da hora.
    """
    hoje = pd.Timestamp(datetime.now(ZoneInfo("America/Sao_Paulo")).date())
    periodos = []
    for mes in meses or []:
        inicio = pd.to_datetime("01/" + str(mes), dayfirst=True, errors="coerce")
        if pd.isna(inicio):
            continue
        fim_mes = inicio + pd.offsets.MonthEnd(0)
        if inicio > hoje:
            continue
        fim = min(fim_mes, hoje) if (inicio.year == hoje.year and inicio.month == hoje.month) else fim_mes
        periodos.append((inicio, fim))
    return periodos


def meta_cpfl_stc_meses_ate_momento(meses):
    return int(sum(meta_cpfl_stc_periodo(inicio, fim) for inicio, fim in _periodos_meses_cpfl_ate_momento(meses)))


def contar_cortes_cpfl_stc_meses_ate_momento(notas, meses, contrato="STC Jundiai"):
    parcial = preparar_parcial_do_dia(notas, incluir_recusas=True)
    if parcial.empty or "DATA_DT" not in parcial.columns or "EH_CORTE" not in parcial.columns:
        return 0
    base = parcial.copy()
    base["DATA_DT"] = pd.to_datetime(base["DATA_DT"], dayfirst=True, errors="coerce")
    base = base[base["DATA_DT"].notna()].copy()
    if contrato != "Todos" and "CONTRATO" in base.columns:
        base = base[base["CONTRATO"] == contrato].copy()
    if "EH_RECUSA" in base.columns:
        base = base[pd.to_numeric(base["EH_RECUSA"], errors="coerce").fillna(0).astype(int) == 0].copy()
    total = 0
    for inicio, fim in _periodos_meses_cpfl_ate_momento(meses):
        recorte = base[(base["DATA_DT"].dt.normalize() >= inicio.normalize()) & (base["DATA_DT"].dt.normalize() <= fim.normalize())]
        total += int(pd.to_numeric(recorte["EH_CORTE"], errors="coerce").fillna(0).astype(int).sum())
    return int(total)


def contar_express_cpfl_stc_meses_ate_momento(notas, meses, contrato="STC Jundiai"):
    express = express_detalhado_cpfl_cache(notas)
    if express.empty or "DATA_EXPRESS_DT" not in express.columns:
        return 0
    base = express.copy()
    base["DATA_EXPRESS_DT"] = pd.to_datetime(base["DATA_EXPRESS_DT"], errors="coerce")
    base = base[base["DATA_EXPRESS_DT"].notna()].copy()
    if contrato != "Todos" and "CONTRATO" in base.columns:
        base = base[base["CONTRATO"] == contrato].copy()
    total = 0
    for inicio, fim in _periodos_meses_cpfl_ate_momento(meses):
        recorte = base[(base["DATA_EXPRESS_DT"].dt.normalize() >= inicio.normalize()) & (base["DATA_EXPRESS_DT"].dt.normalize() <= fim.normalize())]
        total += int(len(recorte))
    return int(total)


def render_auditoria_express_ranking(
    tipo_periodo, valor_periodo, express_caminho, express_data_max, express_sem_vinculo,
    express_resumo_recurso, total_express_mensal
):
    """Mostra a auditoria do Pagamento Express no fim do Ranking de recursos."""
    if not (tipo_periodo == "Mês" and valor_periodo):
        return

    st.markdown('<div class="section-title">Auditoria do Pagamento Express</div>', unsafe_allow_html=True)
    if express_caminho:
        if express_data_max:
            st.info(f"Pagamento Express conciliado por DE/PARA Nome → Recurso até {express_data_max}.")
        else:
            st.info("Pagamento Express conciliado por DE/PARA Nome → Recurso. A planilha não trouxe data válida para exibir o limite.")
    else:
        st.caption("Pagamento Express: arquivo não localizado.")

    if not express_sem_vinculo.empty:
        st.warning(f"Pagamento Express: {numero(len(express_sem_vinculo))} linha(s) não encontraram nome no DE/PARA Nome → Recurso.")

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
                        "nomes_mapeados_por_depara": int(express_debug.get("NOME_EXPRESS_NORM", pd.Series(dtype=object)).map(DEPARA_NOME_RECURSO_EXPRESS).fillna("").ne("").sum()) if "NOME_EXPRESS_NORM" in express_debug.columns else 0,
                        "recursos_diretos_no_excel": int(express_debug.get("RECURSO_EXPRESS_DIRETO", pd.Series(dtype=object)).fillna("").astype(str).str.strip().replace({"nan":"", "NaN":"", "None":""}).ne("").sum()) if "RECURSO_EXPRESS_DIRETO" in express_debug.columns else 0,
                        "tamanho_depara_nome_recurso": len(DEPARA_NOME_RECURSO_EXPRESS),
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

    for col in ["TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES"]:
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

    for col in ["NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "DIAS_ATIVOS", "QTD_EQUIPES"]:
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
        "verificacoes": 0,
        "por_contrato": {},
    }

    if parcial.empty:
        return resumo

    ultima_data_dt = parcial["DATA_DT"].max()
    parcial_dia = parcial[parcial["DATA_DT"] == ultima_data_dt].copy()

    resumo["data"] = ultima_data_dt.strftime("%d/%m/%Y")
    resumo["notas"] = int(parcial_dia["ORDEM_DE_SERVICO"].nunique())
    if "EH_VERIFICACAO" not in parcial_dia.columns:
        parcial_dia["EH_VERIFICACAO"] = 0
    resumo["cortes"] = int(parcial_dia["EH_CORTE"].sum()) + int(parcial_dia["EH_VERIFICACAO"].sum())
    resumo["religues"] = int(parcial_dia["EH_RELIGUE"].sum())
    resumo["verificacoes"] = int(parcial_dia["EH_VERIFICACAO"].sum())

    for contrato, df_contrato in parcial_dia.groupby("CONTRATO", dropna=False):
        contrato = str(contrato)
        resumo["por_contrato"][contrato] = {
            "notas": int(df_contrato["ORDEM_DE_SERVICO"].nunique()),
            "cortes": int(df_contrato["EH_CORTE"].sum()) + int(df_contrato["EH_VERIFICACAO"].sum()) if "EH_VERIFICACAO" in df_contrato.columns else int(df_contrato["EH_CORTE"].sum()),
            "religues": int(df_contrato["EH_RELIGUE"].sum()),
            "verificacoes": int(df_contrato["EH_VERIFICACAO"].sum()) if "EH_VERIFICACAO" in df_contrato.columns else 0,
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


def mostrar_status_atualizacao(notas, contrato_filtro_notas):
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





# ==============================
# VISÕES POR PERFIL
# ==============================

def filtrar_bases_para_supervisor_stc(bases):
    """Mantém somente Carro/STC Jundiai e Disjuntor Santa Cruz para o perfil Supervisor STC."""
    bases_filtradas = {k: v.copy() for k, v in bases.items()}
    permitidos = set(CONTRATOS_SUPERVISOR_STC)

    for chave in ["contratos", "dias", "carro", "carro_dias"]:
        df = bases_filtradas.get(chave, pd.DataFrame())
        if not df.empty and "CONTRATO" in df.columns:
            bases_filtradas[chave] = df[df["CONTRATO"].isin(permitidos)].copy()

    notas_df = bases_filtradas.get("notas", pd.DataFrame())
    if not notas_df.empty:
        try:
            parcial = preparar_parcial_do_dia(notas_df, incluir_recusas=True)
            if not parcial.empty and "CONTRATO" in parcial.columns and "ORDEM_DE_SERVICO" in parcial.columns:
                ordens_permitidas = (
                    parcial.loc[parcial["CONTRATO"].isin(permitidos), "ORDEM_DE_SERVICO"]
                    .astype(str)
                    .unique()
                    .tolist()
                )
                if "ORDEM_DE_SERVICO" in notas_df.columns:
                    notas_tmp = notas_df.copy()
                    notas_tmp["ORDEM_DE_SERVICO"] = notas_tmp["ORDEM_DE_SERVICO"].astype(str)
                    bases_filtradas["notas"] = notas_tmp[notas_tmp["ORDEM_DE_SERVICO"].isin(ordens_permitidas)].copy()
        except Exception:
            # Em caso de erro, evita liberar dados de Jundiaí por engano.
            bases_filtradas["notas"] = notas_df.iloc[0:0].copy()

    return bases_filtradas


def _contrato_operacional_sidebar(prefixo_key="stc"):
    opcoes = ["Todos"] + CONTRATOS_SUPERVISOR_STC
    key = f"contrato_{prefixo_key}"
    if key not in st.session_state or st.session_state[key] not in opcoes:
        st.session_state[key] = "Todos"

    st.sidebar.markdown("### Contratos permitidos")
    for contrato in opcoes:
        label = "📊 Todos STC/Santa Cruz" if contrato == "Todos" else f"🔹 {contrato}"
        if st.sidebar.button(label, use_container_width=True, key=f"btn_{prefixo_key}_{contrato}"):
            st.session_state[key] = contrato
            st.rerun()
    return st.session_state[key]


def _remover_colunas_financeiras(df):
    if df is None or df.empty:
        return df
    termos_bloqueados = ["FATURAMENTO", "TICKET", "VALOR", "MÍNIMO", "MAXIMO", "MÁXIMO", "MIN", "MAX"]
    colunas = [c for c in df.columns if not any(t in str(c).upper() for t in termos_bloqueados)]
    return df[colunas].copy()


def mostrar_painel_supervisor_leitura():
    st.title("📖 G.Z.U.S. — Supervisor Leitura")
    st.caption("Acesso restrito ao contrato Leitura. Nenhuma tela financeira ou de corte está disponível neste perfil.")
    st.sidebar.header("Supervisor Leitura")
    st.sidebar.info("Contrato Leitura")
    mostrar_base_leitura("Americana")
    st.markdown("---")
    mostrar_base_leitura("Piracicaba")


def mostrar_painel_supervisor_stc(bases):
    """Painel operacional STC/Santa Cruz sem qualquer métrica financeira."""
    st.title("🤖 G.Z.U.S. — Supervisor STC")
    st.caption("Acompanhamento operacional STC")

    notas_stc = bases.get("notas", pd.DataFrame())
    if notas_stc.empty:
        st.info("Nenhuma nota disponível para STC Jundiai ou Disjuntor Santa Cruz.")
        return

    st.sidebar.header("Filtros")
    if st.sidebar.button("🔄 Atualizar dados", use_container_width=True, key="stc_atualizar"):
        st.cache_data.clear()
        st.rerun()

    contrato_escolhido = _contrato_operacional_sidebar("supervisor_stc")

    meses_base = meses_disponiveis_da_base(notas_stc)
    meses_escolhidos = []
    if not meses_base.empty:
        opcoes_meses = meses_base["MES"].tolist()
        mes_padrao = opcoes_meses[0]
        meses_escolhidos = st.sidebar.multiselect(
            "Meses do resumo",
            opcoes_meses,
            default=[mes_padrao],
            key="stc_meses_resumo",
        ) or [mes_padrao]

    parcial_com_recusas = preparar_parcial_do_dia(notas_stc, incluir_recusas=True)
    if contrato_escolhido != "Todos" and not parcial_com_recusas.empty:
        parcial_com_recusas = parcial_com_recusas[parcial_com_recusas["CONTRATO"] == contrato_escolhido].copy()

    abas = st.tabs(["Resumo operacional", "Parcial do dia", "Ranking de recursos", "Comparativo mensal", "Dias da semana", "Notas"])

    with abas[0]:
        st.subheader("Resumo operacional")
        if parcial_com_recusas.empty:
            st.info("Sem dados operacionais para o filtro selecionado.")
        else:
            df_periodo = parcial_com_recusas.copy()
            if meses_escolhidos:
                df_periodo["MES"] = df_periodo["DATA_DT"].dt.strftime("%m/%Y")
                df_periodo = df_periodo[df_periodo["MES"].isin(meses_escolhidos)].copy()

            pagaveis = df_periodo[pd.to_numeric(df_periodo.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 0].copy()
            recusas = df_periodo[pd.to_numeric(df_periodo.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 1].copy()

            total_notas = int(pagaveis["ORDEM_DE_SERVICO"].nunique()) if not pagaveis.empty else 0
            total_cortes = int(pagaveis["EH_CORTE"].sum()) if not pagaveis.empty else 0
            total_religues = int(pagaveis["EH_RELIGUE"].sum()) if not pagaveis.empty else 0
            total_recusas = int(recusas["ORDEM_DE_SERVICO"].nunique()) if not recusas.empty else 0
            recursos_ativos = int(pagaveis["RECURSO"].nunique()) if not pagaveis.empty else 0

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Notas feitas", numero(total_notas))
            c2.metric("Cortes", numero(total_cortes))
            c3.metric("Religues", numero(total_religues))
            c4.metric("Recusas", numero(total_recusas))
            st.metric("Recursos ativos", numero(recursos_ativos))

            if not pagaveis.empty:
                resumo_contrato = (
                    pagaveis.groupby("CONTRATO", dropna=False)
                    .agg(TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"), CORTES=("EH_CORTE", "sum"), RELIGUES=("EH_RELIGUE", "sum"), VERIFICACOES=("EH_VERIFICACAO", "sum"), RECURSOS_ATIVOS=("RECURSO", "nunique"))
                    .reset_index()
                    .sort_values("TOTAL_NOTAS", ascending=False)
                )
                st.markdown("**Produção por contrato**")
                st.dataframe(resumo_contrato, use_container_width=True, hide_index=True)
                st.bar_chart(resumo_contrato, x="CONTRATO", y="TOTAL_NOTAS")

    with abas[1]:
        st.subheader("Parcial do dia")
        if parcial_com_recusas.empty:
            st.info("Sem dados para parcial do dia.")
        else:
            datas = parcial_com_recusas[["DATA", "DATA_DT"]].drop_duplicates().sort_values("DATA_DT", ascending=False)
            data_escolhida = st.selectbox("Escolha o dia", datas["DATA"].tolist(), index=0, key="stc_data_parcial")
            dia = parcial_com_recusas[parcial_com_recusas["DATA"] == data_escolhida].copy()
            pagaveis = dia[pd.to_numeric(dia.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 0].copy()
            recusas = dia[pd.to_numeric(dia.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 1].copy()

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            if "EH_VERIFICACAO" not in pagaveis.columns:
                pagaveis["EH_VERIFICACAO"] = 0
            c1.metric("Notas feitas", numero(pagaveis["ORDEM_DE_SERVICO"].nunique() if not pagaveis.empty else 0))
            c2.metric("Cortes", numero(int(pagaveis["EH_CORTE"].sum()) if not pagaveis.empty else 0))
            c3.metric("Religues", numero(int(pagaveis["EH_RELIGUE"].sum()) if not pagaveis.empty else 0))
            c4.metric("Verificações", numero(int(pagaveis["EH_VERIFICACAO"].sum()) if not pagaveis.empty else 0))
            c5.metric("Recusas", numero(recusas["ORDEM_DE_SERVICO"].nunique() if not recusas.empty else 0))
            c6.metric("Recursos", numero(pagaveis["RECURSO"].nunique() if not pagaveis.empty else 0))

            if not pagaveis.empty:
                resumo = (
                    pagaveis.groupby(["RECURSO", "CONTRATO"], dropna=False)
                    .agg(TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"), CORTES=("EH_CORTE", "sum"), RELIGUES=("EH_RELIGUE", "sum"), VERIFICACOES=("EH_VERIFICACAO", "sum"))
                    .reset_index()
                    .sort_values("TOTAL_NOTAS", ascending=False)
                )
                resumo.insert(0, "POSIÇÃO", range(1, len(resumo) + 1))
                st.dataframe(resumo, use_container_width=True, hide_index=True)

            if not recusas.empty:
                st.markdown("**Recusas do dia por tipo**")
                rec = (
                    recusas.groupby(["RECURSO", "CONTRATO", "RECUSA"], dropna=False)
                    .agg(QTD_RECUSAS=("ORDEM_DE_SERVICO", "nunique"))
                    .reset_index()
                    .sort_values(["QTD_RECUSAS", "RECURSO"], ascending=[False, True])
                )
                st.dataframe(rec, use_container_width=True, hide_index=True)

            recursos_sem_movimento = calcular_recursos_sem_movimento_no_dia(parcial_com_recusas, data_escolhida)
            render_alerta_recursos_sem_movimento(
                recursos_sem_movimento,
                contrato_unico=(contrato_escolhido != "Todos"),
            )

    with abas[2]:
        st.subheader("Ranking de recursos")
        base_exec = montar_base_executores(notas_stc)
        if contrato_escolhido != "Todos" and not base_exec.empty:
            base_exec = base_exec[base_exec["CONTRATO"] == contrato_escolhido].copy()

        if base_exec.empty:
            st.info("Sem base para ranking.")
        else:
            tipo_periodo = st.radio("Período", ["Mês", "Semana", "Dia"], horizontal=True, key="stc_tipo_rank")
            dias_op, semanas_op, meses_op = opcoes_periodo_ranking(base_exec)
            if tipo_periodo == "Dia":
                valor_periodo = st.selectbox("Dia", dias_op, key="stc_rank_dia") if dias_op else None
            elif tipo_periodo == "Semana":
                valor_periodo = st.selectbox("Semana", semanas_op, key="stc_rank_semana") if semanas_op else None
            else:
                valor_periodo = st.selectbox("Mês", meses_op, key="stc_rank_mes") if meses_op else None

            base_filtrada, ranking = ranking_recursos_cacheado(base_exec, "Todos", tipo_periodo, valor_periodo, "Notas")
            # Supervisor STC não visualiza Pagamento Express.

            colunas = ["POSIÇÃO", "RECURSO", "NOTAS", "CORTES", "RELIGUES", "RECUSAS", "DIAS_ATIVOS", "MÉDIA_NOTAS_DIA"]
            colunas = [c for c in colunas if c in ranking.columns]
            st.dataframe(preparar_tabela_ranking(ranking[colunas]), use_container_width=True, hide_index=True)
            if not ranking.empty:
                graf = ranking.head(20)[["RECURSO", "NOTAS"]].copy()
                st.bar_chart(graf, x="RECURSO", y="NOTAS")

            recusas_tipo = calcular_recusas_por_tipo(base_filtrada)
            if not recusas_tipo.empty:
                st.markdown("**Total por tipo de recusa**")
                total_tipo = recusas_tipo.groupby("RECUSA", dropna=False).agg(QTD_RECUSAS=("QTD_RECUSAS", "sum")).reset_index().sort_values("QTD_RECUSAS", ascending=False)
                st.dataframe(preparar_tabela_ranking(total_tipo), use_container_width=True, hide_index=True)
                st.markdown("**Detalhamento por equipe, contrato e tipo de recusa**")
                st.dataframe(preparar_tabela_ranking(recusas_tipo), use_container_width=True, hide_index=True)

    with abas[3]:
        st.subheader("Comparativo mensal operacional")
        if meses_base.empty:
            st.info("Sem meses disponíveis.")
        else:
            opcoes_meses = meses_base["MES"].tolist()
            mes_escolhido = st.selectbox("Escolha o mês", opcoes_meses, index=0, key="stc_comp_mes")
            periodo_escolhido = meses_base.loc[meses_base["MES"] == mes_escolhido, "PERIODO"].iloc[0]
            mes_anterior = (periodo_escolhido - 1).strftime("%m/%Y")
            def resumo_operacional_mes(mes_ref):
                df_mes = parcial_com_recusas.copy()
                if not df_mes.empty:
                    df_mes["MES"] = df_mes["DATA_DT"].dt.strftime("%m/%Y")
                    df_mes = df_mes[df_mes["MES"] == mes_ref].copy()
                pag = df_mes[pd.to_numeric(df_mes.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 0].copy() if not df_mes.empty else pd.DataFrame()
                return {
                    "TOTAL_NOTAS": int(pag["ORDEM_DE_SERVICO"].nunique()) if not pag.empty else 0,
                    "CORTES": int(pag["EH_CORTE"].sum()) + int(pag["EH_VERIFICACAO"].sum()) if not pag.empty else 0,
                    "RELIGUES": int(pag["EH_RELIGUE"].sum()) if not pag.empty else 0,
                    "VERIFICACOES": int(pag["EH_VERIFICACAO"].sum()) if not pag.empty else 0,
                    "VERIFICACOES": int(pag["EH_VERIFICACAO"].sum()) if not pag.empty and "EH_VERIFICACAO" in pag.columns else 0,
                }

            atual = resumo_operacional_mes(mes_escolhido)
            anterior = resumo_operacional_mes(mes_anterior)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Notas", numero(atual["TOTAL_NOTAS"]), variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"]))
            c2.metric("Cortes", numero(atual["CORTES"]), variacao_percentual(atual["CORTES"], anterior["CORTES"]))
            c3.metric("Religues", numero(atual["RELIGUES"]), variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"]))
            c4.metric("Verificações", numero(atual.get("VERIFICACOES", 0)), variacao_percentual(atual.get("VERIFICACOES", 0), anterior.get("VERIFICACOES", 0)))
            tabela = pd.DataFrame([
                {"Indicador": "Notas", mes_escolhido: numero(atual["TOTAL_NOTAS"]), mes_anterior: numero(anterior["TOTAL_NOTAS"]), "Variação": variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"])},
                {"Indicador": "Cortes", mes_escolhido: numero(atual["CORTES"]), mes_anterior: numero(anterior["CORTES"]), "Variação": variacao_percentual(atual["CORTES"], anterior["CORTES"])},
                {"Indicador": "Religues", mes_escolhido: numero(atual["RELIGUES"]), mes_anterior: numero(anterior["RELIGUES"]), "Variação": variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"])},
                {"Indicador": "Verificações", mes_escolhido: numero(atual.get("VERIFICACOES", 0)), mes_anterior: numero(anterior.get("VERIFICACOES", 0)), "Variação": variacao_percentual(atual.get("VERIFICACOES", 0), anterior.get("VERIFICACOES", 0))},
            ])
            st.dataframe(tabela, use_container_width=True, hide_index=True)

    with abas[4]:
        st.subheader("Produção por dia da semana")
        df_dias = parcial_com_recusas.copy()
        if df_dias.empty:
            st.info("Sem dados.")
        else:
            df_dias = df_dias[pd.to_numeric(df_dias.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 0].copy()
            tabela = (
                df_dias.groupby(["CONTRATO", "SEMANA_INICIO", "DIA_SEMANA"], dropna=False)
                .agg(NOTAS=("ORDEM_DE_SERVICO", "nunique"))
                .reset_index()
                .pivot_table(index=["CONTRATO", "SEMANA_INICIO"], columns="DIA_SEMANA", values="NOTAS", aggfunc="sum", fill_value=0)
                .reset_index()
            )
            colunas_dias = [c for c in ORDEM_DIAS if c in tabela.columns]
            tabela["Total semana"] = tabela[colunas_dias].sum(axis=1) if colunas_dias else 0
            st.dataframe(tabela[["CONTRATO", "SEMANA_INICIO"] + colunas_dias + ["Total semana"]], use_container_width=True, hide_index=True)
            por_dia = df_dias.groupby("DIA_SEMANA", as_index=False).agg(NOTAS=("ORDEM_DE_SERVICO", "nunique"))
            por_dia["ordem"] = por_dia["DIA_SEMANA"].map({d: i for i, d in enumerate(ORDEM_DIAS)})
            por_dia = por_dia.sort_values("ordem")
            st.bar_chart(por_dia, x="DIA_SEMANA", y="NOTAS")

    with abas[5]:
        st.subheader("Consulta de notas")
        df_notas = notas_stc.copy()
        termo = st.text_input("Buscar por OS, recurso ou recusa", key="stc_busca_notas")
        if termo:
            termo_norm = str(termo).upper().strip()
            mask = pd.Series(False, index=df_notas.index)
            for col in ["ORDEM_DE_SERVICO", "RECURSO", "RECUSA", "GRUPO_NOTA"]:
                if col in df_notas.columns:
                    mask = mask | df_notas[col].fillna("").astype(str).str.upper().str.contains(termo_norm, na=False)
            df_notas = df_notas[mask].copy()
        df_notas = _remover_colunas_financeiras(df_notas)
        st.dataframe(df_notas.head(1000), use_container_width=True, hide_index=True)

# ==============================
# G.Z.U.S. — CHATBOT LOCAL
# Gestão Inteligente de Serviços
# ==============================

NOME_ASSISTENTE = "G.Z.U.S."
SUBTITULO_ASSISTENTE = "Gestão Inteligente de Serviços"


def _normalizar_chat(texto):
    """Normaliza texto para comparação simples, sem depender de IA/API externa."""
    import unicodedata
    import re

    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = texto.upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def _nome_mes_chat(mes):
    nomes = {
        "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
        "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
        "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro",
    }
    try:
        mm, aa = str(mes).split("/")
        return f"{nomes.get(mm, mm)}/{aa}"
    except Exception:
        return str(mes or "período selecionado")


def _meses_chat_disponiveis(base):
    if base.empty or "DATA_DT" not in base.columns:
        return []

    meses = (
        base[["DATA_DT"]]
        .dropna()
        .assign(
            MES=lambda d: d["DATA_DT"].dt.strftime("%m/%Y"),
            PERIODO=lambda d: d["DATA_DT"].dt.to_period("M"),
        )[["MES", "PERIODO"]]
        .drop_duplicates()
        .sort_values("PERIODO", ascending=False)["MES"]
        .tolist()
    )
    return meses


def _extrair_mes_chat(pergunta, meses_disponiveis, contexto=None):
    import re

    contexto = contexto or {}
    pergunta_norm = _normalizar_chat(pergunta)

    achado = re.search(r"\b(0?[1-9]|1[0-2])[/\-](20\d{2}|\d{2})\b", pergunta_norm)
    if achado:
        mes = int(achado.group(1))
        ano = int(achado.group(2))
        if ano < 100:
            ano += 2000
        candidato = f"{mes:02d}/{ano}"
        if candidato in meses_disponiveis:
            return candidato

    # "mês 03", "mes 3", "no mês 03", "e no 03?"
    achado_mes_num = re.search(r"(?:\bMES\b|\bEM\b|\bNO\b|\bNO MES\b|\bMES DE\b|\bE NO\b)\s*(0?[1-9]|1[0-2])\b", pergunta_norm)
    if achado_mes_num:
        numero_mes = f"{int(achado_mes_num.group(1)):02d}"
        ano_ctx = str(contexto.get("mes", ""))[-4:] if contexto.get("mes") else ""
        if ano_ctx:
            candidato = f"{numero_mes}/{ano_ctx}"
            if candidato in meses_disponiveis:
                return candidato
        for mes_disp in meses_disponiveis:
            if mes_disp.startswith(numero_mes + "/"):
                return mes_disp

    mapa_meses = {
        "JANEIRO": "01", "FEVEREIRO": "02", "MARCO": "03", "MARÇO": "03", "ABRIL": "04",
        "MAIO": "05", "JUNHO": "06", "JULHO": "07", "AGOSTO": "08", "SETEMBRO": "09",
        "OUTUBRO": "10", "NOVEMBRO": "11", "DEZEMBRO": "12",
    }
    for nome, numero_mes in mapa_meses.items():
        if nome in pergunta_norm:
            achado_ano = re.search(r"\b(20\d{2})\b", pergunta_norm)
            if achado_ano:
                candidato = f"{numero_mes}/{achado_ano.group(1)}"
                if candidato in meses_disponiveis:
                    return candidato

            ano_ctx = str(contexto.get("mes", ""))[-4:] if contexto.get("mes") else ""
            if ano_ctx:
                candidato = f"{numero_mes}/{ano_ctx}"
                if candidato in meses_disponiveis:
                    return candidato

            for mes_disp in meses_disponiveis:
                if mes_disp.startswith(numero_mes + "/"):
                    return mes_disp

    if any(t in pergunta_norm for t in ["ESSE MES", "MES ATUAL", "ATUAL", "MENSAL", "NO MES", "MES"]):
        return meses_disponiveis[0] if meses_disponiveis else None

    return None


def _identificar_contrato_chat(pergunta, contratos_disponiveis):
    """Identifica contrato com regras rígidas.

    - STC / carro / JUN58 = STC Jundiai.
    - Santa Cruz = Disjuntor Santa Cruz.
    - Jundiaí sozinho respeita os contratos disponíveis no perfil.
    """
    import re
    pergunta_norm = _normalizar_chat(pergunta)
    disponiveis = set(contratos_disponiveis or [])

    if "TODOS" in pergunta_norm or "GERAL" in pergunta_norm:
        return None

    def tem_termo(termo):
        termo_norm = _normalizar_chat(termo)
        return re.search(rf"(?<![A-Z0-9]){re.escape(termo_norm)}(?![A-Z0-9])", pergunta_norm) is not None

    regras = [
        (["DISJUNTOR SANTA CRUZ", "SANTA CRUZ", "ST CRUZ", "SANTACRUZ"], "Disjuntor Santa Cruz"),
        (["STC JUNDIAI", "STC JUNDIA", "CONTRATO CARRO", "CARRO", "JUN58"], "STC Jundiai"),
        (["STC"], "STC Jundiai"),
        (["DISJUNTOR JUNDIAI", "DISJUNTOR JUNDIA"], "Disjuntor Jundiaí"),
    ]

    for termos, contrato in regras:
        if contrato in disponiveis and any(tem_termo(t) for t in termos):
            return contrato

    if any(tem_termo(t) for t in ["JUNDIAI", "JUNDIA"]):
        if "Disjuntor Jundiaí" in disponiveis:
            return "Disjuntor Jundiaí"
        if "STC Jundiai" in disponiveis:
            return "STC Jundiai"

    for contrato in contratos_disponiveis:
        if _normalizar_chat(contrato) in pergunta_norm:
            return contrato
    return None


def _identificar_recurso_chat(pergunta, recursos_disponiveis):
    import re

    pergunta_norm = _normalizar_chat(pergunta)
    recursos_norm = {str(r).upper().strip(): r for r in recursos_disponiveis if str(r).strip()}

    # Match literal: JUN5981-EMP, ITN8905-EMP etc.
    for recurso_norm, recurso_original in recursos_norm.items():
        if recurso_norm and recurso_norm in pergunta_norm:
            return recurso_original

    # Match sem sufixo/prefixo: "5981 fez quanto?", "equipe 5802", "ITN 8905".
    codigos = re.findall(r"\b\d{4}\b", pergunta_norm)
    for codigo in codigos:
        candidatos = [r for r in recursos_disponiveis if codigo_numerico_recurso(r) == codigo]
        if len(candidatos) == 1:
            return candidatos[0]

        if candidatos:
            # Prioriza JUN para códigos 55/58/59; isso ajuda frases curtas como "5981".
            candidatos_jun = [r for r in candidatos if str(r).upper().startswith("JUN")]
            if candidatos_jun:
                return candidatos_jun[0]
            return candidatos[0]

    return None


def _pergunta_eh_complemento_chat(pergunta_norm):
    """Detecta continuações curtas como 'e no mês 03?' ou 'mas conta express'."""
    termos = [
        "E NO", "E EM", "MAS", "TAMBEM", "TAMBÉM", "AGORA", "COM EXPRESS",
        "CONTA EXPRESS", "CONTAR EXPRESS", "INCLUI EXPRESS", "INCLUIR EXPRESS",
        "NO MES", "MES 0", "MES 1", "E O", "E A", "E DE"
    ]
    curta = len(pergunta_norm.split()) <= 8
    return curta or any(t in pergunta_norm for t in termos)


def _pergunta_ultimos_meses_chat(pergunta_norm):
    return any(t in pergunta_norm for t in [
        "ULTIMOS MESES", "ÚLTIMOS MESES", "ULTIMAS MESES", "NOS ULTIMOS",
        "ULTIMOS 3", "ULTIMOS 4", "ULTIMOS 5", "ULTIMOS 6",
        "POR MES", "MES A MES", "MENSAL DOS", "HISTORICO", "HISTÓRICO"
    ])


def _tipo_consulta_chat(pergunta_norm):
    """Classifica a intenção principal da pergunta do assistente local.

    v2.1: reforça perguntas do tipo "melhor equipe", "melhor contrato",
    "maior produção" e "somando todos os contratos" para não caírem no
    resumo geral por engano.
    """
    if any(t in pergunta_norm for t in ["COMPARE", "COMPARA", "VS", "VERSUS", "MELHOR QUE", "PIOR QUE", "CRESCEU", "CAIU"]):
        return "comparacao"

    termos_ranking = [
        "QUEM MAIS", "QUEM FOI", "TOP", "RANKING", "LIDER", "LÍDER",
        "MAIOR PRODU", "MAIOR PRODUCAO", "MAIOR PRODUÇÃO", "MAIS FEZ",
        "MAIS NOTAS", "CAMPEAO", "CAMPEÃO", "MELHOR EQUIPE",
        "MELHORES EQUIPES", "MELHOR RECURSO", "MELHORES RECURSOS",
        "MELHOR CONTRATO", "MELHORES CONTRATOS", "O MELHOR", "A MELHOR",
        "QUAL FOI A MELHOR", "QUAL FOI O MELHOR", "QUAL A MELHOR",
        "QUAL O MELHOR", "QUAIS CONTRATOS", "QUAIS EQUIPES",
        "SOMANDO TODOS OS CONTRATOS", "TODOS OS CONTRATOS"
    ]
    if any(t in pergunta_norm for t in termos_ranking):
        return "ranking"

    if any(t in pergunta_norm for t in ["RECUSA", "RECUSAS", "CONTA PAGA", "SEM ACESSO", "CASA FECHADA", "CLIENTE"]):
        return "recusas"
    if "EXPRESS" in pergunta_norm:
        return "express"
    if any(t in pergunta_norm for t in ["FATUR", "RECEITA", "VALOR", "R$", "DINHEIRO"]):
        return "faturamento"
    if any(t in pergunta_norm for t in ["COMO FOI", "RESUMO", "QUANTO", "QUANTAS", "QUANTOS", "FEZ", "PRODU", "NOTAS", "RESULTADO"]):
        return "resumo"
    return "resumo"


def _top_n_chat(pergunta_norm, padrao=5):
    import re
    m = re.search(r"\bTOP\s*(\d{1,2})\b", pergunta_norm)
    if not m:
        m = re.search(r"\b(\d{1,2})\s*(?:PRIMEIR|MELHOR|MAIOR|EQUIPE|CONTRATO|RECURSO)", pergunta_norm)
    if m:
        try:
            return max(1, min(int(m.group(1)), 15))
        except Exception:
            return padrao

    # Perguntas no singular devem devolver o campeão, não Top 5.
    if any(t in pergunta_norm for t in [
        "QUEM", "CAMPEAO", "CAMPEÃO", "LIDER", "LÍDER",
        "O MELHOR", "A MELHOR", "QUAL FOI O MELHOR", "QUAL FOI A MELHOR",
        "QUAL O MELHOR", "QUAL A MELHOR", "MELHOR EQUIPE",
        "MELHOR RECURSO", "MELHOR CONTRATO"
    ]):
        if not any(t in pergunta_norm for t in ["TOP", "MELHORES", "QUAIS", "LISTA", "RANKING"]):
            return 1
    return padrao


def _ranking_metrica_chat(pergunta_norm):
    if any(t in pergunta_norm for t in ["RECUSA", "RECUSAS", "CONTA PAGA", "CASA FECHADA", "SEM ACESSO"]):
        return "recusas"
    if any(t in pergunta_norm for t in ["MEDIA", "MÉDIA", "EFICIEN", "POR DIA"]):
        return "media"
    if any(t in pergunta_norm for t in ["FATUR", "VALOR", "RECEITA", "R$"]):
        return "faturamento"
    return "notas"


def _ranking_dimensao_chat(pergunta_norm):
    """Decide se o ranking pedido é por contrato ou por equipe/recurso.

    v2.1: palavras de equipe têm prioridade sobre "contratos" quando a
    pergunta diz algo como "melhor equipe somando todos os contratos".
    """
    termos_recurso = [
        "EQUIPE", "EQUIPES", "RECURSO", "RECURSOS", "CARRO", "CARROS",
        "FUNCIONARIO", "FUNCIONÁRI", "AGENTE", "AGENTES",
        "ELETRICISTA", "EXECUTOR", "EXECUTORES"
    ]
    if any(t in pergunta_norm for t in termos_recurso):
        return "recurso"

    if any(t in pergunta_norm for t in [
        "CONTRATO", "CONTRATOS", "POR CONTRATO", "ENTRE CONTRATOS",
        "QUAL CONTRATO", "QUAIS CONTRATOS"
    ]):
        return "contrato"
    return "recurso"


def _montar_ranking_contratos_chat(df, metrica="notas", pode_ver_financeiro=True):
    """Ranking agregado por CONTRATO, não por RECURSO.

    Mantém separadas notas pagáveis e recusas, igual ao ranking por equipe.
    """
    if df.empty or "CONTRATO" not in df.columns:
        return pd.DataFrame()

    eh_recusa = pd.to_numeric(df.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
    base = df.copy()
    base["_EH_RECUSA"] = eh_recusa.values
    pagaveis = base[base["_EH_RECUSA"] == 0].copy()
    recusas = base[base["_EH_RECUSA"] == 1].copy()

    contratos = sorted(base.get("CONTRATO", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
    if not contratos:
        return pd.DataFrame()

    if pagaveis.empty:
        ranking = pd.DataFrame({"CONTRATO": contratos})
        ranking["NOTAS"] = 0
        ranking["CORTES"] = 0
        ranking["RELIGUES"] = 0
        ranking["VERIFICACOES"] = 0
        ranking["DIAS_ATIVOS"] = 0
        ranking["RECURSOS_ATIVOS"] = 0
        ranking["FATURAMENTO"] = 0.0
    else:
        ranking = (
            pagaveis.groupby("CONTRATO", dropna=False)
            .agg(
                NOTAS=("ORDEM_DE_SERVICO", "nunique"),
                CORTES=("EH_CORTE", "sum"),
                RELIGUES=("EH_RELIGUE", "sum"),
                VERIFICACOES=("EH_VERIFICACAO", "sum"),
                DIAS_ATIVOS=("DATA", "nunique"),
                RECURSOS_ATIVOS=("RECURSO", "nunique"),
                FATURAMENTO=("FATURAMENTO", "sum"),
            )
            .reset_index()
        )

    if recusas.empty:
        rec = pd.DataFrame({"CONTRATO": contratos, "RECUSAS": 0})
    else:
        rec = recusas.groupby("CONTRATO", dropna=False).agg(RECUSAS=("ORDEM_DE_SERVICO", "nunique")).reset_index()

    ranking = ranking.merge(rec, on="CONTRATO", how="outer").fillna(0)
    for col in ["NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "DIAS_ATIVOS", "RECURSOS_ATIVOS", "RECUSAS"]:
        if col not in ranking.columns:
            ranking[col] = 0
        ranking[col] = pd.to_numeric(ranking[col], errors="coerce").fillna(0).astype(int)
    if "FATURAMENTO" not in ranking.columns:
        ranking["FATURAMENTO"] = 0.0
    ranking["FATURAMENTO"] = pd.to_numeric(ranking["FATURAMENTO"], errors="coerce").fillna(0.0)
    ranking["MEDIA_DIA"] = ranking.apply(lambda r: (r["NOTAS"] / r["DIAS_ATIVOS"]) if r["DIAS_ATIVOS"] else 0, axis=1)

    if metrica == "recusas":
        ordem = ["RECUSAS", "NOTAS"]
    elif metrica == "media":
        ordem = ["MEDIA_DIA", "NOTAS"]
    elif metrica == "faturamento" and pode_ver_financeiro:
        ordem = ["FATURAMENTO", "NOTAS"]
    else:
        ordem = ["NOTAS", "RECUSAS"]

    ranking = ranking.sort_values(ordem, ascending=[False] * len(ordem)).reset_index(drop=True)
    ranking.insert(0, "POSICAO", range(1, len(ranking) + 1))
    return ranking


def _montar_ranking_chat(df, metrica="notas", pode_ver_financeiro=True):
    if df.empty:
        return pd.DataFrame()

    eh_recusa = pd.to_numeric(df.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
    base = df.copy()
    base["_EH_RECUSA"] = eh_recusa.values
    pagaveis = base[base["_EH_RECUSA"] == 0].copy()
    recusas = base[base["_EH_RECUSA"] == 1].copy()

    recursos = sorted(base.get("RECURSO", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
    if not recursos:
        return pd.DataFrame()

    if pagaveis.empty:
        ranking = pd.DataFrame({"RECURSO": recursos})
        ranking["NOTAS"] = 0
        ranking["CORTES"] = 0
        ranking["RELIGUES"] = 0
        ranking["VERIFICACOES"] = 0
        ranking["DIAS_ATIVOS"] = 0
        ranking["FATURAMENTO"] = 0.0
    else:
        ranking = (
            pagaveis.groupby("RECURSO", dropna=False)
            .agg(
                NOTAS=("ORDEM_DE_SERVICO", "nunique"),
                CORTES=("EH_CORTE", "sum"),
                RELIGUES=("EH_RELIGUE", "sum"),
                VERIFICACOES=("EH_VERIFICACAO", "sum"),
                DIAS_ATIVOS=("DATA", "nunique"),
                FATURAMENTO=("FATURAMENTO", "sum"),
            )
            .reset_index()
        )

    if recusas.empty:
        rec = pd.DataFrame({"RECURSO": recursos, "RECUSAS": 0})
    else:
        rec = recusas.groupby("RECURSO", dropna=False).agg(RECUSAS=("ORDEM_DE_SERVICO", "nunique")).reset_index()

    ranking = ranking.merge(rec, on="RECURSO", how="outer").fillna(0)
    for col in ["NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "DIAS_ATIVOS", "RECUSAS"]:
        if col not in ranking.columns:
            ranking[col] = 0
        ranking[col] = pd.to_numeric(ranking[col], errors="coerce").fillna(0).astype(int)
    if "FATURAMENTO" not in ranking.columns:
        ranking["FATURAMENTO"] = 0.0
    ranking["FATURAMENTO"] = pd.to_numeric(ranking["FATURAMENTO"], errors="coerce").fillna(0.0)
    ranking["MEDIA_DIA"] = ranking.apply(lambda r: (r["NOTAS"] / r["DIAS_ATIVOS"]) if r["DIAS_ATIVOS"] else 0, axis=1)

    if metrica == "recusas":
        ordem = ["RECUSAS", "NOTAS"]
    elif metrica == "media":
        ordem = ["MEDIA_DIA", "NOTAS"]
    elif metrica == "faturamento" and pode_ver_financeiro:
        ordem = ["FATURAMENTO", "NOTAS"]
    else:
        ordem = ["NOTAS", "RECUSAS"]

    ranking = ranking.sort_values(ordem, ascending=[False] * len(ordem)).reset_index(drop=True)
    ranking.insert(0, "POSICAO", range(1, len(ranking) + 1))
    return ranking


def _principal_recusa_recurso_chat(df, recurso=None):
    tmp = df.copy()
    if recurso and "RECURSO" in tmp.columns:
        tmp = tmp[tmp["RECURSO"] == recurso].copy()
    resumo = _resumo_recusas_tipo(tmp)
    if resumo.empty:
        return ""
    row = resumo.iloc[0]
    return f"Principal motivo de recusa: **{row['RECUSA']}** ({numero(int(row['QTD']))})."


def _resumo_numerico_chat(df):
    if df.empty:
        return {
            "notas": 0,
            "cortes": 0,
            "religues": 0,
            "recusas": 0,
            "faturamento": 0.0,
            "recursos": 0,
            "dias_ativos": 0,
        }

    eh_recusa = pd.to_numeric(df.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
    pagaveis = df[eh_recusa == 0].copy()
    recusas = df[eh_recusa == 1].copy()

    return {
        "notas": int(pagaveis["ORDEM_DE_SERVICO"].nunique()) if "ORDEM_DE_SERVICO" in pagaveis.columns else 0,
        "cortes": int(pd.to_numeric(pagaveis.get("EH_CORTE", 0), errors="coerce").fillna(0).sum()),
        "religues": int(pd.to_numeric(pagaveis.get("EH_RELIGUE", 0), errors="coerce").fillna(0).sum()),
        "recusas": int(recusas["ORDEM_DE_SERVICO"].nunique()) if "ORDEM_DE_SERVICO" in recusas.columns else 0,
        "faturamento": float(pd.to_numeric(pagaveis.get("FATURAMENTO", 0), errors="coerce").fillna(0).sum()),
        "recursos": int(pagaveis["RECURSO"].nunique()) if "RECURSO" in pagaveis.columns else 0,
        "dias_ativos": int(pagaveis["DATA"].nunique()) if "DATA" in pagaveis.columns else 0,
    }


def _express_chat_periodo(notas, mes, contrato=None, recurso=None):
    if not mes:
        return {"express": 0, "faturamento_express": 0.0, "tem_base": False}

    express_resumo, _, _, caminho = calcular_express_mensal(notas, mes)
    if express_resumo.empty:
        return {"express": 0, "faturamento_express": 0.0, "tem_base": bool(caminho)}

    df = express_resumo.copy()
    if contrato:
        df = df[df["CONTRATO"] == contrato].copy()
    if recurso:
        df = df[df["RECURSO"] == recurso].copy()

    return {
        "express": int(pd.to_numeric(df.get("EXPRESS", 0), errors="coerce").fillna(0).sum()) if not df.empty else 0,
        "faturamento_express": float(pd.to_numeric(df.get("FATURAMENTO_EXPRESS", 0), errors="coerce").fillna(0).sum()) if not df.empty else 0.0,
        "tem_base": bool(caminho),
    }


def _resumo_recusas_tipo(df):
    if df.empty or "EH_RECUSA" not in df.columns:
        return pd.DataFrame(columns=["RECUSA", "QTD"])

    eh_recusa = pd.to_numeric(df.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
    recusas = df[eh_recusa == 1].copy()
    if recusas.empty:
        return pd.DataFrame(columns=["RECUSA", "QTD"])

    recusas["RECUSA"] = recusas.get("RECUSA", "").fillna("").astype(str).str.strip()
    recusas.loc[recusas["RECUSA"] == "", "RECUSA"] = "Não informado"

    return (
        recusas.groupby("RECUSA", dropna=False)
        .agg(QTD=("ORDEM_DE_SERVICO", "nunique"))
        .reset_index()
        .sort_values("QTD", ascending=False)
    )


def _destaques_executivos_chat(resumo, express_qtd, resumo_recusas):
    notas = int(resumo.get("notas", 0) or 0)
    recusas = int(resumo.get("recusas", 0) or 0)
    total_operacional = notas + express_qtd + recusas
    taxa_recusa = (recusas / total_operacional * 100) if total_operacional else 0

    destaques = []

    if taxa_recusa >= 20:
        destaques.append(f"⚠️ Destaque: taxa de recusa alta ({taxa_recusa:.1f}%).".replace(".", ","))
    elif taxa_recusa >= 10:
        destaques.append(f"⚠️ Atenção: recusas representam {taxa_recusa:.1f}% do volume operacional.".replace(".", ","))

    if not resumo_recusas.empty:
        top = resumo_recusas.iloc[0]
        qtd_top = int(top["QTD"])
        if qtd_top > 0:
            destaques.append(f"Principal motivo de recusa: **{top['RECUSA']}** ({numero(qtd_top)}).")

    if express_qtd > 0:
        destaques.append(f"Pagamento express identificado: **{numero(express_qtd)}** atendimento(s).")

    if not destaques:
        destaques.append("Sem alerta crítico no período consultado.")

    return destaques


def _escopo_texto_chat(contrato=None, recurso=None):
    if recurso:
        return f"equipe **{recurso}**"
    if contrato:
        return f"contrato **{contrato}**"
    return "visão geral"


def _filtrar_base_chat(base, mes=None, contrato=None, recurso=None):
    df = base.copy()
    if mes:
        df = df[df["DATA_DT"].dt.strftime("%m/%Y") == mes].copy()
    if contrato:
        df = df[df["CONTRATO"] == contrato].copy()
    if recurso:
        df = df[df["RECURSO"] == recurso].copy()
    return df



def responder_chatbot_leitura(pergunta):
    """Chat simples para o perfil Leitura, restrito às parciais de leitura."""
    pergunta_norm = _normalizar_chat(pergunta)
    bases = []
    if "AMERICANA" in pergunta_norm:
        bases = ["Americana"]
    elif "PIRACICABA" in pergunta_norm:
        bases = ["Piracicaba"]
    else:
        bases = ["Americana", "Piracicaba"]

    linhas = [f"📖 **{NOME_ASSISTENTE} — Contrato Leitura**", ""]
    encontrou = False
    for base_nome in bases:
        caminho = caminho_leitura(base_nome)
        if not caminho:
            linhas.append(f"• **{base_nome}:** parcial não encontrada.")
            continue
        try:
            df = ler_parcial_leitura(str(caminho))
        except Exception as e:
            linhas.append(f"• **{base_nome}:** não consegui ler a parcial ({e}).")
            continue
        if df.empty:
            linhas.append(f"• **{base_nome}:** parcial vazia.")
            continue
        encontrou = True
        total_instala = int(df["T. INSTALA"].sum())
        total_visitada = int(df["T. VISITADA"].sum())
        total_faltam = int(df["FALTAM"].sum())
        percentual = (total_visitada / total_instala * 100) if total_instala else 0
        linhas.append(f"### {base_nome}")
        linhas.append(f"• **Leituras totais:** {numero(total_instala)}")
        linhas.append(f"• **Leituras feitas:** {numero(total_visitada)}")
        linhas.append(f"• **Leituras faltantes:** {numero(total_faltam)}")
        linhas.append(f"• **Executado:** {percentual:.1f}%".replace(".", ","))
        atrasados = df.sort_values(["FALTAM", "% EXECUTADO"], ascending=[False, True]).head(3)
        if not atrasados.empty:
            linhas.append("• **Maiores pendências:** " + "; ".join(
                f"{r['AGENTE COMERCIAL']} ({numero(r['FALTAM'])})" for _, r in atrasados.iterrows()
            ))
        linhas.append("")

    if not encontrou:
        linhas.append("Não encontrei parciais de leitura disponíveis para consultar.")
    return "\n".join(linhas)

def responder_chatbot_painel(pergunta, notas, pode_ver_financeiro=True, pode_ver_express=True, modo_leitura=False):
    """Responde perguntas operacionais usando os CSVs já carregados no painel.

    Versão avançada local: entende contrato, equipe, mês, ranking, top N,
    perguntas curtas e continuações de conversa sem depender de API externa.
    """
    if modo_leitura:
        return responder_chatbot_leitura(pergunta)

    if notas.empty:
        return "Ainda não encontrei dados carregados para consultar."

    base = preparar_parcial_do_dia(notas, incluir_recusas=True)
    if base.empty:
        return "Ainda não encontrei dados suficientes para responder. Confira se o `notas_dashboard.csv` foi carregado."

    pergunta_norm = _normalizar_chat(pergunta)
    meses = _meses_chat_disponiveis(base)
    contratos_disp = sorted(base["CONTRATO"].dropna().unique().tolist()) if "CONTRATO" in base.columns else []
    recursos_disp = sorted(base["RECURSO"].dropna().unique().tolist()) if "RECURSO" in base.columns else []

    contexto_anterior = st.session_state.get("chatbot_painel_contexto", {}) if hasattr(st, "session_state") else {}
    complemento = _pergunta_eh_complemento_chat(pergunta_norm)

    mes = _extrair_mes_chat(pergunta, meses, contexto_anterior)
    contrato = _identificar_contrato_chat(pergunta, contratos_disp)
    recurso = _identificar_recurso_chat(pergunta, recursos_disp)

    if complemento:
        if not recurso:
            recurso = contexto_anterior.get("recurso")
        if not contrato and not recurso:
            contrato = contexto_anterior.get("contrato")
        if not mes:
            mes = contexto_anterior.get("mes")

    if not mes and meses:
        mes = meses[0]

    if recurso and not contrato:
        contratos_recurso = base.loc[base["RECURSO"] == recurso, "CONTRATO"].dropna().unique().tolist()
        if len(contratos_recurso) == 1:
            contrato = contratos_recurso[0]

    # ==============================
    # MOTOR LOCAL DE INTENÇÃO / CONTEXTO
    # ==============================
    # O chatbot continua sem API externa, mas passa a separar:
    # 1) o que foi dito agora;
    # 2) o que deve ser herdado da pergunta anterior;
    # 3) o que será consultado na base.
    # Isso corrige continuações como "E em março?", mantendo ranking/contrato/métrica.
    tipo_detectado = _tipo_consulta_chat(pergunta_norm)
    metrica_detectada = _ranking_metrica_chat(pergunta_norm)
    dimensao_detectada = _ranking_dimensao_chat(pergunta_norm)
    top_n_detectado = _top_n_chat(pergunta_norm, padrao=5)

    tem_intencao_forte = any(t in pergunta_norm for t in [
        "QUEM MAIS", "QUEM FOI", "TOP", "RANKING", "LIDER", "LÍDER",
        "MAIS FEZ", "MAIS NOTAS", "CAMPEAO", "CAMPEÃO", "RECUSA",
        "RECUSAS", "FATUR", "EXPRESS", "COMO FOI", "RESUMO", "QUANTO",
        "QUANTAS", "QUANTOS", "PRODU", "NOTAS", "COMPARE", "COMPARA",
        "VS", "VERSUS", "CRESCEU", "CAIU", "MELHOR", "MELHORES",
        "EQUIPE", "EQUIPES", "RECURSO", "RECURSOS", "TODOS OS CONTRATOS",
        "SOMANDO"
    ])

    tipo = tipo_detectado
    metrica_ranking = metrica_detectada
    top_n = top_n_detectado
    dimensao_ranking = dimensao_detectada

    # Escopo geral explícito: evita herdar contrato/equipe anterior quando o usuário
    # pede "todas as equipes" ou "somando todos os contratos".
    escopo_geral_explicito = any(t in pergunta_norm for t in [
        "TODOS OS CONTRATOS", "TODOS CONTRATOS", "SOMANDO TODOS",
        "SOMANDO TUDO", "GERAL", "NO GERAL", "TODAS AS EQUIPES",
        "TODAS EQUIPES", "TODOS OS RECURSOS", "TODOS RECURSOS"
    ])
    if escopo_geral_explicito and not recurso:
        contrato = None

    if complemento and contexto_anterior:
        # Perguntas curtas de sequência normalmente só trocam mês/contrato/equipe.
        # Ex.: "E em março?" mantém "ranking da maior produção no STC".
        if not tem_intencao_forte or pergunta_norm.startswith(("E ", "E NO", "E EM", "E A", "E O")):
            tipo = contexto_anterior.get("tipo", tipo) or tipo
            metrica_ranking = contexto_anterior.get("metrica_ranking", metrica_ranking) or metrica_ranking
            top_n = int(contexto_anterior.get("top_n", top_n) or top_n)
            dimensao_ranking = contexto_anterior.get("dimensao_ranking", dimensao_ranking) or dimensao_ranking

        # Se a continuação só citou mês, mantém o alvo anterior, exceto quando
        # o usuário pedir explicitamente visão geral/todos os contratos.
        if not escopo_geral_explicito and not recurso and not contrato:
            recurso = contexto_anterior.get("recurso")
            contrato = contexto_anterior.get("contrato")

        # Se a continuação citou uma equipe, a equipe tem prioridade sobre contrato.
        if recurso:
            contrato = None
        if escopo_geral_explicito and not recurso:
            contrato = None

    quer_somar_express = "EXPRESS" in pergunta_norm and any(t in pergunta_norm for t in ["CONTAR", "CONTA", "INCLUI", "INCLUIR", "COM", "SOMAR", "TOTAL"])
    quer_express = "EXPRESS" in pergunta_norm or quer_somar_express

    if quer_express and not pode_ver_express:
        return "Não encontrei essa informação na visão atual. Posso consultar produção, cortes, religues, recusas e ranking operacional."

    if not pode_ver_financeiro and tipo == "faturamento":
        tipo = "resumo"

    if hasattr(st, "session_state"):
        st.session_state["chatbot_painel_contexto"] = {
            "mes": mes,
            "contrato": contrato,
            "recurso": recurso,
            "tipo": tipo,
            "metrica_ranking": metrica_ranking,
            "top_n": top_n,
            "dimensao_ranking": dimensao_ranking,
            "pergunta_original": pergunta,
        }

    escopo_txt = _escopo_texto_chat(contrato, recurso)
    periodo_txt = _nome_mes_chat(mes) if mes else "todo o histórico"

    if _pergunta_ultimos_meses_chat(pergunta_norm):
        df_meses = base.copy()
        if contrato:
            df_meses = df_meses[df_meses["CONTRATO"] == contrato].copy()
        if recurso:
            df_meses = df_meses[df_meses["RECURSO"] == recurso].copy()
        eh_recusa_m = pd.to_numeric(df_meses.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int)
        pagaveis_m = df_meses[eh_recusa_m == 0].copy()
        if pagaveis_m.empty:
            return f"Não encontrei notas feitas para {escopo_txt} nos últimos meses."
        pagaveis_m["MES"] = pagaveis_m["DATA_DT"].dt.strftime("%m/%Y")
        pagaveis_m["PERIODO"] = pagaveis_m["DATA_DT"].dt.to_period("M")
        mensal = (
            pagaveis_m.groupby(["MES", "PERIODO"], dropna=False)
            .agg(NOTAS=("ORDEM_DE_SERVICO", "nunique"), CORTES=("EH_CORTE", "sum"), RELIGUES=("EH_RELIGUE", "sum"), FATURAMENTO=("FATURAMENTO", "sum"))
            .reset_index()
            .sort_values("PERIODO", ascending=False)
            .head(6)
            .sort_values("PERIODO", ascending=True)
        )
        linhas = [f"📈 **{NOME_ASSISTENTE} — evolução de {escopo_txt}**", ""]
        for row in mensal.itertuples(index=False):
            express_info = _express_chat_periodo(notas, row.MES, contrato=contrato, recurso=recurso) if pode_ver_express else {"express": 0, "faturamento_express": 0.0}
            express_qtd = int(express_info.get("express", 0) or 0)
            fat_total = float(row.FATURAMENTO) + float(express_info.get("faturamento_express", 0.0) or 0.0)
            linha_mes = f"- **{_nome_mes_chat(row.MES)}:** {numero(row.NOTAS)} notas"
            if pode_ver_express and express_qtd:
                linha_mes += f" (+{numero(express_qtd)} express)"
            if pode_ver_financeiro:
                linha_mes += f" • {dinheiro(fat_total)}"
            linhas.append(linha_mes)
        return "\n".join(linhas)

    if tipo == "comparacao":
        import re
        # Compara dois meses citados na pergunta ou mês atual vs contexto anterior.
        meses_citados = []
        for nome_mes, num_mes in {
            "JANEIRO": "01", "FEVEREIRO": "02", "MARCO": "03", "MARÇO": "03", "ABRIL": "04",
            "MAIO": "05", "JUNHO": "06", "JULHO": "07", "AGOSTO": "08", "SETEMBRO": "09",
            "OUTUBRO": "10", "NOVEMBRO": "11", "DEZEMBRO": "12",
        }.items():
            if nome_mes in pergunta_norm:
                for mes_disp in meses:
                    if mes_disp.startswith(num_mes + "/") and mes_disp not in meses_citados:
                        meses_citados.append(mes_disp)
                        break
        for m in re.findall(r"\b(0?[1-9]|1[0-2])/(20\d{2})\b", pergunta_norm):
            candidato = f"{int(m[0]):02d}/{m[1]}"
            if candidato in meses and candidato not in meses_citados:
                meses_citados.append(candidato)
        if len(meses_citados) < 2 and contexto_anterior.get("mes") and mes and contexto_anterior.get("mes") != mes:
            meses_citados = [contexto_anterior.get("mes"), mes]
        if len(meses_citados) < 2:
            return "Para comparar, me diga dois meses. Exemplo: **comparar STC abril vs março**."

        m1, m2 = meses_citados[0], meses_citados[1]
        df1 = _filtrar_base_chat(base, mes=m1, contrato=contrato, recurso=recurso)
        df2 = _filtrar_base_chat(base, mes=m2, contrato=contrato, recurso=recurso)
        r1 = _resumo_numerico_chat(df1)
        r2 = _resumo_numerico_chat(df2)
        def var(a, b):
            if b == 0:
                return "novo" if a else "0,0%"
            return f"{((a-b)/b)*100:+.1f}%".replace(".", ",")
        alvo = recurso or contrato or "Geral"
        linhas = [f"📊 **Comparativo — {alvo}**", "", f"**{_nome_mes_chat(m1)} → {_nome_mes_chat(m2)}**", ""]
        linhas.append(f"• Produção: **{numero(r1['notas'])} → {numero(r2['notas'])}** ({var(r2['notas'], r1['notas'])})")
        linhas.append(f"• Cortes/religues/verificações: **{numero(r1['cortes'])}/{numero(r1['religues'])} → {numero(r2['cortes'])}/{numero(r2['religues'])}**")
        linhas.append(f"• Recusas: **{numero(r1['recusas'])} → {numero(r2['recusas'])}** ({var(r2['recusas'], r1['recusas'])})")
        if pode_ver_financeiro:
            linhas.append(f"• Faturamento: **{dinheiro(r1['faturamento'])} → {dinheiro(r2['faturamento'])}** ({var(r2['faturamento'], r1['faturamento'])})")
        if r2['notas'] > r1['notas']:
            linhas.append("✅ Tendência: produção cresceu no segundo período.")
        elif r2['notas'] < r1['notas']:
            linhas.append("⚠️ Tendência: produção caiu no segundo período.")
        else:
            linhas.append("➖ Tendência: produção estável.")
        return "\n".join(linhas)

    df = _filtrar_base_chat(base, mes=mes, contrato=contrato, recurso=recurso)
    express_info = _express_chat_periodo(notas, mes, contrato=contrato, recurso=recurso) if (mes and pode_ver_express) else {"express": 0, "faturamento_express": 0.0, "tem_base": False}

    if tipo == "ranking":
        if df.empty:
            return f"Não encontrei dados para montar ranking de {escopo_txt} em **{periodo_txt}**."
        if metrica_ranking == "faturamento" and not pode_ver_financeiro:
            metrica_ranking = "notas"

        titulo_metrica = {"notas": "maior produção", "recusas": "mais recusas", "media": "melhor média por dia", "faturamento": "maior faturamento"}.get(metrica_ranking, "ranking")

        # Se a pergunta pede CONTRATOS, o ranking deve ser agregado por contrato.
        # Isso evita responder com recursos/equipes quando o usuário perguntou "quais contratos".
        if dimensao_ranking == "contrato" and not contrato and not recurso:
            ranking_contratos = _montar_ranking_contratos_chat(df, metrica=metrica_ranking, pode_ver_financeiro=pode_ver_financeiro)
            if ranking_contratos.empty:
                return f"Não há dados suficientes para montar ranking de contratos em **{periodo_txt}**."

            if top_n == 1:
                row = ranking_contratos.iloc[0]
                linhas = [f"🏆 **Contrato com {titulo_metrica} — {periodo_txt}**", ""]
                if metrica_ranking == "recusas":
                    linhas.append(f"**{row['CONTRATO']}** liderou com **{numero(int(row['RECUSAS']))} recusas**.")
                elif metrica_ranking == "media":
                    linhas.append(f"**{row['CONTRATO']}** liderou com **{float(row['MEDIA_DIA']):.1f} notas/dia**.".replace(".", ","))
                elif metrica_ranking == "faturamento" and pode_ver_financeiro:
                    linhas.append(f"**{row['CONTRATO']}** liderou com **{dinheiro(float(row['FATURAMENTO']))}**.")
                else:
                    linhas.append(f"**{row['CONTRATO']}** liderou a produção com **{numero(int(row['NOTAS']))} notas**.")
                linhas.append("")
                detalhe = f"• Cortes / religues / verificações: **{numero(int(row['CORTES']))} / {numero(int(row['RELIGUES']))} / {numero(int(row.get('VERIFICACOES', 0)))}** • Recursos ativos: **{numero(int(row['RECURSOS_ATIVOS']))}** • Recusas: **{numero(int(row['RECUSAS']))}**"
                linhas.append(detalhe.replace(".", ","))
                return "\n".join(linhas)

            linhas = [f"🏆 **Top {top_n} contratos — {titulo_metrica} — {periodo_txt}**", ""]
            for _, row in ranking_contratos.head(top_n).iterrows():
                if metrica_ranking == "recusas":
                    linhas.append(f"{int(row['POSICAO'])}. **{row['CONTRATO']}** — {numero(int(row['RECUSAS']))} recusas • {numero(int(row['NOTAS']))} notas")
                elif metrica_ranking == "media":
                    linhas.append(f"{int(row['POSICAO'])}. **{row['CONTRATO']}** — {float(row['MEDIA_DIA']):.1f} notas/dia • {numero(int(row['NOTAS']))} notas".replace(".", ","))
                elif metrica_ranking == "faturamento" and pode_ver_financeiro:
                    linhas.append(f"{int(row['POSICAO'])}. **{row['CONTRATO']}** — {dinheiro(float(row['FATURAMENTO']))} • {numero(int(row['NOTAS']))} notas")
                else:
                    linhas.append(f"{int(row['POSICAO'])}. **{row['CONTRATO']}** — {numero(int(row['NOTAS']))} notas • Cortes/religues/verificações: {numero(int(row['CORTES']))}/{numero(int(row['RELIGUES']))}/{numero(int(row.get('VERIFICACOES', 0)))} • Recusas: {numero(int(row['RECUSAS']))}")
            linhas.append("")
            linhas.append("Obs.: ranking agregado por **contrato**, não por equipe/recurso.")
            return "\n".join(linhas)

        ranking = _montar_ranking_chat(df, metrica=metrica_ranking, pode_ver_financeiro=pode_ver_financeiro)
        if ranking.empty:
            return f"Não há dados suficientes para montar ranking em **{periodo_txt}**."
        contexto_titulo = contrato or ("Todos os contratos" if dimensao_ranking == "recurso" and escopo_geral_explicito else "Geral")
        if top_n == 1:
            row = ranking.iloc[0]
            linhas = [f"🏆 **Equipe com {titulo_metrica} — {contexto_titulo} — {periodo_txt}**", ""]
            if metrica_ranking == "recusas":
                linhas.append(f"**{row['RECURSO']}** liderou com **{numero(int(row['RECUSAS']))} recusas**.")
            elif metrica_ranking == "media":
                linhas.append(f"**{row['RECURSO']}** liderou com **{float(row['MEDIA_DIA']):.1f} notas/dia**.".replace(".", ","))
            elif metrica_ranking == "faturamento" and pode_ver_financeiro:
                linhas.append(f"**{row['RECURSO']}** liderou com **{dinheiro(float(row['FATURAMENTO']))}**.")
            else:
                linhas.append(f"**{row['RECURSO']}** liderou a produção com **{numero(int(row['NOTAS']))} notas**.")
            linhas.append("")
            detalhe = f"• Cortes / religues / verificações: **{numero(int(row['CORTES']))} / {numero(int(row['RELIGUES']))} / {numero(int(row.get('VERIFICACOES', 0)))}** • Média/dia: **{float(row['MEDIA_DIA']):.1f}** • Recusas: **{numero(int(row['RECUSAS']))}**"
            linhas.append(detalhe.replace(".", ","))
            motivo = _principal_recusa_recurso_chat(df, recurso=row["RECURSO"])
            if motivo:
                linhas.append(f"⚠️ {motivo}")
            return "\n".join(linhas)
        linhas = [f"🏆 **Top {top_n} equipes — {titulo_metrica} — {contexto_titulo} — {periodo_txt}**", ""]
        for _, row in ranking.head(top_n).iterrows():
            if metrica_ranking == "recusas":
                principal = _principal_recusa_recurso_chat(df, recurso=row["RECURSO"])
                extra = f" • {principal.replace('Principal motivo de recusa: ', '')}" if principal else ""
                linhas.append(f"{int(row['POSICAO'])}. **{row['RECURSO']}** — {numero(int(row['RECUSAS']))} recusas • {numero(int(row['NOTAS']))} notas{extra}")
            elif metrica_ranking == "media":
                linhas.append(f"{int(row['POSICAO'])}. **{row['RECURSO']}** — {float(row['MEDIA_DIA']):.1f} notas/dia • {numero(int(row['NOTAS']))} notas".replace(".", ","))
            elif metrica_ranking == "faturamento" and pode_ver_financeiro:
                linhas.append(f"{int(row['POSICAO'])}. **{row['RECURSO']}** — {dinheiro(float(row['FATURAMENTO']))} • {numero(int(row['NOTAS']))} notas")
            else:
                linhas.append(f"{int(row['POSICAO'])}. **{row['RECURSO']}** — {numero(int(row['NOTAS']))} notas • Cortes/religues/verificações: {numero(int(row['CORTES']))}/{numero(int(row['RELIGUES']))}/{numero(int(row.get('VERIFICACOES', 0)))} • Recusas: {numero(int(row['RECUSAS']))}")
        return "\n".join(linhas)

    if tipo == "recusas":
        resumo_recusas = _resumo_recusas_tipo(df)
        if resumo_recusas.empty:
            return f"Não encontrei recusas para {escopo_txt} em **{periodo_txt}**."
        total = int(resumo_recusas["QTD"].sum())
        linhas = [f"🚧 **{NOME_ASSISTENTE} — recusas de {escopo_txt} em {periodo_txt}**", "", f"Total de recusas: **{numero(total)}**", "", "**Por tipo:**"]
        for row in resumo_recusas.head(10).itertuples(index=False):
            linhas.append(f"- {row.RECUSA}: **{numero(row.QTD)}**")
        return "\n".join(linhas)

    resumo = _resumo_numerico_chat(df)
    express_qtd = int(express_info.get("express", 0) or 0)
    faturamento_express = float(express_info.get("faturamento_express", 0.0) or 0.0)
    resumo_recusas = _resumo_recusas_tipo(df)
    if df.empty and express_qtd == 0:
        return f"Não encontrei dados para {escopo_txt} em **{periodo_txt}**."

    if tipo == "express" and not quer_somar_express:
        linhas = [f"⚡ **{NOME_ASSISTENTE} — pagamento express de {escopo_txt} em {periodo_txt}**", "", f"• Express: **{numero(express_qtd)}**"]
        if faturamento_express and pode_ver_financeiro:
            linhas.append(f"• Faturamento express: **{dinheiro(faturamento_express)}**")
        if not express_info.get("tem_base"):
            linhas.extend(["", "Obs.: não encontrei a planilha de Pagamento Express no local configurado."])
        return "\n".join(linhas)

    faturamento_base = float(resumo["faturamento"])
    faturamento_total = faturamento_base + faturamento_express
    total_atendimentos = int(resumo["notas"] + (express_qtd if pode_ver_express else 0))
    volume_com_recusas = total_atendimentos + int(resumo["recusas"])
    taxa_recusa = (resumo["recusas"] / volume_com_recusas * 100) if volume_com_recusas else 0
    media_dia = (resumo["notas"] / resumo["dias_ativos"]) if resumo["dias_ativos"] else 0

    titulo_alvo = recurso or contrato or "Geral"
    linhas = [f"📊 **{titulo_alvo} — {periodo_txt}**", ""]
    linhas.append(f"• **Produção:** {numero(resumo['notas'])} notas" + (f" (+{numero(express_qtd)} express)" if (pode_ver_express and express_qtd) else ""))
    linhas.append(f"• **Total de atendimentos:** {numero(total_atendimentos)}")
    if pode_ver_financeiro:
        linhas.append(f"• **Faturamento:** {dinheiro(faturamento_total)}")
        if faturamento_express and pode_ver_express:
            linhas.append(f"  - Sem express: {dinheiro(faturamento_base)}")
            linhas.append(f"  - Express: {dinheiro(faturamento_express)}")
    linhas.append(f"• **Cortes / religues / verificações:** {numero(resumo['cortes'])} / {numero(resumo['religues'])}")
    linhas.append(f"• **Recusas:** {numero(resumo['recusas'])} ({taxa_recusa:.1f}%)".replace(".", ","))
    if resumo["dias_ativos"]:
        linhas.append(f"• **Média/dia:** {media_dia:.1f} notas".replace(".", ","))
    if not recurso:
        linhas.append(f"• **Recursos ativos:** {numero(resumo['recursos'])}")
    linhas.append("")
    for destaque in _destaques_executivos_chat(resumo, express_qtd if pode_ver_express else 0, resumo_recusas):
        linhas.append(destaque)
    if any(t in pergunta_norm for t in ["COMO FOI", "RESUMO", "RESULTADO"]):
        ranking = _montar_ranking_chat(df, metrica="notas", pode_ver_financeiro=pode_ver_financeiro)
        if not ranking.empty:
            top = ranking.iloc[0]
            linhas.append(f"Destaque de produção: **{top['RECURSO']}** com {numero(int(top['NOTAS']))} notas.")
    if quer_express and not express_info.get("tem_base"):
        linhas.extend(["", "Obs.: não encontrei a planilha de Pagamento Express no local configurado."])
    return "\n".join(linhas)


def mostrar_chatbot_popup(notas, pode_ver_financeiro=True, pode_ver_express=True, modo_leitura=False):
    """Mostra o G.Z.U.S. em formato de popup fixo no canto inferior direito, respeitando o perfil logado."""
    st.markdown(
        """
        <style>
        .gzus-brand-card {
            background: linear-gradient(135deg, rgba(15,23,42,0.98), rgba(29,78,216,0.92));
            color: white;
            border-radius: 18px;
            padding: 12px 14px;
            margin-bottom: 10px;
            border: 1px solid rgba(147,197,253,0.25);
        }
        .gzus-brand-card .gzus-title {
            font-weight: 950;
            letter-spacing: 0.08em;
            font-size: 1rem;
        }
        .gzus-brand-card .gzus-subtitle {
            color: #bfdbfe;
            font-size: 0.78rem;
            margin-top: 2px;
        }
        div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) {
            position: fixed !important;
            right: 22px !important;
            bottom: 22px !important;
            width: min(440px, calc(100vw - 44px)) !important;
            max-height: min(660px, calc(100vh - 44px));
            overflow: auto;
            z-index: 9999 !important;
            background: rgba(15, 23, 42, 0.96);
            border: 1px solid rgba(147, 197, 253, 0.28);
            border-radius: 24px;
            padding: 10px 12px 12px 12px;
            box-shadow: 0 22px 65px rgba(0,0,0,0.38);
        }
        div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) details {
            border: 0 !important;
        }
        div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) summary {
            font-weight: 900;
        }
        div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) [data-testid="stMarkdownContainer"] p,
        div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) [data-testid="stMarkdownContainer"] li {
            font-size: 0.92rem;
        }
        @media (max-width: 768px) {
            div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) {
                left: 10px !important;
                right: 10px !important;
                bottom: 10px !important;
                width: auto !important;
                max-height: 48vh !important;
                padding: 8px 10px 10px 10px !important;
                border-radius: 20px !important;
                overflow-y: auto !important;
                background: rgba(15, 23, 42, 0.98) !important;
                box-shadow: 0 12px 38px rgba(0,0,0,0.48) !important;
            }
            div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) summary {
                font-size: 0.92rem !important;
                color: #f8fafc !important;
            }
            div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) .gzus-brand-card {
                padding: 10px 12px !important;
                margin-bottom: 8px !important;
            }
            div[data-testid="stVerticalBlock"]:has(.chatbot-popup-anchor):not(:has(div[data-testid="stVerticalBlock"] .chatbot-popup-anchor)) input {
                font-size: 16px !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.container():
        st.markdown('<span class="chatbot-popup-anchor"></span>', unsafe_allow_html=True)
        with st.expander("🤖 G.Z.U.S. Assistente", expanded=False):
            st.markdown(
                f"""
                <div class="gzus-brand-card">
                    <div class="gzus-title">G.Z.U.S.</div>
                    <div class="gzus-subtitle">Gestão Inteligente de Serviços</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.caption("Pergunte de forma natural: “5981 fez quanto?”, “carro faturou quanto?”, “como foi abril?”")

            if "chatbot_painel_historico" not in st.session_state:
                st.session_state.chatbot_painel_historico = []

            for item in st.session_state.chatbot_painel_historico[-4:]:
                st.markdown(f"**Você:** {item['pergunta']}")
                st.markdown(item["resposta"])
                st.markdown("---")

            pergunta = st.text_input(
                "Pergunta",
                placeholder="Ex: 5981 fez quanto em abril?",
                key="chatbot_painel_pergunta",
                label_visibility="collapsed",
            )
            col1, col2 = st.columns([2, 1])
            with col1:
                enviar = st.button("Perguntar", use_container_width=True, key="chatbot_painel_enviar")
            with col2:
                limpar = st.button("Limpar", use_container_width=True, key="chatbot_painel_limpar")

            if limpar:
                st.session_state.chatbot_painel_historico = []
                st.session_state["chatbot_painel_contexto"] = {}
                st.rerun()

            if enviar and pergunta.strip():
                resposta = responder_chatbot_painel(
                    pergunta,
                    notas,
                    pode_ver_financeiro=pode_ver_financeiro,
                    pode_ver_express=pode_ver_express,
                    modo_leitura=modo_leitura,
                )
                st.session_state.chatbot_painel_historico.append({"pergunta": pergunta, "resposta": resposta})
                st.rerun()



# Antes de carregar os CSV/Excel, o app PODE puxar do GitHub a versão mais recente.
# Ajuste de velocidade:
# Não fazemos mais git fetch/reset automático durante a navegação.
# Isso era o principal candidato aos 10-15 segundos de espera no Streamlit Cloud.
# A atualização continua existindo pelo botão "Atualizar dados".
_status_sync_github = _ler_status_github_sync() or {
    "ok": True,
    "changed": False,
    "skipped": True,
    "message": "GitHub em modo manual para abrir mais rápido.",
    "quando": datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M:%S"),
}

bases, faltando = carregar_bases_leves()

# Diagnóstico discreto da fonte usada. Se aparecer, o painel já está lendo o gzus.db.
try:
    fontes_usadas = st.session_state.get("fontes_dados_dashboard", {})
    com_sqlite = [k for k, v in fontes_usadas.items() if v == "sqlite"]
    if com_sqlite:
        st.sidebar.caption("🗄️ SQLite ativo: " + ", ".join(com_sqlite))
except Exception:
    pass

if PERFIL_ACESSO == "supervisor_stc":
    bases = filtrar_bases_para_supervisor_stc(bases)

if not bases:
    st.error("Nenhum CSV foi encontrado. Verifique se os arquivos estão na pasta dashboard.")
    st.stop()

contratos_original = bases.get("contratos", pd.DataFrame())
carro_original = bases.get("carro", pd.DataFrame())
dias_original = bases.get("dias", pd.DataFrame())
carro_dias_original = bases.get("carro_dias", pd.DataFrame())
notas = pd.DataFrame()  # carregada sob demanda por SQL filtrado

# Perfis restritos têm telas próprias para evitar exposição acidental de dados financeiros.
if PERFIL_ACESSO == "supervisor_leitura":
    st.title("📖 Leitura temporariamente desativada")
    st.info("A área de Leitura foi removida temporariamente para acelerar o painel principal.")
    st.stop()

if PERFIL_ACESSO == "supervisor_stc":
    mostrar_painel_supervisor_stc(bases)
    if not notas.empty:
        mostrar_chatbot_popup(notas, pode_ver_financeiro=False, pode_ver_express=False)
    st.stop()

st.title("🤖 G.Z.U.S. — Gestão Inteligente de Serviços")
st.caption("Painel operacional com assistente inteligente. Atualização automática com GitHub e banco leve SQLite.")
st.sidebar.caption(f"Perfil: {NOME_ACESSO}")
if isinstance(_status_sync_github, dict) and _status_sync_github.get("quando"):
    st.sidebar.caption(f"GitHub: {_status_sync_github.get('message', '')} ({_status_sync_github.get('quando')})")

if st.session_state.pop("github_dados_atualizados_sem_recarregar", False):
    st.sidebar.success("✅ Dados novos aplicados sem recarregar o painel.")
    st.toast("Dados novos aplicados sem interromper o uso do painel.", icon="✅")

if faltando:
    st.warning("Arquivos não encontrados: " + ", ".join(faltando))

# Popup do assistente: por padrão fica desligado no carregamento inicial para deixar
# o pós-login mais rápido. Para religar, coloque ASSISTENTE_GZUS_AUTO = "true" nos Secrets.
ASSISTENTE_GZUS_AUTO = str(secret_value("ASSISTENTE_GZUS_AUTO", "false") or "false").strip().lower() in ["1", "true", "sim", "s", "yes", "on"]
if ASSISTENTE_GZUS_AUTO and PERFIL_ACESSO == "gerente" and not notas.empty:
    mostrar_chatbot_popup(notas, pode_ver_financeiro=True, pode_ver_express=True)

# ==============================
# FILTROS EM BOTÕES
# ==============================

st.sidebar.header("Filtros")

if st.sidebar.button("🔄 Atualizar dados", use_container_width=True):
    status_manual = sincronizar_github_se_preciso(forcar=True)
    st.cache_data.clear()
    if status_manual.get("changed"):
        st.sidebar.success("Atualizado pelo GitHub. Os dados serão aplicados no próximo ciclo da tela.")
        st.toast("Atualização baixada. Continue usando; a tela aplicará os dados no próximo refresh.", icon="✅")
    elif status_manual.get("ok"):
        st.sidebar.info("Cache limpo. Nenhum commit novo no GitHub.")
    else:
        st.sidebar.warning(status_manual.get("message", "Não consegui consultar o GitHub, mas limpei o cache local."))


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

# Área de Leitura removida temporariamente para reduzir carga inicial.
if st.session_state.get("modo_painel") == "leitura":
    st.session_state.modo_painel = "corte"

modo_painel = st.session_state.modo_painel
contrato_escolhido = st.session_state.contrato_escolhido
contrato_filtro_notas = contrato_para_base_notas(contrato_escolhido)

st.sidebar.markdown("---")
st.sidebar.markdown("**Tela selecionada:**")
st.sidebar.info(contrato_escolhido)

# Status detalhado depende da tabela grande de notas; fica fora do pós-login para acelerar.
# Ele pode voltar depois em versão SQL agregada, sem carregar notas brutas.

# Este período vale para a tela inicial "Resumo".
# Por padrão, fica só no mês mais recente da base, para não somar março + abril sem querer.
meses_base = meses_disponiveis_leves(dias_original, carro_dias_original)
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
        # Para STC, o faturamento estimado pode estar registrado como
        # "Contrato Carro STC estimado". Mantemos as duas visões juntas
        # para a Home/Resumo conseguir exibir mínimo e máximo do STC.
        if contrato_escolhido in ["STC Jundiai", "Contrato Carro STC estimado"]:
            carro_dias = carro_dias[carro_dias["CONTRATO"].isin(["STC Jundiai", "Contrato Carro STC estimado"])]
        else:
            carro_dias = carro_dias[carro_dias["CONTRATO"] == contrato_escolhido]

mostrar_carro = not carro.empty

mostrar_aba_carro = (contrato_escolhido == "Todos") or (contrato_filtro_notas == "STC Jundiai") or (not carro.empty)

nomes_abas = ["Resumo", "Parcial do dia", "Ranking de recursos", "Comparativo mensal", "Dias da semana"]
if mostrar_aba_carro:
    nomes_abas.append("STC")
nomes_abas += ["Notas", "Downloads"]

# Mais leve que st.tabs: no Streamlit, todas as abas executam ao mesmo tempo.
# Com radio, só a tela escolhida roda, reduzindo carregamento após login e troca de filtros.
tela_escolhida = st.radio(
    "Tela",
    nomes_abas,
    horizontal=True,
    label_visibility="collapsed",
    key="tela_principal_gzus",
)

# A tabela grande de notas NÃO é mais carregada aqui.
# Cada tela pesada carrega somente quando for aberta.
notas = pd.DataFrame()

# ==============================
# ABA RESUMO
# ==============================

if tela_escolhida == "Resumo":
    resumo_contrato_periodo, resumo_grupo_periodo = resumo_home_leve(
        dias,
        carro_dias,
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
            resumo_contrato_periodo["CONTRATO"].isin(["STC Jundiai", "Contrato Carro STC estimado"])
        ].copy()

        mostrar_carro_periodo = not carro_periodo.empty

        if mostrar_carro_periodo:
            total_carro_min = carro_periodo["FATURAMENTO_MIN"].sum()
            total_carro_max = carro_periodo["FATURAMENTO_MAX"].sum()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Faturamento contratos", dinheiro(total_contratos))
            c2.metric("STC mínimo", dinheiro(total_carro_min))
            c3.metric("STC máximo", dinheiro(total_carro_max))
            c4.metric("Notas únicas", numero(qtd_notas))
        else:
            c1, c2 = st.columns(2)
            c1.metric("Faturamento contratos", dinheiro(total_contratos))
            c2.metric("Notas únicas", numero(qtd_notas))

        # Para manter a abertura inicial rápida, a meta CPFL detalhada fica nas telas
        # Parcial / Ranking, que carregam notas sob demanda.

        st.subheader("Faturamento por contrato")

        grafico_resumo = resumo_contrato_periodo.copy()
        st.bar_chart(grafico_resumo, x="CONTRATO", y="FATURAMENTO")

        st.markdown("**Resumo com corte + religue + verificação**")

        if resumo_grupo_periodo.empty:
            st.caption("Nesta versão rápida, o detalhamento corte/religue da Home fica fora do carregamento inicial. Use Parcial, Ranking ou Notas para detalhes operacionais.")
        else:
            tabela_resumo = resumo_grupo_periodo.pivot_table(
                index="CONTRATO",
                columns="GRUPO_NOTA",
                values="FATURAMENTO",
                aggfunc="sum",
                fill_value=0,
            ).reset_index()

            for col in ["CORTE", "RELIGUE", "VERIFICACAO"]:
                if col not in tabela_resumo.columns:
                    tabela_resumo[col] = 0

            tabela_resumo["TOTAL"] = tabela_resumo[["CORTE", "RELIGUE", "VERIFICACAO"]].sum(axis=1)
            tabela_resumo = tabela_resumo[["CONTRATO", "CORTE", "RELIGUE", "VERIFICACAO", "TOTAL"]]

            st.dataframe(formatar_tabela(tabela_resumo), use_container_width=True, hide_index=True)

        st.markdown("**Detalhamento por contrato no período**")
        colunas_detalhe_resumo = [
            "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "EXPRESS",
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

if tela_escolhida == "Parcial do dia":
    notas = carregar_notas_rapido(meses_escolhidos_resumo)
    st.subheader("Parcial do dia por recurso")
    if contrato_escolhido != contrato_filtro_notas:
        st.caption(f"Exibindo a base operacional de notas: {contrato_filtro_notas}.")

    # Base com recusas para mostrar na parcial.
    parcial_com_recusas = preparar_parcial_do_dia(notas, incluir_recusas=True)

    if parcial_com_recusas.empty:
        st.info("Ainda não há dados suficientes para montar a parcial do dia.")
    else:
        if contrato_filtro_notas != "Todos" and "CONTRATO" in parcial_com_recusas.columns:
            parcial_com_recusas = parcial_com_recusas[parcial_com_recusas["CONTRATO"] == contrato_filtro_notas]

        datas_disponiveis = (
            parcial_com_recusas[["DATA", "DATA_DT"]]
            .drop_duplicates()
            .sort_values("DATA_DT", ascending=False)
        )

        if datas_disponiveis.empty:
            st.info("Nenhuma data encontrada na base de notas para este contrato/filtro.")
        else:
            opcoes_datas = datas_disponiveis["DATA"].tolist()
            data_escolhida = st.selectbox("Escolha o dia", opcoes_datas, index=0)

            dados_dia_cache = calcular_parcial_dia_processada_cache(parcial_com_recusas, data_escolhida)
            parcial_dia_tudo = dados_dia_cache["parcial_dia_tudo"].copy()
            parcial_dia = dados_dia_cache["parcial_dia"].copy()
            recusas_dia = dados_dia_cache["recusas_dia"].copy()
            totais_dia = dados_dia_cache["totais"]

            if parcial_dia_tudo.empty:
                st.info("Nenhuma nota encontrada para esse dia.")
            else:
                total_notas = int(totais_dia.get("total_notas", 0))
                total_recursos_ativos = int(totais_dia.get("total_recursos_ativos", 0))
                total_cortes = int(totais_dia.get("total_cortes", 0))
                total_religues = int(totais_dia.get("total_religues", 0))
                total_recusas = int(totais_dia.get("total_recusas", 0))
                total_faturamento = float(totais_dia.get("total_faturamento", 0.0))
                total_faturamento_min = float(totais_dia.get("total_faturamento_min", 0.0))
                total_faturamento_max = float(totais_dia.get("total_faturamento_max", 0.0))

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Recursos ativos", numero(total_recursos_ativos))
                c2.metric("Notas feitas", numero(total_notas))
                c3.metric("Cortes", numero(total_cortes))
                c4.metric("Religues", numero(total_religues))
                c5.metric("Recusas", numero(total_recusas))

                if contrato_filtro_notas == "STC Jundiai":
                    data_meta_inicio, data_meta_fim = _periodo_datas_cpfl("Dia", data_escolhida)
                    meta_cpfl = meta_cpfl_stc_periodo(data_meta_inicio, data_meta_fim)
                    express_cpfl = contar_express_cpfl_periodo(notas, "STC Jundiai", data_meta_inicio, data_meta_fim)
                    render_meta_cpfl_stc("Meta CPFL do dia", meta_cpfl, total_cortes, express_cpfl)

                tem_carro_no_dia = "CONTRATO" in parcial_dia.columns and (
                    parcial_dia["CONTRATO"] == "STC Jundiai"
                ).any()

                if tem_carro_no_dia:
                    st.metric("Faturamento estimado", f"{dinheiro(total_faturamento_min)} a {dinheiro(total_faturamento_max)}")
                else:
                    st.metric("Faturamento", dinheiro(total_faturamento))

                st.markdown('<div class="section-title">Ranking do dia por produção</div>', unsafe_allow_html=True)

                resumo_equipe = dados_dia_cache["resumo_equipe"].copy()

                if resumo_equipe.empty:
                    st.info("Nenhuma nota ou recusa encontrada para esse dia.")
                else:
                    resumo_equipe.insert(0, "POSIÇÃO", range(1, len(resumo_equipe) + 1))

                    recursos_sem_movimento = dados_dia_cache["recursos_sem_movimento"].copy()
                    render_alerta_recursos_sem_movimento(
                        recursos_sem_movimento,
                        contrato_unico=(contrato_escolhido != "Todos"),
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
                                alt.Tooltip("VERIFICACOES:Q", title="Verificações"),
                                alt.Tooltip("RECUSAS:Q", title="Recusas"),
                                alt.Tooltip("FATURAMENTO:Q", title="Faturamento", format=",.2f"),
                            ],
                        )
                        .properties(height=330)
                    )

                    st.altair_chart(grafico_parcial, use_container_width=True)

                    def faturamento_linha_equipe(row):
                        if row.get("CONTRATO") == "STC Jundiai":
                            return f"{dinheiro(row.get('FATURAMENTO_MIN', 0))} a {dinheiro(row.get('FATURAMENTO_MAX', 0))}"
                        return dinheiro(row.get("FATURAMENTO", 0))

                    tabela_equipe = resumo_equipe.copy()
                    tabela_equipe["FATURAMENTO"] = tabela_equipe.apply(faturamento_linha_equipe, axis=1)
                    tabela_equipe = tabela_equipe[[
                        "POSIÇÃO", "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "RECUSAS", "FATURAMENTO"
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

if tela_escolhida == "Ranking de recursos":
    notas = carregar_notas_rapido(meses_escolhidos_resumo)
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
            index=contratos_exec.index(contrato_filtro_notas) if contrato_filtro_notas in contratos_exec else 0,
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

            if contrato_ranking == "STC Jundiai" and tipo_periodo in ["Dia", "Semana", "Mês"] and valor_periodo:
                meta_inicio, meta_fim = _periodo_datas_cpfl(tipo_periodo, valor_periodo)
                meta_cpfl = meta_cpfl_stc_periodo(meta_inicio, meta_fim)
                cortes_cpfl = int(base_filtrada_exec.loc[
                    pd.to_numeric(base_filtrada_exec.get("EH_RECUSA", 0), errors="coerce").fillna(0).astype(int) == 0,
                    "EH_CORTE"
                ].sum()) if not base_filtrada_exec.empty and "EH_CORTE" in base_filtrada_exec.columns else 0
                express_cpfl = contar_express_cpfl_periodo(notas, "STC Jundiai", meta_inicio, meta_fim)
                titulo_meta = "Meta CPFL da semana" if tipo_periodo == "Semana" else ("Meta CPFL do mês" if tipo_periodo == "Mês" else "Meta CPFL do dia")
                render_meta_cpfl_stc(titulo_meta, meta_cpfl, cortes_cpfl, express_cpfl)

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
                "POSIÇÃO", "RECURSO", "NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "EXPRESS", "RECUSAS", "DIAS_ATIVOS",
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
                    "GRUPO_NOTA", "EH_CORTE", "EH_RELIGUE", "EH_VERIFICACAO", "FATURAMENTO", "FATURAMENTO_ATRIBUÍDO"
                ]
                detalhe_cols = [c for c in detalhe_cols if c in base_filtrada_exec.columns]
                detalhe_base = base_filtrada_exec.copy()
                if "EH_RECUSA" in detalhe_base.columns:
                    detalhe_base = detalhe_base[pd.to_numeric(detalhe_base["EH_RECUSA"], errors="coerce").fillna(0).astype(int) == 0].copy()
                detalhe = detalhe_base[detalhe_cols].sort_values(["DATA", "RECURSO"], ascending=[False, True])
                st.dataframe(
                    preparar_tabela_ranking(detalhe, colunas_moeda=["FATURAMENTO", "FATURAMENTO_ATRIBUÍDO"]),
                    use_container_width=True,
                    hide_index=True,
                )

            if tipo_periodo == "Mês" and valor_periodo and not express_resumo_recurso.empty:
                st.markdown('<div class="section-title">Pagamento Express conciliado por recurso</div>', unsafe_allow_html=True)
                tabela_express_recurso = express_resumo_recurso.copy().sort_values(
                    ["EXPRESS", "FATURAMENTO_EXPRESS"], ascending=False
                )
                st.dataframe(
                    formatar_tabela(tabela_express_recurso[["RECURSO", "CONTRATO", "EXPRESS", "FATURAMENTO_EXPRESS"]]),
                    use_container_width=True,
                    hide_index=True,
                )

            st.markdown('<div class="section-title">Resumo de recusas</div>', unsafe_allow_html=True)
            recusas_tipo = calcular_recusas_por_tipo(base_filtrada_exec)
            if recusas_tipo.empty:
                st.success("Nenhuma recusa encontrada para os filtros selecionados.")
            else:
                total_recusas_periodo = int(recusas_tipo["QTD_RECUSAS"].sum())
                st.caption(f"Total de recusas no período filtrado: {numero(total_recusas_periodo)}")

                st.markdown("**Total por tipo de recusa**")
                total_por_tipo = (
                    recusas_tipo.groupby("RECUSA", dropna=False)
                    .agg(QTD_RECUSAS=("QTD_RECUSAS", "sum"))
                    .reset_index()
                    .sort_values(["QTD_RECUSAS", "RECUSA"], ascending=[False, True])
                )
                st.dataframe(
                    preparar_tabela_ranking(total_por_tipo),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("**Total por contrato e tipo de recusa**")
                total_por_contrato = (
                    recusas_tipo.groupby(["CONTRATO", "RECUSA"], dropna=False)
                    .agg(QTD_RECUSAS=("QTD_RECUSAS", "sum"))
                    .reset_index()
                    .sort_values(["CONTRATO", "QTD_RECUSAS", "RECUSA"], ascending=[True, False, True])
                )
                st.dataframe(
                    preparar_tabela_ranking(total_por_contrato),
                    use_container_width=True,
                    hide_index=True,
                )

                st.markdown("**Detalhamento por equipe, contrato e tipo de recusa**")
                recusas_tipo = recusas_tipo.sort_values(
                    ["RECURSO", "CONTRATO", "QTD_RECUSAS", "RECUSA"],
                    ascending=[True, True, False, True],
                )
                st.dataframe(
                    preparar_tabela_ranking(recusas_tipo),
                    use_container_width=True,
                    hide_index=True,
                )

            render_auditoria_express_ranking(
                tipo_periodo, valor_periodo, express_caminho, express_data_max, express_sem_vinculo,
                express_resumo_recurso, total_express_mensal
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

if tela_escolhida == "Comparativo mensal":
    # Comparativo precisa enxergar meses diferentes; carrega notas somente ao abrir esta tela.
    notas = carregar_notas_rapido(None)
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

        atual = calcular_resumo_mensal(notas, mes_escolhido, contrato_filtro_notas)
        anterior = calcular_resumo_mensal(notas, mes_anterior, contrato_filtro_notas)

        st.markdown(f"**Comparando: {mes_escolhido} x {mes_anterior}**")

        data_max_comp = data_maxima_do_mes(notas, mes_escolhido)
        if data_max_comp is not None:
            ultimo_dia_comp = data_max_comp.to_period("M").end_time.normalize()
            if data_max_comp.normalize() < ultimo_dia_comp:
                st.info(
                    f"Atenção: {mes_escolhido} ainda é parcial. "
                    f"Os dados vão até {data_max_comp.strftime('%d/%m/%Y')}."
                )

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Faturamento", dinheiro(atual["FATURAMENTO"]), variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"]))
        c2.metric("Notas", numero(atual["TOTAL_NOTAS"]), variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"]))
        c3.metric("Cortes", numero(atual["CORTES"]), variacao_percentual(atual["CORTES"], anterior["CORTES"]))
        c4.metric("Religues", numero(atual["RELIGUES"]), variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"]))
        c5.metric("Verificações", numero(atual.get("VERIFICACOES", 0)), variacao_percentual(atual.get("VERIFICACOES", 0), anterior.get("VERIFICACOES", 0)))
        c6.metric("Express", numero(atual.get("EXPRESS", 0)), variacao_percentual(atual.get("EXPRESS", 0), anterior.get("EXPRESS", 0)))

        tabela_comparativo = pd.DataFrame([
            {"Indicador": "Faturamento", mes_escolhido: dinheiro(atual["FATURAMENTO"]), mes_anterior: dinheiro(anterior["FATURAMENTO"]), "Variação": variacao_percentual(atual["FATURAMENTO"], anterior["FATURAMENTO"])},
            {"Indicador": "Notas", mes_escolhido: numero(atual["TOTAL_NOTAS"]), mes_anterior: numero(anterior["TOTAL_NOTAS"]), "Variação": variacao_percentual(atual["TOTAL_NOTAS"], anterior["TOTAL_NOTAS"])},
            {"Indicador": "Cortes", mes_escolhido: numero(atual["CORTES"]), mes_anterior: numero(anterior["CORTES"]), "Variação": variacao_percentual(atual["CORTES"], anterior["CORTES"])},
            {"Indicador": "Religues", mes_escolhido: numero(atual["RELIGUES"]), mes_anterior: numero(anterior["RELIGUES"]), "Variação": variacao_percentual(atual["RELIGUES"], anterior["RELIGUES"])},
            {"Indicador": "Verificações", mes_escolhido: numero(atual.get("VERIFICACOES", 0)), mes_anterior: numero(anterior.get("VERIFICACOES", 0)), "Variação": variacao_percentual(atual.get("VERIFICACOES", 0), anterior.get("VERIFICACOES", 0))},
            {"Indicador": "Express", mes_escolhido: numero(atual.get("EXPRESS", 0)), mes_anterior: numero(anterior.get("EXPRESS", 0)), "Variação": variacao_percentual(atual.get("EXPRESS", 0), anterior.get("EXPRESS", 0))},
            {"Indicador": "Faturamento Express", mes_escolhido: dinheiro(atual.get("FATURAMENTO_EXPRESS", 0)), mes_anterior: dinheiro(anterior.get("FATURAMENTO_EXPRESS", 0)), "Variação": variacao_percentual(atual.get("FATURAMENTO_EXPRESS", 0), anterior.get("FATURAMENTO_EXPRESS", 0))},
        ])
        st.dataframe(tabela_comparativo, use_container_width=True, hide_index=True)

        st.markdown("**Evolução mês a mês**")
        linhas_evolucao = []
        for mes in reversed(opcoes_meses):
            r = calcular_resumo_mensal(notas, mes, contrato_filtro_notas)
            linhas_evolucao.append({
                "MES": mes,
                "FATURAMENTO": r["FATURAMENTO"],
                "NOTAS": r["TOTAL_NOTAS"],
                "CORTES": r["CORTES"],
                "RELIGUES": r["RELIGUES"],
                "VERIFICACOES": r.get("VERIFICACOES", 0),
                "EXPRESS": r.get("EXPRESS", 0),
                "FATURAMENTO_EXPRESS": r.get("FATURAMENTO_EXPRESS", 0),
            })
        evolucao = pd.DataFrame(linhas_evolucao)

        if not evolucao.empty:
            st.bar_chart(evolucao, x="MES", y="FATURAMENTO")
            st.dataframe(formatar_tabela(evolucao), use_container_width=True, hide_index=True)

        st.markdown("**Resumo por contrato no mês escolhido**")
        resumo_mes, _ = resumo_por_periodo(notas, [mes_escolhido], contrato_filtro_notas)
        if resumo_mes.empty:
            st.info("Nenhum dado encontrado para esse mês.")
        else:
            colunas = ["CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "VERIFICACOES", "EXPRESS", "FATURAMENTO", "FATURAMENTO_EXPRESS", "FATURAMENTO_MIN", "FATURAMENTO_MAX"]
            colunas = [c for c in colunas if c in resumo_mes.columns]
            st.dataframe(formatar_tabela(resumo_mes[colunas]), use_container_width=True, hide_index=True)

# ==============================
# ABA DIAS
# ==============================

if tela_escolhida == "Dias da semana":
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

if mostrar_aba_carro and tela_escolhida == "STC":
        st.subheader("Contrato do carro — estimativa")

        if not carro.empty:
            c1, c2 = st.columns(2)
            c1.metric("Mínimo estimado", dinheiro(carro["FATURAMENTO_MIN"].sum()))
            c2.metric("Máximo estimado", dinheiro(carro["FATURAMENTO_MAX"].sum()))

            st.dataframe(formatar_tabela(carro), use_container_width=True, hide_index=True)
        else:
            st.info("Selecione o contrato do carro no menu lateral para visualizar.")

        st.subheader("STC estimado por dia")

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

if tela_escolhida == "Notas":
    notas = carregar_notas_rapido(meses_escolhidos_resumo)
    st.subheader("Consulta de notas")

    if not notas.empty:
        df_notas = notas.copy()

        # A base de notas acumulada não tem contrato salvo. Por isso, para filtrar por contrato,
        # reaproveitamos a classificação feita na parcial.
        parcial_para_filtro = preparar_parcial_do_dia(notas)
        if contrato_filtro_notas != "Todos" and not parcial_para_filtro.empty:
            ordens_do_contrato = parcial_para_filtro.loc[
                parcial_para_filtro["CONTRATO"] == contrato_filtro_notas,
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
# TXT SUPERVISÃO / TXT DO DIA
# ==============================

# ==============================
# TXT SUPERVISÃO
# ==============================

def _valor_txt_supervisao(valor):
    """Formata célula para TXT tabulado, preservando vazio como vazio."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    texto = str(valor).strip()
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]
    # Evita quebrar a estrutura de colunas do TXT.
    texto = texto.replace("\t", " ").replace("\r", " ").replace("\n", " ")
    return texto


def _data_hora_txt_supervisao(row):
    """Usa DATA_ENCERRAMENTO quando existir; senão usa DATA."""
    valor = row.get("DATA_ENCERRAMENTO", "")
    if _valor_txt_supervisao(valor) == "":
        valor = row.get("DATA", "")

    dt = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        # Se houver hora, mantém dd/mm/aaaa hh:mm. Se não houver, mantém só a data.
        if dt.hour or dt.minute or dt.second:
            return dt.strftime("%d/%m/%Y %H:%M")
        return dt.strftime("%d/%m/%Y")

    return _valor_txt_supervisao(valor)


def _status_txt_supervisao(row):
    """Status no formato do TXT operacional: FINALIZADA ou REJEITADA.

    Alguns CSVs/bancos trazem STATUS vazio ou com outro texto interno.
    Para o TXT dos supervisores, a regra mais estável é:
    - se há RECUSA/observação de recusa => REJEITADA
    - senão => FINALIZADA
    Se o STATUS já vier FINALIZADA/REJEITADA, preserva.
    """
    status = _valor_txt_supervisao(row.get("STATUS", "")).upper()
    recusa = _valor_txt_supervisao(row.get("RECUSA", ""))
    if status in ["FINALIZADA", "REJEITADA"]:
        return status
    return "REJEITADA" if recusa else "FINALIZADA"


def gerar_txt_supervisao(notas, data_escolhida=None, contrato_filtro="Todos", grupo_filtro="Todos"):
    """Gera TXT no formato operacional usado para colar no Excel.

    Formato sem cabeçalho e separado por TAB:
    OS | TIPO | RECURSO | STATUS | DATA/HORA | ELETRICISTA1 | ELETRICISTA2 | RECUSA
    """
    if notas is None or notas.empty:
        return "", pd.DataFrame()

    df = notas.copy()

    # Garante colunas esperadas sem quebrar caso o banco leve mude.
    for col in [
        "ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "STATUS", "RECUSA",
        "ELETRICISTA1", "ELETRICISTA2", "DATA", "DATA_ENCERRAMENTO",
    ]:
        if col not in df.columns:
            df[col] = ""

    # Usa a parcial apenas para descobrir o contrato operacional, mas mantém as colunas originais
    # para o TXT sair igual ao material que era copiado do extrator/local.
    parcial = preparar_parcial_do_dia(df, incluir_recusas=True)
    if not parcial.empty and "ORDEM_DE_SERVICO" in parcial.columns:
        mapa_contrato = (
            parcial[["ORDEM_DE_SERVICO", "CONTRATO"]]
            .drop_duplicates(subset=["ORDEM_DE_SERVICO"], keep="last")
        )
        mapa_contrato["ORDEM_DE_SERVICO"] = mapa_contrato["ORDEM_DE_SERVICO"].astype(str)
        df["ORDEM_DE_SERVICO"] = df["ORDEM_DE_SERVICO"].astype(str)
        df = df.merge(mapa_contrato, on="ORDEM_DE_SERVICO", how="left")
    else:
        df["CONTRATO"] = ""

    if contrato_filtro and contrato_filtro != "Todos" and "CONTRATO" in df.columns:
        df = df[df["CONTRATO"] == contrato_filtro].copy()

    if grupo_filtro and grupo_filtro != "Todos" and "GRUPO_NOTA" in df.columns:
        grupo_filtro_norm = normalizar_grupo_nota(grupo_filtro)
        grupo_norm = df["GRUPO_NOTA"].apply(normalizar_grupo_nota)
        if grupo_filtro_norm == "CORTE":
            df = df[grupo_norm.isin(["CORTE", "VERIFICACAO"])].copy()
        else:
            df = df[grupo_norm == grupo_filtro_norm].copy()

    df["DATA_HORA_TXT"] = df.apply(_data_hora_txt_supervisao, axis=1)
    df["DATA_TXT_DT"] = pd.to_datetime(df["DATA_HORA_TXT"], dayfirst=True, errors="coerce")

    if data_escolhida:
        data_ref = pd.to_datetime(data_escolhida, dayfirst=True, errors="coerce")
        if pd.notna(data_ref):
            df = df[df["DATA_TXT_DT"].dt.strftime("%d/%m/%Y") == data_ref.strftime("%d/%m/%Y")].copy()

    if df.empty:
        return "", df

    df["STATUS_TXT"] = df.apply(_status_txt_supervisao, axis=1)

    # Ordena por data/hora e recurso para ficar estável para o Excel.
    df = df.sort_values(["DATA_TXT_DT", "RECURSO", "ORDEM_DE_SERVICO"], na_position="last").copy()

    linhas = []
    for _, row in df.iterrows():
        grupo_txt = normalizar_grupo_nota(row.get("GRUPO_NOTA", ""))
        if grupo_txt == "VERIFICACAO":
            grupo_txt = "CORTE"
        campos = [
            _valor_txt_supervisao(row.get("ORDEM_DE_SERVICO", "")),
            grupo_txt,
            _valor_txt_supervisao(row.get("RECURSO", "")).upper(),
            _valor_txt_supervisao(row.get("STATUS_TXT", "")).upper(),
            _valor_txt_supervisao(row.get("DATA_HORA_TXT", "")),
            _valor_txt_supervisao(row.get("ELETRICISTA1", "")),
            _valor_txt_supervisao(row.get("ELETRICISTA2", "")),
            _valor_txt_supervisao(row.get("RECUSA", "")),
        ]
        linhas.append("\t".join(campos))

    return "\n".join(linhas), df


def mostrar_copiador_txt_supervisao(texto_txt):
    """Mostra botão de copiar via navegador e uma área de texto como plano B."""
    texto_js = json.dumps(texto_txt)
    texto_html = html.escape(texto_txt)
    components.html(
        f"""
        <div style="font-family: system-ui, -apple-system, Segoe UI, sans-serif;">
          <button id="btn-copy-gzus" style="
              border: 0;
              border-radius: 12px;
              padding: 10px 16px;
              font-weight: 800;
              cursor: pointer;
              background: #1d4ed8;
              color: white;
              margin-bottom: 8px;">
            📋 Copiar TXT para a área de transferência
          </button>
          <span id="copy-status-gzus" style="margin-left: 10px; font-size: 14px; color: #166534;"></span>
          <textarea id="txt-gzus" style="
              width: 100%;
              height: 260px;
              margin-top: 8px;
              box-sizing: border-box;
              font-family: Consolas, monospace;
              font-size: 12px;
              white-space: pre;
              border: 1px solid #cbd5e1;
              border-radius: 12px;
              padding: 10px;">{texto_html}</textarea>
        </div>
        <script>
          const textoGzus = {texto_js};
          const btn = document.getElementById('btn-copy-gzus');
          const status = document.getElementById('copy-status-gzus');
          const area = document.getElementById('txt-gzus');
          btn.addEventListener('click', async () => {{
            try {{
              await navigator.clipboard.writeText(textoGzus);
              status.textContent = 'Copiado!';
            }} catch (e) {{
              area.focus();
              area.select();
              document.execCommand('copy');
              status.textContent = 'Copiado pelo modo compatível!';
            }}
          }});
        </script>
        """,
        height=340,
    )

# ==============================
# ABA DOWNLOAD
# ==============================

if tela_escolhida == "Downloads":
    st.subheader("TXT do dia")
    st.caption("Gera o TXT tabulado exatamente no formato operacional para colar no Excel dos supervisores.")

    notas_txt = carregar_notas_rapido(meses_escolhidos_resumo)

    if notas_txt.empty:
        st.info("Base de notas não encontrada para gerar o TXT.")
    else:
        base_datas_txt = notas_txt.copy()
        col_data_txt = "DATA_ENCERRAMENTO" if "DATA_ENCERRAMENTO" in base_datas_txt.columns else "DATA"
        datas_txt = pd.to_datetime(base_datas_txt.get(col_data_txt, pd.Series(dtype=str)), dayfirst=True, errors="coerce")
        datas_disponiveis_txt = (
            pd.DataFrame({"DATA_DT": datas_txt})
            .dropna()
            .drop_duplicates()
            .sort_values("DATA_DT", ascending=False)
        )
        datas_disponiveis_txt["DATA"] = datas_disponiveis_txt["DATA_DT"].dt.strftime("%d/%m/%Y")
        opcoes_datas_txt = datas_disponiveis_txt["DATA"].drop_duplicates().tolist()

        if not opcoes_datas_txt:
            st.info("Não encontrei datas na base de notas para gerar o TXT.")
        else:
            c1, c2, c3 = st.columns([1.2, 1.2, 1.2])
            data_txt = c1.selectbox("Dia do TXT", opcoes_datas_txt, index=0, key="download_txt_supervisao_data")

            contratos_txt = ["Todos"]
            try:
                parcial_txt_filtro = preparar_parcial_do_dia(notas_txt, incluir_recusas=True)
                if not parcial_txt_filtro.empty and "CONTRATO" in parcial_txt_filtro.columns:
                    contratos_txt += sorted(parcial_txt_filtro["CONTRATO"].dropna().astype(str).unique().tolist())
            except Exception:
                pass

            contrato_default_idx = 0
            if contrato_filtro_notas in contratos_txt:
                contrato_default_idx = contratos_txt.index(contrato_filtro_notas)
            contrato_txt = c2.selectbox("Contrato", contratos_txt, index=contrato_default_idx, key="download_txt_supervisao_contrato")

            grupos_txt = ["Todos"] + sorted(
                notas_txt.get("GRUPO_NOTA", pd.Series(dtype=str)).dropna().astype(str).str.upper().unique().tolist()
            )
            grupo_txt = c3.selectbox("Tipo", grupos_txt, index=0, key="download_txt_supervisao_grupo")

            texto_txt, df_txt = gerar_txt_supervisao(
                notas_txt,
                data_escolhida=data_txt,
                contrato_filtro=contrato_txt,
                grupo_filtro=grupo_txt,
            )

            if not texto_txt:
                st.warning("Nenhuma linha encontrada para os filtros escolhidos.")
            else:
                st.success(f"TXT gerado com {numero(len(df_txt))} linhas. Use copiar ou baixe o arquivo.")
                mostrar_copiador_txt_supervisao(texto_txt)
                nome_data = data_txt.replace("/", "-")
                st.download_button(
                    "⬇️ Baixar TXT do dia",
                    texto_txt.encode("utf-8"),
                    file_name=f"resultado_final_{nome_data}.txt",
                    mime="text/plain",
                    use_container_width=True,
                )

                with st.expander("Prévia em tabela", expanded=False):
                    colunas_previas = [
                        "ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "STATUS_TXT",
                        "DATA_HORA_TXT", "ELETRICISTA1", "ELETRICISTA2", "RECUSA",
                    ]
                    colunas_previas = [c for c in colunas_previas if c in df_txt.columns]
                    st.dataframe(df_txt[colunas_previas].head(500), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Arquivos carregados")

    banco_atual = caminho_banco_gzus()
    if banco_atual:
        with open(banco_atual, "rb") as f:
            st.download_button("Baixar gzus.db", f, file_name="gzus.db")

    for chave, nome in ARQUIVOS.items():
        caminho = caminho_arquivo(nome)

        if caminho:
            with open(caminho, "rb") as f:
                st.download_button(f"Baixar {caminho.name}", f, file_name=caminho.name)
