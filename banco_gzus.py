"""
banco_gzus.py
================
SQLite V2 para o sistema G.Z.U.S. com tabela processada e índices de performance.

Objetivo:
- Manter o sistema atual funcionando com CSV/Excel.
- Criar uma copia organizada em SQLite: gzus.db.
- Permitir que o dashboard passe a ler do banco aos poucos.

Como usar no terminal, dentro da pasta do projeto:
    python banco_gzus.py importar
    python banco_gzus.py resumo

Requisitos:
    pip install pandas openpyxl

Observacao importante:
Este arquivo NAO apaga seus CSV/Excel. Ele apenas le e grava uma copia no banco.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd


# ==============================
# CONFIGURACAO BASICA
# ==============================

PASTA_ATUAL = Path(".")
PASTA_DASHBOARD = Path("dashboard")
PASTA_LEITURA = PASTA_DASHBOARD / "leitura"
BANCO_PADRAO = PASTA_DASHBOARD / "gzus.db"

# Mesmos arquivos principais usados no dashboard atual.
ARQUIVOS_CSV_DASHBOARD = {
    "notas": "notas_dashboard.csv",
    "faturamento_contratos": "faturamento_contratos_dashboard.csv",
    "faturamento_dias": "faturamento_dias_dashboard.csv",
    "faturamento_carro_estimado": "faturamento_carro_estimado_dashboard.csv",
    "faturamento_carro_dias": "faturamento_carro_dias_dashboard.csv",
}

# Arquivos de leitura gerados pelo extrator CWSI.
PADROES_EXCEL_LEITURA = [
    "Tarefas_Americana*.xlsx",
    "Tarefas_Piracicaba*.xlsx",
    "Parcial_Americana*.xlsx",
    "Parcial_Piracicaba*.xlsx",
    "Resumo_D_por_base_municipio*.xlsx",
]


# ==============================
# FUNCOES AUXILIARES
# ==============================

def agora_iso() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def garantir_pastas() -> None:
    PASTA_DASHBOARD.mkdir(exist_ok=True)
    PASTA_LEITURA.mkdir(parents=True, exist_ok=True)


def conectar(caminho_banco: Path | str = BANCO_PADRAO) -> sqlite3.Connection:
    garantir_pastas()
    conn = sqlite3.connect(str(caminho_banco))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    return conn


def limpar_nome_tabela(nome: str) -> str:
    permitido = []
    for ch in str(nome).lower().strip():
        if ch.isalnum():
            permitido.append(ch)
        elif ch in [" ", "-", ".", "/", "\\"]:
            permitido.append("_")
    saida = "".join(permitido).strip("_")
    while "__" in saida:
        saida = saida.replace("__", "_")
    return saida or "tabela"


def caminho_csv(nome_arquivo: str) -> Optional[Path]:
    candidatos = [
        PASTA_DASHBOARD / nome_arquivo,
        PASTA_ATUAL / nome_arquivo,
        PASTA_ATUAL / nome_arquivo.replace(".csv", "(1).csv"),
    ]
    for c in candidatos:
        if c.exists():
            return c

    achados = list(PASTA_DASHBOARD.glob(nome_arquivo.replace(".csv", "*.csv")))
    achados += list(PASTA_ATUAL.glob(nome_arquivo.replace(".csv", "*.csv")))
    return max(achados, key=lambda p: p.stat().st_mtime) if achados else None


def detectar_base_por_nome(caminho: Path) -> str:
    nome = caminho.name.upper()
    if "AMERICANA" in nome:
        return "AMERICANA"
    if "PIRACICABA" in nome:
        return "PIRACICABA"
    return ""


def detectar_tipo_leitura(caminho: Path) -> str:
    nome = caminho.name.upper()
    if nome.startswith("TAREFAS_"):
        return "TAREFAS"
    if nome.startswith("PARCIAL_"):
        return "PARCIAL"
    if nome.startswith("RESUMO_D"):
        return "RESUMO_D"
    return "OUTRO"


def adicionar_metadados(df: pd.DataFrame, caminho: Path, origem: str) -> pd.DataFrame:
    df = df.copy()
    df["_origem_arquivo"] = caminho.name
    df["_origem_caminho"] = str(caminho)
    df["_origem_tipo"] = origem
    df["_importado_em"] = agora_iso()
    return df


def salvar_dataframe(conn: sqlite3.Connection, df: pd.DataFrame, tabela: str, modo: str = "replace") -> int:
    """Salva um DataFrame no SQLite e retorna a quantidade de linhas."""
    tabela = limpar_nome_tabela(tabela)
    if df is None or df.empty:
        return 0

    # SQLite aceita tipos simples melhor. Datas viram texto ISO quando possivel.
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")

    df.to_sql(tabela, conn, if_exists=modo, index=False)
    return len(df)


def criar_tabela_controle(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS controle_importacao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo TEXT NOT NULL,
            caminho TEXT,
            tabela TEXT,
            linhas INTEGER,
            importado_em TEXT NOT NULL,
            tamanho_bytes INTEGER,
            modificado_em REAL
        )
        """
    )
    conn.commit()


def registrar_importacao(conn: sqlite3.Connection, caminho: Path, tabela: str, linhas: int) -> None:
    conn.execute(
        """
        INSERT INTO controle_importacao
        (arquivo, caminho, tabela, linhas, importado_em, tamanho_bytes, modificado_em)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            caminho.name,
            str(caminho),
            limpar_nome_tabela(tabela),
            int(linhas),
            agora_iso(),
            int(caminho.stat().st_size) if caminho.exists() else 0,
            float(caminho.stat().st_mtime) if caminho.exists() else 0,
        ),
    )
    conn.commit()


# ==============================
# IMPORTADORES
# ==============================

def importar_csvs_dashboard(conn: sqlite3.Connection) -> dict[str, int]:
    """Importa os CSVs principais do dashboard para tabelas SQLite."""
    resultado: dict[str, int] = {}

    for tabela, nome_arquivo in ARQUIVOS_CSV_DASHBOARD.items():
        caminho = caminho_csv(nome_arquivo)
        if not caminho:
            resultado[tabela] = 0
            continue

        # Tenta primeiro o separador usado pelos CSVs do painel.
        # Se vier errado, tenta autodetectar.
        try:
            df = pd.read_csv(caminho, sep=";", encoding="utf-8-sig")
            if len(df.columns) <= 1:
                df = pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")

        df = adicionar_metadados(df, caminho, origem="CSV_DASHBOARD")
        linhas = salvar_dataframe(conn, df, tabela, modo="replace")
        registrar_importacao(conn, caminho, tabela, linhas)
        resultado[tabela] = linhas

    return resultado


def localizar_excels_leitura(pastas: Optional[Iterable[Path]] = None) -> list[Path]:
    pastas_busca = list(pastas or [PASTA_LEITURA, PASTA_DASHBOARD, PASTA_ATUAL / "leitura", PASTA_ATUAL])
    achados: dict[str, Path] = {}

    for pasta in pastas_busca:
        if not pasta.exists():
            continue
        for padrao in PADROES_EXCEL_LEITURA:
            for caminho in pasta.glob(padrao):
                if caminho.is_file() and caminho.suffix.lower() in [".xlsx", ".xls"]:
                    achados[str(caminho.resolve())] = caminho.resolve()

    return sorted(achados.values(), key=lambda p: (p.stat().st_mtime, p.name), reverse=True)


def importar_excels_leitura(conn: sqlite3.Connection, limite_arquivos: Optional[int] = None) -> dict[str, int]:
    """Importa as abas de leitura dos Excels gerados pelo extrator."""
    arquivos = localizar_excels_leitura()
    if limite_arquivos:
        arquivos = arquivos[:limite_arquivos]

    todas_tarefas = []
    todos_resumos = []
    resultado = {"leitura_tarefas": 0, "leitura_resumos": 0, "arquivos_lidos": 0}

    for caminho in arquivos:
        tipo = detectar_tipo_leitura(caminho)
        base = detectar_base_por_nome(caminho)

        try:
            abas = pd.read_excel(caminho, sheet_name=None)
        except Exception as e:
            print(f"Aviso: não consegui ler {caminho.name}: {e}")
            continue

        resultado["arquivos_lidos"] += 1

        for nome_aba, df in abas.items():
            if df is None or df.empty:
                continue

            df = adicionar_metadados(df, caminho, origem=f"EXCEL_LEITURA_{tipo}")
            df["_base_arquivo"] = base
            df["_aba_excel"] = nome_aba

            nome_aba_norm = nome_aba.upper().strip()
            colunas_norm = {str(c).upper().strip() for c in df.columns}

            # Abas detalhadas costumam ter TAREFA. Resumos costumam ter TOTAL TAREFAS ou FEITA/PENDENTE.
            if "TAREFA" in colunas_norm or nome_aba_norm == "TAREFAS":
                todas_tarefas.append(df)
            elif any(c in colunas_norm for c in ["TOTAL TAREFAS", "FEITA", "PENDENTE", "PARCIAL"]):
                todos_resumos.append(df)

    if todas_tarefas:
        df_tarefas = pd.concat(todas_tarefas, ignore_index=True)
        linhas = salvar_dataframe(conn, df_tarefas, "leitura_tarefas", modo="replace")
        resultado["leitura_tarefas"] = linhas

    if todos_resumos:
        df_resumos = pd.concat(todos_resumos, ignore_index=True)
        linhas = salvar_dataframe(conn, df_resumos, "leitura_resumos", modo="replace")
        resultado["leitura_resumos"] = linhas

    return resultado



# ==============================
# PERFORMANCE V2 - NOTAS PROCESSADAS
# ==============================

def _txt(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def _numero(valor, padrao=0.0) -> float:
    try:
        if pd.isna(valor):
            return float(padrao)
        if isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".") if "," in valor else valor
        return float(valor)
    except Exception:
        return float(padrao)


def _eh_disjuntor_jundiai(recurso) -> bool:
    recurso_norm = _txt(recurso).upper()
    return recurso_norm.startswith("JUN55") or recurso_norm.startswith("JUN59") or recurso_norm.startswith("SAL55")


def _eh_disjuntor_santa_cruz(recurso) -> bool:
    import re
    recurso_norm = _txt(recurso).upper()
    m = re.search(r"(\d+)", recurso_norm)
    if not m:
        return False
    primeiros_numeros = m.group(1)
    return primeiros_numeros.startswith("89") or primeiros_numeros.startswith("20")


def criar_notas_processadas(conn: sqlite3.Connection) -> int:
    try:
        df = pd.read_sql_query('SELECT * FROM "notas"', conn)
    except Exception:
        return 0

    if df.empty:
        return 0
    if len(df.columns) <= 1 or any(";" in str(c) for c in df.columns):
        return 0

    for col in ["ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "RECUSA", "ELETRICISTA1", "ELETRICISTA2", "DATA"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].apply(_txt).astype(str).str.strip()

    if "QTD_EXECUTORES" not in df.columns:
        df["QTD_EXECUTORES"] = ((df["ELETRICISTA1"] != "").astype(int) + (df["ELETRICISTA2"] != "").astype(int))
    else:
        df["QTD_EXECUTORES"] = pd.to_numeric(df["QTD_EXECUTORES"], errors="coerce").fillna(0).astype(int)

    df["GRUPO_NOTA"] = df["GRUPO_NOTA"].astype(str).str.upper().str.strip()
    df["RECURSO"] = df["RECURSO"].astype(str).str.upper().str.strip()
    df["RECUSA"] = df["RECUSA"].fillna("").astype(str).str.strip()
    df["EH_RECUSA"] = (df["RECUSA"] != "").astype(int)

    tarifas = {
        "JUNDIAI_CORTE": 13.72,
        "JUNDIAI_RELIGUE": 27.43,
        "SANTA_CORTE": 11.98,
        "SANTA_RELIGUE": 23.97,
        "STC_CORTE_MIN": 38.18,
        "STC_RELIGUE_MIN": 36.36,
        "STC_CORTE_MAX": 45.45,
        "STC_RELIGUE_MAX": 50.91,
    }

    contratos, fats, fats_min, fats_max, eh_corte, eh_religue, manter = [], [], [], [], [], [], []
    for _, row in df.iterrows():
        recurso = row.get("RECURSO", "")
        grupo = row.get("GRUPO_NOTA", "")
        qtd_exec = int(_numero(row.get("QTD_EXECUTORES", 0), 0))
        recusa = _txt(row.get("RECUSA", "")) != ""
        contrato = ""
        fat = fat_min = fat_max = 0.0
        if _eh_disjuntor_jundiai(recurso):
            contrato = "Disjuntor Jundiaí"
            if not recusa:
                fat = {"CORTE": tarifas["JUNDIAI_CORTE"], "RELIGUE": tarifas["JUNDIAI_RELIGUE"]}.get(grupo, 0.0)
                fat_min = fat_max = fat
        elif _eh_disjuntor_santa_cruz(recurso):
            contrato = "Disjuntor Santa Cruz"
            if not recusa:
                fat = {"CORTE": tarifas["SANTA_CORTE"], "RELIGUE": tarifas["SANTA_RELIGUE"]}.get(grupo, 0.0)
                fat_min = fat_max = fat
        elif str(recurso).startswith("JUN58") and qtd_exec >= 2:
            contrato = "STC Jundiai"
            if not recusa:
                fat_min = {"CORTE": tarifas["STC_CORTE_MIN"], "RELIGUE": tarifas["STC_RELIGUE_MIN"]}.get(grupo, 0.0)
                fat_max = {"CORTE": tarifas["STC_CORTE_MAX"], "RELIGUE": tarifas["STC_RELIGUE_MAX"]}.get(grupo, 0.0)
                fat = fat_min
        manter.append(bool(contrato))
        contratos.append(contrato); fats.append(fat); fats_min.append(fat_min); fats_max.append(fat_max)
        eh_corte.append(1 if (grupo == "CORTE" and not recusa) else 0)
        eh_religue.append(1 if (grupo == "RELIGUE" and not recusa) else 0)

    out = df.loc[manter].copy()
    if out.empty:
        return 0
    out["CONTRATO"] = [c for c, m in zip(contratos, manter) if m]
    out["FATURAMENTO"] = [v for v, m in zip(fats, manter) if m]
    out["FATURAMENTO_MIN"] = [v for v, m in zip(fats_min, manter) if m]
    out["FATURAMENTO_MAX"] = [v for v, m in zip(fats_max, manter) if m]
    out["EH_CORTE"] = [v for v, m in zip(eh_corte, manter) if m]
    out["EH_RELIGUE"] = [v for v, m in zip(eh_religue, manter) if m]
    out["DATA_DT"] = pd.to_datetime(out["DATA"], dayfirst=True, errors="coerce")
    out = out.dropna(subset=["DATA_DT"]).copy()
    out["DATA"] = out["DATA_DT"].dt.strftime("%d/%m/%Y")
    out["DATA_DT"] = out["DATA_DT"].dt.strftime("%Y-%m-%d")
    out["MES"] = pd.to_datetime(out["DATA_DT"], errors="coerce").dt.strftime("%m/%Y")
    return salvar_dataframe(conn, out, "notas_processadas", modo="replace")


def criar_resumos_performance(conn: sqlite3.Connection) -> dict[str, int]:
    resultado = {"notas_processadas": criar_notas_processadas(conn)}
    try:
        conn.execute('DROP TABLE IF EXISTS resumo_dia_contrato')
        conn.execute("""
            CREATE TABLE resumo_dia_contrato AS
            SELECT DATA, DATA_DT, CONTRATO,
                   COUNT(DISTINCT CASE WHEN EH_RECUSA = 0 THEN ORDEM_DE_SERVICO END) AS TOTAL_NOTAS,
                   SUM(EH_CORTE) AS CORTES,
                   SUM(EH_RELIGUE) AS RELIGUES,
                   SUM(CASE WHEN EH_RECUSA = 1 THEN 1 ELSE 0 END) AS RECUSAS,
                   COUNT(DISTINCT RECURSO) AS RECURSOS_ATIVOS,
                   SUM(FATURAMENTO) AS FATURAMENTO,
                   SUM(FATURAMENTO_MIN) AS FATURAMENTO_MIN,
                   SUM(FATURAMENTO_MAX) AS FATURAMENTO_MAX
            FROM notas_processadas
            GROUP BY DATA, DATA_DT, CONTRATO
        """)
        resultado["resumo_dia_contrato"] = conn.execute('SELECT COUNT(*) FROM resumo_dia_contrato').fetchone()[0]
    except Exception:
        resultado["resumo_dia_contrato"] = 0
    try:
        conn.execute('DROP TABLE IF EXISTS resumo_dia_recurso')
        conn.execute("""
            CREATE TABLE resumo_dia_recurso AS
            SELECT DATA, DATA_DT, CONTRATO, RECURSO,
                   COUNT(DISTINCT CASE WHEN EH_RECUSA = 0 THEN ORDEM_DE_SERVICO END) AS TOTAL_NOTAS,
                   SUM(EH_CORTE) AS CORTES,
                   SUM(EH_RELIGUE) AS RELIGUES,
                   SUM(CASE WHEN EH_RECUSA = 1 THEN 1 ELSE 0 END) AS RECUSAS,
                   SUM(FATURAMENTO) AS FATURAMENTO,
                   SUM(FATURAMENTO_MIN) AS FATURAMENTO_MIN,
                   SUM(FATURAMENTO_MAX) AS FATURAMENTO_MAX
            FROM notas_processadas
            GROUP BY DATA, DATA_DT, CONTRATO, RECURSO
        """)
        resultado["resumo_dia_recurso"] = conn.execute('SELECT COUNT(*) FROM resumo_dia_recurso').fetchone()[0]
    except Exception:
        resultado["resumo_dia_recurso"] = 0
    conn.commit()
    return resultado


def criar_indices(conn: sqlite3.Connection) -> None:
    """Cria indices simples para acelerar filtros comuns. Ignora se a coluna nao existir."""
    indices = [
        ("notas", "CONTRATO"),
        ("notas", "RECURSO"),
        ("notas_processadas", "DATA"),
        ("notas_processadas", "DATA_DT"),
        ("notas_processadas", "CONTRATO"),
        ("notas_processadas", "RECURSO"),
        ("notas_processadas", "MES"),
        ("resumo_dia_contrato", "DATA"),
        ("resumo_dia_contrato", "CONTRATO"),
        ("resumo_dia_recurso", "DATA"),
        ("resumo_dia_recurso", "CONTRATO"),
        ("resumo_dia_recurso", "RECURSO"),
        ("faturamento_contratos", "CONTRATO"),
        ("faturamento_dias", "CONTRATO"),
        ("leitura_tarefas", "BASE"),
        ("leitura_tarefas", "MUNICÍPIO"),
        ("leitura_tarefas", "D OPERACIONAL"),
        ("leitura_tarefas", "STATUS OPERACIONAL"),
        ("leitura_tarefas", "_base_arquivo"),
    ]

    for tabela, coluna in indices:
        try:
            nome_indice = limpar_nome_tabela(f"idx_{tabela}_{coluna}")
            conn.execute(f'CREATE INDEX IF NOT EXISTS "{nome_indice}" ON "{tabela}" ("{coluna}")')
        except Exception:
            pass
    conn.commit()


def importar_tudo(caminho_banco: Path | str = BANCO_PADRAO, limite_excels: Optional[int] = None) -> dict:
    conn = conectar(caminho_banco)
    criar_tabela_controle(conn)

    resultado = {
        "banco": str(caminho_banco),
        "csvs_dashboard": importar_csvs_dashboard(conn),
        "excels_leitura": importar_excels_leitura(conn, limite_arquivos=limite_excels),
    }
    resultado["performance_v2"] = criar_resumos_performance(conn)
    criar_indices(conn)
    conn.close()
    return resultado


# ==============================
# LEITURA PARA O DASHBOARD
# ==============================

def tabela_existe(conn: sqlite3.Connection, tabela: str) -> bool:
    cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (limpar_nome_tabela(tabela),))
    return cur.fetchone() is not None


def ler_tabela(tabela: str, caminho_banco: Path | str = BANCO_PADRAO, limite: Optional[int] = None) -> pd.DataFrame:
    conn = conectar(caminho_banco)
    tabela_limpa = limpar_nome_tabela(tabela)
    if not tabela_existe(conn, tabela_limpa):
        conn.close()
        return pd.DataFrame()
    sql = f'SELECT * FROM "{tabela_limpa}"'
    if limite:
        sql += f" LIMIT {int(limite)}"
    df = pd.read_sql_query(sql, conn)
    conn.close()
    return df


def consulta_sql(sql: str, caminho_banco: Path | str = BANCO_PADRAO, params: Optional[tuple] = None) -> pd.DataFrame:
    conn = conectar(caminho_banco)
    df = pd.read_sql_query(sql, conn, params=params or ())
    conn.close()
    return df


def resumo_banco(caminho_banco: Path | str = BANCO_PADRAO) -> pd.DataFrame:
    conn = conectar(caminho_banco)
    tabelas = pd.read_sql_query(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name",
        conn,
    )
    linhas = []
    for tabela in tabelas["name"].tolist():
        try:
            qtd = conn.execute(f'SELECT COUNT(*) FROM "{tabela}"').fetchone()[0]
        except Exception:
            qtd = None
        linhas.append({"tabela": tabela, "linhas": qtd})
    conn.close()
    return pd.DataFrame(linhas)


# ==============================
# LINHA DE COMANDO
# ==============================

def main() -> None:
    parser = argparse.ArgumentParser(description="Banco SQLite do sistema G.Z.U.S.")
    parser.add_argument("acao", choices=["importar", "resumo"], help="O que deseja fazer")
    parser.add_argument("--banco", default=str(BANCO_PADRAO), help="Caminho do arquivo .db")
    parser.add_argument("--limite-excels", type=int, default=None, help="Limita quantos Excels de leitura importar")
    args = parser.parse_args()

    caminho_banco = Path(args.banco)

    if args.acao == "importar":
        resultado = importar_tudo(caminho_banco, limite_excels=args.limite_excels)
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
        print("\nImportação finalizada.")
        print(f"Banco criado/atualizado em: {caminho_banco.resolve()}")

    elif args.acao == "resumo":
        df = resumo_banco(caminho_banco)
        if df.empty:
            print("Banco vazio ou ainda não criado.")
        else:
            print(df.to_string(index=False))


if __name__ == "__main__":
    main()
