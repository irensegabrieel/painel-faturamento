"""
banco_gzus.py — G.Z.U.S. otimizado

Uso dentro da pasta do projeto painel-faturamento:
    python banco_gzus.py importar
    python banco_gzus.py resumo

O que ele faz:
- Mantém dashboard/gzus.db como banco completo/local.
- Cria dashboard/gzus_dashboard.db como banco leve para subir ao GitHub/Streamlit.
- Cria índices e tabelas-resumo para acelerar filtros.
"""
from __future__ import annotations

import argparse
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

PASTA_ATUAL = Path('.')
PASTA_DASHBOARD = Path('dashboard')
PASTA_LEITURA = PASTA_DASHBOARD / 'leitura'
BANCO_COMPLETO = PASTA_DASHBOARD / 'gzus.db'
BANCO_DASHBOARD = PASTA_DASHBOARD / 'gzus_dashboard.db'

ARQUIVOS_CSV = {
    'notas': 'notas_dashboard.csv',
    'faturamento_contratos': 'faturamento_contratos_dashboard.csv',
    'faturamento_dias': 'faturamento_dias_dashboard.csv',
    'faturamento_carro_estimado': 'faturamento_carro_estimado_dashboard.csv',
    'faturamento_carro_dias': 'faturamento_carro_dias_dashboard.csv',
}

PADROES_LEITURA = [
    'Tarefas_Americana*.xlsx',
    'Tarefas_Piracicaba*.xlsx',
    'Parcial_Americana*.xlsx',
    'Parcial_Piracicaba*.xlsx',
    'Resumo_D_por_base_municipio*.xlsx',
]


def agora_iso() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=' ')


def garantir_pastas() -> None:
    PASTA_DASHBOARD.mkdir(exist_ok=True)
    PASTA_LEITURA.mkdir(parents=True, exist_ok=True)


def conectar(caminho: Path | str) -> sqlite3.Connection:
    garantir_pastas()
    conn = sqlite3.connect(str(caminho))
    conn.execute('PRAGMA journal_mode=DELETE')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.execute('PRAGMA temp_store=MEMORY')
    return conn


def caminho_csv(nome: str) -> Optional[Path]:
    candidatos = [
        PASTA_DASHBOARD / nome,
        PASTA_ATUAL / nome,
        PASTA_ATUAL / nome.replace('.csv', '(1).csv'),
    ]
    for c in candidatos:
        if c.exists():
            return c
    achados = list(PASTA_DASHBOARD.glob(nome.replace('.csv', '*.csv'))) + list(PASTA_ATUAL.glob(nome.replace('.csv', '*.csv')))
    return max(achados, key=lambda p: p.stat().st_mtime) if achados else None


def ler_csv_dashboard(caminho: Path) -> pd.DataFrame:
    # Os CSVs do painel normalmente usam ;. Se vier diferente, tenta autodetectar.
    try:
        df = pd.read_csv(caminho, sep=';', encoding='utf-8-sig')
        if len(df.columns) <= 1:
            raise ValueError('CSV parece ter separador diferente')
    except Exception:
        df = pd.read_csv(caminho, sep=None, engine='python', encoding='utf-8-sig')

    # Remove colunas vazias criadas por exportações ocasionais.
    df = df.loc[:, [c for c in df.columns if str(c).strip() and not str(c).startswith('Unnamed')]].copy()

    # Conversões seguras.
    for col in df.columns:
        if 'FATURAMENTO' in str(col).upper():
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if col in ['QTD_NOTAS', 'QTD_EXECUTORES', 'DIA_SEMANA_NUM']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    return df


def salvar_df(conn: sqlite3.Connection, df: pd.DataFrame, tabela: str) -> int:
    if df is None or df.empty:
        try:
            conn.execute(f'DROP TABLE IF EXISTS "{tabela}"')
        except Exception:
            pass
        return 0
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
    df.to_sql(tabela, conn, if_exists='replace', index=False)
    return len(df)


def registrar(conn: sqlite3.Connection, arquivo: Path, tabela: str, linhas: int) -> None:
    conn.execute('''
        CREATE TABLE IF NOT EXISTS controle_importacao (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo TEXT,
            tabela TEXT,
            linhas INTEGER,
            importado_em TEXT,
            tamanho_bytes INTEGER,
            modificado_em REAL
        )
    ''')
    conn.execute('''
        INSERT INTO controle_importacao
        (arquivo, tabela, linhas, importado_em, tamanho_bytes, modificado_em)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (arquivo.name, tabela, int(linhas), agora_iso(), int(arquivo.stat().st_size), float(arquivo.stat().st_mtime)))


def importar_csvs(conn: sqlite3.Connection, incluir_metadados: bool = False) -> dict[str, int]:
    resultado = {}
    for tabela, nome in ARQUIVOS_CSV.items():
        c = caminho_csv(nome)
        if not c:
            resultado[tabela] = 0
            continue
        df = ler_csv_dashboard(c)
        if incluir_metadados:
            df['_origem_arquivo'] = c.name
            df['_importado_em'] = agora_iso()
        linhas = salvar_df(conn, df, tabela)
        registrar(conn, c, tabela, linhas)
        resultado[tabela] = linhas
    return resultado


def tabela_existe(conn: sqlite3.Connection, tabela: str) -> bool:
    return conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tabela,)).fetchone() is not None


def colunas(conn: sqlite3.Connection, tabela: str) -> set[str]:
    if not tabela_existe(conn, tabela):
        return set()
    return {r[1] for r in conn.execute(f'PRAGMA table_info("{tabela}")').fetchall()}


def criar_indices(conn: sqlite3.Connection) -> None:
    specs = [
        ('notas', 'DATA'), ('notas', 'CONTRATO'), ('notas', 'RECURSO'), ('notas', 'GRUPO_NOTA'),
        ('faturamento_contratos', 'CONTRATO'),
        ('faturamento_dias', 'CONTRATO'), ('faturamento_dias', 'DATA'),
        ('faturamento_carro_dias', 'CONTRATO'), ('faturamento_carro_dias', 'DATA'),
    ]
    for tabela, coluna in specs:
        try:
            if tabela_existe(conn, tabela) and coluna in colunas(conn, tabela):
                idx = f'idx_{tabela}_{coluna}'.replace(' ', '_').replace('.', '').replace('-', '_')
                conn.execute(f'CREATE INDEX IF NOT EXISTS "{idx}" ON "{tabela}" ("{coluna}")')
        except Exception:
            pass


def criar_resumos(conn: sqlite3.Connection) -> dict[str, int]:
    """Cria tabelas pequenas para o painel usar depois sem recalcular tudo."""
    out = {}
    if not tabela_existe(conn, 'notas'):
        return out
    cols = colunas(conn, 'notas')

    # resumo_dia
    try:
        if {'DATA', 'CONTRATO'}.issubset(cols):
            grupo = 'GRUPO_NOTA' if 'GRUPO_NOTA' in cols else None
            recurso = 'RECURSO' if 'RECURSO' in cols else None
            ordem = 'ORDEM_DE_SERVICO' if 'ORDEM_DE_SERVICO' in cols else None
            expr_corte = "SUM(CASE WHEN UPPER(COALESCE(GRUPO_NOTA,'')) LIKE '%CORTE%' THEN 1 ELSE 0 END)" if grupo else '0'
            expr_religue = "SUM(CASE WHEN UPPER(COALESCE(GRUPO_NOTA,'')) LIKE '%RELIG%' THEN 1 ELSE 0 END)" if grupo else '0'
            expr_recursos = f'COUNT(DISTINCT "{recurso}")' if recurso else '0'
            expr_notas = f'COUNT(DISTINCT "{ordem}")' if ordem else 'COUNT(*)'
            conn.execute('DROP TABLE IF EXISTS resumo_dia')
            conn.execute(f'''
                CREATE TABLE resumo_dia AS
                SELECT
                    DATA,
                    CONTRATO,
                    {expr_notas} AS TOTAL_NOTAS,
                    {expr_corte} AS CORTES,
                    {expr_religue} AS RELIGUES,
                    {expr_recursos} AS RECURSOS_ATIVOS
                FROM notas
                GROUP BY DATA, CONTRATO
            ''')
            out['resumo_dia'] = conn.execute('SELECT COUNT(*) FROM resumo_dia').fetchone()[0]
    except Exception as e:
        out['erro_resumo_dia'] = str(e)

    # ranking_recursos_dia
    try:
        if {'DATA', 'CONTRATO', 'RECURSO'}.issubset(cols):
            grupo = 'GRUPO_NOTA' if 'GRUPO_NOTA' in cols else None
            ordem = 'ORDEM_DE_SERVICO' if 'ORDEM_DE_SERVICO' in cols else None
            fat = 'FATURAMENTO' if 'FATURAMENTO' in cols else None
            expr_corte = "SUM(CASE WHEN UPPER(COALESCE(GRUPO_NOTA,'')) LIKE '%CORTE%' THEN 1 ELSE 0 END)" if grupo else '0'
            expr_religue = "SUM(CASE WHEN UPPER(COALESCE(GRUPO_NOTA,'')) LIKE '%RELIG%' THEN 1 ELSE 0 END)" if grupo else '0'
            expr_notas = f'COUNT(DISTINCT "{ordem}")' if ordem else 'COUNT(*)'
            expr_fat = f'SUM(COALESCE("{fat}",0))' if fat else '0'
            conn.execute('DROP TABLE IF EXISTS ranking_recursos_dia')
            conn.execute(f'''
                CREATE TABLE ranking_recursos_dia AS
                SELECT
                    DATA,
                    CONTRATO,
                    RECURSO,
                    {expr_notas} AS NOTAS,
                    {expr_corte} AS CORTES,
                    {expr_religue} AS RELIGUES,
                    {expr_fat} AS FATURAMENTO_ATRIBUIDO
                FROM notas
                WHERE COALESCE(RECURSO,'') <> ''
                GROUP BY DATA, CONTRATO, RECURSO
            ''')
            out['ranking_recursos_dia'] = conn.execute('SELECT COUNT(*) FROM ranking_recursos_dia').fetchone()[0]
    except Exception as e:
        out['erro_ranking_recursos_dia'] = str(e)

    # meses disponíveis: tabelinha pequena para inicializar filtros sem varrer tudo no app.
    try:
        if 'DATA' in cols:
            conn.execute('DROP TABLE IF EXISTS meses_notas')
            conn.execute('''
                CREATE TABLE meses_notas AS
                SELECT DISTINCT substr(DATA, 4, 7) AS MES
                FROM notas
                WHERE DATA IS NOT NULL AND length(DATA) >= 10
            ''')
            out['meses_notas'] = conn.execute('SELECT COUNT(*) FROM meses_notas').fetchone()[0]
    except Exception as e:
        out['erro_meses_notas'] = str(e)

    criar_indices(conn)
    for tabela, coluna in [('resumo_dia', 'DATA'), ('resumo_dia', 'CONTRATO'), ('ranking_recursos_dia', 'DATA'), ('ranking_recursos_dia', 'CONTRATO'), ('ranking_recursos_dia', 'RECURSO')]:
        try:
            if tabela_existe(conn, tabela) and coluna in colunas(conn, tabela):
                conn.execute(f'CREATE INDEX IF NOT EXISTS idx_{tabela}_{coluna} ON {tabela} ("{coluna}")')
        except Exception:
            pass
    return out


def localizar_excels_leitura() -> list[Path]:
    achados = {}
    for pasta in [PASTA_LEITURA, PASTA_DASHBOARD, PASTA_ATUAL / 'leitura', PASTA_ATUAL]:
        if not pasta.exists():
            continue
        for padrao in PADROES_LEITURA:
            for c in pasta.glob(padrao):
                if c.is_file() and c.suffix.lower() in ['.xlsx', '.xls']:
                    achados[str(c.resolve())] = c.resolve()
    return sorted(achados.values(), key=lambda p: (p.stat().st_mtime, p.name), reverse=True)


def importar_leitura_completa(conn: sqlite3.Connection, limite: Optional[int] = None) -> dict[str, int]:
    arquivos = localizar_excels_leitura()
    if limite:
        arquivos = arquivos[:limite]
    tarefas, resumos = [], []
    out = {'arquivos_lidos': 0, 'leitura_tarefas': 0, 'leitura_resumos': 0}
    for caminho in arquivos:
        try:
            abas = pd.read_excel(caminho, sheet_name=None)
        except Exception as e:
            print(f'Aviso: não consegui ler {caminho.name}: {e}')
            continue
        out['arquivos_lidos'] += 1
        for aba, df in abas.items():
            if df is None or df.empty:
                continue
            df = df.copy()
            df['_origem_arquivo'] = caminho.name
            df['_aba_excel'] = aba
            cols_norm = {str(c).upper().strip() for c in df.columns}
            if 'TAREFA' in cols_norm or aba.upper().strip() == 'TAREFAS':
                tarefas.append(df)
            elif any(c in cols_norm for c in ['TOTAL TAREFAS', 'TOTAL TAREFA', 'FEITA', 'PENDENTE', 'PARCIAL']):
                resumos.append(df)
    if tarefas:
        out['leitura_tarefas'] = salvar_df(conn, pd.concat(tarefas, ignore_index=True), 'leitura_tarefas')
    if resumos:
        out['leitura_resumos'] = salvar_df(conn, pd.concat(resumos, ignore_index=True), 'leitura_resumos')
    return out


def otimizar(conn: sqlite3.Connection) -> None:
    conn.commit()
    try:
        conn.execute('VACUUM')
    except Exception:
        pass


def importar_tudo(limite_excels: Optional[int] = None) -> dict:
    garantir_pastas()
    resultado = {'completo': {}, 'dashboard': {}}

    # Banco completo local: pode guardar leitura bruta também.
    with conectar(BANCO_COMPLETO) as conn:
        resultado['completo']['csvs'] = importar_csvs(conn, incluir_metadados=True)
        resultado['completo']['resumos'] = criar_resumos(conn)
        resultado['completo']['leitura'] = importar_leitura_completa(conn, limite=limite_excels)
        conn.commit()
    with conectar(BANCO_COMPLETO) as conn:
        otimizar(conn)

    # Banco leve do Streamlit: só o que o app precisa rápido.
    with conectar(BANCO_DASHBOARD) as conn:
        resultado['dashboard']['csvs'] = importar_csvs(conn, incluir_metadados=False)
        resultado['dashboard']['resumos'] = criar_resumos(conn)
        conn.commit()
    with conectar(BANCO_DASHBOARD) as conn:
        otimizar(conn)

    resultado['arquivos'] = {
        'gzus_db': str(BANCO_COMPLETO.resolve()),
        'gzus_dashboard_db': str(BANCO_DASHBOARD.resolve()),
        'tamanho_gzus_kb': round(BANCO_COMPLETO.stat().st_size / 1024, 1) if BANCO_COMPLETO.exists() else 0,
        'tamanho_dashboard_kb': round(BANCO_DASHBOARD.stat().st_size / 1024, 1) if BANCO_DASHBOARD.exists() else 0,
    }
    return resultado


def resumo_banco(caminho: Path | str) -> pd.DataFrame:
    caminho = Path(caminho)
    if not caminho.exists():
        return pd.DataFrame()
    with sqlite3.connect(str(caminho)) as conn:
        tabelas = pd.read_sql_query("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name", conn)
        linhas = []
        for t in tabelas['name'].tolist():
            try:
                qtd = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
            except Exception:
                qtd = None
            linhas.append({'banco': caminho.name, 'tabela': t, 'linhas': qtd})
        return pd.DataFrame(linhas)


def main() -> None:
    parser = argparse.ArgumentParser(description='Banco SQLite do sistema G.Z.U.S.')
    parser.add_argument('acao', choices=['importar', 'resumo'], help='O que deseja fazer')
    parser.add_argument('--limite-excels', type=int, default=None, help='Limita quantos Excels de leitura importar no banco completo')
    args = parser.parse_args()

    if args.acao == 'importar':
        r = importar_tudo(limite_excels=args.limite_excels)
        print(json.dumps(r, ensure_ascii=False, indent=2))
        print('\nImportação finalizada.')
        print(f'Banco completo: {BANCO_COMPLETO.resolve()}')
        print(f'Banco leve do dashboard: {BANCO_DASHBOARD.resolve()}')
    else:
        dfs = [resumo_banco(BANCO_COMPLETO), resumo_banco(BANCO_DASHBOARD)]
        df = pd.concat([d for d in dfs if not d.empty], ignore_index=True) if any(not d.empty for d in dfs) else pd.DataFrame()
        if df.empty:
            print('Bancos ainda não criados.')
        else:
            print(df.to_string(index=False))


if __name__ == '__main__':
    main()
