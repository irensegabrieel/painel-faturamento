from __future__ import annotations

import os

import argparse
import json
import re
import unicodedata
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
    try:
        df = pd.read_csv(caminho, sep=';', encoding='utf-8-sig')
        if len(df.columns) <= 1:
            raise ValueError('CSV parece ter separador diferente')
    except Exception:
        df = pd.read_csv(caminho, sep=None, engine='python', encoding='utf-8-sig')

    df = df.loc[:, [c for c in df.columns if str(c).strip() and not str(c).startswith('Unnamed')]].copy()

    for col in df.columns:
        col_upper = str(col).upper()
        if 'FATURAMENTO' in col_upper:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if col in ['QTD_NOTAS', 'QTD_EXECUTORES', 'DIA_SEMANA_NUM']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    return df


def salvar_df(conn: sqlite3.Connection, df: pd.DataFrame, tabela: str) -> int:
    if df is None or df.empty:
        conn.execute(f'DROP TABLE IF EXISTS "{tabela}"')
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
            caminho TEXT,
            tabela TEXT,
            linhas INTEGER,
            importado_em TEXT,
            tamanho_bytes INTEGER,
            modificado_em REAL
        )
    ''')
    conn.execute('''
        INSERT INTO controle_importacao
        (arquivo, caminho, tabela, linhas, importado_em, tamanho_bytes, modificado_em)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (arquivo.name, str(arquivo), tabela, int(linhas), agora_iso(), int(arquivo.stat().st_size), float(arquivo.stat().st_mtime)))


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
    return conn.execute('SELECT 1 FROM sqlite_master WHERE type=\'table\' AND name=?', (tabela,)).fetchone() is not None


def colunas(conn: sqlite3.Connection, tabela: str) -> set[str]:
    if not tabela_existe(conn, tabela):
        return set()
    return {r[1] for r in conn.execute(f'PRAGMA table_info("{tabela}")').fetchall()}



def normalizar_grupo_nota(valor: str) -> str:
    texto = str(valor or '').strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    if 'VERIFIC' in texto:
        return 'VERIFICACAO'
    if 'RELIG' in texto:
        return 'RELIGUE'
    if 'CORTE' in texto:
        return 'CORTE'
    return texto


def tarifa_float(nome: str, padrao: float) -> float:
    try:
        return float(os.getenv(nome, padrao))
    except Exception:
        return float(padrao)

def _eh_disjuntor_jundiai(recurso: str) -> bool:
    r = str(recurso or '').strip().upper()
    return r.startswith('JUN55') or r.startswith('JUN59') or r.startswith('SAL55')


def _eh_disjuntor_santa_cruz(recurso: str) -> bool:
    r = str(recurso or '').strip().upper()
    m = re.search(r'(\d+)', r)
    if not m:
        return False
    return m.group(1).startswith('89') or m.group(1).startswith('20')


def _derivar_contrato(recurso: str, qtd_exec: int) -> str:
    r = str(recurso or '').strip().upper()
    try:
        qtd = int(qtd_exec or 0)
    except Exception:
        qtd = 0
    if _eh_disjuntor_jundiai(r):
        return 'Disjuntor Jundiaí'
    if _eh_disjuntor_santa_cruz(r):
        return 'Disjuntor Santa Cruz'
    if r.startswith('JUN58') and qtd >= 2:
        return 'STC Jundiai'
    return ''


def criar_notas_processadas(conn: sqlite3.Connection) -> dict[str, int | str]:
    out: dict[str, int | str] = {}
    if not tabela_existe(conn, 'notas'):
        return out

    df = pd.read_sql_query('SELECT * FROM "notas"', conn)
    if df.empty:
        salvar_df(conn, df, 'notas_processadas')
        out['notas_processadas'] = 0
        return out

    for col in ['ORDEM_DE_SERVICO', 'GRUPO_NOTA', 'RECURSO', 'RECUSA', 'ELETRICISTA1', 'ELETRICISTA2', 'DATA']:
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].fillna('').astype(str).str.strip()

    if 'QTD_EXECUTORES' not in df.columns:
        df['QTD_EXECUTORES'] = ((df['ELETRICISTA1'] != '').astype(int) + (df['ELETRICISTA2'] != '').astype(int))
    else:
        df['QTD_EXECUTORES'] = pd.to_numeric(df['QTD_EXECUTORES'], errors='coerce').fillna(0).astype(int)

    df['GRUPO_NOTA'] = df['GRUPO_NOTA'].apply(normalizar_grupo_nota)
    df['RECURSO'] = df['RECURSO'].astype(str).str.upper().str.strip()
    df['RECUSA'] = df['RECUSA'].fillna('').astype(str).str.strip()

    df['CONTRATO_DERIVADO'] = [_derivar_contrato(r, q) for r, q in zip(df['RECURSO'], df['QTD_EXECUTORES'])]
    df['CONTRATO'] = df['CONTRATO_DERIVADO']

    df['EH_RECUSA'] = (df['RECUSA'] != '').astype(int)
    df['EH_VERIFICACAO'] = ((df['GRUPO_NOTA'] == 'VERIFICACAO') & (df['CONTRATO'] == 'Disjuntor Santa Cruz') & (df['EH_RECUSA'] == 0)).astype(int)
    df['EH_CORTE'] = (((df['GRUPO_NOTA'] == 'CORTE') | ((df['GRUPO_NOTA'] == 'VERIFICACAO') & (df['CONTRATO'] == 'Disjuntor Jundiaí'))) & (df['EH_RECUSA'] == 0) & (df['CONTRATO'] != '')).astype(int)
    df['EH_RELIGUE'] = ((df['GRUPO_NOTA'].str.contains('RELIG', na=False)) & (df['EH_RECUSA'] == 0) & (df['CONTRATO'] != '')).astype(int)

    tjc = tarifa_float('TARIFA_DISJUNTOR_JUNDIAI_CORTE', 13.72)
    tjr = tarifa_float('TARIFA_DISJUNTOR_JUNDIAI_RELIGUE', 27.43)
    tsc = tarifa_float('TARIFA_DISJUNTOR_SANTA_CRUZ_CORTE', 11.98)
    tsr = tarifa_float('TARIFA_DISJUNTOR_SANTA_CRUZ_RELIGUE', 23.97)
    tsv = tarifa_float('TARIFA_DISJUNTOR_SANTA_CRUZ_VERIFICACAO', 23.97)

    def _faturamento_linha(row):
        if int(row.get('EH_RECUSA', 0) or 0) == 1:
            return 0.0
        contrato = str(row.get('CONTRATO', ''))
        grupo = str(row.get('GRUPO_NOTA', ''))
        if contrato == 'Disjuntor Jundiaí':
            return {'CORTE': tjc, 'VERIFICACAO': tjc, 'RELIGUE': tjr}.get(grupo, 0.0)
        if contrato == 'Disjuntor Santa Cruz':
            return {'CORTE': tsc, 'VERIFICACAO': tsv, 'RELIGUE': tsr}.get(grupo, 0.0)
        return 0.0

    df['FATURAMENTO'] = df.apply(_faturamento_linha, axis=1)
    df['EH_NOTA_VALIDA'] = ((df['EH_RECUSA'] == 0) & (df['CONTRATO'] != '')).astype(int)

    data_dt = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')
    df['DATA_SQL'] = data_dt.dt.strftime('%Y-%m-%d')
    df['DATA'] = data_dt.dt.strftime('%d/%m/%Y').fillna(df['DATA'])

    linhas = salvar_df(conn, df, 'notas_processadas')
    out['notas_processadas'] = linhas

    try:
        max_data = pd.to_datetime(df['DATA_SQL'], errors='coerce').max()
        out['ultima_data_notas_processadas'] = max_data.strftime('%d/%m/%Y') if pd.notna(max_data) else ''
    except Exception:
        pass

    return out


def criar_indices(conn: sqlite3.Connection) -> None:
    specs = [
        ('notas', 'DATA'), ('notas', 'RECURSO'), ('notas', 'GRUPO_NOTA'),
        ('notas_processadas', 'DATA'), ('notas_processadas', 'DATA_SQL'),
        ('notas_processadas', 'CONTRATO'), ('notas_processadas', 'RECURSO'),
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


def criar_resumos(conn: sqlite3.Connection) -> dict[str, int | str]:
    out: dict[str, int | str] = {}
    out.update(criar_notas_processadas(conn))

    if not tabela_existe(conn, 'notas_processadas'):
        return out

    try:
        conn.execute('DROP TABLE IF EXISTS resumo_dia')
        conn.execute('''
            CREATE TABLE resumo_dia AS
            SELECT
                DATA,
                DATA_SQL,
                CONTRATO,
                COUNT(DISTINCT CASE WHEN EH_NOTA_VALIDA = 1 THEN ORDEM_DE_SERVICO END) AS TOTAL_NOTAS,
                SUM(EH_CORTE) AS CORTES,
                SUM(EH_RELIGUE) AS RELIGUES,
                SUM(EH_VERIFICACAO) AS VERIFICACOES,
                SUM(EH_RECUSA) AS RECUSAS,
                COUNT(DISTINCT CASE WHEN CONTRATO <> '' THEN RECURSO END) AS RECURSOS_ATIVOS
            FROM notas_processadas
            WHERE COALESCE(CONTRATO,'') <> ''
            GROUP BY DATA, DATA_SQL, CONTRATO
        ''')
        out['resumo_dia'] = conn.execute('SELECT COUNT(*) FROM resumo_dia').fetchone()[0]
    except Exception as e:
        out['erro_resumo_dia'] = str(e)

    try:
        conn.execute('DROP TABLE IF EXISTS resumo_dia_total')
        conn.execute('''
            CREATE TABLE resumo_dia_total AS
            SELECT
                DATA,
                DATA_SQL,
                'Todos' AS CONTRATO,
                SUM(TOTAL_NOTAS) AS TOTAL_NOTAS,
                SUM(CORTES) AS CORTES,
                SUM(RELIGUES) AS RELIGUES,
                SUM(VERIFICACOES) AS VERIFICACOES,
                SUM(RECUSAS) AS RECUSAS,
                SUM(RECURSOS_ATIVOS) AS RECURSOS_ATIVOS
            FROM resumo_dia
            GROUP BY DATA, DATA_SQL
        ''')
        out['resumo_dia_total'] = conn.execute('SELECT COUNT(*) FROM resumo_dia_total').fetchone()[0]
    except Exception as e:
        out['erro_resumo_dia_total'] = str(e)

    try:
        conn.execute('DROP TABLE IF EXISTS ranking_recursos_dia')
        conn.execute('''
            CREATE TABLE ranking_recursos_dia AS
            SELECT
                DATA,
                DATA_SQL,
                CONTRATO,
                RECURSO,
                COUNT(DISTINCT CASE WHEN EH_NOTA_VALIDA = 1 THEN ORDEM_DE_SERVICO END) AS NOTAS,
                SUM(EH_CORTE) AS CORTES,
                SUM(EH_RELIGUE) AS RELIGUES,
                SUM(EH_VERIFICACAO) AS VERIFICACOES,
                SUM(EH_RECUSA) AS RECUSAS
            FROM notas_processadas
            WHERE COALESCE(CONTRATO,'') <> '' AND COALESCE(RECURSO,'') <> ''
            GROUP BY DATA, DATA_SQL, CONTRATO, RECURSO
        ''')
        out['ranking_recursos_dia'] = conn.execute('SELECT COUNT(*) FROM ranking_recursos_dia').fetchone()[0]
    except Exception as e:
        out['erro_ranking_recursos_dia'] = str(e)

    try:
        conn.execute('DROP TABLE IF EXISTS meses_notas')
        conn.execute('''
            CREATE TABLE meses_notas AS
            SELECT DISTINCT substr(DATA, 4, 7) AS MES
            FROM notas_processadas
            WHERE DATA IS NOT NULL AND length(DATA) >= 10
        ''')
        out['meses_notas'] = conn.execute('SELECT COUNT(*) FROM meses_notas').fetchone()[0]
    except Exception as e:
        out['erro_meses_notas'] = str(e)

    criar_indices(conn)
    for tabela, coluna in [
        ('resumo_dia', 'DATA'), ('resumo_dia', 'DATA_SQL'), ('resumo_dia', 'CONTRATO'),
        ('resumo_dia_total', 'DATA'), ('resumo_dia_total', 'DATA_SQL'),
        ('ranking_recursos_dia', 'DATA'), ('ranking_recursos_dia', 'DATA_SQL'),
        ('ranking_recursos_dia', 'CONTRATO'), ('ranking_recursos_dia', 'RECURSO'),
    ]:
        try:
            if tabela_existe(conn, tabela) and coluna in colunas(conn, tabela):
                conn.execute(f'CREATE INDEX IF NOT EXISTS idx_{tabela}_{coluna} ON "{tabela}" ("{coluna}")')
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
            if 'TAREFA' in cols_norm or aba.upper().strip() in ['TAREFAS', 'TAREFAS_CONSOLIDADAS']:
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


def limpar_banco_dashboard_leve(conn: sqlite3.Connection) -> dict[str, str]:
    """Remove tabelas brutas do banco leve para ficar abaixo do limite do GitHub Web.

    O banco completo local continua com tudo.
    O banco dashboard fica apenas com resumos e tabelas pequenas.
    """
    removidas = {}
    for tabela in [
        'notas',
        'notas_processadas',
        'leitura_tarefas',
        'leitura_resumos',
    ]:
        try:
            if tabela_existe(conn, tabela):
                conn.execute(f'DROP TABLE IF EXISTS "{tabela}"')
                removidas[tabela] = 'removida'
        except Exception as e:
            removidas[tabela] = f'erro: {e}'
    return removidas


def importar_tudo(limite_excels: Optional[int] = None) -> dict:
    garantir_pastas()
    resultado = {'completo': {}, 'dashboard': {}}

    with conectar(BANCO_COMPLETO) as conn:
        resultado['completo']['csvs'] = importar_csvs(conn, incluir_metadados=True)
        resultado['completo']['resumos'] = criar_resumos(conn)
        resultado['completo']['leitura'] = importar_leitura_completa(conn, limite=limite_excels)
        conn.commit()
    with conectar(BANCO_COMPLETO) as conn:
        otimizar(conn)

    with conectar(BANCO_DASHBOARD) as conn:
        # O banco leve precisa importar as notas temporariamente para gerar resumos,
        # mas NÃO deve manter a tabela bruta, senão passa do limite de upload do GitHub Web.
        resultado['dashboard']['csvs'] = importar_csvs(conn, incluir_metadados=False)
        resultado['dashboard']['resumos'] = criar_resumos(conn)
        resultado['dashboard']['limpeza_banco_leve'] = limpar_banco_dashboard_leve(conn)
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
        tabelas = pd.read_sql_query('SELECT name FROM sqlite_master WHERE type=\'table\' ORDER BY name', conn)
        linhas = []
        for t in tabelas['name'].tolist():
            try:
                qtd = conn.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
            except Exception:
                qtd = None
            linhas.append({'banco': caminho.name, 'tabela': t, 'linhas': qtd})
        return pd.DataFrame(linhas)


def diagnostico_datas() -> None:
    for caminho in [BANCO_COMPLETO, BANCO_DASHBOARD]:
        if not caminho.exists():
            continue
        with sqlite3.connect(str(caminho)) as conn:
            print(f'\nDatas em {caminho.name}:')
            for tabela in ['notas', 'notas_processadas', 'resumo_dia']:
                if tabela_existe(conn, tabela):
                    try:
                        q = f'''
                            SELECT DATA, COUNT(*) AS LINHAS
                            FROM "{tabela}"
                            GROUP BY DATA
                            ORDER BY substr(DATA,7,4)||substr(DATA,4,2)||substr(DATA,1,2) DESC
                            LIMIT 8
                        '''
                        print(f'\n{tabela}:')
                        print(pd.read_sql_query(q, conn).to_string(index=False))
                    except Exception as e:
                        print(f'{tabela}: erro {e}')


def main() -> None:
    parser = argparse.ArgumentParser(description='Banco SQLite do sistema G.Z.U.S.')
    parser.add_argument('acao', choices=['importar', 'resumo', 'datas'], help='O que deseja fazer')
    parser.add_argument('--limite-excels', type=int, default=None, help='Limita quantos Excels de leitura importar no banco completo')
    args = parser.parse_args()

    if args.acao == 'importar':
        r = importar_tudo(limite_excels=args.limite_excels)
        print(json.dumps(r, ensure_ascii=False, indent=2))
        print('\nImportação finalizada.')
        print(f'Banco completo: {BANCO_COMPLETO.resolve()}')
        print(f'Banco leve do dashboard: {BANCO_DASHBOARD.resolve()}')
    elif args.acao == 'datas':
        diagnostico_datas()
    else:
        dfs = [resumo_banco(BANCO_COMPLETO), resumo_banco(BANCO_DASHBOARD)]
        df = pd.concat([d for d in dfs if not d.empty], ignore_index=True) if any(not d.empty for d in dfs) else pd.DataFrame()
        if df.empty:
            print('Bancos ainda não criados.')
        else:
            print(df.to_string(index=False))


if __name__ == '__main__':
    main()
