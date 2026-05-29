#!/usr/bin/env python3
from __future__ import annotations
import os, sqlite3, unicodedata
from pathlib import Path
import pandas as pd
BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_DIR = BASE_DIR / "dashboard"
NOTAS_CSV = DASHBOARD_DIR / "notas_dashboard.csv"
DB_PATH = DASHBOARD_DIR / "gzus_dashboard.db"
TXT_SUPERVISOR_STC_CSV = DASHBOARD_DIR / "txt_supervisor_stc_santa_cruz.csv"
def env_float(nome, padrao):
    try: return float(os.getenv(nome, padrao))
    except Exception: return float(padrao)
def log(msg): print(msg, flush=True)
def eh_disjuntor_jundiai(recurso):
    r=str(recurso or '').strip().upper(); return r.startswith('JUN55') or r.startswith('JUN59') or r.startswith('SAL55')
def eh_disjuntor_santa_cruz(recurso):
    import re
    r=str(recurso or '').strip().upper(); m=re.search(r'(\d+)', r)
    return bool(m and (m.group(1).startswith('89') or m.group(1).startswith('20')))

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

def preparar_notas_processadas(notas):
    df=notas.copy()
    for col in ['ORDEM_DE_SERVICO','GRUPO_NOTA','RECURSO','RECUSA','ELETRICISTA1','ELETRICISTA2','DATA','DATA_ENCERRAMENTO']:
        if col not in df.columns: df[col]=''
        df[col]=df[col].fillna('').astype(str).str.strip()
    if 'QTD_EXECUTORES' not in df.columns:
        df['QTD_EXECUTORES']=((df['ELETRICISTA1']!='').astype(int)+(df['ELETRICISTA2']!='').astype(int))
    else: df['QTD_EXECUTORES']=pd.to_numeric(df['QTD_EXECUTORES'], errors='coerce').fillna(0).astype(int)
    df['GRUPO_NOTA']=df['GRUPO_NOTA'].apply(normalizar_grupo_nota); df['RECURSO']=df['RECURSO'].str.upper(); df['EH_RECUSA']=(df['RECUSA']!='').astype(int)
    data_raw=df['DATA'] if 'DATA' in df.columns and not (df['DATA'].fillna('').astype(str).str.strip()=='').all() else df['DATA_ENCERRAMENTO']
    df['DATA_DT']=pd.to_datetime(data_raw, dayfirst=True, errors='coerce'); df['DATA']=df['DATA_DT'].dt.strftime('%d/%m/%Y').fillna(''); df=df[df['DATA']!=''].copy()
    tjc=env_float('TARIFA_DISJUNTOR_JUNDIAI_CORTE',13.72); tjr=env_float('TARIFA_DISJUNTOR_JUNDIAI_RELIGUE',27.43); tsc=env_float('TARIFA_DISJUNTOR_SANTA_CRUZ_CORTE',11.98); tsr=env_float('TARIFA_DISJUNTOR_SANTA_CRUZ_RELIGUE',23.97); tsv=env_float('TARIFA_DISJUNTOR_SANTA_CRUZ_VERIFICACAO',23.97)
    def classificar(row):
        recurso=row['RECURSO']; grupo=row['GRUPO_NOTA']; rec=int(row['EH_RECUSA'])==1; qtd=int(row.get('QTD_EXECUTORES',0) or 0)
        contrato=''; fat=fmin=fmax=0.0
        if eh_disjuntor_jundiai(recurso):
            contrato='Disjuntor Jundiaí';
            if not rec: fat={'CORTE':tjc,'VERIFICACAO':tjc,'RELIGUE':tjr}.get(grupo,0.0); fmin=fmax=fat
        elif eh_disjuntor_santa_cruz(recurso):
            contrato='Disjuntor Santa Cruz';
            if not rec: fat={'CORTE':tsc,'RELIGUE':tsr,'VERIFICACAO':tsv}.get(grupo,0.0); fmin=fmax=fat
        elif recurso.startswith('JUN58') and qtd>=2:
            contrato='STC Jundiai';
            if not rec: fmin={'CORTE':38.18,'RELIGUE':36.36}.get(grupo,0.0); fmax={'CORTE':45.45,'RELIGUE':50.91}.get(grupo,0.0); fat=fmin
        elif recurso.startswith('JUN') or recurso.startswith('SAL'):
            contrato='STC Jundiai'
        return pd.Series({'CONTRATO':contrato,'FATURAMENTO':fat,'FATURAMENTO_MIN':fmin,'FATURAMENTO_MAX':fmax,'EH_CORTE':1 if ((grupo=='CORTE') or (grupo=='VERIFICACAO' and contrato=='Disjuntor Jundiaí')) and not rec else 0,'EH_RELIGUE':1 if grupo=='RELIGUE' and not rec else 0,'EH_VERIFICACAO':1 if grupo=='VERIFICACAO' and contrato=='Disjuntor Santa Cruz' and not rec else 0})
    df=pd.concat([df.reset_index(drop=True), df.apply(classificar, axis=1).reset_index(drop=True)], axis=1); df=df[df['CONTRATO'].fillna('').astype(str).str.strip()!=''].copy()
    df['ORDEM_SERVICO_PAGAVEL']=df['ORDEM_DE_SERVICO'].where(df['EH_RECUSA']==0, pd.NA); df['ORDEM_SERVICO_RECUSA']=df['ORDEM_DE_SERVICO'].where(df['EH_RECUSA']==1, pd.NA); df['DATA_PAGAVEL']=df['DATA'].where(df['EH_RECUSA']==0, pd.NA)
    df['MES']=df['DATA_DT'].dt.strftime('%m/%Y'); df['SEMANA_INICIO_DT']=df['DATA_DT']-pd.to_timedelta(df['DATA_DT'].dt.weekday, unit='D'); df['SEMANA']=df['SEMANA_INICIO_DT'].dt.strftime('%d/%m/%Y')
    return df
def gerar_parcial(df):
    detalhe_cols=['DATA','CONTRATO','ORDEM_DE_SERVICO','GRUPO_NOTA','RECURSO','RECUSA','ELETRICISTA1','ELETRICISTA2','QTD_EXECUTORES','EH_RECUSA','EH_CORTE','EH_RELIGUE','EH_VERIFICACAO','FATURAMENTO','FATURAMENTO_MIN','FATURAMENTO_MAX']
    detalhe=df[[c for c in detalhe_cols if c in df.columns]].copy()
    parcial=df.groupby(['DATA','CONTRATO','RECURSO'], dropna=False).agg(NOTAS=('ORDEM_SERVICO_PAGAVEL','nunique'),CORTES=('EH_CORTE','sum'),RELIGUES=('EH_RELIGUE','sum'),VERIFICACOES=('EH_VERIFICACAO','sum'),RECUSAS=('ORDEM_SERVICO_RECUSA','nunique'),FATURAMENTO=('FATURAMENTO','sum'),FATURAMENTO_MIN=('FATURAMENTO_MIN','sum'),FATURAMENTO_MAX=('FATURAMENTO_MAX','sum')).reset_index()
    return parcial, detalhe
def calcular_ranking(base):
    if base.empty: return pd.DataFrame()
    r=base.groupby('RECURSO', dropna=False).agg(NOTAS=('ORDEM_SERVICO_PAGAVEL','nunique'),CORTES=('EH_CORTE','sum'),RELIGUES=('EH_RELIGUE','sum'),VERIFICACOES=('EH_VERIFICACAO','sum'),RECUSAS=('ORDEM_SERVICO_RECUSA','nunique'),FATURAMENTO_ATRIBUÍDO=('FATURAMENTO','sum'),FATURAMENTO_MIN_ATRIBUÍDO=('FATURAMENTO_MIN','sum'),FATURAMENTO_MAX_ATRIBUÍDO=('FATURAMENTO_MAX','sum'),DIAS_ATIVOS=('DATA_PAGAVEL','nunique')).reset_index()
    r['MÉDIA_NOTAS_DIA']=(r['NOTAS']/r['DIAS_ATIVOS'].replace(0,pd.NA)).fillna(0).round(2); r['TICKET_MÉDIO']=(r['FATURAMENTO_ATRIBUÍDO']/r['NOTAS'].replace(0,pd.NA)).fillna(0).round(2)
    r=r.sort_values(['NOTAS','FATURAMENTO_ATRIBUÍDO'], ascending=[False,False]).reset_index(drop=True); r['POSICAO_NOTAS']=range(1,len(r)+1)
    r=r.sort_values(['FATURAMENTO_ATRIBUÍDO','NOTAS'], ascending=[False,False]).reset_index(drop=True); r['POSICAO_FATURAMENTO']=range(1,len(r)+1)
    return r
def gerar_ranking(df):
    partes=[]; contratos=['Todos']+sorted([c for c in df['CONTRATO'].dropna().astype(str).unique().tolist() if c]); periodos=[('Total','Total',df)]
    for data,sub in df.groupby('DATA', dropna=False): periodos.append(('Dia',data,sub))
    for semana,sub in df.groupby('SEMANA', dropna=False): periodos.append(('Semana',semana,sub))
    for mes,sub in df.groupby('MES', dropna=False): periodos.append(('Mês',mes,sub))
    for tipo,valor,basep in periodos:
        for contrato in contratos:
            base=basep if contrato=='Todos' else basep[basep['CONTRATO']==contrato]
            r=calcular_ranking(base)
            if not r.empty:
                r.insert(0,'VALOR_PERIODO',valor); r.insert(0,'TIPO_PERIODO',tipo); r.insert(0,'CONTRATO',contrato); partes.append(r)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
def gerar_opcoes(df):
    linhas=[]
    def add(chave, valores):
        for i,v in enumerate(valores):
            if pd.notna(v) and str(v).strip(): linhas.append({'CHAVE':chave,'VALOR':str(v),'ORDEM':i})
    contratos=['Todos']+sorted([c for c in df['CONTRATO'].dropna().astype(str).unique().tolist() if c]); add('contratos_ranking', contratos)
    for contrato in contratos:
        base=df if contrato=='Todos' else df[df['CONTRATO']==contrato]
        datas=base[['DATA','DATA_DT']].drop_duplicates().sort_values('DATA_DT', ascending=False)['DATA'].tolist(); semanas=base[['SEMANA','SEMANA_INICIO_DT']].drop_duplicates().sort_values('SEMANA_INICIO_DT', ascending=False)['SEMANA'].tolist(); mdf=base[['MES','DATA_DT']].copy(); mdf['PER']=mdf['DATA_DT'].dt.to_period('M'); meses=mdf.drop_duplicates('MES').sort_values('PER', ascending=False)['MES'].tolist()
        add(f'datas_parcial::{contrato}', datas); add(f'ranking_dias::{contrato}', datas); add(f'ranking_semanas::{contrato}', semanas); add(f'ranking_meses::{contrato}', meses)
    return pd.DataFrame(linhas)
def gerar_txt_supervisor_stc_csv(notas, base):
    """Gera CSV leve e pré-filtrado para o login Supervisor STC.

    Contém somente STC Jundiai e Disjuntor Santa Cruz.
    Não inclui Disjuntor Jundiaí.
    É usado pelo app.py para abrir o login do supervisor sem carregar/classificar
    todo o notas_dashboard.csv.
    """
    if notas is None or notas.empty or base is None or base.empty:
        return pd.DataFrame()

    permitidos = ["STC Jundiai", "Disjuntor Santa Cruz"]

    mapa = (
        base[["ORDEM_DE_SERVICO", "CONTRATO"]]
        .dropna(subset=["ORDEM_DE_SERVICO"])
        .drop_duplicates(subset=["ORDEM_DE_SERVICO"], keep="last")
        .copy()
    )
    mapa["ORDEM_DE_SERVICO"] = mapa["ORDEM_DE_SERVICO"].astype(str).str.strip()

    saida = notas.copy()
    if "ORDEM_DE_SERVICO" not in saida.columns:
        return pd.DataFrame()

    saida["ORDEM_DE_SERVICO"] = saida["ORDEM_DE_SERVICO"].fillna("").astype(str).str.strip()
    saida = saida.merge(mapa, on="ORDEM_DE_SERVICO", how="left")
    saida = saida[saida["CONTRATO"].isin(permitidos)].copy()

    colunas_preferidas = [
        "ORDEM_DE_SERVICO",
        "GRUPO_NOTA",
        "CONTRATO",
        "RECURSO",
        "STATUS",
        "DATA",
        "DATA_ENCERRAMENTO",
        "ELETRICISTA1",
        "ELETRICISTA2",
        "RECUSA",
        "QTD_EXECUTORES",
    ]
    for col in colunas_preferidas:
        if col not in saida.columns:
            saida[col] = ""

    outras = [c for c in saida.columns if c not in colunas_preferidas]
    saida = saida[colunas_preferidas + outras].copy()
    return saida

def main():
    if not NOTAS_CSV.exists(): log(f'⚠️ notas_dashboard.csv não encontrado: {NOTAS_CSV}'); return 0
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True); log(f'📚 Lendo histórico de notas: {NOTAS_CSV}'); notas=pd.read_csv(NOTAS_CSV, sep=';', encoding='utf-8-sig', dtype=str); log(f'📊 Notas carregadas: {len(notas):,}'.replace(',','.'))
    base=preparar_notas_processadas(notas); parcial, detalhe=gerar_parcial(base); ranking=gerar_ranking(base); opcoes=gerar_opcoes(base); txt_supervisor_stc=gerar_txt_supervisor_stc_csv(notas, base)
    with sqlite3.connect(DB_PATH) as conn:
        parcial.to_sql('instant_parcial_recurso', conn, if_exists='replace', index=False); detalhe.to_sql('instant_parcial_detalhe', conn, if_exists='replace', index=False); ranking.to_sql('instant_ranking', conn, if_exists='replace', index=False); opcoes.to_sql('instant_opcoes', conn, if_exists='replace', index=False)
        conn.execute('CREATE INDEX IF NOT EXISTS idx_instant_parcial ON instant_parcial_recurso(DATA, CONTRATO, RECURSO)'); conn.execute('CREATE INDEX IF NOT EXISTS idx_instant_detalhe ON instant_parcial_detalhe(DATA, CONTRATO, RECURSO)'); conn.execute('CREATE INDEX IF NOT EXISTS idx_instant_ranking ON instant_ranking(CONTRATO, TIPO_PERIODO, VALOR_PERIODO)'); conn.execute('CREATE INDEX IF NOT EXISTS idx_instant_opcoes ON instant_opcoes(CHAVE, ORDEM)')
    txt_supervisor_stc.to_csv(TXT_SUPERVISOR_STC_CSV, sep=';', encoding='utf-8-sig', index=False); log(f'✅ Tabelas instantâneas gravadas em {DB_PATH}'); log(f'✅ CSV leve do Supervisor STC gravado em {TXT_SUPERVISOR_STC_CSV}: {len(txt_supervisor_stc):,} linhas'.replace(',','.')); log(f'   instant_parcial_recurso: {len(parcial):,}'.replace(',','.')); log(f'   instant_parcial_detalhe: {len(detalhe):,}'.replace(',','.')); log(f'   instant_ranking: {len(ranking):,}'.replace(',','.')); return 0
if __name__=='__main__': raise SystemExit(main())
