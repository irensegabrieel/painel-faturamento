#!/usr/bin/env python3
from __future__ import annotations

import sqlite3
import unicodedata
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_DIR = BASE_DIR / "dashboard"
NOTAS_CSV = DASHBOARD_DIR / "notas_dashboard.csv"
DB_PATH = DASHBOARD_DIR / "gzus_dashboard.db"
TXT_SUPERVISOR_STC_CSV = DASHBOARD_DIR / "txt_supervisor_stc_santa_cruz.csv"

CONTRATOS_SUPERVISOR_STC = ["STC Jundiai", "Disjuntor Santa Cruz"]


def log(msg: str) -> None:
    print(msg, flush=True)


def sem_acentos(valor) -> str:
    texto = str(valor or "").strip().upper()
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


def eh_disjuntor_jundiai(recurso) -> bool:
    r = str(recurso or "").strip().upper()
    return r.startswith("JUN55") or r.startswith("JUN59") or r.startswith("SAL55")


def eh_disjuntor_santa_cruz(recurso) -> bool:
    import re
    r = str(recurso or "").strip().upper()
    m = re.search(r"(\d+)", r)
    return bool(m and (m.group(1).startswith("89") or m.group(1).startswith("20")))


def normalizar_grupo_nota(valor) -> str:
    texto = sem_acentos(valor)
    if "VERIFIC" in texto:
        return "VERIFICACAO"
    if "RELIG" in texto:
        return "RELIGUE"
    if "CORTE" in texto:
        return "CORTE"
    return texto


def normalizar_contrato(valor) -> str:
    texto = sem_acentos(valor).replace("Í", "I")
    if "DISJUNTOR" in texto and "SANTA" in texto:
        return "Disjuntor Santa Cruz"
    if "DISJUNTOR" in texto and ("JUNDIAI" in texto or "JUNDIA" in texto):
        return "Disjuntor Jundiaí"
    if "STC" in texto and ("JUNDIAI" in texto or "JUNDIA" in texto):
        return "STC Jundiai"
    return str(valor or "").strip()


def preparar_notas_processadas(notas: pd.DataFrame) -> pd.DataFrame:
    df = notas.copy()

    for col in [
        "ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO", "RECUSA",
        "ELETRICISTA1", "ELETRICISTA2", "DATA", "DATA_ENCERRAMENTO",
    ]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "QTD_EXECUTORES" not in df.columns:
        df["QTD_EXECUTORES"] = ((df["ELETRICISTA1"] != "").astype(int) + (df["ELETRICISTA2"] != "").astype(int))
    else:
        df["QTD_EXECUTORES"] = pd.to_numeric(df["QTD_EXECUTORES"], errors="coerce").fillna(0).astype(int)

    df["GRUPO_NOTA"] = df["GRUPO_NOTA"].apply(normalizar_grupo_nota)
    df["RECURSO"] = df["RECURSO"].str.upper().str.strip()
    df["EH_RECUSA"] = (df["RECUSA"] != "").astype(int)

    data_raw = df["DATA"] if not (df["DATA"].fillna("").astype(str).str.strip() == "").all() else df["DATA_ENCERRAMENTO"]
    df["DATA_DT"] = pd.to_datetime(data_raw, dayfirst=True, errors="coerce")
    df["DATA"] = df["DATA_DT"].dt.strftime("%d/%m/%Y").fillna("")
    df = df[df["DATA"] != ""].copy()

    def classificar(row):
        recurso = row["RECURSO"]
        grupo = row["GRUPO_NOTA"]
        recusa = int(row["EH_RECUSA"] or 0) == 1
        qtd = int(row.get("QTD_EXECUTORES", 0) or 0)
        contrato = ""

        if eh_disjuntor_jundiai(recurso):
            contrato = "Disjuntor Jundiaí"
        elif eh_disjuntor_santa_cruz(recurso):
            contrato = "Disjuntor Santa Cruz"
        elif recurso.startswith("JUN58") and qtd >= 2:
            contrato = "STC Jundiai"
        elif recurso.startswith("JUN") or recurso.startswith("SAL"):
            contrato = "STC Jundiai"

        # Regra operacional nova: VERIFICAÇÃO conta como CORTE quando aplicável.
        eh_corte = 1 if (not recusa and grupo in ["CORTE", "VERIFICACAO"]) else 0
        eh_religue = 1 if (not recusa and grupo == "RELIGUE") else 0
        eh_verificacao = 1 if (not recusa and grupo == "VERIFICACAO") else 0
        return pd.Series({
            "CONTRATO": contrato,
            "EH_CORTE": eh_corte,
            "EH_RELIGUE": eh_religue,
            "EH_VERIFICACAO": eh_verificacao,
        })

    df = pd.concat([df.reset_index(drop=True), df.apply(classificar, axis=1).reset_index(drop=True)], axis=1)
    df["CONTRATO"] = df["CONTRATO"].apply(normalizar_contrato)
    df = df[df["CONTRATO"].fillna("").astype(str).str.strip() != ""].copy()

    df["ORDEM_SERVICO_PAGAVEL"] = df["ORDEM_DE_SERVICO"].where(df["EH_RECUSA"] == 0, pd.NA)
    df["ORDEM_SERVICO_RECUSA"] = df["ORDEM_DE_SERVICO"].where(df["EH_RECUSA"] == 1, pd.NA)
    df["DATA_PAGAVEL"] = df["DATA"].where(df["EH_RECUSA"] == 0, pd.NA)
    df["MES"] = df["DATA_DT"].dt.strftime("%m/%Y")
    df["SEMANA_INICIO_DT"] = df["DATA_DT"] - pd.to_timedelta(df["DATA_DT"].dt.weekday, unit="D")
    df["SEMANA"] = df["SEMANA_INICIO_DT"].dt.strftime("%d/%m/%Y")
    return df


def gerar_parcial(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    detalhe_cols = [
        "DATA", "MES", "SEMANA", "CONTRATO", "ORDEM_DE_SERVICO", "GRUPO_NOTA", "RECURSO",
        "RECUSA", "ELETRICISTA1", "ELETRICISTA2", "QTD_EXECUTORES", "EH_RECUSA",
        "EH_CORTE", "EH_RELIGUE", "EH_VERIFICACAO",
    ]
    detalhe = df[[c for c in detalhe_cols if c in df.columns]].copy()

    parcial = (
        df.groupby(["DATA", "MES", "SEMANA", "CONTRATO", "RECURSO"], dropna=False)
        .agg(
            NOTAS=("ORDEM_SERVICO_PAGAVEL", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            RECUSAS=("ORDEM_SERVICO_RECUSA", "nunique"),
        )
        .reset_index()
    )
    return parcial, detalhe


def calcular_ranking(base: pd.DataFrame) -> pd.DataFrame:
    if base.empty:
        return pd.DataFrame()
    r = (
        base.groupby("RECURSO", dropna=False)
        .agg(
            NOTAS=("ORDEM_SERVICO_PAGAVEL", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            VERIFICACOES=("EH_VERIFICACAO", "sum"),
            RECUSAS=("ORDEM_SERVICO_RECUSA", "nunique"),
            DIAS_ATIVOS=("DATA_PAGAVEL", "nunique"),
        )
        .reset_index()
    )
    r["MÉDIA_NOTAS_DIA"] = (r["NOTAS"] / r["DIAS_ATIVOS"].replace(0, pd.NA)).fillna(0).round(2)
    r = r.sort_values(["NOTAS", "CORTES", "RELIGUES", "RECUSAS"], ascending=[False, False, False, False]).reset_index(drop=True)
    r.insert(0, "POSIÇÃO", range(1, len(r) + 1))
    return r


def _com_contrato_e_periodo(ranking: pd.DataFrame, contrato: str, **periodo) -> pd.DataFrame:
    if ranking.empty:
        return ranking
    out = ranking.copy()
    out.insert(0, "CONTRATO", contrato)
    for col, valor in reversed(list(periodo.items())):
        out.insert(1, col, valor)
    return out


def gerar_rankings_separados(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    contratos = ["Todos"] + sorted([c for c in df["CONTRATO"].dropna().astype(str).unique().tolist() if c])
    saida = {"dia": [], "semana": [], "mes": [], "total": []}

    for contrato in contratos:
        base_contrato = df if contrato == "Todos" else df[df["CONTRATO"] == contrato]
        r_total = calcular_ranking(base_contrato)
        if not r_total.empty:
            saida["total"].append(_com_contrato_e_periodo(r_total, contrato))

        for data, sub in base_contrato.groupby("DATA", dropna=False):
            r = calcular_ranking(sub)
            if not r.empty:
                saida["dia"].append(_com_contrato_e_periodo(r, contrato, DATA=data))

        for semana, sub in base_contrato.groupby("SEMANA", dropna=False):
            r = calcular_ranking(sub)
            if not r.empty:
                saida["semana"].append(_com_contrato_e_periodo(r, contrato, SEMANA=semana))

        for mes, sub in base_contrato.groupby("MES", dropna=False):
            r = calcular_ranking(sub)
            if not r.empty:
                saida["mes"].append(_com_contrato_e_periodo(r, contrato, MES=mes))

    return {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame() for k, v in saida.items()}


def gerar_ranking_compat(df: pd.DataFrame) -> pd.DataFrame:
    partes = []
    rank = gerar_rankings_separados(df)
    for tipo, tabela, coluna in [
        ("Dia", rank["dia"], "DATA"),
        ("Semana", rank["semana"], "SEMANA"),
        ("Mês", rank["mes"], "MES"),
        ("Total", rank["total"], None),
    ]:
        if tabela.empty:
            continue
        t = tabela.copy()
        t.insert(1, "TIPO_PERIODO", tipo)
        t.insert(2, "VALOR_PERIODO", "Total" if coluna is None else t[coluna])
        partes.append(t)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()


def gerar_opcoes(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []

    def add(chave, valores):
        for i, v in enumerate(valores):
            if pd.notna(v) and str(v).strip():
                linhas.append({"CHAVE": chave, "VALOR": str(v), "ORDEM": i})

    contratos = ["Todos"] + sorted([c for c in df["CONTRATO"].dropna().astype(str).unique().tolist() if c])
    add("contratos_ranking", contratos)
    for contrato in contratos:
        base = df if contrato == "Todos" else df[df["CONTRATO"] == contrato]
        datas = base[["DATA", "DATA_DT"]].drop_duplicates().sort_values("DATA_DT", ascending=False)["DATA"].tolist()
        semanas = base[["SEMANA", "SEMANA_INICIO_DT"]].drop_duplicates().sort_values("SEMANA_INICIO_DT", ascending=False)["SEMANA"].tolist()
        mdf = base[["MES", "DATA_DT"]].copy()
        mdf["PER"] = mdf["DATA_DT"].dt.to_period("M")
        meses = mdf.drop_duplicates("MES").sort_values("PER", ascending=False)["MES"].tolist()
        add(f"datas_parcial::{contrato}", datas)
        add(f"ranking_dias::{contrato}", datas)
        add(f"ranking_semanas::{contrato}", semanas)
        add(f"ranking_meses::{contrato}", meses)
    return pd.DataFrame(linhas)


def gerar_txt_supervisor_stc_csv(notas: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    if notas is None or notas.empty or base is None or base.empty:
        return pd.DataFrame()
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
    saida = saida[saida["CONTRATO"].isin(CONTRATOS_SUPERVISOR_STC)].copy()

    colunas_preferidas = [
        "ORDEM_DE_SERVICO", "GRUPO_NOTA", "CONTRATO", "RECURSO", "STATUS", "DATA",
        "DATA_ENCERRAMENTO", "ELETRICISTA1", "ELETRICISTA2", "RECUSA", "QTD_EXECUTORES",
    ]
    for col in colunas_preferidas:
        if col not in saida.columns:
            saida[col] = ""
    outras = [c for c in saida.columns if c not in colunas_preferidas and "FATUR" not in c.upper() and "VALOR" not in c.upper()]
    return saida[colunas_preferidas + outras].copy()


def criar_indices(conn: sqlite3.Connection) -> None:
    indices = [
        "CREATE INDEX IF NOT EXISTS idx_instant_parcial_dia ON instant_parcial_dia(DATA, CONTRATO, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_parcial_recurso ON instant_parcial_recurso(DATA, CONTRATO, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_detalhe ON instant_parcial_detalhe(DATA, CONTRATO, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_ranking_dia ON instant_ranking_dia(CONTRATO, DATA, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_ranking_semana ON instant_ranking_semana(CONTRATO, SEMANA, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_ranking_mes ON instant_ranking_mes(CONTRATO, MES, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_ranking_total ON instant_ranking_total(CONTRATO, RECURSO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_ranking ON instant_ranking(CONTRATO, TIPO_PERIODO, VALOR_PERIODO)",
        "CREATE INDEX IF NOT EXISTS idx_instant_opcoes ON instant_opcoes(CHAVE, ORDEM)",
    ]
    for sql in indices:
        conn.execute(sql)


def main() -> int:
    if not NOTAS_CSV.exists():
        log(f"⚠️ notas_dashboard.csv não encontrado: {NOTAS_CSV}")
        return 0

    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
    log(f"📚 Lendo histórico de notas: {NOTAS_CSV}")
    notas = pd.read_csv(NOTAS_CSV, sep=";", encoding="utf-8-sig", dtype=str)
    log(f"📊 Notas carregadas: {len(notas):,}".replace(",", "."))

    base = preparar_notas_processadas(notas)
    parcial, detalhe = gerar_parcial(base)
    rankings = gerar_rankings_separados(base)
    ranking_compat = gerar_ranking_compat(base)
    opcoes = gerar_opcoes(base)
    txt_supervisor_stc = gerar_txt_supervisor_stc_csv(notas, base)

    with sqlite3.connect(DB_PATH) as conn:
        parcial.to_sql("instant_parcial_dia", conn, if_exists="replace", index=False)
        parcial.to_sql("instant_parcial_recurso", conn, if_exists="replace", index=False)
        detalhe.to_sql("instant_parcial_detalhe", conn, if_exists="replace", index=False)
        rankings["dia"].to_sql("instant_ranking_dia", conn, if_exists="replace", index=False)
        rankings["semana"].to_sql("instant_ranking_semana", conn, if_exists="replace", index=False)
        rankings["mes"].to_sql("instant_ranking_mes", conn, if_exists="replace", index=False)
        rankings["total"].to_sql("instant_ranking_total", conn, if_exists="replace", index=False)
        ranking_compat.to_sql("instant_ranking", conn, if_exists="replace", index=False)
        opcoes.to_sql("instant_opcoes", conn, if_exists="replace", index=False)
        criar_indices(conn)

    txt_supervisor_stc.to_csv(TXT_SUPERVISOR_STC_CSV, sep=";", encoding="utf-8-sig", index=False)
    log(f"✅ Tabelas instantâneas operacionais gravadas em {DB_PATH}")
    log(f"✅ CSV leve do Supervisor STC gravado em {TXT_SUPERVISOR_STC_CSV}: {len(txt_supervisor_stc):,} linhas".replace(",", "."))
    for nome, tabela in [
        ("instant_parcial_dia", parcial),
        ("instant_parcial_detalhe", detalhe),
        ("instant_ranking_dia", rankings["dia"]),
        ("instant_ranking_semana", rankings["semana"]),
        ("instant_ranking_mes", rankings["mes"]),
        ("instant_ranking_total", rankings["total"]),
    ]:
        log(f"   {nome}: {len(tabela):,}".replace(",", "."))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
