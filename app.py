from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Painel de Faturamento", page_icon="📊", layout="wide")

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


@st.cache_data(show_spinner=False)
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
    for col in df2.columns:
        if "FATURAMENTO" in col or col in ["CORTE", "RELIGUE", "TOTAL", "MÍNIMO", "MÁXIMO"]:
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["QTD_NOTAS", "NOTAS"]:
            df2[col] = df2[col].apply(numero)
    return df2


def carregar_bases():
    bases = {}
    faltando = []
    for chave, nome in ARQUIVOS.items():
        caminho = caminho_arquivo(nome)
        if caminho:
            bases[chave] = ler_csv(caminho)
        else:
            faltando.append(nome)
    return bases, faltando


bases, faltando = carregar_bases()

st.title("📊 Painel de Faturamento")
st.caption("Painel gerado a partir dos arquivos CSV da pasta dashboard.")

if faltando:
    st.warning("Arquivos não encontrados: " + ", ".join(faltando))

if not bases:
    st.error("Nenhum CSV foi encontrado. Coloque este app.py na mesma pasta dos CSVs ou deixe os CSVs dentro de uma pasta chamada dashboard.")
    st.stop()

contratos_original = bases.get("contratos", pd.DataFrame())
carro_original = bases.get("carro", pd.DataFrame())
dias_original = bases.get("dias", pd.DataFrame())
carro_dias_original = bases.get("carro_dias", pd.DataFrame())
notas = bases.get("notas", pd.DataFrame())

st.sidebar.header("Filtros")
contratos_lista = []
for base in [contratos_original, dias_original, carro_original, carro_dias_original]:
    if not base.empty and "CONTRATO" in base.columns:
        contratos_lista += base["CONTRATO"].dropna().unique().tolist()
contratos_lista = sorted(set(contratos_lista))
contratos_selecionados = st.sidebar.multiselect("Contratos", contratos_lista, default=contratos_lista)

contratos = contratos_original.copy()
carro = carro_original.copy()
dias = dias_original.copy()
carro_dias = carro_dias_original.copy()

if contratos_selecionados:
    if not contratos.empty and "CONTRATO" in contratos.columns:
        contratos = contratos[contratos["CONTRATO"].isin(contratos_selecionados)]
    if not dias.empty and "CONTRATO" in dias.columns:
        dias = dias[dias["CONTRATO"].isin(contratos_selecionados)]
    if not carro.empty and "CONTRATO" in carro.columns:
        carro = carro[carro["CONTRATO"].isin(contratos_selecionados)]
    if not carro_dias.empty and "CONTRATO" in carro_dias.columns:
        carro_dias = carro_dias[carro_dias["CONTRATO"].isin(contratos_selecionados)]
else:
    contratos = contratos.iloc[0:0]
    dias = dias.iloc[0:0]
    carro = carro.iloc[0:0]
    carro_dias = carro_dias.iloc[0:0]

mostrar_carro = not carro.empty

aba_resumo, aba_dias, aba_carro, aba_notas, aba_download = st.tabs([
    "Resumo", "Dias da semana", "Carro estimado", "Notas", "Downloads"
])

with aba_resumo:
    total_contratos = contratos["FATURAMENTO"].sum() if "FATURAMENTO" in contratos.columns else 0
    qtd_notas = int(notas["ORDEM_DE_SERVICO"].nunique()) if "ORDEM_DE_SERVICO" in notas.columns else 0

    if mostrar_carro:
        total_carro_min = carro["FATURAMENTO_MIN"].sum() if "FATURAMENTO_MIN" in carro.columns else 0
        total_carro_max = carro["FATURAMENTO_MAX"].sum() if "FATURAMENTO_MAX" in carro.columns else 0
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
    if not contratos.empty:
        resumo_contrato = contratos.groupby("CONTRATO", as_index=False)["FATURAMENTO"].sum()
        st.bar_chart(resumo_contrato, x="CONTRATO", y="FATURAMENTO")

        st.markdown("**Resumo com corte + religue**")
        tabela_resumo = contratos.pivot_table(
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

        st.markdown("**Detalhamento**")
        st.dataframe(formatar_tabela(contratos), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum contrato de disjuntor selecionado.")

    if mostrar_carro:
        st.subheader("Contrato do carro selecionado")
        tabela_carro_resumo = carro.copy()
        st.dataframe(formatar_tabela(tabela_carro_resumo), use_container_width=True, hide_index=True)

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
        tabela = tabela[["CONTRATO", "SEMANA_INICIO"] + colunas_dias + ["Total semana"]]
        st.dataframe(formatar_tabela(tabela), use_container_width=True, hide_index=True)

        por_dia = dias.groupby("DIA_SEMANA", as_index=False)["FATURAMENTO"].sum()
        por_dia["ordem"] = por_dia["DIA_SEMANA"].map({d: i for i, d in enumerate(ORDEM_DIAS)})
        por_dia = por_dia.sort_values("ordem")
        st.bar_chart(por_dia, x="DIA_SEMANA", y="FATURAMENTO")
    else:
        st.info("Nenhum contrato de disjuntor selecionado.")

with aba_carro:
    st.subheader("Contrato do carro — estimativa")
    if not carro.empty:
        c1, c2 = st.columns(2)
        c1.metric("Mínimo estimado", dinheiro(carro["FATURAMENTO_MIN"].sum()))
        c2.metric("Máximo estimado", dinheiro(carro["FATURAMENTO_MAX"].sum()))
        st.dataframe(formatar_tabela(carro), use_container_width=True, hide_index=True)
    else:
        st.info("Selecione 'Contrato Carro STC estimado' no filtro lateral para visualizar.")

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
        st.info("Selecione o contrato do carro no filtro lateral para visualizar os dias.")

with aba_notas:
    st.subheader("Consulta de notas")
    if not notas.empty:
        grupo = st.selectbox("Grupo de nota", ["Todos"] + sorted(notas.get("GRUPO_NOTA", pd.Series(dtype=str)).dropna().unique().tolist()))
        qtd_exec = st.selectbox("Quantidade de executores", ["Todos", 1, 2])
        df_notas = notas.copy()
        if grupo != "Todos" and "GRUPO_NOTA" in df_notas.columns:
            df_notas = df_notas[df_notas["GRUPO_NOTA"] == grupo]
        if qtd_exec != "Todos" and "QTD_EXECUTORES" in df_notas.columns:
            df_notas = df_notas[df_notas["QTD_EXECUTORES"] == qtd_exec]
        st.dataframe(df_notas.head(2000), use_container_width=True, hide_index=True)
        st.caption("Mostrando até 2000 linhas para não deixar o painel pesado.")
    else:
        st.info("Base de notas não encontrada.")

with aba_download:
    st.subheader("Arquivos carregados")
    for chave, nome in ARQUIVOS.items():
        caminho = caminho_arquivo(nome)
        if caminho:
            with open(caminho, "rb") as f:
                st.download_button(f"Baixar {caminho.name}", f, file_name=caminho.name)
