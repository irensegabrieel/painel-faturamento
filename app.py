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

# Atualiza a página automaticamente a cada 60 segundos.
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


@st.cache_data(ttl=60, show_spinner=False)
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
        if "FATURAMENTO" in col or col in ["CORTE", "RELIGUE", "TOTAL", "MÍNIMO", "MÁXIMO", "VALOR"]:
            df2[col] = df2[col].apply(dinheiro)
        elif col in ["QTD_NOTAS", "NOTAS", "CORTES", "RELIGUES", "TOTAL_NOTAS"]:
            df2[col] = df2[col].apply(numero)

    return df2


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
            faturamento = {"CORTE": 11.34, "RELIGUE": 12.68}.get(grupo, 0.0)
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


def preparar_comparativo_mensal(notas):
    parcial = preparar_parcial_do_dia(notas)

    if parcial.empty:
        return pd.DataFrame(), pd.DataFrame()

    parcial = parcial.copy()
    parcial["ANO_MES"] = parcial["DATA_DT"].dt.to_period("M").astype(str)
    parcial["MES"] = parcial["DATA_DT"].dt.strftime("%m/%Y")
    parcial["MES_DT"] = parcial["DATA_DT"].dt.to_period("M").dt.to_timestamp()

    mensal = (
        parcial.groupby(["MES_DT", "ANO_MES", "MES", "CONTRATO"], dropna=False)
        .agg(
            TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            FATURAMENTO=("FATURAMENTO", "sum"),
            FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
            FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
        )
        .reset_index()
        .sort_values("MES_DT")
    )

    mensal_equipes = (
        parcial.groupby(["MES_DT", "ANO_MES", "MES", "CONTRATO", "RECURSO"], dropna=False)
        .agg(
            TOTAL_NOTAS=("ORDEM_DE_SERVICO", "nunique"),
            CORTES=("EH_CORTE", "sum"),
            RELIGUES=("EH_RELIGUE", "sum"),
            FATURAMENTO=("FATURAMENTO", "sum"),
            FATURAMENTO_MIN=("FATURAMENTO_MIN", "sum"),
            FATURAMENTO_MAX=("FATURAMENTO_MAX", "sum"),
        )
        .reset_index()
        .sort_values(["MES_DT", "CONTRATO", "RECURSO"])
    )

    return mensal, mensal_equipes


def variacao_percentual(atual, anterior):
    try:
        atual = float(atual)
        anterior = float(anterior)
        if anterior == 0:
            return "-"
        return f"{((atual - anterior) / anterior) * 100:.1f}%".replace(".", ",")
    except Exception:
        return "-"


# ==============================
# CARREGAMENTO
# ==============================

bases, faltando = carregar_bases()

st.title("📊 Painel de Faturamento")
st.caption("Painel atualizado automaticamente a cada 60 segundos.")

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
st.sidebar.markdown("**Selecionado:**")
st.sidebar.info(contrato_escolhido)

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

aba_resumo, aba_parcial, aba_comparativo, aba_dias, aba_carro, aba_notas, aba_download = st.tabs([
    "Resumo", "Parcial do dia", "Comparativo mensal", "Dias da semana", "Carro estimado", "Notas", "Downloads"
])

# ==============================
# ABA RESUMO
# ==============================

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
        st.info("Nenhum contrato selecionado.")

    if mostrar_carro:
        st.subheader("Contrato do carro selecionado")
        st.dataframe(formatar_tabela(carro), use_container_width=True, hide_index=True)

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

                if contrato_escolhido == "Contrato Carro STC estimado":
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

                if contrato_escolhido == "Contrato Carro STC estimado":
                    tabela_equipe = resumo_equipe[[
                        "RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO_MIN", "FATURAMENTO_MAX"
                    ]]
                else:
                    tabela_equipe = resumo_equipe[[
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
# ABA COMPARATIVO MENSAL
# ==============================

with aba_comparativo:
    st.subheader("Comparativo mensal")
    st.caption("Compara o mês escolhido com o mês anterior, usando a base acumulada de notas.")

    mensal, mensal_equipes = preparar_comparativo_mensal(notas)

    if mensal.empty:
        st.info("Ainda não há dados suficientes para montar o comparativo mensal.")
    else:
        if contrato_escolhido != "Todos" and "CONTRATO" in mensal.columns:
            mensal = mensal[mensal["CONTRATO"] == contrato_escolhido]
            mensal_equipes = mensal_equipes[mensal_equipes["CONTRATO"] == contrato_escolhido]

        if mensal.empty:
            st.info("Não há dados mensais para o contrato selecionado.")
        else:
            meses_disponiveis = (
                mensal[["MES", "ANO_MES", "MES_DT"]]
                .drop_duplicates()
                .sort_values("MES_DT", ascending=False)
            )

            opcoes_meses = meses_disponiveis["MES"].tolist()
            mes_escolhido = st.selectbox("Escolha o mês para comparar", opcoes_meses, index=0)

            mes_dt = meses_disponiveis.loc[meses_disponiveis["MES"] == mes_escolhido, "MES_DT"].iloc[0]
            mes_anterior_dt = mes_dt - pd.DateOffset(months=1)
            mes_anterior_label = mes_anterior_dt.strftime("%m/%Y")

            atual = mensal[mensal["MES_DT"] == mes_dt].copy()
            anterior = mensal[mensal["MES_DT"] == mes_anterior_dt].copy()

            total_atual = atual["FATURAMENTO"].sum()
            total_anterior = anterior["FATURAMENTO"].sum() if not anterior.empty else 0
            notas_atual = int(atual["TOTAL_NOTAS"].sum())
            notas_anterior = int(anterior["TOTAL_NOTAS"].sum()) if not anterior.empty else 0
            cortes_atual = int(atual["CORTES"].sum())
            cortes_anterior = int(anterior["CORTES"].sum()) if not anterior.empty else 0
            religues_atual = int(atual["RELIGUES"].sum())
            religues_anterior = int(anterior["RELIGUES"].sum()) if not anterior.empty else 0

            st.markdown(f"**Comparando:** {mes_escolhido} x {mes_anterior_label}")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Faturamento", dinheiro(total_atual), variacao_percentual(total_atual, total_anterior))
            c2.metric("Notas", numero(notas_atual), variacao_percentual(notas_atual, notas_anterior))
            c3.metric("Cortes", numero(cortes_atual), variacao_percentual(cortes_atual, cortes_anterior))
            c4.metric("Religues", numero(religues_atual), variacao_percentual(religues_atual, religues_anterior))

            st.markdown("**Evolução mês a mês**")

            tabela_mensal = mensal.copy().sort_values(["MES_DT", "CONTRATO"])
            colunas_mensal = [
                "MES", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO"
            ]
            st.dataframe(
                formatar_tabela(tabela_mensal[colunas_mensal]),
                use_container_width=True,
                hide_index=True,
            )

            grafico = (
                mensal.groupby("MES_DT", as_index=False)["FATURAMENTO"]
                .sum()
                .sort_values("MES_DT")
            )
            grafico["MES"] = grafico["MES_DT"].dt.strftime("%m/%Y")
            st.bar_chart(grafico, x="MES", y="FATURAMENTO")

            st.markdown("**Comparativo por contrato no mês escolhido**")
            por_contrato = atual[["CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO"]].copy()
            st.dataframe(formatar_tabela(por_contrato), use_container_width=True, hide_index=True)

            st.markdown("**Equipes no mês escolhido**")
            equipes_mes = mensal_equipes[mensal_equipes["MES_DT"] == mes_dt].copy()
            equipes_mes = equipes_mes.sort_values(["CONTRATO", "RECURSO"])
            colunas_equipes = ["RECURSO", "CONTRATO", "TOTAL_NOTAS", "CORTES", "RELIGUES", "FATURAMENTO"]
            st.dataframe(
                formatar_tabela(equipes_mes[colunas_equipes]),
                use_container_width=True,
                hide_index=True,
            )

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

        st.dataframe(df_notas.head(2000), use_container_width=True, hide_index=True)
        st.caption("Mostrando até 2000 linhas para não deixar o painel pesado.")
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
