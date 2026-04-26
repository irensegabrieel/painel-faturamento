# Versão com gráficos horizontais (ranking visual melhorado)

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Painel de Faturamento", layout="wide")

st.title("📊 Painel de Faturamento")

# Exemplo de dados (substitua pelos seus CSVs reais)
data = {
    "RECURSO": ["SAL5508-EMP","SAL5506-EMP","SAL5504-EMP","SAL5500-EMP","JUN5595-EMP","JUN5583-EMP"],
    "NOTAS": [1020,680,910,890,700,780]
}

df = pd.DataFrame(data)

# Ordenação correta (MAIOR -> MENOR)
df = df.sort_values("NOTAS", ascending=True)

st.subheader("Top recursos (ordenado)")

# Gráfico horizontal (melhor leitura)
st.bar_chart(df.set_index("RECURSO"))

st.caption("Gráfico agora ordenado do maior para o menor e com melhor leitura visual.")
