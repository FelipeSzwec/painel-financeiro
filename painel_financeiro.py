import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Título
st.title("📋 Painel de Controle Financeiro dos Filiados")

# Carrega a planilha
df = pd.read_excel("Controle Financeiro/Ff.xlsx")

# Renomeia colunas
df = df.rename(columns={
    'Filiados': 'Filiado',
    'Valor Taxa': 'Valor',
    'Valor Recebido': 'Valor Recebido',
    'Mês Referente': 'Mês Referente'
})

# Preenche valores nulos
df['Valor'] = df['Valor'].fillna(0)
df['Valor Recebido'] = df['Valor Recebido'].fillna(0)
df['Mês Referente'] = df['Mês Referente'].fillna("-")

# Mapeia meses
meses = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}
meses_nome = {v: k.capitalize() for k, v in meses.items()}
mes_atual = datetime.now().month
nome_mes_atual = meses_nome[mes_atual]

# Define o status de cada filiado
def classificar_status(row):
    mes_ref = str(row['Mês Referente']).strip().lower()
    if mes_ref in meses:
        mes_pago = meses[mes_ref]
        if mes_pago >= mes_atual - 1:
            return "✅ Em dia"
        else:
            return "⚠️ Atrasado"
    return "⏳ Pendente"

df["Status"] = df.apply(classificar_status, axis=1)

# Mostra resumo
st.subheader(f"📊 Resumo de {nome_mes_atual}:")
col1, col2, col3 = st.columns(3)
col1.metric("✅ Em dia", df[df['Status'] == '✅ Em dia'].shape[0])
col2.metric("⚠️ Atrasado", df[df['Status'] == '⚠️ Atrasado'].shape[0])
col3.metric("⏳ Pendentes", df[df['Status'] == '⏳ Pendente'].shape[0])

# Filtros
filtro_status = st.multiselect("Filtrar por Status", options=df["Status"].unique(), default=df["Status"].unique())
filtro_nome = st.text_input("🔍 Buscar Filiado")

df_filtrado = df[df["Status"].isin(filtro_status)]
if filtro_nome:
    df_filtrado = df_filtrado[df_filtrado["Filiado"].str.contains(filtro_nome, case=False)]

# Tabela com dados
st.subheader("📄 Lista de Filiados")
st.dataframe(df_filtrado, use_container_width=True)

# Exportar dados filtrados para Excel
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name='Relatório')
buffer.seek(0)

st.download_button(
    label="📥 Baixar planilha filtrada",
    data=buffer,
    file_name=f"Relatorio_{nome_mes_atual}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


