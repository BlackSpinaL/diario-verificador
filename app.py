import streamlit as st
import camelot
import pandas as pd

st.title("📑 Verificador de Notas - Diário Escolar")

uploaded_files = st.file_uploader(
    "Faça upload dos diários em PDF",
    type="pdf",
    accept_multiple_files=True
)

# Colunas padrão que queremos verificar
colunas_padrao = ["MATRICULA", "NOME DO ALUNO", "AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]
colunas_selecionadas = st.multiselect("Selecione as colunas para verificar:", colunas_padrao[2:])  # só notas

# Dicionário de equivalências (ajuste automático)
mapa_colunas = {
    "MATRÍCULA": "MATRICULA",
    "MATRICULA": "MATRICULA",
    "NOME": "NOME DO ALUNO",
    "NOME DO ALUNO": "NOME DO ALUNO",
    "ALUNO": "NOME DO ALUNO",
    "AV1/AP1": "AP1/AV1",
    "AP1": "AP1/AV1",
    "AS/AP2": "AP2/AV2",
    "AP2": "AP2/AV2",
    "TOTAL": "TOTAL PARCIAL",
    "FINAL": "FINAL"
}

resultados = []

if uploaded_files and colunas_selecionadas:
    for file in uploaded_files:
        tables = camelot.read_pdf(file, pages="all")

        for t in tables:
            df = t.df
            df.columns = df.iloc[0]  # primeira linha como cabeçalho
            df = df.drop(0)

            # Normalizar cabeçalhos
            df.columns = [mapa_colunas.get(col.strip().upper(), col.strip().upper()) for col in df.columns]

            # Verificação de células vazias
            for idx, row in df.iterrows():
                for col in colunas_selecionadas:
                    if col in df.columns:
                        if row[col] == "" or pd.isna(row[col]):
                            resultados.append({
                                "Arquivo": file.name,
                                "Matrícula": row.get("MATRICULA", ""),
                                "Nome": row.get("NOME DO ALUNO", ""),
                                "Coluna faltando": col
                            })

    # Mostrar relatório consolidado
    if resultados:
        st.warning("⚠️ Foram encontradas pendências:")
        resultado_df = pd.DataFrame(resultados)
        st.dataframe(resultado_df)

        # Exportar para Excel
        st.download_button(
            label="📥 Baixar relatório em Excel",
            data=resultado_df.to_excel(index=False, engine="openpyxl"),
            file_name="pendencias_notas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("✅ Todos os diários estão completos nas colunas selecionadas!")
