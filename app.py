import streamlit as st
import pdfplumber
import pandas as pd

st.title("📑 Verificador de Notas - Diário Escolar")

# Upload de múltiplos PDFs
uploaded_files = st.file_uploader(
    "Faça upload dos diários em PDF",
    type="pdf",
    accept_multiple_files=True
)

# Seleção das colunas que deseja verificar
colunas = ["AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]
colunas_selecionadas = st.multiselect("Selecione as colunas para verificar:", colunas)

resultados = []

if uploaded_files and colunas_selecionadas:
    for file in uploaded_files:
        with pdfplumber.open(file) as pdf:
            texto = ""
            for page in pdf.pages:
                texto += page.extract_text() + "\n"

        # ⚠️ Aqui você deve adaptar a extração para o formato real do seu PDF
        # Exemplo fictício de dados simulados:
        dados = {
            "Matrícula": ["122328", "117087", "144298"],
            "Nome": ["Álvaro Tostes", "Ana Laura Almeida", "Ana Laura Abreu"],
            "AP1/AV1": ["", "30", "25"],
            "AP2/AV2": ["20", "", "28"],
            "FINAL": ["", "60", "53"]
        }
        df = pd.DataFrame(dados)

        # Verificação de células vazias
        for idx, row in df.iterrows():
            for col in colunas_selecionadas:
                if row[col] == "" or pd.isna(row[col]):
                    resultados.append({
                        "Arquivo": file.name,
                        "Matrícula": row["Matrícula"],
                        "Nome": row["Nome"],
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
