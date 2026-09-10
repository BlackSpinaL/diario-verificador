import streamlit as st
import pdfplumber
import pandas as pd

st.title("📑 Verificador de Notas - Diário Escolar")

# Upload do PDF
uploaded_file = st.file_uploader("Faça upload do diário em PDF", type="pdf")

# Seleção das colunas que deseja verificar
colunas = ["AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]
colunas_selecionadas = st.multiselect("Selecione as colunas para verificar:", colunas)

if uploaded_file is not None and colunas_selecionadas:
    with pdfplumber.open(uploaded_file) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"

    # Simulação: transformar texto em tabela (aqui você ajusta conforme o formato do PDF)
    # Exemplo fictício: criar DataFrame com colunas
    dados = {
        "Matrícula": ["122328", "117087", "144298"],
        "Nome": ["Álvaro Tostes", "Ana Laura Almeida", "Ana Laura Abreu"],
        "AP1/AV1": ["", "30", "25"],
        "AP2/AV2": ["20", "", "28"],
        "FINAL": ["", "60", "53"]
    }
    df = pd.DataFrame(dados)

    # Verificação de células vazias
    pendencias = []
    for idx, row in df.iterrows():
        for col in colunas_selecionadas:
            if row[col] == "" or pd.isna(row[col]):
                pendencias.append({
                    "Matrícula": row["Matrícula"],
                    "Nome": row["Nome"],
                    "Coluna faltando": col
                })

    # Mostrar resultado
    if pendencias:
        st.warning("⚠️ Foram encontradas pendências:")
        st.dataframe(pd.DataFrame(pendencias))
    else:
        st.success("✅ Todas as colunas selecionadas estão preenchidas!")
