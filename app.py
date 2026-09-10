import streamlit as st
import camelot
import pandas as pd
from io import BytesIO
import re

st.title("📑 Verificador de Notas - Diário Escolar")

uploaded_files = st.file_uploader(
    "Faça upload dos diários em PDF",
    type="pdf",
    accept_multiple_files=True
)

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
    "TE": "TE",
    "AE": "AE",
    "ND": "ND",
    "TOTAL": "TOTAL PARCIAL",
    "TOTAL PARCIAL": "TOTAL PARCIAL",
    "FINAL": "FINAL"
}

colunas_disponiveis = set()

if uploaded_files:
    for file in uploaded_files:
        tables = camelot.read_pdf(file, pages="1")
        for t in tables:
            df = t.df
            df.columns = df.iloc[0]
            df = df.drop(0)
            df.columns = [mapa_colunas.get(col.strip().upper(), col.strip().upper()) for col in df.columns]
            colunas_disponiveis.update(df.columns)

colunas_selecionadas = st.multiselect("Selecione as colunas para verificar:", sorted(colunas_disponiveis))

if st.button("▶️ Rodar verificação") and uploaded_files and colunas_selecionadas:
    resultados_por_turma = {}

    for file in uploaded_files:
        turma_resultados = []
        professor_nome = "Professor não identificado"

        header_tables = camelot.read_pdf(file, pages="1")
        for ht in header_tables:
            texto_cabecalho = " ".join(ht.df.values.flatten())
            match_prof = re.search(r'PROFESSOR\s+([A-Z\s]+)', texto_cabecalho, re.IGNORECASE)
            if match_prof:
                professor_nome = match_prof.group(1).title()
                break

        tables = camelot.read_pdf(file, pages="all")
        for t in tables:
            df = t.df
            df.columns = df.iloc[0]
            df = df.drop(0)
            df.columns = [mapa_colunas.get(col.strip().upper(), col.strip().upper()) for col in df.columns]

            for idx, row in df.iterrows():
                for col in colunas_selecionadas:
                    if col in df.columns:
                        if row[col] == "" or pd.isna(row[col]):
                            turma_resultados.append({
                                "Matrícula": row.get("MATRICULA", ""),
                                "Nome": row.get("NOME DO ALUNO", ""),
                                "Coluna faltando": col
                            })

        if turma_resultados:
            match = re.search(r'(\d{5})', file.name)
            turma_codigo = match.group(1) if match else file.name[:30]
            resultados_por_turma[turma_codigo] = {
                "professor": professor_nome,
                "df": pd.DataFrame(turma_resultados)
            }

    if resultados_por_turma:
        st.warning("⚠️ Foram encontradas pendências!")
        for turma, dados in resultados_por_turma.items():
            st.subheader(f"📘 Turma: {turma} — {dados['professor']}")
            st.dataframe(dados["df"])

        # Excel com abas por turma
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            for turma, dados in resultados_por_turma.items():
                sheet_name = f"{turma}_{dados['professor'][:15]}"
                dados["df"].to_excel(writer, sheet_name=sheet_name[:31], index=False)

        st.download_button(
            label="📥 Baixar relatório detalhado por turma (Excel)",
            data=output_excel.getvalue(),
            file_name="pendencias_por_turma.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # PDF com páginas por turma
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet

        output_pdf = BytesIO()
        doc = SimpleDocTemplate(output_pdf, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()

        for turma, dados in resultados_por_turma.items():
            elements.append(Paragraph(f"📘 Turma: {turma} — {dados['professor']}", styles['Heading2']))
            data = [dados["df"].columns.tolist()] + dados["df"].values.tolist()
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black)
            ]))
            elements.append(table)
            elements.append(Spacer(1, 20))

        doc.build(elements)

        st.download_button(
            label="📥 Baixar relatório detalhado por turma (PDF)",
            data=output_pdf.getvalue(),
            file_name="pendencias_por_turma.pdf",
            mime="application/pdf"
        )
    else:
        st.success("✅ Todos os diários estão completos nas colunas selecionadas!")
