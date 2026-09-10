import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.title("📑 Verificador de Notas - Diário Escolar")

uploaded_files = st.file_uploader(
    "Faça upload dos diários em PDF",
    type="pdf",
    accept_multiple_files=True
)

# Colunas padrão que queremos verificar
colunas_padrao = ["AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]

# Dicionário de equivalências para padronizar cabeçalhos
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

colunas_selecionadas = st.multiselect(
    "Selecione as colunas para verificar:",
    colunas_padrao
)

# Botão para rodar a verificação
if st.button("▶️ Rodar verificação") and uploaded_files and colunas_selecionadas:
    resultados_por_turma = {}

    for file in uploaded_files:
        turma_resultados = []
        professor_nome = "Professor não identificado"
        
        # Usando pdfplumber para extrair as tabelas
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # Extrai a tabela da página
                table = page.extract_table()
                
                if not table:
                    continue
                
                # A primeira linha geralmente é o cabeçalho
                # Vamos tentar encontrar a linha que contém "MATRICULA" ou "NOME"
                header_index = -1
                for i, row in enumerate(table):
                    row_str = " ".join([str(cell) for cell in row if cell])
                    if "MATRICULA" in row_str.upper() or "NOME" in row_str.upper():
                        header_index = i
                        break
                
                if header_index == -1:
                    continue # Pula se não achar o cabeçalho
                
                # Pega o cabeçalho e limpa
                raw_header = table[header_index]
                clean_header = []
                for col in raw_header:
                    if col:
                        # Remove quebras de linha e espaços extras
                        c = str(col).replace("\n", " ").strip().upper()
                        # Aplica o mapa de colunas
                        c_mapped = mapa_colunas.get(c, c)
                        clean_header.append(c_mapped)
                    else:
                        clean_header.append("VAZIO")
                
                # Pega os dados a partir da linha seguinte ao cabeçalho
                data_rows = table[header_index + 1:]
                
                # Cria o DataFrame
                df = pd.DataFrame(data_rows, columns=clean_header)
                
                # Substitui strings vazias ou None por NaN para facilitar
                df = df.replace(["", "None", "none", "nan"], pd.NA)
                
                # Tenta capturar o nome do professor no texto da página
                texto_pagina = page.extract_text()
                if texto_pagina:
                    match_prof = re.search(r'PROFESSOR\s+([A-Z\s]+)', texto_pagina, re.IGNORECASE)
                    if match_prof and professor_nome == "Professor não identificado":
                        professor_nome = match_prof.group(1).strip().title()

                # Itera sobre as linhas do DataFrame
                for idx, row in df.iterrows():
                    # Verifica se a linha é válida (tem matrícula ou nome)
                    matricula = str(row.get("MATRICULA", "")).strip()
                    nome = str(row.get("NOME DO ALUNO", "")).strip()
                    
                    if not matricula or not nome or matricula == "None":
                        continue

                    # Verifica as colunas selecionadas
                    for col in colunas_selecionadas:
                        if col in df.columns:
                            valor_bruto = row[col]
                            
                            # Limpeza do valor: remove espaços, converte vírgula para ponto
                            valor_str = str(valor_bruto).strip().replace(",", ".").replace(" ", "")
                            
                            try:
                                numero = float(valor_str)
                            except (ValueError, TypeError):
                                numero = None

                            # Considerar vazio, NaN ou zero como pendência
                            # Adicionado "00.00" e "0.00" na lista
                            if valor_str in ["", "0", "00", "0.0", "0.00", "00.00", "nan", "<NA>"] or pd.isna(valor_bruto) or (numero is not None and numero == 0.0):
                                turma_resultados.append({
                                    "Matrícula": matricula,
                                    "Nome": nome,
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
                # Limita o nome da aba a 31 caracteres (limite do Excel)
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
            
            # Converte o DataFrame para uma lista de listas para o ReportLab
            data = [dados["df"].columns.tolist()] + dados["df"].values.tolist()
            
            # Cria a tabela com estilo
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('FONTSIZE', (0,0), (-1,-1), 8) # Reduz a fonte para caber melhor
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
