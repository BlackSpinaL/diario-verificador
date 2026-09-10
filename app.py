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

# Colunas padrão que queremos verificar (nesta ordem, sempre as 7 últimas
# colunas da tabela principal de notas do diário)
colunas_padrao = ["AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]

colunas_selecionadas = st.multiselect(
    "Selecione as colunas para verificar:",
    colunas_padrao
)


def extrair_tabela_notas(df_bruto):
    """
    Recebe um DataFrame cru extraído pelo camelot (flavor='stream') e devolve
    um DataFrame já limpo com colunas MATRICULA, NOME DO ALUNO + colunas_padrao,
    contendo só as linhas de alunos.

    IMPORTANTE: no PDF do diário (Colégio Tiradentes - PMMG), o cabeçalho da
    tabela de notas é multi-linha e o camelot (flavor=stream) o quebra em
    várias linhas diferentes do DataFrame - nunca fica tudo alinhado numa
    única linha 0. Por isso, usar `df.columns = df.iloc[0]` (como no código
    original) pega uma linha de cabeçalho "errada" (normalmente um fragmento
    do topo da página), e nenhuma das colunas de notas (AP1/AV1, TOTAL
    PARCIAL, FINAL etc.) é encontrada depois. Resultado: a checagem
    "if col in df.columns" nunca é verdadeira para NINGUÉM, e o app relata
    "tudo completo" mesmo quando faltam notas (como a da aluna Vitória).

    A extração abaixo NÃO depende do texto do cabeçalho. Em vez disso:
      - identifica as linhas de alunos pela numeração sequencial (001, 002...)
        que sempre aparece na 1ª coluna;
      - usa a POSIÇÃO fixa das colunas: coluna 1 = matrícula, coluna 2 = nome,
        e as 7 ÚLTIMAS colunas da tabela = AP1/AV1, AP2/AV2, TE, AE, ND,
        TOTAL PARCIAL, FINAL (nessa ordem), que é o layout fixo do diário.
    """
    linhas_alunos = []
    for _, row in df_bruto.iterrows():
        primeira_celula = str(row.iloc[0]).strip()
        if re.match(r'^\d{3}$', primeira_celula):  # ex.: 001, 002, ...
            linhas_alunos.append(row)

    if not linhas_alunos:
        return None  # não é a tabela principal de notas

    n_cols = df_bruto.shape[1]
    if n_cols < 9:  # não tem colunas suficientes p/ conter as 7 notas + matricula + nome
        return None

    registros = []
    for row in linhas_alunos:
        matricula = str(row.iloc[1]).strip()
        nome = str(row.iloc[2]).strip()
        notas = row.iloc[n_cols - 7:n_cols].tolist()  # 7 últimas colunas
        registro = {"MATRICULA": matricula, "NOME DO ALUNO": nome}
        for nome_coluna, valor in zip(colunas_padrao, notas):
            registro[nome_coluna] = valor
        registros.append(registro)

    return pd.DataFrame(registros)


def eh_tabela_de_notas(df_bruto):
    """Confirma que a tabela extraída é a tabela principal de notas
    (e não, por ex., a tabela de faltas por etapa da página 2),
    procurando o texto 'AV1/AP1' em qualquer célula."""
    texto = " ".join(df_bruto.astype(str).values.flatten()).upper()
    return "AV1/AP1" in texto or "AP1/AV1" in texto


if st.button("▶️ Rodar verificação") and uploaded_files and colunas_selecionadas:
    resultados_por_turma = {}

    for file in uploaded_files:
        turma_resultados = []
        professor_nome = "Professor não identificado"

        # Nome do professor: procurar em todas as tabelas da página 1
        header_tables = camelot.read_pdf(file, pages="1", flavor="stream", strip_text="\n")
        for ht in header_tables:
            texto_cabecalho = " ".join(ht.df.astype(str).values.flatten())
            match_prof = re.search(r'PROFESSOR\s+([A-Z\s]+)', texto_cabecalho, re.IGNORECASE)
            if match_prof:
                professor_nome = match_prof.group(1).title()
                break

        tables = camelot.read_pdf(file, pages="all", flavor="stream", strip_text="\n")
        for t in tables:
            df_bruto = t.df
            if not eh_tabela_de_notas(df_bruto):
                continue  # pula tabelas que não são a tabela principal de notas

            df = extrair_tabela_notas(df_bruto)
            if df is None:
                continue

            for _, row in df.iterrows():
                for col in colunas_selecionadas:
                    valor = str(row[col]).strip().replace(",", ".")
                    try:
                        numero = float(valor)
                    except ValueError:
                        numero = None

                    if valor in ["", "0", "00", "0.0", "0,0"] or pd.isna(row[col]) or numero == 0.0:
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
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
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
