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

colunas_padrao = ["AP1/AV1", "AP2/AV2", "TE", "AE", "ND", "TOTAL PARCIAL", "FINAL"]

colunas_selecionadas = st.multiselect(
    "Selecione as colunas para verificar:",
    colunas_padrao
)


def extrair_tabela_notas(df_bruto):
    linhas_alunos = []
    for _, row in df_bruto.iterrows():
        primeira_celula = str(row.iloc[0]).strip()
        if re.match(r'^\d{3}$', primeira_celula):
            linhas_alunos.append(row)

    if not linhas_alunos:
        return None

    n_cols = df_bruto.shape[1]
    if n_cols < 9:
        return None

    registros = []
    for row in linhas_alunos:
        matricula = str(row.iloc[1]).strip()
        nome = str(row.iloc[2]).strip()
        notas = row.iloc[n_cols - 7:n_cols].tolist()
        registro = {"MATRICULA": matricula, "NOME DO ALUNO": nome}
        for nome_coluna, valor in zip(colunas_padrao, notas):
            registro[nome_coluna] = valor
        registros.append(registro)

    return pd.DataFrame(registros)


def eh_tabela_de_notas(df_bruto):
    texto = " ".join(df_bruto.astype(str).values.flatten()).upper()
    return "AV1/AP1" in texto or "AP1/AV1" in texto


def valor_em_branco(valor):
    if valor is None:
        return True
    s = str(valor).strip()
    if s == "":
        return True
    s_norm = s.replace(",", ".")
    try:
        num = float(s_norm)
        return num == 0.0
    except ValueError:
        return False


if st.button("▶️ Rodar verificação") and uploaded_files and colunas_selecionadas:
    resultados_por_turma = {}

    for file in uploaded_files:
        professor_nome = "Professor não identificado"
        disciplina_nome = "Disciplina não identificada"

        header_tables = camelot.read_pdf(file, pages="1", flavor="stream", strip_text="\n")
        for ht in header_tables:
            texto_cabecalho = " ".join(ht.df.astype(str).values.flatten())
            match_prof = re.search(r'PROFESSOR\s+([A-Z\s]+?)\s{3,}', texto_cabecalho, re.IGNORECASE)
            if match_prof:
                professor_nome = match_prof.group(1).title()
            match_disc = re.search(r'DISCIPLINA\s+(.+?)\s+MESES', texto_cabecalho, re.IGNORECASE)
            if match_disc:
                disciplina_nome = match_disc.group(1).strip().title()
            if match_prof and match_disc:
                break

        # Chave da turma: uma por arquivo
        match = re.search(r'(\d{5})', file.name)
        turma_codigo = match.group(1) if match else file.name[:30]

        # Inicializa a estrutura da turma uma única vez
        if turma_codigo not in resultados_por_turma:
            resultados_por_turma[turma_codigo] = {
                "professor": professor_nome,
                "linhas": []  # cada linha será um dict aluno+disciplina+colunas
            }

        tables = camelot.read_pdf(file, pages="all", flavor="stream", strip_text="\n")
        for t in tables:
            df_bruto = t.df
            if not eh_tabela_de_notas(df_bruto):
                continue

            df = extrair_tabela_notas(df_bruto)
            if df is None:
                continue

            for _, row in df.iterrows():
                # Monta a linha já no formato final: uma por aluno
                linha = {
                    "Matrícula": row.get("MATRICULA", ""),
                    "Nome": row.get("NOME DO ALUNO", ""),
                    "Disciplina": disciplina_nome,
                }
                tem_pendencia = False
                for col in colunas_selecionadas:
                    if valor_em_branco(row[col]):
                        linha[col] = "X"
                        tem_pendencia = True
                    else:
                        linha[col] = ""

                if tem_pendencia:
                    resultados_por_turma[turma_codigo]["linhas"].append(linha)

    # Remove turmas sem nenhuma pendência
    resultados_por_turma = {
        k: v for k, v in resultados_por_turma.items() if v["linhas"]
    }

    if resultados_por_turma:
        st.warning("⚠️ Foram encontradas pendências!")

        dfs_finais = {}
        for turma, dados in resultados_por_turma.items():
            df_turma = pd.DataFrame(dados["linhas"])
            ordem = ["Matrícula", "Nome", "Disciplina"] + colunas_selecionadas
            df_turma = df_turma.reindex(columns=ordem, fill_value="")
            dfs_finais[turma] = {"df": df_turma, **dados}

            st.subheader(f"📘 Turma: {turma} — {dados['professor']}")
            st.dataframe(df_turma)

        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            for turma, dados in dfs_finais.items():
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
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet

        output_pdf = BytesIO()
        doc = SimpleDocTemplate(output_pdf, pagesize=landscape(A4))
        elements = []
        styles = getSampleStyleSheet()

        for turma, dados in dfs_finais.items():
            elements.append(Paragraph(
                f"📘 Turma: {turma} — {dados['professor']}",
                styles['Heading2']
            ))
            data = [dados["df"].columns.tolist()] + dados["df"].values.tolist()
            table = Table(data, repeatRows=1)
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
