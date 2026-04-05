import streamlit as st
import re
from datetime import date
import tempfile
import os
from docx import Document

# =============================
# CONFIG STREAMLIT
# =============================
st.set_page_config(page_title="Gerador RN - Allianz", layout="wide")
st.title("Gerador de RN - Modelo Word")
st.success("✅ App carregado com sucesso")

TEMPLATE = "MODELO RN (1).docx"

# =============================
# FUNÇÕES AUXILIARES
# =============================
def format_cnpj(cnpj: str) -> str:
    nums = re.sub(r"\D", "", cnpj or "")
    if len(nums) == 14:
        return f"{nums[:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:]}"
    return cnpj


def set_cell_text(cell, text, paragraph_index=0):
    """
    ESCREVE APENAS NA CÉLULA DA DIREITA
    Preserva formatação e NUNCA toca na coluna esquerda
    """
    while len(cell.paragraphs) <= paragraph_index:
        cell.add_paragraph("")

    p = cell.paragraphs[paragraph_index]

    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def replace_in_cell(cell, old, new):
    """
    Substitui texto dentro da célula (ex: xxxx.xxxx do e-mail)
    SEM apagar estrutura
    """
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)


def find_table(doc, anchor_text):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if anchor_text.upper() in cell.text.upper():
                    return table
    return None


def find_row(table, left_label):
    for i, row in enumerate(table.rows):
        if left_label.upper() in row.cells[0].text.upper():
            return i
    return None


# =============================
# FORMULÁRIO
# =============================
if not os.path.exists(TEMPLATE):
    st.error(f"Arquivo {TEMPLATE} não encontrado no repositório")
    st.stop()

with st.form("rn_form"):
    col1, col2 = st.columns(2)

    with col1:
        rn = st.text_input("PROC. Nº (RN)")
        destinatario = st.text_input("DESTINATÁRIO / To")
        subscritor = st.text_input("REMETENTE - Subscritor")
        filial = st.text_input("REMETENTE - Comercial / Filial")
        segurado = st.text_input("SEGURADO")
        cnpj_raw = st.text_input("CNPJ")

    with col2:
        email_user = st.text_input("E-mail (antes do @allianz.com.br)")
        data_doc = st.date_input("DATA / DATE", value=date.today())
        paginas = st.number_input("PÁGINAS / PAGES", value=13, min_value=1)
        cotacao = st.text_input("COTAÇÃO", value="Riscos Nomeados")

    submit = st.form_submit_button("Gerar Word")


# =============================
# PROCESSAMENTO
# =============================
if submit:
    doc = Document(TEMPLATE)

    data = {
        "rn": rn,
        "destinatario": destinatario,
        "subscritor": subscritor,
        "filial": filial,
        "email_user": email_user,
        "data": data_doc.strftime("%d/%m/%Y"),
        "paginas": str(paginas),
        "cotacao": cotacao,
        "segurado": segurado,
        "cnpj": format_cnpj(cnpj_raw),
    }

    # ========= TABELA DA CAPA =========
    cover = find_table(doc, "PROC. Nº")
    if cover:
        i = find_row(cover, "PROC. Nº")
        if i is not None and data["rn"]:
            set_cell_text(cover.cell(i, 1), f"RN - {data['rn']}")

        i = find_row(cover, "DESTINATÁRIO")
        if i is not None and data["destinatario"]:
            set_cell_text(cover.cell(i, 1), data["destinatario"])

        i = find_row(cover, "REMETENTE")
        if i is not None:
            if data["subscritor"]:
                set_cell_text(cover.cell(i, 1), data["subscritor"], 0)
            if data["filial"]:
                set_cell_text(cover.cell(i, 1), data["filial"], 1)

        i = find_row(cover, "DEPTO/DIVISION")
        if i is not None and data["email_user"]:
            replace_in_cell(cover.cell(i, 1), "xxxx.xxxx", data["email_user"])

        i = find_row(cover, "DATA/DATE")
        if i is not None:
            set_cell_text(cover.cell(i, 1), data["data"])

        i = find_row(cover, "PÁGINAS/PAGES")
        if i is not None:
            set_cell_text(
                cover.cell(i, 1),
                f"{data['paginas']} (incluindo esta capa/including the cover page)"
            )

    # ========= TABELA DE COTAÇÃO =========
    quote = find_table(doc, "COTAÇÃO")
    if quote:
        i = find_row(quote, "COTAÇÃO")
        if i is not None:
            set_cell_text(quote.cell(i, 1), data["cotacao"])

        i = find_row(quote, "SEGURADO")
        if i is not None and data["segurado"]:
            set_cell_text(quote.cell(i, 1), data["segurado"])

        i = find_row(quote, "CNPJ")
        if i is not None and data["cnpj"]:
            set_cell_text(quote.cell(i, 1), data["cnpj"])

    # ========= SALVAR =========
    with tempfile.TemporaryDirectory() as tmp:
        output_path = os.path.join(tmp, "RN_preenchido.docx")
        doc.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                "⬇️ Baixar RN preenchido",
                data=f,
                file_name="RN_preenchido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
