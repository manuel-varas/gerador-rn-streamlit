import streamlit as st
import re
from datetime import date
import tempfile
import os
from rn_fill import fill_rn_docx

st.set_page_config(page_title="Gerador RN – Allianz", layout="wide")
st.title("Gerador de RN – Riscos Nomeados")

# ---------- FUNÇÕES AUXILIARES ----------
def format_cnpj(cnpj):
    nums = re.sub(r"\D", "", cnpj)
    if len(nums) == 14:
        return f"{nums[:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:]}"
    return cnpj

# ---------- FORMULÁRIO ----------
with st.form("rn_form"):
    col1, col2 = st.columns(2)

    with col1:
        rn = st.text_input("RN")
        destinatario = st.text_input("Destinatário")
        remetente = st.text_input("Remetente / From")
        segurado = st.text_input("Segurado")
        cnpj_raw = st.text_input("CNPJ")

    with col2:
        subscritor = st.text_input("Subscritor")
        filial = st.text_input("Comercial / Filial")
        email_user = st.text_input("E-mail (antes do @allianz.com.br)")
        data_doc = st.date_input("Data", value=date.today())
        paginas = st.number_input("Páginas", value=13)

    cotacao = st.text_input("Cotação", value="Riscos Nomeados")

    submit = st.form_submit_button("Gerar Word")

# ---------- PROCESSAMENTO ----------
if submit:
    data = {
        "rn": rn,
        "destinatario": destinatario,
        "remetente": remetente,
        "subscritor": subscritor,
        "filial": filial,
        "email_user": email_user,
        "data": data_doc.strftime("%d/%m/%Y"),
        "paginas": str(paginas),
        "segurado": segurado,
        "cnpj": format_cnpj(cnpj_raw),
        "cotacao": cotacao
    }

    with tempfile.TemporaryDirectory() as tmp:
        output_path = os.path.join(tmp, "RN_preenchido.docx")

        fill_rn_docx(
            template_path="MODELO RN (1).docx",
            output_path=output_path,
            data=data
        )

        with open(output_path, "rb") as f:
            st.download_button(
                label="⬇️ Baixar RN preenchido",
                data=f,
                file_name="RN_preenchido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
