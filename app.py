import streamlit as st
import tempfile
import os
from rn_fill import fill_rn_docx

st.set_page_config(page_title="Gerador RN – Allianz", layout="wide")
st.title("Gerador de RN – Modelo Word")

with st.form("rn_form"):
    col1, col2 = st.columns(2)

    with col1:
        rn = st.text_input("RN")
        destinatario = st.text_input("Destinatário")
        segurado = st.text_input("Segurado")
        cnpj = st.text_input("CNPJ")
        participacao = st.text_input("Participação (%)", placeholder="40,00")
        assinatura = st.text_input("Assinatura")

    with col2:
        subscritor = st.text_input("Subscritor")
        filial = st.text_input("Comercial / Filial")
        email_user = st.text_input("E-mail (antes do @allianz.com.br)")
        data_doc = st.text_input("Data", placeholder="dd/mm/aaaa")
        vig_ini = st.text_input("Vigência início")
        vig_fim = st.text_input("Vigência fim")

    submit = st.form_submit_button("Gerar Word")

if submit:
    data = {
        "rn": rn,
        "destinatario": destinatario,
        "subscritor": subscritor,
        "filial": filial,
        "email_user": email_user,
        "data": data_doc,
        "segurado": segurado,
        "cnpj": cnpj,
        "participacao_pct": participacao,
        "assinatura_nome": assinatura,
        "vig_inicio": vig_ini,
        "vig_fim": vig_fim,
        "locais": [],
        "vr": []
    }

    with tempfile.TemporaryDirectory() as tmp:
        out_docx = os.path.join(tmp, "RN_preenchido.docx")
        fill_rn_docx("MODELO RN (1).docx", out_docx, data)

        with open(out_docx, "rb") as f:
            st.download_button(
                "⬇️ Baixar Word preenchido",
                data=f,
                file_name="RN_preenchido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )