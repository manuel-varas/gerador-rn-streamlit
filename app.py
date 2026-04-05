# app.py
import streamlit as st
import re
from datetime import date
import tempfile
import os

st.set_page_config(page_title="Gerador RN – Allianz", layout="wide")
st.title("Gerador de RN – Modelo Word")
st.success("✅ App carregou (se você está vendo isso, o script está rodando).")

def format_cnpj(cnpj: str) -> str:
    nums = re.sub(r"\D", "", cnpj or "")
    if len(nums) == 14:
        return f"{nums[:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:]}"
    return cnpj

try:
    from rn_fill import fill_rn_docx
except Exception as e:
    st.error("Erro ao importar rn_fill.py")
    st.exception(e)
    st.stop()

TEMPLATE = "MODELO RN (1).docx"
if not os.path.exists(TEMPLATE):
    st.error(f"Não encontrei o arquivo do modelo: {TEMPLATE}")
    st.info("Confirme se o arquivo está no repositório com esse nome EXATO.")
    st.stop()

with st.form("rn_form"):
    col1, col2 = st.columns(2)

    with col1:
        rn = st.text_input("PROC. Nº (RN)", value="")
        destinatario = st.text_input("DESTINATÁRIO/To", value="")
        subscritor = st.text_input("REMETENTE/FROM - Subscritor", value="")
        filial = st.text_input("REMETENTE/FROM - Comercial/Filial", value="")
        segurado = st.text_input("SEGURADO", value="")
        cnpj_raw = st.text_input("CNPJ (digite números ou com pontuação)", value="")

    with col2:
        email_user = st.text_input("E-mail (antes do @allianz.com.br)", value="")
        data_doc = st.date_input("DATA/DATE", value=date.today())
        paginas = st.number_input("PÁGINAS/PAGES", min_value=1, max_value=200, value=13, step=1)
        cotacao = st.text_input("COTAÇÃO", value="Riscos Nomeados")

    submit = st.form_submit_button("Gerar Word")

if submit:
    cnpj_fmt = format_cnpj(cnpj_raw)
    data_fmt = data_doc.strftime("%d/%m/%Y")

    st.caption(f"📌 Data formatada: **{data_fmt}** | CNPJ formatado: **{cnpj_fmt}**")

    data = {
        "rn": rn,
        "destinatario": destinatario,
        "subscritor": subscritor,
        "filial": filial,
        "email_user": email_user,
        "data": data_fmt,
        "paginas": str(paginas),
        "cotacao": cotacao,
        "segurado": segurado,
        "cnpj": cnpj_fmt,
    }

    try:
        with tempfile.TemporaryDirectory() as tmp:
            out_docx = os.path.join(tmp, "RN_preenchido.docx")
            fill_rn_docx(TEMPLATE, out_docx, data)

            with open(out_docx, "rb") as f:
                st.download_button(
                    "⬇️ Baixar RN preenchido",
                    data=f,
                    file_name="RN_preenchido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
    except Exception as e:
        st.error("Erro ao preencher o Word (rn_fill.py)")
        st.exception(e)
