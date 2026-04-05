import streamlit as st
import re
from datetime import date
import tempfile
import os
import json
import urllib.request
import copy

from docx import Document

# =============================
# CONFIG STREAMLIT
# =============================
st.set_page_config(page_title="Gerador RN - Allianz", layout="wide")
st.title("Gerador de RN - Modelo Word")
st.success("✅ App carregado com sucesso")

TEMPLATE = "MODELO RN (1).docx"

# =============================
# SESSION STATE (locais dinâmicos)
# =============================
if "n_locais" not in st.session_state:
    st.session_state.n_locais = 10  # padrão inicial

if "locais_data" not in st.session_state:
    # cada item: cep, endereco_base, numero, complemento, atividade
    st.session_state.locais_data = [
        {"cep": "", "endereco_base": "", "numero": "", "complemento": "", "atividade": ""}
        for _ in range(st.session_state.n_locais)
    ]

# versionador para permitir atualizar defaults sem mexer em session_state de widgets já montados
if "locais_version" not in st.session_state:
    st.session_state.locais_version = 0


def _sync_locais_list():
    """Garante que locais_data tenha tamanho = n_locais."""
    n = int(st.session_state.n_locais)
    cur = st.session_state.locais_data
    if len(cur) < n:
        cur.extend(
            [{"cep": "", "endereco_base": "", "numero": "", "complemento": "", "atividade": ""} for _ in range(n - len(cur))]
        )
    elif len(cur) > n:
        st.session_state.locais_data = cur[:n]


def aumentar_locais(mais=10):
    st.session_state.n_locais = int(st.session_state.n_locais) + int(mais)
    _sync_locais_list()


def reduzir_locais(menos=10):
    st.session_state.n_locais = max(10, int(st.session_state.n_locais) - int(menos))
    _sync_locais_list()


# =============================
# FUNÇÕES AUXILIARES
# =============================
def format_cnpj(cnpj: str) -> str:
    nums = re.sub(r"\D", "", cnpj or "")
    if len(nums) == 14:
        return f"{nums[:2]}.{nums[2:5]}.{nums[5:8]}/{nums[8:12]}-{nums[12:]}"
    return cnpj


def format_cep(cep: str) -> str:
    nums = re.sub(r"\D", "", cep or "")
    if len(nums) == 8:
        return f"{nums[:5]}-{nums[5:]}"
    return cep


def viacep_lookup(cep: str) -> str:
    """Busca CEP no ViaCEP. Retorna string base do endereço ou '' se falhar."""
    nums = re.sub(r"\D", "", cep or "")
    if len(nums) != 8:
        return ""

    url = f"https://viacep.com.br/ws/{nums}/json/"
    try:
        with urllib.request.urlopen(url, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("erro"):
            return ""
        logradouro = data.get("logradouro", "")
        bairro = data.get("bairro", "")
        localidade = data.get("localidade", "")
        uf = data.get("uf", "")
        complemento = data.get("complemento", "")
        parts = [p for p in [logradouro, complemento, bairro, f"{localidade}-{uf}"] if p]
        return " - ".join(parts).strip()
    except Exception:
        return ""


def montar_endereco_final(endereco_base: str, numero: str, complemento: str) -> str:
    """Monta no formato: base - Nº xx - complemento"""
    partes = []
    base = (endereco_base or "").strip()
    if base:
        partes.append(base)

    num = (numero or "").strip()
    if num:
        partes.append(f"Nº {num}")

    comp = (complemento or "").strip()
    if comp:
        partes.append(comp)

    return " - ".join(partes)


# =============================
# FUNÇÕES WORD
# =============================
def set_cell_text(cell, text, paragraph_index=0):
    """Escreve preservando estilo (não destrói a célula)."""
    while len(cell.paragraphs) <= paragraph_index:
        cell.add_paragraph("")
    p = cell.paragraphs[paragraph_index]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def replace_in_cell_all(cell, old, new, max_replacements=None):
    count = 0
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                if max_replacements is None:
                    r.text = r.text.replace(old, new)
                else:
                    while old in r.text and count < max_replacements:
                        r.text = r.text.replace(old, new, 1)
                        count += 1
                        if count >= max_replacements:
                            break
    return count


def find_table(doc, anchor_text):
    a = anchor_text.upper()
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                if a in c.text.upper():
                    return t
    return None


def find_row(table, left_label_contains):
    needle = left_label_contains.upper()
    for i, row in enumerate(table.rows):
        if len(row.cells) >= 2 and needle in row.cells[0].text.upper():
            return i
    return None


def find_locais_table(doc):
    """Acha a tabela com cabeçalho Local/Endereço/Atividade."""
    for t in doc.tables:
        if len(t.rows) == 0:
            continue
        header = " ".join(c.text.strip().upper() for c in t.rows[0].cells)
        if ("LOCAL" in header) and ("ENDEREÇO" in header) and ("ATIVIDADE" in header) and len(t.columns) >= 3:
            return t
    return None


def ensure_table_rows_with_style(table, desired_data_rows, header_rows=1, template_row_index=None):
    current_rows = len(table.rows)
    target_rows = header_rows + desired_data_rows
    if current_rows >= target_rows:
        return
    if template_row_index is None:
        template_row_index = current_rows - 1
    template_tr = table.rows[template_row_index]._tr
    for _ in range(target_rows - current_rows):
        new_tr = copy.deepcopy(template_tr)
        table._tbl.append(new_tr)


def safe_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


# =============================
# UI
# =============================
if not os.path.exists(TEMPLATE):
    st.error(f"Arquivo {TEMPLATE} não encontrado no repositório")
    st.stop()

tabs = st.tabs(["Página 1 - Capa/Cotação", "Página 2 - Segurado/Vigência/Locais"])

with st.form("rn_form"):

    # ---------- Página 1 ----------
    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            rn = st.text_input("PROC. Nº (RN)")
            destinatario = st.text_input("DESTINATÁRIO / To")
            subscritor = st.text_input("REMETENTE - Subscritor")
            filial = st.text_input("REMETENTE - Comercial / Filial")
            segurado_p1 = st.text_input("SEGURADO (Página 1)")
            cnpj_raw_p1 = st.text_input("CNPJ (Página 1)")

        with col2:
            email_user = st.text_input("E-mail (antes do @allianz.com.br)")
            data_doc = st.date_input("DATA / DATE", value=date.today())
            paginas = st.number_input("PÁGINAS / PAGES", value=13, min_value=1)
            cotacao = st.text_input("COTAÇÃO", value="Riscos Nomeados")

    # ---------- Página 2 ----------
    with tabs[1]:
        st.subheader("I - Segurado / Cossegurados")
        c1, c2 = st.columns(2)
        with c1:
            segurado_p2 = st.text_input("Segurado (Página 2)", value="")
            cossegurados = st.text_input("Cossegurados (Página 2)", value="")
        with c2:
            cnpj_raw_p2 = st.text_input("CNPJ Segurado (Página 2)", value="")
            cosseg_cnpj_raw = st.text_input("CNPJ Cossegurados (Página 2)", value="")

        st.subheader("III - Atividade Principal")
        atividade_principal = st.text_input("Atividade Principal", value="")

        st.subheader("IV - Vigência do seguro (substitui xx/xx/xxxx)")
        v1, v2 = st.columns(2)
        with v1:
            vig_inicio = st.date_input("Início de vigência", value=date.today(), key="vig_inicio_p2")
        with v2:
            vig_fim = st.date_input("Término de vigência", value=date.today(), key="vig_fim_p2")

        st.subheader("V - Locais em Risco/VR")
        b1, b2, b3 = st.columns([1, 1, 2])
        with b1:
            st.button("➕ +10 locais", on_click=aumentar_locais, kwargs={"mais": 10})
        with b2:
            st.button("➖ -10 locais", on_click=reduzir_locais, kwargs={"menos": 10})
        with b3:
            st.caption(f"Total de locais na interface: {st.session_state.n_locais}")

        _sync_locais_list()
        ver = int(st.session_state.locais_version)

        h1, h2, h3, h4, h5, h6, h7 = st.columns([0.6, 1.0, 0.9, 2.3, 0.8, 1.2, 2.0])
        h1.markdown("**Local**")
        h2.markdown("**CEP**")
        h3.markdown("**Buscar**")
        h4.markdown("**Endereço**")
        h5.markdown("**Nº**")
        h6.markdown("**Complemento**")
        h7.markdown("**Atividade**")

        for i in range(int(st.session_state.n_locais)):
            row = st.session_state.locais_data[i]

            c_local, c_cep, c_btn, c_end, c_num, c_comp, c_atv = st.columns([0.6, 1.0, 0.9, 2.3, 0.8, 1.2, 2.0])
            c_local.write(f"{i+1:02d}")

            cep_key = f"cep_{i}_{ver}"
            end_key = f"end_{i}_{ver}"
            num_key = f"num_{i}_{ver}"
            comp_key = f"comp_{i}_{ver}"
            atv_key = f"atv_{i}_{ver}"

            cep_val = c_cep.text_input("", value=row.get("cep", ""), key=cep_key, placeholder="00000-000")
            end_val = c_end.text_input("", value=row.get("endereco_base", ""), key=end_key, placeholder="Rua..., Bairro..., Cidade-UF")
            num_val = c_num.text_input("", value=row.get("numero", ""), key=num_key, placeholder="XX")
            comp_val = c_comp.text_input("", value=row.get("complemento", ""), key=comp_key, placeholder="Complemento")
            atv_val = c_atv.text_input("", value=row.get("atividade", ""), key=atv_key, placeholder="Atividade")

            # Persistir digitado
            st.session_state.locais_data[i]["cep"] = cep_val
            st.session_state.locais_data[i]["endereco_base"] = end_val
            st.session_state.locais_data[i]["numero"] = num_val
            st.session_state.locais_data[i]["complemento"] = comp_val
            st.session_state.locais_data[i]["atividade"] = atv_val

            if c_btn.button("CEP", key=f"buscar_{i}_{ver}"):
                base = viacep_lookup(cep_val)
                if base:
                    st.session_state.locais_data[i]["cep"] = format_cep(cep_val)
                    st.session_state.locais_data[i]["endereco_base"] = base
                    st.session_state.locais_version += 1
                    st.toast(f"CEP {format_cep(cep_val)} encontrado!", icon="✅")
                    safe_rerun()
                else:
                    st.toast("CEP não encontrado ou sem acesso. Preencha manualmente.", icon="⚠️")

            # Preview do endereço final no padrão pedido
            endereco_final_preview = montar_endereco_final(
                st.session_state.locais_data[i]["endereco_base"],
                st.session_state.locais_data[i]["numero"],
                st.session_state.locais_data[i]["complemento"],
            )
            c_end.caption(endereco_final_preview if endereco_final_preview else "")

    submit = st.form_submit_button("Gerar Word")


# =============================
# PROCESSAMENTO (Word)
# =============================
if submit:
    doc = Document(TEMPLATE)

    data = {
        # pág 1
        "rn": rn,
        "destinatario": destinatario,
        "subscritor": subscritor,
        "filial": filial,
        "email_user": email_user,
        "data": data_doc.strftime("%d/%m/%Y"),
        "paginas": str(paginas),
        "cotacao": cotacao,
        "segurado_p1": segurado_p1,
        "cnpj_p1": format_cnpj(cnpj_raw_p1),

        # pág 2
        "segurado_p2": segurado_p2,
        "cnpj_p2": format_cnpj(cnpj_raw_p2),
        "cossegurados": cossegurados,
        "cosseg_cnpj": format_cnpj(cosseg_cnpj_raw),
        "atividade_principal": atividade_principal,
        "vig_inicio": vig_inicio.strftime("%d/%m/%Y"),
        "vig_fim": vig_fim.strftime("%d/%m/%Y"),
        "locais": st.session_state.locais_data,
    }

    # ========= PÁGINA 1 =========
    cover = find_table(doc, "PROC. Nº")
    if cover:
        i = find_row(cover, "PROC. Nº")
        if i is not None and data["rn"]:
            set_cell_text(cover.cell(i, 1), f"RN - {data['rn']}")

        i = find_row(cover, "DESTINATÁRIO")
        if i is not None and data["destinatario"]:
            set_cell_text(cover.cell(i, 1), data["destinatario"])

        i = find_row(cover, "REMETENTE/FROM")
        if i is not None:
            if data["subscritor"]:
                set_cell_text(cover.cell(i, 1), data["subscritor"], 0)
            if data["filial"]:
                set_cell_text(cover.cell(i, 1), data["filial"], 1)

        i = find_row(cover, "DEPTO/DIVISION")
        if i is not None and data["email_user"]:
            replace_in_cell_all(cover.cell(i, 1), "xxxx.xxxx", data["email_user"])

        i = find_row(cover, "DATA/DATE")
        if i is not None:
            set_cell_text(cover.cell(i, 1), data["data"])

        i = find_row(cover, "PÁGINAS/PAGES")
        if i is not None:
            set_cell_text(
                cover.cell(i, 1),
                f"{data['paginas']} (incluindo esta capa/including the cover page)"
            )

    quote = find_table(doc, "COTAÇÃO:")
    if quote:
        i = find_row(quote, "COTAÇÃO")
        if i is not None:
            set_cell_text(quote.cell(i, 1), data["cotacao"])

        i = find_row(quote, "SEGURADO")
        if i is not None and data["segurado_p1"]:
            set_cell_text(quote.cell(i, 1), data["segurado_p1"])

        i = find_row(quote, "CNPJ")
        if i is not None and data["cnpj_p1"]:
            set_cell_text(quote.cell(i, 1), data["cnpj_p1"])

    # ========= PÁGINA 2 - I =========
    t_seg = find_table(doc, "I – Segurado")
    if t_seg:
        if len(t_seg.rows) >= 4 and len(t_seg.columns) >= 2:
            set_cell_text(t_seg.cell(1, 0), data["segurado_p2"])
            set_cell_text(t_seg.cell(1, 1), data["cnpj_p2"])
            set_cell_text(t_seg.cell(3, 0), data["cossegurados"])
            set_cell_text(t_seg.cell(3, 1), data["cosseg_cnpj"])

    # ========= PÁGINA 2 - III =========
    t_iii = find_table(doc, "III – Objeto Segurado / Atividade Principal")
    if t_iii:
        if len(t_iii.rows) >= 5:
            set_cell_text(t_iii.cell(4, 0), data["atividade_principal"])

    # ========= PÁGINA 2 - IV =========
    t_vig = find_table(doc, "IV – Vigência do seguro")
    if t_vig:
        if len(t_vig.rows) >= 2 and len(t_vig.columns) >= 2:
            cell = t_vig.cell(1, 1)
            replace_in_cell_all(cell, "xx/xx/xxxx", data["vig_inicio"], max_replacements=1)
            replace_in_cell_all(cell, "xx/xx/xxxx", data["vig_fim"], max_replacements=1)

    # ========= PÁGINA 2 - V (Locais) =========
    t_locais = find_locais_table(doc)
    if t_locais:
        desired = len(data["locais"])
        ensure_table_rows_with_style(t_locais, desired_data_rows=desired, header_rows=1)

        for i in range(desired):
            row_index = 1 + i
            local_num = f"{i+1:02d}"

            end_base = (data["locais"][i].get("endereco_base") or "").strip()
            num = (data["locais"][i].get("numero") or "").strip()
            comp = (data["locais"][i].get("complemento") or "").strip()
            endereco_final = montar_endereco_final(end_base, num, comp)

            atv = (data["locais"][i].get("atividade") or "").strip()

            set_cell_text(t_locais.cell(row_index, 0), local_num)
            set_cell_text(t_locais.cell(row_index, 1), endereco_final)
            set_cell_text(t_locais.cell(row_index, 2), atv)

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
