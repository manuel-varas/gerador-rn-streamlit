import streamlit as st
import re
from datetime import date
import tempfile
import os
import json
import urllib.request
import urllib.error
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
    # lista de dicts: cada item {cep, endereco, atividade}
    st.session_state.locais_data = [{"cep": "", "endereco": "", "atividade": ""} for _ in range(st.session_state.n_locais)]

def _sync_locais_list():
    """Garante que locais_data tenha tamanho = n_locais."""
    n = st.session_state.n_locais
    cur = st.session_state.locais_data
    if len(cur) < n:
        cur.extend([{"cep": "", "endereco": "", "atividade": ""} for _ in range(n - len(cur))])
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
    """
    Busca CEP no ViaCEP.
    Se não der, retorna "".
    """
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
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return ""

def set_cell_text(cell, text, paragraph_index=0):
    """
    Escreve preservando estilo (não destrói a célula).
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

def replace_in_cell_all(cell, old, new, max_replacements=None):
    """
    Substitui 'old' por 'new' nos runs da célula.
    Se max_replacements for 1, troca só a primeira ocorrência.
    """
    count = 0
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                if max_replacements is None:
                    r.text = r.text.replace(old, new)
                else:
                    # troca controlada
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

def ensure_table_rows_with_style(table, desired_data_rows, header_rows=1, template_row_index=None):
    """
    Garante que a tabela tenha header_rows + desired_data_rows linhas.
    Para manter estilo, clona o XML do 'template_row_index' (por padrão a última linha atual).
    """
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

# =============================
# UI - FORMULÁRIO COM PÁGINAS
# =============================
if not os.path.exists(TEMPLATE):
    st.error(f"Arquivo {TEMPLATE} não encontrado no repositório")
    st.stop()

tabs = st.tabs(["Página 1 - Capa/Cotação", "Página 2 - Segurado/Vigência/Locais"])

with st.form("rn_form"):
    # ---------------------------
    # PÁGINA 1 (mantida)
    # ---------------------------
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

    # ---------------------------
    # PÁGINA 2 (NOVO)
    # ---------------------------
    with tabs[1]:
        st.subheader("I - Segurado / Cossegurados (campos verdes)")
        c1, c2 = st.columns(2)
        with c1:
            segurado_p2 = st.text_input("Segurado (Página 2)", value=segurado_p1)
            cossegurados = st.text_input("Cossegurados (Página 2)", value="")
        with c2:
            cnpj_raw_p2 = st.text_input("CNPJ Segurado (Página 2)", value=cnpj_raw_p1)
            cosseg_cnpj_raw = st.text_input("CNPJ Cossegurados (Página 2)", value="")

        st.subheader("III - Atividade Principal (campo verde)")
        atividade_principal = st.text_input("Atividade Principal", value="")

        st.subheader("IV - Vigência do seguro (substitui xx/xx/xxxx)")
        v1, v2 = st.columns(2)
        with v1:
            vig_inicio = st.date_input("Início de vigência", value=date.today(), key="vig_inicio_p2")
        with v2:
            vig_fim = st.date_input("Término de vigência", value=date.today(), key="vig_fim_p2")

        st.subheader("V - Locais em Risco/VR (Endereço e Atividade)")
        b1, b2, b3 = st.columns([1, 1, 2])
        with b1:
            st.button("➕ +10 locais", on_click=aumentar_locais, kwargs={"mais": 10})
        with b2:
            st.button("➖ -10 locais", on_click=reduzir_locais, kwargs={"menos": 10})
        with b3:
            st.caption(f"Total de locais na interface: {st.session_state.n_locais}")

        _sync_locais_list()

        # Cabeçalho
        h1, h2, h3, h4 = st.columns([0.6, 1.0, 2.5, 2.0])
        h1.markdown("**Local**")
        h2.markdown("**CEP**")
        h3.markdown("**Endereço**")
        h4.markdown("**Atividade**")

        for i in range(st.session_state.n_locais):
            row = st.session_state.locais_data[i]
            c_local, c_cep, c_end, c_atv = st.columns([0.6, 1.0, 2.5, 2.0])

            c_local.write(f"{i+1:02d}")

            cep_key = f"cep_{i}"
            end_key = f"end_{i}"
            atv_key = f"atv_{i}"

            row["cep"] = c_cep.text_input("", value=row["cep"], key=cep_key, placeholder="00000-000")
            row["endereco"] = c_end.text_input("", value=row["endereco"], key=end_key, placeholder="Rua..., nº..., Bairro..., Cidade-UF")
            row["atividade"] = c_atv.text_input("", value=row["atividade"], key=atv_key, placeholder="Atividade do local")

            # Botão pequeno de busca (por linha)
            bcol1, bcol2 = st.columns([1, 6])
            if bcol1.button("Buscar CEP", key=f"buscar_{i}"):
                end = viacep_lookup(row["cep"])
                if end:
                    row["endereco"] = end
                    st.session_state[end_key] = end
                    st.toast(f"CEP {format_cep(row['cep'])} encontrado!", icon="✅")
                else:
                    st.toast("CEP não encontrado ou sem acesso. Preencha manualmente.", icon="⚠️")

        st.session_state.locais_data = st.session_state.locais_data  # persistência

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

    # ========= PÁGINA 1 (capa) =========
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

    # ========= PÁGINA 2 - I (Segurado / Cossegurados) =========
    t_seg = find_table(doc, "I – Segurado")
    if t_seg:
        # Estrutura do modelo: row 1 (vazio) é segurado; row 3 (vazio) é cossegurados [1](https://allianzms.sharepoint.com/sites/DE1000-connect-az-reinsurance-advisory-services/SiteAssets/SitePages/Ge/GIS_Training_MasteringGeospatial_2025.pdf?web=1)
        if len(t_seg.rows) >= 4 and len(t_seg.columns) >= 2:
            set_cell_text(t_seg.cell(1, 0), data["segurado_p2"])
            set_cell_text(t_seg.cell(1, 1), data["cnpj_p2"])
            set_cell_text(t_seg.cell(3, 0), data["cossegurados"])
            set_cell_text(t_seg.cell(3, 1), data["cosseg_cnpj"])

    # ========= PÁGINA 2 - III (Atividade Principal) =========
    t_iii = find_table(doc, "III – Objeto Segurado / Atividade Principal")
    if t_iii:
        # No modelo, a última linha é vazia (após "Atividade Principal:") [1](https://allianzms.sharepoint.com/sites/DE1000-connect-az-reinsurance-advisory-services/SiteAssets/SitePages/Ge/GIS_Training_MasteringGeospatial_2025.pdf?web=1)
        if len(t_iii.rows) >= 5:
            set_cell_text(t_iii.cell(4, 0), data["atividade_principal"])

    # ========= PÁGINA 2 - IV (Vigência do seguro) =========
    t_vig = find_table(doc, "IV – Vigência do seguro")
    if t_vig:
        # Substitui os dois "xx/xx/xxxx" dentro da célula da direita [1](https://allianzms.sharepoint.com/sites/DE1000-connect-az-reinsurance-advisory-services/SiteAssets/SitePages/Ge/GIS_Training_MasteringGeospatial_2025.pdf?web=1)
        if len(t_vig.rows) >= 2 and len(t_vig.columns) >= 2:
            cell = t_vig.cell(1, 1)
            # troca primeira ocorrência = início
            replace_in_cell_all(cell, "xx/xx/xxxx", data["vig_inicio"], max_replacements=1)
            # troca segunda ocorrência = fim
            replace_in_cell_all(cell, "xx/xx/xxxx", data["vig_fim"], max_replacements=1)

    # ========= PÁGINA 2 - V (Locais em Risco/VR) =========
    # A tabela logo após "V – Locais em Risco/VR" tem cabeçalho Local/Endereço/Atividade e 10 linhas no modelo [1](https://allianzms.sharepoint.com/sites/DE1000-connect-az-reinsurance-advisory-services/SiteAssets/SitePages/Ge/GIS_Training_MasteringGeospatial_2025.pdf?web=1)
    t_locais = find_table(doc, "Endereço")
    # Para evitar pegar tabelas erradas, validamos que a primeira linha contém os 3 termos
    if t_locais:
        header = " ".join(c.text.strip() for c in t_locais.rows[0].cells).upper()
        if "LOCAL" in header and "ENDEREÇO" in header and "ATIVIDADE" in header:
            desired = len(data["locais"])
            # garante linhas suficientes mantendo estilo (clonando última linha)
            ensure_table_rows_with_style(t_locais, desired_data_rows=desired, header_rows=1)

            # preencher linhas (row 1..N)
            for i in range(desired):
                row_index = 1 + i
                local_num = f"{i+1:02d}"
                end = (data["locais"][i].get("endereco") or "").strip()
                atv = (data["locais"][i].get("atividade") or "").strip()

                # coluna 0: número do local (mantém padrão)
                set_cell_text(t_locais.cell(row_index, 0), local_num)
                # coluna 1: endereço
                set_cell_text(t_locais.cell(row_index, 1), end)
                # coluna 2: atividade
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
