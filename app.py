import streamlit as st
import re
from datetime import date
import tempfile
import os
import json
import urllib.request
import copy

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# =============================
# CONFIG STREAMLIT
# =============================
st.set_page_config(page_title="Gerador RN - Allianz", layout="wide")
st.title("Gerador de RN - Modelo Word")
st.success("✅ App carregado com sucesso")

TEMPLATE = "MODELO RN (1).docx"

# =============================
# SESSION STATE (tamanho mestre)
# =============================
if "n_locais" not in st.session_state:
    st.session_state.n_locais = 10  # padrão inicial

# Página 2 (Locais)
if "locais_data" not in st.session_state:
    st.session_state.locais_data = [
        {"cep": "", "endereco_base": "", "numero": "", "complemento": "", "atividade": ""}
        for _ in range(st.session_state.n_locais)
    ]

# Página 3 (VR)
if "vr_data" not in st.session_state:
    st.session_state.vr_data = [
        {"predio": "", "mmu": "", "mmp": "", "lucros": ""}
        for _ in range(st.session_state.n_locais)
    ]

# Versionador para CEP (evitar erro de session_state ao atualizar campo)
if "locais_version" not in st.session_state:
    st.session_state.locais_version = 0


def safe_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def _sync_lists():
    """Garante que locais_data e vr_data tenham tamanho = n_locais."""
    n = int(st.session_state.n_locais)

    L = st.session_state.locais_data
    if len(L) < n:
        L.extend([{"cep": "", "endereco_base": "", "numero": "", "complemento": "", "atividade": ""} for _ in range(n - len(L))])
    elif len(L) > n:
        st.session_state.locais_data = L[:n]

    V = st.session_state.vr_data
    if len(V) < n:
        V.extend([{"predio": "", "mmu": "", "mmp": "", "lucros": ""} for _ in range(n - len(V))])
    elif len(V) > n:
        st.session_state.vr_data = V[:n]


def aumentar_locais(mais=10):
    st.session_state.n_locais = int(st.session_state.n_locais) + int(mais)
    _sync_lists()


def reduzir_locais(menos=10):
    st.session_state.n_locais = max(10, int(st.session_state.n_locais) - int(menos))
    _sync_lists()


# =============================
# FORMATAÇÃO / PARSE
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


def parse_brl_number(val: str) -> float:
    """
    Aceita:
      2000000
      R$ 2.000.000,00
      2.000.000,00
      2000000,50
      1,234.56
    """
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s:
        return 0.0

    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    if not s:
        return 0.0

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")   # BR
        else:
            s = s.replace(",", "")                      # EN
    else:
        if "," in s:
            s = s.replace(".", "").replace(",", ".")    # BR

    try:
        return float(s)
    except Exception:
        return 0.0


def fmt_brl_number(x: float) -> str:
    """1.234.567,89 (sem prefixo)"""
    s = f"{x:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def fmt_brl_money(x: float) -> str:
    """R$ 1.234.567,89"""
    return f"R$ {fmt_brl_number(x)}"


def format_vr_input_keep_prefix(key: str):
    """
    Mantém o prefixo 'R$' automaticamente no campo.
    Ao sair do input:
      2000000 -> R$ 2.000.000,00
    """
    raw = st.session_state.get(key, "")
    value = parse_brl_number(raw)
    if value > 0:
        st.session_state[key] = fmt_brl_money(value)
    else:
        st.session_state[key] = ""


# =============================
# WORD HELPERS
# =============================
def set_cell_text(cell, text, paragraph_index=0):
    while len(cell.paragraphs) <= paragraph_index:
        cell.add_paragraph("")
    p = cell.paragraphs[paragraph_index]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def clear_cell_keep_format(cell):
    tc = cell._tc
    for p in list(tc.p_lst):
        tc.remove(p)


def replace_in_cell_all(cell, old, new):
    for p in cell.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)


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
    for t in doc.tables:
        if len(t.rows) == 0:
            continue
        header = " ".join(c.text.strip().upper() for c in t.rows[0].cells)
        if ("LOCAL" in header) and ("ENDEREÇO" in header) and ("ATIVIDADE" in header) and len(t.columns) >= 3:
            return t
    return None


def find_vr_table(doc):
    # header row 1 contém PRÉDIO/MMU/MMP/LUCROS/LOCAL no modelo
    for t in doc.tables:
        if len(t.rows) < 2:
            continue
        header = " ".join(c.text.strip().upper() for c in t.rows[1].cells)
        if ("PRÉDIO" in header) and ("MMU" in header) and ("MMP" in header) and ("LUCROS" in header) and ("LOCAL" in header):
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


def vr_adjust_rows(table, desired_rows):
    totals_idx = None
    for i, row in enumerate(table.rows):
        if any("TOTAIS" in c.text.upper() for c in row.cells):
            totals_idx = i
            break
    if totals_idx is None:
        return

    data_start = 2
    current_data = totals_idx - data_start

    if desired_rows > current_data:
        template_tr = table.rows[totals_idx - 1]._tr
        for _ in range(desired_rows - current_data):
            new_tr = copy.deepcopy(template_tr)
            table._tbl.insert(totals_idx, new_tr)
            totals_idx += 1

    elif desired_rows < current_data:
        for _ in range(current_data - desired_rows):
            remove_idx = totals_idx - 1
            table._tbl.remove(table.rows[remove_idx]._tr)
            totals_idx -= 1


# =============================
# UI
# =============================
if not os.path.exists(TEMPLATE):
    st.error(f"Arquivo {TEMPLATE} não encontrado no repositório.")
    st.stop()

tabs = st.tabs([
    "Página 1 - Capa/Cotação",
    "Página 2 - Segurado/Vigência/Locais",
    "Página 3 - Valor em Risco (R$)"
])

_sync_lists()

with st.form("rn_form"):

    # -----------------------------
    # Página 1
    # -----------------------------
    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            rn = st.text_input("PROC. Nº (RN)")
            destinatario = st.text_input("DESTINATÁRIO/To")
            subscritor = st.text_input("REMETENTE - Subscritor")
            filial = st.text_input("REMETENTE - Comercial / Filial")
            segurado_p1 = st.text_input("SEGURADO (Página 1)")
            cnpj_raw_p1 = st.text_input("CNPJ (Página 1)")
        with col2:
            email_user = st.text_input("E-mail (antes do @allianz.com.br)")
            data_doc = st.date_input("DATA/DATE", value=date.today())
            paginas = st.number_input("PÁGINAS/PAGES", value=13, min_value=1)
            cotacao = st.text_input("COTAÇÃO", value="Riscos Nomeados")

    # -----------------------------
    # Página 2
    # -----------------------------
    with tabs[1]:
        st.subheader("IV - Vigência do seguro")
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
            st.caption(f"Total de locais (Locais/VR): {st.session_state.n_locais}")

        _sync_lists()
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

            endereco_final_preview = montar_endereco_final(
                st.session_state.locais_data[i]["endereco_base"],
                st.session_state.locais_data[i]["numero"],
                st.session_state.locais_data[i]["complemento"],
            )
            c_end.caption(endereco_final_preview if endereco_final_preview else "")

    # -----------------------------
    # Página 3 - VR (com prefixo R$ no campo)
    # -----------------------------
    with tabs[2]:
        st.subheader("Valor em Risco (R$)")
        st.caption("Total DM = Prédio + MMU + MMP | VR Total = DM + Lucros")

        bb1, bb2, bb3 = st.columns([1, 1, 2])
        with bb1:
            st.button("➕ +10 linhas VR", on_click=aumentar_locais, kwargs={"mais": 10})
        with bb2:
            st.button("➖ -10 linhas VR", on_click=reduzir_locais, kwargs={"menos": 10})
        with bb3:
            st.caption(f"Total de linhas (Locais/VR): {st.session_state.n_locais}")

        _sync_lists()
        n = int(st.session_state.n_locais)

        c0, c1, c2, c3, c4, c5 = st.columns([0.6, 1.3, 1.3, 1.3, 1.4, 1.5])
        c0.markdown("**Local**")
        c1.markdown("**Prédio**")
        c2.markdown("**MMU**")
        c3.markdown("**MMP**")
        c4.markdown("**Total DM**")
        c5.markdown("**Lucros**")

        total_pred = total_mmu = total_mmp = total_dm = total_luc = 0.0

        for i in range(n):
            row = st.session_state.vr_data[i]
            r0, r1, r2, r3, r4, r5 = st.columns([0.6, 1.3, 1.3, 1.3, 1.4, 1.5])
            r0.write(f"{i+1:02d}")

            pred_key = f"vr_pred_{i}"
            mmu_key  = f"vr_mmu_{i}"
            mmp_key  = f"vr_mmp_{i}"
            luc_key  = f"vr_luc_{i}"

            pred_s = r1.text_input("", value=row.get("predio", ""), key=pred_key,
                                   placeholder="R$ 0,00",
                                   on_change=format_vr_input_keep_prefix, args=(pred_key,))
            mmu_s  = r2.text_input("", value=row.get("mmu", ""), key=mmu_key,
                                   placeholder="R$ 0,00",
                                   on_change=format_vr_input_keep_prefix, args=(mmu_key,))
            mmp_s  = r3.text_input("", value=row.get("mmp", ""), key=mmp_key,
                                   placeholder="R$ 0,00",
                                   on_change=format_vr_input_keep_prefix, args=(mmp_key,))
            luc_s  = r5.text_input("", value=row.get("lucros", ""), key=luc_key,
                                   placeholder="R$ 0,00",
                                   on_change=format_vr_input_keep_prefix, args=(luc_key,))

            pred = parse_brl_number(pred_s)
            mmu  = parse_brl_number(mmu_s)
            mmp  = parse_brl_number(mmp_s)
            luc  = parse_brl_number(luc_s)
            dm = pred + mmu + mmp

            st.session_state.vr_data[i]["predio"] = pred_s
            st.session_state.vr_data[i]["mmu"] = mmu_s
            st.session_state.vr_data[i]["mmp"] = mmp_s
            st.session_state.vr_data[i]["lucros"] = luc_s

            r4.write(fmt_brl_money(dm))

            total_pred += pred
            total_mmu += mmu
            total_mmp += mmp
            total_dm += dm
            total_luc += luc

        st.markdown("---")
        t0, t1, t2, t3, t4, t5 = st.columns([0.6, 1.3, 1.3, 1.3, 1.4, 1.5])
        t0.markdown("**Totais**")
        t1.markdown(f"**{fmt_brl_money(total_pred)}**")
        t2.markdown(f"**{fmt_brl_money(total_mmu)}**")
        t3.markdown(f"**{fmt_brl_money(total_mmp)}**")
        t4.markdown(f"**{fmt_brl_money(total_dm)}**")
        t5.markdown(f"**{fmt_brl_money(total_luc)}**")

        vr_total = total_dm + total_luc
        st.markdown(f"### Valor em Risco Total (DM + Lucros) = **{fmt_brl_money(vr_total)}**")

    submit = st.form_submit_button("Gerar Word")


# -----------------------------
# GERAR WORD (página 1/2/3)
# -----------------------------
if submit:
    doc = Document(TEMPLATE)
    n = int(st.session_state.n_locais)

    # Página 1 - capa
    cover = find_table(doc, "PROC. Nº")
    if cover:
        i = find_row(cover, "PROC. Nº")
        if i is not None and rn:
            set_cell_text(cover.cell(i, 1), f"RN - {rn}")

        i = find_row(cover, "DESTINATÁRIO")
        if i is not None and destinatario:
            set_cell_text(cover.cell(i, 1), destinatario)

        i = find_row(cover, "REMETENTE/FROM")
        if i is not None:
            if subscritor:
                set_cell_text(cover.cell(i, 1), subscritor, 0)
            if filial:
                set_cell_text(cover.cell(i, 1), filial, 1)

        i = find_row(cover, "DEPTO/DIVISION")
        if i is not None and email_user:
            replace_in_cell_all(cover.cell(i, 1), "xxxx.xxxx", email_user)

        i = find_row(cover, "DATA/DATE")
        if i is not None:
            set_cell_text(cover.cell(i, 1), data_doc.strftime("%d/%m/%Y"))

        i = find_row(cover, "PÁGINAS/PAGES")
        if i is not None:
            set_cell_text(cover.cell(i, 1), f"{paginas} (incluindo esta capa/including the cover page)")

    # Página 1 - cotação
    quote = find_table(doc, "COTAÇÃO:")
    if quote:
        i = find_row(quote, "COTAÇÃO")
        if i is not None:
            set_cell_text(quote.cell(i, 1), cotacao)

        i = find_row(quote, "SEGURADO")
        if i is not None and segurado_p1:
            set_cell_text(quote.cell(i, 1), segurado_p1)

        i = find_row(quote, "CNPJ")
        if i is not None and cnpj_raw_p1:
            set_cell_text(quote.cell(i, 1), format_cnpj(cnpj_raw_p1))

    # Página 2 - Vigência (corrigida)
    t_vig = find_table(doc, "IV – Vigência do seguro")
    if t_vig and len(t_vig.rows) >= 2 and len(t_vig.columns) >= 2:
        cell = t_vig.cell(1, 1)
        clear_cell_keep_format(cell)
        p1 = cell.add_paragraph(f"Das 24 horas do dia {vig_inicio.strftime('%d/%m/%Y')}")
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2 = cell.add_paragraph(f"Às 24 horas do dia {vig_fim.strftime('%d/%m/%Y')}")
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Página 2 - Locais
    t_locais = find_locais_table(doc)
    if t_locais:
        ensure_table_rows_with_style(t_locais, desired_data_rows=n, header_rows=1)
        for i in range(n):
            row_index = 1 + i
            local_num = f"{i+1:02d}"
            end_base = (st.session_state.locais_data[i].get("endereco_base") or "").strip()
            numv = (st.session_state.locais_data[i].get("numero") or "").strip()
            comp = (st.session_state.locais_data[i].get("complemento") or "").strip()
            endereco_final = montar_endereco_final(end_base, numv, comp)
            atv = (st.session_state.locais_data[i].get("atividade") or "").strip()

            set_cell_text(t_locais.cell(row_index, 0), local_num)
            set_cell_text(t_locais.cell(row_index, 1), endereco_final)
            set_cell_text(t_locais.cell(row_index, 2), atv)

    # Página 3 - VR
    t_vr = find_vr_table(doc)
    if t_vr:
        vr_adjust_rows(t_vr, n)

        totals_idx = None
        for i, row in enumerate(t_vr.rows):
            if any("TOTAIS" in c.text.upper() for c in row.cells):
                totals_idx = i
                break

        data_start = 2
        total_pred = total_mmu = total_mmp = total_dm = total_luc = 0.0

        for i in range(n):
            row_idx = data_start + i
            local_num = f"{i+1:02d}"

            pred = parse_brl_number(st.session_state.vr_data[i].get("predio", ""))
            mmu  = parse_brl_number(st.session_state.vr_data[i].get("mmu", ""))
            mmp  = parse_brl_number(st.session_state.vr_data[i].get("mmp", ""))
            luc  = parse_brl_number(st.session_state.vr_data[i].get("lucros", ""))
            dm = pred + mmu + mmp

            total_pred += pred
            total_mmu += mmu
            total_mmp += mmp
            total_dm += dm
            total_luc += luc

            set_cell_text(t_vr.cell(row_idx, 0), local_num)
            set_cell_text(t_vr.cell(row_idx, 1), fmt_brl_money(pred))
            set_cell_text(t_vr.cell(row_idx, 2), fmt_brl_money(mmu))
            set_cell_text(t_vr.cell(row_idx, 3), fmt_brl_money(mmp))
            set_cell_text(t_vr.cell(row_idx, 4), fmt_brl_money(dm))
            set_cell_text(t_vr.cell(row_idx, 5), fmt_brl_money(luc))

        if totals_idx is not None:
            set_cell_text(t_vr.cell(totals_idx, 1), fmt_brl_money(total_pred))
            set_cell_text(t_vr.cell(totals_idx, 2), fmt_brl_money(total_mmu))
            set_cell_text(t_vr.cell(totals_idx, 3), fmt_brl_money(total_mmp))
            set_cell_text(t_vr.cell(totals_idx, 4), fmt_brl_money(total_dm))
            set_cell_text(t_vr.cell(totals_idx, 5), fmt_brl_money(total_luc))

            vr_total = total_dm + total_luc
            vr_total_row = totals_idx + 1
            try:
                set_cell_text(t_vr.cell(vr_total_row, 5), fmt_brl_money(vr_total))
            except Exception:
                try:
                    set_cell_text(t_vr.cell(vr_total_row, 4), fmt_brl_money(vr_total))
                except Exception:
                    pass

    # Salvar e baixar
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
``
