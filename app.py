# rn_fill.py
# VERSÃO: RN_FILL_V2_RIGHT_ONLY_2026-04-04

def fill_rn_docx(template_path, output_path, data):
    from docx import Document

    doc = Document(template_path)

    # ---------- helpers ----------
    def set_paragraph_text_preserve(paragraph, text):
        if paragraph.runs:
            paragraph.runs[0].text = text
            for r in paragraph.runs[1:]:
                r.text = ""
        else:
            paragraph.add_run(text)

    def set_cell_text_preserve(cell, text, p_index=0):
        while len(cell.paragraphs) <= p_index:
            cell.add_paragraph("")
        set_paragraph_text_preserve(cell.paragraphs[p_index], text)

    def replace_in_cell_runs(cell, old, new):
        for p in cell.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new)

    def find_table_by_anchor(anchor_text):
        a = anchor_text.upper()
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    if a in c.text.upper():
                        return t
        return None

    def find_row_index_by_left_label(table, label_contains):
        needle = label_contains.upper()
        for idx, row in enumerate(table.rows):
            if len(row.cells) >= 2 and needle in row.cells[0].text.upper():
                return idx
        return None

    # ---------- CAPA ----------
    cover = find_table_by_anchor("PROC. Nº")
    if cover and len(cover.columns) >= 2:
        # PROC / RN
        i = find_row_index_by_left_label(cover, "PROC. Nº")
        if i is not None and data.get("rn"):
            set_cell_text_preserve(cover.cell(i, 1), f"RN - {data['rn']}")

        # DESTINATÁRIO
        i = find_row_index_by_left_label(cover, "DESTINATÁRIO")
        if i is not None and data.get("destinatario"):
            set_cell_text_preserve(cover.cell(i, 1), data["destinatario"])

        # REMETENTE (2 linhas na célula da direita)
        i = find_row_index_by_left_label(cover, "REMETENTE")
        if i is not None:
            if data.get("subscritor"):
                set_cell_text_preserve(cover.cell(i, 1), data["subscritor"], 0)
            if data.get("filial"):
                set_cell_text_preserve(cover.cell(i, 1), data["filial"], 1)

        # DEPTO/DIVISION (só troca o xxxx.xxxx)
        i = find_row_index_by_left_label(cover, "DEPTO/DIVISION")
        if i is not None and data.get("email_user"):
            replace_in_cell_runs(cover.cell(i, 1), "xxxx.xxxx", data["email_user"])

        # DATA/DATE
        i = find_row_index_by_left_label(cover, "DATA/DATE")
        if i is not None and data.get("data"):
            set_cell_text_preserve(cover.cell(i, 1), data["data"])

        # PÁGINAS/PAGES
        i = find_row_index_by_left_label(cover, "PÁGINAS/PAGES")
        if i is not None and data.get("paginas"):
            set_cell_text_preserve(
                cover.cell(i, 1),
                f"{data['paginas']} (incluindo esta capa/including the cover page)"
            )

    # ---------- COTAÇÃO ----------
    quote = find_table_by_anchor("COTAÇÃO:")
    if quote and len(quote.columns) >= 2:
        i = find_row_index_by_left_label(quote, "COTAÇÃO")
        if i is not None:
            set_cell_text_preserve(quote.cell(i, 1), data.get("cotacao", "Riscos Nomeados"))

        i = find_row_index_by_left_label(quote, "SEGURADO")
        if i is not None and data.get("segurado"):
            set_cell_text_preserve(quote.cell(i, 1), data["segurado"])

        i = find_row_index_by_left_label(quote, "CNPJ")
        if i is not None and data.get("cnpj"):
            set_cell_text_preserve(quote.cell(i, 1), data["cnpj"])

    doc.save(output_path)
``
