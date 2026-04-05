# rn_fill.py
# Preenche SOMENTE as células da direita, preservando os rótulos e a formatação do Word.

def fill_rn_docx(template_path, output_path, data):
    from docx import Document

    doc = Document(template_path)

    # ---------- Helpers (preservar formatação/runs) ----------
    def set_paragraph_text_preserve(paragraph, text):
        """
        Define o texto no 1º run do parágrafo e zera os demais runs,
        preservando estilo do parágrafo/célula.
        """
        if paragraph.runs:
            paragraph.runs[0].text = text
            for r in paragraph.runs[1:]:
                r.text = ""
        else:
            paragraph.add_run(text)

    def set_cell_line(cell, line_text, paragraph_index=0):
        """
        Escreve em um parágrafo específico da célula (sem mexer na coluna esquerda).
        """
        # garante parágrafos suficientes
        while len(cell.paragraphs) <= paragraph_index:
            cell.add_paragraph("")
        set_paragraph_text_preserve(cell.paragraphs[paragraph_index], line_text)

    def replace_in_cell_runs(cell, old, new):
        """
        Substitui texto dentro dos runs existentes, preservando formatação (ideal p/ xxxx.xxxx).
        """
        for p in cell.paragraphs:
            for r in p.runs:
                if old in r.text:
                    r.text = r.text.replace(old, new)

    def find_table_by_anchor(anchor_text):
        """
        Encontra uma tabela que contenha o anchor_text em qualquer célula.
        """
        a = anchor_text.upper()
        for t in doc.tables:
            for row in t.rows:
                for c in row.cells:
                    if a in c.text.upper():
                        return t
        return None

    # ---------- 1) CAPA (tabela com PROC. Nº / DESTINATÁRIO / REMETENTE / DATA / PÁGINAS) ----------
    # No seu modelo, essa tabela existe e tem os rótulos na coluna esquerda. [2](https://allianzms-my.sharepoint.com/personal/ana_araujo1_allianz_com_br/_layouts/15/Doc.aspx?sourcedoc=%7B2954F353-592E-49FB-945A-F5D4715A0C2B%7D&file=MODELO%20RN.docx&action=default&mobileredirect=true&DefaultItemOpen=1)[3](https://allianzms-my.sharepoint.com/personal/manuel_jobcenterext_allpronet_com_br/_layouts/15/Doc.aspx?sourcedoc=%7BECB0E7C2-B90E-4D7A-9EB9-DD473A22F216%7D&file=MODELO%20RN%20(1).docx&action=default&mobileredirect=true&DefaultItemOpen=1)
    cover = find_table_by_anchor("PROC. Nº")
    if cover and len(cover.rows) >= 6 and len(cover.columns) >= 2:
        rn = (data.get("rn") or "").strip()
        destinatario = (data.get("destinatario") or "").strip()
        subscritor = (data.get("subscritor") or "").strip()
        filial = (data.get("filial") or "").strip()
        email_user = (data.get("email_user") or "").strip()
        data_doc = (data.get("data") or "").strip()
        pages = (data.get("paginas") or "").strip()

        # RN (linha 0, coluna direita)
        if rn:
            set_cell_line(cover.cell(0, 1), f"RN - {rn}", 0)

        # Destinatário (linha 1, coluna direita)
        if destinatario:
            set_cell_line(cover.cell(1, 1), destinatario, 0)

        # Remetente/FROM (linha 2, coluna direita) -> DUAS LINHAS
        # Mantém o rótulo REMETENTE/FROM na coluna esquerda intacto.
        if subscritor:
            set_cell_line(cover.cell(2, 1), subscritor, 0)
        if filial:
            set_cell_line(cover.cell(2, 1), filial, 1)

        # E-mail: troca só o placeholder "xxxx.xxxx" dentro dos runs, sem destruir layout
        if email_user:
            replace_in_cell_runs(cover.cell(3, 1), "xxxx.xxxx", email_user)

        # Data (linha 4, coluna direita)
        if data_doc:
            set_cell_line(cover.cell(4, 1), data_doc, 0)

        # Páginas (linha 5, coluna direita) -> mantém o sufixo padrão
        if pages:
            set_cell_line(
                cover.cell(5, 1),
                f"{pages} (incluindo esta capa/including the cover page)",
                0
            )

    # ---------- 2) TABELA DE COTAÇÃO (COTAÇÃO / SEGURADO / CNPJ) ----------
    quote = find_table_by_anchor("COTAÇÃO:")
    if quote and len(quote.rows) >= 3 and len(quote.columns) >= 2:
        cotacao = (data.get("cotacao") or "Riscos Nomeados").strip()
        segurado = (data.get("segurado") or "").strip()
        cnpj = (data.get("cnpj") or "").strip()

        # Preenche SOMENTE a coluna direita (col 1). A esquerda fica intacta.
        set_cell_line(quote.cell(0, 1), cotacao, 0)   # COTAÇÃO
        if segurado:
            set_cell_line(quote.cell(1, 1), segurado, 0)  # SEGURADO
        if cnpj:
            set_cell_line(quote.cell(2, 1), cnpj, 0)      # CNPJ

    doc.save(output_path)
