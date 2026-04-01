def fill_rn_docx(template_path, output_path, data):
    from docx import Document  # 👈 IMPORT AQUI DENTRO (IMPORT LAZY)

    doc = Document(template_path)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in data.items():
                    if isinstance(value, str) and value:
                        if key.upper() in cell.text.upper():
                            cell.text = value

    doc.save(output_path)
