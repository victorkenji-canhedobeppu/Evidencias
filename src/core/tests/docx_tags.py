from docx import Document

template_path = r"C:\Users\Geologia\PythonProjects\RelatorioEvidencias\src\templates\RSP-116RJ-000+000-GER-EXE-RT-Z9-001-R06.docx"
document = Document(template_path)


def _set_content_control_text(document, tag: str, value: str):
    """Encontra um Content Control pelo seu 'tag' e define o seu texto."""
    for sdt in document.inline_shapes:
        if sdt.type == 3:  # 3 é o tipo para SDT (Structured Document Tag)
            if sdt._inline.docPr.title == tag:
                print(sdt._inline.sdt.sdtContent.r.t.text)
                return
    for sdt in document.paragraphs:
        for run in sdt.runs:
            if run._r.tag.val == tag:
                print(run.text)


_set_content_control_text(document, "Código Interno", "")
