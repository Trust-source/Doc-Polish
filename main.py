from docx import Document

def extract_paragraphs(docx_path):
    doc = Document(docx_path)
    paragraphs = []

    for para in doc.paragraphs:

        if para.runs:
            first_run = para.runs[0]
            bold = first_run.bold
            if first_run.font.size:
                font_size = first_run.font.size.pt
        paragraphs.append({
            "text": para.text,
            "style": para.style.name,
            "bold": bold,
            "font_size": font_size
        })

    return paragraphs


print(extract_paragraphs("sample.docx"))