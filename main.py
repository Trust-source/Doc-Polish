from docx import Document

def extract_paragraphs(docx_path):
    doc = Document(docx_path)
    paragraphs = []

    for para in doc.paragraphs:
        paragraphs.append({
            "text": para.text,
            "style": para.style.name
        })

    return paragraphs


print(extract_paragraphs("sample.docx"))