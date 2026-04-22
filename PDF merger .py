from PyPDF3 import PdfFileReader, PdfFileWriter
from docx import Document

def crop_pdf(input_path, output_path, start_page, end_page):
    pdf_reader = PdfFileReader(input_path)
    pdf_writer = PdfFileWriter()

    total_pages = pdf_reader.getNumPages()

    if start_page < 1 or end_page > total_pages:
        print("Invalid page range!")
        return

    for page_num in range(start_page - 1, end_page):
        page = pdf_reader.getPage(page_num)
        pdf_writer.addPage(page)

    with open(output_path, 'wb') as out:
        pdf_writer.write(out)

    print(f"Cropped PDF saved as: {output_path}")


def pdf_to_word(pdf_path, docx_path):
    pdf_reader = PdfFileReader(pdf_path)
    document = Document()

    for page_num in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page_num)
        text = page.extractText()

        document.add_paragraph(text)

    document.save(docx_path)
    print(f"Converted to Word: {docx_path}")


print("PDF Crop & Convert Tool")

input_pdf = input("Enter PDF file path: ")
output_pdf = input("Enter output PDF name: ")

start = int(input("Enter start page: "))
end = int(input("Enter end page: "))

crop_pdf(input_pdf, output_pdf, start, end)

choice = input("Convert to Word? (y/n): ")

if choice.lower() == 'y':
    output_docx = output_pdf.replace(".pdf", ".docx")
    pdf_to_word(output_pdf, output_docx)