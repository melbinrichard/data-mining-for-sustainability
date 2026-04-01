import PyPDF2
import os

# Open the PDF file
input_pdf = "../Annual Reports/Shell/shell-annual-report-2023.pdf"

# Extract filename and append "_Extract" before the file extension
base_name = os.path.splitext(input_pdf)[0]
last_four_chars = base_name[-4:]

# Create the new output filename using the last four characters
output_pdf = f"../Annual Report Extract/Shell/{last_four_chars}.pdf"

with open(input_pdf, "rb") as file:
    reader = PyPDF2.PdfReader(file)
    writer = PyPDF2.PdfWriter()

    for page_num in range(30, 32):
        page = reader.pages[page_num]
        writer.add_page(page)

    # Write the extracted pages to a new PDF
    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

print(f"Pages have been extracted and saved to {output_pdf}")
