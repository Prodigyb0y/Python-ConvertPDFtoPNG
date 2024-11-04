
import win32com.client as win32
import os
import fitz

# Caminhos dos arquivos e pastas
pdf_folder = r"pasta_pdf"
png_folder = r"pasta_png"

# Lista de nomes dos arquivos PDF
pdf_files = ["nome_arquivos"]

# Dimensões desejadas (largura, altura)
desired_width = 0
desired_height = 0

# Verificar se a pasta de PNGs existe, se não, criar
os.makedirs(png_folder, exist_ok=True)

# Loop para converter cada PDF
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    doc = fitz.open(pdf_path)  # Abrir o PDF

    # Salvar cada página do PDF como PNG com as dimensões desejadas
    for i, page in enumerate(doc):
        mat = fitz.Matrix(desired_width / page.rect.width, desired_height / page.rect.height)
        pix = page.get_pixmap(matrix=mat)  # Obter o mapa de pixels com as dimensões desejadas
        png_name = f"{os.path.splitext(pdf_file)[0]}_pagina_{i+1}.png"
        png_path = os.path.join(png_folder, png_name)
        pix.save(png_path)  # Salvar a imagem

    print(f"PDF '{pdf_file}' convertido com sucesso!")

 
