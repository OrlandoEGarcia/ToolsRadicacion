import fitz  # PyMuPDF

# Ruta del archivo PDF de entrada
pdf_path = "M:/Descargas Autorizaciones/2025-02-13/2025000434298.pdf"
# Ruta del archivo de texto de salida
txt_path = "output.txt"

# Abrir el PDF
doc = fitz.open(pdf_path)

# Extraer el texto
text = ""
for page in doc:
    text += page.get_text("text") + "\n"

# Guardar el texto en un archivo
with open(txt_path, "w", encoding="utf-8") as txt_file:
    txt_file.write(text)

print(f"Texto extraído y guardado en {txt_path}")
