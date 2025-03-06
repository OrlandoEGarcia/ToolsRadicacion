import fitz  # PyMuPDF
import re

# Ruta del archivo PDF de entrada
pdf_path = "M:/Descargas Autorizaciones/2025-02-13/2025000434298.pdf"
# Ruta del archivo de texto de salida
txt_path = "documento_id.txt"
# Abrir el PDF
doc = fitz.open(pdf_path)
# Extraer el texto completo
text = ""
for page in doc:
    text += page.get_text("text") + "\n"

# Buscar el número de documento de identificación
match = re.search(r"Permiso especial de permanencia\s*\n(\d+)\nNúmero documento de identificación", text)

if match:
    documento_id = match.group(1)
    # Guardar en un archivo de texto
    with open(txt_path, "w", encoding="utf-8") as txt_file:
        txt_file.write(documento_id)
