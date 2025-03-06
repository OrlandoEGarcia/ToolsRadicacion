import pdfplumber
import fitz  # PyMuPDF
import os
import tkinter as tk
import re
import pandas as pd
from openpyxl import Workbook
from tkinter import filedialog, messagebox, ttk


# Configuración de conexión
server = '192.168.59.230'  # Ejemplo: 'localhost' o '192.168.1.100'
database = 'sifacturacion'      # Nombre de la base de datos
username = 'AdminFacturacion'        # Usuario de SQL Server
password = 'SI.Admin.23$%*'     # Contraseña del usuario

# Crear un archivo de Excel
#wb = Workbook()


def extraer_datos():
    
    seleccionar_carpeta()

    if not carpeta_pdf:
        messagebox.showwarning("Error", "Debe seleccionar una carpeta primero.")
        return

    archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]


    # Crear un archivo de Excel
    wb = Workbook()
    ws = wb.active
    # Escribir los encabezados
    ws.append(["AUTORIZACION","DOCUMENTO","CODIGO"])

    for archivo in archivos_pdf:
        pdf_path = os.path.join(carpeta_pdf, archivo)
        try:
            pdf = fitz.open(pdf_path)
        except Exception as e:
            print(f"Error al abrir {archivo}: {e}")
            continue

        # Extraer el texto del PDF
        text = ""

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
            # Guardar en un archivo de texto+
            print(pdf_path,documento_id) 


                # Buscamos los datos correspondientes a los codigos autorizados
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text()

        # Extraer el número de autorización
        autorizacion = re.search(r"NÚMERO DE AUTORIZACIÓN:\s*(\d+)", text).group(1)

        # Buscar la sección de "SERVICIOS AUTORIZADOS"
        servicios_autorizados = re.search(r"SERVICIOS AUTORIZADOS(.*?)(?=Notas auditor|$)", text, re.DOTALL).group(1)

        # Extraer todos los códigos de 6 dígitos en esa sección
        codigos = re.findall(r"\b\d{6}\b", servicios_autorizados)

        for codigo in codigos:
            ws.append([autorizacion, documento_id, codigo])

    # Guardar el archivo de Excel
    wb.save(carpeta_pdf+"/todos_los_codigos.xlsx")
    leerArchivoExcel('ReporteTotal.xlsx','coincidentes.xlsx','nocoincidentes.xlsx')

    print("Archivo Excel generado con éxito.")



def leerArchivoExcel(nombreArchivo,nombreDestino1,nombreDestino2):
    archivo_xls = os.path.join(carpeta_pdf, 'todos_los_codigos.xlsx')
    archivo_rpt = os.path.join(carpeta_pdf, nombreArchivo)

    excel1=archivo_rpt
    excel2=archivo_xls

    df_datos=pd.read_excel(excel1)
    df_reporte=pd.read_excel(excel2)

    col_datos1="DOCUMENTO"
    col_datos2="CODIGO"

    col_rep1="documento"
    col_rep2="cups"

    df_datos['Clave']=df_datos[col_rep1].astype(str)+'_'+df_datos[col_rep2].astype(str)
    df_reporte['Clave']=df_reporte[col_datos1].astype(str)+'_'+df_reporte[col_datos2].astype(str)

    df_fusionado = pd.merge(df_datos, df_reporte, on='Clave', how='left', indicator=True)

    coincidentes = df_fusionado[df_fusionado['_merge'] == 'both']
    no_coincidentes = df_fusionado[df_fusionado['_merge'] == 'left_only']

    coincidentes = coincidentes.drop(columns=['_merge'])
    no_coincidentes = no_coincidentes.drop(columns=['_merge'])

    archivo_destinocoincidentes=os.path.join(carpeta_pdf, nombreDestino1)
    archivo_destinonocoincidentes=os.path.join(carpeta_pdf, nombreDestino2)
    coincidentes.to_excel(archivo_destinocoincidentes, index=False)
    no_coincidentes.to_excel(archivo_destinonocoincidentes, index=False)


# Función para seleccionar la carpeta
def seleccionar_carpeta():
    global carpeta_pdf
    carpeta_pdf = filedialog.askdirectory()
    if carpeta_pdf:
        label_estado.config(text=f"Carpeta seleccionada: {carpeta_pdf}")

# Configuración de la interfaz gráfica con tkinter
root = tk.Tk()
root.title("Extractor de Datos de PDF")
root.geometry("500x300")
root.config(bg="#f0f0f0")

# Título
label_titulo = tk.Label(root, text="Extractor de Datos de PDF", font=("Arial", 16, "bold"), bg="#f0f0f0")
label_titulo.pack(pady=10)

# Instrucción
label = tk.Label(root, text="Seleccione la carpeta que contiene los archivos PDF:", font=("Arial", 12), bg="#f0f0f0")
label.pack(pady=5)

# Botón para ejecutar extracción
boton_extraer = tk.Button(root, text="Ejecutar Extraccion Datos", command=extraer_datos, font=("Arial", 12), bg="#2196F3", fg="white")
boton_extraer.pack(pady=10)

# Etiqueta de estado
label_estado = tk.Label(root, text="Ninguna carpeta seleccionada", font=("Arial", 10), bg="#f0f0f0", fg="grey")
label_estado.pack(pady=5)

# Barra de progreso
barra_progreso = ttk.Progressbar(root, mode="indeterminate")
barra_progreso.pack(pady=10, fill=tk.X, padx=20)

root.mainloop()
