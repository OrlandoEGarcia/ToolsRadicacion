import fitz  # PyMuPDF
import pandas as pd
import os
import re 
import tkinter as tk
import openpyxl
from tabulate import tabulate

import pyodbc

from tkinter import filedialog, messagebox, ttk

from tkcalendar import DateEntry
from tkinter import messagebox

import pyodbc
from openpyxl import Workbook


# Configuración de conexión
server = '192.168.59.230'  # Ejemplo: 'localhost' o '192.168.1.100'
database = 'sifacturacion'      # Nombre de la base de datos
username = 'AdminFacturacion'        # Usuario de SQL Server
password = 'SI.Admin.23$%*'     # Contraseña del usuario



# Función para manejar la fecha seleccionada
def verificarAutorizaciones():
    fecha = date_entry.get()  # Obtiene la fecha como cadena en formato DD-MM-YYYY
    nombre = entrada.get()
    #messagebox.showinfo("Fecha seleccionada", f"Has seleccionado: {fecha}")
    seleccionar_carpeta()
    
    
    # Crear un archivo Excel con openpyxl
    archivo_excel = os.path.join(carpeta_pdf, "AutorizacionesFaltantes.xlsx")
    workbook = Workbook()
    hoja = workbook.active
    hoja.title = "Resultados"
    hoja.append(["IdAdmision","FechaIngreso","NoFactura","FechaFactura","NoRemision","IdUsuario","NoHistoria"])  # Encabezados del archivo    
    
    
    # Crear la conexión a la base de datos
    try:
        conexion = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        print("Conexión exitosa a la base de datos.")

        # Crear un cursor para ejecutar la consulta
        cursor = conexion.cursor()

        # Consulta SQL con dos parámetros en el WHERE
        consulta="set dateformat dmy \r\n select a.IdAdmision,a.FechaIngreso,NoFactura,FechaFactura,NoRemision,A.IdUsuario,a.NoHistoria from mFacturas f INNER JOIN mAdmisiones A ON F.IdAdmision=A.IdAdmision inner join cUsuarios U on f.IdcUsuario=u.IdcUsuario where Usuario= ?  AND FechaFactura BETWEEN  ?  AND  ? "

        #consulta = "SELECT columna1, columna2 FROM TablaEjemplo WHERE columna3 = ? AND columna4 = ?"
        parametro1 = nombre
        parametro2 = fecha
        parametro3 = fecha + " 23:59:59"
        

        # Ejecutar la consulta pasando los dos parámetros
        cursor.execute(consulta, (parametro1, parametro2,parametro3))


        # Recorrer los resultados
        print("Resultados de la consulta:")
        for fila in cursor:
            autorizacion=fila.NoRemision+".pdf"
            rutaautorizacion = os.path.join(carpeta_pdf, autorizacion)

            if os.path.exists(rutaautorizacion):
                print(f"El archivo '{rutaautorizacion}' encontrado")
            else:
                hoja.append([fila.IdAdmision, fila.FechaIngreso, fila.NoFactura,fila.FechaFactura,autorizacion,fila.IdUsuario,fila.NoHistoria])
                
            #if not rutaautorizacion.exists():
       
        workbook.save(archivo_excel)

        # Cerrar la conexión
        cursor.close()

        # Cerrar la conexión
        conexion.close()

    except Exception as e:
        print("Ocurrió un error:", e)
    
    messagebox.showinfo("INFORMACION","fue creado el archivo de cruce de autorizaciones")




# Función para seleccionar la carpeta
def seleccionar_carpeta():
    global carpeta_pdf
    carpeta_pdf = filedialog.askdirectory()
    if carpeta_pdf:
        label_estado.config(text=f"Carpeta seleccionada: {carpeta_pdf}")



# Función para centrar la ventana
def centrar_ventana(ventana, ancho, alto):
    # Obtener dimensiones de la pantalla
    screen_width = ventana.winfo_screenwidth()
    screen_height = ventana.winfo_screenheight()
    
    # Calcular la posición del centro
    x = (screen_width // 2) - (ancho // 2)
    y = (screen_height // 2) - (alto // 2)
    
    # Configurar la geometría centrada
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Selector de Fecha Formato DD-MM-YYYY")

# Dimensiones de la ventana
ancho_ventana = 600
alto_ventana = 400

# Centrar la ventana
centrar_ventana(ventana, ancho_ventana, alto_ventana)

# Etiqueta para instrucción
label = tk.Label(ventana, text="Selecciona una fecha:")
label.pack(pady=10)

# Selector de fecha con formato DD-MM-YYYY
date_entry = DateEntry(
    ventana,
    width=20,
    background="darkblue",
    foreground="white",
    borderwidth=2,
    date_pattern="dd-MM-yyyy"  # Configurar el formato de la fecha
)
date_entry.pack(pady=10)


# Etiqueta de instrucción
label = tk.Label(ventana, text="Ingresa un dato:")
label.pack(pady=10)

# Campo de entrada (Entry) para que el usuario ingrese un texto
entrada = tk.Entry(ventana, width=30)
entrada.pack(pady=10)


# Botón para obtener la fecha


boton_extraer = tk.Button(ventana, text="VERFICAR AUTORIZACIONES", command=verificarAutorizaciones, font=("Arial", 12), bg="#2196F3", fg="white")
boton_extraer.pack(pady=10)


# Etiqueta de estado
label_estado = tk.Label(ventana, text="Ninguna carpeta seleccionada", font=("Arial", 10), bg="#f0f0f0", fg="grey")
label_estado.pack(pady=5)

# Ejecutar el bucle principal
ventana.mainloop()
