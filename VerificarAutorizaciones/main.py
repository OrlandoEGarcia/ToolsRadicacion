from tkinter import *
from tkinter import ttk
from tkcalendar import *


import fitz  # PyMuPDF
import pandas as pd
import os
import re 
import tkinter as tk
import openpyxl
from tabulate import tabulate

import pyodbc

from tkinter import filedialog, messagebox

from tkcalendar import DateEntry
from tkinter import messagebox

import pyodbc
from openpyxl import Workbook

from datetime import datetime


# Configuración de conexión
server = '192.168.59.230'  # Ejemplo: 'localhost' o '192.168.1.100'
database = 'sifacturacion'      # Nombre de la base de datos
username = 'AdminFacturacion'        # Usuario de SQL Server
password = 'SI.Admin.23$%*'     # Contraseña del usuario


Lista = ['IBARRA59','BOLAÑOS27','NARVAEZ27','BELALCAZAR108','']




def verificarAutorizaciones():
    fecha = calendario.get_date()  # Obtiene la fecha como cadena en formato DD-MM-YYYY
    fecha2=calendariof.get_date()
    
    fechaNombre=fecha.replace("/","")
    fechaNombre2 =fecha2.replace("/","")
    fnombre=fechaNombre+"_"+fechaNombre2
    nombre = combo.get()

    
    #messagebox.showinfo("Fecha seleccionada", f"Has seleccionado: {fecha}")
    
    # Crear un archivo Excel con openpyxl
    #archivo_excel = os.path.join(carpeta_pdfAut, fnombre+"_AutorizacionesFaltantes"+".xlsx")
    archivo_excel = nombre+"_"+fnombre+"_"+"_Autorizaciones"+".xlsx"
    workbook = Workbook()
    hoja = workbook.active
    hoja.title = "Resultados"
    hoja.append(["IdAdmision","FechaIngreso","NoFactura","FechaFactura","NoRemision","IdUsuario","NoHistoria"])  # Encabezados del archivo    
    

    # Crear un archivo Excel con openpyxl
    #archivo_excelFirmas = os.path.join(carpeta_pdfFir, fnombre+"_FirmasFaltantes"+".xlsx")
    archivo_excelFirmas =nombre+"_"+fnombre+"_"+"_Firmas"+".xlsx"
    workbook2 = Workbook()
    hoja2 = workbook2.active
    hoja2.title = "Resultados"
    hoja2.append(["Usuario"])  # Encabezados del archivo    



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
        parametro3 = fecha2 + " 23:59:59"
        

        # Ejecutar la consulta pasando los dos parámetros
        cursor.execute(consulta, (parametro1, parametro2,parametro3))


        # Recorrer los resultados
        #print("Resultados de la consulta:")
        for fila in cursor:
            autorizacion=fila.NoRemision+".pdf"
            firma=fila.IdUsuario+".pdf"
            rutaautorizacion = os.path.join(carpeta_pdfAut, autorizacion)
            rutafirma =os.path.join(carpeta_pdfFir, firma)

            if os.path.exists(rutaautorizacion):
                encaut=1
                #print(f"El archivo '{rutaautorizacion}' encontrado")
            else:
                hoja.append([fila.IdAdmision, fila.FechaIngreso, fila.NoFactura,fila.FechaFactura,autorizacion,fila.IdUsuario,fila.NoHistoria])
                
            if os.path.exists(rutafirma):
                encfir=1
                #print(f"El archivo '{rutafirma}' encontrado")
            else:
                hoja2.append([fila.IdUsuario])

            #if not rutaautorizacion.exists():
       
        workbook.save(archivo_excel)
        workbook2.save(archivo_excelFirmas)

        # Cerrar la conexión
        cursor.close()

        # Cerrar la conexión
        conexion.close()

    except Exception as e:
        print("Ocurrió un error:", e)
    
    messagebox.showinfo("INFORMACION","fue creado el archivo de cruce de autorizaciones")




def seleccionar_carpetaAutorizaciones():
    global carpeta_pdfAut
    carpeta_pdfAut = filedialog.askdirectory()
    if carpeta_pdfAut:
        textoAutorizaciones.config(text=f"{carpeta_pdfAut}")

def seleccionar_carpetaFirmas():
    global carpeta_pdfFir
    carpeta_pdfFir = filedialog.askdirectory()
    if carpeta_pdfFir:
        textoFirmas.config(text=f"{carpeta_pdfFir}")


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


alto=400
ancho=560

ventana = Tk()
ventana.minsize(ancho,alto)
ventana.title("VERIFICACION DE ARCHIVOS")

# Centrar la ventana
centrar_ventana(ventana, ancho, alto)

calendario= Calendar(ventana, setmode="day", date_pattern='dd/mm/yyyy')
calendario.place(x=10,y=10)

calendariof= Calendar(ventana, setmode="day", date_pattern='dd/mm/yyyy')
calendariof.place(x=300,y=10)


# Crear el ComboBox y asignar las opciones
textoCombo = ttk.Label(ventana,text="USUARIO: ")
textoCombo.place(x=10,y=210)

combo = ttk.Combobox(ventana, values=Lista, state='readonly',width=25)
combo.place(x=100,y=210)

    
rutaAutorizaciones=Button(ventana,text="Ruta Autorizaciones",command=seleccionar_carpetaAutorizaciones, width=20)
rutaAutorizaciones.place(x=10,y=240)

textoAutorizaciones = ttk.Label(ventana,text="Selecione la carpeta de autorizaciones")
textoAutorizaciones.place(x=170,y=240)


rutaFirmas=Button(ventana,text="Ruta Firmas",command=seleccionar_carpetaFirmas, width=20)
rutaFirmas.place(x=10,y=280)

textoFirmas = ttk.Label(ventana,text="Selecione la carpeta de firmas")
textoFirmas.place(x=170,y=280)


btnVerificar=Button(ventana,text="Verificar",command=verificarAutorizaciones, font=("Arial", 12), bg="#2196F3", fg="white", width=50)
btnVerificar.place(x=20,y=320)




# Ejecutar el bucle principal
ventana.mainloop()



