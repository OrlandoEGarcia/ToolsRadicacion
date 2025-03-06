import tkinter as tk
from tkinter import ttk

# Crear la ventana principal
root = tk.Tk()
root.title("Ejemplo de ComboBox")
root.geometry("300x200")

# Lista de opciones para el ComboBox
opciones = ["Opción 1", "Opción 2", "Opción 3", "Opción 4"]

# Crear el ComboBox y asignar las opciones
combo = ttk.Combobox(root, values=opciones)
combo.set("Selecciona una opción")  # Establecer valor inicial
combo.pack(pady=20)

# Función para mostrar la selección
def mostrar_seleccion():
    seleccion = combo.get()
    label.config(text=f"Seleccionaste: {seleccion}")

# Botón para mostrar la selección
boton = tk.Button(root, text="Mostrar selección", command=mostrar_seleccion)
boton.pack(pady=10)

# Label para mostrar la opción seleccionada
label = tk.Label(root, text="Selecciona una opción")
label.pack(pady=10)

# Ejecutar el bucle de la aplicación
root.mainloop()
