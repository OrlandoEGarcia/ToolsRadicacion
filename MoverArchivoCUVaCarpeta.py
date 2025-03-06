import os
import shutil

# Ruta de la carpeta con los archivos
source_folder = "J:/2025_Radicacion/_DescargaCuv"

# Recorrer todos los archivos en la carpeta
for filename in os.listdir(source_folder):
    if filename.endswith('.json'):
        # Extraer la parte del nombre del archivo que se usará como nombre de carpeta
        try:
            folder_name = filename.split('_')[2].split('.')[0]
        except IndexError:
            print(f"No se pudo extraer el nombre de carpeta de: {filename}")
            continue

        # Crear la ruta de la nueva carpeta
        folder_path = os.path.join(source_folder, folder_name)

        # Crear la carpeta si no existe
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Mover el archivo a la nueva carpeta
        source_file = os.path.join(source_folder, filename)
        destination_file = os.path.join(folder_path, filename)

        try:
            shutil.move(source_file, destination_file)
            print(f"Archivo movido: {filename} -> {folder_path}")
        except Exception as e:
            print(f"Error al mover {filename}: {e}")

print("Proceso finalizado.")
