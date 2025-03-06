import os

# Ruta principal donde están las subcarpetas
base_path = "J:/_Revision"

# Prefijos de los archivos esperados
expected_files = [
    "FRCUV_900077584",
    "FRJSON_900077584",
    "FRXML_900077584"
]

# Función para verificar si una subcarpeta contiene los archivos necesarios
def folder_has_all_files(folder_path):
    files = os.listdir(folder_path)
    
    # Verificar los archivos con los prefijos específicos
    for prefix in expected_files:
        if not any(f.startswith(prefix) for f in files):
            return False
    
    # Verificar si hay al menos un archivo zip
    if not any(f.endswith('.zip') for f in files):
        return False

    return True

# Recorrer todas las subcarpetas
for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)

    # Verificar si es una carpeta
    if os.path.isdir(folder_path):
        if not folder_has_all_files(folder_path):
            # Renombrar la carpeta si falta algún archivo
            new_folder_name = f"_{folder}"
            new_folder_path = os.path.join(base_path, new_folder_name)
            os.rename(folder_path, new_folder_path)
            print(f"Carpeta renombrada: {folder} -> {new_folder_name}")

print("Proceso finalizado.")
