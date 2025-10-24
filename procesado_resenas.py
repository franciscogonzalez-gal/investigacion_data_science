"""
Módulo: procesado_resenas.py
Resumen:
    Herramientas para cargar múltiples archivos CSV de reseñas (exportados de Trustpilot),
    combinar los datos en un único DataFrame, normalizar el nombre de la compañía e
    exportar las reseñas combinadas a Excel y CSV.

Propósito:
    - Leer todos los archivos .csv en una carpeta de entrada.
    - Añadir una columna 'company' derivada del nombre de archivo (se elimina el prefijo
      "trustpilot_reviews_" y se convierte a mayúsculas).
    - Concatenar todos los DataFrames leídos en un único DataFrame.
    - Filtrar (por ejemplo eliminar registros de la compañía "SAMPLE").
    - Guardar el resultado en formatos .xlsx y .csv.

Requisitos:
    - Python 3.8+.
    - pandas instalado (pip install pandas).
    - Permisos de lectura en la carpeta de entrada y de escritura en la carpeta de salida.

Estructura de archivos esperada:
    carpeta_entrada (por defecto "review_data")
        ├─ trustpilot_reviews_COMPANY1.csv
        ├─ trustpilot_reviews_COMPANY2.csv
        └─ ...
    salida esperada:
        output/resenas_combinadas.xlsx
        output/resenas_combinadas.csv

Funciones públicas:
    - cargar_resenas_csv(carpeta: str) -> pandas.DataFrame
        Lee todos los .csv de la carpeta indicada, añade la columna 'company' y
        devuelve un DataFrame con todas las reseñas concatenadas.
        Excepciones:
            - FileNotFoundError: si no se encuentran archivos CSV en la carpeta.

    - guardar_resenas_excel(df: pandas.DataFrame, ruta_salida: str) -> None
        Guarda el DataFrame dado en formato Excel (.xlsx). No crea automáticamente
        el directorio padre; debe existir o provocará un error de escritura.

    - guardar_resenas_csv(df: pandas.DataFrame, ruta_salida: str) -> None
        Guarda el DataFrame dado en formato CSV. Igual consideración de directorio.

Uso (línea de comandos):
    python procesado_resenas.py
    - Ajustar la variable `carpeta_entrada` dentro de `main()` o modificar las rutas
      si se desea usar ubicaciones diferentes.

Notas de implementación:
    - El nombre de la compañía se obtiene desde el nombre de archivo removiendo el
      prefijo literal "trustpilot_reviews_". Si los nombres de archivo no siguen este
      patrón, la etiqueta 'company' puede quedar con el nombre completo del archivo.
    - Pandas usa la codificación por defecto para leer CSV; para CSV con codificaciones
      distintas (p. ej. 'utf-8-sig' o 'latin-1') puede ser necesario pasar `encoding=...`
      a pd.read_csv.
    - Se añade la columna 'company' en mayúsculas: df['company'] = nombre_sin_extension.upper()
    - El módulo actualmente filtra explícitamente filas con company == 'SAMPLE' en main().
    - Mejoras recomendadas: añadir manejo de logging, validación de columnas esperadas,
      y crear automáticamente directorios de salida si no existen (os.makedirs(..., exist_ok=True)).

Ejemplo mínimo de personalización:
    - Cambiar `carpeta_entrada` por la ruta absoluta de la carpeta con CSVs.
    - Crear la carpeta `output/` antes de ejecutar o adaptar las funciones para crearla.

Licencia / Autor:
    - Crear o añadir la información de autor y licencia según proceda en el proyecto.
"""

import pandas as pd
import os


def cargar_resenas_csv(carpeta: str) -> dict:
    dataframes = {}
    for archivo in os.listdir(carpeta):
        if archivo.endswith(".csv"):
            ruta_completa = os.path.join(carpeta, archivo)
            nombre_sin_extension = os.path.splitext(archivo)[0]
            #dejar el nombre de la comania como nombre del dataframe, removiendo el texto trustpilot_reviews_
            nombre_sin_extension = nombre_sin_extension.replace("trustpilot_reviews_", "")
            df = pd.read_csv(ruta_completa)
            df['company'] = nombre_sin_extension.upper()
            dataframes[nombre_sin_extension] = df
    if not dataframes:
        raise FileNotFoundError(f"No se encontraron archivos CSV en la carpeta: {carpeta}")
    else:
        data = pd.concat(dataframes.values(), ignore_index=True)
    return data

def guardar_resenas_excel(df: pd.DataFrame, ruta_salida: str) -> None:
    df.to_excel(ruta_salida, index=False)
    print(f"Reseñas guardadas en: {ruta_salida}")
    
#guadar las resenas en CSV
def guardar_resenas_csv(df: pd.DataFrame, ruta_salida: str) -> None:
    df.to_csv(ruta_salida, index=False)
    print(f"Reseñas guardadas en: {ruta_salida}")

def main():
    carpeta_entrada = "review_data"
    # Cargar reseñas desde CSVs
    df_resenas = cargar_resenas_csv(carpeta_entrada)
    
    #eliminar las filas de la compania SAMPLE
    df_resenas = df_resenas[df_resenas['company'] != 'SAMPLE']

    # Guardar reseñas en un archivo Excel
    guardar_resenas_excel(df_resenas, "output/resenas_combinadas.xlsx")
    
    # Guardar reseñas en un archivo CSV
    guardar_resenas_csv(df_resenas, "output/resenas_combinadas.csv")
    
   
    
if __name__ == "__main__":
    main()
