import pandas as pd
import os
from libreria_conexion_big_query import create_bigquery_table_from_dataframe, write_dataframe_to_bigquery


GCP_PROJECT_ID = "alpine-realm-352216"  # Cambia por tu proyecto GCP
GCP_DATASET_ID = "galileo"       # Cambia por tu dataset
GCP_TABLE_ID = "resenas_clasificadas"  # Cambia por tu tabla

GOOGLE_JSON_CREDENTIALS_PATH = "alpine-realm-352216-66fb3ad8f36b.json"  # Cambia por la ruta a tus credenciales



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
    
    create_bigquery_table_from_dataframe(
        dataframe=df_resenas, 
        dataframe_name=GCP_TABLE_ID,
        project_id=GCP_PROJECT_ID, 
        dataset_id=GCP_DATASET_ID,
        credentials_path=GOOGLE_JSON_CREDENTIALS_PATH
    )
    
    write_dataframe_to_bigquery(
        dataframe=df_resenas, 
        project_id=GCP_PROJECT_ID, 
        dataset_id=GCP_DATASET_ID,
        table_id=GCP_TABLE_ID,
        credentials_path=GOOGLE_JSON_CREDENTIALS_PATH
    )
    print(f"Datos escritos en BigQuery: {GCP_PROJECT_ID}.{GCP_DATASET_ID}.{GCP_TABLE_ID}")
    
if __name__ == "__main__":
    main()
