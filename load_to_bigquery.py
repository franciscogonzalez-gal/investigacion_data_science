import pandas as pd
from libreria_conexion_big_query import create_bigquery_table_from_dataframe, write_dataframe_to_bigquery


GCP_PROJECT_ID = "alpine-realm-352216"  # Cambia por tu proyecto GCP
GCP_DATASET_ID = "galileo"       # Cambia por tu dataset
GCP_TABLE_ID = "resenas_clasificadas"  # Cambia por tu tabla

GOOGLE_JSON_CREDENTIALS_PATH = "alpine-realm-352216-66fb3ad8f36b.json"  # Cambia por la ruta a tus credenciales

def main():
    df_resenas = pd.read_excel("output/resenas_clasificadas.xlsx")
    
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