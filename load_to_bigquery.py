# -*- coding: utf-8 -*-
"""
Módulo: load_to_bigquery.py
Descripción:
    Script para crear una tabla en Google BigQuery a partir de un DataFrame de pandas
    y para cargar los datos de reseñas preprocesadas a dicha tabla. El DataFrame se
    lee desde un archivo Excel local (por defecto: output/resenas_clasificadas.xlsx).

Funciones principales:
    - create_bigquery_table_from_dataframe(...):
        Crea la tabla en BigQuery inferiendo el esquema a partir del DataFrame.
        (Delegada al módulo libreria_conexion_big_query).
    - write_dataframe_to_bigquery(...):
        Inserta los registros del DataFrame en la tabla BigQuery especificada.
        (Delegada al módulo libreria_conexion_big_query).
    - main():
        Orquesta la lectura del Excel y llama a las funciones anteriores para crear
        la tabla y cargar los datos.

Requisitos:
    - Python 3.8+ (recomendado)
    - pandas
    - google-cloud-bigquery (si la librería de conexión lo utiliza)
    - libreria_conexion_big_query.py accesible en el mismo directorio o en PYTHONPATH
    - Archivo de credenciales JSON de Google Cloud con permisos de BigQuery:
        - roles/bigquery.dataEditor (o superior) para escribir datos
        - roles/bigquery.admin (si se requiere crear datasets/tablas)

Parámetros de configuración (variables del módulo):
    - GCP_PROJECT_ID: ID del proyecto GCP donde se alojará la tabla.
    - GCP_DATASET_ID: ID del dataset BigQuery.
    - GCP_TABLE_ID: Nombre de la tabla a crear/escribir.
    - GOOGLE_JSON_CREDENTIALS_PATH: Ruta local al archivo JSON de credenciales.

Archivo de entrada por defecto:
    - output/resenas_clasificadas.xlsx
      (Debe contener las columnas esperadas por su pipeline; el esquema será inferido.)

Cómo ejecutar:
    1) Colocar el archivo de credenciales JSON en una ruta accesible y
       actualizar GOOGLE_JSON_CREDENTIALS_PATH o establecer la variable de entorno
       GOOGLE_APPLICATION_CREDENTIALS con la ruta al JSON.
    2) Ejecutar desde la terminal (Windows PowerShell o CMD):
       python c:/Users/... /Investigacion/load_to_bigquery.py

Buenas prácticas y notas de seguridad:
    - No subir el archivo de credenciales JSON a repositorios públicos.
    - Preferir el uso de cuentas de servicio con permisos mínimos necesarios.
    - Para cargas grandes, preferir métodos de carga por archivos (GCS -> BigQuery)
      en lugar de insert_rows si la librería lo soporta.
    - Validar tamaño y tipos de datos del DataFrame antes de la carga para evitar
      fallos por incompatibilidades de esquema.

Manejo de errores recomendado:
    - Añadir manejo de excepciones alrededor de las llamadas a la API de BigQuery
      para capturar errores de autenticación, permisos o límites de cuota.
    - Validar que el archivo Excel exista y que pandas pueda leerlo correctamente.


Referencias:
    - https://cloud.google.com/bigquery/docs
    - Documentación de google-cloud-bigquery (si aplica)
"""

import pandas as pd
from libreria_conexion_big_query import create_bigquery_table_from_dataframe, write_dataframe_to_bigquery
from logger_library import setup_logger

GCP_PROJECT_ID = "alpine-realm-352216"  # Cambia por tu proyecto GCP
GCP_DATASET_ID = "galileo"       # Cambia por tu dataset
GCP_TABLE_ID = "resenas_clasificadas"  # Cambia por tu tabla

GOOGLE_JSON_CREDENTIALS_PATH = "alpine-realm-352216-66fb3ad8f36b.json"  # Cambia por la ruta a tus credenciales

def main():
    logger = setup_logger("load_to_bigquery")
    logger.info("Iniciando carga de datos a BigQuery...")
    df_resenas = pd.read_excel("output/resenas_clasificadas.xlsx")
    
    create_bigquery_table_from_dataframe(
        dataframe=df_resenas, 
        dataframe_name=GCP_TABLE_ID,
        project_id=GCP_PROJECT_ID, 
        dataset_id=GCP_DATASET_ID,
        credentials_path=GOOGLE_JSON_CREDENTIALS_PATH,
        replace_if_exists=True
    )
    
    write_dataframe_to_bigquery(
        dataframe=df_resenas, 
        project_id=GCP_PROJECT_ID, 
        dataset_id=GCP_DATASET_ID,
        table_id=GCP_TABLE_ID,
        credentials_path=GOOGLE_JSON_CREDENTIALS_PATH
    )
    logger.info(f"Datos escritos en BigQuery: {GCP_PROJECT_ID}.{GCP_DATASET_ID}.{GCP_TABLE_ID}")

if __name__ == "__main__":
    main()