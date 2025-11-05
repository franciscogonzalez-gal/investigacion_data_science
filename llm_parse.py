# -*- coding: utf-8 -*-
"""
llm_parse.py — Clasificación de reseñas usando la Responses API de OpenAI (JSON -> DataFrame -> Excel)

Resumen
-------
Este módulo procesa un CSV de reseñas (columna por defecto 'body'), envía cada reseña a la Responses API
de OpenAI usando un system prompt controlado, y guarda los resultados en un archivo Excel. El modelo
debe devolver exclusivamente un objeto JSON con tres campos obligatorios:
  - sentiment_label: "Positiva" | "Negativa"
  - general_category: una etiqueta de la lista GENERAL
  - specific_category: una etiqueta válida de la lista ESPECÍFICAS (según polaridad)

Características principales
---------------------------
- Lector CSV con pandas.
- Cliente OpenAI (openai.OpenAI) configurado mediante la variable de entorno OPENAI_API_KEY.
- Envío de prompt de sistema + prompt de usuario (reseña).
- Parsing robusto del texto de respuesta a JSON (con varios fallbacks).
- Validación de campos obligatorios.
- Reintentos con backoff exponencial configurable.
- Exportación final a Excel (openpyxl).

Requisitos
----------
- Python 3.8+
- Paquetes: openai (OpenAI Python client), pandas, openpyxl, python-dotenv (opcional)
  Instalación: pip install openai pandas openpyxl python-dotenv

Variables / Rutas configurables (constantes del módulo)
------------------------------------------------------
- INPUT_CSV: ruta del CSV de entrada (por defecto "output/resenas_combinadas.csv")
  - El CSV debe contener, como mínimo, la columna TEXT_COLUMN (por defecto "body").
- TEXT_COLUMN: nombre de la columna que contiene el texto de la reseña.
- ID_COLUMN: columna opcional con identificadores de reseñas (ej. "review_id").
- OUTPUT_XLSX: ruta del archivo Excel de salida.
- MODEL_NAME: nombre del modelo a usar (p. ej. "gpt-5").
- MAX_RETRIES / INITIAL_BACKOFF: controlan los reintentos y el backoff exponencial.

Seguridad y configuración de la API
-----------------------------------
- La clave de API debe estar en la variable de entorno OPENAI_API_KEY.
  - Windows (ej.): setx OPENAI_API_KEY "sk-..."
  - macOS/Linux (ej.): export OPENAI_API_KEY="sk-..."
- Se puede usar un archivo .env en el directorio de trabajo; python-dotenv carga automáticamente las variables.

Formato esperado de la respuesta del modelo
------------------------------------------
- El modelo debe responder únicamente con JSON válido, sin texto adicional.
- Estructura obligatoria:
  {
    "sentiment_label": "Positiva" | "Negativa",
    "general_category": "<etiqueta GENERAL>",
    "specific_category": "<etiqueta específica válida para la polaridad>"
  }
- Si el modelo devuelve texto extra (comentarios, Markdown), el módulo intentará extraer el primer bloque JSON
  o fallará con una excepción descriptiva.

Manejo de errores y fallbacks
-----------------------------
- El cliente implementa reintentos con backoff exponencial (controlados por MAX_RETRIES e INITIAL_BACKOFF).
- Si la respuesta no puede parsearse a JSON o faltan campos, se captura el error y la fila se registra con
  campos nulos y una clave "error" con el mensaje, para mantener trazabilidad en el Excel de salida.
- Errores fatales (como ausencia de OPENAI_API_KEY o inexistencia del CSV de entrada) lanzan excepciones
  para que el proceso termine con diagnóstico claro.

Salida / columnas en el Excel resultante
----------------------------------------
Cada fila de salida contendrá, como mínimo:
- review_id (si INPUT CSV tiene ID_COLUMN)
- review (texto original)
- sentiment_label
- general_category
- specific_category
Además, si existen en el CSV original, se copian columnas comunes: author, rating, date_published, location,
source_url, title, company. Se recomienda revisar el Excel resultante para validar la calidad del parsing.

Cómo usar
--------
1. Preparar el CSV de entrada en INPUT_CSV con una columna que contenga la reseña (TEXT_COLUMN).
2. Exportar/definir OPENAI_API_KEY en el entorno.
3. Ejecutar: python llm_parse.py
   - Para tests locales o depuración, se puede reducir MAX_RETRIES o cambiar MODEL_NAME.
4. Revisar el archivo OUTPUT_XLSX generado.

Consideraciones prácticas
-------------------------
- Coste y latencia: llamar a la API por cada reseña tiene coste y latencia. Para lotes grandes, considerar
  batching o límites de tasa (throttling).
- Auditoría: se conserva la reseña en el resultado para trazabilidad. Guardar también la respuesta cruda si
  se necesita auditar decisiones del modelo.
- Consistencia: el system prompt fuerza formato JSON estricto; aun así, conviene revisar muestras para validar
  que las etiquetas devueltas respetan las listas permitidas.

Extendiendo el script
---------------------
- Añadir caching local de respuestas para evitar reconsultas.
- Paralelizar llamadas respetando límites de la API.
- Añadir logging a fichero en lugar de prints.
- Añadir validación más estricta de las etiquetas devueltas frente a las listas permitidas.

Licencia y privacidad
---------------------
- No incluir datos sensibles ni PII en las reseñas si la política de privacidad lo prohíbe.
- Este script es una plantilla; adaptar su uso a la normativa de protección de datos aplicable.

"""

import os
import re
import time
import json
import pandas as pd
from typing import Dict, Any
from dotenv import load_dotenv
from openai import OpenAI
from logger_library import setup_logger

load_dotenv()
# -----------------------------
# Configuración
# -----------------------------
INPUT_CSV = "output/resenas_combinadas.csv"      # CSV generado por tu scraper (columna 'body')
TEXT_COLUMN = "body"                   # Columna con el texto de la reseña
ID_COLUMN = "review_id"               # Opcional, si existe en tu CSV
OUTPUT_XLSX = "output/resenas_clasificadas.xlsx"


MODEL_NAME = "gpt-5"                # Ajusta si usas otra variante (p.ej. "gpt-5.1")
MAX_RETRIES = 4
INITIAL_BACKOFF = 2.0

# -----------------------------
# System Prompt (del mensaje anterior)
# -----------------------------
SYSTEM_PROMPT = r"""
Eres un clasificador de reseñas de clientes en español (y puedes manejar texto mixto ES/EN). Recibirás exactamente una reseña como texto de entrada. Debes responder únicamente con un objeto JSON válido (sin texto adicional, sin comentarios, sin Markdown).

Objetivo: devolver 4 campos obligatorios:
- "sentiment_label": "Positiva" o "Negativa".
- "general_category": una etiqueta de la lista GENERAL.
- "specific_category": una etiqueta de la lista específica según la polaridad.

Reglas:
1) Si la reseña contiene varios temas, elige el más central.
2) Si sentimiento y categoría chocan, etiqueta por el sentimiento actual.
3) Si no hay información suficiente para una categoría específica, usa "Otros".
4) Mantén la capitalización exacta de las etiquetas permitidas.

GENERAL:
- "Entrega"
- "Recogida y logística inversa"
- "Seguimiento y comunicación"
- "Servicio al cliente"
- "Compensación y reembolso"
- "Calidad del producto entregado"
- "Repartidor"
- "Experiencia general"
- "Valor percibido"
- "Fidelización"
- "Responsabilidad y recuperación"

ESPECÍFICAS NEGATIVAS:
- "Falta de entrega"
- "Retraso en la entrega"
- "Entrega en dirección incorrecta"
- "Entrega sin aviso o contacto"
- "Entrega dañada"
- "Entrega fuera de horario o zona"
- "No se presentó a recoger"
- "Retraso en recogida"
- "Problemas con punto de recogida"
- "Seguimiento incorrecto o sin actualizar"
- "Comunicación inexistente o deficiente"
- "Información confusa o contradictoria"
- "Falta de respuesta a reclamaciones"
- "Atención poco profesional o grosera"
- "Derivación o evasión de responsabilidad"
- "No reembolsan producto o envío"
- "Procesos de reclamo ineficaces"
- "Daño físico al producto"
- "Contenido incompleto o perdido"
- "Repartidor poco profesional"
- "Proceso ineficiente o burocrático"
- "Costo excesivo frente a servicio"
- "Empresa no confiable"
- "Otros"

ESPECÍFICAS POSITIVAS:
- "Entrega puntual"
- "Entrega rápida"
- "Entrega correcta"
- "Entrega en buenas condiciones"
- "Entrega flexible o conveniente"
- "Buen seguimiento"
- "Comunicación efectiva"
- "Aviso previo o confirmación"
- "Atención rápida y resolutiva"
- "Atención amable o profesional"
- "Buena gestión de reclamaciones"
- "Repartidor amable o educado"
- "Repartidor puntual o responsable"
- "Repartidor proactivo"
- "Servicio confiable"
- "Satisfacción general"
- "Profesionalismo"
- "Rapidez de respuesta"
- "Buena relación calidad-precio"
- "Expectativas superadas"
- "Recomendación a otros"
- "Repetición de compra o uso"
- "Resolución satisfactoria de errores"
- "Compromiso con el cliente"
- "Otros"

Responde siempre con solo JSON con esta estructura:
{
  "sentiment_label": "Positiva" | "Negativa",
  "general_category": "<una etiqueta GENERAL>",
  "specific_category": "<una etiqueta específica válida para la polaridad>"
}
""".strip()


# -----------------------------
# Cliente OpenAI
# -----------------------------
def get_openai_client() -> OpenAI:
    # Carga variables desde un archivo .env en el directorio de trabajo (si existe)
    

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta la variable de entorno OPENAI_API_KEY. Añádela en .env o exporta la variable.")
    return OpenAI(api_key=api_key)


def call_openai_json(review_text: str, client: OpenAI) -> Dict[str, Any]:
    """
    Llama a la Responses API con system prompt y devuelve un dict con el JSON.
    Implementa reintentos con backoff y parsing robusto.
    """
    last_err = None
    backoff = INITIAL_BACKOFF

    for _ in range(MAX_RETRIES):
        try:
            resp = client.responses.create(
                model=MODEL_NAME,
                reasoning={"effort": "low"},
                input=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    # Si quieres añadir un developer prompt (ej. "Talk like a pirate."),
                    # añade aquí otro dict con role="developer".
                    {"role": "user", "content": review_text},
                ],
            )

            # En tu ejemplo: response.output_text
            raw = getattr(resp, "output_text", None)
            if raw is None:
                # Algunas versiones devuelven en resp.output[0].content[0].text
                # Fallback genérico:
                raw = str(resp)

            data = coerce_to_json(raw)
            validate_fields(data)
            return data

        except Exception as e:
            last_err = e
            time.sleep(backoff)
            backoff = min(backoff * 2, 30)

    raise RuntimeError(f"Fallo clasificando reseña tras reintentos: {last_err}")


def coerce_to_json(s: str) -> Dict[str, Any]:
    """
    Intenta parsear 's' a JSON. Si viene con texto extra o Markdown, limpia y vuelve a intentar.
    """
    s = s.strip()

    # Intento directo
    try:
        return json.loads(s)
    except Exception:
        pass

    # Extraer el primer bloque {...} más externo
    # Captura greedy para incluir anidados, luego valida
    m = re.search(r"\{.*\}", s, flags=re.DOTALL)
    if m:
        candidate = m.group(0)
        try:
            return json.loads(candidate)
        except Exception:
            pass

    # Último recurso: componer un dict mínimo si el modelo devolvió claves como líneas
    # (no debería ocurrir con el prompt, pero dejamos este fallback)
    raise ValueError(f"No se pudo parsear a JSON: {s[:300]}...")


def validate_fields(d: Dict[str, Any]) -> None:
    required = {"sentiment_label", "general_category", "specific_category"}
    missing = required - set(d.keys())
    if missing:
        raise ValueError(f"Faltan campos obligatorios: {missing}")


# -----------------------------
# Carga CSV, clasificación y guardado
# -----------------------------
def main():
    # Cargar reseñas
    logger = setup_logger("llm_parse")
    if not os.path.exists(INPUT_CSV):
        raise FileNotFoundError(f"No se encontró el CSV de entrada: {INPUT_CSV}")

    df = pd.read_csv(INPUT_CSV)
    if TEXT_COLUMN not in df.columns:
        raise KeyError(f"El CSV debe contener la columna '{TEXT_COLUMN}' con el texto de la reseña.")

    client = get_openai_client()

    results = []
    for idx, row in df.iterrows():
        review_text = str(row.get(TEXT_COLUMN, "")).strip()
        if not review_text:
            continue

        try:
            data = call_openai_json(review_text, client)
            logger.info(f"Fila {idx} (ID: {row.get(ID_COLUMN, 'N/A')}): Clasificada correctamente.")
        except Exception as e:
            logger.info(f"Error en fila {idx} (ID: {row.get(ID_COLUMN, 'N/A')}): {e}")
            # Si algo falla, registramos un resultado con error para que no se pierda la fila
            data = {
                "sentiment_label": None,
                "general_category": None,
                "specific_category": None,
                "error": str(e),
            }

        # Enlazar con campos originales útiles (id/fecha/origen si están)
        enriched = {
            "review_id": row.get(ID_COLUMN) if ID_COLUMN in df.columns else None,
            "review": review_text,
            "sentiment_label": data.get("sentiment_label"),
            "general_category": data.get("general_category"),
            "specific_category": data.get("specific_category"),
            # Copiamos la reseña devuelta por el modelo por trazabilidad (puede coincidir con la original)
        }

        # Copia extra de columnas originales si quieres mantenerlas (fecha, rating, author, etc.)
        for col in ["author", "rating", "date_published", "location", "source_url", "title","company"]:
            if col in df.columns:
                enriched[col] = row.get(col)

        results.append(enriched)

    out_df = pd.DataFrame(results)

    # Guardar a Excel (openpyxl)
    out_df.to_excel(OUTPUT_XLSX, index=False)
    logger.info(f"Listo. Guardado en: {OUTPUT_XLSX}\nFilas: {len(out_df)}")
    
   

if __name__ == "__main__":
    main()
