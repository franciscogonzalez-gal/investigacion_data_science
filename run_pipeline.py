# -*- coding: utf-8 -*-
"""run_pipeline.py — Orquestador del pipeline completo

Lee el listado de empresas a scrapear desde 'Empresas.xlsx', ejecuta el loop de
web scraping (Trustpilot), luego ejecuta:
  1) procesado_resenas
  2) llm_parse
  3) load_to_bigquery

Uso:
  python run_pipeline.py

Opcional:
  python run_pipeline.py --empresas-xlsx Empresas.xlsx --sheet Hoja1 --playwright

Notas:
- Este orquestador llama a web_scrapping usando su interfaz CLI interna (argparse)
  de forma segura, aislando sys.argv por cada ejecución.
- Requiere que existan/estén configuradas las credenciales:
  - OPENAI_API_KEY (para llm_parse)
  - Credenciales GCP JSON según load_to_bigquery.py
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Iterable, Optional

import pandas as pd

import load_to_bigquery
import llm_parse
import procesado_resenas
import web_scrapping
from logger_library import setup_logger


@dataclass(frozen=True)
class EmpresaSpec:
    company: str
    url: str


@contextmanager
def patched_argv(new_argv: list[str]):
    """Reemplaza temporalmente sys.argv para invocar módulos con argparse."""
    old_argv = sys.argv[:]
    sys.argv = new_argv
    try:
        yield
    finally:
        sys.argv = old_argv


def slugify_company(value: str) -> str:
    s = str(value or "").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    s = s.strip("._-")
    return s.lower() or "empresa"


def normalize_trustpilot_url(value: str) -> str:
    v = str(value or "").strip()
    if not v:
        raise ValueError("URL/empresa vacía")

    if v.startswith("http://") or v.startswith("https://"):
        return v

    # Si viene como host/path sin esquema
    if "trustpilot.com" in v:
        return "https://" + v.lstrip("/")

    # Si viene como dominio (p. ej. sending.es)
    domain = re.sub(r"^https?://", "", v).strip().strip("/")
    return f"https://es.trustpilot.com/review/{domain}"


def _pick_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    lowered = {str(c).strip().lower(): str(c) for c in df.columns}
    for c in candidates:
        if c.lower() in lowered:
            return lowered[c.lower()]
    return None


def read_empresas_from_excel(
    xlsx_path: str,
    sheet: Optional[str] = None,
    company_col: Optional[str] = None,
    url_col: Optional[str] = None,
) -> list[EmpresaSpec]:
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"No se encontró el archivo Excel: {xlsx_path}")

    df = pd.read_excel(xlsx_path, sheet_name=sheet if sheet is not None else 0)
    if df.empty:
        return []

    # Normaliza nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]

    detected_url_col = url_col or _pick_column(
        df,
        [
            "trustpilot_url",
            "url",
            "link",
            "enlace",
            "pagina",
            "página",
        ],
    )

    detected_company_col = company_col or _pick_column(
        df,
        [
            "company",
            "empresa",
            "nombre",
            "name",
            "dominio",
            "domain",
        ],
    )

    # Si no hay columnas reconocibles, usa la primera columna
    if not detected_url_col and not detected_company_col:
        detected_company_col = str(df.columns[0])

    empresas: list[EmpresaSpec] = []
    for _, row in df.iterrows():
        raw_company = row.get(detected_company_col) if detected_company_col else None
        raw_url = row.get(detected_url_col) if detected_url_col else None

        # Derivación flexible:
        # - si hay URL, úsala
        # - si no hay URL, usa el valor de company como dominio/empresa
        if raw_url is not None and str(raw_url).strip():
            url = normalize_trustpilot_url(str(raw_url))
        elif raw_company is not None and str(raw_company).strip():
            url = normalize_trustpilot_url(str(raw_company))
        else:
            continue

        # Company para nombre de archivo
        if raw_company is not None and str(raw_company).strip():
            company = str(raw_company).strip()
        else:
            # si no hay company, intenta inferir desde la URL
            company = url.split("/review/")[-1].strip("/") or url

        empresas.append(EmpresaSpec(company=company, url=url))

    # Dedup por URL
    unique: dict[str, EmpresaSpec] = {}
    for e in empresas:
        unique[e.url] = e
    return list(unique.values())


def run_web_scraping_for_empresas(
    empresas: list[EmpresaSpec],
    review_data_dir: str,
    timeout: int,
    pause: float,
    max_pages: int,
    playwright: bool,
    conservative: bool,
) -> None:
    logger = setup_logger("run_pipeline.web_scraping")
    os.makedirs(review_data_dir, exist_ok=True)

    for e in empresas:
        slug = slugify_company(e.company)
        out_csv = os.path.join(review_data_dir, f"trustpilot_reviews_{slug}.csv")

        argv = [
            "web_scrapping.py",
            "--url",
            e.url,
            "--out",
            out_csv,
            "--timeout",
            str(timeout),
            "--pause",
            str(pause),
            "--max-pages",
            str(max_pages),
        ]
        if playwright:
            argv.append("--playwright")
        if conservative:
            argv.append("--conservative")

        logger.info(f"Scrape: {e.company} -> {e.url}")
        try:
            with patched_argv(argv):
                # No pasamos company=... para no sobreescribir la URL.
                web_scrapping.main()
        except Exception as ex:
            logger.exception(f"Fallo scraping para '{e.company}' ({e.url}): {ex}")


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Ejecuta el pipeline completo (scrape -> procesado -> LLM -> BigQuery).")
    ap.add_argument("--empresas-xlsx", default="Empresas.xlsx", help="Ruta al Excel con el listado de empresas.")
    ap.add_argument("--sheet", default=None, help="Nombre de hoja (sheet) a leer. Si se omite, usa la primera.")
    ap.add_argument("--company-col", default=None, help="Nombre de columna para empresa/company (opcional).")
    ap.add_argument("--url-col", default=None, help="Nombre de columna para URL Trustpilot (opcional).")

    ap.add_argument("--timeout", type=int, default=25, help="Timeout (s) por request en scraping.")
    ap.add_argument("--pause", type=float, default=2.0, help="Pausa (s) entre páginas en scraping.")
    ap.add_argument("--max-pages", type=int, default=500, help="Máximo de páginas por empresa.")
    ap.add_argument("--playwright", action="store_true", help="Usa Playwright como fallback si requests no extrae.")
    ap.add_argument("--conservative", action="store_true", help="Aborta si robots.txt no es legible (modo conservador).")
    return ap.parse_args()


def main() -> None:
    logger = setup_logger("run_pipeline")
    args = parse_args()

    os.makedirs("output", exist_ok=True)
    os.makedirs("review_data", exist_ok=True)

    logger.info("Leyendo Empresas.xlsx...")
    empresas = read_empresas_from_excel(
        xlsx_path=args.empresas_xlsx,
        sheet=args.sheet,
        company_col=args.company_col,
        url_col=args.url_col,
    )
    if not empresas:
        raise RuntimeError("No se encontraron empresas válidas en el Excel.")

    logger.info(f"Empresas a scrapear: {len(empresas)}")
    run_web_scraping_for_empresas(
        empresas=empresas,
        review_data_dir="review_data",
        timeout=args.timeout,
        pause=args.pause,
        max_pages=args.max_pages,
        playwright=args.playwright,
        conservative=args.conservative,
    )

    logger.info("Ejecutando procesado_resenas...")
    procesado_resenas.main()

    logger.info("Ejecutando llm_parse...")
    llm_parse.main()

    logger.info("Ejecutando load_to_bigquery...")
    load_to_bigquery.main()

    logger.info("Pipeline completo finalizado.")


if __name__ == "__main__":
    main()
