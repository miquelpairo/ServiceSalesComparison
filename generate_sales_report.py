# -*- coding: utf-8 -*-
"""
BUCHI - Sales Report Orchestrator
===================================
Lee el Excel exportado de Power BI, calcula los dos periodos YTD
(año anterior vs año actual hasta la misma fecha), y genera el HTML
con generate_sales_comparison_html().

Uso standalone:
    python generate_sales_report.py
    python generate_sales_report.py --file exports/pbi_sales_20260301_094523.xlsx
    python generate_sales_report.py --file exports/pbi_sales_20260301_094523.xlsx --output reports/

Uso desde Node-RED / selenium_pbi_export.py:
    Llama a run(xlsx_path) y recibe el path del HTML generado.
"""

import argparse
import sys
import os
from datetime import datetime, date
from pathlib import Path
import pandas as pd

# ---------------------------------------------------------------------------
# Path setup: forzar working directory al directorio del script
# Esto es necesario cuando el script lo lanza Node-RED u otro proceso externo,
# ya que el cwd puede ser distinto y los CSS/templates no se encuentran.
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).parent.resolve()
os.chdir(SCRIPT_DIR)  # forzar cwd al directorio del script
sys.path.insert(0, str(SCRIPT_DIR))

from sales_comparison_report_generator import generate_sales_comparison_html
from column_mappings import (
    detect_format,
    get_mapping_for_format,
    get_additional_columns,
    validate_format,
)

# =============================================================================
# CONFIG
# =============================================================================

# Carpeta donde selenium_pbi_export.py deja los xlsx
DEFAULT_EXPORTS_DIR = Path(r"C:\Users\pairo\OneDrive\Documentos\MP Server\service-report\exports")

# Carpeta donde se guardan los HTML generados
DEFAULT_OUTPUT_DIR  = Path(r"C:\Users\pairo\OneDrive\Documentos\MP Server\service-report\reports")

# =============================================================================
# HELPERS
# =============================================================================

def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def find_latest_export(exports_dir: Path) -> Path:
    """Devuelve el xlsx mas reciente en la carpeta de exports."""
    files = sorted(exports_dir.glob("pbi_sales_*.xlsx"), key=lambda f: f.stat().st_mtime)
    if not files:
        raise FileNotFoundError(f"No hay ficheros pbi_sales_*.xlsx en {exports_dir}")
    return files[-1]


def load_and_normalize(xlsx_path: Path) -> pd.DataFrame:
    """
    Carga el Excel, detecta el formato y normaliza las columnas
    al esquema interno estandar.
    """
    log(f"Cargando: {xlsx_path.name}")
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    log(f"  Filas cargadas: {len(df)}  |  Columnas: {list(df.columns)}")

    # Detectar formato
    fmt = detect_format(df.columns.tolist())
    log(f"  Formato detectado: {fmt}")
    if fmt == "unknown":
        raise ValueError(
            f"Formato de columnas no reconocido.\n"
            f"Columnas encontradas: {df.columns.tolist()}"
        )

    # Validar columnas requeridas
    is_valid, missing = validate_format(df, fmt)
    if not is_valid:
        raise ValueError(f"Columnas requeridas no encontradas: {missing}")

    # Renombrar al esquema interno
    mapping = get_mapping_for_format(fmt)          # {nombre_interno: nombre_real}
    rename_map = {v: k for k, v in mapping.items()}  # invertir: real -> interno
    df = df.rename(columns=rename_map)

    # Conservar columnas adicionales (Market Org, Territory, Country, SFDC Link, etc.)
    # mixed y new comparten las mismas columnas extra -> usar siempre 'new'
    extra = get_additional_columns('new')
    all_cols = list(mapping.keys()) + [c for c in extra if c in df.columns]
    df = df[[c for c in all_cols if c in df.columns]]

    # Convertir Date
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    # EUR es el campo interno; en mixed format viene de LC
    if "EUR" in df.columns:
        df["EUR"] = pd.to_numeric(df["EUR"], errors="coerce").fillna(0)

    # Columnas calculadas que Streamlit añade
    df["Amount"]     = df["EUR"]
    df["Year"]       = df["Date"].dt.year
    df["Month"]      = df["Date"].dt.month
    df["Month_Name"] = df["Date"].dt.strftime("%B")

    # FC y CHF numericas si existen
    for col in ["FC", "CHF"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    log(f"  Normalizacion OK. Rango de fechas: {df['Date'].min().date()} -> {df['Date'].max().date()}")
    return df


def compute_ytd_periods(df: pd.DataFrame, reference_date: date = None):
    """
    Calcula los dos periodos YTD:
      - P1: 1 enero año_anterior  ->  reference_date con año_anterior
      - P2: 1 enero año_actual    ->  reference_date

    reference_date por defecto = hoy
    Devuelve (df_p1, df_p2, nombre_p1, nombre_p2)
    """
    if reference_date is None:
        reference_date = date.today()

    year_actual   = reference_date.year
    year_anterior = year_actual - 1

    # Misma fecha del año anterior
    try:
        ref_anterior = reference_date.replace(year=year_anterior)
    except ValueError:
        # 29 feb en año no bisiesto
        ref_anterior = date(year_anterior, 2, 28)

    start_p1 = date(year_anterior, 1, 1)
    end_p1   = ref_anterior

    start_p2 = date(year_actual, 1, 1)
    end_p2   = reference_date

    log(f"  Periodo 1: {start_p1} -> {end_p1}")
    log(f"  Periodo 2: {start_p2} -> {end_p2}")

    mask_p1 = (df["Date"].dt.date >= start_p1) & (df["Date"].dt.date <= end_p1)
    mask_p2 = (df["Date"].dt.date >= start_p2) & (df["Date"].dt.date <= end_p2)

    df_p1 = df[mask_p1].copy()
    df_p2 = df[mask_p2].copy()

    nombre_p1 = f"YTD {year_anterior} ({start_p1.strftime('%d/%m')} - {end_p1.strftime('%d/%m/%Y')})"
    nombre_p2 = f"YTD {year_actual} ({start_p2.strftime('%d/%m')} - {end_p2.strftime('%d/%m/%Y')})"

    log(f"  Filas P1: {len(df_p1)}  |  Filas P2: {len(df_p2)}")
    return df_p1, df_p2, nombre_p1, nombre_p2


def build_comparativa(df_p1: pd.DataFrame, df_p2: pd.DataFrame,
                      nombre_p1: str, nombre_p2: str) -> pd.DataFrame:
    """
    Construye el DataFrame de comparativa por Business Partner + ItemIdAndName
    con el mismo esquema que genera Streamlit.
    """
    key_cols   = ["Business Partner Name", "ItemIdAndName"]
    # Columnas extra a preservar (se toma el primer valor de cada grupo)
    extra_cols = ["SalesRepresentative", "Set", "Productline", "ProductType"]

    def agg_amount(df, year_label):
        return df.groupby(key_cols)["EUR"].sum().rename(f"Amount {year_label}")

    def agg_qty(df, year_label):
        return df.groupby(key_cols)["Qty"].sum().rename(f"Quantity {year_label}")

    # Extraer año de los nombres de periodo (p.e. "YTD 2025 (...)" -> "2025")
    import re as _re
    def extract_year(nombre):
        m = _re.search(r"(\d{4})", nombre)
        return m.group(1) if m else nombre

    y1 = extract_year(nombre_p1)
    y2 = extract_year(nombre_p2)

    amt_p1 = agg_amount(df_p1, y1)
    amt_p2 = agg_amount(df_p2, y2)
    qty_p1 = agg_qty(df_p1, y1)
    qty_p2 = agg_qty(df_p2, y2)

    comp = pd.concat([amt_p1, amt_p2, qty_p1, qty_p2], axis=1).fillna(0)

    comp[f"Amount Difference"]   = comp[f"Amount {y2}"]   - comp[f"Amount {y1}"]
    comp[f"Quantity Difference"] = comp[f"Quantity {y2}"] - comp[f"Quantity {y1}"]
    comp["Growth %"] = comp.apply(
        lambda r: (r[f"Amount Difference"] / r[f"Amount {y1}"] * 100)
                  if r[f"Amount {y1}"] != 0 else 0,
        axis=1
    )

    # Añadir columnas extra (SalesRep, Set, etc.) desde ambos periodos combinados
    combined = pd.concat([df_p1, df_p2], ignore_index=True)
    for col in extra_cols:
        if col in combined.columns:
            extra = (
                combined[combined[col].notna()]
                .groupby(key_cols)[col]
                .first()
            )
            comp = comp.join(extra, how="left")

    return comp


def get_available_lists(df_p1: pd.DataFrame, df_p2: pd.DataFrame):
    """Extrae las listas de filtros disponibles (union de ambos periodos)."""
    combined = pd.concat([df_p1, df_p2], ignore_index=True)

    def unique_sorted(col):
        if col not in combined.columns:
            return []
        return sorted(combined[col].dropna().astype(str).unique().tolist())

    return {
        "available_types":        unique_sorted("ProductType"),
        "available_sets":         unique_sorted("Set"),
        "available_reps":         unique_sorted("SalesRepresentative"),
        "available_market_orgs":  unique_sorted("Market Organization Name"),
        "available_territories":  unique_sorted("Sales Territory"),
        "available_countries":    unique_sorted("Country"),
        "available_segments":     unique_sorted("Segment"),
    }


# =============================================================================
# MAIN FUNCTION (importable desde otros scripts)
# =============================================================================

def run(xlsx_path: Path = None, output_dir: Path = None, reference_date: date = None) -> Path:
    """
    Orquesta la generacion del HTML report.

    Args:
        xlsx_path:      Path al xlsx exportado de PBI. Si None, coge el mas reciente.
        output_dir:     Carpeta donde guardar el HTML. Si None, usa DEFAULT_OUTPUT_DIR.
        reference_date: Fecha de corte para YTD. Si None, usa hoy.

    Returns:
        Path del HTML generado.
    """
    log("="*60)
    log("BUCHI Sales Report Generator")
    log("="*60)

    # Resolver paths
    if xlsx_path is None:
        xlsx_path = find_latest_export(DEFAULT_EXPORTS_DIR)
        log(f"Usando export mas reciente: {xlsx_path.name}")
    else:
        xlsx_path = Path(xlsx_path)

    if output_dir is None:
        output_dir = DEFAULT_OUTPUT_DIR
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Cargar y normalizar
    df = load_and_normalize(xlsx_path)

    # 2. Calcular periodos YTD
    log("Calculando periodos YTD...")
    df_p1, df_p2, nombre_p1, nombre_p2 = compute_ytd_periods(df, reference_date)

    if len(df_p1) == 0 and len(df_p2) == 0:
        raise ValueError("Ambos periodos estan vacios. Revisa el rango de fechas del Excel.")

    # 3. Construir comparativa
    log("Construyendo comparativa...")
    comparativa = build_comparativa(df_p1, df_p2, nombre_p1, nombre_p2)
    log(f"  Comparativa: {len(comparativa)} filas")

    # 4. Listas de filtros
    log("Extrayendo listas de filtros...")
    lists = get_available_lists(df_p1, df_p2)

    # Los nombres cortos (solo año) son los que usa generate_sales_comparison_html
    # internamente para buscar columnas "Amount {nombre}" en la comparativa.
    # Los nombres largos solo van a config_info para mostrar en el report.
    import re as _re
    def extract_year(nombre):
        m = _re.search(r"(\d{4})", nombre)
        return m.group(1) if m else nombre

    nombre_p1_short = extract_year(nombre_p1)  # "2025"
    nombre_p2_short = extract_year(nombre_p2)  # "2026"

    # 5. Config info
    config_info = {
        "source_file":    xlsx_path.name,
        "generated_at":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "period_1":       nombre_p1,
        "period_2":       nombre_p2,
        "rows_p1":        len(df_p1),
        "rows_p2":        len(df_p2),
        "reference_date": str(reference_date or date.today()),
    }

    # 6. Detectar si hay SFDC links
    has_sfdc = "SFDC Link" in df_p1.columns or "SFDC Link" in df_p2.columns

    # 7. Generar HTML
    log("Generando HTML...")
    html = generate_sales_comparison_html(
        df1_filtrado=df_p1,
        df2_filtrado=df_p2,
        comparativa=comparativa,
        config_info=config_info,
        nombre_periodo_1=nombre_p1_short,
        nombre_periodo_2=nombre_p2_short,
        has_sfdc_links=has_sfdc,
        **lists,
    )

    # 8. Guardar
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"sales_report_{ts}.html"
    out_path.write_text(html, encoding="utf-8")
    log(f"HTML guardado: {out_path}")
    log("="*60)

    return out_path


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="BUCHI Sales Report Generator")
    parser.add_argument(
        "--file", "-f",
        help="Path al xlsx exportado de PBI (por defecto: el mas reciente en exports/)",
        default=None
    )
    parser.add_argument(
        "--output", "-o",
        help="Carpeta de salida para el HTML (por defecto: reports/)",
        default=None
    )
    parser.add_argument(
        "--date", "-d",
        help="Fecha de corte YTD en formato YYYY-MM-DD (por defecto: hoy)",
        default=None
    )
    args = parser.parse_args()

    ref_date = None
    if args.date:
        try:
            ref_date = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"[ERROR] Formato de fecha invalido: {args.date}. Usa YYYY-MM-DD.")
            sys.exit(1)

    try:
        out = run(
            xlsx_path=Path(args.file) if args.file else None,
            output_dir=Path(args.output) if args.output else None,
            reference_date=ref_date,
        )
        print(f"SUCCESS:{out}")
        sys.exit(0)
    except Exception as e:
        import traceback
        print(f"[ERROR] {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()