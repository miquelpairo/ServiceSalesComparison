"""
Sales Comparison HTML Report Generator
Generates standalone HTML reports with interactive filters and charts using Jinja2
"""

import json
import os
from datetime import datetime
from typing import Dict, List
import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
from core.report_utils import load_buchi_css, get_sidebar_styles, get_common_report_styles
from app_config.plotting import BUCHI_COLORS

# =============================================================================
# CDN EMBEDDING - Makes HTML work offline / from file:// (e.g. iPhone Safari)
# =============================================================================

# Cache folder: same directory as this script
_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.cdn_cache')

# CDN resources to embed. Order matters for JS (jquery -> popper -> bootstrap).
_CDN_RESOURCES = [
    {
        "tag":   "bootstrap_css",
        "url":   "https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css",
        "file":  "bootstrap.4.5.2.min.css",
        "type":  "css",
    },
    {
        "tag":   "jquery_js",
        "url":   "https://code.jquery.com/jquery-3.5.1.min.js",
        "file":  "jquery.3.5.1.min.js",
        "type":  "js",
    },
    {
        "tag":   "popper_js",
        "url":   "https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js",
        "file":  "popper.1.16.1.min.js",
        "type":  "js",
    },
    {
        "tag":   "bootstrap_js",
        "url":   "https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js",
        "file":  "bootstrap.4.5.2.min.js",
        "type":  "js",
    },
    {
        "tag":   "plotly_js",
        "url":   "https://cdn.plot.ly/plotly-2.26.0.min.js",
        "file":  "plotly.2.26.0.min.js",
        "type":  "js",
    },
]


def _fetch_cdn_resource(url: str, cache_path: str) -> str:
    """Download a CDN resource and cache it locally. Returns the content."""
    import urllib.request
    print(f"  [CDN] Downloading: {url}", flush=True)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        content = r.read().decode("utf-8", errors="replace")
    os.makedirs(os.path.dirname(cache_path), exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  [CDN] Cached: {os.path.basename(cache_path)} ({len(content)//1024} KB)", flush=True)
    return content


def _get_cdn_content(resource: dict) -> str:
    """Return CDN resource content, using cache if available."""
    cache_path = os.path.join(_CACHE_DIR, resource["file"])
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as f:
            return f.read()
    return _fetch_cdn_resource(resource["url"], cache_path)


def _embed_cdn_resources(html: str) -> str:
    """
    Replace CDN <link> and <script src=...> tags with inline content.
    Falls back to original CDN tags if download fails.

    IMPORTANT: uses lambda replacements in re.sub to avoid re interpreting
    backslashes in minified JS/CSS content (e.g. \\D, \\S regex chars).
    """
    import re

    for res in _CDN_RESOURCES:
        try:
            content = _get_cdn_content(res)
        except Exception as e:
            print(f"  [CDN] WARNING: Could not embed {res['file']}: {e}", flush=True)
            continue  # keep original CDN tag — better than nothing

        if res["type"] == "css":
            inline_css = f'<style>{content}</style>'
            lib_name = res["file"].split(".")[0].lower()  # "bootstrap"
            pattern = (
                r'<link\s[^>]*href=["\'][^"\']*' + lib_name + r'[^"\']*\.css["\'][^>]*>'
            )
            html = re.sub(pattern, lambda m: inline_css, html, count=1, flags=re.IGNORECASE)

        else:
            inline_js = f'<script>{content}</script>'
            lib_name = res["file"].split(".")[0].lower()  # "jquery", "popper", "bootstrap", "plotly"

            # Try exact URL match first
            pattern_exact = (
                r'<script\s[^>]*src=["\']' + re.escape(res["url"]) + r'["\'][^>]*>\s*</script>'
            )
            new_html = re.sub(pattern_exact, lambda m: inline_js, html, flags=re.IGNORECASE)
            if new_html != html:
                html = new_html
                continue

            # Fallback: match by library name (handles .min vs non-min, version differences)
            pattern_name = (
                r'<script\s[^>]*src=["\'][^"\']*' + lib_name + r'[^"\']*\.js["\'][^>]*>\s*</script>'
            )
            html = re.sub(pattern_name, lambda m: inline_js, html, count=1, flags=re.IGNORECASE)

    return html


def generate_sales_comparison_html(
    df1_filtrado: pd.DataFrame,
    df2_filtrado: pd.DataFrame,
    comparativa: pd.DataFrame,
    config_info: Dict,
    nombre_periodo_1: str,
    nombre_periodo_2: str,
    available_types: List[str],
    available_sets: List[str],
    available_reps: List[str],
    available_market_orgs: List[str] = None,
    available_territories: List[str] = None,
    available_countries: List[str] = None,
    available_segments: List[str] = None,
    has_sfdc_links: bool = False  # ✅ NUEVO PARÁMETRO
) -> str:
    """
    Generate standalone HTML file with embedded data and interactive filters using Jinja2
    
    Args:
        df1_filtrado: Period 1 filtered data
        df2_filtrado: Period 2 filtered data
        comparativa: Comparison DataFrame with differences
        config_info: Configuration dictionary with analysis parameters
        nombre_periodo_1: Period 1 name
        nombre_periodo_2: Period 2 name
        available_types: List of product types
        available_sets: List of sets
        available_reps: List of sales representatives

        # Handle optional new filter lists
        available_market_orgs = available_market_orgs or []
        available_territories = available_territories or []
        available_countries = available_countries or []
        available_segments = available_segments or []
        
    Returns:
        str: Complete HTML content
    """
    
    # =========================================================================
    # PREPARE DATA FOR EXPORT
    # =========================================================================
    
    # Copy dataframes to avoid modifying originals
    df1_export = df1_filtrado.copy()
    df2_export = df2_filtrado.copy()
    comparativa_export = comparativa.reset_index()
    
    # Convert dates to strings for JSON serialization
    if 'Date' in df1_export.columns:
        df1_export['Date'] = pd.to_datetime(df1_export['Date']).dt.strftime('%Y-%m-%d')
    if 'Date' in df2_export.columns:
        df2_export['Date'] = pd.to_datetime(df2_export['Date']).dt.strftime('%Y-%m-%d')
    
    # Add period identifier to distinguish in combined tables
    df1_export['_Period'] = nombre_periodo_1
    df2_export['_Period'] = nombre_periodo_2
    
    # Handle NaN values before JSON conversion
    df1_export = df1_export.fillna('')
    df2_export = df2_export.fillna('')
    comparativa_export = comparativa_export.fillna(0)
    
    # Convert DataFrames to JSON (records format for easy JS consumption)
    df1_json = df1_export.to_json(orient='records')
    df2_json = df2_export.to_json(orient='records')
    comparativa_json = comparativa_export.to_json(orient='records')
    
    # =========================================================================
    # CALCULATE SUMMARY METRICS
    # =========================================================================
    
    total_p1 = float(df1_filtrado['EUR'].sum())
    total_p2 = float(df2_filtrado['EUR'].sum())
    diferencia_total = total_p2 - total_p1
    porcentaje_cambio = (diferencia_total / total_p1 * 100) if total_p1 != 0 else 0
    
    # Calculate retention metrics
    registros_comunes = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] > 0) & 
                                        (comparativa[f"Amount {nombre_periodo_2}"] > 0)])
    registros_solo_p1 = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] > 0) & 
                                        (comparativa[f"Amount {nombre_periodo_2}"] == 0)])
    registros_solo_p2 = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] == 0) & 
                                        (comparativa[f"Amount {nombre_periodo_2}"] > 0)])
    
    if registros_comunes + registros_solo_p1 > 0:
        tasa_retencion = (registros_comunes / (registros_comunes + registros_solo_p1)) * 100
    else:
        tasa_retencion = 0
    
    if registros_comunes + registros_solo_p2 > 0:
        tasa_captacion = (registros_solo_p2 / (registros_comunes + registros_solo_p2)) * 100
    else:
        tasa_captacion = 0
    
    metrics = {
        'total_p1': total_p1,
        'total_p2': total_p2,
        'diferencia_total': diferencia_total,
        'porcentaje_cambio': porcentaje_cambio,
        'total_records': len(comparativa),
        'registros_comunes': registros_comunes,
        'registros_solo_p1': registros_solo_p1,
        'registros_solo_p2': registros_solo_p2,
        'tasa_retencion': tasa_retencion,
        'tasa_captacion': tasa_captacion
    }
    
    # =========================================================================
    # QUICK FILTER KEYWORDS
    # =========================================================================
    
    quick_filter_keywords = ['CARE', 'Exact', 'Start', 'Circle', 'Maintain', 'IQ', 'OQ', 'Install', 'Academy', 'Plus']
    
    # =========================================================================
    # LOAD BUCHI STYLES
    # =========================================================================
    
    buchi_css = load_buchi_css()
    sidebar_styles = get_sidebar_styles()
    common_styles = get_common_report_styles()
    
    # =========================================================================
    # TIMESTAMP
    # =========================================================================
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    timestamp_short = datetime.now().strftime("%Y-%m-%d")
    year = datetime.now().year
    
    # =========================================================================
    # SETUP JINJA2 ENVIRONMENT
    # =========================================================================
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(script_dir, 'templates')
    
    # Check if templates directory exists, if not create it
    if not os.path.exists(templates_dir):
        os.makedirs(templates_dir)
        raise FileNotFoundError(
            f"Templates directory created at: {templates_dir}\n"
            f"Please place 'sales_comparison_report.jinja2' in this directory."
        )
    
    # Check if template file exists
    template_path = os.path.join(templates_dir, 'sales_comparison_report.jinja2')
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Template not found at: {template_path}\n"
            f"Expected location: {template_path}\n"
            f"Please create the template file in the templates folder."
        )
    
    # Create Jinja2 environment
    env = Environment(
        loader=FileSystemLoader(templates_dir),
        autoescape=select_autoescape(['html', 'xml'])
    )
    
    # Load template
    template = env.get_template('sales_comparison_report.jinja2')
    
    # =========================================================================
    # RENDER TEMPLATE WITH DATA
    # =========================================================================
    
    html_content = template.render(
        # Data
        df1_json=df1_json,
        df2_json=df2_json,
        comparativa_json=comparativa_json,
        
        # Lists for filters
        available_types=available_types,
        available_sets=available_sets,
        available_reps=available_reps,
        available_market_orgs=available_market_orgs,
        available_territories=available_territories,
        available_countries=available_countries,
        available_segments=available_segments,
        quick_filter_keywords=quick_filter_keywords,
        
        # Period names
        nombre_p1=nombre_periodo_1,
        nombre_p2=nombre_periodo_2,
        
        # Metrics
        metrics=metrics,
        
        # Config info
        config_info=config_info,
        
        # Styles
        buchi_css=buchi_css,
        sidebar_styles=sidebar_styles,
        common_styles=common_styles,
        colors=BUCHI_COLORS,
        
        # Timestamps
        timestamp=timestamp,
        timestamp_short=timestamp_short,
        year=year,
        
        # ✅ SFDC Links flag
        has_sfdc_links=has_sfdc_links
    )
    
    # =========================================================================
    # EMBED CDN RESOURCES (makes HTML work from file:// on iPhone/Safari)
    # =========================================================================
    print("[CDN] Embedding libraries for offline use...", flush=True)
    html_content = _embed_cdn_resources(html_content)

    return html_content