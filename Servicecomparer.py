import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Comparador de Ventas", 
    layout="wide",
    page_icon="📊"
)

# CSS personalizado para mejor apariencia
st.markdown("""
<style>
    .step-header {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .success-box {
        background-color: #d4edda;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #28a745;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
        margin: 10px 0;
    }
    .info-box {
        background-color: #d1ecf1;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #17a2b8;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.title("📊 Comparador de Ventas por Periodo")
st.markdown("### Herramienta para analizar y comparar ventas entre dos periodos")

# Información inicial y guía
with st.expander("ℹ️ **CÓMO USAR ESTA APLICACIÓN** - Haz clic aquí para ver la guía", expanded=False):
    st.markdown("""
    ## 🎯 Propósito
    Esta herramienta te permite comparar ventas exportadas desde **Power BI** entre dos periodos diferentes.
    
    ## 📋 Pasos a seguir:
    
    ### **PASO 1:** Preparar tus archivos
    - Exporta los datos de ventas desde Power BI en formato **Excel (.xlsx)** o **CSV (.csv)**
    - Necesitas **DOS archivos**: uno para cada periodo que deseas comparar
    - Asegúrate de que incluyan estas columnas:
      - `Date` - Fecha de la transacción
      - `Business Partner Name` - Nombre del cliente
      - `ItemIdAndName` - Producto o servicio
      - `ProductType` - Tipo de producto
      - `Qty` - Cantidad vendida
      - `EUR` - Importe en euros
      - `SalesRepresentative` - Representante de ventas
      - `Set` y `Productline` - Agrupaciones
    
    ### **PASO 2:** Cargar archivos (sección de abajo)
    - Nombra cada periodo (ej: "Q1 2024", "Enero 2024")
    - Sube el archivo de cada periodo
    
    ### **PASO 3:** Verificar columnas
    - La app detecta automáticamente las columnas estándar
    - Si tus columnas tienen nombres diferentes, ajústalas manualmente
    
    ### **PASO 4:** Aplicar filtros
    - Selecciona el rango de fechas que deseas analizar
    - Elige qué tipos de productos incluir
    
    ### **PASO 5:** Descargar resultados
    - Revisa el resumen de métricas
    - Descarga el archivo Excel con el análisis completo
    
    ---
    
    ## 📊 ¿Qué obtendrás?
    Un archivo Excel con **5 hojas**:
    1. **Comparativa** - Tabla completa con todas las ventas
    2. **Datos Originales** - Transacciones completas de ambos periodos
    3. **Solo en Periodo 1** - Ventas que no se repitieron (clientes perdidos)
    4. **Solo en Periodo 2** - Nuevas ventas (clientes ganados)
    5. **Comunes en ambos** - Ventas recurrentes (clientes fidelizados)
    
    ## 💡 Ejemplos de uso:
    - Comparar **Q1 2024 vs Q1 2023** → Crecimiento interanual
    - Comparar **Enero vs Febrero** → Evolución mensual
    - Filtrar por representante → Evaluar performance individual
    """)

# Separador visual
st.markdown("---")

# Función para leer el archivo en Excel o CSV
def load_file(file):
    if file is not None:
        try:
            if file.name.endswith(".csv"):
                return pd.read_csv(file, encoding='utf-8')
            else:
                return pd.read_excel(file)
        except Exception as e:
            st.error(f"❌ Error al cargar el archivo: {str(e)}")
            return None
    return None

# =============================================================================
# PASO 1: CARGA DE ARCHIVOS
# =============================================================================
st.markdown("## 📁 PASO 1: Carga de Archivos")
st.markdown("Sube los dos archivos de ventas exportados desde Power BI")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📅 Periodo 1")
    nombre_periodo_1 = st.text_input(
        "Nombre del período 1 (ej: 'Q1 2024', 'Enero 2024')", 
        value="Periodo 1", 
        key="nombre_p1",
        help="Dale un nombre descriptivo para identificarlo fácilmente"
    )
    file1 = st.file_uploader(
        f"Archivo de ventas - {nombre_periodo_1}", 
        type=["xlsx", "csv"], 
        key="file1",
        help="Exporta desde Power BI: Datos → Exportar datos → .xlsx o .csv"
    )
    
with col2:
    st.markdown("### 📅 Periodo 2")
    nombre_periodo_2 = st.text_input(
        "Nombre del período 2 (ej: 'Q1 2023', 'Enero 2023')", 
        value="Periodo 2", 
        key="nombre_p2",
        help="Dale un nombre descriptivo para identificarlo fácilmente"
    )
    file2 = st.file_uploader(
        f"Archivo de ventas - {nombre_periodo_2}", 
        type=["xlsx", "csv"], 
        key="file2",
        help="Exporta desde Power BI: Datos → Exportar datos → .xlsx o .csv"
    )

# Procesamiento y comparación
if file1 and file2:
    st.success("✅ **Archivos cargados correctamente**")
    
    df1 = load_file(file1)
    df2 = load_file(file2)
    
    if df1 is None or df2 is None:
        st.stop()
    
    # =============================================================================
    # VISTA PREVIA DE DATOS
    # =============================================================================
    st.markdown("---")
    st.markdown("### 👀 Vista Previa de los Datos")
    st.info("🔍 Verifica que los datos se hayan cargado correctamente antes de continuar")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**{nombre_periodo_1}** ({len(df1)} registros)")
        st.dataframe(df1.head(5), use_container_width=True)
    with col2:
        st.markdown(f"**{nombre_periodo_2}** ({len(df2)} registros)")
        st.dataframe(df2.head(5), use_container_width=True)
    
    # =============================================================================
    # PASO 2: ASIGNACIÓN DE COLUMNAS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🛠️ PASO 2: Asignación de Columnas")
    st.markdown("La aplicación detecta automáticamente las columnas. **Solo ajusta si es necesario.**")
    
    # Asignar nombres de columnas fijos con posibilidad de personalizar
    default_cols = {
        'Fecha': 'Date',
        'Cliente': 'Business Partner Name',
        'Producto': 'ItemIdAndName',
        'Tipo de producto': 'ProductType',
        'Cantidad': 'Qty',
        'Importe': 'EUR',
        'SalesRepresentative': 'SalesRepresentative',
        'Set': 'Set',
        'Productline': 'Productline'
    }
    
    with st.expander("🔧 Ajustar mapeo de columnas (solo si es necesario)", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"**Columnas de {nombre_periodo_1}**")
            col_fecha_1 = st.selectbox(
                f"📅 Fecha", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Fecha']) if default_cols['Fecha'] in df1.columns else 0,
                key="fecha_1"
            )
            col_cliente_1 = st.selectbox(
                f"👤 Cliente", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Cliente']) if default_cols['Cliente'] in df1.columns else 0,
                key="cliente_1"
            )
            col_producto_1 = st.selectbox(
                f"📦 Producto", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Producto']) if default_cols['Producto'] in df1.columns else 0,
                key="producto_1"
            )
            col_tipo_1 = st.selectbox(
                f"🏷️ Tipo de producto", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Tipo de producto']) if default_cols['Tipo de producto'] in df1.columns else 0,
                key="tipo_1"
            )
            col_cantidad_1 = st.selectbox(
                f"🔢 Cantidad", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Cantidad']) if default_cols['Cantidad'] in df1.columns else 0,
                key="cantidad_1"
            )
            col_precio_1 = st.selectbox(
                f"💰 Importe", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Importe']) if default_cols['Importe'] in df1.columns else 0,
                key="precio_1"
            )
            col_sales_rep_1 = st.selectbox(
                f"👔 SalesRepresentative", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['SalesRepresentative']) if default_cols['SalesRepresentative'] in df1.columns else 0,
                key="sales_1"
            )
            col_set_1 = st.selectbox(
                f"📊 Set", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Set']) if default_cols['Set'] in df1.columns else 0,
                key="set_1"
            )
            col_productline_1 = st.selectbox(
                f"📈 Productline", 
                df1.columns, 
                index=df1.columns.get_loc(default_cols['Productline']) if default_cols['Productline'] in df1.columns else 0,
                key="productline_1"
            )
        
        with col2:
            st.markdown(f"**Columnas de {nombre_periodo_2}**")
            col_fecha_2 = st.selectbox(
                f"📅 Fecha", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Fecha']) if default_cols['Fecha'] in df2.columns else 0,
                key="fecha_2"
            )
            col_cliente_2 = st.selectbox(
                f"👤 Cliente", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Cliente']) if default_cols['Cliente'] in df2.columns else 0,
                key="cliente_2"
            )
            col_producto_2 = st.selectbox(
                f"📦 Producto", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Producto']) if default_cols['Producto'] in df2.columns else 0,
                key="producto_2"
            )
            col_tipo_2 = st.selectbox(
                f"🏷️ Tipo de producto", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Tipo de producto']) if default_cols['Tipo de producto'] in df2.columns else 0,
                key="tipo_2"
            )
            col_cantidad_2 = st.selectbox(
                f"🔢 Cantidad", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Cantidad']) if default_cols['Cantidad'] in df2.columns else 0,
                key="cantidad_2"
            )
            col_precio_2 = st.selectbox(
                f"💰 Importe", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Importe']) if default_cols['Importe'] in df2.columns else 0,
                key="precio_2"
            )
            col_sales_rep_2 = st.selectbox(
                f"👔 SalesRepresentative", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['SalesRepresentative']) if default_cols['SalesRepresentative'] in df2.columns else 0,
                key="sales_2"
            )
            col_set_2 = st.selectbox(
                f"📊 Set", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Set']) if default_cols['Set'] in df2.columns else 0,
                key="set_2"
            )
            col_productline_2 = st.selectbox(
                f"📈 Productline", 
                df2.columns, 
                index=df2.columns.get_loc(default_cols['Productline']) if default_cols['Productline'] in df2.columns else 0,
                key="productline_2"
            )
    else:
        # Usar valores por defecto si no se expande
        col_fecha_1 = default_cols['Fecha'] if default_cols['Fecha'] in df1.columns else df1.columns[0]
        col_cliente_1 = default_cols['Cliente'] if default_cols['Cliente'] in df1.columns else df1.columns[0]
        col_producto_1 = default_cols['Producto'] if default_cols['Producto'] in df1.columns else df1.columns[0]
        col_tipo_1 = default_cols['Tipo de producto'] if default_cols['Tipo de producto'] in df1.columns else df1.columns[0]
        col_cantidad_1 = default_cols['Cantidad'] if default_cols['Cantidad'] in df1.columns else df1.columns[0]
        col_precio_1 = default_cols['Importe'] if default_cols['Importe'] in df1.columns else df1.columns[0]
        col_sales_rep_1 = default_cols['SalesRepresentative'] if default_cols['SalesRepresentative'] in df1.columns else df1.columns[0]
        col_set_1 = default_cols['Set'] if default_cols['Set'] in df1.columns else df1.columns[0]
        col_productline_1 = default_cols['Productline'] if default_cols['Productline'] in df1.columns else df1.columns[0]
        
        col_fecha_2 = default_cols['Fecha'] if default_cols['Fecha'] in df2.columns else df2.columns[0]
        col_cliente_2 = default_cols['Cliente'] if default_cols['Cliente'] in df2.columns else df2.columns[0]
        col_producto_2 = default_cols['Producto'] if default_cols['Producto'] in df2.columns else df2.columns[0]
        col_tipo_2 = default_cols['Tipo de producto'] if default_cols['Tipo de producto'] in df2.columns else df2.columns[0]
        col_cantidad_2 = default_cols['Cantidad'] if default_cols['Cantidad'] in df2.columns else df2.columns[0]
        col_precio_2 = default_cols['Importe'] if default_cols['Importe'] in df2.columns else df2.columns[0]
        col_sales_rep_2 = default_cols['SalesRepresentative'] if default_cols['SalesRepresentative'] in df2.columns else df2.columns[0]
        col_set_2 = default_cols['Set'] if default_cols['Set'] in df2.columns else df2.columns[0]
        col_productline_2 = default_cols['Productline'] if default_cols['Productline'] in df2.columns else df2.columns[0]
    
    # =============================================================================
    # PASO 3: FILTROS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🎯 PASO 3: Aplicar Filtros")
    
    # Filtro por fecha
    st.markdown("### 📅 Filtro por Rango de Fechas")
    st.info("💡 **Tip:** Puedes comparar el mismo mes de diferentes años, o periodos personalizados")
    
    # Formatear fechas para eliminar la hora (00:00:00)
    df1[col_fecha_1] = pd.to_datetime(df1[col_fecha_1], errors='coerce').dt.date
    df2[col_fecha_2] = pd.to_datetime(df2[col_fecha_2], errors='coerce').dt.date
    
    min_date_1, max_date_1 = pd.to_datetime(df1[col_fecha_1]).min(), pd.to_datetime(df1[col_fecha_1]).max()
    min_date_2, max_date_2 = pd.to_datetime(df2[col_fecha_2]).min(), pd.to_datetime(df2[col_fecha_2]).max()
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Rango disponible en {nombre_periodo_1}:**")
        st.caption(f"Desde {min_date_1.date()} hasta {max_date_1.date()}")
        start_date_1, end_date_1 = st.date_input(
            f"Selecciona rango para {nombre_periodo_1}", 
            [min_date_1, max_date_1],
            key="date_range_1"
        )
    
    with col2:
        st.markdown(f"**Rango disponible en {nombre_periodo_2}:**")
        st.caption(f"Desde {min_date_2.date()} hasta {max_date_2.date()}")
        start_date_2, end_date_2 = st.date_input(
            f"Selecciona rango para {nombre_periodo_2}", 
            [min_date_2, max_date_2],
            key="date_range_2"
        )
    
    df1 = df1[(pd.to_datetime(df1[col_fecha_1]) >= pd.to_datetime(start_date_1)) & (pd.to_datetime(df1[col_fecha_1]) <= pd.to_datetime(end_date_1))]
    df2 = df2[(pd.to_datetime(df2[col_fecha_2]) >= pd.to_datetime(start_date_2)) & (pd.to_datetime(df2[col_fecha_2]) <= pd.to_datetime(end_date_2))]
    
    # Filtro por tipo de producto
    st.markdown("### 🏷️ Filtro por Tipo de Producto")
    tipos_disponibles = sorted(set(df1[col_tipo_1].dropna().unique()) | set(df2[col_tipo_2].dropna().unique()))
    
    st.info(f"💡 **Tipos disponibles:** {', '.join(tipos_disponibles)}")
    tipos_seleccionados = st.multiselect(
        "Selecciona los tipos de producto a incluir en el análisis", 
        tipos_disponibles, 
        default=tipos_disponibles,
        help="Puedes excluir tipos como 'Alquiler', 'Muestra', etc."
    )
    
    if not tipos_seleccionados:
        st.warning("⚠️ Debes seleccionar al menos un tipo de producto")
        st.stop()
    
    df1_filtrado = df1[df1[col_tipo_1].isin(tipos_seleccionados)].copy()
    df2_filtrado = df2[df2[col_tipo_2].isin(tipos_seleccionados)].copy()
    
    st.success(f"✅ Filtros aplicados: {len(df1_filtrado)} registros en {nombre_periodo_1}, {len(df2_filtrado)} registros en {nombre_periodo_2}")
    
    # =============================================================================
    # PROCESAMIENTO DE DATOS
    # =============================================================================
    
    # Convertir cantidad e importe a numérico
    df1_filtrado[col_cantidad_1] = pd.to_numeric(df1_filtrado[col_cantidad_1], errors='coerce')
    df1_filtrado[col_precio_1] = pd.to_numeric(df1_filtrado[col_precio_1], errors='coerce')
    df2_filtrado[col_cantidad_2] = pd.to_numeric(df2_filtrado[col_cantidad_2], errors='coerce')
    df2_filtrado[col_precio_2] = pd.to_numeric(df2_filtrado[col_precio_2], errors='coerce')
    
    # El importe ya es el total
    df1_filtrado["Importe"] = df1_filtrado[col_precio_1]
    df2_filtrado["Importe"] = df2_filtrado[col_precio_2]
    
    # Agrupar por cliente, producto, sales representative, set y productline
    with st.spinner("🔄 Procesando datos y generando comparativa..."):
        grouped_1 = df1_filtrado.groupby([col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]).agg({
            col_cantidad_1: "sum",
            "Importe": "sum"
        }).rename(columns={col_cantidad_1: f"Cantidad {nombre_periodo_1}", "Importe": f"Importe {nombre_periodo_1}"})
        
        grouped_2 = df2_filtrado.groupby([col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]).agg({
            col_cantidad_2: "sum",
            "Importe": "sum"
        }).rename(columns={col_cantidad_2: f"Cantidad {nombre_periodo_2}", "Importe": f"Importe {nombre_periodo_2}"})
        
        comparativa = pd.merge(grouped_1, grouped_2, how="outer", left_index=True, right_index=True).fillna(0)
        comparativa["Diferencia Cantidad"] = comparativa[f"Cantidad {nombre_periodo_2}"] - comparativa[f"Cantidad {nombre_periodo_1}"]
        comparativa["Diferencia Importe"] = comparativa[f"Importe {nombre_periodo_2}"] - comparativa[f"Importe {nombre_periodo_1}"]
    
    # =============================================================================
    # PASO 4: RESULTADOS Y MÉTRICAS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📊 PASO 4: Resultados del Análisis")
    
    # Mostrar resumen de la comparativa
    st.markdown("### 💰 Resumen Financiero")
    
    total_periodo_1 = comparativa[f"Importe {nombre_periodo_1}"].sum()
    total_periodo_2 = comparativa[f"Importe {nombre_periodo_2}"].sum()
    diferencia_total = total_periodo_2 - total_periodo_1
    porcentaje_cambio = (diferencia_total / total_periodo_1 * 100) if total_periodo_1 != 0 else 0
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            f"💵 Total {nombre_periodo_1}", 
            f"€{total_periodo_1:,.2f}",
            help="Suma total de ventas del primer periodo"
        )
    with col2:
        st.metric(
            f"💵 Total {nombre_periodo_2}", 
            f"€{total_periodo_2:,.2f}",
            help="Suma total de ventas del segundo periodo"
        )
    with col3:
        st.metric(
            "📈 Diferencia Total", 
            f"€{diferencia_total:,.2f}",
            delta=f"{porcentaje_cambio:.1f}%",
            help="Variación entre ambos periodos"
        )
    
    st.markdown("### 📋 Análisis de Registros")
    
    registros_comunes = len(comparativa[(comparativa[f"Importe {nombre_periodo_1}"] > 0) & (comparativa[f"Importe {nombre_periodo_2}"] > 0)])
    registros_solo_p1 = len(comparativa[(comparativa[f"Importe {nombre_periodo_1}"] > 0) & (comparativa[f"Importe {nombre_periodo_2}"] == 0)])
    registros_solo_p2 = len(comparativa[(comparativa[f"Importe {nombre_periodo_1}"] == 0) & (comparativa[f"Importe {nombre_periodo_2}"] > 0)])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            "✅ Registros Comunes", 
            registros_comunes,
            help="Ventas que aparecen en ambos periodos (clientes recurrentes)"
        )
    with col2:
        st.metric(
            f"🔴 Solo en {nombre_periodo_1}", 
            registros_solo_p1,
            help="Ventas que NO se repitieron (clientes perdidos)"
        )
    with col3:
        st.metric(
            f"🟢 Solo en {nombre_periodo_2}", 
            registros_solo_p2,
            help="Nuevas ventas (clientes ganados)"
        )
    with col4:
        st.metric(
            "📊 Total Registros", 
            len(comparativa),
            help="Total de combinaciones únicas cliente-producto"
        )
    
    # KPIs adicionales
    st.markdown("### 📈 KPIs de Retención y Captación")
    
    if registros_comunes + registros_solo_p1 > 0:
        tasa_retencion = (registros_comunes / (registros_comunes + registros_solo_p1)) * 100
    else:
        tasa_retencion = 0
    
    if registros_comunes + registros_solo_p2 > 0:
        tasa_captacion = (registros_solo_p2 / (registros_comunes + registros_solo_p2)) * 100
    else:
        tasa_captacion = 0
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(
            "🎯 Tasa de Retención", 
            f"{tasa_retencion:.1f}%",
            help="% de ventas que se mantuvieron del P1 al P2. Objetivo: > 80%"
        )
    with col2:
        st.metric(
            "🚀 Tasa de Captación", 
            f"{tasa_captacion:.1f}%",
            help="% de nuevas ventas en P2. Indica crecimiento"
        )
    
    # Vista previa de la comparativa
    st.markdown("### 👀 Vista Previa de la Comparativa (primeras 20 filas)")
    st.dataframe(
        comparativa.reset_index().head(20), 
        use_container_width=True,
        height=400
    )
    
    # =============================================================================
    # PASO 5: DESCARGA
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📥 PASO 5: Descargar Resultados")
    
    st.info("🔧 Generando archivo Excel con todas las hojas de análisis...")
    
    # Preparar archivo de descarga
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_path = tmp_file.name
    
    # Crear Excel
    with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
        # Hoja principal de comparativa
        comparativa_out = comparativa.reset_index()
        comparativa_out.to_excel(writer, index=False, sheet_name='Comparativa')
        
        # Datos originales completos
        df1_filtrado_copy = df1_filtrado.copy()
        df1_filtrado_copy['Origen'] = nombre_periodo_1
        df2_filtrado_copy = df2_filtrado.copy()
        df2_filtrado_copy['Origen'] = nombre_periodo_2
        
        datos_originales = pd.concat([df1_filtrado_copy, df2_filtrado_copy], ignore_index=True)
        datos_originales.to_excel(writer, index=False, sheet_name='Datos Originales')
        
        # Servicios únicos en cada periodo y comunes
        keys_1 = set(df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1))
        keys_2 = set(df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1))
        
        unicos_1 = keys_1 - keys_2
        unicos_2 = keys_2 - keys_1
        comunes = keys_1 & keys_2
        
        # DataFrames de únicos y comunes
        if unicos_1:
            df_unicos_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(unicos_1)]
            df_unicos_1.to_excel(writer, index=False, sheet_name=f'Solo en {nombre_periodo_1}'[:31])
        
        if unicos_2:
            df_unicos_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(unicos_2)]
            df_unicos_2.to_excel(writer, index=False, sheet_name=f'Solo en {nombre_periodo_2}'[:31])
        
        if comunes:
            df_comunes_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(comunes)].copy()
            df_comunes_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(comunes)].copy()
            
            df_comunes_1['Periodo'] = nombre_periodo_1
            df_comunes_2['Periodo'] = nombre_periodo_2
            
            df_comunes_combinado = pd.concat([df_comunes_1, df_comunes_2], ignore_index=True)
            df_comunes_combinado.to_excel(writer, index=False, sheet_name='Comunes en ambos')
    
    # Leer archivo final
    with open(tmp_path, 'rb') as f:
        output = BytesIO(f.read())
    
    # Limpiar
    os.unlink(tmp_path)
    
    # Nombre del archivo con fecha
    fecha_actual = datetime.now().strftime("%Y%m%d")
    nombre_archivo = f"comparativa_{nombre_periodo_1}_vs_{nombre_periodo_2}_{fecha_actual}.xlsx"
    
    st.success("✅ **Archivo Excel generado correctamente**")
    
    # Información sobre el contenido
    with st.expander("📋 ¿Qué contiene el archivo Excel?", expanded=True):
        st.markdown(f"""
        El archivo contiene **5 hojas** con diferentes análisis:
        
        1. **📊 Comparativa** - Tabla completa con todas las combinaciones cliente-producto y sus diferencias
        2. **📄 Datos Originales** - Todas las transacciones originales de ambos periodos (con columna 'Origen')
        3. **🔴 Solo en {nombre_periodo_1}** - {registros_solo_p1} registros que NO aparecen en {nombre_periodo_2} (clientes perdidos)
        4. **🟢 Solo en {nombre_periodo_2}** - {registros_solo_p2} registros nuevos (clientes ganados)
        5. **✅ Comunes en ambos** - {registros_comunes} registros que aparecen en ambos periodos (clientes recurrentes)
        
        ### 💡 Usos recomendados:
        - **Hoja 3 (Solo en P1):** Identifica clientes a recuperar o servicios no renovados
        - **Hoja 4 (Solo en P2):** Celebra nuevas captaciones y expansión
        - **Hoja 5 (Comunes):** Analiza crecimiento en la base fidelizada
        """)
    
    # Botón de descarga destacado
    st.download_button(
        label="📥 DESCARGAR ARCHIVO EXCEL COMPLETO",
        data=output.getvalue(),
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.balloons()
    st.success(f"🎉 **¡Análisis completado!** Archivo listo para descargar: `{nombre_archivo}`")

else:
    # Mensaje cuando no hay archivos cargados
    st.info("👆 **Comienza subiendo los dos archivos de ventas en la sección de arriba**")
    
    st.markdown("""
    ### 📌 Recordatorio:
    1. Exporta tus datos desde **Power BI** en formato Excel (.xlsx) o CSV (.csv)
    2. Necesitas **DOS archivos**: uno para cada periodo
    3. Asegúrate de que incluyan las columnas estándar de Power BI
    
    ¿Necesitas ayuda? Abre la guía completa en la parte superior de la página.
    """)
