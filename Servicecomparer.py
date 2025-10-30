import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO

st.set_page_config(page_title="Comparador de Ventas", layout="wide")

st.title("📊 Comparador de Ventas por Periodo")
st.markdown("Sube dos archivos de ventas para comparar productos, clientes o cantidades entre dos periodos.")

# Función para leer el archivo en Excel o CSV
def load_file(file):
   if file is not None:
       if file.name.endswith(".csv"):
           return pd.read_csv(file)
       else:
           return pd.read_excel(file)
   return None

# Carga de archivos y nombres de períodos
col1, col2 = st.columns(2)

with col1:
   st.subheader("Período 1")
   nombre_periodo_1 = st.text_input("Nombre del período 1", value="Periodo 1", key="nombre_p1")
   file1 = st.file_uploader(f"Archivo de ventas - {nombre_periodo_1}", type=["xlsx", "csv"], key="file1")
with col2:
   st.subheader("Período 2")
   nombre_periodo_2 = st.text_input("Nombre del período 2", value="Periodo 2", key="nombre_p2")
   file2 = st.file_uploader(f"Archivo de ventas - {nombre_periodo_2}", type=["xlsx", "csv"], key="file2")

# Procesamiento y comparación
if file1 and file2:
   df1 = load_file(file1)
   df2 = load_file(file2)

   st.subheader("Vista previa")
   col1, col2 = st.columns(2)
   with col1:
       st.write(f"**{nombre_periodo_1}**")
       st.dataframe(df1.head(), use_container_width=True)
   with col2:
       st.write(f"**{nombre_periodo_2}**")
       st.dataframe(df2.head(), use_container_width=True)

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

   st.subheader("🛠️ Asignación de columnas")
   
   col1, col2 = st.columns(2)
   
   with col1:
       st.write(f"**{nombre_periodo_1}**")
       col_fecha_1 = st.selectbox(f"Columna de Fecha ({nombre_periodo_1})", df1.columns, 
                                  index=df1.columns.get_loc(default_cols['Fecha']) if default_cols['Fecha'] in df1.columns else 0)
       col_cliente_1 = st.selectbox(f"Columna de Cliente ({nombre_periodo_1})", df1.columns, 
                                    index=df1.columns.get_loc(default_cols['Cliente']) if default_cols['Cliente'] in df1.columns else 0)
       col_producto_1 = st.selectbox(f"Columna de Producto ({nombre_periodo_1})", df1.columns, 
                                     index=df1.columns.get_loc(default_cols['Producto']) if default_cols['Producto'] in df1.columns else 0)
       col_tipo_1 = st.selectbox(f"Columna de Tipo de producto ({nombre_periodo_1})", df1.columns, 
                                 index=df1.columns.get_loc(default_cols['Tipo de producto']) if default_cols['Tipo de producto'] in df1.columns else 0)
       col_cantidad_1 = st.selectbox(f"Columna de Cantidad ({nombre_periodo_1})", df1.columns, 
                                     index=df1.columns.get_loc(default_cols['Cantidad']) if default_cols['Cantidad'] in df1.columns else 0)
       col_precio_1 = st.selectbox(f"Columna de Importe ({nombre_periodo_1})", df1.columns, 
                                   index=df1.columns.get_loc(default_cols['Importe']) if default_cols['Importe'] in df1.columns else 0)
       col_sales_rep_1 = st.selectbox(f"Columna de SalesRepresentative ({nombre_periodo_1})", df1.columns, 
                                      index=df1.columns.get_loc(default_cols['SalesRepresentative']) if default_cols['SalesRepresentative'] in df1.columns else 0)
       col_set_1 = st.selectbox(f"Columna de Set ({nombre_periodo_1})", df1.columns, 
                                index=df1.columns.get_loc(default_cols['Set']) if default_cols['Set'] in df1.columns else 0)
       col_productline_1 = st.selectbox(f"Columna de Productline ({nombre_periodo_1})", df1.columns, 
                                        index=df1.columns.get_loc(default_cols['Productline']) if default_cols['Productline'] in df1.columns else 0)

   with col2:
       st.write(f"**{nombre_periodo_2}**")
       col_fecha_2 = st.selectbox(f"Columna de Fecha ({nombre_periodo_2})", df2.columns, 
                                  index=df2.columns.get_loc(default_cols['Fecha']) if default_cols['Fecha'] in df2.columns else 0)
       col_cliente_2 = st.selectbox(f"Columna de Cliente ({nombre_periodo_2})", df2.columns, 
                                    index=df2.columns.get_loc(default_cols['Cliente']) if default_cols['Cliente'] in df2.columns else 0)
       col_producto_2 = st.selectbox(f"Columna de Producto ({nombre_periodo_2})", df2.columns, 
                                     index=df2.columns.get_loc(default_cols['Producto']) if default_cols['Producto'] in df2.columns else 0)
       col_tipo_2 = st.selectbox(f"Columna de Tipo de producto ({nombre_periodo_2})", df2.columns, 
                                 index=df2.columns.get_loc(default_cols['Tipo de producto']) if default_cols['Tipo de producto'] in df2.columns else 0)
       col_cantidad_2 = st.selectbox(f"Columna de Cantidad ({nombre_periodo_2})", df2.columns, 
                                     index=df2.columns.get_loc(default_cols['Cantidad']) if default_cols['Cantidad'] in df2.columns else 0)
       col_precio_2 = st.selectbox(f"Columna de Importe ({nombre_periodo_2})", df2.columns, 
                                   index=df2.columns.get_loc(default_cols['Importe']) if default_cols['Importe'] in df2.columns else 0)
       col_sales_rep_2 = st.selectbox(f"Columna de SalesRepresentative ({nombre_periodo_2})", df2.columns, 
                                      index=df2.columns.get_loc(default_cols['SalesRepresentative']) if default_cols['SalesRepresentative'] in df2.columns else 0)
       col_set_2 = st.selectbox(f"Columna de Set ({nombre_periodo_2})", df2.columns, 
                                index=df2.columns.get_loc(default_cols['Set']) if default_cols['Set'] in df2.columns else 0)
       col_productline_2 = st.selectbox(f"Columna de Productline ({nombre_periodo_2})", df2.columns, 
                                        index=df2.columns.get_loc(default_cols['Productline']) if default_cols['Productline'] in df2.columns else 0)

   # Mostrar columnas seleccionadas
   st.markdown(f"**{nombre_periodo_1}:** {col_cliente_1}, {col_producto_1}, {col_tipo_1}, {col_cantidad_1}, {col_precio_1}, {col_fecha_1}, {col_sales_rep_1}")
   st.markdown(f"**{nombre_periodo_2}:** {col_cliente_2}, {col_producto_2}, {col_tipo_2}, {col_cantidad_2}, {col_precio_2}, {col_fecha_2}, {col_sales_rep_2}")

   # Filtro por fecha
   st.subheader("🗓️ Filtro por Fecha")
   # Formatear fechas para eliminar la hora (00:00:00)
   df1[col_fecha_1] = pd.to_datetime(df1[col_fecha_1], errors='coerce').dt.date
   df2[col_fecha_2] = pd.to_datetime(df2[col_fecha_2], errors='coerce').dt.date

   min_date_1, max_date_1 = pd.to_datetime(df1[col_fecha_1]).min(), pd.to_datetime(df1[col_fecha_1]).max()
   min_date_2, max_date_2 = pd.to_datetime(df2[col_fecha_2]).min(), pd.to_datetime(df2[col_fecha_2]).max()

   start_date_1, end_date_1 = st.date_input(f"Rango de fechas {nombre_periodo_1}", [min_date_1, max_date_1])
   start_date_2, end_date_2 = st.date_input(f"Rango de fechas {nombre_periodo_2}", [min_date_2, max_date_2])

   df1 = df1[(pd.to_datetime(df1[col_fecha_1]) >= pd.to_datetime(start_date_1)) & (pd.to_datetime(df1[col_fecha_1]) <= pd.to_datetime(end_date_1))]
   df2 = df2[(pd.to_datetime(df2[col_fecha_2]) >= pd.to_datetime(start_date_2)) & (pd.to_datetime(df2[col_fecha_2]) <= pd.to_datetime(end_date_2))]

   # Filtro por tipo de producto
   st.subheader("🎯 Filtro por Tipo de Producto")
   tipos_disponibles = sorted(set(df1[col_tipo_1].dropna().unique()) | set(df2[col_tipo_2].dropna().unique()))
   tipos_seleccionados = st.multiselect("Selecciona tipos a incluir", tipos_disponibles, default=tipos_disponibles)

   df1_filtrado = df1[df1[col_tipo_1].isin(tipos_seleccionados)].copy()
   df2_filtrado = df2[df2[col_tipo_2].isin(tipos_seleccionados)].copy()

   # Convertir cantidad e importe a numérico
   df1_filtrado[col_cantidad_1] = pd.to_numeric(df1_filtrado[col_cantidad_1], errors='coerce')
   df1_filtrado[col_precio_1] = pd.to_numeric(df1_filtrado[col_precio_1], errors='coerce')
   df2_filtrado[col_cantidad_2] = pd.to_numeric(df2_filtrado[col_cantidad_2], errors='coerce')
   df2_filtrado[col_precio_2] = pd.to_numeric(df2_filtrado[col_precio_2], errors='coerce')

   # Ya es el importe total
   df1_filtrado["Importe"] = df1_filtrado[col_precio_1]
   df2_filtrado["Importe"] = df2_filtrado[col_precio_2]

   # Agrupar por cliente, producto, sales representative, set y productline (incluyendo campos adicionales)
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

   # Mostrar resumen de la comparativa
   st.subheader("📊 Resumen de la Comparativa")
   
   total_periodo_1 = comparativa[f"Importe {nombre_periodo_1}"].sum()
   total_periodo_2 = comparativa[f"Importe {nombre_periodo_2}"].sum()
   diferencia_total = total_periodo_2 - total_periodo_1
   
   col1, col2, col3, col4, col5 = st.columns(5)
   with col1:
       st.metric(f"Total {nombre_periodo_1}", f"€{total_periodo_1:,.2f}")
   with col2:
       st.metric(f"Total {nombre_periodo_2}", f"€{total_periodo_2:,.2f}")
   with col3:
       st.metric("Diferencia Total", f"€{diferencia_total:,.2f}")
   with col4:
       st.metric("Registros comunes", len(comparativa[(comparativa[f"Importe {nombre_periodo_1}"] > 0) & (comparativa[f"Importe {nombre_periodo_2}"] > 0)]))
   with col5:
       st.metric("Total registros", len(comparativa))

   # Vista previa de la comparativa
   st.subheader("Vista previa de la comparativa")
   st.dataframe(comparativa.reset_index().head(20), use_container_width=True)
   
   st.info("🔧 Preparando archivo Excel con formato...")

   # Preparar archivo de descarga
   with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
       tmp_path = tmp_file.name

   st.write(f"📁 Archivo temporal creado: {tmp_path}")

   # Crear Excel básico primero
   with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
       # Hoja principal de comparativa
       comparativa_out = comparativa.reset_index()
       comparativa_out.to_excel(writer, index=False, sheet_name='Comparativa')
       st.write("✅ Hoja 'Comparativa' creada")

       # Datos originales completos
       df1_filtrado_copy = df1_filtrado.copy()
       df1_filtrado_copy['Origen'] = nombre_periodo_1
       df2_filtrado_copy = df2_filtrado.copy()
       df2_filtrado_copy['Origen'] = nombre_periodo_2
       
       # Unir ambos dataframes para la hoja de datos originales
       datos_originales = pd.concat([df1_filtrado_copy, df2_filtrado_copy], ignore_index=True)
       datos_originales.to_excel(writer, index=False, sheet_name='Datos Originales')
       st.write("✅ Hoja 'Datos Originales' creada")

       # Servicios únicos en cada periodo y comunes
       keys_1 = set(df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1))
       keys_2 = set(df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1))

       unicos_1 = keys_1 - keys_2
       unicos_2 = keys_2 - keys_1
       comunes = keys_1 & keys_2
       
       st.write(f"📊 Análisis: {len(unicos_1)} únicos P1, {len(unicos_2)} únicos P2, {len(comunes)} comunes")

       # DataFrames de únicos y comunes
       if unicos_1:
           df_unicos_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(unicos_1)]
           df_unicos_1.to_excel(writer, index=False, sheet_name=f'Solo en {nombre_periodo_1}')
           st.write(f"✅ Hoja 'Solo en {nombre_periodo_1}' creada con {len(df_unicos_1)} registros")
       
       if unicos_2:
           df_unicos_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(unicos_2)]
           df_unicos_2.to_excel(writer, index=False, sheet_name=f'Solo en {nombre_periodo_2}')
           st.write(f"✅ Hoja 'Solo en {nombre_periodo_2}' creada con {len(df_unicos_2)} registros")
       
       if comunes:
           df_comunes_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(comunes)].copy()
           df_comunes_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(comunes)].copy()
           
           df_comunes_1['Periodo'] = nombre_periodo_1
           df_comunes_2['Periodo'] = nombre_periodo_2
           
           df_comunes_combinado = pd.concat([df_comunes_1, df_comunes_2], ignore_index=True)
           df_comunes_combinado.to_excel(writer, index=False, sheet_name='Comunes en ambos')
           st.write(f"✅ Hoja 'Comunes en ambos' creada con {len(df_comunes_combinado)} registros")

   st.success("📝 Excel creado correctamente")
   
   # Leer archivo final
   with open(tmp_path, 'rb') as f:
       output = BytesIO(f.read())
   
   # Limpiar
   os.unlink(tmp_path)

   st.download_button(
       label="📥 Descargar comparativa en Excel",
       data=output.getvalue(),
       file_name=f"comparativa_{nombre_periodo_1}_vs_{nombre_periodo_2}.xlsx",
       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
   )

else:
   st.info("Por favor, sube ambos archivos para comenzar el análisis.")