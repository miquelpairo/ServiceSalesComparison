import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Sales Comparison Tool", 
    layout="wide",
    page_icon="📊"
)

# Custom CSS for better appearance
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

# Main title
st.title("📊 Sales Comparison by Period")
st.markdown("### Tool to analyze and compare sales between two periods")

# Initial information and guide
with st.expander("ℹ️ **HOW TO USE THIS APPLICATION** - Click here to see the guide", expanded=False):
    st.markdown("""
    ## 🎯 Purpose
    This tool allows you to compare sales exported from **Power BI** between two different periods.
    
    ## 📋 Steps to follow:
    
    ### **STEP 1:** Prepare your files
    - Export sales data from Power BI in **Excel (.xlsx)** or **CSV (.csv)** format
    - You need **TWO files**: one for each period you want to compare
    - Make sure they include these columns:
      - `Date` - Transaction date
      - `Business Partner Name` - Customer name
      - `ItemIdAndName` - Product or service
      - `ProductType` - Product type
      - `Qty` - Quantity sold
      - `EUR` - Amount in euros
      - `SalesRepresentative` - Sales representative
      - `Set` and `Productline` - Groupings
    
    ### **STEP 2:** Upload files (section below)
    - Name each period (e.g., "Q1 2024", "January 2024")
    - Upload the file for each period
    
    ### **STEP 3:** Verify columns
    - The app automatically detects standard columns
    - If your columns have different names, adjust them manually
    
    ### **STEP 4:** Apply filters
    - Select the date range you want to analyze
    - Choose which product types to include
    
    ### **STEP 5:** Download results
    - Review the metrics summary
    - Download the Excel file with the complete analysis
    
    ---
    
    ## 📊 What will you get?
    An Excel file with **5 sheets**:
    1. **Comparison** - Complete table with all sales
    2. **Original Data** - Complete transactions from both periods
    3. **Only in Period 1** - Sales that didn't repeat (lost customers)
    4. **Only in Period 2** - New sales (gained customers)
    5. **Common in both** - Recurring sales (loyal customers)
    
    ## 💡 Usage examples:
    - Compare **Q1 2024 vs Q1 2023** → Year-over-year growth
    - Compare **January vs February** → Monthly evolution
    - Filter by representative → Evaluate individual performance
    """)

# Visual separator
st.markdown("---")

# Function to read file in Excel or CSV
def load_file(file):
    if file is not None:
        try:
            if file.name.endswith(".csv"):
                return pd.read_csv(file, encoding='utf-8')
            else:
                return pd.read_excel(file)
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            return None
    return None

# =============================================================================
# STEP 1: FILE UPLOAD
# =============================================================================
st.markdown("## 📁 STEP 1: File Upload")
st.markdown("Upload the two sales files exported from Power BI")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📅 Period 1")
    nombre_periodo_1 = st.text_input(
        "Name of period 1 (e.g., 'Q1 2024', 'January 2024')", 
        value="Period 1", 
        key="nombre_p1",
        help="Give it a descriptive name to identify it easily"
    )
    file1 = st.file_uploader(
        f"Sales file - {nombre_periodo_1}", 
        type=["xlsx", "csv"], 
        key="file1",
        help="Export from Power BI: Data → Export data → .xlsx or .csv"
    )
    
with col2:
    st.markdown("### 📅 Period 2")
    nombre_periodo_2 = st.text_input(
        "Name of period 2 (e.g., 'Q1 2023', 'January 2023')", 
        value="Period 2", 
        key="nombre_p2",
        help="Give it a descriptive name to identify it easily"
    )
    file2 = st.file_uploader(
        f"Sales file - {nombre_periodo_2}", 
        type=["xlsx", "csv"], 
        key="file2",
        help="Export from Power BI: Data → Export data → .xlsx or .csv"
    )

# Processing and comparison
if file1 and file2:
    st.success("✅ **Files loaded successfully**")
    
    df1 = load_file(file1)
    df2 = load_file(file2)
    
    if df1 is None or df2 is None:
        st.stop()
    
    # =============================================================================
    # DATA PREVIEW
    # =============================================================================
    st.markdown("---")
    st.markdown("### 👀 Data Preview")
    st.info("🔍 Verify that data has been loaded correctly before continuing")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**{nombre_periodo_1}** ({len(df1)} records)")
        st.dataframe(df1.head(5), use_container_width=True)
    with col2:
        st.markdown(f"**{nombre_periodo_2}** ({len(df2)} records)")
        st.dataframe(df2.head(5), use_container_width=True)
    
    # =============================================================================
    # STEP 2: COLUMN ASSIGNMENT
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🛠️ STEP 2: Column Assignment")
    st.markdown("The application automatically detects columns. **Only adjust if necessary.**")
    
    # Assign fixed column names with customization option
    default_cols = {
        'Date': 'Date',
        'Customer': 'Business Partner Name',
        'Product': 'ItemIdAndName',
        'Product Type': 'ProductType',
        'Quantity': 'Qty',
        'Amount': 'EUR',
        'SalesRepresentative': 'SalesRepresentative',
        'Set': 'Set',
        'Productline': 'Productline'
    }
    
    # Always use default values (without expander to simplify)
    col_fecha_1 = default_cols['Date'] if default_cols['Date'] in df1.columns else df1.columns[0]
    col_cliente_1 = default_cols['Customer'] if default_cols['Customer'] in df1.columns else df1.columns[0]
    col_producto_1 = default_cols['Product'] if default_cols['Product'] in df1.columns else df1.columns[0]
    col_tipo_1 = default_cols['Product Type'] if default_cols['Product Type'] in df1.columns else df1.columns[0]
    col_cantidad_1 = default_cols['Quantity'] if default_cols['Quantity'] in df1.columns else df1.columns[0]
    col_precio_1 = default_cols['Amount'] if default_cols['Amount'] in df1.columns else df1.columns[0]
    col_sales_rep_1 = default_cols['SalesRepresentative'] if default_cols['SalesRepresentative'] in df1.columns else df1.columns[0]
    col_set_1 = default_cols['Set'] if default_cols['Set'] in df1.columns else df1.columns[0]
    col_productline_1 = default_cols['Productline'] if default_cols['Productline'] in df1.columns else df1.columns[0]
    
    col_fecha_2 = default_cols['Date'] if default_cols['Date'] in df2.columns else df2.columns[0]
    col_cliente_2 = default_cols['Customer'] if default_cols['Customer'] in df2.columns else df2.columns[0]
    col_producto_2 = default_cols['Product'] if default_cols['Product'] in df2.columns else df2.columns[0]
    col_tipo_2 = default_cols['Product Type'] if default_cols['Product Type'] in df2.columns else df2.columns[0]
    col_cantidad_2 = default_cols['Quantity'] if default_cols['Quantity'] in df2.columns else df2.columns[0]
    col_precio_2 = default_cols['Amount'] if default_cols['Amount'] in df2.columns else df2.columns[0]
    col_sales_rep_2 = default_cols['SalesRepresentative'] if default_cols['SalesRepresentative'] in df2.columns else df2.columns[0]
    col_set_2 = default_cols['Set'] if default_cols['Set'] in df2.columns else df2.columns[0]
    col_productline_2 = default_cols['Productline'] if default_cols['Productline'] in df2.columns else df2.columns[0]
    
    st.success(f"✅ Columns detected automatically. If your files use different names, contact the administrator.")
    
    with st.expander("🔍 View detected columns", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Columns from {nombre_periodo_1}:**")
            st.write(f"- Date: `{col_fecha_1}`")
            st.write(f"- Customer: `{col_cliente_1}`")
            st.write(f"- Product: `{col_producto_1}`")
            st.write(f"- Type: `{col_tipo_1}`")
            st.write(f"- Quantity: `{col_cantidad_1}`")
            st.write(f"- Amount: `{col_precio_1}`")
        with col2:
            st.markdown(f"**Columns from {nombre_periodo_2}:**")
            st.write(f"- Date: `{col_fecha_2}`")
            st.write(f"- Customer: `{col_cliente_2}`")
            st.write(f"- Product: `{col_producto_2}`")
            st.write(f"- Type: `{col_tipo_2}`")
            st.write(f"- Quantity: `{col_cantidad_2}`")
            st.write(f"- Amount: `{col_precio_2}`")
    
    # =============================================================================
    # STEP 3: FILTERS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🎯 STEP 3: Apply Filters")
    
    # Date filter
    st.markdown("### 📅 Date Range Filter")
    st.info("💡 **Tip:** You can compare the same month from different years, or custom periods")
    
    # Format dates to remove time (00:00:00)
    df1[col_fecha_1] = pd.to_datetime(df1[col_fecha_1], errors='coerce').dt.date
    df2[col_fecha_2] = pd.to_datetime(df2[col_fecha_2], errors='coerce').dt.date
    
    min_date_1, max_date_1 = pd.to_datetime(df1[col_fecha_1]).min(), pd.to_datetime(df1[col_fecha_1]).max()
    min_date_2, max_date_2 = pd.to_datetime(df2[col_fecha_2]).min(), pd.to_datetime(df2[col_fecha_2]).max()
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Available range in {nombre_periodo_1}:**")
        st.caption(f"From {min_date_1.date()} to {max_date_1.date()}")
        start_date_1, end_date_1 = st.date_input(
            f"Select range for {nombre_periodo_1}", 
            [min_date_1, max_date_1],
            key="date_range_1"
        )
    
    with col2:
        st.markdown(f"**Available range in {nombre_periodo_2}:**")
        st.caption(f"From {min_date_2.date()} to {max_date_2.date()}")
        start_date_2, end_date_2 = st.date_input(
            f"Select range for {nombre_periodo_2}", 
            [min_date_2, max_date_2],
            key="date_range_2"
        )
    
    df1 = df1[(pd.to_datetime(df1[col_fecha_1]) >= pd.to_datetime(start_date_1)) & (pd.to_datetime(df1[col_fecha_1]) <= pd.to_datetime(end_date_1))]
    df2 = df2[(pd.to_datetime(df2[col_fecha_2]) >= pd.to_datetime(start_date_2)) & (pd.to_datetime(df2[col_fecha_2]) <= pd.to_datetime(end_date_2))]
    
    # Product type filter
    st.markdown("### 🏷️ Product Type Filter")
    tipos_disponibles = sorted(set(df1[col_tipo_1].dropna().unique()) | set(df2[col_tipo_2].dropna().unique()))
    
    st.info(f"💡 **Available types:** {', '.join(tipos_disponibles)}")
    tipos_seleccionados = st.multiselect(
        "Select product types to include in the analysis", 
        tipos_disponibles, 
        default=tipos_disponibles,
        help="You can exclude types like 'Rental', 'Sample', etc."
    )
    
    if not tipos_seleccionados:
        st.warning("⚠️ You must select at least one product type")
        st.stop()
    
    df1_filtrado = df1[df1[col_tipo_1].isin(tipos_seleccionados)].copy()
    df2_filtrado = df2[df2[col_tipo_2].isin(tipos_seleccionados)].copy()
    
    st.success(f"✅ Filters applied: {len(df1_filtrado)} records in {nombre_periodo_1}, {len(df2_filtrado)} records in {nombre_periodo_2}")
    
    # =============================================================================
    # DATA PROCESSING
    # =============================================================================
    
    # Convert quantity and amount to numeric
    df1_filtrado[col_cantidad_1] = pd.to_numeric(df1_filtrado[col_cantidad_1], errors='coerce')
    df1_filtrado[col_precio_1] = pd.to_numeric(df1_filtrado[col_precio_1], errors='coerce')
    df2_filtrado[col_cantidad_2] = pd.to_numeric(df2_filtrado[col_cantidad_2], errors='coerce')
    df2_filtrado[col_precio_2] = pd.to_numeric(df2_filtrado[col_precio_2], errors='coerce')
    
    # Amount is already the total
    df1_filtrado["Amount"] = df1_filtrado[col_precio_1]
    df2_filtrado["Amount"] = df2_filtrado[col_precio_2]
    
    # Group by customer, product, sales representative, set and productline
    with st.spinner("🔄 Processing data and generating comparison..."):
        grouped_1 = df1_filtrado.groupby([col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]).agg({
            col_cantidad_1: "sum",
            "Amount": "sum"
        }).rename(columns={col_cantidad_1: f"Quantity {nombre_periodo_1}", "Amount": f"Amount {nombre_periodo_1}"})
        
        grouped_2 = df2_filtrado.groupby([col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]).agg({
            col_cantidad_2: "sum",
            "Amount": "sum"
        }).rename(columns={col_cantidad_2: f"Quantity {nombre_periodo_2}", "Amount": f"Amount {nombre_periodo_2}"})
        
        comparativa = pd.merge(grouped_1, grouped_2, how="outer", left_index=True, right_index=True).fillna(0)
        comparativa["Quantity Difference"] = comparativa[f"Quantity {nombre_periodo_2}"] - comparativa[f"Quantity {nombre_periodo_1}"]
        comparativa["Amount Difference"] = comparativa[f"Amount {nombre_periodo_2}"] - comparativa[f"Amount {nombre_periodo_1}"]
    
    # =============================================================================
    # STEP 4: RESULTS AND METRICS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📊 STEP 4: Analysis Results")
    
    # Show comparison summary
    st.markdown("### 💰 Financial Summary")
    
    total_periodo_1 = comparativa[f"Amount {nombre_periodo_1}"].sum()
    total_periodo_2 = comparativa[f"Amount {nombre_periodo_2}"].sum()
    diferencia_total = total_periodo_2 - total_periodo_1
    porcentaje_cambio = (diferencia_total / total_periodo_1 * 100) if total_periodo_1 != 0 else 0
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            f"💵 Total {nombre_periodo_1}", 
            f"€{total_periodo_1:,.2f}",
            help="Total sales sum for the first period"
        )
    with col2:
        st.metric(
            f"💵 Total {nombre_periodo_2}", 
            f"€{total_periodo_2:,.2f}",
            help="Total sales sum for the second period"
        )
    with col3:
        st.metric(
            "📈 Total Difference", 
            f"€{diferencia_total:,.2f}",
            delta=f"{porcentaje_cambio:.1f}%",
            help="Variation between both periods"
        )
    
    st.markdown("### 📋 Record Analysis")
    
    registros_comunes = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] > 0) & (comparativa[f"Amount {nombre_periodo_2}"] > 0)])
    registros_solo_p1 = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] > 0) & (comparativa[f"Amount {nombre_periodo_2}"] == 0)])
    registros_solo_p2 = len(comparativa[(comparativa[f"Amount {nombre_periodo_1}"] == 0) & (comparativa[f"Amount {nombre_periodo_2}"] > 0)])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            "✅ Common Records", 
            registros_comunes,
            help="Sales appearing in both periods (recurring customers)"
        )
    with col2:
        st.metric(
            f"🔴 Only in {nombre_periodo_1}", 
            registros_solo_p1,
            help="Sales that did NOT repeat (lost customers)"
        )
    with col3:
        st.metric(
            f"🟢 Only in {nombre_periodo_2}", 
            registros_solo_p2,
            help="New sales (gained customers)"
        )
    with col4:
        st.metric(
            "📊 Total Records", 
            len(comparativa),
            help="Total unique customer-product combinations"
        )
    
    # Additional KPIs
    st.markdown("### 📈 Retention and Acquisition KPIs")
    
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
            "🎯 Retention Rate", 
            f"{tasa_retencion:.1f}%",
            help="% of sales maintained from P1 to P2. Target: > 80%"
        )
    with col2:
        st.metric(
            "🚀 Acquisition Rate", 
            f"{tasa_captacion:.1f}%",
            help="% of new sales in P2. Indicates growth"
        )
    
    # Preview of comparison
    st.markdown("### 👀 Comparison Preview (first 20 rows)")
    st.dataframe(
        comparativa.reset_index().head(20), 
        use_container_width=True,
        height=400
    )
    
    # =============================================================================
    # STEP 5: DOWNLOAD
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📥 STEP 5: Download Results")
    
    st.info("🔧 Generating Excel file with all analysis sheets...")
    
    # Prepare download file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_path = tmp_file.name
    
    # Create Excel
    with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
        # Main comparison sheet
        comparativa_out = comparativa.reset_index()
        comparativa_out.to_excel(writer, index=False, sheet_name='Comparison')
        
        # Complete original data
        df1_filtrado_copy = df1_filtrado.copy()
        df1_filtrado_copy['Source'] = nombre_periodo_1
        df2_filtrado_copy = df2_filtrado.copy()
        df2_filtrado_copy['Source'] = nombre_periodo_2
        
        datos_originales = pd.concat([df1_filtrado_copy, df2_filtrado_copy], ignore_index=True)
        datos_originales.to_excel(writer, index=False, sheet_name='Original Data')
        
        # Unique services in each period and common ones
        keys_1 = set(df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1))
        keys_2 = set(df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1))
        
        unicos_1 = keys_1 - keys_2
        unicos_2 = keys_2 - keys_1
        comunes = keys_1 & keys_2
        
        # DataFrames of unique and common
        if unicos_1:
            df_unicos_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(unicos_1)]
            df_unicos_1.to_excel(writer, index=False, sheet_name=f'Only in {nombre_periodo_1}'[:31])
        
        if unicos_2:
            df_unicos_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(unicos_2)]
            df_unicos_2.to_excel(writer, index=False, sheet_name=f'Only in {nombre_periodo_2}'[:31])
        
        if comunes:
            df_comunes_1 = df1_filtrado[df1_filtrado[[col_cliente_1, col_producto_1, col_sales_rep_1, col_set_1, col_productline_1, col_tipo_1]].apply(tuple, axis=1).isin(comunes)].copy()
            df_comunes_2 = df2_filtrado[df2_filtrado[[col_cliente_2, col_producto_2, col_sales_rep_2, col_set_2, col_productline_2, col_tipo_2]].apply(tuple, axis=1).isin(comunes)].copy()
            
            df_comunes_1['Period'] = nombre_periodo_1
            df_comunes_2['Period'] = nombre_periodo_2
            
            df_comunes_combinado = pd.concat([df_comunes_1, df_comunes_2], ignore_index=True)
            df_comunes_combinado.to_excel(writer, index=False, sheet_name='Common in both')
    
    # Read final file
    with open(tmp_path, 'rb') as f:
        output = BytesIO(f.read())
    
    # Clean up
    os.unlink(tmp_path)
    
    # Filename with date
    fecha_actual = datetime.now().strftime("%Y%m%d")
    nombre_archivo = f"comparison_{nombre_periodo_1}_vs_{nombre_periodo_2}_{fecha_actual}.xlsx"
    
    st.success("✅ **Excel file generated successfully**")
    
    # Information about content
    with st.expander("📋 What does the Excel file contain?", expanded=True):
        st.markdown(f"""
        The file contains **5 sheets** with different analyses:
        
        1. **📊 Comparison** - Complete table with all customer-product combinations and their differences
        2. **📄 Original Data** - All original transactions from both periods (with 'Source' column)
        3. **🔴 Only in {nombre_periodo_1}** - {registros_solo_p1} records that do NOT appear in {nombre_periodo_2} (lost customers)
        4. **🟢 Only in {nombre_periodo_2}** - {registros_solo_p2} new records (gained customers)
        5. **✅ Common in both** - {registros_comunes} records appearing in both periods (loyal customers)
        
        ### 💡 Recommended uses:
        - **Sheet 3 (Only in P1):** Identify customers to recover or non-renewed services
        - **Sheet 4 (Only in P2):** Celebrate new acquisitions and expansion
        - **Sheet 5 (Common):** Analyze growth in the loyal base
        """)
    
    # Highlighted download button
    st.download_button(
        label="📥 DOWNLOAD COMPLETE EXCEL FILE",
        data=output.getvalue(),
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.balloons()
    st.success(f"🎉 **Analysis completed!** File ready to download: `{nombre_archivo}`")

else:
    # Message when no files are loaded
    st.info("👆 **Start by uploading the two sales files in the section above**")
    
    st.markdown("""
    ### 📌 Reminder:
    1. Export your data from **Power BI** in Excel (.xlsx) or CSV (.csv) format
    2. You need **TWO files**: one for each period
    3. Make sure they include standard Power BI columns
    
    Need help? Open the complete guide at the top of the page.
    """)
