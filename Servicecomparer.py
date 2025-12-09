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
st.markdown("### Tool to analyze and compare sales between two periods from a single file")

# Initial information and guide
with st.expander("ℹ️ **HOW TO USE THIS APPLICATION** - Click here to see the guide", expanded=False):
    st.markdown("""
    ## 🎯 Purpose
    This tool allows you to compare sales between two different periods from a **single file** exported from Power BI.
    
    ## 📋 Steps to follow:
    
    ### **STEP 1:** Prepare your file
 - Export sales data from Power BI in **Excel (.xlsx)** or **CSV (.csv)** format

**📌 How to export from Power BI:**
1. Open your sales report in Power BI
2. Apply filter on **ProductType** column
3. Select ONLY these product types:
   - ✅ Cost Item (Service)
   - ✅ Service Contract
   - ✅ Spare Part (Service)
   - ✅ Training
   - ✅ Wear Part (Service)
4. Click on the visual/table → **More options (...)** → **Export data**
5. Choose Excel (.xlsx) or CSV format
6. Save the file

    - Include data covering **both periods** you want to compare
    - Make sure it includes these columns:
      - `Date` - Transaction date
      - `Business Partner Name` - Customer name
      - `ItemIdAndName` - Product or service
      - `ProductType` - Product type
      - `Qty` - Quantity sold
      - `EUR` - Amount in local currency
      - `SalesRepresentative` - Sales representative
      - `Set` and `Productline` - Groupings
    
    ### **STEP 2:** Upload the file
    - Upload your consolidated sales file
    
    ### **STEP 3:** Define periods
    - Select the date range for **Period 1**
    - Select the date range for **Period 2**
    - You can use the calendar or type dates manually (YYYY-MM-DD format)
    
    ### **STEP 4:** Apply filters
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

# Function to parse manual date input
def parse_date_input(date_str):
    """Parse date string in format YYYY-MM-DD"""
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except:
        return None

# =============================================================================
# STEP 1: FILE UPLOAD
# =============================================================================
st.markdown("## 📁 STEP 1: File Upload")
st.markdown("Upload your sales file exported from Power BI (should contain data for both periods)")

file = st.file_uploader(
    "Sales file (Excel or CSV)", 
    type=["xlsx", "csv"], 
    key="main_file",
    help="Export from Power BI: Data → Export data → .xlsx or .csv"
)

# Processing and comparison
if file:
    st.success("✅ **File loaded successfully**")
    
    df = load_file(file)
    
    if df is None:
        st.stop()
    
    # =============================================================================
    # DATA PREVIEW
    # =============================================================================
    st.markdown("---")
    st.markdown("### 👀 Data Preview")
    st.info(f"🔍 Loaded {len(df)} records. Verify that the data is correct before continuing")
    
    st.dataframe(df.head(10), use_container_width=True)
    
    # =============================================================================
    # STEP 2: COLUMN ASSIGNMENT
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🛠️ STEP 2: Column Assignment")
    st.markdown("The application detects columns automatically. **Only adjust if necessary.**")
    
    # Assign fixed column names
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
    
    # Detect columns
    col_fecha = default_cols['Date'] if default_cols['Date'] in df.columns else df.columns[0]
    col_cliente = default_cols['Customer'] if default_cols['Customer'] in df.columns else df.columns[0]
    col_producto = default_cols['Product'] if default_cols['Product'] in df.columns else df.columns[0]
    col_tipo = default_cols['Product Type'] if default_cols['Product Type'] in df.columns else df.columns[0]
    col_cantidad = default_cols['Quantity'] if default_cols['Quantity'] in df.columns else df.columns[0]
    col_precio = default_cols['Amount'] if default_cols['Amount'] in df.columns else df.columns[0]
    col_sales_rep = default_cols['SalesRepresentative'] if default_cols['SalesRepresentative'] in df.columns else df.columns[0]
    col_set = default_cols['Set'] if default_cols['Set'] in df.columns else df.columns[0]
    col_productline = default_cols['Productline'] if default_cols['Productline'] in df.columns else df.columns[0]
    
    st.success(f"✅ Columns detected automatically.")
    
    with st.expander("🔍 View detected columns", expanded=False):
        st.markdown("**Detected columns:**")
        st.write(f"- Date: `{col_fecha}`")
        st.write(f"- Customer: `{col_cliente}`")
        st.write(f"- Product: `{col_producto}`")
        st.write(f"- Type: `{col_tipo}`")
        st.write(f"- Quantity: `{col_cantidad}`")
        st.write(f"- Amount: `{col_precio}`")
        st.write(f"- Sales Rep: `{col_sales_rep}`")
        st.write(f"- Set: `{col_set}`")
        st.write(f"- Productline: `{col_productline}`")
    
    # Format dates
    df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce').dt.date
    
    # Get date range from data
    min_date = pd.to_datetime(df[col_fecha]).min()
    max_date = pd.to_datetime(df[col_fecha]).max()
    
    # =============================================================================
    # STEP 3: PERIOD DEFINITION
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📅 STEP 3: Define Periods to Compare")
    st.markdown(f"**Available date range in file:** {min_date.date()} to {max_date.date()}")
    st.info("💡 **Tip:** You can use the calendar picker or type dates manually in YYYY-MM-DD format (e.g., 2024-01-01)")
    
    col1, col2 = st.columns(2)
    
    # ==================== PERIOD 1 ====================
    with col1:
        st.markdown("### 📅 Period 1")
        
        nombre_periodo_1 = st.text_input(
            "Name for Period 1 (e.g., 'Q1 2024', 'January 2024')", 
            value="Period 1", 
            key="nombre_p1",
            help="Give it a descriptive name"
        )
        
        # Toggle between calendar and manual input
        use_calendar_p1 = st.checkbox("Use calendar picker", value=False, key="calendar_p1")
        
        if use_calendar_p1:
            # Calendar picker
            date_range_1 = st.date_input(
                f"Select date range for {nombre_periodo_1}", 
                [min_date, min_date + pd.Timedelta(days=30)],
                min_value=min_date,
                max_value=max_date,
                key="date_range_p1_cal"
            )
            if len(date_range_1) == 2:
                start_date_1, end_date_1 = date_range_1
            else:
                start_date_1 = date_range_1[0]
                end_date_1 = date_range_1[0]
        else:
            # Manual input
            col_start, col_end = st.columns(2)
            with col_start:
                start_str_1 = st.text_input(
                    "Start date (YYYY-MM-DD)",
                    value=min_date.strftime("%Y-%m-%d"),
                    key="start_manual_p1"
                )
                start_date_1 = parse_date_input(start_str_1)
                if start_date_1 is None:
                    st.error("Invalid date format. Use YYYY-MM-DD")
                    st.stop()
            
            with col_end:
                end_str_1 = st.text_input(
                    "End date (YYYY-MM-DD)",
                    value=(min_date + pd.Timedelta(days=30)).strftime("%Y-%m-%d"),
                    key="end_manual_p1"
                )
                end_date_1 = parse_date_input(end_str_1)
                if end_date_1 is None:
                    st.error("Invalid date format. Use YYYY-MM-DD")
                    st.stop()
        
        st.caption(f"Selected: {start_date_1} to {end_date_1}")
    
    # ==================== PERIOD 2 ====================
    with col2:
        st.markdown("### 📅 Period 2")
        
        nombre_periodo_2 = st.text_input(
            "Name for Period 2 (e.g., 'Q1 2023', 'January 2023')", 
            value="Period 2", 
            key="nombre_p2",
            help="Give it a descriptive name"
        )
        
        # Toggle between calendar and manual input
        use_calendar_p2 = st.checkbox("Use calendar picker", value=False, key="calendar_p2")
        
        if use_calendar_p2:
            # Calendar picker
            date_range_2 = st.date_input(
                f"Select date range for {nombre_periodo_2}", 
                [max_date - pd.Timedelta(days=30), max_date],
                min_value=min_date,
                max_value=max_date,
                key="date_range_p2_cal"
            )
            if len(date_range_2) == 2:
                start_date_2, end_date_2 = date_range_2
            else:
                start_date_2 = date_range_2[0]
                end_date_2 = date_range_2[0]
        else:
            # Manual input
            col_start, col_end = st.columns(2)
            with col_start:
                start_str_2 = st.text_input(
                    "Start date (YYYY-MM-DD)",
                    value=(max_date - pd.Timedelta(days=30)).strftime("%Y-%m-%d"),
                    key="start_manual_p2"
                )
                start_date_2 = parse_date_input(start_str_2)
                if start_date_2 is None:
                    st.error("Invalid date format. Use YYYY-MM-DD")
                    st.stop()
            
            with col_end:
                end_str_2 = st.text_input(
                    "End date (YYYY-MM-DD)",
                    value=max_date.strftime("%Y-%m-%d"),
                    key="end_manual_p2"
                )
                end_date_2 = parse_date_input(end_str_2)
                if end_date_2 is None:
                    st.error("Invalid date format. Use YYYY-MM-DD")
                    st.stop()
        
        st.caption(f"Selected: {start_date_2} to {end_date_2}")
    
    # Validate date ranges
    if start_date_1 > end_date_1:
        st.error(f"❌ Period 1: Start date must be before end date")
        st.stop()
    
    if start_date_2 > end_date_2:
        st.error(f"❌ Period 2: Start date must be before end date")
        st.stop()
    
    # Filter data by periods
    df1 = df[(pd.to_datetime(df[col_fecha]) >= pd.to_datetime(start_date_1)) & 
             (pd.to_datetime(df[col_fecha]) <= pd.to_datetime(end_date_1))].copy()
    
    df2 = df[(pd.to_datetime(df[col_fecha]) >= pd.to_datetime(start_date_2)) & 
             (pd.to_datetime(df[col_fecha]) <= pd.to_datetime(end_date_2))].copy()
    
    st.success(f"✅ **Periods defined:** {len(df1)} records in {nombre_periodo_1}, {len(df2)} records in {nombre_periodo_2}")
    
    # Warning if periods overlap
    if not (end_date_1 < start_date_2 or end_date_2 < start_date_1):
        st.warning("⚠️ **Warning:** The selected periods overlap. This may affect the comparison results.")
    
    # =============================================================================
    # STEP 4: FILTERS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 🎯 STEP 4: Apply Filters")
    
    # Product type filter
    st.markdown("### 🏷️ Product Type Filter")
    tipos_disponibles = sorted(set(df1[col_tipo].dropna().unique()) | set(df2[col_tipo].dropna().unique()))
    
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
    
    df1_filtrado = df1[df1[col_tipo].isin(tipos_seleccionados)].copy()
    df2_filtrado = df2[df2[col_tipo].isin(tipos_seleccionados)].copy()
    
    st.success(f"✅ Filters applied: {len(df1_filtrado)} records in {nombre_periodo_1}, {len(df2_filtrado)} records in {nombre_periodo_2}")

    # Filter by Product/Service Name
    st.markdown("### 🔍 Filter by Product/Service Name (Optional)")

    search_text = st.text_input(
        "Search in ItemIdAndName (leave empty to include all)",
        value="",
        placeholder="e.g., 'pump', 'valve', 'maintenance'...",
        help="Filter products/services that contain this text (case insensitive)"
    )

    if search_text:
        # Apply filter
        df1_filtrado = df1_filtrado[df1_filtrado[col_producto].str.contains(search_text, case=False, na=False)]
        df2_filtrado = df2_filtrado[df2_filtrado[col_producto].str.contains(search_text, case=False, na=False)]
        
        st.success(f"✅ Name filter applied: {len(df1_filtrado)} records in {nombre_periodo_1}, {len(df2_filtrado)} records in {nombre_periodo_2}")
        
        if len(df1_filtrado) == 0 or len(df2_filtrado) == 0:
            st.warning("⚠️ No records found with that search term. Try a different keyword.")
            st.stop()
    else:
        st.info("ℹ️ No name filter applied - showing all products")
    
    # =============================================================================
    # DATA PROCESSING
    # =============================================================================
    
    # Convert quantity and amount to numeric
    df1_filtrado[col_cantidad] = pd.to_numeric(df1_filtrado[col_cantidad], errors='coerce')
    df1_filtrado[col_precio] = pd.to_numeric(df1_filtrado[col_precio], errors='coerce')
    df2_filtrado[col_cantidad] = pd.to_numeric(df2_filtrado[col_cantidad], errors='coerce')
    df2_filtrado[col_precio] = pd.to_numeric(df2_filtrado[col_precio], errors='coerce')
    
    # Amount is already the total
    df1_filtrado["Amount"] = df1_filtrado[col_precio]
    df2_filtrado["Amount"] = df2_filtrado[col_precio]
    
    # Group by customer, product, sales representative, set and productline
    with st.spinner("🔄 Processing data and generating comparison..."):
        grouped_1 = df1_filtrado.groupby([col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]).agg({
            col_cantidad: "sum",
            "Amount": "sum"
        }).rename(columns={col_cantidad: f"Quantity {nombre_periodo_1}", "Amount": f"Amount {nombre_periodo_1}"})
        
        grouped_2 = df2_filtrado.groupby([col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]).agg({
            col_cantidad: "sum",
            "Amount": "sum"
        }).rename(columns={col_cantidad: f"Quantity {nombre_periodo_2}", "Amount": f"Amount {nombre_periodo_2}"})
        
        comparativa = pd.merge(grouped_1, grouped_2, how="outer", left_index=True, right_index=True).fillna(0)
        comparativa["Quantity Difference"] = comparativa[f"Quantity {nombre_periodo_2}"] - comparativa[f"Quantity {nombre_periodo_1}"]
        comparativa["Amount Difference"] = comparativa[f"Amount {nombre_periodo_2}"] - comparativa[f"Amount {nombre_periodo_1}"]
    
    # =============================================================================
    # STEP 5: RESULTS AND METRICS
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📊 STEP 5: Analysis Results")
    
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
    # STEP 6: DOWNLOAD
    # =============================================================================
    st.markdown("---")
    st.markdown("## 📥 STEP 6: Download Results")
    
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
        keys_1 = set(df1_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1))
        keys_2 = set(df2_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1))
        
        unicos_1 = keys_1 - keys_2
        unicos_2 = keys_2 - keys_1
        comunes = keys_1 & keys_2
        
        # DataFrames of unique and common
        if unicos_1:
            df_unicos_1 = df1_filtrado[df1_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1).isin(unicos_1)]
            df_unicos_1.to_excel(writer, index=False, sheet_name=f'Only in {nombre_periodo_1}'[:31])
        
        if unicos_2:
            df_unicos_2 = df2_filtrado[df2_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1).isin(unicos_2)]
            df_unicos_2.to_excel(writer, index=False, sheet_name=f'Only in {nombre_periodo_2}'[:31])
        
        if comunes:
            df_comunes_1 = df1_filtrado[df1_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1).isin(comunes)].copy()
            df_comunes_2 = df2_filtrado[df2_filtrado[[col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]].apply(tuple, axis=1).isin(comunes)].copy()
            
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
    
    
    st.success(f"🎉 **Analysis completed!** File ready to download: `{nombre_archivo}`")

else:
    # Message when no file is loaded
    st.info("👆 **Start by uploading your sales file in the section above**")
    
    st.markdown("""
    ### 📌 Reminder:
    1. Export your data from **Power BI** in Excel (.xlsx) or CSV (.csv) format
    2. Make sure the file includes data for **both periods** you want to compare
    3. Include standard Power BI columns (Date, Business Partner Name, ItemIdAndName, etc.)
    
    Need help? Open the complete guide at the top of the page.
    """)
