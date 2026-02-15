import streamlit as st
import pandas as pd
import tempfile
import os
from io import BytesIO
from datetime import datetime, timedelta
from calendar import monthrange
import logging
import plotly.express as px
import plotly.graph_objects as go
from buchi_streamlit_theme import apply_buchi_styles
from sales_comparison_report_generator import generate_sales_comparison_html
from column_mappings import (
    detect_format, 
    get_mapping_for_format, 
    get_additional_columns,
    validate_format,
    get_format_info,
    REQUIRED_COLUMNS
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="Sales Comparison Tool", 
    layout="wide",
    page_icon="📊"
)

# Apply BUCHI corporate styles
apply_buchi_styles()

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'file_key' not in st.session_state:
    st.session_state.file_key = None

# Main title with BUCHI styling
st.markdown('<div class="main-header">📊 Sales Comparison by Period</div>', unsafe_allow_html=True)
st.markdown("### Advanced tool to analyze and compare sales between two periods")

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
    - Filter by sales representative (optional)
    - Search by product/service name (optional)
    
    ### **STEP 5:** Review analytics
    - View financial metrics
    - Analyze retention and acquisition rates
    - Review top 10 rankings
    - Interactive visualizations
    
    ### **STEP 6:** Download results
    - Download the Excel file with the complete analysis
    - Export to CSV if needed
    
    ---
    
    ## 📊 What will you get?
    An Excel file with **6 sheets**:
    1. **Comparison** - Complete table with all sales
    2. **Original Data** - Complete transactions from both periods
    3. **Only in Period 1** - Sales that didn't repeat (lost customers)
    4. **Only in Period 2** - New sales (gained customers)
    5. **Common in both** - Recurring sales (loyal customers)
    6. **Configuration** - Analysis parameters for reproducibility
    
    ## 💡 Usage examples:
    - Compare **Q1 2024 vs Q1 2023** → Year-over-year growth
    - Compare **January vs February** → Monthly evolution
    - Filter by representative → Evaluate individual performance
    """)

# Visual separator
st.markdown("---")

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

@st.cache_data
@st.cache_data
def load_file(file):
    """Load file and apply format detection and mapping"""
    if file is not None:
        try:
            # Load raw data
            if file.name.endswith(".csv"):
                df = pd.read_csv(file, encoding='utf-8')
            else:
                df = pd.read_excel(file)
            
            logger.info(f"File loaded: {len(df)} rows, {len(df.columns)} columns")
            logger.info(f"Available columns: {df.columns.tolist()}")
            
            # Detect format
            format_type = detect_format(df.columns.tolist())
            logger.info(f"Detected format: {format_type}")
            
            if format_type == 'unknown':
                st.error("❌ **Unknown file format detected**")
                st.error(f"Available columns: {', '.join(df.columns.tolist())}")
                st.info("💡 This application supports Power BI exports in original or new multi-currency format")
                return None
            
            # Show format info
            format_info = get_format_info(format_type)
            st.sidebar.success(f"✅ **{format_info['name']}** detected")
            st.sidebar.info(f"📊 {format_info['description']}")
            st.sidebar.caption(f"💰 Currency: {format_info['currency']}")
            
            # Validate format
            is_valid, missing_cols = validate_format(df, format_type)
            if not is_valid:
                st.error(f"❌ Missing required columns for {format_type} format: {', '.join(missing_cols)}")
                return None
            
            # Apply column mapping
            mapping = get_mapping_for_format(format_type)
            reverse_mapping = {v: k for k, v in mapping.items()}
            
            # Rename columns to standardized names
            df = df.rename(columns=reverse_mapping)

            logger.info("Column mapping applied successfully")
            logger.info(f"Standardized columns: {df.columns.tolist()}")

            # ✅ NUEVO: Rellenar Business Partner Name vacíos usando Id - Name como fallback
            if 'Business Partner Name' in df.columns:
                empty_mask = df['Business Partner Name'].isna()
                empty_count = empty_mask.sum()
                
                if empty_count > 0:
                    logger.info(f"Found {empty_count} empty Business Partner Names")
                    
                    # Buscar columna con nombre del cliente (Id - Name)
                    id_name_col = None
                    for col in df.columns:
                        if 'Id - Name' in col or 'ID - Name' in col:
                            id_name_col = col
                            break
                    
                    if id_name_col and id_name_col in df.columns:
                        # Extraer parte después del " - " 
                        def extract_name_from_id(value):
                            if pd.isna(value):
                                return '(Unknown Customer)'
                            value_str = str(value)
                            if ' - ' in value_str:
                                return value_str.split(' - ', 1)[1].strip()
                            return value_str.strip()
                        
                        # Aplicar fallback solo a los vacíos
                        df.loc[empty_mask, 'Business Partner Name'] = df.loc[empty_mask, id_name_col].apply(extract_name_from_id)
                        logger.info(f"Filled {empty_count} empty Business Partner Names from '{id_name_col}'")
                    else:
                        # Si no existe Id - Name, usar Unknown
                        df.loc[empty_mask, 'Business Partner Name'] = '(Unknown Customer)'
                        logger.info(f"No Id-Name column found, filled with '(Unknown Customer)'")
            
            # Process dates
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['Year'] = df['Date'].dt.year
            df['Month'] = df['Date'].dt.month
            df['Month_Name'] = df['Date'].dt.strftime('%B')
            
            # Store format type in session state
            st.session_state.file_format = format_type
            
            return df
            
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            logger.error(f"Error loading file: {str(e)}", exc_info=True)
            return None
    return None

def validate_dataframe(df, required_columns):
    """Validate that DataFrame has required columns (already mapped)"""
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing required standardized columns: {', '.join(missing_cols)}")
        st.info("📋 Available standardized columns: " + ", ".join(df.columns.tolist()))
        logger.warning(f"Missing standardized columns: {missing_cols}")
        return False
    logger.info("All required standardized columns present")
    return True

def safe_date_conversion(df, col_fecha):
    """Convert dates safely with feedback"""
    try:
        df = df.copy()  # Create explicit copy to avoid SettingWithCopyWarning
        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
        invalid_dates = df[col_fecha].isna().sum()
        if invalid_dates > 0:
            st.warning(f"⚠️ {invalid_dates} invalid dates found and excluded")
            logger.warning(f"{invalid_dates} invalid dates excluded")
        df = df.dropna(subset=[col_fecha])
        df[col_fecha] = df[col_fecha].dt.date
        logger.info(f"Dates converted successfully: {len(df)} valid records")
        return df
    except Exception as e:
        logger.error(f"Error processing dates: {str(e)}")
        st.error(f"❌ Error processing dates: {str(e)}")
        return None

def parse_date_input(date_str):
    """Parse date string in format YYYY-MM-DD"""
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except:
        return None

def create_comparison_chart(total_p1, total_p2, nombre_p1, nombre_p2):
    """Create comparison bar chart"""
    df_viz = pd.DataFrame({
        'Period': [nombre_p1, nombre_p2],
        'Sales': [total_p1, total_p2]
    })
    
    fig = px.bar(
        df_viz, 
        x='Period', 
        y='Sales',
        title='📊 Total Sales Comparison',
        text_auto='.2s',
        color='Period',
        color_discrete_sequence=['#FF6B6B', '#4ECDC4']
    )
    fig.update_traces(texttemplate='€%{text}', textposition='outside')
    fig.update_layout(
        showlegend=False,
        height=400,
        yaxis_title="Sales (€)",
        xaxis_title=""
    )
    return fig

def create_retention_pie_chart(comunes, solo_p1, solo_p2, nombre_p1, nombre_p2):
    """Create retention pie chart"""
    df_pie = pd.DataFrame({
        'Category': ['Common', f'Only {nombre_p1}', f'Only {nombre_p2}'],
        'Count': [comunes, solo_p1, solo_p2]
    })
    
    fig_pie = px.pie(
        df_pie, 
        values='Count', 
        names='Category',
        title='🎯 Record Distribution',
        color_discrete_sequence=['#4ECDC4', '#FF6B6B', '#95E1D3']
    )
    fig_pie.update_layout(height=400)
    return fig_pie

def create_growth_chart(comparativa, nombre_p1, nombre_p2):
    """Create growth trend chart for top products"""
    top_growth = comparativa.nlargest(10, 'Amount Difference').reset_index()
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=nombre_p1,
        y=top_growth[top_growth.columns[1]],  # Product name
        x=top_growth[f"Amount {nombre_p1}"],
        orientation='h',
        marker=dict(color='#FF6B6B')
    ))
    fig.add_trace(go.Bar(
        name=nombre_p2,
        y=top_growth[top_growth.columns[1]],
        x=top_growth[f"Amount {nombre_p2}"],
        orientation='h',
        marker=dict(color='#4ECDC4')
    ))
    
    fig.update_layout(
        title='📈 Top 10 Products by Growth',
        barmode='group',
        height=500,
        xaxis_title="Sales (€)",
        yaxis_title="Product"
    )
    return fig

# ============================================================================
# SIDEBAR: FILE UPLOAD
# ============================================================================
st.sidebar.markdown("## 📁 Data Upload")
uploaded_file = st.sidebar.file_uploader(
    "Upload Sales File (Excel or CSV)", 
    type=["xlsx", "csv"], 
    key="main_file",
    help="Export from Power BI: Data → Export data → .xlsx or .csv"
)

if uploaded_file:
    # Identificador estable del fichero (nombre + tamaño) para evitar recalcular en cada rerun
    file_key = f"{uploaded_file.name}:{uploaded_file.size}"
    
    # Si cambia el archivo, reiniciar session state
    if st.session_state.get("file_key") != file_key:
        st.session_state.file_key = file_key
        st.session_state.df = None
    
    # Cargar archivo (solo si no está ya cargado)
    if st.session_state.df is None:
        df = load_file(uploaded_file)
        
        if df is None:
            st.stop()
        
        st.session_state.df = df
    else:
        df = st.session_state.df
    
    st.sidebar.success(f"✅ {len(df):,} records loaded")
    
    # Validate required columns (now using standardized names from mapping)
    if not validate_dataframe(df, REQUIRED_COLUMNS):
        st.stop()
    
    # Get available options for filters
    available_years = sorted(df['Year'].dropna().unique().astype(int).tolist()) if 'Year' in df.columns else []
    available_reps = sorted(df['SalesRepresentative'].dropna().unique().tolist())
    available_types = sorted(df['ProductType'].dropna().unique().tolist())
    available_sets = sorted(df['Set'].dropna().unique().tolist())
    
    # Rellenar End User Segment y Market Organization vacíos
    if 'End User Segment' in df.columns:
        df['End User Segment'] = df['End User Segment'].fillna('Unknown')

    if 'Market Organization Name' in df.columns:
        df['Market Organization Name'] = df['Market Organization Name'].fillna('Unknown')

    if 'Sales Territory' in df.columns:
        df['Sales Territory'] = df['Sales Territory'].fillna('Unknown')

    if 'Country' in df.columns:
        df['Country'] = df['Country'].fillna('Unknown')
    
    
    month_options = list(range(1, 13))
    month_labels = {
        1: 'January', 2: 'February', 3: 'March', 4: 'April',
        5: 'May', 6: 'June', 7: 'July', 8: 'August',
        9: 'September', 10: 'October', 11: 'November', 12: 'December'
    }
    
    # Reset function
    def reset_all_filters():
        st.session_state["type_filter"] = available_types
        st.session_state["set_filter"] = available_sets
        st.session_state["rep_filter"] = available_reps
        st.session_state["search_filter"] = ""
        st.session_state["customer_filter"] = ""
        st.session_state["selected_quick_filters"] = []
                # Reset new filters (only if they exist)
        if 'Market Organization Name' in df.columns:
            st.session_state["mo_filter"] = sorted(df['Market Organization Name'].dropna().unique().tolist())
        if 'Sales Territory' in df.columns:
            st.session_state["ter_filter"] = sorted(df['Sales Territory'].dropna().unique().tolist())
        if 'Country' in df.columns:
            st.session_state["country_filter"] = sorted(df['Country'].dropna().unique().tolist())
        if 'End User Segment' in df.columns:
            st.session_state["seg_filter"] = sorted(df['End User Segment'].dropna().unique().tolist())
    
    # =========================================================================
    # SIDEBAR: FILTERS (COLLAPSIBLE)
    # =========================================================================
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🎛️ Filters")
    
    # Product Type filter - COLLAPSIBLE
    with st.sidebar.expander("🏷️ Product Type", expanded=False):
        col_type1, col_type2 = st.sidebar.columns(2)
        with col_type1:
            st.button("✅ All", key="type_all", use_container_width=True,
                     on_click=lambda: st.session_state.update({"type_filter": available_types}))
        with col_type2:
            st.button("❌ None", key="type_none", use_container_width=True,
                     on_click=lambda: st.session_state.update({"type_filter": []}))
        
        tipos_seleccionados = st.multiselect(
            "Select product types",
            available_types,
            default=available_types,
            key="type_filter",
            label_visibility="collapsed",
            help="You can exclude types like 'Rental', 'Sample', etc."
        )
    
    # Set filter - COLLAPSIBLE
    with st.sidebar.expander("📦 Set", expanded=False):
        col_set1, col_set2 = st.sidebar.columns(2)
        with col_set1:
            st.button("✅ All", key="set_all", use_container_width=True,
                     on_click=lambda: st.session_state.update({"set_filter": available_sets}))
        with col_set2:
            st.button("❌ None", key="set_none", use_container_width=True,
                     on_click=lambda: st.session_state.update({"set_filter": []}))
        
        selected_sets = st.multiselect(
            "Select sets",
            available_sets,
            default=available_sets,
            key="set_filter",
            label_visibility="collapsed"
        )
    
    # Sales Representative filter - COLLAPSIBLE
    with st.sidebar.expander("👤 Sales Representative", expanded=False):
        col_rep1, col_rep2 = st.sidebar.columns(2)
        with col_rep1:
            st.button("✅ All", key="rep_all", use_container_width=True,
                     on_click=lambda: st.session_state.update({"rep_filter": available_reps}))
        with col_rep2:
            st.button("❌ None", key="rep_none", use_container_width=True,
                     on_click=lambda: st.session_state.update({"rep_filter": []}))
        
        selected_reps = st.multiselect(
            "Select representatives",
            available_reps,
            default=available_reps,
            key="rep_filter",
            label_visibility="collapsed"
        )
    
    # Search filter - COLLAPSIBLE with Quick Filters
    with st.sidebar.expander("🔍 Search Service", expanded=False):
        # AND/OR selector
        quick_mode = st.radio(
            "Quick filter mode:",
            options=['AND', 'OR'],
            horizontal=True,
            key="quick_filter_mode",
            help="AND = All keywords must match | OR = Any keyword matches"
        )
        
        # Initialize quick filters in session state
        if 'selected_quick_filters' not in st.session_state:
            st.session_state.selected_quick_filters = []
        
        # Quick filter keywords
        quick_filter_keywords = ['CARE', 'Exact', 'Start', 'Circle', 'Maintain', 'IQ/OQ', 'OQ', 'Install', 'Academy']
        
        # Display buttons in 3 columns
        cols = st.columns(3)
        for idx, keyword in enumerate(quick_filter_keywords):
            col = cols[idx % 3]
            # Check if button is active
            is_active = keyword in st.session_state.selected_quick_filters
            if col.button(
                keyword, 
                key=f"quick_{keyword}",
                use_container_width=True,
                type="primary" if is_active else "secondary"
            ):
                # Toggle keyword
                if is_active:
                    st.session_state.selected_quick_filters.remove(keyword)
                else:
                    st.session_state.selected_quick_filters.append(keyword)
                st.rerun()
        
        # Manual search input
        search_text = st.text_input(
            "Or type custom search",
            placeholder="e.g., 'pump', 'valve', 'maintenance'...",
            key="search_filter",
            help="Filter products/services that contain this text (case insensitive)",
            label_visibility="collapsed"
        )
    
    # Customer filter - COLLAPSIBLE
    with st.sidebar.expander("👥 Customer", expanded=False):
        customer_search = st.text_input(
            "Filter by Customer Name",
            placeholder="e.g., 'Universidad', 'Hospital'...",
            key="customer_filter",
            help="Filter by customer name (case insensitive)",
            label_visibility="collapsed"
        )
    

    # Market Organization filter - COLLAPSIBLE
    with st.sidebar.expander("🏢 Market Organization", expanded=False):
        # Get available market organizations (only if column exists)
        available_market_orgs = []
        if 'Market Organization Name' in df.columns:
            available_market_orgs = sorted(df['Market Organization Name'].dropna().unique().tolist())
        
        if available_market_orgs:
            col_mo1, col_mo2 = st.sidebar.columns(2)
            with col_mo1:
                st.button("✅ All", key="mo_all", use_container_width=True,
                         on_click=lambda: st.session_state.update({"mo_filter": available_market_orgs}))
            with col_mo2:
                st.button("❌ None", key="mo_none", use_container_width=True,
                         on_click=lambda: st.session_state.update({"mo_filter": []}))
            
            selected_market_orgs = st.multiselect(
                "Select market organizations",
                available_market_orgs,
                default=available_market_orgs,
                key="mo_filter",
                label_visibility="collapsed"
            )
        else:
            st.info("Not available in this file format")
    
    # Sales Territory filter - COLLAPSIBLE
    with st.sidebar.expander("🌍 Sales Territory", expanded=False):
        available_territories = []
        if 'Sales Territory' in df.columns:
            available_territories = sorted(df['Sales Territory'].dropna().unique().tolist())
        
        if available_territories:
            col_ter1, col_ter2 = st.sidebar.columns(2)
            with col_ter1:
                st.button("✅ All", key="ter_all", use_container_width=True,
                         on_click=lambda: st.session_state.update({"ter_filter": available_territories}))
            with col_ter2:
                st.button("❌ None", key="ter_none", use_container_width=True,
                         on_click=lambda: st.session_state.update({"ter_filter": []}))
            
            selected_territories = st.multiselect(
                "Select territories",
                available_territories,
                default=available_territories,
                key="ter_filter",
                label_visibility="collapsed"
            )
        else:
            st.info("Not available in this file format")
    
    # Country filter - COLLAPSIBLE
    with st.sidebar.expander("🌎 Country", expanded=False):
        available_countries = []
        if 'Country' in df.columns:
            available_countries = sorted(df['Country'].dropna().unique().tolist())
        
        if available_countries:
            col_country1, col_country2 = st.sidebar.columns(2)
            with col_country1:
                st.button("✅ All", key="country_all", use_container_width=True,
                         on_click=lambda: st.session_state.update({"country_filter": available_countries}))
            with col_country2:
                st.button("❌ None", key="country_none", use_container_width=True,
                         on_click=lambda: st.session_state.update({"country_filter": []}))
            
            selected_countries = st.multiselect(
                "Select countries",
                available_countries,
                default=available_countries,
                key="country_filter",
                label_visibility="collapsed"
            )
        else:
            st.info("Not available in this file format")
    
    # End User Segment filter - COLLAPSIBLE
    with st.sidebar.expander("👥 End User Segment", expanded=False):
        available_segments = []
        if 'End User Segment' in df.columns:
            available_segments = sorted(df['End User Segment'].dropna().unique().tolist())
        
        if available_segments:
            col_seg1, col_seg2 = st.sidebar.columns(2)
            with col_seg1:
                st.button("✅ All", key="seg_all", use_container_width=True,
                         on_click=lambda: st.session_state.update({"seg_filter": available_segments}))
            with col_seg2:
                st.button("❌ None", key="seg_none", use_container_width=True,
                         on_click=lambda: st.session_state.update({"seg_filter": []}))
            
            selected_segments = st.multiselect(
                "Select segments",
                available_segments,
                default=available_segments,
                key="seg_filter",
                label_visibility="collapsed"
            )
        else:
            st.info("Not available in this file format")


    # RESET BUTTON
    st.sidebar.markdown("---")
    st.sidebar.button(
        "🔄 Reset All Filters",
        type="primary",
        use_container_width=True,
        on_click=reset_all_filters,
        key="reset_btn"
    )
    
    # =========================================================================
    # DATA PREVIEW
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">👀 Data Preview</div>', unsafe_allow_html=True)
    st.info(f"🔍 Loaded {len(df):,} records. Verify that the data is correct before continuing")
    
    with st.expander("📋 View first 20 rows", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)
    
    # =========================================================================
    # COLUMN ASSIGNMENT
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">🛠️ Column Assignment</div>', unsafe_allow_html=True)
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
    
    # =========================================================================
    # COLUMN ASSIGNMENT (Now automatic via mapping)
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">🛠️ Column Assignment</div>', unsafe_allow_html=True)
    
    # Get format info
    format_type = st.session_state.get('file_format', 'unknown')
    format_info = get_format_info(format_type)
    
    st.success(f"✅ **Columns mapped automatically** using **{format_info['name']}**")
    
    # Assign standardized column names (already mapped by load_file)
    col_fecha = 'Date'
    col_cliente = 'Business Partner Name'
    col_producto = 'ItemIdAndName'
    col_tipo = 'ProductType'
    col_cantidad = 'Qty'
    col_precio = 'EUR'  # This is now mapped from LC or EUR depending on format
    col_sales_rep = 'SalesRepresentative'
    col_set = 'Set'
    col_productline = 'Productline'
    
    with st.expander("🔍 View column mapping details", expanded=False):
        st.markdown("**Column mapping applied:**")
        
        mapping = get_mapping_for_format(format_type)
        
        if mapping is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Standardized Name**")
                for std_name in mapping.keys():
                    st.write(f"- {std_name}")
            with col2:
                st.markdown(f"**Original Name ({format_info['name']})**")
                for orig_name in mapping.values():
                    st.write(f"- `{orig_name}`")
            
            # Show additional columns preserved
            additional_cols = get_additional_columns(format_type)
            if additional_cols:
                st.markdown("---")
                st.markdown("**📦 Additional columns preserved:**")
                preserved = [col for col in additional_cols if col in df.columns]
                if preserved:
                    st.write(", ".join([f"`{col}`" for col in preserved]))
                else:
                    st.write("None")
        else:
            st.warning("⚠️ Mapping information not available")
            st.write(f"**Current columns in DataFrame:**")
            st.write(", ".join([f"`{col}`" for col in df.columns.tolist()]))
    
    # Format dates safely
    df = safe_date_conversion(df, col_fecha)
    
    if df is None or df.empty:
        st.error("❌ Tras convertir fechas, no queda ningún registro válido. Revisa el formato de la columna Date del export de Power BI.")
        st.stop()
    
    # Get date range from data
    min_date = pd.to_datetime(df[col_fecha]).min()
    max_date = pd.to_datetime(df[col_fecha]).max()
    
    logger.info(f"Date range: {min_date.date()} to {max_date.date()}")
    
    # Initialize period selection state
    if 'period_preset' not in st.session_state:
        st.session_state.period_preset = None
    if 'period_year_1' not in st.session_state:
        st.session_state.period_year_1 = max_date.year
    if 'period_year_2' not in st.session_state:
        st.session_state.period_year_2 = max_date.year - 1
    if 'period_months' not in st.session_state:
        st.session_state.period_months = None
    
    # =========================================================================
    # PERIOD DEFINITION WITH QUICK PRESETS (REDESIGN + FIX P1/P2 ORDER)
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">📅 Define Periods to Compare</div>', unsafe_allow_html=True)
    st.markdown(f"**Available date range in file:** {min_date.date()} to {max_date.date()}")

    # -------------------------------------------------------------------------
    # Helpers / constants
    # -------------------------------------------------------------------------
    month_labels_full = {
        1: 'January', 2: 'February', 3: 'March', 4: 'April',
        5: 'May', 6: 'June', 7: 'July', 8: 'August',
        9: 'September', 10: 'October', 11: 'November', 12: 'December'
    }

    # Reusa tu definición de presets (ya la tenías). La dejo aquí integrada:
    month_presets = {
        "Full Year": None,          # especial
        "H1 (Jan-Jun)": (1, 6),
        "H2 (Jul-Dec)": (7, 12),
        "Q1 (Jan-Mar)": (1, 3),
        "Q2 (Apr-Jun)": (4, 6),
        "Q3 (Jul-Sep)": (7, 9),
        "Q4 (Oct-Dec)": (10, 12),
        "Single Month": "single",
        "Custom": "custom"
    }

    def clamp(n, min_v, max_v):
        return max(min_v, min(n, max_v))

    def get_default_year_indexes(years_sorted):
        # Queremos: P2 = último año (más nuevo), P1 = penúltimo (si existe)
        idx_p2 = 0
        idx_p1 = min(1, len(years_sorted) - 1)  # si solo hay 1, será 0
        return idx_p1, idx_p2

    def apply_quick_to_session_state(preset_key, year_1, year_2, start_m=None, end_m=None):
        # Full year
        if preset_key == "Full Year" or month_presets.get(preset_key) is None:
            st.session_state.start_date_1 = f"{year_1}-01-01"
            st.session_state.end_date_1 = f"{year_1}-12-31"
            st.session_state.start_date_2 = f"{year_2}-01-01"
            st.session_state.end_date_2 = f"{year_2}-12-31"
            st.session_state.nombre_p1 = str(year_1)
            st.session_state.nombre_p2 = str(year_2)
            return

        # Month range (Q/H/single/custom)
        if start_m is None or end_m is None:
            st.error("⚠️ Please select a valid month range.")
            st.stop()

        st.session_state.start_date_1 = f"{year_1}-{start_m:02d}-01"
        last_day_1 = monthrange(year_1, end_m)[1]
        st.session_state.end_date_1 = f"{year_1}-{end_m:02d}-{last_day_1}"

        st.session_state.start_date_2 = f"{year_2}-{start_m:02d}-01"
        last_day_2 = monthrange(year_2, end_m)[1]
        st.session_state.end_date_2 = f"{year_2}-{end_m:02d}-{last_day_2}"

        # Nombres
        if start_m == end_m:
            label = f"{month_labels_full[start_m]}"
        else:
            label = f"{month_labels_full[start_m]}-{month_labels_full[end_m]}"

        st.session_state.nombre_p1 = f"{label} {year_1}"
        st.session_state.nombre_p2 = f"{label} {year_2}"


    # -------------------------------------------------------------------------
    # Validate years
    # -------------------------------------------------------------------------
    if not available_years or len(available_years) == 0:
        st.error("❌ No years found in data. Please check your Date column.")
        st.stop()

    years_sorted = sorted(available_years, reverse=True)  # Newest first
    default_p1_idx, default_p2_idx = get_default_year_indexes(years_sorted)

    # -------------------------------------------------------------------------
    # Quick selector (solo rellena Period Details de abajo)
    # -------------------------------------------------------------------------
    st.markdown("### ⚡ Quick Period Selection")
    st.info("💡 **Tip:** Choose a preset + years and click Apply to auto-fill the editable Period Details below.")

    col_preset, col_apply = st.columns([3, 1])

    with col_preset:
        preset_key = st.selectbox(
            "Quick preset:",
            options=list(month_presets.keys()),
            index=0,
            key="quick_preset_select_new",
            help="This will autofill the Period Details below"
        )

    # Inicializa inputs para el apply
    year_1 = None
    year_2 = None
    month_range = None  # (start_m, end_m)

    preset_value = month_presets[preset_key]

    # Caso 1: Full Year o Q/H (rango fijo)
    if preset_value is None or isinstance(preset_value, tuple):
        col_y1, col_y2 = st.columns(2)
        with col_y1:
            year_1 = st.selectbox(
                "Period 1 Year:",
                options=years_sorted,
                index=default_p1_idx,  # ✅ P1 = penúltimo (si existe)
                key="quick_year_1_new"
            )
        with col_y2:
            year_2 = st.selectbox(
                "Period 2 Year:",
                options=years_sorted,
                index=default_p2_idx,  # ✅ P2 = último
                key="quick_year_2_new"
            )

        if isinstance(preset_value, tuple):
            month_range = preset_value

    # Caso 2: Single Month
    elif preset_value == "single":
        col_m, col_y1, col_y2 = st.columns([2, 1, 1])
        with col_m:
            month = st.selectbox(
                "Month:",
                options=list(range(1, 13)),
                format_func=lambda x: month_labels_full[x],
                key="quick_single_month_new"
            )
            month_range = (month, month)

        with col_y1:
            year_1 = st.selectbox(
                "Period 1 Year:",
                options=years_sorted,
                index=default_p1_idx,
                key="quick_single_year_1_new"
            )
        with col_y2:
            year_2 = st.selectbox(
                "Period 2 Year:",
                options=years_sorted,
                index=default_p2_idx,
                key="quick_single_year_2_new"
            )

    # Caso 3: Custom Month Range
    elif preset_value == "custom":
        col_m1, col_m2, col_y1, col_y2 = st.columns([1, 1, 1, 1])
        with col_m1:
            start_m = st.selectbox(
                "From month:",
                options=list(range(1, 13)),
                format_func=lambda x: month_labels_full[x],
                key="quick_custom_start_month_new"
            )
        with col_m2:
            end_m = st.selectbox(
                "To month:",
                options=list(range(1, 13)),
                index=clamp(start_m - 1, 0, 11),
                format_func=lambda x: month_labels_full[x],
                key="quick_custom_end_month_new"
            )
        month_range = (start_m, end_m)

        with col_y1:
            year_1 = st.selectbox(
                "Period 1 Year:",
                options=years_sorted,
                index=default_p1_idx,
                key="quick_custom_year_1_new"
            )
        with col_y2:
            year_2 = st.selectbox(
                "Period 2 Year:",
                options=years_sorted,
                index=default_p2_idx,
                key="quick_custom_year_2_new"
            )

    # Apply + Swap
    with col_apply:
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🚀 Apply", type="primary", use_container_width=True, key="apply_preset_new"):
            if year_1 is None or year_2 is None:
                st.error("⚠️ Please select years for both periods.")
                st.stop()

            if month_range is None:
                # Full year
                apply_quick_to_session_state(preset_key, year_1, year_2)
            else:
                sm, em = month_range
                apply_quick_to_session_state(preset_key, year_1, year_2, start_m=sm, end_m=em)

            st.success("✅ Dates applied! You can edit them manually below if needed.")
            st.rerun()

        if st.button("🔁 Swap P1 ↔ P2", use_container_width=True, key="swap_periods_new"):
            # Intercambia los period details ya aplicados (si existen)
            for a, b in [("start_date_1", "start_date_2"), ("end_date_1", "end_date_2"), ("nombre_p1", "nombre_p2")]:
                st.session_state[a], st.session_state[b] = st.session_state.get(b), st.session_state.get(a)
            st.rerun()


    # -------------------------------------------------------------------------
    # Period Details (editable) - tu bloque original casi intacto
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 📝 Period Details (editable)")

    col1, col2 = st.columns(2)

    # ==================== PERIOD 1 ====================
    with col1:
        st.markdown("#### 📅 Period 1")

        default_name_1 = st.session_state.get('nombre_p1', str(max_date.year - 1))
        nombre_periodo_1 = st.text_input(
            "Period 1 Name",
            value=default_name_1,
            key="nombre_p1_input",
            help="Edit to customize the period name"
        )

        col_start, col_end = st.columns(2)
        with col_start:
            default_start_1 = st.session_state.get('start_date_1', f"{max_date.year - 1}-01-01")
            start_str_1 = st.text_input(
                "Start date (YYYY-MM-DD)",
                value=default_start_1,
                key="start_manual_p1"
            )
            start_date_1 = parse_date_input(start_str_1)
            if start_date_1 is None:
                st.error("Invalid date format. Use YYYY-MM-DD")
                st.stop()

        with col_end:
            default_end_1 = st.session_state.get('end_date_1', f"{max_date.year - 1}-12-31")
            end_str_1 = st.text_input(
                "End date (YYYY-MM-DD)",
                value=default_end_1,
                key="end_manual_p1"
            )
            end_date_1 = parse_date_input(end_str_1)
            if end_date_1 is None:
                st.error("Invalid date format. Use YYYY-MM-DD")
                st.stop()

        st.caption(f"📊 Selected: {start_date_1} to {end_date_1}")

    # ==================== PERIOD 2 ====================
    with col2:
        st.markdown("#### 📅 Period 2")

        default_name_2 = st.session_state.get('nombre_p2', str(max_date.year))
        nombre_periodo_2 = st.text_input(
            "Period 2 Name",
            value=default_name_2,
            key="nombre_p2_input",
            help="Edit to customize the period name"
        )

        col_start, col_end = st.columns(2)
        with col_start:
            default_start_2 = st.session_state.get('start_date_2', f"{max_date.year}-01-01")
            start_str_2 = st.text_input(
                "Start date (YYYY-MM-DD)",
                value=default_start_2,
                key="start_manual_p2"
            )
            start_date_2 = parse_date_input(start_str_2)
            if start_date_2 is None:
                st.error("Invalid date format. Use YYYY-MM-DD")
                st.stop()

        with col_end:
            default_end_2 = st.session_state.get('end_date_2', f"{max_date.year}-12-31")
            end_str_2 = st.text_input(
                "End date (YYYY-MM-DD)",
                value=default_end_2,
                key="end_manual_p2"
            )
            end_date_2 = parse_date_input(end_str_2)
            if end_date_2 is None:
                st.error("Invalid date format. Use YYYY-MM-DD")
                st.stop()

        st.caption(f"📊 Selected: {start_date_2} to {end_date_2}")

    # Validate date ranges
    if start_date_1 > end_date_1:
        st.error("❌ Period 1: Start date must be before end date")
        st.stop()

    if start_date_2 > end_date_2:
        st.error("❌ Period 2: Start date must be before end date")
        st.stop()

    # Filter data by periods
    with st.spinner("🔄 Filtering data by periods..."):
        df1 = df[(pd.to_datetime(df[col_fecha]) >= pd.to_datetime(start_date_1)) &
                (pd.to_datetime(df[col_fecha]) <= pd.to_datetime(end_date_1))].copy()

        df2 = df[(pd.to_datetime(df[col_fecha]) >= pd.to_datetime(start_date_2)) &
                (pd.to_datetime(df[col_fecha]) <= pd.to_datetime(end_date_2))].copy()


    # ✅ DEBUG: Check clientes vacíos
    print(f"🔍 DEBUG - df1 clientes vacíos: {df1[col_cliente].isna().sum()}")
    print(f"🔍 DEBUG - df2 clientes vacíos: {df2[col_cliente].isna().sum()}")

    logger.info(f"Period 1: {len(df1)} records, Period 2: {len(df2)} records")

    st.success(f"✅ **Periods defined:** {len(df1):,} records in {nombre_periodo_1}, {len(df2):,} records in {nombre_periodo_2}")

    # Warning if periods overlap
    if not (end_date_1 < start_date_2 or end_date_2 < start_date_1):
        st.warning("⚠️ **Warning:** The selected periods overlap. This may affect the comparison results.")
        logger.warning("Periods overlap detected")
    
    # =========================================================================
    # APPLY FILTERS FROM SIDEBAR
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">🎯 Filtering Data</div>', unsafe_allow_html=True)
    
    # Apply product type filter from sidebar
    df1_filtrado = df1[df1[col_tipo].isin(tipos_seleccionados)].copy()
    df2_filtrado = df2[df2[col_tipo].isin(tipos_seleccionados)].copy()
    
    # Apply set filter from sidebar
    df1_filtrado = df1_filtrado[df1_filtrado[col_set].isin(selected_sets)]
    df2_filtrado = df2_filtrado[df2_filtrado[col_set].isin(selected_sets)]
    
    # Apply sales representative filter from sidebar
    if selected_reps and len(selected_reps) > 0:
        df1_filtrado = df1_filtrado[df1_filtrado[col_sales_rep].isin(selected_reps)]
        df2_filtrado = df2_filtrado[df2_filtrado[col_sales_rep].isin(selected_reps)]
        logger.info(f"Sales rep filter applied: {selected_reps}")
    
    # Apply search filter from sidebar (quick filters or manual search)
    if st.session_state.selected_quick_filters:
        quick_mode = st.session_state.get('quick_filter_mode', 'AND')
        if quick_mode == 'AND':
            mask1 = pd.Series([True] * len(df1_filtrado), index=df1_filtrado.index)
            mask2 = pd.Series([True] * len(df2_filtrado), index=df2_filtrado.index)
            for keyword in st.session_state.selected_quick_filters:
                mask1 &= df1_filtrado[col_producto].str.contains(keyword, case=False, na=False)
                mask2 &= df2_filtrado[col_producto].str.contains(keyword, case=False, na=False)
            df1_filtrado = df1_filtrado[mask1]
            df2_filtrado = df2_filtrado[mask2]
        else:  # OR
            mask1 = pd.Series([False] * len(df1_filtrado), index=df1_filtrado.index)
            mask2 = pd.Series([False] * len(df2_filtrado), index=df2_filtrado.index)
            for keyword in st.session_state.selected_quick_filters:
                mask1 |= df1_filtrado[col_producto].str.contains(keyword, case=False, na=False)
                mask2 |= df2_filtrado[col_producto].str.contains(keyword, case=False, na=False)
            df1_filtrado = df1_filtrado[mask1]
            df2_filtrado = df2_filtrado[mask2]
        logger.info(f"Quick filters applied ({quick_mode}): {st.session_state.selected_quick_filters}")
    elif search_text:
        df1_filtrado = df1_filtrado[df1_filtrado[col_producto].str.contains(search_text, case=False, na=False)]
        df2_filtrado = df2_filtrado[df2_filtrado[col_producto].str.contains(search_text, case=False, na=False)]
        logger.info(f"Manual search filter applied: '{search_text}'")
    
    # Apply customer filter from sidebar
    customer_search = st.session_state.get('customer_filter', '')
    if customer_search:
        df1_filtrado = df1_filtrado[df1_filtrado[col_cliente].str.contains(customer_search, case=False, na=False)]
        df2_filtrado = df2_filtrado[df2_filtrado[col_cliente].str.contains(customer_search, case=False, na=False)]
        logger.info(f"Customer filter applied: '{customer_search}'")


    # Apply Market Organization filter
    if 'Market Organization Name' in df1_filtrado.columns and 'mo_filter' in st.session_state:
        selected_market_orgs = st.session_state.get('mo_filter', [])
        if selected_market_orgs and len(selected_market_orgs) > 0:
            df1_filtrado = df1_filtrado[df1_filtrado['Market Organization Name'].isin(selected_market_orgs)]
            df2_filtrado = df2_filtrado[df2_filtrado['Market Organization Name'].isin(selected_market_orgs)]
            logger.info(f"Market Organization filter applied: {selected_market_orgs}")
    
    # Apply Sales Territory filter
    if 'Sales Territory' in df1_filtrado.columns and 'ter_filter' in st.session_state:
        selected_territories = st.session_state.get('ter_filter', [])
        if selected_territories and len(selected_territories) > 0:
            df1_filtrado = df1_filtrado[df1_filtrado['Sales Territory'].isin(selected_territories)]
            df2_filtrado = df2_filtrado[df2_filtrado['Sales Territory'].isin(selected_territories)]
            logger.info(f"Sales Territory filter applied: {selected_territories}")
    
    # Apply Country filter
    if 'Country' in df1_filtrado.columns and 'country_filter' in st.session_state:
        selected_countries = st.session_state.get('country_filter', [])
        if selected_countries and len(selected_countries) > 0:
            df1_filtrado = df1_filtrado[df1_filtrado['Country'].isin(selected_countries)]
            df2_filtrado = df2_filtrado[df2_filtrado['Country'].isin(selected_countries)]
            logger.info(f"Country filter applied: {selected_countries}")
    
    # Apply End User Segment filter
    if 'End User Segment' in df1_filtrado.columns and 'seg_filter' in st.session_state:
        selected_segments = st.session_state.get('seg_filter', [])
        if selected_segments and len(selected_segments) > 0:
            df1_filtrado = df1_filtrado[df1_filtrado['End User Segment'].isin(selected_segments)]
            df2_filtrado = df2_filtrado[df2_filtrado['End User Segment'].isin(selected_segments)]
            logger.info(f"End User Segment filter applied: {selected_segments}")


    # Show filtering summary
    st.success(f"✅ **Filters applied:** {len(df1_filtrado):,} records in {nombre_periodo_1}, {len(df2_filtrado):,} records in {nombre_periodo_2}")
    
    # ✅ DEBUG: Check clientes vacíos tras filtros
    print(f"🔍 DEBUG - df1_filtrado clientes vacíos: {df1_filtrado[col_cliente].isna().sum()}")
    print(f"🔍 DEBUG - df2_filtrado clientes vacíos: {df2_filtrado[col_cliente].isna().sum()}")

    if len(df1_filtrado) == 0 or len(df2_filtrado) == 0:
        st.warning("⚠️ No records found with the selected filters. Try adjusting your filters.")
        st.stop()
    
    # =========================================================================
    # DATA PROCESSING
    # =========================================================================
    
    # Convert quantity and amount to numeric
    with st.spinner("🔄 Processing numerical data..."):
        df1_filtrado[col_cantidad] = pd.to_numeric(df1_filtrado[col_cantidad], errors='coerce')
        df1_filtrado[col_precio] = pd.to_numeric(df1_filtrado[col_precio], errors='coerce')
        df2_filtrado[col_cantidad] = pd.to_numeric(df2_filtrado[col_cantidad], errors='coerce')
        df2_filtrado[col_precio] = pd.to_numeric(df2_filtrado[col_precio], errors='coerce')
        
        # Amount is already the total
        df1_filtrado["Amount"] = df1_filtrado[col_precio]
        df2_filtrado["Amount"] = df2_filtrado[col_precio]
    
        # ✅ DEBUG: Verificar clientes antes de groupby
        print(f"🔍 DEBUG - Clientes únicos en df1_filtrado: {df1_filtrado[col_cliente].nunique()}")
        print(f"🔍 DEBUG - Clientes vacíos en df1_filtrado: {df1_filtrado[col_cliente].isna().sum()}")
        print(f"🔍 DEBUG - Sample de clientes: {df1_filtrado[col_cliente].head(10).tolist()}")

    # Group by customer, product, sales representative, set and productline
    # Group by customer, product, sales representative, set and productline
    with st.spinner("🔄 Processing data and generating comparison..."):
        progress_bar = st.progress(0)
        
        # Group period 1
        progress_bar.progress(25)
        grouped_1 = df1_filtrado.groupby([col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]).agg({
            col_cantidad: "sum",
            "Amount": "sum"
        })
        # Rename columns explicitly BEFORE merge
        grouped_1.columns = [f"Quantity {nombre_periodo_1}", f"Amount {nombre_periodo_1}"]
        
        # Debug
        logger.info(f"Grouped_1 columns: {grouped_1.columns.tolist()}")
        
        # Group period 2
        progress_bar.progress(50)
        grouped_2 = df2_filtrado.groupby([col_cliente, col_producto, col_sales_rep, col_set, col_productline, col_tipo]).agg({
            col_cantidad: "sum",
            "Amount": "sum"
        })
        # Rename columns explicitly BEFORE merge
        grouped_2.columns = [f"Quantity {nombre_periodo_2}", f"Amount {nombre_periodo_2}"]
        
        # Debug
        logger.info(f"Grouped_2 columns: {grouped_2.columns.tolist()}")
        
        # Create comparison - now column names are already unique so no _x/_y suffixes
        progress_bar.progress(75)
        comparativa = pd.merge(
            grouped_1, 
            grouped_2, 
            how="outer", 
            left_index=True, 
            right_index=True,
            suffixes=('', '')  # No suffixes needed since names are already unique
        ).fillna(0)
        
        # Debug - show what columns we actually have
        logger.info(f"Comparativa columns after merge: {comparativa.columns.tolist()}")
        st.caption(f"🔍 Debug: Comparativa columns = {comparativa.columns.tolist()}")
        
        # Now calculate differences - column names should be correct
        qty_col_1 = f"Quantity {nombre_periodo_1}"
        qty_col_2 = f"Quantity {nombre_periodo_2}"
        amt_col_1 = f"Amount {nombre_periodo_1}"
        amt_col_2 = f"Amount {nombre_periodo_2}"
        
        # Verify columns exist before accessing
        if qty_col_1 not in comparativa.columns or qty_col_2 not in comparativa.columns:
            st.error(f"❌ ERROR: Expected columns not found!")
            st.error(f"Looking for: '{qty_col_1}' and '{qty_col_2}'")
            st.error(f"Available columns: {comparativa.columns.tolist()}")
            st.stop()
        
        comparativa["Quantity Difference"] = comparativa[qty_col_2] - comparativa[qty_col_1]
        comparativa["Amount Difference"] = comparativa[amt_col_2] - comparativa[amt_col_1]
        
        # Calculate growth percentage
        comparativa["Growth %"] = ((comparativa[amt_col_2] - comparativa[amt_col_1]) / 
                                comparativa[amt_col_1] * 100)
        comparativa["Growth %"] = comparativa["Growth %"].replace([float('inf'), -float('inf')], 0)
        
        progress_bar.progress(100)
        logger.info(f"Comparison generated: {len(comparativa)} unique records")
    
    # =========================================================================
    # STEP 5: RESULTS AND METRICS
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">📊 Analysis Results</div>', unsafe_allow_html=True)
    
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
            f"{registros_comunes:,}",
            help="Sales appearing in both periods (recurring customers)"
        )
    with col2:
        st.metric(
            f"🔴 Only in {nombre_periodo_1}", 
            f"{registros_solo_p1:,}",
            help="Sales that did NOT repeat (lost customers)"
        )
    with col3:
        st.metric(
            f"🟢 Only in {nombre_periodo_2}", 
            f"{registros_solo_p2:,}",
            help="New sales (gained customers)"
        )
    with col4:
        st.metric(
            "📊 Total Records", 
            f"{len(comparativa):,}",
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
    
    # =========================================================================
    # COMPARATIVE VISUALIZATIONS
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">📊 Comparative Analysis</div>', unsafe_allow_html=True)

    # Total comparison at top
    st.markdown("### 💰 Overall Comparison")
    col1, col2 = st.columns(2)
    with col1:
        st.metric(f"Total {nombre_periodo_1}", f"€{total_periodo_1:,.2f}")
    with col2:
        st.metric(f"Total {nombre_periodo_2}", f"€{total_periodo_2:,.2f}", delta=f"€{diferencia_total:,.2f}")

    st.markdown("---")

    # REEMPLAZA TODO EL CÓDIGO DESDE AQUÍ HASTA "TOP 10 ANALYSIS" CON ESTO:

    # =========================================================================
    # TABS FOR COMPARATIVE VISUALIZATIONS
    # =========================================================================
    tab1, tab2, tab3, tab4 = st.tabs(["👤 Sales Rep", "🏷️ Product Type", "📦 Set", "📈 Performance Over Time"])

    with tab1:
        st.markdown("#### Comparison by Sales Representative")
        rep_comparison = pd.DataFrame({
            nombre_periodo_1: df1_filtrado.groupby(col_sales_rep)[col_precio].sum(),
            nombre_periodo_2: df2_filtrado.groupby(col_sales_rep)[col_precio].sum()
        }).fillna(0).reset_index()
        rep_comparison = rep_comparison.sort_values(nombre_periodo_2, ascending=True)
        
        fig_rep = go.Figure()
        fig_rep.add_trace(go.Bar(
            name=nombre_periodo_1,
            y=rep_comparison[col_sales_rep],
            x=rep_comparison[nombre_periodo_1],
            orientation='h',
            marker=dict(color='#FF6B6B'),
            text=rep_comparison[nombre_periodo_1].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_rep.add_trace(go.Bar(
            name=nombre_periodo_2,
            y=rep_comparison[col_sales_rep],
            x=rep_comparison[nombre_periodo_2],
            orientation='h',
            marker=dict(color='#4ECDC4'),
            text=rep_comparison[nombre_periodo_2].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_rep.update_layout(
            barmode='group',
            height=max(400, len(rep_comparison) * 40),
            xaxis_title="Sales (€)",
            yaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_rep, use_container_width=True)

    with tab2:
        st.markdown("#### Comparison by Product Type")
        type_comparison = pd.DataFrame({
            nombre_periodo_1: df1_filtrado.groupby(col_tipo)[col_precio].sum(),
            nombre_periodo_2: df2_filtrado.groupby(col_tipo)[col_precio].sum()
        }).fillna(0).reset_index()
        type_comparison = type_comparison.sort_values(nombre_periodo_2, ascending=True)
        
        fig_type = go.Figure()
        fig_type.add_trace(go.Bar(
            name=nombre_periodo_1,
            y=type_comparison[col_tipo],
            x=type_comparison[nombre_periodo_1],
            orientation='h',
            marker=dict(color='#FF6B6B'),
            text=type_comparison[nombre_periodo_1].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_type.add_trace(go.Bar(
            name=nombre_periodo_2,
            y=type_comparison[col_tipo],
            x=type_comparison[nombre_periodo_2],
            orientation='h',
            marker=dict(color='#4ECDC4'),
            text=type_comparison[nombre_periodo_2].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_type.update_layout(
            barmode='group',
            height=max(400, len(type_comparison) * 40),
            xaxis_title="Sales (€)",
            yaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_type, use_container_width=True)

    with tab3:
        st.markdown("#### Comparison by Set")
        set_comparison = pd.DataFrame({
            nombre_periodo_1: df1_filtrado.groupby(col_set)[col_precio].sum(),
            nombre_periodo_2: df2_filtrado.groupby(col_set)[col_precio].sum()
        }).fillna(0).reset_index()
        set_comparison = set_comparison.sort_values(nombre_periodo_2, ascending=True)
        
        fig_set = go.Figure()
        fig_set.add_trace(go.Bar(
            name=nombre_periodo_1,
            y=set_comparison[col_set],
            x=set_comparison[nombre_periodo_1],
            orientation='h',
            marker=dict(color='#FF6B6B'),
            text=set_comparison[nombre_periodo_1].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_set.add_trace(go.Bar(
            name=nombre_periodo_2,
            y=set_comparison[col_set],
            x=set_comparison[nombre_periodo_2],
            orientation='h',
            marker=dict(color='#4ECDC4'),
            text=set_comparison[nombre_periodo_2].apply(lambda x: f'€{x:,.0f}'),
            textposition='auto',
        ))
        fig_set.update_layout(
            barmode='group',
            height=max(400, len(set_comparison) * 50),
            xaxis_title="Sales (€)",
            yaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_set, use_container_width=True)

    with tab4:
        st.markdown("#### Month-by-Month Sales (Jan vs Jan, Feb vs Feb, ...)")
        
        df1m = df1_filtrado.copy()
        df2m = df2_filtrado.copy()

        df1m["_DateDT"] = pd.to_datetime(df1m[col_fecha], errors="coerce")
        df2m["_DateDT"] = pd.to_datetime(df2m[col_fecha], errors="coerce")
        df1m = df1m.dropna(subset=["_DateDT"])
        df2m = df2m.dropna(subset=["_DateDT"])

        df1m["MonthNum"] = df1m["_DateDT"].dt.month
        df2m["MonthNum"] = df2m["_DateDT"].dt.month

        monthly_p1 = df1m.groupby("MonthNum")[col_precio].sum().reset_index()
        monthly_p1.columns = ["MonthNum", nombre_periodo_1]

        monthly_p2 = df2m.groupby("MonthNum")[col_precio].sum().reset_index()
        monthly_p2.columns = ["MonthNum", nombre_periodo_2]

        monthly_combined = pd.merge(monthly_p1, monthly_p2, on="MonthNum", how="outer").fillna(0)
        monthly_combined = monthly_combined.sort_values("MonthNum")

        month_labels_full = {
            1: "January", 2: "February", 3: "March", 4: "April",
            5: "May", 6: "June", 7: "July", 8: "August",
            9: "September", 10: "October", 11: "November", 12: "December"
        }
        monthly_combined["MonthLabel"] = monthly_combined["MonthNum"].map(month_labels_full)

        fig_time = go.Figure()

        fig_time.add_trace(go.Bar(
            name=nombre_periodo_1,
            x=monthly_combined["MonthLabel"],
            y=monthly_combined[nombre_periodo_1],
            marker=dict(color="#FF6B6B"),
            text=monthly_combined[nombre_periodo_1].apply(lambda x: f"€{x:,.0f}" if x > 0 else ""),
            textposition="outside",
        ))

        fig_time.add_trace(go.Bar(
            name=nombre_periodo_2,
            x=monthly_combined["MonthLabel"],
            y=monthly_combined[nombre_periodo_2],
            marker=dict(color="#4ECDC4"),
            text=monthly_combined[nombre_periodo_2].apply(lambda x: f"€{x:,.0f}" if x > 0 else ""),
            textposition="outside",
        ))

        fig_time.update_layout(
            barmode="group",
            height=520,
            xaxis_title="",
            yaxis_title="Sales (€)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=70, r=10, b=10, l=10),
        )

        st.plotly_chart(fig_time, use_container_width=True)

    # =========================================================================
    # BY CUSTOMER (P1 vs P2)  ✅ updates with filters + robust numeric sorting
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">👥 By Customer</div>', unsafe_allow_html=True)

    customer_col = col_cliente          # "Business Partner Name"
    eur_col = col_precio                # "EUR" (ya estandarizado)

    # Fuente CORRECTA: los dataframes filtrados
    df1_src = df1_filtrado.copy()
    df2_src = df2_filtrado.copy()

    # Validación defensiva
    if customer_col not in df1_src.columns or customer_col not in df2_src.columns:
        st.error(f"❌ No encuentro la columna de cliente '{customer_col}' en df1_filtrado/df2_filtrado.")
        st.stop()

    if eur_col not in df1_src.columns or eur_col not in df2_src.columns:
        st.error(f"❌ No encuentro la columna de facturación '{eur_col}' en df1_filtrado/df2_filtrado.")
        st.stop()

    def to_float_money(s: pd.Series) -> pd.Series:
        """
        Convierte valores tipo:
        - '1.234,56' (EU)
        - '1,234.56' (US)
        - '1234,56'
        - '1234.56'
        - '€1.234,56'
        a float.
        """
        x = s.astype(str).str.strip()

        # Limpieza básica
        x = (x.str.replace("€", "", regex=False)
            .str.replace("\u00A0", "", regex=False)  # non-breaking space
            .str.replace(" ", "", regex=False))

        has_comma = x.str.contains(",", regex=False)
        has_dot = x.str.contains(r"\.", regex=True)

        both = has_comma & has_dot
        only_comma = has_comma & ~has_dot

        # Ambos separadores: decide decimal por el separador más a la derecha
        xb = x[both]
        dec_is_comma = xb.str.rfind(",") > xb.str.rfind(".")
        # decimal coma: quita miles '.' y convierte ',' -> '.'
        xb1 = xb[dec_is_comma].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        # decimal punto: quita miles ','
        xb2 = xb[~dec_is_comma].str.replace(",", "", regex=False)

        x.loc[xb1.index] = xb1
        x.loc[xb2.index] = xb2

        # Solo coma: asumimos coma decimal
        x.loc[only_comma] = (x.loc[only_comma]
                            .str.replace(".", "", regex=False)
                            .str.replace(",", ".", regex=False))

        return pd.to_numeric(x, errors="coerce").fillna(0.0)

    # Normaliza cliente + EUR a numérico real
    df1_src[customer_col] = df1_src[customer_col].fillna("(Unknown)").astype(str)
    df2_src[customer_col] = df2_src[customer_col].fillna("(Unknown)").astype(str)

    df1_src[eur_col] = to_float_money(df1_src[eur_col])
    df2_src[eur_col] = to_float_money(df2_src[eur_col])

    # Agregados
    p1_by_customer = df1_src.groupby(customer_col, dropna=False)[eur_col].sum().rename(nombre_periodo_1)
    p2_by_customer = df2_src.groupby(customer_col, dropna=False)[eur_col].sum().rename(nombre_periodo_2)

    by_customer = (
        pd.concat([p1_by_customer, p2_by_customer], axis=1)
        .fillna(0.0)
        .reset_index()
        .rename(columns={customer_col: "Customer"})
    )

    # Deltas (siempre numéricos)
    by_customer["Diff (P2-P1)"] = by_customer[nombre_periodo_2] - by_customer[nombre_periodo_1]
    by_customer["Growth %"] = by_customer.apply(
        lambda r: (r["Diff (P2-P1)"] / r[nombre_periodo_1] * 100)
                if r[nombre_periodo_1] != 0
                else (100.0 if r[nombre_periodo_2] > 0 else 0.0),
        axis=1
    )

    # Orden inicial por P2 desc (numérico real)
    by_customer = by_customer.sort_values(nombre_periodo_2, ascending=False).reset_index(drop=True)

    # Tabla (mantén números; el formateo lo hace column_config)
    st.dataframe(
        by_customer,
        use_container_width=True,
        height=520,
        column_config={
            nombre_periodo_1: st.column_config.NumberColumn(nombre_periodo_1, format="€%.2f"),
            nombre_periodo_2: st.column_config.NumberColumn(nombre_periodo_2, format="€%.2f"),
            "Diff (P2-P1)": st.column_config.NumberColumn("Diff (P2-P1)", format="€%.2f"),
            "Growth %": st.column_config.NumberColumn("Growth %", format="%.1f%%"),
        }
    )

    # Descarga CSV
    csv_bytes = by_customer.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download (CSV)",
        data=csv_bytes,
        file_name=f"by_customer_{nombre_periodo_1}_vs_{nombre_periodo_2}.csv",
        mime="text/csv",
        use_container_width=True
    )



    # =========================================================================
    # TOP 10 ANALYSIS
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">🏆 Top 10 Analysis</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["📊 By Customer", "📦 By Product", "📈 By Growth", "👤 By Sales Rep"])
    
    with tab1:
        st.markdown("#### Top 10 Customers by Sales Volume")
        top_clientes = comparativa.groupby(level=0)[[f"Amount {nombre_periodo_1}", 
                                                       f"Amount {nombre_periodo_2}"]].sum()
        top_clientes['Total'] = top_clientes[f"Amount {nombre_periodo_1}"] + top_clientes[f"Amount {nombre_periodo_2}"]
        top_clientes = top_clientes.sort_values('Total', ascending=False).head(10)
        
        # Format for display
        top_clientes_display = top_clientes.copy()
        for col in top_clientes_display.columns:
            top_clientes_display[col] = top_clientes_display[col].apply(lambda x: f"€{x:,.2f}")
        
        st.dataframe(top_clientes_display, use_container_width=True)
    
    with tab2:
        st.markdown("#### Top 10 Products by Sales Volume")
        top_productos = comparativa.groupby(level=1)[[f"Amount {nombre_periodo_1}", 
                                                        f"Amount {nombre_periodo_2}"]].sum()
        top_productos['Total'] = top_productos[f"Amount {nombre_periodo_1}"] + top_productos[f"Amount {nombre_periodo_2}"]
        top_productos = top_productos.sort_values('Total', ascending=False).head(10)
        
        # Format for display
        top_productos_display = top_productos.copy()
        for col in top_productos_display.columns:
            top_productos_display[col] = top_productos_display[col].apply(lambda x: f"€{x:,.2f}")
        
        st.dataframe(top_productos_display, use_container_width=True)
    
    with tab3:
        st.markdown("#### Top 10 by Highest Growth")
        crecimiento = comparativa.copy()
        crecimiento = crecimiento[crecimiento[f"Amount {nombre_periodo_1}"] > 0]  # Only items that existed in P1
        top_crecimiento = crecimiento.nlargest(10, 'Amount Difference').reset_index()
        
        # Select relevant columns
        display_cols = [col_cliente, col_producto, f"Amount {nombre_periodo_1}", 
                       f"Amount {nombre_periodo_2}", 'Amount Difference', 'Growth %']
        top_crecimiento_display = top_crecimiento[display_cols].copy()
        
        # Format monetary columns
        for col in [f"Amount {nombre_periodo_1}", f"Amount {nombre_periodo_2}", 'Amount Difference']:
            top_crecimiento_display[col] = top_crecimiento_display[col].apply(lambda x: f"€{x:,.2f}")
        top_crecimiento_display['Growth %'] = top_crecimiento_display['Growth %'].apply(lambda x: f"{x:.1f}%")
        
        st.dataframe(top_crecimiento_display, use_container_width=True)
    
    with tab4:
        st.markdown("#### Performance by Sales Representative")
        rep_analysis = df1_filtrado.groupby(col_sales_rep)[col_precio].sum().to_frame(nombre_periodo_1)
        rep_analysis[nombre_periodo_2] = df2_filtrado.groupby(col_sales_rep)[col_precio].sum()
        rep_analysis = rep_analysis.fillna(0)
        rep_analysis['Difference'] = rep_analysis[nombre_periodo_2] - rep_analysis[nombre_periodo_1]
        rep_analysis['Growth %'] = ((rep_analysis[nombre_periodo_2] - rep_analysis[nombre_periodo_1]) / 
                                    rep_analysis[nombre_periodo_1] * 100)
        rep_analysis = rep_analysis.replace([float('inf'), -float('inf')], 0)
        rep_analysis = rep_analysis.sort_values('Difference', ascending=False)
        
        # Format for display
        rep_analysis_display = rep_analysis.copy()
        for col in [nombre_periodo_1, nombre_periodo_2, 'Difference']:
            rep_analysis_display[col] = rep_analysis_display[col].apply(lambda x: f"€{x:,.2f}")
        rep_analysis_display['Growth %'] = rep_analysis_display['Growth %'].apply(lambda x: f"{x:.1f}%")
        
        st.dataframe(rep_analysis_display, use_container_width=True)
    
    # Preview of comparison
    st.markdown("---")
    st.markdown("### 👀 Full Comparison Preview (first 20 rows)")
    preview_df = comparativa.reset_index().head(20).copy()
    
    # Format monetary columns for preview
    for col in preview_df.columns:
        if 'Amount' in col or 'Difference' in col:
            if col != 'Growth %':
                preview_df[col] = preview_df[col].apply(lambda x: f"€{x:,.2f}" if pd.notnull(x) else "€0.00")
    
    if 'Growth %' in preview_df.columns:
        preview_df['Growth %'] = preview_df['Growth %'].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else "0.0%")
    
    st.dataframe(
        preview_df, 
        use_container_width=True,
        height=400
    )
    



    # =========================================================================
    # STEP 6: DOWNLOAD HTML DASHBOARD
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="section-header">📥 Download Interactive Dashboard</div>', unsafe_allow_html=True)

    # Prepare configuration for reproducibility
    config_info = {
        'Source File': uploaded_file.name,
        'Analysis Date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'Period 1 Name': nombre_periodo_1,
        'Period 1 Start': str(start_date_1),
        'Period 1 End': str(end_date_1),
        'Period 2 Name': nombre_periodo_2,
        'Period 2 Start': str(start_date_2),
        'Period 2 End': str(end_date_2),
        'Product Types': ', '.join(tipos_seleccionados),
        'Product Name Filter': search_text if search_text else 'None',
        'Sales Rep Filter': ', '.join(selected_reps) if 'selected_reps' in locals() and selected_reps else 'All',
        'Total Records P1': len(df1_filtrado),
        'Total Records P2': len(df2_filtrado),
        'Total Sales P1': f"€{total_periodo_1:,.2f}",
        'Total Sales P2': f"€{total_periodo_2:,.2f}",
        'Difference': f"€{diferencia_total:,.2f}",
        'Growth %': f"{porcentaje_cambio:.1f}%"
    }

    # =========================================================================
    # HTML DASHBOARD DOWNLOAD
    # =========================================================================
    st.markdown("### 🌐 Interactive HTML Dashboard")

    st.info("💡 **Generate a self-contained HTML file** with all data, filters, charts, and interactive tables. Perfect for sharing with colleagues or presenting in meetings.")

    if st.button("🚀 GENERATE HTML DASHBOARD", type="primary", use_container_width=True):
        with st.spinner("🔄 Generating interactive HTML dashboard..."):
            # Prepare additional filter lists for new columns
            available_market_orgs = []
            available_territories = []
            available_countries = []
            available_segments = []
            
            if 'Market Organization Name' in df.columns:
                available_market_orgs = sorted(df['Market Organization Name'].dropna().unique().tolist())
            
            if 'Sales Territory' in df.columns:
                available_territories = sorted(df['Sales Territory'].dropna().unique().tolist())
            
            if 'Country' in df.columns:
                available_countries = sorted(df['Country'].dropna().unique().tolist())
            
            if 'End User Segment' in df.columns:
                available_segments = sorted(df['End User Segment'].dropna().unique().tolist())
            
            # ✅ Check if SFDC Link column exists in filtered data
            has_sfdc_links = 'SFDC Link' in df1_filtrado.columns and df1_filtrado['SFDC Link'].notna().sum() > 0

            html_content = generate_sales_comparison_html(
                df1_filtrado=df1_filtrado,
                df2_filtrado=df2_filtrado,
                comparativa=comparativa,
                config_info=config_info,
                nombre_periodo_1=nombre_periodo_1,
                nombre_periodo_2=nombre_periodo_2,
                available_types=available_types,
                available_sets=available_sets,
                available_reps=available_reps,
                available_market_orgs=available_market_orgs,
                available_territories=available_territories,
                available_countries=available_countries,
                available_segments=available_segments,
                has_sfdc_links=has_sfdc_links  # ✅ NUEVO PARÁMETRO
            )
            
            fecha_actual = datetime.now().strftime("%Y%m%d")
            nombre_archivo_html = f"comparison_{nombre_periodo_1}_vs_{nombre_periodo_2}_{fecha_actual}.html"
            
            st.download_button(
                label="📥 DOWNLOAD HTML DASHBOARD",
                data=html_content,
                file_name=nombre_archivo_html,
                mime="text/html",
                use_container_width=True,
                key="download_html"
            )
            
            st.success("✅ **HTML dashboard generated successfully!**")

    # =========================================================================
    # HTML INFO BOX
    # =========================================================================
    st.markdown("---")

    with st.expander("ℹ️ About the HTML Dashboard", expanded=False):
        st.markdown("""
        ### 🌐 Interactive HTML Dashboard Features:
        
        **📊 Self-contained & Portable:**
        - Single HTML file with all data embedded
        - No internet connection required after download
        - Share with colleagues who don't have Python/Streamlit
        - Works on any device with a web browser
        
        **🔗 SFDC Integration:**
        - Customer names are clickable links to Salesforce (when available)
        - Direct access to account details with one click
        - Opens in new tab for seamless workflow
        
        **🎛️ Interactive Filters:**
        - Product Type, Set, Sales Representative filters
        - Quick filter tags (CARE, Exact, Start, etc.) with AND/OR mode
        - Customer search
        - Market Organization, Territory, Country, Segment filters
        - All/None buttons for each filter group
        - Active filters display with chips
        
        **📈 Dynamic Visualizations:**
        - Comparison by Sales Representative (grouped bars)
        - Comparison by Product Type (grouped bars)
        - Comparison by Set (grouped bars)
        - Performance Over Time - Month-by-Month (grouped bars)
        - All charts update dynamically with filters
        
        **📋 Multiple Data Views (Tabs):**
        - **Full Comparison:** Complete comparison table with differences
        - **By Customer:** Aggregated customer-level analysis
        - **Only in P1:** Records that didn't repeat (lost customers)
        - **Only in P2:** New records (gained customers)
        - **Common Records:** Recurring sales (loyal customers)
        
        **📊 Real-time Metrics:**
        - Total sales for each period
        - Growth percentage
        - Customer counts (retained, lost, new)
        - Record counts (common, only P1, only P2)
        - Retention and acquisition rates
        - Expansion and contraction analysis
        
        **🔍 Sortable Tables:**
        - Click column headers to sort
        - Visual sort indicators (↑↓)
        - CSV export buttons for each table
        
        **💾 State Persistence:**
        - Filter selections remembered during session
        - Smooth, responsive interface
        
        ### 💡 Best Use Cases:
        - 📧 Email to management for review
        - 📱 Open on mobile devices
        - 💼 Present in meetings without internet
        - 🔄 Archive for future reference
        - 👥 Share with non-technical stakeholders
        - 🔗 Quick access to Salesforce records
        """)

    st.success(f"🎉 **Analysis completed!** Generate the HTML dashboard above to share your results.")

    logger.info(f"Analysis completed successfully.")

# Footer
st.markdown("---")
st.markdown("""
<div class='footer'>
    <p><strong>📊 Sales Comparison Tool v2.0</strong> | BUCHI Analytics Suite</p>
    <p>Enhanced with Advanced Analytics: Interactive visualizations · Top 10 rankings · Retention KPIs · Multi-format export</p>
</div>
""", unsafe_allow_html=True)