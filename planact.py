# streamlit_app.py

import streamlit as st
import pandas as pd
import os
import datetime
from io import BytesIO

# ------------------------------
# 1. Page Configuration
# ------------------------------
st.set_page_config(
    page_title="📊 Plan vs Actuals Report",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------
# 2. Custom CSS Styling
# ------------------------------
custom_css = """
<style>
    /* Overall background */
    .stApp {
        background-color: #ffffff;
    }

    /* Sidebar styling */
    .sidebar .sidebar-content {
        background-color: #2c3e50;
        color: white;
    }

    .sidebar .sidebar-content a {
        color: #ecf0f1;
    }

    /* Header styling */
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #34495e;
        text-align: center;
        margin-top: 20px;
        margin-bottom: 10px;
    }

    /* Subheader styling */
    .main-subheader {
        font-size: 1.2rem;
        color: #7f8c8d;
        text-align: center;
        margin-bottom: 40px;
    }

    /* Button styling */
    .stButton > button {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 12px 28px;
        border-radius: 8px;
        font-size: 16px;
        cursor: pointer;
        transition: background-color 0.3s ease;
    }

    .stButton > button:hover {
        background-color: #2980b9;
    }

    /* DataFrame styling */
    .dataframe thead tr th {
        background-color: #34495e;
        color: white;
        font-weight: bold;
    }

    .dataframe tbody tr:nth-child(even) {
        background-color: #f2f2f2;
    }

    .dataframe tbody tr:hover {
        background-color: #e0f7fa;
    }

    /* Footer styling */
    .footer {
        font-size: 0.9rem;
        color: #95a5a6;
        text-align: center;
        margin-top: 50px;
    }

    /* Download button styling */
    .stDownloadButton button {
        background-color: #2ecc71;
        color: white;
        border: none;
        padding: 12px 28px;
        border-radius: 8px;
        font-size: 16px;
        cursor: pointer;
        transition: background-color 0.3s ease;
    }

    .stDownloadButton button:hover {
        background-color: #27ae60;
    }

    /* Spinner styling */
    .element-container .stSpinner {
        border-top-color: #3498db;
    }
</style>
"""

# Inject custom CSS
st.markdown(custom_css, unsafe_allow_html=True)

# ------------------------------
# 3. Helper Functions
# ------------------------------
def to_excel(df):
    """Convert DataFrame to Excel format in memory."""
    # Flatten MultiIndex columns if present
    if isinstance(df.columns, pd.MultiIndex):
        df = df.copy()
        df.columns = ['_'.join([str(c) for c in col]).strip() for col in df.columns.values]
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Define processing functions

def read_shopfloor_data(uploaded_file):
    """Read and preprocess shopfloor CSV data."""
    df_shopfloor = pd.read_csv(uploaded_file)
    # Convert 'Date' to datetime and normalize
    df_shopfloor['Date'] = pd.to_datetime(df_shopfloor['Date'], errors='coerce').dt.normalize()
    # Drop rows with invalid dates
    df_shopfloor = df_shopfloor.dropna(subset=['Date'])
    # Convert 'Module' to integer and format with leading zeros
    df_shopfloor['Module_Upd'] = "BAI III Team " + df_shopfloor['Module'].astype(int).astype(str).str.zfill(2)
    # Group by Schedule, Date, Module_Upd and sum 'Sewingout[130]-Good'
    df1_sfcs = df_shopfloor.groupby(['Schedule','Date','Module_Upd']).agg({
        'Sewingout[130]-Good' : 'sum',
    }).reset_index()
    df1_sfcs.rename(columns={'Sewingout[130]-Good': 'Actuals'}, inplace=True)
    return df1_sfcs

def read_order_book(uploaded_file):
    """Read and preprocess order book Excel data."""
    df1 = pd.read_excel(uploaded_file)
    # Create 'Sample Code' by extracting substring from 'Cust Style No'
    df1['Sample Code'] = df1['Cust Style No'].str[2:-4]
    
    # Calculate additional metrics
    df1['Cut Balance'] = df1['Cum Cut Qty'] - df1['CO Qty']
    df1['Sew_Good'] = df1['Cum SewOut Qty'] - df1['Cum Sew Out Rej Qty']
    df1['Cut %'] = (df1['Cum Cut Qty'] / df1['CO Qty']) * 100
    df1['Sew%'] = (df1['Cum SewOut Qty'] / df1['CO Qty']) * 100
    df1['IMS'] = df1['Cum Sew In Qty'] - df1['Cum SewOut Qty']
    df1['Rej%'] = (df1['Cum Sew Out Rej Qty'] / df1['Cum SewOut Qty']) * 100
    df1['Bal_to_Ship'] = df1['Delivered Qty'] - df1['CO Qty']
    df1['Del%'] = (df1['Delivered Qty'] / df1['CO Qty']) * 100
    df1['Bal_to_sew%'] = (df1['Sew_Good'] / df1['CO Qty']) * 100
    df1['Bal_to_sew'] = df1['Sew_Good'] - df1['CO Qty']
    
    # Process 'Schedule No'
    df1['Schedule No'] = pd.to_numeric(df1['Schedule No'], errors='coerce').fillna(0).astype(int).astype(str)
    
    # Group by relevant columns
    df1_group = df1.groupby(['Schedule No','VPO No','Sample Code','Group Tech Class']).agg({
        'CO Qty': 'sum',
        'Sew_Good': 'sum',
        'PED': 'max'
    }).reset_index()
    
    return df1_group

def create_order_book_pivot(ob_group_df):
    """Create pivot table from order book grouped data."""
    pivot_table = ob_group_df.pivot_table(
        index=['Schedule No', 'VPO No', 'Sample Code','Group Tech Class','PED'], 
        values=['CO Qty', 'Sew_Good'], 
        aggfunc='sum'
    ).reset_index()  # Reset the index to make 'Schedule No' a column

    # Check for duplicates
    pivot_duplicates = pivot_table.duplicated(subset=['Schedule No', 'VPO No', 'Sample Code','Group Tech Class','PED']).any()
    return pivot_table, pivot_duplicates

def process_product_mapping(uploaded_file):
    """Process product mapping data."""
    df_mapping = pd.read_excel(uploaded_file)
    df_mapping = df_mapping.drop(columns=['Sub Item', 'IND Only'], errors='ignore')
    df_mapping.rename(columns={'Master Item': 'Product'}, inplace=True)
    return df_mapping

def merge_order_book_with_mapping(ob_group_df, mapping_df):
    """Merge order book grouped data with product mapping."""
    mapping_ob = ob_group_df.merge(mapping_df, left_on='Sample Code', right_on='Style', how='left')
    return mapping_ob

def process_loading_plan(uploaded_file):
    """Process loading plan Excel data."""
    data = pd.read_excel(uploaded_file)
    
    # Convert 'Schedule No' to string and strip spaces
    data['Schedule No'] = data['Schedule No'].astype(str).str.strip()
    
    # Drop rows with NaN in 'Schedule No'
    data_cleaned = data.dropna(subset=['Schedule No'])
    
    # Drop 'Unnamed: 15' if exists
    if 'Unnamed: 15' in data_cleaned.columns:
        data_cleaned = data_cleaned.drop(columns=['Unnamed: 15'])
    
    # Identify date columns starting from the 16th column (zero-based indexing)
    date_columns = data_cleaned.columns[15:]
    date_columns = [col for col in date_columns if 'Unnamed' not in str(col)]
    
    # Convert date columns to string format
    date_columns_str = []
    for col in date_columns:
        if isinstance(col, datetime.datetime):
            date_columns_str.append(col.strftime('%Y-%m-%d'))
        else:
            # Try to parse the column as a date
            try:
                parsed_date = pd.to_datetime(col, errors='coerce')
                if pd.notnull(parsed_date):
                    date_columns_str.append(parsed_date.strftime('%Y-%m-%d'))
                else:
                    date_columns_str.append(str(col))
            except:
                date_columns_str.append(str(col))
    
    # Rename columns
    data_cleaned.columns = list(data_cleaned.columns[:15]) + date_columns_str
    
    # Melt the DataFrame
    melted_data = data_cleaned.melt(
        id_vars=['Schedule No'], 
        value_vars=date_columns_str, 
        var_name='Date', 
        value_name='Quantity'
    )
    
    # Convert 'Quantity' to numeric
    melted_data['Quantity'] = pd.to_numeric(melted_data['Quantity'], errors='coerce')
    
    # Drop rows with NaN or empty 'Schedule No'
    melted_data_cleaned = melted_data.dropna(subset=['Schedule No'])
    melted_data_cleaned = melted_data_cleaned[melted_data_cleaned['Schedule No'].str.strip().str.lower() != 'nan']
    
    # Ensure 'Schedule No' is string and strip spaces
    melted_data_cleaned['Schedule No'] = melted_data_cleaned['Schedule No'].astype(str).str.strip()
    
    # Group by 'Date' and 'Schedule No'
    grouped_schedule_qty = melted_data_cleaned.groupby(['Date', 'Schedule No']).sum().reset_index()
    
    # Convert 'Schedule No' to integer and then to string to remove decimal
    grouped_schedule_qty['Schedule No'] = grouped_schedule_qty['Schedule No'].astype(float).astype(int).astype(str)
    
    return grouped_schedule_qty

def merge_plan_vs_actuals(loading_plan_df, sfcs_df):
    """Merge loading plan with shopfloor actuals."""
    # Ensure 'Date' columns are datetime
    loading_plan_df['Date'] = pd.to_datetime(loading_plan_df['Date'], errors='coerce')
    sfcs_df['Date'] = pd.to_datetime(sfcs_df['Date'], errors='coerce')
    
    # Rename 'Schedule' to 'Schedule No' in sfcs_df
    sfcs_df = sfcs_df.rename(columns={'Schedule': 'Schedule No'})
    
    # Ensure 'Schedule No' is string and strip spaces
    loading_plan_df['Schedule No'] = loading_plan_df['Schedule No'].astype(str).str.strip()
    sfcs_df['Schedule No'] = sfcs_df['Schedule No'].astype(str).str.strip()
    
    # Merge on ['Schedule No', 'Date']
    merged_df = pd.merge(loading_plan_df, sfcs_df, on=['Schedule No', 'Date'], how='outer')
    
    # Create pivot tables
    pivot_table = merged_df.pivot_table(
        values=['Quantity', 'Actuals'], 
        index=['Schedule No'],
        columns='Date', 
        aggfunc='sum', 
        fill_value=0
    )
    
    # Reorder pivot to have 'Quantity' first then 'Actuals'
    pivot_table = pivot_table.reindex(['Quantity', 'Actuals'], level=0, axis=1)
    
    # Convert all values to integers
    pivot_table = pivot_table.fillna(0).astype(int)
    
    return merged_df, pivot_table

def merge_with_signoff(signoff_file, sfcs_df):
    """Merge signoff data with shopfloor actuals."""
    # Load signoff data
    data_signoff = pd.read_excel(signoff_file)
    
    # Remove 'Unnamed' columns from sfcs_df
    sfcs_df = sfcs_df.loc[:, ~sfcs_df.columns.str.contains('^Unnamed')]
    
    # Ensure 'Date' columns are datetime
    data_signoff['Date'] = pd.to_datetime(data_signoff['Date'], errors='coerce')
    sfcs_df['Date'] = pd.to_datetime(sfcs_df['Date'], errors='coerce')
    
    # Rename 'Schedule' to 'Schedule No' in sfcs_df
    sfcs_df = sfcs_df.rename(columns={'Schedule': 'Schedule No'})
    
    # Ensure 'Schedule No' is string and strip spaces
    data_signoff['Schedule No'] = data_signoff['Schedule No'].astype(str).str.strip()
    sfcs_df['Schedule No'] = sfcs_df['Schedule No'].astype(str).str.strip()
    
    # Drop 'Module_Upd' if exists
    if 'Module_Upd' in sfcs_df.columns:
        sfcs_df = sfcs_df.drop(columns=['Module_Upd'])
    
    # Group by ['Schedule No', 'Date'] and sum
    data_signoff_grouped = data_signoff.groupby(['Schedule No', 'Date']).sum().reset_index()
    data_sfcs_grouped = sfcs_df.groupby(['Schedule No', 'Date']).sum().reset_index()
    
    # Merge signoff with sfcs
    merged_df1 = pd.merge(data_signoff_grouped, data_sfcs_grouped, on=['Schedule No', 'Date'], how='outer')
    
    # Identify and count rows with NaN in 'Actuals'
    missing_actuals = merged_df1[merged_df1['Actuals'].isna()]
    missing_actuals_count = missing_actuals.shape[0]
    st.warning(f"**Number of rows with NaN values in 'Actuals': {missing_actuals_count}**")
    
    # Fill NaN in 'Actuals' with 0
    merged_df1['Actuals'] = merged_df1['Actuals'].fillna(0)
    
    return merged_df1

def merge_vpolevel(merger_df, ob_pivot_df):
    """Merge merger data with order book pivot data."""
    # Ensure 'Schedule No' is string and uppercase
    merger_df['Schedule No'] = merger_df['Schedule No'].astype(str).str.strip().str.upper()
    ob_pivot_df['Schedule No'] = ob_pivot_df['Schedule No'].fillna(0).astype(float).astype(int).astype(str).str.strip().str.upper()
    
    # Merge on 'Schedule No'
    merged_data = pd.merge(merger_df, ob_pivot_df, on='Schedule No', how='left')
    
    # Remove 'Unnamed' columns
    merged_data_cleaned = merged_data.loc[:, ~merged_data.columns.str.contains('^Unnamed')]
    
    return merged_data_cleaned

def merge_with_product_mapping(vpolevel_df, product_mapping_df):
    """Merge VPO level data with product mapping."""
    # Ensure 'Sample Code' and 'Style' are strings
    vpolevel_df['Sample Code'] = vpolevel_df['Sample Code'].astype(str)
    product_mapping_df['Style'] = product_mapping_df['Style'].astype(str)
    
    # Merge on 'Sample Code' and 'Style'
    merged_data = pd.merge(vpolevel_df, product_mapping_df, left_on='Sample Code', right_on='Style', how='left')
    
    # Drop 'Unnamed: 0' if exists
    if 'Unnamed: 0' in merged_data.columns:
        merged_data_cleaned = merged_data.drop(columns=['Unnamed: 0'])
    else:
        merged_data_cleaned = merged_data
    
    return merged_data_cleaned

# ------------------------------
# 4. Sidebar: File Uploads
# ------------------------------
st.sidebar.markdown('<h2 style="color: white;">📁 Upload Input Files</h2>', unsafe_allow_html=True)
st.sidebar.markdown("<br>", unsafe_allow_html=True)

# File uploaders
uploaded_shopfloor = st.sidebar.file_uploader(
    "Upload Shopfloor CSV (`sfcs_09.19.csv`)", type=["csv"]
)

uploaded_order_book = st.sidebar.file_uploader(
    "Upload Order Book Excel (`Copy of 18th Sep Order Book.xlsx`)", type=["xlsx", "xls"]
)

uploaded_product_mapping = st.sidebar.file_uploader(
    "Upload Product Mapping Excel (`Style Product Mapping Summary.xlsx`)", type=["xlsx", "xls"]
)

uploaded_loading_plan = st.sidebar.file_uploader(
    "Upload Loading Plan Excel (`BAI3_Loading Plan -Sept 24.xlsx`)", type=["xlsx", "xls"]
)

uploaded_signoff = st.sidebar.file_uploader(
    "Upload Signoff Excel (`Sign off August.xlsx`)", type=["xlsx", "xls"]
)

# Add spacing before the button
st.sidebar.markdown("<br><br>", unsafe_allow_html=True)

# ------------------------------
# 5. Main Page: Header
# ------------------------------
st.markdown('<h1 class="main-header">📊 Plan vs Actuals</h1>', unsafe_allow_html=True)
st.markdown('<p class="main-subheader">Analyzing and comparing planned production against actual performance metrics.</p>', unsafe_allow_html=True)

# ------------------------------
# 6. Processing and Display
# ------------------------------
if st.sidebar.button("Generate Report"):
    # Ensure all required files are uploaded
    required_files = {
        "Shopfloor CSV": uploaded_shopfloor,
        "Order Book Excel": uploaded_order_book,
        "Product Mapping Excel": uploaded_product_mapping,
        "Loading Plan Excel": uploaded_loading_plan,
        "Signoff Excel": uploaded_signoff
    }

    missing_files = [name for name, file in required_files.items() if file is None]

    if missing_files:
        st.error(f"⚠️ Please upload the following required files: {', '.join(missing_files)}")
    else:
        try:
            # Step 1: Process Shopfloor Data
            with st.spinner("📥 Processing Shopfloor Data..."):
                df_shopfloor = read_shopfloor_data(uploaded_shopfloor)
                st.session_state.df_shopfloor = df_shopfloor

            # Step 2: Process Order Book
            with st.spinner("📥 Processing Order Book Data..."):
                df_order_book_grouped = read_order_book(uploaded_order_book)
                st.session_state.df_order_book_grouped = df_order_book_grouped

            # Step 3: Create Order Book Pivot Table
            with st.spinner("📊 Creating Order Book Pivot Table..."):
                pivot_ob, pivot_duplicates = create_order_book_pivot(df_order_book_grouped)
                st.session_state.pivot_ob = pivot_ob

            # Step 4: Process Product Mapping
            with st.spinner("📥 Processing Product Mapping Data..."):
                df_product_mapping = process_product_mapping(uploaded_product_mapping)
                st.session_state.df_product_mapping = df_product_mapping

            # Step 5: Merge Order Book with Product Mapping
            with st.spinner("🔗 Merging Order Book with Product Mapping..."):
                mapping_ob = merge_order_book_with_mapping(df_order_book_grouped, df_product_mapping)
                st.session_state.mapping_ob = mapping_ob

            # Step 6: Process Loading Plan
            with st.spinner("📥 Processing Loading Plan Data..."):
                grouped_loading_plan = process_loading_plan(uploaded_loading_plan)
                st.session_state.grouped_loading_plan = grouped_loading_plan

            # Step 7: Merge Plan vs Actuals
            with st.spinner("🔀 Merging Plan vs Actuals..."):
                merged_df, pivot_plan_vs_actuals = merge_plan_vs_actuals(grouped_loading_plan, df_shopfloor)
                st.session_state.merged_df = merged_df
                st.session_state.pivot_plan_vs_actuals = pivot_plan_vs_actuals

            # Step 8: Merge with Signoff Data
            with st.spinner("🔗 Merging with Signoff Data..."):
                merged_signoff_sfcs = merge_with_signoff(uploaded_signoff, df_shopfloor)
                st.session_state.merged_signoff_sfcs = merged_signoff_sfcs

            # Step 9: Merge with Order Book Pivot
            with st.spinner("🔗 Merging with Order Book Pivot Data..."):
                vpolevel_data = merge_vpolevel(merged_df, pivot_ob)
                st.session_state.vpolevel_data = vpolevel_data

            # Step 10: Merge VPO Level with Product Mapping
            with st.spinner("🔗 Merging VPO Level Data with Product Mapping..."):
                uq_plan_vs_actuals = merge_with_product_mapping(vpolevel_data, df_product_mapping)
                st.session_state.uq_plan_vs_actuals = uq_plan_vs_actuals

            # Final DataFrame
            final_df = st.session_state.uq_plan_vs_actuals.copy()

            # Display the final DataFrame
            st.markdown("### 📈 UQ Plan vs Actuals Data Preview")
            st.dataframe(final_df.head())

            # Provide download link for the final report
            st.markdown("### 🔽 Download Final Report")
            excel_final = to_excel(final_df)
            st.download_button(
                label="📥 Download UQ Plan vs Actuals as Excel",
                data=excel_final,
                file_name='uq_plan_vs_actuals.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.success("✅ Final report is ready for download.")

        except Exception as e:
            st.error(f"❌ An error occurred during processing: {e}")

# ------------------------------
# 7. Footer
# ------------------------------
st.markdown("""
---
<div class="footer">
    &copy; 2024 Developed by <a href="https://your-website.com" style="color: #2c3e50; text-decoration: none;">Your Name</a>
</div>
""", unsafe_allow_html=True)
