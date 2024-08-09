import pandas as pd
import streamlit as st
from PIL import Image

# Load the logo image with a transparent background
logo = r"C:\Users\srikanthve\OneDrive - Brandix Lanka Pvt Ltd\Desktop\logobrandix.jpg"

# Set up the sidebar with the logo on top (no resizing)
st.sidebar.image(logo, use_column_width=False)

# Custom CSS for a more professional title style
st.markdown(
    """
    <style>
    .title {
        font-family: 'Helvetica Neue', sans-serif;
        font-size: 40px;
        font-weight: bold;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True
)

# Display the title with the custom style
st.markdown('<div class="title">Plan vs Actuals</div>', unsafe_allow_html=True)

# Step 1: Read and Process Shopfloor Data
def process_shopfloor_data(df_shopfloor):
    df_shopfloor['Date'] = pd.to_datetime(df_shopfloor['Date']).dt.normalize()
    df_shopfloor['Module_Upd'] = "BAI III Team " + df_shopfloor['Module'].astype(int).astype(str).str.zfill(2)
    df1_sfcs = df_shopfloor.groupby(['Schedule', 'Date', 'Module_Upd']).agg({'Sewingout[130]-Good': 'sum'}).reset_index()
    df1_sfcs.rename(columns={'Sewingout[130]-Good': 'Actuals'}, inplace=True)
    return df1_sfcs

# Step 2: Read and Process Order Book Data
def process_order_book_data(df1):
    if 'Cust Style No' in df1.columns:
        df1['Sample Code'] = df1['Cust Style No'].str[2:-4]
    else:
        st.error("The column 'Cust Style No' does not exist in the uploaded file.")
        return None
    
    df1['Cut Balance'] = df1['Cum Cut Qty'] - df1['CO Qty']
    df1['Sew_Good'] = df1['Cum SewOut Qty'] - df1['Cum Sew Out Rej Qty']
    df1['Cut %'] = df1['Cum Cut Qty'] / df1['CO Qty'] * 100
    df1['Sew%'] = df1['Cum SewOut Qty'] / df1['CO Qty'] * 100
    df1['IMS'] = df1['Cum Sew In Qty'] - df1['Cum SewOut Qty']
    df1['Rej%'] = df1['Cum Sew Out Rej Qty'] / df1['Cum SewOut Qty'] * 100
    df1['Bal_to_Ship'] = df1['Delivered Qty'] - df1['CO Qty']
    df1['Del%'] = df1['Delivered Qty'] / df1['CO Qty'] * 100
    df1['Bal_to_sew%'] = df1['Sew_Good'] / df1['CO Qty'] * 100
    df1['Bal_to_sew'] = df1['Sew_Good'] - df1['CO Qty']
    
    df1['Schedule No'] = pd.to_numeric(df1['Schedule No'], errors='coerce')
    df1['Schedule No'] = df1['Schedule No'].fillna(0).astype(int).astype(str)
    
    df1_group = df1.groupby(['Schedule No', 'VPO No', 'Sample Code', 'Group Tech Class']).agg({
        'CO Qty': 'sum',
        'Sew_Good': 'sum',
        'PED': 'max'
    }).reset_index()
    
    return df1_group

# Step 3: Merge and Clean Additional Data
def merge_additional_data(vpolevel_path, productmapping_path):
    # Load the pre-defined files
    vpolevel_data1 = pd.read_excel(vpolevel_path)
    productmapping_data1 = pd.read_excel(productmapping_path)
    
    # Convert 'Sample Code' and 'Style' to string type in both dataframes
    vpolevel_data1['Sample Code'] = vpolevel_data1['Sample Code'].astype(str)
    productmapping_data1['Style'] = productmapping_data1['Style'].astype(str)
    
    # Merge the two dataframes on 'Sample Code' from vpolevel and 'Style' from productmapping
    merged_data1 = pd.merge(vpolevel_data1, productmapping_data1, left_on='Sample Code', right_on='Style', how='left')
    
    # Remove the "Unnamed: 0" column that came from productmapping file if it exists
    if 'Unnamed: 0' in merged_data1.columns:
        merged_data_cleaned1 = merged_data1.drop(columns=['Unnamed: 0'])
    else:
        merged_data_cleaned1 = merged_data1
    
    return merged_data_cleaned1

# Main Application

# Upload Shopfloor Data
uploaded_shopfloor_file = st.file_uploader("Choose a Shopfloor data file", type=["csv", "xlsx"])
uploaded_order_book_file = st.file_uploader("Choose an Order Book data file", type=["csv", "xlsx"])

if uploaded_shopfloor_file and uploaded_order_book_file:
    if uploaded_shopfloor_file.name.endswith('.csv'):
        df_shopfloor = pd.read_csv(uploaded_shopfloor_file)
    elif uploaded_shopfloor_file.name.endswith('.xlsx'):
        df_shopfloor = pd.read_excel(uploaded_shopfloor_file)
    
    processed_shopfloor_data = process_shopfloor_data(df_shopfloor)
    
    if uploaded_order_book_file.name.endswith('.csv'):
        df_order_book = pd.read_csv(uploaded_order_book_file)
    elif uploaded_order_book_file.name.endswith('.xlsx'):
        df_order_book = pd.read_excel(uploaded_order_book_file)
    
    processed_order_book_data = process_order_book_data(df_order_book)
    
    if processed_order_book_data is not None:
        # Define the paths to the preloaded files
        vpolevel_path = r'C:\India Plants\Marketing\Ajay Report\vpolevel1.xlsx'
        productmapping_path = r'C:\India Plants\Marketing\Ajay Report\productmapping.xlsx'
        
        merged_data_cleaned1 = merge_additional_data(vpolevel_path, productmapping_path)
        
        # Save the cleaned merged dataframe to an Excel file
        merged_data_cleaned1_output_path = 'uq_planvsactuals1.xlsx'
        merged_data_cleaned1.to_excel(merged_data_cleaned1_output_path, index=False)
        
        st.download_button(
            label="Download Merged Data",
            data=open(merged_data_cleaned1_output_path, 'rb').read(),
            file_name="uq_planvsactuals1.xlsx"
        )
