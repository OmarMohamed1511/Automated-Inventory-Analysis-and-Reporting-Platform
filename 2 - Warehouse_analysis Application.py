import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import base64
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# Set page config
st.set_page_config(page_title="Warehouse Stock Analysis", layout="wide")

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'df2' not in st.session_state:
    st.session_state.df2 = None
if 'processed' not in st.session_state:
    st.session_state.processed = False

# Title
st.title("Warehouse Stock Analysis Dashboard")

# File upload section
col1, col2 = st.columns(2)
with col1:
    stock_source_file = st.file_uploader("Upload Stock Source File", type=['xlsx'])
with col2:
    fabric_stock_file = st.file_uploader("Upload Fabric Stock File", type=['xlsx'])

# Process data when files are uploaded
if stock_source_file and fabric_stock_file and not st.session_state.processed:
    with st.spinner("Processing data..."):
        # Read files
        df = pd.read_excel(stock_source_file, engine='openpyxl', sheet_name='Sheet')
        df2 = pd.read_excel(fabric_stock_file, engine='openpyxl', sheet_name='Sheet')
        
        # Extract current date from fabric stock filename
        fabric_filename = fabric_stock_file.name
        try:
            date_part = fabric_filename.split('stock')[1].split('.')[0].strip()
            current_date = date_part.replace(' ', '-')
        except:
            current_date = datetime.now().strftime('%d-%m-%Y')
        
        # Aggregate by Warehouse
        df_group1 = df.groupby('Warehouse')['Quantity'].sum().reset_index()
        df2["one"] = 1
        df_group2 = df2.groupby('Ware House')['one'].sum().reset_index()
        
        # Add PF_Active row
        df_group1 = pd.concat([df_group1, df_group2[df_group2['Ware House'] == 'PF_Active'].rename(columns={'Ware House': 'Warehouse', 'one': 'Quantity'})], ignore_index=True)
        df_grouped = df_group1
        
        # Warehouse order
        warehouses = {'G_Active_1': 11, 'G_Active_2': 12, 'G_MD_1': 13, 'G_MD_2': 14, 
                      'HGBU_Extra': 18, 'Pre_Ship_1': 15, 'Pre_Ship_2': 16, 'WIPLines1': 9, 
                      'WIPLines2': 10, 'WIP_Cut_1': 2, 'WIP_Emb_1': 17, 'WIP_P1': 4, 
                      'WIP_Pri_1': 3, 'WIP_Sew_1': 5, 'WIP_Sew_2': 6, 'WIP_Sew_P1': 7, 
                      'WIP_Sew_P2': 8, 'PF_Active': 1}
        
        df_grouped['Order'] = df_grouped['Warehouse'].map(warehouses)
        df_grouped = df_grouped.sort_values(by='Order').reset_index(drop=True)
        
        # Calculate number of days
        df["number of days"] = (pd.to_datetime(current_date, format='%d-%m-%Y') - pd.to_datetime(df['Last Movement Date'], format='%d-%m-%Y')).dt.days
        df2["number of days"] = (pd.to_datetime(current_date, format='%d-%m-%Y') - pd.to_datetime(df2['last transaction date'], format='%d-%m-%Y')).dt.days
        
        # Days categories
        df["days cat"] = pd.cut(df["number of days"], bins=[-np.inf, 15, 30, 60, 90, 180, np.inf], 
                                labels=["0 - 15 days", "16 - 30 days", "31 - 60 days", "61 - 90 days", "91 - 180 days", "180+ days"])
        df2["days cat"] = pd.cut(df2["number of days"], bins=[-np.inf, 15, 30, 60, 90, 180, np.inf], 
                                 labels=["0 - 15 days", "16 - 30 days", "31 - 60 days", "61 - 90 days", "91 - 180 days", "180+ days"])
        
        # Statistical significance
        t_95_table = {2: 4.303, 3: 3.182, 4: 2.776, 5: 2.571, 6: 2.447, 7: 2.365, 8: 2.306, 
                      9: 2.262, 10: 2.228, 11: 2.201, 12: 2.179, 13: 2.160, 14: 2.145, 15: 2.131, 
                      16: 2.120, 17: 2.110, 18: 2.101, 19: 2.093, 20: 2.086, 21: 2.080, 22: 2.074, 
                      23: 2.069, 24: 2.064, 25: 2.060, 26: 2.056, 27: 2.052, 28: 2.048, 29: 2.045, 30: 2.042}
        
        statistical_sig = {}
        for i in df["Warehouse"].unique():
            temp_df = df[df["Warehouse"] == i]
            mean = temp_df["number of days"].mean()
            sigma = temp_df["number of days"].std()
            if len(temp_df) > 30:
                CI_pos = mean + 1.96 * (sigma / np.sqrt(len(temp_df)))
            else:
                CI_pos = mean + t_95_table[len(temp_df)] * (sigma / np.sqrt(len(temp_df)))
            statistical_sig[i] = CI_pos
        
        statistical_sig2 = {}
        for i in df2["Ware House"].unique():
            temp_df = df2[df2["Ware House"] == i]
            mean = temp_df["number of days"].mean()
            sigma = temp_df["number of days"].std()
            if len(temp_df) > 30:
                CI_pos = mean + 1.96 * (sigma / np.sqrt(len(temp_df)))
            else:
                CI_pos = mean + t_95_table[len(temp_df)] * (sigma / np.sqrt(len(temp_df)))
            statistical_sig2[i] = CI_pos
        
        df["Critical"] = df["number of days"] > df["Warehouse"].map(statistical_sig)
        df2["Critical"] = df2["number of days"] > df2["Ware House"].map(statistical_sig2)
        
        # Pivot tables
        pivot_table = pd.pivot_table(df, values='Quantity', index='Warehouse', columns='days cat', aggfunc='sum', fill_value=0)
        pivot_table2 = pd.pivot_table(df2, values='one', index='Ware House', columns='days cat', aggfunc='sum', fill_value=0)
        
        pivot_table = pd.concat([pivot_table, pivot_table2.loc[['PF_Active']].rename(index={'PF_Active': 'PF_Active'})], axis=0).fillna(0)
        pivot_table['Order'] = pivot_table.index.map(warehouses)
        pivot_table = pivot_table.sort_values(by='Order').drop(columns=['Order'])
        
        # Time category totals
        time_cat_totals = df.groupby('days cat')['Quantity'].sum().reset_index()
        
        # Critical totals
        crucial_totals = df[df['Critical']].groupby('Warehouse')['Quantity'].sum().reset_index()
        
        # Add PF_Active critical totals from df2
        if 'PF_Active' in df2['Ware House'].unique():
            pf_critical = df2[(df2['Ware House'] == 'PF_Active') & (df2['Critical'])].groupby('Ware House')['one'].sum().reset_index()
            if not pf_critical.empty:
                pf_critical.columns = ['Warehouse', 'Quantity']
                crucial_totals = pd.concat([crucial_totals, pf_critical], ignore_index=True)
        
        crucial_totals['Order'] = crucial_totals['Warehouse'].map(warehouses)
        crucial_totals = crucial_totals.sort_values(by='Order').drop(columns=['Order']).reset_index(drop=True)
        
        # Store in session state
        st.session_state.df = df
        st.session_state.df2 = df2
        st.session_state.df_grouped = df_grouped
        st.session_state.pivot_table = pivot_table
        st.session_state.time_cat_totals = time_cat_totals
        st.session_state.crucial_totals = crucial_totals
        st.session_state.current_date = current_date
        st.session_state.processed = True
        st.rerun()

# Main dashboard
if st.session_state.processed:
    df = st.session_state.df
    df2 = st.session_state.df2
    df_grouped = st.session_state.df_grouped
    pivot_table = st.session_state.pivot_table
    time_cat_totals = st.session_state.time_cat_totals
    crucial_totals = st.session_state.crucial_totals
    current_date = st.session_state.current_date
    
    # Display current date
    st.info(f"Analysis Date: {current_date}")
    
    # Cards for total quantities by warehouse
    st.subheader("Total Quantity by Warehouse")
    qty_sorted = df_grouped.sort_values(by='Quantity', ascending=False).reset_index(drop=True)
    qty_sorted['Rank'] = qty_sorted.index + 1
    
    # Create cards using Streamlit columns
    num_cols = 6
    cols = st.columns(num_cols)
    for index, row in df_grouped.iterrows():
        col_idx = index % num_cols
        rank = qty_sorted.loc[qty_sorted['Warehouse'] == row['Warehouse'], 'Rank'].values[0]
        is_top3 = rank <= 3
        
        with cols[col_idx]:
            if is_top3:
                st.markdown(f"""
                <div style="border: 2px solid #ff6b6b; border-radius: 8px; padding: 12px; text-align: center; background-color: #ffe0e0;">
                    <p style="margin: 0; font-weight: bold; font-size: 14px;">{row['Warehouse']}</p>
                    <p style="margin: 5px 0 0 0; font-size: 20px; font-weight: bold; color: #d63031;">{int(row['Quantity']):,}</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="border: 1px solid #ccc; border-radius: 8px; padding: 12px; text-align: center; background-color: #f0f8ff;">
                    <p style="margin: 0; font-weight: bold; font-size: 14px;">{row['Warehouse']}</p>
                    <p style="margin: 5px 0 0 0; font-size: 20px; font-weight: bold; color: #2c3e50;">{int(row['Quantity']):,}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Bar chart
    st.subheader("Total Quantity Distribution")
    fig, ax = plt.subplots(figsize=(12, 6), dpi=600)
    ax.bar(df_grouped['Warehouse'], df_grouped['Quantity'], color='skyblue')
    ax.set_xlabel('Warehouse')
    ax.set_ylabel('Total Quantity')
    ax.set_title('Total Quantity by Warehouse')
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(fig, dpi=600)
    
    # Pivot table
    st.subheader("Quantity Distribution by Time Category")
    st.dataframe(pivot_table, use_container_width=True)
    
    # Sidebar filters
    st.sidebar.header("Filters")
    filter_type = st.sidebar.radio("Filter Type", ["Days", "Statistical"])
    
    warehouse_list = ['All'] + list(df_grouped['Warehouse'].unique())
    selected_warehouse = st.sidebar.selectbox("Select Warehouse", warehouse_list)
    
    # Right side content based on filter
    col_left, col_right = st.columns([2, 1])
    
    with col_right:
        if filter_type == "Days":
            st.subheader("Time Category Summary")
            
            # Calculate time category totals based on selected warehouse
            if selected_warehouse == 'All':
                filtered_time_cats = time_cat_totals
            else:
                if selected_warehouse == 'PF_Active':
                    filtered_time_cats = df2[df2['Ware House'] == selected_warehouse].groupby('days cat')['one'].sum().reset_index()
                    filtered_time_cats.columns = ['days cat', 'Quantity']
                else:
                    filtered_time_cats = df[df['Warehouse'] == selected_warehouse].groupby('days cat')['Quantity'].sum().reset_index()
            
            # Cards for time categories using columns
            for index, row in filtered_time_cats.iterrows():
                st.markdown(f"""
                <div style="border: 1px solid #9e9e9e; border-radius: 8px; padding: 12px; margin: 8px 0; text-align: center; background-color: #f0f4c3;">
                    <p style="margin: 0; font-weight: bold; font-size: 14px;">{row['days cat']}</p>
                    <p style="margin: 5px 0 0 0; font-size: 20px; font-weight: bold; color: #33691e;">{int(row['Quantity']):,}</p>
                </div>
                """, unsafe_allow_html=True)
            
            # Days category filter
            days_categories = ["0 - 15 days", "16 - 30 days", "31 - 60 days", "61 - 90 days", "91 - 180 days", "180+ days"]
            selected_days = st.multiselect("Select Days Categories", days_categories, default=days_categories)
            
        else:  # Statistical
            st.subheader("Critical Items Summary")
            
            # Cards for critical quantities
            display_crucial = crucial_totals if selected_warehouse == 'All' else crucial_totals[crucial_totals['Warehouse'] == selected_warehouse]
            for index, row in display_crucial.iterrows():
                st.markdown(f"""
                <div style="border: 1px solid #ff9800; border-radius: 8px; padding: 12px; margin: 8px 0; text-align: center; background-color: #ffe0b2;">
                    <p style="margin: 0; font-weight: bold; font-size: 14px;">{row['Warehouse']}</p>
                    <p style="margin: 5px 0 0 0; font-size: 20px; font-weight: bold; color: #e65100;">{int(row['Quantity']):,}</p>
                </div>
                """, unsafe_allow_html=True)
    
    with col_left:
        st.subheader("Detailed Items")
        
        if filter_type == "Days":
            # Filter by warehouse and days
            if selected_warehouse == 'All':
                filtered_df = df[df['days cat'].isin(selected_days)]
                if not filtered_df.empty:
                    st.write("**Stock Source (All Warehouses)**")
                    display_df = filtered_df[['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Warehouse']].copy()
                    st.dataframe(display_df, use_container_width=True, height=400)
                    
                    # Download button
                    csv = display_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Current View",
                        data=csv,
                        file_name=f"All_Warehouses_{current_date}.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No items found for the selected day categories.")
            else:
                if selected_warehouse == 'PF_Active':
                    filtered_df2 = df2[(df2['Ware House'] == selected_warehouse) & (df2['days cat'].isin(selected_days))]
                    if not filtered_df2.empty:
                        st.write("**Fabric Stock (PF_Active)**")
                        display_df2 = filtered_df2[['Project', 'Lot No', 'Style-color', 'Gramaj']].copy()
                        st.dataframe(display_df2, use_container_width=True, height=400)
                        
                        # Download button
                        csv = display_df2.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Download Current View",
                            data=csv,
                            file_name=f"{selected_warehouse}_{current_date}.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No items found for the selected day categories.")
                else:
                    filtered_df = df[(df['Warehouse'] == selected_warehouse) & (df['days cat'].isin(selected_days))]
                    if not filtered_df.empty:
                        st.write(f"**Stock Source ({selected_warehouse})**")
                        display_df = filtered_df[['Project', 'Color', 'Size', 'Quantity', 'Customer']].copy()
                        st.dataframe(display_df, use_container_width=True, height=400)
                        
                        # Download button
                        csv = display_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Download Current View",
                            data=csv,
                            file_name=f"{selected_warehouse}_{current_date}.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No items found for the selected day categories.")
        
        else:  # Statistical
            # Filter by critical items
            if selected_warehouse == 'All':
                filtered_df = df[df['Critical']]
                if not filtered_df.empty:
                    st.write("**Critical Items (All Warehouses)**")
                    display_df = filtered_df[['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Warehouse']].copy()
                    st.dataframe(display_df, use_container_width=True, height=400)
                    
                    # Download button
                    csv = display_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Critical Items",
                        data=csv,
                        file_name=f"Critical_All_Warehouses_{current_date}.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No critical items found.")
            else:
                if selected_warehouse == 'PF_Active':
                    filtered_df2 = df2[(df2['Ware House'] == selected_warehouse) & (df2['Critical'])]
                    if not filtered_df2.empty:
                        st.write("**Critical Fabric Stock (PF_Active)**")
                        display_df2 = filtered_df2[['Project', 'Lot No', 'Style-color', 'Gramaj']].copy()
                        st.dataframe(display_df2, use_container_width=True, height=400)
                        
                        # Download button
                        csv = display_df2.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Download Critical Items",
                            data=csv,
                            file_name=f"Critical_{selected_warehouse}_{current_date}.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No critical items found.")
                else:
                    filtered_df = df[(df['Warehouse'] == selected_warehouse) & (df['Critical'])]
                    if not filtered_df.empty:
                        st.write(f"**Critical Items ({selected_warehouse})**")
                        display_df = filtered_df[['Project', 'Color', 'Size', 'Quantity', 'Customer']].copy()
                        st.dataframe(display_df, use_container_width=True, height=400)
                        
                        # Download button
                        csv = display_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Download Critical Items",
                            data=csv,
                            file_name=f"Critical_{selected_warehouse}_{current_date}.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No critical items found.")
    
    # Download all warehouses button
    st.sidebar.markdown("---")
    st.sidebar.subheader("Download All Warehouses")
    if st.sidebar.button("📦 Download All Warehouses (ZIP)"):
        import zipfile
        from io import BytesIO
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Apply filters based on filter type
            if filter_type == "Days":
                # Export each warehouse from df with day filter applied
                for warehouse in df['Warehouse'].unique():
                    warehouse_df = df[(df['Warehouse'] == warehouse) & (df['days cat'].isin(selected_days))][['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Last Movement Date', 'number of days']].copy()
                    if not warehouse_df.empty:
                        csv_data = warehouse_df.to_csv(index=False)
                        zip_file.writestr(f"{warehouse}_{current_date}_DaysFilter.csv", csv_data)
                
                # Export PF_Active if exists with day filter
                if 'PF_Active' in df2['Ware House'].unique():
                    pf_active_df = df2[(df2['Ware House'] == 'PF_Active') & (df2['days cat'].isin(selected_days))][['Project', 'Lot No', 'Style-color', 'Gramaj', 'last transaction date', 'number of days']].copy()
                    if not pf_active_df.empty:
                        csv_data = pf_active_df.to_csv(index=False)
                        zip_file.writestr(f"PF_Active_{current_date}_DaysFilter.csv", csv_data)
            
            else:  # Statistical filter
                # Export each warehouse from df with critical filter applied
                for warehouse in df['Warehouse'].unique():
                    warehouse_df = df[(df['Warehouse'] == warehouse) & (df['Critical'])][['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Last Movement Date', 'number of days']].copy()
                    if not warehouse_df.empty:
                        csv_data = warehouse_df.to_csv(index=False)
                        zip_file.writestr(f"{warehouse}_{current_date}_Critical.csv", csv_data)
                
                # Export PF_Active if exists with critical filter
                if 'PF_Active' in df2['Ware House'].unique():
                    pf_active_df = df2[(df2['Ware House'] == 'PF_Active') & (df2['Critical'])][['Project', 'Lot No', 'Style-color', 'Gramaj', 'last transaction date', 'number of days']].copy()
                    if not pf_active_df.empty:
                        csv_data = pf_active_df.to_csv(index=False)
                        zip_file.writestr(f"PF_Active_{current_date}_Critical.csv", csv_data)
        
        zip_buffer.seek(0)
        filter_suffix = "DaysFilter" if filter_type == "Days" else "Critical"
        st.sidebar.download_button(
            label="💾 Download ZIP File",
            data=zip_buffer,
            file_name=f"All_Warehouses_{current_date}_{filter_suffix}.zip",
            mime="application/zip"
        )
    
    # Reset button
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Reset and Upload New Files"):
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()
    
    # Email Section
    st.markdown("---")
    st.header("📧 Send Email Report")
    
    # Department to warehouse mapping based on line 59 warehouses
    # Warehouses: G_Active_1, G_Active_2, G_MD_1, G_MD_2, HGBU_Extra, Pre_Ship_1, Pre_Ship_2, 
    #             WIPLines1, WIPLines2, WIP_Cut_1, WIP_Emb_1, WIP_P1, WIP_Pri_1, 
    #             WIP_Sew_1, WIP_Sew_2, WIP_Sew_P1, WIP_Sew_P2, PF_Active
    
    department_warehouse_mapping = {
        "Garment Active (G_Active)": {
            "email": "garment.active@company.com",
            "warehouses": ['G_Active_1', 'G_Active_2']
        },
        "Garment MD (G_MD)": {
            "email": "garment.md@company.com",
            "warehouses": ['G_MD_1', 'G_MD_2']
        },
        "Pre-Shipment": {
            "email": "preshipment@company.com",
            "warehouses": ['Pre_Ship_1', 'Pre_Ship_2']
        },
        "WIP Lines": {
            "email": "wiplines@company.com",
            "warehouses": ['WIPLines1', 'WIPLines2']
        },
        "WIP Sewing": {
            "email": "wipsewing@company.com",
            "warehouses": ['WIP_Sew_1', 'WIP_Sew_2', 'WIP_Sew_P1', 'WIP_Sew_P2']
        },
        "WIP Cutting & Print": {
            "email": "wipcutting@company.com",
            "warehouses": ['WIP_Cut_1', 'WIP_Pri_1', 'WIP_P1']
        },
        "WIP Embroidery": {
            "email": "wipembroidery@company.com",
            "warehouses": ['WIP_Emb_1']
        },
        "Fabric Department (PF_Active)": {
            "email": "fabric@company.com",
            "warehouses": ['PF_Active']
        },
        "HGBU Extra": {
            "email": "hgbu@company.com",
            "warehouses": ['HGBU_Extra']
        },
        "All Warehouses": {
            "email": "management@company.com",
            "warehouses": list(df['Warehouse'].unique()) + (['PF_Active'] if 'PF_Active' in df2['Ware House'].unique() else [])
        }
    }
    
    col_email1, col_email2 = st.columns([1, 1])
    
    with col_email1:
        st.subheader("Email Configuration")
        
        # Department selection
        selected_department = st.selectbox("Select Department", list(department_warehouse_mapping.keys()))
        department_warehouses = department_warehouse_mapping[selected_department]["warehouses"]
        
        # Recipient email input
        recipient_email = st.text_input("Recipient Email Address", placeholder="recipient@company.com")
        
        st.info(f"📦 Warehouses: {', '.join(department_warehouses)}")
        
        # Sender credentials
        sender_email = st.text_input("Your Gmail Address", placeholder="your.email@gmail.com")
        
        # Email subject
        email_subject = st.text_input("Email Subject", 
                                     value=f"Warehouse Stock Report - {selected_department} - {current_date}")
    
    with col_email2:
        st.subheader("Files to be Sent")
        
        # Automatically determine files based on department warehouses and current filters
        files_to_send = []
        total_items = 0
        
        if filter_type == "Days":
            st.write(f"**Filter: Days ({', '.join(selected_days)})**")
            filter_suffix = "DaysFilter"
            
            for warehouse in department_warehouses:
                if warehouse == 'PF_Active':
                    if 'PF_Active' in df2['Ware House'].unique():
                        warehouse_df = df2[(df2['Ware House'] == warehouse) & (df2['days cat'].isin(selected_days))]
                        if not warehouse_df.empty:
                            files_to_send.append(warehouse)
                            total_items += len(warehouse_df)
                else:
                    if warehouse in df['Warehouse'].unique():
                        warehouse_df = df[(df['Warehouse'] == warehouse) & (df['days cat'].isin(selected_days))]
                        if not warehouse_df.empty:
                            files_to_send.append(warehouse)
                            total_items += len(warehouse_df)
        else:  # Statistical
            st.write("**Filter: Critical Items**")
            filter_suffix = "Critical"
            
            for warehouse in department_warehouses:
                if warehouse == 'PF_Active':
                    if 'PF_Active' in df2['Ware House'].unique():
                        warehouse_df = df2[(df2['Ware House'] == warehouse) & (df2['Critical'])]
                        if not warehouse_df.empty:
                            files_to_send.append(warehouse)
                            total_items += len(warehouse_df)
                else:
                    if warehouse in df['Warehouse'].unique():
                        warehouse_df = df[(df['Warehouse'] == warehouse) & (df['Critical'])]
                        if not warehouse_df.empty:
                            files_to_send.append(warehouse)
                            total_items += len(warehouse_df)
        
        if files_to_send:
            st.success(f"✅ **{len(files_to_send)} file(s) ready to send**")
            st.write("**Files that will be attached:**")
            for file in files_to_send:
                st.write(f"• {file}_{current_date}_{filter_suffix}.xlsx")
            st.metric("Total Items", total_items)
        else:
            st.warning("⚠️ No data available for selected department with current filters")
    
    # Email template
    st.subheader("Email Template Preview")
    
    # Determine filter description based on filter type
    if filter_type == "Days":
        # Check if any late categories (over 60 days) are selected
        late_categories = [cat for cat in selected_days if cat in ['61 - 90 days', '91 - 180 days', '180+ days']]
        if late_categories:
            filter_description = "Projects that stayed over 60 days"
        else:
            filter_description = f"Days Filter: {', '.join(selected_days)}"
    else:
        filter_description = "Critical projects"
    
    # Get top 3 projects
    top_projects_text = ""
    if files_to_send:
        if filter_type == "Days":
            # Get top 3 most late projects (longest time)
            all_data_frames = []
            for warehouse in files_to_send:
                if warehouse == 'PF_Active':
                    if 'PF_Active' in df2['Ware House'].unique():
                        warehouse_df = df2[(df2['Ware House'] == warehouse) & (df2['days cat'].isin(selected_days))][['Project', 'number of days']].copy()
                        all_data_frames.append(warehouse_df)
                else:
                    if warehouse in df['Warehouse'].unique():
                        warehouse_df = df[(df['Warehouse'] == warehouse) & (df['days cat'].isin(selected_days))][['Project', 'number of days']].copy()
                        all_data_frames.append(warehouse_df)
            
            if all_data_frames:
                combined_df = pd.concat(all_data_frames, ignore_index=True)
                top_3 = combined_df.nlargest(3, 'number of days')
                top_projects_text = "\nTop 3 Most Late Projects:\n"
                for idx, row in top_3.iterrows():
                    top_projects_text += f"  {idx+1}. {row['Project']} - {int(row['number of days'])} days\n"
        else:  # Critical
            # Get top 3 critical projects (longest time)
            all_data_frames = []
            for warehouse in files_to_send:
                if warehouse == 'PF_Active':
                    if 'PF_Active' in df2['Ware House'].unique():
                        warehouse_df = df2[(df2['Ware House'] == warehouse) & (df2['Critical'])][['Project', 'number of days']].copy()
                        all_data_frames.append(warehouse_df)
                else:
                    if warehouse in df['Warehouse'].unique():
                        warehouse_df = df[(df['Warehouse'] == warehouse) & (df['Critical'])][['Project', 'number of days']].copy()
                        all_data_frames.append(warehouse_df)
            
            if all_data_frames:
                combined_df = pd.concat(all_data_frames, ignore_index=True)
                top_3 = combined_df.nlargest(3, 'number of days')
                top_projects_text = "\nTop 3 Most Critical Projects:\n"
                for idx, row in top_3.iterrows():
                    top_projects_text += f"  {idx+1}. {row['Project']} - {int(row['number of days'])} days\n"
    
    email_body_template = f"""
Dear {selected_department} Team,

Please find attached the Warehouse Stock Report for your review.

Report Details:
- Report Date: {current_date}
- Department: {selected_department}
- Filter Applied: {filter_description}
- Number of Files: {len(files_to_send)}
- Total Items: {total_items}
- Warehouses Included: {', '.join(files_to_send) if files_to_send else 'None'}{top_projects_text}

The attached Excel file(s) contain detailed information about stock items based on the applied filters.

Please review the data and take necessary actions as required.

Best regards,
Bassem
Planning Department
"""
    
    st.text_area("Email Body", email_body_template, height=300, disabled=True)
    
    # Send email button
    st.markdown("---")
    
    # Check credentials
    credentials_valid = bool(sender_email and recipient_email)
    
    if st.button("📤 Send Email", type="primary", disabled=not (credentials_valid and files_to_send)):
        if not sender_email:
            st.error("Please provide your email address")
        elif not recipient_email:
            st.error("Please provide recipient email address")
        elif not files_to_send:
            st.error("No data available for the selected department with current filters")
        else:
            with st.spinner("Preparing and sending email..."):
                try:
                    # Create message
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient_email
                    msg['Subject'] = email_subject
                    
                    # Attach email body
                    msg.attach(MIMEText(email_body_template, 'plain'))
                    
                    # Create Excel files for warehouses in this department
                    for warehouse_name in files_to_send:
                        # Get the filtered data for this warehouse
                        if filter_type == "Days":
                            if warehouse_name == 'PF_Active':
                                warehouse_data = df2[(df2['Ware House'] == warehouse_name) & 
                                                    (df2['days cat'].isin(selected_days))][
                                    ['Project', 'Lot No', 'Style-color', 'Gramaj', 'last transaction date', 'number of days']
                                ].copy()
                            else:
                                warehouse_data = df[(df['Warehouse'] == warehouse_name) & 
                                                   (df['days cat'].isin(selected_days))][
                                    ['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Last Movement Date', 'number of days']
                                ].copy()
                        else:  # Statistical
                            if warehouse_name == 'PF_Active':
                                warehouse_data = df2[(df2['Ware House'] == warehouse_name) & 
                                                    (df2['Critical'])][
                                    ['Project', 'Lot No', 'Style-color', 'Gramaj', 'last transaction date', 'number of days']
                                ].copy()
                            else:
                                warehouse_data = df[(df['Warehouse'] == warehouse_name) & 
                                                   (df['Critical'])][
                                    ['Project', 'Color', 'Size', 'Quantity', 'Customer', 'Last Movement Date', 'number of days']
                                ].copy()
                        
                        # Convert to Excel
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            warehouse_data.to_excel(writer, sheet_name='Stock Report', index=False)
                        excel_buffer.seek(0)
                        
                        # Attach file
                        filename = f"{warehouse_name}_{current_date}_{filter_suffix}.xlsx"
                        
                        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        part.set_payload(excel_buffer.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename={filename}')
                        msg.attach(part)
                    
                    # Send email based on selected method
                    # Send via Google Cloud Gmail API
                    import os
                    from google_auth_oauthlib.flow import InstalledAppFlow
                    import pickle
                    
                    SCOPES = ['https://www.googleapis.com/auth/gmail.send']
                    creds = None
                    
                    # Check for saved token
                    if os.path.exists('token.pickle'):
                        with open('token.pickle', 'rb') as token:
                            creds = pickle.load(token)
                    
                    # If no valid credentials, authenticate
                    if not creds or not creds.valid:
                        if creds and creds.expired and creds.refresh_token:
                            creds.refresh(Request())
                        else:
                            credentials_json_path = r"your_credentials_json.json"
                            
                            if os.path.exists(credentials_json_path):
                                flow = InstalledAppFlow.from_client_secrets_file(
                                    credentials_json_path, SCOPES)
                                creds = flow.run_local_server(port=0)
                                
                                # Save credentials for next run
                                with open('token.pickle', 'wb') as token:
                                    pickle.dump(creds, token)
                            else:
                                st.error("Credentials file not found at the specified path")
                                st.stop()
                    
                    # Build Gmail API service
                    service = build('gmail', 'v1', credentials=creds)
                    
                    # Encode message
                    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
                    message = {'raw': raw_message}
                    
                    # Send message
                    service.users().messages().send(userId='me', body=message).execute()
                    
                    st.success(f"✅ Email sent successfully to {selected_department} ({recipient_email})")
                    st.balloons()
                    
                except smtplib.SMTPAuthenticationError:
                    st.error("❌ Authentication failed. Please check your email and app password. For Gmail, make sure you're using an App Password, not your regular password.")
                except Exception as e:
                    st.error(f"❌ Failed to send email: {str(e)}\n\nPlease check your credentials and try again.")

else:
    st.info("Please upload both Stock Source and Fabric Stock files to begin analysis.")