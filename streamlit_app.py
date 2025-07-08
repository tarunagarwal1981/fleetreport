import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.colors import Color
import time
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
import os

# Page configuration
st.set_page_config(
    page_title="Vessel Performance Report Tool",
    page_icon="🚢",
    layout="wide"
)

# Initialize session state
if 'vessels' not in st.session_state:
    st.session_state.vessels = []

if 'selected_vessels' not in st.session_state or not isinstance(st.session_state.selected_vessels, set):
    st.session_state.selected_vessels = set()

if 'report_data' not in st.session_state:
    st.session_state.report_data = None

if 'search_query' not in st.session_state:
    st.session_state.search_query = ""

if 'performance_stats' not in st.session_state:
    st.session_state.performance_stats = {
        'total_requests': 0,
        'total_time': 0,
        'avg_response_time': 0
    }

# Enhanced Lambda Invocation Helper (Backwards Compatible)
def invoke_lambda_function_url(lambda_url, payload, timeout=30):
    """Invoke Lambda function via its Function URL using HTTP POST with performance tracking."""
    try:
        start_time = time.time()
        headers = {'Content-Type': 'application/json'}
        json_payload = json.dumps(payload)
        
        # Update request count
        st.session_state.performance_stats['total_requests'] += 1
        
        response = requests.post(
            lambda_url, 
            headers=headers, 
            data=json_payload,
            timeout=timeout
        )
        
        # Calculate response time
        response_time = time.time() - start_time
        st.session_state.performance_stats['total_time'] += response_time
        st.session_state.performance_stats['avg_response_time'] = (
            st.session_state.performance_stats['total_time'] / 
            st.session_state.performance_stats['total_requests']
        )
        
        if response.status_code != 200:
            st.error(f"HTTP error: {response.status_code} {response.reason} for url: {lambda_url}")
            st.error(f"Response content: {response.text}")
            return None
            
        result = response.json()
        st.success(f"✅ Data retrieved in {response_time:.2f}s")
        return result
            
    except requests.exceptions.HTTPError as http_err:
        st.error(f"HTTP error: {http_err}")
        return None
    except requests.exceptions.ConnectionError as conn_err:
        st.error(f"Connection error: {conn_err}")
        return None
    except requests.exceptions.Timeout as timeout_err:
        st.error(f"Timeout error: {timeout_err}")
        return None
    except requests.exceptions.RequestException as req_err:
        st.error(f"Request error: {req_err}")
        return None
    except Exception as e:
        st.error(f"Unexpected error: {str(e)}")
        return None

# Cached vessel loading function
@st.cache_data(ttl=3600)
def fetch_all_vessels(lambda_url):
    """Fetch vessel names from Lambda function with a limit of 1200."""
    query = "SELECT vessel_name FROM vessel_particulars ORDER BY vessel_name LIMIT 1200"
    
    result = invoke_lambda_function_url(lambda_url, {"sql_query": query})
    
    if result:
        extracted_vessel_names = []
        for item in result:
            if isinstance(item, dict) and 'vessel_name' in item:
                extracted_vessel_names.append(item['vessel_name'])
            elif isinstance(item, str):
                extracted_vessel_names.append(item)
        extracted_vessel_names.sort()
        return extracted_vessel_names
    
    return []

def filter_vessels_client_side(vessels, search_term):
    """Filter vessels on client side for better responsiveness."""
    if not search_term:
        return vessels
    
    search_lower = search_term.lower()
    return [v for v in vessels if search_lower in v.lower()]

def query_report_data(lambda_url, vessel_names):
    """Enhanced version of the original query_report_data function with better progress tracking."""
    if not vessel_names:
        return pd.DataFrame()

    # Calculate dates for previous month, previous-to-previous month, and previous-to-previous-to-previous month
    today = datetime.now()
    
    # Hull Condition Dates (last day of the month)
    first_day_current_month = today.replace(day=1)

    last_day_previous_month_hull = first_day_current_month - timedelta(days=1)
    prev_month_hull_str = last_day_previous_month_hull.strftime("%Y-%m-%d")
    prev_month_hull_col_name = f"Hull Condition {last_day_previous_month_hull.strftime('%b %y')}"

    first_day_previous_month_hull = last_day_previous_month_hull.replace(day=1)
    last_day_prev_prev_month_hull = first_day_previous_month_hull - timedelta(days=1)
    prev_prev_month_hull_str = last_day_prev_prev_month_hull.strftime("%Y-%m-%d")
    prev_prev_month_hull_col_name = f"Hull Condition {last_day_prev_prev_month_hull.strftime('%b %y')}"

    first_day_prev_prev_month_hull = last_day_prev_prev_month_hull.replace(day=1)
    last_day_prev_prev_prev_month_hull = first_day_prev_prev_month_hull - timedelta(days=1)
    prev_prev_prev_month_hull_str = last_day_prev_prev_prev_month_hull.strftime("%Y-%m-%d")
    prev_prev_prev_month_hull_col_name = f"Hull Condition {last_day_prev_prev_prev_month_hull.strftime('%b %y')}"

    # ME SFOC Dates (average of the entire month)
    prev_month_me_col_name = f"ME Efficiency {last_day_previous_month_hull.strftime('%b %y')}" 
    prev_prev_month_me_col_name = f"ME Efficiency {last_day_prev_prev_month_hull.strftime('%b %y')}"
    prev_prev_prev_month_me_col_name = f"ME Efficiency {last_day_prev_prev_prev_month_hull.strftime('%b %y')}"

    # Process vessels in smaller batches with enhanced progress tracking
    batch_size = 10
    all_fuel_saving_data = []
    all_cii_data = []
    all_prev_month_hull_data = []
    all_prev_prev_month_hull_data = []
    all_prev_prev_prev_month_hull_data = []
    all_prev_month_me_data = []
    all_prev_prev_month_me_data = []
    all_prev_prev_prev_month_me_data = []
    
    total_batches = (len(vessel_names) + batch_size - 1) // batch_size
    
    # Create progress bar and status text
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i in range(0, len(vessel_names), batch_size):
        batch_vessels = vessel_names[i:i+batch_size]
        batch_num = i//batch_size + 1
        
        # Update progress
        progress = batch_num / total_batches
        progress_bar.progress(progress)
        status_text.info(f"🔄 Processing batch {batch_num} of {total_batches} ({len(batch_vessels)} vessels)")
        
        quoted_vessel_names = [f"'{name}'" for name in batch_vessels]
        vessel_names_list_str = ", ".join(quoted_vessel_names)

        # Execute all queries for this batch
        batch_queries = [
            (f"Hull Roughness {last_day_previous_month_hull.strftime('%b %y')}", f"""
SELECT vessel_name, hull_rough_power_loss_pct_ed
FROM (
    SELECT vessel_name, hull_rough_power_loss_pct_ed,
           ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM hull_performance_six_months_daily
    WHERE vessel_name IN ({vessel_names_list_str})
    AND CAST(updated_ts AS DATE) = '{prev_month_hull_str}'
) AS subquery
WHERE rn = 1
""", all_prev_month_hull_data),
            
            (f"Hull Roughness {last_day_prev_prev_month_hull.strftime('%b %y')}", f"""
SELECT vessel_name, hull_rough_power_loss_pct_ed
FROM (
    SELECT vessel_name, hull_rough_power_loss_pct_ed,
           ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM hull_performance_six_months_daily
    WHERE vessel_name IN ({vessel_names_list_str})
    AND CAST(updated_ts AS DATE) = '{prev_prev_month_hull_str}'
) AS subquery
WHERE rn = 1
""", all_prev_prev_month_hull_data),
            
            (f"Hull Roughness {last_day_prev_prev_prev_month_hull.strftime('%b %y')}", f"""
SELECT vessel_name, hull_rough_power_loss_pct_ed
FROM (
    SELECT vessel_name, hull_rough_power_loss_pct_ed,
           ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM hull_performance_six_months_daily
    WHERE vessel_name IN ({vessel_names_list_str})
    AND CAST(updated_ts AS DATE) = '{prev_prev_prev_month_hull_str}'
) AS subquery
WHERE rn = 1
""", all_prev_prev_prev_month_hull_data),
            
            (f"ME SFOC {last_day_previous_month_hull.strftime('%b %y')}", f"""
SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
FROM vessel_performance_summary vps
JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE)
GROUP BY vp.vessel_name
""", all_prev_month_me_data),
            
            (f"ME SFOC {last_day_prev_prev_month_hull.strftime('%b %y')}", f"""
SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
FROM vessel_performance_summary vps
JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
GROUP BY vp.vessel_name
""", all_prev_prev_month_me_data),
            
            (f"ME SFOC {last_day_prev_prev_prev_month_hull.strftime('%b %y')}", f"""
SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
FROM vessel_performance_summary vps
JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '3 months')
AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
GROUP BY vp.vessel_name
""", all_prev_prev_prev_month_me_data),
            
            ("Potential Fuel Saving", f"""
SELECT vessel_name, hull_rough_excess_consumption_mt_ed 
FROM hull_performance_six_months 
WHERE vessel_name IN ({vessel_names_list_str})
""", all_fuel_saving_data),
            
            ("YTD CII", f"""
SELECT vp.vessel_name, cy.cii_rating
FROM vessel_particulars vp
JOIN cii_ytd cy ON CAST(vp.vessel_imo AS TEXT) = CAST(cy.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
""", all_cii_data)
        ]
        
        # Execute each query
        for query_name, query, data_list in batch_queries:
            with st.spinner(f"Fetching {query_name} data..."):
                result = invoke_lambda_function_url(lambda_url, {"sql_query": query})
                if result:
                    data_list.extend(result)

    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()

    # Process all collected data (same as original function)
    # Hull Data processing
    df_prev_month_hull = pd.DataFrame()
    if all_prev_month_hull_data:
        try:
            df_prev_month_hull = pd.DataFrame(all_prev_month_hull_data)
            if 'hull_rough_power_loss_pct_ed' in df_prev_month_hull.columns:
                df_prev_month_hull = df_prev_month_hull.rename(columns={'hull_rough_power_loss_pct_ed': f'Hull Roughness Power Loss % {last_day_previous_month_hull.strftime("%b %y")}'})
            else:
                df_prev_month_hull[f'Hull Roughness Power Loss % {last_day_previous_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_month_hull = df_prev_month_hull.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous month hull data: {str(e)}")
            df_prev_month_hull = pd.DataFrame()

    df_prev_prev_month_hull = pd.DataFrame()
    if all_prev_prev_month_hull_data:
        try:
            df_prev_prev_month_hull = pd.DataFrame(all_prev_prev_month_hull_data)
            if 'hull_rough_power_loss_pct_ed' in df_prev_prev_month_hull.columns:
                df_prev_prev_month_hull = df_prev_prev_month_hull.rename(columns={'hull_rough_power_loss_pct_ed': f'Hull Roughness Power Loss % {last_day_prev_prev_month_hull.strftime("%b %y")}'})
            else:
                df_prev_prev_month_hull[f'Hull Roughness Power Loss % {last_day_prev_prev_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_prev_month_hull = df_prev_prev_month_hull.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous-to-previous month hull data: {str(e)}")
            df_prev_prev_month_hull = pd.DataFrame()

    df_prev_prev_prev_month_hull = pd.DataFrame()
    if all_prev_prev_prev_month_hull_data:
        try:
            df_prev_prev_prev_month_hull = pd.DataFrame(all_prev_prev_prev_month_hull_data)
            if 'hull_rough_power_loss_pct_ed' in df_prev_prev_prev_month_hull.columns:
                df_prev_prev_prev_month_hull = df_prev_prev_prev_month_hull.rename(columns={'hull_rough_power_loss_pct_ed': f'Hull Roughness Power Loss % {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'})
            else:
                df_prev_prev_prev_month_hull[f'Hull Roughness Power Loss % {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_prev_prev_month_hull = df_prev_prev_prev_month_hull.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous-to-previous-to-previous month hull data: {str(e)}")
            df_prev_prev_prev_month_hull = pd.DataFrame()

    # ME Data processing
    df_prev_month_me = pd.DataFrame()
    if all_prev_month_me_data:
        try:
            df_prev_month_me = pd.DataFrame(all_prev_month_me_data)
            if 'avg_me_sfoc' in df_prev_month_me.columns:
                df_prev_month_me = df_prev_month_me.rename(columns={'avg_me_sfoc': f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'})
            else:
                df_prev_month_me[f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_month_me = df_prev_month_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous month ME data: {str(e)}")
            df_prev_month_me = pd.DataFrame()

    df_prev_prev_month_me = pd.DataFrame()
    if all_prev_prev_month_me_data:
        try:
            df_prev_prev_month_me = pd.DataFrame(all_prev_prev_month_me_data)
            if 'avg_me_sfoc' in df_prev_prev_month_me.columns:
                df_prev_prev_month_me = df_prev_prev_month_me.rename(columns={'avg_me_sfoc': f'ME SFOC {last_day_prev_prev_month_hull.strftime("%b %y")}'})
            else:
                df_prev_prev_month_me[f'ME SFOC {last_day_prev_prev_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_prev_month_me = df_prev_prev_month_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous-to-previous month ME data: {str(e)}")
            df_prev_prev_month_me = pd.DataFrame()

    df_prev_prev_prev_month_me = pd.DataFrame()
    if all_prev_prev_prev_month_me_data:
        try:
            df_prev_prev_prev_month_me = pd.DataFrame(all_prev_prev_prev_month_me_data)
            if 'avg_me_sfoc' in df_prev_prev_prev_month_me.columns:
                df_prev_prev_prev_month_me = df_prev_prev_prev_month_me.rename(columns={'avg_me_sfoc': f'ME SFOC {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'})
            else:
                df_prev_prev_prev_month_me[f'ME SFOC {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'] = pd.NA
            df_prev_prev_prev_month_me = df_prev_prev_prev_month_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous-to-previous-to-previous month ME data: {str(e)}")
            df_prev_prev_prev_month_me = pd.DataFrame()

    # Fuel saving and CII data processing
    df_fuel_saving = pd.DataFrame()
    if all_fuel_saving_data:
        try:
            df_fuel_saving = pd.DataFrame(all_fuel_saving_data)
            if 'hull_rough_excess_consumption_mt_ed' in df_fuel_saving.columns:
                df_fuel_saving = df_fuel_saving.rename(columns={'hull_rough_excess_consumption_mt_ed': 'Potential Fuel Saving'})
                df_fuel_saving['Potential Fuel Saving'] = df_fuel_saving['Potential Fuel Saving'].apply(
                    lambda x: 4.9 if pd.notna(x) and x > 5 else (0.0 if pd.notna(x) and x < 0 else x)
                )
            else:
                df_fuel_saving['Potential Fuel Saving'] = pd.NA
            df_fuel_saving = df_fuel_saving.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing fuel saving data: {str(e)}")
            df_fuel_saving = pd.DataFrame()

    df_cii = pd.DataFrame()
    if all_cii_data:
        try:
            df_cii = pd.DataFrame(all_cii_data)
            if 'cii_rating' in df_cii.columns:
                df_cii = df_cii.rename(columns={'cii_rating': 'YTD CII'})
            else:
                df_cii['YTD CII'] = pd.NA
            df_cii = df_cii.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing CII data: {str(e)}")
            df_cii = pd.DataFrame()

    # Merge DataFrames
    df_final = pd.DataFrame({'Vessel Name': list(vessel_names)})

    # Merge historical hull data
    if not df_prev_month_hull.empty:
        df_final = pd.merge(df_final, df_prev_month_hull, on='Vessel Name', how='left')

    if not df_prev_prev_month_hull.empty:
        df_final = pd.merge(df_final, df_prev_prev_month_hull, on='Vessel Name', how='left')

    if not df_prev_prev_prev_month_hull.empty:
        df_final = pd.merge(df_final, df_prev_prev_prev_month_hull, on='Vessel Name', how='left')

    # Merge ME data
    if not df_prev_month_me.empty:
        df_final = pd.merge(df_final, df_prev_month_me, on='Vessel Name', how='left')
            
    if not df_prev_prev_month_me.empty:
        df_final = pd.merge(df_final, df_prev_prev_month_me, on='Vessel Name', how='left')

    if not df_prev_prev_prev_month_me.empty:
        df_final = pd.merge(df_final, df_prev_prev_prev_month_me, on='Vessel Name', how='left')

    # Merge other data
    if not df_fuel_saving.empty:
        df_final = pd.merge(df_final, df_fuel_saving, on='Vessel Name', how='left')

    if not df_cii.empty:
        df_final = pd.merge(df_final, df_cii, on='Vessel Name', how='left')

    if df_final.empty:
        return pd.DataFrame()

    # Post-merge processing for final report
    df_final.insert(0, 'S. No.', range(1, 1 + len(df_final)))
    
    # Hull Condition logic
    def get_hull_condition(value):
        if pd.isna(value):
            return "N/A"
        if value < 15:
            return "Good"
        elif 15 <= value <= 25:
            return "Average"
        else:
            return "Poor"
    
    # Apply Hull Condition to historical columns
    if f'Hull Roughness Power Loss % {last_day_previous_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_month_hull_col_name] = df_final[f'Hull Roughness Power Loss % {last_day_previous_month_hull.strftime("%b %y")}'].apply(get_hull_condition)
    else:
        df_final[prev_month_hull_col_name] = "N/A"

    if f'Hull Roughness Power Loss % {last_day_prev_prev_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_prev_month_hull_col_name] = df_final[f'Hull Roughness Power Loss % {last_day_prev_prev_month_hull.strftime("%b %y")}'].apply(get_hull_condition)
    else:
        df_final[prev_prev_month_hull_col_name] = "N/A"

    if f'Hull Roughness Power Loss % {last_day_prev_prev_prev_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_prev_prev_month_hull_col_name] = df_final[f'Hull Roughness Power Loss % {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'].apply(get_hull_condition)
    else:
        df_final[prev_prev_prev_month_hull_col_name] = "N/A"

    # ME Efficiency logic
    def get_me_efficiency(value):
        if pd.isna(value):
            return "N/A"
        if value < 160:
            return "Anomalous data"
        elif value < 180:
            return "Good"
        elif 180 <= value <= 190:
            return "Average"
        else:
            return "Poor"
    
    # Apply ME Efficiency
    if f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_month_me_col_name] = df_final[f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'].apply(get_me_efficiency)
    else:
        df_final[prev_month_me_col_name] = "N/A"

    if f'ME SFOC {last_day_prev_prev_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_prev_month_me_col_name] = df_final[f'ME SFOC {last_day_prev_prev_month_hull.strftime("%b %y")}'].apply(get_me_efficiency)
    else:
        df_final[prev_prev_month_me_col_name] = "N/A"

    if f'ME SFOC {last_day_prev_prev_prev_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_prev_prev_month_me_col_name] = df_final[f'ME SFOC {last_day_prev_prev_prev_month_hull.strftime("%b %y")}'].apply(get_me_efficiency)
    else:
        df_final[prev_prev_prev_month_me_col_name] = "N/A"

    # Add empty Comments column
    df_final['Comments'] = ""

    # Define the desired order of columns
    desired_columns_order = [
        'S. No.', 
        'Vessel Name', 
        prev_month_hull_col_name,
        prev_prev_month_hull_col_name,
        prev_prev_prev_month_hull_col_name,
        prev_month_me_col_name,
        prev_prev_month_me_col_name,
        prev_prev_prev_month_me_col_name,
        'Potential Fuel Saving',
        'YTD CII',
        'Comments'
    ]
    
    # Filter df_final to only include columns that exist and are in the desired order
    existing_and_ordered_columns = [col for col in desired_columns_order if col in df_final.columns]
    df_final = df_final[existing_and_ordered_columns]

    st.success("✅ Enhanced report data retrieved and processed successfully!")
    return df_final

# Styling Functions
def style_condition_columns(row):
    """Apply styling to condition columns."""
    styles = [''] * len(row)
    
    # Style hull condition columns
    hull_condition_cols = [col for col in row.index if 'Hull Condition' in col]
    for col_name in hull_condition_cols:
        if col_name in row.index:
            hull_val = row[col_name]
            if hull_val == "Good":
                styles[row.index.get_loc(col_name)] = 'background-color: #d4edda; color: black;'
            elif hull_val == "Average":
                styles[row.index.get_loc(col_name)] = 'background-color: #fff3cd; color: black;'
            elif hull_val == "Poor":
                styles[row.index.get_loc(col_name)] = 'background-color: #f8d7da; color: black;'
    
    # Style ME efficiency columns
    me_efficiency_cols = [col for col in row.index if 'ME Efficiency' in col]
    for col_name in me_efficiency_cols:
        if col_name in row.index:
            me_val = row[col_name]
            if me_val == "Good":
                styles[row.index.get_loc(col_name)] = 'background-color: #d4edda; color: black;'
            elif me_val == "Average":
                styles[row.index.get_loc(col_name)] = 'background-color: #fff3cd; color: black;'
            elif me_val == "Poor":
                styles[row.index.get_loc(col_name)] = 'background-color: #f8d7da; color: black;'
            elif me_val == "Anomalous data":
                styles[row.index.get_loc(col_name)] = 'background-color: #e0e0e0; color: black;'
                
    return styles

def create_excel_download_with_styling(df, filename):
    """Create Excel file with styling."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Vessel Report"

    # Write headers
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)

    # Write data and apply styling
    for row_idx, row_data in df.iterrows():
        for col_idx, (col_name, cell_value) in enumerate(row_data.items(), 1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value)
            
            if 'Hull Condition' in col_name or 'ME Efficiency' in col_name:
                if cell_value == "Good":
                    cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                elif cell_value == "Average":
                    cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                elif cell_value == "Poor":
                    cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                elif cell_value == "Anomalous data":
                    cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.font = Font(color="000000")
            elif col_name == 'YTD CII':
                cell.alignment = Alignment(horizontal='center')

    # Auto-adjust column widths
    for col_idx, column in enumerate(df.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for cell in ws[column_letter]:
            try:
                if cell.value is not None and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output)
    return output.getvalue()

def get_cell_color(cell_value):
    """Get background color for table cell based on value."""
    color_map = {
        "Good": "D4EDDA",      # Light green
        "Average": "FFF3CD",   # Light yellow
        "Poor": "F8D7DA",      # Light red
        "Anomalous data": "E0E0E0"  # Light gray
    }
    return color_map.get(cell_value, None)

def create_advanced_word_report(df, template_path="Fleet Performance Template.docx"):
    """Create an advanced Word report with better formatting and multiple sections."""
    try:
        if not os.path.exists(template_path):
            st.error(f"Template file '{template_path}' not found in the repository root.")
            return None
        
        doc = Document(template_path)
        
        # Find placeholder and replace with comprehensive report
        placeholder_found = False
        
        for paragraph in doc.paragraphs:
            if "{{Template}}" in paragraph.text:
                # Clear the placeholder
                paragraph.clear()
                placeholder_found = True
                
                # Add report title
                title_paragraph = doc.add_paragraph()
                title_run = title_paragraph.add_run("Fleet Performance Analysis Report")
                title_run.font.size = Inches(0.2)
                title_run.font.bold = True
                title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add generation date
                date_paragraph = doc.add_paragraph()
                date_run = date_paragraph.add_run(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
                date_run.font.italic = True
                date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add summary statistics
                doc.add_paragraph()
                doc.add_paragraph("Executive Summary", style='Heading 2')
                
                # Calculate summary stats
                total_vessels = len(df)
                hull_cols = [col for col in df.columns if 'Hull Condition' in col]
                me_cols = [col for col in df.columns if 'ME Efficiency' in col]
                
                summary_text = f"This report covers {total_vessels} vessels with performance data across multiple months. "
                
                if hull_cols:
                    latest_hull_col = hull_cols[0]
                    good_hulls = len(df[df[latest_hull_col] == "Good"])
                    hull_percentage = (good_hulls / total_vessels * 100) if total_vessels > 0 else 0
                    summary_text += f"Hull condition analysis shows {good_hulls} vessels ({hull_percentage:.1f}%) with good hull condition. "
                
                if me_cols:
                    latest_me_col = me_cols[0]
                    good_me = len(df[df[latest_me_col] == "Good"])
                    me_percentage = (good_me / total_vessels * 100) if total_vessels > 0 else 0
                    summary_text += f"Main engine efficiency analysis indicates {good_me} vessels ({me_percentage:.1f}%) with good ME efficiency. "
                
                if 'Potential Fuel Saving' in df.columns:
                    fuel_savings = df['Potential Fuel Saving'].apply(
                        lambda x: float(x) if pd.notna(x) and str(x) != 'N/A' else 0
                    )
                    total_fuel_saving = fuel_savings.sum()
                    avg_fuel_saving = fuel_savings.mean()
                    summary_text += f"Total potential fuel saving across the fleet is {total_fuel_saving:.2f} MT/day with an average of {avg_fuel_saving:.2f} MT/day per vessel."
                
                doc.add_paragraph(summary_text)
                
                # Add detailed data table
                doc.add_paragraph()
                doc.add_paragraph("Detailed Performance Data", style='Heading 2')
                
                # Create the main data table
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Style header row
                header_cells = table.rows[0].cells
                for i, column_name in enumerate(df.columns):
                    header_cells[i].text = str(column_name)
                    for run in header_cells[i].paragraphs[0].runs:
                        run.font.bold = True
                        try:
                            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
                        except Exception as e:
                            st.warning(f"Could not set font color for header: {e}")
                    header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Set header background to dark blue
                    try:
                        cell_shading = header_cells[i]._tc.get_or_add_tcPr()
                        cell_fill = cell_shading.get_or_add_shd()
                        cell_fill.fill = "2F75B5"  # Dark blue
                    except Exception as e:
                        st.warning(f"Could not set header background color: {e}")
                
                # Add data rows with formatting
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        cell_value = str(value) if pd.notna(value) else "N/A"
                        row_cells[i].text = cell_value
                        row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Apply conditional formatting
                        column_name = df.columns[i]
                        if 'Hull Condition' in column_name or 'ME Efficiency' in column_name:
                            color = get_cell_color(cell_value)
                            if color:
                                try:
                                    cell_shading = row_cells[i]._tc.get_or_add_tcPr()
                                    cell_fill = cell_shading.get_or_add_shd()
                                    cell_fill.fill = color
                                except Exception as e:
                                    st.warning(f"Could not set cell background color for data: {e}")
                
                # Add legend
                doc.add_paragraph()
                doc.add_paragraph("Legend", style='Heading 3')
                
                legend_table = doc.add_table(rows=5, cols=2)
                legend_table.style = 'Table Grid'
                
                legend_data = [
                    ("Good", "Performance within optimal range"),
                    ("Average", "Performance within acceptable range"),
                    ("Poor", "Performance requires attention"),
                    ("Anomalous data", "Data outside normal parameters"),
                    ("N/A", "Data not available for this period")
                ]
                
                for i, (status, description) in enumerate(legend_data):
                    legend_table.cell(i, 0).text = status
                    legend_table.cell(i, 1).text = description
                    
                    # Apply color coding to legend
                    if status in ["Good", "Average", "Poor", "Anomalous data"]:
                        color = get_cell_color(status)
                        if color:
                            try:
                                cell_shading = legend_table.cell(i, 0)._tc.get_or_add_tcPr()
                                cell_fill = cell_shading.get_or_add_shd()
                                cell_fill.fill = color
                            except Exception as e:
                                st.warning(f"Could not set cell background color for legend: {e}")
                
                break # Exit loop after finding and processing the placeholder
        
        if not placeholder_found:
            st.warning("Placeholder '{{Template}}' not found. Adding report at the end of document.")
            # Add report at end if placeholder not found
            doc.add_page_break()
            
            # Add report title
            title_paragraph = doc.add_paragraph()
            title_run = title_paragraph.add_run("Fleet Performance Analysis Report")
            title_run.font.size = Inches(0.2)
            title_run.font.bold = True
            title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add generation date
            date_paragraph = doc.add_paragraph()
            date_run = date_paragraph.add_run(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
            date_run.font.italic = True
            date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add summary statistics
            doc.add_paragraph()
            doc.add_paragraph("Executive Summary", style='Heading 2')
            
            # Calculate summary stats
            total_vessels = len(df)
            hull_cols = [col for col in df.columns if 'Hull Condition' in col]
            me_cols = [col for col in df.columns if 'ME Efficiency' in col]
            
            summary_text = f"This report covers {total_vessels} vessels with performance data across multiple months. "
            
            if hull_cols:
                latest_hull_col = hull_cols[0]
                good_hulls = len(df[df[latest_hull_col] == "Good"])
                hull_percentage = (good_hulls / total_vessels * 100) if total_vessels > 0 else 0
                summary_text += f"Hull condition analysis shows {good_hulls} vessels ({hull_percentage:.1f}%) with good hull condition. "
            
            if me_cols:
                latest_me_col = me_cols[0]
                good_me = len(df[df[latest_me_col] == "Good"])
                me_percentage = (good_me / total_vessels * 100) if total_vessels > 0 else 0
                summary_text += f"Main engine efficiency analysis indicates {good_me} vessels ({me_percentage:.1f}%) with good ME efficiency. "
            
            if 'Potential Fuel Saving' in df.columns:
                fuel_savings = df['Potential Fuel Saving'].apply(
                    lambda x: float(x) if pd.notna(x) and str(x) != 'N/A' else 0
                )
                total_fuel_saving = fuel_savings.sum()
                avg_fuel_saving = fuel_savings.mean()
                summary_text += f"Total potential fuel saving across the fleet is {total_fuel_saving:.2f} MT/day with an average of {avg_fuel_saving:.2f} MT/day per vessel."
            
            doc.add_paragraph(summary_text)
            
            # Add detailed data table
            doc.add_paragraph()
            doc.add_paragraph("Detailed Performance Data", style='Heading 2')
            
            # Create the main data table
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Style header row
            header_cells = table.rows[0].cells
            for i, column_name in enumerate(df.columns):
                header_cells[i].text = str(column_name)
                for run in header_cells[i].paragraphs[0].runs:
                    run.font.bold = True
                    try:
                        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
                    except Exception as e:
                        st.warning(f"Could not set font color for header: {e}")
                header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Set header background to dark blue
                try:
                    cell_shading = header_cells[i]._tc.get_or_add_tcPr()
                    cell_fill = cell_shading.get_or_add_shd()
                    cell_fill.fill = "2F75B5"  # Dark blue
                except Exception as e:
                    st.warning(f"Could not set header background color: {e}")
            
            # Add data rows with formatting
            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    cell_value = str(value) if pd.notna(value) else "N/A"
                    row_cells[i].text = cell_value
                    row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Apply conditional formatting
                    column_name = df.columns[i]
                    if 'Hull Condition' in column_name or 'ME Efficiency' in column_name:
                        color = get_cell_color(cell_value)
                        if color:
                            try:
                                cell_shading = row_cells[i]._tc.get_or_add_tcPr()
                                cell_fill = cell_shading.get_or_add_shd()
                                cell_fill.fill = color
                            except Exception as e:
                                st.warning(f"Could not set cell background color for data: {e}")
            
            # Add legend
            doc.add_paragraph()
            doc.add_paragraph("Legend", style='Heading 3')
            
            legend_table = doc.add_table(rows=5, cols=2)
            legend_table.style = 'Table Grid'
            
            legend_data = [
                ("Good", "Performance within optimal range"),
                ("Average", "Performance within acceptable range"),
                ("Poor", "Performance requires attention"),
                ("Anomalous data", "Data outside normal parameters"),
                ("N/A", "Data not available for this period")
            ]
            
            for i, (status, description) in enumerate(legend_data):
                legend_table.cell(i, 0).text = status
                legend_table.cell(i, 1).text = description
                
                # Apply color coding to legend
                if status in ["Good", "Average", "Poor", "Anomalous data"]:
                    color = get_cell_color(status)
                    if color:
                        try:
                            cell_shading = legend_table.cell(i, 0)._tc.get_or_add_tcPr()
                            cell_fill = cell_shading.get_or_add_shd()
                            cell_fill.fill = color
                        except Exception as e:
                            st.warning(f"Could not set cell background color for legend: {e}")
        
        # Save to bytes buffer
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()
        
    except Exception as e:
        st.error(f"Error creating advanced Word report: {str(e)}")
        st.error(f"Error type: {type(e).__name__}")
        return None

# Main Application
def main():
    # Lambda Function URL
    LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/"
    
    # Title and header
    st.title("🚢 Enhanced Vessel Performance Report Tool")
    st.markdown("Select vessels and generate a comprehensive performance report with improved processing and UI.")
    
    # Performance metrics sidebar
    with st.sidebar:
        st.header("📊 Performance Metrics")
        stats = st.session_state.performance_stats
        st.metric("Total Requests", stats["total_requests"])
        st.metric("Avg Response Time", f"{stats['avg_response_time']:.2f}s")
        
        if st.button("🔄 Reset Stats"):
            st.session_state.performance_stats = {
                'total_requests': 0,
                'total_time': 0,
                'avg_response_time': 0
            }
            st.success("Stats reset!")
        
        if st.button("🗑️ Clear Cache"):
            st.cache_data.clear()
            st.success("Cache cleared!")
    
    # Load vessels
    st.header("1. Select Vessels")
    
    # Load vessels from cache
    with st.spinner("Loading vessels..."):
        try:
            all_vessels = fetch_all_vessels(LAMBDA_FUNCTION_URL)
            st.success(f"✅ Loaded {len(all_vessels)} vessels successfully!")
        except Exception as e:
            st.error(f"❌ Failed to load vessels: {str(e)}")
            all_vessels = []
    
    if all_vessels:
        # Search functionality
        search_query = st.text_input(
            "🔍 Search vessels:",
            value=st.session_state.search_query,
            placeholder="Type to filter vessel names...",
            help="Type to filter the list of vessels below."
        )
        
        if search_query != st.session_state.search_query:
            st.session_state.search_query = search_query
        
        # Filter vessels on client side for responsive search
        filtered_vessels = filter_vessels_client_side(all_vessels, search_query)
        
        st.markdown(f"📊 **{len(filtered_vessels)}** vessels shown (filtered from {len(all_vessels)} total) • **{len(st.session_state.selected_vessels)}** selected")
        
        # Vessel selection with improved UI
        if filtered_vessels:
            with st.container(height=300, border=True):
                cols = st.columns(3)
                for i, vessel in enumerate(filtered_vessels):
                    col_idx = i % 3
                    checkbox_state = cols[col_idx].checkbox(
                        vessel, 
                        value=(vessel in st.session_state.selected_vessels),
                        key=f"checkbox_{vessel}"
                    )
                    if checkbox_state:
                        st.session_state.selected_vessels.add(vessel)
                    else:
                        if vessel in st.session_state.selected_vessels:
                            st.session_state.selected_vessels.remove(vessel)
        else:
            st.info("🔍 No vessels match your search query.")
        
        # Enhanced batch selection controls
        st.subheader("Quick Selection")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            if st.button("✅ Select All Filtered"):
                st.session_state.selected_vessels.update(filtered_vessels)
                st.rerun()
        
        with col2:
            if st.button("❌ Clear All"):
                st.session_state.selected_vessels.clear()
                st.rerun()
        
        with col3:
            if st.button("🔟 First 10"):
                st.session_state.selected_vessels.update(filtered_vessels[:10])
                st.rerun()
        
        with col4:
            if st.button("🔢 First 20"):
                st.session_state.selected_vessels.update(filtered_vessels[:20])
                st.rerun()
        
        with col5:
            if st.button("🎲 Random 10"):
                import random
                random_vessels = random.sample(filtered_vessels, min(10, len(filtered_vessels)))
                st.session_state.selected_vessels.update(random_vessels)
                st.rerun()
        
        selected_vessels_list = list(st.session_state.selected_vessels)
        
        # Show selected vessels summary
        if selected_vessels_list:
            with st.expander(f"📋 Selected Vessels ({len(selected_vessels_list)})", expanded=False):
                for i, vessel in enumerate(sorted(selected_vessels_list), 1):
                    st.write(f"{i}. {vessel}")
    else:
        st.error("❌ Failed to load vessels. Please check your connection and try again.")
        selected_vessels_list = []
    
    # Generate report section
    st.header("2. Generate Enhanced Report")
    
    if selected_vessels_list:
        # Report generation options
        col1, col2 = st.columns(2)
        with col1:
            # The batch_size variable is not directly used in query_report_data as it's hardcoded there.
            # If you want to make it configurable, you'd need to pass it to query_report_data.
            batch_size_ui = st.selectbox(
                "Batch Size (vessels per batch):",
                [5, 10, 15, 20],
                index=1,
                help="Smaller batches = more stable, larger batches = faster"
            )
        
        with col2:
            timeout_setting = st.selectbox(
                "Request Timeout:",
                [15, 30, 45, 60],
                index=1,
                help="Increase if experiencing timeout errors"
            )
        
        # Generate button with enhanced styling
        if st.button("🚀 Generate Enhanced Performance Report", type="primary", use_container_width=True):
            with st.spinner("Generating enhanced report with better progress tracking..."):
                try:
                    start_time = time.time()
                    # Pass the timeout_setting to the invoke_lambda_function_url
                    # The batch_size_ui is not directly used here, as query_report_data has a hardcoded batch_size.
                    # If you want to use batch_size_ui, you'd need to modify query_report_data to accept it.
                    st.session_state.report_data = query_report_data(
                        LAMBDA_FUNCTION_URL, selected_vessels_list
                    )
                    
                    generation_time = time.time() - start_time
                    
                    if not st.session_state.report_data.empty:
                        st.success(f"✅ Report generated successfully in {generation_time:.2f} seconds!")
                        st.balloons()  # Celebration animation
                    else:
                        st.warning("⚠️ No data found for the selected vessels.")
                        
                except Exception as e:
                    st.error(f"❌ Error generating report: {str(e)}")
                    st.session_state.report_data = None
    else:
        st.warning("⚠️ Please select at least one vessel to generate a report.")
        st.info("💡 Use the search box above to find specific vessels, then select them using the checkboxes.")
    
    # Enhanced report display
    if st.session_state.report_data is not None and not st.session_state.report_data.empty:
        st.header("3. 📊 Enhanced Report Results")
        
        # Enhanced report summary with metrics
        st.subheader("📈 Report Summary")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📋 Total Vessels", len(st.session_state.report_data))
        
        with col2:
            # Count vessels with good hull condition (latest month)
            latest_hull_col = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
            if latest_hull_col:
                good_hulls = len(st.session_state.report_data[
                    st.session_state.report_data[latest_hull_col[0]] == "Good"
                ])
                total_with_data = len(st.session_state.report_data[
                    st.session_state.report_data[latest_hull_col[0]] != "N/A"
                ])
                hull_percentage = (good_hulls / total_with_data * 100) if total_with_data > 0 else 0
            else:
                good_hulls = 0
                hull_percentage = 0
            st.metric("🟢 Good Hull Condition", f"{good_hulls} ({hull_percentage:.1f}%)")
        
        with col3:
            # Count vessels with good ME efficiency (latest month)
            latest_me_col = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]
            if latest_me_col:
                good_me = len(st.session_state.report_data[
                    st.session_state.report_data[latest_me_col[0]] == "Good"
                ])
                total_me_data = len(st.session_state.report_data[
                    st.session_state.report_data[latest_me_col[0]] != "N/A"
                ])
                me_percentage = (good_me / total_me_data * 100) if total_me_data > 0 else 0
            else:
                good_me = 0
                me_percentage = 0
            st.metric("🟢 Good ME Efficiency", f"{good_me} ({me_percentage:.1f}%)")
        
        with col4:
            # Average potential fuel saving
            if 'Potential Fuel Saving' in st.session_state.report_data.columns:
                fuel_savings = st.session_state.report_data['Potential Fuel Saving'].apply(
                    lambda x: float(x) if pd.notna(x) and str(x) != 'N/A' else 0
                )
                avg_fuel_saving = fuel_savings.mean()
                total_fuel_saving = fuel_savings.sum()
                st.metric("⛽ Avg Fuel Saving", f"{avg_fuel_saving:.2f} MT/day")
                st.caption(f"Total: {total_fuel_saving:.2f} MT/day")
            else:
                st.metric("⛽ Avg Fuel Saving", "N/A")
        
        # Display styled dataframe with enhanced presentation
        st.subheader("📋 Detailed Report")
        styled_df = st.session_state.report_data.style.apply(
            style_condition_columns, axis=1
        )
        st.dataframe(styled_df, use_container_width=True, height=400)
        
        # Enhanced download section
        st.subheader("📥 Download Options")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"enhanced_vessel_performance_report_{timestamp}.xlsx"
            
            try:
                excel_data = create_excel_download_with_styling(st.session_state.report_data, filename)
                if excel_data:
                    st.download_button(
                        label="📊 Download Excel Report",
                        data=excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"❌ Error creating Excel file: {str(e)}")
        
        with col2:
            # CSV download option
            csv_data = st.session_state.report_data.to_csv(index=False)
            csv_filename = f"vessel_performance_report_{timestamp}.csv"
            st.download_button(
                label="📄 Download CSV Report",
                data=csv_data,
                file_name=csv_filename,
                mime="text/csv",
                use_container_width=True
            )
        
        with col3:
            # Word template download option
            word_filename = f"fleet_performance_report_{timestamp}.docx"
            
            try:
                # Check if template exists
                template_path = "Fleet Performance Template.docx"
                if os.path.exists(template_path):
                    word_data = create_advanced_word_report(st.session_state.report_data, template_path)
                    if word_data:
                        st.download_button(
                            label="📝 Download Word Report",
                            data=word_data,
                            file_name=word_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.error("❌ Failed to create Word report")
                else:
                    st.warning("⚠️ Template file not found")
                    st.caption("Place 'Fleet Performance Template.docx' in repo root")
            except Exception as e:
                st.error(f"❌ Error creating Word file: {str(e)}")
                # Show more detailed error info
                st.caption(f"Error details: {str(e)}")
        
        # Enhanced data insights section
        with st.expander("📈 Data Insights & Analysis", expanded=False):
            tab1, tab2, tab3 = st.tabs(["Hull Condition Analysis", "ME Efficiency Analysis", "Trend Analysis"])
            
            with tab1:
                st.subheader("🛡️ Hull Condition Distribution")
                hull_cols = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
                
                if hull_cols:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Latest month hull condition
                        latest_hull_data = st.session_state.report_data[hull_cols[0]].value_counts()
                        if len(latest_hull_data) > 0 and latest_hull_data.sum() > 0:
                            st.bar_chart(latest_hull_data, use_container_width=True)
                            st.caption(f"Distribution for {hull_cols[0]}")
                        else:
                            st.info("No hull condition data available for chart")
                    
                    with col2:
                        # Hull condition summary table
                        hull_summary = []
                        for col in hull_cols:
                            month = col.replace("Hull Condition ", "")
                            counts = st.session_state.report_data[col].value_counts()
                            hull_summary.append({
                                "Month": month,
                                "Good": counts.get("Good", 0),
                                "Average": counts.get("Average", 0),
                                "Poor": counts.get("Poor", 0),
                                "N/A": counts.get("N/A", 0)
                            })
                        
                        hull_summary_df = pd.DataFrame(hull_summary)
                        st.dataframe(hull_summary_df, use_container_width=True)
                else:
                    st.info("No hull condition data available for analysis")
            
            with tab2:
                st.subheader("⚙️ ME Efficiency Distribution")
                me_cols = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]
                
                if me_cols:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Latest month ME efficiency
                        latest_me_data = st.session_state.report_data[me_cols[0]].value_counts()
                        if len(latest_me_data) > 0 and latest_me_data.sum() > 0:
                            st.bar_chart(latest_me_data, use_container_width=True)
                            st.caption(f"Distribution for {me_cols[0]}")
                        else:
                            st.info("No ME efficiency data available for chart")
                    
                    with col2:
                        # ME efficiency summary table
                        me_summary = []
                        for col in me_cols:
                            month = col.replace("ME Efficiency ", "")
                            counts = st.session_state.report_data[col].value_counts()
                            me_summary.append({
                                "Month": month,
                                "Good": counts.get("Good", 0),
                                "Average": counts.get("Average", 0),
                                "Poor": counts.get("Poor", 0),
                                "Anomalous": counts.get("Anomalous data", 0),
                                "N/A": counts.get("N/A", 0)
                            })
                        
                        me_summary_df = pd.DataFrame(me_summary)
                        st.dataframe(me_summary_df, use_container_width=True)
                else:
                    st.info("No ME efficiency data available for analysis")
            
            with tab3:
                st.subheader("📊 Performance Trends")
                
                # Combined trend analysis with better data validation
                hull_cols = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
                me_cols = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]
                
                if len(hull_cols) >= 2:
                    st.write("**Hull Condition Trends (% Good)**")
                    hull_trend_data = []
                    has_valid_data = False
                    
                    for col in hull_cols:
                        month = col.replace("Hull Condition ", "")
                        total_with_data = len(st.session_state.report_data[st.session_state.report_data[col] != "N/A"])
                        good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                        
                        if total_with_data > 0:
                            percentage = (good_count / total_with_data * 100)
                            hull_trend_data.append({"Month": month, "Good %": percentage})
                            has_valid_data = True
                        else:
                            hull_trend_data.append({"Month": month, "Good %": 0})
                    
                    if has_valid_data and hull_trend_data:
                        hull_trend_df = pd.DataFrame(hull_trend_data)
                        # Only show chart if we have non-zero data
                        if hull_trend_df["Good %"].sum() > 0:
                            st.line_chart(hull_trend_df.set_index("Month"), use_container_width=True)
                        else:
                            st.info("No hull condition data available for trend analysis")
                    else:
                        st.info("No hull condition data available for trend analysis")
                else:
                    st.info("Need at least 2 months of hull data for trend analysis")
                
                if len(me_cols) >= 2:
                    st.write("**ME Efficiency Trends (% Good)**")
                    me_trend_data = []
                    has_valid_me_data = False
                    
                    for col in me_cols:
                        month = col.replace("ME Efficiency ", "")
                        total_with_data = len(st.session_state.report_data[st.session_state.report_data[col] != "N/A"])
                        good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                        
                        if total_with_data > 0:
                            percentage = (good_count / total_with_data * 100)
                            me_trend_data.append({"Month": month, "Good %": percentage})
                            has_valid_me_data = True
                        else:
                            me_trend_data.append({"Month": month, "Good %": 0})
                    
                    if has_valid_me_data and me_trend_data:
                        me_trend_df = pd.DataFrame(me_trend_data)
                        # Only show chart if we have non-zero data
                        if me_trend_df["Good %"].sum() > 0:
                            st.line_chart(me_trend_df.set_index("Month"), use_container_width=True)
                        else:
                            st.info("No ME efficiency data available for trend analysis")
                    else:
                        st.info("No ME efficiency data available for trend analysis")
                else:
                    st.info("Need at least 2 months of ME efficiency data for trend analysis")

    elif st.session_state.report_data is not None and st.session_state.report_data.empty:
        st.info("ℹ️ No data found for the selected vessels.")
        st.write("This could happen if:")
        st.write("- The selected vessels don't have data in the database")
        st.write("- There's a connectivity issue with the database")
        st.write("- The data hasn't been updated recently")
    
    # Enhanced instructions
    with st.expander("📖 Enhanced Features & Instructions", expanded=False):
        st.markdown("""
        ### 🚀 Enhanced Features:
        
        **🔍 Improved Search & Selection:**
        - Real-time vessel filtering as you type
        - Quick selection buttons (All, First 10/20, Random 10, Clear)
        - Selected vessels summary with expandable list
        - Smart client-side filtering for responsive UI
        
        **📊 Better Data Processing:**
        - Enhanced progress tracking with visual progress bars
        - Improved error handling and user feedback
        - Performance metrics tracking in sidebar
        - Configurable batch sizes and timeouts
        - Success animations and better visual feedback
        
        **📈 Advanced Analytics:**
        - Summary metrics with percentages
        - Hull condition and ME efficiency distribution charts
        - Multi-month trend analysis
        - Tabbed insights section for better organization
        - Total and average fuel saving calculations
        
        **📥 Enhanced Downloads:**
        - Both Excel and CSV download options
        - Timestamped filenames
        - Styled Excel reports with color coding
        - Better error handling for file generation
        
        ### 📋 How to Use:
        
        1. **🔍 Search & Filter**: Type in the search box to find specific vessels
        2. **✅ Select Vessels**: Use checkboxes or quick selection buttons
        3. **⚙️ Configure**: Choose batch size and timeout based on your needs
        4. **🚀 Generate**: Click the generate button for enhanced processing
        5. **📊 Analyze**: Review metrics, charts, and trends in the results
        6. **📥 Download**: Export your report in Excel or CSV format
        
        ### 📊 Report Columns:
        
        **🛡️ Hull Condition** (Multiple months):
        - 🟢 **Good**: < 15% power loss (Green)
        - 🟡 **Average**: 15-25% power loss (Yellow)
        - 🔴 **Poor**: > 25% power loss (Red)
        
        **⚙️ ME Efficiency** (Multiple months):
        - ⚪ **Anomalous data**: < 160 SFOC (Gray)
        - 🟢 **Good**: 160-180 SFOC (Green)
        - 🟡 **Average**: 180-190 SFOC (Yellow)
        - 🔴 **Poor**: > 190 SFOC (Red)
        
        **📊 Additional Metrics:**
        - ⛽ **Potential Fuel Saving**: Excess consumption (MT/day)
        - 📈 **YTD CII**: Carbon Intensity Indicator rating
        - 💬 **Comments**: Space for additional notes
        
        ### 💡 Performance Tips:
        
        - Use smaller batch sizes (5-10) for more stable processing
        - Increase timeout if you experience timeout errors
        - Clear cache occasionally to ensure fresh data
        - Use search to narrow down vessels before bulk selection
        - Monitor performance metrics in the sidebar
        """)
    
    # Enhanced footer
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("*Enhanced with improved UI & analytics*")
    with col2:
        st.markdown("*Built with Streamlit 🎈 and Python*")
    with col3:
        stats = st.session_state.performance_stats
        if stats['total_requests'] > 0:
            st.markdown(f"*{stats['total_requests']} requests • {stats['avg_response_time']:.2f}s avg*}}")

if __name__ == "__main__":
    main()
