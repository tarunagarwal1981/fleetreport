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

# --- Lambda Invocation Helper ---
def invoke_lambda_function_url(lambda_url, payload):
    """Invoke Lambda function via its Function URL using HTTP POST."""
    try:
        headers = {'Content-Type': 'application/json'}
        json_payload = json.dumps(payload)
        
        response = requests.post(
            lambda_url, 
            headers=headers, 
            data=json_payload,
            timeout=30
        )
        
        if response.status_code != 200:
            st.error(f"HTTP error: {response.status_code} {response.reason} for url: {lambda_url}")
            return None
            
        return response.json()
            
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

# --- Data Fetching Functions ---
@st.cache_data(ttl=3600)
def fetch_all_vessels(lambda_url):
    """Fetch vessel names from Lambda function with a limit of 1200."""
    query = "SELECT vessel_name FROM vessel_particulars ORDER BY vessel_name LIMIT 1200"
    
    with st.spinner("Loading vessels..."):
        result = invoke_lambda_function_url(lambda_url, {"sql_query": query})
        
        if result:
            extracted_vessel_names = []
            for item in result:
                if isinstance(item, dict) and 'vessel_name' in item:
                    extracted_vessel_names.append(item['vessel_name'])
                elif isinstance(item, str):
                    extracted_vessel_names.append(item)
            extracted_vessel_names.sort()
            st.success(f"Loaded {len(extracted_vessel_names)} vessels.")
            return extracted_vessel_names
        
        st.error("Failed to load vessel names.")
        return []

def query_report_data(lambda_url, vessel_names):
    """Fetch hull roughness power loss, ME SFOC, Potential Fuel Saving, YTD CII, and historical hull conditions for selected vessels and process for report."""
    if not vessel_names:
        return pd.DataFrame()

    # Calculate dates for previous month and previous-to-previous month
    today = datetime.now()
    
    # Last day of previous month for Hull Condition
    first_day_current_month = today.replace(day=1)
    last_day_previous_month_hull = first_day_current_month - timedelta(days=1)
    prev_month_hull_str = last_day_previous_month_hull.strftime("%Y-%m-%d")
    prev_month_hull_col_name = f"Hull Condition {last_day_previous_month_hull.strftime('%b %y')}"

    # Last day of previous-to-previous month for Hull Condition
    first_day_previous_month_hull = last_day_previous_month_hull.replace(day=1)
    last_day_prev_prev_month_hull = first_day_previous_month_hull - timedelta(days=1)
    prev_prev_month_hull_str = last_day_prev_prev_month_hull.strftime("%Y-%m-%d")
    prev_prev_month_hull_col_name = f"Hull Condition {last_day_prev_prev_month_hull.strftime('%b %y')}"

    # Dates for ME SFOC (average of the entire month)
    # Previous month (already in use)
    # Previous-to-previous month
    first_day_prev_prev_month_me = today.replace(day=1) - timedelta(days=1) # This gets us to last day of prev month
    first_day_prev_prev_month_me = first_day_prev_prev_month_me.replace(day=1) - timedelta(days=1) # This gets us to last day of prev-prev month
    first_day_prev_prev_month_me = first_day_prev_prev_month_me.replace(day=1) # This gets us to first day of prev-prev month
    
    end_day_prev_prev_month_me = first_day_prev_prev_month_me + timedelta(days=32) # Go past end of month
    end_day_prev_prev_month_me = end_day_prev_prev_month_me.replace(day=1) - timedelta(days=1) # Get last day of prev-prev month

    prev_month_me_col_name = f"ME Efficiency {last_day_previous_month_hull.strftime('%b %y')}" # Reusing prev month date for naming
    prev_prev_month_me_col_name = f"ME Efficiency {end_day_prev_prev_month_me.strftime('%b %y')}"


    # Process vessels in smaller batches to avoid timeout/size issues
    batch_size = 10
    all_hull_data = [] # This will no longer be used for the main "Hull Condition" column
    all_me_data = [] # This is for the previous month ME SFOC
    all_fuel_saving_data = []
    all_cii_data = []
    all_prev_month_hull_data = []
    all_prev_prev_month_hull_data = []
    all_prev_prev_month_me_data = [] # New list for previous-to-previous month ME SFOC
    
    total_batches = (len(vessel_names) + batch_size - 1) // batch_size
    
    for i in range(0, len(vessel_names), batch_size):
        batch_vessels = vessel_names[i:i+batch_size]
        batch_num = i//batch_size + 1
        st.info(f"Processing batch {batch_num} of {total_batches} ({len(batch_vessels)} vessels)")
        
        quoted_vessel_names = [f"'{name}'" for name in batch_vessels]
        vessel_names_list_str = ", ".join(quoted_vessel_names)

        # --- Query 1: Hull Roughness Power Loss (Previous Month) - Single row per vessel per day ---
        sql_query_prev_month_hull = f"""
SELECT
    vessel_name,
    hull_rough_power_loss_pct_ed
FROM (
    SELECT
        vessel_name,
        hull_rough_power_loss_pct_ed,
        ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM
        hull_performance_six_months_daily
    WHERE
        vessel_name IN ({vessel_names_list_str})
        AND CAST(updated_ts AS DATE) = '{prev_month_hull_str}'
) AS subquery
WHERE rn = 1
"""
        st.info(f"Fetching Hull Roughness data for {last_day_previous_month_hull.strftime('%b %y')} for batch...")
        prev_month_hull_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_prev_month_hull})
        if prev_month_hull_result:
            all_prev_month_hull_data.extend(prev_month_hull_result)

        # --- Query 2: Hull Roughness Power Loss (Previous-to-Previous Month) - Single row per vessel per day ---
        sql_query_prev_prev_month_hull = f"""
SELECT
    vessel_name,
    hull_rough_power_loss_pct_ed
FROM (
    SELECT
        vessel_name,
        hull_rough_power_loss_pct_ed,
        ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM
        hull_performance_six_months_daily
    WHERE
        vessel_name IN ({vessel_names_list_str})
        AND CAST(updated_ts AS DATE) = '{prev_prev_month_hull_str}'
) AS subquery
WHERE rn = 1
"""
        st.info(f"Fetching Hull Roughness data for {last_day_prev_prev_month_hull.strftime('%b %y')} for batch...")
        prev_prev_month_hull_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_prev_prev_month_hull})
        if prev_prev_month_hull_result:
            all_prev_prev_month_hull_data.extend(prev_prev_month_hull_result)

        # --- Query 3: ME SFOC (Previous Month) ---
        sql_query_me = f"""
SELECT
    vp.vessel_name,
    AVG(vps.me_sfoc) AS avg_me_sfoc
FROM
    vessel_performance_summary vps
JOIN
    vessel_particulars vp
ON
    CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE
    vp.vessel_name IN ({vessel_names_list_str})
    AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
    AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE)
GROUP BY
    vp.vessel_name
"""
        
        st.info(f"Fetching ME SFOC data for {last_day_previous_month_hull.strftime('%b %y')} for batch...")
        
        me_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_me})

        if me_result:
            all_me_data.extend(me_result)
        
        # --- Query 4: ME SFOC (Previous-to-Previous Month) ---
        sql_query_prev_prev_month_me = f"""
SELECT
    vp.vessel_name,
    AVG(vps.me_sfoc) AS avg_me_sfoc
FROM
    vessel_performance_summary vps
JOIN
    vessel_particulars vp
ON
    CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE
    vp.vessel_name IN ({vessel_names_list_str})
    AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
    AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
GROUP BY
    vp.vessel_name
"""
        st.info(f"Fetching ME SFOC data for {end_day_prev_prev_month_me.strftime('%b %y')} for batch...")
        prev_prev_month_me_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_prev_prev_month_me})
        if prev_prev_month_me_result:
            all_prev_prev_month_me_data.extend(prev_prev_month_me_result)


        # --- Query 5: Potential Fuel Saving ---
        sql_query_fuel_saving = f"SELECT vessel_name, hull_rough_excess_consumption_mt_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str})"
        
        st.info(f"Fetching Potential Fuel Saving data for batch...")
        
        fuel_saving_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_fuel_saving})

        if fuel_saving_result:
            all_fuel_saving_data.extend(fuel_saving_result)

        # --- Query 6: YTD CII ---
        sql_query_cii = f"""
SELECT
    vp.vessel_name,
    cy.cii_rating
FROM
    vessel_particulars vp
JOIN
    cii_ytd cy
ON
    CAST(vp.vessel_imo AS TEXT) = CAST(cy.vessel_imo AS TEXT)
WHERE
    vp.vessel_name IN ({vessel_names_list_str})
"""
        st.info(f"Fetching YTD CII data for batch...")
        
        cii_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_cii})

        if cii_result:
            all_cii_data.extend(cii_result)


    # Process all collected data
    # Hull Data (Previous Month)
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
    else:
        st.warning(f"No hull roughness data found for {last_day_previous_month_hull.strftime('%b %y')}.")

    # Hull Data (Previous-to-Previous Month)
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
    else:
        st.warning(f"No hull roughness data found for {last_day_prev_prev_month_hull.strftime('%b %y')}.")

    # ME Data (Previous Month)
    df_me = pd.DataFrame()
    if all_me_data:
        try:
            df_me = pd.DataFrame(all_me_data)
            if 'avg_me_sfoc' in df_me.columns:
                df_me = df_me.rename(columns={'avg_me_sfoc': f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'})
            else:
                df_me[f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'] = pd.NA
            df_me = df_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous month ME data: {str(e)}")
            df_me = pd.DataFrame()
    else:
        st.error("Failed to retrieve previous month ME SFOC data.")

    # ME Data (Previous-to-Previous Month)
    df_prev_prev_month_me = pd.DataFrame()
    if all_prev_prev_month_me_data:
        try:
            df_prev_prev_month_me = pd.DataFrame(all_prev_prev_month_me_data)
            if 'avg_me_sfoc' in df_prev_prev_month_me.columns:
                df_prev_prev_month_me = df_prev_prev_month_me.rename(columns={'avg_me_sfoc': f'ME SFOC {end_day_prev_prev_month_me.strftime("%b %y")}'})
            else:
                df_prev_prev_month_me[f'ME SFOC {end_day_prev_prev_month_me.strftime("%b %y")}'] = pd.NA
            df_prev_prev_month_me = df_prev_prev_month_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing previous-to-previous month ME data: {str(e)}")
            df_prev_prev_month_me = pd.DataFrame()
    else:
        st.error("Failed to retrieve previous-to-previous month ME SFOC data.")


    df_fuel_saving = pd.DataFrame()
    if all_fuel_saving_data:
        try:
            df_fuel_saving = pd.DataFrame(all_fuel_saving_data)
            if 'hull_rough_excess_consumption_mt_ed' in df_fuel_saving.columns:
                df_fuel_saving = df_fuel_saving.rename(columns={'hull_rough_excess_consumption_mt_ed': 'Potential Fuel Saving'})
                # Apply the capping logic: if > 5, set to 4.9; if < 0, set to 0
                df_fuel_saving['Potential Fuel Saving'] = df_fuel_saving['Potential Fuel Saving'].apply(
                    lambda x: 4.9 if pd.notna(x) and x > 5 else (0.0 if pd.notna(x) and x < 0 else x)
                )
            else:
                df_fuel_saving['Potential Fuel Saving'] = pd.NA
            df_fuel_saving = df_fuel_saving.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing fuel saving data: {str(e)}")
            df_fuel_saving = pd.DataFrame()
    else:
        st.error("Failed to retrieve Potential Fuel Saving data.")

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
    else:
        st.error("Failed to retrieve YTD CII data.")


    # --- Merge DataFrames ---
    df_final = pd.DataFrame({'Vessel Name': list(vessel_names)})

    # Merge historical hull data
    if not df_prev_month_hull.empty:
        df_final = pd.merge(df_final, df_prev_month_hull, on='Vessel Name', how='left')

    if not df_prev_prev_month_hull.empty:
        df_final = pd.merge(df_final, df_prev_prev_month_hull, on='Vessel Name', how='left')

    # Merge ME data
    if not df_me.empty:
        df_final = pd.merge(df_final, df_me, on='Vessel Name', how='left')
            
    if not df_prev_prev_month_me.empty:
        df_final = pd.merge(df_final, df_prev_prev_month_me, on='Vessel Name', how='left')

    # Merge other data
    if not df_fuel_saving.empty:
        df_final = pd.merge(df_final, df_fuel_saving, on='Vessel Name', how='left')

    if not df_cii.empty:
        df_final = pd.merge(df_final, df_cii, on='Vessel Name', how='left')

    if df_final.empty:
        return pd.DataFrame()

    # --- Post-merge processing for final report ---
    # Add S. No. column
    df_final.insert(0, 'S. No.', range(1, 1 + len(df_final)))
    
    # Hull Condition logic (reusable function)
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


    # Add ME Efficiency column logic (reusable function)
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
    
    # Apply ME Efficiency to previous month and previous-to-previous month columns
    if f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}' in df_final.columns:
        df_final[prev_month_me_col_name] = df_final[f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}'].apply(get_me_efficiency)
    else:
        df_final[prev_month_me_col_name] = "N/A"

    if f'ME SFOC {end_day_prev_prev_month_me.strftime("%b %y")}' in df_final.columns:
        df_final[prev_prev_month_me_col_name] = df_final[f'ME SFOC {end_day_prev_prev_month_me.strftime("%b %y")}'].apply(get_me_efficiency)
    else:
        df_final[prev_prev_month_me_col_name] = "N/A"

    # Add empty Comments column
    df_final['Comments'] = ""

    # Define the desired order of columns
    desired_columns_order = [
        'S. No.', 
        'Vessel Name', 
        prev_month_hull_col_name, # Previous Month Hull Condition
        prev_prev_month_hull_col_name, # Previous-to-Previous Month Hull Condition
        prev_month_me_col_name, # Previous Month ME Efficiency
        prev_prev_month_me_col_name, # Previous-to-Previous Month ME Efficiency
        'Potential Fuel Saving',
        'YTD CII',
        'Comments',
        f'Hull Roughness Power Loss % {last_day_previous_month_hull.strftime("%b %y")}', # Raw data for previous month
        f'Hull Roughness Power Loss % {last_day_prev_prev_month_hull.strftime("%b %y")}', # Raw data for previous-to-previous month
        f'ME SFOC {last_day_previous_month_hull.strftime("%b %y")}', # Raw data for previous month ME
        f'ME SFOC {end_day_prev_prev_month_me.strftime("%b %y")}' # Raw data for previous-to-previous month ME
    ]
    
    # Filter df_final to only include columns that exist and are in the desired order
    existing_and_ordered_columns = [col for col in desired_columns_order if col in df_final.columns]
    df_final = df_final[existing_and_ordered_columns]

    st.success("Report data retrieved and processed successfully!")
    return df_final

# --- Styling for Streamlit DataFrame ---
def style_condition_columns(row):
    styles = [''] * len(row)
    
    # List of all hull condition columns to style
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
    
    # List of all ME Efficiency columns to style
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
            
    # YTD CII styling (no specific color, but ensure it's handled)
    if 'YTD CII' in row.index:
        # No specific styling needed for CII, but including it here ensures it's processed
        pass 
            
    return styles

# --- Excel Export Function ---
def create_excel_download_with_styling(df, filename):
    """Convert DataFrame to Excel format with styling and auto-width using openpyxl."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Vessel Report"

    # Write headers
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)

    # Write data and apply cell styling
    for row_idx, row_data in df.iterrows():
        for col_idx, (col_name, cell_value) in enumerate(row_data.items(), 1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value)
            
            # Apply styling for 'Hull Condition' and 'ME Efficiency' columns
            if 'Hull Condition' in col_name or 'ME Efficiency' in col_name: # Check if it's any hull condition or ME Efficiency column
                if cell_value == "Good":
                    cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                elif cell_value == "Average":
                    cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                elif cell_value == "Poor":
                    cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                elif cell_value == "Anomalous data":
                    cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.font = Font(color="000000")
            # No specific styling for 'YTD CII' but ensure it's written
            elif col_name == 'YTD CII':
                cell.alignment = Alignment(horizontal='center') # Example: center align CII values

    # Set column widths
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

# --- Main App Layout ---
st.title("🚢 Vessel Performance Report Tool")
st.markdown("Select vessels and generate a report on their excess power.")

# --- Lambda Function URL ---
LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/" 

# --- Load Vessels Automatically ---
if not st.session_state.vessels:
    st.session_state.vessels = fetch_all_vessels(LAMBDA_FUNCTION_URL)

# --- Main Content Area ---
st.header("1. Select Vessels")

if st.session_state.vessels:
    # Search and filter vessels
    st.session_state.search_query = st.text_input(
        "Search vessels:",
        value=st.session_state.search_query,
        placeholder="Type to filter vessel names...",
        help="Type to filter the list of vessels below."
    )

    filtered_vessels = [
        v for v in st.session_state.vessels 
        if st.session_state.search_query.lower() in v.lower()
    ]

    st.markdown(f"📊 {len(st.session_state.vessels)} vessels available. {len(st.session_state.selected_vessels)} selected.")

    # Use a container for scrollable list of checkboxes
    with st.container(height=300, border=True):
        if filtered_vessels:
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
            st.info("No vessels match your search query.")
    
    selected_vessels_list = list(st.session_state.selected_vessels)
else:
    st.error("Failed to load vessels. Please refresh the page to try again.")
    selected_vessels_list = []

st.header("2. Generate Report")

if selected_vessels_list:
    if st.button("🚀 Generate Excess Power Report", type="primary"):
        with st.spinner("Generating report..."):
            st.session_state.report_data = query_report_data(
                LAMBDA_FUNCTION_URL, 
                selected_vessels_list
            )
else:
    st.warning("Please select at least one vessel to generate a report.")

# --- Display Report ---
if st.session_state.report_data is not None and not st.session_state.report_data.empty:
    st.header("3. Report Results")
    
    styled_df = st.session_state.report_data.style.apply(
        style_condition_columns, axis=1
    )
    st.dataframe(styled_df, use_container_width=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"excess_power_report_{timestamp}.xlsx"

    excel_data = create_excel_download_with_styling(st.session_state.report_data, filename)

    if excel_data:
        st.download_button(
            label="📥 Download Report as Excel",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif st.session_state.report_data is not None and st.session_state.report_data.empty:
    st.info("No data found for the selected vessels or query.")

# --- Instructions ---
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Step-by-step Guide:
    
    1. **Select Vessels**: Use the search bar to filter vessels, then check the boxes to select them.
    2. **Generate Report**: Click "Generate Excess Power Report" to fetch the data and display it.
    3. **Download**: If data is available, a download button will appear to get the report as an Excel file.
    
    ### Report Columns:
    
    - **Hull Condition (Previous Month, Previous-to-Previous Month)**: Based on Hull Roughness Power Loss %
      - Good: < 15%
      - Average: 15-25%
      - Poor: > 25%
    
    - **ME Efficiency (Previous Month, Previous-to-Previous Month)**: Based on ME SFOC
      - Anomalous data: < 160
      - Good: 160-180
      - Average: 180-190
      - Poor: > 190
    
    - **Potential Fuel Saving**: Excess fuel consumption in metric tons per day
      - Capped at 4.9 if > 5
      - Set to 0 if < 0
    
    - **YTD CII**: The Carbon Intensity Indicator rating for the vessel.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈 and Python*")
