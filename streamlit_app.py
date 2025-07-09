import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Color
import time
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

# Page configuration with improved styling
st.set_page_config(
    page_title="Vessel Performance Report Tool",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
  
    .section-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        color: white;
        font-weight: bold;
    }
  
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2E86AB;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 0.5rem 0;
    }
  
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        transition: all 0.3s ease;
    }
  
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
  
    .success-box {
        background: linear-gradient(90deg, #56ab2f 0%, #a8e6cf 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        margin: 1rem 0;
    }
  
    .info-box {
        background: linear-gradient(90deg, #3498db 0%, #85c1e9 100%);
        padding: 1rem;
        border-radius: 8px;
        color: white;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'vessels' not in st.session_state:
    st.session_state.vessels = []

if 'selected_vessels' not in st.session_state or not isinstance(st.session_state.selected_vessels, set):
    st.session_state.selected_vessels = set()

if 'report_data' not in st.session_state:
    st.session_state.report_data = None

if 'search_query' not in st.session_state:
    st.session_state.search_query = ""

if 'report_months' not in st.session_state:
    st.session_state.report_months = 2

# Enhanced Lambda Invocation Helper
def invoke_lambda_function_url(lambda_url, payload, timeout=60):
    """Invoke Lambda function via its Function URL using HTTP POST with performance tracking."""
    try:
        start_time = time.time()
        headers = {'Content-Type': 'application/json'}
        json_payload = json.dumps(payload)

        response = requests.post(
            lambda_url,
            headers=headers,
            data=json_payload,
            timeout=timeout
        )

        response_time = time.time() - start_time

        if response.status_code != 200:
            st.error(f"HTTP error: {response.status_code} {response.reason} for url: {lambda_url}")
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

def query_report_data(lambda_url, vessel_names, num_months):
    """Enhanced version of the original query_report_data function with better progress tracking and fixed date logic."""
    if not vessel_names:
        return pd.DataFrame()

    today = datetime.now()
    # Get the first day of the current month
    first_day_current_month = today.replace(day=1)

    # Prepare date strings and column names based on num_months
    hull_dates_info = []
    me_dates_info = []

    for i in range(num_months):
        # Calculate the target month (going backwards from current month)
        # For i=0: current month - 1 (previous month)
        # For i=1: current month - 2 (two months ago)
        months_back = i + 1
        
        # Get the first day of the target month
        if first_day_current_month.month <= months_back:
            # Need to go back to previous year
            target_year = first_day_current_month.year - 1
            target_month = 12 - (months_back - first_day_current_month.month)
        else:
            target_year = first_day_current_month.year
            target_month = first_day_current_month.month - months_back
        
        target_month_first_day = datetime(target_year, target_month, 1)
        
        # Hull Condition Dates (last day of the target month)
        if target_month == 12:
            next_month_first = datetime(target_year + 1, 1, 1)
        else:
            next_month_first = datetime(target_year, target_month + 1, 1)
        
        target_month_last_day = next_month_first - timedelta(days=1)

        hull_date_str = target_month_last_day.strftime("%Y-%m-%d")
        hull_col_name = f"Hull Condition {target_month_last_day.strftime('%b %y')}"
        hull_power_loss_col_name = f"Hull Roughness Power Loss % {target_month_last_day.strftime('%b %y')}"
        hull_dates_info.append({
            'date_str': hull_date_str,
            'col_name': hull_col_name,
            'power_loss_col_name': hull_power_loss_col_name,
            'months_back': months_back
        })

        # ME SFOC Dates (average of the entire target month)
        me_col_name = f"ME Efficiency {target_month_last_day.strftime('%b %y')}"
        me_dates_info.append({
            'col_name': me_col_name,
            'months_back': months_back,
            'target_month_first': target_month_first_day
        })

    # Process vessels in smaller batches with enhanced progress tracking
    batch_size = 10
    all_fuel_saving_data = []
    all_cii_data = []
    all_hull_data_by_month = {info['power_loss_col_name']: [] for info in hull_dates_info}
    all_me_data_by_month = {info['col_name']: [] for info in me_dates_info}

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

        batch_queries = []

        # Hull Roughness queries
        for hull_info in hull_dates_info:
            batch_queries.append((hull_info['power_loss_col_name'], f"""
SELECT vessel_name, hull_rough_power_loss_pct_ed
FROM (
    SELECT vessel_name, hull_rough_power_loss_pct_ed,
           ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
    FROM hull_performance_six_months_daily
    WHERE vessel_name IN ({vessel_names_list_str})
    AND CAST(updated_ts AS DATE) = '{hull_info['date_str']}'
) AS subquery
WHERE rn = 1
""", all_hull_data_by_month[hull_info['power_loss_col_name']]))

        # ME SFOC queries with corrected date logic
        for me_info in me_dates_info:
            batch_queries.append((me_info['col_name'], f"""
SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
FROM vessel_performance_summary vps
JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{me_info['months_back']} month')
AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{me_info['months_back'] - 1} month')
GROUP BY vp.vessel_name
""", all_me_data_by_month[me_info['col_name']]))

        # Fixed queries
        batch_queries.append(("Potential Fuel Saving", f"""
SELECT vessel_name, hull_rough_excess_consumption_mt_ed
FROM hull_performance_six_months
WHERE vessel_name IN ({vessel_names_list_str})
""", all_fuel_saving_data))

        batch_queries.append(("YTD CII", f"""
SELECT vp.vessel_name, cy.cii_rating
FROM vessel_particulars vp
JOIN cii_ytd cy ON CAST(vp.vessel_imo AS TEXT) = CAST(cy.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
""", all_cii_data))

        # Execute each query
        for query_name, query, data_list in batch_queries:
            with st.spinner(f"Fetching {query_name} data..."):
                result = invoke_lambda_function_url(lambda_url, {"sql_query": query})
                if result:
                    data_list.extend(result)

    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()

    # Process all collected data
    df_final = pd.DataFrame({'Vessel Name': list(vessel_names)})

    # Hull Data processing and merging
    for hull_info in hull_dates_info:
        df_hull = pd.DataFrame()
        if all_hull_data_by_month[hull_info['power_loss_col_name']]:
            try:
                df_hull = pd.DataFrame(all_hull_data_by_month[hull_info['power_loss_col_name']])
                if 'hull_rough_power_loss_pct_ed' in df_hull.columns:
                    df_hull = df_hull.rename(columns={'hull_rough_power_loss_pct_ed': hull_info['power_loss_col_name']})
                else:
                    df_hull[hull_info['power_loss_col_name']] = pd.NA
                df_hull = df_hull.rename(columns={'vessel_name': 'Vessel Name'})
            except Exception as e:
                st.error(f"Error processing {hull_info['col_name']} data: {str(e)}")
                df_hull = pd.DataFrame()
        if not df_hull.empty:
            df_final = pd.merge(df_final, df_hull, on='Vessel Name', how='left')

    # ME Data processing and merging
    for me_info in me_dates_info:
        df_me = pd.DataFrame()
        if all_me_data_by_month[me_info['col_name']]:
            try:
                df_me = pd.DataFrame(all_me_data_by_month[me_info['col_name']])
                if 'avg_me_sfoc' in df_me.columns:
                    df_me = df_me.rename(columns={'avg_me_sfoc': me_info['col_name']})
                else:
                    df_me[me_info['col_name']] = pd.NA
                df_me = df_me.rename(columns={'vessel_name': 'Vessel Name'})
            except Exception as e:
                st.error(f"Error processing {me_info['col_name']} data: {str(e)}")
                df_me = pd.DataFrame()
        if not df_me.empty:
            df_final = pd.merge(df_final, df_me, on='Vessel Name', how='left')

    # Fuel saving and CII data processing
    df_fuel_saving = pd.DataFrame()
    if all_fuel_saving_data:
        try:
            df_fuel_saving = pd.DataFrame(all_fuel_saving_data)
            if 'hull_rough_excess_consumption_mt_ed' in df_fuel_saving.columns:
                df_fuel_saving = df_fuel_saving.rename(columns={'hull_rough_excess_consumption_mt_ed': 'Potential Fuel Saving (MT/Day)'})
                df_fuel_saving['Potential Fuel Saving (MT/Day)'] = df_fuel_saving['Potential Fuel Saving (MT/Day)'].apply(
                    lambda x: 4.9 if pd.notna(x) and x > 5 else (0.0 if pd.notna(x) and x < 0 else x)
                ).round(2)
            else:
                df_fuel_saving['Potential Fuel Saving (MT/Day)'] = pd.NA
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
    for hull_info in hull_dates_info:
        if hull_info['power_loss_col_name'] in df_final.columns:
            df_final[hull_info['col_name']] = df_final[hull_info['power_loss_col_name']].apply(get_hull_condition)
        else:
            df_final[hull_info['col_name']] = "N/A"

    # ME Efficiency logic and comments
    def get_me_efficiency_and_comment(value):
        if pd.isna(value):
            return "N/A", ""
        if value < 160:
            return "Anomalous data", f"SFOC value is {value:.1f} g/kWh, unusually low, indicating potential data anomaly."
        elif value < 180:
            return "Good", ""
        elif 180 <= value <= 190:
            return "Average", ""
        else:
            return "Poor", ""

    # Initialize a temporary column for comments
    df_final['temp_comments_list'] = [[] for _ in range(len(df_final))]

    # Apply ME Efficiency and populate comments
    for me_info in me_dates_info:
        if me_info['col_name'] in df_final.columns:
            # Apply the function to get both status and comment
            df_final[[me_info['col_name'], 'current_me_comment']] = df_final[me_info['col_name']].apply(
                lambda x: pd.Series(get_me_efficiency_and_comment(x))
            )
            # Append comments for anomalous data to the temporary list
            df_final['temp_comments_list'] = df_final.apply(
                lambda row: row['temp_comments_list'] + [f"ME Efficiency ({me_info['col_name'].split(' ')[-2:]}): {row['current_me_comment']}"] if row['current_me_comment'] else row['temp_comments_list'],
                axis=1
            )
            df_final = df_final.drop(columns=['current_me_comment'])

    # Join all collected comments into the final 'Comments' column
    df_final['Comments'] = df_final['temp_comments_list'].apply(lambda x: " ".join(x).strip())
    df_final = df_final.drop(columns=['temp_comments_list'])

    # Define the desired order of columns (hide Potential Fuel Saving)
    desired_columns_order = ['S. No.', 'Vessel Name']
    for hull_info in hull_dates_info:
        desired_columns_order.append(hull_info['col_name'])
    for me_info in me_dates_info:
        desired_columns_order.append(me_info['col_name'])
    # desired_columns_order.append('Potential Fuel Saving (MT/Day)')  # Hidden
    desired_columns_order.extend(['YTD CII', 'Comments'])

    # Filter df_final to only include columns that exist and are in the desired order
    existing_and_ordered_columns = [col for col in desired_columns_order if col in df_final.columns]
    df_final = df_final[existing_and_ordered_columns]

    st.success("✅ Enhanced report data retrieved and processed successfully!")
    return df_final

# Enhanced styling function with CII color coding
def style_condition_columns(row):
    """Apply styling to condition columns including CII rating text color."""
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
                styles[row.index.get_loc(col_name)] = 'background-color: #FF0000; color: white;'  # Red background

    # Style YTD CII column with text color only (updated A and B colors)
    if 'YTD CII' in row.index:
        cii_val = str(row['YTD CII']).upper() if pd.notna(row['YTD CII']) else "N/A"
        if cii_val == "A":
            styles[row.index.get_loc('YTD CII')] = 'color: #006400; font-weight: bold;'  # Dark green
        elif cii_val == "B":
            styles[row.index.get_loc('YTD CII')] = 'color: #90EE90; font-weight: bold;'  # Light green
        elif cii_val == "C":
            styles[row.index.get_loc('YTD CII')] = 'color: #FFD700; font-weight: bold;'  # Yellow
        elif cii_val == "D":
            styles[row.index.get_loc('YTD CII')] = 'color: #FF8C00; font-weight: bold;'  # Orange
        elif cii_val == "E":
            styles[row.index.get_loc('YTD CII')] = 'color: #FF0000; font-weight: bold;'  # Red

    return styles

def create_excel_download_with_styling(df, filename):
    """Create Excel file with styling including CII color coding."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Vessel Report"

    # Write headers
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    # Write data and apply styling
    for row_idx, row_data in df.iterrows():
        for col_idx, (col_name, cell_value) in enumerate(row_data.items(), 1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value)
            cell.alignment = Alignment(wrap_text=True, vertical='top')

            if 'Hull Condition' in col_name or 'ME Efficiency' in col_name:
                if cell_value == "Good":
                    cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                elif cell_value == "Average":
                    cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                elif cell_value == "Poor":
                    cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                elif cell_value == "Anomalous data":
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    cell.font = Font(color="FFFFFF")  # White text on red background
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            elif col_name == 'YTD CII':
                # Remove CII color coding for Excel to avoid color errors
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            elif col_name == 'Comments':
                cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')
            elif col_name == 'Potential Fuel Saving (MT/Day)':
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    # Auto-adjust column widths and set specific width for Comments and S. No.
    for col_idx, column in enumerate(df.columns, 1):
        column_letter = get_column_letter(col_idx)
        if column == 'Comments':
            ws.column_dimensions[column_letter].width = 40
        elif column == 'S. No.':
            ws.column_dimensions[column_letter].width = 8
        else:
            max_length = 0
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
        "Good": "D4EDDA",
        "Average": "FFF3CD",
        "Poor": "F8D7DA",
        "Anomalous data": "#f8d7da"
    }
    return color_map.get(cell_value, None)

def get_cii_text_color(cii_value):
    """Get text color for CII rating (updated A and B colors)."""
    cii_val = str(cii_value).upper() if pd.notna(cii_value) else "N/A"
    color_map = {
        "A": (0, 100, 0),      # Dark green
        "B": (144, 238, 144),  # Light green
        "C": (255, 215, 0),    # Yellow
        "D": (255, 140, 0),    # Orange
        "E": (255, 0, 0)       # Red
    }
    return color_map.get(cii_val, None)

def set_cell_border(cell, **kwargs):
    """Set borders for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for border_name in ("top", "left", "bottom", "right"):
        if border_name in kwargs:
            border_element = OxmlElement(f"w:{border_name}")
            for attr, value in kwargs[border_name].items():
                border_element.set(qn(f"w:{attr}"), str(value))
            tcPr.append(border_element)

def set_cell_shading(cell, color_hex):
    """Set background color for a table cell using direct XML manipulation."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)

def create_enhanced_word_report(df, template_path="Fleet Performance Template.docx", num_months=2):
    """Create an enhanced Word report with improved table formatting and CII color coding."""
    try:
        if not os.path.exists(template_path):
            st.error(f"Template file '{template_path}' not found in the repository root.")
            return None

        doc = Document(template_path)
      
        # Find placeholder and replace with report
        placeholder_found = False
      
        for paragraph in doc.paragraphs:
            if "{{Template}}" in paragraph.text:
                paragraph.clear()
                placeholder_found = True
              
                # Add report title with better styling
                title_paragraph = doc.add_paragraph()
                title_run = title_paragraph.add_run("Fleet Performance Analysis Report")
                title_run.font.size = Pt(24)
                title_run.font.bold = True
                title_run.font.color.rgb = RGBColor(47, 117, 181)  # Blue color
                title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
              
                # Add generation date
                date_paragraph = doc.add_paragraph()
                date_run = date_paragraph.add_run(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
                date_run.font.italic = True
                date_run.font.size = Pt(12)
                date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
              
                # Add some spacing
                doc.add_paragraph("")
              
                # Create table with simplified structure
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
              
                # Set table width to page width
                table.autofit = False
                table.allow_autofit = False
              
                # Define column widths based on content type
                page_width = Inches(8.5)  # Standard page width minus margins
                total_width = Inches(7.5)  # Usable width
              
                # Calculate column widths - optimized for better distribution with smaller S. No.
                col_widths = {}
                for col_name in df.columns:
                    if col_name == 'S. No.':
                        col_widths[col_name] = 288000  # 0.2 inches in EMUs (reduced from 0.3)
                    elif col_name == 'Vessel Name':
                        col_widths[col_name] = 1728000  # 1.2 inches in EMUs
                    elif col_name == 'Comments':
                        col_widths[col_name] = 5760000  # 4.0 inches in EMUs
                    elif col_name == 'Potential Fuel Saving (MT/Day)':
                        col_widths[col_name] = 1152000  # 0.8 inches in EMUs
                    elif col_name == 'YTD CII':
                        col_widths[col_name] = 576000  # 0.4 inches in EMUs
                    elif 'Hull Condition' in col_name or 'ME Efficiency' in col_name:
                        col_widths[col_name] = 864000  # 0.6 inches in EMUs
                    else:
                        col_widths[col_name] = 864000  # 0.6 inches in EMUs
              
                # Set column widths using integer values
                for i, col_name in enumerate(df.columns):
                    table.columns[i].width = col_widths[col_name]
              
                # Style header row
                header_cells = table.rows[0].cells
                for i, col_name in enumerate(df.columns):
                    cell = header_cells[i]
                    cell.text = col_name
                  
                    # Header styling
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = cell.paragraphs[0].runs[0]
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(10)
                  
                    # Header background color
                    set_cell_shading(cell, "2F75B5")
                  
                    # Header borders
                    set_cell_border(
                        cell,
                        top={"sz": 6, "val": "single", "color": "000000"},
                        left={"sz": 6, "val": "single", "color": "000000"},
                        bottom={"sz": 6, "val": "single", "color": "000000"},
                        right={"sz": 6, "val": "single", "color": "000000"}
                    )
              
                # Add data rows
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        cell = row_cells[i]
                        cell_value = str(value) if pd.notna(value) else "N/A"
                        cell.text = cell_value
                      
                        # Cell styling based on column type
                        column_name = df.columns[i]
                      
                        # Set font size
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(9)
                      
                        # Apply conditional formatting
                        if 'Hull Condition' in column_name or 'ME Efficiency' in column_name:
                            color_hex = get_cell_color(cell_value)
                            if color_hex:
                                set_cell_shading(cell, color_hex)
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        elif column_name == 'YTD CII':
                            # Apply CII text color coding
                            cii_color = get_cii_text_color(cell_value)
                            if cii_color:
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.color.rgb = RGBColor(*cii_color)
                                        run.font.bold = True
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        elif column_name == 'Comments':
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                      
                        # Set text wrapping and vertical alignment
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                      
                        # Enable text wrapping
                        tcW = OxmlElement("w:tcW")
                        tcW.set(qn("w:type"), "auto")
                        tcPr.append(tcW)
                      
                        # Set vertical alignment to top
                        vAlign = OxmlElement('w:vAlign')
                        vAlign.set(qn('w:val'), 'top')
                        tcPr.append(vAlign)
                      
                        # Add borders
                        set_cell_border(
                            cell,
                            top={"sz": 6, "val": "single", "color": "000000"},
                            left={"sz": 6, "val": "single", "color": "000000"},
                            bottom={"sz": 6, "val": "single", "color": "000000"},
                            right={"sz": 6, "val": "single", "color": "000000"}
                        )
              
                # Add page break before appendix
                doc.add_page_break()
              
                # Add Appendix section
                appendix_title = doc.add_paragraph()
                appendix_run = appendix_title.add_run("Appendix")
                appendix_run.font.size = Pt(20)
                appendix_run.font.bold = True
                appendix_run.font.color.rgb = RGBColor(255, 255, 255)
                appendix_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
              
                # Set appendix title background
                appendix_title_format = appendix_title.paragraph_format
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:val'), 'clear')
                shading_elm.set(qn('w:color'), 'auto')
                shading_elm.set(qn('w:fill'), '00B0F0')
                appendix_title_format._element.get_or_add_pPr().append(shading_elm)
              
                # Add spacing
                doc.add_paragraph("")
              
                # General Conditions section
                general_heading = doc.add_paragraph()
                general_run = general_heading.add_run("General Conditions")
                general_run.font.size = Pt(14)
                general_run.font.bold = True
                general_run.font.color.rgb = RGBColor(47, 117, 181)
              
                # Add bullet points for general conditions
                conditions = [
                    "Analysis Period is Last Six Months or after the Last Event whichever is later",
                    "Days with Good Weather (BF<=4) are considered for analysis",
                    "Days with Steaming hrs greater than 17 considered for analysis",
                    "Data is compared with Original Sea Trial"
                ]
              
                for condition in conditions:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.first_line_indent = Inches(-0.25)
                    run = p.add_run("• " + condition)
                    run.font.size = Pt(10)
              
                # Hull Performance section
                doc.add_paragraph("")
                hull_heading = doc.add_paragraph()
                hull_run = hull_heading.add_run("Hull Performance")
                hull_run.font.size = Pt(14)
                hull_run.font.bold = True
                hull_run.font.color.rgb = RGBColor(47, 117, 181)
              
                # Hull performance criteria with colors
                hull_criteria = [
                    ("Excess Power < 15% – Rating Good", RGBColor(0, 176, 80)),
                    ("15% < Excess Power < 25% – Rating Average", RGBColor(255, 192, 0)),
                    ("Excess Power > 25% – Rating Poor", RGBColor(255, 0, 0))
                ]
              
                for criteria, color in hull_criteria:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.first_line_indent = Inches(-0.25)
                    run = p.add_run("• " + criteria)
                    run.font.size = Pt(10)
                    run.font.color.rgb = color
              
                # Machinery Performance section
                doc.add_paragraph("")
                machinery_heading = doc.add_paragraph()
                machinery_run = machinery_heading.add_run("Machinery Performance")
                machinery_run.font.size = Pt(14)
                machinery_run.font.bold = True
                machinery_run.font.color.rgb = RGBColor(47, 117, 181)
              
                # Machinery performance criteria with colors
                machinery_criteria = [
                    ("SFOC (g/kWh) within ±10 from Shop test condition are considered as \"Good\"", RGBColor(0, 176, 80)),
                    ("SFOC (g/kWh) Greater than 10 and less than 20 are considered as \"Average\"", RGBColor(255, 192, 0)),
                    ("SFOC (g/kWh) Above 20 are considered as \"Poor\"", RGBColor(255, 0, 0))
                ]
              
                for criteria, color in machinery_criteria:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                    p.paragraph_format.first_line_indent = Inches(-0.25)
                    run = p.add_run("• " + criteria)
                    run.font.size = Pt(10)
                    run.font.color.rgb = color

                # CII Rating section
                # doc.add_paragraph("")
                # cii_heading = doc.add_paragraph()
                # cii_run = cii_heading.add_run("CII Rating")
                # cii_run.font.size = Pt(14)
                # cii_run.font.bold = True
                # cii_run.font.color.rgb = RGBColor(47, 117, 181)
              
                # CII performance criteria with colors
                # cii_criteria = [
                #     ("Rating A – Significantly Better Performance", RGBColor(144, 238, 144)),  # Light green
                #     ("Rating B – Better Performance", RGBColor(0, 100, 0)),                    # Dark green
                #     ("Rating C – Moderate Performance", RGBColor(255, 215, 0)),               # Yellow
                #     ("Rating D – Minor Inferior Performance", RGBColor(255, 140, 0)),         # Orange
                #     ("Rating E – Inferior Performance", RGBColor(255, 0, 0))                  # Red
                # ]
              
                # for criteria, color in cii_criteria:
                #     p = doc.add_paragraph()
                #     p.paragraph_format.left_indent = Inches(0.25)
                #     p.paragraph_format.first_line_indent = Inches(-0.25)
                #     run = p.add_run("• " + criteria)
                #     run.font.size = Pt(10)
                #     run.font.color.rgb = color
              
                break
      
        if not placeholder_found:
            st.warning("Placeholder '{{Template}}' not found. Adding report at the end of document.")
            # If no placeholder found, add content at the end
            doc.add_page_break()
          
            # Add the same content as above but at the end
            title_paragraph = doc.add_paragraph()
            title_run = title_paragraph.add_run("Fleet Performance Analysis Report")
            title_run.font.size = Pt(24)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(47, 117, 181)
            title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
          
            # Continue with the rest of the report...
      
        # Save to bytes buffer
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()
      
    except Exception as e:
        st.error(f"Error creating Word report: {str(e)}")
        return None

# Function to reset session state
def reset_page():
    st.session_state.selected_vessels = set()
    st.session_state.report_data = None
    st.session_state.search_query = ""
    st.session_state.report_months = 2
    st.cache_data.clear()
    st.session_state.trigger_rerun = True

# Main Application
def main():
    # Check for rerun flag
    if 'trigger_rerun' in st.session_state and st.session_state.trigger_rerun:
        st.session_state.trigger_rerun = False
        st.rerun()

    # Lambda Function URL
    LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/"

    # Enhanced Header with gradient styling
    st.markdown("""
    <div class="main-header">
        <h1>🚢 Enhanced Vessel Performance Report Tool</h1>
        <p>Generate comprehensive performance reports with advanced analytics and beautiful formatting</p>
    </div>
    """, unsafe_allow_html=True)

    # Reset button with enhanced styling
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔄 Reset All", type="secondary", use_container_width=True):
            reset_page()

    # Section 1: Vessel Selection
    st.markdown('<div class="section-header">1. 🎯 Select Vessels</div>', unsafe_allow_html=True)

    # Load vessels
    with st.spinner("Loading vessels..."):
        try:
            all_vessels = fetch_all_vessels(LAMBDA_FUNCTION_URL)
            st.markdown(f'<div class="success-box">✅ Successfully loaded {len(all_vessels)} vessels!</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"❌ Failed to load vessels: {str(e)}")
            all_vessels = []

    if all_vessels:
        # Enhanced search with better styling
        search_query = st.text_input(
            "🔍 Search vessels:",
            value=st.session_state.search_query,
            placeholder="Type vessel name to filter...",
            help="Start typing to filter the vessel list in real-time"
        )

        if search_query != st.session_state.search_query:
            st.session_state.search_query = search_query

        # Filter vessels
        filtered_vessels = filter_vessels_client_side(all_vessels, search_query)

        # Enhanced metrics display
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h3>📊 Total Vessels</h3>
                <h2>{len(all_vessels)}</h2>
            </div>
            """, unsafe_allow_html=True)
      
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h3>🔍 Filtered</h3>
                <h2>{len(filtered_vessels)}</h2>
            </div>
            """, unsafe_allow_html=True)
      
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <h3>✅ Selected</h3>
                <h2>{len(st.session_state.selected_vessels)}</h2>
            </div>
            """, unsafe_allow_html=True)

        # Vessel selection with improved UI
        if filtered_vessels:
            st.subheader("Select Vessels:")
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
            st.markdown('<div class="info-box">🔍 No vessels match your search criteria</div>', unsafe_allow_html=True)

        selected_vessels_list = list(st.session_state.selected_vessels)

        # Enhanced selected vessels display
        if selected_vessels_list:
            with st.expander(f"📋 Selected Vessels ({len(selected_vessels_list)})", expanded=False):
                for i, vessel in enumerate(sorted(selected_vessels_list), 1):
                    st.write(f"{i}. {vessel}")
    else:
        st.error("❌ Failed to load vessels. Please check your connection and try again.")
        selected_vessels_list = []

    # Section 2: Report Generation
    st.markdown('<div class="section-header">2. 🚀 Generate Performance Report</div>', unsafe_allow_html=True)

    # Enhanced report duration selection
    st.subheader("📅 Select Report Duration:")
    st.session_state.report_months = st.radio(
        "",
        options=[1, 2, 3],
        format_func=lambda x: f"📊 {x} Month{'s' if x > 1 else ''} Analysis",
        index=1,
        horizontal=True,
        help="Choose the number of months for historical analysis"
    )

    if selected_vessels_list:
        # Enhanced generate button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Generate Performance Report", type="primary", use_container_width=True):
                with st.spinner("Generating comprehensive report with advanced analytics..."):
                    try:
                        start_time = time.time()
                        st.session_state.report_data = query_report_data(
                            LAMBDA_FUNCTION_URL, selected_vessels_list, st.session_state.report_months
                        )

                        generation_time = time.time() - start_time

                        if not st.session_state.report_data.empty:
                            st.markdown(f'<div class="success-box">✅ Report generated successfully in {generation_time:.2f} seconds!</div>', unsafe_allow_html=True)
                        else:
                            st.warning("⚠️ No data found for the selected vessels.")

                    except Exception as e:
                        st.error(f"❌ Error generating report: {str(e)}")
                        st.session_state.report_data = None
    else:
        st.markdown('<div class="info-box">⚠️ Please select at least one vessel to generate a report</div>', unsafe_allow_html=True)

    # Section 3: Report Display and Download
    if st.session_state.report_data is not None and not st.session_state.report_data.empty:
        st.markdown('<div class="section-header">3. 📊 Report Results & Analytics</div>', unsafe_allow_html=True)

        # Enhanced report display
        st.subheader("📋 Performance Data Table")
        styled_df = st.session_state.report_data.style.apply(style_condition_columns, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=400)

        # Enhanced download section
        st.subheader("📥 Download Options")
        col1, col2, col3 = st.columns(3)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        with col1:
            filename = f"vessel_performance_report_{timestamp}.xlsx"
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
            word_filename = f"fleet_performance_report_{timestamp}.docx"
            try:
                template_path = "Fleet Performance Template.docx"
                if os.path.exists(template_path):
                    word_data = create_enhanced_word_report(st.session_state.report_data, template_path, st.session_state.report_months)
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
            except Exception as e:
                st.error(f"❌ Error creating Word file: {str(e)}")

        # Enhanced analytics section
        with st.expander("📈 Advanced Analytics & Insights", expanded=False):
            tab1, tab2, tab3, tab4 = st.tabs(["🛡️ Hull Analysis", "⚙️ Engine Analysis", "📊 Trend Analysis", "🌍 CII Analysis"])

            with tab1:
                st.subheader("Hull Condition Distribution")
                hull_cols = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]

                if hull_cols:
                    col1, col2 = st.columns(2)
                    with col1:
                        latest_hull_data = st.session_state.report_data[hull_cols[0]].value_counts()
                        if len(latest_hull_data) > 0:
                            st.bar_chart(latest_hull_data, use_container_width=True)
                            st.caption(f"Distribution for {hull_cols[0]}")

                    with col2:
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
                      
                        if hull_summary:
                            hull_summary_df = pd.DataFrame(hull_summary)
                            st.dataframe(hull_summary_df, use_container_width=True)

            with tab2:
                st.subheader("ME Efficiency Distribution")
                me_cols = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]

                if me_cols:
                    col1, col2 = st.columns(2)
                    with col1:
                        latest_me_data = st.session_state.report_data[me_cols[0]].value_counts()
                        if len(latest_me_data) > 0:
                            st.bar_chart(latest_me_data, use_container_width=True)
                            st.caption(f"Distribution for {me_cols[0]}")

                    with col2:
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
                      
                        if me_summary:
                            me_summary_df = pd.DataFrame(me_summary)
                            st.dataframe(me_summary_df, use_container_width=True)

            with tab3:
                st.subheader("Performance Trends")
              
                hull_cols = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
                me_cols = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]

                if len(hull_cols) >= 2:
                    st.write("**Hull Condition Trends (% Good)**")
                    hull_trend_data = []
                    for col in hull_cols:
                        month = col.replace("Hull Condition ", "")
                        total_with_data = len(st.session_state.report_data[st.session_state.report_data[col] != "N/A"])
                        good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                      
                        if total_with_data > 0:
                            percentage = (good_count / total_with_data * 100)
                            hull_trend_data.append({"Month": month, "Good %": percentage})
                  
                    if hull_trend_data:
                        hull_trend_df = pd.DataFrame(hull_trend_data)
                        if hull_trend_df["Good %"].sum() > 0:
                            st.line_chart(hull_trend_df.set_index("Month"), use_container_width=True)

                if len(me_cols) >= 2:
                    st.write("**ME Efficiency Trends (% Good)**")
                    me_trend_data = []
                    for col in me_cols:
                        month = col.replace("ME Efficiency ", "")
                        total_with_data = len(st.session_state.report_data[st.session_state.report_data[col] != "N/A"])
                        good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                      
                        if total_with_data > 0:
                            percentage = (good_count / total_with_data * 100)
                            me_trend_data.append({"Month": month, "Good %": percentage})
                  
                    if me_trend_data:
                        me_trend_df = pd.DataFrame(me_trend_data)
                        if me_trend_df["Good %"].sum() > 0:
                            st.line_chart(me_trend_df.set_index("Month"), use_container_width=True)

            with tab4:
                st.subheader("CII Rating Distribution")
                if 'YTD CII' in st.session_state.report_data.columns:
                    col1, col2 = st.columns(2)
                    with col1:
                        cii_data = st.session_state.report_data['YTD CII'].value_counts()
                        if len(cii_data) > 0:
                            st.bar_chart(cii_data, use_container_width=True)
                            st.caption("CII Rating Distribution")

                    with col2:
                        cii_summary = []
                        counts = st.session_state.report_data['YTD CII'].value_counts()
                        total_vessels = len(st.session_state.report_data)
                        
                        for rating in ['A', 'B', 'C', 'D', 'E']:
                            count = counts.get(rating, 0)
                            percentage = (count / total_vessels * 100) if total_vessels > 0 else 0
                            cii_summary.append({
                                "Rating": rating,
                                "Count": count,
                                "Percentage": f"{percentage:.1f}%"
                            })
                        
                        # Add N/A if exists
                        na_count = counts.get('N/A', 0) + sum(1 for x in st.session_state.report_data['YTD CII'] if pd.isna(x))
                        if na_count > 0:
                            na_percentage = (na_count / total_vessels * 100) if total_vessels > 0 else 0
                            cii_summary.append({
                                "Rating": "N/A",
                                "Count": na_count,
                                "Percentage": f"{na_percentage:.1f}%"
                            })
                        
                        if cii_summary:
                            cii_summary_df = pd.DataFrame(cii_summary)
                            st.dataframe(cii_summary_df, use_container_width=True)
                else:
                    st.info("No CII data available for analysis")

    # Enhanced Help Section
    with st.expander("📖 User Guide & Features", expanded=False):
        st.markdown("""
        ### 🌟 Enhanced Features:

        **🎨 Modern UI Design:**
        - Gradient backgrounds and modern styling
        - Responsive layout with improved visual hierarchy
        - Color-coded metrics and status indicators

        **🔍 Smart Vessel Selection:**
        - Real-time search and filtering
        - Multi-column layout for easy browsing
        - Visual metrics showing selection status

        **📊 Advanced Analytics:**
        - Interactive charts and visualizations
        - Multi-month trend analysis
        - Performance distribution insights
        - CII rating analysis with color coding

        **📥 Professional Reports:**
        - Enhanced Excel reports with color coding
        - Beautifully formatted Word documents
        - Optimized table layouts with proper spacing
        - CII rating color coding in all formats

        ### 📋 How to Use:

        1. **🔍 Search**: Use the search box to find specific vessels
        2. **✅ Select**: Check vessels you want to analyze
        3. **📅 Configure**: Choose analysis period (1-3 months)
        4. **🚀 Generate**: Click to create comprehensive report
        5. **📊 Analyze**: Review charts and performance metrics
        6. **📥 Download**: Export in your preferred format

        ### 🎯 Performance Indicators:

        **🛡️ Hull Condition:**
        - 🟢 **Good**: < 15% excess power
        - 🟡 **Average**: 15-25% excess power
        - 🔴 **Poor**: > 25% excess power

        **⚙️ Engine Efficiency:**
        - 🟢 **Good**: 160-180 g/kWh SFOC
        - 🟡 **Average**: 180-190 g/kWh SFOC
        - 🔴 **Poor**: > 190 g/kWh SFOC
        - ⚪ **Anomalous**: < 160 g/kWh SFOC
        """)

if __name__ == "__main__":
    main()
