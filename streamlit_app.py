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
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn # For XML manipulation
from docx.oxml import OxmlElement # For XML manipulation
from docx.shared import Pt # For font size
from docx.enum.text import WD_ALIGN_PARAGRAPH # For paragraph alignment
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

if 'report_months' not in st.session_state:
    st.session_state.report_months = 2 # Default to 2 months

# Enhanced Lambda Invocation Helper (Backwards Compatible)
def invoke_lambda_function_url(lambda_url, payload, timeout=60): # Timeout fixed to 60 seconds
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
    """Enhanced version of the original query_report_data function with better progress tracking."""
    if not vessel_names:
        return pd.DataFrame()

    today = datetime.now()
    first_day_current_month = today.replace(day=1)

    # Prepare date strings and column names based on num_months
    hull_dates_info = []
    me_dates_info = []

    for i in range(num_months):
        # Hull Condition Dates (last day of the month)
        target_month_end = first_day_current_month - timedelta(days=1) - timedelta(days=30 * i) # Approximate
        target_month_end = target_month_end.replace(day=1) - timedelta(days=1) # Get last day of the i-th previous month

        hull_date_str = target_month_end.strftime("%Y-%m-%d")
        hull_col_name = f"Hull Condition {target_month_end.strftime('%b %y')}"
        hull_power_loss_col_name = f"Hull Roughness Power Loss % {target_month_end.strftime('%b %y')}"
        hull_dates_info.append({
            'date_str': hull_date_str,
            'col_name': hull_col_name,
            'power_loss_col_name': hull_power_loss_col_name,
            'interval_str': f"INTERVAL '{i+1} month'"
        })

        # ME SFOC Dates (average of the entire month)
        me_col_name = f"ME Efficiency {target_month_end.strftime('%b %y')}"
        me_dates_info.append({
            'col_name': me_col_name,
            'interval_start_str': f"INTERVAL '{i+1} month'",
            'interval_end_str': f"INTERVAL '{i} month'"
        })

    # Process vessels in smaller batches with enhanced progress tracking
    batch_size = 10 # Batch size fixed to 10
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

        # ME SFOC queries
        for me_info in me_dates_info:
            batch_queries.append((me_info['col_name'], f"""
SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
FROM vessel_performance_summary vps
JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
WHERE vp.vessel_name IN ({vessel_names_list_str})
AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - {me_info['interval_start_str']})
AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - {me_info['interval_end_str']})
GROUP BY vp.vessel_name
""", all_me_data_by_month[me_info['col_name']]))

        # Fixed queries (Potential Fuel Saving, YTD CII)
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
                ).round(2) # Round to 2 decimal places
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
            return "Anomalous data", "SFOC value is unusually low, indicating potential data anomaly."
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
            df_final = df_final.drop(columns=['current_me_comment']) # Drop temporary column

    # Join all collected comments into the final 'Comments' column
    df_final['Comments'] = df_final['temp_comments_list'].apply(lambda x: " ".join(x).strip())
    df_final = df_final.drop(columns=['temp_comments_list']) # Drop the temporary list column

    # Define the desired order of columns
    desired_columns_order = ['S. No.', 'Vessel Name']
    for hull_info in hull_dates_info:
        desired_columns_order.append(hull_info['col_name'])
    for me_info in me_dates_info:
        desired_columns_order.append(me_info['col_name'])
    desired_columns_order.extend(['Potential Fuel Saving (MT/Day)', 'YTD CII', 'Comments'])

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
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center') # Wrap text for headers

    # Write data and apply styling
    for row_idx, row_data in df.iterrows():
        for col_idx, (col_name, cell_value) in enumerate(row_data.items(), 1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value)
            cell.alignment = Alignment(wrap_text=True, vertical='top') # Wrap text for all data cells

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
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center') # Center for conditions
            elif col_name == 'YTD CII':
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center') # Center for CII
            elif col_name == 'Comments':
                cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top') # Left align for comments
            elif col_name == 'Potential Fuel Saving (MT/Day)':
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center') # Center for fuel saving

    # Auto-adjust column widths and set specific width for Comments and S. No.
    for col_idx, column in enumerate(df.columns, 1):
        column_letter = get_column_letter(col_idx)
        if column == 'Comments':
            ws.column_dimensions[column_letter].width = 40 # Fixed wider width for comments
        elif column == 'S. No.':
            ws.column_dimensions[column_letter].width = 8 # Smaller fixed width for S. No.
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
        "Good": "D4EDDA",      # Light green
        "Average": "FFF3CD",   # Light yellow
        "Poor": "F8D7DA",      # Light red
        "Anomalous data": "E0E0E0"  # Light gray
    }
    return color_map.get(cell_value, None)

# Helper function to set cell borders
def set_cell_border(cell, **kwargs):
    """
    Set borders for a table cell.
    Usage: set_cell_border(cell, top={"sz": 12, "val": "single", "color": "#000000"})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Create a border element for each side
    for border_name in ("top", "left", "bottom", "right"):
        if border_name in kwargs:
            border_element = OxmlElement(f"w:{border_name}")
            for attr, value in kwargs[border_name].items():
                # Ensure 'sz' is an integer representing eighths of a point
                # The value passed to sz should already be an integer (e.g., 6 for 0.75pt)
                border_element.set(qn(f"w:{attr}"), str(value))
            tcPr.append(border_element)

# Helper function to set cell shading (background color)
def set_cell_shading(cell, color_hex):
    """
    Set background color for a table cell using direct XML manipulation.
    color_hex should be an RGB hex string (e.g., "FF0000" for red).
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear") # "clear" means solid fill
    shd.set(qn("w:color"), "auto") # "auto" means default text color
    shd.set(qn("w:fill"), color_hex) # The fill color
    tcPr.append(shd)

def create_advanced_word_report(df, template_path="Fleet Performance Template.docx", num_months=2):
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
                title_run.font.size = Pt(24) # Larger font for title
                title_run.font.bold = True
                title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Add generation date
                date_paragraph = doc.add_paragraph()
                date_run = date_paragraph.add_run(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
                date_run.font.italic = True
                date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Determine column groups for multi-level header
                hull_cols_base = [col for col in df.columns if 'Hull Condition' in col]
                me_cols_base = [col for col in df.columns if 'ME Efficiency' in col]

                # Create the main data table with two header rows
                table = doc.add_table(rows=2, cols=len(df.columns))
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                table.allow_autofit = False

                # Calculate column widths
                # Total content width for the table (e.g., 6.5 inches)
                total_content_width = Inches(6.5)
                
                # Fixed widths for specific columns
                s_no_col_width = Inches(0.5) # Smaller width for S. No.
                vessel_name_col_width = Inches(1.2)
                comments_col_width = Inches(2.5) # Increased width for comments
                fuel_saving_col_width = Inches(1.0)
                cii_col_width = Inches(0.7)

                # Calculate remaining width for dynamic columns
                fixed_width_sum = s_no_col_width + vessel_name_col_width + comments_col_width + fuel_saving_col_width + cii_col_width
                
                dynamic_cols_count = len(hull_cols_base) + len(me_cols_base)
                # Ensure dynamic_col_width is not zero if dynamic_cols_count is zero
                dynamic_col_width = (total_content_width - fixed_width_sum) / dynamic_cols_count if dynamic_cols_count > 0 else Inches(0.1)

                col_widths = {}
                for i, col_name in enumerate(df.columns):
                    if col_name == 'S. No.':
                        col_widths[col_name] = s_no_col_width
                    elif col_name == 'Vessel Name':
                        col_widths[col_name] = vessel_name_col_width
                    elif col_name == 'Comments':
                        col_widths[col_name] = comments_col_width
                    elif col_name == 'Potential Fuel Saving (MT/Day)':
                        col_widths[col_name] = fuel_saving_col_width
                    elif col_name == 'YTD CII':
                        col_widths[col_name] = cii_col_width
                    else:
                        col_widths[col_name] = dynamic_col_width

                for i, column_name in enumerate(df.columns):
                    table.columns[i].width = col_widths[column_name]

                # Define border size (6 for 0.75pt, which is 6 eighths of a point)
                border_sz = 6 

                # Populate the first header row (main categories)
                hdr_cells_0 = table.rows[0].cells
                
                # S. No.
                cell_s_no = hdr_cells_0[df.columns.get_loc('S. No.')]
                cell_s_no.text = "S. No."
                cell_s_no.merge(table.rows[1].cells[df.columns.get_loc('S. No.')]) # Merge with cell below
                cell_s_no.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell_s_no, "2F75B5")
                set_cell_border(cell_s_no, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # Vessel Name
                cell_vessel_name = hdr_cells_0[df.columns.get_loc('Vessel Name')]
                cell_vessel_name.text = "Vessel Name"
                cell_vessel_name.merge(table.rows[1].cells[df.columns.get_loc('Vessel Name')]) # Merge with cell below
                cell_vessel_name.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell_vessel_name, "2F75B5")
                set_cell_border(cell_vessel_name, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # Hull Condition Group
                if hull_cols_base:
                    first_hull_col_idx = df.columns.get_loc(hull_cols_base[0])
                    last_hull_col_idx = df.columns.get_loc(hull_cols_base[-1])
                    
                    merged_cell_hull = hdr_cells_0[first_hull_col_idx]
                    # Merge cells for "Hull Condition" across the number of months
                    for i in range(first_hull_col_idx + 1, last_hull_col_idx + 1):
                        merged_cell_hull.merge(hdr_cells_0[i])
                    merged_cell_hull.text = "Hull Condition"
                    merged_cell_hull.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_shading(merged_cell_hull, "2F75B5")
                    set_cell_border(merged_cell_hull, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})
                
                # ME Efficiency Group
                if me_cols_base:
                    first_me_col_idx = df.columns.get_loc(me_cols_base[0])
                    last_me_col_idx = df.columns.get_loc(me_cols_base[-1])

                    merged_cell_me = hdr_cells_0[first_me_col_idx]
                    # Merge cells for "ME Efficiency" across the number of months
                    for i in range(first_me_col_idx + 1, last_me_col_idx + 1):
                        merged_cell_me.merge(hdr_cells_0[i])
                    merged_cell_me.text = "ME Efficiency"
                    merged_cell_me.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_shading(merged_cell_me, "2F75B5")
                    set_cell_border(merged_cell_me, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # Potential Fuel Saving (MT/Day)
                fuel_saving_col_idx = df.columns.get_loc('Potential Fuel Saving (MT/Day)')
                cell_fuel_saving = hdr_cells_0[fuel_saving_col_idx]
                cell_fuel_saving.text = "Potential Fuel Saving (MT/Day)"
                cell_fuel_saving.merge(table.rows[1].cells[fuel_saving_col_idx]) # Merge with cell below
                cell_fuel_saving.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell_fuel_saving, "2F75B5")
                set_cell_border(cell_fuel_saving, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # YTD CII
                cii_col_idx = df.columns.get_loc('YTD CII')
                cell_cii = hdr_cells_0[cii_col_idx]
                cell_cii.text = "YTD CII"
                cell_cii.merge(table.rows[1].cells[cii_col_idx]) # Merge with cell below
                cell_cii.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell_cii, "2F75B5")
                set_cell_border(cell_cii, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # Comments
                comments_col_idx = df.columns.get_loc('Comments')
                cell_comments = hdr_cells_0[comments_col_idx]
                cell_comments.text = "Comments"
                cell_comments.merge(table.rows[1].cells[comments_col_idx]) # Merge with cell below
                cell_comments.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell_comments, "2F75B5")
                set_cell_border(cell_comments, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})


                # Populate the second header row (sub-categories for Hull and ME)
                hdr_cells_1 = table.rows[1].cells
                
                # Hull Condition sub-headers
                for i, col_name in enumerate(hull_cols_base):
                    cell = hdr_cells_1[df.columns.get_loc(col_name)]
                    month_info = col_name.replace("Hull Condition ", "")
                    if i == 0:
                        cell.text = "Current Month"
                    elif i == 1:
                        cell.text = "Previous Month"
                    elif i == 2:
                        cell.text = "2 Months Ago"
                    else:
                        cell.text = month_info # Fallback
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_shading(cell, "2F75B5")
                    set_cell_border(cell, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # ME Efficiency sub-headers
                for i, col_name in enumerate(me_cols_base):
                    cell = hdr_cells_1[df.columns.get_loc(col_name)]
                    month_info = col_name.replace("ME Efficiency ", "")
                    if i == 0:
                        cell.text = "Current Month"
                    elif i == 1:
                        cell.text = "Previous Month"
                    elif i == 2:
                        cell.text = "2 Months Ago"
                    else:
                        cell.text = month_info # Fallback
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_cell_shading(cell, "2F75B5")
                    set_cell_border(cell, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

                # Set font color for all header cells to white
                for row_idx in range(2):
                    for cell in table.rows[row_idx].cells:
                        for run in cell.paragraphs[0].runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)


                # Add data rows with formatting
                for _, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        cell = row_cells[i]
                        cell_value = str(value) if pd.notna(value) else "N/A"
                        cell.text = cell_value
                        
                        # Enable text wrapping for all cells
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        
                        # Add w:tcW element to enable auto width and wrapping
                        tcW = OxmlElement("w:tcW")
                        tcW.set(qn("w:type"), "auto") # Auto width
                        tcPr.append(tcW)
                        
                        # Add vertical alignment to top for all cells
                        vAlign = OxmlElement('w:vAlign')
                        vAlign.set(qn('w:val'), 'top')
                        tcPr.append(vAlign)

                        # Apply conditional formatting and alignment
                        column_name = df.columns[i]
                        if 'Hull Condition' in column_name or 'ME Efficiency' in column_name:
                            color_hex = get_cell_color(cell_value)
                            if color_hex:
                                set_cell_shading(cell, color_hex)
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER # Center align for conditions
                        elif column_name == 'Comments':
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Left align for comments
                        else:
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER # Default center

                        # Apply borders to data cells
                        set_cell_border(
                            cell,
                            top={"sz": border_sz, "val": "single", "color": "000000"},
                            left={"sz": border_sz, "val": "single", "color": "000000"},
                            bottom={"sz": border_sz, "val": "single", "color": "000000"},
                            right={"sz": border_sz, "val": "single", "color": "000000"},
                        )

                # --- Add Appendix Section ---
                doc.add_page_break()

                # Appendix Title with blue background
                appendix_title_paragraph = doc.add_paragraph()
                appendix_title_run = appendix_title_paragraph.add_run("Appendix")
                appendix_title_run.font.size = Pt(20)
                appendix_title_run.font.bold = True
                appendix_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Set background color for the Appendix title paragraph
                appendix_title_paragraph_format = appendix_title_paragraph.paragraph_format
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:val'), 'clear')
                shading_elm.set(qn('w:color'), 'auto')
                shading_elm.set(qn('w:fill'), '00B0F0') # Blue color
                appendix_title_paragraph_format._element.get_or_add_pPr().append(shading_elm)


                # General Conditions
                doc.add_paragraph() # Add a small space
                doc.add_paragraph("General Conditions", style='Heading 3')

                # Custom bullet points for General Conditions
                def add_custom_bullet(doc, text):
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25) # Indent for bullet
                    p.paragraph_format.first_line_indent = Inches(-0.25) # Outdent bullet
                    run = p.add_run("•\t" + text) # Add bullet character and tab
                    run.font.size = Pt(10) # Adjust font size if needed

                add_custom_bullet(doc, "Analysis Period is Last Six Months or the after the Last Event which ever is later")
                add_custom_bullet(doc, "Days with Good Weather (BF<=4) are considered for analysis.")
                add_custom_bullet(doc, "Days with Steaming hrs greater than 17 considered for analysis.")
                add_custom_bullet(doc, "Data is compared with Original Sea Trial")

                # Hull Performance
                doc.add_paragraph() # Add a small space
                doc.add_paragraph("Hull Performance", style='Heading 3')

                # Helper to add bullet points with specific colors and custom formatting
                def add_colored_custom_bullet(doc, text, color_rgb):
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25) # Indent for bullet
                    p.paragraph_format.first_line_indent = Inches(-0.25) # Outdent bullet
                    run = p.add_run("•\t" + text) # Add bullet character and tab
                    run.font.color.rgb = color_rgb
                    run.font.size = Pt(10) # Adjust font size if needed

                add_colored_custom_bullet(doc, "Excess Power < 15 %– Rating Good", RGBColor(0, 176, 80)) # Green
                add_colored_custom_bullet(doc, "15< Excess Power < 25 % – Rating Average", RGBColor(255, 192, 0)) # Orange
                add_colored_custom_bullet(doc, "Excess Power > 25 % – Rating Poor", RGBColor(255, 0, 0)) # Red

                # Machinery Performance
                doc.add_paragraph() # Add a small space
                doc.add_paragraph("Machinery Performance", style='Heading 3')
                add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) within +/- 10 from Shop test condition are considered as \"Good\"", RGBColor(0, 176, 80)) # Green
                add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) Greater than 10 and less than 20 are considered as \"Average\"", RGBColor(255, 192, 0)) # Orange
                add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) Above 20 are considered as \"Poor\"", RGBColor(255, 0, 0)) # Red


                break # Exit loop after finding and processing the placeholder

        if not placeholder_found:
            st.warning("Placeholder '{{Template}}' not found. Adding report at the end of document.")
            # Add report at end if placeholder not found
            doc.add_page_break()

            # Add report title
            title_paragraph = doc.add_paragraph()
            title_run = title_paragraph.add_run("Fleet Performance Analysis Report")
            title_run.font.size = Pt(24) # Larger font for title
            title_run.font.bold = True
            title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Add generation date
            date_paragraph = doc.add_paragraph()
            date_run = date_paragraph.add_run(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
            date_run.font.italic = True
            date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Determine column groups for multi-level header
            hull_cols_base = [col for col in df.columns if 'Hull Condition' in col]
            me_cols_base = [col for col in df.columns if 'ME Efficiency' in col]

            # Create the main data table with two header rows
            table = doc.add_table(rows=2, cols=len(df.columns))
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            table.allow_autofit = False

            # Calculate column widths
            total_content_width = Inches(6.5)
            
            # Fixed widths for specific columns
            s_no_col_width = Inches(0.5) # Smaller width for S. No.
            vessel_name_col_width = Inches(1.2)
            comments_col_width = Inches(2.5) # Increased width for comments
            fuel_saving_col_width = Inches(1.0)
            cii_col_width = Inches(0.7)

            # Calculate remaining width for dynamic columns
            fixed_width_sum = s_no_col_width + vessel_name_col_width + comments_col_width + fuel_saving_col_width + cii_col_width
            
            dynamic_cols_count = len(hull_cols_base) + len(me_cols_base)
            dynamic_col_width = (total_content_width - fixed_width_sum) / dynamic_cols_count if dynamic_cols_count > 0 else Inches(0.1)

            col_widths = {}
            for i, col_name in enumerate(df.columns):
                if col_name == 'S. No.':
                    col_widths[col_name] = s_no_col_width
                elif col_name == 'Vessel Name':
                    col_widths[col_name] = vessel_name_col_width
                elif col_name == 'Comments':
                    col_widths[col_name] = comments_col_width
                elif col_name == 'Potential Fuel Saving (MT/Day)':
                    col_widths[col_name] = fuel_saving_col_width
                elif col_name == 'YTD CII':
                    col_widths[col_name] = cii_col_width
                else:
                    col_widths[col_name] = dynamic_col_width

            for i, column_name in enumerate(df.columns):
                table.columns[i].width = col_widths[column_name]

            # Define border size (6 for 0.75pt, which is 6 eighths of a point)
            border_sz = 6

            # Populate the first header row (main categories)
            hdr_cells_0 = table.rows[0].cells
            
            # S. No.
            cell_s_no = hdr_cells_0[df.columns.get_loc('S. No.')]
            cell_s_no.text = "S. No."
            cell_s_no.merge(table.rows[1].cells[df.columns.get_loc('S. No.')]) # Merge with cell below
            cell_s_no.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell_s_no, "2F75B5")
            set_cell_border(cell_s_no, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # Vessel Name
            cell_vessel_name = hdr_cells_0[df.columns.get_loc('Vessel Name')]
            cell_vessel_name.text = "Vessel Name"
            cell_vessel_name.merge(table.rows[1].cells[df.columns.get_loc('Vessel Name')]) # Merge with cell below
            cell_vessel_name.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell_vessel_name, "2F75B5")
            set_cell_border(cell_vessel_name, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # Hull Condition Group
            if hull_cols_base:
                first_hull_col_idx = df.columns.get_loc(hull_cols_base[0])
                last_hull_col_idx = df.columns.get_loc(hull_cols_base[-1])
                
                merged_cell_hull = hdr_cells_0[first_hull_col_idx]
                # Merge cells for "Hull Condition" across the number of months
                for i in range(first_hull_col_idx + 1, last_hull_col_idx + 1):
                    merged_cell_hull.merge(hdr_cells_0[i])
                merged_cell_hull.text = "Hull Condition"
                merged_cell_hull.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(merged_cell_hull, "2F75B5")
                set_cell_border(merged_cell_hull, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})
            
            # ME Efficiency Group
            if me_cols_base:
                first_me_col_idx = df.columns.get_loc(me_cols_base[0])
                last_me_col_idx = df.columns.get_loc(me_cols_base[-1])

                merged_cell_me = hdr_cells_0[first_me_col_idx]
                # Merge cells for "ME Efficiency" across the number of months
                for i in range(first_me_col_idx + 1, last_me_col_idx + 1):
                    merged_cell_me.merge(hdr_cells_0[i])
                merged_cell_me.text = "ME Efficiency"
                merged_cell_me.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(merged_cell_me, "2F75B5")
                set_cell_border(merged_cell_me, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # Potential Fuel Saving (MT/Day)
            fuel_saving_col_idx = df.columns.get_loc('Potential Fuel Saving (MT/Day)')
            cell_fuel_saving = hdr_cells_0[fuel_saving_col_idx]
            cell_fuel_saving.text = "Potential Fuel Saving (MT/Day)"
            cell_fuel_saving.merge(table.rows[1].cells[fuel_saving_col_idx]) # Merge with cell below
            cell_fuel_saving.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell_fuel_saving, "2F75B5")
            set_cell_border(cell_fuel_saving, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # YTD CII
            cii_col_idx = df.columns.get_loc('YTD CII')
            cell_cii = hdr_cells_0[cii_col_idx]
            cell_cii.text = "YTD CII"
            cell_cii.merge(table.rows[1].cells[cii_col_idx]) # Merge with cell below
            cell_cii.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell_cii, "2F75B5")
            set_cell_border(cell_cii, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # Comments
            comments_col_idx = df.columns.get_loc('Comments')
            cell_comments = hdr_cells_0[comments_col_idx]
            cell_comments.text = "Comments"
            cell_comments.merge(table.rows[1].cells[comments_col_idx]) # Merge with cell below
            cell_comments.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_cell_shading(cell_comments, "2F75B5")
            set_cell_border(cell_comments, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})


            # Populate the second header row (sub-categories for Hull and ME)
            hdr_cells_1 = table.rows[1].cells
            
            # Hull Condition sub-headers
            for i, col_name in enumerate(hull_cols_base):
                cell = hdr_cells_1[df.columns.get_loc(col_name)]
                month_info = col_name.replace("Hull Condition ", "")
                if i == 0:
                    cell.text = "Current Month"
                elif i == 1:
                    cell.text = "Previous Month"
                elif i == 2:
                    cell.text = "2 Months Ago"
                else:
                    cell.text = month_info # Fallback
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell, "2F75B5")
                set_cell_border(cell, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # ME Efficiency sub-headers
            for i, col_name in enumerate(me_cols_base):
                cell = hdr_cells_1[df.columns.get_loc(col_name)]
                month_info = col_name.replace("ME Efficiency ", "")
                if i == 0:
                    cell.text = "Current Month"
                elif i == 1:
                    cell.text = "Previous Month"
                elif i == 2:
                    cell.text = "2 Months Ago"
                else:
                    cell.text = month_info # Fallback
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_cell_shading(cell, "2F75B5")
                set_cell_border(cell, top={"sz": border_sz, "val": "single", "color": "000000"}, left={"sz": border_sz, "val": "single", "color": "000000"}, bottom={"sz": border_sz, "val": "single", "color": "000000"}, right={"sz": border_sz, "val": "single", "color": "000000"})

            # Set font color for all header cells to white
            for row_idx in range(2):
                for cell in table.rows[row_idx].cells:
                    for run in cell.paragraphs[0].runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)


            # Add data rows with formatting
            for _, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row):
                    cell = row_cells[i]
                    cell_value = str(value) if pd.notna(value) else "N/A"
                    cell.text = cell_value
                    
                    # Enable text wrapping for all cells
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    
                    # Add w:tcW element to enable auto width and wrapping
                    tcW = OxmlElement("w:tcW")
                    tcW.set(qn("w:type"), "auto") # Auto width
                    tcPr.append(tcW)
                    
                    # Add vertical alignment to top for all cells
                    vAlign = OxmlElement('w:vAlign')
                    vAlign.set(qn('w:val'), 'top')
                    tcPr.append(vAlign)

                    # Apply conditional formatting and alignment
                    column_name = df.columns[i]
                    if 'Hull Condition' in column_name or 'ME Efficiency' in column_name:
                        color_hex = get_cell_color(cell_value)
                        if color_hex:
                            set_cell_shading(cell, color_hex)
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER # Center align for conditions
                    elif column_name == 'Comments':
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Left align for comments
                    else:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER # Default center

                    # Apply borders to data cells
                    set_cell_border(
                        cell,
                        top={"sz": border_sz, "val": "single", "color": "000000"},
                        left={"sz": border_sz, "val": "single", "color": "000000"},
                        bottom={"sz": border_sz, "val": "single", "color": "000000"},
                        right={"sz": border_sz, "val": "single", "color": "000000"},
                    )

            # --- Add Appendix Section (if placeholder not found) ---
            doc.add_page_break()

            # Appendix Title with blue background
            appendix_title_paragraph = doc.add_paragraph()
            appendix_title_run = appendix_title_paragraph.add_run("Appendix")
            appendix_title_run.font.size = Pt(20)
            appendix_title_run.font.bold = True
            appendix_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Set background color for the Appendix title paragraph
            appendix_title_paragraph_format = appendix_title_paragraph.paragraph_format
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:val'), 'clear')
            shading_elm.set(qn('w:color'), 'auto')
            shading_elm.set(qn('w:fill'), '00B0F0') # Blue color
            appendix_title_paragraph_format._element.get_or_add_pPr().append(shading_elm)

            # General Conditions
            doc.add_paragraph() # Add a small space
            doc.add_paragraph("General Conditions", style='Heading 3')

            # Custom bullet points for General Conditions
            def add_custom_bullet(doc, text):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25) # Indent for bullet
                p.paragraph_format.first_line_indent = Inches(-0.25) # Outdent bullet
                run = p.add_run("•\t" + text) # Add bullet character and tab
                run.font.size = Pt(10) # Adjust font size if needed

            add_custom_bullet(doc, "Analysis Period is Last Six Months or the after the Last Event which ever is later")
            add_custom_bullet(doc, "Days with Good Weather (BF<=4) are considered for analysis.")
            add_custom_bullet(doc, "Days with Steaming hrs greater than 17 considered for analysis.")
            add_custom_bullet(doc, "Data is compared with Original Sea Trial")

            # Hull Performance
            doc.add_paragraph() # Add a small space
            doc.add_paragraph("Hull Performance", style='Heading 3')

            # Helper to add bullet points with specific colors and custom formatting
            def add_colored_custom_bullet(doc, text, color_rgb):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25) # Indent for bullet
                p.paragraph_format.first_line_indent = Inches(-0.25) # Outdent bullet
                run = p.add_run("•\t" + text) # Add bullet character and tab
                run.font.color.rgb = color_rgb
                run.font.size = Pt(10) # Adjust font size if needed

            add_colored_custom_bullet(doc, "Excess Power < 15 %– Rating Good", RGBColor(0, 176, 80)) # Green
            add_colored_custom_bullet(doc, "15< Excess Power < 25 % – Rating Average", RGBColor(255, 192, 0)) # Orange
            add_colored_custom_bullet(doc, "Excess Power > 25 % – Rating Poor", RGBColor(255, 0, 0)) # Red

            # Machinery Performance
            doc.add_paragraph() # Add a small space
            doc.add_paragraph("Machinery Performance", style='Heading 3')
            add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) within +/- 10 from Shop test condition are considered as \"Good\"", RGBColor(0, 176, 80)) # Green
            add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) Greater than 10 and less than 20 are considered as \"Average\"", RGBColor(255, 192, 0)) # Orange
            add_colored_custom_bullet(doc, "SFOC(Grms/kW.hr) Above 20 are considered as \"Poor\"", RGBColor(255, 0, 0)) # Red


        # Save to bytes buffer
        output = io.BytesIO()
        doc.save(output)
        return output.getvalue()

    except Exception as e:
        st.error(f"Error creating advanced Word report: {str(e)}")
        st.error(f"Error type: {type(e).__name__}")
        return None

# Function to reset session state
def reset_page():
    st.session_state.selected_vessels = set()
    st.session_state.report_data = None
    st.session_state.search_query = ""
    st.session_state.report_months = 2 # Reset to default
    st.cache_data.clear() # Clear cache for fresh vessel list
    # Set a flag to trigger rerun in the main loop
    st.session_state.trigger_rerun = True

# Main Application
def main():
    # Check for rerun flag
    if 'trigger_rerun' in st.session_state and st.session_state.trigger_rerun:
        st.session_state.trigger_rerun = False
        st.rerun()

    # Lambda Function URL
    LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/"

    # Title and header
    st.title("🚢 Enhanced Vessel Performance Report Tool")
    st.markdown("Select vessels and generate a comprehensive performance report with improved processing and UI.")

    # Add a reset button at the top
    st.button("🔄 Reset All", on_click=reset_page, type="secondary", help="Clear all selections and reset the page.")

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

        selected_vessels_list = list(st.session_state.selected_vessels)

        # Show selected vessels summary
        if selected_vessels_list:
            with st.expander(f"📋 Selected Vessels ({len(selected_vessels_list)})", expanded=False):
                for i, vessel in enumerate(sorted(selected_vessels_list), 1):
                    st.write(f"- {vessel}") # Changed to bullet points for better readability
    else:
        st.error("❌ Failed to load vessels. Please check your connection and try again.")
        selected_vessels_list = []

    # Generate report section
    st.header("2. Generate Enhanced Report")

    # Report duration selection
    st.session_state.report_months = st.radio(
        "Select Report Duration:",
        options=[1, 2, 3],
        format_func=lambda x: f"{x} Month{'s' if x > 1 else ''}",
        index=1, # Default to 2 months (index 1)
        horizontal=True,
        help="Choose to generate report for the previous 1, 2, or 3 months."
    )

    if selected_vessels_list:
        # Generate button with enhanced styling
        if st.button("🚀 Generate Enhanced Performance Report", type="primary", use_container_width=True):
            with st.spinner("Generating enhanced report with better progress tracking..."):
                try:
                    start_time = time.time()
                    st.session_state.report_data = query_report_data(
                        LAMBDA_FUNCTION_URL, selected_vessels_list, st.session_state.report_months
                    )

                    generation_time = time.time() - start_time

                    if not st.session_state.report_data.empty:
                        st.success(f"✅ Report generated successfully in {generation_time:.2f} seconds!")
                        # Removed st.balloons()
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

        # Display styled dataframe with enhanced presentation
        st.subheader("📋 Performance Data Table") # Changed heading
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
                    word_data = create_advanced_word_report(st.session_state.report_data, template_path, st.session_state.report_months)
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
        - Selected vessels summary with expandable list
        - Smart client-side filtering for responsive UI

        **📊 Better Data Processing:**
        - Enhanced progress tracking with visual progress bars
        - Improved error handling and user feedback
        - Success animations and better visual feedback

        **📈 Advanced Analytics:**
        - Hull condition and ME efficiency distribution charts
        - Multi-month trend analysis
        - Tabbed insights section for better organization

        **📥 Enhanced Downloads:**
        - Both Excel and CSV download options
        - Timestamped filenames
        - Styled Excel reports with color coding and text wrapping
        - Word reports with improved formatting, text wrapping, and specific comments for anomalous data
        - Better error handling for file generation

        ### 📋 How to Use:

        1. **🔍 Search & Filter**: Type in the search box to find specific vessels
        2. **✅ Select Vessels**: Use checkboxes to select vessels
        3. **🚀 Generate**: Click the generate button for enhanced processing
        4. **📊 Analyze**: Review metrics, charts, and trends in the results
        5. **📥 Download**: Export your report in Excel, CSV, or Word format

        ### 📊 Report Columns:

        **🛡️ Hull Condition** (Multiple months):
        - 🟢 **Good**: < 15% power loss (Green)
        - 🟡 **Average**: 15-25% power loss (Yellow)
        - 🔴 **Poor**: > 25% power loss (Red)

        **⚙️ ME Efficiency** (Multiple months):
        - ⚪ **Anomalous data**: < 160 SFOC (Gray) - *A comment will be added for these entries.*
        - 🟢 **Good**: 160-180 SFOC (Green)
        - 🟡 **Average**: 180-190 SFOC (Yellow)
        - 🔴 **Poor**: > 190 SFOC (Red)

        **📊 Additional Metrics:**
        - ⛽ **Potential Fuel Saving (MT/Day)**: Excess consumption (MT/day)
        - 📈 **YTD CII**: Carbon Intensity Indicator rating
        - 💬 **Comments**
        """)

if __name__ == "__main__":
    main()
