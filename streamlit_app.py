import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.colors import Color

# Page configuration
st.set_page_config(
    page_title="Vessel Data & Report Tool",
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

# --- Lambda Invocation Helper (IMPROVED WITH BETTER ERROR HANDLING) ---
def invoke_lambda_function_url(lambda_url, payload):
    """Invoke Lambda function via its Function URL using HTTP POST."""
    st.info(f"Attempting to invoke Lambda URL: {lambda_url}")
    
    try:
        headers = {'Content-Type': 'application/json'}
        # Ensure payload is properly formatted
        json_payload = json.dumps(payload)
        
        st.info(f"Sending payload: {json_payload}")
        
        response = requests.post(
            lambda_url, 
            headers=headers, 
            data=json_payload,
            timeout=30  # Increased timeout to 30 seconds
        )
        
        st.info(f"Response status code: {response.status_code}")
        
        # Try to get response text regardless of status code
        try:
            response_text = response.text[:500]  # First 500 chars to avoid flooding the UI
            st.info(f"Response text preview: {response_text}...")
        except:
            st.warning("Could not get response text")
        
        # Now raise for status to handle errors
        response.raise_for_status()
        
        # Try to parse JSON response
        try:
            return response.json()
        except json.JSONDecodeError:
            st.error(f"Response is not valid JSON: {response.text[:200]}...")
            return None
            
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
    """Fetch all vessel names from Lambda function using its URL with progressive fallbacks."""
    # First try with a LIMIT to see if it's a size/timeout issue
    test_query = "SELECT vessel_name FROM vessel_particulars LIMIT 10"
    st.info("Testing vessel query with LIMIT 10...")
    
    test_result = invoke_lambda_function_url(lambda_url, {"sql_query": test_query})
    
    if test_result:
        st.success("Limited query successful, now trying with LIMIT 100...")
        # Try with a larger limit
        medium_query = "SELECT vessel_name FROM vessel_particulars LIMIT 100"
        medium_result = invoke_lambda_function_url(lambda_url, {"sql_query": medium_query})
        
        if medium_result:
            st.success("LIMIT 100 query successful, now trying with LIMIT 1000...")
            # Try with an even larger limit
            large_query = "SELECT vessel_name FROM vessel_particulars LIMIT 1000"
            large_result = invoke_lambda_function_url(lambda_url, {"sql_query": large_query})
            
            if large_result:
                extracted_vessel_names = []
                for item in large_result:
                    if isinstance(item, dict) and 'vessel_name' in item:
                        extracted_vessel_names.append(item['vessel_name'])
                    elif isinstance(item, str):
                        extracted_vessel_names.append(item)
                extracted_vessel_names.sort()
                st.success(f"Loaded {len(extracted_vessel_names)} vessels.")
                return extracted_vessel_names
            else:
                # Fall back to medium result
                st.warning("LIMIT 1000 query failed, using LIMIT 100 results instead.")
                extracted_vessel_names = []
                for item in medium_result:
                    if isinstance(item, dict) and 'vessel_name' in item:
                        extracted_vessel_names.append(item['vessel_name'])
                    elif isinstance(item, str):
                        extracted_vessel_names.append(item)
                extracted_vessel_names.sort()
                return extracted_vessel_names
        else:
            # Fall back to test result
            st.warning("LIMIT 100 query failed, using LIMIT 10 results instead.")
            extracted_vessel_names = []
            for item in test_result:
                if isinstance(item, dict) and 'vessel_name' in item:
                    extracted_vessel_names.append(item['vessel_name'])
                elif isinstance(item, str):
                    extracted_vessel_names.append(item)
            extracted_vessel_names.sort()
            return extracted_vessel_names
    
    # If all else fails, try a different approach - query for vessel count first
    st.warning("All vessel queries failed. Trying alternative approach...")
    count_query = "SELECT COUNT(*) as count FROM vessel_particulars"
    count_result = invoke_lambda_function_url(lambda_url, {"sql_query": count_query})
    
    if count_result and isinstance(count_result, list) and len(count_result) > 0:
        try:
            vessel_count = int(count_result[0]['count'])
            st.info(f"Found {vessel_count} vessels in database. Fetching in batches...")
            
            # Fetch in batches of 50
            all_vessels = []
            batch_size = 50
            for offset in range(0, vessel_count, batch_size):
                batch_query = f"SELECT vessel_name FROM vessel_particulars ORDER BY vessel_name LIMIT {batch_size} OFFSET {offset}"
                st.info(f"Fetching batch {offset//batch_size + 1} of {(vessel_count + batch_size - 1)//batch_size}...")
                
                batch_result = invoke_lambda_function_url(lambda_url, {"sql_query": batch_query})
                if batch_result:
                    for item in batch_result:
                        if isinstance(item, dict) and 'vessel_name' in item:
                            all_vessels.append(item['vessel_name'])
                        elif isinstance(item, str):
                            all_vessels.append(item)
                else:
                    st.warning(f"Failed to fetch batch at offset {offset}")
            
            if all_vessels:
                all_vessels.sort()
                st.success(f"Loaded {len(all_vessels)} vessels using batch approach.")
                return all_vessels
        except (KeyError, ValueError, TypeError) as e:
            st.error(f"Error processing count result: {str(e)}")
    
    st.error("Failed to load vessel names using all approaches.")
    return []

def query_report_data(lambda_url, vessel_names):
    """Fetch hull roughness power loss, ME SFOC, and Potential Fuel Saving for selected vessels and process for report."""
    if not vessel_names:
        return pd.DataFrame() # Return empty DataFrame if no vessels selected

    # Process vessels in smaller batches to avoid timeout/size issues
    batch_size = 10
    all_hull_data = []
    all_me_data = []
    all_fuel_saving_data = []
    
    for i in range(0, len(vessel_names), batch_size):
        batch_vessels = vessel_names[i:i+batch_size]
        st.info(f"Processing batch {i//batch_size + 1} of {(len(vessel_names) + batch_size - 1)//batch_size} ({len(batch_vessels)} vessels)")
        
        quoted_vessel_names = [f"'{name}'" for name in batch_vessels]
        vessel_names_list_str = ", ".join(quoted_vessel_names)

        # --- Query 1: Hull Roughness Power Loss ---
        sql_query_hull = f"SELECT vessel_name, hull_rough_power_loss_pct_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
        
        st.info(f"Fetching Hull Roughness data for batch...")
        
        hull_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_hull})
        
        if hull_result:
            all_hull_data.extend(hull_result)
        
        # --- Query 2: ME SFOC ---
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
            vp.vessel_name;
        """
        
        st.info(f"Fetching ME SFOC data for batch...")
        
        me_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_me})

        if me_result:
            all_me_data.extend(me_result)
        
        # --- Query 3: Potential Fuel Saving ---
        sql_query_fuel_saving = f"SELECT vessel_name, hull_rough_excess_consumption_mt_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
        
        st.info(f"Fetching Potential Fuel Saving data for batch...")
        
        fuel_saving_result = invoke_lambda_function_url(lambda_url, {"sql_query": sql_query_fuel_saving})

        if fuel_saving_result:
            all_fuel_saving_data.extend(fuel_saving_result)

    # Process all collected data
    df_hull = pd.DataFrame()
    if all_hull_data:
        try:
            df_hull = pd.DataFrame(all_hull_data)
            if 'hull_rough_power_loss_pct_ed' in df_hull.columns:
                df_hull = df_hull.rename(columns={'hull_rough_power_loss_pct_ed': 'Hull Roughness Power Loss %'})
            else:
                st.warning("Column 'hull_rough_power_loss_pct_ed' not found in Lambda response for hull data.")
                df_hull['Hull Roughness Power Loss %'] = pd.NA
            df_hull = df_hull.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing hull data: {str(e)}")
            df_hull = pd.DataFrame()
    else:
        st.error("Failed to retrieve hull roughness data.")

    df_me = pd.DataFrame()
    if all_me_data:
        try:
            df_me = pd.DataFrame(all_me_data)
            if 'avg_me_sfoc' in df_me.columns:
                df_me = df_me.rename(columns={'avg_me_sfoc': 'ME SFOC'})
            else:
                st.warning("Column 'avg_me_sfoc' not found in Lambda response for ME data.")
                df_me['ME SFOC'] = pd.NA
            df_me = df_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing ME data: {str(e)}")
            df_me = pd.DataFrame()
    else:
        st.error("Failed to retrieve ME SFOC data.")

    df_fuel_saving = pd.DataFrame()
    if all_fuel_saving_data:
        try:
            df_fuel_saving = pd.DataFrame(all_fuel_saving_data)
            if 'hull_rough_excess_consumption_mt_ed' in df_fuel_saving.columns:
                df_fuel_saving = df_fuel_saving.rename(columns={'hull_rough_excess_consumption_mt_ed': 'Potential Fuel Saving'})
                # Apply the capping logic: if > 5, set to 4.9; if < 0, set to 0
                df_fuel_saving['Potential Fuel Saving'] = df_fuel_saving['Potential Fuel Saving'].apply(
                    lambda x: 4.9 if pd.notna(x) and x > 5 else (0.0 if pd.notna(x) and x < 0 else x) # Ensure 0.0 for float consistency
                )
            else:
                st.warning("Column 'hull_rough_excess_consumption_mt_ed' not found in Lambda response for fuel saving data.")
                df_fuel_saving['Potential Fuel Saving'] = pd.NA
            df_fuel_saving = df_fuel_saving.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing fuel saving data: {str(e)}")
            df_fuel_saving = pd.DataFrame()
    else:
        st.error("Failed to retrieve Potential Fuel Saving data.")

    # --- Merge DataFrames ---
    df_final = pd.DataFrame()
    # Start with a DataFrame containing all selected vessel names to ensure all are present
    # even if they have no data for specific metrics.
    df_final = pd.DataFrame({'Vessel Name': list(vessel_names)}) 

    if not df_hull.empty:
        df_final = pd.merge(df_final, df_hull, on='Vessel Name', how='left')
    
    if not df_me.empty:
        df_final = pd.merge(df_final, df_me, on='Vessel Name', how='left')
            
    if not df_fuel_saving.empty:
        df_final = pd.merge(df_final, df_fuel_saving, on='Vessel Name', how='left')

    if df_final.empty:
        st.warning("No data retrieved for any of the requested metrics.")
        return pd.DataFrame()

    # --- Post-merge processing for final report ---
    # Add S. No. column (after merge to ensure correct numbering)
    df_final.insert(0, 'S. No.', range(1, 1 + len(df_final)))
    
    # Add Hull Condition column
    def get_hull_condition(value):
        if pd.isna(value):
            return "N/A"
        if value < 15:
            return "Good"
        elif 15 <= value <= 25:
            return "Average"
        else: # value > 25
            return "Poor"
    
    if 'Hull Roughness Power Loss %' in df_final.columns:
        df_final['Hull Condition'] = df_final['Hull Roughness Power Loss %'].apply(get_hull_condition)
    else:
        df_final['Hull Condition'] = "N/A" # Default if column is missing

    # Add ME Efficiency column
    def get_me_efficiency(value):
        if pd.isna(value):
            return "N/A"
        if value < 160: # New logic: Anomalous data
            return "Anomalous data"
        elif value < 180:
            return "Good"
        elif 180 <= value <= 190:
            return "Average"
        else: # value > 190
            return "Poor"
    
    if 'ME SFOC' in df_final.columns:
        df_final['ME Efficiency'] = df_final['ME SFOC'].apply(get_me_efficiency)
    else:
        df_final['ME Efficiency'] = "N/A" # Default if ME SFOC is missing

    # Add empty Comments column
    df_final['Comments'] = ""

    # Define the desired order of columns
    desired_columns_order = [
        'S. No.', 
        'Vessel Name', 
        'Hull Condition', 
        'ME Efficiency', 
        'Potential Fuel Saving',
        'Comments', # New empty column
        'Hull Roughness Power Loss %', # Moved to last
        'ME SFOC' # Moved to last
    ]
    
    # Filter df_final to only include columns that exist and are in the desired order
    # This ensures that if a column (like 'ME SFOC' or 'Hull Roughness Power Loss %')
    # was not retrieved, it won't cause an error, and the 'Condition' columns
    # will still be created if their source data exists.
    existing_and_ordered_columns = [col for col in desired_columns_order if col in df_final.columns]
    df_final = df_final[existing_and_ordered_columns]

    st.success("Report data retrieved and processed successfully!")
    return df_final

# --- Styling for Streamlit DataFrame ---
def style_condition_columns(row):
    styles = [''] * len(row)
    
    # Hull Condition styling
    if 'Hull Condition' in row.index:
        hull_val = row['Hull Condition']
        if hull_val == "Good":
            styles[row.index.get_loc('Hull Condition')] = 'background-color: #d4edda; color: black;' # Light green
        elif hull_val == "Average":
            styles[row.index.get_loc('Hull Condition')] = 'background-color: #fff3cd; color: black;' # Light orange
        elif hull_val == "Poor":
            styles[row.index.get_loc('Hull Condition')] = 'background-color: #f8d7da; color: black;' # Light red
    
    # ME Efficiency styling
    if 'ME Efficiency' in row.index:
        me_val = row['ME Efficiency']
        if me_val == "Good":
            styles[row.index.get_loc('ME Efficiency')] = 'background-color: #d4edda; color: black;' # Light green
        elif me_val == "Average":
            styles[row.index.get_loc('ME Efficiency')] = 'background-color: #fff3cd; color: black;' # Light orange
        elif me_val == "Poor":
            styles[row.index.get_loc('ME Efficiency')] = 'background-color: #f8d7da; color: black;' # Light red
        elif me_val == "Anomalous data": # New styling for anomalous data
            styles[row.index.get_loc('ME Efficiency')] = 'background-color: #e0e0e0; color: black;' # Light grey
            
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
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value) # +2 because headers are row 1, data starts row 2 (0-indexed df, 1-indexed excel)
            
            # Apply styling for 'Hull Condition' and 'ME Efficiency' columns
            if col_name in ['Hull Condition', 'ME Efficiency']:
                if cell_value == "Good":
                    cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Light green
                elif cell_value == "Average":
                    cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Light orange
                elif cell_value == "Poor":
                    cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Light red
                elif cell_value == "Anomalous data": # New styling for anomalous data in Excel
                    cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid") # Light grey
                cell.font = Font(color="000000") # Black font color

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
        adjusted_width = (max_length + 2) * 1.2 # Add a little padding
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output)
    return output.getvalue()

# --- Main App Layout ---
st.title("🚢 Vessel Performance Report Tool")
st.markdown("Select vessels and generate a report on their excess power.")

# --- Hardcode Lambda Function URL here ---
LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/" 

# Sidebar for Configuration
with st.sidebar:
    st.header("Configuration")
    
    # Display the hardcoded URL, or allow input if you prefer
    lambda_url = st.text_input(
        "Lambda Function URL",
        value=LAMBDA_FUNCTION_URL,
        help="The URL of your AWS Lambda Function (Auth type: NONE)"
    )
    
    # Add a test connection button
    if st.button("Test Lambda Connection"):
        test_result = invoke_lambda_function_url(
            lambda_url, 
            {"sql_query": "SELECT 1 as test"}
        )
        if test_result:
            st.success("✅ Connection successful!")
        else:
            st.error("❌ Connection failed. Check URL and Lambda configuration.")
    
    st.info("💡 **Lambda Function URL**: This is the HTTP endpoint for your Lambda. Ensure its 'Auth type' is set to 'NONE' for public access.")

# --- Vessel Loading ---
if st.button("Load Vessels") or (st.session_state.vessels and lambda_url == LAMBDA_FUNCTION_URL):
    st.session_state.vessels = fetch_all_vessels(lambda_url)

# --- Main Content Area ---
st.header("1. Select Vessels")

if st.session_state.vessels:
    # New UI/UX for vessel selection
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
            for vessel in filtered_vessels:
                checkbox_state = st.checkbox(
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
    st.info("Click 'Load Vessels' to fetch vessel names from the database.")
    selected_vessels_list = []

st.header("2. Generate Report")

if selected_vessels_list:
    if st.button("🚀 Generate Excess Power Report", type="primary"):
        with st.spinner("Generating report..."):
            st.session_state.report_data = query_report_data(
                lambda_url, 
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
    
    1. **Load Vessels**: Click the "Load Vessels" button to fetch all vessel names from the database.
    2. **Select Vessels**: Use the search bar to filter vessels, then check the boxes to select them.
    3. **Generate Report**: Click "Generate Excess Power Report" to fetch the data and display it.
    4. **Download**: If data is available, a download button will appear to get the report as an Excel file.
    
    ### Troubleshooting:
    
    - If vessel loading fails, check the Lambda Function URL in the sidebar.
    - Use the "Test Lambda Connection" button to verify connectivity.
    - The app will try to load vessels in smaller batches if the full query fails.
    - Check the console logs for detailed error messages.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈 and Python*")
