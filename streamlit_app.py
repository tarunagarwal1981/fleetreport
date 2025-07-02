import streamlit as st
import boto3
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
if 'selected_vessels' not in st.session_state:
    st.session_state.selected_vessels = []
if 'report_data' not in st.session_state:
    st.session_state.report_data = None

# --- Lambda Invocation Helper ---
def invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token=None, region='ap-south-1'):
    """Invoke Lambda function directly using AWS SDK"""
    try:
        lambda_client = boto3.client(
            'lambda',
            aws_access_key_id=aws_access_key,
            aws_secret_access_key=aws_secret_key,
            aws_session_token=aws_session_token,
            region_name=region
        )
        
        response = lambda_client.invoke(
            FunctionName=function_name,
            InvocationType='RequestResponse',
            Payload=json.dumps(payload)
        )
        
        response_payload = json.loads(response['Payload'].read())
        
        if response.get('StatusCode') == 200:
            if 'statusCode' in response_payload and 'body' in response_payload:
                if response_payload['statusCode'] != 200:
                    st.error(f"Lambda returned error status: {response_payload.get('body', 'Unknown error')}")
                    return None
                if isinstance(response_payload['body'], str):
                    return json.loads(response_payload['body'])
                else:
                    return response_payload['body']
            else:
                return response_payload
        else:
            st.error(f"AWS invoke error (non-200 status from AWS API): {response_payload}")
            return None
            
    except Exception as e:
        st.error(f"Error invoking Lambda: {str(e)}")
        return None

# --- Data Fetching Functions ---
@st.cache_data(ttl=3600) # Cache results for 1 hour to avoid re-fetching on every rerun
def fetch_all_vessels(function_name, aws_access_key, aws_secret_key, aws_session_token):
    """Fetch all vessel names from Lambda function."""
    query = "select vessel_name from vessel_particulars"
    st.info("Loading all vessel names...")
    result = invoke_lambda_function(function_name, {"sql_query": query}, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        extracted_vessel_names = []
        for item in result:
            if isinstance(item, dict) and 'vessel_name' in item:
                extracted_vessel_names.append(item['vessel_name'])
            elif isinstance(item, str): # In case the Lambda returns a list of strings directly
                extracted_vessel_names.append(item)
        st.success(f"Loaded {len(extracted_vessel_names)} vessels.")
        return extracted_vessel_names
    st.error("Failed to load vessel names.")
    return []

def query_report_data(function_name, vessel_names, aws_access_key, aws_secret_key, aws_session_token):
    """Fetch hull roughness power loss, ME SFOC, and Potential Fuel Saving for selected vessels and process for report."""
    if not vessel_names:
        return pd.DataFrame() # Return empty DataFrame if no vessels selected

    quoted_vessel_names = [f"'{name}'" for name in vessel_names]
    vessel_names_list_str = ", ".join(quoted_vessel_names)

    # --- Query 1: Hull Roughness Power Loss ---
    sql_query_hull = f"SELECT vessel_name, hull_rough_power_loss_pct_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
    
    st.info(f"Fetching Hull Roughness data for {len(vessel_names)} vessels...")
    st.code(sql_query_hull, language="sql")
    
    hull_result = invoke_lambda_function(function_name, {"sql_query": sql_query_hull}, aws_access_key, aws_secret_key, aws_session_token)
    
    df_hull = pd.DataFrame()
    if hull_result:
        try:
            df_hull = pd.DataFrame(hull_result)
            if 'hull_rough_power_loss_pct_ed' in df_hull.columns:
                df_hull = df_hull.rename(columns={'hull_rough_power_loss_pct_ed': 'Hull Roughness Power Loss %'})
            else:
                st.warning("Column 'hull_rough_power_loss_pct_ed' not found in Lambda response for hull data.")
                df_hull['Hull Roughness Power Loss %'] = pd.NA
            df_hull = df_hull.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing hull data: {str(e)}")
            st.json(hull_result)
            df_hull = pd.DataFrame()
    else:
        st.error("Failed to retrieve hull roughness data.")

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
    
    st.info(f"Fetching ME SFOC data for {len(vessel_names)} vessels...")
    st.code(sql_query_me, language="sql")

    me_result = invoke_lambda_function(function_name, {"sql_query": sql_query_me}, aws_access_key, aws_secret_key, aws_session_token)

    df_me = pd.DataFrame()
    if me_result:
        try:
            df_me = pd.DataFrame(me_result)
            if 'avg_me_sfoc' in df_me.columns:
                df_me = df_me.rename(columns={'avg_me_sfoc': 'ME SFOC'})
            else:
                st.warning("Column 'avg_me_sfoc' not found in Lambda response for ME data.")
                df_me['ME SFOC'] = pd.NA
            df_me = df_me.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing ME data: {str(e)}")
            st.json(me_result)
            df_me = pd.DataFrame()
    else:
        st.error("Failed to retrieve ME SFOC data.")

    # --- Query 3: Potential Fuel Saving ---
    sql_query_fuel_saving = f"SELECT vessel_name, hull_rough_excess_consumption_mt_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
    
    st.info(f"Fetching Potential Fuel Saving data for {len(vessel_names)} vessels...")
    st.code(sql_query_fuel_saving, language="sql")

    fuel_saving_result = invoke_lambda_function(function_name, {"sql_query": sql_query_fuel_saving}, aws_access_key, aws_secret_key, aws_session_token)

    df_fuel_saving = pd.DataFrame()
    if fuel_saving_result:
        try:
            df_fuel_saving = pd.DataFrame(fuel_saving_result)
            if 'hull_rough_excess_consumption_mt_ed' in df_fuel_saving.columns:
                df_fuel_saving = df_fuel_saving.rename(columns={'hull_rough_excess_consumption_mt_ed': 'Potential Fuel Saving'})
                # Apply the capping logic: if > 5, set to 4.9
                df_fuel_saving['Potential Fuel Saving'] = df_fuel_saving['Potential Fuel Saving'].apply(lambda x: 4.9 if pd.notna(x) and x > 5 else x)
            else:
                st.warning("Column 'hull_rough_excess_consumption_mt_ed' not found in Lambda response for fuel saving data.")
                df_fuel_saving['Potential Fuel Saving'] = pd.NA
            df_fuel_saving = df_fuel_saving.rename(columns={'vessel_name': 'Vessel Name'})
        except Exception as e:
            st.error(f"Error processing fuel saving data: {str(e)}")
            st.json(fuel_saving_result)
            df_fuel_saving = pd.DataFrame()
    else:
        st.error("Failed to retrieve Potential Fuel Saving data.")

    # --- Merge DataFrames ---
    # Start with df_hull, then merge df_me, then merge df_fuel_saving
    df_final = pd.DataFrame()
    if not df_hull.empty:
        df_final = df_hull
    
    if not df_me.empty:
        if df_final.empty:
            df_final = df_me
        else:
            df_final = pd.merge(df_final, df_me, on='Vessel Name', how='outer')
            
    if not df_fuel_saving.empty:
        if df_final.empty:
            df_final = df_fuel_saving
        else:
            df_final = pd.merge(df_final, df_fuel_saving, on='Vessel Name', how='outer')

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
        if value < 180:
            return "Good"
        elif 180 <= value <= 190:
            return "Average"
        else: # value > 190
            return "Poor"
    
    if 'ME SFOC' in df_final.columns:
        df_final['ME Efficiency'] = df_final['ME SFOC'].apply(get_me_efficiency)
    else:
        df_final['ME Efficiency'] = "N/A" # Default if ME SFOC is missing

    # Define the desired order of columns
    desired_columns_order = [
        'S. No.', 
        'Vessel Name', 
        'Hull Condition', 
        'Hull Roughness Power Loss %', 
        'ME Efficiency', 
        'ME SFOC',
        'Potential Fuel Saving' # New column
    ]
    
    # Filter df_final to only include columns that exist and are in the desired order
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

# Sidebar for AWS Configuration
with st.sidebar:
    st.header("AWS Configuration")
    
    aws_access_key = st.text_input(
        "AWS Access Key ID",
        type="password",
        help="Your AWS Access Key ID"
    )
    
    aws_secret_key = st.text_input(
        "AWS Secret Access Key",
        type="password",
        help="Your AWS Secret Access Key"
    )
    
    aws_session_token = st.text_input(
        "AWS Session Token (Optional)",
        type="password",
        help="Required if using temporary credentials (e.g., from STS AssumeRole)"
    )
    
    function_name = st.text_input(
        "Lambda Function Name",
        value="",
        help="The name of your Lambda function (e.g., 'my-vessel-query-function')"
    )
    
    st.info("💡 **Function Name**: Find this in your AWS Lambda console. It's the name, not the full ARN or URL.")
    st.warning("⚠️ **Permissions**: Ensure the provided AWS credentials have `lambda:InvokeFunction` permission on your Lambda.")

# --- Automatic Vessel Loading ---
# Only attempt to fetch if credentials and function name are provided
if all([aws_access_key, aws_secret_key, function_name]) and not st.session_state.vessels:
    st.session_state.vessels = fetch_all_vessels(function_name, aws_access_key, aws_secret_key, aws_session_token)

# --- Main Content Area ---
st.header("1. Select Vessels")

if st.session_state.vessels:
    # Use st.multiselect for better handling of many options
    st.session_state.selected_vessels = st.multiselect(
        "Choose Vessels for Report:",
        options=st.session_state.vessels,
        default=st.session_state.selected_vessels, # Keep previously selected
        placeholder="Search and select vessels..."
    )
    st.info(f"📊 {len(st.session_state.vessels)} vessels available. {len(st.session_state.selected_vessels)} selected.")
else:
    st.warning("Please provide AWS credentials and Lambda Function Name in the sidebar to load vessels.")

st.header("2. Generate Report")

if st.session_state.selected_vessels and all([aws_access_key, aws_secret_key, function_name]):
    if st.button("🚀 Generate Excess Power Report", type="primary"):
        with st.spinner("Generating report..."):
            st.session_state.report_data = query_report_data(
                function_name, 
                st.session_state.selected_vessels, 
                aws_access_key, 
                aws_secret_key, 
                aws_session_token
            )
else:
    missing_items = []
    if not st.session_state.selected_vessels:
        missing_items.append("Select vessels")
    if not all([aws_access_key, aws_secret_key, function_name]):
        missing_items.append("Configure AWS credentials and Lambda Function Name")
    
    if missing_items:
        st.warning(f"Please complete: {', '.join(missing_items)}")

# --- Display Report ---
if st.session_state.report_data is not None and not st.session_state.report_data.empty:
    st.header("3. Report Results")
    
    # Apply styling for Streamlit dataframe
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
    ### Setup Requirements:
    
    1.  **AWS Credentials**: You need AWS Access Key ID, Secret Access Key, and optionally a Session Token.
        *   **Access Key ID & Secret Access Key**: These are long-term credentials for an IAM user.
        *   **Session Token**: This is required if your credentials are *temporary* (e.g., obtained from AWS STS `AssumeRole` or from an EC2 instance role). If you're using long-term IAM user keys, you typically won't have a session token.
    2.  **Lambda Function Name**: The actual name of your Lambda function (e.g., `my-vessel-query-function`).
    
    ### Step-by-step Guide:
    
    1.  **Configure AWS**: Enter your AWS credentials and Lambda function name in the sidebar. Vessels will attempt to load automatically.
    2.  **Select Vessels**: Use the multi-select dropdown to choose the vessels you want to include in the report. You can search within the dropdown.
    3.  **Generate Report**: Click "Generate Excess Power Report" to fetch the data and display it.
    4.  **Download**: If data is available, a download button will appear to get the report as an Excel file.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈 and AWS SDK*")
