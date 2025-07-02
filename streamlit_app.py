import streamlit as st
import boto3
import pandas as pd
import json
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
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
    """Fetch hull roughness power loss for selected vessels and process for report."""
    if not vessel_names:
        return pd.DataFrame() # Return empty DataFrame if no vessels selected

    # Corrected way to format vessel names for SQL IN clause
    quoted_vessel_names = [f"'{name}'" for name in vessel_names]
    vessel_names_list_str = ", ".join(quoted_vessel_names)

    # Ensure the column name matches your database exactly
    sql_query_string = f"SELECT vessel_name, hull_rough_power_loss_pct_ed FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
    
    st.info(f"Generating report for {len(vessel_names)} vessels...")
    st.code(sql_query_string, language="sql") # Show the query being sent

    result = invoke_lambda_function(function_name, {"sql_query": sql_query_string}, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        try:
            df = pd.DataFrame(result)
            
            # Ensure the column exists before renaming
            if 'hull_rough_power_loss_pct_ed' in df.columns:
                df = df.rename(columns={'hull_rough_power_loss_pct_ed': 'Hull Roughness Power Loss %'})
            else:
                st.warning("Column 'hull_rough_power_loss_pct_ed' not found in Lambda response. Please check your query or Lambda output.")
                # If the column is missing, create it with NaNs to avoid errors
                df['Hull Roughness Power Loss %'] = pd.NA

            # Add S. No. column
            df.insert(0, 'S. No.', range(1, 1 + len(df)))
            
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
            
            df['Hull Condition'] = df['Hull Roughness Power Loss %'].apply(get_hull_condition)

            # Reorder columns for final display
            df = df[['S. No.', 'vessel_name', 'Hull Condition', 'Hull Roughness Power Loss %']]
            df = df.rename(columns={'vessel_name': 'Vessel Name'}) # Rename vessel_name after reordering

            st.success("Report data retrieved and processed successfully!")
            return df
        except Exception as e:
            st.error(f"Error processing report data: {str(e)}")
            st.json(result) # Show raw result for debugging
            return pd.DataFrame()
    st.error("Failed to retrieve report data.")
    return pd.DataFrame()

# --- Styling for Streamlit DataFrame ---
def style_hull_condition(val):
    if val == "Good":
        return 'background-color: #d4edda; color: black;' # Light green
    elif val == "Average":
        return 'background-color: #fff3cd; color: black;' # Light orange
    elif val == "Poor":
        return 'background-color: #f8d7da; color: black;' # Light red
    return ''

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
            
            # Apply styling for 'Hull Condition' column
            if col_name == 'Hull Condition':
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
    styled_df = st.session_state.report_data.style.applymap(
        style_hull_condition, subset=['Hull Condition']
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
