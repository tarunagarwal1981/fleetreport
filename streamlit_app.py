import streamlit as st
import boto3
import pandas as pd
import json
import io
from datetime import datetime

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
    """Fetch hull roughness power loss for selected vessels."""
    if not vessel_names:
        return pd.DataFrame() # Return empty DataFrame if no vessels selected

    vessel_names_list_str = ", ".join([f"'{name}'" for name in vessel_names])
    sql_query_string = f"SELECT vessel_name, hull_roughness_power_loss FROM hull_performance_six_months WHERE vessel_name IN ({vessel_names_list_str});"
    
    st.info(f"Generating report for {len(vessel_names)} vessels...")
    st.code(sql_query_string, language="sql") # Show the query being sent

    result = invoke_lambda_function(function_name, {"sql_query": sql_query_string}, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        try:
            df = pd.DataFrame(result)
            # Rename columns for display
            df = df.rename(columns={
                'vessel_name': 'Vessel Name',
                'hull_roughness_power_loss': 'Excess Power'
            })
            st.success("Report data retrieved successfully!")
            return df
        except Exception as e:
            st.error(f"Error processing report data: {str(e)}")
            st.json(result) # Show raw result for debugging
            return pd.DataFrame()
    st.error("Failed to retrieve report data.")
    return pd.DataFrame()

def create_excel_download(data, filename):
    """Convert DataFrame to Excel format for download"""
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Vessel Data', index=False)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

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
    st.dataframe(st.session_state.report_data, use_container_width=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"excess_power_report_{timestamp}.xlsx"

    excel_data = create_excel_download(st.session_state.report_data, filename)

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
