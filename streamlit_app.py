import streamlit as st
import boto3
import pandas as pd
import json
import io
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Vessel Data Export & Report",
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

# --- AWS Lambda Invocation Functions ---
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

def fetch_vessel_names_from_lambda(function_name, aws_access_key, aws_secret_key, aws_session_token):
    """Fetches vessel names using a specific SQL query."""
    query = "select vessel_name from vessel_particulars"
    payload = {"sql_query": query}
    
    result = invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        extracted_names = []
        for item in result:
            if isinstance(item, dict) and 'vessel_name' in item:
                extracted_names.append(item['vessel_name'])
            elif isinstance(item, str):
                extracted_names.append(item)
        return extracted_names
    return []

def query_vessel_data_from_lambda(function_name, sql_query_string, aws_access_key, aws_secret_key, aws_session_token):
    """Sends a SQL query to Lambda and gets results."""
    payload = {"sql_query": sql_query_string}
    return invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token)

# --- Excel Download Function ---
def create_excel_download(data, filename):
    """Convert data to Excel format for download"""
    try:
        if isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, dict):
            if 'data' in data:
                df = pd.DataFrame(data['data'])
            elif 'results' in data:
                df = pd.DataFrame(data['results'])
            else:
                df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Vessel Data', index=False)

        return buffer.getvalue()
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

# --- Main App Layout ---
st.title("🚢 Vessel Performance Report Tool")
st.markdown("Select vessels and generate a report on their hull performance.")

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
    
    lambda_function_name = st.text_input(
        "Lambda Function Name",
        value="",
        help="The name of your Lambda function (e.g., 'my-vessel-query-function')"
    )
    
    st.info("💡 **Function Name**: Find this in your AWS Lambda console. It's the name, not the full ARN or URL.")
    st.warning("⚠️ **Permissions**: Ensure the provided AWS credentials have `lambda:InvokeFunction` permission on your Lambda.")

# --- Auto-fetch vessels on page load ---
if lambda_function_name and aws_access_key and aws_secret_key and not st.session_state.vessels:
    with st.spinner("Fetching vessel list..."):
        fetched_vessels = fetch_vessel_names_from_lambda(
            lambda_function_name, aws_access_key, aws_secret_key, aws_session_token
        )
        if fetched_vessels:
            st.session_state.vessels = sorted(list(set(fetched_vessels))) # Ensure unique and sorted
            st.success(f"Loaded {len(st.session_state.vessels)} vessels.")
        else:
            st.error("Failed to load vessels. Please check AWS configuration and Lambda logs.")

# --- Main Content Area ---
st.header("1. Select Vessels")

if st.session_state.vessels:
    # Multiselect for vessel selection
    st.session_state.selected_vessels = st.multiselect(
        "Choose Vessels for Report:",
        options=st.session_state.vessels,
        default=st.session_state.selected_vessels, # Maintain selection across reruns
        placeholder="Select one or more vessels...",
        help="Type to search or select from the list."
    )
    
    if st.session_state.selected_vessels:
        st.success(f"✅ {len(st.session_state.selected_vessels)} vessels selected.")
    else:
        st.info("No vessels selected yet.")
else:
    st.warning("Please configure AWS credentials and Lambda Function Name in the sidebar to load vessels.")

st.header("2. Generate Report")

if st.button("📊 Generate Performance Report", type="primary", 
             disabled=not (st.session_state.selected_vessels and lambda_function_name and aws_access_key and aws_secret_key)):
    
    if not st.session_state.selected_vessels:
        st.warning("Please select at least one vessel to generate a report.")
    else:
        report_results = []
        with st.spinner("Fetching performance data... This may take a moment for many vessels."):
            for vessel_name in st.session_state.selected_vessels:
                # Construct the specific query for hull_roughness_power_loss
                report_query = f"select hull_roughness_power_loss from sdb_reporting_layer.digital_desk.hull_performance_six_months where vessel_name = '{vessel_name}';"
                
                data = query_vessel_data_from_lambda(
                    lambda_function_name, report_query, aws_access_key, aws_secret_key, aws_session_token
                )
                
                if data and len(data) > 0:
                    # Assuming hull_roughness_power_loss is the first (or only) key in the returned dict
                    # And we want the first value if multiple are returned (e.g., if query returns multiple rows)
                    excess_power_value = data[0].get('hull_roughness_power_loss')
                    report_results.append({
                        "Vessel Name": vessel_name,
                        "Excess Power": excess_power_value if excess_power_value is not None else "N/A"
                    })
                else:
                    report_results.append({
                        "Vessel Name": vessel_name,
                        "Excess Power": "No Data"
                    })
        
        if report_results:
            st.session_state.report_data = pd.DataFrame(report_results)
            st.success("Report generated successfully!")
        else:
            st.error("No data found for the selected vessels to generate a report.")

# --- Display Report ---
if st.session_state.report_data is not None:
    st.subheader("Performance Report:")
    st.dataframe(st.session_state.report_data, use_container_width=True)
    
    # Add download button for the report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"vessel_performance_report_{timestamp}.xlsx"
    
    excel_report_data = create_excel_download(st.session_state.report_data.to_dict(orient='records'), report_filename)
    
    if excel_report_data:
        st.download_button(
            label="📥 Download Performance Report (Excel)",
            data=excel_report_data,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- Instructions ---
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Setup Requirements:
    
    1.  **AWS Credentials**: You need AWS Access Key ID, Secret Access Key, and optionally a Session Token.
        *   **Access Key ID & Secret Access Key**: These are long-term credentials for an IAM user.
        *   **Session Token**: This is required if your credentials are *temporary* (e.g., obtained from AWS STS `AssumeRole` or from an EC2 instance role). If you're using long-term IAM user keys, you typically won't have a session token.
    2.  **Lambda Function Name**: The actual name of your Lambda function (e.g., `my-vessel-query-function`).
    
    ### Step-by-step Guide:
    
    1.  **Configure AWS**: Enter your AWS credentials (Access Key ID, Secret Access Key, and Session Token if applicable) and Lambda function name in the sidebar.
    2.  **Vessel List Loads Automatically**: Once AWS config is valid, the vessel list will automatically load.
    3.  **Select Vessels**: Use the multi-select dropdown to choose the vessels you want to include in the report. You can type to search.
    4.  **Generate Report**: Click "Generate Performance Report" to fetch the `hull_roughness_power_loss` data for the selected vessels and display it in a table.
    5.  **Download Report**: An Excel download button will appear below the report table.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈*")
