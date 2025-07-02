import streamlit as st
import boto3
import pandas as pd
import json
import io
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Vessel Data Export",
    page_icon="🚢",
    layout="wide"
)

# Initialize session state
if 'vessels' not in st.session_state:
    st.session_state.vessels = []
if 'selected_vessels' not in st.session_state:
    st.session_state.selected_vessels = []

def invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token=None, region='ap-south-1'):
    """Invoke Lambda function directly using AWS SDK"""
    try:
        # Create Lambda client
        lambda_client = boto3.client(
            'lambda',
            aws_access_key_id=aws_access_key,
            aws_secret_access_key=aws_secret_key,
            aws_session_token=aws_session_token, # Pass session token if provided
            region_name=region
        )
        
        # Invoke the function
        response = lambda_client.invoke(
            FunctionName=function_name,
            InvocationType='RequestResponse',
            Payload=json.dumps(payload)
        )
        
        # Parse response
        response_payload = json.loads(response['Payload'].read())
        
        # Debug: Show raw Lambda response payload
        st.write("**Debug - Raw Lambda Response Payload:**")
        st.json(response_payload)

        if response.get('StatusCode') == 200:
            # Your Lambda returns a JSON object with statusCode and body if it's a Function URL response
            # or just the data if it's a direct invoke.
            # We need to handle both cases.
            if 'statusCode' in response_payload and 'body' in response_payload:
                if response_payload['statusCode'] != 200:
                    st.error(f"Lambda returned error status: {response_payload.get('body', 'Unknown error')}")
                    return None
                # If body is a string, parse it
                if isinstance(response_payload['body'], str):
                    return json.loads(response_payload['body'])
                else:
                    return response_payload['body']
            else:
                # Assume it's a direct invoke response (just the data)
                return response_payload
        else:
            st.error(f"AWS invoke error (non-200 status from AWS API): {response_payload}")
            return None
            
    except Exception as e:
        st.error(f"Error invoking Lambda: {str(e)}")
        return None

def fetch_vessels(function_name, query, aws_access_key, aws_secret_key, aws_session_token):
    """Fetch vessel list from Lambda function using direct invoke"""
    payload = {"sql_query": query}
    
    st.write("**Debug - Payload being sent to Lambda:**")
    st.json(payload)
    
    result = invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        st.write("**Debug - Processed Response from Lambda:**")
        st.json(result)
        return result
    
    return []

def query_vessel_data(function_name, sql_query_string, aws_access_key, aws_secret_key, aws_session_token):
    """Send SQL query to Lambda function and get results"""
    payload = {"sql_query": sql_query_string}
    
    st.write("**Debug - Export Payload being sent to Lambda:**")
    st.json(payload)
    
    result = invoke_lambda_function(function_name, payload, aws_access_key, aws_secret_key, aws_session_token)
    
    if result:
        st.write("**Debug - Processed Export Response from Lambda:**")
        st.json(result)
        return result
    
    return None

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

# Main app
st.title("🚢 Vessel Data Export Tool")
st.markdown("Select vessels and export their data to Excel using direct Lambda invoke")

# Sidebar for configuration
with st.sidebar:
    st.header("AWS Configuration")
    
    # AWS credentials
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
    
    # Lambda function name
    function_name = st.text_input(
        "Lambda Function Name",
        value="",  # User needs to provide this
        help="The name of your Lambda function (e.g., 'my-vessel-query-function')"
    )
    
    st.info("💡 **Function Name**: Find this in your AWS Lambda console. It's the name, not the full ARN or URL.")
    st.warning("⚠️ **Permissions**: Ensure the provided AWS credentials have `lambda:InvokeFunction` permission on your Lambda.")

    # SQL Query Template
    st.header("SQL Query Configuration")
    base_query = st.text_area(
        "Base SQL Query",
        placeholder="""SELECT * FROM vessel_data
WHERE vessel_name IN ({vessel_names})
AND date >= '2024-01-01'
ORDER BY vessel_name, date""",
        height=100,
        help="Use {vessel_names} placeholder for selected vessels"
    )

# Test section
st.header("🔧 Test Section")
st.markdown("Let's test with your known working query:")

test_query = "SELECT vessel_name FROM hull_performance WHERE hull_roughness_power_loss IS NOT NULL OR hull_roughness_speed_loss IS NOT NULL GROUP BY 1;"

if st.button("Test Direct Lambda Invoke"):
    if not all([aws_access_key, aws_secret_key, function_name]):
        st.error("Please provide AWS Access Key ID, Secret Access Key, and Lambda Function Name.")
    else:
        with st.spinner("Testing direct Lambda invoke..."):
            test_result = fetch_vessels(function_name, test_query, aws_access_key, aws_secret_key, aws_session_token)
            if test_result:
                st.success("✅ Direct invoke successful!")
                st.json(test_result)

st.markdown("---")

# Main content area
col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Load Vessels")

    vessel_name_query = "select vessel_name from vessel_particulars"

    if st.button("Fetch Vessels", disabled=not all([aws_access_key, aws_secret_key, function_name])):
        with st.spinner("Loading vessels..."):
            vessels_data = fetch_vessels(function_name, vessel_name_query, aws_access_key, aws_secret_key, aws_session_token)

            if vessels_data:
                extracted_vessel_names = []
                for item in vessels_data:
                    if isinstance(item, dict) and 'vessel_name' in item:
                        extracted_vessel_names.append(item['vessel_name'])
                    elif isinstance(item, str):
                        extracted_vessel_names.append(item)

                st.session_state.vessels = extracted_vessel_names
                st.success(f"Loaded {len(st.session_state.vessels)} vessels")

    if st.session_state.vessels:
        st.info(f"📊 {len(st.session_state.vessels)} vessels available")

with col2:
    st.header("2. Select Vessels")

    if st.session_state.vessels:
        col2a, col2b = st.columns(2)
        with col2a:
            if st.button("Select All"):
                st.session_state.selected_vessels = st.session_state.vessels.copy()
        with col2b:
            if st.button("Clear All"):
                st.session_state.selected_vessels = []

        st.subheader("Choose Vessels:")

        for i, vessel_name in enumerate(st.session_state.vessels):
            is_selected = vessel_name in st.session_state.selected_vessels
            if st.checkbox(vessel_name, value=is_selected, key=f"vessel_{i}"):
                if vessel_name not in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.append(vessel_name)
            else:
                if vessel_name in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.remove(vessel_name)

        if st.session_state.selected_vessels:
            st.success(f"✅ {len(st.session_state.selected_vessels)} vessels selected")
    else:
        st.warning("No vessels loaded. Please fetch vessels first.")

# Query execution section
st.header("3. Export Data")

if st.session_state.selected_vessels and base_query and all([aws_access_key, aws_secret_key, function_name]):
    col3a, col3b = st.columns([3, 1])

    with col3a:
        vessel_names_list = [f"'{name}'" for name in st.session_state.selected_vessels]
        vessel_names_str = ", ".join(vessel_names_list)
        preview_query = base_query.replace("{vessel_names}", vessel_names_str)

        with st.expander("Preview SQL Query"):
            st.code(preview_query, language="sql")

    with col3b:
        export_button = st.button("🚀 Export Data", type="primary")

    if export_button:
        with st.spinner("Querying data..."):
            result_data = query_vessel_data(function_name, preview_query, aws_access_key, aws_secret_key, aws_session_token)

            if result_data:
                st.success("✅ Data retrieved successfully!")

                try:
                    if isinstance(result_data, list) and result_data:
                        preview_df = pd.DataFrame(result_data[:5])
                    elif isinstance(result_data, dict):
                        if 'data' in result_data:
                            preview_df = pd.DataFrame(result_data['data'][:5])
                        else:
                            preview_df = pd.DataFrame([result_data])

                    st.subheader("Data Preview:")
                    st.dataframe(preview_df)

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"vessel_data_{timestamp}.xlsx"

                    excel_data = create_excel_download(result_data, filename)

                    if excel_data:
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"Error processing data: {str(e)}")
                    st.json(result_data)

else:
    missing_items = []
    if not st.session_state.selected_vessels:
        missing_items.append("Select vessels")
    if not base_query:
        missing_items.append("Configure SQL query")
    if not aws_access_key:
        missing_items.append("AWS Access Key")
    if not aws_secret_key:
        missing_items.append("AWS Secret Key")
    if not function_name:
        missing_items.append("Lambda Function Name")

    if missing_items:
        st.warning(f"Please complete: {', '.join(missing_items)}")

# Instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Setup Requirements:
    
    1. **AWS Credentials**: You need AWS Access Key ID, Secret Access Key, and optionally a Session Token.
       - **Access Key ID & Secret Access Key**: These are long-term credentials for an IAM user.
       - **Session Token**: This is required if your credentials are *temporary* (e.g., obtained from AWS STS `AssumeRole` or from an EC2 instance role). If you're using long-term IAM user keys, you typically won't have a session token.
    2. **Lambda Function Name**: The actual name of your Lambda function (e.g., `my-vessel-query-function`).
    
    ### Finding Your Lambda Function Name:
    
    1. Go to AWS Lambda Console.
    2. Find your function in the list.
    3. The function name is displayed (e.g., "my-vessel-query-function").
    4. **Don't use** the full ARN or Function URL - just the name.
    
    ### Step-by-step Guide:
    
    1. **Configure AWS**: Enter your AWS credentials (Access Key ID, Secret Access Key, and Session Token if applicable) and Lambda function name in the sidebar.
    2. **Test**: Use the "Test Direct Lambda Invoke" button to verify the connection.
    3. **Set SQL Query**: Configure your base SQL query using `{vessel_names}` placeholder.
    4. **Load Vessels**: Click "Fetch Vessels" to get the list.
    5. **Select Vessels**: Use checkboxes to select vessels.
    6. **Export**: Click "Export Data" to download Excel file.
    """)

# Footer
st.markdown("---")
st.markdown("*Using AWS SDK for direct Lambda invoke with full credential support*")
