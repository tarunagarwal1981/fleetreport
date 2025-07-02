import streamlit as st
import requests
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

def make_api_request(api_url, query_payload, access_token=None):
    """
    Makes a POST request to the API URL with the given payload and optional access token.
    """
    try:
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        
        if access_token:
            # Assuming it's a Bearer token or similar.
            # If it's an API Gateway API Key, it might be 'x-api-key': access_token
            headers['Authorization'] = f'Bearer {access_token}'
            # Or if it's an API Gateway API Key:
            # headers['x-api-key'] = access_token
            
        st.write("**Debug - Request Headers:**")
        st.json(headers)
        st.write("**Debug - Request Payload:**")
        st.json(query_payload)

        response = requests.post(
            api_url,
            json=query_payload, # Use json parameter for automatic serialization and Content-Type
            headers=headers,
            timeout=30
        )

        st.write(f"**Debug - Response Status Code:** {response.status_code}")
        st.write(f"**Debug - Raw Response Text:** {response.text}")
        
        response.raise_for_status() # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()

    except requests.exceptions.RequestException as e:
        st.error(f"Error making API request: {str(e)}")
        if hasattr(e, 'response') and e.response is not None:
            st.error(f"Response content: {e.response.text}")
        return None

def fetch_vessels(api_url, query, access_token):
    """Fetch vessel list from Lambda API using a POST request with a SQL query."""
    payload = {"sql_query": query}
    return make_api_request(api_url, payload, access_token)

def query_vessel_data(api_url, sql_query_string, access_token):
    """Send SQL query to Lambda API and get results"""
    payload = {"sql_query": sql_query_string}
    return make_api_request(api_url, payload, access_token)

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
st.markdown("Select vessels and export their data to Excel")

# Sidebar for configuration
with st.sidebar:
    st.header("API Configuration")
    
    # API URLs
    vessel_api_url = st.text_input(
        "Vessel List API URL",
        value="https://6mfmavicpuezjic6mtwtbuw56e0pjysg.lambda-url.ap-south-1.on.aws/",
        help="Lambda API endpoint that returns list of vessels"
    )
    
    query_api_url = st.text_input(
        "Data Query API URL",
        value="https://6mfmavicpuezjic6mtwtbuw56e0pjysg.lambda-url.ap-south-1.on.aws/",
        help="Lambda API endpoint that executes SQL queries"
    )
    
    # Access Token input
    access_token = st.text_input(
        "Access Token (Optional)",
        type="password",
        help="If your Lambda Function URL or API Gateway requires an access token (e.g., API Key, Bearer Token)"
    )

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

if st.button("Test API Connection"):
    if not vessel_api_url:
        st.error("Please provide the API URL.")
    else:
        with st.spinner("Testing API connection..."):
            test_result = fetch_vessels(vessel_api_url, test_query, access_token)
            if test_result:
                st.success("✅ API connection successful!")
                st.json(test_result)

st.markdown("---")

# Main content area
col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Load Vessels")

    vessel_name_query = "select vessel_name from vessel_particulars"

    if st.button("Fetch Vessels", disabled=not vessel_api_url):
        with st.spinner("Loading vessels..."):
            vessels_data = fetch_vessels(vessel_api_url, vessel_name_query, access_token)

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

if st.session_state.selected_vessels and base_query and vessel_api_url:
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
            result_data = query_vessel_data(query_api_url, preview_query, access_token)

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
    if not vessel_api_url:
        missing_items.append("Set API URL")

    if missing_items:
        st.warning(f"Please complete: {', '.join(missing_items)}")

# Instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Setup Requirements:
    
    1. **API URLs**: Enter your Lambda Function URLs.
    2. **Access Token (Optional)**: If your Lambda Function URL or API Gateway requires an access token, enter it here.
       - **Common types**:
         - **API Key**: Often sent in `x-api-key` header.
         - **Bearer Token**: Often sent in `Authorization: Bearer YOUR_TOKEN` header.
       - The code currently assumes a `Bearer` token. If it's an `x-api-key`, you'll need to uncomment that line in `make_api_request`.
    
    ### Step-by-step Guide:
    
    1. **Configure APIs**: Enter your Lambda Function URLs in the sidebar.
    2. **Enter Access Token**: If required, paste your access token into the "Access Token" field.
    3. **Test**: Use the "Test API Connection" button to verify the connection.
    4. **Set SQL Query**: Configure your base SQL query using `{vessel_names}` placeholder.
    5. **Load Vessels**: Click "Fetch Vessels" to get the list.
    6. **Select Vessels**: Use checkboxes to select vessels.
    7. **Export**: Click "Export Data" to download Excel file.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈*")
