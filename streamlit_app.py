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

def fetch_vessels(api_url, query):
    """Fetch vessel list from Lambda API using a POST request with a SQL query."""
    try:
        headers = {'Content-Type': 'application/json'}
        payload = {"sql_query": query} # Send the SQL query in the payload
        response = requests.post(api_url, data=json.dumps(payload), headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching vessels: {str(e)}")
        return []

def query_vessel_data(api_url, sql_query_string): # Changed parameter name for clarity
    """Send SQL query to Lambda API and get results"""
    try:
        headers = {'Content-Type': 'application/json'}
        # The payload should ONLY contain 'sql_query' as per your Lambda
        payload = {"sql_query": sql_query_string}
        response = requests.post(api_url,
                               data=json.dumps(payload), # Send the correct payload
                               headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Error querying data: {str(e)}")
        return None

def create_excel_download(data, filename):
    """Convert data to Excel format for download"""
    try:
        # Convert JSON data to DataFrame
        if isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, dict):
            # Handle different JSON structures
            if 'data' in data:
                df = pd.DataFrame(data['data'])
            elif 'results' in data:
                df = pd.DataFrame(data['results'])
            else:
                df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)

        # Create Excel buffer
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
    # Use the same URL for both, as your Lambda handles queries
    vessel_api_url = st.text_input(
        "Vessel List API URL",
        value="https://6mfmavicpuezjic6mtwtbuw56e0pjysg.lambda-url.ap-south-1.on.aws/", # Pre-fill with your Lambda URL
        help="Lambda API endpoint that returns list of vessels"
    )

    query_api_url = st.text_input(
        "Data Query API URL",
        value="https://6mfmavicpuezjic6mtwtbuw56e0pjysg.lambda-url.ap-south-1.on.aws/", # Pre-fill with your Lambda URL
        help="Lambda API endpoint that executes SQL queries"
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

# Main content area
col1, col2 = st.columns([1, 1])

with col1:
    st.header("1. Load Vessels")

    # Define the specific query for fetching vessel names
    vessel_name_query = "select vessel_name from vessel_particulars"

    if st.button("Fetch Vessels", disabled=not vessel_api_url):
        with st.spinner("Loading vessels..."):
            # Pass the specific query to fetch_vessels
            vessels_data = fetch_vessels(vessel_api_url, vessel_name_query)

            if vessels_data:
                # Your Lambda returns a list of objects like [{"vessel_name": "Vessel1"}]
                # We need to extract just the vessel names for selection
                extracted_vessel_names = []
                for item in vessels_data:
                    if isinstance(item, dict) and 'vessel_name' in item:
                        extracted_vessel_names.append(item['vessel_name'])
                    elif isinstance(item, str): # In case it returns just strings
                        extracted_vessel_names.append(item)

                st.session_state.vessels = extracted_vessel_names
                st.success(f"Loaded {len(st.session_state.vessels)} vessels")

    # Display loaded vessels count
    if st.session_state.vessels:
        st.info(f"📊 {len(st.session_state.vessels)} vessels available")

with col2:
    st.header("2. Select Vessels")

    if st.session_state.vessels:
        # Select all/none buttons
        col2a, col2b = st.columns(2)
        with col2a:
            if st.button("Select All"):
                st.session_state.selected_vessels = st.session_state.vessels.copy()
        with col2b:
            if st.button("Clear All"):
                st.session_state.selected_vessels = []

        # Vessel selection checkboxes
        st.subheader("Choose Vessels:")

        # Handle different vessel data formats
        for i, vessel_name in enumerate(st.session_state.vessels):
            # Checkbox for vessel selection
            is_selected = vessel_name in st.session_state.selected_vessels
            if st.checkbox(vessel_name, value=is_selected, key=f"vessel_{i}"):
                if vessel_name not in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.append(vessel_name)
            else:
                if vessel_name in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.remove(vessel_name)

        # Show selection summary
        if st.session_state.selected_vessels:
            st.success(f"✅ {len(st.session_state.selected_vessels)} vessels selected")
    else:
        st.warning("No vessels loaded. Please fetch vessels first.")

# Query execution section
st.header("3. Export Data")

if st.session_state.selected_vessels and base_query and query_api_url:
    col3a, col3b = st.columns([3, 1])

    with col3a:
        # Show query preview
        vessel_names_list = [f"'{name}'" for name in st.session_state.selected_vessels]

        vessel_names_str = ", ".join(vessel_names_list)
        preview_query = base_query.replace("{vessel_names}", vessel_names_str)

        with st.expander("Preview SQL Query"):
            st.code(preview_query, language="sql")

    with col3b:
        export_button = st.button("🚀 Export Data", type="primary")

    if export_button:
        with st.spinner("Querying data..."):
            # Prepare query payload for the main data query
            # Pass only the preview_query string to query_vessel_data
            result_data = query_vessel_data(query_api_url, preview_query)

            if result_data:
                st.success("✅ Data retrieved successfully!")

                # Show data preview
                try:
                    if isinstance(result_data, list) and result_data:
                        preview_df = pd.DataFrame(result_data[:5])  # Show first 5 rows
                    elif isinstance(result_data, dict):
                        if 'data' in result_data:
                            preview_df = pd.DataFrame(result_data['data'][:5])
                        else:
                            preview_df = pd.DataFrame([result_data])

                    st.subheader("Data Preview:")
                    st.dataframe(preview_df)

                    # Create download
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
                    # Still show raw data
                    st.json(result_data)

else:
    missing_items = []
    if not st.session_state.selected_vessels:
        missing_items.append("Select vessels")
    if not base_query:
        missing_items.append("Configure SQL query")
    if not query_api_url:
        missing_items.append("Set query API URL")

    if missing_items:
        st.warning(f"Please complete: {', '.join(missing_items)}")

# Instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Step-by-step Guide:

    1. **Configure APIs**: In the sidebar, enter your Lambda API URLs
       - Vessel List API: Should return a JSON array/object with vessel information
       - Data Query API: Should accept JSON payload with SQL query and return data

    2. **Set SQL Query**: Configure your base SQL query using `{vessel_names}` placeholder

    3. **Load Vessels**: Click "Fetch Vessels" to get the list from your API

    4. **Select Vessels**: Use checkboxes to select which vessels you want data for

    5. **Export**: Click "Export Data" to run the query and download Excel file

    ### API Requirements:

    **Vessel List API Response Format:**
    ```json
    [
        {"name": "Vessel1", "type": "Cargo"},
        {"name": "Vessel2", "type": "Tanker"}
    ]
    ```

    **Query API Request Format:**
    ```json
    {
        "query": "SELECT * FROM vessel_data WHERE vessel_name IN ('Vessel1', 'Vessel2')",
        "vessel_names": ["'Vessel1'", "'Vessel2'"],
        "selected_vessels": [...]
    }
    ```

    **Query API Response Format:**
    ```json
    [
        {"vessel_name": "Vessel1", "date": "2024-01-01", "value": 100},
        {"vessel_name": "Vessel2", "date": "2024-01-01", "value": 200}
    ]
    ```
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit 🎈*")
