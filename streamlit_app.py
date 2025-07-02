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

def fetch_vessels(api_url):
    """Fetch vessel list from Lambda API"""
    try:
        response = requests.get(api_url)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching vessels: {str(e)}")
        return []

def query_vessel_data(api_url, query_payload):
    """Send SQL query to Lambda API and get results"""
    try:
        headers = {'Content-Type': 'application/json'}
        response = requests.post(api_url, 
                               data=json.dumps(query_payload), 
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
    vessel_api_url = st.text_input(
        "Vessel List API URL",
        placeholder="https://your-lambda-url.com/vessels",
        help="Lambda API endpoint that returns list of vessels"
    )
    
    query_api_url = st.text_input(
        "Data Query API URL", 
        placeholder="https://your-lambda-url.com/query",
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
    
    if st.button("Fetch Vessels", disabled=not vessel_api_url):
        with st.spinner("Loading vessels..."):
            vessels_data = fetch_vessels(vessel_api_url)
            
            if vessels_data:
                # Handle different response formats
                if isinstance(vessels_data, list):
                    st.session_state.vessels = vessels_data
                elif isinstance(vessels_data, dict):
                    # Try common keys
                    for key in ['vessels', 'data', 'results', 'items']:
                        if key in vessels_data:
                            st.session_state.vessels = vessels_data[key]
                            break
                    else:
                        st.session_state.vessels = [vessels_data]
                
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
        for i, vessel in enumerate(st.session_state.vessels):
            # Extract vessel name/identifier
            if isinstance(vessel, dict):
                vessel_name = vessel.get('name', vessel.get('vessel_name', 
                                       vessel.get('id', f"Vessel_{i}")))
                display_name = f"{vessel_name}"
                if 'type' in vessel:
                    display_name += f" ({vessel['type']})"
            else:
                vessel_name = str(vessel)
                display_name = vessel_name
            
            # Checkbox for vessel selection
            is_selected = vessel in st.session_state.selected_vessels
            if st.checkbox(display_name, value=is_selected, key=f"vessel_{i}"):
                if vessel not in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.append(vessel)
            else:
                if vessel in st.session_state.selected_vessels:
                    st.session_state.selected_vessels.remove(vessel)
        
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
        vessel_names_list = []
        for vessel in st.session_state.selected_vessels:
            if isinstance(vessel, dict):
                name = vessel.get('name', vessel.get('vessel_name', vessel.get('id')))
            else:
                name = str(vessel)
            vessel_names_list.append(f"'{name}'")
        
        vessel_names_str = ", ".join(vessel_names_list)
        preview_query = base_query.replace("{vessel_names}", vessel_names_str)
        
        with st.expander("Preview SQL Query"):
            st.code(preview_query, language="sql")
    
    with col3b:
        export_button = st.button("🚀 Export Data", type="primary")
    
    if export_button:
        with st.spinner("Querying data..."):
            # Prepare query payload
            query_payload = {
                "query": preview_query,
                "vessel_names": vessel_names_list,
                "selected_vessels": st.session_state.selected_vessels
            }
            
            # Execute query
            result_data = query_vessel_data(query_api_url, query_payload)
            
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
