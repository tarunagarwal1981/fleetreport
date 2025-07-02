import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import datetime
import urllib.parse

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

def fetch_vessels_approach_1(api_url, query):
    """Approach 1: Try with query parameters"""
    try:
        # Try sending as query parameter
        params = {"sql_query": query}
        response = requests.get(api_url, params=params, timeout=30)
        
        st.write("**Approach 1 - GET with query params:**")
        st.write(f"URL: {response.url}")
        st.write(f"Status: {response.status_code}")
        st.write(f"Response: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        return None
    except Exception as e:
        st.write(f"Approach 1 failed: {e}")
        return None

def fetch_vessels_approach_2(api_url, query):
    """Approach 2: Try with form data"""
    try:
        # Try sending as form data
        data = {"sql_query": query}
        response = requests.post(api_url, data=data, timeout=30)
        
        st.write("**Approach 2 - POST with form data:**")
        st.write(f"Status: {response.status_code}")
        st.write(f"Response: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        return None
    except Exception as e:
        st.write(f"Approach 2 failed: {e}")
        return None

def fetch_vessels_approach_3(api_url, query):
    """Approach 3: Try with URL encoded JSON"""
    try:
        # Try sending JSON as URL encoded
        payload = json.dumps({"sql_query": query})
        data = {"json": payload}
        response = requests.post(api_url, data=data, timeout=30)
        
        st.write("**Approach 3 - POST with URL encoded JSON:**")
        st.write(f"Status: {response.status_code}")
        st.write(f"Response: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        return None
    except Exception as e:
        st.write(f"Approach 3 failed: {e}")
        return None

def fetch_vessels_approach_4(api_url, query):
    """Approach 4: Try with raw JSON string as body"""
    try:
        # Try sending raw JSON string
        payload = json.dumps({"sql_query": query})
        headers = {'Content-Type': 'text/plain'}
        response = requests.post(api_url, data=payload, headers=headers, timeout=30)
        
        st.write("**Approach 4 - POST with raw JSON string:**")
        st.write(f"Status: {response.status_code}")
        st.write(f"Response: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        return None
    except Exception as e:
        st.write(f"Approach 4 failed: {e}")
        return None

def fetch_vessels_approach_5(api_url, query):
    """Approach 5: Try with different Content-Type"""
    try:
        # Try with application/x-www-form-urlencoded
        payload = {"sql_query": query}
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = requests.post(api_url, data=payload, headers=headers, timeout=30)
        
        st.write("**Approach 5 - POST with form-urlencoded:**")
        st.write(f"Status: {response.status_code}")
        st.write(f"Response: {response.text}")
        
        if response.status_code == 200:
            return response.json()
        return None
    except Exception as e:
        st.write(f"Approach 5 failed: {e}")
        return None

def fetch_vessels(api_url, query):
    """Try multiple approaches to send the request"""
    st.write("### Trying Multiple Request Approaches:")
    
    # Try approach 1: GET with query params
    result = fetch_vessels_approach_1(api_url, query)
    if result:
        return result
    
    # Try approach 2: POST with form data
    result = fetch_vessels_approach_2(api_url, query)
    if result:
        return result
    
    # Try approach 3: POST with URL encoded JSON
    result = fetch_vessels_approach_3(api_url, query)
    if result:
        return result
    
    # Try approach 4: POST with raw JSON string
    result = fetch_vessels_approach_4(api_url, query)
    if result:
        return result
    
    # Try approach 5: POST with form-urlencoded
    result = fetch_vessels_approach_5(api_url, query)
    if result:
        return result
    
    st.error("All approaches failed!")
    return []

def query_vessel_data(api_url, sql_query_string):
    """Use the same successful approach for data querying"""
    # We'll implement this once we find which approach works
    st.info("Will implement data querying once vessel fetching works")
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
st.title("🚢 Vessel Data Export Tool - Debug Mode")
st.markdown("Testing different request approaches to find what works with your Lambda")

# Sidebar for configuration
with st.sidebar:
    st.header("API Configuration")

    vessel_api_url = st.text_input(
        "Vessel List API URL",
        value="https://6mfmavicpuezjic6mtwtbuw56e0pjysg.lambda-url.ap-south-1.on.aws/",
        help="Lambda API endpoint"
    )

# Test section
st.header("🔧 Multi-Approach Test")
st.markdown("Let's try different ways to send the request:")

test_query = "SELECT vessel_name FROM hull_performance WHERE hull_roughness_power_loss IS NOT NULL OR hull_roughness_speed_loss IS NOT NULL GROUP BY 1;"

if st.button("Test All Approaches"):
    if vessel_api_url:
        with st.spinner("Testing all approaches..."):
            test_result = fetch_vessels(vessel_api_url, test_query)
            if test_result:
                st.success("✅ Found a working approach!")
                st.json(test_result)
            else:
                st.error("❌ None of the approaches worked")
    else:
        st.error("Please enter the API URL first")

st.markdown("---")

# Simple vessel loading test
st.header("1. Load Vessels (Simple Test)")
vessel_name_query = "select vessel_name from vessel_particulars"

if st.button("Test Vessel Loading", disabled=not vessel_api_url):
    with st.spinner("Testing vessel loading..."):
        vessels_data = fetch_vessels(vessel_api_url, vessel_name_query)
        if vessels_data:
            st.success(f"✅ Successfully loaded data!")
            st.json(vessels_data)

# Instructions
with st.expander("📖 What This Does"):
    st.markdown("""
    This debug version tries 5 different approaches to send requests to your Lambda:
    
    1. **GET with query parameters** - `?sql_query=SELECT...`
    2. **POST with form data** - `Content-Type: application/x-www-form-urlencoded`
    3. **POST with URL encoded JSON** - JSON wrapped in form data
    4. **POST with raw JSON string** - `Content-Type: text/plain`
    5. **POST with form-urlencoded** - Explicit form encoding
    
    Once we find which approach works, we can use that for the full application.
    """)

# Footer
st.markdown("---")
st.markdown("*Debug Mode - Finding the right request format*")
