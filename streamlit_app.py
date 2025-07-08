import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.colors import Color
import time

# Page configuration
st.set_page_config(
    page_title="Vessel Performance Report Tool",
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

if 'cache_stats' not in st.session_state:
    st.session_state.cache_stats = {}

# Enhanced Lambda Service Class
class EnhancedLambdaService:
    def __init__(self, lambda_url):
        self.lambda_url = lambda_url
        self.request_count = 0
        self.cache_hits = 0
        
    def call_lambda(self, payload, timeout=30):
        """Enhanced Lambda call with better error handling and performance tracking."""
        try:
            self.request_count += 1
            start_time = time.time()
            
            headers = {'Content-Type': 'application/json'}
            json_payload = json.dumps(payload)
            
            response = requests.post(
                self.lambda_url, 
                headers=headers, 
                data=json_payload,
                timeout=timeout
            )
            
            response_time = time.time() - start_time
            
            if response.status_code != 200:
                st.error(f"HTTP error: {response.status_code} {response.reason}")
                return None
                
            result = response.json()
            
            # Track cache usage
            if isinstance(result, dict) and result.get('cached', False):
                self.cache_hits += 1
                st.info(f"✅ Data retrieved from cache in {response_time:.2f}s")
            else:
                st.success(f"📊 Data retrieved in {response_time:.2f}s")
                
            return result
            
        except requests.exceptions.Timeout:
            st.error("Request timed out. Please try again or select fewer vessels.")
            return None
        except requests.exceptions.ConnectionError:
            st.error("Connection error. Please check your internet connection.")
            return None
        except requests.exceptions.RequestException as e:
            st.error(f"Request error: {str(e)}")
            return None
        except Exception as e:
            st.error(f"Unexpected error: {str(e)}")
            return None

    def get_vessels_list(self, search_term=None, use_cache=True):
        """Get vessel list using enhanced Lambda with server-side filtering."""
        payload = {
            "query_type": "vessel_list",
            "params": [
                f"%{search_term}%" if search_term else None,
                1200  # limit
            ],
            "filters": {
                "pageSize": 1200
            },
            "use_cache": use_cache
        }
        
        result = self.call_lambda(payload)
        if result and 'data' in result:
            return [item['vessel_name'] for item in result['data'] if 'vessel_name' in item]
        return []

    def get_vessel_performance_data(self, vessel_names, use_cache=True):
        """Get comprehensive vessel performance data using predefined queries."""
        if not vessel_names:
            return {}

        # Calculate date ranges
        today = datetime.now()
        first_day_current_month = today.replace(day=1)
        
        # Previous month
        last_day_prev_month = first_day_current_month - timedelta(days=1)
        prev_month_str = last_day_prev_month.strftime("%Y-%m-%d")
        
        # Previous-to-previous month
        first_day_prev_month = last_day_prev_month.replace(day=1)
        last_day_prev_prev_month = first_day_prev_month - timedelta(days=1)
        prev_prev_month_str = last_day_prev_prev_month.strftime("%Y-%m-%d")
        
        # Previous-to-previous-to-previous month
        first_day_prev_prev_month = last_day_prev_prev_month.replace(day=1)
        last_day_prev_prev_prev_month = first_day_prev_prev_month - timedelta(days=1)
        prev_prev_prev_month_str = last_day_prev_prev_prev_month.strftime("%Y-%m-%d")

        all_data = {}
        batch_size = 10
        total_batches = (len(vessel_names) + batch_size - 1) // batch_size
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i in range(0, len(vessel_names), batch_size):
            batch_vessels = vessel_names[i:i+batch_size]
            batch_num = i//batch_size + 1
            
            status_text.text(f"Processing batch {batch_num} of {total_batches} ({len(batch_vessels)} vessels)")
            progress_bar.progress(batch_num / total_batches)
            
            # Process each data type for this batch
            batch_data = self._get_batch_data(batch_vessels, prev_month_str, prev_prev_month_str, 
                                            prev_prev_prev_month_str, last_day_prev_month, 
                                            last_day_prev_prev_month, last_day_prev_prev_prev_month, use_cache)
            
            # Merge batch data into all_data
            for key, data in batch_data.items():
                if key not in all_data:
                    all_data[key] = []
                all_data[key].extend(data)
        
        progress_bar.empty()
        status_text.empty()
        
        return all_data

    def _get_batch_data(self, vessel_names, prev_month_str, prev_prev_month_str, 
                       prev_prev_prev_month_str, last_day_prev_month, 
                       last_day_prev_prev_month, last_day_prev_prev_prev_month, use_cache):
        """Get all performance data for a batch of vessels using enhanced Lambda queries."""
        
        quoted_vessel_names = [f"'{name}'" for name in vessel_names]
        vessel_names_list_str = ", ".join(quoted_vessel_names)
        
        batch_data = {}
        
        # 1. Hull Performance Data (Previous Month)
        hull_prev_query = f"""
        SELECT vessel_name, hull_rough_power_loss_pct_ed
        FROM (
            SELECT vessel_name, hull_rough_power_loss_pct_ed,
                   ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
            FROM hull_performance_six_months_daily
            WHERE vessel_name IN ({vessel_names_list_str})
            AND CAST(updated_ts AS DATE) = '{prev_month_str}'
        ) AS subquery
        WHERE rn = 1
        """
        
        result = self.call_lambda({"sql_query": hull_prev_query, "use_cache": use_cache})
        batch_data['hull_prev'] = result if result else []
        
        # 2. Hull Performance Data (Previous-to-Previous Month)
        hull_prev_prev_query = f"""
        SELECT vessel_name, hull_rough_power_loss_pct_ed
        FROM (
            SELECT vessel_name, hull_rough_power_loss_pct_ed,
                   ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
            FROM hull_performance_six_months_daily
            WHERE vessel_name IN ({vessel_names_list_str})
            AND CAST(updated_ts AS DATE) = '{prev_prev_month_str}'
        ) AS subquery
        WHERE rn = 1
        """
        
        result = self.call_lambda({"sql_query": hull_prev_prev_query, "use_cache": use_cache})
        batch_data['hull_prev_prev'] = result if result else []
        
        # 3. Hull Performance Data (Previous-to-Previous-to-Previous Month)
        hull_prev_prev_prev_query = f"""
        SELECT vessel_name, hull_rough_power_loss_pct_ed
        FROM (
            SELECT vessel_name, hull_rough_power_loss_pct_ed,
                   ROW_NUMBER() OVER (PARTITION BY vessel_name, CAST(updated_ts AS DATE) ORDER BY updated_ts DESC) as rn
            FROM hull_performance_six_months_daily
            WHERE vessel_name IN ({vessel_names_list_str})
            AND CAST(updated_ts AS DATE) = '{prev_prev_prev_month_str}'
        ) AS subquery
        WHERE rn = 1
        """
        
        result = self.call_lambda({"sql_query": hull_prev_prev_prev_query, "use_cache": use_cache})
        batch_data['hull_prev_prev_prev'] = result if result else []
        
        # 4. ME SFOC Data (Previous Month)
        me_prev_query = f"""
        SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
        FROM vessel_performance_summary vps
        JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
        WHERE vp.vessel_name IN ({vessel_names_list_str})
        AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
        AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE)
        GROUP BY vp.vessel_name
        """
        
        result = self.call_lambda({"sql_query": me_prev_query, "use_cache": use_cache})
        batch_data['me_prev'] = result if result else []
        
        # 5. ME SFOC Data (Previous-to-Previous Month)
        me_prev_prev_query = f"""
        SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
        FROM vessel_performance_summary vps
        JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
        WHERE vp.vessel_name IN ({vessel_names_list_str})
        AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
        AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 month')
        GROUP BY vp.vessel_name
        """
        
        result = self.call_lambda({"sql_query": me_prev_prev_query, "use_cache": use_cache})
        batch_data['me_prev_prev'] = result if result else []
        
        # 6. ME SFOC Data (Previous-to-Previous-to-Previous Month)
        me_prev_prev_prev_query = f"""
        SELECT vp.vessel_name, AVG(vps.me_sfoc) AS avg_me_sfoc
        FROM vessel_performance_summary vps
        JOIN vessel_particulars vp ON CAST(vps.vessel_imo AS TEXT) = CAST(vp.vessel_imo AS TEXT)
        WHERE vp.vessel_name IN ({vessel_names_list_str})
        AND vps.reportdate >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '3 months')
        AND vps.reportdate < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '2 months')
        GROUP BY vp.vessel_name
        """
        
        result = self.call_lambda({"sql_query": me_prev_prev_prev_query, "use_cache": use_cache})
        batch_data['me_prev_prev_prev'] = result if result else []
        
        # 7. Fuel Saving Data
        fuel_saving_query = f"""
        SELECT vessel_name, hull_rough_excess_consumption_mt_ed 
        FROM hull_performance_six_months 
        WHERE vessel_name IN ({vessel_names_list_str})
        """
        
        result = self.call_lambda({"sql_query": fuel_saving_query, "use_cache": use_cache})
        batch_data['fuel_saving'] = result if result else []
        
        # 8. CII Data
        cii_query = f"""
        SELECT vp.vessel_name, cy.cii_rating
        FROM vessel_particulars vp
        JOIN cii_ytd cy ON CAST(vp.vessel_imo AS TEXT) = CAST(cy.vessel_imo AS TEXT)
        WHERE vp.vessel_name IN ({vessel_names_list_str})
        """
        
        result = self.call_lambda({"sql_query": cii_query, "use_cache": use_cache})
        batch_data['cii'] = result if result else []
        
        return batch_data

    def get_performance_stats(self):
        """Get performance statistics for the Lambda service."""
        cache_rate = (self.cache_hits / self.request_count * 100) if self.request_count > 0 else 0
        return {
            "total_requests": self.request_count,
            "cache_hits": self.cache_hits,
            "cache_rate": cache_rate
        }

# Data Processing Functions
def process_vessel_performance_data(all_data, vessel_names):
    """Process the raw data into a formatted DataFrame for the report."""
    if not all_data:
        return pd.DataFrame()

    # Calculate date labels
    today = datetime.now()
    first_day_current_month = today.replace(day=1)
    
    last_day_prev_month = first_day_current_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    last_day_prev_prev_month = first_day_prev_month - timedelta(days=1)
    first_day_prev_prev_month = last_day_prev_prev_month.replace(day=1)
    last_day_prev_prev_prev_month = first_day_prev_prev_month - timedelta(days=1)
    
    # Create month labels
    prev_month_hull_col = f"Hull Condition {last_day_prev_month.strftime('%b %y')}"
    prev_prev_month_hull_col = f"Hull Condition {last_day_prev_prev_month.strftime('%b %y')}"
    prev_prev_prev_month_hull_col = f"Hull Condition {last_day_prev_prev_prev_month.strftime('%b %y')}"
    
    prev_month_me_col = f"ME Efficiency {last_day_prev_month.strftime('%b %y')}"
    prev_prev_month_me_col = f"ME Efficiency {last_day_prev_prev_month.strftime('%b %y')}"
    prev_prev_prev_month_me_col = f"ME Efficiency {last_day_prev_prev_prev_month.strftime('%b %y')}"

    # Create base DataFrame
    df_final = pd.DataFrame({'Vessel Name': vessel_names})
    
    # Process hull data
    def process_hull_data(data, col_name):
        if data:
            df_hull = pd.DataFrame(data)
            if 'hull_rough_power_loss_pct_ed' in df_hull.columns:
                df_hull[col_name] = df_hull['hull_rough_power_loss_pct_ed'].apply(get_hull_condition)
            else:
                df_hull[col_name] = "N/A"
            return df_hull[['vessel_name', col_name]].rename(columns={'vessel_name': 'Vessel Name'})
        else:
            return pd.DataFrame({'Vessel Name': [], col_name: []})
    
    # Process ME data
    def process_me_data(data, col_name):
        if data:
            df_me = pd.DataFrame(data)
            if 'avg_me_sfoc' in df_me.columns:
                df_me[col_name] = df_me['avg_me_sfoc'].apply(get_me_efficiency)
            else:
                df_me[col_name] = "N/A"
            return df_me[['vessel_name', col_name]].rename(columns={'vessel_name': 'Vessel Name'})
        else:
            return pd.DataFrame({'Vessel Name': [], col_name: []})
    
    # Merge all data
    data_frames = [
        process_hull_data(all_data.get('hull_prev', []), prev_month_hull_col),
        process_hull_data(all_data.get('hull_prev_prev', []), prev_prev_month_hull_col),
        process_hull_data(all_data.get('hull_prev_prev_prev', []), prev_prev_prev_month_hull_col),
        process_me_data(all_data.get('me_prev', []), prev_month_me_col),
        process_me_data(all_data.get('me_prev_prev', []), prev_prev_month_me_col),
        process_me_data(all_data.get('me_prev_prev_prev', []), prev_prev_prev_month_me_col),
    ]
    
    # Merge DataFrames
    for df in data_frames:
        if not df.empty:
            df_final = pd.merge(df_final, df, on='Vessel Name', how='left')
    
    # Process fuel saving data
    if all_data.get('fuel_saving'):
        df_fuel = pd.DataFrame(all_data['fuel_saving'])
        if 'hull_rough_excess_consumption_mt_ed' in df_fuel.columns:
            df_fuel['Potential Fuel Saving'] = df_fuel['hull_rough_excess_consumption_mt_ed'].apply(
                lambda x: 4.9 if pd.notna(x) and x > 5 else (0.0 if pd.notna(x) and x < 0 else x)
            )
        else:
            df_fuel['Potential Fuel Saving'] = pd.NA
        df_fuel = df_fuel.rename(columns={'vessel_name': 'Vessel Name'})
        df_final = pd.merge(df_final, df_fuel[['Vessel Name', 'Potential Fuel Saving']], on='Vessel Name', how='left')
    
    # Process CII data
    if all_data.get('cii'):
        df_cii = pd.DataFrame(all_data['cii'])
        if 'cii_rating' in df_cii.columns:
            df_cii = df_cii.rename(columns={'cii_rating': 'YTD CII'})
        else:
            df_cii['YTD CII'] = pd.NA
        df_cii = df_cii.rename(columns={'vessel_name': 'Vessel Name'})
        df_final = pd.merge(df_final, df_cii[['Vessel Name', 'YTD CII']], on='Vessel Name', how='left')
    
    # Add S. No. and Comments
    df_final.insert(0, 'S. No.', range(1, len(df_final) + 1))
    df_final['Comments'] = ""
    
    # Fill NaN values with "N/A"
    df_final = df_final.fillna("N/A")
    
    return df_final

def get_hull_condition(value):
    """Determine hull condition based on power loss percentage."""
    if pd.isna(value):
        return "N/A"
    if value < 15:
        return "Good"
    elif 15 <= value <= 25:
        return "Average"
    else:
        return "Poor"

def get_me_efficiency(value):
    """Determine ME efficiency based on SFOC value."""
    if pd.isna(value):
        return "N/A"
    if value < 160:
        return "Anomalous data"
    elif value < 180:
        return "Good"
    elif 180 <= value <= 190:
        return "Average"
    else:
        return "Poor"

# Cached vessel loading function
@st.cache_data(ttl=3600)
def load_vessels_cached(lambda_url, search_term=None):
    """Cached function to load vessels."""
    # Create a temporary service instance for caching
    temp_service = EnhancedLambdaService(lambda_url)
    return temp_service.get_vessels_list(search_term, use_cache=True)

# Styling Functions
def style_condition_columns(row):
    """Apply styling to condition columns."""
    styles = [''] * len(row)
    
    # Style hull condition columns
    hull_condition_cols = [col for col in row.index if 'Hull Condition' in col]
    for col_name in hull_condition_cols:
        if col_name in row.index:
            hull_val = row[col_name]
            if hull_val == "Good":
                styles[row.index.get_loc(col_name)] = 'background-color: #d4edda; color: black;'
            elif hull_val == "Average":
                styles[row.index.get_loc(col_name)] = 'background-color: #fff3cd; color: black;'
            elif hull_val == "Poor":
                styles[row.index.get_loc(col_name)] = 'background-color: #f8d7da; color: black;'
    
    # Style ME efficiency columns
    me_efficiency_cols = [col for col in row.index if 'ME Efficiency' in col]
    for col_name in me_efficiency_cols:
        if col_name in row.index:
            me_val = row[col_name]
            if me_val == "Good":
                styles[row.index.get_loc(col_name)] = 'background-color: #d4edda; color: black;'
            elif me_val == "Average":
                styles[row.index.get_loc(col_name)] = 'background-color: #fff3cd; color: black;'
            elif me_val == "Poor":
                styles[row.index.get_loc(col_name)] = 'background-color: #f8d7da; color: black;'
            elif me_val == "Anomalous data":
                styles[row.index.get_loc(col_name)] = 'background-color: #e0e0e0; color: black;'
                
    return styles

def create_excel_download_with_styling(df, filename):
    """Create Excel file with styling."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Vessel Report"

    # Write headers
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name).font = Font(bold=True)

    # Write data and apply styling
    for row_idx, row_data in df.iterrows():
        for col_idx, (col_name, cell_value) in enumerate(row_data.items(), 1):
            cell = ws.cell(row=row_idx + 2, column=col_idx, value=cell_value)
            
            if 'Hull Condition' in col_name or 'ME Efficiency' in col_name:
                if cell_value == "Good":
                    cell.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                elif cell_value == "Average":
                    cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                elif cell_value == "Poor":
                    cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                elif cell_value == "Anomalous data":
                    cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
                cell.font = Font(color="000000")
            elif col_name == 'YTD CII':
                cell.alignment = Alignment(horizontal='center')

    # Auto-adjust column widths
    for col_idx, column in enumerate(df.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for cell in ws[column_letter]:
            try:
                if cell.value is not None and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output)
    return output.getvalue()

# Main Application
def main():
    # Lambda Function URL
    LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/"
    
    # Initialize Lambda service
    if 'lambda_service' not in st.session_state:
        st.session_state.lambda_service = EnhancedLambdaService(LAMBDA_FUNCTION_URL)
    
    lambda_service = st.session_state.lambda_service
    
    # Title and header
    st.title("🚢 Enhanced Vessel Performance Report Tool")
    st.markdown("Select vessels and generate a comprehensive performance report with improved caching and server-side processing.")
    
    # Performance metrics sidebar
    with st.sidebar:
        st.header("📊 Performance Metrics")
        stats = lambda_service.get_performance_stats()
        st.metric("Total Requests", stats["total_requests"])
        st.metric("Cache Hits", stats["cache_hits"])
        st.metric("Cache Rate", f"{stats['cache_rate']:.1f}%")
        
        if st.button("🔄 Clear Cache"):
            st.cache_data.clear()
            st.success("Cache cleared!")
    
    # Load vessels
    st.header("1. Select Vessels")
    
    # Search functionality
    search_query = st.text_input(
        "Search vessels:",
        value=st.session_state.search_query,
        placeholder="Type to filter vessel names...",
        help="Type to filter the list of vessels below."
    )
    
    if search_query != st.session_state.search_query:
        st.session_state.search_query = search_query
    
    # Load vessels with search
    with st.spinner("Loading vessels..."):
        vessels = load_vessels_cached(LAMBDA_FUNCTION_URL, search_query if search_query else None)
    
    if vessels:
        st.markdown(f"📊 {len(vessels)} vessels available. {len(st.session_state.selected_vessels)} selected.")
        
        # Vessel selection with improved UI
        with st.container(height=300, border=True):
            if vessels:
                cols = st.columns(3)
                for i, vessel in enumerate(vessels):
                    col_idx = i % 3
                    checkbox_state = cols[col_idx].checkbox(
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
        
        # Batch selection controls
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Select All"):
                st.session_state.selected_vessels = set(vessels)
                st.rerun()
        with col2:
            if st.button("Clear Selection"):
                st.session_state.selected_vessels = set()
                st.rerun()
        with col3:
            if st.button("Select First 10"):
                st.session_state.selected_vessels = set(vessels[:10])
                st.rerun()
        
        selected_vessels_list = list(st.session_state.selected_vessels)
    else:
        st.error("Failed to load vessels. Please check your connection and try again.")
        selected_vessels_list = []
    
    # Generate report section
    st.header("2. Generate Report")
    
    # Report options
    col1, col2 = st.columns(2)
    with col1:
        use_cache = st.checkbox("Use caching for faster results", value=True, 
                               help="Enable caching to speed up repeated queries")
    with col2:
        force_refresh = st.checkbox("Force refresh data", value=False,
                                   help="Bypass cache and fetch fresh data")
    
    if selected_vessels_list:
        if st.button("🚀 Generate Enhanced Performance Report", type="primary"):
            with st.spinner("Generating enhanced report with server-side processing..."):
                try:
                    # Get performance data using enhanced Lambda
                    all_data = lambda_service.get_vessel_performance_data(
                        selected_vessels_list, 
                        use_cache=(use_cache and not force_refresh)
                    )
                    
                    # Process data into final report format
                    st.session_state.report_data = process_vessel_performance_data(
                        all_data, selected_vessels_list
                    )
                    
                    if not st.session_state.report_data.empty:
                        st.success("✅ Report generated successfully with enhanced processing!")
                    else:
                        st.warning("⚠️ No data found for the selected vessels.")
                        
                except Exception as e:
                    st.error(f"❌ Error generating report: {str(e)}")
                    st.session_state.report_data = None
    else:
        st.warning("Please select at least one vessel to generate a report.")
    
    # Display report results
    if st.session_state.report_data is not None and not st.session_state.report_data.empty:
        st.header("3. Enhanced Report Results")
        
        # Report summary
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Vessels", len(st.session_state.report_data))
        with col2:
            # Count vessels with good hull condition (latest month)
            latest_hull_col = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
            if latest_hull_col:
                good_hulls = len(st.session_state.report_data[
                    st.session_state.report_data[latest_hull_col[0]] == "Good"
                ])
            else:
                good_hulls = 0
            st.metric("Good Hull Condition", good_hulls)
        with col3:
            # Count vessels with good ME efficiency (latest month)
            latest_me_col = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]
            if latest_me_col:
                good_me = len(st.session_state.report_data[
                    st.session_state.report_data[latest_me_col[0]] == "Good"
                ])
            else:
                good_me = 0
            st.metric("Good ME Efficiency", good_me)
        with col4:
            # Average potential fuel saving
            if 'Potential Fuel Saving' in st.session_state.report_data.columns:
                avg_fuel_saving = st.session_state.report_data['Potential Fuel Saving'].apply(
                    lambda x: float(x) if pd.notna(x) and str(x) != 'N/A' else 0
                ).mean()
                st.metric("Avg Fuel Saving (MT/day)", f"{avg_fuel_saving:.2f}")
            else:
                st.metric("Avg Fuel Saving (MT/day)", "N/A")
        
        # Display styled dataframe
        styled_df = st.session_state.report_data.style.apply(
            style_condition_columns, axis=1
        )
        st.dataframe(styled_df, use_container_width=True)
        
        # Download section
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"enhanced_vessel_performance_report_{timestamp}.xlsx"
        
        try:
            excel_data = create_excel_download_with_styling(st.session_state.report_data, filename)
            if excel_data:
                st.download_button(
                    label="📥 Download Enhanced Report as Excel",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error creating Excel file: {str(e)}")
            
        # Data insights section
        with st.expander("📈 Data Insights"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Hull Condition Summary")
                if latest_hull_col:
                    hull_summary = st.session_state.report_data[latest_hull_col[0]].value_counts()
                    st.bar_chart(hull_summary)
                else:
                    st.info("No hull condition data available")
            
            with col2:
                st.subheader("ME Efficiency Summary")
                if latest_me_col:
                    me_summary = st.session_state.report_data[latest_me_col[0]].value_counts()
                    st.bar_chart(me_summary)
                else:
                    st.info("No ME efficiency data available")
        
        # Trend analysis
        with st.expander("📊 Trend Analysis"):
            st.subheader("Hull Condition Trends")
            hull_cols = [col for col in st.session_state.report_data.columns if 'Hull Condition' in col]
            if len(hull_cols) >= 2:
                trend_data = []
                for col in hull_cols:
                    good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                    average_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Average"])
                    poor_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Poor"])
                    
                    month = col.replace("Hull Condition ", "")
                    trend_data.append({
                        "Month": month,
                        "Good": good_count,
                        "Average": average_count,
                        "Poor": poor_count
                    })
                
                trend_df = pd.DataFrame(trend_data)
                st.line_chart(trend_df.set_index("Month"))
            else:
                st.info("Need at least 2 months of data for trend analysis")
            
            st.subheader("ME Efficiency Trends")
            me_cols = [col for col in st.session_state.report_data.columns if 'ME Efficiency' in col]
            if len(me_cols) >= 2:
                me_trend_data = []
                for col in me_cols:
                    good_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Good"])
                    average_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Average"])
                    poor_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Poor"])
                    anomalous_count = len(st.session_state.report_data[st.session_state.report_data[col] == "Anomalous data"])
                    
                    month = col.replace("ME Efficiency ", "")
                    me_trend_data.append({
                        "Month": month,
                        "Good": good_count,
                        "Average": average_count,
                        "Poor": poor_count,
                        "Anomalous": anomalous_count
                    })
                
                me_trend_df = pd.DataFrame(me_trend_data)
                st.line_chart(me_trend_df.set_index("Month"))
            else:
                st.info("Need at least 2 months of data for trend analysis")

    elif st.session_state.report_data is not None and st.session_state.report_data.empty:
        st.info("No data found for the selected vessels.")
    
    # Enhanced instructions
    with st.expander("📖 Enhanced Features & Instructions"):
        st.markdown("""
        ### 🚀 New Enhanced Features:
        
        **Performance Improvements:**
        - ⚡ **Server-side Processing**: Data filtering and processing now happens in Lambda for faster results
        - 🗄️ **Intelligent Caching**: Frequently accessed data is cached for instant retrieval
        - 📊 **Batch Processing**: Large vessel selections are processed in optimized batches
        - 📈 **Performance Metrics**: Track cache hits and request performance in the sidebar
        
        **Enhanced UI Features:**
        - 🔍 **Smart Search**: Real-time vessel filtering with server-side search
        - 📋 **Batch Selection**: Select all, clear all, or select first 10 vessels with one click
        - 📊 **Data Insights**: Automatic charts and trend analysis for hull and ME efficiency
        - 📈 **Trend Analysis**: Compare performance across multiple months
        - 📋 **Report Summary**: Key metrics displayed at the top of results
        
        **Caching Options:**
        - ✅ **Use Caching**: Enable for faster repeated queries (recommended)
        - 🔄 **Force Refresh**: Bypass cache to get the latest data
        - 🗑️ **Clear Cache**: Reset all cached data (available in sidebar)
        
        ### 📋 How to Use:
        
        1. **Search & Select**: Use the search bar to filter vessels, then select using checkboxes or batch controls
        2. **Configure Options**: Choose caching preferences based on your needs
        3. **Generate Report**: Click the enhanced generate button for server-optimized processing
        4. **Analyze Results**: Review summary metrics, charts, and trend analysis
        5. **Download**: Export the styled report to Excel
        
        ### 📊 Report Columns (Enhanced):
        
        **Hull Condition** (Multiple months for trend analysis):
        - 🟢 **Good**: < 15% power loss
        - 🟡 **Average**: 15-25% power loss  
        - 🔴 **Poor**: > 25% power loss
        
        **ME Efficiency** (Multiple months for trend analysis):
        - ⚪ **Anomalous data**: < 160 SFOC
        - 🟢 **Good**: 160-180 SFOC
        - 🟡 **Average**: 180-190 SFOC
        - 🔴 **Poor**: > 190 SFOC
        
        **Additional Metrics:**
        - 🔋 **Potential Fuel Saving**: Excess consumption (MT/day) - automatically capped at 4.9
        - 📊 **YTD CII**: Carbon Intensity Indicator rating
        - 💬 **Comments**: Space for additional notes
        
        ### 🔧 Performance Tips:
        
        - Enable caching for repeated analysis of the same vessels
        - Use search to narrow down vessel lists before selection
        - Process large vessel selections in smaller batches if needed
        - Use force refresh only when you need the latest data
        - Check performance metrics in sidebar to monitor system efficiency
        """)
    
    # Footer with enhanced info
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("*Enhanced with server-side processing*")
    with col2:
        st.markdown("*Built with Streamlit 🎈 and Python*")
    with col3:
        if st.session_state.lambda_service:
            stats = st.session_state.lambda_service.get_performance_stats()
            st.markdown(f"*{stats['total_requests']} requests • {stats['cache_rate']:.0f}% cached*")

if __name__ == "__main__":
    main()
