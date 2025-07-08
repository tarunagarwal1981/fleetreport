import streamlit as st
import pandas as pd
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64
import json
import requests
import datetime

# --- Helper Functions for Word Report (as previously defined) ---

def set_cell_border(cell, **kwargs):
    """
    Set cell border for a given cell.
    Usage: set_cell_border(cell, top={"sz": 12, "val": "single", "color": "#FF0000"}, ...)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Create a dictionary of border properties
    borders = {
        'top': 'w:topBdr',
        'left': 'w:leftBdr',
        'bottom': 'w:bottomBdr',
        'right': 'w:rightBdr',
        'insideH': 'w:insideH',
        'insideV': 'w:insideV'
    }

    for border_name, border_tag in borders.items():
        if border_name in kwargs:
            border_properties = kwargs[border_name]
            if border_properties is not None:
                border_element = OxmlElement(border_tag)
                for key, value in border_properties.items():
                    border_element.set(qn(f'w:{key}'), str(value))
                tcPr.append(border_element)

def create_advanced_word_report(df, vessel_name, start_date, end_date):
    document = Document("Fleet Performance Template.docx")

    # Add title
    document.add_heading(f"Vessel Performance Report - {vessel_name}", level=1)

    # Add report generation date
    document.add_paragraph(f"Report Generated: {datetime.date.today().strftime('%Y-%m-%d')}")
    document.add_paragraph(f"Period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")

    # Add the DataFrame as a table
    table = document.add_table(rows=1, cols=df.shape[1])
    # No style applied here, as per previous discussion
    table.autofit = True

    # Add table header
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = col_name
        # Set header cell background color and bold text
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'D9D9D9') # Light grey color
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Set borders for header cells
        set_cell_border(
            hdr_cells[i],
            top={"sz": 8, "val": "single", "color": "000000"},
            left={"sz": 8, "val": "single", "color": "000000"},
            bottom={"sz": 8, "val": "single", "color": "000000"},
            right={"sz": 8, "val": "single", "color": "000000"}
        )

    # Add data rows
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, cell_value in enumerate(row):
            row_cells[i].text = str(cell_value)
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Apply conditional formatting for cell shading based on value (example: highlight 'High' in red)
            if isinstance(cell_value, str) and cell_value.lower() == 'high':
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'FF0000') # Red color
                row_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
                row_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) # White text

            # Set borders for data cells
            set_cell_border(
                row_cells[i],
                top={"sz": 4, "val": "single", "color": "000000"},
                left={"sz": 4, "val": "single", "color": "000000"},
                bottom={"sz": 4, "val": "single", "color": "000000"},
                right={"sz": 4, "val": "single", "color": "000000"}
            )

    # Save the document to a BytesIO object
    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit Application ---

st.set_page_config(layout="wide")

st.title("Vessel Performance Report Generator")

# Sidebar for filters
st.sidebar.header("Report Filters")

# Vessel Name Input
vessel_name = st.sidebar.text_input("Vessel Name", "ExampleVessel")

# Date Range Input
start_date = st.sidebar.date_input("Start Date", datetime.date(2023, 1, 1))
end_date = st.sidebar.date_input("End Date", datetime.date(2023, 12, 31))

# Batch Size and Timeout (fixed values)
batch_size = 10
timeout_seconds = 60

# Placeholder for fetched data
fetched_data = None

# Function to fetch data from Lambda
def fetch_data_from_lambda(vessel, start, end):
    lambda_url = "YOUR_LAMBDA_ENDPOINT_HERE" # Replace with your actual Lambda endpoint
    headers = {"Content-Type": "application/json"}
    payload = {
        "vessel_name": vessel,
        "start_date": start.strftime("%Y-%m-%d"),
        "end_date": end.strftime("%Y-%m-%d"),
        "batch_size": batch_size,
        "timeout_seconds": timeout_seconds
    }
    try:
        response = requests.post(lambda_url, headers=headers, data=json.dumps(payload), timeout=timeout_seconds)
        response.raise_for_status()  # Raise an HTTPError for bad responses (4xx or 5xx)
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"Error fetching data from Lambda: {e}")
        return None

if st.sidebar.button("Generate Report"):
    with st.spinner("Generating report..."):
        data = fetch_data_from_lambda(vessel_name, start_date, end_date)
        if data and "body" in data:
            try:
                # Assuming 'body' contains a JSON string of the data
                report_data = json.loads(data["body"])
                fetched_data = pd.DataFrame(report_data)
                st.success("Report generated successfully!")
            except json.JSONDecodeError as e:
                st.error(f"Error decoding JSON from Lambda response: {e}")
                fetched_data = None
        else:
            st.warning("No data received from Lambda or an error occurred.")
            fetched_data = None

if fetched_data is not None:
    st.subheader("Generated Report Data")
    st.dataframe(fetched_data)

    st.subheader("Download Options")

    # Excel Download
    excel_buffer = io.BytesIO()
    fetched_data.to_excel(excel_buffer, index=False, engine='xlsxwriter')
    excel_buffer.seek(0)
    st.download_button(
        label="Download Excel Report",
        data=excel_buffer,
        file_name=f"Vessel_Performance_Report_{vessel_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # CSV Download
    csv_buffer = io.StringIO()
    fetched_data.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)
    st.download_button(
        label="Download CSV Report",
        data=csv_buffer.getvalue(),
        file_name=f"Vessel_Performance_Report_{vessel_name}.csv",
        mime="text/csv"
    )

    # Word Download
    word_buffer = create_advanced_word_report(fetched_data, vessel_name, start_date, end_date)
    st.download_button(
        label="Download Word Report",
        data=word_buffer,
        file_name=f"Vessel_Performance_Report_{vessel_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
