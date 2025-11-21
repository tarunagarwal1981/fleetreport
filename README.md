# 🚢 Vessel Performance Report Tool

A comprehensive Streamlit application for generating vessel performance reports with advanced analytics, multi-format exports, and beautiful visualizations.

## Overview

This tool enables fleet managers and maritime operations teams to:
- Select multiple vessels from a searchable database
- Generate comprehensive performance reports with historical data
- Analyze hull condition, main engine efficiency, and CII ratings
- Export reports in Excel, Word, and CSV formats
- View interactive analytics and trend visualizations

## Features

### 🎯 Vessel Selection
- Load up to 1,200 vessels from the database
- Real-time search and filtering
- Multi-select interface with visual metrics
- Persistent selection across sessions

### 📊 Performance Metrics
- **Hull Condition**: Power loss percentage analysis
  - Good: < 15% excess power
  - Average: 15-25% excess power
  - Poor: > 25% excess power
- **ME Efficiency**: Main Engine Specific Fuel Oil Consumption (SFOC)
  - Good: < 180 g/kWh
  - Average: 180-190 g/kWh
  - Poor: > 190 g/kWh
  - Anomalous: < 160 g/kWh (flagged for review)
- **Potential Fuel Saving**: Excess consumption in MT/Day
- **YTD CII**: Carbon Intensity Indicator ratings (A-E scale)

### 📅 Report Configuration
- Configurable analysis period: 1, 2, or 3 months
- Historical trend analysis
- Batch processing for efficient data retrieval

### 📥 Export Options
- **Excel**: Styled reports with conditional formatting and color coding
- **Word**: Professional documents using template with methodology appendix
- **CSV**: Simple data export for further analysis

### 📈 Analytics Dashboard
- Hull condition distribution charts
- ME efficiency analysis
- Performance trend visualizations
- CII rating distribution and statistics

## Prerequisites

- Python 3.7+
- Access to the AWS Lambda function URL (configured in the app)
- `Fleet Performance Template.docx` file in the repository root (for Word export)

## Installation

1. Install the requirements:

   ```bash
   pip install -r requirements.txt
   ```

2. Ensure the Word template file is present:
   - `Fleet Performance Template.docx` should be in the repository root
   - The template should contain a `{{Template}}` placeholder where the report will be inserted

## Usage

1. Run the application:

   ```bash
   streamlit run streamlit_app.py
   ```

2. **Select Vessels**:
   - Use the search box to filter vessels
   - Check the boxes next to vessels you want to analyze
   - View selected vessels in the expandable section

3. **Configure Report**:
   - Choose the analysis period (1-3 months)
   - Click "Generate Performance Report"

4. **Review Results**:
   - View the performance data table with color-coded indicators
   - Explore the analytics dashboard for insights
   - Download reports in your preferred format

## Technical Details

### Architecture
- **Frontend**: Streamlit web application
- **Backend**: AWS Lambda function (SQL query execution)
- **Database**: Contains vessel particulars, performance summaries, hull performance data, and CII ratings

### Data Sources
The application queries the following database tables:
- `vessel_particulars`: Vessel information and metadata
- `hull_performance_six_months_daily`: Daily hull performance metrics
- `vessel_performance_summary`: Main engine performance data
- `cii_ytd`: Year-to-date Carbon Intensity Indicator ratings

### Key Functions
- `invoke_lambda_function_url()`: HTTP POST requests to Lambda with error handling
- `fetch_all_vessels()`: Cached vessel list retrieval (1-hour TTL)
- `query_report_data()`: Complex data aggregation across multiple queries
- `create_excel_download_with_styling()`: Excel generation with conditional formatting
- `create_enhanced_word_report()`: Word document generation from template

## Configuration

The Lambda function URL is currently hardcoded in the application:
```python
LAMBDA_FUNCTION_URL = "https://yrgj6p4lt5sgv6endohhedmnmq0eftti.lambda-url.ap-south-1.on.aws/"
```

To change this, modify the `LAMBDA_FUNCTION_URL` variable in `streamlit_app.py`.

## Performance Indicators

### Hull Condition Ratings
- 🟢 **Good**: Excess power < 15%
- 🟡 **Average**: Excess power 15-25%
- 🔴 **Poor**: Excess power > 25%

### ME Efficiency Ratings
- 🟢 **Good**: SFOC < 180 g/kWh
- 🟡 **Average**: SFOC 180-190 g/kWh
- 🔴 **Poor**: SFOC > 190 g/kWh
- ⚪ **Anomalous**: SFOC < 160 g/kWh (data quality issue)

### CII Ratings
- **A**: Significantly better performance
- **B**: Better performance
- **C**: Moderate performance
- **D**: Minor inferior performance
- **E**: Inferior performance

## Requirements

See `requirements.txt` for the complete list of dependencies:
- streamlit
- pandas
- boto3
- openpyxl
- python-docx
- requests

## License

See `LICENSE` file for details.

## Support

For issues or questions, please refer to the application's built-in help section or contact the development team.
