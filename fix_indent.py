from pathlib import Path
path = Path('streamlit_app.py')
text = path.read_text(encoding='utf-8')
text = text.replace('cleaned_records = []\n        for item in result:', 'cleaned_records = []\n    for item in result:', 1)
text = text.replace('cleaned_records = []\n    for item in result:\n        vessel = None\n            if isinstance(item, dict) and \'vessel_name\' in item:', 'cleaned_records = []\n    for item in result:\n        vessel = None\n        if isinstance(item, dict) and \'vessel_name\' in item:', 1)
path.write_text(text, encoding='utf-8')
