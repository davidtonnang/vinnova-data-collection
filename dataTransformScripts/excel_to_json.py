import pandas as pd
import json
import os
from datetime import datetime, date, time

def datetime_handler(obj):
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    elif isinstance(obj, time):
        return obj.strftime('%H:%M:%S')
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

def excel_to_json(excel_file, output_file=None):
    try:
        # Read all sheets from the Excel file
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        # Convert each sheet to a dictionary
        json_data = {}
        for sheet_name, df in excel_data.items():
            # Convert DataFrame to records
            records = df.to_dict('records')
            # Clean the data by replacing NaN with None
            cleaned_records = []
            for record in records:
                cleaned_record = {k: (None if pd.isna(v) else v) for k, v in record.items()}
                cleaned_records.append(cleaned_record)
            json_data[sheet_name] = cleaned_records
        
        # If no output file specified, create one based on input filename
        if output_file is None:
            output_file = os.path.splitext(excel_file)[0] + '.json'
        
        # Write to JSON file
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2, default=datetime_handler)
        
        print(f"Successfully converted {excel_file} to {output_file}")
        return True
        
    except Exception as e:
        print(f"Error converting Excel file: {str(e)}")
        return False

if __name__ == "__main__":
    # Get the Excel file from the current directory
    excel_file = "Projektportfölj avancerad digitalisering - 2025-01-09 - FG - INNEHÅLLER PERSONUPPGIFTER.xlsx"
    
    if os.path.exists(excel_file):
        excel_to_json(excel_file)
    else:
        print(f"Error: Could not find the Excel file '{excel_file}'")