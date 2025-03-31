import pandas as pd
import json
from datetime import datetime
import os

def convert_json_to_excel(json_file_path):
    try:
        # Read the JSON file
        print(f"Reading JSON file: {json_file_path}")
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"\nData type: {type(data)}")
        print(f"Number of records: {len(data)}")
        if data:
            print("\nSample of first record keys:", list(data[0].keys()))
            print("\nSample of first record values:")
            for key, value in data[0].items():
                print(f"{key}: {type(value)} - Length: {len(str(value)) if isinstance(value, str) else 'N/A'}")
        
        # Convert nested dictionaries to strings
        print("\nProcessing data...")
        processed_data = []
        for i, item in enumerate(data):
            if i % 1000 == 0:  # Progress indicator
                print(f"Processing record {i}/{len(data)}")
            processed_item = {}
            for key, value in item.items():
                if isinstance(value, (dict, list)):
                    processed_item[key] = json.dumps(value, ensure_ascii=False)
                else:
                    processed_item[key] = value
            processed_data.append(processed_item)
        
        # Convert to DataFrame
        print("\nConverting to DataFrame...")
        df = pd.DataFrame(processed_data)
        
        # Create Excel filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"vinnova_calls_excel_{timestamp}.xlsx"
        
        print(f"\nCreating Excel file: {excel_filename}")
        print(f"DataFrame shape: {df.shape}")
        print("\nColumn info:")
        for col in df.columns:
            non_null = df[col].count()
            print(f"{col}: {non_null} non-null values")
        
        # Create Excel writer with xlsxwriter engine
        writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
        
        # Write DataFrame to Excel
        df.to_excel(writer, sheet_name='Vinnova Calls', index=False)
        
        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Vinnova Calls']
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#D9E1F2',
            'border': 1
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        })
        
        # Format the header row
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Set column widths and format cells
        print("\nFormatting Excel file...")
        for idx, col in enumerate(df):
            # Get the maximum length of the column
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            
            # Set column width (with a maximum of 100 characters)
            worksheet.set_column(idx, idx, min(max_length + 2, 100))
            
            # Format all cells in the column
            for row_num in range(len(df)):
                worksheet.write(row_num + 1, idx, df.iloc[row_num][col], cell_format)
        
        # Save the Excel file
        print("\nSaving Excel file...")
        writer.close()
        
        print(f"\nExcel file created successfully: {excel_filename}")
        print(f"Total rows: {len(df)}")
        print(f"Columns: {', '.join(df.columns)}")
        
    except Exception as e:
        print(f"Error converting JSON to Excel: {str(e)}")
        import traceback
        print("\nFull error traceback:")
        print(traceback.format_exc())

if __name__ == "__main__":
    # Find the most recent JSON file
    json_files = [f for f in os.listdir('.') if f.startswith('vinnova_calls_selected_fields_') and f.endswith('.json')]
    
    if not json_files:
        print("No JSON files found!")
        print("Looking for files starting with 'vinnova_calls_selected_fields_'")
        print("Available files in directory:")
        for f in os.listdir('.'):
            print(f"- {f}")
        exit(1)
    
    # Get the most recent file
    latest_json = max(json_files, key=os.path.getctime)
    print(f"Found JSON file: {latest_json}")
    convert_json_to_excel(latest_json) 