import requests
import pandas as pd
from datetime import datetime
import json
from urllib.parse import quote
from config import API_KEY

class VinnovaAPI:
    def __init__(self, api_key):
        self.api_key = api_key
        self.base_url = "https://api.vinnova.se/gdp/v1"
        self.headers = {
            "Authorization": api_key,
            "accept": "application/json"
        }

    def get_metadata(self):
        try:
            response = requests.get(
                f"{self.base_url}/metadata",
                headers=self.headers,
                timeout=10
            )
            
            print(f"Response Status: {response.status_code}")
            print(f"Response Headers: {response.headers}")
            print(f"Request Headers: {self.headers}")
            
            response.raise_for_status()
            
            try:
                return response.json()
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON response: {str(e)}")
                print(f"Raw response: {response.text[:500]}...")
                return None
                
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from Vinnova API: {str(e)}")
            return None

    def get_calls(self, oppningsdatum_from, oppningsdatum_to):
        try:
            # Construct the URL with opening date parameters
            url = (f"{self.base_url}/utlysningar"
                  f"?franOppningsdatum={oppningsdatum_from}"
                  f"&tillOppningsdatum={oppningsdatum_to}")
            
            print("\nGenerated URL:")
            print(url)
            print("\nRequest Headers:", self.headers)
            
            response = requests.get(
                url,
                headers=self.headers,
                timeout=10
            )
            
            print(f"Response Status: {response.status_code}")
            print(f"Response Headers: {response.headers}")
            
            response.raise_for_status()
            
            try:
                return response.json()
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON response: {str(e)}")
                print(f"Raw response: {response.text[:500]}...")
                return None
                
        except requests.exceptions.RequestException as e:
            print(f"Error fetching calls from Vinnova API: {str(e)}")
            return None

def format_long_text(text, max_length=80):
    """Format long text by adding line breaks at word boundaries."""
    if not isinstance(text, str):
        return text
    
    words = text.split()
    lines = []
    current_line = []
    current_length = 0
    
    for word in words:
        if current_length + len(word) + 1 <= max_length:
            current_line.append(word)
            current_length += len(word) + 1
        else:
            lines.append(' '.join(current_line))
            current_line = [word]
            current_length = len(word)
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return '\n'.join(lines)

def display_metadata(api_key):
    # Initialize the API client
    vinnova = VinnovaAPI(api_key)
    
    # Fetch metadata
    print("Fetching metadata from Vinnova API...")
    data = vinnova.get_metadata()
    
    if data is None:
        print("No data was retrieved from the API.")
        return
    
    try:
        # Display the data in a table format using pandas for readable output
        if isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)
        
        print("\nVinnova Metadata:")
        print("=" * 80)
        print(df.to_string(index=False))
        
        # Save to JSON file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_filename = f"vinnova_metadata_{timestamp}.json"
        
        # Save with proper formatting and UTF-8 encoding
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"\nData has been saved to {json_filename}")
        
    except Exception as e:
        print(f"Error processing the data: {str(e)}")
        print("Please check the API response to verify the data structure.")

def display_calls(api_key, oppningsdatum_from, oppningsdatum_to):
    # Initialize the API client
    vinnova = VinnovaAPI(api_key)
    
    # Fetch calls
    print(f"Fetching calls with opening dates from {oppningsdatum_from} to {oppningsdatum_to}...")
    data = vinnova.get_calls(oppningsdatum_from, oppningsdatum_to)
    
    if data is None:
        print("No data was retrieved from the API.")
        return
    
    try:
        # Convert the data to a pandas DataFrame
        df = pd.DataFrame(data)
        
        # Select and rename specific columns for display
        display_columns = {
            'diarienummer': 'Case Number',
            'titel': 'Title',
            'titelEng': 'Title (English)',
            'oppningsdatum': 'Opening Date',
            'stangningsdatum': 'Closing Date'
        }
        
        # Create a simplified view of the data
        if len(df) > 0:
            display_df = df[display_columns.keys()].rename(columns=display_columns)
            
            print("\nVinnova Calls Summary:")
            print("=" * 100)
            pd.set_option('display.max_colwidth', 50)  # Limit column width for better display
            print(display_df.to_string(index=False))
            print("\nTotal number of calls:", len(df))
            
            # Save complete data to JSON file with formatted text
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            json_filename = f"vinnova_calls_{timestamp}.json"
            
            # Format long text fields before saving
            formatted_data = []
            for item in data:
                formatted_item = item.copy()
                # Format long text fields
                for field in ['beskrivning', 'beskrivningEng', 'titel', 'titelEng']:
                    if field in formatted_item:
                        formatted_item[field] = format_long_text(formatted_item[field])
                formatted_data.append(formatted_item)
            
            # Save with proper formatting and UTF-8 encoding
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(formatted_data, f, ensure_ascii=False, indent=2)
            
            print(f"\nComplete data has been saved to {json_filename}")
            
        else:
            print("No calls found for the specified date range.")
            
    except Exception as e:
        print(f"Error processing the data: {str(e)}")
        print("Raw data structure:", data)

def extract_specific_fields(api_key, oppningsdatum_from, oppningsdatum_to):
    # Initialize the API client
    vinnova = VinnovaAPI(api_key)
    
    # Fetch calls
    print(f"Fetching calls with opening dates from {oppningsdatum_from} to {oppningsdatum_to}...")
    data = vinnova.get_calls(oppningsdatum_from, oppningsdatum_to)
    
    if data is None:
        print("No data was retrieved from the API.")
        return
    
    try:
        # Extract only the specified fields
        selected_fields = ['titel', 'titelEng', 'beskrivning', 'beskrivningEng', 'oppningsdatum', 'stangningsdatum']
        extracted_data = []
        
        for item in data:
            extracted_item = {}
            for field in selected_fields:
                if field in item:
                    # Format long text fields for better readability
                    if field in ['beskrivning', 'beskrivningEng', 'titel', 'titelEng']:
                        extracted_item[field] = format_long_text(item[field])
                    else:
                        extracted_item[field] = item[field]
                else:
                    extracted_item[field] = None
            extracted_data.append(extracted_item)
        
        # Save to JSON file with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_filename = f"vinnova_calls_selected_fields_{timestamp}.json"
        
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(extracted_data, f, ensure_ascii=False, indent=2)
        
        print(f"\nExtracted data has been saved to {json_filename}")
        print(f"Total number of records: {len(extracted_data)}")
        
        # Display a sample of the data
        if extracted_data:
            print("\nSample of extracted data (first record):")
            print("=" * 80)
            print(json.dumps(extracted_data[0], ensure_ascii=False, indent=2))
            
    except Exception as e:
        print(f"Error processing the data: {str(e)}")
        print("Raw data structure:", data)

if __name__ == "__main__":
    # Set opening dates to span from 2020 to 2026
    oppningsdatum_from = "2020-01-01"  # Opening start date
    oppningsdatum_to = "2026-12-31"    # Opening end date
    
    # Extract and store specific fields
    extract_specific_fields(API_KEY, oppningsdatum_from, oppningsdatum_to) 