import csv
import pandas as pd
import ssl
import urllib.error
import urllib.request
import sqlite3
import os
from pathlib import Path

# Load sensitive data from environment variables
SPREADSHEET_ID = "107ETlafnwDzyn9DcVxGpJ77UzULnuWPms3L2F6VR1cw" #os.getenv('SPREADSHEET_ID')

if not SPREADSHEET_ID:
    raise ValueError("Environment variable 'SPREADSHEET_ID' is not set. Please set it before running this script.")

CONTEXT = ssl._create_unverified_context()

# Configuration for sheet mappings
SHEET_CONFIGS = [
    {'sheet_name': 'tblOrigin', 'origin_file': 'pharmO.csv', 'final_file': 'pharmF.csv', 'output': 'P.csv', 'label': 'Pharmacy'},
    {'sheet_name': 'tblOO', 'origin_file': 'officerO.csv', 'final_file': 'officerF.csv', 'output': 'O.csv', 'label': 'Officer'},
    {'sheet_name': 'tblHO', 'origin_file': 'helperO.csv', 'final_file': 'helperF.csv', 'output': 'H.csv', 'label': 'Helper'},
]

def download_csv_from_sheet(sheet_name, output_file):
    """Download CSV data from Google Sheets and save to file.
    
    Args:
        sheet_name (str): Name of the sheet in the Google Spreadsheet
        output_file (str): Path to save the CSV file
    """
    try:
        url = f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}'
        response = urllib.request.urlopen(url, context=CONTEXT)
        data = response.read().decode('utf-8')
        rows = list(csv.reader(data.splitlines()))
        
        if len(rows) < 2:
            print(f"Warning: No data found in sheet {sheet_name}")
            return False
            
        header = rows[0]
        data_rows = rows[1:]
        
        with open(output_file, 'w', newline='') as f:
            for row in data_rows:
                if len(row) > 1:
                    name = row[1]
                    for i in range(4, len(header)):
                        label = header[i]
                        value = row[i]
                        if value.strip():
                            f.write(f'{label},{name},{value.lower()}\n')
        
        print(f"Successfully downloaded {sheet_name} to {output_file}")
        return True
        
    except urllib.error.URLError as e:
        print(f"Error downloading {sheet_name}: {e}")
        return False
    except Exception as e:
        print(f"Unexpected error in download_csv_from_sheet: {e}")
        return False

def process_matching(filename, filename2, output_filename):
    """Match data between two CSV files based on DateJob and Job columns.
    
    Args:
        filename (str): Path to first CSV file
        filename2 (str): Path to second CSV file
        output_filename (str): Path to save the matched results
    """
    try:
        # Check if files exist
        if not Path(filename).exists() or not Path(filename2).exists():
            print(f"Error: Input files not found")
            return False
        
        conn = sqlite3.connect(':memory:')
        cursor = conn.cursor()
        
        # Create tables
        cursor.execute('''CREATE TABLE sheet1 (DateJob TEXT NOT NULL, Name TEXT NOT NULL, Job TEXT NOT NULL)''')
        cursor.execute('''CREATE TABLE sheet2 (DateJob TEXT NOT NULL, Name TEXT NOT NULL, Job TEXT NOT NULL)''')
        
        # Read and insert data from first file
        with open(filename, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 3:
                    cursor.execute(
                        'INSERT INTO sheet1 (DateJob, Name, Job) VALUES (?, ?, ?)',
                        (row[0], row[1], row[2])
                    )
        
        # Read and insert data from second file
        with open(filename2, 'r') as f2:
            reader = csv.reader(f2)
            for row in reader:
                if len(row) >= 3:
                    cursor.execute(
                        'INSERT INTO sheet2 (DateJob, Name, Job) VALUES (?, ?, ?)',
                        (row[0], row[1], row[2])
                    )
        
        conn.commit()
        
        # Perform matching query
        query = '''SELECT s1.DateJob, s1.Name, s1.Job, s2.Job, s2.Name  FROM sheet1 as s1  INNER JOIN sheet2 as s2  ON s1.DateJob = s2.DateJob AND s1.Job = s2.Job'''        
        cursor.execute(query)
        results = cursor.fetchall()
        
        # Write results to output file
        with open(output_filename, 'w', newline='') as f3:
            for row in results:
                if row[1] == row[4]:
                    f3.write(f'{row[0]},{row[1]},,{row[2]},,{row[3]},\n')
                else:
                    f3.write(f'{row[0]},{row[1]},,{row[2]},,{row[3]},{row[4]}\n')
        pd = f3
        pd.to_csv("output.txt", index=False, encoding="utf-8")
        conn.close()
        print(f"Successfully matched data and saved to {output_filename}")
        return True
        
    except sqlite3.Error as e:
        print(f"Database error in process_matching: {e}")
        return False
    except Exception as e:
        print(f"Error in process_matching: {e}")
        return False

def download_all_sheets():
    """Download all configured sheets from Google Spreadsheet."""
    print("Starting download of all sheets...")
    
    for config in SHEET_CONFIGS:
        sheet_name = config['sheet_name']
        origin_file = config['origin_file']
        final_file = config['final_file']
        
        # Download origin sheet
        download_csv_from_sheet(sheet_name, origin_file)
        
        # Download final sheet (assuming pattern)
        final_sheet = sheet_name.replace('O', 'F')
        download_csv_from_sheet(final_sheet, final_file)

def process_all_matches():
    """Process matching for all configured sheet pairs."""
    print("Starting matching process...")
    
    for config in SHEET_CONFIGS:
        origin_file = config['origin_file']
        final_file = config['final_file']
        output = config['output']
        label = config['label']
        
        print(f"Processing {label}...")
        if process_matching(origin_file, final_file, output):
            print(f"{label} processing completed successfully")
        else:
            print(f"{label} processing failed")

def main():
    """Main entry point for the OT scheduling script."""
    print("=" * 50)
    print("OT Scheduling Script Started")
    print("=" * 50)
    
    try:
        # Download all sheets
        download_all_sheets()
        
        print("\n" + "=" * 50)
        print("Starting data matching process")
        print("=" * 50 + "\n")
        
        # Process all matches
        process_all_matches()
        
        print("\n" + "=" * 50)
        print("OT Scheduling Script Completed Successfully")
        print("=" * 50)
        
    except Exception as e:
        print(f"Fatal error in main: {e}")
        return False
    
    return True


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
