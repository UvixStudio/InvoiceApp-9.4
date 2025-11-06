#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gmail Invoice Scanner - Local Download Script
Downloads all invoice PDFs from Google Sheets to local folder
"""

import os
import sys
import requests
from pathlib import Path
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials

# Configuration
SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"  # Replace with your actual spreadsheet ID
SETTINGS_SHEET = "Settings"
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

def connect_to_sheet():
    """Connect to Google Sheets using credentials"""
    try:
        # Try to use credentials.json file
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPES)
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        print("❌ Error: credentials.json file not found!")
        print("Please download your Google Service Account credentials and save as 'credentials.json'")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error connecting to Google Sheets: {e}")
        sys.exit(1)

def get_settings(sheet):
    """Read settings from Settings sheet"""
    try:
        settings_ws = sheet.worksheet(SETTINGS_SHEET)
        
        # Read values from Settings sheet
        start_date = settings_ws.acell('B1').value  # Start Date
        end_date = settings_ws.acell('B2').value    # End Date
        base_path = settings_ws.acell('I2').value   # Export folder path (was Folder ID)
        
        print(f"📅 Date Range: {start_date} to {end_date}")
        print(f"📁 Base Path: {base_path}")
        
        return {
            'start_date': start_date,
            'end_date': end_date,
            'base_path': base_path
        }
    except Exception as e:
        print(f"❌ Error reading settings: {e}")
        sys.exit(1)

def create_date_folder(base_path, start_date, end_date):
    """Create folder based on date range"""
    try:
        # Format: YYYY-MM-DD_to_YYYY-MM-DD
        folder_name = f"{start_date}_to_{end_date}"
        full_path = Path(base_path) / folder_name
        
        # Create folder if doesn't exist
        full_path.mkdir(parents=True, exist_ok=True)
        
        print(f"✅ Target folder: {full_path}")
        return full_path
    except Exception as e:
        print(f"❌ Error creating folder: {e}")
        sys.exit(1)

def get_existing_files(folder_path):
    """Get list of existing files in folder"""
    try:
        existing = set()
        if folder_path.exists():
            for file in folder_path.iterdir():
                if file.is_file():
                    existing.add(file.name)
        return existing
    except Exception as e:
        print(f"⚠️ Warning: Could not read existing files: {e}")
        return set()

def find_invoice_sheet(sheet):
    """Find the most recent invoices sheet"""
    try:
        worksheets = sheet.worksheets()
        invoice_sheets = [ws for ws in worksheets if ws.title.startswith('invoices ')]
        
        if not invoice_sheets:
            print("❌ No invoice sheets found!")
            return None
        
        # Sort by date in title and get most recent
        invoice_sheets.sort(key=lambda x: x.title, reverse=True)
        latest_sheet = invoice_sheets[0]
        
        print(f"📊 Using sheet: {latest_sheet.title}")
        return latest_sheet
    except Exception as e:
        print(f"❌ Error finding invoice sheet: {e}")
        return None

def download_file(url, filepath, filename):
    """Download a single file from Google Drive link"""
    try:
        # Extract file ID from Google Drive URL
        if 'drive.google.com' in url:
            if '/file/d/' in url:
                file_id = url.split('/file/d/')[1].split('/')[0]
            elif 'id=' in url:
                file_id = url.split('id=')[1].split('&')[0]
            else:
                print(f"⚠️ Could not parse URL: {url}")
                return False
            
            # Direct download URL
            download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        else:
            download_url = url
        
        # Download file
        response = requests.get(download_url, stream=True)
        
        if response.status_code == 200:
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        else:
            print(f"⚠️ Failed to download {filename}: HTTP {response.status_code}")
            return False
            
    except Exception as e:
        print(f"⚠️ Error downloading {filename}: {e}")
        return False

def main():
    """Main execution function"""
    print("=" * 60)
    print("📥 Gmail Invoice Scanner - Local Download")
    print("=" * 60)
    print()
    
    # Connect to Google Sheets
    print("🔗 Connecting to Google Sheets...")
    client = connect_to_sheet()
    sheet = client.open_by_key(SPREADSHEET_ID)
    print("✅ Connected successfully!")
    print()
    
    # Read settings
    print("⚙️ Reading settings...")
    settings = get_settings(sheet)
    print()
    
    # Create target folder
    print("📁 Creating target folder...")
    target_folder = create_date_folder(
        settings['base_path'],
        settings['start_date'],
        settings['end_date']
    )
    print()
    
    # Get existing files
    print("🔍 Checking existing files...")
    existing_files = get_existing_files(target_folder)
    print(f"Found {len(existing_files)} existing files")
    print()
    
    # Find invoice sheet
    print("📊 Finding invoice data...")
    invoice_ws = find_invoice_sheet(sheet)
    if not invoice_ws:
        sys.exit(1)
    print()
    
    # Get all data from sheet
    print("📥 Reading invoice data...")
    all_data = invoice_ws.get_all_values()
    
    if len(all_data) < 14:  # Header + Mini-Log (rows 1-13)
        print("❌ No invoice data found!")
        sys.exit(1)
    
    # Process invoices (starting from row 14, index 13)
    headers = all_data[13]  # Row 14 is headers
    data_rows = all_data[14:]  # Data starts from row 15
    
    # Find column indices
    try:
        download_link_col = headers.index('Download Link')
        attachment_name_col = headers.index('Attachment Name')
    except ValueError as e:
        print(f"❌ Could not find required columns: {e}")
        sys.exit(1)
    
    print(f"Found {len(data_rows)} invoices to process")
    print()
    
    # Download files
    print("⬇️ Starting downloads...")
    print("-" * 60)
    
    downloaded = 0
    skipped = 0
    failed = 0
    
    for idx, row in enumerate(data_rows, start=1):
        if len(row) <= max(download_link_col, attachment_name_col):
            continue
        
        download_link = row[download_link_col]
        attachment_name = row[attachment_name_col]
        
        if not download_link or not attachment_name:
            continue
        
        # Check if file already exists
        if attachment_name in existing_files:
            print(f"⏭️ [{idx}/{len(data_rows)}] Skipping (exists): {attachment_name}")
            skipped += 1
            continue
        
        # Download file
        filepath = target_folder / attachment_name
        print(f"⬇️ [{idx}/{len(data_rows)}] Downloading: {attachment_name}")
        
        if download_file(download_link, filepath, attachment_name):
            downloaded += 1
            print(f"   ✅ Saved to: {filepath}")
        else:
            failed += 1
    
    # Summary
    print()
    print("=" * 60)
    print("📊 Download Summary")
    print("=" * 60)
    print(f"✅ Downloaded: {downloaded} files")
    print(f"⏭️ Skipped (already exist): {skipped} files")
    print(f"❌ Failed: {failed} files")
    print(f"📁 Location: {target_folder}")
    print()
    print("🎉 Done!")

if __name__ == "__main__":
    main()
