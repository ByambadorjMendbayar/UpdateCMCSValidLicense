"""
CMCS Valid License Updater
===========================

REQUIREMENTS:
-------------
1. Python 3.8 or higher
2. Required packages will be installed automatically if missing:
   - tqdm
   - beautifulsoup4
   - polars
   - requests
   - json5
   - openpyxl, fastexcel, xlsxwriter (for Excel file operations)
   - lxml (for HTML parsing)

3. Required files in the same folder:
   - old_valid_licences.xlsx
   - old_valid_licence_coordinates.xlsx

4. Internet connection to access https://cmcs.mrpam.gov.mn

USAGE:
------
Simply double-click this script and wait for completion.
The script will automatically:
- Install missing packages
- Log in to the CMCS system
- Download current valid licenses
- Compare with existing data
- Update Excel files with new information

OUTPUT FILES:
-------------
- valid_licences.xlsx (updated license information)
- valid_licence_coordinates.xlsx (updated coordinate data)
- old_valid_licences.xlsx (backup for next run)
- old_valid_licence_coordinates.xlsx (backup for next run)

"""

import sys
import os
import subprocess

# Change working directory to script's location
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

print("=" * 70)
print("CMCS Valid License Updater")
print("=" * 70)
print()

# Check and install required packages
print("Checking required packages...")
required_packages = {
    'tqdm': 'tqdm',
    'bs4': 'beautifulsoup4',
    'polars': 'polars',
    'requests': 'requests',
    'json5': 'json5',
    'openpyxl': 'openpyxl',
    'fastexcel': 'fastexcel',
    'xlsxwriter': 'XlsxWriter',
    'lxml': 'lxml'
}

missing_packages = []
for import_name, package_name in required_packages.items():
    try:
        __import__(import_name)
    except ImportError:
        missing_packages.append(package_name)

if missing_packages:
    print(f"[!] Missing packages detected: {', '.join(missing_packages)}")
    print("Installing missing packages... This may take a few minutes.")
    print()
    
    for package in missing_packages:
        try:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--quiet"])
            print(f"[OK] {package} installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"[X] ERROR: Failed to install {package}")
            print(f"   Details: {str(e)}")
            print("\nPlease install manually using: pip install -r requirements.txt")
            input("\nPress Enter to exit...")
            sys.exit(1)
    
    print()
    print("[OK] All required packages installed successfully")
    print()
else:
    print("[OK] All required packages are already installed")
    print()

# Now import all required packages
from tqdm import tqdm
from bs4 import BeautifulSoup
import polars as pl
import requests
import json5
import time
import math
import re

# Set UTF-8 encoding for console output
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

print("=" * 70)
print("CMCS Valid License Updater")
print("=" * 70)
print()

# Check for required files
print("Checking required files...")
required_files = ["old_valid_licences.xlsx", "old_valid_licence_coordinates.xlsx"]
missing_files = [f for f in required_files if not os.path.exists(f)]

if missing_files:
    print(f"\n[X] ERROR: Missing required files:")
    for file in missing_files:
        print(f"   - {file}")
    print("\nPlease ensure these files are in the same folder as this script.")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("[OK] All required files found")
print()

# Login credentials
login_data = {
    'UserName': "Riotinto",
    'Password': "PASSWORDHERE"
}

print("Loading existing license data...")
try:
    old_valid_licences_df = pl.read_excel("./old_valid_licences.xlsx")
    old_valid_licences_coordinates_df = pl.read_excel("./old_valid_licence_coordinates.xlsx")
    print(f"[OK] Loaded {len(old_valid_licences_df)} existing licenses")
    print()
except Exception as e:
    print(f"\n[X] ERROR: Failed to read existing Excel files")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Connecting to CMCS system...")
session = requests.Session()
session.headers.update({
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest',
    'Origin': 'https://cmcs.mrpam.gov.mn',
    'Referer': 'https://cmcs.mrpam.gov.mn/CMCS/',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36'
})
session.cookies.update({'_cmcsCulture': 'en-GB'})

try:
    response = session.get("https://cmcs.mrpam.gov.mn/CMCS/Account/Login")
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    
    csrf_token = None
    verification_token_input = soup.find('input', {'name': '__RequestVerificationToken'})
    if verification_token_input:
        csrf_token = verification_token_input.get('value')
        print("[OK] Connected to CMCS system")
    else:
        print("[X] ERROR: Could not obtain security token from CMCS")
        input("\nPress Enter to exit...")
        sys.exit(1)
except Exception as e:
    print(f"\n[X] ERROR: Failed to connect to CMCS system")
    print(f"   Details: {str(e)}")
    print("\nPlease check your internet connection and try again.")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Logging in...")
try:
    if csrf_token:
        login_data['__RequestVerificationToken'] = csrf_token
    login_response = session.post(
        "https://cmcs.mrpam.gov.mn/CMCS/Account/Login",
        data=login_data,
        allow_redirects=True
    )
    
    if login_response.status_code == 200:
        print("[OK] Successfully logged in")
        print()
    else:
        print(f"[X] ERROR: Login failed (Status code: {login_response.status_code})")
        input("\nPress Enter to exit...")
        sys.exit(1)
except Exception as e:
    print(f"\n[X] ERROR: Login failed")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Retrieving current valid licenses from CMCS...")
try:
    index_count_response = session.get(f'https://cmcs.mrpam.gov.mn/CMCS/License/IndexCount/2?_={str(int(time.time() * 1000))}')
    valid_licences_count = int(index_count_response.content)
    print(f"[OK] Found {valid_licences_count} valid licenses in CMCS system")
    print()
except Exception as e:
    print(f"\n[X] ERROR: Failed to retrieve license count")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

current_valid_licences_list = []
try:
    for page in tqdm(range(1, math.ceil(valid_licences_count/1000) + 1), desc='Fetching valid licences', file=sys.stdout):
        
        valid_licences_chunk = session.post(
            url='https://cmcs.mrpam.gov.mn/CMCS//License/GridData',
            data= {
                'indexType': '2',
                '_search': 'false',
                'nd': str(int(time.time() * 1000)),
                'rows': '1000',
                'page': str(page),
                'sidx': 'Id',
                'sord': 'desc'
            }
        )
        
        current_valid_licences_list.extend([row['cell'] for row in valid_licences_chunk.json()['rows']])
    
    print()
except Exception as e:
    print(f"\n[X] ERROR: Failed to download license data")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Processing license data...")
try:
    current_valid_licences_df = (
        pl.DataFrame(
            current_valid_licences_list, 
            schema=['ID', 'Code', 'Name', 'Type', 'Status', 'Holder', 'Area', 'DisplayText'],
            orient='row'
        ).with_row_index(
            name='OBJECTID',
            offset=1
        ).select(
            ['OBJECTID', 'ID', 'Code', 'Name', 'Type', 'Status', 'Holder', 'Area']
        ).cast({
            'OBJECTID': pl.Int64,
            'ID': pl.Int64,
            'Code': pl.String,
            'Name': pl.String,
            'Type': pl.String,
            'Status': pl.String,
            'Holder': pl.String,
            'Area': pl.Float64
        })
    )
        
    new_valid_licences_df = (
        pl.concat([old_valid_licences_df, current_valid_licences_df])
        .unique(subset='ID', keep='last')
        .with_columns(
            pl.when(~pl.col('ID').is_in(current_valid_licences_df['ID'].to_list()))
            .then(pl.lit('NotValid'))
            .otherwise(pl.col('Status'))
            .alias('Status')
        ).drop('OBJECTID')
        .with_row_index('OBJECTID', offset=1)
        .cast({
            'OBJECTID': pl.Int64
        })
    )
    
    added_valid_licences = [licence for licence in current_valid_licences_list if licence[0] not in old_valid_licences_df['ID'].to_list()]
    
    print(f"[OK] Processed license data")
    print(f"  - New licenses found: {len(added_valid_licences)}")
    print()
    
except Exception as e:
    print(f"\n[X] ERROR: Failed to process license data")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

# Check if there are new licenses to process
if len(added_valid_licences) == 0:
    print("=" * 70)
    print("[i] INFO: No new licenses found")
    print("=" * 70)
    print("\nThe CMCS system has no new valid licenses since the last update.")
    print("All data is up to date. No changes were made to the files.")
    print()
    input("Press Enter to exit...")
    sys.exit(0)

print(f"Downloading coordinate data for {len(added_valid_licences)} new licenses...")
print("This may take a few minutes...")
print()

added_licence_coordinates_list = []
failed_coordinates = []

try:
    for licence in tqdm(added_valid_licences, desc='Fetching coordinates for new valid licences', file=sys.stdout):
        try:
            resp = session.post(f'https://cmcs.mrpam.gov.mn/CMCS/License/Details/{licence[0]}')
            soup = BeautifulSoup(resp.text, 'html.parser')
            script_tag = soup.find('script', type='text/javascript')
            
            coordinate_data = re.search(r'i\s*=\s*(\{.+?\}),(?:e=new f|$)', script_tag.string, re.DOTALL)
            cleaned_coordinate_data = coordinate_data.group(1).strip()
            cleaned_coordinate_data = re.sub(r';\s*$', '', cleaned_coordinate_data)
            cleaned_coordinate_data = re.sub(r':\s*!0\b', ': true', cleaned_coordinate_data)
            cleaned_coordinate_data = re.sub(r':\s*!1\b', ': false', cleaned_coordinate_data)
            added_licence_coordinates_list.append(json5.loads(cleaned_coordinate_data))
        except Exception as e:
            failed_coordinates.append((licence[0], str(e)))
            continue
    
    print()
    
    if failed_coordinates:
        print(f"[!] WARNING: Failed to retrieve coordinates for {len(failed_coordinates)} licenses:")
        for lic_id, error in failed_coordinates[:5]:  # Show first 5 failures
            print(f"   - License ID {lic_id}")
        if len(failed_coordinates) > 5:
            print(f"   ... and {len(failed_coordinates) - 5} more")
        print()
        
except Exception as e:
    print(f"\n[X] ERROR: Failed to download coordinate data")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Processing coordinate data...")
try:
    added_licence_coordinates_df = pl.DataFrame()
    for coordinates in added_licence_coordinates_list:
        df_tmp = (
            pl.DataFrame(
                coordinates['Geometry']['rings'][0],
                schema=['Longitude', 'Latitude'],
                orient='row'
            ).with_columns(
                pl.lit(int(coordinates['Id'])).alias('ID'),
                (pl.int_range(pl.len()) + 1).alias('Point')
            ).select(
                ['ID', 'Point', 'Longitude', 'Latitude']
            ).cast({
                'ID': pl.Int64,
                'Point': pl.Int64,
                'Longitude': pl.Float64,
                'Latitude': pl.Float64
            })
        )
        added_licence_coordinates_df = pl.concat([added_licence_coordinates_df, df_tmp])
    
    new_licence_coordinates_df = (
        pl.concat([old_valid_licences_coordinates_df, added_licence_coordinates_df])
        .unique(subset=['ID', 'Point'], keep='last')
    )
    print("[OK] Coordinate data processed")
    print()
    
except Exception as e:
    print(f"\n[X] ERROR: Failed to process coordinate data")
    print(f"   Details: {str(e)}")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("Saving updated data to Excel files...")
try:
    new_valid_licences_df.write_excel("./valid_licences.xlsx")
    new_licence_coordinates_df.write_excel("./valid_licence_coordinates.xlsx")
    
    new_valid_licences_df.write_excel("./old_valid_licences.xlsx")
    new_licence_coordinates_df.write_excel("./old_valid_licence_coordinates.xlsx")
    
    print("[OK] Files saved successfully")
    print()
    
except Exception as e:
    print(f"\n[X] ERROR: Failed to save Excel files")
    print(f"   Details: {str(e)}")
    print("\nPlease ensure the files are not open in Excel and try again.")
    input("\nPress Enter to exit...")
    sys.exit(1)

print("=" * 70)
print("[OK] UPDATE COMPLETED SUCCESSFULLY")
print("=" * 70)
print()
print(f"Summary:")
print(f"  - Total licenses in system: {len(new_valid_licences_df)}")
print(f"  - New licenses added: {len(added_valid_licences)}")
print(f"  - Coordinate points added: {len(added_licence_coordinates_df)}")
print()
print("Output files:")
print("  - valid_licences.xlsx")
print("  - valid_licence_coordinates.xlsx")
print()
input("Press Enter to exit...")