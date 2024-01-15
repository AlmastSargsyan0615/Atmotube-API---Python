import os
import json
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import requests

def flatten_coords(item):
    if isinstance(item, dict):
        flattened_item = {**item, **item.get('coords', {})}
        del flattened_item['coords']
        return flattened_item
    else:
        # Handle the case where item is not a dictionary (e.g., it's a string)
        return {'time': '', 'voc': '', 't': '', 'h': '', 'p': '', 'pm1': '', 'pm25': '', 'pm10': '', 'lat': '', 'lon': ''}

def save_to_excel_and_csv(data, start_date, end_date, mac_address):
    flattened_data = [flatten_coords(item) for item in data.get('data', {}).get('items', [])]
    
    df = pd.DataFrame(flattened_data)

    df = df.rename(columns={
        'time': 'Date (GMT)',
        'voc': 'VOC (ppm)',
        't': 'Temperature (C)',
        'h': 'Humidity (%)',
        'p': 'Pressure (mbar)',
        'pm1': 'PM1 (ug/m3)',
        'pm25': 'PM2.5 (ug/m3)',
        'pm10': 'PM10 (ug/m3)',
        'lat': 'Latitude',
        'lon': 'Longitude'
    })

    mac_address_folder = mac_address.replace(":", "-")

    excel_directory = f'excel_data/{mac_address_folder}'

    # Create the directory if it doesn't exist
    os.makedirs(excel_directory, exist_ok=True)
    
    excel_file_path = f'{excel_directory}/{start_date}_{end_date}.xlsx'

    # Check if the file already exists
    if os.path.exists(excel_file_path):
        # If the file exists, remove it
        os.remove(excel_file_path)

    if not data['data']['items']:  # Check if items are empty
        excel_file_path = f'{excel_directory}/empty_{start_date}_{end_date}.xlsx'

    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        worksheet = writer.sheets['Sheet1']
        
        for row in worksheet.iter_rows(min_row=2, max_col=1, max_row=worksheet.max_row):
            for cell in row:
                cell.number_format = 'yyyy-mm-dd hh:mm'

    csv_directory = f'csv_data/{mac_address_folder}'
        
    # Create the directory if it doesn't exist
    os.makedirs(csv_directory, exist_ok=True)
        
    csv_file_path = f'{csv_directory}/{start_date}_{end_date}.csv'

    # Check if the file already exists
    if os.path.exists(csv_file_path):
        # If the file exists, remove it
        os.remove(csv_file_path)

    if not data['data']['items']:  # Check if items are empty
        csv_file_path = f'{csv_directory}/empty_{start_date}_{end_date}.csv'

    # Save data to CSV file
    df.to_csv(csv_file_path, index=False, sep=';')


def get_correct_end_date(start_date):
    end_date = min(datetime.now(), start_date + timedelta(days=7))
    return end_date

def get_user_input_start_date():
    while True:
        start_date_str = input("Please input start date (YYYY-MM-DD): ")
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
            return start_date
        except ValueError:
            print("Invalid date format. Please use the format YYYY-MM-DD.")

if __name__ == "__main__":
    # Read configuration from config.json
    with open('config.json', 'r') as config_file:
        config = json.load(config_file)
    
    url = config['url']
    api_key = config['api_key']
    atmotube_mac_addresses = config['atmotube_mac_addresses']
    start_date_str = config['start_date']

    # Check if start_date is provided as a command-line argument
    start_date_str = input("Please input start date (YYYY-MM-DD): ") if not '--start_date' in ' '.join(os.sys.argv) else None

    if start_date_str:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    else:
        start_date = get_user_input_start_date()

    end_date = get_correct_end_date(start_date)

    for mac_address in atmotube_mac_addresses:
        params = {
            'api_key': api_key,
            'mac': mac_address,
            'order': 'desc',
            'format': 'json',
            'offset': 0,
            'limit': 50,
            'start_date': start_date.strftime('%Y-%m-%d'),
            'end_date': end_date.strftime('%Y-%m-%d')
        }

        response = requests.get(url, params=params)

        if response.status_code == 200:
            data = response.json()
            print(f"Total records for MAC {mac_address}: {data['data']['total']}")
            
            for item in data['data']['items']:
                flattened_item = flatten_coords(item)
                flattened_item['time'] = datetime.strptime(item['time'], '%Y-%m-%dT%H:%M:%S.%fZ').strftime('%Y-%m-%d %H:%M')
                print(flattened_item)

            save_to_excel_and_csv(data, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'), mac_address)
        else:
            print(f"Error: {response.status_code}, {response.text}")
