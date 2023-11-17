import pandas as pd
from datetime import datetime
import numpy as np
import requests

def send_telegram_message(bot_token, chat_id, message):
    try:
        url = f'https://api.telegram.org/bot{bot_token}/sendMessage'
        params = {'chat_id': chat_id, 'text': message}
        response = requests.post(url, params=params)

        if response.ok:
            print("Message sent successfully.")
        else:
            print(f"Error sending message. Response status code: {response.status_code}")
    except Exception as e:
        print(f"Error sending message: {str(e)}")

# Replace 'YOUR_BOT_TOKEN' with your actual bot token
bot_token = '6825338365:AAG3q86XuzYCNeVVRIxci3KR_n9APE2STkA'

# Real 'AGALQA'
chat_id = -1001409713167
# TEST group
# chat_id = -1001409713167
# Message to send

def read_excel_to_dict(excel_file):
    # Create an empty dictionary to store data for each sheet
    all_data_dict = {}

    # Read each sheet and populate the dictionary
    xls = pd.ExcelFile(excel_file)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_file, sheet_name)

        # Create a dictionary for the current sheet
        data_dict = {}

        # Iterate through rows and populate the dictionary
        for _, row in df.iterrows():
            # Check for NaN values in the 'Туылған күні' column
            if pd.notna(row['Туылған күні ']):
                date_key = row['Туылған күні '].strftime("%d.%m")  # Assuming the column name is 'Birthday'
                name_value = row['Аты-жөні']  # Assuming the column name is 'Full Name'

                data_dict[date_key] = name_value

        # Store the data dictionary for the current sheet
        all_data_dict[sheet_name] = data_dict

    return all_data_dict

def get_birthday_names_today(excel_file):
    # Call the function to read data from Excel to dictionaries for each sheet
    result_dict = read_excel_to_dict(excel_file)

    # Get today's date in the format day.month
    today_date = datetime.now().strftime("%d.%m")

    # Collect the list of names for birthdays today
    today_birthday_names = []
    for data_dict in result_dict.values():
        if today_date in data_dict:
            today_birthday_names.append(data_dict[today_date])

    return today_birthday_names

# Replace 'path/to/data.xlsx' with the actual path to your Excel file
excel_file_path = 'data.xlsx'

# Call the function to get the list of names for birthdays today
today_birthday_names = get_birthday_names_today(excel_file_path)

# Print the resulting list of names
print("Today's Birthday Names:")
print(today_birthday_names)

for name in today_birthday_names:
    # Call the function to send the message
    message_text = f'{name}, АҒАЛҚА қоғамдық бірлестігінің атынанан тұған күніңізбен құттықтаймыз!'
    send_telegram_message(bot_token, chat_id, message_text)
