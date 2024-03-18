import openpyxl
import subprocess
import os

# Function to read content and phone numbers from Excel
def read_excel(file_path, sheet_name):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)
    return data

# Function to check if a message has been sent previously
def check_message_sent(phone_number):
    sent_messages_file = 'sent_messages.txt'  # File to track sent messages
    if os.path.exists(sent_messages_file):
        with open(sent_messages_file, 'r') as file:
            sent_messages = file.readlines()
            if phone_number + '\n' in sent_messages:
                return True
    return False

# Function to mark a message as sent
def mark_message_sent(phone_number):
    sent_messages_file = 'sent_messages.txt'  # File to track sent messages
    with open(sent_messages_file, 'a') as file:
        file.write(phone_number + '\n')

# Function to send iMessage
def send_imessage(phone_number, message):
    applescript = f'tell application "Messages" to send "{message}" to buddy "{phone_number}"'
    try:
        subprocess.run(['osascript', '-e', applescript], check=True)
        mark_message_sent(phone_number)  # Mark the message as sent only if sending is successful
    except subprocess.CalledProcessError as e:
        print(f"Failed to send message to {phone_number}: {e}")

# Usage
file_path = '/Users/tristanshin/Desktop/pythonProject/test1.xlsx'
sheet_name = 'Sheet1'  # Replace with your sheet name
predetermined_message = (
    "Type your message here"
)

excel_data = read_excel(file_path, sheet_name)

for row in excel_data:
    phone_number = str(row[9])  # Assuming phone numbers are in the 10th column
    if not check_message_sent(phone_number):
        send_imessage(phone_number, predetermined_message)
