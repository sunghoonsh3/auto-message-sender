# Political Campaign Messaging Program

This program automates the process of sending designated messages to a list of recipients extracted from an Excel file. It then marks those recipients in the original Excel file to indicate that messages have been sent to them. The program consists of two main parts: `send_message.py` for sending messages and `mark_the_recipient.py` for marking recipients in the Excel file.

## Features

- **Automated Messaging**: Send messages automatically to phone numbers extracted from an Excel file.
- **Recipient Tracking**: Keep track of which recipients have been messaged and mark them in the Excel sheet.
- **Customization**: Easily customize the message content and Excel file details.

## Prerequisites

- Python 3.x
- `openpyxl` library for handling Excel files
- `pandas` library for data manipulation
- Access to Apple's Messages app for sending iMessages (only works on macOS)

## Setup

1. **Install Python Dependencies**: Run the following command to install the required Python libraries.
pip install openpyxl pandas

2. **Configure Script Parameters**: 
- In `send_message.py,` set `file_path` to the location of your Excel file.
- Ensure the `sheet_name` variable matches the name of the sheet in your Excel file.
- Update the `predetermined_message` variable with the message you wish to send.
- Adjust the `phone_number` variable to match the column in your Excel file that contains phone numbers. This is currently set to extract phone numbers from the 10th column (`str(row[9])`).

3. **Run the Program**:
- First, run `send_message.py` to send out messages to your list of recipients. This script will also create a `sent_messages.txt` file to track which numbers have received messages.
  ```
  python send_message.py
  ```
- Next, run `mark_the_recipient.py` to mark in the Excel file which recipients have received the message. This script will generate a new Excel file named `marked_recipients.xlsx` indicating the status.
  ```
  python mark_the_recipient.py
  ```

## Important Notes

- **Excel File Configuration**: Make sure your Excel file is properly set up with emails in the correct column. The default script configuration of the emails is in the 5th column.
- **Sheet Name**: The scripts default sheet name to `Sheet1`. If your sheet is named differently, ensure you change the `sheet_name` variable in both scripts.
- **Messaging Limitations**: The messaging function is designed to work with Apple's Messages app and will only work on macOS systems.
- **Privacy and Compliance**: Please ensure you consent to message the individuals in your Excel file and comply with any applicable laws regarding automated messaging.
- **Sample Excel File**: You can refer to the phone_data.xlsx file provided in the repository to see the correct formatting for the Excel file.

## Contribution

Feel free to fork this repository and submit pull requests to contribute to this project. Before making major changes, please open an issue to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)
