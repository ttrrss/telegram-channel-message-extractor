from telethon.sync import TelegramClient
from telethon.tl.types import InputChannel
from datetime import datetime, timedelta
import pandas as pd

# Replace 'API_ID' and 'API_HASH' with your own values (obtain from https://my.telegram.org/)
api_id = 'api_id'
api_hash = 'api_hash'
phone_number = 'phone_number'  # Your phone number linked to your Telegram account
channel_username = 'channel_username'  # Replace with your channel username


print('Initializing...')

# Function to get the last xxx messages
def get_last_xxx_messages(api_id, api_hash, phone_number, channel_username):
    with TelegramClient('session_name', api_id, api_hash) as client:
        # Connect to Telegram
        client.connect()

        # Ensure you're authorized
        if not client.is_user_authorized():
            client.send_code_request(phone_number)
            client.sign_in(phone_number, input('Enter the code: '))

        # Resolve the channel
        channel = client.get_entity(channel_username)

        # Get the last xxx messages from the channel
        messages = client.get_messages(InputChannel(channel_id=channel.id, access_hash=channel.access_hash), limit=message_count)

        return messages

# Function to save messages to Excel
def save_to_excel(messages):
    data = []

    for message in messages:
        message_date = message.date
        message_date_str = message_date.strftime('%d,%m,%Y')
        message_text = message.text

        data.append({'Date': message_date_str, 'Message': message_text,})

   # Create DataFrame from the list of dictionaries
    df = pd.DataFrame(data)

    # Specify the directory path you want to save data to; like (C:\Python\Data Scraping\...)
    directory = r'Your Directory'

    # Check if the directory exists, create it if it doesn't
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    # Save DataFrame to Excel in the specified directory
    df.to_excel(os.path.join(directory, 'telegram_messages.xlsx'), index=False)

if __name__ == "__main__":

    # Enter number of messages wanted to extract
    message_count = input("Enter number of messages to extract: ")
    message_count = int(message_count)

    # Get the last xxx messages from the channel
    last_xxx_messages = get_last_xxx_messages(api_id, api_hash, phone_number, channel_username)

    # Save messages to Excel
    if last_xxx_messages:
        save_to_excel(last_xxx_messages)
        print("Messages saved to 'telegram_messages.xlsx'")
    else:
        print("No messages found.")