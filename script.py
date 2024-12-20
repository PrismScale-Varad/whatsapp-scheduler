import pandas as pd
import pywhatkit
from datetime import datetime, timedelta

# Load the Excel file
file_path = "schedule.xlsx"  # Replace with your Excel file path
data = pd.read_excel(file_path)

# Iterate through the data and send WhatsApp messages
for index, row in data.iterrows():
    phone_number = f"+91{row['Phone Number']}"  # Adjust the country code if needed
    message = row['Message']
    time = row['Time']
    scheduled_time = datetime.now().replace(hour=time.hour, minute=time.minute, second=0, microsecond=0)
    
    if scheduled_time < datetime.now():
        scheduled_time += timedelta(days=1)
    
    hours, minutes = scheduled_time.hour, scheduled_time.minute

    print(f"Scheduling message to {phone_number} at {hours}:{minutes}")
    pywhatkit.sendwhatmsg(phone_number, message, hours, minutes)
