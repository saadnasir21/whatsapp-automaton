import pandas as pd
import pywhatkit as pwk
import pyautogui
import time
import tkinter as tk
from tkinter import filedialog
import traceback

# Simple helper to send a yes/no poll-like prompt
def send_yes_no_poll():
    """Send a placeholder text asking the recipient to reply yes or no."""
    pyautogui.typewrite("\nPlease reply with Yes or No to confirm your order.")
    pyautogui.press('enter')

# Set up GUI for file selection
def select_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="Select the CSV file", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        print("No file selected. Exiting...")
        return None
    return file_path

# Function to send a WhatsApp message
def send_confirmation_message(name, items, total_balance, phone_number, last_success, send_poll=False):
    # Format item list with quantities
    item_list = "\n".join(f"- {item} x{quantity}" for item, quantity in items)
    
    message = (
        f"Hi {name},\n\n"
        "Thank you for your order of the following items:\n"
        f"{item_list}\n\n"
        f"Your total is Rs{total_balance:.2f}. Could you kindly confirm your order "
        "so we can proceed with dispatching it?\n\n"
        "Looking forward to your confirmation!\n\n"
        "Best regards,\n"
        "ASH Homes Customer Support"
    )

    print(f"Preparing to send message to {name} at {phone_number}...")

    try:
        pwk.sendwhatmsg_instantly(phone_number, message, wait_time=20, tab_close=False)

        print("Waiting for WhatsApp to load and draft message...")
        time.sleep(20)  # Extra wait to ensure message is typed completely

        pyautogui.press('enter')  # Send the message once

        if send_poll:
            send_yes_no_poll()

        time.sleep(10)
        pyautogui.press('esc')  # Dismiss any popup if appears
        
        last_success["name"] = name
        last_success["phone"] = phone_number
        print(f"Message sent to {name} successfully!")

        pyautogui.hotkey('ctrl', 'w')  # Close the browser tab
        time.sleep(1)
        pyautogui.press('enter')  # Confirm tab close if needed
        print("Tab closed. Waiting before next message...")
        time.sleep(20)  # Wait before sending the next message

    except Exception as e:
        print(f"Failed to send message to {name}. Error: {str(e)}")
        print(f"Last successful message was sent to {last_success['name']} at {last_success['phone']}.")
        print("Traceback:", traceback.format_exc())

# Function to process multiple messages
def process_messages(df):
    last_success = {"name": None, "phone": None}
    grouped_orders = df.groupby('Name')

    for order_name, group in grouped_orders:
        billing_name = group.iloc[0]['Billing Name']
        # Create a list of tuples containing item names and their respective quantities
        items = list(zip(group['Lineitem name'], group['Lineitem quantity']))
        total_payment = group['Total'].sum()
        shipping_phone = group.iloc[0]['Shipping Phone'].strip()

        # Format the phone number correctly
        shipping_phone = "+92" + shipping_phone[-10:]

        send_confirmation_message(
            billing_name,
            items,
            total_payment,
            shipping_phone,
            last_success,
            send_poll=True,
        )

# Main execution
file_path = select_file()
if file_path:
    try:
        df = pd.read_csv(file_path, dtype={'Shipping Phone': str, 'Total': float, 'Lineitem quantity': int})  # Specify datatypes
        process_messages(df)
    except Exception as e:
        print(f"An error occurred while reading the file or processing messages: {str(e)}")
        print("Traceback:", traceback.format_exc())


