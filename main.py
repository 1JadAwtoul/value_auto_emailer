
import os
import logging
import time
import win32com.client
import sqlite3

# Setup logging
logging.basicConfig(filename='send_invoices.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

logging.info('Script started')

# Configuration
INVOICE_FOLDER = 'C:\\Users\\aawtoul\\Desktop\\invoices'  # Update this path
TARGET_EMAIL = 'aawtoul@jetrord.com'  # Update this email
DELAY_SECONDS = 15  # Delay between emails in seconds

# Connect to the SQLite database
try:
    conn = sqlite3.connect('sent_invoices.db')
    cursor = conn.cursor()
    logging.info('Database connection established')
except Exception as e:
    logging.error(f'Failed to connect to database: {e}')
    exit()

# Email sending function
def send_email_with_attachment(recipient, subject, attachment_path):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = f'INVOICE #: {subject}'
        mail.Body = f'Please find invoice #{subject} attached'
        mail.Attachments.Add(attachment_path)
        mail.Send()
        logging.info(f'Email sent for invoice: {subject}')
    except Exception as e:
        logging.error(f'Failed to send email for invoice {subject}. Error: {e}')

# Function to process and send invoices
def send_invoices():
    for filename in os.listdir(INVOICE_FOLDER):
        if filename.endswith('.pdf'):
            invoice_number = os.path.splitext(filename)[0]

            # Check if the invoice has been sent
            cursor.execute('SELECT invoice_id FROM sent_invoices WHERE invoice_id = ?', (invoice_number,))
            if cursor.fetchone() is None:
                # Send the email
                send_email_with_attachment(TARGET_EMAIL, invoice_number, os.path.join(INVOICE_FOLDER, filename))

                # Record the sent invoice in the database
                cursor.execute("INSERT INTO sent_invoices (invoice_id) VALUES (?)", (invoice_number,))
                conn.commit()
                logging.info(f'Invoice {invoice_number} recorded as sent')

                # Wait for a specified delay before sending the next email
                logging.info(f'Waiting for {DELAY_SECONDS} seconds before sending the next email')
                time.sleep(DELAY_SECONDS)
            else:
                logging.info(f'Invoice {invoice_number} has already been sent')

send_invoices()

conn.close()
logging.info('Script finished')
