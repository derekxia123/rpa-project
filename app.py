import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

#Q1
while True:
    directory = single_input("Please enter the directory where the ProductData, MasterInvoice, and Order files are located.")

    if os.path.exists(directory) == True:
        directory = directory + '\\'
        directory = directory.replace('\\\\', '\\')

        filenames = os.listdir(directory)

        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break

    show_message("Error", "That folder doesn't exist or is missing required files. Please try again.")

#Q2
descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for index, row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

#Q3
# INSERT ANSWER FOR QUESTION 3 HERE
get_pdf(directory + 'MasterInvoice.pdf')

def get_order_data(invoice_num, filename, directory):

    order_file = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet = order_file['OrderForm']            

    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value
    company = selected_sheet['C15'].value
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = company
    order_data['CustomerAddress1'] = address
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    total = 0
    product_num = 0  
    excel_row = 24      

    while True:
        unit = selected_sheet[f'B{excel_row}'].value
        if pd.isnull(unit) == True:    
            break

        quantity = selected_sheet[f'C{excel_row}'].value
        price = product_prices[unit]
        amount = price * quantity
        total = total + amount

        order_data[f'Description{product_num}'] = product_descriptions[unit]
        order_data[f'Price{product_num}'] = f'${price:,.2f}'
        order_data[f'Quantity{product_num}'] = str(quantity)
        order_data[f'Amount{product_num}'] = f'${amount:,.2f}'

        product_num = product_num + 1
        excel_row = excel_row + 1

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
#print(order_data)
#print(name)
#print(email)
#print(date)

#Q4
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

#Q5
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):

    gmail_user = 'derekxia123@gmail.com'   
    gmail_password = 'xsmktvphiqdrkvlw'

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'derekxia123@gmail.com', '11/13/2023', directory)

#Q6
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):

    gmail_user = 'derekxia123@gmail.com'   
    gmail_password = 'xsmktvphiqdrkvlw'

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'derekxia123@gmail.com', '11/13/2023', directory)
# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
