# 1 
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

needed_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
while True:
    directory = single_input('Please select the directory where the files are.')
    if directory [-1] != '\\':
       directory = directory + '\\'
    files_in_dir = os.listdir(directory)

    all_files_found = True

    for file in needed_files:
        if file not in files_in_dir:
            all_files_found = False

    if all_files_found == True:
        break
    else:
        show_message('Error', 'Required files are missing. Please try again.')

  #2
descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
display(descriptions.head())

product_descriptions = {}
for index, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
#print(product_descriptions)

prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
#display(prices.head())

product_prices = {}
for index, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
#print(product_prices)

#3
def get_order_data(invoice_num, filename, directory):
    order = pd.read_excel(directory + filename, header=None)
    display(order)
    order_data={}
    total = 0

    for index, row in order.iterrows():
        if index == 13:
            name = row[2]
        if index == 12:
            email = row[2]
        if index == 8:
            date = row [2]
            date = date.strftime('%m/%d/%Y')
        if index == 23:
            unit_number1=row[1]
            quantity1=row[2]
        if index == 24:
            unit_number2=row[1]
            quantity2=row[2]
        if index == 25:
            unit_number3=row[1]
            quantity3=row[2]
        if index == 14:
            company = row[2]
        if index == 15:
            address = row [2]
        if index == 16:
            city=row[2]
        if index == 17:
            state=row[2]
        if index == 18:
           zipcode=row[2]
    order_data['Data'] = date
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(company)
    order_data['CustomerAddress1'] = str(address)
    order_data['CustomerAddress2'] = str(city) + ', ' + str(state) + ' ' + str(zipcode)

    for index, row in order.iterrows():
        if index >= 23 and index <= 25:
            if pd.isnull(row[1]) == False:
                unit_number = row[1]
                quantity = int(row[2])
                price = product_prices[unit_number]
                amount = price * quantity
                total = total + amount

                item = str(index - 23)
                order_data['Description' + item] = str(product_descriptions[unit_number])
                order_data['Price' + item] = f'${price:,.2f}'
                order_data['Quantity' + item] = str(quantity)
                order_data['Amount' + item] = f'${amount:,.2f}'
    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date       

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

#4
def create_invoice(invoice_num, order_data, directory):

    template_path = directory + 'MasterInvoice.pdf'
    new_pdf_path = directory + 'Invoice ' + str(invoice_num) + '.pdf'

    create_pdf(template_path, new_pdf_path, order_data)




#5
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""
    attachment = directory + 'Invoice ' + str(invoice_num) + '.pdf'
    message = EmailMessage()
    message['Subject'] = subject
    message['To'] = email
    message.set_content(body)
    message = add_attachment(attachment, message)
    gmail_user = 'kmichellebb@gmail.com'
    message['From'] = gmail_user
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(gmail_user, gmail_password)
    server.send_message(message)
    server.quit()

#6
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date =  get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email (invoice_num, name, email, date, directory)
        invoice_num=invoice_num+1
