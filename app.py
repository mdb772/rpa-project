while True: 
    invoice_data_path = select_file(title='Please select the data.')
    directory = invoice_data_path.split('/')
    del directory[-1]
    directory = '\\'.join(directory)
    directory = directory + '\\'
    filelist = os.listdir(directory)
    if ('ProductData.xlsx' in filelist) and ('MasterInvoice.pdf' in filelist) and ('Order1.xlsx' in filelist) and ('Order2.xlsx' in filelist) and ('Order3.xlsx' in filelist) and ('Order4.xlsx' in filelist):
        break
    else:
        show_message('Error', 'The file you selected is incorrect. Please try again.')
print(directory)

product_descriptions = {}
product_prices = {}

df_descriptions = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
for index,row in df_descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
print(product_descriptions)

df_prices = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')
for index,row in df_prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
print(product_prices)

def get_order_data(invoice_num,filename,directory):
    order_data = {}
    temp_file_path = directory + filename
    template = openpyxl.load_workbook(temp_file_path)
    selected_sheet = template['OrderForm']

    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = str(selected_sheet['C9'].value)
    date = date.split(' ')[0]
    date_split = date.split('-')
    date = date_split[1] + '/' + date_split[2] + '/' + date_split[0]
    temp_total = 0

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = selected_sheet['C17'].value + ', ' + selected_sheet['C18'].value + ' ' + str(selected_sheet['C19'].value)
    order_data['Price0'] = f'${product_prices[selected_sheet['B24'].value]:,.2f}'
    order_data['Quantity0'] = str(selected_sheet['C24'].value)
    order_data['Amount0'] = f'${product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value:,.2f}'
    order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]
    temp_total = temp_total + (product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value)
    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = f'${product_prices[selected_sheet['B25'].value]:,.2f}'
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        order_data['Amount1'] = f'${product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        temp_total = temp_total + (product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value)
    if selected_sheet['B26'].value is not None:    
        order_data['Price2'] = f'${product_prices[selected_sheet['B26'].value]:,.2f}'
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        order_data['Amount2'] = f'${product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
        temp_total = temp_total + (product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value)
    if selected_sheet['B27'].value is not None:
        order_data['Price3'] = f'${product_prices[selected_sheet['B27'].value]:,.2f}'
        order_data['Quantity3'] = str(selected_sheet['C27'].value)
        order_data['Amount3'] = f'${product_prices[selected_sheet['B27'].value] * selected_sheet['C27'].value:,.2f}'
        order_data['Description3'] = product_descriptions[selected_sheet['B27'].value]
        temp_total = temp_total + (product_prices[selected_sheet['B27'].value] * selected_sheet['C27'].value)
    if selected_sheet['B28'].value is not None:        
        order_data['Price4'] = f'${product_prices[selected_sheet['B28'].value]:,.2f}'
        order_data['Quantity4'] = str(selected_sheet['C28'].value)
        order_data['Amount4'] = f'${product_prices[selected_sheet['B28'].value] * selected_sheet['C28'].value:,.2f}'
        order_data['Description4'] = product_descriptions[selected_sheet['B28'].value]
        temp_total = temp_total + (product_prices[selected_sheet['B28'].value] * selected_sheet['C28'].value)
    order_data['Total'] = f'${temp_total:,.2f}'
    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)
create_invoice(8515,order_data,directory)


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num,name,email,date,directory):

    gmail_user = "maximeek@gmail.com"

    recipients = [email]

    subject = 'Your Order from PythonTech Co. - Invoice ' + str(invoice_num)

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

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','maximeek@gmail.com','11/13/2023',directory)

filenames = os.listdir(directory)
temp_invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(temp_invoice_num,filename,directory)
        create_invoice(temp_invoice_num,order_data,directory)
        send_email(temp_invoice_num,name,email,date,directory)
        temp_invoice_num = temp_invoice_num + 1