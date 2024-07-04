import os
import pdfplumber
import pandas as pd
from datetime import datetime
import subprocess
import pyodbc
import logging

# Configure logging if needed
# logging.basicConfig(level=logging.INFO)
# logger = logging.getLogger(__name__)

# Function to convert PDF to HTML using pdftohtml
def pdf_to_html(pdf_path):
    html_path = pdf_path.replace('.pdf', '.html')
    try:
        result = subprocess.run(['pdftohtml', pdf_path, '-stdout'], capture_output=True, text=True, check=True)
        html_content = result.stdout
        return html_content
    except subprocess.CalledProcessError as e:
        print(f"Error converting {pdf_path} to HTML:")
        print(e)
        return None
    except Exception as e:
        print(f"Unexpected error converting {pdf_path} to HTML:")
        print(e)
        return None

# Function to establish database connection
def get_python_database_connection():
    try:
        con = pyodbc.connect(
            "DRIVER={ODBC Driver 18 for Sql Server};"
            "SERVER=192.168.8.10,1433;"
            "DATABASE=Python;"
            "UID=sa;"
            "PWD=1234;"
            "Trusted_Connection=no;"
        )
        return con
    except pyodbc.Error as e:
        # If using logging:
        # logger.error(f"Database connection error: {e}")
        print(f"Database connection error: {e}")
        return None

# Function to insert data into the database
def insert_into_database(data):
    con = None
    try:
        con = get_python_database_connection()
        if con:
            cursor = con.cursor()

            # Example SQL query for inserting data into a table named 'Tenders'
            sql_query = """
                INSERT INTO Tenders (TenderNumber, TenderEndSubmissionDateTime, ContactNumber, TenderType,
                TenderOpeningDateTime, ContactAddress, NameOfWebSite, CrawlingDateTime, EarnestMoneyDeposite,
                TenderEstimatedCost, Address, RequirementWorkBrief, TenderProdNo, ContactPhone2,
                TenderDetailWorkDescription, HTMLcontent, Document, OrganizationName)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """

            # Extract values from the data dictionary
            values = (
                data['TenderNumber'], data['TenderEndSubmissionDateTime'], data['ContactNumber'], data['TenderType'],
                data['TenderOpeningDateTime'], data['ContactAddress'], data['NameOfWebSite'], data['CrawlingDateTime'],
                data['EarnestMoneyDeposite'], data['TenderEstimatedCost'], data['Address'], data['RequirementWorkBrief'],
                data['TenderProdNo'], data['ContactPhone2'], data['TenderDetailWorkDescription'], data['HTMLcontent'],
                data['Document'], data['OrganizationName']
            )

            # Execute the SQL query
            cursor.execute(sql_query, values)
            con.commit()

            print("Data inserted successfully into the database.")
        else:
            print("Error: Database connection is not established.")
    except Exception as e:
        # If using logging:
        # logger.error(f"Error inserting data into the database: {str(e)}")
        print(f"Error inserting data into the database: {str(e)}")
    finally:
        if con:
            con.close()

# Define the directory path and Excel file path
directory_path = 'GEM/2024/B'  # Replace with your directory path
excel_path = "GemProjectBidTest.xlsx"  # Replace with your desired Excel file path

# Define all column headers including existing ones and those extracted from PDFs
all_columns = [
    'TenderNumber', 'TenderEndSubmissionDateTime', 'ContactNumber', 'TenderType',
    'TenderOpeningDateTime', 'ContactAddress', 'NameOfWebSite', 'CrawlingDateTime',
    'EarnestMoneyDeposite', 'TenderEstimatedCost', 'Address', 'RequirementWorkBrief',
    'TenderProdNo', 'ContactPhone2', 'TenderDetailWorkDescription', 'HTMLcontent', 'Document', 'OrganizationName'
]

def clean_value(value):
    if value:
        return value.replace('\n', ' ').strip()
    return ''

def clean_address(address):
    if address:
        cleaned_address = address.replace('\n', ' ').replace('*', '').strip()
        return cleaned_address
    return ''

def extract_horizontal_table(page):
    tables = page.extract_tables()
    address_set = set()

    for table in tables:
        for row_index, row in enumerate(table):
            for cell_index, cell in enumerate(row):
                if cell and ("Address" in cell or "पता" in cell):  # Adjust language as necessary
                    if row_index + 1 < len(table):
                        next_row = table[row_index + 1]
                        if cell_index < len(next_row) and next_row[cell_index]:
                            cleaned_address = clean_address(next_row[cell_index])
                            address_set.add(cleaned_address)

    return list(address_set)

def extract_from_pdf(pdf_path, pdf_path_for_excel):
    bid_opening_date = None
    total_quantity = None
    ministry = None
    department_name = None
    organisation_name = None
    emd_amount = None
    estimated_bid_value = None
    boq_title = None
    item_category = None
    msme_exemption = None
    startup_exemption = None

    address_set = set()

    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            all_tables.extend(tables)
            address_set.update(extract_horizontal_table(page))

        for table in all_tables:
            for row in table:
                if len(row) > 1:
                    key = row[0].strip().lower()
                    value = row[1].strip() if len(row) > 1 and row[1] else ''
                    cleaned_value = clean_value(value)

                    if 'bid opening date' in key:
                        bid_opening_date = cleaned_value.split(' ')[0]  # Extract only date part
                    elif 'total quantity' in key:
                        total_quantity = cleaned_value
                    elif 'ministry' in key:
                        ministry = cleaned_value
                    elif 'department name' in key:
                        department_name = cleaned_value
                    elif 'organisation name' in key:
                        organisation_name = cleaned_value
                    elif 'emd amount' in key:
                        emd_amount = cleaned_value
                    elif 'estimated bid value' in key:
                        estimated_bid_value = cleaned_value
                    elif 'boq title' in key:
                        boq_title = cleaned_value
                    elif 'item category' in key:
                        item_category = cleaned_value
                    elif 'mse exemption' in key:
                        msme_exemption = cleaned_value
                    elif 'startup exemption' in key:
                        startup_exemption = cleaned_value

    contact_address = ', '.join(filter(None, [ministry, department_name]))
    organization_name = organisation_name or department_name or ministry
    requirement_work_brief = f"supply of {boq_title} - {item_category} | Quantity | {total_quantity} - MSME Exemption | {msme_exemption} - Startup Exemption | {startup_exemption}"
    tender_prod_no = 'Y' if msme_exemption and msme_exemption.lower() == 'yes' else 'N'
    contact_phone_2 = 'Y' if startup_exemption and startup_exemption.lower() == 'yes' else 'N'
    crawling_date_time = datetime.now().strftime("%Y-%m-%d %H:%M")  # Ensure datetime module is imported

    # Convert PDF to HTML
    html_content = pdf_to_html(pdf_path)

    data = {
        'TenderNumber': pdf_path.replace('_0.pdf', ''),  # Extract tender number from pdf_path
        'TenderEndSubmissionDateTime': bid_opening_date,
        'ContactNumber': total_quantity,
        'TenderType': 'buy',
        'TenderOpeningDateTime': bid_opening_date,
        'ContactAddress': contact_address,
        'NameOfWebSite': 'https://bidplus.gem.gov.in/all-bids',
        'CrawlingDateTime': crawling_date_time,
        'EarnestMoneyDeposite': emd_amount,
        'TenderEstimatedCost': estimated_bid_value,
        'Address': ' '.join(list(address_set)),
        'RequirementWorkBrief': requirement_work_brief,
        'TenderProdNo': tender_prod_no,
        'ContactPhone2': contact_phone_2,
        'TenderDetailWorkDescription': f"{boq_title} - {item_category}",
        'HTMLcontent': html_content,  # Add HTML content to the DataFrame
        'Document': '',  # Empty field for Document
        'OrganizationName': organisation_name  # Organization Name field
    }

    return data

# Load existing Excel file if it exists
existing_tenders = set()
df_existing = pd.DataFrame(columns=all_columns)  # Create an empty DataFrame with all columns

if os.path.exists(excel_path):
    try:
        df_existing = pd.read_excel(excel_path)
        if 'TenderNumber' in df_existing.columns:
            existing_tenders.update(df_existing['TenderNumber'].astype(str))  # Convert to string for consistency
        else:
            print("Warning: 'TenderNumber' column not found in existing Excel file.")
    except Exception as e:
        print(f"Error loading existing Excel file: {str(e)}")

# Extract data from the PDF directory and get updated existing tenders
extracted_data_list = []

for filename in os.listdir(directory_path):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(directory_path, filename)

        # Extract tender number from pdf_path
        tender_number = pdf_path.replace('_0.pdf', '')

        if tender_number in existing_tenders:
            print(f"Tender {tender_number} has already been crawled. Skipping...")
            continue

        extracted_data = extract_from_pdf(pdf_path, pdf_path)  # Pass pdf_path as the pdf_path parameter
        if extracted_data:
            extracted_data_list.append(extracted_data)

            # Insert extracted data into the database
            insert_into_database(extracted_data)

        # Add the tender number to existing tenders set
        existing_tenders.add(tender_number)

# Create a DataFrame with the extracted data
df_new = pd.DataFrame(extracted_data_list)

# Append new data to the existing DataFrame
df_combined = pd.concat([df_existing, df_new], ignore_index=True, sort=False)

# Save the combined DataFrame to the Excel file
with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
    df_combined.to_excel(writer, index=False)

print(f"Data has been successfully written to {excel_path}")
