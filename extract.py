import os
import pdfplumber
from datetime import datetime
import subprocess
import pyodbc

# Function to establish database connection
def get_python_database_connection():
    try:
        con = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for Sql Server};"
            "SERVER=192.168.8.10,1433;"
            "DATABASE=Python;"
            "UID=sa;"
            "PWD=1234;"
            "Trusted_Connection=no;"
        )
        return con
    except pyodbc.Error as e:
        print(f"Database connection error: {e}")
        return None

# Function to create Tenders table if it doesn't exist
def create_temp_table_tenders():
    conn = get_python_database_connection()
    if conn is not None:
        try:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Tenders')
                    BEGIN
                    CREATE TABLE [dbo].[Tenders](
                        [TenderNumber] [nvarchar](max) NULL,
                        [TenderEndSubmissionDateTime] [nvarchar](max) NULL,
                        [ContactNumber] [nvarchar](max) NULL,
                        [TenderType] [nvarchar](max) NULL,
                        [TenderOpeningDateTime] [nvarchar](max) NULL,
                        [ContactAddress] [nvarchar](max) NULL,
                        [NameOfWebSite] [nvarchar](max) NULL,
                        [CrawlingDateTime] [nvarchar](max) NULL,
                        [EarnestMoneyDeposite] [nvarchar](max) NULL,
                        [TenderEstimatedCost] [nvarchar](max) NULL,
                        [Address] [nvarchar](max) NULL,
                        [RequirementWorkBrief] [nvarchar](max) NULL,
                        [TenderProdNo] [nvarchar](max) NULL,
                        [ContactPhone2] [nvarchar](max) NULL,
                        [TenderDetailWorkDescription] [nvarchar](max) NULL,
                        [HTMLcontent] [nvarchar](max) NULL,
                        [Document] [nvarchar](max) NULL,
                        [OrganizationName] [nvarchar](max) NULL
                    );
                    END
                    """
                )
                print("Successfully created table Tenders.")
        except Exception as e:
            print(f"Failed to create table: {e}")
        finally:
            conn.commit()
            conn.close()
    else:
        print("Failed to create connection to database.")

# Function to convert PDF to HTML
def pdf_to_html(pdf_path):
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

# Function to clean value
def clean_value(value):
    if value:
        return value.replace('\n', ' ').strip()
    return ''

# Function to extract horizontal tables for address information
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
                            cleaned_address = clean_value(next_row[cell_index])
                            address_set.add(cleaned_address)

    return address_set

# Function to extract data from PDF
def extract_from_pdf(pdf_path):
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
    tender_prod_no = 'Y' if msme_exemption.lower() == 'yes' else 'N'
    contact_phone_2 = 'Y' if startup_exemption.lower() == 'yes' else 'N'
    crawling_date_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Convert PDF to HTML
    html_content = pdf_to_html(pdf_path)

    data = {
        'TenderNumber': os.path.basename(pdf_path).replace('.pdf', ''),  # Extract TenderNumber from filename
        'TenderEndSubmissionDateTime': bid_opening_date,
        'ContactNumber': total_quantity,
        'TenderType': 'buy',
        'TenderOpeningDateTime': bid_opening_date,
        'ContactAddress': contact_address,
        'NameOfWebSite': 'https://bidplus.gem.gov.in/all-bids',
        'CrawlingDateTime': crawling_date_time,
        'EarnestMoneyDeposite': emd_amount,
        'TenderEstimatedCost': estimated_bid_value,
        'Address': ', '.join(list(address_set)),
        'RequirementWorkBrief': requirement_work_brief,
        'TenderProdNo': tender_prod_no,
        'ContactPhone2': contact_phone_2,
        'TenderDetailWorkDescription': f"{boq_title} - {item_category}",
        'HTMLcontent': html_content,
        'Document': '',  # Placeholder for Document field
        'OrganizationName': organization_name
    }

    return data

# Function to insert data into database with check for existing TenderNumber
def insert_into_database(data):
    conn = get_python_database_connection()
    if conn is not None:
        try:
            with conn.cursor() as cursor:
                # Check if TenderNumber already exists
                cursor.execute("SELECT 1 FROM Tenders WHERE TenderNumber = ?", (data['TenderNumber'],))
                existing_record = cursor.fetchone()

                if existing_record:
                    print(f"TenderNumber '{data['TenderNumber']}' already exists in the database. Skipping insertion.")
                else:
                    # Insert new record
                    cursor.execute(
                        """
                        INSERT INTO Tenders (
                            TenderNumber, TenderEndSubmissionDateTime, ContactNumber, TenderType,
                            TenderOpeningDateTime, ContactAddress, NameOfWebSite, CrawlingDateTime,
                            EarnestMoneyDeposite, TenderEstimatedCost, Address, RequirementWorkBrief,
                            TenderProdNo, ContactPhone2, TenderDetailWorkDescription, HTMLcontent, Document, OrganizationName
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            data['TenderNumber'], data['TenderEndSubmissionDateTime'], data['ContactNumber'], data['TenderType'],
                            data['TenderOpeningDateTime'], data['ContactAddress'], data['NameOfWebSite'], data['CrawlingDateTime'],
                            data['EarnestMoneyDeposite'], data['TenderEstimatedCost'], data['Address'], data['RequirementWorkBrief'],
                            data['TenderProdNo'], data['ContactPhone2'], data['TenderDetailWorkDescription'], data['HTMLcontent'], data['Document'], data['OrganizationName']
                        )
                    )
                    conn.commit()
                    print(f"Data inserted successfully for TenderNumber: {data['TenderNumber']}")
        except Exception as e:
            print(f"Error inserting data into the database: {e}")
        finally:
            conn.close()
    else:
        print("Failed to establish connection to database.")

# Define directory path containing PDF files
directory_path = 'GEM/2024/B'

# Ensure Tenders table exists in the database
create_temp_table_tenders()

# Process each PDF file in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(directory_path, filename)
        extracted_data = extract_from_pdf(pdf_path)
        if extracted_data:
            insert_into_database(extracted_data)
        else:
            print(f"Failed to extract data from {pdf_path}")
    else:
        continue
