import os
import pdfplumber
import pandas as pd
from datetime import datetime
import subprocess

# Function to convert PDF to HTML using pdftohtml
def pdf_to_html(pdf_path):
    html_path = pdf_path.replace('.pdf', '.html')
    try:
        subprocess.run(['pdftohtml', '-s', '-noframes', '-nodrm', '-noimages', pdf_path, html_path], check=True)
        with open(html_path, 'r', encoding='utf-8') as html_file:
            html_content = html_file.read()
        return html_content
    except subprocess.CalledProcessError as e:
        print(f"Error converting {pdf_path} to HTML:")
        print(e)
        return None
    except Exception as e:
        print(f"Unexpected error converting {pdf_path} to HTML:")
        print(e)
        return None
    finally:
        if os.path.exists(html_path):
            os.remove(html_path)  # Remove temporary HTML file

# Define the directory path and Excel file path
directory_path = 'GEM/2024/B'  # Replace with your directory path
excel_path = "GemProjectBidTest.xlsx"  # Replace with your desired Excel file path

# Define all column headers including existing ones and those extracted from PDFs
all_columns = [
    'TenderNumber', 'TenderEndSubmissionDateTime', 'ContactNumber', 'TenderType',
    'TenderOpeningDateTime', 'ContactAddress', 'NameOfWebSite', 'CrawlingDateTime',
    'Value4', 'Value7', 'Value8', 'Value9', 'OrganizationName', 'EarnestMoneyDeposite',
    'TenderEstimatedCost', 'Address', 'Value1', 'Value3', 'Value5', 'Value_10',
    'Value_11', 'RequirementWorkBrief', 'TenderProdNo', 'ContactPhone2', 'TenderDetailWorkDescription',
    'HTMLcontent'  # New column for HTML content
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
    tender_prod_no = 'Y' if msme_exemption.lower() == 'yes' else 'N'
    contact_phone_2 = 'Y' if startup_exemption.lower() == 'yes' else 'N'
    crawling_date_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Convert PDF to HTML
    html_content = pdf_to_html(pdf_path)

    data = {
        'TenderNumber': pdf_path_for_excel.replace('_0.pdf', ''),
        'TenderEndSubmissionDateTime': bid_opening_date,
        'ContactNumber': total_quantity,
        'TenderType': 'buy',
        'TenderOpeningDateTime': bid_opening_date,
        'ContactAddress': contact_address,
        'NameOfWebSite': 'https://bidplus.gem.gov.in/all-bids',
        'CrawlingDateTime': crawling_date_time,
        'Value4': boq_title,
        'Value7': organisation_name,
        'Value8': department_name,
        'Value9': ministry,
        'OrganizationName': organization_name,
        'EarnestMoneyDeposite': emd_amount,
        'TenderEstimatedCost': estimated_bid_value,
        'Address': ' '.join(list(address_set)),
        'Value1': boq_title,
        'Value3': item_category,
        'Value5': total_quantity,
        'Value_10': f"- MSME Exemption | {msme_exemption}",
        'Value_11': f"- Startup Exemption | {startup_exemption}",
        'RequirementWorkBrief': requirement_work_brief,
        'TenderProdNo': tender_prod_no,
        'ContactPhone2': contact_phone_2,
        'TenderDetailWorkDescription': f"{boq_title} - {item_category}",
        'HTMLcontent': html_content
    }

    return data

# Load existing Excel file if it exists
existing_tenders = set()
df_existing = pd.DataFrame(columns=all_columns)

if os.path.exists(excel_path):
    try:
        df_existing = pd.read_excel(excel_path)
        if 'TenderNumber' in df_existing.columns:
            existing_tenders.update(df_existing['TenderNumber'].astype(str))
        else:
            print("Warning: 'TenderNumber' column not found in existing Excel file.")
    except Exception as e:
        print(f"Error loading existing Excel file: {str(e)}")

# Extract data from the PDF directory and get updated existing tenders
extracted_data_list = []

for filename in os.listdir(directory_path):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(directory_path, filename)
        tender_number = pdf_path.replace('_0.pdf', '')

        if tender_number in existing_tenders:
            print(f"Tender {tender_number} has already been crawled. Skipping...")
            continue

        extracted_data = extract_from_pdf(pdf_path, pdf_path)
        if extracted_data:
            extracted_data_list.append(extracted_data)

        existing_tenders.add(tender_number)

# Create a DataFrame with the extracted data
df_new = pd.DataFrame(extracted_data_list)

# Append new data to the existing DataFrame
df_combined = pd.concat([df_existing, df_new], ignore_index=True, sort=False)

# Save the combined DataFrame to the Excel file
with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
    df_combined.to_excel(writer, index=False)

print(f"Data has been successfully written to {excel_path}")
