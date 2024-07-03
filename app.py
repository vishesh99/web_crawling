import csv
import requests
import os
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


# Function to configure Chrome WebDriver with options for downloading files
def configure_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),  # Set the download directory
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    chromedriver_path = 'chromedriver-mac-arm64/chromedriver'  # Replace with the path to your chromedriver executable
    driver_service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=driver_service, options=chrome_options)
    return driver


# Function to search for a keyword and extract hrefs from the website
def search_and_extract_hrefs(driver, keyword):
    base_url = "https://bidplus.gem.gov.in"
    search_url = "https://bidplus.gem.gov.in/all-bids"
    hrefs = []

    try:
        driver.get(search_url)
        time.sleep(3)  # Wait for some time to ensure page loads (adjust as needed)

        # Find the search input field by name
        search_box = driver.find_element(By.NAME, 'searchBid')
        search_box.clear()  # Clear any existing text in the input field
        search_box.send_keys(keyword)  # Enter the keyword into the search input field
        search_box.send_keys(Keys.RETURN)  # Simulate pressing Enter key to perform the search

        # Wait for search results to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".bid_no_hover[href]"))
        )

        # Get the page source and parse it with BeautifulSoup
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find all anchor elements with class 'bid_no_hover' and href attribute
        bid_links = soup.select('.bid_no_hover[href]')

        # Extract href attributes and prepend base_url
        for link in bid_links:
            href = link['href']
            full_url = base_url + href
            hrefs.append(full_url)
            print(f"Keyword: {keyword}, Href: {full_url}")

    except Exception as e:
        print(f"Error during search and extraction: {e}")

    return hrefs


# Function to download a document from a given URL
def download_document(url, filename):
    try:
        # Ensure directory exists
        os.makedirs(os.path.dirname(filename), exist_ok=True)

        # Download file
        response = requests.get(url)
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded: {filename}")
    except Exception as e:
        print(f"Failed to download {filename}: {e}")


# Load keywords from CSV
def load_keywords_from_csv(csv_filename):
    keywords = []
    with open(csv_filename, newline='') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            keywords.extend(row)
    return keywords


# Main function to search each keyword in the CSV, extract hrefs, and download documents
def main():
    input_csv_filename = 'your_file.csv'  # Replace with your input CSV filename
    keywords = load_keywords_from_csv(input_csv_filename)

    driver = configure_driver()

    try:
        for keyword in keywords:
            print(f"Searching for keyword: {keyword}")
            hrefs = search_and_extract_hrefs(driver, keyword)
            if hrefs:
                for index, href in enumerate(hrefs):
                    filename = f"{keyword}_{index}.pdf"  # Adjust filename logic as needed
                    download_document(href, filename)
            else:
                print(f"No hrefs found for keyword '{keyword}'")
    finally:
        driver.quit()  # Close the WebDriver session


if __name__ == "__main__":
    main()
