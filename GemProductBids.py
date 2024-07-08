import os
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Function to configure Chrome WebDriver with options for downloading files
def configure_driver():
    chrome_options = Options()
    # Remove headless argument to see the live browser actions
    # chrome_options.add_argument("--headless")  # Run in headless mode
    chrome_options.add_argument("--disable-gpu")  # Applicable to Windows
    chrome_options.add_argument("--no-sandbox")  # Bypass OS security model
    chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
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

# Function to tick the checkbox, click the dropdown, and extract hrefs from the website
def extract_hrefs(driver):
    base_url = "https://bidplus.gem.gov.in"
    page_url = "https://bidplus.gem.gov.in/all-bids"
    hrefs = []

    try:
        driver.get(page_url)
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'product'))
        )

        # Tick the checkbox
        checkbox = driver.find_element(By.ID, 'product')
        if not checkbox.is_selected():
            checkbox.click()
        time.sleep(1)  # Allow some time for the page to update

        # Click the dropdown button
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'currentSort'))
        )
        dropdown_button = driver.find_element(By.ID, 'currentSort')
        dropdown_button.click()
        time.sleep(1)  # Allow some time for the dropdown to appear

        # Click the 'Bid Start Date: Latest First' option
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Bid Start Date: Latest First']"))
        )
        sort_option = driver.find_element(By.XPATH, "//a[text()='Bid Start Date: Latest First']")
        sort_option.click()
        time.sleep(3)  # Wait for some time to ensure page loads (adjust as needed)

        # Wait for search results to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".bid_no_hover[href]"))
        )

        # Get the page source and parse it with BeautifulSoup
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find the first 5 anchor elements with class 'bid_no_hover' and href attribute containing 'showbidDocument'
        bid_links = soup.select('.bid_no_hover[href]')
        for link in bid_links:
            href = link['href']
            if "showbidDocument" in href:  # Check if the href contains 'showbidDocument'
                full_url = base_url + href if href.startswith("/") else base_url + "/" + href
                bid_number = link.text.strip().split('/')[-1]
                hrefs.append((full_url, bid_number))
                print(f"Href: {full_url}, Bid Number: {bid_number}")
                if len(hrefs) == 5:  # Stop after finding 5 valid links
                    break

    except Exception as e:
        print(f"Error during extraction: {e}")

    return hrefs

# Function to download a document from a given URL
def download_document(url, bid_number):
    try:
        # Ensure directory 'docs' exists
        docs_dir = os.path.join(os.getcwd(), 'docs')
        os.makedirs(docs_dir, exist_ok=True)

        # Download file with bid number as filename
        filename = os.path.join(docs_dir, f"GEM_2024_B_{bid_number}.pdf")

        response = requests.get(url)
        response.raise_for_status()  # Check for request errors
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded: {filename}")
    except Exception as e:
        print(f"Failed to download {filename}: {e}")

# Main function to extract hrefs and download documents
def main():
    driver = configure_driver()

    try:
        print("Processing the first 5 files on the website")
        hrefs = extract_hrefs(driver)
        if hrefs:
            for href, bid_number in hrefs:
                download_document(href, bid_number)
        else:
            print("No hrefs found on the website")
    finally:
        driver.quit()  # Close the WebDriver session

if __name__ == "__main__":
    main()
