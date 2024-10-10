import os
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time

url = "https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=9m35m&oscode=wt64a&productcode=optiplex-3090-ultra"

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")

# Use webdriver_manager to automatically manage the ChromeDriver
try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service as ChromeService

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get(url)
    # Wait for a specific element to be present instead of using time.sleep
    wait = WebDriverWait(driver, 20)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')
    
    # Find the link that is in the same "a" as the id of latest-download
    latest_download_link = soup.find("a", id="latest-download")['href']
    
    # Save the latest download link to a file named dell.txt in the same folder as this python file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "dell.txt")
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(latest_download_link)

    # Remove the dell.html file in the same folder as this python file
    html_file_path = os.path.join(script_dir, "dell.html")
    if os.path.exists(html_file_path):
        os.remove(html_file_path)

except ModuleNotFoundError as e:
    exit(1)
except Exception as e:
    exit(1)
finally:
    if 'driver' in locals():
        driver.quit()
