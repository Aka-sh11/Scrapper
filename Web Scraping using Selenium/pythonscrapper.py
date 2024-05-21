import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By  # Import By class for element location
from docx import Document

# Read URLs from Word document
def read_urls_from_docx(file_path):
    urls = []
    try:
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            for word in paragraph.text.split():
                if word.startswith("http"):
                    urls.append(word)
    except Exception as e:
        print("Error reading URLs from the Word document:", e)
    return urls

# Configure Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode
service = Service()  # Replace with path to chromedriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# Initialize data lists
data = []

# Read URLs from the Word document
urls = read_urls_from_docx("urls.docx")

# Scrape data from each URL
for url in urls:
    try:
        driver.get(url)
        time.sleep(3)  # Wait for page to load

        # Extract title
        title = driver.title

        # Extract headings
        headings = driver.find_elements(By.XPATH, "//h1 | //h2 | //h3 | //h4 | //h5 | //h6")
        headings_text = [heading.text for heading in headings]

        # Extract images
        images = driver.find_elements(By.TAG_NAME, "img")
        image_srcs = [image.get_attribute("src") for image in images]

        # Extract paragraphs
        paragraphs = driver.find_elements(By.TAG_NAME, "p")
        paragraphs_text = [paragraph.text for paragraph in paragraphs]

        # Append data to list
        data.append({
            "URL": url,
            "Title": title,
            "Headings": headings_text,
            "Images": image_srcs,
            "Paragraphs": paragraphs_text
        })
    except Exception as e:
        print(f"Error scraping data from {url}: {e}")

# Close the WebDriver
driver.quit()

# Create a DataFrame
df = pd.DataFrame(data)

# Write DataFrame to Excel file
df.to_excel("scraped_data.xlsx", index=False)

print("Data scraped and saved to scraped_data.xlsx")
