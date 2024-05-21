# Web Scraping with Selenium and Python

This project uses Selenium WebDriver and python-docx to scrape data from a list of URLs provided in a Word document.

## Dependencies

- Python
- Selenium WebDriver
- python-docx
- pandas

## How to Run

1. Ensure you have all the dependencies installed.
2. Replace `Service()` in the script with the path to your chromedriver executable.
3. Place a Word document named "urls.docx" in the same directory as the script. This document should contain the URLs you want to scrape.
4. Run the script. The scraped data will be saved to an Excel file named "scraped_data.xlsx" in the same directory.

## What the Script Does

1. Reads URLs from a Word document.
2. Sets up a headless Chrome WebDriver.
3. Navigates to each URL, waits for the page to load, and then extracts the page title, headings, image sources, and paragraph text.
4. Stores the extracted data in a list of dictionaries, where each dictionary corresponds to a webpage.
5. Converts the data into a pandas DataFrame and saves it to an Excel file.
