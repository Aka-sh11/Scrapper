import pandas as pd
from bs4 import BeautifulSoup
import requests
import csv
import json

# Load the Excel file
df = pd.read_excel(
    'Scrapping Python Assigment- Flair Insights.xlsx', header=None)

# Extract the links
# replace 0 with the index of the column that contains the links
links = df.iloc[:, 0].tolist()

texts = []
all_links = []
images = []

for link in links:
    # Fetch the HTML page
    response = requests.get(link)

    # Parse the HTML page with BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find all text within the HTML page
    texts.append([element.text for element in soup.find_all('p')])

    # Find all links within the HTML page
    all_links.append([a['href'] for a in soup.find_all('a', href=True)])

    # Find all images within the HTML page
    images.append([img['src'] for img in soup.find_all('img', src=True)])

# Store the scraped data in a CSV file
with open('output.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Text", "Links", "Images"])
    for text, link, image in zip(texts, all_links, images):
        writer.writerow([text, link, image])

# Store the scraped data in a JSON file
data = [{"Text": text, "Link": link, "Image": image}
        for text, link, image in zip(texts, all_links, images)]
with open('output.json', 'w') as json_file:
    json.dump(data, json_file)
