import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

input = r'C:\Users\user\Documents\assignment\Input.xlsx'
df= pd.read_excel(input)
output_directory = 'extracted_articles'
os.makedirs(output_directory, exist_ok=True)

def extract_article_text(url):
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        response.raise_for_status()  # Check if the request was successful

        # Parse the HTML content of the page using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract the title and text of the article
        title = soup.title.text.strip()
        article_text = ' '.join([p.text.strip() for p in soup.find_all('p')])

        return title, article_text
    except Exception as e:
        print(f"Error extracting text from {url}: {e}")
        return None, None

for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    title, article_text = extract_article_text(url)

    # Save the extracted text to a text file
    if title and article_text:
        output_file_path = os.path.join(output_directory, f'{url_id}.txt')
        with open(output_file_path, 'w', encoding='utf-8') as output_file:
            output_file.write(f'{title}\n\n{article_text}')

print("Extraction completed. Text files saved in 'extracted_articles' directory.")