from flask import Flask, request, render_template, redirect, url_for, send_from_directory, abort
import pandas as pd
import requests
import re
from bs4 import BeautifulSoup
import os

app = Flask(__name__)

# Ensure the static folder exists
if not os.path.exists('static'):
    os.makedirs('static')

# Path for the Excel file
excel_file_path = 'static/organization_info.xlsx'

# Function to remove duplicates from a list
def remove_duplicates(data_list):
    return list(dict.fromkeys(data_list))

# Function to extract emails using regular expressions
def extract_emails(html_text):
    try:
        emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,3}", html_text)
        return remove_duplicates(emails)
    except Exception as e:
        print(f"Error extracting emails: {e}")
        return []

# Function to extract phone numbers using regular expressions
def extract_phone_numbers(html_text):
    try:
        phone_numbers = re.findall(r"(\d{2} \d{3,4} \d{3,4})", html_text)
        phone_numbers += re.findall(r"((?:\d{2,3}|\(\d{2,3}\))?(?:\s|-|\.)?\d{3,4}(?:\s|-|\.)\d{4})", html_text)
        return remove_duplicates(phone_numbers)
    except Exception as e:
        print(f"Error extracting phone numbers: {e}")
        return []

# Function to extract contact data from a URL
def extract_contact_data(url):
    try:
        response = requests.get(url)
        print(f"Processing URL: {url}")
        soup = BeautifulSoup(response.text, 'lxml')

        emails = extract_emails(soup.get_text())
        phones = extract_phone_numbers(soup.get_text())

        # Check if a contact page is available and extract data from it
        contact_link = soup.find('a', string=re.compile('contact', re.IGNORECASE))
        contact_page_url = None
        if contact_link:
            contact_page_url = contact_link['href']
            if not contact_page_url.startswith('http'):
                contact_page_url = f"{response.url.rstrip('/')}/{contact_page_url.lstrip('/')}"
            print(f"Found contact page URL: {contact_page_url}")
            contact_response = requests.get(contact_page_url)
            contact_soup = BeautifulSoup(contact_response.text, 'lxml')

            emails += extract_emails(contact_soup.get_text())
            phones += extract_phone_numbers(contact_soup.get_text())

        return {
            'URL': url,
            'Email': ', '.join(remove_duplicates(emails)),
            'Phone': ', '.join(remove_duplicates(phones)),
            'Contact Page URL': contact_page_url if contact_page_url else 'N/A'
        }
    except Exception as e:
        print(f"Error extracting data from {url}: {e}")
        return None

# Function to save contact data to an Excel file
def save_to_excel(contact_data_list, excel_file=excel_file_path):
    df = pd.DataFrame(contact_data_list)
    df.to_excel(excel_file, index=False)

# Ensure the file exists, or create an empty placeholder file
def ensure_file_exists():
    if not os.path.isfile(excel_file_path):
        pd.DataFrame(columns=['URL', 'Email', 'Phone', 'Contact Page URL']).to_excel(excel_file_path, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        df = pd.read_excel(file)
        urls = df.iloc[:, 0].dropna().tolist()
        contact_data_list = []
        for url in urls:
            contact_data = extract_contact_data(url)
            if contact_data:
                contact_data_list.append(contact_data)
        save_to_excel(contact_data_list)
        return redirect(url_for('download_file'))

@app.route('/manual', methods=['POST'])
def manual_entry():
    urls = request.form['urls'].splitlines()
    contact_data_list = []
    for url in urls:
        contact_data = extract_contact_data(url)
        if contact_data:
            contact_data_list.append(contact_data)
    save_to_excel(contact_data_list)
    return redirect(url_for('download_file'))

@app.route('/download')
def download_file():
    ensure_file_exists()  # Ensure the file exists before attempting to serve it
    return send_from_directory('static', 'organization_info.xlsx')

if __name__ == "__main__":
    app.run(debug=True)
