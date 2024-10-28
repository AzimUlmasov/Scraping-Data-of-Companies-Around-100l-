import docx
from googlesearch import search
import requests
from bs4 import BeautifulSoup
import re
import csv

# Function to extract company names from DOCX file
def extract_company_names_from_docx(file_path):
    doc = docx.Document(file_path)
    company_names = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text and not text.startswith("Company Name") and not text.startswith("Location"):  # Skip headers
            if not re.match(r'^[A-Z].*\s\(.*\)$', text):  # Ignore locations (usually have parentheses)
                company_names.append(text)
    return company_names

# Function to scrape email addresses from a website
def extract_emails_from_website(url):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Raise HTTPError for bad responses
        soup = BeautifulSoup(response.content, 'html.parser')
        emails = set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", soup.text))
        return emails
    except requests.RequestException as e:
        print(f"Error accessing {url}: {e}")
        return set()

# Function to search the company name in Google and return the top website
def get_company_website(company_name):
    query = company_name + " official website"
    try:
        for result in search(query, num=1, stop=1, pause=2):
            return result
    except Exception as e:
        print(f"Error searching for {company_name}: {e}")
        return None

# Main function to get websites and emails for companies listed in the DOCX file
def get_company_info(docx_file):
    company_names = extract_company_names_from_docx(docx_file)
    company_info = []

    for company in company_names:
        print(f"Searching for {company}...")
        website = get_company_website(company)
        if website:
            print(f"Found website: {website}")
            emails = extract_emails_from_website(website)
            company_info.append({
                'company': company,
                'website': website,
                'emails': ', '.join(emails) if emails else 'No emails found'
            })
        else:
            company_info.append({
                'company': company,
                'website': 'No website found',
                'emails': 'N/A'
            })

    return company_info

# Run the process for the uploaded DOCX file
docx_file_path = r"C:\path\to\your\Maschinenbau.docx"  # Update with the actual path to your docx file
company_data = get_company_info(docx_file_path)

# Output results to console
for data in company_data:
    print(f"Company: {data['company']}")
    print(f"Website: {data['website']}")
    print(f"Emails: {data['emails']}")
    print("-" * 40)

# Save the results to a CSV file
with open("company_data_output.csv", "w", newline="", encoding="utf-8") as csvfile:
    fieldnames = ['Company', 'Website', 'Emails']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    for data in company_data:
        writer.writerow({
            'Company': data['company'],
            'Website': data['website'],
            'Emails': data['emails']
        })
