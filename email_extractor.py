import docx
import requests
from bs4 import BeautifulSoup
import re

# Function to extract email addresses from a website
def extract_emails_from_website(url):
    try:
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.content, 'html.parser')
        emails = set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", soup.text))
        return emails
    except requests.RequestException as e:
        print(f"Error accessing {url}: {e}")
        return set()

# Function to extract URLs from the DOCX file
def extract_urls_from_docx(file_path):
    doc = docx.Document(file_path)
    urls = []
    for paragraph in doc.paragraphs:
        text = paragraph.text
        urls_in_text = re.findall(r'(https?://\S+)', text)  # Find all URLs in the text
        urls.extend(urls_in_text)
    return urls

# Main function to process the DOCX file and extract emails
def main(docx_file):
    urls = extract_urls_from_docx(docx_file)
    all_emails = {}
    no_emails = []

    for url in urls:
        print(f"Extracting emails from {url}...")
        emails = extract_emails_from_website(url)
        if emails:
            all_emails[url] = emails
            print(f"Found emails: {emails}")
        else:
            no_emails.append(url)
            print(f"No emails found on {url}")

    # Save the results to a text file
    with open("emails_output.txt", "w") as f:
        for url, emails in all_emails.items():
            f.write(f"{url}: {', '.join(emails)}\n")
    
    # Save the websites with no emails to a separate file
    with open("no_emails_output.txt", "w") as f_no_email:
        for url in no_emails:
            f_no_email.write(f"{url}\n")

if __name__ == "__main__":
    # Update the file path to the correct location on your machine
    docx_file = r"C:\Users\azimu\Downloads\Companies.docx"
    main(docx_file)
