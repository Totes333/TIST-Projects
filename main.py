import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
import re
import os
from docx import Document
from spellchecker import SpellChecker

def extract_text_from_pdfs(pdf_files):
    texts = []
    for pdf_file in pdf_files:
        try:
            # Use PyMuPDF to extract all text
            text = extract_all_text(pdf_file)
            texts.append(text)
        except Exception as e:
            print(f"Failed to process {pdf_file}: {e}")
    return texts

def extract_all_text(pdf_file):
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()  # Get text from the entire page
    return text


# Step 1: Download PDFs
def download_pdfs(base_url, download_folder):
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    response = requests.get(base_url)
    soup = BeautifulSoup(response.content, 'html.parser')
    pdf_links = [link['href'] for link in soup.find_all('a', href=True) if link['href'].endswith('.pdf')]

    downloaded_files = []
    for pdf_link in pdf_links:
        if not pdf_link.startswith('http'):
            pdf_link = base_url + pdf_link

        pdf_filename = os.path.join(download_folder, pdf_link.split('/')[-1])
        # Check if file already exists before downloading
        if not os.path.exists(pdf_filename):
            pdf_response = requests.get(pdf_link)
            with open(pdf_filename, 'wb') as pdf_file:
                pdf_file.write(pdf_response.content)
            print(f"Downloaded: {pdf_filename}")
        else:
            print(f"Skipping download: {pdf_filename} (already exists)")

        downloaded_files.append(pdf_filename)

    return downloaded_files


# Step 2: Extract Text from PDFs
def extract_text_from_pdfs(pdf_files):
    pdf_texts = []
    for pdf_file in pdf_files:
        try:
            # Assuming using PyMuPDF
            with open(pdf_file, 'rb') as pdf_obj:
                doc = fitz.open(pdf_obj)
                text = ""
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    text += page.get_text()  # Extract text from each page
            pdf_texts.append(text)  # Append extracted text for each PDF
        except Exception as e:
            print(f"Failed to open file {pdf_file}: {e}")
    return pdf_texts


# Step 3: Identify Relevant Sections
def find_english(text):
    """
    Identifies and returns English text sections from a single PDF text string.

    Args:
        text: The extracted text string from a single PDF file.

    Returns:
        A list of strings where each element contains the language name and
        the corresponding text section following the "  VERSION" identifier.
    """
    # Remove leading/trailing whitespace and potential extra newlines
    cleaned_text = text.strip()

    # Remove stray letters by themselves (excluding single letter abbreviations)
    cleaned_text = re.sub(r"\b\w\b", "", cleaned_text, flags=re.UNICODE)

    # Split into sections based on language identifier and version marker
    sections = re.split(r'(?=\b\w+  VERSION\b)', cleaned_text, flags=re.IGNORECASE | re.MULTILINE)

    if re.match(r'^\b\w+  VERSION\b', sections[0] , re.IGNORECASE):
        sections = sections
    else:
        sections.pop(0)

    for i in reversed(range(len(sections))):
        if (not (bool(re.match('ENGLISH  VERSION', sections[i], re.I)))):
            sections.pop(i)

    delim = ' '
    all_english = delim.join(sections)

    return all_english




# Step 4: Filter Useful Information
def clean(text):
    """
    Cleans text by removing extraneous symbols and organizing sentences.

    Args:
        text: The text string from a newsletter section.

    Returns:
        A string containing the cleaned and organized text.
    """
    # Remove leading/trailing whitespace and extra newlines
    cleaned_text = text.strip()

    # Replace underscores with spaces
    cleaned_text = cleaned_text.replace('_', ' ')

    # remove "ENGLISH VERSION" case insensitive

    cleaned_text = re.sub(r'ENGLISH  VERSION','', cleaned_text, flags=re.IGNORECASE)

    # Remove "Page" case sensitive

    cleaned_text = re.sub(r'Page', '', cleaned_text)

    #remove all line breaks

    cleaned_text = cleaned_text.replace('\n', ' ').replace('\r', ' ')

    # Remove multiple consecutive spaces
    cleaned_text = re.sub(r'\s\s+', ' ', cleaned_text)

    # Separate lines of text into individual sentences using sentence delimiters
    cleaned_text = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', cleaned_text, flags=re.MULTILINE)

    # search and save month/year and newsletter title as variables
    #print("cleaned_text: ", cleaned_text[-3])

    final_text = '\n'.join(cleaned_text)

    # Remove multiple consecutive spaces
    final_text = re.sub(r'\s\s+', ' ', final_text)

    final_text = re.sub(r'\s([?.!"](?:\s|$))', r'\1', final_text)

    try:

        match = re.search(
            r'(\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}\b)',
            final_text[-300:])

        if(match == None):
            match = re.search(
                r'(\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}\b)',
                final_text[-500:])




    except:

        print("ERROR: newsletter empty")



    try:

        if match:
            # Extract the matched substring
            month_year = match.group(0)
            # Insert month_year at beginning of newsletter
            final_text = month_year + ": " + final_text

    except:

        print("ERROR: match not found or null")

    match = re.search(r'Mazingira Bora', final_text)

    if match:
        # Extract the matched substring
        title = match.group(0)
        # Insert title at beginning of newsletter
        final_text = "\n\n\n" + title + " " + final_text

    return final_text


# Example usage
base_url = 'https://www.tist.org/i2/moreinfo.php'
download_folder = './pdfs/'

pdf_files = [os.path.join(download_folder, file) for file in os.listdir(download_folder) if file.endswith('.pdf')] #download_pdfs(base_url, download_folder)
pdf_files.sort()
print(pdf_files)
pdf_texts = extract_text_from_pdfs(pdf_files)

print("Raw Newsletter 114: ", pdf_texts[113])

messy_newsletters = []

for i in range(len(pdf_texts)):
    messy_newsletters.append(find_english(pdf_texts[i]))

newsletters = []

for i in range(len(messy_newsletters)):
    newsletters.append(clean(messy_newsletters[i]))

    print("Newsletters", i + 1, ": ", newsletters[i][:100])

# Output newsletter information to a document


# Define download folder path (assuming you have write access)
downloads_folder = os.path.expanduser("~/Downloads")  # Expand user path for downloads

# Create a new document
document = Document()

# Add each string as a new paragraph
for article in newsletters:
    paragraph = document.add_paragraph(article)

# Define filename (replace with your desired name)
filename = "my_document.docx"

# Create the full path to the document
filepath = os.path.join(downloads_folder, filename)

# Save the document
document.save(filepath)

print(f"Document saved successfully: {filepath}")

# EOF