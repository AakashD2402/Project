import fitz  # PyMuPDF for digital PDFs
import pytesseract  # OCR for scanned PDFs
from pdf2image import convert_from_path  # Convert scanned PDFs to images
import os
import pandas as pd
import ssl
import nltk
from nltk.tokenize import word_tokenize
from collections import Counter

# Fix SSL issue when downloading nltk data
ssl._create_default_https_context = ssl._create_unverified_context  # Ignore SSL errors
nltk.download('punkt')

# Define the paths to the subfolders
pdf_folder = "pdf_files"
subfolders = ["Annual Assurance Reports", "Service Inquiries"]  # Subfolder names
output_csv = "extracted_words.csv"

# Function to clean text
def text_cleanup(text):
    text = text.lower()  # Convert to lowercase
    text = text.replace("（", "").replace(")", "").replace("\n", " ")
    text = text.replace(",", "").replace(".", "")  # Remove punctuation
    return text

# Function to extract text from digital PDFs
def extract_text_from_digital(pdf_path):
    doc = fitz.open(pdf_path)
    extracted_text = []
    
    for page in doc:
        text = page.get_text("text")
        extracted_text.append(text_cleanup(text))

    return " ".join(extracted_text)

# Function to extract text from scanned PDFs using OCR
def extract_text_from_scanned(pdf_path):
    images = convert_from_path(pdf_path)  # Convert PDF pages to images
    extracted_text = []

    for img in images:
        text = pytesseract.image_to_string(img)  # Apply OCR
        extracted_text.append(text_cleanup(text))

    return " ".join(extracted_text)

# Function to determine if a PDF is digital or scanned
def is_pdf_digital(pdf_path):
    doc = fitz.open(pdf_path)
    for page in doc:
        if page.get_text("text").strip():  # If text is found, it's digital
            return True
    return False  # If no text is found, it's a scanned PDF

# Process all PDFs in the specified subfolders
all_extracted_data = []

for subfolder in subfolders:
    folder_path = os.path.join(pdf_folder, subfolder)
    
    # Ensure the folder exists
    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(folder_path, filename)
                
                if is_pdf_digital(pdf_path):
                    extracted_text = extract_text_from_digital(pdf_path)
                else:
                    extracted_text = extract_text_from_scanned(pdf_path)
                
                # Tokenize and get unique words
                words = word_tokenize(extracted_text)
                words = [word for word in words if word.isalpha()]  # Keep only words
                unique_words = list(dict.fromkeys(words))  # Remove duplicates

                # Store the results
                all_extracted_data.append({
                    "PDF File": filename,
                    "Folder": subfolder,  # Capturing the subfolder name
                    "Extracted Words": " ".join(unique_words)
                })

# Convert to DataFrame and save to CSV
df = pd.DataFrame(all_extracted_data)
df.to_csv(output_csv, index=False)

print(f"✅ Extraction complete! Saved results to {output_csv}")
