# Conversion and Sorting Program
# Version 0.0.1
# By Brian Mahan
# Date: 2023-10-01
# Description: This program processes PDF files, extracts text and images, and sorts them based on templates defined in a JSON file.
# It uses OCR for text extraction and allows the user to specify output directories.
# It also includes a GUI for user interaction and progress tracking.

import os
import fitz  # PyMuPDF
import json
import re
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import logging
from tkinter import Tk, Button, Label, IntVar
from tkinter.filedialog import askopenfilenames, askdirectory
from tkinter import ttk
import threading
from concurrent.futures import ThreadPoolExecutor

# Configure logging
logging.basicConfig(
    filename='pdf_processing.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Load template JSON
def load_templates(json_path="C:\\Users\\Brian Mahan\\OneDrive - waynecompany.com\\Desktop\\Python scripts\\Sorting program\\Template.json"):
    try:
        with open(json_path, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        logging.error("Template JSON file not found.")
        return {}

# Normalize file path
def normalize_path(file_path):
    return os.path.abspath(os.path.normpath(file_path.strip()))

# Create company folder in the output directory
def create_company_folder(base_output_folder, company_name):
    """Creates a subfolder for each company inside the base output folder."""
    company_folder = os.path.join(base_output_folder, company_name)
    if not os.path.exists(company_folder):
        os.makedirs(company_folder)
        logging.info(f"Created folder for {company_name}: {company_folder}")
    return company_folder

# Identify company based on header text
def identify_company(text, templates):
    for company, rules in templates.items():
        if any(keyword in text for keyword in rules["header_keywords"]):
            return company
    return None

# Extract data using regex based on company rules
def extract_data(text, templates, company):
    extracted_info = {}
    if company in templates:
        for key, pattern in templates[company]["regex_patterns"].items():
            match = re.search(pattern, text)
            extracted_info[key] = match.group() if match else None
    return extracted_info

# Process each page using OCR
def process_page(args, templates):
    page_number, pdf_path, output_folder = args
    try:
        pdf_document = fitz.open(pdf_path)  # Open the PDF for page count
        images = convert_from_path(pdf_path, dpi=150)  # Batch conversion at lower DPI
        
        if page_number >= len(images):
            return f"Page {page_number + 1} out of range for {pdf_path}."

        image = images[page_number]

        # Resize image for faster processing
        image = image.resize((image.width // 2, image.height // 2))

        # Run OCR
        ocr_text = pytesseract.image_to_string(image, lang='eng', config="--psm 3")

        # Identify company and extract relevant data
        company = identify_company(ocr_text, templates)
        extracted_data = extract_data(ocr_text, templates, company)

        # Extract order number for file naming
        order_number = extracted_data.get("Order Number") if extracted_data else None
        file_name = f"{order_number}.png" if order_number else f"page_{page_number + 1}.png"

        if company:
            # Create folder for the company
            company_folder = create_company_folder(output_folder, company)
            
            # Save the image into the company's folder with the extracted name
            output_path = os.path.join(company_folder, file_name)
            image.save(output_path, "PNG")
            logging.info(f"Page saved as image for {company}: {output_path}")
        else:
            logging.warning(f"No match found for page {page_number + 1} in {pdf_path}")

        pdf_document.close()
        return f"Page {page_number + 1} processed successfully."

    except Exception as e:
        logging.error(f"Error processing page {page_number + 1}: {e}")
        return f"Error processing page {page_number + 1}: {e}"

# Process multiple files with parallel execution
def process_files(pdf_file_paths, output_folder, progress_var):
    try:
        templates = load_templates()
        total_pages = sum([len(fitz.open(pdf_file)) for pdf_file in pdf_file_paths])
        completed_pages = 0

        with ThreadPoolExecutor() as executor:
            future_results = []
            for pdf_file_path in pdf_file_paths:
                pdf_document = fitz.open(pdf_file_path)
                total_pages_pdf = len(pdf_document)
                pdf_document.close()

                tasks = [(page_number, pdf_file_path, output_folder) for page_number in range(total_pages_pdf)]
                future_results.extend(executor.map(lambda args: process_page(args, templates), tasks))

            for result in future_results:
                completed_pages += 1
                progress_var.set((completed_pages / total_pages) * 100)  # Update progress
                root.update_idletasks()

    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")

# Start PDF conversion from GUI
def start_convert_pdf():
    try:
        Tk().withdraw()
        pdf_file_paths = askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])
        if not pdf_file_paths:
            raise FileNotFoundError("No files selected. Exiting...")

        output_folder_name = askdirectory(title="Select an Output Folder")
        if not output_folder_name:
            raise ValueError("No output folder selected. Exiting...")
        output_folder_name = normalize_path(output_folder_name)

        if not os.path.exists(output_folder_name):
            create_output_folder(output_folder_name)

        # Process all selected files in parallel
        threading.Thread(target=lambda: process_files(pdf_file_paths, output_folder_name, progress_var)).start()

    except Exception as e:
        logging.error(f"An error occurred: {e}")

# Create output directory
def create_output_folder(folder_name):
    try:
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        logging.info(f"Output folder '{folder_name}' created successfully.")
    except OSError as e:
        logging.error(f"Error creating folder {folder_name}: {e}")
        return False
    return True

# Exit GUI
def exit_program():
    root.destroy()

# GUI Setup
root = Tk()
root.title("PDF Processor with Template Matching")
root.geometry("400x250")

title_label = Label(root, text="PDF Processor", font=("Helvetica", 16))
title_label.pack(pady=10)

convert_button = Button(root, text="Convert PDF", command=start_convert_pdf, width=20)
convert_button.pack(pady=10)

exit_button = Button(root, text="Exit", command=exit_program, width=20)
exit_button.pack(pady=10)

progress_var = IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(pady=20, padx=20, fill='x')

root.mainloop()