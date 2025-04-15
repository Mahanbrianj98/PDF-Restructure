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
from fpdf import FPDF  # For creating PDFs

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

# Extract layout features from a page
def extract_layout_features(page):
    blocks = page.get_text("blocks")
    layout_features = []
    for block in blocks:
        x0, y0, x1, y1 = block[:4]  # Bounding box
        text = block[4]  # Text content
        layout_features.append({
            "text": text.strip(),
            "bounding_box": (x0, y0, x1, y1)
        })
    return layout_features

# Compare layout features between two pages
def compare_layouts(features1, features2):
    similarity = 0
    for feature1, feature2 in zip(features1, features2):
        box1 = feature1["bounding_box"]
        box2 = feature2["bounding_box"]
        if abs(box1[0] - box2[0]) < 10 and abs(box1[1] - box2[1]) < 10:
            similarity += 1
    return similarity / max(len(features1), len(features2))

# Identify page based on text and layout features
def identify_page(text, layout_features, templates):
    for company, rules in templates.items():
        text_match = any(keyword in text for keyword in rules["header_keywords"])
        layout_similarity = 0
        if "layout_features" in rules:
            template_features = rules["layout_features"]
            layout_similarity = compare_layouts(layout_features, template_features)
        if text_match or layout_similarity > 0.8:
            return company
    return None

# Process each page
def process_page(args, templates, company_images):
    page_number, pdf_path, output_folder = args
    try:
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[page_number]

        # Extract text and layout features
        text = page.get_text("text")
        layout_features = extract_layout_features(page)

        # Identify company
        company = identify_page(text, layout_features, templates)

        # Extract page image for PDF creation
        images = convert_from_path(pdf_path, dpi=150)
        image = images[page_number]
        image = image.resize((image.width // 2, image.height // 2))  # Resize for faster processing

        if company:
            if company not in company_images:
                company_images[company] = []
            company_images[company].append(image)
            logging.info(f"Page {page_number + 1} classified under {company}.")
        else:
            logging.warning(f"Page {page_number + 1} not matched to any company.")

        pdf_document.close()

    except Exception as e:
        logging.error(f"Error processing page {page_number + 1}: {e}")

# Create PDFs for each company
def create_company_pdfs(company_images, output_folder):
    for company, images in company_images.items():
        company_folder = create_company_folder(output_folder, company)
        output_pdf_path = os.path.join(company_folder, f"{company}.pdf")

        pdf = FPDF()
        pdf.set_auto_page_break(0)
        
        for image in images:
            temp_path = os.path.join(company_folder, "temp_image.jpg")
            image.save(temp_path, "JPEG")  # Save temporary JPEG for FPDF
            pdf.add_page()
            pdf.image(temp_path, x=0, y=0, w=210, h=297)  # A4 dimensions in mm
            os.remove(temp_path)  # Clean up temp file after adding to PDF

        pdf.output(output_pdf_path)
        logging.info(f"Created PDF for {company}: {output_pdf_path}")

# Process multiple files
def process_files(pdf_file_paths, output_folder, progress_var):
    try:
        templates = load_templates()
        company_images = {}

        total_pages = sum([len(fitz.open(pdf_file)) for pdf_file in pdf_file_paths])
        completed_pages = 0

        with ThreadPoolExecutor() as executor:
            for pdf_file_path in pdf_file_paths:
                pdf_document = fitz.open(pdf_file_path)
                total_pages_pdf = len(pdf_document)
                pdf_document.close()

                tasks = [(page_number, pdf_file_path, output_folder) for page_number in range(total_pages_pdf)]
                for result in executor.map(lambda args: process_page(args, templates, company_images), tasks):
                    completed_pages += 1
                    progress_var.set((completed_pages / total_pages) * 100)
                    root.update_idletasks()

        create_company_pdfs(company_images, output_folder)

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
root.title("PDF Processor with Text and Layout Analysis")
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