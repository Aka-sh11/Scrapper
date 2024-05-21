import os
import re
import pandas as pd
import docx
import PyPDF2
from pathlib import Path

# Define the directory containing your .dcx or .pdf files
input_directory = "./test/Sample2"

# Initialize an empty DataFrame to store the extracted data
columns = ["filename", "email", "number", "text"]
output_df = pd.DataFrame(columns=columns)

# Regular expressions for email and phone number extraction
email_regex = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}"
phone_regex = r"\+?[0-9]+[-\s]?\(?[0-9]+\)?[-\s]?[0-9]+"


def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])


def extract_text_from_pdf(file_path):
    with open(file_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        return "\n".join([page.extract_text() for page in pdf_reader.pages])


def extract_emails_and_numbers(text):
    emails = re.findall(email_regex, text)
    numbers = re.findall(phone_regex, text)
    return emails, numbers


# Iterate through files in the input directory
for file in os.listdir(input_directory):
    if file.lower().endswith((".docx", ".pdf")):
        file_path = os.path.join(input_directory, file)
        filename = os.path.basename(file_path)

        if file.lower().endswith(".docx"):
            text = extract_text_from_docx(file_path)
        elif file.lower().endswith(".pdf"):
            text = extract_text_from_pdf(file_path)

        emails, numbers = extract_emails_and_numbers(text)

        # Append data to the DataFrame
        new_row = {
            "filename": filename,
            "email": ", ".join(emails),
            "number": ", ".join(numbers),
            "text": text
        }
        output_df = pd.concat(
            [output_df, pd.DataFrame([new_row])], ignore_index=True)

# Save the DataFrame to an .xls file
output_df.to_excel("output.xlsx", index=False)
print("Extraction completed. Output saved to output.xlsx.")
