import os
import pandas as pd
import re

def extract_attributes(content):
    """Extract company details and attributes from raw text."""
    companies = []
    entries = content.split("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
    
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        company = {
            "Company Name": re.search(r"Company Name\s+(.*?)(?:\n|$)", entry),
            "Company Address": re.search(r"Company Address\s+(.*?)(?:\n|$)", entry),
            "Phone": re.search(r"Phone\s+(.*?)(?:\n|$)", entry),
            "Fax": re.search(r"Fax\s+(.*?)(?:\n|$)", entry),
            "URL": re.search(r"URL\s+(.*?)(?:\n|$)", entry),
            "Company Head's Name": re.search(r"Company Head's Name\s+(.*?)(?:\n|$)", entry),
            "Company Head's Designation": re.search(r"Company Head's Designation\s+(.*?)(?:\n|$)", entry),
            "Company Head's Email ID": re.search(r"Company Head's Email ID\s+(.*?)(?:\n|$)", entry),
            "Contact Person's Name": re.search(r"Contact Person's Name\s+(.*?)(?:\n|$)", entry),
            "Contact Person's Designation": re.search(r"Contact Person's Designation\s+(.*?)(?:\n|$)", entry),
            "Contact Person's Email ID": re.search(r"Contact Person's Email ID\s+(.*?)(?:\n|$)", entry),
            "Company Description": re.search(r"Company Description\s+(.*?)(?:\n|$)", entry),
        }
        clean_company = {key: (match.group(1).strip() if match else "") for key, match in company.items()}
        companies.append(clean_company)

    return pd.DataFrame(companies)

def clean_company_name(df):
    """Clean company names by removing URLs and extra details."""
    df["Company Name"] = df["Company Name"].apply(lambda x: re.sub(r"(http[s]?://\S+|www\.\S+)", "", x).strip())
    return df

def process_text_file(file_path, output_file="company_details_cleaned.xlsx"):
    """Process the text file containing company data and save it as an Excel sheet."""
    with open(file_path, "r", encoding="utf-8") as file:
        content = file.read()

    df = extract_attributes(content)
    df = clean_company_name(df)

    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Data has been successfully extracted and saved to {output_file}!")

if __name__ == '__main__':
    file_path = "emails.txt"
    process_text_file(file_path)
