import openai
import PyPDF2
from googleapiclient.discovery import build
from google.oauth2 import service_account
from typing import List, Dict

# OpenAI API key
openai.api_key = 'sk-proj-6McfO6lCzvmkIgFI7ddgT3BlbkFJSHeO9jl8x5v69jbfOJlS'

# Global variable for row mapping
row_mapping = {}

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to prompt GPT with the list of variables and extracted PDF text
def get_values_from_gpt(pdf_text, variables):
    prompt = f"""
You are provided with the following medical lab results:

{pdf_text}

Please extract and return the values for the following health data variables, do not include the units of measurement - I only want the numbers:
{', '.join(variables)}

Respond in the format:
Variable Name: Value

If a value is not found, do not include that variable in your response.
"""
    response = openai.ChatCompletion.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=4096
    )
    response_text = response.choices[0].message['content'].strip()
    print("Raw GPT response:\n", response_text)
    return response_text

# Function to write data to Google Sheets
def write_to_google_sheets(data: Dict[str, str], sheet_id: str, sheet_name: str):
    SERVICE_ACCOUNT_FILE = 'keys.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Prepare the data for batch update
    batch_data = []
    for variable, value in data.items():
        if variable in row_mapping:
            batch_data.append({
                'range': f'{sheet_name}!B{row_mapping[variable]}',
                'values': [[value]]
            })

    # Perform batch update
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': batch_data
    }
    request = sheet.values().batchUpdate(spreadsheetId=sheet_id, body=body).execute()

    print("Data successfully written to Google Sheets.")

# Function to get gender input
def get_gender():
    while True:
        gender = input("Enter gender (m/f): ").lower()
        if gender in ['m', 'f']:
            return gender
        else:
            print("Invalid input. Please enter 'm' for male or 'f' for female.")

# Function to set gender-specific variables and mappings
def set_gender_specific_data(gender):
    global row_mapping

    if gender == 'm':
        health_data_variables = [
            "BUN", "Creatinine", "eGFR",
            "Folate", "B12", "Apolipoprotein A", "Apolipoprotein B",
            "Lipoprotein (a)", "Total Cholesterol", "Triglycerides", "HDL Cholesterol", "LDL Cholesterol",
            "Hemoglobin", "Hematocrit", "Iron", "MCV", "RBC",
            "TIBC", "Transferrin % Saturation", "RDW %",
            "Ferritin", "Sodium", "Potassium", "Chloride", "Bicarbonate", "Magnesium",
            "GGT", "Albumin", "AST",
            "ALT", "Glucose", "HbA1C",
            "WBC", "Neutrophil Count", "Basophil Count", "Eosinophil Count",
            "Lymphocyte Count", "Platelets", "MPV", "Alk Phos",
            "Phosphorus", "Calcium", "TSH", "Serum Total Testosterone",
            "Sex hormone binding glob.", "Blood pressure (systolic pressure/diastolic pressure)", "Body mass index", "Waist measurement"
        ]

        row_mapping = {
            "BUN": 3, "Creatinine": 4, "eGFR": 5,
            "Folate": 7, "B12": 8, "Apolipoprotein A": 10, "Apolipoprotein B": 11,
            "Lipoprotein (a)": 12, "Total Cholesterol": 13, "Triglycerides": 14, "HDL Cholesterol": 15, "LDL Cholesterol": 16,
            "Hemoglobin": 18, "Hematocrit": 19, "Iron": 20, "MCV": 21, "RBC": 22,
            "TIBC": 23, "Transferrin % Saturation": 24, "RDW %": 25,
            "Ferritin": 26, "Sodium": 28, "Potassium": 29, "Chloride": 30, "Bicarbonate": 31, "Magnesium": 32,
            "GGT": 34, "Albumin": 35, "AST": 37,
            "ALT": 38, "Glucose": 39, "HbA1C": 40,
            "WBC": 42, "Neutrophil Count": 43, "Basophil Count": 44, "Eosinophil Count": 45,
            "Lymphocyte Count": 46, "Platelets": 47, "MPV": 48, "Alk Phos": 50,
            "Phosphorus": 51, "Calcium": 52, "TSH": 55, "Serum Total Testosterone": 57,
            "Sex hormone binding glob.": 58, "Blood pressure (systolic pressure/diastolic pressure)": 61, "Body mass index": 63, "Waist measurement": 64
        }
    else:  # female
        health_data_variables = [
            "BUN", "Creatinine", "eGFR",
            "Folate", "B12", "Apolipoprotein A", "Apolipoprotein B",
            "Lipoprotein (a)", "Total Cholesterol", "Triglycerides", "HDL Cholesterol", "LDL Cholesterol",
            "Hemoglobin", "Hematocrit", "Iron", "MCV", "RBC",
            "TIBC", "Transferrin % Saturation", "RDW %",
            "Ferritin", "Sodium", "Potassium", "Chloride", "Bicarbonate", "Magnesium",
            "GGT", "Albumin", "AST",
            "ALT", "Glucose", "HbA1C",
            "WBC", "Neutrophil Count", "Basophil Count", "Eosinophil Count",
            "Lymphocyte Count", "Platelets", "MPV", "Alk Phos",
            "Phosphorus", "Calcium", "TSH", "Estradiol",
            "FSH", "Anti-Mullerian hormone", "Prolactin", "Blood pressure (systolic pressure/diastolic pressure)", "Body mass index", "Waist measurement"
        ]

        row_mapping = {
            "BUN": 3, "Creatinine": 4, "eGFR": 5,
            "Folate": 7, "B12": 8, "Apolipoprotein A": 10, "Apolipoprotein B": 11,
            "Lipoprotein (a)": 12, "Total Cholesterol": 13, "Triglycerides": 14, "HDL Cholesterol": 15, "LDL Cholesterol": 16,
            "Hemoglobin": 18, "Hematocrit": 19, "Iron": 20, "MCV": 21, "RBC": 22,
            "TIBC": 23, "Transferrin % Saturation": 24, "RDW %": 25,
            "Ferritin": 26, "Sodium": 28, "Potassium": 29, "Chloride": 30, "Bicarbonate": 31, "Magnesium": 32,
            "GGT": 34, "Albumin": 35, "AST": 37,
            "ALT": 38, "Glucose": 39, "HbA1C": 40,
            "WBC": 42, "Neutrophil Count": 43, "Basophil Count": 44, "Eosinophil Count": 45,
            "Lymphocyte Count": 46, "Platelets": 47, "MPV": 48, "Alk Phos": 50,
            "Phosphorus": 51, "Calcium": 52, "TSH": 55, "Estradiol": 57,
            "FSH": 57, "Anti-Mullerian hormone": 58, "Prolactin": 59, "Blood pressure (systolic pressure/diastolic pressure)": 62, "Body mass index": 64, "Waist measurement": 65
        }

    return health_data_variables

# Main function
def main():
    # Get gender input
    gender = get_gender()

    # Set sheet name based on gender
    sheet_name = "Template Sheet - Men" if gender == 'm' else "Template Sheet - Women"

    # Set gender-specific variables and mappings
    health_data_variables = set_gender_specific_data(gender)

    # Extract text from the PDF
    pdf_text = extract_text_from_pdf('lab.pdf')

    # Get the values for each health data variable from GPT
    response_text = get_values_from_gpt(pdf_text, health_data_variables)

    # Print the GPT response to the console
    print("OpenAI's Response:\n", response_text)

    # Parse GPT response into a dictionary
    extracted_data = {}
    for line in response_text.split('\n'):
        if ':' in line:
            variable, value = line.split(':', 1)
            variable = variable.strip()
            value = value.strip()
            if value and value.lower() != "not available in the provided lab results":
                print(f"Parsed: {variable} = {value}")
                extracted_data[variable] = value

    print("\nExtracted Data:")
    for var, value in extracted_data.items():
        print(f"{var}: {value}")

    # Write the extracted data to Google Sheets
    SHEET_ID = '1Nfe8arVgI4TZ5l6ZYVmFpn6vaBKMenTR6hbSZYOxD5o'
    write_to_google_sheets(extracted_data, SHEET_ID, sheet_name)

    print("Process completed.")
    print("Raw GPT response:\n", response_text)

if __name__ == "__main__":
    main()