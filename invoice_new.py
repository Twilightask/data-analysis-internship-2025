import os
import re
import pandas as pd
import pdfplumber

PDF_FOLDER = r"C:\Users\Aayush\OneDrive\Desktop\Invoice folder"
OUTPUT_FILE = "invoices_data_fixed.xlsx"

def extract_invoice_data(pdf_path):
    data = {"File name": os.path.basename(pdf_path)}

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

        # Invoice no
        match = re.search(r"Invoice No\s+([A-Z0-9-]+)", text)
        data["Invoice no"] = match.group(1) if match else None

        # Invoice date
        match = re.search(r"Invoice Date\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})", text)
        data["Invoice date"] = match.group(1) if match else None

        # Price
        match = re.search(r"Total Amount\s+([\d,.]+)", text)
        data["Price"] = match.group(1) if match else None

        # Price + Taxes
        match = re.search(r"Total Due \(INR\)\s+([\d,.]+)", text)
        data["Price + Taxes"] = match.group(1) if match else None

                # --- Extract from the TABLE (Passenger, Details, Travel date) ---
                # --- Extract from the TABLE (Passenger, Details, Travel date) ---
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                # First row is headers; next rows are passenger details
                for row in table[1:]:
                    if not row:
                        continue
                    
                    # Safely extract with index checks
                    passenger = row[0].strip() if len(row) > 0 and row[0] else ""
                    airline   = row[3].strip() if len(row) > 3 and row[3] else ""
                    sector    = row[5].strip() if len(row) > 5 and row[5] else ""
                    flight    = row[6].strip() if len(row) > 6 and row[6] else ""
                    travel    = row[7].strip() if len(row) > 7 and row[7] else ""

                    data["Passenger name"] = passenger
                    data["Details"] = f"{airline} {sector} {flight}".strip()
                    data["Travel date"] = travel

                    return data  # keep only the first passenger for now



# Process all PDFs
all_data = []
for file in os.listdir(PDF_FOLDER):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(PDF_FOLDER, file)
        invoice_data = extract_invoice_data(pdf_path)
        all_data.append(invoice_data)

# Save to Excel
df = pd.DataFrame(all_data, columns=[
    "File name", "Invoice no", "Invoice date",
    "Passenger name", "Details", "Travel date",
    "Price", "Price + Taxes"
])
df.to_excel(OUTPUT_FILE, index=False)
print(f"Data extraction completed! Saved to {OUTPUT_FILE}")
