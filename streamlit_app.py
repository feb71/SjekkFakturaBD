import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Function to read invoice number from PDF
def get_invoice_number(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Could not read invoice number from PDF: {e}")
        return None

# Function to extract relevant data from PDF file
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False  # Control variable to start data collection

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"No text found on page {page.page_number} in the PDF file.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    if doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 5:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue  # Skip rows without a valid item number
                            
                            description = " ".join(columns[2:-3])
                            try:
                                quantity = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else columns[-2]
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
                            except ValueError as e:
                                st.error(f"Could not convert to float: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhetspris_Faktura": unit_price,
                                "Totalt pris": total_price,
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("No data found in the PDF file.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Could not read data from PDF: {e}")
        return pd.DataFrame()

# Function to split the description and move 'M', 'M2', and 'STK' to the 'Enhet' column
def split_description(data, doc_type):
    if doc_type == "Faktura":
        # Extract the unit 'M', 'M2', or 'STK' from 'Beskrivelse_Faktura'
        data['Enhet'] = data['Beskrivelse_Faktura'].str.extract(r'(M2|M|STK)$', expand=False)
        data['Beskrivelse_Faktura'] = data['Beskrivelse_Faktura'].str.replace(r'\b(M2|M|STK)$', '', regex=True).str.strip()
        
        # Extract the 'Antall_Faktura' from the remaining description
        data['Antall_Faktura'] = data['Beskrivelse_Faktura'].str.extract(r'(\d+)$', expand=False).astype(float)
        data['Beskrivelse_Faktura'] = data['Beskrivelse_Faktura'].str.replace(r'\s*\d+$', '', regex=True).str.strip()
    
    return data

# Function to convert DataFrame to an Excel file
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Main function for the Streamlit app
def main():
    st.title("Compare Invoice with Offer")

    # File upload section
    invoice_file = st.file_uploader("Upload invoice from Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Upload offer from Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Retrieve invoice number
        st.info("Retrieving invoice number from the invoice...")
        invoice_number = get_invoice_number(invoice_file)
        
        if invoice_number:
            st.success(f"Invoice number found: {invoice_number}")
            
            # Extract data from the invoice PDF
            st.info("Loading invoice...")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Split the description to separate 'Antall_Faktura' and 'Enhet'
            if not invoice_data.empty:
                invoice_data = split_description(invoice_data, "Faktura")

            # Load offer data from the Excel file
            st.info("Loading offer...")
            offer_data = pd.read_excel(offer_file)

            if not offer_data.empty:
                # Ensure column names match your existing data structure
                offer_data.rename(columns={"VARENR": "Varenummer", "BESKRIVELSE": "Beskrivelse_Tilbud",
                                           "ANTALL": "Antall_Tilbud", "ENHET": "Enhet",
                                           "ENHETSPRIS": "Enhetspris_Tilbud", "TOTALPRIS": "Totalt pris"}, inplace=True)

                # Save the offer data as Excel file
                offer_excel_data = convert_df_to_excel(offer_data)
                st.download_button(
                    label="Download the offer as Excel",
                    data=offer_excel_data,
                    file_name="offer_data.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                # Compare invoice and offer data
                st.write("Comparing data...")
                merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", suffixes=('_Tilbud', '_Faktura'))

                # Convert columns to numeric
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Identify discrepancies
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                discrepancies = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                            (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                st.subheader("Discrepancies between Invoice and Offer")
                st.dataframe(discrepancies)

                # Save only item data to XLSX
                excel_data = convert_df_to_excel(invoice_data)
                
                st.success("Item numbers saved as Excel file.")
                
                st.download_button(
                    label="Download discrepancy report as Excel",
                    data=convert_df_to_excel(discrepancies),
                    file_name="discrepancy_report.xlsx"
                )
                
                st.download_button(
                    label="Download all item numbers as Excel",
                    data=excel_data,
                    file_name="invoice_items.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("Could not read offer data from the Excel file.")
        else:
            st.error("Invoice number was not found in the PDF file.")

if __name__ == "__main__":
    main()
