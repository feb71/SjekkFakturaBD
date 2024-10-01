import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Function to read the invoice number from the PDF
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
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Function to extract data from PDF file
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False  # Control variable to start data collection

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    # Start collecting data after finding "Artikkel" or "VARENR" based on document type
                    if doc_type == "Tilbud" and "VARENR" in line:
                        start_reading = True
                        continue  # Skip the line containing "VARENR" to the next line
                    elif doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue  # Skip the line containing "Artikkel" to the next line

                    if start_reading:
                        columns = line.split()
                        if doc_type == "Faktura" and len(columns) >= 5:
                            item_number = columns[1] 
                            if not item_number.isdigit():
                                continue  # Skip lines where the item is not a valid item number
                            
                            description = " ".join(columns[2:-3])  
                            try:
                                quantity = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else columns[-2]
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
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
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Function to process description columns for Faktura data
def process_description_faktura(df):
    # Extract the quantity (Antall) which is typically the last number in the description
    df['Antall_Faktura'] = df['Beskrivelse_Faktura'].str.extract(r'(\d+)$').astype(float)
    # Remove the extracted quantity from the description
    df['Beskrivelse_Faktura'] = df['Beskrivelse_Faktura'].str.replace(r'\s*\d+$', '', regex=True)
    
    # Extract the unit (Enhet), which is usually 'M', 'M2', 'STK', etc., appearing towards the end of the description
    df['Enhet'] = df['Beskrivelse_Faktura'].str.extract(r'(\bM2|\bM|\bSTK\b)$', expand=False)
    # Remove the extracted unit from the description
    df['Beskrivelse_Faktura'] = df['Beskrivelse_Faktura'].str.replace(r'\s*\b(M2|M|STK)\b$', '', regex=True)
    
    return df

# Function to read the offer data from Excel
def read_offer_from_excel(file):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Kunne ikke lese Excel-filen: {e}")
        return pd.DataFrame()

# Function to convert DataFrame to an Excel file
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Main function for the Streamlit app
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Upload section
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Get invoice number
        st.info("Henter fakturanummer fra faktura...")
        invoice_number = get_invoice_number(invoice_file)
        
        if invoice_number:
            st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Extract data from PDF and Excel files
            st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
            st.info("Laster inn tilbud fra Excel...")
            offer_data = read_offer_from_excel(offer_file)

            # Process the description in the Faktura data
            if not invoice_data.empty:
                invoice_data = process_description_faktura(invoice_data)

            if not offer_data.empty:
                # Save the offer data to an Excel file for download
                offer_excel_data = convert_df_to_excel(offer_data)
                
                st.download_button(
                    label="Last ned tilbudet som Excel",
                    data=offer_excel_data,
                    file_name="tilbud_data.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                # Compare invoice with offer
                st.write("Sammenligner data...")
                merged_data = pd.merge(offer_data, invoice_data, left_on="VARENR", right_on="Varenummer", suffixes=('_Tilbud', '_Faktura'))

                # Convert columns to numeric
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["ANTALL"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["ENHETSPRIS"], errors='coerce')

                # Find discrepancies
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                    (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

                # Save invoice data to XLSX
                all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Totalt pris"]]
                excel_data = convert_df_to_excel(all_items)

                st.success("Varenummer er lagret som Excel-fil.")
                
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=convert_df_to_excel(avvik),
                    file_name="avvik_rapport.xlsx"
                )
                
                st.download_button(
                    label="Last ned alle varenummer som Excel",
                    data=excel_data,
                    file_name="faktura_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("Kunne ikke lese tilbudsdata fra Excel-filen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
