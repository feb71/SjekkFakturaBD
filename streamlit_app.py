import streamlit as st
import pandas as pd
import pdfplumber

def extract_data_from_pdf(pdf_file, source_type):
    data_rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        # Assuming the first column contains "VARENR" when it is not empty or None
                        if row and len(row) > 0 and row[0]:
                            # Extracting relevant columns as per the expected layout
                            data_rows.append(row)
    return data_rows

def process_invoice_data(data_rows):
    # Extracting columns from the data rows into a DataFrame
    columns = ['Varenummer', 'Beskrivelse', 'Mengde', 'Enhet', 'Enhetspris', 'Totalpris']
    data = pd.DataFrame(data_rows, columns=columns)
    
    # Ensure the relevant columns are converted to numeric where applicable
    data['Varenummer'] = pd.to_numeric(data['Varenummer'], errors='coerce')
    data['Mengde'] = pd.to_numeric(data['Mengde'], errors='coerce')
    data['Enhetspris'] = pd.to_numeric(data['Enhetspris'], errors='coerce')
    data['Totalpris'] = pd.to_numeric(data['Totalpris'], errors='coerce')
    
    # Removing rows where 'Varenummer' might be missing or not relevant
    data = data[data['Varenummer'].notna()]
    
    return data

def main():
    st.title("Sammenlign faktura og tilbud")
    
    # Upload PDF files
    invoice_file = st.file_uploader("Last opp faktura (PDF)", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud (PDF)", type="pdf")
    
    if invoice_file:
        st.info("Leser faktura...")
        invoice_data_rows = extract_data_from_pdf(invoice_file, 'faktura')
        if invoice_data_rows:
            invoice_data = process_invoice_data(invoice_data_rows)
            st.success("Fakturadata lest:")
            st.write(invoice_data)
        else:
            st.error("Kunne ikke lese data fra faktura PDF.")
    
    if offer_file:
        st.info("Leser tilbud...")
        offer_data_rows = extract_data_from_pdf(offer_file, 'tilbud')
        if offer_data_rows:
            offer_data = process_invoice_data(offer_data_rows)
            st.success("Tilbudsdata lest:")
            st.write(offer_data)
        else:
            st.error("Kunne ikke lese data fra tilbud PDF.")

    # Comparison logic
    if invoice_file and offer_file and not invoice_data.empty and not offer_data.empty:
        st.info("Sammenligner data...")
        
        merged_data = pd.merge(invoice_data, offer_data, how="left", left_on="Varenummer", right_on="Varenummer", suffixes=('_Faktura', '_Tilbud'))
        
        # Calculate deviations
        merged_data['Avvik_Antall'] = merged_data['Mengde_Faktura'] - merged_data['Mengde_Tilbud']
        merged_data['Avvik_Enhetspris'] = merged_data['Enhetspris_Faktura'] - merged_data['Enhetspris_Tilbud']
        merged_data['Prosent_avvik_pris'] = (merged_data['Avvik_Enhetspris'] / merged_data['Enhetspris_Tilbud']) * 100
        
        # Displaying only relevant columns
        deviation_report = merged_data[['Varenummer', 'Beskrivelse_Faktura', 'Mengde_Faktura', 'Enhetspris_Faktura', 'Totalpris_Faktura',
                                        'Mengde_Tilbud', 'Enhetspris_Tilbud', 'Totalpris_Tilbud', 'Avvik_Antall', 'Avvik_Enhetspris', 'Prosent_avvik_pris']]
        
        st.success("Avviksrapport:")
        st.write(deviation_report)
    else:
        st.warning("Last opp både faktura og tilbud for å sammenligne.")

if __name__ == "__main__":
    main()
