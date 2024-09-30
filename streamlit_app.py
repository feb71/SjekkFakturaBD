import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
import re

def extract_data_from_pdf(file):
    # Åpne PDF-filen
    doc = fitz.open(stream=file.read(), filetype="pdf")
    data_rows = []

    for page in doc:
        text = page.get_text("text")
        lines = text.split("\n")

        for line in lines:
            # Bruk regex for å finne linjer som inneholder rene tall for VARENR
            match = re.search(r'^\d+$', line.split()[0])
            if match:
                data_rows.append(line.split())

    return data_rows

def process_offer_data(offer_data_rows):
    columns = ["VARENR", "Beskrivelse", "Antall", "Enhet", "Pris"]
    data = pd.DataFrame(offer_data_rows, columns=columns)
    
    # Konverter VARENR til en ren numerisk kolonne og fjern ikke-numeriske verdier
    data['VARENR'] = pd.to_numeric(data['VARENR'], errors='coerce')
    data = data.dropna(subset=['VARENR'])
    data['VARENR'] = data['VARENR'].astype(int)
    
    return data

def main():
    st.title("PDF Data Extractor")

    uploaded_file = st.file_uploader("Last opp tilbuds-PDF", type="pdf")

    if uploaded_file is not None:
        st.info("Leser data fra PDF...")
        
        # Ekstraher data fra PDF
        offer_data_rows = extract_data_from_pdf(uploaded_file)
        
        if offer_data_rows:
            # Prosesser dataene
            offer_data = process_offer_data(offer_data_rows)
            
            if not offer_data.empty:
                st.success("Data ble funnet og tolket.")
                st.write("Data funnet i tilbudsfilen:")
                st.dataframe(offer_data)
            else:
                st.error("Ingen gyldige VARENR funnet i PDF-filen.")
        else:
            st.error("Kunne ikke lese data fra PDF-filen. Sjekk om filene er riktige og prøv igjen.")

if __name__ == "__main__":
    main()
