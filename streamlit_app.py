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
            columns = line.split()
            
            # Sjekker om linjen inneholder nok kolonner
            if len(columns) >= 5:
                varenr = columns[0]
                
                # Sjekk om varenr er et gyldig tall og at det ikke inneholder punktum (for å unngå ting som 43.21.1)
                if varenr.isdigit():
                    # Hent ut de relevante kolonnene basert på antall kolonner i linjen
                    beskrivelse = " ".join(columns[1:-3])
                    antall = columns[-3]
                    enhet = columns[-2]
                    pris = columns[-1]
                    
                    # Legg til i dataen
                    data_rows.append([varenr, beskrivelse, antall, enhet, pris])

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
