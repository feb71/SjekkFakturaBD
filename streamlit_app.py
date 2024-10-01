import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Funksjon for å lese fakturanummer fra PDF
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

# Funksjon for å lese PDF-filen og hente ut relevante data fra faktura
def extract_data_from_pdf(file, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False  # Kontrollvariabel for å starte innsamlingen av data

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    # Start innsamlingen etter å ha funnet "Artikkel"
                    if "Artikkel" in line:
                        start_reading = True
                        continue  # Hopp over linjen som inneholder "Artikkel" til neste linje

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 5:  # Forventer at vi har nok kolonner i linjen
                            item_number = columns[1]  # Henter artikkelnummeret fra riktig kolonne (andre kolonne)
                            if not item_number.isdigit():
                                continue  # Skipper linjer der elementet ikke er et gyldig artikkelnummer
                            
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
                                "Beskrivelse": description,
                                "Antall": quantity,
                                "Enhetspris": unit_price,
                                "Totalt pris": total_price,
                                "Type": "Faktura"
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplastingsseksjon
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        st.info("Henter fakturanummer fra faktura...")
        invoice_number = get_invoice_number(invoice_file)
        
        if invoice_number:
            st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filen for faktura
            st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(invoice_file, invoice_number)

            # Lese tilbudet fra Excel-filen
            st.info("Laster inn tilbud fra Excel...")
            try:
                offer_data = pd.read_excel(offer_file)
            except Exception as e:
                st.error(f"Kunne ikke lese tilbudsdata fra Excel-filen: {e}")
                return

            if not offer_data.empty:
                # Sammenligne faktura mot tilbud
                st.write("Sammenligner data...")
                merged_data = pd.merge(offer_data, invoice_data, left_on="Varenummer", right_on="Varenummer", suffixes=('_Tilbud', '_Faktura'))

                # Konverter kolonner til numerisk
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Finne avvik
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                    (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

                # Lagre avviksrapporten som Excel-fil
                avvik_excel_data = convert_df_to_excel(avvik)
                
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=avvik_excel_data,
                    file_name="avvik_rapport.xlsx"
                )
                
                st.success("Sammenligningen er ferdig!")
            else:
                st.error("Tilbudsdataen er tom.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
