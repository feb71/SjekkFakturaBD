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

# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
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
                    # Start innsamlingen etter å ha funnet "Artikkel" eller "VARENR" basert på dokumenttypen
                    if doc_type == "Tilbud" and "VARENR" in line:
                        start_reading = True
                        continue  # Hopp over linjen som inneholder "VARENR" til neste linje
                    elif doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue  # Hopp over linjen som inneholder "Artikkel" til neste linje

                    if start_reading:
                        columns = line.split()
                        if doc_type == "Tilbud" and len(columns) > 3 and columns[0].isdigit() and len(columns[0]) == 7:
                            item_number = columns[0]  
                            description = " ".join(columns[1:-3])
                            try:
                                quantity = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else columns[-2]
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            data.append({
                                "UnikID": item_number,
                                "Varenummer": item_number,
                                "Beskrivelse_Tilbud": description,
                                "Antall_Tilbud": quantity,
                                "Enhetspris_Tilbud": unit_price,
                                "Totalt pris": total_price,
                                "Type": doc_type
                            })

                        elif doc_type == "Faktura" and len(columns) >= 5:
                            item_number = columns[1] 
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

# Funksjon for å dele opp beskrivelse basert på siste elementer
def split_description(data, doc_type):
    if doc_type == "Tilbud":
        data['Antall_Tilbud'] = data['Beskrivelse_Tilbud'].str.extract(r'(\d+)$', expand=False).astype(float)
        data['Beskrivelse_Tilbud'] = data['Beskrivelse_Tilbud'].str.replace(r'\s*\d+$', '', regex=True)
        
        data['Enhet'] = data['Beskrivelse_Tilbud'].str.extract(r'(\b(?:M|M2|STK)\b)$', expand=False)
        data['Beskrivelse_Tilbud'] = data['Beskrivelse_Tilbud'].str.replace(r'\s*\b(?:M|M2|STK)\b$', '', regex=True)
    elif doc_type == "Faktura":
        data['Antall_Faktura'] = data['Beskrivelse_Faktura'].str.extract(r'(\d+)$', expand=False).astype(float)
        data['Beskrivelse_Faktura'] = data['Beskrivelse_Faktura'].str.replace(r'\s*\d+$', '', regex=True)
        
        data['Enhet'] = data['Beskrivelse_Faktura'].str.extract(r'(\b(?:M|M2|STK)\b)$', expand=False)
        data['Beskrivelse_Faktura'] = data['Beskrivelse_Faktura'].str.replace(r'\s*\b(?:M|M2|STK)\b$', '', regex=True)
    
    return data

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
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl", type="pdf")

    if invoice_file and offer_file:
        # Hent fakturanummer
        st.info("Henter fakturanummer fra faktura...")
        invoice_number = get_invoice_number(invoice_file)
        
        if invoice_number:
            st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filer
            st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
            st.info("Laster inn tilbud...")
            offer_data = extract_data_from_pdf(offer_file, "Tilbud")

            # Del opp beskrivelsen
            if not invoice_data.empty:
                invoice_data = split_description(invoice_data, "Faktura")
            if not offer_data.empty:
                offer_data = split_description(offer_data, "Tilbud")

            if not offer_data.empty:
                offer_excel_data = convert_df_to_excel(offer_data)
                st.download_button(
                    label="Last ned tilbudet som Excel",
                    data=offer_excel_data,
                    file_name="tilbud_data.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

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

                # Lagre kun artikkeldataene til XLSX
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
                st.error("Kunne ikke lese tilbudsdata fra PDF-filen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
