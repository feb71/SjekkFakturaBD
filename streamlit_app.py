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
                        if len(columns) >= 5:  # Forventer at vi har nok kolonner i linjen
                            item_number = columns[1]  # Henter artikkelnummeret fra riktig kolonne (andre kolonne)
                            if not item_number.isdigit():
                                continue  # Skipper linjer der elementet ikke er et gyldig artikkelnummer
                            
                            description = " ".join(columns[2:-3])  # Justert for å fange beskrivelsen riktig

                            if doc_type == "Faktura":
                                # For faktura, les beskrivelse bakfra og del på siste mellomrom
                                split_desc = description.rsplit(' ', 1)
                                if len(split_desc) > 1:
                                    amount = split_desc[1]
                                    description = split_desc[0]
                                else:
                                    amount = columns[-3]  # Hvis ingen mellomrom, ta antall fra kolonnene
                                
                                try:
                                    quantity = float(amount.replace('.', '').replace(',', '.'))
                                    unit_price = float(columns[-2].replace('.', '').replace(',', '.'))
                                    total_price = float(columns[-1].replace('.', '').replace(',', '.'))
                                except ValueError as e:
                                    st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                    continue

                            else:
                                # Justeringer for tilbudet
                                try:
                                    quantity = float(columns[-4].replace('.', '').replace(',', '.')) if columns[-4].replace('.', '').replace(',', '').isdigit() else columns[-4]
                                    unit_price = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
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
                                "Type": doc_type
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

            if not offer_data.empty:
                # Lagre tilbudet som Excel-fil
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
                all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse", "Antall", "Enhetspris", "Totalt pris"]]
                
                # Konverter DataFrame til XLSX
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
