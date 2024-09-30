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
                            try:
                                # Juster for riktig plassering av kolonner for tilbud
                                if doc_type == "Tilbud":
                                    quantity = float(columns[-4].replace('.', '').replace(',', '.')) if columns[-4].replace('.', '').replace(',', '').isdigit() else columns[-4]
                                    unit_price = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                    total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
                                else:
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
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å splitte opp kolonnene riktig
def split_columns(dataframe, doc_type):
    if doc_type == "Tilbud":
        dataframe["Antall_Tilbud"] = dataframe["Beskrivelse"].apply(lambda x: re.findall(r'(\d+)$', x)[0] if re.findall(r'(\d+)$', x) else None)
        dataframe["Enhet_Tilbud"] = dataframe["Beskrivelse"].apply(lambda x: re.findall(r'(M|STK)$', x)[0] if re.findall(r'(M|STK)$', x) else None)
        dataframe["Beskrivelse_Tilbud"] = dataframe["Beskrivelse"].apply(lambda x: re.sub(r'(\d+\s)?(M|STK)?$', '', x).strip())
        dataframe["Enhetspris_Tilbud"] = dataframe["Antall"]
        dataframe.drop(columns=["Antall"], inplace=True)

    elif doc_type == "Faktura":
        dataframe["Antall_Faktura"] = dataframe["Beskrivelse"].apply(lambda x: re.findall(r'(\d+)$', x)[0] if re.findall(r'(\d+)$', x) else None)
        dataframe["Beskrivelse_Faktura"] = dataframe["Beskrivelse"].apply(lambda x: re.sub(r'\d+$', '', x).strip())

    return dataframe

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

            # Splitte opp kolonnene
            offer_data = split_columns(offer_data, "Tilbud")
            invoice_data = split_columns(invoice_data, "Faktura")

            # Legg til suffikser for å unngå konflikter under sammenligning
            offer_data = offer_data.add_suffix('_Tilbud')
            invoice_data = invoice_data.add_suffix('_Faktura')

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
                merged_data = pd.merge(offer_data, invoice_data, left_on="Varenummer_Tilbud", right_on="Varenummer_Faktura", suffixes=('_Tilbud', '_Faktura'))

                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(merged_data)

                # Lagre kun artikkeldataene til XLSX
                excel_data = convert_df_to_excel(merged_data)

                st.success("Varenummer er lagret som Excel-fil.")
                
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=excel_data,
                    file_name="avvik_rapport.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("Kunne ikke lese tilbudsdata fra PDF-filen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
