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
            start_reading = False

            st.write(f"Antall sider i PDF-filen: {len(pdf.pages)}")

            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()

                st.write(f"Råtekst fra side {page_number}:")
                st.write(text)

                if text is None:
                    st.error(f"Ingen tekst funnet på side {page_number} i PDF-filen.")
                    continue

                lines = text.split('\n')
                st.write(f"Leser side {page_number} med {len(lines)} linjer.")

                for i, line in enumerate(lines):
                    st.write(f"Leste linje: {line}")

                    # Startbetingelse for tilbud
                    if doc_type == "Tilbud" and "VARENR" in line:
                        start_reading = True
                        st.info(f"Starter å lese tilbud fra linje {i} på side {page_number}.")
                        continue
                    # Startbetingelse for faktura
                    elif doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        st.info(f"Starter å lese faktura fra linje {i} på side {page_number}.")
                        continue

                    if start_reading:
                        columns = line.split()
                        
                        # Sjekk om første kolonne er et tall (altså VARENR)
                        if len(columns) > 0 and columns[0].isdigit():
                            varenummer = columns[0]
                            beskrivelse = " ".join(columns[1:])
                            
                            # Hent andre kolonneverdier (f.eks. antall, enhetspris osv.) avhengig av layout
                            if len(columns) >= 5:
                                antall = columns[-4]   # Juster basert på korrekt kolonneplassering
                                enhetspris = columns[-3]
                                totalpris = columns[-1]
                            else:
                                antall = ""
                                enhetspris = ""
                                totalpris = ""

                            data.append([invoice_number, varenummer, beskrivelse, antall, enhetspris, totalpris])

            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen. Sjekk om layouten eller teksten blir riktig tolket.")
            else:
                st.success("Data ble funnet og tolket.")
                
            return data

    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF-filen: {e}")
        return None


# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


# Assuming you have extracted the offer data into a DataFrame named `offer_data`
if not offer_data.empty:
    st.write("Tabell lest fra tilbudet:")
    st.dataframe(offer_data)
else:
    st.write("Ingen data ble funnet i tilbudet.")


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

            # Feilsøking: Skriv ut kolonnenavnene
            st.write("Kolonner fra faktura:", invoice_data.columns)
            st.write("Kolonner fra tilbud:", offer_data.columns)

            if not offer_data.empty:
                # Lagre tilbudet som Excel-fil
                offer_excel_data = convert_df_to_excel(offer_data)
                
                st.download_button(
                    label="Last ned tilbudet som Excel",
                    data=offer_excel_data,
                    file_name="tilbud_data.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                # Sjekk kolonnenavn etter innlesing
                st.write("Kolonner fra faktura:", invoice_data.columns)
                st.write("Kolonner fra tilbud:", offer_data.columns)

                st.write("Antall rader i tilbudsdata:", len(offer_data))
                st.write("Antall rader i fakturadata:", len(invoice_data))


                # Sammenligne faktura mot tilbud
                st.write("Sammenligner data...")
                merged_data = pd.merge(invoice_data, offer_data, how='left', left_on="Varenummer", right_on="Varenummer", suffixes=('_Faktura', '_Tilbud'))


                # Konverter kolonner til numerisk
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Finne avvik
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                
                # Beregne prosentforskjellen
                merged_data["Prosent_avvik_pris"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

                # Filtrer kun rader med avvik
                avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                    (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0)) |
                                    (merged_data["Prosent_avvik_pris"].notna() & (merged_data["Prosent_avvik_pris"] != 0))]

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
