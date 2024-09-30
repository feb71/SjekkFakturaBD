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

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                st.write(f"Leser side {page.page_number} med {len(lines)} linjer.")
                
                for i, line in enumerate(lines):
                    st.write(f"Leste linje: {line}")

                    # Sjekker om vi starter å lese fra VARENR
                    if doc_type == "Tilbud" and "VARENR" in line:
                        start_reading = True
                        continue
                    elif doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        # Sjekk om denne linjen kan være starten på en ny "VARENR"
                        if len(columns) > 0 and columns[0].startswith('■'):
                            if len(lines) > i + 1:
                                next_line = lines[i + 1].split()
                                if len(next_line) > 0 and next_line[0].isdigit():
                                    item_number = next_line[0]  # Bruker neste linjes første verdi som VARENR
                                    
                                    # Sjekk om beskrivelsen, mengde, og priser finnes i denne linjen eller neste
                                    description = " ".join(next_line[1:-3])  # Justert for å fange beskrivelsen riktig
                                    try:
                                        quantity = float(next_line[-3].replace('.', '').replace(',', '.'))
                                        unit_price = float(next_line[-2].replace('.', '').replace(',', '.'))
                                        total_price = float(next_line[-1].replace('.', '').replace(',', '.'))
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
