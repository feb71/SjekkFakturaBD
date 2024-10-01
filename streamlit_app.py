import streamlit as st
import pandas as pd
import pdfplumber
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

# Funksjon for å lese fakturadata fra PDF
def extract_invoice_data(file, invoice_number):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        columns = line.split()
                        if len(columns) >= 5:
                            item_number = columns[1]
                            if item_number.isdigit():
                                description = " ".join(columns[2:-3])
                                quantity = float(columns[-3].replace(',', '.'))
                                unit_price = float(columns[-2].replace(',', '.'))
                                total_price = float(columns[-1].replace(',', '.'))
                                
                                # Plassere riktig verdi i 'Enhet' og 'Antall'
                                enhet = ""
                                antall = quantity
                                
                                # Sjekker om 'Antall_Faktura' inneholder en enhet
                                if isinstance(quantity, float):
                                    if description.split()[-1].isalpha():
                                        enhet = description.split()[-1]
                                        description = " ".join(description.split()[:-1])
                                
                                unique_id = f"{invoice_number}_{item_number}"
                                data.append({
                                    "UnikID": unique_id,
                                    "Varenummer": item_number,
                                    "Beskrivelse_Faktura": description,
                                    "Antall_Faktura": antall,
                                    "Enhetspris_Faktura": unit_price,
                                    "Totalt pris": total_price,
                                    "Type": "Faktura",
                                    "Enhet": enhet
                                })
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese fakturadata fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å lese tilbudsdata fra Excel
def extract_offer_data(file):
    try:
        df = pd.read_excel(file)
        df.rename(columns=lambda x: x.strip(), inplace=True)  # Fjern eventuelle ledende og etterfølgende mellomrom
        if 'VARENR' in df.columns:
            df.rename(columns={'VARENR': 'Varenummer', 'BESKRIVELSE': 'Beskrivelse_Tilbud', 
                               'ANTALL': 'Antall_Tilbud', 'ENHET': 'Enhet_Tilbud',
                               'ENHETSPRIS': 'Enhetspris_Tilbud', 'TOTALPRIS': 'Totalt pris'}, inplace=True)
        return df
    except Exception as e:
        st.error(f"Kunne ikke lese tilbudsdata fra Excel: {e}")
        return pd.DataFrame()

# Funksjon for å konvertere DataFrame til Excel
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Ark1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplastingsseksjon
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        st.info("Henter fakturanummer fra faktura...")
        invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            st.success(f"Fakturanummer funnet: {invoice_number}")

            # Ekstraher data fra PDF og Excel
            st.info("Laster inn faktura...")
            invoice_data = extract_invoice_data(invoice_file, invoice_number)
            st.info("Laster inn tilbud...")
            offer_data = extract_offer_data(offer_file)

            if not offer_data.empty and not invoice_data.empty:
                # Sammenlign faktura mot tilbud
                st.info("Sammenligner data...")
                merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", suffixes=('_Tilbud', '_Faktura'))

                # Konverter kolonner til numerisk
                try:
                    merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                    merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                    merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')
                    merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')

                    # Finne avvik
                    merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                    merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                    avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                        (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                    st.subheader("Avvik mellom Faktura og Tilbud")
                    st.dataframe(avvik)

                    # Lagre avviksrapporten til Excel
                    st.download_button(
                        label="Last ned avviksrapport som Excel",
                        data=convert_df_to_excel(avvik),
                        file_name="avvik_rapport.xlsx"
                    )

                except KeyError as e:
                    st.error(f"En nødvendig kolonne mangler i dataene: {e}")
            else:
                st.error("Kunne ikke lese data fra tilbuds- eller fakturafilen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
