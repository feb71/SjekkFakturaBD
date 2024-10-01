import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Funksjon for å lese Excel-filen og hente ut relevante data
def extract_data_from_excel(file, doc_type):
    try:
        df = pd.read_excel(file)
        
        # Standardiser kolonnenavn i tilfelle de har forskjellig kapitalisering
        df.columns = df.columns.str.strip().str.lower()
        
        if doc_type == "Tilbud":
            df.rename(columns={
                'varenr': 'Varenummer',
                'beskrivelse': 'Beskrivelse_Tilbud',
                'antall': 'Antall_Tilbud',
                'enhet': 'Enhet_Tilbud',
                'enhetspris': 'Enhetspris_Tilbud',
                'totalpris': 'Totalt pris_Tilbud'
            }, inplace=True)
            
        elif doc_type == "Faktura":
            df.rename(columns={
                'varenr': 'Varenummer',
                'beskrivelse': 'Beskrivelse_Faktura',
                'antall': 'Antall_Faktura',
                'enhet': 'Enhet_Faktura',
                'enhetspris': 'Enhetspris_Faktura',
                'totalpris': 'Totalt pris_Faktura'
            }, inplace=True)
        
        return df
    except Exception as e:
        st.error(f"Kunne ikke lese data fra Excel: {e}")
        return pd.DataFrame()

# Funksjon for å behandle antall og enhet korrekt
def process_description_for_faktura(df):
    try:
        # Flytt verdien fra 'Antall_Faktura' til 'Enhet_Faktura' hvis det er en gyldig enhet
        df['Enhet_Faktura'] = df['Antall_Faktura'].apply(lambda x: x if isinstance(x, str) and x in ['M', 'M2', 'STK'] else None)
        
        # Hvis 'Enhet_Faktura' er satt, fjern verdien fra 'Antall_Faktura'
        df['Antall_Faktura'] = df.apply(lambda row: None if row['Enhet_Faktura'] else row['Antall_Faktura'], axis=1)

        # Splitt beskrivelsen for å flytte siste verdi inn i 'Antall_Faktura'
        df['Antall_Faktura'] = df.apply(lambda row: re.search(r'(\d+)$', row['Beskrivelse_Faktura']).group(1) if pd.isna(row['Antall_Faktura']) else row['Antall_Faktura'], axis=1)
        df['Antall_Faktura'] = pd.to_numeric(df['Antall_Faktura'], errors='coerce')
        
        # Fjern antallsverdien fra 'Beskrivelse_Faktura'
        df['Beskrivelse_Faktura'] = df['Beskrivelse_Faktura'].str.replace(r'\s*\d+$', '', regex=True)
        
        return df
    except Exception as e:
        st.error(f"Feil ved behandling av beskrivelse for faktura: {e}")
        return df

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplastingsseksjon
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl (Excel)", type="xlsx")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        st.info("Laster inn faktura...")
        invoice_data = extract_data_from_excel(invoice_file, "Faktura")

        st.info("Laster inn tilbud...")
        offer_data = extract_data_from_excel(offer_file, "Tilbud")

        # Behandle data
        invoice_data = process_description_for_faktura(invoice_data)

        if not offer_data.empty and not invoice_data.empty:
            # Sammenligne faktura mot tilbud
            st.write("Sammenligner data...")
            merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how="inner", suffixes=('_Tilbud', '_Faktura'))

            # Finne avvik
            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]

            st.subheader("Avvik mellom Faktura og Tilbud")
            st.dataframe(merged_data)

            # Konverter til Excel for nedlasting
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_data.to_excel(writer, index=False, sheet_name='Avvik')
            st.download_button(
                label="Last ned avviksrapport som Excel",
                data=output.getvalue(),
                file_name="avvik_rapport.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()
