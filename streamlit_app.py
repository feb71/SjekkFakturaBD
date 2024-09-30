import streamlit as st
import pdfplumber
import pandas as pd

# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    with pdfplumber.open(file) as pdf:
        data = []
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                if "KABELRØR" in line or "FIBERDUK" in line:
                    columns = line.split()
                    if len(columns) >= 4:
                        item_number = columns[0]
                        description = " ".join(columns[1:-3])
                        quantity = float(columns[-3].replace(',', '.'))
                        unit_price = float(columns[-2].replace(',', '.'))
                        total_price = float(columns[-1].replace(',', '.'))
                        unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                        data.append({
                            "UnikID": unique_id,
                            "Fakturanummer": invoice_number,
                            "Varenummer": item_number,
                            "Beskrivelse": description,
                            "Antall": quantity,
                            "Enhetspris": unit_price,
                            "Totalt pris": total_price,
                            "Type": doc_type
                        })
        return pd.DataFrame(data)

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplastingsseksjon
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl", type="pdf")
    
    invoice_number = st.text_input("Skriv inn fakturanummer for unike ID-er")

    if invoice_file and offer_file and invoice_number:
        # Ekstraher data fra PDF-filer
        st.info("Laster inn faktura...")
        invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
        st.info("Laster inn tilbud...")
        offer_data = extract_data_from_pdf(offer_file, "Tilbud")

        # Lagre tilbudet som Excel-fil
        offer_data.to_excel("tilbud_data.xlsx", index=False)
        st.success("Tilbudet er lagret som tilbud_data.xlsx")

        # Sammenligne faktura mot tilbud
        st.write("Sammenligner data...")
        merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", suffixes=('_Tilbud', '_Faktura'))

        # Finne avvik
        merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
        merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
        avvik = merged_data[(merged_data["Avvik_Antall"] != 0) | (merged_data["Avvik_Enhetspris"] != 0)]

        st.subheader("Avvik mellom Faktura og Tilbud")
        st.dataframe(avvik)

        # Lagre alle varenummer til CSV
        all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse", "Antall", "Enhetspris", "Totalt pris"]]
        all_items.to_csv("faktura_varer.csv", index=False)

        st.success("Varenummer er lagret som faktura_varer.csv")
        st.download_button(
            label="Last ned avviksrapport som Excel",
            data=avvik.to_excel(index=False),
            file_name="avvik_rapport.xlsx"
        )
        
        st.download_button(
            label="Last ned alle varenummer som CSV",
            data=all_items.to_csv(index=False),
            file_name="faktura_varer.csv",
            mime='text/csv'
        )

if __name__ == "__main__":
    main()
