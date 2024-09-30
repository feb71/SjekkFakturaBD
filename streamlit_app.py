import pandas as pd
import streamlit as st
import fitz  # PyMuPDF library

def extract_pdf_table(pdf_path):
    doc = fitz.open(pdf_path)
    data = []

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        lines = text.split('\n')

        for line in lines:
            # We assume that the line contains data if it has more than a threshold number of words or certain keywords
            parts = line.split()
            if len(parts) > 1 and parts[0].isdigit():
                # Assuming VARENR is a purely numeric value
                varenr = next((part for part in parts if part.isdigit()), None)
                
                # Extract other relevant data from the parts of the line
                if varenr:
                    try:
                        amount_index = parts.index(varenr) + 1
                        description = " ".join(parts[amount_index:-3])  # Assuming the last 3 columns are quantity, unit, and price
                        quantity = parts[-3]
                        unit = parts[-2]
                        price = parts[-1]
                        
                        data.append([varenr, description, quantity, unit, price])
                    except IndexError:
                        pass  # Skip lines that don't match the expected structure

    doc.close()
    columns = ["VARENR", "Beskrivelse", "Antall", "Enhet", "Pris"]
    return pd.DataFrame(data, columns=columns)

def main():
    st.title("PDF Data Extractor")

    uploaded_file = st.file_uploader("Last opp tilbuds-PDF", type="pdf")

    if uploaded_file is not None:
        with open("uploaded_offer.pdf", "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.write("Leser data fra PDF...")

        try:
            offer_data = extract_pdf_table("uploaded_offer.pdf")

            if not offer_data.empty:
                st.write("Data funnet i tilbudsfilen:")
                st.dataframe(offer_data)

                # You can implement additional processing or comparison logic here
                # For example, save to Excel
                offer_data.to_excel("tilbud_data.xlsx", index=False)
                st.success("Tilbudet er lagret som tilbud_data.xlsx")
                st.download_button(
                    label="Last ned tilbudet som Excel",
                    data=open("tilbud_data.xlsx", "rb"),
                    file_name="tilbud_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("Ingen gyldige VARENR-data funnet i PDF-en.")
        except Exception as e:
            st.error(f"Det oppstod en feil under behandling: {e}")

if __name__ == "__main__":
    main()
