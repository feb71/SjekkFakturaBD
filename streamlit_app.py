import streamlit as st
import pandas as pd
import fitz  # PyMuPDF library for reading PDFs
from io import BytesIO

# Function to extract data from a PDF file
def extract_pdf_data(file):
    text = ""
    pdf_document = fitz.open(stream=file.read(), filetype="pdf")
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    
    pdf_document.close()
    return text

# Function to parse the text into a structured DataFrame for both invoice and offer
def parse_invoice_data(text):
    lines = text.split('\n')
    data = []
    
    for line in lines:
        # Here, adjust the parsing logic to match the structure of the invoice
        if line.strip() and len(line.split()) > 4:  # Adjust condition as per your invoice structure
            parts = line.split()
            varenr = parts[0]
            description = " ".join(parts[1:-3])  # Adjust indices accordingly
            quantity = parts[-3]
            unit = parts[-2]
            price = parts[-1]
            data.append([varenr, description, quantity, unit, price])
    
    return pd.DataFrame(data, columns=["VARENR", "Beskrivelse", "Antall", "Enhet", "Pris"])

def parse_offer_data(text):
    lines = text.split('\n')
    data = []
    
    for line in lines:
        # Adjust parsing logic to match the structure of the offer
        if line.strip() and len(line.split()) > 4:  # Adjust condition as per your offer structure
            parts = line.split()
            varenr = parts[0]
            description = " ".join(parts[1:-3])  # Adjust indices accordingly
            quantity = parts[-3]
            unit = parts[-2]
            price = parts[-1]
            data.append([varenr, description, quantity, unit, price])
    
    return pd.DataFrame(data, columns=["VARENR", "Beskrivelse", "Antall", "Enhet", "Pris"])

# Main Streamlit application
st.title("Sammenlign faktura og tilbud")

# Upload section for invoice and offer
invoice_file = st.file_uploader("Last opp faktura (PDF)", type="pdf")
offer_file = st.file_uploader("Last opp tilbud (PDF)", type="pdf")

if invoice_file is not None and offer_file is not None:
    # Process and display invoice data
    st.subheader("Fakturadata:")
    invoice_text = extract_pdf_data(invoice_file)
    invoice_data = parse_invoice_data(invoice_text)
    st.dataframe(invoice_data)

    # Process and display offer data
    st.subheader("Tilbudsdata:")
    offer_text = extract_pdf_data(offer_file)
    offer_data = parse_offer_data(offer_text)
    st.dataframe(offer_data)

    # Comparison logic
    comparison = pd.merge(invoice_data, offer_data, on='VARENR', how='outer', suffixes=('_Faktura', '_Tilbud'))
    
    # Calculate discrepancies in price
    comparison['Avvik_Enhetspris'] = comparison['Pris_Faktura'].astype(float) - comparison['Pris_Tilbud'].astype(float)
    comparison['Prosent_avvik_pris'] = (comparison['Avvik_Enhetspris'] / comparison['Pris_Tilbud'].astype(float)) * 100

    st.subheader("Avviksrapport:")
    st.dataframe(comparison)

    # Export comparison to Excel
    st.download_button(
        label="Last ned avviksrapporten som Excel",
        data=comparison.to_excel(index=False),
        file_name="avviksrapport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
