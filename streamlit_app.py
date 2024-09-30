import streamlit as st
import pandas as pd
import fitz  # PyMuPDF library for reading PDFs
from io import BytesIO

# Function to extract table data from a PDF
def extract_pdf_data(file_path):
    with fitz.open(file_path) as pdf:
        text = ""
        for page_num in range(pdf.page_count):
            page = pdf.load_page(page_num)
            text += page.get_text()
    return text

# Function to clean and parse the extracted text into a DataFrame
def parse_pdf_data(text):
    lines = text.split("\n")
    data = []
    for line in lines:
        # Assuming the data has specific characteristics to identify relevant lines
        # Adjust these parsing rules based on your actual data structure
        if line.strip().isdigit():  # Check if line is just digits, assuming it’s VARENR
            continue
        columns = line.split()
        if len(columns) > 4:  # Adjust this number based on actual column expectations
            data.append(columns)
    return pd.DataFrame(data, columns=["VARENR", "Beskrivelse", "Antall", "Enhet", "Pris"])

# Main Streamlit application
st.title("Sammenlign faktura og tilbud")

# Upload section for the invoice PDF
invoice_file = st.file_uploader("Last opp faktura (PDF)", type="pdf")
offer_file = st.file_uploader("Last opp tilbud (PDF)", type="pdf")

# Check if both files are uploaded
if invoice_file and offer_file:
    # Extract and display invoice data
    st.subheader("Fakturadata:")
    invoice_text = extract_pdf_data(invoice_file)
    invoice_data = parse_pdf_data(invoice_text)
    st.dataframe(invoice_data)

    # Extract and display offer data
    st.subheader("Tilbudsdata:")
    offer_text = extract_pdf_data(offer_file)
    offer_data = parse_pdf_data(offer_text)
    st.dataframe(offer_data)

    # Compare the invoice data with the offer data
    comparison_result = pd.merge(
        invoice_data,
        offer_data,
        how="left",
        left_on="VARENR",
        right_on="VARENR",
        suffixes=("_Faktura", "_Tilbud")
    )

    # Calculate percentage difference between the invoice and offer prices
    comparison_result["Prosent_avvik_pris"] = (
        (comparison_result["Pris_Faktura"].astype(float) - comparison_result["Pris_Tilbud"].astype(float))
        / comparison_result["Pris_Tilbud"].astype(float)
    ) * 100

    # Display the comparison result
    st.subheader("Sammenligning av faktura og tilbud")
    st.dataframe(comparison_result)

    # Provide a download button for the comparison result as an Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        comparison_result.to_excel(writer, index=False, sheet_name='Sammenligning')
        writer.save()
        st.download_button(
            label="Last ned sammenligningsresultat til Excel",
            data=output.getvalue(),
            file_name="sammenligning_resultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Vennligst last opp både fakturaen og tilbudet for å sammenligne.")
