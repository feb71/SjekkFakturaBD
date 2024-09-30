import fitz  # PyMuPDF

def extract_data_from_pdf(pdf_path):
    # Open the PDF file
    doc = fitz.open(pdf_path)
    extracted_data = []

    # Iterate over pages
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        text = page.get_text("blocks")
        
        # Process text blocks and filter rows with valid 7-digit VARENR
        for block in text:
            lines = block[4].splitlines()
            for line in lines:
                columns = line.split()
                # Check if the first column is a valid 7-digit number
                if len(columns) > 0 and columns[0].isdigit() and len(columns[0]) == 7:
                    # Extract the relevant columns for VARENR, description, etc.
                    varenr = columns[0]
                    description = " ".join(columns[1:-3])  # Assuming the description spans multiple columns
                    amount = columns[-3]
                    unit = columns[-2]
                    price = columns[-1]
                    
                    extracted_data.append([varenr, description, amount, unit, price])
    
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(extracted_data, columns=['VARENR', 'Beskrivelse', 'Antall', 'Enhet', 'Pris'])
    
    return df

# Usage example in Streamlit
uploaded_file = st.file_uploader("Last opp tilbuds-PDF", type="pdf")

if uploaded_file is not None:
    st.info("Leser data fra PDF...")
    offer_data = extract_data_from_pdf(uploaded_file)
    
    if not offer_data.empty:
        st.success("Data ble funnet og tolket.")
        st.dataframe(offer_data)
        offer_data.to_excel("tilbud_data.xlsx", index=False)
        st.success("Tilbudet er lagret som tilbud_data.xlsx")
        st.download_button(label="Last ned tilbudet som Excel", data=open("tilbud_data.xlsx", "rb"), file_name="tilbud_data.xlsx")
    else:
        st.warning("Ingen gyldige VARENR ble funnet i PDF-filen.")
