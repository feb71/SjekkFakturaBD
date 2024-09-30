import streamlit as st
import pandas as pd
import tabula

# Funksjon for å sjekke om en verdi er numerisk
def is_numeric(value):
    try:
        float(value)
        return True
    except ValueError:
        return False

# Funksjon for å lese data fra PDF
def read_pdf_data(pdf_file, is_offer=False):
    try:
        df_list = tabula.read_pdf(pdf_file, pages='all', multiple_tables=True, pandas_options={'header': None})
        data = pd.concat(df_list, ignore_index=True)
        
        # Gi passende kolonnenavn basert på om det er faktura eller tilbud
        if is_offer:
            data.columns = ["VARENR", "BESKRIVELSE", "MENGDE", "ENHET", "Enhetspris_tilbud", "Beløp Materiell"]
        else:
            data.columns = ["Linje", "VARENR", "BESKRIVELSE", "MENGDE", "ENHET", "Salgspris", "Rab. %", "Beløp"]
        
        # Fjern unødvendige rader som ikke inneholder numeriske VARENR
        data = data[data['VARENR'].apply(is_numeric)]
        
        # Konverter VARENR til string for å sikre riktig behandling
        data['VARENR'] = data['VARENR'].astype(str)
        return data
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {str(e)}")
        return pd.DataFrame()  # Returner en tom DataFrame hvis noe går galt

# Last opp faktura og tilbud
st.title("Sammenlign faktura og tilbud")
faktura_file = st.file_uploader("Last opp faktura (PDF)", type="pdf")
tilbud_file = st.file_uploader("Last opp tilbud (PDF)", type="pdf")

if faktura_file and tilbud_file:
    # Les fakturadata
    faktura_data = read_pdf_data(faktura_file, is_offer=False)
    st.write("Fakturadata:")
    st.dataframe(faktura_data)
    
    # Les tilbudsdata
    tilbud_data = read_pdf_data(tilbud_file, is_offer=True)
    st.write("Tilbudsdata:")
    st.dataframe(tilbud_data)
    
    if not faktura_data.empty and not tilbud_data.empty:
        # Sammenlign tilbud og faktura basert på VARENR
        merged_data = pd.merge(faktura_data, tilbud_data, on='VARENR', how='outer', suffixes=('_Faktura', '_Tilbud'))

        # Beregn avvik
        merged_data['Avvik_Antall'] = merged_data['MENGDE_Faktura'] - merged_data['MENGDE_Tilbud']
        merged_data['Avvik_Enhetspris'] = merged_data['Salgspris'] - merged_data['Enhetspris_tilbud']
        merged_data['Prosent_avvik_pris'] = ((merged_data['Salgspris'] - merged_data['Enhetspris_tilbud']) / merged_data['Enhetspris_tilbud']) * 100
        
        # Fyll inn NaN-verdier med 0 for enklere lesing
        merged_data.fillna(0, inplace=True)

        # Vis avviksrapport
        st.write("Avviksrapport:")
        st.dataframe(merged_data)

        # Last ned-knapp for avviksrapport
        @st.cache_data
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Avviksrapport')
            processed_data = output.getvalue()
            return processed_data

        excel_data = convert_df_to_excel(merged_data)
        st.download_button(
            label="Last ned avviksrapport i Excel-format",
            data=excel_data,
            file_name='avviksrapport.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("Kunne ikke lese data fra PDF-filene. Sjekk om filene er riktige og prøv igjen.")
