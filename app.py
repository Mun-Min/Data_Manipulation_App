import streamlit as st 
import pandas as pd
import io
import base64

st.markdown('# Data Manipulation App')

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

# Display uploaded file
if uploaded_file:
    st.write("Uploaded file:")
    st.write(uploaded_file.name)

    # Load Excel file
    sheet_names = pd.read_excel(uploaded_file, sheet_name=None).keys()
    selected_sheet = st.selectbox("Select a sheet", sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    # Display data
    st.write("Data:")
    st.write(df)

    # Perform data manipulation
    columns = list(df.columns)
    round_column = st.selectbox("Select a column to round", columns)
    decimal_places = st.number_input("Enter the number of decimal places to round to", min_value=0, max_value=10)
    
    try:
        rounded_df = df.round({round_column: decimal_places})
    except TypeError:
        st.error("No columns can be rounded.")
        st.stop()
    
    st.write("Rounded Data:")
    st.write(rounded_df)

    # Export data to new Excel file
    output_file = st.text_input("Enter output file name", "output")
    if st.button("Export to Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            rounded_df.to_excel(writer, sheet_name=selected_sheet, index=False)

        output.seek(0)
        b64 = base64.b64encode(output.read()).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{output_file}.xlsx">Download file</a>'
        download_button = st.markdown(href, unsafe_allow_html=True)
        st.write("Exported data to Excel file:", output_file + ".xlsx")
        st.write("Please check the download folder and open the file in Excel.")
