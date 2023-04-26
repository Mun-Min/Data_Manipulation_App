import streamlit as st
import pandas as pd
from streamlit_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode, JsCode

st.markdown('# Data Manipulation App')

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

# Display uploaded file
if uploaded_file:
    st.write("Uploaded file:")
    st.write(uploaded_file.name)

    # Load Excel file
    sheet_names, df = getGoogleSheet(uploaded_file)

    # Display data
    st.write("Data:")
    st.write(df)

    # Perform data manipulation
    columns = list(df.columns)
    sort_column = st.selectbox("Select a column to sort by", columns)
    sorted_df = df.sort_values(by=[sort_column])
    st.write("Sorted Data:")
    st.write(sorted_df)

    # Define grid options
    gb = GridOptionsBuilder.from_dataframe(sorted_df)
    gb.configure_default_column(groupable=True, value=True, enableRowGroup=True, aggFunc='sum', editable=True)
    gridOptions = gb.build()

    # Display editable table
    data_return_mode = 'AS_INPUT'
    grid_result = AgGrid(
        sorted_df, 
        gridOptions=gridOptions,
        data_return_mode=data_return_mode,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        width='100%',
        height='400px',
        allow_unsafe_jscode=True,
    )

    # Update dataframe with edited values
    if data_return_mode == 'AS_OUTPUT':
        edited_df = pd.DataFrame(grid_result['data'])
        st.write("Edited Data:")
        st.write(edited_df)

        # Export data to new Excel file
        output_file = st.text_input("Enter output file name", "output.xlsx")
        if st.button("Export to Excel"):
            with pd.ExcelWriter(output_file) as writer:
                edited_df.to_excel(writer, sheet_name=selected_sheet, index=False)

            st.write("Exported data to Excel file:", output_file)
    else:
        st.write("No data edited.")
