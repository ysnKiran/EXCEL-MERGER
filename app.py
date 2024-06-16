import streamlit as st
import pandas as pd

def merge_excel_files(df1, other_files, id_column, columns_to_merge):
    # Read all worksheets from the main Excel file
    xlsx1 = pd.read_excel(df1, sheet_name=None)
    print(f"Worksheets in the main file: {list(xlsx1.keys())}")

    # Initialize a dictionary to store the merged data for each worksheet
    merged_data_dict = {}

    # Loop through each worksheet in the main file
    for sheet_name, df1_sheet in xlsx1.items():
        print(f"Processing worksheet '{sheet_name}' from the main file")

        # Initialize the merged DataFrame for the current worksheet
        merged_df = df1_sheet.copy()

        # Loop through each uploaded file
        for uploaded_file in other_files:
            # Read the single worksheet from the uploaded file
            df2 = pd.read_excel(uploaded_file)
            print(f"Merging data from file: {uploaded_file.name}")

            # Merge the specified columns from df2 to df1_sheet based on the common ID column
            merged_df = pd.merge(merged_df, df2[[id_column] + columns_to_merge], on=id_column, how='left')

        # Store the merged data for the current worksheet in the dictionary
        merged_data_dict[sheet_name] = merged_df

    return merged_data_dict

# Set the app title
st.set_page_config(page_title="Excel Merger")

# Add a title
st.title("Excel Mergers")

# File uploader for the main Excel file
file_sheet1 = st.file_uploader("Choose the main Excel file", type=["xlsx"])

# File uploader for other Excel files
other_files = st.file_uploader("Choose other Excel files", type=["xlsx"], accept_multiple_files=True)

# Input for the ID column name
id_column = st.text_input("Enter the name of the ID column")

# Input for columns to merge
columns_to_merge = st.text_input("Enter the names of columns to merge (separated by commas)")

# Button to trigger the merging process
if st.button("Merge Excel Files"):
    if file_sheet1 is not None and other_files and id_column and columns_to_merge:
        # Convert the input string to a list
        columns_to_merge_list = [col.strip() for col in columns_to_merge.split(",")]

        # Call the merge function
        merged_data_dict = merge_excel_files(file_sheet1, other_files, id_column, columns_to_merge_list)

        # Display the merged data for each worksheet
        for sheet_name, merged_data in merged_data_dict.items():
            st.subheader(f"Merged Data - {sheet_name}")
            st.write(merged_data)

        # Option to download the merged data as an Excel file
        with pd.ExcelWriter("merged_data.xlsx") as writer:
            for sheet_name, merged_data in merged_data_dict.items():
                merged_data.to_excel(writer, sheet_name=sheet_name, index=False)

        with open("merged_data.xlsx", "rb") as file:
            st.download_button(label="Download Merged Excel File", data=file, file_name="merged_data.xlsx")

    else:
        st.warning("Please provide all the required inputs.")