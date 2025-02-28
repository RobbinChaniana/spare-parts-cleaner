import streamlit as st
import pandas as pd
import re

# Function to clean descriptions
def clean_description(description):
    description = re.sub(r"[:,].*?(?=\s\d|\s\()", "", description)  # Remove supplier names
    description = description.replace('"', ' IN')  # Replace inch symbols
    description = description.upper().strip()  # Convert to uppercase
    return description

# Function to process uploaded Excel file
def process_spare_parts(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    parts_data = []
    current_part = None
    current_description = None

    for i, row in df_raw.iterrows():
        col_0 = str(row[0]).strip() if pd.notna(row[0]) else ""
        col_1 = str(row[1]).strip() if pd.notna(row[1]) else ""
        col_3 = str(row[3]).strip() if pd.notna(row[3]) else ""

        if col_0 == "Part No.:" and col_1:
            part_info = col_1.split(" - ", 1)
            if len(part_info) == 2:
                current_part, current_description = part_info
                current_part = current_part.strip()
                current_description = current_description.strip()

        elif col_1 == "Item -" and "Quantity:" in col_3:
            try:
                quantity = int(col_3.replace("Quantity:", "").strip())
            except ValueError:
                quantity = 0

            if current_part and current_description and quantity > 0:
                parts_data.append((quantity, current_part, current_description))

    # Create DataFrame
    df_cleaned = pd.DataFrame(parts_data, columns=["Quantity", "Part Number", "Description"])
    
    # Group duplicate part numbers and sum their quantities
    df_cleaned = df_cleaned.groupby(["Part Number", "Description"], as_index=False).sum()
    
    # Apply description formatting
    df_cleaned["Description"] = df_cleaned["Description"].apply(clean_description)

    return df_cleaned

# Streamlit UI
st.title("Spare Parts List Cleaner")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    cleaned_data = process_spare_parts(uploaded_file)
    
    # Display the cleaned table
    st.write("### Cleaned Spare Parts List:")
    st.dataframe(cleaned_data)

    # Download cleaned file
    output_file = "cleaned_spare_parts.xlsx"
    cleaned_data.to_excel(output_file, index=False)
    
    with open(output_file, "rb") as f:
        st.download_button("Download Cleaned File", f, file_name="cleaned_spare_parts.xlsx")

