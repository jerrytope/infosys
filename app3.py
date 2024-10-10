import streamlit as st
import os
import pandas as pd
from io import BytesIO
import zipfile
import tempfile
import shutil

# First Function: Split the data by class and save it to individual files
def split_class_data(input_file, output_dir, base_filename):
    df = pd.read_excel(input_file)

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Split the DataFrame by Class
    if 'Class' not in df.columns:
        raise ValueError("The input file must contain a 'Class' column.")

    df['Class'] = df['Class'].astype(str)
    classes = df['Class'].unique()

    saved_files = []
    for cls in classes:
        clean_class_name = cls.replace('\n', ' ').strip()
        class_df = df[df['Class'] == cls]
        # columns_to_drop = [0, 2, 3, 22, 23, 24, 25, 26]
        # class_df = class_df.drop(df.columns[columns_to_drop], axis=1, errors='ignore')
        columns_to_drop = [0, 2, 3]

        # Adding the last 5 columns dynamically to the drop list
        columns_to_drop += list(range(-5, 0))

        # Dropping the specified columns
        class_df = class_df.drop(class_df.columns[columns_to_drop], axis=1, errors='ignore')

        # to get the reg no and names only 
        # class_df = class_df.iloc[:, [1, 2]]

        filename = os.path.join(output_dir, f"{base_filename}_{clean_class_name}.xlsx")
        class_df.to_excel(filename, index=False, engine='openpyxl')
        saved_files.append(filename)

    return saved_files

# Second Function: Extract score per subject for each class file
def extract_columns_to_workbooks(input_file, output_prefix):
    df = pd.read_excel(input_file)
    first_column = df.iloc[:, 0]

    for i in range(1, df.shape[1]):
        extracted_data = pd.concat([first_column, df.iloc[:, i]], axis=1)
        column_name = df.columns[i]
        output_file = f"{output_prefix}_{column_name}.xlsx"
        extracted_data.to_excel(output_file, index=False)

# Function to zip files in a directory
def zip_files(output_dir, zip_file_name):
    with zipfile.ZipFile(zip_file_name, 'w') as zip_file:
        for foldername, _, filenames in os.walk(output_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zip_file.write(file_path, os.path.relpath(file_path, output_dir))
    return zip_file_name

# Main process function
def process_class_data(input_file, output_dir, final_output_dir, base_filename):
    saved_files = split_class_data(input_file, output_dir, base_filename)

    os.makedirs(final_output_dir, exist_ok=True)

    for saved_file in saved_files:
        class_name = os.path.basename(saved_file).replace(".xlsx", "")
        output_prefix = os.path.join(final_output_dir, class_name)
        extract_columns_to_workbooks(saved_file, output_prefix)

# Streamlit Application
def main():
    st.title("Class Data Processor")

    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        with tempfile.TemporaryDirectory() as temp_dir:
            class_output_dir = os.path.join(temp_dir, "class_data")
            final_output_dir = os.path.join(temp_dir, "final_data")
            
            base_filename = os.path.splitext(uploaded_file.name)[0]
            
            with st.spinner('Processing data...'):
                input_file = BytesIO(uploaded_file.getvalue())
                
                try:
                    process_class_data(input_file, class_output_dir, final_output_dir, base_filename)

                    zip_filename = os.path.join(temp_dir, f"{base_filename}_processed_files.zip")
                    zip_file_path = zip_files(final_output_dir, zip_filename)

                    st.success("Data processing completed!")

                    with open(zip_file_path, "rb") as f:
                        st.download_button(
                            label="Download All Processed Files as ZIP",
                            data=f,
                            file_name=os.path.basename(zip_file_path),
                            mime='application/zip'
                        )
                except Exception as e:
                    st.error(f"Error during processing: {e}")

if __name__ == "__main__":
    main()
