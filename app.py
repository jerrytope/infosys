import streamlit as st
import os
import pandas as pd
from io import BytesIO

# First Function: Split the data by class and save it to individual files
def split_class_data(input_file, output_dir):
    # Load the Excel file
    df = pd.read_excel(input_file)

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Split the DataFrame by Class
    classes = df['Class'].unique()

    saved_files = []
    for cls in classes:
        # Clean up the class name by removing invalid characters (e.g., newline characters)
        clean_class_name = cls.replace('\n', ' ').strip()

        # Filter the DataFrame for the current class
        class_df = df[df['Class'] == cls]
        columns_to_drop = [0, 2, 3, 22, 23, 24, 25, 26]
        class_df = class_df.drop(df.columns[columns_to_drop], axis=1)

        # Create a new Excel file for each class with the cleaned class name
        filename = os.path.join(output_dir, f"jss1 second term {clean_class_name}.xlsx")
        class_df.to_excel(filename, index=False, engine='openpyxl')
        saved_files.append(filename)

    return saved_files

# Second Function: Extract score per subject for each class file
def extract_columns_to_workbooks(input_file, output_prefix):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)
    
    # Get the first column (assumed to be the student's names or admission IDs)
    first_column = df.iloc[:, 0]
    
    # Loop through the rest of the columns, starting from the second column
    for i in range(1, df.shape[1]):
        # Extract the first column and the current column (subject scores)
        extracted_data = pd.concat([first_column, df.iloc[:, i]], axis=1)
        
        # Get the column name for the current column
        column_name = df.columns[i]
        
        # Create an output filename using the name of the column (subject)
        output_file = f"{output_prefix}_{column_name}.xlsx"
        
        # Save the extracted columns to a new Excel file
        extracted_data.to_excel(output_file, index=False)
        print(f"Workbook created: {output_file}")

# Main process function
def process_class_data(input_file, output_dir, final_output_dir):
    # Step 1: Split class data and save individual class workbooks
    saved_files = split_class_data(input_file, output_dir)

    # Ensure the final output directory exists for extracted subject files
    os.makedirs(final_output_dir, exist_ok=True)

    # Step 2: For each saved class file, extract score per subject
    for saved_file in saved_files:
        # Get the class name from the saved file name to use as a prefix
        class_name = os.path.basename(saved_file).replace(".xlsx", "")
        output_prefix = os.path.join(final_output_dir, class_name)

        # Extract subject scores into separate files
        extract_columns_to_workbooks(saved_file, output_prefix)

# Streamlit Application
def main():
    st.title("Class Data Processor")

    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        # Create temporary directories for saving outputs
        class_output_dir = "class_data"
        final_output_dir = "final_data"
        
        with st.spinner('Processing data...'):
            # Save uploaded file to a buffer to pass it to pandas
            input_file = BytesIO(uploaded_file.getvalue())
            
            # Run the main data processing function
            process_class_data(input_file, class_output_dir, final_output_dir)

            st.success("Data processing completed!")

            # Display links to download the generated files
            st.write("### Download Processed Files:")
            
            # List the files in the final output directory and create download links
            for filename in os.listdir(final_output_dir):
                file_path = os.path.join(final_output_dir, filename)
                with open(file_path, "rb") as f:
                    btn = st.download_button(
                        label=f"Download {filename}",
                        data=f,
                        file_name=filename
                    )

if __name__ == "__main__":
    main()
