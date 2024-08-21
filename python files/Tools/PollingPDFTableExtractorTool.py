import os
import pandas as pd
import camelot
import warnings

# Base folder containing subdirectories for each election year
base_folder = r"C:\Users\Akshat\Desktop\testpdf"

# Output folder for Excel files
output_folder = r"C:\Users\Akshat\Desktop\testexcel"

# List of subdirectories for each election year
subdirectories = ["AC2009", "AC2014", "AC2019"]

# Initialize a list to store names of PDF files causing errors and valid PDF files
error_files = []
valid_files = []

# Function to clean and concatenate strings in DataFrame cells
def clean_and_concat(cell):
    if isinstance(cell, str):
        return ' '.join(cell.split())
    return cell

# Define the full path to the error.txt file
error_file_path = os.path.join(output_folder, "error.txt")

# Ignore warnings
warnings.filterwarnings("ignore")  

def validate():
    errors = []

    # Check if base_folder exists
    if not os.path.exists(base_folder):
        errors.append(f"Base folder '{base_folder}' does not exist.")
    
    # Check if output_folder exists, create if not
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            errors.append(f"Could not create output folder '{output_folder}': {e}")
    
    # Check each subdirectory
    for subdir in subdirectories:
        pdf_folder = os.path.join(base_folder, subdir)
        if not os.path.exists(pdf_folder):
            errors.append(f"PDF folder '{pdf_folder}' does not exist for subdir '{subdir}'.")
        else:
            # Check if there are any PDF files in the subdirectory
            pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]
            if not pdf_files:
                errors.append(f"No PDF files found in '{pdf_folder}'.")
            else:
                # Check if each PDF file contains tables
                for pdf_file in pdf_files:
                    pdf_path = os.path.join(pdf_folder, pdf_file)
                    try:
                        tables = camelot.read_pdf(pdf_path, pages="all") # read all pages of pdf file
                        if not tables or tables.n <= 1: # Ensure at least two tables
                            errors.append(f"PDF file '{pdf_path}' does not contain tables.") 
                        else:
                            valid_files.append(pdf_path)
                    except Exception as e:
                        errors.append(f"Error reading '{pdf_path}': {e}")
    
    return errors

def execute():
    # Iterate over valid PDF files
    for pdf_path in valid_files:
        pdf_file = os.path.basename(pdf_path)
        subdir = os.path.basename(os.path.dirname(pdf_path))
        output_subfolder = os.path.join(output_folder, subdir)
        os.makedirs(output_subfolder, exist_ok=True)  # Create output subfolder if it doesn't exist

        try:
            # Read tables from all pages of the PDF
            tables = camelot.read_pdf(pdf_path, pages="all")

            # Initialize an empty list to store DataFrames for each table
            dataframes = []

            # Iterate through each table
            for table in tables:
                # Extract the DataFrame from each table and append it to the list
                dataframes.append(table.df)

            # Concatenate all DataFrames into a single DataFrame
            combined_df = pd.concat(dataframes, ignore_index=True)

            # Apply the clean_and_concat function to every cell in the DataFrame
            combined_df_cleaned = combined_df.applymap(clean_and_concat)

            # Remove duplicate rows from combined_df_cleaned
            combined_df_cleaned_unique = combined_df_cleaned.drop_duplicates()

            # Export the unique DataFrame to an Excel file with the same name as the PDF file
            excel_file = os.path.splitext(pdf_file)[0] + ".xlsx"
            excel_path = os.path.join(output_subfolder, excel_file)
            combined_df_cleaned_unique.to_excel(excel_path, index=False)

            print(f"Excel file '{excel_file}' exported successfully.")

        except Exception as e:
            print(f"Error occurred while processing '{pdf_file}':", e)
            error_files.append(pdf_file)

    # Write the names of PDF files causing errors to the error.txt file
    if error_files:
        with open(error_file_path, "a") as f:
            f.write("\n".join(error_files) + "\n")

        print("Error log appended to 'error.txt'.")
    else:
        print("No errors occurred during processing.")

def main():
    validation_errors = validate()
    if validation_errors:
        print("Validation errors found:")
        with open(error_file_path, "w") as f:
            for error in validation_errors:
                print(f"- {error}")
                f.write(f"{error}\n")

    if valid_files:
        print("Proceeding to execution for valid files.")
        execute()
    else:
        print("No valid files to process.")

if __name__ == "__main__":
    main()
