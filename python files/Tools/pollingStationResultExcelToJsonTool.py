import os
import pandas as pd
import json
import shutil
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Define input and output folders and subdirectories
input_folder = r"C:\Users\Akshat\Desktop\Election_Ananlysis_Project\PollingStationResult_excel"
output_folder = r"C:\Users\Akshat\Desktop\Election_Ananlysis_Project\PollingJson"
error_folder = r"C:\Users\Akshat\Desktop\Election_Ananlysis_Project\error_file"
subdirectories = {"AC2009": 2009, "AC2014": 2014, "AC2019": 2019}

# first row first column .. that is first column ..
search_strings = [
    "Serial No of Polling Station", "Sr.No of Polling Station",
    "Sr No. of Polling Station", "Sr. No of Polling Station",
    "Serial No.of Polling Station", "Round No",
    "Polling Station", "Polling Station No.",
    "Sr.No of P.S", "Polling Serial No. Station",
    "SrNo of Polling Station", "Sr. No. of Polling Station"
]
# unwanted words in polling station number column
words_to_drop = [
    "Total","Total_","total","_total" 
]
    
# unwanted headers 
drop_columns = [
    "No. of Tendered Votes", "No.of Tender votes", "No. of tendered votes",
    "No.of Tender votes", "Tendered votes", "No. of Tender-ed Votes", "No Of Tendered Votes",
    "No.of tendered votes", "No.of Tender votes", "Total of Valid votes", "Total Of Valid Votes",
    "Total of valid votes", "TotalValid Votes", "No. of rejected votes", "No of Rejected votes",
    "No. of rejected votes", "No. of Rejected Votes", "No Of Rejected Votes", "Total",
    "Total Voters", "Male", "Female", "Sr No.", "Sr. No.", "Total Votes","Page 1 of 17"
    "Serail No. of Polling Station","No of rejected votes","No of tendered votes"
]

def validate_excel_files(input_folder, subdirectories):
    """
    Validate the Excel files by checking for the presence of required files.
    Return lists of valid files and error files.
    """
    valid_files = []
    error_files = []

    for subdirectory in subdirectories:
        input_subfolder = os.path.join(input_folder, subdirectory)
        if not os.path.exists(input_subfolder):
            print(f"Input subdirectory {input_subfolder} does not exist. Skipping...")
            continue

        for filename in os.listdir(input_subfolder):
            if filename.endswith('.xlsx') or filename.endswith('.xls'):
                excel_path = os.path.join(input_subfolder, filename)

# check if in Excel have polling station number in colppend in valid_files else append in error file
                try:
                    df = pd.read_excel(excel_path) # read as dataframe               
                    if not df.empty and any(df.apply(lambda row: any(string in str(row) for string in search_strings), axis=1)):
                        valid_files.append(excel_path)
                    else:
                        error_files.append(excel_path)
                except Exception as e:
                    print(f"An error occurred while reading {excel_path}: {e}")
                    error_files.append(excel_path)

    return valid_files, error_files

def process_excel_file(excel_path, year):

# Process a single Excel file and return the JSON data.
    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            print(f"DataFrame is empty for {excel_path}. Skipping...")
            return []

        polling_station_index = [idx for idx, row in df.iterrows() 
                                 if any(string in str(row) for string in search_strings)]

        if not polling_station_index:
            print("Serial No of Polling Station not found in the DataFrame. Skipping...")
            return []

        start_index = polling_station_index[0]
        df = df.iloc[start_index:].reset_index(drop=True)
        df.ffill(inplace=True)
        df.columns = df.iloc[1]
        df = df[2:]
        df.rename(columns={df.columns[0]: "polling_station"}, inplace=True)

# drop the rows
        df = df[~df.iloc[:, 0].isin(words_to_drop)]

# drop columns 
        df.drop(columns=[col for col in df.columns if any(word in col for word in drop_columns)], inplace=True)

# drop row index 
        try:
            if df.iloc[0, 1] == '1' and df.iloc[0, 2] == '2' and df.iloc[0, 3] == '3' and df.iloc[0, 4] == '4' and df.iloc[0, 5] == '5':
                df.drop(df.index[0], inplace=True)
        except IndexError as e:
            print(f"IndexError: {e} - The DataFrame might not have enough columns.")
        except Exception as e:
            print(f"An error occurred: {e}")

# Drop row in polling station column where polling_station row length 7 and above 
        df= df[df["polling_station"].str.len()<7]

# if NOTA is not NOTA then rename :
        try:
            df.rename(columns={"Votes for 'NOTA' option": "NOTA"}, inplace=True)
        except Exception as e:
            print(f"error {e}")        

        constituency_no = os.path.splitext(os.path.basename(excel_path))[0]
        json_data = []
        for i in range(df.shape[0]):
            for j in range(1, df.shape[1]):
                if pd.notna(df.iloc[i, j]):
                    data = {
                        "constituency_no": constituency_no,
                        "polling_station_no": df.iloc[i, 0],
                        "candidate_name": df.columns[j],
                        "vote": int(df.iloc[i, j]),
                        "year": year
                    }
                    json_data.append(data)

        return json_data

    except Exception as e:
        print(f"An error occurred while processing {excel_path}: {e}")
        return []

def execute_json_conversion(valid_files, output_folder, year_dict):
    """
    Process the valid Excel files, transform the data, and save the output as JSON files.
    """
    for excel_path in valid_files:
        subdirectory = os.path.basename(os.path.dirname(excel_path))
        year = year_dict.get(subdirectory, None)
        if year is None:
            print(f"Year not found for subdirectory {subdirectory}. Skipping...")
            continue

        json_path = os.path.join(output_folder, subdirectory, os.path.splitext(os.path.basename(excel_path))[0] + '.json')
        json_data = process_excel_file(excel_path, year)

        if json_data:
            os.makedirs(os.path.dirname(json_path), exist_ok=True)
            with open(json_path, "w") as f:
                json.dump(json_data, f, indent=4)
            print(f"JSON file exported successfully for {excel_path}.")
        else:
            move_to_error_folder(excel_path, error_folder, subdirectory)

def move_to_error_folder(file_path, error_folder, subdirectory):
    """
    Move the specified file to the error folder, preserving subdirectory structure.
    """
    destination_folder = os.path.join(error_folder, subdirectory)
    os.makedirs(destination_folder, exist_ok=True)
    try:
        shutil.move(file_path, os.path.join(destination_folder, os.path.basename(file_path)))
        print(f"Moved {file_path} to {destination_folder}.")
    except Exception as e:
        print(f"An error occurred while moving {file_path} to {destination_folder}: {e}")

def validate_and_move_error_files(error_files, error_folder):
    """
    Move the error files to a specified error folder, preserving subdirectory structure.
    """
    for file in error_files:
        if file.endswith('.xlsx'):  # Ensure only .xlsx files are moved
            subdirectory = os.path.basename(os.path.dirname(file))
            move_to_error_folder(file, error_folder, subdirectory)

def main(input_folder, output_folder, subdirectories, error_folder):
    """
    Main function to validate the Excel files and process the valid files.
    """
    valid_files, error_files = validate_excel_files(input_folder, subdirectories)

    print("Valid files:")
    for file in valid_files:
        print(file)

    print("Error files:")
    for file in error_files:
        print(file)

    execute_json_conversion(valid_files, output_folder, subdirectories)
    validate_and_move_error_files(error_files, error_folder)

if __name__ == "__main__":
    main(input_folder, output_folder, subdirectories, error_folder)
