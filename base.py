import pandas as pd
import os
# Function to read Excel files from a directory
def read_excel_files(directory):
    all_data = []
    for file in os.listdir(directory):
        if file.endswith(".xlsx"):
            try:
                df = pd.read_excel(os.path.join(directory, file), engine='openpyxl')
                all_data.append(df)
            except Exception as e:
                print(f"Error reading file {file}: {e}")
    return all_data


# Function to merge Excel files into a single DataFrame
def merge_excel_files(dataframes):
    return pd.concat(dataframes, ignore_index=True)

# Function to write merged DataFrame to a new Excel file
def write_to_excel(df, output_file):
    df.to_excel(output_file, index=False)
    print(f"Merged data written to {output_file}")

# Main function
def main():
    # Directory containing Excel files
    directory = "C:\\Users\\soyeb\\PycharmProjects\\Centralised_Data_Hub\\excels"

    # Read Excel files
    excel_data = read_excel_files(directory)

    # Merge Excel files
    merged_data = merge_excel_files(excel_data)

    # Output file name
    output_file = "merged_data.xlsx"

    # Write merged data to Excel file
    write_to_excel(merged_data, output_file)

if __name__ == "__main__":
    main()
