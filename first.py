import pandas as pd
import hashlib
import chardet

# MD5 hash 
def calculate_md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

#file encoding
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))
    return result['encoding']

#process and compare
def process_files(file1_path, file2_path):
    # Calculate MD5 hashes
    md5_file1 = calculate_md5(file1_path)
    md5_file2 = calculate_md5(file2_path)

    print(f"MD5 hash of {file1_path}: {md5_file1}")
    print(f"MD5 hash of {file2_path}: {md5_file2}")

    # Compare hashes
    if md5_file1 == md5_file2:
        print("The files are identical (based on MD5 hash).")
    else:
        print("The files are different (based on MD5 hash).")
        # Detect encoding and read the CSV files
        encoding1 = detect_encoding(file1_path)
        encoding2 = detect_encoding(file2_path)
        
        df1 = pd.read_csv(file1_path, index_col=0, encoding=encoding1)
        df2_copy = pd.read_csv(file2_path, index_col=0, encoding=encoding2)

        # Remove any unnamed columns
        df1 = df1.loc[:, ~df1.columns.str.contains('^Unnamed')]
        df2_copy = df2_copy.loc[:, ~df2_copy.columns.str.contains('^Unnamed')]

        # Align the indices and columns of both DataFrames
        df1, df2_copy = df1.align(df2_copy, join='outer')

        # Compare the DataFrames and get the differences
        difference = df1.compare(df2_copy)

        # Save the differences to a new CSV file
        diff_filepath = 'differences.csv'
        difference.to_csv(diff_filepath)

        changes = difference.index.unique()
        print(f'The following rows have changes:\n{changes}\n')
        print('Detailed changes:')
        print(difference)

# Prompt the user to enter file paths
file1_path = input("Enter the file path for the first CSV file: ")
file2_path = input("Enter the file path for the second CSV file: ")

# Process the files
process_files(file1_path, file2_path)
