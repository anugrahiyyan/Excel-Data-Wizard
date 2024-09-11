import pandas as pd
import os

def extract_emails_and_phones_excel(df, email_column, phone_column):
    selected_columns = [email_column]
    if phone_column in df.columns or isinstance(phone_column, int):
        selected_columns.append(phone_column)
    emails_and_phones = df[selected_columns]
    col_names = ['Email']
    if len(selected_columns) > 1:
        col_names.append('Phone Number')
    emails_and_phones.columns = col_names
    return emails_and_phones

def extract_emails_and_phones_multi_column(row):
    try:
        email = str(row[1]).strip()
        phone = str(row[2]).strip()
        return pd.Series([email, phone], index=['Email', 'Phone Number'])
    except Exception as e:
        print(f"Error processing row: {row} - {e}")
    return pd.Series([None, None], index=['Email', 'Phone Number'])

all_data = []

input_choice = input("Do you want to input multiple files (f) or a folder (d)? Enter 'f' for files and 'd' for folder: ")

if input_choice.lower() == 'f':
    file_paths = input("Enter the paths to your files (separated by commas): ").split(',')
    for file_path in file_paths:
        file_path = file_path.strip()
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            df = pd.read_excel(file_path)
            email_column = input(f"Enter the column name or index for Email in file {file_path} (e.g., 'E' or 4): ")
            phone_column = input(f"Enter the column name or index for Phone Number in file {file_path} (e.g., 'F' or 5), or leave blank if none: ")
            try:
                email_column = int(email_column) - 1
            except ValueError:
                pass
            if phone_column.strip() != "":
                try:
                    phone_column = int(phone_column) - 1
                except ValueError:
                    pass
            else:
                phone_column = None
            emails_and_phones = extract_emails_and_phones_excel(df, email_column, phone_column)
            all_data.append(emails_and_phones)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=None)
            df_extracted = df.apply(extract_emails_and_phones_multi_column, axis=1)
            all_data.append(df_extracted)

elif input_choice.lower() == 'd':
    folder_path = input("Please enter the path to the folder containing files: ")
    email_column = input("Enter the column name or index for Email in all files (e.g., 'E' or 4), or leave blank if none: ")
    phone_column = input("Enter the column name or index for Phone Number in all files (e.g., 'F' or 5), or leave blank if none: ")
    try:
        email_column = int(email_column) - 1
    except ValueError:
        email_column = None
    if phone_column.strip() != "":
        try:
            phone_column = int(phone_column) - 1
        except ValueError:
            phone_column = None
    else:
        phone_column = None

    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            df = pd.read_excel(file_path)
            emails_and_phones = extract_emails_and_phones_excel(df, email_column, phone_column)
            all_data.append(emails_and_phones)
        elif file_name.endswith('.csv'):
            df = pd.read_csv(file_path, header=None)
            df_extracted = df.apply(extract_emails_and_phones_multi_column, axis=1)
            all_data.append(df_extracted)

final_data = pd.concat(all_data, ignore_index=True)

print(final_data)

output_path = input("Enter the output file path (e.g., extracted_data.xlsx): ")
try:
    final_data.to_excel(output_path, index=False)
    print(f"Extracted data saved to {output_path}")
except Exception as e:
    print(f"Error saving file: {e}")
