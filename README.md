# Excel-Data-Wizard

# Phone Column Skipping:
1. The script now checks if the phone column exists in the file.
2. If the phone column is missing, it will only extract the email data and skip the phone column.

# User Option to Skip:
1. When inputting the file paths, the user can leave the phone column blank if it's not present in the file. The script will still work with just the email column.

# Column Check in DataFrame:
1. Inside the extract_emails_and_phones function, the code checks if the phone column exists before attempting to extract it, ensuring the script doesn't throw an error if it's absent.

# Usage:
1. If the phone column exists in some files and not in others, the script will handle this seamlessly. For files with no phone number, it will still capture the email.
You can either provide the phone column or leave it blank if you know the file doesn't have one.
