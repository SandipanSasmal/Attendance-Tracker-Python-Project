from MyQR import myqr
import os
import base64
import pandas as pd

# Folder to save QR codes
qr_folder = 'generated_qrs'

# Create the folder if it doesn't exist
if not os.path.exists(qr_folder):
    os.makedirs(qr_folder)

# Load the student data from the Excel file
try:
    # Read the Excel file
    df = pd.read_excel('students.xlsx')

    # Normalize column names to lowercase for case-insensitive matching
    df.columns = df.columns.str.lower()

    # Ensure required columns exist
    if 'name' not in df.columns or 'university roll number' not in df.columns:
        raise ValueError("Excel file must have columns 'Name' and 'University Roll Number' (case-insensitive).")

    # Extract the student names and university roll numbers
    students = df[['name', 'university roll number']].dropna()  # Drop rows with missing data
    students_list = students[['name', 'university roll number']].apply(lambda x: [x[0].strip(), str(x[1]).strip()], axis=1).values.tolist()  # List of (name, roll number) pairs

    print("Student Data Loaded:")
    print(students)

except FileNotFoundError:
    print("Error: 'students.xlsx' not found. Please ensure the file is in the same directory as this script.")
    exit()
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# List to track students with existing QR codes
existing_qrs = []

# Check if QR code already exists for a student
for student in students_list:
    student_name = student[0]
    student_roll_number = student[1]

    # Define the QR code filename
    qr_filename = f"{student_name}_{student_roll_number}.png"
    qr_filepath = os.path.join(qr_folder, qr_filename)

    if os.path.exists(qr_filepath):
        # If QR file exists, add to the list
        existing_qrs.append(f"{student_name} (University Roll Number: {student_roll_number})")
    else:
        # If QR code doesn't exist, generate a new one
        try:
            data = f"{student_name}|{student_roll_number}"  # Format: name|university roll number
            encoded_data = base64.b64encode(data.encode('utf-8')).decode('utf-8')

            # Generate QR code
            version, level, qr_name = myqr.run(
                str(encoded_data),            # Data for the QR code
                level='H',                    # Error correction level
                version=1,                    # Version of the QR code
                picture='bg.jpg',             # Path to the background image
                colorized=True,               # Colorize the QR code
                contrast=1.0,                 # Contrast (integer)
                brightness=1.0,               # Brightness (integer)
                save_name=qr_filename,        # File name for the QR code image
                save_dir=qr_folder            # Save directory (specific folder)
            )
            print(f"QR code generated for: {student_name} (University Roll Number: {student_roll_number}) -> {qr_name}")
        except Exception as e:
            print(f"Error generating QR code for {student_name}: {e}")

# If any QR codes already exist, print the list
if existing_qrs:
    print("\nQR code already exists for the following students:")
    for qr in existing_qrs:
        print(qr)

else:
    print("\nAll QR codes generated successfully.")
