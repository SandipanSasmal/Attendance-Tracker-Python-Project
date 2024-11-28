import cv2
import pyzbar.pyzbar as pyzbar
import time
import base64
from datetime import datetime
import pandas as pd
import os

# Function to check the admin passcode
def check_admin_passcode():
    try:
        with open('admin_pass.txt', 'r') as f:
            correct_passcode = f.read().strip()
    except FileNotFoundError:
        print("Error: 'admin_pass.txt' not found. Please ensure the file is in the same directory as this script.")
        exit()

    passcode_input = input("Enter Admin passcode: ").strip()

    if passcode_input == correct_passcode:
        print("Passcode correct. Proceeding with attendance script...")
    else:
        print("Incorrect passcode. Exiting...")
        exit()

# Check admin passcode before running the rest of the script
check_admin_passcode()

# Load the student data from Excel file
try:
    df = pd.read_excel('students.xlsx')
    df.columns = df.columns.str.lower()  # Normalize column names
    if 'name' not in df.columns or 'university roll number' not in df.columns:
        raise ValueError("Excel file must have 'Name' and 'University Roll Number' columns.")
    students = df[['name', 'university roll number']].dropna()
    students_list = students.apply(lambda x: [x[0].strip(), str(x[1]).strip()], axis=1).values.tolist()
    print("Student Data Loaded:")
    print(students)
except FileNotFoundError:
    print("Error: 'students.xlsx' not found. Ensure it is in the same directory.")
    exit()
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

names = []  # Track present students
attendance_data = []  # Collect attendance data

# Prompt user for subject code
subject_code = input("Enter the subject code: ").strip()

# Set up directory structure
current_date = datetime.now()
month_name = current_date.strftime("%B %Y")
date_str = current_date.strftime("%a %d-%m-%y")
base_path = os.path.join(os.getcwd(), month_name, subject_code)
os.makedirs(base_path, exist_ok=True)
excel_file_path = os.path.join(base_path, "attendance.xlsx")

# Check if the sheet for today's date already exists
sheet_exists = False
if os.path.exists(excel_file_path):
    with pd.ExcelWriter(excel_file_path, mode="a", engine="openpyxl") as writer:
        if date_str in writer.book.sheetnames:
            sheet_exists = True

if sheet_exists:
    print(f"Attendance for {subject_code} already exists for {date_str}.")
    choice = input("Would you like to overwrite (o), create a new sheet (n), or cancel (c)? ").strip().lower()
    if choice == "c":
        print("Cancelled. Exiting the program.")
        exit()
    elif choice == "o":
        print(f"Overwriting attendance for {date_str}.")
    elif choice == "n":
        # Find the next available suffix for the sheet name
        suffix = 2
        new_sheet_name = f"{date_str} {suffix}nd"
        with pd.ExcelWriter(excel_file_path, mode="a", engine="openpyxl") as writer:
            while new_sheet_name in writer.book.sheetnames:
                suffix += 1
                new_sheet_name = f"{date_str} {suffix}nd"
        date_str = new_sheet_name
        print(f"Creating new sheet: {date_str}.")
    else:
        print("Invalid choice. Exiting the program.")
        exit()

# Function to write attendance to the Excel file
def write_to_excel(data, date_str):
    if os.path.exists(excel_file_path):
        with pd.ExcelWriter(excel_file_path, mode="a", engine="openpyxl") as writer:
            if date_str in writer.book.sheetnames:
                del writer.book[date_str]  # Overwrite case
            pd.DataFrame(data, columns=['Student Name', 'University Roll Number', 'Attendance Status', 'Time']).to_excel(writer, sheet_name=date_str, index=False)
    else:
        with pd.ExcelWriter(excel_file_path, engine="openpyxl") as writer:
            pd.DataFrame(data, columns=['Student Name', 'University Roll Number', 'Attendance Status', 'Time']).to_excel(writer, sheet_name=date_str, index=False)
    print(f"Attendance saved to {excel_file_path} in sheet {date_str}.")

# Function to mark attendance
def enterData(z):
    if z not in names:
        names.append(z)
        decoded_data = base64.b64decode(z).decode('utf-8').strip()
        name, roll_number = decoded_data.split('|')
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        attendance_data.append([name, roll_number, 'PRESENT', timestamp])
        print(f"Marked PRESENT: {name} (University Roll Number: {roll_number}) at {timestamp}")

# Function to validate scanned data
def checkData(data):
    data = data.decode('utf-8')
    decoded_data = base64.b64decode(data).decode('utf-8').strip()
    name, roll_number = decoded_data.split('|')
    if [name.strip(), roll_number.strip()] in students_list:
        enterData(data)
    else:
        print(f"Invalid student: {name} (University Roll Number: {roll_number})")

# Start scanning
cap = cv2.VideoCapture(0)
while True:
    _, frame = cap.read()
    decoded_objects = pyzbar.decode(frame)
    for obj in decoded_objects:
        checkData(obj.data)
        time.sleep(1)
    cv2.imshow('Scan your ID Card', frame)
    if cv2.waitKey(1) & 0xFF == 27:
        break
cap.release()
cv2.destroyAllWindows()

# Mark absentees
for _, row in students.iterrows():
    student_name = row['name']
    student_roll_number = row['university roll number']
    data = f"{student_name}|{student_roll_number}"
    encoded_data = base64.b64encode(data.encode('utf-8')).decode('utf-8')
    if encoded_data not in names:
        attendance_data.append([student_name, student_roll_number, 'ABSENT', 'N/A'])
        print(f"Marked ABSENT: {student_name} (University Roll Number: {student_roll_number})")

# Save attendance
write_to_excel(attendance_data, date_str)
