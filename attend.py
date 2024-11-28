import cv2
import pyzbar.pyzbar as pyzbar
import time
import base64
from datetime import datetime
import pandas as pd

# Load the student data from Excel file
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
    students_list = students[['name', 'university roll number']].apply(lambda x: [x[0].strip(), str(x[1]).strip()], axis=1).values.tolist()  # List of (name, roll number) pairs with trimmed spaces

    print("Student Data Loaded:")
    print(students)

except FileNotFoundError:
    print("Error: 'students.xlsx' not found. Please ensure the file is in the same directory as this script.")
    exit()
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

names = []  # List to keep track of students who are marked present
attendance_data = []  # List to hold attendance data for Excel

# Function to mark attendance
def enterData(z):
    if z in names:
        pass
    else:
        names.append(z)
        # Decode the Base64-encoded name and university roll number
        decoded_data = base64.b64decode(z).decode('utf-8').strip()
        name, roll_number = decoded_data.split('|')  # Split into name and university roll number
        # Get the current timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Add to attendance data
        attendance_data.append([name, roll_number, 'PRESENT', timestamp])
        print(f"Written to attendance: {name} (University Roll Number: {roll_number}) at {timestamp}")
        return names

print('Reading code.....')

# Function to check if the scanned data is valid
def checkData(data):
    data = data.decode('utf-8')  # Decode QR code data from bytes to string
    decoded_data = base64.b64decode(data).decode('utf-8').strip()  # Decode the Base64 to original data
    name, roll_number = decoded_data.split('|')  # Split into name and university roll number

    print(f"Decoded QR Data: Name = {name}, University Roll Number = {roll_number}")

    # Trim spaces from name and roll number before checking
    name = name.strip()
    roll_number = roll_number.strip()

    # Check if the (name, roll_number) pair is in the students list
    if [name, roll_number] in students_list:
        if roll_number not in [item[1] for item in names]:  # Check if this student is already marked
            print(f"{name} (University Roll Number: {roll_number}): PRESENT")
            enterData(data)
        else:
            print(f"{name} (University Roll Number: {roll_number}): ALREADY PRESENT")
    else:
        print(f"Student with University Roll Number: {roll_number} NOT FOUND")

# Start the webcam feed
cap = cv2.VideoCapture(0)

while True:
    _, frame = cap.read()
    decodedObject = pyzbar.decode(frame)  # Decode the QR code in the frame

    for obj in decodedObject:
        checkData(obj.data)  # Process the QR code data
        time.sleep(1)

    cv2.imshow('Scan your ID Card', frame)

    # Exit when 'Escape' key is pressed
    if cv2.waitKey(1) & 0xFF == 27:  # 27 is ASCII code for 'Esc'
        break

# Clean up resources
cv2.destroyAllWindows()
cap.release()

# Mark absent students
for _, row in students.iterrows():
    student_name = row['name']
    student_roll_number = row['university roll number']
    # Encode the student name and roll number in Base64 to match the QR code format
    data = f"{student_name}|{student_roll_number}"
    encoded_data = base64.b64encode(data.encode('utf-8')).decode('utf-8')
    if encoded_data not in names:
        attendance_data.append([student_name, student_roll_number, 'ABSENT', 'N/A'])
        print(f"{student_name} (University Roll Number: {student_roll_number}): MARKED ABSENT")

# Create DataFrame for attendance
attendance_df = pd.DataFrame(attendance_data, columns=['Student Name', 'University Roll Number', 'Attendance Status', 'Time'])

# Convert 'University Roll Number' to string before sorting
attendance_df['University Roll Number'] = attendance_df['University Roll Number'].astype(str)

# Sort the attendance by University Roll Number
attendance_df = attendance_df.sort_values(by='University Roll Number')

# Save to Excel
attendance_filename = f'attendance_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
attendance_df.to_excel(attendance_filename, index=False)

print(f"Attendance has been saved to {attendance_filename}")
