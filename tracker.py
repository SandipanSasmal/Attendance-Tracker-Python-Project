import os
import pandas as pd
from datetime import datetime

def calculate_attendance(data):
  """
  Calculates attendance percentage for a student in a subject.

  Args:
      data (pandas.DataFrame): A DataFrame containing attendance data.

  Returns:
      float: Attendance percentage as a decimal.
  """
  total_classes = len(data)
  present_count = len(data[data["Attendance Status"] == "PRESENT"])
  if total_classes == 0:
    return 0.0  # Include students with 0 attendance
  return present_count / total_classes * 100  # Percentage calculation

def process_attendance_data(month_dir):
  """
  Processes attendance data from a specific month folder.

  Args:
      month_dir (str): Path to the month subfolder within attendance directory.
  """
  month_str = os.path.basename(month_dir)  # Extract month string (e.g., Dec 2024)
  subject_files = [f for f in os.listdir(month_dir) if f.endswith(".xlsx")]

  # Load student data
  student_data = pd.read_excel("students.xlsx")

  # Create monthly tracker file
  tracker_file = os.path.join("attendance_tracker/monthly", f"attendance-{month_str}.xlsx")
  writer = pd.ExcelWriter(tracker_file, engine="openpyxl")

  for subject_file in subject_files:
    subject_code = subject_file.split(".")[0]  # Extract subject code
    subject_data = pd.read_excel(os.path.join(month_dir, subject_file))

    # Merge student and attendance data
    merged_data = pd.merge(student_data, subject_data, on=["Name", "University Roll Number"], how="left")

    # Calculate attendance percentage
    merged_data["Attendance Percentage"] = merged_data.apply(calculate_attendance, axis=1)

    # Create sheet and save data
    merged_data.to_excel(writer, sheet_name=subject_code, index=False)

  writer.save()
  print(f"Attendance data processed for {month_str}.")

def check_and_create_directories():
  """
  Checks and creates necessary directories for the tracker.
  """
  if not os.path.exists("attendance_tracker/monthly"):
    os.makedirs("attendance_tracker/monthly")

  # Iterate through attendance subdirectories
  for month_dir in os.listdir("attendance"):
    month_path = os.path.join("attendance", month_dir)
    if os.path.isdir(month_path):
      process_attendance_data(month_path)  # Process data for each month

if __name__ == "__main__":
  check_and_create_directories()
  print("Attendance tracker script finished.")