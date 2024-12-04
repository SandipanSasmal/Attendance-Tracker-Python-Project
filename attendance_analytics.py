import pandas as pd
import os

def load_attendance_data(attendance_dir):
    """
    Load attendance data from all sheets in Excel files in the specified directory.

    Args:
        attendance_dir (str): Path to the attendance directory.

    Returns:
        dict: A dictionary with keys as (month, stream) and values as DataFrames.
    """
    attendance_data = {}
    for month_dir in os.listdir(attendance_dir):
        month_path = os.path.join(attendance_dir, month_dir)
        if os.path.isdir(month_path):
            for stream_dir in os.listdir(month_path):
                stream_path = os.path.join(month_path, stream_dir)
                if os.path.isdir(stream_path):
                    for file in os.listdir(stream_path):
                        if file.endswith(".xlsx"):
                            file_path = os.path.join(stream_path, file)
                            all_sheets_df = pd.concat(pd.read_excel(file_path, sheet_name=None), ignore_index=True)
                            if stream_dir not in attendance_data:
                                attendance_data[stream_dir] = []
                            attendance_data[stream_dir].append(all_sheets_df)
                            print(f"Loaded data for {month_dir} - {stream_dir}: {all_sheets_df.shape[0]} records")
    return attendance_data

def calculate_student_analytics(attendance_data):
    """
    Calculate attendance analytics for each student.

    Args:
        attendance_data (dict): A dictionary with keys as (stream) and values as lists of DataFrames.

    Returns:
        dict: A dictionary with student analytics.
    """
    student_analytics = {}
    for stream, dfs in attendance_data.items():
        combined_df = pd.concat(dfs, ignore_index=True)
        for _, row in combined_df.iterrows():
            student_name = row['Student Name']
            roll_number = row['University Roll Number']
            status = row['Attendance Status']
            key = stream
            if key not in student_analytics:
                student_analytics[key] = {}
            if (student_name, roll_number) not in student_analytics[key]:
                student_analytics[key][(student_name, roll_number)] = {'total_classes': 0, 'present': 0}
            student_analytics[key][(student_name, roll_number)]['total_classes'] += 1
            if status == 'PRESENT':
                student_analytics[key][(student_name, roll_number)]['present'] += 1
    return student_analytics

def save_analytics_to_excel(student_analytics, output_dir):
    """
    Save analytics to an Excel file in the specified directory structure.

    Args:
        student_analytics (dict): A dictionary with student analytics.
        output_dir (str): Path to the output directory.
    """
    for stream, data in student_analytics.items():
        stream_dir = os.path.join(output_dir, stream)
        os.makedirs(stream_dir, exist_ok=True)
        output_file = os.path.join(stream_dir, "analytics.xlsx")
        
        # Create a DataFrame for the analytics
        analytics_data = []
        for (student_name, roll_number), stats in data.items():
            total_classes = stats['total_classes']
            present = stats['present']
            attendance_percentage = (present / total_classes) * 100 if total_classes > 0 else 0
            analytics_data.append([
                student_name, roll_number, total_classes, present, attendance_percentage
            ])
        
        df = pd.DataFrame(analytics_data, columns=[
            'Student Name', 'University Roll Number', 'Total Classes', 'Present', 'Attendance Percentage'
        ])
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Analytics', index=False)
        
        print(f"Saved analytics for {stream} to {output_file}")

def save_cumulative_analytics(student_analytics, output_dir):
    """
    Save cumulative analytics to an Excel file in the specified directory structure.

    Args:
        student_analytics (dict): A dictionary with student analytics.
        output_dir (str): Path to the output directory.
    """
    os.makedirs(output_dir, exist_ok=True)
    cumulative_data = {}
    for stream, data in student_analytics.items():
        for (student_name, roll_number), stats in data.items():
            if (student_name, roll_number) not in cumulative_data:
                cumulative_data[(student_name, roll_number)] = {}
            if stream not in cumulative_data[(student_name, roll_number)]:
                cumulative_data[(student_name, roll_number)][stream] = {'total_classes': 0, 'present': 0}
            cumulative_data[(student_name, roll_number)][stream]['total_classes'] += stats['total_classes']
            cumulative_data[(student_name, roll_number)][stream]['present'] += stats['present']

    # Create a DataFrame for the cumulative analytics
    cumulative_analytics_data = []
    for (student_name, roll_number), subjects in cumulative_data.items():
        row = [student_name, roll_number]
        for subject, stats in subjects.items():
            total_classes = stats['total_classes']
            present = stats['present']
            attendance_percentage = (present / total_classes) * 100 if total_classes > 0 else 0
            row.append(attendance_percentage)
        cumulative_analytics_data.append(row)

    # Ensure columns match the data
    columns = ['Student Name', 'University Roll Number'] + list(next(iter(cumulative_data.values())).keys())
    df = pd.DataFrame(cumulative_analytics_data, columns=columns)

    # Save to Excel
    output_file = os.path.join(output_dir, "cumulative_analytics.xlsx")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cumulative Analytics', index=False)
    
    print(f"Saved cumulative analytics to {output_file}")

if __name__ == "__main__":
    attendance_dir = "attendance"  # Path to the attendance directory
    output_dir = "analytics"  # Output directory for analytics
    cumulative_output_dir = "analytics_sems"  # Output directory for cumulative analytics

    attendance_data = load_attendance_data(attendance_dir)
    if not attendance_data:
        print("No attendance data found.")
    else:
        student_analytics = calculate_student_analytics(attendance_data)
        save_analytics_to_excel(student_analytics, output_dir)
        save_cumulative_analytics(student_analytics, cumulative_output_dir)
        print(f"Analytics saved to {output_dir}")
        print(f"Cumulative analytics saved to {cumulative_output_dir}")