import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import os
from datetime import datetime
import base64
import cv2
import pyzbar.pyzbar as pyzbar
import time
import hashlib

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Tracker")
        self.root.geometry("600x400")

        self.student_data = None
        self.attendance_data = []
        self.names = []

        self.create_main_menu()

    def create_main_menu(self):
        self.clear_screen()

        self.main_label = tk.Label(self.root, text="Attendance Tracker", font=("Arial", 16))
        self.main_label.pack(pady=20)

        self.student_button = tk.Button(self.root, text="Student Portal", command=self.create_student_portal)
        self.student_button.pack(pady=10)

        self.admin_button = tk.Button(self.root, text="Admin Controls", command=self.create_login_screen)
        self.admin_button.pack(pady=10)

    def create_login_screen(self):
        self.clear_screen()

        self.login_label = tk.Label(self.root, text="Admin Login", font=("Arial", 16))
        self.login_label.pack(pady=20)

        self.username_label = tk.Label(self.root, text="Username:")
        self.username_label.pack(pady=5)
        self.username_entry = tk.Entry(self.root)
        self.username_entry.pack(pady=5)

        self.password_label = tk.Label(self.root, text="Password:")
        self.password_label.pack(pady=5)
        self.password_entry = tk.Entry(self.root, show="*")
        self.password_entry.pack(pady=5)

        self.login_button = tk.Button(self.root, text="Login", command=self.admin_login)
        self.login_button.pack(pady=20)

        self.back_button = tk.Button(self.root, text="Back", command=self.create_main_menu)
        self.back_button.pack(pady=10)

    def admin_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        # For simplicity, we use hardcoded credentials. In a real application, use a secure method.
        # Example of hashing a password (store the hashed password in a secure database)
        # hashed_password = hashlib.sha256(password.encode()).hexdigest()
        # Check the hashed password against the stored hashed password
        if username == "admin" and password == "password":
            self.create_admin_portal()
        else:
            messagebox.showerror("Error", "Invalid credentials")

    def create_admin_portal(self):
        self.clear_screen()

        self.subject_label = tk.Label(self.root, text="Enter Subject Code:")
        self.subject_label.pack(pady=5)
        self.subject_entry = tk.Entry(self.root)
        self.subject_entry.pack(pady=5)

        self.mark_button = tk.Button(self.root, text="Mark Attendance", command=self.mark_attendance)
        self.mark_button.pack(pady=10)

        self.analytics_button = tk.Button(self.root, text="Generate Analytics", command=self.generate_analytics)
        self.analytics_button.pack(pady=10)

        self.view_analytics_button = tk.Button(self.root, text="View Analytics", command=self.view_analytics)
        self.view_analytics_button.pack(pady=10)

        self.back_button = tk.Button(self.root, text="Back", command=self.create_main_menu)
        self.back_button.pack(pady=10)

    def create_student_portal(self):
        self.clear_screen()

        self.subject_label = tk.Label(self.root, text="Enter Subject Code:")
        self.subject_label.pack(pady=5)
        self.subject_entry = tk.Entry(self.root)
        self.subject_entry.pack(pady=5)

        self.mark_button = tk.Button(self.root, text="Mark Attendance", command=self.mark_attendance)
        self.mark_button.pack(pady=10)

        self.back_button = tk.Button(self.root, text="Back", command=self.create_main_menu)
        self.back_button.pack(pady=10)

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def mark_attendance(self):
        subject_code = self.subject_entry.get().strip()
        if not subject_code:
            messagebox.showerror("Error", "Please enter a subject code.")
            return

        current_date = datetime.now()
        month_name = current_date.strftime("%B %Y")
        date_str = current_date.strftime("%a %d-%m-%y")

        attendance_root = os.path.join(os.getcwd(), "attendance")
        os.makedirs(attendance_root, exist_ok=True)

        base_path = os.path.join(attendance_root, month_name, subject_code)
        os.makedirs(base_path, exist_ok=True)

        excel_file_path = os.path.join(base_path, "attendance.xlsx")

        sheet_exists = False
        if os.path.exists(excel_file_path):
            with pd.ExcelWriter(excel_file_path, mode="a", engine="openpyxl") as writer:
                if date_str in writer.book.sheetnames:
                    sheet_exists = True

        if sheet_exists:
            choice = messagebox.askquestion("Sheet Exists", f"Attendance for {subject_code} already exists for {date_str}. Overwrite?")
            if choice == 'no':
                return

        self.attendance_data = []
        self.names = []

        cap = cv2.VideoCapture(0)
        while True:
            _, frame = cap.read()
            decoded_objects = pyzbar.decode(frame)
            for obj in decoded_objects:
                self.check_data(obj.data)
                time.sleep(1)
            cv2.imshow('Scan your ID Card', frame)
            if cv2.waitKey(1) & 0xFF == 27:
                break
        cap.release()
        cv2.destroyAllWindows()

        for _, row in self.student_data.iterrows():
            student_name = row['name']
            student_roll_number = row['university roll number']
            data = f"{student_name}|{student_roll_number}"
            encoded_data = base64.b64encode(data.encode('utf-8')).decode('utf-8')
            if encoded_data not in self.names:
                self.attendance_data.append([student_name, student_roll_number, 'ABSENT', 'N/A'])

        self.write_to_excel(self.attendance_data, date_str, excel_file_path)
        messagebox.showinfo("Success", "Attendance Marked Successfully")

    def check_data(self, data):
        data = data.decode('utf-8')
        decoded_data = base64.b64decode(data).decode('utf-8').strip()
        name, roll_number = decoded_data.split('|')
        if [name.strip(), roll_number.strip()] in self.student_list:
            self.enter_data(data)
        else:
            print(f"Invalid student: {name} (University Roll Number: {roll_number})")

    def enter_data(self, z):
        if z not in self.names:
            self.names.append(z)
            decoded_data = base64.b64decode(z).decode('utf-8').strip()
            name, roll_number = decoded_data.split('|')
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.attendance_data.append([name, roll_number, 'PRESENT', timestamp])
            print(f"Marked PRESENT: {name} (University Roll Number: {roll_number}) at {timestamp}")

    def write_to_excel(self, data, date_str, excel_file_path):
        with pd.ExcelWriter(excel_file_path, mode="a", engine="openpyxl") as writer:
            pd.DataFrame(data, columns=['Student Name', 'University Roll Number', 'Attendance Status', 'Time']).to_excel(writer, sheet_name=date_str, index=False)
        print(f"Attendance saved to {excel_file_path} in sheet {date_str}.")

    def generate_analytics(self):
        attendance_dir = "attendance"
        output_dir = "analytics"

        attendance_data = self.load_attendance_data(attendance_dir)
        if not attendance_data:
            messagebox.showerror("Error", "No attendance data found.")
            return

        student_analytics = self.calculate_student_analytics(attendance_data)
        self.save_analytics_to_excel(student_analytics, output_dir)
        messagebox.showinfo("Success", f"Analytics saved to {output_dir}")

    def load_attendance_data(self, attendance_dir):
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
                                df = pd.read_excel(file_path)
                                attendance_data[(month_dir, stream_dir)] = df
                                print(f"Loaded data for {month_dir} - {stream_dir}: {df.shape[0]} records")
        return attendance_data

    def calculate_student_analytics(self, attendance_data):
        student_analytics = {}
        for (month, stream), df in attendance_data.items():
            for _, row in df.iterrows():
                student_name = row['Student Name']
                roll_number = row['University Roll Number']
                status = row['Attendance Status']
                key = (month, stream)
                if key not in student_analytics:
                    student_analytics[key] = {}
                if (student_name, roll_number) not in student_analytics[key]:
                    student_analytics[key][(student_name, roll_number)] = {'total_classes': 0, 'present': 0}
                student_analytics[key][(student_name, roll_number)]['total_classes'] += 1
                if status == 'PRESENT':
                    student_analytics[key][(student_name, roll_number)]['present'] += 1
        return student_analytics

    def save_analytics_to_excel(self, student_analytics, output_dir):
        for (month, stream), data in student_analytics.items():
            month_dir = os.path.join(output_dir, month)
            stream_dir = os.path.join(month_dir, stream)
            os.makedirs(stream_dir, exist_ok=True)
            output_file = os.path.join(stream_dir, "analytics.xlsx")

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

            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Analytics', index=False)

            print(f"Saved analytics for {stream} in {month} to {output_file}")

    def view_analytics(self):
        analytics_dir = "analytics"
        if not os.path.exists(analytics_dir):
            messagebox.showerror("Error", "No analytics data found.")
            return

        file_path = filedialog.askopenfilename(initialdir=analytics_dir, filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            os.startfile(file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()