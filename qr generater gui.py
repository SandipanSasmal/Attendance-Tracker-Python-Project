import tkinter as tk
from tkinter import scrolledtext, messagebox
import subprocess
import os

# Path to the Excel file
excel_file_path = r"C:\Users\Sandipan Sasmal\Desktop\python mini proj\students.xlsx"

# Function to run generate.py and capture its output
def run_generate_script():
    try:
        # Clear the terminal output
        terminal_output.config(state=tk.NORMAL)
        terminal_output.delete(1.0, tk.END)

        # Run generate.py and capture the output
        process = subprocess.Popen(
            ["python", "generate.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # Read the output and error streams
        stdout, stderr = process.communicate()

        # Display the output in the terminal
        terminal_output.insert(tk.END, stdout)
        if stderr:
            terminal_output.insert(tk.END, stderr)

        # Check if the script ran successfully
        if process.returncode == 0:
            messagebox.showinfo("Success", "QR codes generated successfully.")
        else:
            messagebox.showerror("Error", "An error occurred while generating QR codes.")
    except Exception as e:
        messagebox.showerror("Error", f"Error running generate.py: {e}")
    finally:
        terminal_output.config(state=tk.DISABLED)

# Create the main window
root = tk.Tk()
root.title("QR Code Generator")
root.geometry("700x500")
root.configure(bg="#f0f0f0")

# Create and place widgets
title_label = tk.Label(root, text="GENERATE QR", font=("Arial", 24, "bold"), bg="#f0f0f0")
title_label.pack(pady=20)

file_label = tk.Label(root, text=f"File used: {excel_file_path}", font=("Arial", 12), bg="#f0f0f0")
file_label.pack(pady=10)

generate_button = tk.Button(root, text="Generate QR Codes", command=run_generate_script, font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
generate_button.pack(pady=20)

# Terminal output box
terminal_frame = tk.Frame(root, bg="#f0f0f0")
terminal_frame.pack(pady=10, fill=tk.BOTH, expand=True)

terminal_output = scrolledtext.ScrolledText(terminal_frame, height=15, bg="black", fg="white", font=("Consolas", 10), padx=10, pady=10)
terminal_output.pack(fill=tk.BOTH, expand=True)
terminal_output.config(state=tk.DISABLED)

# Run the main loop
root.mainloop()