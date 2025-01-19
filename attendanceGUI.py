import tkinter as tk
from tkinter import scrolledtext, simpledialog, messagebox
import subprocess
import threading

# Function to run the attend.py script
def run_attend_script():
    """
    Runs the attend.py script and captures its output in real time.
    Opens a dialog box if the script requests the Admin Passcode.
    """
    def execute_script():
        try:
            # Set the script path (update to the correct location of attend.py)
            script_path = "attend.py"  # Ensure attend.py is in the same directory or provide full path
            
            # Run the script and capture its output
            process = subprocess.Popen(
                ["python", script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.PIPE,
                text=True,
                bufsize=1
            )
            
            def handle_output(line):
                """Process each line of output from the script."""
                terminal.insert(tk.END, line)
                terminal.see(tk.END)  # Scroll to the end

                # Check if the script asks for the passcode
                if "Enter Admin passcode:" in line:
                    # Open a dialog box for passcode entry
                    passcode = simpledialog.askstring("Admin Passcode", "Enter Admin passcode:", show='*')
                    if passcode:
                        process.stdin.write(passcode + "\n")  # Send the passcode to the script
                        process.stdin.flush()
                    else:
                        process.terminate()  # Terminate the process if no passcode is provided
                        messagebox.showerror("Error", "Admin passcode is required. Process terminated.")
            
            # Display stdout in the terminal
            for line in iter(process.stdout.readline, ''):
                handle_output(line)
            process.stdout.close()

            # Display stderr in the terminal
            for line in iter(process.stderr.readline, ''):
                terminal.insert(tk.END, f"ERROR: {line}")
                terminal.see(tk.END)
            process.stderr.close()

            process.wait()
        except FileNotFoundError:
            terminal.insert(tk.END, "ERROR: attend.py not found. Please check the file path.\n")
        except Exception as e:
            terminal.insert(tk.END, f"ERROR: {str(e)}\n")
    
    # Run the script execution in a separate thread to prevent freezing the GUI
    thread = threading.Thread(target=execute_script)
    thread.start()

# Create the main GUI window
root = tk.Tk()
root.title("Take Attendance")
root.geometry("600x400")

# Title label
title_label = tk.Label(root, text="Take Attendance", font=("Arial", 20))
title_label.pack(pady=10)

# Proceed button
proceed_button = tk.Button(root, text="Proceed", font=("Arial", 14), command=run_attend_script)
proceed_button.pack(pady=10)

# Terminal area (ScrolledText widget to display script output)
terminal = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Consolas", 12))
terminal.pack(pady=10, fill=tk.BOTH, expand=True)

# Run the main loop
root.mainloop()
