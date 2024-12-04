import tkinter as tk
from tkinter import Text, Scrollbar
from PIL import Image, ImageTk

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Tracker")
        self.root.geometry("800x600")

        self.bg_image = None
        self.bg_label = None

        self.create_main_menu()
        self.root.bind("<Configure>", self.resize_background)

    def create_main_menu(self):
        self.clear_screen()

        # Load and set the background image
        self.bg_image_original = Image.open("bg.jpg")
        self.update_background_image()

        self.main_label = tk.Label(self.root, text="Attendance Tracker", font=("Arial", 24), bg="white", fg="black")
        self.main_label.pack(pady=20)

        self.scanner_button = tk.Button(self.root, text="Scanner", command=self.scanner, font=("Arial", 14))
        self.scanner_button.pack(pady=10)

        self.admin_button = tk.Button(self.root, text="Admin Control", command=self.create_login_screen, font=("Arial", 14))
        self.admin_button.pack(pady=10)

        self.attendance_button = tk.Button(self.root, text="Attendance", command=self.create_student_portal, font=("Arial", 14))
        self.attendance_button.pack(pady=10)

        # Output terminal with scrollbar
        self.output_frame = tk.Frame(self.root)
        self.output_frame.pack(fill=tk.BOTH, padx=10, pady=10, side=tk.BOTTOM)

        self.output_terminal = Text(self.output_frame, height=9, bg="black", fg="white", font=("Arial", 10), padx=10, pady=10)
        self.output_terminal.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = Scrollbar(self.output_frame, command=self.output_terminal.yview, bg="grey")
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.output_terminal.config(yscrollcommand=self.scrollbar.set)
        self.output_terminal.insert(tk.END, "OUTPUT\n")
        self.output_terminal.insert(tk.END, "." * 100 + "\n")
        self.output_terminal.config(state=tk.DISABLED)

    def update_background_image(self):
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        bg_image_resized = self.bg_image_original.resize((width, height), Image.LANCZOS)
        bg_image_resized = bg_image_resized.convert("RGBA")
        alpha = 128  # Semi-transparency
        bg_image_resized = Image.blend(bg_image_resized, Image.new("RGBA", bg_image_resized.size, (255, 255, 255, alpha)), alpha / 255.0)
        self.bg_image = ImageTk.PhotoImage(bg_image_resized)

        if self.bg_label is None:
            self.bg_label = tk.Label(self.root, image=self.bg_image)
            self.bg_label.place(relwidth=1, relheight=1)
        else:
            self.bg_label.config(image=self.bg_image)

    def resize_background(self, event):
        self.update_background_image()

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def scanner(self):
        self.output_terminal.config(state=tk.NORMAL)
        self.output_terminal.insert(tk.END, "Scanner functionality not implemented yet.\n\n")
        self.output_terminal.config(state=tk.DISABLED)

    def create_login_screen(self):
        self.output_terminal.config(state=tk.NORMAL)
        self.output_terminal.insert(tk.END, "Admin Control functionality not implemented yet.\n\n")
        self.output_terminal.config(state=tk.DISABLED)

    def create_student_portal(self):
        self.output_terminal.config(state=tk.NORMAL)
        self.output_terminal.insert(tk.END, "Attendance functionality not implemented yet.\n\n")
        self.output_terminal.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()