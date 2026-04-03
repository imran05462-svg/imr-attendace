import tkinter as tk
from tkinter import filedialog, messagebox
from processor import main_process
import os

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Teacher Attendance Processor")
        self.root.geometry("500x480")
        self.root.configure(padx=20, pady=20)

        # Title
        title_label = tk.Label(root, text="Attendance Dashboard Generator", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # File Selection Frame
        file_frame = tk.Frame(root)
        file_frame.pack(fill="x", pady=10)

        self.file_path_var = tk.StringVar(value="No file selected")
        file_label = tk.Label(file_frame, text="Log File:", font=("Arial", 10, "bold"))
        file_label.pack(side="left", padx=(0, 10))

        self.path_entry = tk.Entry(file_frame, textvariable=self.file_path_var, width=40, state="readonly")
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_btn = tk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side="right")

        # Layout Selection
        layout_frame = tk.LabelFrame(root, text="Report Layout", padx=10, pady=10)
        layout_frame.pack(fill="x", pady=10)
        
        self.layout_var = tk.StringVar(value="horizontal")
        tk.Radiobutton(layout_frame, text="Compact Dashboard (Stacked In/Out - Best for Prints)", variable=self.layout_var, value="horizontal").pack(side="top", anchor="w")
        tk.Radiobutton(layout_frame, text="Single Page (Entire Month on One Sheet)", variable=self.layout_var, value="single").pack(side="top", anchor="w")
        tk.Radiobutton(layout_frame, text="Vertical Report (Individual Teacher Mode)", variable=self.layout_var, value="vertical").pack(side="top", anchor="w")

        # Buttons Frame
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=20)

        # Process Button (Excel)
        self.process_btn = tk.Button(btn_frame, text="Generate Excel", 
                                     command=lambda: self.run_processing("excel"), 
                                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                     padx=15, pady=8)
        self.process_btn.pack(side="left", padx=10)

        # PDF Button
        self.pdf_btn = tk.Button(btn_frame, text="Export to PDF", 
                                 command=lambda: self.run_processing("pdf"), 
                                 bg="#f44336", fg="white", font=("Arial", 11, "bold"),
                                 padx=15, pady=8)
        self.pdf_btn.pack(side="left", padx=10)

        # Status
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = tk.Label(root, textvariable=self.status_var, font=("Arial", 10))
        self.status_label.pack()

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Attendance Log File",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.status_var.set("File selected. Ready to generate.")

    def run_processing(self, export_type="excel"):
        input_file = self.file_path_var.get()
        if input_file == "No file selected" or not input_file:
            messagebox.showwarning("Warning", "Please select an input file first.")
            return

        layout = self.layout_var.get()
        self.status_var.set(f"Generating {layout.capitalize()} {export_type.upper()}...")
        self.root.update_idletasks()

        output_file = "attendance_output.xlsx" if export_type == "excel" else "attendance_output.pdf"
        
        try:
            if export_type == "excel" and os.path.exists(output_file):
                try:
                    with open(output_file, 'a'): pass
                except PermissionError:
                    messagebox.showerror("Error", f"Close '{output_file}' and try again.")
                    self.status_var.set("Error: File is open.")
                    return

            main_process(input_file, export_type, layout)
            self.status_var.set(f"Done! {export_type.upper()} saved.")
            messagebox.showinfo("Success", f"{export_type.capitalize()} dashboard saved as {output_file}")
        except Exception as e:
            self.status_var.set("An error occurred.")
            messagebox.showerror("Error", f"Failed:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
