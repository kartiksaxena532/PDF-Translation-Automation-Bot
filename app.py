# app.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading

# Import the formatting logic from the other file
from docx_formatter import process_document

class DocxFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Formatter")
        self.root.geometry("500x250") # Adjusted height as checkbox is removed
        self.root.resizable(False, False)

        self.filepath = None

        # Style
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10), padding=5)
        style.configure("TLabel", font=("Segoe UI", 10))

        # Main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.select_button = ttk.Button(file_frame, text="Select DOCX File", command=self.select_file)
        self.select_button.pack(side=tk.LEFT, padx=(0, 10))

        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=350)
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Options section has been removed.

        # Action button
        self.format_button = ttk.Button(main_frame, text="Format Document", command=self.start_formatting_thread, state=tk.DISABLED)
        self.format_button.pack(pady=20)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready. Please select a file.")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w", padding=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def select_file(self):
        """Opens a file dialog to select a .docx file."""
        filepath = filedialog.askopenfilename(
            title="Select a DOCX file",
            filetypes=[("Word Documents", "*.docx")]
        )
        if filepath:
            self.filepath = filepath
            filename = os.path.basename(filepath)
            self.file_label.config(text=filename)
            self.format_button.config(state=tk.NORMAL)
            self.status_var.set(f"File selected: {filename}")

    def start_formatting_thread(self):
        """Starts the formatting process in a separate thread to keep the GUI responsive."""
        if not self.filepath:
            messagebox.showerror("Error", "No file selected.")
            return

        self.format_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)
        self.status_var.set("Processing... Please wait.")

        # Run the potentially long-running task in a new thread
        thread = threading.Thread(
            target=self.run_formatting,
            args=(self.filepath,) # Removed the TOC argument
        )
        thread.start()

    def run_formatting(self, filepath):
        """The actual formatting logic that runs in the background thread."""
        try:
            # Removed the TOC argument from the function call
            output_path = process_document(filepath)
            self.on_formatting_complete(output_path)
        except Exception as e:
            self.on_formatting_error(e)

    def on_formatting_complete(self, output_path):
        """Updates the GUI after successful formatting."""
        self.status_var.set("Formatting complete!")
        # Updated success message, removed TOC instructions
        messagebox.showinfo(
            "Success",
            f"Document formatted successfully!\n\nSaved as:\n{output_path}"
        )
        self.reset_ui()

    def on_formatting_error(self, error):
        """Updates the GUI after a formatting error."""
        self.status_var.set("An error occurred.")
        messagebox.showerror("Error", f"An error occurred during formatting:\n\n{error}")
        self.reset_ui()

    def reset_ui(self):
        """Resets the UI to its initial state."""
        self.format_button.config(state=tk.NORMAL if self.filepath else tk.DISABLED)
        self.select_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxFormatterApp(root)
    root.mainloop()