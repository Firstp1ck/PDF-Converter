import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter

def select_pdf():
    global pdf_path
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, pdf_path)

def select_output():
    global output_path
    output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_path)

def convert_pdf_to_word():
    if pdf_path and output_path:
        try:
            converter = Converter(pdf_path)
            converter.convert(output_path, start=0, end=None)
            converter.close()
            messagebox.showinfo("Success", "PDF converted to Word successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Warning", "Please select both PDF and output file paths.")

# Set up the Tkinter window
window = tk.Tk()
window.title("PDF to Word Converter")

# Variables to hold file paths
pdf_path = ""
output_path = ""

# Create the UI components
pdf_label = tk.Label(window, text="PDF File:")
pdf_label.grid(column=0, row=0, padx=10, pady=10)
pdf_entry = tk.Entry(window, width=50)
pdf_entry.grid(column=1, row=0, padx=10, pady=10)
pdf_button = tk.Button(window, text="Select PDF", command=select_pdf)
pdf_button.grid(column=2, row=0, padx=10, pady=10)

output_label = tk.Label(window, text="Output Word File:")
output_label.grid(column=0, row=1, padx=10, pady=10)
output_entry = tk.Entry(window, width=50)
output_entry.grid(column=1, row=1, padx=10, pady=10)
output_button = tk.Button(window, text="Select Output", command=select_output)
output_button.grid(column=2, row=1, padx=10, pady=10)

convert_button = tk.Button(window, text="Convert to Word", command=convert_pdf_to_word)
convert_button.grid(column=1, row=2, padx=10, pady=10)

# Start the Tkinter event loop
window.mainloop()
