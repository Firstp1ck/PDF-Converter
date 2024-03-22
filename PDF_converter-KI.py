import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document

# Pfad zu Tesseract-OCR - an Ihr System anpassen
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def pdf_to_word(pdf_path, word_path):
    document = Document()
    pdf = fitz.open(pdf_path)

    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = pytesseract.image_to_string(img, lang='deu')
        document.add_paragraph(text)

    document.save(word_path)
    pdf.close()
    messagebox.showinfo("Erfolg", "Die Konvertierung wurde erfolgreich abgeschlossen!")

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_path_entry.delete(0, tk.END)
        pdf_path_entry.insert(0, file_path)

def select_word():
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
    if file_path:
        word_path_entry.delete(0, tk.END)
        word_path_entry.insert(0, file_path)

def convert_button_action():
    pdf_path = pdf_path_entry.get()
    word_path = word_path_entry.get()
    if pdf_path and word_path:
        pdf_to_word(pdf_path, word_path)
    else:
        messagebox.showwarning("Warnung", "Bitte wählen Sie sowohl eine PDF- als auch eine Word-Datei aus.")

# GUI
root = tk.Tk()
root.title("PDF zu Word Konverter")

tk.Label(root, text="PDF Datei:").pack(padx=10, pady=5)
pdf_path_entry = tk.Entry(root, width=50)
pdf_path_entry.pack(padx=10, pady=5)
tk.Button(root, text="Datei auswählen", command=select_pdf).pack(padx=10, pady=5)

tk.Label(root, text="Speichern als Word-Datei:").pack(padx=10, pady=5)
word_path_entry = tk.Entry(root, width=50)
word_path_entry.pack(padx=10, pady=5)
tk.Button(root, text="Speicherort wählen", command=select_word).pack(padx=10, pady=5)

tk.Button(root, text="Konvertieren", command=convert_button_action).pack(padx=10, pady=20)

root.mainloop()
