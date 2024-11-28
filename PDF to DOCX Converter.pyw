import os
import customtkinter as ctk
from tkinter import filedialog
from pdf2docx import Converter
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES  # Drag-and-Drop support

# Configuration for CustomTkinter (look and feel)
ctk.set_appearance_mode("Light")  # Options: "Light", "Dark", or "System"
ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

# Function to select the PDF file
def select_pdf():
    pdf_file = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if pdf_file:
        pdf_path.set(pdf_file)

# Function to select the save location for the DOCX file
def select_save_location():
    docx_file = filedialog.asksaveasfilename(
        title="Choose save location",
        defaultextension=".docx",
        filetypes=[("Word Files", "*.docx")]
    )
    if docx_file:
        docx_path.set(docx_file)

# Drag-and-Drop function
def on_drop(event):
    pdf_path.set(event.data.strip("{}"))  # Removes braces for paths with spaces

# Function to convert PDF to DOCX
def convert_pdf_to_docx():
    pdf = pdf_path.get()
    docx = docx_path.get()
    
    if not os.path.exists(pdf):
        result_label.configure(text="PDF file not found.", text_color="red")
        return
    
    # If no save location is selected, use the same path and name as the PDF file
    if not docx:
        docx = os.path.splitext(pdf)[0] + ".docx"
    
    try:
        converter = Converter(pdf)
        converter.convert(docx)
        converter.close()  # Close the converter after completion
        
        if os.path.exists(docx):
            result_label.configure(text="Conversion completed!", text_color="green")
        else:
            result_label.configure(text="Conversion failed.", text_color="red")
    except Exception as e:
        result_label.configure(text=f"Error: {str(e)}", text_color="red")

# Main application with Drag-and-Drop support
app = TkinterDnD.Tk()  # Use TkinterDnD instead of ctk.CTk for Drag-and-Drop support
app.title("PDF to DOCX Converter")
app.geometry("400x350")

# Variables to store file paths
pdf_path = tk.StringVar()
docx_path = tk.StringVar()

# User Interface
label_pdf = ctk.CTkLabel(app, text="PDF File:")
label_pdf.pack(pady=10)

entry_pdf = ctk.CTkEntry(app, textvariable=pdf_path, width=300)
entry_pdf.pack(pady=5)

# Drag-and-Drop area
entry_pdf.drop_target_register(DND_FILES)
entry_pdf.dnd_bind('<<Drop>>', on_drop)

button_pdf = ctk.CTkButton(app, text="Select PDF", command=select_pdf)
button_pdf.pack(pady=5)

label_docx = ctk.CTkLabel(app, text="DOCX File Save Location:")
label_docx.pack(pady=10)

entry_docx = ctk.CTkEntry(app, textvariable=docx_path, width=300)
entry_docx.pack(pady=5)

button_docx = ctk.CTkButton(app, text="Choose Save Location", command=select_save_location)
button_docx.pack(pady=5)

convert_button = ctk.CTkButton(app, text="Convert", command=convert_pdf_to_docx)
convert_button.pack(pady=20)

result_label = ctk.CTkLabel(app, text="")
result_label.pack(pady=10)

# Start the GUI
app.mainloop()
