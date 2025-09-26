from pdf2docx import Converter
import os
from tkinter import filedialog, messagebox
import tkinter as tk

def convert_pdf_to_word(pdf_path, output_folder):
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Generate output file path
    output_path = os.path.join(output_folder, os.path.splitext(os.path.basename(pdf_path))[0] + ".docx")

    # Convert PDF to Word
    cv = Converter(pdf_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()
    

def batch_convert_pdfs_to_word(input_folder, output_folder):
    
    # Loop through all PDF files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)
            convert_pdf_to_word(pdf_path, output_folder)

if __name__ == "__main__":
    # Specify your input and output folders
    # Show message to choose the destination folder
    messagebox.showinfo("Επιλογή φακέλου", "Επιλέξτε τον φάκελο που υπαρχουν τα αρχεία PDF.")

    # Prompt user to select the destination folder
    root = tk.Tk()
    root.withdraw()
    input_folder = filedialog.askdirectory()
    messagebox.showinfo("Επιλογή φακέλου", "Επιλέξτε τον φάκελο που θα αποθηκευτούν τα αρχεία Docx.")
    output_folder = filedialog.askdirectory()
    messagebox.showinfo("Eπεξεργασία","Παρακαλώ περιμένετε...")
    # Convert PDFs to Word
    batch_convert_pdfs_to_word(input_folder, output_folder)
    root.destroy()
    messagebox.showinfo("Επιτυχία","Η διαδικασία Ολοκληρώθηκε.")