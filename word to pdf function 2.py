import os
from docx2pdf import convert
from tkinter import Tk, filedialog

# Initialize tkinter
root = Tk()
root.withdraw()  # Hide the main window

# Ask user to select a Word document file
file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])

if file_path:
    # Extract directory and filename from the selected file path
    directory, filename = os.path.split(file_path)

    try:
        # Construct the output PDF file path
        pdf_output_path = os.path.join(directory, os.path.splitext(filename)[0] + ".pdf")

        # Convert the Word document to PDF
        convert(file_path, pdf_output_path)

        print("Conversion complete. PDF saved as:", pdf_output_path)
    except Exception as e:
        print("An error occurred:", e)
else:
    print("No file selected. Conversion aborted.")