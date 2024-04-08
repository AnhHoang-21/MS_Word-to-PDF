import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os

def convert_to_pdf():
    # Open file dialog to select input Word document
    input_file = filedialog.askopenfilename(title="Select Word Document", filetypes=[("Word Documents", "*.docx")])
    if not input_file:
        return

    # Get the directory of the input file
    input_folder = os.path.dirname(input_file)

    # Construct output file path in the same folder as input file
    output_file = os.path.join(input_folder, "output.pdf")

    try:
        print(f"Converting '{input_file}' to '{output_file}'...")
        convert(input_file, output_file)
        print("Conversion successful!")
        messagebox.showinfo("Conversion Complete", f"The PDF file has been saved at: {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")
        messagebox.showerror("Conversion Error", f"An error occurred during the conversion: {e}")

# Create the Tkinter window
root = tk.Tk()
root.title("Word to PDF Converter")

# Create the "Convert" button
convert_button = tk.Button(root, text="Convert Word to PDF", command=convert_to_pdf)
convert_button.pack(padx=20, pady=20)

# Start the Tkinter event loop
root.mainloop()
