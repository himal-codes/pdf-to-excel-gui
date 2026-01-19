import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os

def convert_pdf_to_excel():
    # 1. Ask user to select a PDF
    pdf_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    
    if not pdf_path:
        return  # User cancelled

    try:
        # 2. Ask user where to save the Excel file
        excel_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if not excel_path:
            return # User cancelled

        print(f"📄 Processing: {pdf_path}...")
        
        # 3. The Conversion Logic
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table)
                    all_data.append(df)

        # 4. Save to Excel
        if all_data:
            final_df = pd.concat(all_data)
            final_df.to_excel(excel_path, index=False, header=False)
            messagebox.showinfo("Success!", f"✅ Converted successfully!\nSaved to: {excel_path}")
        else:
            messagebox.showwarning("Oops", "⚠️ No tables found in this PDF.")

    except Exception as e:
        messagebox.showerror("Error", f"❌ An error occurred:\n{e}")

# --- GUI SETUP ---
root = tk.Tk()
root.title("PDF to Excel Converter")
root.geometry("300x150")

label = tk.Label(root, text="Himal's PDF Converter", font=("Arial", 12, "bold"))
label.pack(pady=10)

btn = tk.Button(root, text="Select PDF to Convert", command=convert_pdf_to_excel, bg="green", fg="white", font=("Arial", 10))
btn.pack(pady=20)

root.mainloop()
