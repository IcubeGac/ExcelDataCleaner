import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import os
from rapidfuzz import fuzz, process

# List of correct port names
correct_port_names = [
    "CHENNAI", "ENNORE", "KARAIKAL", "KATTUPPALI", "DHAMRA", "HALDIA", "KAKINADA", 
    "KOLKATA", "KRISHNAPATNAM", "PARADIP", "TUTICORIN", "VIZAG", "GANGAVARAM", 
    "COCHIN", "BEDI", "BHAVNAGAR", "SIKKA", "VADINAR", "MULDWARKA", "OKHA", 
    "PORBANDAR", "SALAYA", "SIKKA", "VADINAR", "JNPT", "JAKHAU", "KANDLA", 
    "NAVLAKHI", "SANGHI", "MORMUGAO", "REDI", "PANJIM", "RANPAR", "MUMBAI", 
    "DABHOL", "DHARAMTAR", "DIGHI", "HAJIBUNDER", "JAIGAD", "REVDANDA", 
    "MUNDRA", "NEW MANGALORE", "KARWAR", "PIPAVAV", "DAHEJ", "HAZIRA", "MAGDALLA"
]

# Data cleaning functions
def remove_duplicates(df):
    return df.drop_duplicates()

def fill_missing_values(df, value=0):
    return df.fillna(value)

def drop_missing_values(df):
    return df.dropna()

def convert_dates_to_format(df, column, date_format="%Y-%m-%d"):
    df[column] = pd.to_datetime(df[column], errors='coerce').dt.strftime(date_format)
    return df

def replace_values(df, column, to_replace, value):
    df[column] = df[column].replace(to_replace, value)
    return df

def correct_port_name(port_name, correct_port_names):
    best_match = process.extractOne(port_name, correct_port_names, scorer=fuzz.ratio)
    if best_match and best_match[1] > 80:  # Set a threshold for match confidence
        return best_match[0]
    else:
        return port_name

def convert_text_to_uppercase(df):
    return df.applymap(lambda x: x.upper() if isinstance(x, str) else x)

def remove_leading_trailing_spaces(df, column):
    df[column] = df[column].str.strip()
    return df

def combine_import_export_sheets(import_df, export_df):
    import_df['Source'] = 'import'
    export_df['Source'] = 'export'
    combined_df = pd.concat([import_df, export_df], ignore_index=True)
    return combined_df

class DataCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Cleaner")
        self.root.geometry("600x400")
        self.root.configure(bg='light gray')

        self.file_path = None
        self.sheets = None
        self.current_sheet = None
        self.df = None

        self.create_widgets()

    def create_widgets(self):
        self.upload_btn = tk.Button(self.root, text="Upload Excel File", command=self.upload_file, width=20, height=2, bg='light blue', fg='black', font=('Arial', 12, 'bold'))
        self.upload_btn.pack(pady=10)

        self.sheet_label = tk.Label(self.root, text="Select Sheet", bg='light gray', font=('Arial', 12))
        self.sheet_label.pack(pady=5)

        self.sheet_var = tk.StringVar(self.root)
        self.sheet_menu = tk.OptionMenu(self.root, self.sheet_var, ())
        self.sheet_menu.configure(bg='light blue', fg='black', font=('Arial', 10))
        self.sheet_menu.pack(pady=5)
        self.sheet_var.trace("w", self.change_sheet)

        self.cleaning_frame = tk.Frame(self.root, bg='light gray')
        self.cleaning_frame.pack(pady=10)

        self.duplicate_btn = tk.Button(self.cleaning_frame, text="Remove Duplicates", command=self.apply_remove_duplicates, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.duplicate_btn.grid(row=0, column=0, padx=5, pady=5)

        self.fill_na_btn = tk.Button(self.cleaning_frame, text="Fill Missing Values", command=self.apply_fill_missing_values, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.fill_na_btn.grid(row=0, column=1, padx=5, pady=5)

        self.drop_na_btn = tk.Button(self.cleaning_frame, text="Drop Missing Values", command=self.apply_drop_missing_values, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.drop_na_btn.grid(row=0, column=2, padx=5, pady=5)

        self.convert_datetime_btn = tk.Button(self.cleaning_frame, text="Convert Dates", command=self.apply_convert_dates_to_format, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.convert_datetime_btn.grid(row=1, column=0, padx=5, pady=5)

        self.replace_values_btn = tk.Button(self.cleaning_frame, text="Replace Values", command=self.apply_replace_values, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.replace_values_btn.grid(row=1, column=1, padx=5, pady=5)

        self.correct_port_name_btn = tk.Button(self.cleaning_frame, text="Correct Port Names", command=self.apply_correct_port_name, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.correct_port_name_btn.grid(row=1, column=2, padx=5, pady=5)

        self.uppercase_btn = tk.Button(self.cleaning_frame, text="Convert Text to Uppercase", command=self.apply_convert_text_to_uppercase, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.uppercase_btn.grid(row=2, column=0, padx=5, pady=5)

        self.strip_spaces_btn = tk.Button(self.cleaning_frame, text="Remove Leading/Trailing Spaces", command=self.apply_remove_leading_trailing_spaces, width=20, height=2, bg='light green', fg='black', font=('Arial', 10))
        self.strip_spaces_btn.grid(row=2, column=1, padx=5, pady=5)

        self.download_btn = tk.Button(self.root, text="Download Cleaned File", command=self.download_file, width=20, height=2, bg='light blue', fg='black', font=('Arial', 12, 'bold'))
        self.download_btn.pack(pady=10)

        # Frame for Import/Export Sheets
        self.import_export_frame = tk.Frame(self.root, bg='light gray')
        self.import_export_frame.pack(pady=10)
        self.combine_btn = tk.Button(self.import_export_frame, text="Combine 'IMPORT' and 'EXPORT' Sheets", command=self.combine_import_export_sheets, width=30, height=2, bg='light yellow', fg='black', font=('Arial', 10, 'bold'))
        self.combine_btn.pack(pady=5)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            self.sheets = pd.read_excel(self.file_path, sheet_name=None)
            self.sheet_var.set('')
            self.sheet_menu['menu'].delete(0, 'end')
            for sheet in self.sheets.keys():
                self.sheet_menu['menu'].add_command(label=sheet, command=tk._setit(self.sheet_var, sheet))
            self.sheet_var.set(next(iter(self.sheets)))  # Set to the first sheet
            messagebox.showinfo("File Upload", f"Successfully uploaded {os.path.basename(self.file_path)}")
            self.check_import_export_sheets()

    def change_sheet(self, *args):
        sheet_name = self.sheet_var.get()
        if sheet_name:
            self.current_sheet = sheet_name
            self.df = self.sheets[sheet_name]
            self.check_import_export_sheets()  # Ensure the button visibility is updated

    def apply_remove_duplicates(self):
        if self.df is not None:
            self.df = remove_duplicates(self.df)
            self.sheets[self.current_sheet] = self.df
            messagebox.showinfo("Data Cleaning", "Duplicates removed successfully!")

    def apply_fill_missing_values(self):
        if self.df is not None:
            self.df = fill_missing_values(self.df)
            self.sheets[self.current_sheet] = self.df
            messagebox.showinfo("Data Cleaning", "Missing values filled successfully!")

    def apply_drop_missing_values(self):
        if self.df is not None:
            self.df = drop_missing_values(self.df)
            self.sheets[self.current_sheet] = self.df
            messagebox.showinfo("Data Cleaning", "Rows with missing values dropped successfully!")

    def apply_convert_dates_to_format(self):
        if self.df is not None:
            date_column = simpledialog.askstring("Input", "Enter the date column name:")
            date_format = simpledialog.askstring("Input", "Enter the desired date format (e.g., %Y-%m-%d):", initialvalue="%Y-%m-%d")
            if date_column in self.df.columns:
                self.df = convert_dates_to_format(self.df, date_column, date_format)
                self.sheets[self.current_sheet] = self.df
                messagebox.showinfo("Data Cleaning", f"Dates converted to format {date_format} successfully!")
            else:
                messagebox.showerror("Error", "Invalid column name!")

    def apply_replace_values(self):
        if self.df is not None:
            column = simpledialog.askstring("Input", "Enter the column name:")
            to_replace = simpledialog.askstring("Input", "Enter the value to replace:")
            value = simpledialog.askstring("Input", "Enter the replacement value:")
            if column in self.df.columns:
                self.df = replace_values(self.df, column, to_replace, value)
                self.sheets[self.current_sheet] = self.df
                messagebox.showinfo("Data Cleaning", f"Values replaced successfully!")
            else:
                messagebox.showerror("Error", "Invalid column name!")

    def apply_correct_port_name(self):
        if self.df is not None:
            column = simpledialog.askstring("Input", "Enter the port column name:")
            if column in self.df.columns:
                self.df[column] = self.df[column].apply(correct_port_name, correct_port_names=correct_port_names)
                self.sheets[self.current_sheet] = self.df
                messagebox.showinfo("Data Cleaning", "Port names corrected successfully!")
            else:
                messagebox.showerror("Error", "Invalid column name!")

    def apply_convert_text_to_uppercase(self):
        if self.df is not None:
            self.df = convert_text_to_uppercase(self.df)
            self.sheets[self.current_sheet] = self.df
            messagebox.showinfo("Data Cleaning", "Text converted to uppercase successfully!")

    def apply_remove_leading_trailing_spaces(self):
        if self.df is not None:
            column = simpledialog.askstring("Input", "Enter the column name:")
            if column in self.df.columns:
                self.df = remove_leading_trailing_spaces(self.df, column)
                self.sheets[self.current_sheet] = self.df
                messagebox.showinfo("Data Cleaning", "Leading/trailing spaces removed successfully!")
            else:
                messagebox.showerror("Error", "Invalid column name!")

    def combine_import_export_sheets(self):
        if 'IMPORT' in self.sheets and 'EXPORT' in self.sheets:
            import_df = self.sheets['IMPORT']
            export_df = self.sheets['EXPORT']
            combined_df = combine_import_export_sheets(import_df, export_df)
            self.sheets['Combined'] = combined_df
            self.df = combined_df
            self.sheet_var.set('Combined')
            self.sheet_menu['menu'].add_command(label='Combined', command=tk._setit(self.sheet_var, 'Combined'))
            messagebox.showinfo("Sheet Combination", "Import and Export sheets combined successfully!")
        else:
            messagebox.showerror("Error", "Both 'IMPORT' and 'EXPORT' sheets must be present to combine.")

    def check_import_export_sheets(self):
        if 'IMPORT' in self.sheets and 'EXPORT' in self.sheets:
            self.combine_btn.pack()
        else:
            self.combine_btn.pack_forget()

    def download_file(self):
        if self.file_path and self.sheets:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if file_path:
                with pd.ExcelWriter(file_path) as writer:
                    for sheet_name, df in self.sheets.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Download", "Cleaned file downloaded successfully!")
            else:
                messagebox.showwarning("Download", "No file path selected!")
        else:
            messagebox.showwarning("Download", "No file to download!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataCleanerApp(root)
    root.mainloop()
