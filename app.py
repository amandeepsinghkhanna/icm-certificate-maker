import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

REQ_COLS = ["student_id", "student_name", "student_campus", "student_status"]
REPLACEMENT_KEYWORDS_DICT = {
    "STUDENT_NAME": "",
    "STUDENT_ID": "",
    "STUDENT_CAMPUS": "",
    "STUDENT_STATUS": "",
}


# User defined function to read and validate the student information file:
def read_student_info(student_info_filepath, req_cols):
    """
    Reads a ".xlsx" file with the student information, checks if the required
    columns exist in the file and returns the file as a pandas DataFrame.

    Arguments:
        1. student_info_filepath - str - Path to the ".csv" file with the student information.
        2. req_cols - list - List of str values representing the mandatory columns form the
            student information file.

    Returns:
        1. student_info_df - pandas.core.DataFrame - The student information as a pandas DataFrame.
    """
    # Reading the student information file:
    student_info_df = pd.read_excel(student_info_filepath)
    # Standardisation of the column names:
    student_info_file_cols = [
        f.lower().strip().replace(" ", "_") for f in student_info_df.columns
    ]
    # Checking for missing columns in the files:
    missing_cols = set(req_cols) - set(student_info_file_cols)
    if len(missing_cols) == 0:  # If no columns are missing.
        return student_info_df
    else:  # If at-least one column is missing.
        raise ValueError(f"The columns {missing_cols} are missing.")


def select_student_info_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    student_info_filepath_var.set(filepath)


def select_certificate_template_file():
    filepath = filedialog.askopenfilename(filetypes=[("DOCX Files", "*.docx")])
    certificate_template_filepath_var.set(filepath)


def browse_output_dir():
    directory = filedialog.askdirectory()
    output_dir_var.set(directory)


def start_certificate_maker():
    student_info_df = read_student_info(
        student_info_filepath=certificate_template_filepath_var.get(), req_cols=REQ_COLS
    )
    tk.messagebox.show("Student Info file Processing is complete. Wait for the next pop-up to appear!")


app = tk.Tk()
student_info_filepath_var = tk.StringVar()
certificate_template_filepath_var = tk.StringVar()
output_dir_var = tk.StringVar()
app.title("ICM Africa Certificate Maker")
app.geometry("600x400")
tk.Label(app, text="Select Student Info File:").pack()
tk.Button(app, text="Browse", command=select_student_info_file).pack()
tk.Label(app, textvariable=student_info_filepath_var).pack()
tk.Label(app, text="Select Certificate Template File:").pack()
tk.Button(app, text="Browse", command=select_certificate_template_file).pack()
tk.Label(app, textvariable=certificate_template_filepath_var).pack()
tk.Label(app, text="Select Output Folder:").pack()
tk.Button(app, text="Browse", command=browse_output_dir).pack()
tk.Label(app, textvariable=output_dir_var).pack()
tk.Button(app, text="Generate Certificates", command=start_certificate_maker).pack()
app.mainloop()
