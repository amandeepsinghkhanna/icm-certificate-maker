import os
import re
import docx
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
    student_info_df = pd.read_excel(student_info_filepath, engine="openpyxl")
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


# User-defined function to replace the keywords in the docx file:
def fill_certificate(doc_obj, replacement_keywords_dict):
    """
    Replaces placeholders in a Word document (.docx) with values from a provided dictionary.

    This function iterates through all sections and paragraphs in the header of the document,
    searching for text patterns specified in the `replacement_keywords_dict`. It then replaces
    these patterns with the corresponding values from the dictionary.

    Args:
        doc_obj (docx.Document): The Word document object to be modified.
        replacement_keywords_dict (dict): A dictionary where keys are regular expression patterns
                                          and values are the corresponding replacement strings.

    Returns:
        docx.Document: The modified Word document object with placeholders replaced.
    """
    for section in doc_obj.sections:
        header = section.header
        for paragraph in header.paragraphs:
            for (
                replacement_pattern,
                replacement_value,
            ) in replacement_keywords_dict.items():
                paragraph.text = re.sub(
                    replacement_pattern,
                    replacement_value,
                    paragraph.text,
                    flags=re.IGNORECASE,
                )
    return doc_obj


def create_certificates(
    certificate_template, student_info_df, replacement_keywords_dict, output_dir="./"
):
    student_info_lst = student_info_df.to_dict(orient="records")
    for student_info_dict in student_info_lst:
        working_certificate_copy = certificate_template
        for key in replacement_keywords_dict.keys():
            replacement_keywords_dict[key] = str(student_info_dict[key.lower()])
        fill_certificate(
            doc_obj=working_certificate_copy,
            replacement_keywords_dict=replacement_keywords_dict,
        )
        certificate_filepath = os.path.join(
            output_dir, f"{student_info_dict['student_name']}.docx"
        )
        working_certificate_copy.save(certificate_filepath)


# User-defined function to read the certificate temple stored as a ".docx" file:
def read_certificate_template(certificate_template_path):
    """
    Reads the certificate template stored as a word document ".docx" file format.

    Arguments:
    1. certificate_template_path - str - Path to the ".docx" file with the certificate
        template.
    """
    # Reading the ".docx" file with the certificate template:
    with open(certificate_template_path, "rb") as word_file:
        certificate_template = docx.Document(word_file)
    return certificate_template


def start_certificate_maker():
    student_info_df = read_student_info(
        student_info_filepath=student_info_filepath_var.get(), req_cols=REQ_COLS
    )
    certificate_template = read_certificate_template(
        certificate_template_path=certificate_template_filepath_var.get()
    )
    create_certificates(
        certificate_template=certificate_template,
        student_info_df=student_info_df,
        replacement_keywords_dict=REPLACEMENT_KEYWORDS_DICT,
        output_dir=output_dir_var.get()
    )
    tk.messagebox.showinfo("Process Notifier", "The creation of certification is complete")


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
