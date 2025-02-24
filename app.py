from flask import Flask, render_template, request
import pandas as pd
import os
import re

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {"xls", "xlsx", "ods"}  # Allowed file types

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_finish_code(part_number):
    """
    Extracts the finish code from a part number.
    Assumes the finish code is in the format '-XX' (e.g., '-01').
    """
    match = re.search(r"-\d{2}$", part_number)  # Matches a dash followed by two digits at the end
    return match.group(0) if match else part_number

def clean_shortages(shortage_str):
    """
    Cleans the shortages string by:
    - Removing content in brackets
    - Removing the word 'shortage'
    - Replacing commas and slashes with spaces
    - Ignoring 'WIP - PRODUCTION'
    """
    shortage_str = re.sub(r"\(.*?\)", "", shortage_str)  # Remove content inside brackets
    shortage_str = shortage_str.replace("shortage", "").strip()  # Remove the word 'shortage'
    shortage_str = shortage_str.replace(",", " ").replace("/", " ")  # Replace commas and slashes with spaces
    if "WIP - PRODUCTION" in shortage_str:
        return ""  # Ignore WIP - PRODUCTION entries
    return shortage_str.split()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file part"
        file = request.files["file"]
        if file.filename == "":
            return "No selected file"
        if file and allowed_file(file.filename):
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
            file.save(filepath)
            
            try:
                # Read the file based on extension
                if file.filename.endswith(".xlsx"):
                    df = pd.read_excel(filepath, engine="openpyxl")
                elif file.filename.endswith(".xls"):
                    df = pd.read_excel(filepath)  # xlrd is used automatically
                elif file.filename.endswith(".ods"):
                    df = pd.read_excel(filepath, engine="odf")
                else:
                    return "Unsupported file format"
            except Exception as e:
                return f"Error reading file: {str(e)}"
            
            # Convert to HTML table for display
            data = df.to_html(classes="table table-striped")
            
            # Ensure there are at least 7 columns before proceeding
            if df.shape[1] >= 7:
                selected_columns = df.iloc[:, [0, 5]].dropna()  # Extract SO Number (1st col) & Shortages (6th col)
                selected_columns.columns = ["SO Number", "Shortages"]
                
                # Process shortages
                selected_columns["Shortages"] = selected_columns["Shortages"].astype(str).apply(clean_shortages)
                selected_columns = selected_columns.explode("Shortages")  # Expand shortages into separate rows
                selected_columns.dropna(inplace=True)  # Remove any NaN values

                # Filter out empty rows and those containing "Shortage" or "-"
                selected_columns = selected_columns[(selected_columns["Shortages"] != "") & (selected_columns["Shortages"] != "-")]

                # Extract Finish Code
                selected_columns["Finish Code"] = selected_columns["Shortages"].apply(extract_finish_code)
                
                # Group by Finish Code and collect associated SO Numbers
                grouped_shortages = selected_columns.groupby("Finish Code")["SO Number"].apply(lambda x: ', '.join(map(str, x.unique()))).reset_index()
                
                # Convert tables to HTML for display
                shortages_table = selected_columns.to_html(classes="table table-bordered", index=False)
                grouped_table = grouped_shortages.to_html(classes="table table-bordered", index=False)
            else:
                shortages_table = "<p>Not enough columns in the uploaded file.</p>"
                grouped_table = "<p>Not enough data for grouping.</p>"

            return render_template("index.html", table=data, shortages_table=shortages_table, grouped_table=grouped_table)
        else:
            return "Invalid file type. Please upload an Excel file (.xls, .xlsx, .ods)"
    
    return render_template("index.html", table=None, shortages_table=None, grouped_table=None)

if __name__ == "__main__":
    app.run(debug=True)
