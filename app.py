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
    match = re.search(r"-(\d{2}|[A-Za-z]{2,3})$", part_number)  # Matches either two digits or 2-3 letters at the end
    return match.group(1) if match else ""  # Return the matched finish code or empty string

def clean_shortages(shortage_str):
    shortage_str = re.sub(r"\(.*?\)", "", shortage_str)  # Remove content inside brackets
    shortage_str = shortage_str.replace("shortage", "").strip()
    shortage_str = shortage_str.replace(",", " ").replace("/", " ")
    if "WIP - PRODUCTION" in shortage_str:
        return ""
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
                if file.filename.endswith(".xlsx"):
                    df = pd.read_excel(filepath, engine="openpyxl")
                elif file.filename.endswith(".xls"):
                    df = pd.read_excel(filepath)
                elif file.filename.endswith(".ods"):
                    df = pd.read_excel(filepath, engine="odf")
                else:
                    return "Unsupported file format"
            except Exception as e:
                return f"Error reading file: {str(e)}"
            
            data = df.to_html(classes="table table-striped")
            
            if df.shape[1] >= 7:
                selected_columns = df.iloc[:, [0, 5]].dropna()
                selected_columns.columns = ["SO Number", "Shortages"]
                
                selected_columns["Shortages"] = selected_columns["Shortages"].astype(str).apply(clean_shortages)
                selected_columns = selected_columns.explode("Shortages")
                selected_columns.dropna(inplace=True)
                selected_columns = selected_columns[(selected_columns["Shortages"] != "") & (selected_columns["Shortages"] != "-")]
                
                selected_columns["Finish Code"] = selected_columns["Shortages"].apply(extract_finish_code)
                selected_columns = selected_columns[selected_columns["Finish Code"] != ""]  # Remove rows where Finish Code is empty
                
                # Group by Finish Code and aggregate SO Numbers and Shortages
                final_grouped = selected_columns.groupby("Finish Code").agg({
                    "SO Number": lambda x: ', '.join(map(str, x.unique())),  # Collect unique SO Numbers
                    "Shortages": lambda x: ', '.join(map(str, x.unique()))   # Collect unique Part Codes
                }).reset_index()
                
                shortages_table = selected_columns.to_html(classes="table table-bordered", index=False)
                grouped_table = final_grouped.to_html(classes="table table-bordered", index=False)
            else:
                shortages_table = "<p>Not enough columns in the uploaded file.</p>"
                grouped_table = "<p>Not enough data for grouping.</p>"
            
            return render_template("index.html", table=data, shortages_table=shortages_table, grouped_table=grouped_table)
        else:
            return "Invalid file type. Please upload an Excel file (.xls, .xlsx, .ods)"
    
    return render_template("index.html", table=None, shortages_table=None, grouped_table=None)

if __name__ == "__main__":
    app.run(debug=True)