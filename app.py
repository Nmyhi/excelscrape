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
    match = re.search(r"-\w+$", part_number)
    return match.group(0) if match else part_number

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
                    df = pd.read_excel(filepath)  # xlrd is used automatically
                elif file.filename.endswith(".ods"):
                    df = pd.read_excel(filepath, engine="odf")
                else:
                    return "Unsupported file format"
            except Exception as e:
                return f"Error reading file: {str(e)}"
            
            data = df.to_html(classes="table table-striped")  # Convert to HTML table
            
            # Extract SO Number (Column 2) and Shortages (Column 7) and create new table
            if df.shape[1] >= 7:  # Ensure at least 7 columns exist
                selected_columns = df.iloc[:, [1, 6]].dropna()
                selected_columns.columns = ["SO Number", "Shortages"]
                shortages_table = selected_columns.to_html(classes="table table-bordered", index=False)
                
                # Process shortages to group by finish code
                shortages_expanded = selected_columns.copy()
                shortages_expanded["Shortages"] = shortages_expanded["Shortages"].astype(str).str.split()
                shortages_expanded = shortages_expanded.explode("Shortages")
                shortages_expanded["Finish Code"] = shortages_expanded["Shortages"].apply(extract_finish_code)
                grouped_shortages = shortages_expanded.groupby("Finish Code")["Shortages"].apply(lambda x: ', '.join(x.unique())).reset_index()
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