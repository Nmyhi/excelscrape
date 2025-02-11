from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {"xls", "xlsx", "ods"}  # Allowed file types

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

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
            return render_template("index.html", table=data)
        else:
            return "Invalid file type. Please upload an Excel file (.xls, .xlsx, .ods)"
    
    return render_template("index.html", table=None)

if __name__ == "__main__":
    app.run(debug=True)
