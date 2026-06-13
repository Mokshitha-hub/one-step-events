from flask import Flask, render_template, request, redirect
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/about")
def about():
    return render_template("about.html")

@app.route("/services")
def services():
    return render_template("services.html")

@app.route("/contact")
def contact():
    return render_template("contact.html")

@app.route("/dashboard")
def dashboard():

    file_name = "event_enquiries.xlsx"

    if os.path.exists(file_name):
        df = pd.read_excel(file_name)
        enquiries = df.to_dict(orient="records")
    else:
        enquiries = []

    return render_template(
        "dashboard.html",
        enquiries=enquiries
    )
@app.route("/submit-enquiry", methods=["POST"])
def submit_enquiry():

    full_name = request.form["full_name"]
    mobile = request.form["mobile"]
    email = request.form["email"]
    expected_date = request.form["expected_date"]
    event_type = request.form["event_type"]
    budget = request.form["budget"]
    vision = request.form["vision"]

    enquiry = {
        "Full Name": full_name,
        "Mobile Number": mobile,
        "Email": email,
        "Expected Date": expected_date,
        "Event Type": event_type,
        "Vision": vision
    }

    file_name = "event_enquiries.xlsx"
    
    new_row = pd.DataFrame([{
    "Full Name": full_name,
    "Mobile Number": mobile,
    "Email": email,
    "Expected Date": expected_date,
    "Event Type": event_type,
    "Budget": budget,
    "Vision": vision
}])
    if os.path.exists(file_name):
        df = pd.read_excel(file_name)
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row

    df.to_excel(file_name, index=False)

    wb = load_workbook(file_name)
    ws = wb.active

    end_row = ws.max_row
    end_col = ws.max_column

    table_range = f"A1:G{end_row}"
    tab = Table(displayName="EventEnquiries", ref=table_range)
    
    style = TableStyleInfo(
    name="TableStyleMedium9",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)
    wb.save(file_name)
    
    return render_template("success.html")



if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5002, debug=True)