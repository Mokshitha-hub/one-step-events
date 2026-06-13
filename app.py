from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from uuid import uuid4
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

app = Flask(__name__)

FILE_NAME = "event_enquiries.xlsx"


def load_enquiries():
    schema_changed = False

    if os.path.exists(FILE_NAME):
        df = pd.read_excel(FILE_NAME)
    else:
        df = pd.DataFrame(columns=[
            "Enquiry ID",
            "Full Name",
            "Mobile Number",
            "Email",
            "Expected Date",
            "Event Type",
            "Budget",
            "Vision",
            "Status",
        ])

    if "Enquiry ID" not in df.columns:
        df["Enquiry ID"] = [str(uuid4()) for _ in range(len(df))]
        schema_changed = True

    if "Status" not in df.columns:
        df["Status"] = "Pending"
        schema_changed = True

    df["Status"] = df["Status"].fillna("Pending")

    if schema_changed:
        save_enquiries(df)

    return df


def save_enquiries(df):
    df.to_excel(FILE_NAME, index=False)

    wb = load_workbook(FILE_NAME)
    ws = wb.active

    if "EventEnquiries" in ws.tables:
        del ws.tables["EventEnquiries"]

    end_row = ws.max_row
    table_range = f"A1:I{end_row}"
    tab = Table(displayName="EventEnquiries", ref=table_range)

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)
    wb.save(FILE_NAME)

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

    df = load_enquiries()
    enquiries = df.to_dict(orient="records")
    approved_count = int((df["Status"] == "Approved").sum())
    pending_count = int((df["Status"] != "Approved").sum())

    return render_template(
        "dashboard.html",
        enquiries=enquiries,
        approved_count=approved_count,
        pending_count=pending_count
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

    df = load_enquiries()
    new_row = pd.DataFrame([{
        "Enquiry ID": str(uuid4()),
        "Full Name": full_name,
        "Mobile Number": mobile,
        "Email": email,
        "Expected Date": expected_date,
        "Event Type": event_type,
        "Budget": budget,
        "Vision": vision,
        "Status": "Pending",
    }])

    df = pd.concat([df, new_row], ignore_index=True)
    save_enquiries(df)
    
    return render_template("success.html")


@app.route("/update-status", methods=["POST"])
def update_status():
    enquiry_id = request.form["enquiry_id"]
    new_status = "Approved" if request.form.get("approve_event") == "yes" else "Pending"

    df = load_enquiries()
    df.loc[df["Enquiry ID"].astype(str) == str(enquiry_id), "Status"] = new_status
    save_enquiries(df)

    return redirect(url_for("dashboard"))



if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5002, debug=True)