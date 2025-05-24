

from flask import Flask, request, render_template
import requests
import openpyxl
import os

app = Flask(__name__)

# UltraMsg API credentials
INSTANCE_KEY = "instance101363"  # Replace with your UltraMsg instance key
TOKEN = "m2ck8a6cfg7bnhmz"       # Replace with your UltraMsg API token
API_URL = f"https://api.ultramsg.com/{INSTANCE_KEY}/messages/chat"

# Excel file path
EXCEL_FILE = "Student_Data.xlsx"

# Ensure the Excel file exists and has the correct headers
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Student Data"
    sheet.append([
        "Student Name", 
        "Phone Number", 
        "Email", 
        "Parent Name", 
        "Parent Contact", 
        # "Parent Email"
    ])
    workbook.save(EXCEL_FILE)

# Save data to Excel
def save_to_excel(data):
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    sheet.append(data)
    workbook.save(EXCEL_FILE)

# Send WhatsApp message
def send_whatsapp_message(phone_number, message):
    if not phone_number.startswith("+91"):
        phone_number = "+91" + phone_number.lstrip("0")

    payload = {
        "to": phone_number,
        "body": message
    }

    try:
        response = requests.post(f"{API_URL}?token={TOKEN}", json=payload)
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

# Home route
@app.route("/")
def index():
    return render_template("index.html")

# Handle form submission
@app.route("/send_message", methods=["POST"])
def send_message():
    student_name = request.form.get("student_name")
    student_number = request.form.get("student_number")
    student_email = request.form.get("student_email")
    parent_name = request.form.get("parent_name")
    parent_contact = request.form.get("parent_contact")
    # parent_email = request.form.get("parent_email")

    if not all([student_name, student_number, student_email, parent_name, parent_contact]):
        return "Error: All fields are required."

    # Save to Excel
    save_to_excel([
        student_name, 
        student_number, 
        student_email, 
        parent_name, 
        parent_contact, 
        # parent_email
    ])

    # WhatsApp message
    message = f"""
Hello {student_name},


Welcome to Vikrant Group Of Institutions, Indore.

Thank you very much for visiting the campus.

Vikrant Group has been into existence since more then 20 years now, with more than 10,000/- Alumni working in India and abroad. 

VGI is running UG & PG courses in following institutes:- 

Engineering
Management 
Nursing 
Pharmacy
Law

All courses are approved by UGC, AICTE, PCI, BCI and MPNRC and affiliated to state Govt. Universities.

For various Discount and Scholarship Schemes click on the link below:- 
https://www.vitm.edu.in/scholarship.html

Stay updated and connected by following our official Instagram page: [vikrant.indore]
https://www.instagram.com/vikrant.indore

We hope your visit is both enjoyable and inspiring. Feel free to reach out for more details.

Thanks"""

    result = send_whatsapp_message(student_number, message)

    if "error" in result:
        return f"Message not sent: {result['error']}"
    return f"Message sent successfully to {student_name} ({student_number})!"

if __name__ == "__main__":
    app.run(debug=True)
