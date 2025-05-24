

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

Welcome to Vikrant Institute of Technology and Management (VITM), Indore!*

Welcome to visit our campus. At VITM, we are committed to nurturing innovative minds and fostering excellence in education. 
Our state-of-the-art facilities, experienced faculty, and vibrant campus life make us a premier destination for aspiring professionals.

Explore more about our college on our official website: http://vitm.edu.in

Stay updated and connected by following our official Instagram page: [vikrant.indore](https://www.instagram.com/vikrant.indore)

We hope your visit is both enjoyable and inspiring. Feel free to reach out if you need any assistance!"""

    result = send_whatsapp_message(student_number, message)

    if "error" in result:
        return f"Message not sent: {result['error']}"
    return f"Message sent successfully to {student_name} ({student_number})!"

if __name__ == "__main__":
    app.run(debug=True)
