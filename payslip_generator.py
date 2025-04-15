import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Create a directory for payslips if it doesn't exist
if not os.path.exists("payslips"):
    os.makedirs("payslips")

# Load the Excel file
try:
    df = pd.read_excel("employees.xlsx")
except FileNotFoundError:
    print("⚠️ Error: employees.xlsx file not found!")
    exit()

# Loop through each employee and generate payslip
for index, row in df.iterrows():
    emp_id = str(row["Employee ID"])
    name = row["Name"]
    email = row["Email"]
    basic = row["Basic Salary"]
    allowance = row["Allowances"]
    deduction = row["Deductions"]
    net_salary = basic + allowance - deduction

    # Create PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="Monthly Payslip", ln=True, align='C')
    pdf.set_font("Arial", "", 12)
    pdf.cell(200, 10, txt=f"Employee ID: {emp_id}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${basic:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${allowance:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${deduction:.2f}", ln=True)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, txt=f"Net Salary: ${net_salary:.2f}", ln=True)
    payslip_path = f"payslips/{emp_id}.pdf"
    pdf.output(payslip_path)

    print(f"✅ Payslip created for {name} ({emp_id})")

    # Send Email
    try:
        yag = yagmail.SMTP(user="tatendaleeroy64@gmail.com", password="vcya pmum bisw fcda")
        yag.send(
            to=email,
            subject="Your Payslip for This Month",
            contents=f"Hello {name},\n\nPlease find attached your payslip for this month.\n\nRegards,\nHR Team",
            attachments=payslip_path
        )
        print(f"📧 Email sent to {email}")
    except Exception as e:
        print(f"❌ Failed to send email to {email}: {e}")
