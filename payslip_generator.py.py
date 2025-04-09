import pandas as pd
from fpdf import FPDF
import yagmail
import os

# Load Gmail credentials (using environment variables or external configuration file is more secure)
GMAIL_USER = os.getenv("GMAIL_USER", "vdumbatsuro2@gmail.com")  # Replace with your Gmail address
GMAIL_PASS = os.getenv("GMAIL_PASS", "rzdefvqfjqhjnpku")    # Replace with your App Password (if using 2FA)

# Debugging: Check if credentials are correctly set
if not GMAIL_USER or not GMAIL_PASS:
    print("❌ Gmail credentials not set. Please check your script.")
    exit()
else:
    print("✅ Gmail credentials loaded successfully.")

# === STEP 1: Read employee data from Excel ===\;mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
def read_employee_data(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()  # ✅ Remove extra spaces in column names
        print("✔ Employee data loaded successfully.")
        print("🔍 Columns found:", df.columns.tolist())  # Helpful for debugging
        return df
    except FileNotFoundError:
        print("❌ Excel file not found. Check your path.")
        exit()
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        exit()

# === STEP 2: Calculate net salary ===
def calculate_net_salary(basic_salary, allowances, deductions):
    return basic_salary + allowances - deductions

# === STEP 3: Generate payslip PDF ===
def generate_payslip(employee_id, name, basic_salary, allowances, deductions, net_salary):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, f"Payslip for {name}", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", '', 12)
    pdf.cell(200, 10, f"Employee ID: {employee_id}", ln=True)
    pdf.cell(200, 10, f"Name: {name}", ln=True)
    pdf.cell(200, 10, f"Basic Salary: ${basic_salary:.2f}", ln=True)
    pdf.cell(200, 10, f"Allowances: ${allowances:.2f}", ln=True)
    pdf.cell(200, 10, f"Deductions: ${deductions:.2f}", ln=True)
    pdf.cell(200, 10, f"Net Salary: ${net_salary:.2f}", ln=True)

    os.makedirs("payslips", exist_ok=True)
    file_path = f"payslips/{employee_id}.pdf"
    pdf.output(file_path)
    return file_path

# === STEP 4: Send payslip via email ===
def send_email(recipient_email, file_path):
    try:
        yag = yagmail.SMTP(GMAIL_USER, GMAIL_PASS)
        subject = "Your Payslip for This Month"
        body = "Hi, please find your payslip attached. Let us know if you have any questions."
        yag.send(to=recipient_email, subject=subject, contents=body, attachments=file_path)
        print(f"✔ Payslip emailed to {recipient_email}")
    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}: {e}")

# === STEP 5: Main Program ===
def main():
    excel_file = r"C:\Users\uncommonstudent\New folder (2)\New folder (2)\Employee details.xlsx"
    employees = read_employee_data(excel_file)

    for index, row in employees.iterrows():
        employee_id = row['Employee ID']
        name = row['Name']
        email = row['Email']
        basic_salary = row['Basic Salary']
        allowances = row['Allowances']
        deductions = row['Deductions']

        net_salary = calculate_net_salary(basic_salary, allowances, deductions)
        payslip_path = generate_payslip(employee_id, name, basic_salary, allowances, deductions, net_salary)
        send_email(email, payslip_path)

# Run the program
if __name__ == "__main__":
    main()


