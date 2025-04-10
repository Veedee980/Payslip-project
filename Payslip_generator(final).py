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

# === STEP 1: Read employee data from Excel ===
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

    # Set background color (light peach) - RGB for light peach is (255, 229, 204)
    pdf.set_fill_color(255, 229, 204)  # Light Peach color
    pdf.rect(0, 0, 210, 297, 'F')  # Fill the entire page with the peach background

    # Set the company name and title in the header with purple color
    pdf.set_font("Arial", 'B', 24)
    pdf.set_text_color(128, 0, 128)  # Purple color for company name
    pdf.cell(200, 10, "ReWear", ln=True, align='C')
    pdf.ln(10)  # Add a line break

    # Set the title with a darker purple color
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(128, 0, 128)  # Purple color for title
    pdf.cell(200, 10, f"Payslip for {name}", ln=True, align='C')
    pdf.ln(10)  # Add a line break

    # Add a line separator
    pdf.set_line_width(0.5)
    pdf.set_draw_color(128, 0, 128)  # Purple color for the line
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)

    # Employee Details Section
    pdf.set_font("Arial", 'B', 12)
    pdf.set_text_color(128, 0, 128)  # Purple color for text
    pdf.cell(100, 10, f"Employee ID: {employee_id}", ln=True)
    pdf.cell(100, 10, f"Name: {name}", ln=True)
    pdf.ln(5)

    # Salary Details Section
    pdf.set_font("Arial", '', 12)
    pdf.cell(100, 10, f"Basic Salary:", align='R')
    pdf.cell(50, 10, f"${basic_salary:.2f}", align='R')
    pdf.ln(5)

    pdf.cell(100, 10, f"Allowances:", align='R')
    pdf.cell(50, 10, f"${allowances:.2f}", align='R')
    pdf.ln(5)

    pdf.cell(100, 10, f"Deductions:", align='R')
    pdf.cell(50, 10, f"${deductions:.2f}", align='R')
    pdf.ln(5)

    pdf.cell(100, 10, f"Net Salary:", align='R')
    pdf.cell(50, 10, f"${net_salary:.2f}", align='R')
    pdf.ln(10)

    # Add a line separator
    pdf.set_line_width(0.5)
    pdf.set_draw_color(128, 0, 128)  # Purple color for the line
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # Footer Section with Contact info
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(128, 0, 128)  # Purple color for footer text
    pdf.cell(200, 10, "Thank you for your hard work! For queries, contact HR.", ln=True, align='C')
    
    # Create directory for saving the payslips if it doesn't exist
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



