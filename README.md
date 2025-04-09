🧾 Automated Payslip Generator & Email Sender
This Python script automates the process of generating payslips as PDF files from employee salary data stored in an Excel sheet and then sends them to each employee via email.

📌 Features
Reads employee salary details from an Excel file

Calculates net salary (Basic + Allowances - Deductions)

Generates a personalized PDF payslip for each employee

Sends the payslip to their email address using Gmail

🛠️ Technologies Used
Python

pandas for reading Excel files

fpdf for creating PDF payslips

yagmail for sending emails

os for file and environment variable handling

⚙️ Setup Instructions
Clone the repository

bash
Copy
Edit
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
Install dependencies
Make sure you have Python installed. Then run:

bash
Copy
Edit
pip install pandas fpdf yagmail
Set up Gmail

You need a Gmail address and an App Password (if you use 2FA).

Set your credentials as environment variables:

bash
Copy
Edit
export GMAIL_USER='your-email@gmail.com'
export GMAIL_PASS='your-app-password'
Add your Excel file
Update the excel_file path in the script to point to your Excel sheet. Make sure it includes these columns:

nginx
Copy
Edit
Employee ID | Name | Email | Basic Salary | Allowances | Deductions
Run the script

bash
Copy
Edit
python your_script.py
📂 Output
PDFs will be saved to the payslips/ folder

Each employee will receive their individual payslip via email

✅ Example
PDF Content:

yaml
Copy
Edit
Payslip for John Doe
Employee ID: 12345
Basic Salary: $2000.00
Allowances: $300.00
Deductions: $100.00
Net Salary: $2200.00
🚨 Notes
Keep your email credentials secure. Use environment variables or .env files (never hardcode them).

Make sure "Allow less secure apps" is disabled if using an App Password with Gmail.

Double-check the email column in your Excel file before running the script.

