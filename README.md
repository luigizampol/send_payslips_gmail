# Employee Payslip Distribution Script
This Python script automates the process of sending employee payslips via Gmail. It reads a spreadsheet containing employee names and email addresses, then attaches the latest payslip PDF from each employee's folder and sends it as an email attachment. This script can be useful for HR departments or organizations to efficiently distribute payslips to employees.

# Usage
Before using this script, make sure to follow these steps:

Spreadsheet Preparation: Create an Excel spreadsheet containing employee information with columns for names and email addresses. Update the script with the correct column names for names and email addresses (name_column and email_column).

Folder Structure: Organize employee payslips in individual folders named after each employee. These folders should be located within an "Assets" folder. Each employee folder should have a "Payslips" subfolder where the payslip PDF files are stored.

Gmail Configuration: Configure the script to use a Gmail account to send emails. Provide the Gmail login, sender email address, and App Password for authentication (email_login, email_sender, and email_password variables).

Email Content: Customize the email subject, body (in HTML format), and payslip description according to your needs (subject, body, and payslip_description variables).

Script Execution: Run the script using Python. It will iterate through the spreadsheet, find the latest payslip PDF for each employee, and send the payslip as an email attachment.

# Configuration
Before running the script, ensure you have set the following variables in the script:

file_name: The name of the Excel spreadsheet containing employee names and email addresses.

name_column: The name of the column in the spreadsheet where employee names are located.

email_column: The name of the column in the spreadsheet where employee email addresses are located.

email_login, email_sender, email_password: Gmail account credentials for sending emails.

payslip_description: Description of the payslip content (e.g., "August Salary").

subject: Email subject line. By default, it is the same as payslip_description.

body: The email content in HTML format.

assets_folder: The name of the folder containing individual employee folders.

payslips_folder: The name of the subfolder within each employee folder where payslip PDFs are stored.

# Dependencies
This script requires the following Python libraries:

smtplib: For sending emails.
ssl: For creating an SSL context.
email.message, email.mime.text, email.mime.multipart, email.mime.application: For creating and formatting email messages.
openpyxl: For reading Excel spreadsheets.

# Disclaimer
Make sure to handle sensitive employee information and email credentials with care. Test the script with a small group of test emails before mass distribution.


# License
This script is provided under the MIT License.