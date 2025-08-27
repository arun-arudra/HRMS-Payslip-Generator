# HRMS Payslip Generator

A Python-based **HRMS (Human Resource Management System) tool** that generates employee payslips in **PDF format** from an Excel sheet and can optionally **send them via email**.

This tool is useful for small businesses or teams who want to automate payslip generation without relying on third-party SaaS payroll systems.

---

## ✨ Features
- Generate professional PDF payslips using ReportLab.
- Store employee details in `employees.xlsx`.
- Automatically calculate **earnings, deductions, and net payable salary**.
- Embed company branding and logos by simply replacing the `logo.svg` file.
- Email payslips directly to employees via SMTP.
- Maintain a log of sent payslips (`.payslip_sent_log.json`).
- Configurable to send **current month only** or **all past payslips** since joining.

---

## 📂 Project Structure
hrms-payslip\
├── employees.xlsx # Employee database (auto-created if missing)\
├── logo.svg       # Your company logo (replace with your own)\
├── payslips/      # Generated PDF payslips\
├── hrms.py        # Main script\
├── config.py      # Configuration file for company details and email settings\
├── install.py     # Installation script for dependencies\
├── .payslip_sent_log.json # Log of sent payslips\
├── README.md      # Project documentation\
├── LICENSE        # GNU GPL license\
└── requirements.txt # Python dependencies

---

## 🚀 Getting Started

### 1. Simple Installation

Run the installation script to check for and install all required Python libraries.

```bash
python install.py
```
This command will automatically install dependencies from ```requirements.txt``` and then run the main script.

### 2. Prepare Employee Data
The script will automatically create a template file named ```employees.xlsx``` in the project directory if it's missing.

Open this file and fill in your employee details, making sure to use the exact column headers provided in the template.

### 3. Customize Your Branding & Email
Open ```config.py``` to change the company name, address, and email settings.

Logo: Replace the ```logo.svg``` file in the main directory with your own company logo. The script will automatically use this file for the PDF.


## 📜 License
This project is licensed under the GNU General Public License (GPL v3). You are free to use, modify, and distribute it under the same license terms.

## 🤝 Contribution
Contributions are welcome! If you'd like to improve this project, please follow these steps:

Fork the repository.

Create a feature branch (```git checkout -b feature-name```).

Commit your changes (```git commit -m 'Add new feature'```).

Push to the branch (```git push origin feature-name```).

Open a Pull Request.
