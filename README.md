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
hrms-payslip/
├── employees.xlsx # Employee database (auto-created if missing)
├── logo.svg       # Your company logo (replace with your own)
├── payslips/      # Generated PDF payslips
├── hrms.py        # Main script
├── config.py      # Configuration file for company details and email settings
├── install.py     # Installation script for dependencies
├── .payslip_sent_log.json # Log of sent payslips
├── README.md      # Project documentation
├── LICENSE        # GNU GPL license
└── requirements.txt # Python dependencies

---

## 🚀 Getting Started

### 1. Simple Installation

Run the installation script to check for and install all required Python libraries.

```bash
python install.py
