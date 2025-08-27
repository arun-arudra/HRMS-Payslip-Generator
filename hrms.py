import os
import io
import json
import logging
from datetime import datetime
from pathlib import Path
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
import calendar

# Import configurations from the new config.py file
from config import (
    COMPANY_NAME,
    COMPANY_ADDRESS,
    PRIMARY_COLOR_HEX,
    EMAIL_CONFIG,
    SEND_ALL_PAST_PAYSLIPS
)

# -------------------------
# ========== PATHS & LOGGING ==========
# -------------------------
BASE_DIR = Path(__file__).resolve().parent
EMP_XLSX = BASE_DIR / "employees.xlsx"
PAYSLIPS_DIR = BASE_DIR / "payslips"
SENT_LOG_JSON = BASE_DIR / ".payslip_sent_log.json"
LOGO_SVG_FILE = BASE_DIR / "logo.svg"

# Default SVG logo if no file exists
DEFAULT_LOGO_SVG_CODE = """
<svg width="36" height="40" viewBox="0 0 36 40" fill="none" xmlns="http://www.w3.org/2000/svg">
<path fill-rule="evenodd" clip-rule="evenodd" d="M0 15V31H5C5.52527 31 6.04541 31.1035 6.53076 31.3045C7.01599 31.5055 7.45703 31.8001 7.82837 32.1716C8.19983 32.543 8.49451 32.984 8.69556 33.4693C8.89648 33.9546 9 34.4747 9 35V40H21L36 25V9H31C30.4747 9 29.9546 8.89655 29.4692 8.69553C28.984 8.49451 28.543 8.19986 28.1716 7.82843C27.8002 7.457 27.5055 7.01602 27.3044 6.53073C27.1035 6.04544 27 5.5253 27 5V0H15L0 15ZM17 30H10V19L19 10H26V21L17 30Z" fill="#0004E8"></path>
</svg>
"""

# Define color from config
PRIMARY_COLOR = colors.HexColor(PRIMARY_COLOR_HEX)

# Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger("hrms_payslip")

# -------------------------
# ========== HELPERS ========
# -------------------------
def get_svg_logo():
    """Reads the SVG logo from file or creates a default one."""
    if not LOGO_SVG_FILE.exists():
        logger.info(f"{LOGO_SVG_FILE} not found. Creating a default logo file.")
        with open(LOGO_SVG_FILE, "w", encoding="utf-8") as f:
            f.write(DEFAULT_LOGO_SVG_CODE.strip())
    
    with open(LOGO_SVG_FILE, "r", encoding="utf-8") as f:
        return f.read()

def create_dummy_excel(file_path):
    """Create a dummy employee excel file if missing (with required columns)."""
    if file_path.exists():
        logger.info(f"Employee file found: {file_path}")
        return
    logger.info("employees.xlsx not found — creating template...")
    # These headers are derived from the provided payslip PDF
    dummy_data = {
        "Employee ID": ["AA001"],
        "FullName": ["Arun Kumar"],
        "Date of Joining": ["27-08-2025"],
        "Department": ["Design"],
        "Sub Department": ["N/A"],
        "Designation": ["Graphic Designer"],
        "Payment Mode": ["Bank Transfer"],
        "Bank": ["ICICI Bank"],
        "Bank IFSC": ["ICIC0000001"],
        "Bank Account": ["9xx0100XXXXXXX"],
        "PAN": ["XXXXKXXXXX"],
        "UAN": ["N/A"],
        "PF Number": ["MOH/001/0001"],
        "Email": ["arun@arudra.com"],
        "Annual CTC": [578400.00],
        "Basic": [23500.00],
        "HRA": [11750.00],
        "Medical Allowance": [4700.00],
        "Transport Allowance": [1600.00],
        "Special Allowance": [3100.00],
        "Professional Allowance": [1175.00],
        "Performance Pay": [1175.00],
        "Courier Reimb": [1200.00],
        "Total Working Days": [20],
        "Actual Payable Days": [19],
        "Professional Tax": [200.00],
        "Performance Bonus": [1000.00],
        "Performance Bonus Recovery": [0.0],
        "PF": [500.00],
    }
    df = pd.DataFrame(dummy_data)
    df.to_excel(file_path, index=False)
    logger.info(f"Template created at {file_path}. Please open and fill employee rows.")


def load_sent_log():
    if SENT_LOG_JSON.exists():
        try:
            with open(SENT_LOG_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_sent_log(d):
    with open(SENT_LOG_JSON, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2)

def month_year_string(dt=None):
    dt = dt or datetime.now()
    return dt.strftime("%B"), dt.year

def num_to_words_indian(n):
    # integer to Indian words (simple version)
    ones = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve",
            "Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"]
    tens = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"]
    def two(n):
        if n < 20:
            return ones[n]
        else:
            return tens[n//10] + ("" if n%10==0 else " " + ones[n%10])
    def three(n):
        s = ""
        if n >= 100:
            s += ones[n//100] + " Hundred"
            if n%100:
                s += " "
        if n%100:
            s += two(n%100)
        return s
    if n == 0:
        return "Zero"
    parts = []
    crore = n // 10000000
    if crore:
        parts.append(three(crore) + " Crore")
    n = n % 10000000
    lakh = n // 100000
    if lakh:
        parts.append(three(lakh) + " Lakh")
    n = n % 100000
    thousand = n // 1000
    if thousand:
        parts.append(three(thousand) + " Thousand")
    n = n % 1000
    if n:
        parts.append(three(n))
    return " ".join(parts)

# -------------------------
# ========== PDF BUILD =====
# -------------------------
def create_payslip_pdf(row, month_name, year, out_pdf_path, logo_svg_code):
    """
    Create the PDF matching the provided design.
    """
    WIDTH, HEIGHT = A4
    c = canvas.Canvas(str(out_pdf_path), pagesize=A4)
    left_margin = 14 * mm
    right_margin = 14 * mm
    usable_width = WIDTH - left_margin - right_margin
    top_margin = HEIGHT - 16 * mm
    y_pos = top_margin

    # Define colors
    PAYSLIP_REGULAR_COLOR = colors.HexColor("#505050")
    TEXT_COLOR = colors.HexColor("#000000")
    LABEL_COLOR = colors.HexColor("#858585")
    LINE_COLOR = colors.HexColor("#DCDCDC")

    # ----- Header Section -----
    # Left side: PAYSLIP May 2023
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left_margin, y_pos, "PAYSLIP")
    text_width_payslip = c.stringWidth("PAYSLIP", "Helvetica-Bold", 18)
    c.setFillColor(PAYSLIP_REGULAR_COLOR)
    c.setFont("Helvetica", 18)
    c.drawString(left_margin + text_width_payslip, y_pos, f" {month_name.upper()} {year}")
    y_pos -= 5 * mm

    # Company name and address
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left_margin, y_pos, COMPANY_NAME.upper())
    y_pos -= 5 * mm
    c.setFont("Helvetica", 8)
    address_lines = COMPANY_ADDRESS.split("\n")
    for line in address_lines:
        c.drawString(left_margin, y_pos, line)
        y_pos -= 5 * mm

    # Right side: Logo
    if logo_svg_code:
        try:
            drawing = svg2rlg(io.StringIO(logo_svg_code))
            scale_w = (40*mm) / drawing.width if drawing.width > 0 else 1.0
            scale_h = (20*mm) / drawing.height if drawing.height > 0 else 1.0
            scale = min(scale_w, scale_h, 1.0)
            drawing.width *= scale
            drawing.height *= scale

            logo_x = WIDTH - right_margin - drawing.width
            logo_y = top_margin - drawing.height - 10*mm

            renderPDF.draw(drawing, c, logo_x, logo_y)
        except Exception as e:
            logger.warning(f"SVG render failed from embedded code: {e}")
    
    y_pos -= 10 * mm

    # ----- Employee Name & Line -----
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left_margin, y_pos, str(row.get("FullName", "")).upper())
    y_pos -= 5 * mm
    c.setStrokeColor(TEXT_COLOR)
    c.setLineWidth(1.5)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 5 * mm

    # ----- Employee Details Grid 1 -----
    c.setLineWidth(0.5)
    col_width = usable_width / 4
    
    # Titles
    c.setFillColor(LABEL_COLOR)
    c.setFont("Helvetica", 7)
    c.drawString(left_margin, y_pos, "Employee Number")
    c.drawString(left_margin + col_width, y_pos, "Date Joined")
    c.drawString(left_margin + col_width * 2, y_pos, "Department")
    c.drawString(left_margin + col_width * 3, y_pos, "Sub Department")
    y_pos -= 4 * mm

    # Data
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left_margin, y_pos, str(row.get("Employee ID", 'N/A')))
    
    date_joined_str = str(row.get("Date of Joining", 'N/A'))
    try:
        date_obj = pd.to_datetime(date_joined_str)
        formatted_date = date_obj.strftime("%d %b %Y").upper()
    except (ValueError, TypeError):
        formatted_date = date_joined_str
    
    c.drawString(left_margin + col_width, y_pos, formatted_date)
    c.drawString(left_margin + col_width * 2, y_pos, str(row.get("Department", 'N/A')))
    c.drawString(left_margin + col_width * 3, y_pos, str(row.get("Sub Department", 'N/A')))
    y_pos -= 5 * mm
    c.setStrokeColor(LINE_COLOR)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 5 * mm

    # ----- Employee Details Grid 2 -----
    # Titles
    c.setFillColor(LABEL_COLOR)
    c.setFont("Helvetica", 7)
    c.drawString(left_margin, y_pos, "Designation")
    c.drawString(left_margin + col_width, y_pos, "Payment Mode")
    c.drawString(left_margin + col_width * 2, y_pos, "Bank")
    c.drawString(left_margin + col_width * 3, y_pos, "Bank IFSC")
    y_pos -= 4 * mm
    
    # Data
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left_margin, y_pos, str(row.get("Designation", 'N/A')))
    c.drawString(left_margin + col_width, y_pos, str(row.get("Payment Mode", 'N/A')))
    c.drawString(left_margin + col_width * 2, y_pos, str(row.get("Bank", 'N/A')))
    c.drawString(left_margin + col_width * 3, y_pos, str(row.get("Bank IFSC", 'N/A')))
    y_pos -= 5 * mm
    c.setStrokeColor(LINE_COLOR)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 5 * mm
    
    # ----- Employee Details Grid 3 -----
    # Titles
    c.setFillColor(LABEL_COLOR)
    c.setFont("Helvetica", 7)
    c.drawString(left_margin, y_pos, "Bank Account")
    c.drawString(left_margin + col_width, y_pos, "PAN")
    c.drawString(left_margin + col_width * 2, y_pos, "UAN")
    c.drawString(left_margin + col_width * 3, y_pos, "PF Number")
    y_pos -= 4 * mm
    
    # Data
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left_margin, y_pos, str(row.get("Bank Account", 'N/A')))
    c.drawString(left_margin + col_width, y_pos, str(row.get("PAN", 'N/A')))
    c.drawString(left_margin + col_width * 2, y_pos, str(row.get("UAN", 'N/A')))
    c.drawString(left_margin + col_width * 3, y_pos, str(row.get("PF Number", 'N/A')))
    y_pos -= 5 * mm
    c.setStrokeColor(LINE_COLOR)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 10 * mm

    # ----- Salary Details Section -----
    c.setStrokeColor(TEXT_COLOR)
    c.setLineWidth(1.5)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 5 * mm
    
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(TEXT_COLOR)
    c.drawString(left_margin, y_pos, "SALARY DETAILS")
    y_pos -= 5 * mm
    
    c.setLineWidth(0.5)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 5 * mm
    
    try:
        total_working_days = float(row.get("Total Working Days", 0))
        actual_payable_days = float(row.get("Actual Payable Days", 0))
    except (ValueError, TypeError):
        total_working_days = 0.0
        actual_payable_days = 0.0
    
    loss_of_pay_days = total_working_days - actual_payable_days
    days_payable = actual_payable_days

    # Salary Details table
    c.setLineWidth(0.5)
    c.setFillColor(LABEL_COLOR)
    c.setFont("Helvetica", 7)
    table_headers = ["Actual Payable Days", "Total Working Days", "Loss of Pay Days", "Days Payable"]
    table_values = [
        f"{actual_payable_days}",
        f"{total_working_days}",
        f"{loss_of_pay_days}",
        f"{days_payable}"
    ]
    
    col_width_details = usable_width / len(table_headers)
    
    for i, header in enumerate(table_headers):
        c.drawString(left_margin + i * col_width_details, y_pos, header)
    y_pos -= 4 * mm
    
    c.setFillColor(TEXT_COLOR)
    c.setFont("Helvetica-Bold", 9)
    for i, value in enumerate(table_values):
        c.drawString(left_margin + i * col_width_details, y_pos, value)
    y_pos -= 5 * mm
    c.setStrokeColor(LINE_COLOR)
    c.line(left_margin, y_pos, left_margin + usable_width, y_pos)
    y_pos -= 10 * mm
    
    # ----- Earnings and Deductions Columns -----
    # Vertical divider line
    c.setStrokeColor(LINE_COLOR)
    c.setLineWidth(1)
    c.line(left_margin + usable_width * 0.5, y_pos + 5 * mm, left_margin + usable_width * 0.5, y_pos - 60 * mm)

    left_col_x = left_margin
    right_col_x = left_margin + usable_width * 0.5
    col_width_sal = usable_width * 0.5
    horizontal_padding = 5 * mm

    # Earnings block
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(TEXT_COLOR)
    c.drawString(left_col_x, y_pos, "EARNINGS")
    y_earn = y_pos - 7*mm
    
    total_earn = 0.0
    c.setFont("Helvetica", 8.5)
    
    prorate_items = ["Basic", "HRA", "Special Allowance"]
    for label in prorate_items:
        amt = row.get(label, 0)
        try:
            amt_f = float(amt) if pd.notna(amt) else 0.0
        except Exception:
            amt_f = 0.0

        prorated_amt = (amt_f / total_working_days) * actual_payable_days if total_working_days > 0 else 0
        
        c.drawString(left_col_x, y_earn, label)
        c.drawRightString(left_col_x + col_width_sal - 4*mm, y_earn, f"{prorated_amt:,.2f}")
        y_earn -= 5*mm
        total_earn += prorated_amt
    
    non_prorate_items = ["Medical Allowance", "Transport Allowance", "Professional Allowance", "Performance Pay", "Courier Reimb"]
    for label in non_prorate_items:
        amt = row.get(label, 0)
        try:
            amt_f = float(amt) if pd.notna(amt) else 0.0
        except Exception:
            amt_f = 0.0
        
        c.drawString(left_col_x, y_earn, label)
        c.drawRightString(left_col_x + col_width_sal - 4*mm, y_earn, f"{amt_f:,.2f}")
        y_earn -= 5*mm
        total_earn += amt_f

    pb_earn = float(row.get("Performance Bonus", 0)) if pd.notna(row.get("Performance Bonus", 0)) else 0.0
    if pb_earn > 0:
        c.drawString(left_col_x, y_earn, "Performance Bonus")
        c.drawRightString(left_col_x + col_width_sal - 4*mm, y_earn, f"{pb_earn:,.2f}")
        y_earn -= 5*mm
        total_earn += pb_earn
    
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left_col_x, y_earn - 3*mm, "Total Earnings (A)")
    c.drawRightString(left_col_x + col_width_sal - 4*mm, y_earn - 3*mm, f"{total_earn:,.2f}")

    # Deductions block
    y_pos_ded = y_pos
    c.setFont("Helvetica-Bold", 10)
    c.drawString(right_col_x + horizontal_padding, y_pos_ded, "TAXES & DEDUCTIONS")
    y_ded = y_pos_ded - 7*mm
    total_ded = 0.0
    
    c.setFont("Helvetica", 8.5)

    pt_amt = float(row.get("Professional Tax", 0)) if pd.notna(row.get("Professional Tax", 0)) else 0.0
    c.drawString(right_col_x + horizontal_padding, y_ded, "Professional Tax")
    c.drawRightString(right_col_x + col_width_sal - 4*mm, y_ded, f"{pt_amt:,.2f}")
    y_ded -= 5 * mm
    total_ded += pt_amt

    pf_amt = float(row.get("PF", 0)) if pd.notna(row.get("PF", 0)) and float(row.get("PF", 0)) > 0 else 0.0
    if pf_amt > 0:
        c.drawString(right_col_x + horizontal_padding, y_ded, "PF (Provident Fund)")
        c.drawRightString(right_col_x + col_width_sal - 4*mm, y_ded, f"{pf_amt:,.2f}")
        y_ded -= 5 * mm
        total_ded += pf_amt
    
    pb_recovery = float(row.get("Performance Bonus Recovery", 0)) if pd.notna(row.get("Performance Bonus Recovery", 0)) else 0.0
    if pb_recovery > 0:
        c.drawString(right_col_x + horizontal_padding, y_ded, "Performance Bonus")
        c.drawRightString(right_col_x + col_width_sal - 4*mm, y_ded, f"{pb_recovery:,.2f}")
        y_ded -= 5*mm
        total_ded += pb_recovery
    
    c.setFont("Helvetica-Bold", 8)
    c.drawString(right_col_x + horizontal_padding, y_ded - 3*mm, "Total Deductions (C)")
    c.drawRightString(right_col_x + col_width_sal - 4*mm, y_ded - 3*mm, f"{total_ded:,.2f}")

    # Bottom line
    y_summary = min(y_earn, y_ded) - 20*mm
    c.setLineWidth(1.5)
    c.setStrokeColor(TEXT_COLOR)
    c.line(left_margin, y_summary, left_margin + usable_width, y_summary)
    
    # ----- Summary Section -----
    y_summary -= 5 * mm
    net_salary = total_earn - total_ded
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin, y_summary, "Net Salary Payable (A-C)")
    c.drawRightString(left_margin + usable_width, y_summary, f"{net_salary:,.2f}")
    y_summary -= 7 * mm
    
    c.setFont("Helvetica", 8)
    net_int = int(round(net_salary))
    words = num_to_words_indian(net_int) + " only"
    c.drawString(left_margin, y_summary, "Net Salary in words")
    c.drawRightString(left_margin + usable_width, y_summary, words)
    
    y_summary -= 15*mm

    # ----- Footer -----
    c.setFont("Helvetica", 7)
    c.setFillColor(PAYSLIP_REGULAR_COLOR)
    c.drawString(left_margin, y_summary, "Note: All amounts displayed in this payslip are in INR")
    c.drawString(left_margin, y_summary - 5*mm, "This is computer generated statement, does not require signature.")

    c.showPage()
    c.save()


# -------------------------
# ========== EMAIL =========
# -------------------------
def send_email_with_attachment(to_email, subject, body, attachment_path):
    cfg = EMAIL_CONFIG
    if not cfg.get("SMTP_USERNAME") or not cfg.get("SMTP_PASSWORD"):
        logger.warning("Email credentials not provided - skipping email send.")
        return False, "SMTP not configured"

    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = formataddr((cfg.get("FROM_NAME"), cfg.get("FROM_EMAIL")))
        msg['To'] = to_email
        msg.set_content(body)

        with open(attachment_path, "rb") as f:
            pdf_data = f.read()
        msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=os.path.basename(attachment_path))

        if cfg.get("USE_TLS", True):
            server = smtplib.SMTP(cfg["SMTP_HOST"], cfg["SMTP_PORT"], timeout=30)
            server.ehlo()
            server.starttls()
            server.login(cfg["SMTP_USERNAME"], cfg["SMTP_PASSWORD"])
        else:
            server = smtplib.SMTP_SSL(cfg["SMTP_HOST"], cfg["SMTP_PORT"], timeout=30)
            server.login(cfg["SMTP_USERNAME"], cfg["SMTP_PASSWORD"])

        server.send_message(msg)
        server.quit()
        logger.info(f"Email sent to {to_email}")
        return True, "Sent"
    except Exception as e:
        logger.exception(f"Failed to send email to {to_email}: {e}")
        return False, str(e)

# -------------------------
# ========== MAIN ==========
# -------------------------
def main():
    logo_svg_code = get_svg_logo()
    create_dummy_excel(EMP_XLSX)
    month_name, year = month_year_string()
    logger.info(f"Running payslip generation for {month_name} {year}")

    sent_log = load_sent_log()
    sent_key = f"{year}-{month_name}"

    if sent_log.get("last_sent") == sent_key:
        logger.info(f"Payslips already processed for {month_name} {year} (log). Exiting.")
        return

    try:
        df = pd.read_excel(EMP_XLSX)
    except Exception as e:
        logger.exception(f"Failed to read {EMP_XLSX}: {e}")
        return

    PAYSLIPS_DIR.mkdir(parents=True, exist_ok=True)
    created_files = []
    
    email_cfg_ready = EMAIL_CONFIG.get("SMTP_USERNAME") and EMAIL_CONFIG.get("SMTP_PASSWORD")

    for idx, row in df.iterrows():
        if pd.isna(row.get("FullName")) or pd.isna(row.get("Employee ID")):
            continue
        
        fullname = str(row.get("FullName")).strip()
        employee_email = row.get("Email")
        
        date_of_joining_str = str(row.get("Date of Joining", 'N/A'))
        try:
            date_of_joining = pd.to_datetime(date_of_joining_str)
        except (ValueError, TypeError):
            date_of_joining = datetime.now()

        start = date_of_joining.replace(day=1)
        end = datetime.now().replace(day=1)

        months_to_process = []
        if SEND_ALL_PAST_PAYSLIPS:
            while start <= end:
                months_to_process.append(start)
                next_month = start.replace(day=28) + pd.Timedelta(days=4)
                start = next_month.replace(day=1)
        else:
            months_to_process = [end]

        for date_to_process in months_to_process:
            month_name_gen = calendar.month_name[date_to_process.month]
            year_gen = date_to_process.year
            
            emp_folder = PAYSLIPS_DIR / fullname / str(year_gen) / month_name_gen
            emp_folder.mkdir(parents=True, exist_ok=True)

            safe_name = f"{fullname}-payslip-{month_name_gen}-{year_gen}.pdf"
            out_pdf = emp_folder / safe_name

            create_payslip_pdf(row, month_name_gen, year_gen, out_pdf, logo_svg_code)
            logger.info(f"Created payslip: {out_pdf}")
            created_files.append((row, out_pdf))

            if pd.isna(employee_email) or not employee_email:
                logger.warning(f"No email for {fullname}, skipping email for this payslip.")
                continue

            subject = f"Payslip For {month_name_gen} {year_gen} - {COMPANY_NAME}"
            body = f"Dear {fullname},\n\nPlease find enclosed Payslip for the month of {month_name_gen} {year_gen}. We suggest that you save it in your personal records for any future reference.\n\nImportant:\n- Please ensure that you check the entries in your payslip and for any queries or concerns, you may approach your HR Manager or Payroll Admin.\n\nRegards,\n{EMAIL_CONFIG.get('FROM_NAME')}"

            if email_cfg_ready:
                ok, msg = send_email_with_attachment(employee_email, subject, body, str(out_pdf))
                if not ok:
                    logger.error(f"Email failed for {employee_email}: {msg}")
            else:
                logger.info(f"Email not configured. Skipping email for {employee_email} (payslip created).")

    sent_log["last_sent"] = sent_key
    sent_log.setdefault("history", []).append({
        "month": month_name,
        "year": year,
        "timestamp": datetime.now().isoformat(),
        "created": [str(p) for (_, p) in created_files]
    })
    save_sent_log(sent_log)
    logger.info("Payslip generation process completed. Thank you Arun")

if __name__ == "__main__":
    main()
