# config.py

# Branding
# -------------------------
COMPANY_NAME = "Arun Arudra"
COMPANY_ADDRESS = "#19299910, 622nd Floor, 2000th Main, 1998th Phase,\n Arudra Nagar, Bengaluru – 560 000."
PRIMARY_COLOR_HEX = "#000000"

# Email Configuration
# -------------------------
# Fill these details to enable automatic email sending.
# For services like Gmail, you may need to use an "App Password".
EMAIL_CONFIG = {
    "SMTP_HOST": "mail.gmail.in",
    "SMTP_PORT": 587,
    "SMTP_USERNAME": "arun@arudra.in",
    "SMTP_PASSWORD": "1234567890",
    "FROM_NAME": "Arun Arudra",
    "FROM_EMAIL": "arun@arudra.in",
    "USE_TLS": True
}

# Script Behavior
# -------------------------
# Set to True to generate and send payslips for all months from joining date to the current month.
# Set to False to only generate and send for the current month.
SEND_ALL_PAST_PAYSLIPS = False
