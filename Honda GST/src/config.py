import os
from dotenv import load_dotenv

load_dotenv()

EMAIL_CONFIG = {
    "host": os.getenv("EMAIL_HOST", "smtp.office365.com"),
    "port": int(os.getenv("EMAIL_PORT", 587)),
    "user": os.getenv("EMAIL_USER"),
    "pass": os.getenv("EMAIL_PASS"),
}