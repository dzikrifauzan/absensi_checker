"""
Pengiriman email ringkasan laporan via SMTP.
"""

__version__ = "1.0.0"

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

from config import CONFIG
from logging_setup import log


class EmailSender:
    def __init__(self):
        self.smtp_host = CONFIG["SMTP_HOST"]
        self.smtp_port = CONFIG["SMTP_PORT"]
        self.smtp_user = CONFIG["SMTP_USER"]
        self.smtp_password = CONFIG["SMTP_PASSWORD"]
        self.from_email = CONFIG["FROM_EMAIL"]
        self.to_email = CONFIG["TO_EMAIL"]

    def send_report(self, subject: str, html_body: str, attachment_path: Path = None):
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = self.from_email
        msg["To"] = self.to_email

        msg.attach(MIMEText(html_body, "html"))

        if attachment_path and attachment_path.exists():
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={attachment_path.name}"
            )
            msg.attach(part)

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.smtp_user, self.smtp_password)
            server.sendmail(self.from_email, self.to_email, msg.as_string())

        log.info(f"Email terkirim ke {self.to_email}")