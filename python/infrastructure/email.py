from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from email.header import Header
import smtplib
import os
import logging

class Email:
    def __init__(self, sender: str, smtp_host: str = 'smtp.gmail.com', smtp_port: int = 587, smtp_user: str = None, smtp_password: str = None):
        """Email helper.

        sender: default From address
        smtp_host/smtp_port: SMTP server
        smtp_user/smtp_password: optional credentials used to login if provided
        """
        self.__sender = sender
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.smtp_user = smtp_user
        self.smtp_password = smtp_password

    def send_mail(self, p_recip_i, p_subject_i, p_msgbody_i):
        msg = MIMEMultipart()
        msg['From'] = self.__sender
        msg['To'] = p_recip_i
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = Header(p_subject_i, 'utf-8')
        html = f'<html> {p_msgbody_i}'
        msg.attach(MIMEText(html, 'html'))

        # SMTP send with optional authentication
        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.connect(self.smtp_host, self.smtp_port)
            server.ehlo()
            try:
                server.starttls()
                server.ehlo()
            except Exception:
                # server may not support starttls; continue
                pass
            # login only if credentials were supplied
            if self.smtp_user and self.smtp_password:
                server.login(self.smtp_user, self.smtp_password)
            server.sendmail(self.__sender, p_recip_i, msg.as_string())

    def format_message(self, title: str, subject: str, body_html: str, recipients: list):
        """Return a MIMEMultipart message formatted with title/subject/body and recipient list.

        Inputs:
        - title: visible title (kept in X-Title header)
        - subject: email Subject
        - body_html: HTML body
        - recipients: list of recipient email addresses or a comma-separated string

        Output: email.message.Message
        """
        msg = MIMEMultipart()
        # store title in a custom header so GUI can show it if needed
        msg['X-Email-Title'] = Header(title, 'utf-8')
        msg['From'] = self.__sender
        if isinstance(recipients, (list, tuple)):
            msg['To'] = ', '.join(recipients)
        else:
            msg['To'] = recipients
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = Header(subject, 'utf-8')
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))
        return msg

    def create_msg_file(self, message, filepath: str) -> str:
        """Create a preview email file. On Windows with pywin32 available, try to create a .msg file via Outlook.

        Falls back to writing the raw RFC822 .eml content. Returns the path written.
        """
        # Try to create .msg using win32com if available (Outlook must be installed).
        try:
            import win32com.client  # type: ignore
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # olMailItem
            # map fields
            mail.Subject = str(message.get('Subject', ''))
            mail.HTMLBody = message.get_payload()[0].get_payload(decode=True).decode('utf-8') if message.is_multipart() else message.get_payload(decode=True)
            mail.To = message.get('To', '')
            # Save as .msg
            if not filepath.lower().endswith('.msg'):
                filepath = filepath + '.msg'
            mail.SaveAs(filepath)
            return filepath
        except Exception:
            # fallback: write .eml (RFC822) file
            if not filepath.lower().endswith('.eml'):
                filepath = filepath + '.eml'
            with open(filepath, 'wb') as fh:
                fh.write(message.as_bytes())
            return filepath

    def send_multiple_mails(self, p_recip_df_i):
        for index, row in p_recip_df_i.iterrows():
            try:
                self.send_mail(row['email'], row['subject'], row['msgbody'])
            except Exception as e:
                logging.error(f"Failed to send email to {row['email']}: {e}")

        logging.info("All emails have been processed.")    