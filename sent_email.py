from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import os

def send_email_with_attachments(pdf_folder, sender_email, sender_password, recipient_email, subject, body=""):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    try:
        # Tạo email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Đính kèm tất cả file PDF trong thư mục pdf_folder
        for filename in os.listdir(pdf_folder):
            if filename.endswith(".pdf"):
                filepath = os.path.join(pdf_folder, filename)
                with open(filepath, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={filename}',
                )
                msg.attach(part)

        # Kết nối tới Gmail SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Bật mã hóa
        server.login(sender_email, sender_password)  # Đăng nhập
        server.sendmail(sender_email, recipient_email, msg.as_string())  # Gửi email
        server.quit()

        print("Email đã được gửi thành công.")
    except Exception as e:
        print(f"Lỗi khi gửi email: {e}")