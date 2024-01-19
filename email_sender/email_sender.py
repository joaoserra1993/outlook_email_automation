import os
import smtplib
import sys
import time
from csv import reader
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from random import randint

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, \
    QVBoxLayout, QWidget, QHBoxLayout, \
    QLineEdit, QTextEdit, QFileDialog


class AnotherWindow(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setWindowTitle('Email Sender')
        self.label = QLabel("Emails sent % d" % randint(0, 100))
        layout.addWidget(self.label)
        self.setLayout(layout)

    def update_label(self, emails_sent):
        self.label.setText("Emails sent % d" % emails_sent)


class MainWindow(QMainWindow):
    next_window_signal = pyqtSignal()

    def __init__(self):
        super().__init__()

        self.setWindowTitle('Email Sender')
        self.resize(800, 600)

        layout = QVBoxLayout()

        email_layout = QHBoxLayout()
        email_label = QLabel('Email:')
        self.email = QLineEdit()
        email_layout.addWidget(email_label)
        email_layout.addWidget(self.email)
        layout.addLayout(email_layout)

        password_layout = QHBoxLayout()
        password_label = QLabel('Password:')
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        password_layout.addWidget(password_label)
        password_layout.addWidget(self.password)
        layout.addLayout(password_layout)

        subject_layout = QHBoxLayout()
        subject_label = QLabel('Subject:')
        self.subject = QLineEdit()
        subject_layout.addWidget(subject_label)
        subject_layout.addWidget(self.subject)
        layout.addLayout(subject_layout)

        global files_attached
        files_attached = QLabel('')
        layout.addWidget(files_attached,
                         alignment=Qt.AlignmentFlag.AlignCenter)

        paragraph_one_layout = QVBoxLayout()
        paragraph_one_label = QLabel('Paragraph One:')
        self.paragraph_one = QTextEdit()
        paragraph_one_layout.addWidget(paragraph_one_label)
        paragraph_one_layout.addWidget(self.paragraph_one)
        layout.addLayout(paragraph_one_layout)

        paragraph_two_layout = QVBoxLayout()
        paragraph_two_label = QLabel('Paragraph Two:')
        self.paragraph_two = QTextEdit()
        paragraph_two_layout.addWidget(paragraph_two_label)
        paragraph_two_layout.addWidget(self.paragraph_two)
        layout.addLayout(paragraph_two_layout)

        open_button = QPushButton('Select Files')
        layout.addWidget(open_button)
        open_button.clicked.connect(self.open_files)

        button = QPushButton('Send email')
        layout.addWidget(button)
        button.clicked.connect(self.send_email)
        button.clicked.connect(self.next_window_signal.emit)

        self.setLayout(layout)
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.another_window = AnotherWindow()

    def open_files(self):
        global files
        files, _ = QFileDialog.getOpenFileNames(self, 'Select files')
        file_names = [os.path.basename(file_path) for file_path in files]
        files_attached.setText(', '.join(file_names))

    def send_email(self):
        sender_email = self.email.text()
        sender_password = self.password.text()
        subject = self.subject.text()
        paragraph_one = self.paragraph_one.toPlainText()
        paragraph_two = self.paragraph_two.toPlainText()

        with open(
                "/email_sender/email_list.csv",
                "r") as my_file:
            file_reader = reader(my_file)
            emails_sent = 0
            for recipient_email in file_reader:
                msg = MIMEMultipart('alternative')
                msg['From'] = sender_email
                msg['Subject'] = subject
                msg['To'] = recipient_email[0]
                recipient = recipient_email[1]
                html = (f'<html>'
                        f'<p>Dear {recipient},<p>'
                        f'<p>{paragraph_one}</p>'
                        f'<p>{paragraph_two}</p>'
                        f'<p>Your text pre defined</p>'
                        f'<p>O <b>Company</b>More text information</p>'
                        f'<p>more text <b>webiste</b></p><br><br>'
                        f'<p>Sincerely<br></p>'
                        f'--'
                        f'<p style="font-size: 20px; font-weight: 900; "><b>Peter</b></p>'
                        f'Peter Company S.A.'
                        f'<p style="color: rgb(211, 211, 211);">Street streetname<br>'
                        f'4000-000, Porto, Portugal</p>'
                        f'email.pt'
                        f'</html>')

                # Attach the HTML to the email
                html_text = MIMEText(html, 'html')
                msg.attach(html_text)

                # Attach the files to the email
                for file_path in files:
                    with open(file_path, "rb") as file:
                        part = MIMEApplication(file.read(),
                                               Name=os.path.basename(
                                                   file_path))
                        part[
                            'Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                        msg.attach(part)

                # SMTP server settings for Outlook
                smtp_server = 'smtp-mail.outlook.com'
                smtp_port = 587

                try:
                    # Create a secure SSL/TLS connection to the SMTP server
                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()

                    # Login to your Outlook email account
                    server.login(sender_email, sender_password)

                    time.sleep(5)

                    # Send the email
                    server.sendmail(sender_email, recipient_email,
                                    msg.as_string())

                except smtplib.SMTPException as e:
                    print("Error sending email:", str(e), msg['To'])

                finally:
                    # Close the connection to the SMTP server
                    server.quit()

                emails_sent += 1
                if emails_sent == 20:
                    time.sleep(15)

        # Update the AnotherWindow with the number of emails sent
        self.another_window.update_label(emails_sent)
        self.another_window.show()

        # Close the main window
        self.close()


# Create instances of the windows
app = QApplication(sys.argv)
main_window = MainWindow()

# Show the main window
main_window.show()

# Run the application
sys.exit(app.exec())
