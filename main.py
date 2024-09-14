import sys
import imaplib
import email
from email.header import decode_header
import datetime
import json
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QComboBox, QDateEdit,
                             QTextEdit, QMessageBox, QInputDialog, QListWidget)
from PyQt5.QtGui import QFont, QColor, QPalette
from PyQt5.QtCore import Qt, QDate, QThread, pyqtSignal

CONFIG_FILE = "configurations.json"

def decode_email_content(part):
    try:
        charset = part.get_content_charset()
        if not charset:
            charset = 'utf-8'  # default to utf-8 if charset is not specified
        return part.get_payload(decode=True).decode(charset, errors='ignore')
    except UnicodeDecodeError:
        return part.get_payload(decode=True).decode('latin-1', errors='ignore')

class ArchiverThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    senders_signal = pyqtSignal(list)  # Signal to send list of senders

    def __init__(self, imap_server, imap_port, username, password, keywords, archive_date, action):
        super().__init__()
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.username = username
        self.password = password
        self.keywords = keywords
        self.archive_date = archive_date
        self.action = action
        self.cancel_event = False
        self.selected_senders = []  # Initialize selected_senders

    def run(self):
        try:
            # Connect to the server
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            # Login to your account
            mail.login(self.username, self.password)

            if self.action == "archive":
                self.archive_emails(mail)
            elif self.action == "delete_drafts":
                self.delete_draft_emails(mail)
            elif self.action == "collect_senders":
                self.collect_senders(mail)

            # Logout
            mail.logout()
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")
        self.finished_signal.emit()

    def archive_emails(self, mail):
        try:
            # Select the inbox
            mail.select("inbox")

            # Search for emails from the given date
            result, data = mail.search(None, f'(SINCE "{self.archive_date.strftime("%d-%b-%Y")}")')

            if result != 'OK':
                self.log_signal.emit("No messages found!\n")
                return

            deleted_emails = 0
            matched_keywords = {}

            # Iterate through the emails
            for num in data[0].split():
                if self.cancel_event:
                    self.log_signal.emit("Archiving cancelled.\n")
                    break

                try:
                    result, msg_data = mail.fetch(num, "(RFC822)")
                    if result != 'OK':
                        self.log_signal.emit(f"ERROR getting message {num}\n")
                        continue

                    msg = email.message_from_bytes(msg_data[0][1])

                    # Decode the email subject
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")

                    matched = False

                    # Log subject for debugging
                    self.log_signal.emit(f"Subject: {subject}\n")

                    # Check if any keyword is in the email subject
                    for keyword in self.keywords:
                        if keyword.lower().strip() in subject.lower():
                            matched_keywords[keyword] = matched_keywords.get(keyword, 0) + 1
                            matched = True
                            break

                    # Check if the sender is in the selected senders list
                    sender = msg.get("From")
                    if sender and any(s.lower().strip() in sender.lower() for s in self.selected_senders):
                        matched = True

                    if matched:
                        self.log_signal.emit(f"Matched keyword in subject: {subject}\n")
                        mail.store(num, '+X-GM-LABELS', '\\Archive')
                        mail.store(num, '+FLAGS', '\\Deleted')
                        deleted_emails += 1
                        continue

                    # Check if any keyword is in the email body
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                try:
                                    body = decode_email_content(part)
                                except Exception as e:
                                    self.log_signal.emit(f"Error decoding body: {e}\n")
                                    continue
                                # Log body for debugging
                                self.log_signal.emit(f"Body: {body[:100]}\n")  # Log the first 100 characters of the body

                                for keyword in self.keywords:
                                    if keyword.lower().strip() in body.lower():
                                        matched_keywords[keyword] = matched_keywords.get(keyword, 0) + 1
                                        matched = True
                                        break
                            if matched:
                                self.log_signal.emit(f"Matched keyword in body: {subject}\n")
                                mail.store(num, '+X-GM-LABELS', '\\Archive')
                                mail.store(num, '+FLAGS', '\\Deleted')
                                deleted_emails += 1
                                break
                    else:
                        try:
                            body = decode_email_content(msg)
                        except Exception as e:
                            self.log_signal.emit(f"Error decoding body: {e}\n")
                            continue
                        # Log body for debugging
                        self.log_signal.emit(f"Body: {body[:100]}\n")  # Log the first 100 characters of the body

                        for keyword in self.keywords:
                            if keyword.lower().strip() in body.lower():
                                matched_keywords[keyword] = matched_keywords.get(keyword, 0) + 1
                                matched = True
                                break
                        if matched:
                            self.log_signal.emit(f"Matched keyword in body: {subject}\n")
                            mail.store(num, '+X-GM-LABELS', '\\Archive')
                            mail.store(num, '+FLAGS', '\\Deleted')
                            deleted_emails += 1

                except Exception as e:
                    self.log_signal.emit(f"Exception occurred: {str(e)}\n")

            # Expunge the emails marked for deletion
            mail.expunge()

            # Create a summary message
            summary = f"Total emails archived: {deleted_emails}\n\n"
            for keyword, count in matched_keywords.items():
                summary += f"'{keyword}': {count} emails\n"

            self.log_signal.emit(summary)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")

    def delete_draft_emails(self, mail):
        try:
            # Select the Drafts mailbox
            mail.select('"[Gmail]/Drafts"')

            # Search for all draft emails
            result, data = mail.search(None, "ALL")

            if result != 'OK':
                self.log_signal.emit("No draft messages found!\n")
                return

            deleted_emails = 0

            # Iterate through the draft emails
            for num in data[0].split():
                if self.cancel_event:
                    self.log_signal.emit("Deleting drafts cancelled.\n")
                    break

                try:
                    mail.store(num, '+FLAGS', '\\Deleted')
                    deleted_emails += 1

                except Exception as e:
                    self.log_signal.emit(f"Exception occurred: {str(e)}\n")

            # Expunge the emails marked for deletion
            mail.expunge()

            # Create a summary message
            summary = f"Total drafts deleted: {deleted_emails}\n"
            self.log_signal.emit(summary)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")

    def collect_senders(self, mail):
        try:
            # Select the inbox
            mail.select("inbox")

            # Search for emails from the given date
            result, data = mail.search(None, f'(SINCE "{self.archive_date.strftime("%d-%b-%Y")}")')

            if result != 'OK':
                self.log_signal.emit("No messages found!\n")
                return

            senders = set()

            # Iterate through the emails
            for num in data[0].split():
                if self.cancel_event:
                    self.log_signal.emit("Collecting senders cancelled.\n")
                    break

                try:
                    result, msg_data = mail.fetch(num, "(RFC822)")
                    if result != 'OK':
                        self.log_signal.emit(f"ERROR getting message {num}\n")
                        continue

                    msg = email.message_from_bytes(msg_data[0][1])
                    sender = msg.get("From")
                    if sender:
                        senders.add(sender)

                except Exception as e:
                    self.log_signal.emit(f"Exception occurred: {str(e)}\n")

            self.senders_signal.emit(list(senders))
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")

    def cancel(self):
        self.cancel_event = True

class EmailArchiverApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Email Archiver")
        self.setGeometry(100, 100, 1200, 800)  # Increase default window size
        self.setMinimumSize(1200, 800)

        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#F0F0F0"))
        palette.setColor(QPalette.WindowText, QColor("#333333"))
        palette.setColor(QPalette.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.AlternateBase, QColor("#F0F0F0"))
        palette.setColor(QPalette.ToolTipBase, QColor("#FFFFD7"))
        palette.setColor(QPalette.ToolTipText, QColor("#000000"))
        palette.setColor(QPalette.Text, QColor("#333333"))
        palette.setColor(QPalette.Button, QColor("#28a745"))  # Green color for start button
        palette.setColor(QPalette.ButtonText, QColor("#FFFFFF"))
        palette.setColor(QPalette.BrightText, QColor("#FFFFFF"))
        palette.setColor(QPalette.Highlight, QColor("#007ACC"))
        palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))

        self.setPalette(palette)

        font = QFont()
        font.setPointSize(12)
        self.setFont(font)

        main_layout = QHBoxLayout()
        main_layout.setSpacing(15)  # Set spacing between widgets

        left_layout = QVBoxLayout()
        left_layout.setSpacing(15)

        # Search Emails Section
        search_label = QLabel("Search Emails:")
        left_layout.addWidget(search_label)

        # Collect Senders Button
        self.collect_senders_button = QPushButton("Collect Senders", self)
        self.collect_senders_button.setStyleSheet("background-color: #007ACC; color: #FFFFFF; border-radius: 10px; padding: 10px;")
        self.collect_senders_button.clicked.connect(self.collect_senders)
        left_layout.addWidget(self.collect_senders_button)

        # Sender List
        self.sender_list = QListWidget(self)
        self.sender_list.setSelectionMode(QListWidget.MultiSelection)
        self.sender_list.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        left_layout.addWidget(self.sender_list)

        # Start Archiving Button
        self.archive_button = QPushButton("Start Archiving", self)
        self.archive_button.setStyleSheet("background-color: #28a745; color: #FFFFFF; border-radius: 10px; padding: 10px;")  # Green color for start button
        self.archive_button.clicked.connect(self.start_archiving)
        left_layout.addWidget(self.archive_button)

        # Delete Draft Emails Button
        self.delete_drafts_button = QPushButton("Delete Draft Emails", self)
        self.delete_drafts_button.setStyleSheet("background-color: #dc3545; color: #FFFFFF; border-radius: 10px; padding: 10px;")
        self.delete_drafts_button.clicked.connect(self.delete_drafts)
        left_layout.addWidget(self.delete_drafts_button)

        # Cancel Archiving Button
        self.cancel_button = QPushButton("Cancel Archiving", self)
        self.cancel_button.setStyleSheet("background-color: #dc3545; color: #FFFFFF; border-radius: 10px; padding: 10px;")
        self.cancel_button.clicked.connect(self.cancel_archiving)
        left_layout.addWidget(self.cancel_button)

        # Log Display
        self.logs = QTextEdit(self)
        self.logs.setReadOnly(True)
        self.logs.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        left_layout.addWidget(self.logs)

        right_layout = QVBoxLayout()
        right_layout.setSpacing(15)

        # IMAP Server
        imap_server_layout = QVBoxLayout()
        self.imap_server_label = QLabel("IMAP Server:")
        self.imap_server_label.setToolTip("The IMAP server of your email provider (e.g., imap.gmail.com for Gmail).\nFor other providers, refer to their documentation.")
        self.imap_server_input = QLineEdit(self)
        self.imap_server_input.setPlaceholderText("imap.gmail.com")
        self.imap_server_input.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        imap_server_layout.addWidget(self.imap_server_label)
        imap_server_layout.addWidget(self.imap_server_input)
        right_layout.addLayout(imap_server_layout)

        # IMAP Port
        imap_port_layout = QVBoxLayout()
        self.imap_port_label = QLabel("IMAP Port:")
        self.imap_port_label.setToolTip("The IMAP port of your email provider (e.g., 993 for Gmail).\nFor other providers, refer to their documentation.")
        self.imap_port_input = QLineEdit(self)
        self.imap_port_input.setPlaceholderText("993")
        self.imap_port_input.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        imap_port_layout.addWidget(self.imap_port_label)
        imap_port_layout.addWidget(self.imap_port_input)
        right_layout.addLayout(imap_port_layout)

        # Email
        email_layout = QVBoxLayout()
        self.email_label = QLabel("Email:")
        self.email_label.setToolTip("Your email address.")
        self.email_input = QLineEdit(self)
        self.email_input.setPlaceholderText("your_email@gmail.com")
        self.email_input.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        email_layout.addWidget(self.email_label)
        email_layout.addWidget(self.email_input)
        right_layout.addLayout(email_layout)

        # App Password
        app_password_layout = QVBoxLayout()
        self.password_label = QLabel("App Password:")
        self.password_label.setToolTip("Your app-specific password.\nFor Gmail: https://support.google.com/accounts/answer/185833\nFor other providers, refer to their documentation.")
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        app_password_layout.addWidget(self.password_label)
        app_password_layout.addWidget(self.password_input)
        right_layout.addLayout(app_password_layout)

        # Keywords
        keywords_layout = QVBoxLayout()
        self.keywords_label = QLabel("Keywords (comma-separated):")
        self.keywords_label.setToolTip("The keywords to search for in the emails (e.g., 'unfortunately, thank you for your interest').")
        self.keywords_input = QLineEdit(self)
        self.keywords_input.setPlaceholderText("unfortunately, thank you for your interest")
        self.keywords_input.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        keywords_layout.addWidget(self.keywords_label)
        keywords_layout.addWidget(self.keywords_input)
        right_layout.addLayout(keywords_layout)

        # Archive Emails Since Date Picker
        date_layout = QVBoxLayout()
        self.date_label = QLabel("Archive Emails Since:")
        self.date_picker = QDateEdit(self)
        self.date_picker.setDate(QDate.currentDate().addDays(-7))
        self.date_picker.setCalendarPopup(True)
        date_layout.addWidget(self.date_label)
        date_layout.addWidget(self.date_picker)
        right_layout.addLayout(date_layout)

        # Save Configuration Button
        self.save_button = QPushButton("Save Configuration", self)
        self.save_button.setStyleSheet("background-color: #007ACC; color: #FFFFFF; border-radius: 10px; padding: 10px;")
        self.save_button.clicked.connect(self.save_configuration)
        right_layout.addWidget(self.save_button)

        # Load Configuration Label
        self.load_label = QLabel("Load Configuration:")
        right_layout.addWidget(self.load_label)

        # Configuration Dropdown
        self.config_dropdown = QComboBox(self)
        self.config_dropdown.setStyleSheet("color: #333333; border-radius: 10px; padding: 5px;")
        self.config_dropdown.currentIndexChanged.connect(self.load_configuration)
        right_layout.addWidget(self.config_dropdown)
        self.update_config_dropdown()

        main_layout.addLayout(left_layout, 1)
        main_layout.addLayout(right_layout, 1)

        self.setLayout(main_layout)
        self.show()

    def save_configuration(self):
        current_config_name = self.config_dropdown.currentText()
        if current_config_name:
            message_box = QMessageBox(self)
            message_box.setIcon(QMessageBox.Question)
            message_box.setWindowTitle("Save Configuration")
            message_box.setText(f'Do you want to overwrite the configuration "{current_config_name}" or save as a new configuration?')
            message_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            yes_button = message_box.button(QMessageBox.Yes)
            no_button = message_box.button(QMessageBox.No)
            yes_button.setText("Overwrite")
            no_button.setText("Save As New")
            message_box.exec_()

            if message_box.clickedButton() == yes_button:
                config_name = current_config_name
            else:
                config_name, ok = QInputDialog.getText(self, 'Save Configuration', 'Enter configuration name:', QLineEdit.Normal, "", flags=Qt.WindowFlags())
                if not ok or not config_name:
                    return
        else:
            config_name, ok = QInputDialog.getText(self, 'Save Configuration', 'Enter configuration name:', QLineEdit.Normal, "", flags=Qt.WindowFlags())
            if not ok or not config_name:
                return

        config = {
            "imap_server": self.imap_server_input.text(),
            "imap_port": self.imap_port_input.text(),
            "email": self.email_input.text(),
            "app_password": self.password_input.text(),
            "keywords": self.keywords_input.text()
        }

        try:
            with open(CONFIG_FILE, 'r') as file:
                configurations = json.load(file)
        except FileNotFoundError:
            configurations = {}

        configurations[config_name] = config
        with open(CONFIG_FILE, 'w') as file:
            json.dump(configurations, file, indent=4)
        self.update_config_dropdown()

    def load_configuration(self):
        config_name = self.config_dropdown.currentText()
        if config_name:
            with open(CONFIG_FILE, 'r') as file:
                configurations = json.load(file)
            config = configurations.get(config_name, {})
            self.imap_server_input.setText(config.get("imap_server", ""))
            self.imap_port_input.setText(config.get("imap_port", ""))
            self.email_input.setText(config.get("email", ""))
            self.password_input.setText(config.get("app_password", ""))
            self.keywords_input.setText(config.get("keywords", ""))

    def update_config_dropdown(self):
        self.config_dropdown.clear()
        try:
            with open(CONFIG_FILE, 'r') as file:
                configurations = json.load(file)
            self.config_dropdown.addItems(configurations.keys())
        except FileNotFoundError:
            pass

    def start_archiving(self):
        imap_server = self.imap_server_input.text()
        imap_port = int(self.imap_port_input.text())
        username = self.email_input.text()
        password = self.password_input.text()
        keywords = self.keywords_input.text().split(',')
        archive_date = self.date_picker.date().toPyDate()

        # Get selected senders
        selected_senders = [item.text() for item in self.sender_list.selectedItems()]

        # Ensure any existing thread is properly cleaned up before starting a new one
        if hasattr(self, 'archiving_thread') and self.archiving_thread.isRunning():
            self.logs.append("Previous archiving thread is still running.\n")
            return

        self.logs.clear()
        self.logs.append("Archiving started...\n")

        self.archiving_thread = ArchiverThread(imap_server, imap_port, username, password, keywords, archive_date, "archive")
        self.archiving_thread.selected_senders = selected_senders  # Set selected_senders
        self.archiving_thread.log_signal.connect(self.logs.append)
        self.archiving_thread.finished_signal.connect(self.archiving_finished)
        self.archiving_thread.start()

    def delete_drafts(self):
        imap_server = self.imap_server_input.text()
        imap_port = int(self.imap_port_input.text())
        username = self.email_input.text()
        password = self.password_input.text()

        # Ensure any existing thread is properly cleaned up before starting a new one
        if hasattr(self, 'archiving_thread') and self.archiving_thread.isRunning():
            self.logs.append("Previous archiving thread is still running.\n")
            return

        self.logs.clear()
        self.logs.append("Deleting drafts started...\n")

        self.archiving_thread = ArchiverThread(imap_server, imap_port, username, password, [], None, "delete_drafts")
        self.archiving_thread.log_signal.connect(self.logs.append)
        self.archiving_thread.finished_signal.connect(self.archiving_finished)
        self.archiving_thread.start()

    def collect_senders(self):
        imap_server = self.imap_server_input.text()
        imap_port = int(self.imap_port_input.text())
        username = self.email_input.text()
        password = self.password_input.text()
        archive_date = self.date_picker.date().toPyDate()

        # Ensure any existing thread is properly cleaned up before starting a new one
        if hasattr(self, 'archiving_thread') and self.archiving_thread.isRunning():
            self.logs.append("Previous archiving thread is still running.\n")
            return

        self.logs.clear()
        self.logs.append("Collecting senders started...\n")

        self.archiving_thread = ArchiverThread(imap_server, imap_port, username, password, [], archive_date, "collect_senders")
        self.archiving_thread.log_signal.connect(self.logs.append)
        self.archiving_thread.senders_signal.connect(self.populate_sender_list)
        self.archiving_thread.finished_signal.connect(self.archiving_finished)
        self.archiving_thread.start()

    def populate_sender_list(self, senders):
        self.sender_list.clear()
        self.sender_list.addItems(senders)

    def cancel_archiving(self):
        if hasattr(self, 'archiving_thread') and self.archiving_thread.isRunning():
            self.archiving_thread.cancel()
            self.logs.append("Cancelling archiving process...\n")

    def archiving_finished(self):
        self.logs.append("Archiving finished.\n")

def main():
    app = QApplication(sys.argv)
    ex = EmailArchiverApp()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
