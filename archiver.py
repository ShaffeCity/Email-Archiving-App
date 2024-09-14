import imaplib
import email
from email.header import decode_header
from PyQt5.QtCore import QThread, pyqtSignal

def decode_email_content(part):
    try:
        charset = part.get_content_charset()
        if not charset:
            charset = 'utf-8'  # default to utf-8 if charset is not specified
        return part.get_payload(decode=True).decode(charset, errors='ignore')
    except UnicodeDecodeError:
        return part.get_payload(decode=True).decode('latin-1', errors='ignore')

class Archiver(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    cancel_event = False

    def __init__(self, imap_server, imap_port, username, password, keywords, archive_date, selected_senders):
        super().__init__()
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.username = username
        self.password = password
        self.keywords = keywords
        self.archive_date = archive_date
        self.selected_senders = selected_senders

    def run(self):
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            self.archive_emails(mail)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")
        finally:
            if mail is not None and mail.state != 'LOGOUT':
                try:
                    mail.logout()
                    self.log_signal.emit("Logout successful.")
                except Exception as e:
                    self.log_signal.emit(f"Exception occurred during logout: {str(e)}\n")
            self.finished_signal.emit()

    def archive_emails(self, mail):
        try:
            mail.select("inbox")
            result, data = mail.search(None, f'(SINCE "{self.archive_date.strftime("%d-%b-%Y")}")')
            if result != 'OK':
                self.log_signal.emit("No messages found!\n")
                return

            deleted_emails = 0
            matched_keywords = {}

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
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")

                    matched = False
                    self.log_signal.emit(f"Subject: {subject}\n")

                    for keyword in self.keywords:
                        if keyword.lower().strip() in subject.lower():
                            matched_keywords[keyword] = matched_keywords.get(keyword, 0) + 1
                            matched = True
                            break

                    sender = msg.get("From")
                    if sender and any(s.lower().strip() in sender.lower() for s in self.selected_senders):
                        matched = True

                    if matched:
                        self.log_signal.emit(f"Matched keyword in subject: {subject}\n")
                        mail.store(num, '+X-GM-LABELS', '\\Archive')
                        mail.store(num, '+FLAGS', '\\Deleted')
                        deleted_emails += 1
                        continue

                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                try:
                                    body = decode_email_content(part)
                                except Exception as e:
                                    self.log_signal.emit(f"Error decoding body: {e}\n")
                                    continue
                                self.log_signal.emit(f"Body: {body[:100]}\n")

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
                        self.log_signal.emit(f"Body: {body[:100]}\n")

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

            mail.expunge()
            summary = f"Total emails archived: {deleted_emails}\n\n"
            for keyword, count in matched_keywords.items():
                summary += f"'{keyword}': {count} emails\n"

            self.log_signal.emit(summary)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")
