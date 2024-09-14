import email
import imaplib

from PyQt5.QtCore import QThread, pyqtSignal


class Fetcher(QThread):
    log_signal = pyqtSignal(str)
    senders_signal = pyqtSignal(list)
    cancel_event = False

    def __init__(self, imap_server, imap_port, username, password, archive_date):
        super().__init__()
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.username = username
        self.password = password
        self.archive_date = archive_date

    def run(self):
        try:
            self.log_signal.emit("Starting sender collection...")
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            self.collect_senders(mail)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")
        finally:
            if mail is not None and mail.state != 'LOGOUT':
                try:
                    mail.logout()
                    self.log_signal.emit("Logout successful.")
                except Exception as e:
                    self.log_signal.emit(f"Exception occurred during logout: {str(e)}\n")

    def collect_senders(self, mail):
        try:
            mail.select("inbox")
            result, data = mail.search(None, f'(SINCE "{self.archive_date.strftime("%d-%b-%Y")}")')
            if result != 'OK':
                self.log_signal.emit("No messages found!\n")
                return

            senders = set()
            nums = data[0].split()

            for num in nums:
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
