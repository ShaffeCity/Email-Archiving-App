import imaplib
from PyQt5.QtCore import QThread, pyqtSignal

class Deleter(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    cancel_event = False

    def __init__(self, imap_server, imap_port, username, password):
        super().__init__()
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.username = username
        self.password = password

    def run(self):
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.username, self.password)
            self.delete_draft_emails(mail)
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

    def delete_draft_emails(self, mail):
        try:
            mail.select('"[Gmail]/Drafts"')
            result, data = mail.search(None, "ALL")
            if result != 'OK':
                self.log_signal.emit("No draft messages found!\n")
                return

            deleted_emails = 0

            for num in data[0].split():
                if self.cancel_event:
                    self.log_signal.emit("Deleting drafts cancelled.\n")
                    break
                try:
                    mail.store(num, '+FLAGS', '\\Deleted')
                    deleted_emails += 1
                except Exception as e:
                    self.log_signal.emit(f"Exception occurred: {str(e)}\n")

            mail.expunge()
            summary = f"Total drafts deleted: {deleted_emails}\n"
            self.log_signal.emit(summary)
        except Exception as e:
            self.log_signal.emit(f"Exception occurred: {str(e)}\n")
