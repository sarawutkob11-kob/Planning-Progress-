"""Email senders. Outlook desktop via COM (no Graph / MSAL).

OutlookSender attaches to an Outlook that is ALREADY running and never starts it.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Protocol


@dataclass
class Mail:
    to: str
    cc: list[str] = field(default_factory=list)
    subject: str = ""
    html: str = ""


class SendError(Exception):
    pass


class OutlookNotRunning(SendError):
    pass


class Sender(Protocol):
    name: str
    delivers: bool          # False = nothing leaves the PC (preview only)

    def check(self) -> None:
        """Raise OutlookNotRunning (or SendError) if sending is impossible right now."""

    def send(self, mail: Mail) -> None: ...


class OutlookSender:
    name = "outlook"
    delivers = True

    def __init__(self, account: str = ""):
        self.account = account.strip().lower()
        self._app = None

    def _outlook(self):
        if self._app is not None:
            return self._app
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:            # pragma: no cover - Windows only
            raise SendError("pywin32 is not installed") from exc
        pythoncom.CoInitialize()
        try:
            self._app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as exc:              # pywintypes.com_error: not running
            raise OutlookNotRunning("Outlook is not running on this PC") from exc
        return self._app

    def _account(self, app):
        if not self.account:
            return None
        for acct in app.Session.Accounts:
            if str(acct.SmtpAddress).lower() == self.account:
                return acct
        raise SendError(f"Outlook account {self.account!r} not found")

    def check(self) -> None:
        app = self._outlook()
        self._account(app)

    def send(self, mail: Mail) -> None:
        app = self._outlook()
        try:
            item = app.CreateItem(0)              # olMailItem
            item.To = mail.to
            if mail.cc:
                item.CC = "; ".join(mail.cc)
            item.Subject = mail.subject
            item.HTMLBody = mail.html
            acct = self._account(app)
            if acct is not None:
                item.SendUsingAccount = acct
            item.Send()
        except SendError:
            raise
        except Exception as exc:
            raise SendError(f"Outlook send failed: {exc}") from exc


class PreviewOnlySender:
    """Sends nothing. Used with --preview-only and in dry-run when no pmo_email is set."""
    name = "preview"
    delivers = False

    def check(self) -> None:
        pass

    def send(self, mail: Mail) -> None:
        pass


class FakeSender:
    """Test double: records mails; can simulate Outlook down or a send failure."""
    name = "fake"
    delivers = True

    def __init__(self, running: bool = True, fail_for: set[str] | None = None):
        self.running = running
        self.fail_for = fail_for or set()
        self.sent: list[Mail] = []

    def check(self) -> None:
        if not self.running:
            raise OutlookNotRunning("Outlook is not running on this PC")

    def send(self, mail: Mail) -> None:
        self.check()
        if mail.to in self.fail_for:
            raise SendError(f"simulated failure for {mail.to}")
        self.sent.append(mail)
