"""In-memory Outlook simulator for manifest-driven AI tool tests."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from tool_result import ToolResult


def _strip_html(html: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


@dataclass(slots=True)
class MailItem:
    mode: str
    subject: str
    body_html: str
    from_name: str = "Alex Sender"
    from_email: str = "alex.sender@example.com"
    to: list[dict[str, str]] = field(default_factory=list)
    cc: list[dict[str, str]] = field(default_factory=list)
    bcc: list[dict[str, str]] = field(default_factory=list)


class OutlookSimulator:
    """Minimal mailbox simulator with read and compose states."""

    def __init__(self) -> None:
        self.current_item = MailItem(
            mode="read",
            subject="Quarterly update",
            body_html="<p>Hello team, the quarter closed above plan.</p><p>Please review the attached notes.</p>",
            to=[{"displayName": "Finance Team", "emailAddress": "finance@example.com"}],
        )

    def _ok(self, value: Any) -> ToolResult:
        return ToolResult(success=True, value=value)

    def _error(self, message: str) -> ToolResult:
        return ToolResult(success=False, error=message)

    def _format_recipients(self, recipients: list[dict[str, str]]) -> str:
        return ", ".join(
            f"{recipient.get('displayName', recipient['emailAddress'])} <{recipient['emailAddress']}>"
            for recipient in recipients
        ) or "(none)"

    def get_mail_item(self) -> ToolResult:
        item = self.current_item
        lines = ["Mail Item Overview", "=" * 40, f"Subject: {item.subject}"]
        if item.mode == "read":
            lines.append(f"From: {item.from_name} <{item.from_email}>")
        lines.append(f"To: {self._format_recipients(item.to)}")
        lines.append(f"Mode: {item.mode}")
        return self._ok("\n".join(lines))

    def get_mail_body(self, format: str | None = None) -> ToolResult:  # noqa: A002
        if format == "html":
            return self._ok(self.current_item.body_html)
        return self._ok(_strip_html(self.current_item.body_html))

    def display_new_message(
        self,
        toRecipients: list[str] | None = None,  # noqa: N803
        ccRecipients: list[str] | None = None,  # noqa: N803
        subject: str | None = None,
        htmlBody: str | None = None,  # noqa: N803
    ) -> ToolResult:
        self.current_item = MailItem(
            mode="compose",
            subject=subject or "",
            body_html=htmlBody or "",
            to=[{"displayName": email, "emailAddress": email} for email in (toRecipients or [])],
            cc=[{"displayName": email, "emailAddress": email} for email in (ccRecipients or [])],
        )
        return self._ok("New message form opened.")

    def set_mail_subject(self, subject: str) -> ToolResult:
        if self.current_item.mode != "compose":
            return self._error("Cannot set subject in read mode.")
        self.current_item.subject = subject
        return self._ok(f'Subject set to: "{subject}"')

    def set_mail_body(self, content: str, format: str | None = None) -> ToolResult:  # noqa: A002
        if self.current_item.mode != "compose":
            return self._error("Cannot set body in read mode.")
        self.current_item.body_html = content if format != "text" else f"<p>{content}</p>"
        return self._ok("Mail body updated successfully.")

    def add_mail_recipient(self, field: str, recipients: list[dict[str, str]]) -> ToolResult:
        if self.current_item.mode != "compose":
            return self._error("Cannot add recipients in read mode.")
        bucket = getattr(self.current_item, field, None)
        if bucket is None or not isinstance(bucket, list):
            return self._error(f"Unknown recipient field: {field}")
        bucket.extend(
            {
                "displayName": recipient.get("displayName", recipient["emailAddress"]),
                "emailAddress": recipient["emailAddress"],
            }
            for recipient in recipients
        )
        return self._ok(f"Added {len(recipients)} recipient(s) to {field}.")

    def reply_to_mail(self, htmlBody: str, replyAll: bool | None = None) -> ToolResult:  # noqa: N803
        original = self.current_item
        if original.mode != "read":
            return self._error("Reply only works in read mode.")
        recipients = original.to if replyAll else [{"displayName": original.from_name, "emailAddress": original.from_email}]
        self.current_item = MailItem(
            mode="compose",
            subject=f"RE: {original.subject}",
            body_html=htmlBody,
            to=recipients,
        )
        return self._ok(f"Reply{' all' if replyAll else ''} form opened with provided content.")
