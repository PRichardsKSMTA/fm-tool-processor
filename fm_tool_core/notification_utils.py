from __future__ import annotations

"""Helpers for sending notifications via SMTP or webhooks."""

import logging
import os
from email.message import EmailMessage
from pathlib import Path
import smtplib
from typing import Any

try:  # pragma: no cover - optional dependency
    import requests  # type: ignore
except Exception as _e:  # pragma: no cover
    requests = None  # type: ignore
    logging.basicConfig(level=logging.WARNING)
    logging.warning("requests missing – webhooks disabled (%s)", _e)

log = logging.getLogger(__name__)

from .constants import BID_WEBHOOK_URI


def _send(msg: EmailMessage) -> None:
    """Send *msg* using environment configured SMTP settings."""

    server = os.getenv("SMTP_SERVER")
    port = os.getenv("SMTP_PORT")
    user = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")

    if not (server and port):
        log.warning("SMTP_SERVER or SMTP_PORT not configured")
        return

    try:
        with smtplib.SMTP(server, int(port), timeout=10) as smtp:
            if user and password:
                try:
                    smtp.starttls()
                except Exception:
                    pass
                try:
                    smtp.login(user, password)
                except Exception as exc:
                    log.warning("SMTP login failed: %s", exc)
            smtp.send_message(msg)
    except Exception as exc:
        log.warning("Email send failed: %s", exc)


def send_success_email(
    to_addr: str,
    file_name: str,
    sharepoint_url: str,
    attachment_path: str | Path,
) -> None:
    """Send a success notification email with optional attachment."""

    from_addr = os.getenv("SMTP_FROM")
    if not from_addr:
        log.warning("SMTP_FROM not configured")
        return

    msg = EmailMessage()
    msg["Subject"] = "FM Tool processing succeeded"
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(
        "\n".join(
            [
                f"File {file_name} processed successfully.",
                f"Uploaded to {sharepoint_url}.",
            ]
        )
    )

    path = Path(attachment_path)
    try:
        data = path.read_bytes()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=path.name,
        )
    except Exception as exc:
        log.warning("Could not attach %s: %s", path, exc)

    _send(msg)


def send_failure_email(to_addr: str, error_msg: str) -> None:
    """Send a failure notification email."""

    from_addr = os.getenv("SMTP_FROM")
    if not from_addr:
        log.warning("SMTP_FROM not configured")
        return

    msg = EmailMessage()
    msg["Subject"] = "FM Tool processing failed"
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(f"Processing failed with error:\n{error_msg}")

    _send(msg)


def send_bid_webhook(
    to_addr: str,
    file_name: str,
    sharepoint_url: str,
    message: str,
    extra: dict[str, Any] | None = None,
) -> None:
    """POST a notification to the BID Power Automate webhook."""

    if requests is None:
        log.warning("requests not available")
        return
    if not BID_WEBHOOK_URI:
        log.warning("BID_WEBHOOK_URI not configured")
        return

    payload: dict[str, Any] = {
        "email": to_addr,
        "file_name": file_name,
        "sharepoint_url": sharepoint_url,
        "message": message,
    }
    if extra:
        payload.update(extra)

    try:
        resp = requests.post(BID_WEBHOOK_URI, json=payload, timeout=10)
        if resp.status_code >= 400:
            log.warning(
                "Webhook POST failed: %s %s",
                resp.status_code,
                resp.text,
            )
    except Exception as exc:  # pragma: no cover - network errors
        log.warning("Webhook POST failed: %s", exc)


__all__ = ["send_success_email", "send_failure_email", "send_bid_webhook"]
