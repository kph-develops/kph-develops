"""
Email reporter for HP Printer Collector.

Builds a multipart (plain-text + HTML) email report and sends it via
SMTP with STARTTLS.  If TLS is unavailable the code falls back to an
unencrypted connection so that local/relay SMTP servers still work.

Alert rows are highlighted in the HTML report and called out in the
plain-text version whenever a toner level falls below the configured
threshold (default 20 %).
"""

import logging
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional

logger = logging.getLogger(__name__)

# Regex to pull the numeric part out of a percentage string such as
# "75%", "<10%", "100 %", or "15%*".
_PERCENT_NUM_RE = re.compile(r"(\d+)\s*%")


# ---------------------------------------------------------------------------
# Alert helpers
# ---------------------------------------------------------------------------


def _toner_numeric(value: Optional[str]) -> Optional[int]:
    """
    Convert a toner string like '75%' or '<10%' to an integer.

    Returns None if the string is missing or unparseable.
    The '<' prefix is intentionally ignored; '<10%' is treated as 10.
    """
    if not value:
        return None
    match = _PERCENT_NUM_RE.search(value)
    return int(match.group(1)) if match else None


def _build_alert_lines(result: dict, threshold: int) -> List[str]:
    """Return a list of human-readable alert strings for toners below threshold."""
    alerts = []
    for colour in ("black", "cyan", "yellow", "magenta"):
        raw = result.get(f"toner_{colour}")
        level = _toner_numeric(raw)
        if level is not None and level < threshold:
            alerts.append(f"{colour.capitalize()} toner is LOW: {raw} (threshold {threshold}%)")
    return alerts


# ---------------------------------------------------------------------------
# Plain-text body
# ---------------------------------------------------------------------------


def _build_plain_text(printer_results: List[dict], report_date: str, threshold: int) -> str:
    """Compose the plain-text email body."""
    lines = [
        "HP PRINTER MONTHLY USAGE REPORT",
        f"Report date : {report_date}",
        "=" * 56,
        "",
    ]

    for result in printer_results:
        name = result.get("name", result.get("ip", "Unknown"))
        ip = result.get("ip", "N/A")
        error = result.get("error")

        lines += [f"Printer : {name}", f"IP      : {ip}", ""]

        if error:
            lines += [f"  ERROR: {error}", ""]
        else:
            page_count = result.get("page_count")
            lines += [
                f"  Page count    : {page_count if page_count is not None else 'N/A':,}",
                "  Toner levels",
                f"    Black   : {result.get('toner_black') or 'N/A'}",
                f"    Cyan    : {result.get('toner_cyan') or 'N/A'}",
                f"    Yellow  : {result.get('toner_yellow') or 'N/A'}",
                f"    Magenta : {result.get('toner_magenta') or 'N/A'}",
                "",
            ]

            # Low-toner alerts
            alerts = _build_alert_lines(result, threshold)
            if alerts:
                lines.append("  *** ALERTS ***")
                for alert in alerts:
                    lines.append(f"    ! {alert}")
                lines.append("")

        lines.append("-" * 56)
        lines.append("")

    lines.append("This report was generated automatically by HP Printer Collector.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# HTML body
# ---------------------------------------------------------------------------

_HTML_STYLE = """
<style>
  body { font-family: Arial, sans-serif; font-size: 14px; color: #333; }
  h1   { color: #0073e6; }
  h2   { color: #444; border-bottom: 1px solid #ccc; padding-bottom: 4px; }
  table { border-collapse: collapse; width: 100%; max-width: 560px; margin-bottom: 20px; }
  th    { background: #0073e6; color: white; padding: 8px 12px; text-align: left; }
  td    { padding: 7px 12px; border-bottom: 1px solid #e0e0e0; }
  tr:nth-child(even) td { background: #f7f9fc; }
  .ok    { color: #2e7d32; font-weight: bold; }
  .low   { color: #b71c1c; font-weight: bold; }
  .alert-box { background: #fff3e0; border-left: 4px solid #e65100;
               padding: 10px 14px; margin: 10px 0; border-radius: 2px; }
  .error-box { background: #fce4ec; border-left: 4px solid #c62828;
               padding: 10px 14px; margin: 10px 0; border-radius: 2px; }
  .footer { color: #888; font-size: 12px; margin-top: 30px; }
</style>
"""


def _toner_cell(value: Optional[str], threshold: int) -> str:
    """Return an HTML table cell with colour-coded toner level."""
    if value is None:
        return "<td>N/A</td>"
    level = _toner_numeric(value)
    css_class = "low" if (level is not None and level < threshold) else "ok"
    return f'<td class="{css_class}">{value}</td>'


def _build_html(printer_results: List[dict], report_date: str, threshold: int) -> str:
    """Compose the HTML email body."""
    rows_html = ""

    for result in printer_results:
        name = result.get("name", result.get("ip", "Unknown"))
        ip = result.get("ip", "N/A")
        error = result.get("error")

        rows_html += f"<h2>{name} &mdash; {ip}</h2>\n"

        if error:
            rows_html += f'<div class="error-box"><strong>Error:</strong> {error}</div>\n'
            continue

        page_count = result.get("page_count")
        page_str = f"{page_count:,}" if isinstance(page_count, int) else "N/A"

        # Toner cells with conditional colouring
        black_td = _toner_cell(result.get("toner_black"), threshold)
        cyan_td = _toner_cell(result.get("toner_cyan"), threshold)
        yellow_td = _toner_cell(result.get("toner_yellow"), threshold)
        magenta_td = _toner_cell(result.get("toner_magenta"), threshold)

        rows_html += f"""
<table>
  <tr><th colspan="2">Usage &amp; Toner Summary</th></tr>
  <tr><td><strong>Page Count</strong></td><td>{page_str}</td></tr>
  <tr><td><strong>Black</strong></td>{black_td}</tr>
  <tr><td><strong>Cyan</strong></td>{cyan_td}</tr>
  <tr><td><strong>Yellow</strong></td>{yellow_td}</tr>
  <tr><td><strong>Magenta</strong></td>{magenta_td}</tr>
</table>
"""

        # Low-toner alert box
        alerts = _build_alert_lines(result, threshold)
        if alerts:
            alert_items = "".join(f"<li>{a}</li>" for a in alerts)
            rows_html += f'<div class="alert-box"><strong>&#9888; Alerts</strong><ul>{alert_items}</ul></div>\n'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8">{_HTML_STYLE}</head>
<body>
  <h1>HP Printer Monthly Usage Report</h1>
  <p><strong>Report date:</strong> {report_date}</p>
  {rows_html}
  <p class="footer">This report was generated automatically by HP Printer Collector.</p>
</body>
</html>"""
    return html


# ---------------------------------------------------------------------------
# SMTP sender
# ---------------------------------------------------------------------------


def send_report(
    printer_results: List[dict],
    smtp_cfg: dict,
    recipients: List[str],
    subject: Optional[str] = None,
    threshold: int = 20,
) -> bool:
    """
    Build and send the monthly report email.

    Args:
        printer_results: List of result dicts from scraper.collect_printer_data().
        smtp_cfg:        Dict with keys: host, port, username, password,
                         from_address, use_tls (bool, default True).
        recipients:      List of recipient email addresses.
        subject:         Email subject line (auto-generated if None).
        threshold:       Toner percentage below which an alert is raised.

    Returns:
        True if the email was sent successfully, False otherwise.
    """
    if not recipients:
        logger.error("No email recipients configured; skipping send")
        return False

    report_date = datetime.now().strftime("%B %d, %Y")  # e.g. "April 01, 2026"

    if subject is None:
        subject = f"HP Printer Usage Report – {report_date}"

    # Build the message
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_cfg.get("from_address", smtp_cfg.get("username", ""))
    msg["To"] = ", ".join(recipients)

    plain_text = _build_plain_text(printer_results, report_date, threshold)
    html_text = _build_html(printer_results, report_date, threshold)

    # Attach plain-text first; email clients prefer the last MIME part they
    # can render, so HTML goes second.
    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_text, "html", "utf-8"))

    host = smtp_cfg.get("host", "localhost")
    port = int(smtp_cfg.get("port", 587))
    use_tls = smtp_cfg.get("use_tls", True)
    username = smtp_cfg.get("username", "")
    password = smtp_cfg.get("password", "")

    logger.info(
        "Sending report to %d recipient(s) via %s:%d (TLS=%s)",
        len(recipients),
        host,
        port,
        use_tls,
    )

    try:
        if use_tls:
            # STARTTLS — connects on plain port then upgrades (port 587 default)
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                if username:
                    server.login(username, password)
                server.sendmail(msg["From"], recipients, msg.as_string())
        else:
            # Plain SMTP (local relay, port 25)
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.ehlo()
                if username:
                    server.login(username, password)
                server.sendmail(msg["From"], recipients, msg.as_string())

        logger.info("Email report sent successfully")
        return True

    except smtplib.SMTPAuthenticationError as exc:
        logger.error("SMTP authentication failed for user '%s': %s", username, exc)
    except smtplib.SMTPConnectError as exc:
        logger.error("Cannot connect to SMTP server %s:%d: %s", host, port, exc)
    except smtplib.SMTPException as exc:
        logger.error("SMTP error while sending report: %s", exc)
    except OSError as exc:
        logger.error("Network error while sending email: %s", exc)

    return False
