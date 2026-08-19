#!/usr/bin/env python3
"""Validate SMTP protocol log parsing used by filebeat-smtp.yml."""

from __future__ import annotations

import csv
import io
import re
from pathlib import Path

SAMPLE = Path(__file__).resolve().parent / "samples" / "smtp-protocol.log"
CONFIG = Path(__file__).resolve().parent / "filebeat-smtp.yml"

EVENT_NAMES = {
    "+": "Connect",
    "-": "Disconnect",
    ">": "Send",
    "<": "Receive",
    "*": "Information",
}

EXPECTED_PATHS = [
    r"C:\Queue\TransportLogs\FrontEnd\ProtocolLog\SmtpSend\*.log",
    r"C:\Queue\TransportLogs\FrontEnd\ProtocolLog\SmtpReceive\*.log",
    r"C:\Queue\TransportLogs\Hub\ProtocolLog\SmtpSend\*.log",
    r"C:\Queue\TransportLogs\Hub\ProtocolLog\SmtpReceive\*.log",
]


def parse_smtp_line(line: str) -> dict[str, object]:
    """Mirror Filebeat decode_csv_fields + extra-column join for data."""
    reader = csv.reader(io.StringIO(line), delimiter=",")
    arr = next(reader)
    if len(arr) < 7:
        raise ValueError(f"too few columns: {arr!r}")

    record = {
        "date-time": arr[0],
        "connector-id": arr[1],
        "session-id": arr[2],
        "sequence-number": int(arr[3]) if arr[3] else None,
        "local-endpoint": arr[4],
        "remote-endpoint": arr[5],
        "event": arr[6],
        "data": arr[7] if len(arr) > 7 else "",
        "context": arr[8] if len(arr) > 8 else "",
    }
    if len(arr) > 9:
        record["data"] = ",".join(arr[7:-1])
        record["context"] = arr[-1]
    record["event_name"] = EVENT_NAMES.get(str(record["event"]))
    record["timestamp"] = record["date-time"]
    for src, ip_key, port_key in (
        ("local-endpoint", "local-ip", "local-port"),
        ("remote-endpoint", "remote-ip", "remote-port"),
    ):
        endpoint = record[src]
        if not endpoint:
            continue
        idx = endpoint.rfind(":")
        if idx <= 0:
            continue
        ip = endpoint[:idx]
        port = endpoint[idx + 1 :]
        if ip.startswith("[") and ip.endswith("]"):
            ip = ip[1:-1]
        record[ip_key] = ip
        record[port_key] = int(port) if port.isdigit() else port
    return record


def test_sample_log() -> None:
    lines = [
        line.rstrip("\n")
        for line in SAMPLE.read_text(encoding="utf-8").splitlines()
        if line and not line.startswith("#")
    ]
    assert len(lines) == 11, f"unexpected sample size: {len(lines)}"

    records = [parse_smtp_line(line) for line in lines]
    assert records[0]["event_name"] == "Connect"
    assert records[0]["local-ip"] == "10.0.0.10"
    assert records[0]["local-port"] == 25
    assert records[0]["remote-ip"] == "203.0.113.50"
    assert records[0]["data"] == ""
    assert records[0]["context"] == ""

    banner = records[1]
    assert banner["event_name"] == "Send"
    assert banner["data"].startswith("220 mail.contoso.com")
    assert "Tue, 18 Aug 2026" in str(banner["data"])
    assert banner["context"] == ""

    queued = records[8]
    assert "[InternalId=123, Hostname=MAILBOX01]" in str(queued["data"])
    assert queued["context"] == ""

    disconnect = records[9]
    assert disconnect["event_name"] == "Disconnect"
    assert disconnect["context"] == "Local"

    outbound = records[10]
    assert outbound["event_name"] == "Information"
    assert outbound["local-endpoint"] == ""
    assert outbound["remote-ip"] == "64.8.70.48"
    assert outbound["context"] == "attempting to connect"
    assert outbound["timestamp"] == "2026-08-18T07:00:02.000Z"


def test_config_contains_required_bits() -> None:
    text = CONFIG.read_text(encoding="utf-8")
    for path in EXPECTED_PATHS:
        assert path in text, f"missing path {path}"
    for field in (
        "smtp.protocol.date-time",
        "smtp.protocol.connector-id",
        "smtp.protocol.session-id",
        "smtp.protocol.sequence-number",
        "smtp.protocol.local-endpoint",
        "smtp.protocol.remote-endpoint",
        "smtp.protocol.event",
        "smtp.protocol.data",
        "smtp.protocol.context",
    ):
        assert field in text, f"missing mapping {field}"
    assert "separator: \",\"" in text
    assert "exclude_lines: ['^#']" in text
    assert "topic: \"KE_contoso\"" in text
    assert re.search(r"id: smtp-frontend-send", text)
    assert re.search(r"id: smtp-hub-receive", text)
    assert "arr.slice(7, arr.length - 1).join(\",\")" in text
    assert text.count("prospector.scanner.fingerprint.enabled: false") == 4
    assert text.count("file_identity.native: ~") == 4


if __name__ == "__main__":
    test_sample_log()
    test_config_contains_required_bits()
    print("ok: SMTP Filebeat parse and config checks passed")
