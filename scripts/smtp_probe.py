#!/usr/bin/env python3
"""Probe SMTP capabilities without sending credentials or mail."""

from __future__ import annotations

import argparse
import smtplib
import socket
import ssl
import sys
from typing import Iterable


def _decode_smtp_response(response: bytes | str) -> str:
    if isinstance(response, bytes):
        return response.decode("utf-8", errors="replace")
    return response


def _format_features(features: dict[str, str]) -> Iterable[str]:
    for name in sorted(features):
        value = features[name]
        yield f"  - {name.upper()}: {value}" if value else f"  - {name.upper()}"


def _print_capabilities(client: smtplib.SMTP, label: str) -> None:
    features = client.esmtp_features
    auth = features.get("auth", "")
    print(f"\n{label}")
    print(f"STARTTLS advertised: {'yes' if client.has_extn('starttls') else 'no'}")
    print(f"AUTH advertised: {'yes' if client.has_extn('auth') else 'no'}")
    if auth:
        print(f"AUTH mechanisms: {auth}")
    print("Capabilities:")
    if features:
        print("\n".join(_format_features(features)))
    else:
        print("  <none>")


def _connect(args: argparse.Namespace) -> smtplib.SMTP:
    if args.ssl:
        context = ssl.create_default_context()
        return smtplib.SMTP_SSL(
            host=args.host,
            port=args.port,
            timeout=args.timeout,
            context=context,
            local_hostname=args.local_hostname,
        )

    return smtplib.SMTP(
        host=args.host,
        port=args.port,
        timeout=args.timeout,
        local_hostname=args.local_hostname,
    )


def probe(args: argparse.Namespace) -> int:
    try:
        with _connect(args) as client:
            if args.debug:
                client.set_debuglevel(1)

            code, response = client.ehlo()
            print(f"EHLO status: {code}")
            print("EHLO response:")
            print(_decode_smtp_response(response))
            _print_capabilities(client, "Capabilities before STARTTLS")

            if not args.starttls:
                return 0

            if not client.has_extn("starttls"):
                print("\nSTARTTLS was requested but is not advertised by the server.")
                return 2

            context = ssl.create_default_context()
            code, response = client.starttls(context=context)
            print(f"\nSTARTTLS status: {code}")
            print(_decode_smtp_response(response))

            code, response = client.ehlo()
            print(f"\nEHLO after STARTTLS status: {code}")
            print("EHLO after STARTTLS response:")
            print(_decode_smtp_response(response))
            _print_capabilities(client, "Capabilities after STARTTLS")
            return 0
    except (OSError, smtplib.SMTPException, socket.timeout) as exc:
        print(f"SMTP probe failed: {exc}", file=sys.stderr)
        return 1


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inspect SMTP EHLO capabilities without authenticating."
    )
    parser.add_argument("--host", required=True, help="SMTP server hostname")
    parser.add_argument("--port", type=int, required=True, help="SMTP server port")
    parser.add_argument(
        "--timeout",
        type=float,
        default=10.0,
        help="Connection timeout in seconds, default: 10",
    )
    parser.add_argument(
        "--local-hostname",
        help="Optional hostname to use in EHLO/HELO",
    )
    parser.add_argument(
        "--ssl",
        action="store_true",
        help="Use implicit TLS via SMTP_SSL",
    )
    parser.add_argument(
        "--starttls",
        action="store_true",
        help="Attempt STARTTLS and print capabilities after TLS negotiation",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable smtplib protocol debug output",
    )
    return parser.parse_args(argv)


if __name__ == "__main__":
    raise SystemExit(probe(parse_args(sys.argv[1:])))
