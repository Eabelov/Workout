# Workout

## SMTP AUTH diagnostics

This repository contains a small SMTP diagnostic helper for investigating
`smtplib.SMTPNotSupportedError: SMTP AUTH extension not supported by server`.

The error means that the SMTP client tried to authenticate, but the server did
not advertise the `AUTH` capability in its EHLO response. For an endpoint such
as `smtpsrv.test.ru:54` with TLS and SSL disabled, verify whether the
integration really needs SMTP authentication:

- if "SUZ test" is only the sender identity, configure it as the `From` address
  or envelope sender and leave SMTP username/password empty;
- if SMTP authentication is required, the SMTP server must advertise `AUTH`
  for this port, often after enabling STARTTLS or enabling SMTP AUTH server-side;
- if both TLS and SSL are disabled, do not call `SMTP.login(...)` unless EHLO
  capabilities include `AUTH`.

Run the probe without sending credentials:

```bash
python3 scripts/smtp_probe.py --host smtpsrv.test.ru --port 54
```

See [docs/smtp-auth-debug.md](docs/smtp-auth-debug.md) for the detailed
debugging checklist.
