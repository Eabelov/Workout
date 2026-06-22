# Debugging `SMTPNotSupportedError: SMTP AUTH extension not supported by server`

## What the error means

Python raises this exception when `smtplib.SMTP.login(...)` is called but the
server did not advertise the `AUTH` SMTP extension after EHLO.

For the reported endpoint:

- host: `smtpsrv.test.ru`
- port: `54`
- TLS: disabled
- SSL: disabled
- identity: `SUZ test`

the most likely cause is that the integration is configured with SMTP
credentials, so the client tries to authenticate even though the server does
not support authentication on that port.

## Expected SMTP flow without TLS/SSL

When TLS and SSL are disabled, the client should use a plain SMTP connection:

```text
connect smtpsrv.test.ru:54
EHLO <client-hostname>
MAIL FROM:<sender@example.test>
RCPT TO:<recipient@example.test>
DATA
```

It should not run:

```text
AUTH ...
```

unless the EHLO response contains an `AUTH` capability.

## Client-side checks

1. Confirm whether the "SUZ test" value is meant to be:
   - a sender identity (`From`, `MAIL FROM`, application account), or
   - an SMTP login username.
2. If the server is intended to work without SMTP AUTH, clear SMTP username and
   password in the integration configuration.
3. Ensure the code calls `SMTP.login(...)` only when credentials are configured
   and the server advertises `AUTH`.
4. Keep `use_tls` and `use_ssl` disabled for this endpoint unless the server
   owners confirm otherwise.

## Server-side checks

Ask the SMTP server owners to confirm the EHLO capabilities on port `54`.
Specifically, verify whether the response includes a line like:

```text
250-AUTH PLAIN LOGIN
```

If SMTP authentication is required, the server must advertise `AUTH` on this
port. Some servers advertise `AUTH` only after STARTTLS, so enabling
authentication may also require enabling STARTTLS on the client and server.

## Probe command

Run the included probe from this repository:

```bash
python3 scripts/smtp_probe.py --host smtpsrv.test.ru --port 54
```

The probe prints:

- connection status;
- server greeting;
- EHLO response;
- parsed capabilities;
- whether `AUTH` is advertised.

It does not send credentials and does not call `SMTP.login(...)` when `AUTH` is
missing.

## Interpreting results

If the probe prints:

```text
AUTH advertised: no
```

then the integration must not attempt SMTP AUTH for this server/port. Remove
SMTP login credentials from the client configuration or ask the server owners
to enable SMTP AUTH.

If the probe prints:

```text
STARTTLS advertised: yes
```

you can run an additional STARTTLS probe:

```bash
python3 scripts/smtp_probe.py --host smtpsrv.test.ru --port 54 --starttls
```

If `AUTH advertised` changes from `no` before STARTTLS to `yes` after STARTTLS,
then the server expects STARTTLS before authentication. Enable STARTTLS in the
integration and keep SSL disabled.
