# M365 Email Security Audit Checklist

A practitioner's guide for IT admins and MSPs auditing Microsoft 365 email security posture. Available as a professionally formatted PDF built with Python and ReportLab.

## What's in the guide

- **SPF** — record structure, DNS lookup limits, enforcement levels
- **DKIM** — enabling M365 selectors, key strength, third-party signing
- **DMARC** — policy levels, the path to reject, reporting tags
- **Transport Security** — MTA-STS and TLS-RPT configuration
- **Direct Send Exposure** — testing M365 connector enforcement
- **Third-Party Senders** — identifying and authenticating every ESP
- **Incident Response** — quick reference for five common scenarios
- **Appendices** — glossary, MailFlow Sentinel workflow, next steps by score

## Files

| File | Description |
|---|---|
| `M365_Email_Security_Audit_Checklist.pdf` | Ready-to-use PDF (17 pages) |
| `build_pdf2.py` | Python script that generates the PDF |

## Rebuilding the PDF

Requires Python 3 and [ReportLab](https://www.reportlab.com/):

```bash
pip install reportlab
python3 build_pdf2.py
```

Output: `~/Documents/M365_Email_Security_Audit_Checklist.pdf`

## Related tools

This guide is designed to be used alongside **[MailFlow Sentinel](https://mailflowsentinel.com)** — a free DNS-based email security assessment tool that automates SPF/DKIM/DMARC analysis, Direct Send probing, and third-party sender detection. No account required.

---

Published by [Reboot Labs](https://ther3boot.com) · Free to use and share
