from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak,
    Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import CondPageBreak
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable
from reportlab.lib.colors import HexColor
import os

BLUE       = HexColor("#5AA9FF")
BLUE_DARK  = HexColor("#2563EB")
BLUE_DIM   = HexColor("#1E3A5F")
BG_DARK    = HexColor("#0D1117")
BG_CARD    = HexColor("#161B22")
BG_MID     = HexColor("#1C2433")
TEXT_WHITE = HexColor("#E6EDF3")
TEXT_MUTED = HexColor("#8B949E")
TEXT_BODY  = HexColor("#C9D1D9")
RED        = HexColor("#F85149")
YELLOW     = HexColor("#D29922")
GREEN      = HexColor("#3FB950")
BORDER     = HexColor("#30363D")

W, H = letter

# ── Helper to keep a subsection together ─────────────────────
def kt(*items):
    """Flatten and wrap in KeepTogether"""
    flat = []
    for item in items:
        if isinstance(item, list):
            flat.extend(item)
        else:
            flat.append(item)
    return KeepTogether(flat)

class HRule(Flowable):
    def __init__(self, color=BORDER, thickness=0.5, width=None):
        Flowable.__init__(self)
        self.color = color
        self.thickness = thickness
        self._width = width
        self.height = thickness
    def wrap(self, aW, aH):
        self.width = self._width or aW
        return self.width, self.height
    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

def cover_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG_DARK)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    canvas.setFillColor(BLUE_DARK)
    canvas.rect(0, H - 6, W, 6, fill=1, stroke=0)
    canvas.setFillColor(HexColor("#0D2040"))
    canvas.circle(W + 80, H/2 + 100, 320, fill=1, stroke=0)
    canvas.setFillColor(TEXT_WHITE)
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawString(inch, H - 52, "MailFlow")
    canvas.setFillColor(BLUE)
    canvas.drawString(inch + 72, H - 52, "Sentinel")
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(inch, H - 62, W - inch, H - 62)
    canvas.setFillColor(TEXT_WHITE)
    canvas.setFont("Helvetica-Bold", 32)
    canvas.drawString(inch, H/2 + 80, "M365 Email Security")
    canvas.setFillColor(BLUE)
    canvas.setFont("Helvetica-Bold", 32)
    canvas.drawString(inch, H/2 + 40, "Audit Checklist")
    canvas.setFillColor(TEXT_MUTED)
    canvas.setFont("Helvetica", 13)
    canvas.drawString(inch, H/2 + 10, "A practitioner's guide for IT admins and MSPs")
    canvas.setStrokeColor(BLUE_DARK)
    canvas.setLineWidth(1.5)
    canvas.line(inch, H/2 - 8, inch + 180, H/2 - 8)
    tags = ["SPF", "DKIM", "DMARC", "MTA-STS", "Direct Send", "Third-Party Senders", "Incident Response"]
    x = inch
    y = H/2 - 50
    canvas.setFont("Helvetica-Bold", 8)
    for tag in tags:
        tw = canvas.stringWidth(tag, "Helvetica-Bold", 8) + 14
        canvas.setFillColor(BLUE_DIM)
        canvas.roundRect(x, y, tw, 16, 3, fill=1, stroke=0)
        canvas.setFillColor(BLUE)
        canvas.drawString(x + 7, y + 4, tag)
        x += tw + 6
        if x > W - inch - 60:
            x = inch
            y -= 24
    canvas.setFillColor(BG_CARD)
    canvas.rect(0, 0, W, 80, fill=1, stroke=0)
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(0, 80, W, 80)
    canvas.setFillColor(TEXT_MUTED)
    canvas.setFont("Helvetica", 9)
    canvas.drawString(inch, 50, "Published by MailFlow Sentinel  ·  mailflowsentinel.com")
    canvas.drawString(inch, 34, "Free to use. Share with your team.")
    canvas.setFillColor(BLUE)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawRightString(W - inch, 50, "mailflowsentinel.com")
    canvas.restoreState()

def content_page(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG_DARK)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # Header bar: 42pt tall — topMargin is set to match
    canvas.setFillColor(BG_CARD)
    canvas.rect(0, H - 42, W, 42, fill=1, stroke=0)
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(0, H - 42, W, H - 42)
    canvas.setFillColor(TEXT_WHITE)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawString(inch, H - 26, "MailFlow")
    canvas.setFillColor(BLUE)
    canvas.drawString(inch + 46, H - 26, "Sentinel")
    canvas.setFillColor(TEXT_MUTED)
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(W - inch, H - 26, "M365 Email Security Audit Checklist")
    # Footer bar: 34pt tall — bottomMargin is set to match
    canvas.setFillColor(BG_CARD)
    canvas.rect(0, 0, W, 34, fill=1, stroke=0)
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(0, 34, W, 34)
    canvas.setFillColor(TEXT_MUTED)
    canvas.setFont("Helvetica", 8)
    canvas.drawCentredString(W/2, 13, f"— {doc.page} —")
    canvas.drawString(inch, 13, "mailflowsentinel.com")
    canvas.restoreState()

def make_styles():
    styles = {}
    styles['h1'] = ParagraphStyle('h1', fontSize=20, fontName='Helvetica-Bold',
        textColor=TEXT_WHITE, spaceAfter=6, spaceBefore=10, leading=26)
    styles['h2'] = ParagraphStyle('h2', fontSize=14, fontName='Helvetica-Bold',
        textColor=TEXT_WHITE, spaceAfter=4, spaceBefore=10, leading=20)
    styles['h3'] = ParagraphStyle('h3', fontSize=10, fontName='Helvetica-Bold',
        textColor=BLUE, spaceAfter=3, spaceBefore=6, leading=15)
    styles['body'] = ParagraphStyle('body', fontSize=9, fontName='Helvetica',
        textColor=TEXT_BODY, spaceAfter=5, leading=14)
    styles['muted'] = ParagraphStyle('muted', fontSize=8, fontName='Helvetica',
        textColor=TEXT_MUTED, spaceAfter=3, leading=12)
    styles['code'] = ParagraphStyle('code', fontSize=7.5, fontName='Courier',
        textColor=HexColor("#79C0FF"), spaceAfter=0, leading=12,
        leftIndent=0, rightIndent=0)
    styles['checklist'] = ParagraphStyle('checklist', fontSize=9, fontName='Helvetica',
        textColor=TEXT_BODY, spaceAfter=3, leading=13, leftIndent=18)
    styles['mistake_label'] = ParagraphStyle('ml', fontSize=8.5, fontName='Helvetica-Bold',
        textColor=RED, spaceAfter=1, leading=13)
    styles['mistake'] = ParagraphStyle('m', fontSize=8.5, fontName='Helvetica',
        textColor=HexColor("#FFA657"), spaceAfter=4, leading=13, leftIndent=14)
    styles['section_intro'] = ParagraphStyle('si', fontSize=9, fontName='Helvetica-Oblique',
        textColor=TEXT_MUTED, spaceAfter=8, leading=14)
    return styles

def code_block(text, styles):
    """Each code line gets its own row so the table splits cleanly at page boundaries."""
    lines = text.strip().split('\n')
    rows = [[Paragraph(
        (line or ' ').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'),
        styles['code']
    )] for line in lines]
    t = Table(rows, colWidths=[W - 2*inch])
    # Build per-row padding: 8pt top on first row, 8pt bottom on last, 1pt between rows
    ts = [
        ('BACKGROUND',    (0, 0),  (-1, -1), BG_CARD),
        ('LEFTPADDING',   (0, 0),  (-1, -1), 10),
        ('RIGHTPADDING',  (0, 0),  (-1, -1), 10),
        ('TOPPADDING',    (0, 0),  (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0),  (-1, -1), 1),
        ('TOPPADDING',    (0, 0),  (0,  0),  8),   # first row
        ('BOTTOMPADDING', (0, -1), (-1, -1), 8),   # last row
        ('LINEABOVE',     (0, 0),  (-1,  0), 1, BLUE_DIM),
    ]
    t.setStyle(TableStyle(ts))
    return t

def callout(text, bg=BLUE_DIM, border=BLUE_DARK, text_color=TEXT_BODY):
    s = ParagraphStyle('cb', fontSize=9, fontName='Helvetica', textColor=text_color, leading=14)
    t = Table([[Paragraph(text, s)]], colWidths=[W - 2*inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 14),
        ('RIGHTPADDING', (0,0), (-1,-1), 14),
        ('LINEABOVE', (0,0), (-1,0), 2, border),
    ]))
    return t

def warn_callout(text):
    return callout(f'⚠  {text}', bg=HexColor("#2D2000"), border=YELLOW, text_color=HexColor("#E3B341"))

def checklist_item(text, styles):
    return Paragraph(f'<font color="#5AA9FF">☐</font>&nbsp;&nbsp;{text}', styles['checklist'])

def bullet(text, styles, color="#5AA9FF", icon="▸"):
    return Paragraph(f'<font color="{color}">{icon}</font>&nbsp;&nbsp;{text}',
        ParagraphStyle('bl', fontSize=9, fontName='Helvetica', textColor=TEXT_BODY, leading=14, leftIndent=12, spaceAfter=3))

def numbered(i, text, styles):
    return Paragraph(f'<font color="#5AA9FF"><b>{i}.</b></font>&nbsp;&nbsp;{text}',
        ParagraphStyle('nb', fontSize=9, fontName='Helvetica', textColor=TEXT_BODY, leading=14, leftIndent=12, spaceAfter=4))

def section_header(number, title, subtitle, styles):
    """Returns a single KeepTogether so the number/title/rule never orphan."""
    num_p = Paragraph(f'<font color="#5AA9FF"><b>{number}</b></font>',
        ParagraphStyle('n', fontSize=26, fontName='Helvetica-Bold', textColor=BLUE, leading=32))
    title_p = Paragraph(f'<b>{title}</b>',
        ParagraphStyle('t', fontSize=16, fontName='Helvetica-Bold', textColor=TEXT_WHITE, leading=22))
    sub_p = Paragraph(subtitle,
        ParagraphStyle('s', fontSize=8.5, fontName='Helvetica-Oblique', textColor=TEXT_MUTED, leading=13))
    t = Table([[num_p, [title_p, sub_p]]], colWidths=[44, W - 2*inch - 44])
    t.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    return KeepTogether([Spacer(1, 6), t, Spacer(1, 3), HRule(BLUE_DARK, 1.5), Spacer(1, 8)])

def mistakes(items, styles):
    block = [Paragraph("<b>Common Mistakes</b>", styles['h3'])]
    for label, text in items:
        block.append(Paragraph(f'❌&nbsp;&nbsp;<b>{label}</b>', styles['mistake_label']))
        block.append(Paragraph(text, styles['mistake']))
    block.append(Spacer(1, 4))
    return KeepTogether(block)

def checklist_block(title, items, styles):
    block = [Paragraph(f"<b>{title}</b>", styles['h3'])]
    for item in items:
        block.append(checklist_item(item, styles))
    block.append(Spacer(1, 4))
    return KeepTogether(block)

def build():
    out = os.path.expanduser("~/Documents/M365_Email_Security_Audit_Checklist.pdf")

    doc = SimpleDocTemplate(out, pagesize=letter,
        leftMargin=inch, rightMargin=inch,
        # topMargin matches the 42pt header bar drawn by content_page
        # bottomMargin matches the 34pt footer bar drawn by content_page
        topMargin=0.65*inch, bottomMargin=0.55*inch)

    styles = make_styles()
    story = []

    # ── COVER ─────────────────────────────────────────────────
    story.append(Spacer(1, H - 2*inch))
    story.append(PageBreak())

    # ── TOC ───────────────────────────────────────────────────
    story.append(Paragraph("Table of Contents", styles['h1']))
    story.append(HRule(BLUE_DARK, 1.5))
    story.append(Spacer(1, 10))
    for title, desc in [
        ("Introduction", "Who this is for, what you need, how to use this guide"),
        ("Section 1 — SPF", "Sender Policy Framework: authorization and enforcement"),
        ("Section 2 — DKIM", "DomainKeys Identified Mail: cryptographic signing"),
        ("Section 3 — DMARC", "Enforcement, reporting, and the path to reject"),
        ("Section 4 — Transport Security", "MTA-STS and TLS-RPT: securing mail in transit"),
        ("Section 5 — Direct Send Exposure", "M365 bypass testing and connector enforcement"),
        ("Section 6 — Third-Party Senders", "Identifying and authenticating every sender"),
        ("Section 7 — Incident Response", "Quick reference for common email security incidents"),
        ("Appendix A — Glossary", "Definitions for email security terminology"),
        ("Appendix B — MailFlow Sentinel", "How to use the tool in your audit workflow"),
        ("Appendix C — Next Steps by Score", "Recommended actions by posture level"),
    ]:
        row = Table([[
            Paragraph(f'<b>{title}</b>', ParagraphStyle('tt', fontSize=10, fontName='Helvetica-Bold', textColor=TEXT_WHITE, leading=14)),
            Paragraph(desc, ParagraphStyle('td', fontSize=9, fontName='Helvetica', textColor=TEXT_MUTED, leading=13))
        ]], colWidths=[2.4*inch, W - 2*inch - 2.4*inch])
        row.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LINEBELOW', (0,0), (-1,-1), 0.5, BORDER),
        ]))
        story.append(row)
    story.append(PageBreak())

    # ── INTRODUCTION ──────────────────────────────────────────
    story.append(section_header("", "Introduction", "Who this is for, how to use this guide, and what you will need.", styles))

    story.append(kt(
        Paragraph("<b>Who This Is For</b>", styles['h3']),
        Paragraph("This guide is written for two audiences. The first is the in-house IT administrator responsible for email "
            "at a small or mid-sized organization. You manage Microsoft 365, you handle DNS, and email security likely "
            "sits on your plate alongside a dozen other priorities. This guide gives you a structured process to assess "
            "your current posture, identify gaps, and fix them — without requiring a background in email security.", styles['body']),
        Paragraph("The second is the MSP technician managing email security across multiple client environments. You need a "
            "repeatable process you can run consistently, document clearly, and present to clients in a way that demonstrates "
            "value. This guide is built to be that process.", styles['body']),
        Paragraph("No prior email security expertise is required. Every concept is explained before it is tested. "
            "Every checklist item tells you what to look for and what good looks like.", styles['body']),
        Spacer(1, 6)
    ))

    story.append(kt(
        Paragraph("<b>What This Guide Covers</b>", styles['h3']),
        *[bullet(item, styles) for item in [
            "<b>SPF</b> — who is authorized to send email as your domain",
            "<b>DKIM</b> — cryptographic proof that your email is authentic and unmodified",
            "<b>DMARC</b> — enforcement and reporting for SPF and DKIM failures",
            "<b>Transport Security</b> — ensuring mail in transit is encrypted",
            "<b>Direct Send Exposure</b> — whether your M365 tenant can be reached bypassing your gateway",
            "<b>Third-Party Senders</b> — every service sending email as your domain",
            "<b>Incident Response</b> — what to do when something goes wrong",
        ]],
        Spacer(1, 6)
    ))

    story.append(kt(
        Paragraph("<b>What You Will Need</b>", styles['h3']),
        Paragraph("<b>Access</b>", ParagraphStyle('al', fontSize=8.5, fontName='Helvetica-Bold', textColor=TEXT_MUTED, leading=13, spaceAfter=2)),
        *[bullet(item, styles, icon="·") for item in [
            "DNS management access for the domain being audited",
            "Microsoft 365 admin access — Exchange admin and Security admin roles minimum",
            "Access to the Microsoft 365 Defender portal at security.microsoft.com",
        ]],
        Spacer(1, 3),
        Paragraph("<b>Tools</b>", ParagraphStyle('al', fontSize=8.5, fontName='Helvetica-Bold', textColor=TEXT_MUTED, leading=13, spaceAfter=2)),
        *[bullet(item, styles, icon="·") for item in [
            "MailFlow Sentinel at mailflowsentinel.com — free, no account required",
            "nslookup or dig for DNS verification (built into Windows, macOS, and Linux)",
            "Exchange Online PowerShell for DKIM management (optional)",
        ]],
        Spacer(1, 6)
    ))

    story.append(kt(
        callout("A Word on Urgency: Email remains the primary attack vector for phishing, business email compromise, and "
            "ransomware delivery. Domain spoofing is one of the most damaging and preventable threats an organization faces. "
            "The controls in this guide are foundational hygiene that every domain should have in place."),
        Spacer(1, 4)
    ))
    story.append(PageBreak())

    # ── SECTION 1: SPF ────────────────────────────────────────
    story.append(section_header("01", "SPF", "Sender Policy Framework: defines which mail servers are authorized to send on behalf of your domain.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("SPF is a DNS record that tells receiving mail servers which IP addresses and services are allowed to send "
            "email claiming to be from your domain. SPF alone does not block spoofed email — that is DMARC's job — but "
            "without a valid SPF record, DMARC cannot enforce anything.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "An SPF record exists in DNS (TXT record at the root domain)",
        "The record starts with v=spf1",
        "Microsoft 365 is included — include:spf.protection.outlook.com",
        "All third-party senders are authorized (marketing tools, ticketing systems, etc.)",
        "DNS lookup count is 10 or fewer",
        "The record ends with -all (hardfail) not ~all (softfail)",
        "There is only one SPF record — duplicate records cause failures",
    ], styles))

    story.append(kt(
        Paragraph("<b>What Good Looks Like</b>", styles['h3']),
        Paragraph("M365 only:", styles['muted']),
        code_block("v=spf1 include:spf.protection.outlook.com -all", styles),
        Spacer(1, 3),
        Paragraph("With additional senders:", styles['muted']),
        code_block("v=spf1 include:spf.protection.outlook.com include:sendgrid.net -all", styles),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>The DNS Lookup Limit</b>", styles['h3']),
        Paragraph("SPF has a hard limit of 10 DNS lookups. Every include:, a:, mx:, and redirect= counts toward this limit. "
            "Exceeding 10 causes a PermError — the record is treated as invalid and SPF fails entirely. "
            "Use MailFlow Sentinel to get an exact recursive lookup count.", styles['body']),
        Spacer(1, 4),
        warn_callout("~all (softfail) provides almost no real protection. Most major mail providers treat softfail the same as pass. "
            "Move to -all once all legitimate senders are confirmed in your SPF record."),
        Spacer(1, 4)
    ))

    story.append(CondPageBreak(2*inch))
    story.append(mistakes([
        ("Multiple SPF records", "Only one TXT record starting with v=spf1 is allowed. A second one causes an immediate PermError."),
        ("Forgetting third-party senders", "Marketing platforms, CRMs, helpdesk tools, and HR systems often send as your domain. If not in SPF, their mail fails authentication."),
        ("Using ~all and calling it done", "Softfail is a starting point, not a finish line."),
        ("Exceeding 10 lookups", "Adding includes over time without tracking the count silently breaks SPF."),
    ], styles))
    story.append(PageBreak())

    # ── SECTION 2: DKIM ───────────────────────────────────────
    story.append(section_header("02", "DKIM", "DomainKeys Identified Mail: cryptographic proof that email is authentic and unmodified.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("DKIM adds a cryptographic signature to every outbound email. The receiving server verifies that signature "
            "against a public key published in DNS. For M365, DKIM uses two selectors — selector1 and selector2 — "
            "that Microsoft manages and rotates automatically when enabled.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "DKIM is enabled in the Microsoft 365 Defender portal",
        "selector1 and selector2 CNAME records are published in DNS",
        "Both selectors resolve to valid public keys",
        "Key strength is 2048-bit (not 1024-bit)",
        "DKIM signing is verified — outbound mail carries a valid d= signature",
        "Any third-party senders have their own DKIM signing configured",
    ], styles))

    story.append(kt(
        Paragraph("<b>What Good Looks Like</b>", styles['h3']),
        code_block(
            "selector1._domainkey.yourdomain.com\n"
            "  -> selector1-yourdomain-com._domainkey.yourtenant.onmicrosoft.com\n\n"
            "selector2._domainkey.yourdomain.com\n"
            "  -> selector2-yourdomain-com._domainkey.yourtenant.onmicrosoft.com", styles),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>1024-bit vs 2048-bit Keys</b>", styles['h3']),
        Paragraph("Older M365 tenants may have DKIM keys at 1024-bit. These are considered weak and should be rotated. "
            "MailFlow Sentinel flags 1024-bit keys and deducts points from your score.", styles['body']),
        Paragraph("Rotate via Exchange Online PowerShell:", styles['muted']),
        code_block("Rotate-DkimSigningConfig -KeySize 2048 -Identity yourdomain.com", styles),
        Spacer(1, 4)
    ))

    story.append(CondPageBreak(2*inch))
    story.append(mistakes([
        ("DKIM never enabled in M365", "The CNAME records in DNS mean nothing if DKIM signing is not turned on in the portal."),
        ("DNS records published but signing not enabled", "Records exist but outbound mail does not carry a DKIM signature."),
        ("Assuming M365 handles everything", "M365 handles its own outbound mail. Third-party senders need their own DKIM setup."),
        ("Never checking key strength", "1024-bit keys can sit unnoticed for years. Add a key strength check to your annual review."),
    ], styles))
    story.append(PageBreak())

    # ── SECTION 3: DMARC ──────────────────────────────────────
    story.append(section_header("03", "DMARC", "Enforcement and reporting: tells receivers what to do when SPF or DKIM checks fail.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("DMARC is the enforcement layer tying SPF and DKIM together. Without DMARC, a receiving server detecting "
            "an SPF or DKIM failure has no instruction on what to do. DMARC gives you control over that decision and "
            "adds reporting so you can see authentication results across all mail sent using your domain. DMARC also "
            "introduces alignment — the From: address must match the domain that passed SPF or DKIM.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "A DMARC record exists in DNS (TXT record at _dmarc.yourdomain.com)",
        "The record starts with v=DMARC1",
        "A policy is set — p=none, p=quarantine, or p=reject",
        "Policy is at quarantine or reject (not monitor-only p=none)",
        "pct= is set to 100",
        "rua= is configured pointing to an address you actually monitor",
        "You are reviewing DMARC reports regularly",
        "Subdomain policy (sp=) is explicitly set if needed",
    ], styles))

    story.append(kt(
        Paragraph("<b>What Good Looks Like</b>", styles['h3']),
        code_block("v=DMARC1; p=reject; pct=100; rua=mailto:dmarc@yourdomain.com; fo=1", styles),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>The Three Policy Levels</b>", styles['h3']),
        Table([
            ["Policy", "Action", "Protection"],
            ["p=none", "Deliver normally, report only", "None — monitor only"],
            ["p=quarantine", "Send to spam/junk folder", "Partial"],
            ["p=reject", "Reject at gateway", "Full — the goal"],
        ], colWidths=[1.1*inch, 2.9*inch, 1.8*inch],
        style=TableStyle([
            ('BACKGROUND', (0,0), (-1,0), BLUE_DARK),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [HexColor("#1A2030"), BG_CARD, HexColor("#0D2010")]),
            ('TEXTCOLOR', (0,0), (-1,0), TEXT_WHITE),
            ('TEXTCOLOR', (0,1), (-1,-1), TEXT_BODY),
            ('TEXTCOLOR', (0,3), (0,3), GREEN),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME', (0,1), (0,-1), 'Courier'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, BORDER),
        ])),
        Spacer(1, 6),
        warn_callout("p=none provides zero spoofing protection — it only watches. Many organizations get stuck here indefinitely. Set a deadline to move to enforcement."),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>The Path to Reject</b>", styles['h3']),
        *[numbered(i, step, styles) for i, step in enumerate([
            "Deploy p=none with rua= reporting — run for 2–4 weeks",
            "Analyze reports — identify all legitimate sending sources",
            "Fix authentication gaps — get all senders into SPF and DKIM",
            "Move to p=quarantine at pct=10, gradually increase to 100",
            "Move to p=reject once quarantine at 100% shows no legitimate failures",
        ], 1)],
        Spacer(1, 8),
        Paragraph("<b>Common Mistakes</b>", styles['h3']),
        Paragraph('❌&nbsp;&nbsp;<b>Staying at p=none indefinitely</b>', styles['mistake_label']),
        Paragraph("The most common DMARC failure. p=none is a starting point, not a destination.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>No rua= reporting configured</b>', styles['mistake_label']),
        Paragraph("Enforcement without visibility means you won't know when something breaks.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>pct= below 100 on a reject policy</b>', styles['mistake_label']),
        Paragraph("Partial enforcement gives a false sense of security.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>rua= points only to a vendor</b>', styles['mistake_label']),
        Paragraph("Keep one address you control in case the vendor relationship changes.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Ignoring subdomain policy</b>', styles['mistake_label']),
        Paragraph("If sp= is not set, subdomains inherit the root domain policy. Verify this is intentional.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Not setting fo= for failure reporting</b>', styles['mistake_label']),
        Paragraph("fo=1 instructs receivers to send reports on any authentication failure, not just DMARC failures. Without it you get limited diagnostic data. Add fo=1 to your DMARC record.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Overlooking the ruf= forensic reporting tag</b>', styles['mistake_label']),
        Paragraph("ruf= sends message-level forensic reports on individual failures. Many organizations skip it but it is valuable for diagnosing specific authentication issues. Note that many receivers no longer send forensic reports due to privacy concerns, so rua= remains the priority.", styles['mistake']),
        Spacer(1, 4)
    ))
    story.append(PageBreak())

    # ── SECTION 4: TRANSPORT SECURITY ─────────────────────────
    story.append(section_header("04", "Transport Security", "MTA-STS and TLS-RPT: ensuring mail in transit is encrypted and failures are reported.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("SPF, DKIM, and DMARC protect against spoofing. MTA-STS and TLS-RPT protect the connection itself — "
            "ensuring mail in transit is encrypted and you are notified when encryption fails. Without MTA-STS, "
            "a man-in-the-middle attacker could downgrade an SMTP connection to plaintext.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "An MTA-STS DNS record exists at _mta-sts.yourdomain.com",
        "The MTA-STS policy file is published at https://mta-sts.yourdomain.com/.well-known/mta-sts.txt",
        "Policy mode is set to enforce (not testing or none)",
        "The mx: entries in the policy file match your actual MX records",
        "A TLS-RPT record exists at _smtp._tls.yourdomain.com",
        "TLS-RPT rua= points to an address you monitor",
    ], styles))

    story.append(kt(
        Paragraph("<b>What Good Looks Like</b>", styles['h3']),
        Paragraph("MTA-STS DNS record:", styles['muted']),
        code_block('_mta-sts.yourdomain.com TXT "v=STSv1; id=20240101000000;"', styles),
        Spacer(1, 3),
        Paragraph("MTA-STS policy file:", styles['muted']),
        code_block("version: STSv1\nmode: enforce\nmx: *.mail.protection.outlook.com\nmax_age: 604800", styles),
        Spacer(1, 3),
        Paragraph("TLS-RPT record:", styles['muted']),
        code_block('_smtp._tls.yourdomain.com TXT "v=TLSRPTv1; rua=mailto:tlsrpt@yourdomain.com"', styles),
        Spacer(1, 4)
    ))

    story.append(CondPageBreak(2*inch))
    story.append(mistakes([
        ("Staying in testing mode", "Testing mode provides no protection. Set a deadline to move to enforce."),
        ("MX entries don't match actual MX records", "If the policy file lists the wrong servers, legitimate mail will fail delivery once enforced."),
        ("Forgetting to update the id= tag", "When you change the policy file, the id= value in DNS must change too. Servers cache the policy based on id=."),
        ("Policy file not accessible over HTTPS", "A self-signed cert or HTTP-only URL causes servers to ignore the policy."),
        ("No TLS-RPT configured", "Enforcing MTA-STS without TLS-RPT leaves you with no visibility into delivery failures."),
    ], styles))
    story.append(PageBreak())

    # ── SECTION 5: DIRECT SEND ────────────────────────────────
    story.append(section_header("05", "Direct Send Exposure", "Tests whether M365 can be reached directly, bypassing your gateway and security controls.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("Microsoft 365 accepts inbound mail at a tenant-specific MX endpoint. By default, any external server "
            "can attempt to deliver mail directly to this endpoint, bypassing any gateway sitting in front of it. "
            "If Direct Send is not locked down, an attacker can skip the gateway and deliver mail straight to user "
            "mailboxes — bypassing filtering, sandboxing, and impersonation detection.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "An Enhanced Filtering connector is configured in M365",
        "Connector enforcement is active — external senders receive a 550 5.7.68 rejection",
        "Mail flow is verified through the gateway only",
        "The MX record points to the gateway, not directly to M365",
        "No mail flow rules exist that inadvertently bypass connector restrictions",
        "Direct Send exposure is tested from an external source",
    ], styles))

    story.append(kt(
        Paragraph("<b>What Good Looks Like</b>", styles['h3']),
        Paragraph("When connector enforcement is active, direct connection attempts receive:", styles['muted']),
        code_block("550 5.7.68 TenantInboundAttribution;\nThis mail is not authorized to bypass connector restrictions.", styles),
        Spacer(1, 4),
        callout("This is not a theoretical attack. Direct Send bypass is actively exploited in BEC and phishing campaigns "
            "because it bypasses security tools. MailFlow Sentinel tests this automatically and reports Protected or Not Protected.",
            bg=HexColor("#1A0020"), border=HexColor("#6E40C9")),
        Spacer(1, 4)
    ))

    story.append(CondPageBreak(2*inch))
    story.append(mistakes([
        ("Gateway deployed but connector enforcement never configured", "The gateway filters inbound mail but M365 still accepts direct connections."),
        ("MX record points directly to M365", "No gateway means nothing to enforce. All mail lands directly in M365."),
        ("Connector configured but not tested", "Configuration mistakes are common. Always verify with an external SMTP test."),
        ("Assuming Defender for Office 365 closes the gap", "Defender for Office 365 is inside M365, not a gateway. It does not prevent Direct Send bypass."),
    ], styles))
    story.append(PageBreak())

    # ── SECTION 6: THIRD-PARTY SENDERS ───────────────────────
    story.append(section_header("06", "Third-Party Senders", "Identifying and authenticating every service sending email on behalf of your domain.", styles))

    story.append(kt(
        Paragraph("<b>What it does</b>", styles['h3']),
        Paragraph("Most organizations send email from more than just Microsoft 365. Marketing platforms, CRM systems, "
            "helpdesk tools, HR software, and dozens of SaaS products send email using your domain — often without "
            "IT's knowledge. Every service sending as your domain needs to be identified, authorized in SPF, and "
            "configured for DKIM signing.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(checklist_block("The Checklist", [
        "All third-party services sending email as your domain are identified",
        "Each sender is authorized in your SPF record via include: or ip4:",
        "Each sender has DKIM signing configured with DNS records published",
        "No unauthorized or unknown services are sending as your domain",
        "SPF lookup count remains at or below 10 after all senders are added",
        "Departed vendors have been removed from SPF and DKIM",
    ], styles))

    story.append(kt(
        Paragraph("<b>Sender Inventory Template</b>", styles['h3']),
        Table([
            ["Service", "SPF Auth", "DKIM Config", "Notes"],
            ["Microsoft 365", "✅", "✅", "Primary mail flow"],
            ["[Marketing platform]", "☐", "☐", ""],
            ["[CRM / Sales tool]", "☐", "☐", ""],
            ["[Support platform]", "☐", "☐", ""],
            ["[Transactional email]", "☐", "☐", ""],
            ["[HR / Payroll]", "☐", "☐", ""],
        ], colWidths=[2.0*inch, 0.8*inch, 0.95*inch, 2.05*inch],
        style=TableStyle([
            ('BACKGROUND', (0,0), (-1,0), BLUE_DARK),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [HexColor("#0D2010"), BG_CARD]),
            ('TEXTCOLOR', (0,0), (-1,0), TEXT_WHITE),
            ('TEXTCOLOR', (0,1), (-1,-1), TEXT_BODY),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, BORDER),
        ])),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>How to Find Your Third-Party Senders</b>", styles['h3']),
        Paragraph("<b>SPF includes</b> — Every include: directive points to a sending service.", styles['body']),
        Paragraph("<b>DKIM selector CNAMEs</b> — Services like SendGrid and Mailchimp publish DKIM keys under predictable selector names.", styles['body']),
        Paragraph("<b>Tracking subdomain CNAMEs</b> — Marketing platforms require subdomains for click tracking that reveal platform relationships.", styles['body']),
        Spacer(1, 4),
        callout("MailFlow Sentinel detects third-party senders automatically using all three signals."),
        Spacer(1, 8),
        Paragraph("<b>Common Mistakes</b>", styles['h3']),
        Paragraph('❌&nbsp;&nbsp;<b>Only auditing what IT knows about</b>', styles['mistake_label']),
        Paragraph("The most impactful senders are often set up by marketing or sales without IT involvement.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Adding senders to SPF but skipping DKIM</b>', styles['mistake_label']),
        Paragraph("SPF alone is not enough for strong DMARC alignment.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Never removing departed vendors</b>', styles['mistake_label']),
        Paragraph("Abandoned includes waste lookup budget and create security risk.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Not rechecking after adding senders</b>', styles['mistake_label']),
        Paragraph("Every new include: costs a lookup. Recheck your count every time.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Assuming all senders support DKIM</b>', styles['mistake_label']),
        Paragraph("Some older or budget SaaS platforms do not support DKIM signing. If a vendor cannot provide DKIM records, document it and factor it into your DMARC rollout timeline.", styles['mistake']),
        Paragraph('❌&nbsp;&nbsp;<b>Not testing after configuration</b>', styles['mistake_label']),
        Paragraph("Always send a test email through each third-party sender after configuring SPF and DKIM, and verify the Authentication-Results header shows both passing before closing the ticket.", styles['mistake']),
        Spacer(1, 4)
    ))
    story.append(PageBreak())

    # ── SECTION 7: INCIDENT RESPONSE ─────────────────────────
    story.append(section_header("07", "Incident Response", "Quick reference for the most common email security incidents.", styles))

    story.append(Paragraph("Use this section when something goes wrong. Each scenario includes immediate steps and how to verify resolution.",
        styles['section_intro']))

    for num, title, items in [
        ("Scenario 1", "Mail Stopped Flowing", [
            "Check M365 service health at admin.microsoft.com → Health → Service Health",
            "Verify MX records are resolving — run: nslookup -type=MX yourdomain.com",
            "Check if SPF record was recently modified — syntax errors break mail flow instantly",
            "Verify connector configuration in Exchange admin center is intact",
            "Check gateway status if using a third-party security gateway",
            "Review M365 message trace at admin.microsoft.com → Exchange → Mail flow",
            "Check if the domain is on any blacklists",
        ]),
        ("Scenario 2", "DMARC Failures Spiking", [
            "Pull the latest DMARC aggregate report and identify which IPs are failing",
            "Determine if failing IPs are legitimate senders or unknown sources",
            "If legitimate sender — check if their SPF include or DKIM records changed",
            "If unknown IPs — someone may be spoofing your domain",
            "Verify SPF record still includes all authorized senders",
            "Verify DKIM selectors are still resolving for all active senders",
        ]),
        ("Scenario 3", "Someone Spoofed Our Domain", [
            "Obtain a copy of the spoofed email including full headers",
            "Identify the actual sending IP from the Received: headers",
            "Check your DMARC policy — if at p=none, enforcement is the priority",
            "Verify SPF -all is set — ~all allows spoofed mail to be delivered",
            "Verify DKIM is enabled and signing outbound mail",
            "Report the sending IP to relevant abuse contacts",
        ]),
        ("Scenario 4", "Legitimate Mail Going to Spam", [
            "Check if your sending IP or domain is on any major blacklists",
            "Review recent sending volume — a sudden spike triggers spam filters",
            "Verify SPF, DKIM, and DMARC are all passing for the affected mail stream",
            "Check sending reputation at Google Postmaster Tools and Microsoft SNDS",
            "Verify unsubscribe links are present and functional for marketing mail",
            "Check complaint rates — high rates damage sender reputation quickly",
        ]),
    ]:
        story.append(kt(
            Paragraph(f'<font color="#5AA9FF"><b>{num}</b></font> — <b>{title}</b>',
                ParagraphStyle('sh', fontSize=10, fontName='Helvetica-Bold', textColor=TEXT_WHITE, leading=16, spaceAfter=4, spaceBefore=8)),
            *[checklist_item(item, styles) for item in items],
            Spacer(1, 4)
        ))

    # Scenario 5 + Quick Reference kept together so they never split across pages
    story.append(kt(
        Paragraph('<font color="#5AA9FF"><b>Scenario 5</b></font> — <b>DKIM Signature Failures</b>',
            ParagraphStyle('sh', fontSize=10, fontName='Helvetica-Bold', textColor=TEXT_WHITE, leading=16, spaceAfter=4, spaceBefore=8)),
        *[checklist_item(item, styles) for item in [
            "Identify which selector is failing — selector1, selector2, or third-party",
            "Verify CNAME is published: nslookup -type=CNAME selector1._domainkey.yourdomain.com",
            "Verify the CNAME resolves to a valid public key record",
            "Check if DKIM signing is still enabled in M365 or the relevant platform",
            "If third-party sender — verify their DKIM records have not changed",
            "Check if a DNS migration may have dropped CNAME records",
        ]],
        Spacer(1, 8),
        Paragraph("<b>Quick Reference — Useful Commands</b>", styles['h3']),
        code_block(
            "# Check MX records\n"
            "nslookup -type=MX yourdomain.com\n\n"
            "# Check SPF record\n"
            "nslookup -type=TXT yourdomain.com\n\n"
            "# Check DMARC record\n"
            "nslookup -type=TXT _dmarc.yourdomain.com\n\n"
            "# Check DKIM selector\n"
            "nslookup -type=CNAME selector1._domainkey.yourdomain.com\n\n"
            "# Rotate DKIM to 2048-bit (PowerShell)\n"
            "Rotate-DkimSigningConfig -KeySize 2048 -Identity yourdomain.com", styles),
        Spacer(1, 4)
    ))
    story.append(PageBreak())

    # ── APPENDIX A: GLOSSARY ──────────────────────────────────
    story.append(section_header("A", "Glossary", "Quick reference definitions for email security terminology.", styles))

    for term, definition in [
        ("Alignment", "DMARC requirement that the From: domain matches the domain that passed SPF or DKIM. Closes the spoofing gap SPF and DKIM alone leave open."),
        ("BIMI", "Brand Indicators for Message Identification. Allows verified brand logos in Gmail, Apple Mail, Yahoo Mail. Requires DMARC at quarantine or reject and a VMC."),
        ("CNAME Record", "DNS record pointing one hostname to another. Used in DKIM configuration for third-party senders."),
        ("DKIM", "DomainKeys Identified Mail. Adds a cryptographic signature to outbound email, verified against a public key in DNS."),
        ("DMARC", "Domain-based Message Authentication, Reporting and Conformance. Enforcement layer tying SPF and DKIM together."),
        ("DNS Lookup Limit", "SPF evaluation limited to 10 DNS lookups. Exceeding this causes a PermError — SPF fails entirely."),
        ("Enhanced Filtering", "M365 connector configuration restricting inbound mail to authorized sources. Prevents Direct Send bypass."),
        ("Hardfail (-all)", "SPF mechanism instructing receivers to reject mail from unauthorized senders. Strongest SPF enforcement."),
        ("MTA-STS", "Mail Transfer Agent Strict Transport Security. Requires inbound mail to be delivered over TLS to authorized servers."),
        ("MX Record", "DNS record specifying which mail servers accept email for a domain."),
        ("NDR", "Non-Delivery Report. Bounce message with error codes identifying the reason for delivery failure."),
        ("PermError", "Permanent SPF error from exceeding 10 lookups or a syntax error. Mail fails SPF evaluation entirely."),
        ("p=none", "DMARC policy taking no action on failing mail — monitor and report only. Zero spoofing protection."),
        ("p=quarantine", "DMARC policy sending failing mail to spam/junk. Partial enforcement."),
        ("p=reject", "DMARC policy rejecting failing mail entirely. Maximum protection — the goal."),
        ("Selector", "Label identifying a DKIM public key in DNS. M365 uses selector1 and selector2."),
        ("Softfail (~all)", "SPF mechanism marking unauthorized senders suspicious but allowing delivery. Weak in practice."),
        ("SPF", "Sender Policy Framework. DNS record defining which IPs and services are authorized to send for a domain."),
        ("TLS", "Transport Layer Security. Encryption protocol securing SMTP connections between mail servers."),
        ("VMC", "Verified Mark Certificate. Required for BIMI logo display in Gmail and Apple Mail."),
    ]:
        row = Table([[
            Paragraph(f'<b>{term}</b>', ParagraphStyle('gt', fontSize=8.5, fontName='Helvetica-Bold', textColor=BLUE, leading=13)),
            Paragraph(definition, ParagraphStyle('gd', fontSize=8.5, fontName='Helvetica', textColor=TEXT_BODY, leading=13))
        ]], colWidths=[1.4*inch, W - 2*inch - 1.4*inch])
        row.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LINEBELOW', (0,0), (-1,-1), 0.5, BORDER),
        ]))
        story.append(row)
    story.append(PageBreak())

    # ── APPENDIX B: MAILFLOW SENTINEL ─────────────────────────
    story.append(section_header("B", "MailFlow Sentinel", "How to use the free assessment tool to accelerate your audit workflow.", styles))

    story.append(kt(
        Paragraph("MailFlow Sentinel is a free DNS-based email security assessment tool at mailflowsentinel.com. "
            "It automates the most time-consuming parts of this audit — DNS lookups, SMTP probing, DKIM selector scanning, "
            "and ESP detection — presenting results with findings and recommended fixes. No account required.", styles['body']),
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>What it checks automatically</b>", styles['h3']),
        *[bullet(item, styles, color="#3FB950", icon="✓") for item in [
            "SPF record presence, syntax, exact recursive lookup count, and enforcement level",
            "DMARC policy, pct= value, rua= configuration, and vendor-only reporting detection",
            "DKIM selector scanning across 15 common selectors with key strength detection",
            "MTA-STS DNS record and policy file retrieval",
            "TLS-RPT record presence and BIMI record detection",
            "Direct Send exposure via live SMTP probe with full SMTP trace",
            "Open relay testing and catch-all detection",
            "Mail subdomain scanning",
            "Third-party ESP detection via SPF includes, DKIM CNAMEs, and tracking subdomains",
            "Mail server fingerprinting including TLS version and server banner",
        ]],
        Spacer(1, 4)
    ))

    story.append(kt(
        Paragraph("<b>How to use it in an audit</b>", styles['h3']),
        *[numbered(i, step, styles) for i, step in enumerate([
            "Enter the domain at mailflowsentinel.com and click Run Analysis",
            "Review the score and badge summary — immediate posture snapshot",
            "Work through Findings — each finding maps to a section in this guide",
            "Review Recommended Fixes — CRITICAL items appear in red, prioritize these",
            "Use Share Results to generate a permanent link for customer documentation",
            "Use Copy as Text or Download PDF to include results in your audit report",
        ], 1)],
        Spacer(1, 6),
        callout("A full scan typically completes in under 15 seconds. Visit mailflowsentinel.com to run your first assessment."),
        Spacer(1, 4)
    ))
    story.append(PageBreak())

    # ── APPENDIX C: NEXT STEPS BY SCORE ──────────────────────
    story.append(section_header("C", "Next Steps by Posture Score", "Recommended actions based on your MailFlow Sentinel assessment score.", styles))

    for score, label, color, bg, actions in [
        ("90–100", "Strong Posture", GREEN, HexColor("#0D2010"), [
            "Confirm pct=100 on DMARC reject",
            "Deploy MTA-STS in enforce mode if not already done",
            "Rotate DKIM keys to 2048-bit if any 1024-bit keys remain",
            "Conduct annual sender inventory review",
            "Evaluate BIMI if brand visibility is a priority",
        ]),
        ("75–89", "Good Baseline, Gaps Remain", BLUE, BLUE_DIM, [
            "Identify and fix the specific findings flagged in the audit",
            "Harden SPF from ~all to -all if not already done",
            "Move DMARC from quarantine to reject if at quarantine",
            "Configure MTA-STS if missing",
            "Audit third-party senders for missing DKIM",
        ]),
        ("50–74", "Partial Deployment", YELLOW, HexColor("#2D2000"), [
            "Prioritize DMARC enforcement — move from p=none to p=quarantine immediately",
            "Enable DKIM if not active",
            "Fix any SPF syntax errors or lookup count violations",
            "Verify Direct Send is locked down if using a gateway",
            "Set a 90-day timeline to reach p=reject",
        ]),
        ("Below 50", "High Risk", RED, HexColor("#200D0D"), [
            "Treat this as an urgent remediation project, not a routine audit",
            "Deploy SPF and DKIM immediately if missing",
            "Move DMARC to at minimum p=quarantine within 30 days",
            "Lock down Direct Send if a gateway is in use",
            "Brief stakeholders on the spoofing risk and establish a remediation timeline",
        ]),
    ]:
        header_t = Table([[
            Paragraph(f'<b>{score}</b>', ParagraphStyle('sc', fontSize=13, fontName='Helvetica-Bold', textColor=color, leading=18)),
            Paragraph(f'<b>{label}</b>', ParagraphStyle('sl', fontSize=11, fontName='Helvetica-Bold', textColor=TEXT_WHITE, leading=18))
        ]], colWidths=[1.0*inch, W - 2*inch - 1.0*inch])
        header_t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), bg),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 12),
            ('RIGHTPADDING', (0,0), (-1,-1), 12),
            ('LINEABOVE', (0,0), (-1,0), 2, color),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        # Each action item on its own row so padding and background apply correctly
        action_rows = [[bullet(a, styles)] for a in actions]
        action_t = Table(action_rows, colWidths=[W - 2*inch])
        action_t.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0),  (-1, -1), BG_CARD),
            ('TOPPADDING',    (0, 0),  (0,  0),  8),
            ('TOPPADDING',    (0, 1),  (-1, -1), 2),
            ('BOTTOMPADDING', (0, -1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0),  (-1, -2), 2),
            ('LEFTPADDING',   (0, 0),  (-1, -1), 12),
            ('RIGHTPADDING',  (0, 0),  (-1, -1), 12),
            ('LINEBELOW',     (0, -1), (-1, -1), 0.5, BORDER),
        ]))
        story.append(KeepTogether([header_t, action_t, Spacer(1, 10)]))

    # Footer — keep the closing block together so it doesn't orphan
    story.append(KeepTogether([
        Spacer(1, 20),
        HRule(BLUE_DARK, 1.5),
        Spacer(1, 12),
        Paragraph("Run your free assessment at <b>mailflowsentinel.com</b>",
            ParagraphStyle('fc', fontSize=11, fontName='Helvetica', textColor=TEXT_MUTED, leading=16, alignment=TA_CENTER)),
        Paragraph("Published by MailFlow Sentinel  ·  Free to share with your team",
            ParagraphStyle('fc2', fontSize=8.5, fontName='Helvetica', textColor=TEXT_MUTED, leading=13, alignment=TA_CENTER, spaceBefore=4)),
    ]))

    doc.build(story, onFirstPage=cover_page, onLaterPages=content_page)
    print(f"PDF built: {out}")

if __name__ == "__main__":
    build()
