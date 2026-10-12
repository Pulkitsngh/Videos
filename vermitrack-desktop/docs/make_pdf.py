"""Builds the VermiTrack IT deployment & installation process-flow PDF."""
import sys
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Flowable, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

OUT, SHOTS, LOGO = sys.argv[1], sys.argv[2], sys.argv[3]
F = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DV", F + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", F + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("DVM", F + "DejaVuSansMono.ttf"))
pdfmetrics.registerFontFamily("DV", normal="DV", bold="DVB", italic="DV", boldItalic="DVB")

GREEN = colors.HexColor("#1f5a43"); DARK = colors.HexColor("#163c2e"); INK = colors.HexColor("#18332b")
MUTED = colors.HexColor("#5b6c63"); LINE = colors.HexColor("#d6dfd7"); PALE = colors.HexColor("#eef4ef")
CREAM = colors.HexColor("#fffefa"); AMBER = colors.HexColor("#fff3d6"); AMBER_INK = colors.HexColor("#7a5310")
LANE = {"IT": colors.HexColor("#2f5f8a"), "KP": GREEN, "USER": colors.HexColor("#8a5a2f")}
LANE_BG = {"IT": colors.HexColor("#e8f0f7"), "KP": colors.HexColor("#e7f1eb"), "USER": colors.HexColor("#f6ece2")}

ss = {
    "h1": ParagraphStyle("h1", fontName="DVB", fontSize=17, leading=22, textColor=DARK, spaceBefore=4, spaceAfter=8),
    "h2": ParagraphStyle("h2", fontName="DVB", fontSize=12.5, leading=16, textColor=GREEN, spaceBefore=10, spaceAfter=5),
    "body": ParagraphStyle("body", fontName="DV", fontSize=9.5, leading=14, textColor=INK, spaceAfter=5),
    "small": ParagraphStyle("small", fontName="DV", fontSize=8, leading=11, textColor=MUTED),
    "cell": ParagraphStyle("cell", fontName="DV", fontSize=8.6, leading=11.5, textColor=INK),
    "cellb": ParagraphStyle("cellb", fontName="DVB", fontSize=8.6, leading=11.5, textColor=INK),
    "head": ParagraphStyle("head", fontName="DVB", fontSize=8.6, leading=11.5, textColor=colors.white),
    "step": ParagraphStyle("step", fontName="DV", fontSize=9.3, leading=13.5, textColor=INK),
    "code": ParagraphStyle("code", fontName="DVM", fontSize=8.2, leading=11, textColor=INK, backColor=colors.HexColor("#f3f2ec"), borderPadding=5, spaceBefore=3, spaceAfter=7),
    "cap": ParagraphStyle("cap", fontName="DV", fontSize=7.8, leading=10, textColor=MUTED, alignment=TA_CENTER, spaceAfter=8),
}
P = lambda text, style="body": Paragraph(text, ss[style])
W = A4[0] - 36 * mm


def table(rows, widths, header=True, zebra=True):
    data = [[P(c, "head" if header and i == 0 else ("cellb" if j == 0 else "cell")) if isinstance(c, str) else c for j, c in enumerate(r)] for i, r in enumerate(rows)]
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [("GRID", (0, 0), (-1, -1), 0.5, LINE), ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5)]
    if header: style.append(("BACKGROUND", (0, 0), (-1, 0), GREEN))
    if zebra:
        for i in range(1 if header else 0, len(rows)):
            if i % 2 == 0: style.append(("BACKGROUND", (0, i), (-1, i), PALE))
    t.setStyle(TableStyle(style))
    return t


def note(text, kind="info"):
    bg, ink = (AMBER, AMBER_INK) if kind == "warn" else (PALE, GREEN)
    t = Table([[P(text, "cell")]], colWidths=[W])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), bg), ("LINEBEFORE", (0, 0), (0, -1), 3, ink),
                           ("LEFTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def steps(items, owner):
    rows = []
    for i, text in enumerate(items, 1):
        badge = Table([[P(f"<font color='white'><b>{i}</b></font>", "cell")]], colWidths=[7 * mm], rowHeights=[7 * mm])
        badge.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LANE[owner]), ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        rows.append([badge, P(text, "step")])
    t = Table(rows, colWidths=[10 * mm, W - 10 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    return t


def phase_header(number, title, owner, duration):
    label = {"IT": "IT TEAM", "KP": "KEY PERSON (PULKIT SINGH)", "USER": "EACH USER"}[owner]
    t = Table([[P(f"<font color='white'><b>PHASE {number}</b></font>", "cell"), P(f"<b>{title}</b>", "cellb"), P(f"<font color='{LANE[owner].hexval()}'><b>{label}</b></font><br/>{duration}", "small")]],
              colWidths=[22 * mm, W - 72 * mm, 50 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), LANE[owner]), ("BACKGROUND", (1, 0), (-1, 0), LANE_BG[owner]), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("LEFTPADDING", (0, 0), (-1, -1), 7)]))
    return t


def shot(name, caption, width=W * 0.82):
    img = Image(f"{SHOTS}/{name}"); ratio = img.imageHeight / img.imageWidth
    img.drawWidth, img.drawHeight = width, width * ratio
    frame = Table([[img]], colWidths=[width + 2])
    frame.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, LINE), ("LEFTPADDING", (0, 0), (-1, -1), 1), ("RIGHTPADDING", (0, 0), (-1, -1), 1), ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    return KeepTogether([frame, Spacer(1, 3), P(caption, "cap")])


class Box:
    def __init__(self, x, y, w, h, title, sub="", fill=CREAM, stroke=GREEN, head=None):
        self.__dict__.update(locals())


class Diagram(Flowable):
    """Simple box-and-arrow diagram drawn on the canvas."""
    def __init__(self, width, height, boxes, arrows, lanes=None):
        super().__init__(); self.width, self.height, self.boxes, self.arrows, self.lanes = width, height, boxes, arrows, lanes or []

    def wrap(self, *_):
        return self.width, self.height

    def draw(self):
        c = self.canv
        for (y, h, label, key) in self.lanes:
            c.setFillColor(LANE_BG[key]); c.setStrokeColor(LANE_BG[key]); c.roundRect(0, y, self.width, h, 6, fill=1, stroke=0)
            c.saveState(); c.translate(11, y + h / 2); c.rotate(90); c.setFillColor(LANE[key]); c.setFont("DVB", 7.5); c.drawCentredString(0, -3, label); c.restoreState()
        for (a, b, label) in self.arrows:
            A, B = self.boxes[a], self.boxes[b]
            if abs((A.y + A.h / 2) - (B.y + B.h / 2)) < 4:   # horizontal
                x1, y1, x2, y2 = (A.x + A.w, A.y + A.h / 2, B.x, B.y + B.h / 2) if B.x > A.x else (A.x, A.y + A.h / 2, B.x + B.w, B.y + B.h / 2)
            else:                                             # vertical
                x1, y1, x2, y2 = (A.x + A.w / 2, A.y, B.x + B.w / 2, B.y + B.h) if B.y < A.y else (A.x + A.w / 2, A.y + A.h, B.x + B.w / 2, B.y)
            c.setStrokeColor(MUTED); c.setLineWidth(1.1)
            if abs(x1 - x2) > 2 and abs(y1 - y2) > 2:
                midy = (y1 + y2) / 2; p = c.beginPath(); p.moveTo(x1, y1); p.lineTo(x1, midy); p.lineTo(x2, midy); p.lineTo(x2, y2); c.drawPath(p, stroke=1, fill=0); hx, hy, dx, dy = x2, y2, 0, (1 if y2 > midy else -1)
            else:
                c.line(x1, y1, x2, y2); length = max(((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5, 1); hx, hy, dx, dy = x2, y2, (x2 - x1) / length, (y2 - y1) / length
            c.setFillColor(MUTED); p = c.beginPath(); p.moveTo(hx, hy); p.lineTo(hx - 6 * dx - 3 * dy, hy - 6 * dy + 3 * dx); p.lineTo(hx - 6 * dx + 3 * dy, hy - 6 * dy - 3 * dx); p.close(); c.drawPath(p, fill=1, stroke=0)
            if label:
                c.setFont("DV", 6.8); c.setFillColor(MUTED); c.drawCentredString((x1 + x2) / 2, (y1 + y2) / 2 + 3, label)
        for b in self.boxes.values():
            c.setFillColor(b.fill); c.setStrokeColor(b.stroke); c.setLineWidth(1); c.roundRect(b.x, b.y, b.w, b.h, 5, fill=1, stroke=1)
            if b.head:
                c.setFillColor(b.stroke); c.setFont("DVB", 6.5); c.drawString(b.x + 6, b.y + b.h - 10, b.head)
            c.setFillColor(INK); c.setFont("DVB", 8.2)
            lines = b.title.split("\n"); sub = b.sub.split("\n") if b.sub else []
            total = len(lines) * 10 + len(sub) * 8.5
            ty = b.y + b.h / 2 + total / 2 - 8 - (4 if b.head else 0)
            for line in lines: c.drawCentredString(b.x + b.w / 2, ty, line); ty -= 10
            c.setFont("DV", 6.8); c.setFillColor(MUTED)
            for line in sub: c.drawCentredString(b.x + b.w / 2, ty, line); ty -= 8.5


def architecture():
    w = W; bx = {
        "pc": Box(0, 70, 135, 66, "VermiTrack desktop app", "Windows / Mac\non each user's PC", head="USERS"),
        "entra": Box(185, 120, 145, 52, "Microsoft Entra ID", "Sign-in with @lnbgroup.com\n(no new passwords)", stroke=LANE["IT"], head="LNB GROUP M365"),
        "sp": Box(185, 50, 145, 52, "SharePoint site “VermiTrack”", "7 lists = shared register", stroke=LANE["IT"], head="LNB GROUP M365"),
        "mail": Box(185, -20 + 0, 145, 52, "Outlook (Exchange Online)", "Notifications & reports\nfrom user's own mailbox", stroke=LANE["IT"], head="LNB GROUP M365"),
        "xl": Box(380, 50, w - 380, 52, "Excel / Power BI", "Optional reporting\ndirectly on the lists", stroke=MUTED),
    }
    for b in bx.values(): b.y += 22
    return Diagram(w, 200, bx, [("pc", "entra", "1. sign in"), ("pc", "sp", "2. read / save"), ("pc", "mail", "3. email"), ("sp", "xl", "")])


def process_flow():
    w = W; lane_h = 64; bw = (w - 30) / 4 - 10
    col = lambda i: 26 + i * (bw + 10)
    y_it, y_kp, y_u = 2 * lane_h + 14, lane_h + 8, 2
    B = lambda c, y, t, s: Box(col(c), y + 8, bw, lane_h - 18, t, s, stroke=MUTED)
    bx = {
        "a": B(0, y_it, "1. SharePoint site", "Private team site\n“VermiTrack”"),
        "b": B(1, y_it, "2. App registration", "Entra ID + 4 delegated\npermissions + consent"),
        "c": B(2, y_it, "3. Deploy installer", "Silent install +\nconfig file to all PCs"),
        "d": B(3, y_kp, "4. One-time setup", "Key Person signs in →\n“Set up on SharePoint”"),
        "e": B(3, y_u, "5. Users sign in", "Open app → Microsoft\n365 sign-in → work"),
        "f": B(2, y_kp, "Approve users", "Admin → Approved users\n(+ site membership)"),
        "g": B(1, y_kp, "Import old records", "Optional: web version\nCSV export"),
    }
    lanes = [(y_it, lane_h, "IT TEAM", "IT"), (y_kp, lane_h, "KEY PERSON", "KP"), (y_u, lane_h, "USERS", "USER")]
    for key, b in bx.items():
        b.stroke = LANE["IT"] if b.y > y_it else LANE["KP"] if b.y > y_kp else LANE["USER"]
    return Diagram(w, 3 * lane_h + 18, bx, [("a", "b", ""), ("b", "c", ""), ("c", "d", "IDs + site"), ("d", "e", ""), ("d", "f", ""), ("f", "g", "")], lanes)


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("DV", 7.5); canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 10 * mm, "Navjyoti VermiTrack — IT deployment & installation process  ·  Version 1.0  ·  October 2026")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
    canvas.setStrokeColor(LINE); canvas.line(18 * mm, 13 * mm, A4[0] - 18 * mm, 13 * mm)
    canvas.restoreState()


def cover(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(DARK); canvas.rect(0, A4[1] - 95 * mm, A4[0], 95 * mm, fill=1, stroke=0)
    canvas.restoreState()


story = []
logo = Image(LOGO); logo.drawWidth = logo.drawHeight = 26 * mm
story += [Spacer(1, 8 * mm), logo, Spacer(1, 8 * mm),
          Paragraph("<font color='#d8ea9a'>LNB GROUP · NAVJYOTI · AKOLA UNIT</font>", ParagraphStyle("e", fontName="DVB", fontSize=9, textColor=colors.white, alignment=TA_CENTER)),
          Spacer(1, 4 * mm),
          Paragraph("Navjyoti VermiTrack", ParagraphStyle("t", fontName="DVB", fontSize=30, leading=36, textColor=colors.white, alignment=TA_CENTER)),
          Paragraph("Desktop app deployment &amp; installation process flow", ParagraphStyle("s", fontName="DV", fontSize=13, leading=18, textColor=colors.HexColor("#cfe3d6"), alignment=TA_CENTER)),
          Spacer(1, 30 * mm)]
story.append(table([["Document", "Details"],
                    ["Prepared for", "LNB Group IT team"],
                    ["Requested by", "Pulkit Singh, DVP (Key Person, VermiTrack)"],
                    ["Purpose", "Install the VermiTrack desktop app for all project users and connect it to LNB Group’s Microsoft 365 (Entra ID sign-in, SharePoint data, Outlook notifications)"],
                    ["Software", "Navjyoti VermiTrack desktop 1.0.0, Windows 10/11 (64-bit) and macOS"],
                    ["Installer", "VermiTrack-Windows-Setup.exe (~95 MB) · VermiTrack-Mac.dmg (~200 MB)"],
                    ["Effort for IT", "About 30–45 minutes one-time, plus standard software deployment"]], [40 * mm, W - 40 * mm]))
story.append(Spacer(1, 10 * mm))
story.append(P("<b>Contents</b>: 1. What the software does · 2. How it works · 3. End-to-end process flow · Phases 1–6 · 4. Security summary · 5. Troubleshooting · 6. Sign-off checklist", "small"))
story.append(PageBreak())

# 1. What it does
story += [P("1. What VermiTrack does", "h1"),
          P("VermiTrack replaces the Excel production registers of the Navjyoti vermicompost project with one shared desktop application. Every authorised colleague records data on their own PC, and everyone sees the same live records, reports and charts."),
          table([["Register", "What is captured"],
                 ["Raw material purchase", "Supplier, material type, quantity, rate, vehicle/invoice, moisture, temperature, lab sampling. Lot code is generated automatically."],
                 ["Pre-composting", "Lot set-up, culture application, turning, temperature and moisture readings, transfer to beds"],
                 ["Bed production &amp; operations", "Bed filling, earthworm application, feeding, watering, status, harvest linkage"],
                 ["Harvesting &amp; packaging", "Net yield after sieving, finished batch number, packaging type, packets, moisture"],
                 ["Quality control", "Lab report, pH, EC, NPK, organic carbon, micronutrients, FCO compliance (raw material and finished goods)"],
                 ["Sales register", "Invoice, customer/dealer, batch, quantity, price, revenue"],
                 ["Stock inventory", "Opening stock, production, sales outflow, loss, physical audit"],
                 ["Cost &amp; expenses", "Category, vendor, payment mode, cost, optimisation notes"]], [42 * mm, W - 42 * mm]),
          Spacer(1, 4),
          P("<b>Built-in controls carried over from the existing web version:</b> production batches; batch-wise consolidated reports and charts; CSV export for Excel; automatic serial numbers and register codes; a complete audit trail of every login and change; a 24-hour edit window (after that, only an administrator can edit); deletion only after approval by Pulkit Singh or Pritam Suryavanshi, with a 90-day recycle bin; and per-user limits on registers and batches."),
          P("<b>New in the desktop edition:</b> sign-in with the LNB Group Microsoft 365 account; data held on LNB Group’s own SharePoint; Outlook email for deletion requests and decisions, user invitations and batch reports.")]
story += [Spacer(1, 4), shot("02-home.png", "Home screen: live batch overview with the imported Batch 4 register (screenshot taken in demo mode)", W * 0.78)]
story.append(PageBreak())

# 2. How it works
story += [P("2. How it works", "h1"),
          P("VermiTrack is a desktop application only; <b>there is no external server or third-party cloud</b>. It uses three services LNB Group already has in Microsoft 365, through Microsoft Graph:"),
          architecture(), Spacer(1, 6),
          table([["Component", "Role", "Owner"],
                 ["VermiTrack desktop app", "User interface on each PC. Holds no business data, only connection settings and an encrypted sign-in token.", "IT deploys"],
                 ["Microsoft Entra ID", "Sign-in with existing @lnbgroup.com accounts (MFA and conditional access apply as usual). One app registration.", "IT"],
                 ["SharePoint site “VermiTrack”", "Stores the register in 7 lists. Site membership controls who can reach the data.", "IT creates; Key Person manages"],
                 ["Outlook / Exchange Online", "Sends notifications and reports from the signed-in user’s own mailbox (visible in Sent Items).", "Existing"]],
                [44 * mm, W - 78 * mm, 34 * mm]),
          Spacer(1, 6),
          P("SharePoint lists created automatically in Phase 4", "h2"),
          table([["List", "Contents"],
                 ["VermiTrack Entries", "All register entries (8 registers), one item per entry"],
                 ["VermiTrack Batches", "Production batches (Batch 3, 4, 5 …)"],
                 ["VermiTrack Baselines", "Imported Excel baseline totals"],
                 ["VermiTrack Users", "Approved users, roles and register/batch permissions"],
                 ["VermiTrack Activity", "Audit trail of every action"],
                 ["VermiTrack Deletion Requests", "Deletion requests and approvals"],
                 ["VermiTrack Sign-ins", "Sign-in and session log"]], [55 * mm, W - 55 * mm])]
story.append(PageBreak())

# 3. Process flow
story += [P("3. End-to-end process flow", "h1"),
          P("The rollout has six phases. IT completes Phases 1–3 once, the Key Person completes Phases 4 and 6, and each user completes Phase 5 in about two minutes."),
          process_flow(), Spacer(1, 8),
          table([["#", "Activity", "Done by", "Time", "Output"],
                 ["1", "Create the private SharePoint team site “VermiTrack” and add members", "IT", "5 min", "Site URL"],
                 ["2", "Register the app in Entra ID, add 4 delegated Graph permissions, grant admin consent", "IT (Global / Application Admin)", "10 min", "Client ID + Tenant ID"],
                 ["3", "Install the app on users’ PCs (manually or silently) with the pre-configured connection file", "IT", "Standard deployment", "App on every PC"],
                 ["4", "Sign in, create the lists and load the existing Batch 3–5 data", "Key Person", "3 min", "Live shared register"],
                 ["5", "Users open the app and sign in with their Microsoft 365 account", "Each user", "2 min", "User working"],
                 ["6", "Approve new users; optionally import records from the old web version", "Key Person", "As needed", "Complete data"]],
                [12 * mm, 70 * mm, 34 * mm, 22 * mm, W - 138 * mm])]
story.append(PageBreak())

# Phase 1
story += [phase_header(1, "Create the SharePoint site", "IT", "≈ 5 minutes"), Spacer(1, 6),
          steps(["Open <b>SharePoint admin center</b> → Sites → Active sites → <b>+ Create</b> → <b>Team site</b>. Alternatively, use SharePoint home → <b>+ Create site</b>.",
                 "Site name: <b>VermiTrack</b>. Privacy: <b>Private</b>. Language: English. Owners: <b>Pulkit Singh</b> (and IT).",
                 "Add the project users as <b>Members</b> (edit rights). Initial users: Pulkit Singh, Ashok Pawar, Harish Venkatakrishnan, Akola Team, Pritam Suryavanshi. Add view-only users as <b>Visitors</b>.",
                 "Note the site address, for example <font face='DVM'>https://lnbgroup.sharepoint.com/sites/VermiTrack</font>."], "IT"),
          note("Pulkit Singh must be a site <b>Owner</b>: creating the lists in Phase 4 needs owner rights. Use a <b>Private</b> site; do not grant access to “Everyone except external users”."),
          Spacer(1, 10),
          phase_header(2, "Register the app in Microsoft Entra ID", "IT", "≈ 10 minutes · needs Global / Application Administrator"), Spacer(1, 6),
          steps(["Open <b>entra.microsoft.com</b> → Identity → Applications → <b>App registrations</b> → <b>+ New registration</b>.",
                 "Name: <b>Navjyoti VermiTrack</b>. Supported account types: <b>Accounts in this organizational directory only (single tenant)</b>.",
                 "Redirect URI: platform <b>Public client/native (mobile &amp; desktop)</b>, value <font face='DVM'>http://localhost</font>. Click <b>Register</b>.",
                 "From the Overview page, copy the <b>Application (client) ID</b> and the <b>Directory (tenant) ID</b>.",
                 "Open <b>API permissions</b> → + Add a permission → Microsoft Graph → <b>Delegated permissions</b>, and add the four permissions in the table below.",
                 "Click <b>Grant admin consent for LNB Group</b> → Yes. All four permissions should show a green “Granted” status.",
                 "No client secret or certificate is needed (it is a public desktop client)."], "IT"),
          table([["Permission (delegated)", "Purpose", "Consent"],
                 ["User.Read", "Identify the signed-in user (name, email)", "Granted for all users in step 6"],
                 ["Sites.ReadWrite.All", "Read and save register items in the VermiTrack lists", "Granted for all users in step 6"],
                 ["Sites.Manage.All", "Create the 7 lists on the site (first setup only)", "Granted for all users in step 6"],
                 ["Mail.Send", "Send notifications and reports from the user’s own mailbox", "Granted for all users in step 6"]], [42 * mm, W - 94 * mm, 52 * mm]),
          Spacer(1, 4),
          note("<b>Delegated permissions only.</b> The app acts as the signed-in person and can reach only what that person can already reach. A user who is not a member of the VermiTrack site cannot open its lists. There are no application (background) permissions. Optional hardening: under Enterprise applications → Navjyoti VermiTrack → Properties, set <i>Assignment required = Yes</i> and assign the project users or a security group.")]
story.append(PageBreak())

# Phase 3
story += [phase_header(3, "Install the software on users’ computers", "IT", "Standard deployment"), Spacer(1, 6),
          P("3A. Pre-configure the connection (recommended)", "h2"),
          P("Create the file below on each PC (via Intune, GPO or a script). VermiTrack then opens straight to Microsoft 365 sign-in, and users never type IDs or a setup code."),
          table([["Platform", "File location"],
                 ["Windows", "<font face='DVM'>C:\\ProgramData\\VermiTrack\\vermitrack.config.json</font>"],
                 ["macOS", "<font face='DVM'>/Library/Application Support/VermiTrack/vermitrack.config.json</font>"]], [30 * mm, W - 30 * mm]),
          Paragraph('{<br/>&nbsp;&nbsp;"clientId": "&lt;Application (client) ID from Phase 2&gt;",<br/>&nbsp;&nbsp;"tenantId": "&lt;Directory (tenant) ID from Phase 2&gt;",<br/>&nbsp;&nbsp;"siteUrl": "https://lnbgroup.sharepoint.com/sites/VermiTrack"<br/>}', ss["code"]),
          P("3B. Install the application", "h2"),
          table([["Method", "How"],
                 ["Silent, all users (recommended)", "<font face='DVM'>VermiTrack-Windows-Setup.exe /S /allusers</font> (installs to Program Files; Start menu and desktop shortcuts are created)"],
                 ["Silent, current user", "<font face='DVM'>VermiTrack-Windows-Setup.exe /S /currentuser</font> (no admin rights needed)"],
                 ["Microsoft Intune", "Package the .exe as a Win32 app (IntuneWinAppUtil). Install command: <font face='DVM'>VermiTrack-Windows-Setup.exe /S /allusers</font>. Detection: file <font face='DVM'>%ProgramFiles%\\Navjyoti VermiTrack\\Navjyoti VermiTrack.exe</font>. Deploy the config file as a separate script or as part of the package."],
                 ["Manual (single PC)", "Double-click the installer → choose for all users or only me → Next → Install → Finish."],
                 ["macOS", "Open VermiTrack-Mac.dmg and drag the app to Applications. On first run, right-click → Open."],
                 ["Uninstall", "Settings → Apps → Navjyoti VermiTrack → Uninstall (silent: run the uninstaller with <font face='DVM'>/S</font>)."]], [42 * mm, W - 42 * mm]),
          Spacer(1, 5),
          note("<b>Code signing:</b> version 1.0 is not yet signed with a code-signing certificate, so Edge and Windows SmartScreen warn that the file is uncommon. In Edge: Downloads → ⋯ → Keep → Show more → Keep anyway. On launch: More info → Run anyway. For organisation-wide rollout, IT may (a) allow the file hash in Defender / SmartScreen, or (b) sign the installer with LNB Group’s code-signing certificate. The source code is available for IT review.", "warn"),
          Spacer(1, 6),
          table([["Requirement", "Detail"],
                 ["Operating system", "Windows 10/11 64-bit, or macOS 12+"],
                 ["Disk / memory", "≈ 350 MB installed, 4 GB RAM"],
                 ["Network (HTTPS 443)", "login.microsoftonline.com, graph.microsoft.com, lnbgroup.sharepoint.com"],
                 ["Local data", "%APPDATA%\\vermitrack-desktop: settings, encrypted token cache (Windows DPAPI / macOS Keychain) and log. Register data stays in SharePoint."]], [42 * mm, W - 42 * mm])]
story.append(PageBreak())

# Phase 4
story += [phase_header(4, "One-time setup on SharePoint", "KP", "≈ 3 minutes · Pulkit Singh"), Spacer(1, 6),
          steps(["Open <b>Navjyoti VermiTrack</b> from the Start menu. If IT deployed the config file, it opens on <b>Sign in with Microsoft 365</b>. Otherwise choose <b>First-time setup (administrator)</b> and enter the Client ID, Tenant ID and site URL.",
                 "Click <b>Sign in with Microsoft 365</b>. The browser opens; sign in as pulkit.singh@lnbgroup.com and return to the app.",
                 "Click <b>Set up VermiTrack on this SharePoint site</b>. The app creates the 7 lists and loads the existing data: 149 Batch 3–5 register entries, the Batch 4 baseline and the 5 approved users.",
                 "The Home screen appears with the live Batch 4 figures. Setup is complete for the whole organisation."], "KP"),
          shot("01-setup.png", "First-time setup screen (only shown when IT has not pre-configured the connection)", W * 0.6)]
story.append(PageBreak())

# Phase 5 & 6
story += [phase_header(5, "User sign-in and daily use", "USER", "≈ 2 minutes per user"), Spacer(1, 6),
          steps(["The user must be (a) listed in VermiTrack → Admin → <b>Approved users</b> and (b) a Member of the VermiTrack SharePoint site.",
                 "Open VermiTrack → <b>Sign in with Microsoft 365</b> → sign in with the @lnbgroup.com account (MFA as usual). The sign-in is remembered.",
                 "Use <b>Capture</b> to record entries, <b>Records</b> to search, edit or export, and <b>Reports</b> for batch reports, charts, Excel export and <b>✉ Email report</b>. Data refreshes automatically every 20 seconds."], "USER"),
          note("Without the IT config file, users choose <b>I have a setup code</b> and paste the code the Key Person sends from Admin → Approved users → ✉ Email setup instructions."),
          Spacer(1, 10),
          phase_header(6, "Manage users and bring over existing records", "KP", "As needed"), Spacer(1, 6),
          steps(["<b>Add a user</b>: Admin (initials, top right) → Approved users → + Add an approved user (email = Microsoft 365 sign-in), choose role, registers and batches. Ask IT to add them to the SharePoint site.",
                 "<b>Remove a user</b>: set the account to Inactive in Approved users, and ask IT to remove them from the site.",
                 "<b>Records from the old web version</b> (optional): in the old web app, Records → All registers → Export CSV for each batch; in the desktop app, Admin → <b>Bring records from the web version</b> → choose the file. Duplicates are skipped automatically."], "KP"),
          shot("05-admin.png", "Admin panel: deletion approvals, recycle bin, approved users and audit trail", W * 0.7)]
story.append(PageBreak())

# Security, troubleshooting, checklist
story += [P("4. Security summary", "h1"),
          table([["Topic", "How it is handled"],
                 ["Identity", "Microsoft Entra ID sign-in only; LNB Group MFA and conditional access policies apply. No separate passwords."],
                 ["Data location", "LNB Group’s own SharePoint Online tenant. No external servers, databases or third-party cloud services."],
                 ["Access control", "Two layers: SharePoint site membership (IT) and the VermiTrack approved-user list with roles and register/batch limits (Key Person)."],
                 ["Permissions", "Delegated Microsoft Graph permissions only; no application permissions or client secrets."],
                 ["Audit", "In-app audit trail of every sign-in, entry, edit, deletion request and approval, plus SharePoint version history on every list item."],
                 ["Deletion", "Two-person rule: deletions need approval by Pulkit Singh or Pritam Suryavanshi; restorable for 90 days."],
                 ["On the PC", "Sign-in token encrypted with Windows DPAPI / macOS Keychain. No register data is kept on the PC."],
                 ["Point to note", "The app enforces roles and approvals. A site Member could still edit list items directly in SharePoint, so keep site membership to the VermiTrack team; such edits show in SharePoint version history."]], [32 * mm, W - 32 * mm]),
          Spacer(1, 6), P("5. Troubleshooting", "h1"),
          table([["Message / symptom", "Cause and fix"],
                 ["“Need admin approval” at sign-in", "Admin consent not granted. Complete Phase 2, step 6."],
                 ["“… is not an approved VermiTrack user”", "Add the person under Admin → Approved users with their exact sign-in email."],
                 ["“does not have permission for the VermiTrack SharePoint site”", "Add the user as a Member of the SharePoint site."],
                 ["“Sign-in did not complete”", "Check that the redirect URI <font face='DVM'>http://localhost</font> is under the <b>Public client/native</b> platform; check that the browser can reach login.microsoftonline.com."],
                 ["List creation fails in Phase 4", "The signed-in person must be a site Owner, and Sites.Manage.All must be consented."],
                 ["SmartScreen / Edge blocks the download", "See the code-signing note in Phase 3."],
                 ["Anything else", "File → Open data folder → send vermitrack.log to the Key Person."]], [62 * mm, W - 62 * mm]),
          PageBreak(), P("6. Sign-off checklist", "h1")]
check = lambda t: ["☐", t]
story += [table([["", "IT team"]] + [check(t) for t in [
                    "Private SharePoint team site “VermiTrack” created; Pulkit Singh is an Owner",
                    "Project users added as Members (view-only users as Visitors)",
                    "Entra app registration “Navjyoti VermiTrack” created (single tenant, public client, http://localhost)",
                    "4 delegated Graph permissions added and admin consent granted",
                    "Client ID, Tenant ID and site URL shared with Pulkit Singh",
                    "vermitrack.config.json deployed to target PCs",
                    "Installer deployed (/S /allusers) and SmartScreen allowance or code signing handled"]], [10 * mm, W - 10 * mm], zebra=False),
          Spacer(1, 8),
          table([["", "Key Person"]] + [check(t) for t in [
                    "Signed in and completed “Set up VermiTrack on this SharePoint site”",
                    "Verified Batch 4 figures on the Home and Reports screens",
                    "Approved-user list reviewed; roles and permitted registers/batches set",
                    "Old web-version records imported (if applicable)",
                    "Each user confirmed able to sign in and save a test entry"]], [10 * mm, W - 10 * mm], zebra=False),
          Spacer(1, 14),
          table([["Role", "Name", "Signature / date"], ["IT", "", ""], ["Key Person", "Pulkit Singh", ""]], [40 * mm, 60 * mm, W - 100 * mm], zebra=False)]

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                        title="Navjyoti VermiTrack – IT deployment & installation process", author="Pulkit Singh, LNB Group")
doc.build(story, onFirstPage=cover, onLaterPages=footer)
print("written", OUT)
