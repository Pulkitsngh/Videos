"""Build the illustrated Microsoft 365 setup guide (guide.html) for the Navjyoti tracker.

Figures are simplified drawings of each screen (not screenshots). Numbered markers in a
figure match the numbered steps beside it.
"""
import html
import os

W, H = 720, 400


def esc(s):
    return html.escape(str(s), quote=True)


class Fig:
    def __init__(self, title, h=H):
        self.h = h
        self.parts = [
            f'<rect x="1" y="1" width="{W - 2}" height="{h - 2}" rx="10" class="f-win"/>',
            f'<rect x="1" y="1" width="{W - 2}" height="34" rx="10" class="f-bar"/>',
            f'<rect x="1" y="24" width="{W - 2}" height="11" class="f-bar"/>',
            *[f'<circle cx="{20 + i * 16}" cy="18" r="5" class="f-dot"/>' for i in range(3)],
            f'<text x="{W / 2}" y="23" text-anchor="middle" class="f-t f-muted" font-size="12">{esc(title)}</text>',
        ]

    def rect(self, x, y, w, h, cls="f-panel", rx=6):
        self.parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" class="{cls}"/>')
        return self

    def text(self, x, y, s, size=12, cls="f-t", anchor="start", weight=400):
        self.parts.append(f'<text x="{x}" y="{y}" font-size="{size}" text-anchor="{anchor}" font-weight="{weight}" class="{cls}">{esc(s)}</text>')
        return self

    def mono(self, x, y, s, size=12, cls="f-code"):
        self.parts.append(f'<text x="{x}" y="{y}" font-size="{size}" class="{cls} f-mono">{esc(s)}</text>')
        return self

    def btn(self, x, y, w, h, label, primary=False, size=12):
        self.rect(x, y, w, h, "f-btn-p" if primary else "f-btn", 6)
        return self.text(x + w / 2, y + h / 2 + size * 0.36, label, size, "f-t-inv" if primary else "f-t", "middle", 600)

    def line(self, x1, y1, x2, y2, cls="f-line"):
        self.parts.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" class="{cls}"/>')
        return self

    def arrow(self, x1, y1, x2, y2):
        self.parts.append(f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" class="f-arrow" marker-end="url(#ah)"/>')
        return self

    def hi(self, x, y, w, h, n):
        self.parts.append(f'<rect x="{x - 4}" y="{y - 4}" width="{w + 8}" height="{h + 8}" rx="8" class="f-hi"/>')
        cx, cy = x + w + 4, y - 4
        self.parts.append(f'<circle cx="{cx}" cy="{cy}" r="11" class="f-badge"/>'
                          f'<text x="{cx}" y="{cy + 4.5}" font-size="12" font-weight="700" text-anchor="middle" class="f-t-inv">{n}</text>')
        return self

    def svg(self, label):
        return (f'<div class="fig-scroll"><svg viewBox="0 0 {W} {self.h}" role="img" aria-label="{esc(label)}">'
                f'<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
                f'<path d="M0,0 L10,5 L0,10 z" class="f-arrowhead"/></marker></defs>'
                + "".join(self.parts) + "</svg></div>")


# ---------------------------------------------------------------- figures
def fig_overview():
    f = Fig("How the pieces connect", 330)
    boxes = [
        (30, 70, "Supervisors", "phone / PC", "f-panel"),
        (200, 70, "Power Apps", "data-entry app", "f-c1"),
        (390, 70, "SharePoint lists", "10 registers · the data", "f-c2"),
        (560, 70, "Power BI", "dashboards & reports", "f-c3"),
        (390, 210, "Power Automate", "alerts: overdue beds, FCO…", "f-c4"),
        (560, 210, "Microsoft Teams", "one place for all of it", "f-panel"),
    ]
    for x, y, t, s, c in boxes:
        f.rect(x, y, 140 if x != 390 else 150, 70, c, 10)
        f.text(x + (70 if x != 390 else 75), y + 30, t, 14, "f-t", "middle", 700)
        f.text(x + (70 if x != 390 else 75), y + 50, s, 11, "f-muted", "middle")
    f.arrow(170, 105, 198, 105).arrow(340, 105, 388, 105).arrow(540, 105, 558, 105)
    f.arrow(465, 140, 465, 208).arrow(630, 140, 630, 208).arrow(540, 245, 558, 245)
    f.text(30, 300, "Everyone signs in with their normal company Microsoft 365 account.", 12, "f-muted")
    return f.svg("Supervisors enter data in Power Apps, which saves to SharePoint lists; Power BI reads the lists for reports; Power Automate sends alerts; Teams shows it all")


def fig_create_site():
    f = Fig("SharePoint start page")
    f.rect(20, 50, 680, 40, "f-panel").text(36, 75, "SharePoint", 15, "f-t", weight=700)
    f.btn(40, 110, 130, 34, "+ Create site", True).hi(40, 110, 130, 34, 1)
    f.rect(220, 110, 460, 250, "f-panel", 10).text(240, 140, "Create a site", 15, weight=700)
    f.rect(240, 160, 200, 170, "f-card", 8).text(340, 250, "Team site", 14, "f-t", "middle", 700).text(340, 272, "for working together", 11, "f-muted", "middle")
    f.hi(240, 160, 200, 170, 2)
    f.rect(460, 160, 200, 170, "f-card", 8).text(560, 250, "Communication site", 13, "f-muted", "middle")
    return f.svg("Click Create site, then choose Team site")


def fig_site_name():
    f = Fig("Team site – name and members")
    f.rect(120, 55, 480, 320, "f-panel", 10).text(145, 90, "Give your site a name", 15, weight=700)
    f.text(145, 120, "Site name", 12, "f-muted").rect(145, 128, 430, 34, "f-input").text(157, 150, "Navjyoti Production", 13).hi(145, 128, 430, 34, 3)
    f.text(145, 190, "Site address", 12, "f-muted").rect(145, 198, 430, 30, "f-input").mono(157, 218, "…/sites/NavjyotiProduction", 12)
    f.text(145, 258, "Add members", 12, "f-muted").rect(145, 266, 430, 34, "f-input").text(157, 288, "Supervisors, production team…", 12, "f-muted").hi(145, 266, 430, 34, 4)
    f.btn(475, 325, 100, 34, "Finish", True)
    return f.svg("Name the site Navjyoti Production and add the team as members")


def fig_powershell():
    f = Fig("PowerShell 7", 380)
    f.rect(12, 44, 696, 324, "f-term", 6)
    lines = [
        ("PS> Install-Module PnP.PowerShell -Scope CurrentUser", "f-term-t", 1),
        ("PS> Register-PnPEntraIDAppForInteractiveLogin -ApplicationName \"PnP Navjyoti\" …", "f-term-t", 2),
        ("    Client ID: 1a2b3c4d-…   ← copy this", "f-term-ok", None),
        ("PS> ./Create-NavjyotiLists.ps1 -SiteUrl https://…/sites/NavjyotiProduction -ClientId 1a2b…", "f-term-t", 3),
        ("Created list Batches", "f-term-ok", None),
        ("  Batches : 2 rows loaded", "f-term-ok", None),
        ("Created list RawMaterial", "f-term-ok", None),
        ("  RawMaterial : 15 rows loaded", "f-term-ok", None),
        ("  …", "f-term-t", None),
        ("  Harvest : 80 rows loaded", "f-term-ok", None),
        ("Done. Open …/_layouts/15/viewlsts.aspx to see the lists.", "f-term-ok", None),
    ]
    for i, (s, c, n) in enumerate(lines):
        y = 72 + i * 27
        f.mono(28, y, s, 12, c)
        if n:
            f.hi(22, y - 16, 660, 22, n)
    return f.svg("PowerShell commands: install PnP, register the app, run the Navjyoti script; output shows lists created and rows loaded")


def fig_site_contents():
    f = Fig("Navjyoti Production › Site contents")
    f.text(30, 70, "Site contents", 16, weight=700)
    rows = [("Batches", 2), ("RawMaterial", 15), ("PreCompost", 3), ("Beds", 80), ("Harvest", 80),
            ("QualityControl", 6), ("Sales", 6), ("StockLedger", 63), ("Expenses", 30), ("DailyLog", 0)]
    f.rect(30, 85, 660, 28, "f-head", 4).text(44, 104, "Name", 12, "f-t-inv", weight=700).text(520, 104, "Items", 12, "f-t-inv", weight=700)
    for i, (n, c) in enumerate(rows):
        y = 116 + i * 27
        f.rect(30, y, 660, 26, "f-row" if i % 2 else "f-win", 0).text(44, y + 18, "▦  " + n, 12).text(520, y + 18, str(c), 12, "f-t", weight=600)
    f.hi(30, 116, 660, 26 * 10 + 4, 4)
    return f.svg("Site contents lists the 10 registers with their item counts")


def fig_pa_start():
    f = Fig("make.powerapps.com")
    f.rect(20, 50, 150, 330, "f-panel", 8)
    for i, s in enumerate(["Home", "+ Create", "Apps", "Tables", "Flows"]):
        f.text(40, 85 + i * 32, s, 13, "f-t", weight=600 if s == "+ Create" else 400)
    f.hi(32, 96, 110, 26, 1)
    f.text(195, 80, "Start from", 15, weight=700)
    for i, s in enumerate(["Blank app", "Start with data", "Page design"]):
        f.rect(195 + i * 170, 95, 155, 90, "f-card", 8).text(272 + i * 170, 146, s, 13, "f-t", "middle", 600)
    f.hi(365, 95, 155, 90, 2)
    f.rect(195, 205, 500, 170, "f-panel", 10).text(215, 232, "Connect to data", 14, weight=700)
    f.rect(215, 245, 150, 36, "f-card").text(290, 268, "SharePoint", 13, "f-t", "middle", 600).hi(215, 245, 150, 36, 3)
    f.rect(385, 245, 290, 115, "f-card").text(400, 268, "Site: NavjyotiProduction", 12, "f-muted")
    for i, s in enumerate(["☐ Beds", "☐ Harvest", "☑ DailyLog"]):
        f.text(400, 292 + i * 20, s, 12)
    f.hi(395, 318, 110, 22, 4)
    return f.svg("Power Apps: Create, Start with data, SharePoint, pick the Navjyoti site and the DailyLog list")


def fig_pa_studio():
    f = Fig("Power Apps Studio – Bed board screen")
    f.rect(20, 45, 680, 32, "f-panel").mono(30, 66, "Items  =  SortByColumns(Filter(Beds, Batch.Value = varBatch), \"BedNo\")", 12).hi(20, 45, 680, 32, 5)
    f.rect(20, 88, 150, 290, "f-panel", 8).text(32, 110, "Tree view", 12, "f-muted", weight=700)
    for i, s in enumerate(["scrHome", "scrLog", "scrBeds", "  galBeds", "  conDetail", "scrHarvest", "scrQC"]):
        f.text(34, 136 + i * 24, s, 12, "f-t", weight=700 if s.strip() == "galBeds" else 400)
    f.rect(270, 88, 240, 290, "f-phone", 18)
    colors = ["f-c3", "f-c3", "f-c1", "f-c2", "f-c3", "f-c1", "f-c2", "f-c3", "f-c3", "f-c1", "f-c1", "f-c2"]
    for i in range(12):
        x, y = 285 + (i % 3) * 72, 120 + (i // 3) * 62
        f.rect(x, y, 64, 52, colors[i], 6).text(x + 32, y + 24, f"BED-{i + 1:02d}", 10, "f-t", "middle", 700)
        f.text(x + 32, y + 40, ["Harvested", "Harvested", "Active", "Overdue", "Harvested", "Active", "Overdue", "Harvested", "Harvested", "Active", "Active", "Overdue"][i], 9, "f-muted", "middle")
    f.rect(530, 88, 170, 290, "f-panel", 8).text(542, 110, "Properties", 12, "f-muted", weight=700)
    for i, s in enumerate(["Items", "OnSelect", "Wrap count: 3", "Fill (status colour)"]):
        f.text(544, 138 + i * 26, s, 12)
    return f.svg("Power Apps Studio with the bed board gallery and its Items formula")


def fig_pa_popup():
    f = Fig("Navjyoti app on a phone – tap a bed")
    for px, title in ((90, "Bed board · B5"), (410, "Tap → pop-up")):
        f.rect(px, 45, 220, 340, "f-phone", 22).text(px + 110, 75, title, 12, "f-t", "middle", 700)
        for i in range(9):
            x, y = px + 16 + (i % 3) * 64, 90 + (i // 3) * 60
            f.rect(x, y, 58, 50, ["f-c3", "f-c1", "f-c2"][i % 3], 6).text(x + 29, y + 30, f"BED-{i + 1:02d}", 10, "f-t", "middle", 700)
    f.hi(106, 90, 58, 50, 6)
    f.arrow(318, 200, 402, 200)
    f.rect(410, 45, 220, 340, "f-dim", 22)
    f.rect(425, 130, 190, 220, "f-win", 10).text(440, 156, "BED-01 · Batch 5", 13, weight=700)
    for i, (a, b) in enumerate([("Status", "Active"), ("Filled", "14 Feb"), ("Worms", "4 kg"), ("Expected", "16 Apr"), ("Net yield", "—")]):
        f.text(440, 182 + i * 22, a, 11, "f-muted").text(600, 182 + i * 22, b, 11, "f-t", "end", 600)
    f.btn(440, 305, 75, 30, "Close").btn(525, 305, 75, 30, "Save", True)
    f.hi(425, 130, 190, 220, 7)
    return f.svg("Tapping a bed on the phone app opens a pop-up panel with its details and Save")


def fig_pbi_getdata():
    f = Fig("Power BI Desktop – Get data")
    f.rect(20, 45, 680, 44, "f-panel").btn(34, 52, 100, 30, "Get data", True).hi(34, 52, 100, 30, 1)
    f.rect(150, 105, 420, 270, "f-panel", 10).text(170, 135, "SharePoint Online list", 15, weight=700).hi(170, 118, 200, 24, 2)
    f.text(170, 170, "Site URL", 12, "f-muted").rect(170, 178, 380, 32, "f-input").mono(180, 199, "https://…/sites/NavjyotiProduction", 12).hi(170, 178, 380, 32, 3)
    f.text(170, 238, "Implementation", 12, "f-muted").text(170, 262, "◉ 2.0     ○ 1.0", 13).hi(166, 247, 90, 22, 4)
    f.btn(460, 325, 90, 32, "OK", True)
    return f.svg("Power BI Desktop: Get data, SharePoint Online list, paste the site URL, choose implementation 2.0")


def fig_pbi_model():
    f = Fig("Power BI – Model view")
    f.rect(290, 165, 140, 70, "f-c1", 10).text(360, 196, "Batches", 14, "f-t", "middle", 700).text(360, 216, "BatchCode", 11, "f-muted", "middle")
    tables = [(40, 60, "RawMaterial"), (270, 50, "Beds"), (500, 60, "Harvest"), (40, 175, "PreCompost"), (560, 175, "QualityControl"),
              (40, 290, "Sales"), (215, 305, "StockLedger"), (400, 305, "Expenses"), (560, 290, "DailyLog")]
    for x, y, t in tables:
        f.rect(x, y, 130, 46, "f-card", 8).text(x + 65, y + 22, t, 12, "f-t", "middle", 700).text(x + 65, y + 38, "Batch", 10, "f-muted", "middle")
        f.line(360, 200, x + 65, y + 23, "f-rel")
    f.rect(290, 165, 140, 70, "f-c1", 10).text(360, 196, "Batches", 14, "f-t", "middle", 700).text(360, 216, "BatchCode  1 → *", 11, "f-muted", "middle")
    f.hi(290, 165, 140, 70, 5)
    return f.svg("Model view: Batches connected one-to-many to every list's Batch column")


def fig_pbi_report():
    f = Fig("Power BI report – Overview page")
    f.rect(20, 45, 140, 330, "f-panel", 8).text(32, 70, "Batch", 12, "f-muted", weight=700)
    for i, s in enumerate(["☑ B4", "☑ B5"]):
        f.text(34, 96 + i * 22, s, 12)
    f.hi(28, 80, 124, 50, 6)
    for i, (t, v) in enumerate([("Net yield", "43,890 kg"), ("Conversion", "27.4%"), ("Revenue", "₹3,82,596"), ("Stock", "13,005 kg")]):
        x = 175 + i * 132
        f.rect(x, 45, 122, 62, "f-card", 8).text(x + 10, 66, t, 11, "f-muted").text(x + 10, 94, v, 15, "f-t", weight=700)
    f.rect(175, 118, 320, 257, "f-card", 8).text(187, 138, "Net yield vs sold by month", 12, weight=700)
    for i, (a, b) in enumerate([(150, 140), (35, 0), (70, 15), (52, 40)]):
        x = 205 + i * 70
        f.rect(x, 350 - a, 22, a, "f-c3", 2).rect(x + 24, 350 - b, 22, b, "f-c2", 2)
    f.line(190, 350, 480, 350)
    f.rect(505, 118, 195, 257, "f-card", 8).text(517, 138, "Quality: RM · FG · EX", 12, weight=700)
    for i, (a, b, c) in enumerate([(120, 80, 46), (62, 50, 48), (30, 40, 44)]):
        x = 525 + i * 58
        f.rect(x, 350 - a, 14, a, "f-c1", 2).rect(x + 15, 350 - b, 14, b, "f-c2", 2).rect(x + 30, 350 - c, 14, c, "f-c3", 2)
    f.line(515, 290, 690, 290, "f-ref")
    f.text(690, 285, "FCO", 10, "f-muted", "end")
    f.hi(505, 118, 195, 257, 7)
    return f.svg("Overview report page with batch slicer, KPI cards, net yield vs sold chart and the RM/FG/EX quality chart")


def fig_pbi_refresh():
    f = Fig("app.powerbi.com – Semantic model settings")
    f.rect(20, 45, 680, 330, "f-panel", 10).text(40, 78, "Navjyoti Tracker · Settings", 15, weight=700)
    f.text(40, 115, "▸ Data source credentials", 13, weight=600).btn(470, 98, 200, 28, "Edit credentials").hi(470, 98, 200, 28, 8)
    f.text(60, 142, "Authentication: OAuth2 · organisational account", 12, "f-muted")
    f.text(40, 185, "▾ Refresh", 13, weight=600)
    f.text(60, 212, "Keep your data up to date", 12).rect(470, 198, 60, 24, "f-btn-p", 12).hi(470, 198, 60, 24, 9)
    for i, s in enumerate(["07:00", "10:00", "13:00", "16:00", "19:00"]):
        f.rect(60 + i * 90, 235, 80, 28, "f-input").text(100 + i * 90, 254, s, 12, "f-t", "middle")
    f.text(60, 290, "Time zone: (UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi", 12, "f-muted")
    f.btn(560, 325, 110, 32, "Apply", True)
    return f.svg("Power BI service: set credentials to OAuth2 and turn on scheduled refresh")


def fig_flow():
    f = Fig("Power Automate – Overdue beds flow")
    steps = [("⏰  Recurrence", "Every day at 08:00", "f-c4"), ("▦  Get items · Beds", "Status ne 'Harvested' and ExpectedHarvest lt today", "f-c2"),
             ("◇  Condition", "number of items > 0", "f-card"), ("✉  Post message in Teams", "⚠ Beds past expected harvest + table", "f-c1")]
    for i, (t, s, c) in enumerate(steps):
        y = 50 + i * 85
        f.rect(170, y, 380, 62, c, 8).text(190, y + 26, t, 14, "f-t", weight=700).text(190, y + 47, s, 11, "f-muted")
        if i:
            f.arrow(360, y - 23, 360, y - 2)
        f.hi(170, y, 380, 62, i + 1)
    return f.svg("Flow: recurrence, get overdue beds from SharePoint, condition, post to Teams")


def fig_teams():
    f = Fig("Microsoft Teams – Navjyoti Production › General")
    f.rect(20, 45, 140, 330, "f-panel", 8).text(32, 72, "Navjyoti Production", 11, "f-t", weight=700).text(40, 96, "# General", 12)
    f.rect(175, 45, 525, 40, "f-panel", 6)
    for i, s in enumerate(["Posts", "Files", "Navjyoti app", "Dashboard", "Beds"]):
        f.text(195 + i * 88, 71, s, 12, "f-t", weight=700 if i >= 2 else 400)
    f.hi(362, 55, 240, 22, 2)
    f.text(640, 71, "+", 18, "f-t", weight=700).hi(632, 55, 20, 22, 1)
    f.rect(260, 110, 360, 240, "f-panel", 10).text(280, 140, "Add a tab", 15, weight=700)
    for i, s in enumerate(["Power Apps", "Power BI", "Lists", "Excel"]):
        x = 280 + (i % 2) * 165
        y = 160 + (i // 2) * 80
        f.rect(x, y, 150, 64, "f-card", 8).text(x + 75, y + 38, s, 13, "f-t", "middle", 600)
    return f.svg("Teams: click + to add the Power Apps, Power BI and Lists tabs")


# ---------------------------------------------------------------- page
STAGES = [
    ("site", "Create the SharePoint site", "IT admin · 10 minutes", [
        "Go to <b>office.com</b> → <b>SharePoint</b> → <b>+ Create site</b>.",
        "Choose <b>Team site</b>. It also creates a matching Microsoft Teams team.",
        "Name it <b>Navjyoti Production</b>. Note the site address.",
        "Add supervisors as <b>Members</b> (can edit) and managers as <b>Visitors</b> (view only).",
    ], [fig_create_site(), fig_site_name()], None),
    ("lists", "Create the 10 lists and load Batch 4", "IT admin · 20 minutes", [
        "Install <b>PowerShell 7</b> from the Microsoft Store, open it, and run the install command.",
        "A Microsoft 365 admin registers the PnP app once. Replace <b>yourcompany</b> with the name in your SharePoint address (<i>https://<b>yourcompany</b>.sharepoint.com</i>). A browser window opens to sign in and approve; then copy the <b>Client ID</b> it shows.",
        "Unzip the kit, then in PowerShell go to that folder with <b>cd</b> (not C:\\Windows\\System32). Run the Navjyoti script, replacing <b>yourcompany</b> with your own SharePoint name and <b>PASTE-CLIENT-ID-HERE</b> with the Client ID from step 2.",
        "Open <b>Site contents</b> and check the item counts: RawMaterial 15, PreCompost 3, Beds 80, Harvest 80, QualityControl 6, Sales 6, StockLedger 63, Expenses 30, Batches 2, DailyLog 0.",
    ], [fig_powershell(), fig_site_contents()], "ps"),
    ("app", "Build the data-entry app in Power Apps", "Power user · 2–3 hours", [
        "Go to <b>make.powerapps.com</b> → <b>+ Create</b>.",
        "Choose <b>Start with data</b>.",
        "Pick <b>SharePoint</b> and your Navjyoti Production site.",
        "Tick <b>DailyLog</b> → <b>Create</b>. Power Apps builds a working phone app. Then add the other lists under <b>Data → + Add data</b>.",
        "Add the bed board: a gallery with the Items formula shown, coloured by status. All formulas are in <i>PowerApps-Formulas.md</i> in the kit.",
        "Supervisors tap a bed…",
        "…and a pop-up panel opens with its details and a Save button. Publish and share the app with the team.",
    ], [fig_pa_start(), fig_pa_studio(), fig_pa_popup()], "pa"),
    ("bi", "Build the reports in Power BI", "Power user · 3–4 hours", [
        "Open <b>Power BI Desktop</b> → <b>Get data</b>.",
        "Choose <b>SharePoint Online list</b>.",
        "Paste the site address.",
        "Select <b>Implementation 2.0</b>, tick all 10 lists, then <b>Transform data → Close &amp; apply</b>.",
        "In <b>Model view</b>, drag <b>Batches[BatchCode]</b> onto each list's <b>Batch</b> column. Paste the measures from <i>PowerBI-Measures.dax</i>.",
        "Build the pages with a <b>Batch slicer</b>: Overview, Raw material, Beds &amp; harvest, Quality, Sales &amp; stock, Compare batches.",
        "For the Quality page, put <b>Product</b> (RM / FG / EX) on the legend and add constant lines at the FCO limits. Turn on tooltips and add a drill-through page for pop-up detail.",
        "<b>Publish</b>. In app.powerbi.com open the semantic model <b>Settings</b> and set the credentials to OAuth2.",
        "Turn on <b>scheduled refresh</b> (up to 8 times a day).",
    ], [fig_pbi_getdata(), fig_pbi_model(), fig_pbi_report(), fig_pbi_refresh()], None),
    ("alerts", "Set up alerts in Power Automate", "Power user · 1 hour", [
        "<b>+ Create → Scheduled cloud flow</b>, every day at 08:00.",
        "Add <b>SharePoint – Get items</b> on <i>Beds</i> with the filter shown.",
        "Add a <b>Condition</b>: number of items greater than 0.",
        "Add <b>Teams – Post message</b> with the list of overdue beds. Repeat for watering gaps, FG/EX reports outside FCO and unpaid raw material (see <i>PowerAutomate-Flows.md</i>).",
    ], [fig_flow()], None),
    ("teams", "Put it together in Teams", "Anyone · 10 minutes", [
        "In the Navjyoti Production team, click <b>+</b> on the tab bar.",
        "Add <b>Power Apps</b> (the app), <b>Power BI</b> (the dashboard) and <b>Lists</b> (Beds, Harvest) as tabs.",
    ], [fig_teams()], None),
]

CODE = {
    "ps": [
        ("Install PnP PowerShell", "Install-Module PnP.PowerShell -Scope CurrentUser"),
        ("Register the app (admin, once)", 'Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Navjyoti" -Tenant yourcompany.onmicrosoft.com'),
        ("Go to the unzipped kit folder (example path)", 'cd "$env:USERPROFILE\\Downloads\\Navjyoti_M365_Kit"; Get-ChildItem -Recurse | Unblock-File'),
        ("Create the lists and load Batch 4", "./Create-NavjyotiLists.ps1 -SiteUrl https://yourcompany.sharepoint.com/sites/NavjyotiProduction -ClientId PASTE-CLIENT-ID-HERE"),
    ],
    "pa": [
        ("Bed board gallery · Items", 'SortByColumns(Filter(Beds, Batch.Value = varBatch), "BedNo")'),
        ("Bed board gallery · OnSelect (opens the pop-up)", "Set(varBed, ThisItem); Set(varShowDetail, true)"),
        ("Pop-up container · Visible", "varShowDetail"),
    ],
}


def build(out):
    toc = "".join(f'<a href="#{k}"><span>{i + 1}</span>{esc(t)}</a>' for i, (k, t, *_) in enumerate(STAGES))
    body = []
    for si, (key, title, who, steps, figs, code) in enumerate(STAGES):
        ol = "".join(f"<li><span>{s}</span></li>" for s in steps)
        codes = ""
        if code:
            codes = '<div class="codes">' + "".join(
                f'<div class="code"><div class="code-h"><span>{esc(l)}</span><button type="button" class="copy" data-copy="{esc(c)}">Copy</button></div><pre><code>{esc(c)}</code></pre></div>'
                for l, c in CODE[code]) + "</div>"
        body.append(f'''<section id="{key}" class="stage">
  <header class="stage-h"><span class="stage-n">Stage {si + 1}</span><h2>{esc(title)}</h2><span class="who">{esc(who)}</span></header>
  <div class="stage-b">
    <ol class="steps">{ol}</ol>
    <div class="figs">{"".join(f'<figure>{f}</figure>' for f in figs)}</div>
  </div>
  {codes}
</section>''')
    page = TEMPLATE.replace("{{TOC}}", toc).replace("{{OVERVIEW}}", fig_overview()).replace("{{BODY}}", "\n".join(body))
    with open(out, "w", encoding="utf-8") as f:
        f.write(page)


TEMPLATE = r'''<title>Navjyoti Microsoft 365 Setup</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600;12..96,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Layout: a single reading column; each stage pairs its numbered steps with drawings whose markers carry the same numbers. */
:root {
  --bg: #f2f5f1; --surface: #ffffff; --panel: #e9efe8; --card: #f7faf6; --fg: #16201b; --muted: #5b6b62; --line: #d3ddd4;
  --accent: #1c7148; --accent-fg: #ffffff; --soil: #6b4424; --hi: #e0791f;
  --c1: #dbe8fb; --c2: #fde3d6; --c3: #d8efe2; --c4: #efe7f8; --term: #14201a; --term-t: #d7e6dc; --term-ok: #7fd6a4;
  --f-display: "Bricolage Grotesque", "Segoe UI", system-ui, sans-serif;
  --f-body: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif;
  --f-mono: "IBM Plex Mono", ui-monospace, Menlo, Consolas, monospace;
}
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
  --bg: #0f1512; --surface: #161e1a; --panel: #1f2924; --card: #1a231e; --fg: #e2ebe5; --muted: #95a89c; --line: #2d3a33;
  --accent: #4dbf8a; --accent-fg: #07130d; --soil: #d19a6a; --hi: #f0a050;
  --c1: #1d2c44; --c2: #3b2419; --c3: #173524; --c4: #2a2240; --term: #0a100d; --term-t: #cfe0d5; --term-ok: #7fd6a4;
  color-scheme: dark } }
:root[data-theme="dark"] {
  --bg: #0f1512; --surface: #161e1a; --panel: #1f2924; --card: #1a231e; --fg: #e2ebe5; --muted: #95a89c; --line: #2d3a33;
  --accent: #4dbf8a; --accent-fg: #07130d; --soil: #d19a6a; --hi: #f0a050;
  --c1: #1d2c44; --c2: #3b2419; --c3: #173524; --c4: #2a2240; --term: #0a100d; --term-t: #cfe0d5; --term-ok: #7fd6a4;
  color-scheme: dark }
* { box-sizing: border-box }
body { background: var(--bg); color: var(--fg); font: 15px/1.6 var(--f-body); }
.wrap { max-width: 1120px; margin: 0 auto; padding-inline: 16px; padding-block: 28px 64px; display: grid; gap: 28px }
h1, h2 { font-family: var(--f-display); margin: 0; text-wrap: balance; letter-spacing: -0.01em }
h1 { font-size: clamp(28px, 4vw, 40px); line-height: 1.1 }
.lede { max-width: 68ch; color: var(--muted); margin: 8px 0 0 }
.eyebrow { font: 600 12px/1 var(--f-mono); letter-spacing: .08em; text-transform: uppercase; color: var(--soil) }
.note { background: var(--surface); border: 1px solid var(--line); border-left: 4px solid var(--hi); border-radius: 8px; padding: 12px 14px; max-width: 80ch }
.note b { color: var(--fg) }
nav.toc { display: flex; flex-wrap: wrap; gap: 8px }
nav.toc a { display: inline-flex; gap: 8px; align-items: center; text-decoration: none; color: var(--fg); background: var(--surface); border: 1px solid var(--line); border-radius: 999px; padding: 6px 14px 6px 6px; font-weight: 500; font-size: 14px }
nav.toc a span { display: inline-grid; place-items: center; width: 24px; height: 24px; border-radius: 50%; background: var(--accent); color: var(--accent-fg); font-size: 12px; font-weight: 700 }
nav.toc a:hover { border-color: var(--accent) }
.kit { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 10px }
.kit div { background: var(--surface); border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font-size: 13.5px; min-width: 0 }
.kit code { font-family: var(--f-mono); font-size: 12.5px; color: var(--soil); overflow-wrap: anywhere }
.stage { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 20px; display: grid; gap: 16px; scroll-margin-top: 16px }
.stage-h { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 14px; border-bottom: 1px solid var(--line); padding-bottom: 12px }
.stage-n { font: 600 12px/1 var(--f-mono); letter-spacing: .08em; text-transform: uppercase; color: var(--accent) }
.stage-h h2 { font-size: 22px }
.who { margin-left: auto; font-size: 13px; color: var(--muted); background: var(--panel); border-radius: 999px; padding: 2px 10px }
.stage-b { display: grid; grid-template-columns: minmax(0, 340px) minmax(0, 1fr); gap: 22px; align-items: start }
@media (max-width: 860px) { .stage-b { grid-template-columns: minmax(0, 1fr) } }
ol.steps { margin: 0; padding: 0; list-style: none; counter-reset: s; display: grid; gap: 12px }
ol.steps li { counter-increment: s; display: grid; grid-template-columns: 28px 1fr; gap: 10px; align-items: start }
ol.steps li::before { content: counter(s); display: grid; place-items: center; width: 26px; height: 26px; border-radius: 50%; background: var(--hi); color: #fff; font: 700 13px/1 var(--f-body); margin-top: 1px }
.figs { display: grid; gap: 14px; min-width: 0 }
figure { margin: 0; min-width: 0 }
.fig-scroll { overflow-x: auto; border-radius: 10px }
.fig-scroll svg { display: block; width: 100%; min-width: 560px; height: auto }
.codes { display: grid; gap: 10px }
.code { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; min-width: 0 }
.code-h { display: flex; justify-content: space-between; align-items: center; gap: 8px; background: var(--panel); padding: 6px 10px; font-size: 13px; font-weight: 600 }
.code pre { margin: 0; padding: 10px 12px; overflow-x: auto; background: var(--term); color: var(--term-t); font: 13px/1.5 var(--f-mono) }
.copy { border: 1px solid var(--line); background: var(--surface); color: var(--fg); border-radius: 6px; padding: 3px 10px; font: 600 12px var(--f-body); cursor: pointer }
.copy:focus-visible, nav.toc a:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px }
.foot { color: var(--muted); font-size: 13px; max-width: 80ch }
/* figure drawing classes */
.f-win { fill: var(--surface); stroke: var(--line) } .f-bar { fill: var(--panel) } .f-dot { fill: var(--line) }
.f-panel { fill: var(--panel) } .f-card { fill: var(--card); stroke: var(--line) } .f-input { fill: var(--surface); stroke: var(--line) }
.f-row { fill: var(--card) } .f-head { fill: var(--accent) }
.f-btn { fill: var(--surface); stroke: var(--line) } .f-btn-p { fill: var(--accent) }
.f-c1 { fill: var(--c1) } .f-c2 { fill: var(--c2) } .f-c3 { fill: var(--c3) } .f-c4 { fill: var(--c4) }
.f-phone { fill: var(--card); stroke: var(--muted); stroke-width: 2 } .f-dim { fill: #000; opacity: .35 }
.f-term { fill: var(--term) } .f-term-t { fill: var(--term-t) } .f-term-ok { fill: var(--term-ok) }
.f-t { fill: var(--fg); font-family: var(--f-body) } .f-t-inv { fill: var(--accent-fg); font-family: var(--f-body) } .f-muted { fill: var(--muted); font-family: var(--f-body) }
.f-code { fill: var(--fg) } .f-mono { font-family: var(--f-mono) }
.f-line { stroke: var(--muted); stroke-width: 1.2 } .f-rel { stroke: var(--accent); stroke-width: 1.4; opacity: .6 } .f-ref { stroke: var(--hi); stroke-width: 1.5; stroke-dasharray: 5 4 }
.f-arrow { stroke: var(--muted); stroke-width: 2 } .f-arrowhead { fill: var(--muted) }
.f-hi { fill: none; stroke: var(--hi); stroke-width: 2.5; stroke-dasharray: 6 4 } .f-badge { fill: var(--hi) }
@media (prefers-reduced-motion: reduce) { * { transition: none !important } }
</style>

<div class="wrap">
  <header>
    <div class="eyebrow">Navjyoti · Vermicompost production</div>
    <h1>Setting up the tracker on Microsoft 365</h1>
    <p class="lede">SharePoint holds the registers, Power Apps is the phone app for supervisors, Power BI gives the live reports, and Power Automate sends alerts. Everyone uses their company login. Allow about 1–2 days in total.</p>
  </header>
  <figure>{{OVERVIEW}}</figure>
  <div class="note"><b>About the pictures:</b> they are simplified drawings of each screen, not screenshots. Microsoft changes button positions from time to time, so look for the same words if a screen looks a little different. The orange numbered markers match the numbered steps.</div>
  <nav class="toc" aria-label="Stages">{{TOC}}</nav>
  <section class="stage" aria-label="What is in the kit">
    <header class="stage-h"><span class="stage-n">Before you start</span><h2>What is in the Navjyoti kit</h2><span class="who">Navjyoti_M365_Kit.zip</span></header>
    <div class="kit">
      <div><code>Create-NavjyotiLists.ps1</code><br>Creates the 10 lists with the right column types and loads Batch 4.</div>
      <div><code>data/*.csv</code><br>Batch 4 records, one file per list.</div>
      <div><code>PowerApps-Formulas.md</code><br>Screens and formulas, including pop-up panels.</div>
      <div><code>PowerBI-Measures.dax</code><br>Calculated columns and measures for the reports.</div>
      <div><code>PowerAutomate-Flows.md</code><br>The four alert flows, click by click.</div>
      <div><code>README.md</code><br>The same steps as this page, as text.</div>
    </div>
    <p class="foot">Licences: most business plans (Business Standard or Premium, E3, E5) include SharePoint, Power Apps and Power Automate for SharePoint data. Sharing Power BI reports needs Power BI Pro for each viewer (included in E5). Check with your IT admin.</p>
  </section>
  {{BODY}}
  <section class="stage" id="golive">
    <header class="stage-h"><span class="stage-n">Stage 7</span><h2>Go live with Batch 5</h2><span class="who">Whole team · 1–2 weeks</span></header>
    <ol class="steps">
      <li><span>Fill in Batch 5's start date and stage in the <b>Batches</b> list.</span></li>
      <li><span>From now on, supervisors enter daily logs, beds, harvests, lab reports and sales in the app.</span></li>
      <li><span>Keep the Excel tracker running alongside for 1–2 weeks and compare the totals with Power BI.</span></li>
      <li><span>Once they match, use Microsoft 365 as the main system.</span></li>
    </ol>
  </section>
</div>
<script>
document.addEventListener('click', async e => {
  const b = e.target.closest('.copy'); if (!b) return;
  const text = b.dataset.copy;
  try { await navigator.clipboard.writeText(text); b.textContent = 'Copied'; }
  catch { const pre = b.closest('.code').querySelector('code'); const r = document.createRange(); r.selectNodeContents(pre); const s = getSelection(); s.removeAllRanges(); s.addRange(r); b.textContent = 'Selected, press Ctrl+C'; }
  setTimeout(() => b.textContent = 'Copy', 2000);
});
</script>
'''

if __name__ == "__main__":
    build(os.path.join(os.path.dirname(os.path.abspath(__file__)), "guide.html"))
    print("guide.html written")
