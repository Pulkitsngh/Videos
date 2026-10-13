"""Build the IT deployment guide for the Navjyoti Tracker (Microsoft 365 edition).

    python3 build_it_guide.py   -> it_guide.html (print it to PDF with Chromium; see README)

Screens of the tracker itself (img/app-*.png) are real screenshots of the built app running against a
test copy of the SharePoint lists. Microsoft portal screens are simplified drawings: Microsoft moves
buttons from time to time, so the guide names the labels to look for. Orange numbers match the steps.
"""
import base64, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.dirname(HERE)
ROOT = os.path.dirname(APP)
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "m365"))
from build_guide import Fig, esc, W  # noqa: E402  (same drawing helpers as the Microsoft 365 guide)
from build_cloud_workbook import registers  # noqa: E402
from build_m365_kit import field_type  # noqa: E402

VERSION = open(os.path.join(APP, "VERSION")).read().strip()
SITE = "https://lnbgroup.sharepoint.com/sites/NavjyotiProduction"
APPURL = "https://navjyoti-tracker.azurestaticapps.net/"


def img(name, alt, cls="shot"):
    data = base64.b64encode(open(os.path.join(HERE, "img", name), "rb").read()).decode()
    return f'<img class="{cls}" src="data:image/png;base64,{data}" alt="{esc(alt)}">'


def logo():
    return open(os.path.join(ROOT, "assets", "navjyoti_logo.svg")).read().replace('role="img"', 'role="img" class="logo"')


# ---------------------------------------------------------------- drawings
def fig_architecture():
    f = Fig("How the pieces fit together", 430)
    # PCs
    f.rect(20, 55, 190, 300, "f-c3", 10).text(115, 80, "LNB Group PCs and phones", 13, "f-t", "middle", 700)
    for i, y in enumerate((95, 175)):
        f.rect(42, y, 146, 62, "f-card", 6).rect(42, y, 146, 14, "f-head", 6).text(115, y + 11, "Navjyoti Tracker", 9.5, "f-t-inv", "middle", 600)
        f.text(115, y + 40, "installed app", 11, "f-muted", "middle")
    f.rect(92, 258, 46, 78, "f-phone", 8).text(115, 300, "app", 10, "f-muted", "middle")
    f.text(115, 350, "sign in with work account", 10.5, "f-muted", "middle")
    # cloud boundary
    f.rect(250, 45, 450, 340, "f-panel", 14).text(475, 66, "LNB Group Microsoft 365 / Azure tenant", 12, "f-muted", "middle", 700)
    f.rect(270, 82, 190, 70, "f-c1", 8).text(365, 108, "Microsoft Entra ID", 13, "f-t", "middle", 700).text(365, 128, "sign-in · MFA · access", 11, "f-muted", "middle")
    f.rect(270, 172, 190, 78, "f-c4", 8).text(365, 198, "App hosting", 13, "f-t", "middle", 700).text(365, 217, "Azure Static Web Apps", 11, "f-muted", "middle").text(365, 233, "or internal IIS (HTTPS)", 11, "f-muted", "middle")
    f.rect(270, 270, 190, 70, "f-c2", 8).text(365, 296, "Microsoft Graph", 13, "f-t", "middle", 700).text(365, 316, "secure API to SharePoint", 11, "f-muted", "middle")
    f.rect(490, 150, 190, 190, "f-c3", 8).text(585, 176, "SharePoint site", 13, "f-t", "middle", 700).text(585, 194, "Navjyoti Production", 11, "f-muted", "middle")
    for j, n in enumerate(["Batches · RawMaterial", "PreCompost · Beds", "Harvest · QualityControl", "Sales · Expenses", "Earthworm"]):
        f.rect(505, 206 + j * 25, 160, 20, "f-card", 4).text(585, 220 + j * 25, n, 10, "f-t", "middle")
    f.arrow(210, 120, 268, 117).text(239, 108, "1", 11, "f-t", "middle", 700)
    f.arrow(210, 205, 268, 210).text(239, 198, "2", 11, "f-t", "middle", 700)
    f.arrow(210, 290, 268, 300).text(239, 285, "3", 11, "f-t", "middle", 700)
    f.arrow(460, 300, 488, 280)
    f.text(20, 404, "1 = sign in  ·  2 = load the app (static files only)  ·  3 = read and write records with each person's own SharePoint rights", 11.5, "f-t")
    f.text(20, 422, "No Claude, no Anthropic service and no other outside database is involved.", 11.5, "f-t", weight=700)
    return f.svg("Architecture: PCs sign in with Entra ID, load the app from the host, and read and write SharePoint lists through Microsoft Graph")


def fig_powershell():
    f = Fig("PowerShell 7 — admin PC", 330)
    f.rect(12, 44, W - 24, 274, "f-term", 6)
    lines = [("PS> Install-Module PnP.PowerShell -Scope CurrentUser", 0), ("PS> Register-PnPEntraIDAppForInteractiveLogin -ApplicationName \"PnP Navjyoti\" -Tenant lnbgroup.onmicrosoft.com", 1),
             ("    AzureAppId/ClientId : 0f3c…   (copy this)", 0), ("PS> cd \"$env:USERPROFILE\\Downloads\\Navjyoti_M365_Kit\"; Get-ChildItem -Recurse | Unblock-File", 0),
             (f"PS> ./Create-NavjyotiLists.ps1 -SiteUrl {SITE} -ClientId 0f3c…", 2),
             ("Created list Batches", 3), ("Created list RawMaterial", 3), ("  … PreCompost, Beds, Harvest, QualityControl, Sales, Expenses, Earthworm", 3),
             ("Loaded 15 rows into RawMaterial   ·   80 into Beds   ·   80 into Harvest …", 3), ("Done. 9 lists ready.", 3)]
    y = 70
    for s, n in lines:
        f.mono(26, y, s, 11.5, "f-term-ok" if n == 3 else "f-term-t")
        if n in (1, 2):
            f.hi(22, y - 14, 600 if n == 1 else 640, 19, n)
        y += 25
    return f.svg("PowerShell: install PnP, register the PnP app once, run Create-NavjyotiLists.ps1, nine lists created")


def fig_entra_register():
    f = Fig("Microsoft Entra admin center · App registrations · Register an application", 420)
    f.rect(20, 48, 160, 360, "f-panel", 8)
    for i, s in enumerate(["Home", "Identity", "  Applications", "    App registrations", "    Enterprise apps", "Protection"]):
        f.text(32, 74 + i * 26, s, 11.5, "f-t", weight=700 if "App reg" in s else 400)
    f.text(200, 74, "Register an application", 16, weight=700)
    f.text(200, 102, "Name", 12, "f-muted", weight=600).rect(200, 108, 330, 28, "f-input", 5).text(210, 127, "Navjyoti Tracker", 12).hi(200, 108, 330, 28, 1)
    f.text(200, 160, "Supported account types", 12, "f-muted", weight=600)
    f.parts.append('<circle cx="208" cy="181" r="6" class="f-head"/>')
    f.text(222, 185, "Accounts in this organizational directory only (LNB Group only – Single tenant)", 11).hi(200, 170, 480, 22, 2)
    f.parts.append('<circle cx="208" cy="205" r="6" class="f-input"/>')
    f.text(222, 209, "Accounts in any organizational directory (multitenant)", 11.5, "f-muted")
    f.text(200, 244, "Redirect URI", 12, "f-muted", weight=600)
    f.rect(200, 252, 196, 28, "f-input", 5).text(208, 271, "Single-page application (SPA)", 10.5).text(384, 271, "▾", 11, "f-muted")
    f.rect(404, 252, 296, 28, "f-input", 5).mono(412, 271, APPURL, 10).hi(200, 252, 500, 28, 3)
    f.text(200, 305, "Use the app's address with the slash at the end. You can add more addresses later.", 11, "f-muted")
    f.btn(200, 350, 100, 32, "Register", True).hi(200, 350, 100, 32, 4)
    return f.svg("Entra: new app registration named Navjyoti Tracker, single tenant, single-page application redirect URI")


def fig_entra_ids():
    f = Fig("Navjyoti Tracker · Overview", 250)
    f.text(30, 74, "Navjyoti Tracker", 16, weight=700)
    rows = [("Display name", "Navjyoti Tracker"), ("Application (client) ID", "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"), ("Object ID", "…"),
            ("Directory (tenant) ID", "11111111-2222-3333-4444-555555555555"), ("Supported account types", "My organization only")]
    for i, (k, v) in enumerate(rows):
        y = 104 + i * 26
        f.text(30, y, k, 12, "f-muted").mono(250, y, v, 12)
    f.hi(245, 104 + 26 - 16, 330, 22, 1).hi(245, 104 + 3 * 26 - 16, 330, 22, 2)
    return f.svg("Entra overview page showing the Application (client) ID and the Directory (tenant) ID to copy")


def fig_entra_perms():
    f = Fig("Navjyoti Tracker · API permissions", 330)
    f.btn(30, 52, 150, 30, "+ Add a permission").hi(30, 52, 150, 30, 1)
    f.btn(195, 52, 250, 30, "✓ Grant admin consent for LNB Group", True).hi(195, 52, 250, 30, 4)
    f.rect(30, 100, 660, 28, "f-head", 4)
    for x, h in ((42, "API / Permission name"), (300, "Type"), (400, "Description"), (590, "Status")):
        f.text(x, 119, h, 11.5, "f-t-inv", weight=700)
    for i, (n, d) in enumerate([("Sites.ReadWrite.All", "Edit or delete items in all site collections"), ("User.Read", "Sign in and read user profile")]):
        y = 132 + i * 30
        f.rect(30, y, 660, 28, "f-row", 0).text(42, y + 19, "Microsoft Graph · " + n, 11.5, weight=600).text(300, y + 19, "Delegated", 11.5).text(400, y + 19, d, 11, "f-muted")
        f.text(590, y + 19, "✓ Granted for LNB Group", 11, "f-t", weight=600)
    f.hi(30, 132, 660, 28, 3)
    f.rect(30, 210, 660, 100, "f-card", 8).text(46, 236, "Request API permissions → Microsoft Graph → Delegated permissions", 12, weight=700).hi(40, 220, 420, 24, 2)
    f.rect(46, 250, 230, 26, "f-input", 4).text(56, 268, "Sites.ReadWrite.All", 12)
    f.text(46, 298, "Delegated = the app works with each person's own SharePoint rights, never more.", 11.5, "f-muted")
    return f.svg("API permissions: Microsoft Graph delegated Sites.ReadWrite.All and User.Read, admin consent granted")


def fig_swa_create():
    f = Fig("Azure portal · Create Static Web App", 400)
    f.text(30, 74, "Create Static Web App", 16, weight=700)
    rows = [("Subscription", "LNB Group – Production"), ("Resource group", "rg-navjyoti"), ("Name", "navjyoti-tracker"), ("Plan type", "Free  (for hobby or personal projects)"),
            ("Region for Azure Functions", "East Asia (nearest available)"), ("Deployment source", "Other")]
    for i, (k, v) in enumerate(rows):
        y = 100 + i * 42
        f.text(30, y + 18, k, 12, "f-muted", weight=600).rect(250, y, 300, 28, "f-input", 5).text(260, y + 19, v, 12)
    f.hi(250, 100 + 2 * 42, 300, 28, 1).hi(250, 100 + 3 * 42, 300, 28, 2).hi(250, 100 + 5 * 42, 300, 28, 3)
    f.btn(30, 360, 130, 30, "Review + create", True).hi(30, 360, 130, 30, 4)
    return f.svg("Azure portal: create a Static Web App on the Free plan with deployment source Other")


def fig_swa_overview():
    f = Fig("navjyoti-tracker · Static Web App · Overview", 250)
    f.text(30, 74, "navjyoti-tracker", 16, weight=700)
    f.btn(30, 90, 210, 28, "Manage deployment token").hi(30, 90, 210, 28, 2)
    f.btn(255, 90, 120, 28, "Browse")
    for i, (k, v) in enumerate([("URL", APPURL), ("Plan type", "Free"), ("Source", "Other"), ("Custom domains", "Add navjyoti.lnbgroup.com (optional)")]):
        y = 150 + i * 24
        f.text(30, y, k, 12, "f-muted").mono(200, y, v, 12)
    f.hi(195, 136, 330, 20, 1)
    return f.svg("Static Web App overview: the app URL and the Manage deployment token button")


def fig_edge_install():
    f = Fig("Microsoft Edge", 360)
    f.rect(12, 44, W - 24, 30, "f-input", 15).mono(30, 64, APPURL, 12)
    f.rect(620, 48, 22, 22, "f-c3", 4).text(631, 64, "⊕", 13, "f-t", "middle", 700).hi(620, 48, 22, 22, 1)
    f.rect(12, 82, W - 24, 266, "f-panel", 6).text(40, 120, "Navjyoti tracker (signed in)", 13, "f-muted")
    f.rect(360, 84, 330, 150, "f-card", 10).text(380, 112, "App available", 14, weight=700)
    f.text(380, 136, "Install Navjyoti Vermicompost Tracker", 12).text(380, 156, "Opens in its own window, with a Start menu", 11, "f-muted").text(380, 172, "and taskbar icon.", 11, "f-muted")
    f.btn(380, 190, 90, 30, "Install", True).btn(480, 190, 90, 30, "Not now").hi(380, 190, 90, 30, 2)
    f.rect(40, 250, 300, 80, "f-card", 10).text(56, 276, "Or: ⋯ menu → Apps →", 12).text(56, 296, "Install this site as an app", 12, weight=700).hi(52, 282, 200, 20, 3)
    return f.svg("Edge: the App available install button in the address bar opens the Install dialog")


def fig_intune():
    f = Fig("Microsoft Intune admin center · Configuration · Settings catalog", 360)
    f.text(30, 74, "Microsoft Edge", 15, weight=700).text(30, 96, "Configure list of force-installed Web Apps", 12.5, weight=600).hi(28, 82, 300, 20, 1)
    f.rect(30, 110, 660, 150, "f-term", 6)
    js = ['[', '  {', f'    "url": "{APPURL}",', '    "default_launch_container": "window",', '    "create_desktop_shortcut": true', '  }', ']']
    for i, s in enumerate(js):
        f.mono(46, 134 + i * 18, s, 12, "f-term-t")
    f.hi(30, 110, 660, 150, 2)
    f.text(30, 290, "Assignments → Include groups → Navjyoti Production members", 12).hi(28, 276, 390, 20, 3)
    f.text(30, 320, "Group Policy alternative: Computer Configuration → Administrative Templates → Microsoft Edge →", 11.5, "f-muted")
    f.text(30, 338, "“Configure list of force-installed Web Apps” with the same JSON.", 11.5, "f-muted")
    return f.svg("Intune settings catalog: Edge policy Configure list of force-installed Web Apps with the app URL")


def fig_sp_permissions():
    f = Fig("Navjyoti Production · Settings · Site permissions", 300)
    f.rect(400, 44, 308, 244, "f-card", 0).text(420, 74, "Site permissions", 15, weight=700)
    f.btn(420, 88, 120, 28, "Add members", True).hi(420, 88, 120, 28, 1)
    for i, (g, r, who) in enumerate([("Site owners", "Full control", "IT, production head"), ("Site members", "Edit", "supervisors, data entry"), ("Site visitors", "Read", "managers, auditors")]):
        y = 136 + i * 48
        f.text(420, y, g, 12.5, weight=700).text(600, y, r, 12, "f-t", weight=600).text(420, y + 18, who, 11, "f-muted")
    f.hi(416, 172, 280, 40, 2).hi(416, 220, 280, 40, 3)
    f.rect(12, 44, 380, 244, "f-panel", 0).text(30, 80, "Edit → can add, change and delete records", 12).text(30, 104, "Read → sees everything, cannot save", 12)
    f.text(30, 128, "(the app says “Your access is view-only”)", 11.5, "f-muted")
    f.text(30, 170, "Removing someone from the group ends", 12).text(30, 190, "their access straight away.", 12)
    return f.svg("SharePoint site permissions: owners full control, members edit, visitors read")


def fig_release():
    f = Fig("Releasing a new version", 300)
    steps = [("1", "Change", "index.html or", "src/cloud.js", "f-c1"), ("2", "Raise VERSION", "1.0.0 → 1.1.0", "", "f-c4"), ("3", "Build", "build_desktop_app", "→ dist/ + zip", "f-c2"),
             ("4", "Deploy", "git push (automatic)", "or SWA CLI / IIS copy", "f-c3"), ("5", "PCs update", "banner shows", "“Update now”", "f-c3")]
    for i, (n, t, a, b, c) in enumerate(steps):
        x = 20 + i * 140
        f.rect(x, 70, 124, 120, c, 10).text(x + 62, 100, t, 13, "f-t", "middle", 700).text(x + 62, 126, a, 10.5, "f-muted", "middle").text(x + 62, 143, b, 10.5, "f-muted", "middle")
        f.parts.append(f'<circle cx="{x + 14}" cy="{70}" r="11" class="f-badge"/><text x="{x + 14}" y="74.5" font-size="12" font-weight="700" text-anchor="middle" class="f-t-inv">{n}</text>')
        if i < 4:
            f.arrow(x + 126, 130, x + 138, 130)
    f.text(20, 228, "Records are never touched by a release: they live in SharePoint. Only the screens and logic change.", 12, "f-t", weight=600)
    f.text(20, 252, "If a release adds a column, run Create-NavjyotiLists.ps1 again before step 4 (it only adds what is missing).", 11.5, "f-muted")
    f.text(20, 274, "To roll back, deploy the previous release's zip again.", 11.5, "f-muted")
    return f.svg("Release flow: change, raise VERSION, build, deploy, PCs show the update prompt")


# ---------------------------------------------------------------- content
def lists_appendix():
    rows = []
    for t, (_, cols) in registers().items():
        names = [f"{h} <span>({field_type(t, h, kind, extra) or 'app only'})</span>" for h, k, kind, extra in cols if field_type(t, h, kind, extra)]
        rows.append(f"<tr><td><b>{esc(t)}</b></td><td>{', '.join(names)}</td></tr>")
    return "".join(rows)


def section(n, title, who, body):
    return f'<section class="stage"><header><span class="n">{esc(n)}</span><h2>{esc(title)}</h2><span class="who">{esc(who)}</span></header>{body}</section>'


def steps(*items):
    return "<ol class='steps'>" + "".join(f"<li><span>{i}</span></li>" for i in items) + "</ol>"


def code(label, text):
    return f'<div class="code"><div class="code-h">{esc(label)}</div><pre>{esc(text)}</pre></div>'


def build():
    V = esc(VERSION)
    cfg_example = open(os.path.join(APP, "config.example.js")).read()
    cfg_filled = (cfg_example.replace("PASTE-TENANT-ID-HERE", "11111111-2222-3333-4444-555555555555")
                  .replace("PASTE-CLIENT-ID-HERE", "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee").replace("https://YOUR-TENANT.sharepoint.com", "https://lnbgroup.sharepoint.com"))
    ps_text = ("Install-Module PnP.PowerShell -Scope CurrentUser\n"
               "Register-PnPEntraIDAppForInteractiveLogin -ApplicationName \"PnP Navjyoti\" -Tenant lnbgroup.onmicrosoft.com\n"
               "cd \"$env:USERPROFILE\\Downloads\\Navjyoti_M365_Kit\"; Get-ChildItem -Recurse | Unblock-File\n"
               "./Create-NavjyotiLists.ps1 -SiteUrl " + SITE + " -ClientId PASTE-PNP-CLIENT-ID")
    swa_cmd = "npx @azure/static-web-apps-cli deploy .\\Navjyoti_Tracker_App_v" + V + " --deployment-token &lt;token&gt; --env production"
    fig = lambda f: "<figure>" + f + "</figure>"
    two = lambda a, b: '<div class="cols">' + a + b + "</div>"
    note = lambda t: '<div class="note">' + t + "</div>"
    parts = []

    parts.append(f"""
<section class="cover">
  <div class="cover-logo">{logo()}<span>(LNB Group)</span></div>
  <div class="eyebrow">IT deployment and update guide</div>
  <h1>Navjyoti Vermicompost Tracker</h1>
  <p class="lede">Microsoft 365 edition · version {V} · October 2026</p>
  <div class="summary">
    <div><b>What it is</b><span>The production tracker, installed on PCs as an app. Everyone signs in with their LNB Group work account and works on the same records, which are stored in a SharePoint site in the LNB Group tenant.</span></div>
    <div><b>Who it is for</b><span>The IT administrator who sets it up, and whoever releases new versions later.</span></div>
    <div><b>Time needed</b><span>About half a day: SharePoint 30 min, Entra 15 min, hosting 30 min, PC installs a few minutes each (or automatic through Intune), checks 15 min.</span></div>
    <div><b>Claude</b><span>Not needed. The app runs entirely on LNB Group's Microsoft 365 and hosting. Claude was only used to write it.</span></div>
  </div>
  <div class="toc">
    <b>Contents</b>
    <ol>
      <li>What you are deploying</li><li>Before you start</li><li>Stage 1 · SharePoint site and lists</li><li>Stage 2 · Register the app in Microsoft Entra ID</li>
      <li>Stage 3 · Fill in config.js</li><li>Stage 4 · Host the app</li><li>Stage 5 · Install on PCs and phones</li><li>Stage 6 · Users and permissions</li>
      <li>Stage 7 · Acceptance checks</li><li>Stage 8 · Releasing updates</li><li>Running it: backup, security, monitoring</li><li>Troubleshooting</li><li>Appendices: package files, settings, lists and columns</li>
    </ol>
  </div>
</section>""")

    shot_overview = ('<figure class="frame"><div class="win"><span></span><span></span><span></span><b>Navjyoti Vermicompost Tracker</b></div>'
                     + img("app-overview.png", "The installed tracker, Overview page")
                     + "<figcaption>The installed tracker (real screenshot), showing Batch 4 records read from SharePoint.</figcaption></figure>")
    parts.append(section("1", "What you are deploying", "Overview",
        "<p>The tracker is a <b>progressive web app</b>: a set of static files (HTML, JavaScript, icons) that Microsoft Edge or Google Chrome installs like a desktop program, "
        "with its own window, Start-menu entry and taskbar icon. It has no server code and no database of its own.</p>"
        + fig(fig_architecture()) + f"""
<table class="t"><thead><tr><th>Part</th><th>What it is</th><th>Where it lives</th></tr></thead><tbody>
<tr><td>App package</td><td><code>Navjyoti_Tracker_App_v{V}.zip</code>: the screens, the Microsoft 365 connector, icons</td><td>Azure Static Web Apps (Free plan) or an internal IIS server, over HTTPS</td></tr>
<tr><td>Sign-in</td><td>App registration “Navjyoti Tracker” (single-page application, no client secret)</td><td>Microsoft Entra ID, LNB Group tenant</td></tr>
<tr><td>Data</td><td>9 SharePoint lists: Batches, RawMaterial, PreCompost, Beds, Harvest, QualityControl, Sales, Expenses, Earthworm</td><td>SharePoint site “Navjyoti Production”</td></tr>
<tr><td>Access rights</td><td>SharePoint site permissions: members edit, visitors read</td><td>The site's Microsoft 365 group</td></tr>
<tr><td>Installs</td><td>Edge or Chrome “Install app”, or forced install through Intune / Group Policy</td><td>Each PC; phones through “Add to home screen”</td></tr>
</tbody></table>"""
        + two(note("<b>Same data for everyone.</b> Each open app reads the lists every 30 seconds (setting <code>pollSeconds</code>) and whenever its window comes to the front, so an entry made on one PC shows on the others within half a minute."),
              note("<b>Security model.</b> The app uses <i>delegated</i> Microsoft Graph permissions: it can only do what the signed-in person may do in SharePoint. MFA and Conditional Access apply as for any Microsoft 365 app. The package holds no passwords or secrets."))
        + '<p class="small">Outside services: none, apart from Google Fonts for the typefaces. If your policy blocks Google Fonts the app falls back to Windows fonts and works the same.</p>'
        + shot_overview))

    parts.append(section("2", "Before you start", "Checklist", f"""
<table class="t"><thead><tr><th>You need</th><th>Notes</th></tr></thead><tbody>
<tr><td>An admin account with <b>Global Administrator</b>, or <b>Application Administrator + SharePoint Administrator</b></td><td>To register the app, grant admin consent and create the site</td></tr>
<tr><td>Somewhere to host the files over <b>HTTPS</b></td><td>Azure Static Web Apps, Free plan (recommended, needs an Azure subscription), or an internal IIS server with a certificate the PCs trust. Sign-in and app install do not work over plain HTTP.</td></tr>
<tr><td>A Windows PC with <b>PowerShell 7</b> and the <b>PnP.PowerShell</b> module</td><td>To create the SharePoint lists (Stage 1)</td></tr>
<tr><td><b>Node.js LTS</b> on the admin PC</td><td>Only for Option A deployment with the Static Web Apps command-line tool</td></tr>
<tr><td><b>Python 3</b> with <code>openpyxl</code></td><td>Only to rebuild the package after changes (Stage 8). Not needed for the first deployment.</td></tr>
<tr><td>Microsoft Edge or Google Chrome on the PCs</td><td>Both install the app; Edge can be managed with Intune</td></tr>
</tbody></table>
<h3>Files you receive</h3>
<table class="t"><tbody>
<tr><td><code>Navjyoti_Tracker_App_v{V}.zip</code></td><td>The app, ready to host once <code>config.js</code> is filled in</td></tr>
<tr><td><code>Navjyoti_M365_Kit.zip</code></td><td><code>Create-NavjyotiLists.ps1</code> and the Batch 4 data it loads</td></tr>
<tr><td>Source code (GitHub repository, folder <code>navjyoti-vermicompost</code>)</td><td>Needed for later changes; <code>desktop-app/</code> holds the build script</td></tr>
</tbody></table>
<h3>Decide these values first</h3>
<table class="t fill"><tbody>
<tr><td>Tenant name</td><td><code>lnbgroup.onmicrosoft.com</code></td><td class="blank"></td></tr>
<tr><td>SharePoint site address</td><td><code>{SITE}</code></td><td class="blank"></td></tr>
<tr><td>App address</td><td><code>{APPURL}</code> (given by Azure) or e.g. <code>https://navjyoti.lnbgroup.com/</code></td><td class="blank"></td></tr>
<tr><td>Who edits / who only views</td><td>Supervisors and data entry edit; managers view</td><td class="blank"></td></tr>
</tbody></table>
<p class="small">The addresses in this guide are examples. Write your real values in the empty column.</p>"""))

    parts.append(section("Stage 1", "SharePoint site and lists", "SharePoint admin · 30 min",
        steps("Create the site: <b>sharepoint.com → + Create site → Team site</b>, name <b>Navjyoti Production</b>. If the team already exists in Teams, use its site.",
              "Set the site time zone: <b>Settings (gear) → Site information → View all site settings → Regional settings → Time zone: (UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi</b>. Dates are stored as days in this time zone; PCs should use the same.",
              "On the admin PC, open PowerShell 7 and install PnP: <code>Install-Module PnP.PowerShell -Scope CurrentUser</code>.",
              "Once per tenant, register the PnP sign-in app (marker 1 in the picture) and copy the client ID it prints. This app is only for the script, not for the tracker.",
              "Unzip <code>Navjyoti_M365_Kit.zip</code>, go to the folder, unblock the files, and run <code>Create-NavjyotiLists.ps1</code> (marker 2).",
              "Check <b>Site contents</b>: Batches 2, RawMaterial 15, PreCompost 3, Beds 80, Harvest 80, QualityControl 6, Sales 6, Expenses 30, Earthworm 0.")
        + fig(fig_powershell()) + code("PowerShell 7", ps_text)
        + note("<b>Safe to run again.</b> The script keeps existing lists, columns and rows, adds anything missing, and never loads data into a list that already has rows. Add <code>-SkipData</code> to create empty lists. Lists left from an earlier kit (DailyLog, StockLedger) are not used by the app and can be deleted.")))

    parts.append(section("Stage 2", "Register the app in Microsoft Entra ID", "Entra admin · 15 min",
        steps("Open <b>entra.microsoft.com → Identity → Applications → App registrations → + New registration</b>.",
              "Name: <b>Navjyoti Tracker</b> (marker 1). Supported account types: <b>Accounts in this organizational directory only</b> (marker 2).",
              "Redirect URI: platform <b>Single-page application (SPA)</b>, address = the app address with the slash at the end (marker 3). If you do not know it yet, leave it empty and add it after Stage 4.",
              "Click <b>Register</b> (marker 4). On the overview page copy the <b>Application (client) ID</b> and the <b>Directory (tenant) ID</b>.",
              "<b>API permissions → + Add a permission → Microsoft Graph → Delegated permissions →</b> tick <b>Sites.ReadWrite.All</b> → Add. <b>User.Read</b> is already there.",
              "Click <b>Grant admin consent for LNB Group</b> and confirm. Both rows show “Granted”.",
              "<b>Authentication</b>: add any other addresses the app will use (a custom domain, or <code>http://localhost:8080/</code> for testing). Leave “Implicit grant” unticked; the app uses the authorization-code flow with PKCE.",
              "Optional, recommended: <b>Enterprise applications → Navjyoti Tracker → Properties → Assignment required: Yes</b>, then <b>Users and groups</b> → add the Navjyoti Production group. Only those people can then sign in.")
        + fig(fig_entra_register()) + two(fig(fig_entra_ids()), fig(fig_entra_perms()))
        + note("<b>Stricter option.</b> Sites.ReadWrite.All (delegated) already limits the app to what each person may do. If your policy requires the app to reach one site only, use the Sites.Selected permission instead, grant the app write access to the Navjyoti site (for example with PnP's <code>Grant-PnPEntraIDAppSitePermission</code>), and set <code>scopes</code> in <code>config.js</code> as shown in Appendix B.")))

    parts.append(section("Stage 3", "Fill in config.js", "IT · 10 min",
        steps("Unzip <code>Navjyoti_Tracker_App_v" + V + ".zip</code> into a folder.",
              "Open <code>config.js</code> in Notepad and fill in the tenant ID, the application (client) ID and the SharePoint site address. Save.",
              "If you will rebuild the app later from the source code, put the same file at <code>navjyoti-vermicompost/desktop-app/config.js</code>; every future build then includes it.")
        + code("config.js (example values)", cfg_filled)
        + "<p>These values are not secrets. They only tell the app which tenant, registration and site to use; access is controlled by sign-in and SharePoint permissions. If a value is missing, the app shows this message instead of loading:</p>"
        + img("app-setup.png", "Setup not finished message")))

    parts.append(section("Stage 4", "Host the app", "Azure or IIS admin · 30 min",
        "<h3>Option A · Azure Static Web Apps (recommended)</h3>"
        + steps("<b>portal.azure.com → Create a resource → Static Web App</b>.",
                "Resource group <b>rg-navjyoti</b>, name <b>navjyoti-tracker</b> (marker 1), plan <b>Free</b> (marker 2), deployment source <b>Other</b> (marker 3). <b>Review + create → Create</b> (marker 4).",
                "Open the resource. Copy the <b>URL</b> (marker 1 below) and click <b>Manage deployment token</b> to copy the token (marker 2). Keep the token private: it allows publishing.",
                "On the admin PC (Node.js LTS installed), publish the unzipped folder:<br><code>" + swa_cmd + "</code>",
                "Add the URL (with the slash at the end) as a <b>Single-page application</b> redirect URI in Entra if you did not in Stage 2.",
                "Optional: <b>Custom domains → + Add</b> → e.g. <code>navjyoti.lnbgroup.com</code> (a CNAME record in your DNS). Add that address as a redirect URI too.")
        + two(fig(fig_swa_create()), fig(fig_swa_overview()))
        + '<p class="small">The package includes <code>staticwebapp.config.json</code>, which sets the security headers (the app may only talk to Microsoft sign-in and Microsoft Graph) and stops <code>version.json</code> from being cached, so update notices arrive.</p>'
        + "<h3>Option B · Internal IIS server</h3>"
        + steps("Use a Windows server with IIS and an HTTPS binding whose certificate the PCs trust (company CA or public certificate).",
                "Copy the unzipped files to e.g. <code>C:\\inetpub\\wwwroot\\navjyoti</code>. The included <code>web.config</code> sets the file types and caching.",
                "The address is then e.g. <code>https://apps.lnbgroup.local/navjyoti/</code>. Add it, with the slash at the end, as a SPA redirect URI in Entra.",
                "The PCs still need internet access to Microsoft 365; the server only hands out the app files.")))

    parts.append(section("Stage 5", "Install on PCs and phones", "Each user, or automatic · a few minutes",
        "<h3>On one PC</h3>"
        + steps("Open the app address in <b>Microsoft Edge</b> and sign in with the LNB Group work account.",
                "Click the <b>App available</b> button in the address bar (marker 1) and then <b>Install</b> (marker 2). Or use <b>⋯ → Apps → Install this site as an app</b> (marker 3). In Chrome, use the install icon in the address bar.",
                "The tracker opens in its own window and appears in the Start menu. Right-click its taskbar icon → <b>Pin to taskbar</b>.")
        + fig(fig_edge_install())
        + "<h3>On all PCs at once (Intune or Group Policy)</h3>"
        + steps("<b>intune.microsoft.com → Devices → Configuration → + Create → Windows 10 and later → Settings catalog</b>.",
                "Add the setting <b>Microsoft Edge → Configure list of force-installed Web Apps</b> (marker 1) and paste the JSON (marker 2) with your app address.",
                "Assign it to the Navjyoti Production group (marker 3). PCs install the app at their next policy refresh.")
        + fig(fig_intune())
        + two("<div><h3>Phones and tablets</h3><p>Android: open the address in Edge or Chrome → <b>Install app</b>. iPhone and iPad: open it in Safari → <b>Share → Add to Home Screen</b>. The layout adapts to the small screen.</p></div>",
              '<figure class="phone">' + img("app-phone.png", "The tracker on a phone") + "</figure>")))

    parts.append(section("Stage 6", "Users and permissions", "Site owner · 10 min",
        steps("Open the site → <b>Settings (gear) → Site permissions</b>.",
              "<b>Add members</b> (marker 1): supervisors and data-entry staff go into <b>Site members</b> (Edit, marker 2); they can add, change and delete records.",
              "Managers and auditors go into <b>Site visitors</b> (Read, marker 3): they see everything; if they try to save, the app says “Your access is view-only”.",
              "If you made app assignment required in Stage 2, add the same people or group there too.",
              "When someone leaves, remove them from the group. Their access ends at once, on every PC.")
        + fig(fig_sp_permissions())))

    parts.append(section("Stage 7", "Acceptance checks", "IT with one supervisor · 15 min", f"""
<table class="t check"><thead><tr><th></th><th>Check</th><th>Expected</th></tr></thead><tbody>
<tr><td>☐</td><td>Open the installed app and sign in</td><td>The sidebar footer shows “Signed in as …” and “Version {V}”; the top bar shows “Live · saved to cloud”</td></tr>
<tr><td>☐</td><td>Overview, All batches</td><td>Raw material 161.3 t · net yield 43.9 t · sold 28.8 t · stock on hand 15,080 kg · revenue ₹3,82,596 (Batch 4)</td></tr>
<tr><td>☐</td><td>Add a test expense on PC 1 (it asks for confirmation)</td><td>It appears on PC 2, signed in as another person, within about 30 seconds</td></tr>
<tr><td>☐</td><td>Open the Expenses list in SharePoint</td><td>The same row is there</td></tr>
<tr><td>☐</td><td>Sign in as a site visitor and try to save</td><td>“Your access is view-only…”</td></tr>
<tr><td>☐</td><td>Delete the test expense; click Export Excel</td><td>The row disappears on both PCs; a workbook downloads</td></tr>
</tbody></table>"""
        + two("<figure>" + img("app-signedin.png", "Signed in and version in the sidebar", "shot narrow") + "<figcaption>Sidebar footer after sign-in.</figcaption></figure>",
              "<figure>" + img("app-confirm.png", "Confirmation before saving") + "<figcaption>Every save asks for confirmation and flags possible duplicates.</figcaption></figure>")))

    parts.append(section("Stage 8", "Releasing updates", "Whoever maintains the app",
        "<p>Screens and logic live in one file, <code>navjyoti-vermicompost/index.html</code>; the Microsoft 365 connection is in <code>desktop-app/src/</code>. "
        "A release only replaces the app files on the host. The records stay in SharePoint and are not touched.</p>"
        + fig(fig_release())
        + "<h3>Manual release</h3>"
        + steps("Make the change in <code>index.html</code> (or <code>src/cloud.js</code>). Changes can be made by any developer; Claude is not needed.",
                "Raise the number in <code>desktop-app/VERSION</code>: 1.0.0 → 1.0.1 for fixes, 1.1.0 for new features, 2.0.0 for large changes.",
                "Build: <code>python3 desktop-app/build_desktop_app.py</code>. This writes <code>desktop-app/dist/</code> and <code>Navjyoti_Tracker_App_v&lt;version&gt;.zip</code>.",
                "Deploy <code>dist</code> the same way as in Stage 4 (Static Web Apps command, or copy over the IIS folder).",
                "Open PCs show the notice below within about 10 minutes, or as soon as the window comes to the front; people click <b>Update now</b>. Newly opened windows load the new version straight away.")
        + img("app-update.png", "Update available notice in the app")
        + "<h3>Automatic release from GitHub (optional)</h3>"
        + steps("Copy <code>desktop-app/deploy/github-workflow.yml</code> to <code>.github/workflows/deploy-navjyoti-tracker.yml</code> in the repository.",
                "In GitHub: <b>Settings → Secrets and variables → Actions → New repository secret</b> named <code>AZURE_STATIC_WEB_APPS_API_TOKEN</code>, value = the deployment token.",
                "Commit <code>desktop-app/config.js</code> (not secret) so the automatic build includes it.",
                "From then on, every change merged into <code>main</code> builds and publishes a new version automatically (see the repository's <b>Actions</b> tab).")
        + two(note("<b>New columns or lists.</b> If a change needs a new SharePoint column, add it to <code>registers()</code> in <code>build_cloud_workbook.py</code>, rebuild the kit, and run <code>Create-NavjyotiLists.ps1</code> on the site <i>before</i> deploying. The field map in the app is generated from the same place, so both stay in step."),
              note("<b>Roll back.</b> Keep each release's zip. To go back, deploy the previous zip; PCs pick it up the same way. In GitHub, revert the commit and the workflow redeploys."))))

    parts.append(section("9", "Running it: backup, security, monitoring", "IT", """
<table class="t"><tbody>
<tr><td><b>Backup</b></td><td>SharePoint keeps version history for every item and a 93-day recycle bin. Add a Microsoft Purview retention policy for the site if you keep formal backups. Export Excel in the app gives a monthly offline archive.</td></tr>
<tr><td><b>Sign-in</b></td><td>Microsoft Entra ID with your MFA and Conditional Access policies. Tokens are kept per user in the browser profile; Sign out (sidebar) clears them.</td></tr>
<tr><td><b>Permissions</b></td><td>Delegated Graph permissions: never more than the person's own SharePoint rights. No client secret exists for this app.</td></tr>
<tr><td><b>Network</b></td><td>The app talks only to its host, <code>login.microsoftonline.com</code> and <code>graph.microsoft.com</code> (enforced by the security header), plus Google Fonts for typefaces.</td></tr>
<tr><td><b>Monitoring</b></td><td>Entra sign-in logs (filter on “Navjyoti Tracker”); Purview audit log for list item changes; Azure Static Web Apps metrics for the host.</td></tr>
<tr><td><b>Load</b></td><td>Each open window reads the 9 lists every <code>pollSeconds</code> (default 30) while it is visible, and only redraws when something changed. Microsoft Graph throttling is handled with automatic retries. For many simultaneous users, raise <code>pollSeconds</code> to 60.</td></tr>
<tr><td><b>Limits</b></td><td>Pre-compost lots keep up to 3 turnings in SharePoint (the app warns if more are entered). Dates follow the site time zone; keep the site and PCs on India Standard Time.</td></tr>
</tbody></table>"""))

    parts.append(section("10", "Troubleshooting", "IT", """
<table class="t"><thead><tr><th>What you see</th><th>Cause and fix</th></tr></thead><tbody>
<tr><td>“Setup not finished” banner</td><td><code>config.js</code> still has placeholder values. Fill it in (Stage 3) and redeploy or copy it to the host.</td></tr>
<tr><td>AADSTS50011 (redirect URI does not match)</td><td>Add the exact app address, with the slash at the end, as a <b>Single-page application</b> redirect URI (Stage 2, step 7).</td></tr>
<tr><td>AADSTS65001 or “Need admin approval”</td><td>Admin consent not granted. API permissions → Grant admin consent for LNB Group.</td></tr>
<tr><td>AADSTS50105 (user not assigned)</td><td>Assignment is required and the person is not assigned. Add them under Enterprise applications → Users and groups.</td></tr>
<tr><td>“Missing SharePoint lists: …”</td><td>The lists are not on the site in <code>siteUrl</code>. Run <code>Create-NavjyotiLists.ps1</code> on that site, or correct <code>siteUrl</code>.</td></tr>
<tr><td>“SharePoint site not found”</td><td><code>siteUrl</code> is wrong, or the person has no access to the site. Check the address and site membership.</td></tr>
<tr><td>“Your access is view-only” on save</td><td>The person is a site visitor. Move them to site members if they should edit.</td></tr>
<tr><td>Dates one day off</td><td>The site's regional time zone and the PC's time zone differ. Set both to (UTC+05:30).</td></tr>
<tr><td>No “Install” button in Edge</td><td>The app is not served over HTTPS, or the manifest is blocked. Check the certificate and that <code>.webmanifest</code> is served (included <code>web.config</code> / <code>staticwebapp.config.json</code>).</td></tr>
<tr><td>Other people's entries appear late</td><td>They appear within <code>pollSeconds</code> while the window is visible. Bring the window to the front to refresh at once.</td></tr>
<tr><td>Update notice never appears</td><td><code>version.json</code> is cached by a proxy or server. Make sure it is served with “Cache-Control: no-store” (the included configs do this).</td></tr>
<tr><td>“Offline” banner</td><td>No connection to Microsoft 365. Records refresh on their own when the connection is back; nothing typed is saved while offline.</td></tr>
</tbody></table>"""))

    parts.append(section("A", "Appendices", "Reference", """
<h3>A · Files in the app package</h3>
<table class="t"><tbody>
<tr><td><code>index.html</code></td><td>The tracker screens and logic (built from <code>navjyoti-vermicompost/index.html</code>)</td></tr>
<tr><td><code>config.js</code></td><td>Tenant, app registration and site settings (Stage 3)</td></tr>
<tr><td><code>cloud.js</code></td><td>Microsoft 365 connector: sign-in, SharePoint reads and writes through Microsoft Graph, update check</td></tr>
<tr><td><code>fieldmap.js</code></td><td>Which app field goes to which SharePoint column (generated)</td></tr>
<tr><td><code>version.json</code></td><td>The released version; open apps compare it with their own</td></tr>
<tr><td><code>manifest.webmanifest</code>, <code>sw.js</code>, <code>icons/</code></td><td>Make the app installable, with the Navjyoti icon</td></tr>
<tr><td><code>staticwebapp.config.json</code>, <code>web.config</code></td><td>Security headers and caching for Azure Static Web Apps and for IIS</td></tr>
<tr><td><code>vendor/msal-browser.min.js</code></td><td>Microsoft Authentication Library for JavaScript (MIT licence)</td></tr>
<tr><td><code>vendor/xlsx.full.min.js</code></td><td>SheetJS, for Export Excel (Apache 2.0 licence)</td></tr>
</tbody></table>
<h3>B · config.js settings</h3>
<table class="t"><tbody>
<tr><td><code>tenantId</code></td><td>Directory (tenant) ID from Entra</td></tr>
<tr><td><code>clientId</code></td><td>Application (client) ID of “Navjyoti Tracker”</td></tr>
<tr><td><code>siteUrl</code></td><td>Address of the SharePoint site with the lists</td></tr>
<tr><td><code>pollSeconds</code></td><td>How often an open app checks for other people's changes (default 30, minimum 10)</td></tr>
<tr><td><code>scopes</code> (optional)</td><td>Graph permissions requested. Default <code>['https://graph.microsoft.com/Sites.ReadWrite.All','User.Read']</code>; with Sites.Selected use <code>['https://graph.microsoft.com/Sites.Selected','User.Read']</code></td></tr>
<tr><td><code>redirectUri</code> (optional)</td><td>Only if the sign-in return address must differ from the app address</td></tr>
</tbody></table>
<h3>C · SharePoint lists and columns</h3>
<table class="t cols-list"><tbody>""" + lists_appendix() + """</tbody></table>
<p class="small">Every list also has SharePoint's own Title (unused), Created, Modified, Created By and Modified By columns. Calculated columns are worked out again by the app; the app never writes them.</p>"""))

    html_doc = TEMPLATE.replace("{{BODY}}", "\n".join(parts))
    out = os.path.join(HERE, "it_guide.html")
    open(out, "w", encoding="utf-8").write(html_doc)
    print("written", out)


TEMPLATE = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>Navjyoti Tracker IT Setup Guide</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600;12..96,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Print layout: A4 pages; each stage starts on a new page with numbered steps, drawings and real screenshots. */
:root {
  --bg: #ffffff; --surface: #ffffff; --panel: #e9efe8; --card: #f7faf6; --fg: #16201b; --muted: #55655c; --line: #d3ddd4;
  --accent: #1c7148; --accent-fg: #ffffff; --brand: #143f2b; --soil: #6b4424; --hi: #e0791f;
  --c1: #dbe8fb; --c2: #fde3d6; --c3: #d8efe2; --c4: #efe7f8; --term: #14201a; --term-t: #d7e6dc; --term-ok: #7fd6a4;
  --f-display: "Bricolage Grotesque", "Segoe UI", system-ui, sans-serif;
  --f-body: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif;
  --f-mono: "IBM Plex Mono", Consolas, ui-monospace, monospace;
}
@page { size: A4; margin: 14mm 13mm 16mm }
* { box-sizing: border-box }
body { background: var(--bg); color: var(--fg); font: 10.5pt/1.5 var(--f-body); margin: 0 }
h1, h2, h3 { font-family: var(--f-display); margin: 0; letter-spacing: -0.01em; text-wrap: balance }
h3 { font-size: 13pt; margin: 6px 0 2px; color: var(--brand) }
p { margin: 0 }
code { font-family: var(--f-mono); font-size: 9pt; background: var(--panel); padding: 0 4px; border-radius: 3px; overflow-wrap: anywhere }
.small { font-size: 9pt; color: var(--muted) }
.cover { min-height: 262mm; display: flex; flex-direction: column; gap: 14px; page-break-after: always }
.cover-logo { display: flex; align-items: center; gap: 16px; padding: 18px 22px; background: var(--brand); border-radius: 14px }
.cover-logo .logo { width: 104px; height: auto; background: #fff; border-radius: 12px; padding: 10px }
.cover-logo span { color: #eef6f0; font: 600 16pt var(--f-body) }
.eyebrow { font: 600 9pt/1 var(--f-mono); letter-spacing: .1em; text-transform: uppercase; color: var(--soil); margin-top: 18px }
.cover h1 { font-size: 30pt; line-height: 1.05; color: var(--brand) }
.lede { color: var(--muted); font-size: 12pt }
.summary { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 8px }
.summary div { border: 1px solid var(--line); border-top: 3px solid var(--accent); border-radius: 8px; padding: 10px 12px; display: grid; gap: 4px }
.summary b { color: var(--brand) } .summary span { font-size: 10pt }
.toc { margin-top: auto; background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 12px 16px }
.toc ol { margin: 6px 0 0; padding-left: 20px; columns: 2; column-gap: 28px; font-size: 10pt }
.stage { page-break-before: always; display: grid; gap: 10px }
.stage > header { display: flex; align-items: baseline; gap: 10px; border-bottom: 2px solid var(--brand); padding-bottom: 6px; margin-bottom: 2px }
.stage > header .n { white-space: nowrap; font: 700 9pt/1 var(--f-mono); letter-spacing: .08em; text-transform: uppercase; color: #fff; background: var(--accent); padding: 4px 8px; border-radius: 4px }
.stage > header h2 { font-size: 18pt; color: var(--brand) }
.stage > header .who { margin-left: auto; font-size: 9pt; color: var(--muted); white-space: nowrap }
h3 { break-after: avoid; page-break-after: avoid }
ol.steps { margin: 0; padding: 0; list-style: none; counter-reset: s; display: grid; gap: 6px }
ol.steps li { counter-increment: s; display: grid; grid-template-columns: 22px 1fr; gap: 8px; align-items: start; break-inside: avoid }
ol.steps li::before { content: counter(s); display: grid; place-items: center; width: 20px; height: 20px; border-radius: 50%; background: var(--hi); color: #fff; font: 700 9pt/1 var(--f-body); margin-top: 1px }
figure { margin: 0; break-inside: avoid }
figcaption { font-size: 9pt; color: var(--muted); margin-top: 4px }
.fig-scroll svg { display: block; width: 100%; height: auto }
img.shot { display: block; width: 100%; height: auto; border: 1px solid var(--line); border-radius: 8px; break-inside: avoid }
img.shot.narrow { width: 62% }
.frame { border: 1px solid var(--line); border-radius: 10px; overflow: hidden }
.frame .win { display: flex; align-items: center; gap: 6px; padding: 7px 10px; background: var(--brand); color: #eef6f0; font-size: 9pt }
.frame .win span { width: 9px; height: 9px; border-radius: 50%; background: #8fae9c }
.frame .win b { margin-left: 8px; font-weight: 600 }
.frame img.shot { border: 0; border-radius: 0 }
.frame figcaption { padding: 6px 10px }
.phone img.shot { width: 62%; margin: 0 auto }
.cols { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; align-items: start }
.note { background: var(--card); border: 1px solid var(--line); border-left: 4px solid var(--hi); border-radius: 6px; padding: 8px 12px; font-size: 9.5pt; break-inside: avoid }
table.t { width: 100%; border-collapse: collapse; font-size: 9.5pt; break-inside: auto }
table.t th { text-align: left; background: var(--brand); color: #fff; padding: 5px 8px; font-weight: 600 }
table.t td { border-bottom: 1px solid var(--line); padding: 5px 8px; vertical-align: top }
table.t tr { break-inside: avoid }
table.t.fill td.blank { width: 32%; border-bottom: 1px dashed var(--muted) }
table.t.check td:first-child { width: 22px; font-size: 12pt }
table.cols-list td:first-child { width: 22%; white-space: nowrap }
table.cols-list span { color: var(--muted); font-size: 8.5pt }
.code { border: 1px solid var(--line); border-radius: 6px; overflow: hidden; break-inside: avoid }
.code-h { background: var(--panel); padding: 4px 10px; font-size: 9pt; font-weight: 600 }
.code pre { margin: 0; padding: 8px 10px; background: var(--term); color: var(--term-t); font: 8.6pt/1.45 var(--f-mono); white-space: pre-wrap; overflow-wrap: anywhere }
/* drawing classes shared with the Microsoft 365 guide */
.f-win { fill: var(--surface); stroke: var(--line) } .f-bar { fill: var(--panel) } .f-dot { fill: var(--line) }
.f-panel { fill: var(--panel) } .f-card { fill: var(--card); stroke: var(--line) } .f-input { fill: var(--surface); stroke: var(--line) }
.f-row { fill: var(--card) } .f-head { fill: var(--accent) }
.f-btn { fill: var(--surface); stroke: var(--line) } .f-btn-p { fill: var(--accent) }
.f-c1 { fill: var(--c1) } .f-c2 { fill: var(--c2) } .f-c3 { fill: var(--c3) } .f-c4 { fill: var(--c4) }
.f-phone { fill: var(--card); stroke: var(--muted); stroke-width: 2 }
.f-term { fill: var(--term) } .f-term-t { fill: var(--term-t) } .f-term-ok { fill: var(--term-ok) }
.f-t { fill: var(--fg); font-family: var(--f-body) } .f-t-inv { fill: var(--accent-fg); font-family: var(--f-body) } .f-muted { fill: var(--muted); font-family: var(--f-body) }
.f-code { fill: var(--fg) } .f-mono { font-family: var(--f-mono) }
.f-line { stroke: var(--muted); stroke-width: 1.2 }
.f-arrow { stroke: var(--muted); stroke-width: 2 } .f-arrowhead { fill: var(--muted) }
.f-hi { fill: none; stroke: var(--hi); stroke-width: 2.5; stroke-dasharray: 6 4 } .f-badge { fill: var(--hi) }
</style></head>
<body>
{{BODY}}
</body></html>
'''

if __name__ == "__main__":
    build()
