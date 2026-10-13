"""Build the installable Navjyoti Tracker (LNB Group Microsoft 365 edition) from the tracker page.

    python3 build_desktop_app.py            -> dist/ and ../Navjyoti_Tracker_App_v<VERSION>.zip

The tracker page (../index.html) is the single source for screens and logic. This script:
  * wraps it as a stand-alone web app (head, app manifest, icons, service worker),
  * swaps its data connection for cloud.js, which signs in with Microsoft 365 and stores records in SharePoint,
  * writes fieldmap.js (page fields <-> SharePoint columns) from the Excel tracker's registers,
  * stamps the version from VERSION into the page and version.json, so running PCs are told about updates.
To release a change: edit ../index.html (or src/), raise VERSION, run this script, deploy dist/.
"""
import json, os, re, shutil, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
from build_cloud_workbook import registers  # noqa: E402


def field_map():
    out = {}
    for sheet, (key, cols) in registers().items():
        fields = []
        for h, k, kind, _ in cols:
            if not k or kind.startswith("calc") or (key == "precompost" and h.startswith("Turn")):
                continue
            fields.append([k, h, "date" if kind == "date" else "num" if kind in ("kg", "num", "inr") else "text"])
        out[key] = {"list": sheet, "fields": fields}
        if key == "precompost":
            out[key]["turns"] = 3
    return out


def rep(s, a, b, count=1):
    assert s.count(a) == count, f"build: expected {count}x {a[:70]!r} in index.html, found {s.count(a)}"
    return s.replace(a, b)


def build_page(src, version):
    s = open(src, encoding="utf-8").read()
    # libraries served from the app itself
    s = rep(s, "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "vendor/xlsx.full.min.js")
    # wording that referred to claude.ai
    s = rep(s, "<b>The register is not connected.</b> Open this tracker on claude.ai while signed in to load and save records. Records are stored in the tracker's cloud database, so every signed-in teammate with access sees the same data.",
            "<b>The register is not connected.</b> Sign in with your LNB Group work account. Records are stored in the Navjyoti SharePoint site, so every colleague with access sees the same data. If this stays, see the message at the top of the page or contact IT.")
    s = rep(s, "msg('Not connected. Open the tracker on claude.ai while signed in.')", "msg('Not connected to LNB Group Microsoft 365. Check the connection and sign in again.')")
    # status banner and sidebar footer (sign-in, version)
    s = rep(s, '<div class="work">', '<div class="work">\n<div class="cloud-msg" id="cloudMsg" hidden></div>')
    s = rep(s, '  <nav class="tabs" id="tabs" aria-label="Registers"></nav>\n</aside>',
            '  <nav class="tabs" id="tabs" aria-label="Registers"></nav>\n  <div class="side-foot" id="sideFoot"></div>\n</aside>')
    css = """
.cloud-msg { display: flex; flex-wrap: wrap; gap: 8px 14px; align-items: center; padding: 10px 24px; font-size: 13px; border-bottom: 1px solid var(--line) }
.cloud-msg.warn { background: var(--warn-bg) } .cloud-msg.bad { background: var(--bad-bg) } .cloud-msg.info { background: var(--info-bg) }
.cloud-msg code { font-family: var(--f-mono); font-size: 12px }
.cloud-msg button { border: 1px solid var(--line); background: var(--surface); color: var(--fg); padding: 5px 12px; border-radius: var(--r); cursor: pointer; font: inherit; font-weight: 600 }
.side-foot { margin-top: auto; display: grid; gap: 6px; padding: 12px 8px 0; border-top: 1px solid var(--brand-line); font-size: 12px; color: var(--brand-muted) }
.side-foot b { color: var(--brand-fg); font-weight: 600 }
.side-foot button { all: unset; cursor: pointer; color: var(--brand-fg); text-decoration: underline; justify-self: start }
.side-foot .ver { font-family: var(--f-mono); font-size: 11px }
@media (max-width: 1000px) { .side-foot { display: none } }
</style>"""
    s = rep(s, "</style>", css, count=1)
    head = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#143f2b">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" type="image/png" sizes="32x32" href="icons/favicon-32.png">
<link rel="apple-touch-icon" href="icons/apple-touch-icon.png">
<style>body {{ margin: 0 }} img {{ max-width: 100% }}</style>
<script>window.NAVJYOTI_VERSION = {json.dumps(version)};</script>
<script src="config.js"></script>
<script src="fieldmap.js"></script>
<script src="vendor/msal-browser.min.js"></script>
<script src="cloud.js"></script>
"""
    # the page starts with <title>, links and <style>; they belong in <head>, the markup in <body>
    i = s.index("</style>") + len("</style>")
    return head + s[:i] + "\n</head>\n<body>\n" + s[i:] + "\n</body>\n</html>\n"


def main():
    version = open(os.path.join(HERE, "VERSION")).read().strip()
    dist = os.path.join(HERE, "dist")
    shutil.rmtree(dist, ignore_errors=True)
    os.makedirs(os.path.join(dist, "vendor"))
    open(os.path.join(dist, "index.html"), "w", encoding="utf-8").write(build_page(os.path.join(ROOT, "index.html"), version))
    for f in os.listdir(os.path.join(HERE, "src")):
        shutil.copy(os.path.join(HERE, "src", f), dist)
    for f in os.listdir(os.path.join(HERE, "vendor")):
        shutil.copy(os.path.join(HERE, "vendor", f), os.path.join(dist, "vendor"))
    shutil.copytree(os.path.join(HERE, "icons"), os.path.join(dist, "icons"))
    cfg = os.path.join(HERE, "config.js")
    shutil.copy(cfg if os.path.exists(cfg) else os.path.join(HERE, "config.example.js"), os.path.join(dist, "config.js"))
    with open(os.path.join(dist, "fieldmap.js"), "w") as f:
        f.write("// Generated by build_desktop_app.py from the Excel tracker's registers: page field -> SharePoint column.\n")
        f.write("window.NAVJYOTI_FIELDS = " + json.dumps(field_map(), separators=(",", ":")) + ";\n")
    json.dump({"version": version}, open(os.path.join(dist, "version.json"), "w"))
    z = os.path.join(ROOT, f"Navjyoti_Tracker_App_v{version}.zip")
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
        for base, _, files in os.walk(dist):
            for f in files:
                p = os.path.join(base, f)
                zf.write(p, os.path.relpath(p, dist))
    print(f"built version {version}: {dist}\npackage: {z}" + ("" if os.path.exists(cfg) else "\nnote: config.js not found, so the package uses config.example.js (fill it in before deploying)"))


if __name__ == "__main__":
    main()
