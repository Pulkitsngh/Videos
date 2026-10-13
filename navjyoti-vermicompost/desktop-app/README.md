# Navjyoti Tracker – LNB Group Microsoft 365 edition

The same tracker as the web app (`../index.html`), packaged to run on LNB Group's own Microsoft 365, with no Claude involved:

- people sign in with their LNB Group work account (Microsoft Entra ID),
- records are stored in the SharePoint lists of the Navjyoti Production site (created by `../m365/Create-NavjyotiLists.ps1`),
- the app installs on Windows PCs from Microsoft Edge or Chrome (and on phones), and every PC sees the same data,
- new versions reach every PC when IT deploys them; running copies show "Version x.y.z is available – Update now".

The full, illustrated IT guide is `Navjyoti_Tracker_IT_Setup_Guide.pdf` (built from `docs/build_it_guide.py`).

## Files

| Path | What it is |
|---|---|
| `build_desktop_app.py` | Builds `dist/` and `../Navjyoti_Tracker_App_v<VERSION>.zip` from `../index.html` |
| `VERSION` | The version number shown in the app; raise it for every release |
| `config.example.js` | Settings template: tenant ID, app (client) ID, SharePoint site address. Copy to `config.js` |
| `src/cloud.js` | Microsoft 365 connector: sign-in (MSAL), SharePoint lists via Microsoft Graph, update check |
| `src/sw.js`, `src/manifest.webmanifest` | Make the app installable; network-first so updates arrive at once |
| `src/staticwebapp.config.json` | Azure Static Web Apps headers (security policy, caching) |
| `src/web.config` | The same for an internal IIS server |
| `vendor/` | Microsoft Authentication Library (MIT) and SheetJS (Apache-2.0), served from the app itself |
| `icons/` | App icons made from the Navjyoti logo |
| `deploy/github-workflow.yml` | Optional automatic deployment on every push |

## Release a new version

1. Change `../index.html` (screens and logic) or `src/` (Microsoft 365 connection).
2. Raise `VERSION` (for example 1.0.0 → 1.1.0).
3. `python3 build_desktop_app.py` (needs Python 3 with `openpyxl`).
4. Deploy `dist/` (Azure: `npx @azure/static-web-apps-cli deploy ./dist --deployment-token <token> --env production`; IIS: copy the folder), or push to `main` if the GitHub workflow is set up.
5. Open PCs show the update prompt within about 10 minutes, or when the app window is next brought forward; new windows load the new version straight away.

If SharePoint columns change, run `Create-NavjyotiLists.ps1` again first: it adds missing lists and columns and keeps existing data.
