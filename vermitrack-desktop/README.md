# Navjyoti VermiTrack — desktop edition

Desktop (Windows / macOS) version of the VermiTrack vermicompost production register,
converted from the v19 web app (Next.js on OpenAI Sites / Cloudflare D1).

| | Web version (v19) | Desktop edition |
|---|---|---|
| Runs as | Website / PWA | Electron app with Windows and Mac installers |
| Sign-in | Sign in with ChatGPT + passkeys | Microsoft 365 / Outlook work account (MSAL) |
| Shared data | Cloudflare D1 database | SharePoint lists on your Microsoft 365 site |
| Notifications | — | Outlook emails for deletion requests and decisions, invitations, and reports |

All screens, the 8 registers, the reports and charts, the admin and audit features, the 24-hour
edit rule, deletion approvals with the 90-day recycle bin, and segment/batch permissions are
carried over from the web version.

**Setting it up for the team:** see [SETUP-GUIDE.md](SETUP-GUIDE.md).

## How it works

- `src/App.tsx` is the original UI (`app/page.tsx`), lightly adapted for desktop. It still calls `fetch("/api/...")`.
- `electron/main.ts` serves the UI on an internal `app://vermitrack` protocol and answers `/api/*` itself, so no web server is involved.
- `electron/api.ts` contains the original Next.js route handlers, ported to a storage interface.
- `electron/store.ts` maps each original table to a SharePoint list (`SharePointStore`, through Microsoft Graph). The full row is stored as JSON in `vtData`; the other columns mirror key fields so the lists stay readable in SharePoint and Excel. `LocalStore` (a JSON file) backs demo mode and the tests.
- `electron/auth.ts` handles Microsoft sign-in with the system browser. The token cache is encrypted with the OS keychain.
- `electron/seed.json` holds the Batch 3/4/5 register, the baseline and the approved users, extracted from the original Drizzle migrations. It is loaded into empty lists on first setup.
- `shared/policy.ts` holds the Key Person and the deletion approvers.

## Development

```bash
npm install
npm run dev          # build and open the app (use "Try demo mode" without Microsoft 365)
npm test             # SharePoint store + API against an in-memory Microsoft Graph
npm run test:e2e     # drives the real app with Playwright (Linux needs xvfb)
npm run dist:win     # Windows installer → release/ (on Linux this needs wine + wine32)
npm run dist:mac     # macOS .dmg (must run on a Mac)
```

The GitHub Actions workflow `.github/workflows/vermitrack-desktop.yml` builds the Windows and
Mac installers on every push that touches this folder. Download them from the run's **Artifacts**.
