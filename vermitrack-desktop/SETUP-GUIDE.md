# Navjyoti VermiTrack desktop — Microsoft 365 setup guide

VermiTrack is a desktop app for Windows and Mac. Everyone signs in with their own
Microsoft 365 / Outlook work account, and the register is stored on your company's
SharePoint site, so the whole team sees the same records and reports at the same time.

```
 Colleague PCs (VermiTrack app) ──sign in──▶ Microsoft 365 (your company account)
            │                                        │
            └──── read / write records ────▶ SharePoint site "VermiTrack" (7 lists)
            └──── notifications, reports ──▶ Outlook (sent from your own mailbox)
```

There is no server to rent or maintain. The data stays inside your company's Microsoft 365.

You do the setup **once**, in about 15 minutes. You need someone who can create app registrations in
Microsoft Entra (usually your IT administrator) and who owns, or can create, a SharePoint site.

---

## Step 1: Create the SharePoint site (2 minutes)

1. Open <https://www.office.com>, choose **SharePoint** and click **+ Create site** → **Team site**.
2. Name it **VermiTrack** and set privacy to **Private**.
3. Add every colleague who will use VermiTrack as a **Member**.
   Members need edit rights so they can save entries. Use **Visitor** (read-only) for people who only view reports.
4. Copy the site address from the browser, for example
   `https://lnbgroup.sharepoint.com/sites/VermiTrack`.

## Step 2: Register the app in Microsoft Entra (5 minutes, IT administrator)

1. Open <https://entra.microsoft.com> → **Identity** → **Applications** → **App registrations** → **New registration**.
2. Fill in:
   - **Name:** `Navjyoti VermiTrack`
   - **Supported account types:** *Accounts in this organizational directory only (single tenant)*
   - **Redirect URI:** platform **Public client/native (mobile & desktop)**, value `http://localhost`
3. Click **Register**. On the overview page, copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**
4. Open **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, and add:

   | Permission | Why VermiTrack needs it |
   |---|---|
   | `User.Read` | Know who is signed in |
   | `Sites.ReadWrite.All` | Read and save register entries in the SharePoint lists |
   | `Sites.Manage.All` | Create the VermiTrack lists the first time (Step 3) |
   | `Mail.Send` | Send deletion notifications, invitations and reports from the user's own Outlook |

5. Click **Grant admin consent for <your organisation>**, so colleagues are not asked to approve permissions individually.

> These are *delegated* permissions. VermiTrack can only reach what the signed-in person can
> already reach. A colleague who is not a member of the VermiTrack site cannot open its lists.

## Step 3: Install VermiTrack and connect it (3 minutes, Key Person)

1. Run **Navjyoti VermiTrack Setup 1.0.0.exe** on Windows, or open the `.dmg` on Mac.
   The installer is not code-signed yet:
   - **Windows:** if SmartScreen appears, click **More info → Run anyway**.
   - **Mac:** right-click the app → **Open** the first time. If macOS says the app is damaged, run `xattr -cr "/Applications/Navjyoti VermiTrack.app"` in Terminal once.
2. Choose **First-time setup (administrator)** and enter the client ID, the tenant ID and the SharePoint site address.
   Optionally, add a link where colleagues can download the installer (for example a OneDrive or SharePoint file link).
3. Click **Sign in with Microsoft 365**. Your browser opens; sign in as `pulkit.singh@lnbgroup.com`.
4. Click **Set up VermiTrack on this SharePoint site**. This creates 7 lists on the site and loads:
   - the imported Batch 3, 4 and 5 register (149 entries and the Batch 4 baseline)
   - the 5 approved users from the web version (Pulkit, Ashok, Harish, Akola Team, Pritam)

## Step 4: Bring over records entered in the web version (optional)

The 149 historical entries come from the original Excel import. If your team entered more records
in the online version after that:

1. In the **old web app**, open **Records**, pick a batch, choose **All registers** and click **Export CSV**. Repeat for each batch.
2. In the **desktop app**, open **Admin** (your initials, top right) → **Bring records from the web version** and choose each file.

Records that are already present are skipped, so importing the same file twice is safe.

## Step 5: Add colleagues

1. In **Admin → Approved users**, add or check each person. Their email must be their Microsoft 365 sign-in address.
2. Open their card and click **✉ Email setup instructions**. They receive an Outlook email containing a **setup code**.
3. They install VermiTrack, choose **I have a setup code**, paste it and sign in with their own account.

The setup code is also under **More → Add colleagues**, if you'd rather paste it into Teams.
The code contains only the app ID, tenant ID and site address; it is not a password.

---

## Day-to-day

- **Live data.** Every save goes straight to SharePoint. Each open app refreshes every 20 seconds, and the **Live** pill refreshes immediately.
- **Reports.** Open the **Reports** tab and use **✉ Email report** to send a batch summary from your own Outlook; it appears in your Sent Items. **↓ Export report** saves a full CSV for Excel.
- **Deletion approvals.** A deletion request emails Pulkit Singh and Pritam Suryavanshi. When they approve or reject it, the requester is emailed. Approved deletions stay in a 90-day recycle bin.
- **Excel / Power BI.** The SharePoint lists (for example *VermiTrack Entries*) can be opened in Excel or Power BI. The key columns are readable directly; the *Record data (JSON)* column holds every field.

## Good to know

- **Use the app to make changes.** Roles (Admin, Viewer and so on), segment and batch limits, the 24-hour edit window and deletion approvals are enforced by the VermiTrack app. Anyone who is a site **Member** could still change list items directly in SharePoint, so keep site membership to the VermiTrack team. SharePoint's version history on each list records any direct edits.
- **Demo mode.** On the first screen, *Try demo mode* runs VermiTrack with sample data stored only on that computer. Nothing is shared. Use **More → Leave demo mode** to connect for real.
- **Troubleshooting.** **File → Open data folder** shows `vermitrack.log`. Common messages:
  - *"not an approved VermiTrack user"*: add the person under Admin → Approved users, using their exact sign-in email.
  - *"does not have permission for the VermiTrack SharePoint site"*: add them as a Member of the site.
  - *"Sign-in did not complete"*: check that the redirect URI `http://localhost` is set as a **Public client/native** platform in Step 2.
