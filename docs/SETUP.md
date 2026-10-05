# Running the Outreach Desk on Netlify

The standalone version signs you in with your Cludo Microsoft account, reads your Outlook mail
through Microsoft Graph, and keeps prospect data as JSON files in a private app folder of your
OneDrive. There is no server and no third-party database: everything stays inside Cludo's
Microsoft 365 tenant. Netlify only serves static files.

Two things have to be set up once. Only step 1 needs someone with Azure access.

## 1. Register the app in Microsoft Entra ID (Azure)

Anyone at Cludo with permission to register applications can do this; it takes about five minutes.
If "App registrations" is greyed out, send this section to IT.

1. Go to <https://portal.azure.com> → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Fill in:
   - **Name:** `Cludo Outreach Desk`
   - **Supported account types:** *Accounts in this organizational directory only (Cludo only – Single tenant)*
   - **Redirect URI:** platform **Single-page application (SPA)**, URI
     `https://illustrious-begonia-3a001c.netlify.app/`
     (the trailing slash matters; add the custom domain later if one is set up)
3. Click **Register**. On the **Overview** page copy two values; they go into `config.js` in this repo:
   - **Application (client) ID**
   - **Directory (tenant) ID**
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, tick:
   - `User.Read` — who is signed in (name, email)
   - `Mail.Read` — read Sent Items and Inbox to detect sent emails, replies and bounces (read only; the app never sends mail)
   - `Files.ReadWrite.AppFolder` — one dedicated folder in the user's OneDrive (`Apps/Cludo Outreach Desk`) for the prospect data; nothing else in OneDrive is reachable
5. Still on **API permissions**, click **Grant admin consent for Cludo** if the button is enabled.
   If it is greyed out, ask IT to do it (a Global Administrator or Privileged Role Administrator).
   Without it, the first sign-in shows *"Need admin approval"*.
6. **Authentication** → under *Single-page application* confirm the redirect URI is listed. Leave the
   implicit-grant boxes unticked (the app uses the authorization code flow with PKCE). No client
   secret is needed, and none should be created.

Both IDs are safe to commit: a single-page app has no secret, and sign-in is restricted to Cludo accounts and
to the registered redirect URIs.

## 2. Connect Netlify to this repository

So that every push to `main` goes live without uploading files by hand:

1. In Netlify open the site → **Site configuration** → **Build & deploy** → **Continuous deployment** → **Link repository**.
2. Choose GitHub → `bakalliandrea-hash/cludo-outreach`, branch `main`.
3. Build command: leave empty. Publish directory: `.` (the repository root; `netlify.toml` already says so).
4. Save. Netlify deploys on each push.

## What the app can and cannot reach

- Mail: read-only metadata and previews of your own mailbox (subject, sender, recipients, dates, first lines).
- OneDrive: only the folder `Apps/Cludo Outreach Desk`. Deleting that folder deletes the app's data.
- Nothing is sent, moved or deleted in Outlook. Drafts still open in Outlook for you to send.
- Each signed-in person sees only their own data (it lives in their own OneDrive).
