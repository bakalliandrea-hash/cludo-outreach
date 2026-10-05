# cludo-outreach

**Cludo Outreach Desk** is a single-page tracker for BDR email outreach: a three-email sequence per prospect, queues for
what to send today, drafts that open in Outlook, and a sync that reads Sent Items and the Inbox to detect sent emails,
replies and bounces. Nothing is ever sent from the page.

It runs in two places from the same code:

| | Where | Sign-in | Data | Outlook |
|---|---|---|---|---|
| Claude artifact | <https://claude.ai/artifact/XDK6xQrCLeuwycc1QyZN88> | claude.ai | artifact database, private per user | Microsoft 365 connector |
| Standalone site | <https://illustrious-begonia-3a001c.netlify.app/> | Microsoft (Cludo tenant) | JSON files in the user's own OneDrive (`Apps/Cludo Outreach Desk`) | Microsoft Graph |

The standalone site needs a one-time Azure app registration: see [`docs/SETUP.md`](docs/SETUP.md).

## What it does

- **Campaigns** with Email 1, Follow-up 1 and Follow-up 2; templates with `{first_name}`, `{organisation}`, `{area}` … placeholders; wait times in working days, skipping weekends and a list of public holidays.
- **Import** from a JSON contacts file or a CSV (commas, semicolons or tabs). Prospects are matched by email, so re-importing updates instead of duplicating. A **backup** (export/import) carries the full history between the two versions.
- **Queues:** To contact, Follow-up due, Waiting, Replied, Finished, plus summary stats.
- **Sync** detects sent emails (a resend within 10 minutes counts once), replies (by sender, by conversation, or by the Email 1 subject when a colleague answers from another address), and bounces (a bounced email does not count as sent). Ambiguous matches are listed for manual handling rather than guessed.

## Files

`index.html` (markup) · `app.css` · `app.js` (all logic) · `store-claude.js` / `store-graph.js` (storage + mail adapters) · `config.js` (Azure IDs) · `vendor/msal-browser.min.js` · `netlify.toml` · `test/` (Playwright suites with synthetic data).

## Development

Edit, run the two test suites (`test/claude-tests.js`, `test/graph-tests.js`; see `CLAUDE.md`), push to `main`. Netlify
deploys from `main`; the artifact is republished with the same files. No prospect data lives in this repository.
