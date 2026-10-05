# cludo-outreach

**Cludo Outreach Desk** is a single-page tracker for BDR email outreach. It runs as a Claude artifact.

Live app: <https://claude.ai/artifact/XDK6xQrCLeuwycc1QyZN88>

## What it does

- **Campaigns** with a three-email sequence: Email 1, Follow-up 1 and Follow-up 2. Each campaign has its own templates and wait times (in working days), and a list of public holidays to skip.
- **Prospect import** from a `contacts.json` file or a CSV with the columns `first_name, last_name, email, organisation, title`. Optional columns: `area, website, linkedin, city, subject, body`. Prospects are matched by email address, so importing a file again updates them instead of adding duplicates.
- **Queue tabs:** To contact, Follow-up due, Waiting, Replied and Finished, plus summary stats.
- **Drafts open in Outlook** (web or desktop app), already filled in. The page never sends email itself.
- **Outlook sync** through the Microsoft 365 connector (`outlook_email_search`). It reads Sent Items and the Inbox to find which emails went out, who replied and which emails bounced.

## Files

- `index.html`: the full app (HTML, CSS and JavaScript in one file).

## Running it

The page uses Claude artifact runtime capabilities (`window.claude.use`):

| Capability | Used for |
|---|---|
| `db` | storing campaigns and prospects (private to each user) |
| `user` | identifying who is signed in (`profile` scope) |
| `mcp` | reading Outlook via the Microsoft 365 connector |

Opened outside Claude (for example as a local file), the page loads but shows a notice that storage and Outlook access aren't available. To change the app, edit `index.html` and republish it to the artifact.

## Data

No prospect data is stored in this repository. Campaigns and prospects live only in the artifact's database, under each user's own space.
