# Cludo Outreach Desk

A BDR outreach tracker. One codebase, two ways to run it:

- **Claude artifact** (private, inside claude.ai): https://claude.ai/artifact/XDK6xQrCLeuwycc1QyZN88
- **Standalone site on Netlify** (Microsoft sign-in): https://illustrious-begonia-3a001c.netlify.app/ — setup in `docs/SETUP.md`

## Files

| File | Role |
|---|---|
| `index.html` | markup only; loads the files below |
| `app.css` | styles (color tokens on `:root`, dark-mode overrides) |
| `app.js` | all UI and business logic; talks to a *store adapter*, never to storage or mail directly |
| `store-claude.js` | adapter for the artifact runtime: artifact `db` + Microsoft 365 connector (`outlook_email_search`) |
| `store-graph.js` | adapter for the standalone site: MSAL sign-in, OneDrive app folder (JSON files), Microsoft Graph mail |
| `config.js` | Azure client/tenant IDs for the standalone site (public identifiers, not secrets) |
| `vendor/msal-browser.min.js` | MSAL, vendored from npm `@azure/msal-browser` (CDNs are blocked from the dev environment) |
| `netlify.toml` | static deploy of the repository root |

`app.js` picks the adapter at start: `window.claude.use` present → Claude, else Graph. Both adapters implement the same
interface (`init`, `loadIndex`, `saveCampaign`, `patchCampaign`, `deleteCampaign`, `saveProfile`, `loadProspects`,
`saveProspects`, `mail.{access,request,scan,errorText}`); keep them in step when you add a method.

## Keep the repo, the artifact and Netlify in sync

Every change ends up in all three, in this order:

1. Before editing, read the live artifact (Artifact tool, `action: "read"`). If it differs from the repo, copy the live version in and commit that first.
2. Edit, run the tests (below), commit, push to `main`. Netlify deploys `main` automatically once the site is linked to the repo (see `docs/SETUP.md`).
3. Republish the artifact: `file_path: index.html`, `root` = the folder holding the files, `files: {app.css, app.js, config.js, store-claude.js, store-graph.js}` mapped to themselves. The Artifact tool only reads from the session's working directory or scratchpad, so when this repo is cloned elsewhere, copy those six files into a scratchpad folder first and publish from there. Leave out `capabilities` (stored: `db`, `mcp` Microsoft 365 `outlook_email_search`, `user` with `profile` scope) and `icon`. Afterwards `list` with `scope: "files"` should show the six files.

## Constraints

- The repository is public. Never commit prospect data, contact exports, email addresses of real people or anything read from the artifact's database, OneDrive or Outlook. Test fixtures are synthetic.
- Documents returned by the artifact `db` are **frozen**; the script is non-strict, so writing a property on one fails silently. Always `clone()` before adding fields or editing. (This bug made the prospect list come back empty on every reload.)
- Stored layouts are what users' data lives in; a UI change must not alter them:
  - artifact: `data/users/<uid>/profile`, `data/users/<uid>/<cid>` (campaign), `data/users/<uid>/<cid>/chunks/p<n>` (40 prospects per chunk). `data/users/<uid>/<cid>/prospects/*` is the old one-document-per-prospect layout; `mergeLegacy` in `store-claude.js` folds it in and deletes it — keep it until it has run for every user.
  - OneDrive: `Apps/Cludo Outreach Desk/index.json` (`{campaigns, profile}`) and `prospects-<cid>.json` (`{items}`), written with `If-Match`; Graph's simple upload caps a file at 4 MB (≈5,000 prospects).
- Prospect record: `{first_name, last_name, email, organisation, title, area, website, linkedin, city, subject, body, review, addedAt, log: {sent[], replies[], bounces[], threads[]}, manual: {sent[], replied, handled, ignoreRepliesBefore}, pending}`. Dates are ISO strings; `info()` in `app.js` derives the status from them.
- Nothing is ever sent from the page; drafts open in Outlook and a step counts only when the email appears in Sent Items.

## Tests

Playwright with the preinstalled Chromium; both scripts exit non-zero on a failure and print one PASS/FAIL line per check:

```
NODE_PATH=/opt/node-tools/node_modules node test/claude-tests.js   # artifact adapter: harness-claude.js mocks window.claude with FROZEN documents + the connector
NODE_PATH=/opt/node-tools/node_modules node test/graph-tests.js    # standalone adapter: harness-graph.js mocks msal and fetch() for graph.microsoft.com
```

`test/serve.js` serves the repo at `http://app.test/` inside the browser (everything else is blocked; `overrides` swaps a
file, used to inject a test `config.js`). Fixtures are generated in the scripts; never add real contacts. Run both before every push.
