# Cludo Outreach Desk

The app is one file, `index.html`, published as a Claude artifact:
https://claude.ai/artifact/XDK6xQrCLeuwycc1QyZN88

## Keep the repo and the live artifact in sync

Every change to `index.html` has to end up in both places, in this order:

1. Before editing, read the live artifact (Artifact tool, `action: "read"`, the URL above). If it differs from `index.html` (for example, someone edited it in Cowork), copy the live version into `index.html` and commit that first, so no change gets lost.
2. Edit `index.html`.
3. Commit and push to `main`.
4. Republish `index.html` to the same artifact URL. Leave out `capabilities` so the stored ones (`db`, `mcp` for Microsoft 365 `outlook_email_search`, `user` with `profile` scope) stay as they are. Leave out `icon` too.

## Constraints

- The repository is public. Never commit prospect data, contact exports, email addresses of real people or anything read from the artifact's database or from Outlook.
- The app relies on the artifact runtime (`window.claude.use('db' | 'user' | 'mcp' | 'permissions')`). Outside Claude it only shows a notice.
- Stored data layout (`data/users/{uid}/{campaignId}` and `.../chunks/p{n}`, 40 prospects per chunk) is what existing users' data lives in. A UI change must not change it. If the layout ever has to change, add a migration the way `loadProspects` does.
- Documents returned by the artifact `db` (`doc.data()`, query snapshots) are **frozen**. The script is non-strict, so writing a property on one fails silently. Always `clone()` a document before adding fields or editing it. This was the bug that made the prospect list come back empty on every reload.
- `data/users/<uid>/<cid>/prospects/*` is the old one-document-per-prospect layout. `mergeLegacy` folds anything still there into the chunks and deletes it; keep that function until it has run for every user.
- Colors are tokens on `:root` with dark-mode overrides. Keep new colors as tokens, and keep the layout working at phone width.
