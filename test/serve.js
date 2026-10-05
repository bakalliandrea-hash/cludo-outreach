// Serve the repo's files at http://app.test/ inside Playwright; everything else is blocked. `overrides` replaces a path's content.
const fs = require('fs'), path = require('path');
const ROOT = process.env.APP_ROOT || path.join(__dirname, '..');
const CT = {'.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json'};
async function routeApp(page, {overrides = {}} = {}){
  await page.route('**/*', r => {
    const u = new URL(r.request().url()); if (u.origin !== 'http://app.test') return r.abort();
    if (overrides[u.pathname] !== undefined) return r.fulfill({body: overrides[u.pathname], contentType: CT[path.extname(u.pathname)] || 'text/plain'});
    const f = path.join(ROOT, u.pathname === '/' ? '/index.html' : u.pathname);
    if (!fs.existsSync(f)) return r.fulfill({status: 404, body: 'not found'});
    r.fulfill({body: fs.readFileSync(f), contentType: CT[path.extname(f)] || 'application/octet-stream'});
  });
}
module.exports = {routeApp, ROOT};
