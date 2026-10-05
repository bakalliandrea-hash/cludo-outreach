// Mock of the standalone runtime: a fake `msal` namespace and a fetch() that answers for graph.microsoft.com with an
// in-memory OneDrive app folder (eTags, 412 on If-Match mismatch) and two mail folders with @odata.nextLink paging.
window.__drive = window.__drive || {};            // name → {content, etag}
window.__mail = window.__mail || {sentitems: [], inbox: []};
window.__graphCalls = []; window.__loginCalls = 0; window.__logoutCalls = 0;
window.__signedIn = window.__signedIn || false;
let etagSeq = 1; const newEtag = () => 'W/"' + (etagSeq++) + '"';
window.msal = {
  InteractionRequiredAuthError: class extends Error {},
  PublicClientApplication: class {
    constructor(c){ this.cfg = c; }
    async initialize(){}
    async handleRedirectPromise(){ return null; }
    getAllAccounts(){ return window.__signedIn ? [{homeAccountId: 'acct-1', name: 'Tester', username: 'andrea.bakalli@cludo.com'}] : []; }
    setActiveAccount(){}
    async acquireTokenSilent(){ return {accessToken: 'test-token'}; }
    async loginRedirect(){ window.__loginCalls++; }
    async logoutRedirect(){ window.__logoutCalls++; }
  }
};
const realFetch = window.fetch.bind(window);
const json = (status, body, headers = {}) => new Response(body === undefined ? null : JSON.stringify(body), {status, headers: Object.assign({'Content-Type': 'application/json'}, headers)});
window.fetch = async (input, init = {}) => {
  const url = typeof input === 'string' ? input : input.url, m = init.method || 'GET';
  if (url.startsWith('http://app.test/__dl/')) { const name = decodeURIComponent(url.split('/__dl/')[1]); const f = window.__drive[name]; return f ? json(200, f.content) : json(404, {}); }
  if (!url.startsWith('https://graph.microsoft.com/')) return realFetch(input, init);
  const u = new URL(url); window.__graphCalls.push({m, path: u.pathname + u.search, ifMatch: init.headers && init.headers['If-Match']});
  if (!(init.headers && init.headers.Authorization === 'Bearer test-token')) return json(401, {error: {code: 'InvalidAuthenticationToken', message: 'no token'}});
  const p = decodeURIComponent(u.pathname);
  if (p === '/v1.0/me') return json(200, {displayName: 'Tester', mail: 'Andrea.Bakalli@cludo.com', userPrincipalName: 'andrea.bakalli@cludo.com'});
  let mm = p.match(/^\/v1\.0\/me\/drive\/special\/approot:\/([^:]+)(:\/content)?$/);
  if (mm) {
    const name = mm[1], isContent = !!mm[2], f = window.__drive[name];
    if (m === 'GET') { if (!f) return json(404, {error: {code: 'itemNotFound', message: 'not found'}}); return isContent ? json(200, f.content) : json(200, {name, eTag: f.etag, '@microsoft.graph.downloadUrl': 'http://app.test/__dl/' + encodeURIComponent(name)}); }
    if (m === 'PUT') { const im = init.headers['If-Match']; if (f && im && im !== f.etag) return json(412, {error: {code: 'resourceModified', message: 'etag mismatch'}}); const etag = newEtag(); window.__drive[name] = {content: JSON.parse(init.body), etag}; return json(f ? 200 : 201, {name, eTag: etag}); }
    if (m === 'DELETE') { if (!f) return json(404, {error: {code: 'itemNotFound', message: 'not found'}}); delete window.__drive[name]; return new Response(null, {status: 204}); }
  }
  mm = p.match(/^\/v1\.0\/me\/mailFolders\/(sentitems|inbox)\/messages$/);
  if (mm) {
    const folder = mm[1], filter = u.searchParams.get('$filter') || '', top = +(u.searchParams.get('$top') || 10), skip = +(u.searchParams.get('$skip') || 0);
    const fm = filter.match(/^(\w+) ge (\S+)$/); if (!fm) return json(400, {error: {code: 'BadRequest', message: 'unsupported filter ' + filter}});
    const all = (window.__mail[folder] || []).filter(x => x[fm[1]] >= fm[2]).sort((a, b) => b[fm[1]].localeCompare(a[fm[1]]));
    const page = all.slice(skip, skip + top), out = {value: page};
    if (skip + top < all.length) { const next = new URL(url); next.searchParams.set('$skip', skip + top); out['@odata.nextLink'] = next.toString(); }
    return json(200, out);
  }
  return json(404, {error: {code: 'NotFound', message: 'unmocked ' + m + ' ' + p}});
};
