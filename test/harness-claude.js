// Mock of the Claude artifact runtime: in-memory db seeded from a JSON store, fake user, fake Outlook connector.
window.__writes = [];
window.__mail = window.__mail || {'Sent Items': [], 'Inbox': []};
window.__calls = [];
const STORE = window.__store;
const segs = p => p.split('/');
const freeze = o => { if (o && typeof o === 'object') { Object.values(o).forEach(freeze); Object.freeze(o); } return o; };
const mk = (path, v) => ({id: segs(path).pop(), exists: v !== undefined, data: () => v === undefined ? undefined : freeze(JSON.parse(JSON.stringify(v))), metadata: {fromCache: false, hasPendingWrites: false}});
const db = {
  doc: path => ({
    get: async () => mk(path, STORE[path]),
    set: async v => { if (JSON.stringify(v).length > 256*1024) throw Object.assign(new Error('too big'), {code:'invalid_argument'}); STORE[path] = JSON.parse(JSON.stringify(v)); __writes.push(['set', path]); },
    update: async v => { if (STORE[path] === undefined) throw Object.assign(new Error('missing'), {code:'invalid_argument'}); Object.assign(STORE[path], JSON.parse(JSON.stringify(v))); __writes.push(['update', path]); },
    delete: async () => { delete STORE[path]; __writes.push(['delete', path]); }
  }),
  collection: path => ({ limit: n => ({ get: async () => ({ docs: Object.keys(STORE).filter(k => k.startsWith(path + '/') && segs(k).length === segs(path).length + 1).slice(0, n).map(k => mk(k, STORE[k])), metadata: {fromCache: false, hasPendingWrites: false} }) }) })
};
const user = { me: async () => ({name: 'Tester', email: 'andrea.bakalli@cludo.com'}), id: async () => 'u_test', can: async () => true };
const mcp = { callTool: async (server, tool, args) => {
  __calls.push({server, tool, args});
  const all = (__mail[args.folderName] || []).filter(m => !args.afterDateTime || m.receivedDateTime >= args.afterDateTime);
  const off = args.offset || 0, page = all.slice(off, off + args.limit);
  const content = page.map((m, k) => ({type: 'text', text: JSON.stringify(Object.assign({id: 'id' + (off + k) + args.folderName, offset: off + k}, m))}));
  content.push({type: 'text', text: JSON.stringify(off + args.limit < all.length ? {moreResults: true, nextOffset: off + args.limit, totalResultCount: all.length} : {totalResultCount: all.length})});
  return {content};
}};
const perms = { state: async () => 'granted', request: async () => true };
window.claude = { use: async name => ({db, user, mcp, permissions: perms})[name] };
