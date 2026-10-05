// Store adapter for the Claude artifact runtime: campaigns and prospects live in the artifact database (private per
// user), Outlook is read through the viewer's Microsoft 365 connector. Stored layout:
//   data/users/<uid>/profile                      {senderEmail, campaignId}
//   data/users/<uid>/<cid>                        campaign settings (type: 'campaign')
//   data/users/<uid>/<cid>/chunks/p<n>            {type: 'chunk', n, items: {<docId>: prospect}}, 40 prospects per chunk
//   data/users/<uid>/<cid>/prospects/<docId>      OLD one-document-per-prospect layout: folded into the chunks and deleted on load
// Documents the store hands back are FROZEN: clone them before adding fields or editing, or the change silently does nothing.
window.createClaudeStore = function(){
  const CHUNK = 40, SERVER = 'Microsoft 365', TOOL = 'outlook_email_search', PERM = 'mcp:' + SERVER;
  const clone = v => JSON.parse(JSON.stringify(v)), sleep = ms => new Promise(r => setTimeout(r, ms)), U = () => window.OutreachUtil;
  let db, user, mcp, perms, uid, log = () => {};
  const base = () => `data/users/${uid}`;
  const chunkOf = {};                                                     // docId → chunk number, for the loaded campaign
  const queue = []; let pumping = false;                                  // one write at a time
  function enqueue(fn){ return new Promise((res, rej) => { queue.push({fn, res, rej}); pump(); }); }
  async function pump(){ if (pumping) return; pumping = true; while (queue.length) { const w = queue.shift(); try { w.res(await w.fn()); } catch (e) { w.rej(e); } } pumping = false; }
  async function retrying(fn){ for (let a = 0; ; a++) { try { return await fn(); } catch (e) { if (a < 5 && e && (e.code === 'resource_exhausted' || e.code === 'unavailable')) { await sleep(300 * (a + 1) + Math.random() * 300); continue; } throw e; } } }
  async function readAll(path){ return (await retrying(() => db.collection(path).limit(1000).get())).docs; }
  function assignChunk(id){ if (chunkOf[id] !== undefined) return chunkOf[id]; const counts = {}; for (const n of Object.values(chunkOf)) counts[n] = (counts[n] || 0) + 1; let n = 0; while ((counts[n] || 0) >= CHUNK) n++; chunkOf[id] = n; return n; }
  function writeChunk(cid, n, items){ const sub = {}; for (const [id, c] of Object.entries(items)) if (chunkOf[id] === n) sub[id] = c; return retrying(() => db.doc(`${base()}/${cid}/chunks/p${n}`).set({type: 'chunk', n, items: sub})); }

  // Fold records of the old layout into `items` (prospects that are missing; sent/reply/bounce dates the chunks lack), save, then
  // delete the old copies in the background. A no-op once they are gone.
  async function mergeLegacy(cid, items){
    let old; try { old = await readAll(`${base()}/${cid}/prospects`); } catch (e) { return null; }
    if (!old.length) return null;
    const byKey = {}; for (const [id, c] of Object.entries(items)) byKey[U().keyOf(c.email)] = id;
    let added = 0, merged = 0; const dirty = new Set();
    for (const d of old) {
      const o = clone(d.data() || {}); const k = U().keyOf(o.email); if (!k || !k.includes('@')) continue;
      const id = byKey[k];
      if (!id) { delete o.chunk; o.log = o.log || {sent: [], replies: [], bounces: []}; o.manual = o.manual || {}; const nid = U().docId(o.email); items[nid] = o; byKey[k] = nid; dirty.add(assignChunk(nid)); added++; continue; }
      const c = items[id]; let ch = false; c.log = c.log || {sent: [], replies: [], bounces: []}; const ol = o.log || {};
      for (const f of ['sent', 'replies', 'bounces']) for (const t of ol[f] || []) if (U().addUnique(c.log[f] = c.log[f] || [], t)) ch = true;
      for (const [f, v] of Object.entries(o.manual || {})) if (v && (c.manual = c.manual || {})[f] == null) { c.manual[f] = v; ch = true; }
      if (ch) { dirty.add(chunkOf[id]); merged++; }
    }
    for (const n of dirty) await enqueue(() => writeChunk(cid, n, items));
    log(`old storage layout: ${old.length} records found · ${added} prospects recovered · ${merged} updated · removing the old copies`);
    (async () => { for (const d of old) for (let a = 0; a < 3; a++) { try { await db.doc(`${base()}/${cid}/prospects/${d.id}`).delete(); break; } catch (e) { await sleep(500 * (a + 1)); } } })();
    return {added, merged};
  }

  function parseItems(result){
    const raw = [], texts = [], out = [], seen = new Set();
    // The connector may return the emails as one text block, as many text blocks (one per email), or as a parsed payload.
    for (const b of (result && result.content) || []) if (b && b.type === 'text' && typeof b.text === 'string') texts.push(b.text);
    if (result && result.payload !== undefined) {
      if (typeof result.payload === 'string') texts.push(result.payload);
      else if (Array.isArray(result.payload)) result.payload.forEach(x => raw.push(x));
      else raw.push(result.payload);
    }
    for (const t of texts) {
      try { const v = JSON.parse(t); (Array.isArray(v) ? v : [v]).forEach(x => raw.push(x)); continue; } catch (e) {}
      let depth = 0, start = -1, inStr = false, escp = false;
      for (let k = 0; k < t.length; k++) { const ch = t[k];
        if (inStr) { if (escp) escp = false; else if (ch === '\\') escp = true; else if (ch === '"') inStr = false; continue; }
        if (ch === '"') inStr = true; else if (ch === '{') { if (depth === 0) start = k; depth++; }
        else if (ch === '}') { depth--; if (depth === 0 && start >= 0) { try { raw.push(JSON.parse(t.slice(start, k + 1))); } catch (e) {} start = -1; } } }
    }
    for (const x of raw) {                                                // dedupe (payload often repeats the first content block); keep pagination tails
      if (!x || typeof x !== 'object') continue;
      const key = x.id || x.uri || (x.nextOffset !== undefined || x.moreResults !== undefined ? 'tail' : null);
      if (key) { if (seen.has(key)) continue; seen.add(key); }
      out.push(x);
    }
    return out;
  }
  function timeoutSignal(ms){ try { return AbortSignal.timeout ? AbortSignal.timeout(ms) : undefined; } catch (e) { return undefined; } }

  const store = {
    kind: 'claude',
    setLogger(fn){ log = fn; },
    async init(){
      if (!(window.claude && window.claude.use)) return {ok: false, reason: 'outside', message: 'This copy of the page is not running inside Claude, so it has no storage or Outlook access. Open Cludo Outreach Desk from your Claude artifacts to use it.'};
      [db, user, mcp, perms] = await Promise.all(['db', 'user', 'mcp', 'permissions'].map(n => window.claude.use(n)));
      if (!db || !user) return {ok: false, reason: 'no_storage', message: 'Storage is not available in this view. Sign in to Claude and open the page from your artifacts list.'};
      let me = {}; try { me = await user.me(); uid = await user.id(); } catch (e) {}
      if (!uid) return {ok: false, reason: 'no_identity', message: 'Could not identify you. Sign in to Claude and reload.'};
      let canWrite = null; try { canWrite = await user.can('data.write'); } catch (e) {}
      store.mail.available = !!mcp;
      return {ok: true, viewer: {id: uid, name: me.name || '', email: String(me.email || '').toLowerCase()}, readOnly: canWrite === false};
    },
    describeStorage(){ return 'Everything is stored in your private space in this app; nobody else at Cludo can see your prospects.'; },
    loadFailureText(e){ return 'Could not load your data (' + (e.code || 'error') + (e.message ? ': ' + e.message : '') + '). ' + (e.code === 'invalid_argument' ? 'Ask the owner to share this app with you as a Contributor.' : 'Reload to try again.'); },
    async loadIndex(){
      const read = async () => { const out = {campaigns: {}, profile: {}}; for (const d of await readAll(base())) { const v = d.data() || {}; if (v.type === 'campaign') out.campaigns[d.id] = clone(v); else if (d.id === 'profile') out.profile = clone(v); } return out; };
      let idx = await read();
      if (!Object.keys(idx.campaigns).length) { await sleep(1500); idx = await read(); }   // ask twice before concluding there is nothing: never create a fresh campaign over an answer that was still loading
      return idx;
    },
    saveCampaign(id, camp){ return enqueue(() => retrying(() => db.doc(`${base()}/${id}`).set(camp))); },
    patchCampaign(id, patch){ return enqueue(() => retrying(() => db.doc(`${base()}/${id}`).update(patch))); },
    async deleteCampaign(id){ for (const d of await readAll(`${base()}/${id}/chunks`)) await enqueue(() => db.doc(`${base()}/${id}/chunks/${d.id}`).delete()); await enqueue(() => db.doc(`${base()}/${id}`).delete()); },
    saveProfile(p){ return enqueue(() => retrying(() => db.doc(`${base()}/profile`).set(p))); },
    async loadProspects(cid){
      for (const k of Object.keys(chunkOf)) delete chunkOf[k];
      const items = {};
      for (const d of await readAll(`${base()}/${cid}/chunks`)) { const v = clone(d.data() || {}); const n = typeof v.n === 'number' ? v.n : +String(d.id).replace(/^p/, ''); for (const [id, c] of Object.entries(v.items || {})) { items[id] = c; chunkOf[id] = n; } }
      const recovered = await mergeLegacy(cid, items);
      return {items, recovered};
    },
    async saveProspects(cid, items, changedIds, removedIds){
      const dirty = new Set();
      for (const id of removedIds || []) if (chunkOf[id] !== undefined) { dirty.add(chunkOf[id]); delete chunkOf[id]; }
      for (const id of changedIds || []) if (items[id]) dirty.add(assignChunk(id));
      await Promise.all([...dirty].map(n => enqueue(() => writeChunk(cid, n, items))));
      return null;                                                        // nothing to reconcile: what was written is the local state
    },
    mail: {
      available: false, label: 'Outlook',
      async access(){ if (!perms) return 'unknown'; try { return await Promise.race([perms.state(PERM), sleep(3000).then(() => 'unknown')]); } catch (e) { return 'unknown'; } },
      async request(){ if (!perms) return; try { await Promise.race([perms.request([PERM]), sleep(8000)]); } catch (e) {} },
      async scan(folder, sinceYmd, onItem, onProgress){
        let offset = 0, pages = 0, n = 0, total = 0;
        while (pages < 80) {
          onProgress(folder, n, total, pages);
          let res;
          try { res = await mcp.callTool(SERVER, TOOL, {folderName: folder, afterDateTime: sinceYmd, order: 'newest', limit: 25, offset}, {cache: false, signal: timeoutSignal(90000)}); }
          catch (e) { log(`${folder} page ${pages + 1}: error ${e && e.code} ${e && e.message ? e.message.slice(0, 160) : ''}`); throw e; }
          const items = parseItems(res), tail = items.find(x => x && (x.nextOffset !== undefined || x.moreResults !== undefined)) || {};
          if (tail.totalResultCount) total = tail.totalResultCount;
          let next = null, onPage = 0;
          for (const it of items) { if (it && it.nextOffset !== undefined) { next = it.nextOffset; continue; } if (!it || !it.sender) continue; n++; onPage++; onItem({sender: it.sender, recipients: it.recipients || [], sentDateTime: it.sentDateTime, receivedDateTime: it.receivedDateTime, subject: it.subject || '', preview: it.summary || '', thread: it.conversationId || null}); }
          if ((next === null || next === undefined) && tail.nextOffset === undefined && onPage >= 25) next = offset + onPage;   // no pagination tail returned: keep going while pages are full
          log(`${folder} page ${pages + 1}: ${onPage} emails, next offset ${next}`);
          pages++; if (next === null || next === undefined || onPage === 0) break; offset = next;
        }
        onProgress(folder, n, total, pages, true);
        return n;
      },
      errorText(e){
        const code = e && e.code;
        return code === 'needs_reauth' ? 'Outlook needs to be reconnected: claude.ai → Settings → Connectors → Microsoft 365, then sync again.'
          : code === 'server_not_connected' || code === 'server_not_found' ? 'Add the Microsoft 365 connector in claude.ai → Settings → Connectors, then reload.'
          : code === 'selection_required' ? 'You have more than one Microsoft 365 connector; choose one when the page asks, then sync again.'
          : code === 'not_in_manifest' || code === 'consent_required' || code === 'not_granted' ? 'This page is not allowed to read Outlook yet. Press the button and accept the prompt.'
          : code === 'blocked_by_policy' || code === 'approval_required' ? 'Your organisation’s policy blocks this page from reading Outlook.'
          : code === 'cancelled' ? 'Outlook took too long to answer. Try again.'
          : code === 'server_unavailable' || code === 'upstream_error' || code === 'rate_limited' ? 'Outlook did not answer. Wait a moment and try again.'
          : 'Sync failed (' + (code || 'unknown') + '): ' + (e && e.message ? e.message : '');
      }
    }
  };
  return store;
};
