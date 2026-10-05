// Store adapter for the standalone site (Netlify): Microsoft sign-in with MSAL (authorization code + PKCE, redirect flow),
// Outlook read through Microsoft Graph, and the data kept as JSON files in the app's own OneDrive folder
// (Apps/Cludo Outreach Desk, reachable only with Files.ReadWrite.AppFolder):
//   index.json                 {campaigns: {<cid>: settings}, profile: {senderEmail, campaignId}}
//   prospects-<cid>.json       {items: {<docId>: prospect}}
// Writes carry the file's eTag (If-Match); a 412 means another window saved first, and the change is merged onto the fresh copy.
window.createGraphStore = function(cfg){
  const G = 'https://graph.microsoft.com/v1.0', APP = '/me/drive/special/approot';
  const clone = v => JSON.parse(JSON.stringify(v)), sleep = ms => new Promise(r => setTimeout(r, ms));
  let pca = null, account = null, log = () => {}, index = null;
  const etags = {};
  const redirectUri = () => location.origin + location.pathname;

  function loadMsal(){
    if (window.msal) return Promise.resolve();
    return new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'vendor/msal-browser.min.js'; s.onload = res; s.onerror = () => rej(new Error('Could not load the Microsoft sign-in library')); document.head.appendChild(s); });
  }
  async function token(){
    try { return (await pca.acquireTokenSilent({scopes: cfg.scopes, account})).accessToken; }
    catch (e) {
      const code = (e && e.errorCode) || '';
      if ((window.msal && e instanceof msal.InteractionRequiredAuthError) || /interaction_required|consent_required|login_required|no_tokens_found|monitor_window_timeout/.test(code)) { await pca.acquireTokenRedirect({scopes: cfg.scopes, account}); return new Promise(() => {}); }   // the page is leaving for the sign-in screen
      throw e;
    }
  }
  async function graph(path, opts = {}, attempt = 0){
    const t = await token(), url = path.startsWith('http') ? path : G + path;
    const r = await fetch(url, Object.assign({}, opts, {headers: Object.assign({Authorization: 'Bearer ' + t}, opts.headers || {})}));
    if ((r.status === 429 || r.status === 503 || r.status === 504) && attempt < 4) { const ra = +r.headers.get('Retry-After') || (attempt + 1) * 2; log(`Graph ${r.status}; retrying in ${ra}s`); await sleep(ra * 1000); return graph(path, opts, attempt + 1); }
    return r;
  }
  async function gerr(r){ let msg = 'HTTP ' + r.status; try { const j = await r.json(); if (j.error) msg = j.error.code + ': ' + j.error.message; } catch (e) {} const e = new Error(msg); e.status = r.status; e.code = r.status === 401 || r.status === 403 ? 'not_granted' : 'graph_error'; return e; }
  async function readJsonFile(name){
    const meta = await graph(`${APP}:/${name}`);
    if (meta.status === 404) { delete etags[name]; return {data: null}; }
    if (!meta.ok) throw await gerr(meta);
    const m = await meta.json(); etags[name] = m.eTag;
    let r = null; try { if (m['@microsoft.graph.downloadUrl']) r = await fetch(m['@microsoft.graph.downloadUrl']); } catch (e) {}
    if (!r || !r.ok) { r = await graph(`${APP}:/${name}:/content`); if (!r.ok) throw await gerr(r); }
    return {data: await r.json()};
  }
  async function writeJsonFile(name, data){
    const h = {'Content-Type': 'application/json'}; if (etags[name]) h['If-Match'] = etags[name];
    const r = await graph(`${APP}:/${name}:/content`, {method: 'PUT', headers: h, body: JSON.stringify(data)});
    if (r.status === 412) { const e = new Error('The file changed in another window'); e.code = 'conflict'; throw e; }
    if (!r.ok) throw await gerr(r);
    const m = await r.json(); etags[name] = m.eTag;
  }
  async function saveIndex(mutate){                                          // apply our change to the freshest copy, up to three tries
    for (let a = 0; a < 3; a++) {
      mutate(index);
      try { await writeJsonFile('index.json', index); return; }
      catch (e) { if (e.code !== 'conflict') throw e; index = (await readJsonFile('index.json')).data || {campaigns: {}, profile: {}}; log('settings were saved from another window; merging'); }
    }
    throw new Error('Could not save settings: the file keeps changing in another window');
  }
  const addr = x => x && x.emailAddress ? x.emailAddress.address || x.emailAddress.name || '' : '';

  const store = {
    kind: 'graph',
    setLogger(fn){ log = fn; },
    async init(){
      if (!cfg || !cfg.clientId || !cfg.tenantId) return {ok: false, reason: 'not_configured'};
      await loadMsal();
      pca = new msal.PublicClientApplication({auth: {clientId: cfg.clientId, authority: 'https://login.microsoftonline.com/' + cfg.tenantId, redirectUri: redirectUri(), postLogoutRedirectUri: redirectUri()}, cache: {cacheLocation: 'localStorage'}});
      await pca.initialize();
      let resp = null;
      try { resp = await pca.handleRedirectPromise(); }
      catch (e) { const code = (e && e.errorCode) || ''; log('sign-in failed: ' + code + ' ' + (e && e.errorMessage || e && e.message || '')); return {ok: false, reason: 'signin_failed', detail: /AADSTS65001|consent|admin approval|65004/i.test((e && e.errorMessage) || code) ? 'Cludo has not approved this app yet (admin consent). Ask IT to grant consent in Microsoft Entra ID, then sign in again.' : (e && (e.errorMessage || e.message)) || code}; }
      account = (resp && resp.account) || pca.getAllAccounts()[0] || null;
      if (!account) return {ok: false, reason: 'signed_out'};
      pca.setActiveAccount(account);
      let me = {}; try { const r = await graph('/me?$select=displayName,mail,userPrincipalName'); if (r.ok) me = await r.json(); } catch (e) {}
      return {ok: true, viewer: {id: account.homeAccountId, name: me.displayName || account.name || '', email: String(me.mail || me.userPrincipalName || account.username || '').toLowerCase()}};
    },
    signIn(){ return pca.loginRedirect({scopes: cfg.scopes, prompt: 'select_account'}); },
    signOut(){ return pca.logoutRedirect({account}); },
    describeStorage(){ return 'The data is stored as files in your own OneDrive, in the folder Apps/Cludo Outreach Desk. Nobody else can see it unless you share that folder.'; },
    loadFailureText(e){ return 'Could not load your data' + (e && e.message ? ' (' + e.message + ')' : '') + '. ' + (e && e.code === 'not_granted' ? 'The app lacks permission to read your OneDrive or mail; ask IT to grant consent in Microsoft Entra ID.' : 'Reload to try again.'); },
    async loadIndex(){ index = (await readJsonFile('index.json')).data || {campaigns: {}, profile: {}}; index.campaigns = index.campaigns || {}; index.profile = index.profile || {}; return clone(index); },
    saveCampaign(id, camp){ return saveIndex(ix => { ix.campaigns[id] = clone(camp); }); },
    patchCampaign(id, patch){ return saveIndex(ix => { Object.assign(ix.campaigns[id] = ix.campaigns[id] || {}, clone(patch)); }); },
    async deleteCampaign(id){ await saveIndex(ix => { delete ix.campaigns[id]; }); const r = await graph(`${APP}:/prospects-${id}.json`, {method: 'DELETE'}); if (!r.ok && r.status !== 404) throw await gerr(r); delete etags[`prospects-${id}.json`]; },
    saveProfile(p){ return saveIndex(ix => { ix.profile = clone(p); }); },
    async loadProspects(cid){ const {data} = await readJsonFile(`prospects-${cid}.json`); return {items: (data && data.items) || {}, recovered: null}; },
    async saveProspects(cid, items, changedIds, removedIds){
      const name = `prospects-${cid}.json`;
      try { await writeJsonFile(name, {items}); return null; }
      catch (e) {
        if (e.code !== 'conflict') throw e;
        const fresh = (await readJsonFile(name)).data, merged = (fresh && fresh.items) || {};
        for (const id of changedIds || []) if (items[id]) merged[id] = items[id];
        for (const id of removedIds || []) delete merged[id];
        await writeJsonFile(name, {items: merged});
        log('another window had saved first; merged the changes');
        return merged;
      }
    },
    mail: {
      available: true, label: 'Outlook',
      async access(){ return 'granted'; },                                    // consent is part of sign-in
      async request(){},
      async scan(folder, sinceYmd, onItem, onProgress){
        const f = folder === 'Sent Items' ? 'sentitems' : 'inbox', dateField = f === 'sentitems' ? 'sentDateTime' : 'receivedDateTime';
        let url = `/me/mailFolders/${f}/messages?$filter=${encodeURIComponent(`${dateField} ge ${sinceYmd}T00:00:00Z`)}&$orderby=${encodeURIComponent(dateField + ' desc')}&$top=50&$select=id,conversationId,subject,sentDateTime,receivedDateTime,from,sender,toRecipients,ccRecipients,bodyPreview`;
        let n = 0, pages = 0;
        while (url && pages < 200) {
          onProgress(folder, n, 0, pages);
          const r = await graph(url); if (!r.ok) throw await gerr(r);
          const j = await r.json();
          for (const m of j.value || []) { n++; onItem({sender: addr(m.from) || addr(m.sender), recipients: [...(m.toRecipients || []), ...(m.ccRecipients || [])].map(addr), sentDateTime: m.sentDateTime, receivedDateTime: m.receivedDateTime, subject: m.subject || '', preview: m.bodyPreview || '', thread: m.conversationId || null}); }
          pages++; log(`${folder} page ${pages}: ${(j.value || []).length} emails`);
          url = j['@odata.nextLink'] || null;
        }
        onProgress(folder, n, 0, pages, true);
        return n;
      },
      errorText(e){ return e && e.code === 'not_granted' ? 'Outlook refused the request. Sign out and in again; if it persists, ask IT to grant the app consent in Microsoft Entra ID.' : 'Sync failed: ' + ((e && e.message) || 'unknown error'); }
    }
  };
  return store;
};
