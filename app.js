// Cludo Outreach Desk — shared application logic. Storage and Outlook access come from a store adapter chosen at start:
// store-claude.js inside the Claude artifact runtime, store-graph.js on the standalone site (Microsoft sign-in, OneDrive, Graph).

// ---------- constants ----------
const STEP_NAME = {1: 'Email 1', 2: 'Follow-up 1', 3: 'Follow-up 2'};
const DEFAULT_HOLIDAYS = ['2026-10-12','2026-11-01','2026-12-06','2026-12-08','2026-12-25','2027-01-01','2027-01-06','2027-03-26','2027-05-01','2027-08-15','2027-10-12','2027-11-01','2027-12-06','2027-12-08','2027-12-25'];
const T1_BODY = ['Hola {first_name},',
  'He visto que te ocupas de {area} en {organisation} y te escribo en relación con el buscador de vuestra web.',
  'Cuando los estudiantes y el personal no encuentran información que ya está publicada en la web, acaban acudiendo a la secretaría o al CAU. Un buscador más eficaz les ayuda a resolver sus dudas de forma autónoma y reduce las consultas repetitivas que llegan a estos servicios.',
  'En Cludo trabajamos con varias universidades europeas, entre ellas King\'s College London, Stockholm University y Copenhagen University, para mejorar la búsqueda en sus webs.',
  'En los últimos meses hemos desarrollado una herramienta que analiza el buscador de vuestra web desde el punto de vista de los usuarios. El informe muestra, con ejemplos concretos, qué búsquedas llevan a la información correcta, cuáles presentan problemas y qué podéis mejorar, incluso por vuestra cuenta y sin depender de nosotros.',
  '¿Te interesaría una breve presentación de nuestra solución, tomando como referencia cómo otras universidades han mejorado la búsqueda en su web? El análisis es totalmente gratuito y sin ningún compromiso.',
  'Si no eres tú la persona que lleva este tema, ¿me podrías indicar a quién dirigirme?'].join('\n\n');
const OLD_T1_BODY = 'Hola {first_name},\n\nHe visto que te ocupas de {area} en {organisation} y te escribo en relación con el buscador de vuestra web.\n\n…';   // placeholder shipped by the first version; replaced on load
const DEFAULT_CAMPAIGN = {
  name: 'Spain universities', createdAt: '', waitDays: [2, 3], holidays: DEFAULT_HOLIDAYS, campaignStart: '2026-10-05',
  t: {
    1: {subject: 'Experiencia de búsqueda en la web de {organisation}', body: T1_BODY},
    2: {subject: 'Re: Experiencia de búsqueda en la web de {organisation}', body: 'Hola {first_name},\n\nTe escribo respecto al correo que te envié hace unos días sobre el buscador de vuestra web, por si aún no lo has visto.\n\nCuando tengas un momento, me gustaría saber qué te parece.'},
    3: {subject: 'Re: Experiencia de búsqueda en la web de {organisation}', body: 'Hola {first_name},\n\nQuería hacer un último intento. Puedo prepararte un análisis gratuito del buscador de vuestra web: un resumen concreto de dónde funciona bien la búsqueda y dónde se podría mejorar.\n\nCludo es una herramienta que ayuda a universidades e instituciones a hacer su buscador más conversacional y más fácil de gestionar, incluidas funcionalidades de IA. Ya trabajamos con varias universidades europeas y estamos intentando entender mejor cómo abordan hoy este tema las universidades españolas.\n\nSi te parece, me gustaría hacer una breve llamada: nada comercial, sino un intercambio en el que comentemos los resultados del análisis, vosotros nos deis feedback sobre la herramienta que estamos construyendo y nosotros os enseñemos cómo funciona nuestro buscador. Algo útil para ambas partes.\n\nDime si crees que puede tener sentido.'}
  }
};
const SAME_SEND_MS = 10 * 60000;                      // sends to one prospect this close together are one email (a resend), not a follow-up
const BACKUP_TYPE = 'cludo-outreach-backup';
const CONTACT_FIELDS = ['first_name', 'last_name', 'title', 'area', 'organisation', 'email', 'website', 'linkedin', 'city', 'subject', 'body', 'review'];

// ---------- runtime state ----------
let store = null, viewer = null, SENDER = '', profile = {};
let campaigns = {}, campaignId = null, camp = null;      // campaigns[id] = {...DEFAULT_CAMPAIGN}
let P = {}, loaded = false;                              // prospects of the current campaign keyed by email; loaded: they were read back successfully
let tab = 'todo', toastT, showHandled = false, syncing = false;
const open = new Set();

// ---------- helpers ----------
const $ = s => document.querySelector(s);
const list = $('#list'), toast = $('#toast');
function say(msg){ toast.textContent = msg; toast.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => toast.classList.remove('show'), 2200); }
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c])); }
function enc(s){ return encodeURIComponent(s).replace(/[!'()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase()); }
function clone(v){ return JSON.parse(JSON.stringify(v)); }
function keyOf(email){
  if (email && typeof email === 'object') email = email.address || (email.emailAddress && email.emailAddress.address) || email.email || email.name || '';
  const s = String(email || '').trim().toLowerCase();
  const m = s.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/);   // "Name <x@y.z>" → x@y.z
  return m ? m[0] : s;
}
function docId(email){ return keyOf(email).replace(/[^a-z0-9]/g, '_').slice(0, 180); }   // document ids: safe characters only
function addUnique(arr, iso){ if (iso && !arr.includes(iso)) { arr.push(iso); arr.sort(); return true; } return false; }
window.OutreachUtil = {keyOf, docId, addUnique, clone};
function setStatus(msg, isErr){ const el = $('#syncStatus'); el.textContent = msg; el.className = isErr ? 'err' : ''; }
const LOG = [];
function logLine(msg){ LOG.push(new Date().toLocaleTimeString('en-GB') + '  ' + msg); if (LOG.length > 200) LOG.shift(); const el = $('#syncLog'); el.textContent = LOG.join('\n'); el.scrollTop = el.scrollHeight; $('#logToggle').hidden = false; }
$('#logToggle').addEventListener('click', () => { const el = $('#syncLog'); el.hidden = !el.hidden; });
const prog = {phase: 0, folders: 2};
function showProgress(pct, label){ const bar = $('#syncBar'); bar.hidden = pct == null; if (pct != null) { $('#syncFill').style.width = Math.max(2, Math.min(100, pct)) + '%'; if (label) setStatus(label); } }
function folderProgress(folder, done, total, pages, finished){
  // each folder is half the bar; inside a folder use the server's total when it gives one, else assume 10 pages
  const within = finished ? 1 : total ? Math.min(1, done / total) : Math.min(0.95, pages / 10);
  const pct = Math.round(((prog.phase + within) / prog.folders) * 100);
  showProgress(pct, `Syncing ${folder}… ${done}${total ? ' of ' + total : ''} emails read · ${pct}%`);
  if (finished) prog.phase++;
}

// ---------- dates ----------
function ymd(d){ const x = new Date(d); return x.getFullYear() + '-' + String(x.getMonth() + 1).padStart(2, '0') + '-' + String(x.getDate()).padStart(2, '0'); }
function todayYmd(){ return ymd(new Date()); }
function isWorkDay(y){ const d = new Date(y + 'T12:00:00'); const w = d.getDay(); return w !== 0 && w !== 6 && !(camp.holidays || []).includes(y); }
function addWorkDays(fromIso, n){ let d = new Date(ymd(fromIso) + 'T12:00:00'), left = n; while (left > 0) { d.setDate(d.getDate() + 1); if (isWorkDay(ymd(d))) left--; } return ymd(d); }
function fmt(y){ if (!y) return ''; const d = new Date(y.length === 10 ? y + 'T12:00:00' : y); return isNaN(d) ? '' : d.toLocaleDateString('en-GB', {day: 'numeric', month: 'short'}); }
function dayDiff(a, b){ return Math.round((new Date(a + 'T12:00:00') - new Date(b + 'T12:00:00')) / 86400000); }

// ---------- derived status per prospect ----------
function info(c){
  const L = c.log || {}, M = c.manual || {};
  let sent = [...new Set([...(L.sent || []), ...(M.sent || [])])].sort();
  const once = []; for (const t of sent) if (!once.length || new Date(t) - new Date(once[once.length - 1]) > SAME_SEND_MS) once.push(t); sent = once;
  const bounces = [...(L.bounces || [])].sort(), voided = new Set();
  for (const b of bounces) { const cand = sent.filter(x => !voided.has(x) && x <= b && (new Date(b) - new Date(x)) < 2 * 86400000); if (cand.length) voided.add(cand[cand.length - 1]); }
  sent = sent.filter(x => !voided.has(x));
  let replies = (L.replies || []).filter(r => M.ignoreReplies !== true && !(M.ignoreRepliesBefore && r <= M.ignoreRepliesBefore));   // "not a real reply" ignores what had arrived by then, not what comes later
  if (M.replied) replies.push(M.replied);
  replies = [...new Set(replies)].sort();
  const n = Math.min(sent.length, 3), lastSent = sent[sent.length - 1] || null, lastReply = replies[replies.length - 1] || null;
  const o = {sent, replies, n, lastSent, lastReply, lastBounce: bounces[bounces.length - 1] || null, bounced: voided.size || (bounces.length && !sent.length ? bounces.length : 0), step: n + 1, due: null, status: 'todo'};
  if (M.handled) { o.status = 'handled'; return o; }
  if (lastReply) { o.status = (lastSent && lastSent > lastReply) ? 'handled' : 'replied'; return o; }
  if (n === 0) return o;
  if (n >= 3) { o.status = 'finished'; return o; }
  o.due = addWorkDays(lastSent, (camp.waitDays || [2, 3])[n - 1]);
  o.status = todayYmd() >= o.due ? 'due' : 'waiting';
  return o;
}
// The {area} placeholder: the imported value, else a guess from the job title, else a neutral phrase (and the card asks for review).
function guessArea(c){
  const t = String(c.title || '').toLowerCase();
  if (/\bweb|digital|online|\bseo\b|\bux\b|contenid|content|portal|\bcms\b/.test(t)) return 'la web';
  if (/comunicaci|communication|marketing|prensa|press|relaciones|publicidad|brand/.test(t)) return 'la comunicación';
  if (/\bit\b|\btic\b|inform[aá]tic|sistemas|tecnolog|technolog|\bcio\b|\bcau\b|soporte|support|desarroll|develop|software|\bdatos\b|\bdata\b/.test(t)) return 'los sistemas informáticos';
  return null;
}
function areaFor(c){ return c.area || guessArea(c) || 'la comunicación y la web'; }
function needsReview(c){ return c.review === 'yes' || (!c.area && !guessArea(c)); }
function fill(tpl, c){ return String(tpl || '').replace(/\{(first_name|last_name|organisation|area|title|email)\}/g, (m, k) => k === 'area' ? areaFor(c) : (c[k] || '')); }
function subjectFor(c, step){ if (step === 1 && c.subject) return c.subject; const t = (camp.t || {})[step] || {}; if (step > 1 && c.subject && !t.subject) return 'Re: ' + c.subject; return fill(t.subject, c) || (step > 1 && c.subject ? 'Re: ' + c.subject : ''); }
function bodyFor(c, step){ if (step === 1 && c.body) return c.body; return fill(((camp.t || {})[step] || {}).body, c); }
function links(c, subject, body){
  return { owa: 'https://outlook.office.com/mail/deeplink/compose?to=' + enc(c.email) + '&subject=' + enc(subject) + '&body=' + enc(body) + (SENDER ? '&login_hint=' + enc(SENDER) : ''),
           mailto: 'mailto:' + c.email + '?subject=' + enc(subject) + '&body=' + enc(body) };
}
async function copy(text, label){
  try { await navigator.clipboard.writeText(text); say(label + ' copied'); }
  catch (e) { const ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); say(label + ' copied'); } catch (e2) { say('Select and copy manually'); } document.body.removeChild(ta); }
}

// ---------- rendering ----------
const byOrg = (a, b) => String(a.c.organisation).localeCompare(String(b.c.organisation), 'es', {sensitivity: 'base'}) || String(a.c.last_name).localeCompare(String(b.c.last_name), 'es', {sensitivity: 'base'});
const todoOrder = (a, b) => ((b.i.bounced ? 1 : 0) - (a.i.bounced ? 1 : 0)) || byOrg(a, b);
const TAB_ORDER = {todo: todoOrder, due: byOrg, waiting: byOrg, replied: byOrg, finished: byOrg};
const TAB_LABEL = {todo: 'To contact', due: 'Follow-up due', waiting: 'Waiting', replied: 'Replied', finished: 'Finished'};
const EMPTY = {todo: 'Everyone in this campaign has received the first email.', due: 'No follow-ups due today.', waiting: 'Nobody is waiting on a follow-up.',
  replied: 'No replies yet. Replies are detected from your Inbox on sync.', finished: 'Nobody has completed the three-email sequence yet.'};
const HINT = {todo: 'Haven’t received email 1 yet', due: 'Follow-up date reached', waiting: 'Sent; next follow-up not due yet',
  replied: 'They wrote back; answer them in Outlook and they leave this page', finished: 'Three emails sent, no reply'};

function pillsFor(c, i){
  const p = [];
  if (i.status === 'due') p.push(`<span class="pill due">${esc(STEP_NAME[i.step])} due ${dayDiff(todayYmd(), i.due) > 0 ? esc(fmt(i.due)) : 'today'}</span>`);
  if (i.status === 'waiting') p.push(`<span class="pill mute">${esc(STEP_NAME[i.step])} due ${esc(fmt(i.due))}</span>`);
  if (i.status === 'replied') p.push(`<span class="pill ok">Replied ${esc(fmt(i.lastReply))}</span>`);
  if (i.status === 'finished') p.push(`<span class="pill mute">3 sent · no reply</span>`);
  if (i.status === 'handled') p.push(`<span class="pill ok">Answered ${esc(fmt(i.lastSent))}</span>`);
  if (i.n > 0 && i.status !== 'handled') p.push(`<span class="pill ok">${i.n} sent</span>`);
  const pd = c.pending;
  if (pd && pd.step === i.step && i.status !== 'replied' && i.status !== 'handled') p.push(`<span class="pill pend">Draft opened ${esc(fmt(pd.at))}</span>`);
  if (i.bounced && i.status !== 'handled') p.push(`<span class="pill due">Delivery failed ${esc(fmt(i.lastBounce))} · resend</span>`);
  if (needsReview(c) && i.n === 0) p.push('<span class="pill rev">Review área</span>');
  if (i.n === 0 && c.area) p.push(`<span class="pill">${esc(c.area)}</span>`);
  return p.join('');
}
function timelineHTML(c, i){
  const items = [];
  for (let s = 1; s <= 3; s++) {
    const at = i.sent[s - 1];
    if (at) items.push(`<li class="done">${STEP_NAME[s]} sent ${esc(fmt(at))}</li>`);
    else if (s === i.step && i.due) items.push(`<li class="next ${i.status === 'due' ? 'late' : ''}">${STEP_NAME[s]} due ${esc(fmt(i.due))}</li>`);
    else if (s === 1) items.push(`<li class="next">${STEP_NAME[1]} not sent yet</li>`);
    else items.push(`<li>${STEP_NAME[s]}</li>`);
  }
  if (i.lastReply) items.push(`<li class="reply">Replied ${esc(fmt(i.lastReply))}</li>`);
  if (i.bounced) items.push(`<li class="late">${i.bounced} email${i.bounced > 1 ? 's' : ''} not delivered (bounced ${esc(fmt(i.lastBounce))}); it does not count as sent</li>`);
  return `<ul class="timeline">${items.join('')}</ul>`;
}
function cardHTML(c){
  const i = info(c), isOpen = open.has(c.key);
  return `<div class="card ${i.status === 'due' ? 'due' : ''}" data-key="${esc(c.key)}">
    <div class="head" tabindex="0" role="button" aria-expanded="${isOpen}">
      <div class="meta"><span class="name">${esc(c.first_name)} ${esc(c.last_name)}</span> <span class="org">· ${esc(c.organisation)}</span><br>${esc(c.title)}</div>
      <div class="pills row">${pillsFor(c, i)}</div>
    </div>
    ${isOpen ? bodyHTML(c, i) : ''}
  </div>`;
}
function bodyHTML(c, i){
  const canSend = i.status === 'todo' || i.status === 'due' || i.status === 'waiting', step = i.step, id = docId(c.email);
  const draft = canSend ? `
    <label class="lbl" for="s${id}">${esc(STEP_NAME[step])} · subject</label>
    <input class="subj" id="s${id}" value="${esc(subjectFor(c, step))}">
    <label class="lbl" for="b${id}">${esc(STEP_NAME[step])} · body (editable; your Outlook signature is added on send)</label>
    <textarea class="draft ${step === 1 ? 'first' : ''}" id="b${id}">${esc(bodyFor(c, step))}</textarea>
    <div class="actions">
      <a class="btn primary" data-act="owa" href="#" target="_blank" rel="noopener">Open ${esc(STEP_NAME[step].toLowerCase())} in Outlook web</a>
      <a class="btn" data-act="mailto" href="#" target="_top">Open in Outlook app</a>
      <button class="btn" data-act="copyEmail">Copy address</button>
      <button class="btn" data-act="copyBody">Copy body</button>
      ${c.pending && c.pending.step === step ? '<button class="btn" data-act="unpend">Clear “draft opened”</button>' : ''}
      <button class="btn" data-act="manualSent">Mark ${esc(STEP_NAME[step].toLowerCase())} as sent manually</button>
      ${i.n === 0 ? '' : '<button class="btn" data-act="manualReplied">They replied (mark by hand)</button>'}
      <button class="btn danger" data-act="remove">Remove prospect</button>
    </div>` : i.status === 'replied' ? `
    <div class="actions">
      <a class="btn primary" href="https://outlook.office.com/mail/inbox${SENDER ? '?login_hint=' + enc(SENDER) : ''}" target="_blank" rel="noopener">Open Inbox in Outlook web</a>
      <button class="btn" data-act="copyEmail">Copy address</button>
      <button class="btn" data-act="notReply">Not a real reply (auto-reply)</button>
      <button class="btn" data-act="handled">I answered them, remove from here</button>
    </div>` : i.status === 'finished' ? `
    <div class="actions"><button class="btn" data-act="copyEmail">Copy address</button><button class="btn" data-act="manualReplied">They replied (mark by hand)</button></div>` : `
    <div class="actions"><button class="btn" data-act="unhandle">Bring back to Replied</button></div>`;
  const note = needsReview(c) && step === 1 && canSend ? `No area was given for this prospect and the job title does not say whether they look after the web, communication or IT, so the email says “${esc(areaFor(c))}”. Edit it before sending, or remove this prospect.`
    : i.bounced && canSend ? 'Outlook reported this email as undeliverable, so it is not counted. If the bounce said “bad outbound sender”, your mailbox was blocked by Microsoft; wait for IT to lift the block before resending. If it said the address was not found, fix or remove the prospect.'
    : canSend ? 'Nothing is sent from this page; Outlook opens with the draft ready. The prospect advances only when the email shows up in your Sent Items.' : '';
  return `<div class="body">${timelineHTML(c, i)}
    <div class="fields">
      <div><b>Email</b><span>${esc(c.email)}</span></div>
      ${c.website ? `<div><b>Website</b><span><a href="${esc(c.website)}" target="_blank" rel="noopener">${esc(String(c.website).replace(/^https?:\/\/(www\.)?/, ''))}</a></span></div>` : ''}
      ${c.linkedin ? `<div><b>LinkedIn</b><span><a href="${esc(c.linkedin)}" target="_blank" rel="noopener">profile</a></span></div>` : ''}
      ${c.city ? `<div><b>City</b><span>${esc(c.city)}</span></div>` : ''}
    </div>${draft}${note ? `<p class="note">${note}</p>` : ''}</div>`;
}

function render(){
  const rv = $('#fReview').checked;
  const all = Object.values(P).map(c => ({c, i: info(c)}));
  const counts = {todo: 0, due: 0, waiting: 0, replied: 0, finished: 0, handled: 0};
  for (const x of all) counts[x.i.status]++;
  const rows = all.filter(x => (showHandled && tab === 'replied' ? (x.i.status === 'replied' || x.i.status === 'handled') : x.i.status === tab) && (!rv || needsReview(x.c)));
  rows.sort(TAB_ORDER[tab]);
  let html = '', lastOrg = null;
  for (const x of rows) {
    if (x.c.organisation !== lastOrg) { lastOrg = x.c.organisation; html += `<div class="grp">${esc(lastOrg)} · ${rows.filter(y => y.c.organisation === lastOrg).length}</div>`; }
    html += cardHTML(x.c);
  }
  list.innerHTML = html;
  for (const t of document.querySelectorAll('#tabs .tab')) { const k = t.dataset.tab; t.textContent = TAB_LABEL[k] + ` (${counts[k]})`; t.setAttribute('aria-selected', k === tab); }
  $('#tabHint').textContent = HINT[tab];
  $('#empty').hidden = rows.length > 0 || !all.length; $('#empty').textContent = EMPTY[tab];
  const total = all.length, contacted = total - counts.todo, replied = counts.replied + counts.handled, bounced = all.filter(x => x.i.bounced).length;
  $('#stats').innerHTML = total ? `
    <div class="stat"><b>${total}</b><span>prospects</span></div>
    <div class="stat"><b>${contacted}</b><span>contacted</span></div>
    <div class="stat"><b>${counts.due}</b><span>follow-ups due</span></div>
    <div class="stat"><b>${replied}</b><span>replied${contacted ? ' · ' + Math.round(100 * replied / contacted) + '%' : ''}</span></div>
    <div class="stat"><b>${counts.finished}</b><span>finished, no reply</span></div>
    ${bounced ? `<div class="stat"><b>${bounced}</b><span>bounced</span></div>` : ''}` : '';
  const foot = $('#foot'); foot.hidden = !(tab === 'replied' && counts.handled > 0);
  if (!foot.hidden) foot.innerHTML = `${counts.handled} conversation${counts.handled === 1 ? '' : 's'} answered and handed over to Outlook. <button id="toggleHandled">${showHandled ? 'Hide them' : 'Show them'}</button>`;
  if (!total && loaded && camp) { $('#empty').hidden = false; $('#empty').textContent = 'No prospects in this campaign yet. Use “Import / backup” to add them.'; }
}
function refreshLinks(card, c){
  const subj = card.querySelector('.subj'), body = card.querySelector('.draft'); if (!subj || !body) return;
  const l = links(c, subj.value, body.value);
  card.querySelector('[data-act="owa"]').href = l.owa; card.querySelector('[data-act="mailto"]').href = l.mailto;
}

// ---------- persistence (through the store adapter) ----------
function stripState(c){ const {key, ...rest} = c; return rest; }
function itemsById(){ const out = {}; for (const c of Object.values(P)) out[docId(c.email)] = stripState(c); return out; }
function applyItems(items){ P = {}; for (const c of Object.values(items)) { c.key = keyOf(c.email); if (c.key) P[c.key] = c; } }
const changedIds = new Set(), removedIds = new Set(); let flushTimer = null, flushing = null;
function saveProspect(c){ if (!campaignId) { say('No campaign loaded; reload the page'); return Promise.resolve(); } changedIds.add(docId(c.email)); return scheduleFlush(); }
function removeProspect(c){ const id = docId(c.email); removedIds.add(id); changedIds.delete(id); delete P[c.key]; open.delete(c.key); return scheduleFlush(); }
function scheduleFlush(){ return new Promise(res => { clearTimeout(flushTimer); flushTimer = setTimeout(() => flush().then(res, res), 150); }); }   // coalesce rapid edits into one save
async function flush(){
  if (flushing) await flushing;
  if (!changedIds.size && !removedIds.size) return;
  const ids = [...changedIds], rem = [...removedIds], cid = campaignId; changedIds.clear(); removedIds.clear();
  flushing = (async () => {
    try { const merged = await store.saveProspects(cid, itemsById(), ids, rem); if (merged && cid === campaignId) { applyItems(merged); render(); } }
    catch (e) { ids.forEach(i => changedIds.add(i)); rem.forEach(i => removedIds.add(i)); say('Could not save: ' + (e.message || e.code)); logLine('save failed: ' + (e.message || e.code)); }
  })();
  await flushing; flushing = null;
}
function saveCampaign(){ return store.saveCampaign(campaignId, camp).catch(e => say('Could not save settings: ' + (e.message || e.code))); }
function saveProfile(patch){ Object.assign(profile, patch); return store.saveProfile(profile).catch(e => logLine('profile save failed: ' + (e.message || e.code))); }
function newId(){ let id; do { id = 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); } while (campaigns[id]); return id; }
function freshCampaign(name, start){ return Object.assign({}, DEFAULT_CAMPAIGN, {type: 'campaign', name, createdAt: new Date().toISOString(), campaignStart: start, t: clone(DEFAULT_CAMPAIGN.t), holidays: [...DEFAULT_HOLIDAYS]}); }
async function loadCampaigns(){
  const idx = await store.loadIndex(); campaigns = idx.campaigns || {}; profile = idx.profile || {};
  if (!Object.keys(campaigns).length) { const id = newId(); campaigns[id] = freshCampaign(DEFAULT_CAMPAIGN.name, DEFAULT_CAMPAIGN.campaignStart); await store.saveCampaign(id, campaigns[id]); }
  for (const [id, c] of Object.entries(campaigns)) if (c.t && c.t[1] && c.t[1].body === OLD_T1_BODY) { c.t[1].body = T1_BODY; store.saveCampaign(id, c).catch(() => {}); logLine(`“${c.name}”: the Email 1 template was a placeholder and has been replaced with the full text (edit it under Sequence & settings)`); }
  let last = profile.campaignId || null;                       // the campaign you were in last time, remembered on the server; the browser copy is a fallback
  if (!campaigns[last]) { try { last = localStorage.getItem('desk-campaign'); } catch (e) {} }
  campaignId = campaigns[last] ? last : Object.keys(campaigns).sort((a, b) => (campaigns[a].createdAt || '').localeCompare(campaigns[b].createdAt || ''))[0];
  camp = campaigns[campaignId];
  const sel = $('#campaignSel'); sel.innerHTML = Object.keys(campaigns).map(id => `<option value="${id}" ${id === campaignId ? 'selected' : ''}>${esc(campaigns[id].name)}</option>`).join('');
}
async function loadProspects(){
  P = {}; loaded = false; changedIds.clear(); removedIds.clear(); const cid = campaignId;
  $('#empty').hidden = false; $('#empty').textContent = 'Loading your prospects…';
  const {items, recovered} = await store.loadProspects(cid);
  if (cid !== campaignId) return;
  applyItems(items); loaded = true;
  if (recovered && (recovered.added || recovered.merged)) say(`${recovered.added + recovered.merged} prospects recovered from an older version of this page`);
  open.clear(); render();
}
async function switchCampaign(id){ campaignId = id; camp = campaigns[id]; try { localStorage.setItem('desk-campaign', id); } catch (e) {} saveProfile({campaignId: id}); await loadProspects(); }

// ---------- import, backup ----------
function parseCSV(text){
  text = text.replace(/^﻿/, '');
  const first = text.split(/\r?\n/)[0] || '', delim = [',', ';', '\t'].map(d => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const d = delim === '\t' ? '\\t' : delim, re = new RegExp('("([^"]|"")*"|[^' + d + '\\r\\n]*)(' + d + '|\\r?\\n|$)', 'g');
  const rows = []; let row = [], m;
  while ((m = re.exec(text)) !== null) { let v = m[1]; if (v.startsWith('"')) v = v.slice(1, -1).replace(/""/g, '"'); row.push(v); if (m[3] !== delim) { if (row.length > 1 || row[0] !== '') rows.push(row); row = []; } if (m[3] === '') break; }
  if (!rows.length) return [];
  const head = rows[0].map(h => h.trim().toLowerCase().replace(/\s+/g, '_'));
  const alias = {firstname: 'first_name', first_name: 'first_name', nombre: 'first_name', lastname: 'last_name', last_name: 'last_name', apellidos: 'last_name', email: 'email', correo: 'email', company: 'organisation', company_name: 'organisation', company_name_for_emails: 'organisation', organization: 'organisation', organisation: 'organisation', title: 'title', cargo: 'title', area: 'area', website: 'website', person_linkedin_url: 'linkedin', linkedin: 'linkedin', city: 'city', ciudad: 'city', subject: 'subject', body: 'body', review: 'review'};
  return rows.slice(1).map(r => { const o = {}; head.forEach((h, k) => { const f = alias[h]; if (f && r[k] !== undefined) o[f] = r[k].trim(); }); return o; }).filter(o => o.email);
}
async function importProspects(file){
  if (!campaignId) throw new Error('no campaign loaded; reload the page first');
  if (!loaded) throw new Error('your existing prospects have not finished loading; wait a moment or reload, then import again');   // importing over an unread list would overwrite it
  const text = await file.text(); let arr;
  const isJson = file.name.toLowerCase().endsWith('.json') || /^\s*[\[{]/.test(text);
  try { arr = isJson ? JSON.parse(text) : parseCSV(text); } catch (e) { throw new Error('could not read the file'); }
  if (arr && !Array.isArray(arr) && arr.type === BACKUP_TYPE) return restoreBackup(arr);
  if (!Array.isArray(arr) || !arr.length) throw new Error('no prospects found in the file');
  let added = 0, updated = 0, unchanged = 0; const seen = new Set();
  for (const raw of arr) {
    const email = keyOf(raw.email); if (!email || !email.includes('@') || seen.has(email)) continue;
    seen.add(email);
    const cur = P[email], fields = {};
    for (const f of CONTACT_FIELDS) fields[f] = f === 'email' ? String(raw.email).trim() : (raw[f] || '');
    if (cur && CONTACT_FIELDS.every(k => (cur[k] || '') === fields[k])) { unchanged++; continue; }   // re-import: skip rows that did not change
    const c = Object.assign({}, cur || {log: {sent: [], replies: [], bounces: []}, manual: {}, pending: null, addedAt: new Date().toISOString()}, fields);
    c.key = email; P[email] = c; cur ? updated++ : added++;
    changedIds.add(docId(c.email));
  }
  $('#importMsg').textContent = `Saving ${added + updated} prospects…`;
  await flush();
  return {added, updated, unchanged, failed: changedIds.size};
}
async function restoreBackup(b){
  const wasEmpty = !Object.keys(P).length; let added = 0, updated = 0;
  for (const raw of b.prospects || []) {
    const o = clone(raw); delete o.chunk; delete o.key; const k = keyOf(o.email); if (!k || !k.includes('@')) continue;
    const cur = P[k];
    if (!cur) { o.key = k; o.log = o.log || {sent: [], replies: [], bounces: []}; o.manual = o.manual || {}; P[k] = o; changedIds.add(docId(o.email)); added++; continue; }
    let ch = false; cur.log = cur.log || {sent: [], replies: [], bounces: []}; const ol = o.log || {};
    for (const f of ['sent', 'replies', 'bounces', 'threads']) for (const t of ol[f] || []) if (addUnique(cur.log[f] = cur.log[f] || [], t)) ch = true;
    for (const [f, v] of Object.entries(o.manual || {})) if (v && (cur.manual = cur.manual || {})[f] == null) { cur.manual[f] = v; ch = true; }
    for (const f of CONTACT_FIELDS) if (!cur[f] && o[f]) { cur[f] = o[f]; ch = true; }
    if (ch) { changedIds.add(docId(cur.email)); updated++; }
  }
  let settings = false;
  if (wasEmpty && b.campaign) { for (const f of ['t', 'waitDays', 'holidays', 'campaignStart', 'name']) if (b.campaign[f] !== undefined) camp[f] = clone(b.campaign[f]); await saveCampaign(); $('#campaignSel').querySelector(`option[value="${campaignId}"]`).textContent = camp.name; settings = true; }
  $('#importMsg').textContent = `Restoring ${added + updated} prospects…`;
  await flush();
  return {added, updated, unchanged: (b.prospects || []).length - added - updated, failed: changedIds.size, settings};
}
function backupJson(){ return JSON.stringify({type: BACKUP_TYPE, version: 1, exportedAt: new Date().toISOString(), campaign: Object.assign({id: campaignId}, camp), prospects: Object.values(P).map(stripState)}, null, 1); }
function downloadBackup(){
  const blob = new Blob([backupJson()], {type: 'application/json'}), a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = `outreach-backup-${String(camp.name || 'campaign').replace(/[^\w.-]+/g, '_')}-${todayYmd()}.json`;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  say('Backup downloaded (if nothing happened, use “Copy backup”)');
}

// ---------- settings panel ----------
function fillTemplateFields(t){ for (const s of [1, 2, 3]) { $(`#t${s}s`).value = (t[s] || {}).subject || ''; $(`#t${s}b`).value = (t[s] || {}).body || ''; } }
function openSettings(){
  $('#settingsPanel').hidden = false; $('#importPanel').hidden = true;
  $('#setCampName').textContent = camp.name; $('#campName').value = camp.name;
  fillTemplateFields(camp.t || {});
  $('#senderEmail').value = SENDER; $('#w1').value = camp.waitDays[0]; $('#w2').value = camp.waitDays[1]; $('#hol').value = (camp.holidays || []).join('\n');
  $('#confirmDel').hidden = true;
}
async function saveSettings(){
  camp.name = $('#campName').value.trim() || camp.name;
  const se = keyOf($('#senderEmail').value); if (se !== SENDER) { SENDER = se; saveProfile({senderEmail: SENDER}); }
  camp.t = camp.t || {}; for (const s of [1, 2, 3]) camp.t[s] = {subject: $(`#t${s}s`).value, body: $(`#t${s}b`).value};
  camp.waitDays = [Math.max(1, +$('#w1').value || 2), Math.max(1, +$('#w2').value || 3)];
  camp.holidays = $('#hol').value.split(/\s+/).filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x));
  await saveCampaign(); $('#campaignSel').querySelector(`option[value="${campaignId}"]`).textContent = camp.name; $('#setCampName').textContent = camp.name; render(); say('Saved');
}
async function newCampaign(){
  const id = newId();
  campaigns[id] = freshCampaign('New campaign', todayYmd());
  await store.saveCampaign(id, campaigns[id]);
  $('#campaignSel').insertAdjacentHTML('beforeend', `<option value="${id}">New campaign</option>`); $('#campaignSel').value = id;
  await switchCampaign(id); openSettings(); say('Campaign created. Give it a name and templates, then import prospects.');
}
async function deleteCampaign(){
  const id = campaignId;
  try { await store.deleteCampaign(id); } catch (e) { say('Could not delete: ' + (e.message || e.code)); return; }
  delete campaigns[id]; $('#settingsPanel').hidden = true; if (profile.campaignId === id) saveProfile({campaignId: null});
  await loadCampaigns(); await loadProspects(); say('Campaign deleted');
}

// ---------- events ----------
list.addEventListener('click', async e => {
  const card = e.target.closest('.card'); if (!card) return;
  const c = P[card.dataset.key]; if (!c) return; const i = info(c);
  const act = e.target.closest('[data-act]');
  if (act) {
    const a = act.dataset.act;
    if (a === 'owa' || a === 'mailto') { refreshLinks(card, c); setTimeout(() => { c.pending = {step: i.step, at: new Date().toISOString()}; saveProspect(c); render(); say('Draft opened. It counts once Outlook shows it as sent.'); }, 400); return; }
    e.preventDefault();
    const body = card.querySelector('.draft'); c.manual = c.manual || {};
    if (a === 'copyEmail') copy(c.email, 'Address');
    if (a === 'copyBody' && body) copy(body.value, 'Body');
    if (a === 'unpend') { c.pending = null; }
    if (a === 'manualSent') { c.manual.sent = [...(c.manual.sent || []), new Date().toISOString()]; c.pending = null; open.delete(c.key); say(STEP_NAME[i.step] + ' marked as sent'); }
    if (a === 'manualReplied') { c.manual.replied = new Date().toISOString(); c.manual.ignoreReplies = false; c.manual.ignoreRepliesBefore = null; open.delete(c.key); say('Moved to Replied'); }
    if (a === 'notReply') { c.manual.replied = null; c.manual.ignoreReplies = false; c.manual.ignoreRepliesBefore = new Date().toISOString(); open.delete(c.key); say('Reply ignored; a later real reply will still count'); }
    if (a === 'handled') { c.manual.handled = true; open.delete(c.key); say('Removed from the page'); }
    if (a === 'unhandle') { c.manual.handled = false; open.delete(c.key); }
    if (a === 'remove') { removeProspect(c); render(); say('Prospect removed'); return; }
    if (a !== 'copyEmail' && a !== 'copyBody') saveProspect(c);
    render(); return;
  }
  if (e.target.closest('.head')) { open.has(c.key) ? open.delete(c.key) : open.add(c.key); render(); if (open.has(c.key)) refreshLinks(list.querySelector(`.card[data-key="${CSS.escape(c.key)}"]`), c); }
});
list.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && e.target.classList.contains('head')) { e.preventDefault(); e.target.click(); } });
list.addEventListener('input', e => { const card = e.target.closest('.card'); if (card && e.target.matches('.subj,.draft')) refreshLinks(card, P[card.dataset.key]); });
$('#fReview').addEventListener('input', render);
$('#tabs').addEventListener('click', e => { const t = e.target.closest('[data-tab]'); if (t) { tab = t.dataset.tab; showHandled = false; render(); } });
$('#foot').addEventListener('click', e => { if (e.target.id === 'toggleHandled') { showHandled = !showHandled; render(); } });
$('#campaignSel').addEventListener('change', e => switchCampaign(e.target.value));
$('#btnImport').addEventListener('click', () => { $('#importPanel').hidden = !$('#importPanel').hidden; $('#settingsPanel').hidden = true; $('#importMsg').textContent = ''; });
$('#btnImportClose').addEventListener('click', () => { $('#importPanel').hidden = true; });
$('#btnExport').addEventListener('click', downloadBackup);
$('#btnExportCopy').addEventListener('click', () => copy(backupJson(), 'Backup'));
$('#importFile').addEventListener('change', async e => {
  const f = e.target.files && e.target.files[0]; if (!f) return;
  try { const r = await importProspects(f); $('#importMsg').textContent = `Done: ${r.added} added, ${r.updated} updated${r.unchanged ? ', ' + r.unchanged + ' already up to date' : ''}${r.failed ? ', ' + r.failed + ' not saved (try again)' : ''}${r.settings ? '; campaign settings restored' : ''}.`; render(); say(`${r.added + r.updated} prospects imported`); }
  catch (err) { $('#importMsg').textContent = 'Import failed: ' + err.message; }
  e.target.value = '';
});
$('#btnSettings').addEventListener('click', () => { if ($('#settingsPanel').hidden) openSettings(); else $('#settingsPanel').hidden = true; });
$('#btnSettingsClose').addEventListener('click', () => { $('#settingsPanel').hidden = true; });
$('#btnSaveSettings').addEventListener('click', saveSettings);
$('#btnResetTemplates').addEventListener('click', () => { fillTemplateFields(DEFAULT_CAMPAIGN.t); say('Default templates filled in; press Save to keep them'); });
$('#btnNewCampaign').addEventListener('click', newCampaign);
$('#btnDeleteCampaign').addEventListener('click', () => { $('#confirmDel').hidden = false; });
$('#btnDeleteYes').addEventListener('click', deleteCampaign);
$('#btnSignIn').addEventListener('click', () => { if (store && store.signIn) store.signIn(); });
$('#btnSignOut').addEventListener('click', () => { if (store && store.signOut) store.signOut(); });
$('#copyList').addEventListener('click', () => {
  const rows = Object.values(P).map(c => { const i = info(c); return [c.first_name + ' ' + c.last_name, c.organisation, c.title, c.email, i.status, i.sent[0] || '', i.sent[1] || '', i.sent[2] || '', i.lastReply || '', i.due || ''].map(v => String(v).slice(0, 10)).join('\t'); });
  copy(['Name\tOrganisation\tTitle\tEmail\tStatus\tEmail 1\tFollow-up 1\tFollow-up 2\tReplied\tNext due', ...rows].join('\n'), 'Status list');
});

// ---------- Outlook sync ----------
function normSubject(s){ return String(s || '').toLowerCase().replace(/^(\s*(re|rv|fw|fwd|aw|sv|tr|vs|antw)\s*:\s*)+/i, '').replace(/\s+/g, ' ').trim(); }
function isNdr(it, snd){ return /^undeliverable|delivery status notification|mail delivery failed|returned mail|no se ha podido entregar|no se pudo entregar/i.test(it.subject) || /^(postmaster|mailer-daemon|microsoftexchange)/i.test(snd); }
async function syncOutlook(auto){
  if (!store.mail.available) { setStatus('Outlook sync is not available in this view.', true); return; }
  if (!camp) { setStatus('No campaign loaded; reload the page.', true); return; }
  if (!loaded) { setStatus('Prospects are still loading; try again in a moment.', true); return; }
  if (syncing) { say('Sync already running'); return; }
  syncing = true; $('#syncBtn').textContent = 'Syncing…';
  // Incremental only once something has actually been matched; until then rescan from the campaign start.
  const anyLogged = Object.values(P).some(c => c.log && ((c.log.sent || []).length || (c.log.replies || []).length));
  const since = (camp.lastSync && anyLogged) ? ymd(new Date(new Date(camp.lastSync).getTime() - 2 * 86400000)) : (camp.campaignStart || todayYmd());
  try {
    const changedKeys = new Set(), unmatched = []; let skippedSender = 0, nBounce = 0, nThread = 0, nSubject = 0; prog.phase = 0; showProgress(1, 'Connecting to Outlook…');
    const st = await store.mail.access(); logLine('Outlook access: ' + st);
    if (st === 'denied') { showProgress(null); setStatus('Outlook access was declined for this page. Allow Microsoft 365 in the page’s Permissions menu, then sync again.', true); return; }
    if (st === 'prompt') { showProgress(2, 'Asking for Outlook access… if no prompt appears, allow Microsoft 365 in the page’s Permissions menu'); await store.mail.request(); }
    const touch = c => { c.log = c.log || {sent: [], replies: [], bounces: []}; return c.log; };
    const byThread = {}; for (const c of Object.values(P)) for (const t of (c.log && c.log.threads) || []) byThread[t] = c;
    const noteThread = (c, t) => { if (!t) return; const L = touch(c); if (addUnique(L.threads = L.threads || [], t)) changedKeys.add(c.key); byThread[t] = c; };
    showProgress(3, 'Reading Sent Items…');
    logLine(`Sync start · ${Object.keys(P).length} prospects · since ${since} · sender filter ${SENDER || '(none)'}`);
    const nSent = await store.mail.scan('Sent Items', since, it => {
      const snd = keyOf(it.sender); if (SENDER && snd !== SENDER) { skippedSender++; return false; }
      const at = it.sentDateTime || it.receivedDateTime, targets = new Set();
      for (const r of it.recipients || []) { const c = P[keyOf(r)]; if (c) targets.add(c); }
      if (!targets.size && it.thread && byThread[it.thread]) targets.add(byThread[it.thread]);   // my answer inside a conversation the prospect continued from another address
      for (const c of targets) { if (addUnique(touch(c).sent, at)) changedKeys.add(c.key); if (c.pending) { c.pending = null; changedKeys.add(c.key); } noteThread(c, it.thread); }
      return targets.size > 0;
    }, folderProgress);
    if (skippedSender) logLine(`${skippedSender} sent emails ignored because the sender was not ${SENDER} (change “Your Outlook address” in settings if that is wrong)`);
    // Replies: by sender address first; else by conversation id (when the mail source gives one); else by the Email 1 subject,
    // which names the organisation — a colleague answering from another address still counts.
    const bySubject = {}; for (const c of Object.values(P)) { if (!((c.log && c.log.sent) || []).length) continue; const s = normSubject(subjectFor(c, 1)); if (s) (bySubject[s] = bySubject[s] || []).push(c); }
    const myDomain = (SENDER.split('@')[1] || '').toLowerCase();
    const nIn = await store.mail.scan('Inbox', since, it => {
      const snd = keyOf(it.sender), at = it.receivedDateTime || it.sentDateTime;
      if (isNdr(it, snd)) {
        const text = (it.preview || '') + ' ' + it.subject; let hit = false;
        for (const m of text.toLowerCase().match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/g) || []) { const c = P[m]; if (!c) continue; hit = true; const L = touch(c); if (addUnique(L.bounces = L.bounces || [], at)) { nBounce++; changedKeys.add(c.key); } }
        return hit;
      }
      if (!snd || snd === SENDER) return false;
      let c = P[snd];
      if (!c && it.thread && byThread[it.thread]) { c = byThread[it.thread]; nThread++; }
      if (!c) {
        const cands = bySubject[normSubject(it.subject)] || [];
        if (cands.length === 1) { c = cands[0]; nSubject++; }
        else if (cands.length > 1) { unmatched.push(`${snd} · “${it.subject}” · could be ${cands.map(x => x.first_name + ' ' + x.last_name).join(' or ')}`); return false; }   // two people at that organisation were emailed: a guess could stop follow-ups to the wrong one
      }
      if (!c) return false;
      if (myDomain && snd.endsWith('@' + myDomain)) return false;   // a colleague forwarding internally is not the prospect replying
      if (addUnique(touch(c).replies, at)) changedKeys.add(c.key); noteThread(c, it.thread); return true;
    }, folderProgress);
    for (const k of changedKeys) if (P[k]) saveProspect(P[k]);
    camp.lastSync = new Date().toISOString(); store.patchCampaign(campaignId, {lastSync: camp.lastSync}).catch(() => saveCampaign());   // a patch, not a full write: a stale copy in another tab must not roll settings back
    const emailed = Object.values(P).filter(c => (c.log && c.log.sent || []).length).length, replied = Object.values(P).filter(c => (c.log && c.log.replies || []).length).length;
    setStatus(`Synced ${new Date().toLocaleTimeString('en-GB', {hour: '2-digit', minute: '2-digit'})} · ${nSent} sent + ${nIn} received emails checked since ${fmt(since)} · ${emailed} prospects emailed · ${replied} replied${nBounce ? ' · ' + nBounce + ' new bounce' + (nBounce > 1 ? 's' : '') : ''} · ${changedKeys.size} updated${unmatched.length ? ' · ' + unmatched.length + (unmatched.length > 1 ? ' replies' : ' reply') + ' to match by hand (see details)' : ''}`);
    render();
    if (nThread || nSubject) logLine(`${nThread + nSubject} replies matched by conversation or subject rather than by sender address`);
    for (const u of unmatched) logLine('reply to match by hand (open the prospect and use “They replied”): ' + u);
    logLine(`Sync done · ${changedKeys.size} prospects updated`); showProgress(100); setTimeout(() => showProgress(null), 1500);
  } catch (e) { showProgress(null); setStatus(store.mail.errorText(e), true); logLine('Sync failed: ' + (e && (e.code || e.message))); }
  finally { syncing = false; $('#syncBtn').textContent = 'Sync with Outlook'; }
}
$('#syncBtn').addEventListener('click', () => syncOutlook(false));

// ---------- boot ----------
function gate(msg){ $('#gate').hidden = !msg; if (msg) $('#gateMsg').textContent = msg; }
(async () => {
  render();
  store = (window.claude && window.claude.use) ? createClaudeStore() : createGraphStore(window.OUTREACH_CONFIG || {});
  store.setLogger(logLine);
  $('#storageNote').textContent = store.describeStorage();
  let r; try { r = await store.init(); } catch (e) { gate('Could not start: ' + (e.message || e)); $('#who').textContent = ''; return; }
  if (!r.ok) {
    $('#who').textContent = store.kind === 'graph' ? 'Not signed in' : 'Sign in to Claude to use the desk.';
    if (r.reason === 'not_configured') gate('This copy is not connected to Microsoft yet: add the Azure app IDs to config.js (see docs/SETUP.md in the repository).');
    else if (r.reason === 'signed_out') { gate('Sign in with your Cludo Microsoft account to load your campaigns. The app reads your Outlook mail (read-only) and keeps its data in a private folder of your OneDrive.'); $('#btnSignIn').hidden = false; }
    else if (r.reason === 'signin_failed') { gate('Sign-in failed: ' + (r.detail || 'unknown error')); $('#btnSignIn').hidden = false; }
    else gate(r.message || 'This page cannot run here.');
    return;
  }
  viewer = r.viewer; if (store.signOut) $('#btnSignOut').hidden = false;
  if (r.readOnly) gate('You can view this app but not save anything: ask the owner to add you as a Contributor (or higher) in the share menu, then reload.');
  try { await loadCampaigns(); }
  catch (e) { gate(store.loadFailureText(e)); logLine('load failed: ' + (e.stack || e.message || e.code)); return; }
  SENDER = String(profile.senderEmail || viewer.email || '').toLowerCase();
  $('#who').textContent = `${viewer.name || 'You'}${SENDER ? ' · ' + SENDER : ' · set your Outlook address under Sequence & settings'} · ${store.kind === 'graph' ? 'data in your OneDrive' : 'your prospects are private to you'}`;
  try { await loadProspects(); }
  catch (e) { gate(store.loadFailureText(e)); logLine('load failed: ' + (e.stack || e.message || e.code)); return; }
  if (!store.mail.available) { setStatus('Outlook access is not available in this view; mark steps by hand.', true); $('#syncBtn').hidden = true; return; }
  const st = await store.mail.access();
  if (st === 'granted') { if (Object.keys(P).length) syncOutlook(true); else setStatus('Import prospects, then sync.'); }
  else if (st === 'denied') setStatus('Outlook access was declined for this page. Allow Microsoft 365 in the page’s Permissions menu, then press Sync.', true);
  else { $('#syncBtn').textContent = 'Allow Outlook access and sync'; setStatus('First time here: press the button and accept the Microsoft 365 prompt.'); }
})();
