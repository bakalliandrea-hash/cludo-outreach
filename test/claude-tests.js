// Artifact (Claude adapter) tests against the mocked runtime in harness-claude.js: frozen documents, an in-memory store
// seeded with a synthetic campaign (two chunks plus two records in the old per-prospect layout), and a fake Outlook connector.
const { chromium } = require('playwright'); const { routeApp } = require('./serve'); const fs = require('fs'), path = require('path'), os = require('os');
const HARNESS = fs.readFileSync(path.join(__dirname, 'harness-claude.js'), 'utf8');
const R = []; const ok = (n, c, x = '') => R.push(`${c ? 'PASS' : 'FAIL'}  ${n}${x ? '  — ' + x : ''}`);
const ME = 'andrea.bakalli@cludo.com', OLD_T1 = 'Hola {first_name},\n\nHe visto que te ocupas de {area} en {organisation} y te escribo en relación con el buscador de vuestra web.\n\n…';
const TITLES = ['Responsable web', 'Directora de comunicación', 'Profesor titular', 'Jefe de sistemas'];
function prospect(k){ return {first_name: 'Nombre' + k, last_name: 'Apellido' + k, email: `persona${k}@uni${k % 7}.example`, organisation: 'Universidad ' + (k % 7), title: TITLES[k % 4], area: '', website: '', linkedin: '', city: '', subject: '', body: '', review: '', addedAt: '2026-10-01T08:00:00.000Z', log: {sent: [], replies: [], bounces: []}, manual: {}, pending: null}; }
const did = e => e.toLowerCase().replace(/[^a-z0-9]/g, '_');
function buildStore(){
  const ps = Array.from({length: 45}, (_, k) => prospect(k));
  for (let k = 0; k < 10; k++) ps[k].log.sent = [`2026-10-05T09:0${k}:00.000Z`];                 // waiting
  ps[10].log.sent = ['2026-10-05T09:30:00.000Z']; ps[10].log.bounces = ['2026-10-05T09:30:20.000Z'];   // bounced → back to To contact
  ps[11].log.sent = ['2026-10-05T09:40:00.000Z', '2026-10-05T09:40:16.000Z'];                      // resend 16 s later
  const S = {}, C = 'camp1', B = 'data/users/u_test/';
  S[B + C] = {type: 'campaign', name: 'Synthetic campaign', createdAt: '2026-10-05T08:00:00.000Z', campaignStart: '2026-10-05', waitDays: [2, 3], lastSync: '2026-10-05T09:00:00.000Z', holidays: ['2026-10-12', '2026-11-01'],
    t: {1: {subject: 'Experiencia de búsqueda en la web de {organisation}', body: OLD_T1}, 2: {subject: 'Re: X', body: 'B2 {first_name}'}, 3: {subject: 'Re: X', body: 'B3 {first_name}'}}};
  for (const n of [0, 1]) { const items = {}; ps.slice(n * 40, n * 40 + 40).forEach(p => { items[did(p.email)] = Object.assign({chunk: n}, p); }); S[`${B}${C}/chunks/p${n}`] = {type: 'chunk', n, items}; }
  const a = JSON.parse(JSON.stringify(ps[0])); a.log.sent = ['2026-10-03T08:00:00.000Z'];          // old layout: an extra sent date for prospect 0
  const b = prospect(99); b.email = 'only.legacy@uni9.example';                                      // old layout: a prospect missing from the chunks
  S[`${B}${C}/prospects/${did(a.email)}`] = a; S[`${B}${C}/prospects/${did(b.email)}`] = b;
  return {S, ps, C};
}
async function boot(browser, {store, mail = {'Sent Items': [], 'Inbox': []}, viewport = {width: 1200, height: 900}, scheme = 'light'}){
  const ctx = await browser.newContext({viewport, colorScheme: scheme}); const page = await ctx.newPage(); const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message)); page.on('console', m => { if (m.type() === 'error' && !/ERR_FAILED/.test(m.text())) errs.push('console: ' + m.text()); });
  await page.clock.setFixedTime(new Date('2026-10-05T15:00:00'));
  await page.addInitScript(([s, m, h]) => { window.__store = s; window.__mail = m; eval(h); }, [JSON.parse(JSON.stringify(store)), mail, HARNESS]);
  await routeApp(page); await page.goto('http://app.test/');
  await page.waitForFunction(() => /prospects/.test(document.querySelector('#stats').textContent), null, {timeout: 10000}).catch(() => {});
  await page.waitForTimeout(800); return {ctx, page, errs};
}
const counts = page => page.$$eval('#tabs .tab', ts => Object.fromEntries(ts.map(t => [t.dataset.tab, +t.textContent.match(/\((\d+)\)/)[1]])));
(async () => {
  const browser = await chromium.launch(); const {S, ps, C} = buildStore();
  // ---- load, recovery, persistence ----
  let {ctx, page, errs} = await boot(browser, {store: S});
  let c = await counts(page);
  ok('boots on frozen documents, all prospects load (45 + 1 recovered from the old layout)', Object.values(c).reduce((a, b) => a + b, 0) === 46, JSON.stringify(c));
  ok('queues: 11 waiting (10 + the resend counted once), bounced one back in To contact', c.waiting === 11 && c.todo === 35, JSON.stringify(c));
  ok('old-layout records merged and deleted', (await page.evaluate(() => Object.keys(__store).filter(k => k.includes('/prospects/')).length)) === 0);
  ok('sent date known only to the old layout merged in', await page.evaluate(() => P['persona0@uni0.example'].log.sent.includes('2026-10-03T08:00:00.000Z')));
  ok('Email 1 placeholder template replaced by the full text and saved', await page.evaluate(() => camp.t[1].body.length > 1000 && __store['data/users/u_test/camp1'].t[1].body.length > 1000));
  const L = await page.evaluate(() => ({
    wd: addWorkDays('2026-10-09T10:00:00Z', 2), wd2: addWorkDays('2026-10-07T10:00:00Z', 3),
    ign: info({email: 'x@y.es', log: {sent: ['2026-10-01T09:00:00Z'], replies: ['2026-10-01T10:00:00Z', '2026-10-04T10:00:00Z'], bounces: []}, manual: {ignoreRepliesBefore: '2026-10-02T00:00:00Z'}}).status,
    csv: parseCSV('﻿first_name,last_name,email,organisation,title\n"Ana","De la Peña, M.",ana@uni.es,"Uni ""X""",Web\r\nB,C,b@uni.es,U,T\n'),
    csvSemi: parseCSV('first_name;last_name;email;organisation;title\nAna;P;ana@uni.es;U;T\n').length,
    csvTab: parseCSV('first_name\tlast_name\temail\torganisation\ttitle\nAna\tP\tana@uni.es\tU\tT\n').length,
    areas: Object.values(P).slice(0, 4).map(c => [c.title, areaFor(c), needsReview(c)]),
  }));
  ok('working days skip the weekend and the 12 Oct holiday', L.wd === '2026-10-14' && L.wd2 === '2026-10-13', L.wd + ' ' + L.wd2);
  ok('a real reply after "not a real reply" still counts', L.ign === 'replied', L.ign);
  ok('CSV: BOM, quoted commas, escaped quotes, CRLF; semicolons; tabs', L.csv.length === 2 && L.csv[0].last_name === 'De la Peña, M.' && L.csv[0].organisation === 'Uni "X"' && L.csvSemi === 1 && L.csvTab === 1, JSON.stringify([L.csv[0], L.csvSemi, L.csvTab]));
  ok('area from title: web / communication / IT; academic title flagged', JSON.stringify(L.areas.map(a => a.slice(1))) === JSON.stringify([['la web', false], ['la comunicación', false], ['la comunicación y la web', true], ['los sistemas informáticos', false]]), JSON.stringify(L.areas));
  await page.click('#tabs [data-tab="todo"]'); await page.click('.card .head');
  const draft = await page.$eval('.card .draft', e => e.value), owa = await page.$eval('.card [data-act="owa"]', e => e.href);
  ok('Email 1 draft is complete and the Outlook link carries to/subject/body/login_hint', draft.length > 1000 && !/…\s*$/.test(draft) && /deeplink\/compose\?to=.+&subject=.+&body=.+&login_hint=/.test(owa), draft.length + ' chars');
  const before = await counts(page); await page.click('.card [data-act="manualSent"]'); await page.waitForTimeout(600); const after = await counts(page);
  ok('"mark as sent manually" moves the prospect to Waiting and saves', after.todo === before.todo - 1 && after.waiting === before.waiting + 1 && await page.evaluate(() => Object.values(__store).some(d => d.items && Object.values(d.items).some(x => (x.manual.sent || []).length))));
  await page.click('.card .head'); await page.click('.card [data-act="remove"]'); await page.waitForTimeout(600);
  ok('remove prospect deletes it from storage', (await page.evaluate(() => Object.values(__store).filter(d => d.items).reduce((a, d) => a + Object.keys(d.items).length, 0))) === 45);
  const tmp = os.tmpdir(); fs.writeFileSync(path.join(tmp, 'reimport.json'), JSON.stringify(ps.slice(25, 45)));   // untouched rows fs.writeFileSync(path.join(tmp, 'new.csv'), 'first_name,last_name,email,organisation,title\nTest,Uno,test.uno@example.org,Example Uni,Webmaster\n');
  await page.click('#btnImport'); await page.setInputFiles('#importFile', path.join(tmp, 'reimport.json')); await page.waitForTimeout(800);
  const m1 = await page.$eval('#importMsg', e => e.textContent); await page.setInputFiles('#importFile', path.join(tmp, 'new.csv')); await page.waitForTimeout(800);
  const m2 = await page.$eval('#importMsg', e => e.textContent);
  ok('re-importing the same prospects changes nothing; a CSV adds one', /0 added, 0 updated, 20 already/.test(m1) && /1 added/.test(m2), m1 + ' | ' + m2);
  const total = await page.evaluate(() => Object.keys(P).length);   // 46 − 1 removed + 1 imported
  const backup = await page.evaluate(() => JSON.parse(backupJson()));
  ok('backup JSON carries campaign + prospects with history', backup.type === 'cludo-outreach-backup' && backup.prospects.length === total && total === 46 && backup.prospects.some(p => p.log.sent.length), total + ' prospects');
  await page.click('#btnSettings'); await page.fill('#w1', '4'); await page.fill('#campName', 'Renamed'); await page.click('#btnSaveSettings'); await page.waitForTimeout(400);
  ok('settings save (name, wait days)', await page.evaluate(() => __store['data/users/u_test/camp1'].waitDays[0] === 4 && __store['data/users/u_test/camp1'].name === 'Renamed'));
  await page.click('#btnNewCampaign'); await page.waitForTimeout(600);
  const store2 = await page.evaluate(() => __store); ok('no JS errors so far', errs.length === 0, errs.join(' | ')); await ctx.close();
  ({ctx, page, errs} = await boot(browser, {store: store2}));
  ok('reload: remembered campaign opens, no writes except the sync timestamp', (await page.evaluate(() => camp.name)) === 'New campaign' && (await page.evaluate(() => __writes.filter(w => !(w[0] === 'update')).length)) === 0);
  await page.selectOption('#campaignSel', 'camp1'); await page.waitForTimeout(800);
  ok('switching back loads the prospects again', (await page.evaluate(() => Object.keys(P).length)) === total);
  // restore a backup into an empty campaign
  fs.writeFileSync(path.join(tmp, 'backup.json'), JSON.stringify(backup));
  await page.click('#btnSettings'); await page.click('#btnNewCampaign'); await page.waitForTimeout(600); await page.click('#btnImport'); await page.setInputFiles('#importFile', path.join(tmp, 'backup.json')); await page.waitForTimeout(1000);
  const m3 = await page.$eval('#importMsg', e => e.textContent);
  ok('backup restores prospects with history and the campaign settings into an empty campaign', new RegExp(total + ' added').test(m3) && /settings restored/.test(m3) && await page.evaluate(() => Object.values(P).some(c => c.log.sent.length) && camp.name === 'Synthetic campaign' && camp.waitDays[0] === 2), m3);   // the backup was taken before the rename and the wait-days change
  await ctx.close();
  // ---- sync ----
  const todo = ps[20], waiting = ps[1], waiting2 = ps[5];   // ps[5] is the only contacted person at Universidad 5; Universidad 2 has two (ps[2], ps[9])
  const filler = Array.from({length: 40}, (_, k) => ({subject: 'Newsletter ' + k, sender: 'news' + k + '@example.net', recipients: [ME], receivedDateTime: '2026-10-05T12:' + String(k).padStart(2, '0') + ':00Z', sentDateTime: '2026-10-05T12:00:00Z', summary: ''}));
  const mail = {
    'Sent Items': [{subject: 'Experiencia de búsqueda en la web de ' + todo.organisation, sender: ME, recipients: [todo.email], sentDateTime: '2026-10-05T14:00:00Z', receivedDateTime: '2026-10-05T14:00:01Z', summary: 'Hola'}],
    'Inbox': [...filler,
      {subject: 'Re: Experiencia', sender: waiting.email.toUpperCase(), recipients: [ME], receivedDateTime: '2026-10-05T14:30:00Z', sentDateTime: '2026-10-05T14:29:00Z', summary: 'Gracias'},
      {subject: 'RE: Re: Experiencia de búsqueda en la web de ' + waiting2.organisation, sender: 'colleague@another-domain.example', recipients: [ME], receivedDateTime: '2026-10-05T14:40:00Z', sentDateTime: '2026-10-05T14:39:00Z', summary: 'Le paso a mi compañera'},
      {subject: 'RE: Experiencia de búsqueda en la web de ' + ps[2].organisation, sender: 'someone@elsewhere.example', recipients: [ME], receivedDateTime: '2026-10-05T14:45:00Z', sentDateTime: '2026-10-05T14:44:00Z', summary: 'ambiguous: two prospects at this organisation'},
      {subject: 'Undeliverable: Experiencia de búsqueda en la web de ' + ps[3].organisation, sender: 'MicrosoftExchange329e71ec88ae4615bbc36ab6ce41109e@cludo.com', recipients: [ps[3].email], receivedDateTime: '2026-10-05T14:50:00Z', sentDateTime: '2026-10-05T14:50:00Z', summary: 'Delivery has failed to these recipients or groups:\r\n\r\n' + ps[3].email}]
  };
  ({ctx, page, errs} = await boot(browser, {store: S, mail}));
  await page.waitForFunction(() => /Synced/.test(document.querySelector('#syncStatus').textContent), null, {timeout: 10000}).catch(() => {});
  const st = await page.evaluate(([a, b, d, e]) => [a, b, d, e].map(x => info(P[keyOf(x)]).status), [todo.email, waiting.email, waiting2.email, ps[3].email]);
  ok('sync: Sent Items match → Waiting; reply (case-insensitive) → Replied; reply from another address by subject → Replied; bounce voids the send → To contact', st[0] === 'waiting' && st[1] === 'replied' && st[2] === 'replied' && st[3] === 'todo', JSON.stringify(st));
  const amb = await page.evaluate(() => [info(P['persona2@uni2.example']).status, info(P['persona9@uni2.example']).status, document.querySelector('#syncLog').textContent.includes('match by hand'), document.querySelector('#syncStatus').textContent]);
  ok('sync: ambiguous subject match (two prospects at one organisation) is not guessed; listed for manual matching', amb[0] === 'waiting' && amb[1] === 'waiting' && amb[2] && /1 reply to match by hand/.test(amb[3]), JSON.stringify(amb));
  const icalls = await page.evaluate(() => __calls.filter(x => x.args.folderName === 'Inbox').map(x => x.args.offset));
  ok('sync: Inbox paging follows nextOffset', icalls.join(',') === '0,25', icalls.join(','));
  ok('sync: connector called with folderName + afterDateTime + order only', await page.evaluate(() => __calls.every(x => x.args.folderName && x.args.afterDateTime && x.args.order === 'newest' && !x.args.query && !x.args.recipient)));
  ok('sync: lastSync written with update(), not a full set()', await page.evaluate(() => __writes.some(w => w[0] === 'update' && /camp1$/.test(w[1]))));
  ok('sync: no JS errors', errs.length === 0, errs.join(' | ')); await ctx.close();
  // ---- layout ----
  for (const [label, vp, scheme] of [['desktop-light', {width: 1200, height: 900}, 'light'], ['mobile-dark', {width: 390, height: 844}, 'dark']]) {
    ({ctx, page, errs} = await boot(browser, {store: S, viewport: vp, scheme}));
    await page.click('#tabs [data-tab="todo"]'); await page.click('.card .head');
    const sw = await page.evaluate(() => [document.documentElement.scrollWidth, window.innerWidth]);
    ok(`${label}: no horizontal scroll`, sw[0] <= sw[1], sw.join(' vs ')); await ctx.close();
  }
  await browser.close(); console.log(R.join('\n')); if (R.some(l => l.startsWith('FAIL'))) process.exit(1);
})().catch(e => { console.error(e); process.exit(1); });
