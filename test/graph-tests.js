// Standalone (Graph adapter) tests against the mocked msal + Graph in graph-harness.js. Fixtures are synthetic.
const { chromium } = require('playwright'); const { routeApp } = require('./serve'); const fs = require('fs');
const GH = fs.readFileSync(require('path').join(__dirname, 'harness-graph.js'), 'utf8');
const CONFIG_ON = "window.OUTREACH_CONFIG = {clientId: 'test-client', tenantId: 'test-tenant', scopes: ['User.Read', 'Mail.Read', 'Files.ReadWrite.AppFolder']};";
const CONFIG_OFF = fs.readFileSync(require('path').join(__dirname, '..', 'config.js'), 'utf8');
const R = []; const ok = (n, c, x = '') => R.push(`${c ? 'PASS' : 'FAIL'}  ${n}${x ? '  — ' + x : ''}`);
const me = 'andrea.bakalli@cludo.com', A = a => ({emailAddress: {address: a}});
async function boot(browser, {signedIn = true, drive = {}, mail = {sentitems: [], inbox: []}, config = CONFIG_ON} = {}){
  const ctx = await browser.newContext(); const page = await ctx.newPage(); const errs = [];
  page.on('pageerror', e => errs.push(e.message)); page.on('console', m => { if (m.type() === 'error' && !/ERR_FAILED/.test(m.text())) errs.push(m.text()); });
  await page.clock.setFixedTime(new Date('2026-10-05T15:00:00'));
  await page.addInitScript(([d, m, s, h]) => { window.__drive = d; window.__mail = m; window.__signedIn = s; eval(h); }, [drive, mail, signedIn, GH]);
  await routeApp(page, {overrides: {'/config.js': config, '/vendor/msal-browser.min.js': ''}});
  await page.goto('http://app.test/'); await page.waitForTimeout(1500);
  return {ctx, page, errs};
}
(async () => {
  const browser = await chromium.launch();
  let {ctx, page, errs} = await boot(browser, {config: CONFIG_OFF});
  ok('unconfigured copy says what to do', /config\.js/.test(await page.$eval('#gateMsg', e => e.textContent)));
  await ctx.close();

  ({ctx, page, errs} = await boot(browser, {signedIn: false}));
  ok('signed out: sign-in button shown, nothing loaded', !(await page.$eval('#btnSignIn', e => e.hidden)) && (await page.evaluate(() => __graphCalls.length)) === 0);
  await page.click('#btnSignIn'); ok('sign-in button starts the Microsoft login', (await page.evaluate(() => __loginCalls)) === 1);
  await ctx.close();

  ({ctx, page, errs} = await boot(browser));
  await page.waitForTimeout(1000);
  let drive = await page.evaluate(() => __drive);
  ok('first run creates index.json with one default campaign', drive['index.json'] && Object.keys(drive['index.json'].content.campaigns).length === 1, Object.keys(drive).join(','));
  ok('header shows the signed-in user and OneDrive storage', /Tester · andrea\.bakalli@cludo\.com · data in your OneDrive/.test(await page.$eval('#who', e => e.textContent)), await page.$eval('#who', e => e.textContent));
  ok('sign-out button visible', !(await page.$eval('#btnSignOut', e => e.hidden)));
  fs.writeFileSync(require('path').join(require('os').tmpdir(), 'synthetic.csv'), 'first_name;last_name;email;organisation;title\nAna;Prueba;ana.prueba@uni-a.example;Universidad A;Responsable web\nBruno;Test;bruno@uni-b.example;Universidad B;Director de comunicación\nCarla;Demo;carla@uni-c.example;Universidad C;Profesora\n');
  await page.click('#btnImport'); await page.setInputFiles('#importFile', require('path').join(require('os').tmpdir(), 'synthetic.csv')); await page.waitForTimeout(800);
  const msg = await page.$eval('#importMsg', e => e.textContent);
  drive = await page.evaluate(() => __drive); const cid = Object.keys(drive['index.json'].content.campaigns)[0];
  ok('CSV import (semicolons) writes prospects-<cid>.json', /3 added/.test(msg) && drive[`prospects-${cid}.json`] && Object.keys(drive[`prospects-${cid}.json`].content.items).length === 3, msg);
  const areas = await page.evaluate(() => Object.values(P).map(c => [c.title, areaFor(c), needsReview(c)]));
  ok('area guessed from the job title; academic title flagged for review', JSON.stringify(areas) === JSON.stringify([['Responsable web', 'la web', false], ['Director de comunicación', 'la comunicación', false], ['Profesora', 'la comunicación y la web', true]]), JSON.stringify(areas));
  const drive2 = await page.evaluate(() => __drive); await ctx.close();

  const subj = o => 'Experiencia de búsqueda en la web de ' + o;
  const mail = {
    sentitems: [
      {id: 's1', conversationId: 'conv-A', subject: subj('Universidad A'), from: A(me), toRecipients: [A('Ana.Prueba@uni-a.example')], sentDateTime: '2026-10-05T09:00:00Z', receivedDateTime: '2026-10-05T09:00:01Z', bodyPreview: 'Hola Ana'},
      {id: 's2', conversationId: 'conv-A', subject: subj('Universidad A'), from: A(me), toRecipients: [A('ana.prueba@uni-a.example')], sentDateTime: '2026-10-05T09:00:16Z', receivedDateTime: '2026-10-05T09:00:17Z', bodyPreview: 'Hola Ana (resend)'},
      {id: 's3', conversationId: 'conv-B', subject: subj('Universidad B'), from: A(me), toRecipients: [A('bruno@uni-b.example')], sentDateTime: '2026-10-05T09:10:00Z', receivedDateTime: '2026-10-05T09:10:01Z', bodyPreview: 'Hola Bruno'},
      {id: 's4', conversationId: 'conv-C', subject: subj('Universidad C'), from: A(me), toRecipients: [A('carla@uni-c.example')], sentDateTime: '2026-10-05T09:20:00Z', receivedDateTime: '2026-10-05T09:20:01Z', bodyPreview: 'Hola Carla'},
    ],
    inbox: [
      ...Array.from({length: 60}, (_, k) => ({id: 'n' + k, conversationId: 'news' + k, subject: 'Newsletter ' + k, from: A('news@example.net'), toRecipients: [A(me)], sentDateTime: '2026-10-05T10:00:00Z', receivedDateTime: '2026-10-05T10:' + String(k).padStart(2, '0') + ':00Z', bodyPreview: ''})),
      {id: 'r1', conversationId: 'conv-B', subject: 'RE: ' + subj('Universidad B'), from: A('marketing.central@otra-entidad.example'), toRecipients: [A(me)], sentDateTime: '2026-10-05T11:00:00Z', receivedDateTime: '2026-10-05T11:00:05Z', bodyPreview: 'Le paso a mi compañera'},
      {id: 'ndr', conversationId: 'conv-C', subject: 'Undeliverable: ' + subj('Universidad C'), from: A('MicrosoftExchange329e71ec88ae4615bbc36ab6ce41109e@cludo.com'), toRecipients: [A('carla@uni-c.example')], sentDateTime: '2026-10-05T09:20:30Z', receivedDateTime: '2026-10-05T09:20:31Z', bodyPreview: 'Delivery has failed to these recipients or groups:\r\n\r\ncarla@uni-c.example\r\nThis message couldn\'t be delivered'},
    ]
  };
  ({ctx, page, errs} = await boot(browser, {drive: drive2, mail}));
  await page.waitForFunction(() => /Synced/.test(document.querySelector('#syncStatus').textContent), null, {timeout: 10000}).catch(() => {});
  const n = await page.evaluate(() => Object.keys(P).length);
  ok('reload loads the prospects back from OneDrive', n === 3, n + ' loaded · status: ' + await page.$eval('#syncStatus', e => e.textContent));
  const st = await page.evaluate(() => Object.fromEntries(Object.values(P).map(c => [c.first_name, info(c)])));
  ok('sync: a resend 16 s later counts once → Waiting for Follow-up 1', st.Ana.status === 'waiting' && st.Ana.n === 1, JSON.stringify([st.Ana.status, st.Ana.n]));
  ok('sync: reply from another address matched by conversation id → Replied', st.Bruno.status === 'replied', st.Bruno.status);
  ok('sync: bounce voids the send → To contact, delivery failed', st.Carla.status === 'todo' && st.Carla.bounced === 1, JSON.stringify([st.Carla.status, st.Carla.bounced]));
  const calls = await page.evaluate(() => __graphCalls.filter(c => /mailFolders/.test(c.path)).map(c => decodeURIComponent(c.path)).map(p => /\$skip=/.test(p) ? p.replace(/.*\$skip=(\d+).*/, 'skip=$1') : p.replace(/^\/v1\.0\/me\/mailFolders\/(\w+)\/messages\?.*/, '$1')));
  ok('sync: Inbox follows @odata.nextLink (62 emails → 2 pages)', calls.join(',') === 'sentitems,inbox,skip=50', calls.join(','));
  const filt = await page.evaluate(() => decodeURIComponent(__graphCalls.find(c => /sentitems/.test(c.path)).path));
  ok('sync: Graph query filters on sentDateTime with a matching $orderby', /\$filter=sentDateTime ge 2026-10-05T00:00:00Z&\$orderby=sentDateTime desc&\$top=50/.test(filt), filt);
  // conflict: another window saved Carla's city first; we mark Ana as sent → both survive
  await page.evaluate(() => { const f = __drive[Object.keys(__drive).find(k => k.startsWith('prospects-'))]; f.etag = 'W/"other-window"'; f.content.items['carla_uni_c_example'].city = 'Madrid'; });
  await page.click('#tabs [data-tab="waiting"]'); await page.click('.card .head'); await page.click('.card [data-act="manualSent"]'); await page.waitForTimeout(800);
  const after = await page.evaluate(() => { const f = __drive[Object.keys(__drive).find(k => k.startsWith('prospects-'))]; return {carlaCity: f.content.items['carla_uni_c_example'].city, anaManual: (f.content.items['ana_prueba_uni_a_example'].manual.sent || []).length, localCarla: P['carla@uni-c.example'].city, puts: __graphCalls.filter(c => c.m === 'PUT').length}; });
  ok('conflict (412): our change merged onto the other window\'s copy, nothing lost', after.carlaCity === 'Madrid' && after.anaManual === 1 && after.localCarla === 'Madrid', JSON.stringify(after));
  await page.click('#btnSettings'); await page.click('#btnNewCampaign'); await page.waitForTimeout(600);
  const idx = await page.evaluate(() => __drive['index.json'].content);
  ok('new campaign saved and remembered in index.json profile', Object.keys(idx.campaigns).length === 2 && idx.profile.campaignId && idx.campaigns[idx.profile.campaignId].name === 'New campaign', JSON.stringify(idx.profile));
  const backup = await page.evaluate(() => JSON.parse(backupJson()));
  ok('backup JSON has type, campaign and prospects', backup.type === 'cludo-outreach-backup' && backup.campaign && Array.isArray(backup.prospects));
  await page.click('#btnSignOut'); ok('sign-out calls MSAL logout', (await page.evaluate(() => __logoutCalls)) === 1);
  ok('no JS errors', errs.length === 0, errs.join(' | '));
  await ctx.close(); await browser.close(); console.log(R.join('\n'));
})().catch(e => { console.error(e); process.exit(1); });
