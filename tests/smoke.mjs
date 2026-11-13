/* eat + drink smoke tests — node tests/smoke.mjs   (npm i jsdom)
   Loads index.html in jsdom with MSAL and Graph stubbed, drives the real
   handlers, and checks the rows the app would have written. Nothing touches
   the network or the live workbook. */
import { JSDOM } from 'jsdom';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, '..', 'index.html'), 'utf8')
  .replace(/<script[^>]*src=[^>]*><\/script>/g, '');
// the running build's own version, so the update check stays quiet in tests
const VERSION = (/const APP_VERSION = "([^"]+)"/.exec(html) || [])[1] || 'test';

/* `stale: n` models what Graph actually does: for the n reads after a write it
   still serves the pre-write copy of the table. That is what left a deleted
   place on the list and a place just entered missing from it. */
/* `geocode(q)` answers the place lookup (null = no match); `fix` is what the
   browser's geolocation returns - an array is tried in turn, a number is an
   error code. Every lookup lands in state.geocoded. */
/* `asleep` makes every workbook call answer the way Microsoft's Excel service does
   when it has stopped answering for the account (state.asleep can be flipped back);
   `snapshot` seeds the copy the app keeps on the device; `osm` is what Overpass
   returns (null = it does not answer). */
/* `google(kind, arg)` answers Google's Places API: kind 'search' gets the request body,
   'details' the place ID; a number is the HTTP status it fails with. `googleKey` puts a
   key on the device. Every request lands in state.google. */
function app({ venues = [], menu = [], stale = 0, geocode = () => null, fix = null, asleep = false, snapshot = null, osm = null, google = null, googleKey = '' } = {}){
  const state = { venues: venues.map(r => r.slice()), menu: menu.map(r => r.slice()), calls: [], geocoded: [], fixOpts: [], asleep, osmAsked: 0, google: [] };
  let snapV = state.venues.map(r => r.slice()), snapM = state.menu.map(r => r.slice()), lag = 0;
  const rows = v => ({ value: v.map((r, i) => ({ index: i, values: [r] })) });
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    url: 'https://nameisredacted.github.io/food-drink-app/',
    beforeParse(w){
      if(snapshot) w.localStorage.setItem('eatdrink.snapshot', JSON.stringify(snapshot));
      if(googleKey) w.localStorage.setItem('eatdrink.googleKey', googleKey);
      w.msal = { PublicClientApplication: class {
        async initialize(){} async handleRedirectPromise(){ return null; }
        getAllAccounts(){ return [{ username: 'test' }]; }
        async acquireTokenSilent(){ return { accessToken: 'token' }; }
      } };
      w.fetch = async (url, opts = {}) => {
        const u = String(url), m = (opts.method || 'GET').toUpperCase();
        if(u.includes('version.txt')) return { ok: true, status: 200, async text(){ return VERSION; } };
        if(u.includes('census.gov')){ state.geocoded.push('census'); throw new TypeError('Failed to fetch'); }
        if(u.includes('nominatim')){
          const q = decodeURIComponent(u.split('&q=')[1] || '');
          state.geocoded.push(q);
          const hit = geocode(q);
          return { ok: true, status: 200, async json(){ return hit ? [hit] : []; } };
        }
        if(u.includes('places.googleapis.com')){
          const h = opts.headers || {};
          const kind = u.includes(':searchText') ? 'search' : 'details';
          const arg = kind === 'search' ? JSON.parse(opts.body) : decodeURIComponent(u.split('/places/')[1]);
          state.google.push({ kind, arg, mask: h['X-Goog-FieldMask'], key: h['X-Goog-Api-Key'] });
          const r = google ? google(kind, arg, state.google.length) : 400;
          if(typeof r === 'number')
            return { ok: false, status: r, async json(){ return { error: { message: r === 429 ? 'Quota exceeded for quota metric' : 'API key not valid' } }; } };
          return { ok: true, status: 200, async json(){ return r; } };
        }
        if(u.includes('/api/interpreter')){
          state.osmAsked++;
          return { ok: !!osm, status: osm ? 200 : 504, async json(){ return { elements: osm }; } };
        }
        if(u.includes('/me/drive/root:') && !u.includes(':/workbook'))   // the drive item, not the workbook
          return { ok: true, status: 200, async json(){ return { webUrl: 'https://onedrive.live.com/the-workbook' }; } };
        if(state.asleep){
          const body = '{"error":{"code":"FileOpenUserUnauthorized","message":"You do not have permissions to open this file in the browser."}}';
          return { ok: false, status: 401, async text(){ return body; }, async json(){ return JSON.parse(body); } };
        }
        if(u.includes('/createSession')){ state.sessions = (state.sessions || 0) + 1; return { ok: true, status: 201, async json(){ return { id: 'session-1' }; }, async text(){ return ''; } }; }
        if(u.includes('/closeSession')){ state.sessionsClosed = (state.sessionsClosed || 0) + 1; return { ok: true, status: 204, async json(){ return null; }, async text(){ return ''; } }; }
        if((opts.headers || {})['workbook-session-id']) state.inSession = (state.inSession || 0) + 1;
        const isMenu = u.includes('MenuDetails');
        if(m !== 'GET'){
          if(stale && !lag){ snapV = state.venues.map(r => r.slice()); snapM = state.menu.map(r => r.slice()); }
          const idx = (/itemAt\(index=(\d+)\)/.exec(u) || [])[1];
          const kind = u.includes('/columns') ? 'addcol' : (isMenu ? 'menu' : 'venue');
          state.calls.push({ method: m, kind, body: opts.body && JSON.parse(opts.body), index: idx && +idx });
          const table = isMenu ? state.menu : state.venues;
          if(kind === 'addcol') state.venues = state.venues.map(r => r.concat(['']));
          else if(m === 'PATCH') table[+idx] = JSON.parse(opts.body).values[0];
          else if(m === 'DELETE') table.splice(+idx, 1);
          else if(m === 'POST') table.push(JSON.parse(opts.body).values[0]);
          if(stale) lag = stale;
          return { ok: true, status: 201, async json(){ return {}; }, async text(){ return ''; } };
        }
        // one row read back by index (the lookup checks a row before writing it)
        const gIdx = (/itemAt\(index=(\d+)\)/.exec(u) || [])[1];
        if(gIdx !== undefined && !isMenu){
          const r = state.venues[+gIdx];
          return { ok: true, status: 200, async json(){ return r ? { index: +gIdx, values: [r] } : null; } };
        }
        if(lag > 0){
          const old = isMenu ? rows(snapM) : rows(snapV);
          if(!isMenu) lag--;      // one venue read consumed
          return { ok: true, status: 200, async json(){ return old; }, async text(){ return JSON.stringify(old); } };
        }
        const body = isMenu ? rows(state.menu) : rows(state.venues);
        return { ok: true, status: 200, async json(){ return body; }, async text(){ return JSON.stringify(body); } };
      };
      w.alert = m => { state.alerted = m; }; w.confirm = () => true; w.prompt = () => 'Other';
      const fixes = Array.isArray(fix) ? fix.slice() : [fix];
      Object.defineProperty(w.navigator, 'geolocation', { configurable: true, value: {
        getCurrentPosition(ok, bad, opts){
          state.fixOpts.push(opts);
          const f = fixes.length > 1 ? fixes.shift() : fixes[0];
          setTimeout(() => typeof f === 'number'
            ? bad({ code: f, message: 'code ' + f })
            : ok({ coords: { latitude: f.lat, longitude: f.lng } }), 0);
        } } });
    }
  });
  const w = dom.window;
  return {
    state, w,
    $: id => w.document.getElementById(id),
    click: el => el.dispatchEvent(new w.MouseEvent('click', { bubbles: true })),
    fire: (el, type) => el.dispatchEvent(new w.Event(type, { bubbles: true })),
    enter: el => el.dispatchEvent(new w.KeyboardEvent('keydown', { key: 'Enter', bubbles: true })),
    settle: (ms = 60) => new Promise(r => setTimeout(r, ms)),
  };
}

const V = {
  zuni:    ['Zuni Cafe','American','SF','','','','','','','','',''],
  philzSF: ['Philz Coffee','Coffee','SF','','','','','','','','',''],
  philzPA: ['Philz Coffee','Coffee','Palo Alto','','','','','','','','',''],
  dupeA:   ['Dumpling Home','Chinese','SF','Scallion Pancake','','','','298 Gough St','','','',''],
  dupeB:   ['Dumpling Home','Chinese','SF','Scallion Pancake; pork xlb','y','note','x','','','','',''],
  umbrella:['Tartine Bakery','Bakery','Multiple Locations','','','','x','','','','',''],
  blank:   ['Hang Ah','Chinese','','','','','x','','','','',''],
};
const M = {
  zuni:   ['Zuni Cafe','SF','','Mains','roast chicken'],
  stray:  ['Hang Ah','North Bay, CA','','Mains','har gow'],
  fake:   ['Tartine Bakery','Multiple Locations','','Bakery','morning bun'],
};

let failures = 0;
const results = [];
const check = (name, fn) => {
  try { fn(); results.push('ok   ' + name); }
  catch(e){ failures++; results.push('FAIL ' + name + ' :: ' + e.message); }
};

/* ---------- the ordered sheet ---------- */
{
  const a = app({ venues: [V.zuni, V.philzSF, V.philzPA], menu: [M.zuni] });
  await a.settle(400);
  a.click(a.$('logLink'));

  check('opens empty, no panels', () => {
    if(a.$('logBody').innerHTML !== '') throw new Error('list pre-populated');
    if(a.$('lg_new').style.display !== 'none' || a.$('lg_multi').style.display !== 'none')
      throw new Error('a panel was open');
  });
  check('search filters the log', () => {
    a.$('logSearch').value = 'chicken'; a.fire(a.$('logSearch'), 'input');
    if(!/roast chicken/.test(a.$('logBody').innerHTML)) throw new Error('no match shown');
  });
  check('suggestions stay quiet until typed', () => {
    const el = a.$('lg_venue'); el.value = ''; a.fire(el, 'focus');
    const box = el.parentNode.querySelector('.ac-box');
    if(box && box.className.includes('open')) throw new Error('opened on empty focus');
  });

  a.$('lg_venue').value = 'Juans'; a.fire(a.$('lg_venue'), 'input'); await a.settle();
  check('copy: unknown name', () => {
    if(a.$('lg_hint').textContent !== 'nothing called Juans yet — add it?')
      throw new Error(a.$('lg_hint').textContent);
  });
  a.$('lg_venue').value = 'phi'; a.fire(a.$('lg_venue'), 'input'); await a.settle();
  check('copy: completed name, both locations', () => {
    if(a.$('lg_venue').value !== 'Philz Coffee') throw new Error('inline completion missing');
    if(a.$('lg_hint').textContent !== 'Philz Coffee is already on the list in SF, Palo Alto')
      throw new Error(a.$('lg_hint').textContent);
  });

  a.$('lg_what').value = 'mint mojito'; a.click(a.$('lg_newLoc'));
  a.$('lg_location').value = 'SF'; a.fire(a.$('lg_location'), 'input'); await a.settle();
  check('copy: existing branch, brand question hidden', () => {
    if(a.$('lg_hint').textContent !== 'Philz Coffee in SF is already on the list — this just adds what is new')
      throw new Error(a.$('lg_hint').textContent);
    if(a.$('lg_brandRow').style.display !== 'none') throw new Error('brand row shown');
  });
  a.state.calls.length = 0;
  a.click(a.$('lg_newAdd')); await a.settle(200);
  check('existing branch: item logged, no second venue row', () => {
    if(a.state.calls.some(c => c.method === 'POST' && c.kind === 'venue')) throw new Error('made a duplicate row');
    if(!a.state.menu.some(r => r[4] === 'mint mojito' && r[1] === 'SF')) throw new Error('item not logged');
  });

  a.$('lg_venue').value = 'Philz Coffee'; a.$('lg_what').value = 'iced tea';
  a.click(a.$('lg_newLoc')); a.$('lg_location').value = 'Berkeley'; a.fire(a.$('lg_location'), 'input');
  await a.settle();
  a.state.calls.length = 0;
  a.click(a.$('lg_newAdd')); await a.settle(300);
  check('new branch: row created and the brand marked chain', () => {
    if(a.state.calls.filter(c => c.method === 'POST' && c.kind === 'venue').length !== 1)
      throw new Error('expected one new venue row');
    const philz = a.state.venues.filter(r => r[0] === 'Philz Coffee');
    if(philz.length !== 3) throw new Error('rows: ' + philz.length);
    if(!philz.every(r => r[11] === 'x')) throw new Error('not every branch chained');
  });

  a.$('lg_venue').value = 'Philz Coffee'; a.$('lg_what').value = 'tea';
  a.w.document.querySelectorAll('#logSheet .ac-box').forEach(b => b.classList.remove('open'));
  a.enter(a.$('lg_what'));
  check('repeated name: picker with chain marks and an escape hatch', () => {
    if(a.$('lg_multi').style.display === 'none') throw new Error('picker did not open');
    const labels = [...a.w.document.querySelectorAll('#lg_multiPicks .locpick')].map(l => l.textContent.trim());
    if(labels.length !== 4) throw new Error('options: ' + JSON.stringify(labels));
    if(!labels.some(l => /somewhere else/.test(l))) throw new Error('no "somewhere else"');
    if(!labels.some(l => /chain/.test(l))) throw new Error('chain not shown');
  });
  a.state.calls.length = 0;
  a.$('lg_multiAll').checked = true; a.click(a.$('lg_multiAdd')); await a.settle(300);
  check('log at every location writes one row per branch', () => {
    if(a.state.calls.filter(c => c.method === 'POST' && c.kind === 'menu').length !== 3)
      throw new Error('menu posts: ' + a.state.calls.filter(c => c.kind === 'menu').length);
  });
}

/* ---------- data check repairs ---------- */
{
  const a = app({ venues: [V.dupeA, V.dupeB, V.umbrella, V.philzSF, V.blank], menu: [M.stray, M.fake] });
  await a.settle(400);
  a.w.eval('openDiag()');

  check('every problem is counted', () => {
    if(!/Merge 1 duplicate/.test(a.$('diagMergeDupes').textContent)) throw new Error(a.$('diagMergeDupes').textContent);
    if(!/Fix 1 chain row/.test(a.$('diagFixUmbrellas').textContent)) throw new Error(a.$('diagFixUmbrellas').textContent);
    if(!/Attach 1 logged item/.test(a.$('diagAttachStrays').textContent)) throw new Error(a.$('diagAttachStrays').textContent);
  });

  a.click(a.$('diagMergeDupes')); await a.settle(300);
  check('merge keeps one row and folds the fields in', () => {
    const dh = a.state.venues.filter(r => r[0] === 'Dumpling Home');
    if(dh.length !== 1) throw new Error('rows left: ' + dh.length);
    const [r] = dh;
    if(r[4] !== 'y' || r[5] !== 'note' || r[6] !== 'x') throw new Error('fields lost: ' + JSON.stringify(r));
    if(!/298 Gough/.test(r[7])) throw new Error('address lost');
    if(!/pork xlb/.test(r[3]) || (r[3].match(/Scallion Pancake/g) || []).length !== 1)
      throw new Error('to-order merge: ' + r[3]);
  });

  a.click(a.$('diagFixUmbrellas')); await a.settle(400);
  check('made-up location becomes a chain flag', () => {
    const t = a.state.venues.find(r => r[0] === 'Tartine Bakery');
    if(t[2] !== '' || t[11] !== 'x') throw new Error('row: ' + JSON.stringify(t));
    const m = a.state.menu.find(r => r[0] === 'Tartine Bakery');
    if(m[1] !== '') throw new Error('logged item left behind: ' + JSON.stringify(m));
  });

  a.click(a.$('diagAttachStrays')); await a.settle(300);
  check('stray logged item points at its place', () => {
    const h = a.state.menu.find(r => r[0] === 'Hang Ah');
    if(h[1] !== '' || h[4] !== 'har gow') throw new Error(JSON.stringify(h));
  });

  a.w.eval('renderDiag()');
  check('data check comes back clean', () => {
    for(const id of ['diagMergeDupes','diagFixUmbrellas','diagAttachStrays'])
      if(!a.$(id).disabled) throw new Error(id + ' still offering work');
  });
}

/* ---------- an 11-column workbook (no Chain column yet) ---------- */
{
  const eleven = r => r.slice(0, 11);
  const a = app({ venues: [eleven(V.philzSF), eleven(V.philzPA), eleven(V.zuni)], menu: [M.zuni] });
  await a.settle(400);
  check('reads and writes at the table width', () => {
    if(a.w.eval('hasChainCol()')) throw new Error('claims a Chain column');
    if(a.w.eval('objToRow({name:"x"}).length') !== 11) throw new Error('would write 12 cells');
    if(a.w.eval('unmarkedChainRows().length') !== 2) throw new Error('chain candidates missed');
  });
  a.w.eval('openDiag()');
  a.click(a.$('diagMarkChains')); await a.settle(400);
  check('marking adds the column once, then flags the rows', () => {
    const adds = a.state.calls.filter(c => c.kind === 'addcol');
    if(adds.length !== 1 || adds[0].body.name !== 'Chain') throw new Error('column add: ' + JSON.stringify(adds));
    const philz = a.state.venues.filter(r => r[0] === 'Philz Coffee');
    if(!philz.every(r => r[11] === 'x')) throw new Error('rows: ' + JSON.stringify(philz));
    if(a.state.venues.find(r => r[0] === 'Zuni Cafe')[11] === 'x') throw new Error('single-location row flagged');
  });
}

/* ---------- regressions found in review ---------- */
{
  const a = app({ venues: [['Philz Coffee','Coffee','SF','','','','','','','','','x'],
                           ['Philz Coffee','Coffee','Palo Alto','','','','','','','','','x']],
                  menu: [['Philz Coffee','Palo Alto','','Coffee','mint mojito']] });
  await a.settle(400);
  a.w.eval("openForm(venues.find(v => v.location === 'SF'))");
  a.$('f_category').value = 'Coffee & Tea';
  a.state.calls.length = 0;
  a.click(a.$('saveForm')); await a.settle(300);
  check('editing a venue updates its row instead of adding one', () => {
    if(a.state.calls.some(c => c.method === 'POST' && c.kind === 'venue'))
      throw new Error('the edit created a second row');
    if(a.state.venues.length !== 2) throw new Error('row count: ' + a.state.venues.length);
    const row = a.state.venues.find(r => r[2] === 'SF');
    if(row[1] !== 'Coffee & Tea') throw new Error('edit not applied: ' + JSON.stringify(row));
    if(row[11] !== 'x') throw new Error('chain wiped by the edit form: ' + JSON.stringify(row));
  });

  a.click(a.$('logLink'));
  a.$('logSearch').value = 'mojito'; a.fire(a.$('logSearch'), 'input');
  check('log opens the branch the item belongs to', () => {
    const item = a.w.document.querySelector('#logBody .grp-item');
    if(!item) throw new Error('nothing listed');
    a.click(item);
    const title = a.$('menuPageTitle').textContent;
    const sub = a.w.eval('menuPageVenue.location');
    if(sub !== 'Palo Alto') throw new Error('opened ' + title + ' in ' + sub);
  });
}

/* ---------- closed places: red, read-only, never deleted ---------- */
{
  const closedRow = ['Bellota','Spanish','San Francisco, CA','','y','','x','888 Brannan St','','','','','closed 2026'];
  const openRow   = ['Zuni Cafe','American','San Francisco, CA','','y','','','','','','','',''];
  const a = app({ venues: [closedRow, openRow], menu: [['Bellota','San Francisco, CA','','Mains','jamon']] });
  await a.settle(400);

  check('closed row is flagged in the model', () => {
    if(!a.w.eval("isClosed(venues.find(v => v.name === 'Bellota'))")) throw new Error('not read as closed');
    if(a.w.eval("isClosed(venues.find(v => v.name === 'Zuni Cafe'))")) throw new Error('open row read as closed');
  });
  check('list keeps it, in red, with a tag', () => {
    a.w.eval("state.query = 'a'; render();");
    const el = [...a.w.document.querySelectorAll('#list .item')].find(e => /Bellota/.test(e.textContent));
    if(!el) throw new Error('closed place dropped from the list');
    if(!el.className.includes('closed')) throw new Error('no closed class');
    if(!el.querySelector('.closedtag')) throw new Error('no closed tag');
    // jsdom does not resolve custom properties, so check the token it resolves to
    const colour = a.w.getComputedStyle(el.querySelector('.item-name')).color;
    if(!/--bad|204, 0, 0|#cc0000/.test(colour)) throw new Error('name not red: ' + colour);
    const red = /--bad:\s*(#[0-9a-f]{6})/i.exec(a.w.document.documentElement.innerHTML);
    if(!red || red[1].toLowerCase() !== '#cc0000') throw new Error('the red token moved: ' + (red && red[1]));
  });
  check('detail sheet is a record, not a form', () => {
    a.w.eval("openDetail(venues.find(v => v.name === 'Bellota'))");
    const c = a.$('detailContent');
    if(!c.querySelector('h2.closed')) throw new Error('heading not marked closed');
    if(!/cannot be edited/.test(c.textContent)) throw new Error('no explanation shown');
    if(c.querySelector('.rateBtn')) throw new Error('rating still settable');
    if(c.querySelector('.ate') || c.querySelector('#d_want')) throw new Error('to-order still editable');
    if(a.w.document.getElementById('editBtn')) throw new Error('edit button present');
    if(a.w.document.getElementById('deleteBtn')) throw new Error('delete button present');
    if(!a.w.document.getElementById('openMenuPageBtn')) throw new Error('ordered history not reachable');
    if(!a.w.document.getElementById('closedToggle')) throw new Error('no way to reopen it');
  });
  check('an open place still has its controls', () => {
    a.w.eval("openDetail(venues.find(v => v.name === 'Zuni Cafe'))");
    const c = a.$('detailContent');
    if(!c.querySelector('.rateBtn')) throw new Error('rating gone');
    if(!a.w.document.getElementById('editBtn')) throw new Error('edit gone');
    if(a.$('closedToggle').textContent !== 'mark as closed') throw new Error(a.$('closedToggle').textContent);
  });
  check('the edit form refuses a closed row', () => {
    a.w.eval("openForm(venues.find(v => v.name === 'Bellota'))");
    if(a.$('formSheet').className.includes('open')) throw new Error('form opened for a closed place');
  });
  check('its ordered page keeps history but takes nothing new', () => {
    a.w.eval("openMenuPage(venues.find(v => v.name === 'Bellota'))");
    if(!/jamon/.test(a.$('menuPageBody').innerHTML)) throw new Error('history lost');
    if(a.$('mp_new').style.display !== 'none') throw new Error('+ add items still offered');
    if([...a.w.document.querySelectorAll('#menuPageBody .mp-row button')].some(b => b.style.display !== 'none'))
      throw new Error('remove buttons still live');
  });
  a.w.eval('closeMenuPage()');
  a.click(a.$('logLink'));
  a.$('lg_venue').value = 'Bellota'; a.$('lg_what').value = 'croquetas';
  a.state.calls.length = 0;
  a.enter(a.$('lg_what'));
  await a.settle(200);
  check('logging at a closed place writes nothing', () => {
    if(a.state.calls.some(c => c.method === 'POST')) throw new Error('logged anyway');
  });
  check('the dice never picks a closed place', () => {
    a.click(a.$('closeLog'));
    a.w.eval("state.query = ''; state.category = ''; state.ratings = new Set(['y']); state.near = null;");
    for(let i = 0; i < 25; i++){
      a.click(a.$('diceChip'));
      if(/Bellota/.test(a.$('detailContent').innerHTML)) throw new Error('dice landed on a closed place');
    }
  });
  a.w.eval('closeDetail(); openDetail(venues.find(v => v.name === "Bellota"))');
  a.state.calls.length = 0;
  a.click(a.$('closedToggle'));
  await a.settle(300);
  check('reopening clears the Closed cell and nothing else', () => {
    const p = a.state.calls.filter(c => c.method === 'PATCH' && c.kind === 'venue');
    if(p.length !== 1) throw new Error('patches: ' + p.length);
    const row = p[0].body.values[0];
    if(row[12] !== '') throw new Error('closed cell not cleared: ' + JSON.stringify(row));
    if(row[0] !== 'Bellota' || row[4] !== 'y') throw new Error('other fields disturbed: ' + JSON.stringify(row));
  });
}

/* ---------- the closed label follows the place everywhere ---------- */
{
  const a = app({ venues: [
    ['Bellota','Spanish','San Francisco, CA','','y','','x','','','','','','closed 2026'],
    ['Zuni Cafe','American','San Francisco, CA','','y','','x','','','','','',''],
  ], menu: [
    ['Bellota','San Francisco, CA','','Mains','jamon'],
    ['Zuni Cafe','San Francisco, CA','','Mains','roast chicken'],
  ] });
  await a.settle(400);
  const tag = el => el && el.querySelector('.closedtag');

  check('the list shows the label', () => {
    a.w.eval("state.query = 'bellota'; render();");
    const el = [...a.w.document.querySelectorAll('#list .item')].find(x => /Bellota/.test(x.textContent));
    if(!tag(el)) throw new Error('no label in the list');
  });
  check('the detail sheet shows the label', () => {
    a.w.eval("openDetail(venues.find(v => v.name === 'Bellota'))");
    if(!tag(a.$('detailContent').querySelector('h2'))) throw new Error('no label on the detail heading');
  });
  check('the ordered page shows the label', () => {
    a.w.eval("openMenuPage(venues.find(v => v.name === 'Bellota'))");
    const t = a.$('menuPageTitle');
    if(!tag(t)) throw new Error('no label on the ordered page');
    if(!t.className.includes('closed')) throw new Error('ordered page title not marked closed');
  });
  check('the global log shows the label', () => {
    a.w.eval('closeMenuPage()');
    a.$('logSearch').value = 'jamon';
    a.fire(a.$('logSearch'), 'input');
    const head = a.$('logBody').querySelector('.log-head');
    if(!head) throw new Error('log did not render');
    if(!tag(head)) throw new Error('no label in the global log');
  });
  check('an open place carries no label anywhere', () => {
    a.w.eval("openDetail(venues.find(v => v.name === 'Zuni Cafe'))");
    if(tag(a.$('detailContent'))) throw new Error('open place labelled closed');
    a.w.eval("openMenuPage(venues.find(v => v.name === 'Zuni Cafe'))");
    if(tag(a.$('menuPageTitle').parentNode) && a.$('menuPageTitle').querySelector('.closedtag'))
      throw new Error('open place labelled on its ordered page');
  });
  check('the data check no longer offers a bulk closure button', () => {
    a.w.eval('closeMenuPage(); openDiag()');
    if(a.$('diagMarkClosed')) throw new Error('the Mark closures button is still there');
    if(/Known to have closed/.test(a.$('diagBody').innerHTML)) throw new Error('closure line still in the data check');
    if(!/Marked closed/.test(a.$('diagBody').innerHTML)) throw new Error('the closed count should still be reported');
  });
}

/* ---------- closed places are not favourites ---------- */
{
  const a = app({ venues: [
    ['Bellota','Spanish','San Francisco, CA','','y','','','','','','','','closed 2026'],
    ['Sociale','Italian','San Francisco, CA','','y','','','','','','','',''],
    ['Bar Brucato','Bar','San Francisco, CA','','','','','','','','','','closed Jun 2026']
  ] });
  await a.settle(400);
  const names = () => [...a.w.document.querySelectorAll('#list .item-name')].map(e => e.textContent.replace('closed','').trim());
  check('the y filter hides closed places', () => {
    a.w.eval("state.ratings = new Set(['y']); render();");
    const shown = names();
    if(shown.includes('Bellota')) throw new Error('closed favourite still listed: ' + JSON.stringify(shown));
    if(!shown.includes('Sociale')) throw new Error('open favourite missing: ' + JSON.stringify(shown));
  });
  check('an unfiltered search still finds the record', () => {
    a.w.eval("state.ratings = new Set(); state.query = 'bellota'; render();");
    if(!names().includes('Bellota')) throw new Error('record lost from search: ' + JSON.stringify(names()));
  });
  check('the dice never picks a closed place', () => {
    a.w.eval("state.query = ''; state.ratings = new Set(); render();");
    for(let i = 0; i < 40; i++){
      a.click(a.$('diceChip'));
      const title = a.$('detailContent').querySelector('h2').textContent;
      if(/Bellota|Bar Brucato/.test(title)) throw new Error('dice picked ' + title);
    }
    a.w.eval('closeDetail()');
  });
}

/* ---------- no twin rows, no search dropdown ---------- */
{
  const a = app({ venues: [V.zuni, V.philzSF, V.philzPA], menu: [] });
  await a.settle(400);

  a.w.eval('openForm(null)');
  a.$('f_name').value = 'Zuni Cafe';
  a.$('f_location').value = 'SF';
  a.$('f_toOrder').value = 'roast chicken';
  a.state.calls.length = 0;
  a.click(a.$('saveForm')); await a.settle(300);
  check('adding a name+location already on the list updates that row', () => {
    if(a.state.calls.some(c => c.method === 'POST' && c.kind === 'venue'))
      throw new Error('a twin row was appended');
    if(a.state.venues.length !== 3) throw new Error('row count: ' + a.state.venues.length);
    const row = a.state.venues.find(r => r[0] === 'Zuni Cafe');
    if(row[3] !== 'roast chicken') throw new Error('typed value lost: ' + JSON.stringify(row));
    if(row[1] !== 'American') throw new Error('existing field blanked: ' + JSON.stringify(row));
  });

  a.w.eval('openForm(null)');
  a.$('f_name').value = 'Tartine Bakery';
  a.$('f_location').value = 'SF';
  a.state.calls.length = 0;
  a.click(a.$('saveForm')); await a.settle(300);
  check('a genuinely new name+location still adds a row', () => {
    if(!a.state.calls.some(c => c.method === 'POST' && c.kind === 'venue'))
      throw new Error('nothing was added');
    if(a.state.venues.length !== 4) throw new Error('row count: ' + a.state.venues.length);
  });

  check('the primary search bar has no dropdown', () => {
    if(a.w.document.getElementById('suggestBox')) throw new Error('suggest box still in the markup');
    if(/showSuggest|hideSuggest/.test(html)) throw new Error('suggest handlers still wired up');
  });
  check('typing still filters the list live', () => {
    const box = a.$('searchBox');
    box.value = 'zuni'; a.fire(box, 'input');
    const shown = a.$('list').textContent;
    if(!/Zuni Cafe/.test(shown)) throw new Error('match missing');
    if(/Philz/.test(shown)) throw new Error('non-match still listed');
  });
}

/* ---------- a write the read has not caught up with ---------- */
{
  // 20 stale reads: the read never agrees inside the app's retry window, so
  // only the write-through overlay can keep the list honest
  const a = app({ venues: [V.zuni, V.philzSF, V.philzPA], menu: [M.zuni], stale: 20 });
  await a.settle(400);

  a.w.eval("openDetail(venues.find(v => v.name === 'Zuni Cafe'))");
  a.state.calls.length = 0;
  a.click(a.$('deleteBtn')); await a.settle(1600);   // three reads, none of them fresh
  check('delete takes the place off the list at once', () => {
    if(!a.state.calls.some(c => c.method === 'DELETE')) throw new Error('nothing was deleted');
    if(a.w.eval("venues.some(v => v.name === 'Zuni Cafe')")) throw new Error('still in the model');
    a.w.eval("state.query = 'zuni'; render();");
    if(/Zuni Cafe/.test(a.$('list').textContent)) throw new Error('still on the list');
  });
  check('deleting does not disturb anything else', () => {
    if(a.w.eval('venues.length') !== 2) throw new Error('rows: ' + a.w.eval('venues.length'));
  });
}

{
  const a = app({ venues: [V.zuni, V.philzSF], menu: [], stale: 20 });
  await a.settle(400);

  a.w.eval('openForm(null)');
  a.$('f_name').value = 'Juans Place';
  a.$('f_location').value = 'SF';
  a.$('f_category').value = 'Peruvian';
  a.state.calls.length = 0;
  a.click(a.$('saveForm')); await a.settle(1600);   // three reads, none of them fresh

  check('a place just entered is usable straight away', () => {
    if(!a.state.calls.some(c => c.method === 'POST' && c.kind === 'venue')) throw new Error('nothing written');
    if(!a.w.eval("venues.some(v => v.name === 'Juans Place')")) throw new Error('not in the model');
    if(!a.w.eval("venueNameOptions().includes('Juans Place')")) throw new Error('not offered as a name');
    if(!/Juans Place/.test(a.$('list').textContent)) throw new Error('not on screen');
  });
  check('the row it is waiting on is never written to', () => {
    if(!a.w.eval("venues.find(v => v.name === 'Juans Place').pending")) throw new Error('test assumes it is still pending');
    if(a.w.eval("liveVenues().some(v => v.name === 'Juans Place')")) throw new Error('a repair could aim at it');
  });
  check('a category typed once is offered from then on', () => {
    if(!a.w.eval("categoryOptions().includes('Peruvian')")) throw new Error('new category not suggested');
    if(a.w.eval("JSON.parse(localStorage.getItem('eatdrink.categories')||'[]').indexOf('Peruvian')") < 0)
      throw new Error('new category not kept');
  });
  check('a category the sheet already has is not duplicated', () => {
    const n = a.w.eval("categoryOptions().filter(c => c.toLowerCase() === 'american').length");
    if(n !== 1) throw new Error('american appears ' + n + ' times');
  });
}

/* ---------- an order at a place that is not on the list yet ---------- */
{
  const a = app({ venues: [V.zuni], menu: [M.zuni] });
  await a.settle(400);

  a.click(a.$('logLink')); await a.settle(200);
  a.$('lg_venue').value = 'Juans Place'; a.fire(a.$('lg_venue'), 'input'); await a.settle();
  a.click(a.$('lg_newLoc')); await a.settle();
  check('the + opens the fields for a place the list does not have', () => {
    if(a.$('lg_new').style.display === 'none') throw new Error('fields still hidden');
  });

  a.$('lg_what').value = 'lomo saltado';
  a.$('lg_location').value = 'SF'; a.fire(a.$('lg_location'), 'input');
  a.$('lg_category').value = 'Peruvian'; await a.settle();
  a.state.calls.length = 0;
  a.click(a.$('lg_newAdd')); await a.settle(400);
  check('the order writes the place and the item together', () => {
    const row = a.state.venues.find(r => r[0] === 'Juans Place');
    if(!row) throw new Error('no venue row');
    if(row[1] !== 'Peruvian' || row[2] !== 'SF') throw new Error('venue row: ' + JSON.stringify(row));
    if(row[6] !== 'x') throw new Error('menu flag not set: ' + JSON.stringify(row));
    if(!a.state.menu.some(r => r[0] === 'Juans Place' && r[4] === 'lomo saltado')) throw new Error('item not logged');
  });
  check('the new place is on the list behind the sheet', () => {
    if(!a.w.eval("venues.some(v => v.name === 'Juans Place')")) throw new Error('not in the model');
    if(!/Juans Place/.test(a.$('list').textContent)) throw new Error('not on screen');
  });
  a.$('lg_venue').value = 'Juans Place'; a.fire(a.$('lg_venue'), 'input'); await a.settle();
  check('a second item goes to it without asking again', () => {
    if(!/already on the list/.test(a.$('lg_hint').textContent)) throw new Error(a.$('lg_hint').textContent);
  });
}

/* ---------- coming back to an app that was left open ---------- */
{
  const a = app({ venues: [
    ['Zuni Cafe','American','San Francisco, CA','','y','','','','','','','',''],
  ], menu: [] });
  await a.settle(400);
  // jsdom reports a document that is never rendered as hidden; the whole point
  // of this behaviour is a page that has just come to the foreground
  Object.defineProperty(a.w.document, 'visibilityState', { get: () => 'visible', configurable: true });
  Object.defineProperty(a.w.document, 'hidden', { get: () => false, configurable: true });

  check('a fresh page does not re-read', () => {
    a.state.calls.length = 0;
    a.w.eval("document.dispatchEvent(new Event('visibilitychange'))");
    if(a.state.calls.length) throw new Error('re-read straight after loading');
  });

  // the other device adds a place while this one sits in the background
  a.state.venues.push(['Nopa','American','San Francisco, CA','','','','','','','','','','']);
  a.w.eval('lastLoadAt = 0');
  a.w.eval("document.dispatchEvent(new Event('visibilitychange'))");
  await a.settle(400);
  check('coming back re-reads the sheet and shows the other device\'s edit', () => {
    if(!a.w.eval("venues.some(v => v.name === 'Nopa')")) throw new Error('not in the model');
    a.w.eval("state.query = 'nopa'; render();");
    if(!/Nopa/.test(a.$('list').textContent)) throw new Error('not on screen');
  });

  check('a refresh never runs while a sheet is open', () => {
    a.w.eval("state.query = ''; render(); openDetail(venues.find(v => v.name === 'Zuni Cafe'))");
    a.state.venues.push(['Delfina','Italian','San Francisco, CA','','','','','','','','','','']);
    a.w.eval('lastLoadAt = 0');
    a.state.calls.length = 0;
    a.w.eval("document.dispatchEvent(new Event('visibilitychange'))");
    if(a.state.calls.length) throw new Error('re-read out from under an open sheet');
  });

  a.w.eval('closeDetail()');
  a.w.eval('lastLoadAt = 0');
  a.w.eval("document.dispatchEvent(new Event('visibilitychange'))");
  await a.settle(400);
  check('the deferred edit arrives after the sheet closes', () => {
    if(!a.w.eval("venues.some(v => v.name === 'Delfina')")) throw new Error('never picked it up');
  });
}

/* ---------- nearest ---------- */
{
  const row = (name, loc, addr, lat, lng, geo, chain = '', closed = '') =>
    [name, 'Food', loc, '', '', '', '', addr, lat, lng, geo, chain, closed];
  const a = app({
    venues: [
      row('Zuni Cafe',  'San Francisco, CA', '1658 Market St', '37.7735', '-122.4222', 'zip'),
      row('Nopa',       'San Francisco, CA', '560 Divisadero St', '37.7749', '-122.4374', 'exact'),
      row('Che Fico',   'San Francisco, CA', '', '', '', ''),                 // added through the app
      row('Philz',      'San Francisco, CA', '', '', '', '', 'x'),            // a chain
      row('Gone',       'San Francisco, CA', '', '', '', '', '', 'closed'),   // closed
      row('Wild Crumb', 'Bozeman, MT', '', '', '', ''),                       // nothing to borrow from
    ],
    geocode: q => /1658 Market/.test(q) ? { lat: '37.7736', lon: '-122.4221', category: 'amenity', place_rank: 30 }
               : /Che Fico/.test(q)    ? { lat: '37.7765', lon: '-122.4381', category: 'amenity', place_rank: 30 }
               : null,
    fix: [3, { lat: 37.7750, lng: -122.4370 }],   // precise fix times out, coarse one lands
  });
  await a.settle(400);
  a.click(a.$('nearChip'));
  await a.settle(100);

  check('nearest: a timed-out precise fix falls back to a coarse one', () => {
    if(a.state.fixOpts.length !== 2) throw new Error(a.state.fixOpts.length + ' attempts');
    if(a.state.fixOpts[1].enableHighAccuracy !== false) throw new Error('second try still high accuracy');
    if(!a.w.eval('!!state.near')) throw new Error('no position set');
  });
  const names = () => [...a.$('list').querySelectorAll('.item-name')].map(e => e.textContent.trim());
  check('nearest: only places with a point of their own are ranked', () => {
    if(names().join('|') !== 'Nopa') throw new Error(names().join('|'));
  });
  check('nearest: the rest are counted, not put on a town centre', () => {
    if(!/^3 more with no street address, not ranked/.test(a.$('geoNote').textContent)) throw new Error(a.$('geoNote').textContent);
  });
  a.click(a.$('restToggle'));
  check('nearest: "list them" shows the unplaced ones by town, and says so', () => {
    const n = names();
    if(n.includes('Nopa') || !n.includes('Che Fico') || !n.includes('Zuni Cafe')) throw new Error(n.join('|'));
    if(n.includes('Wild Crumb')) throw new Error('listed a place with nothing to place it by');
    if(!/town centre/.test(a.$('list').querySelector('.dist').textContent)) throw new Error('distance not marked as the town\'s');
    if(!/back to nearest/.test(a.$('geoNote').textContent)) throw new Error(a.$('geoNote').textContent);
  });
  check('nearest: a closed place is in neither list', () => {
    if(names().includes('Gone')) throw new Error('closed place listed');
    a.click(a.$('restToggle'));
    if(names().includes('Gone')) throw new Error('closed place ranked');
    if(names().join('|') !== 'Nopa') throw new Error('did not go back: ' + names().join('|'));
  });

  await a.settle(4000);   // the lookups are spaced a second apart
  check('refine: never calls the Census geocoder', () => {
    if(a.state.geocoded.includes('census')) throw new Error('called census');
  });
  check('refine: by address when there is one, by name + town when not', () => {
    if(!a.state.geocoded.includes('1658 Market St')) throw new Error(JSON.stringify(a.state.geocoded));
    if(!a.state.geocoded.includes('Che Fico, San Francisco, CA')) throw new Error(JSON.stringify(a.state.geocoded));
  });
  check('refine: skips exact, chains, closed and placeless rows', () => {
    const bad = a.state.geocoded.filter(q => /Divisadero|Philz|Gone|Wild Crumb/.test(q));
    if(bad.length) throw new Error(bad.join('; '));
  });
  check('refine: coordinates and precision are written back', () => {
    const z = a.state.venues.find(r => r[0] === 'Zuni Cafe');
    const c = a.state.venues.find(r => r[0] === 'Che Fico');
    if(z[8] !== 37.7736 || z[10] !== 'exact') throw new Error('Zuni ' + z.slice(8, 11));
    if(c[8] !== 37.7765 || c[10] !== 'name') throw new Error('Che Fico ' + c.slice(8, 11));
  });
  check('refine: a place that gets a point joins the ranking', () => {
    const n = names();
    if(!n.includes('Zuni Cafe') || !n.includes('Che Fico')) throw new Error(n.join('|'));
    if(!/^1 more with no street address/.test(a.$('geoNote').textContent)) throw new Error(a.$('geoNote').textContent);
  });
}
{
  const a = app({ venues: [['Zuni Cafe','American','SF','','','','','','37.77','-122.42','exact','','']], fix: 1 });
  await a.settle(400);
  a.click(a.$('nearChip')); await a.settle(100);
  check('nearest: permission denied says so, does not retry', () => {
    if(a.state.fixOpts.length !== 1) throw new Error('retried after a denial');
    if(!/turned off/.test(a.state.alerted || '')) throw new Error(a.state.alerted);
    if(a.$('nearChip').textContent !== 'nearest') throw new Error('chip stuck on ' + a.$('nearChip').textContent);
  });
}

/* ---------- nearest: where the app thinks you are ---------- */
{
  const row = (name, loc, lat, lng) => [name, 'Food', loc, '', '', '', '', '', lat, lng, 'exact', '', ''];
  const a = app({
    venues: [ row('Spoonbar', 'Healdsburg, CA', '38.6110', '-122.8700'),
              row('Zuni Cafe', 'San Francisco, CA', '37.7735', '-122.4222') ],
    geocode: q => /healdsburg/i.test(q) ? { lat: '38.6102', lon: '-122.8694', display_name: 'Healdsburg, Sonoma County, California' } : null,
    fix: { lat: 37.7880, lng: -122.4530 },     // the browser insists on home
  });
  await a.settle(400);
  a.click(a.$('nearChip')); await a.settle(100);
  check('where: the line says where the browser put you', () => {
    if(a.$('nearWhere').style.display === 'none') throw new Error('hidden');
    if(!/your browser puts you near/.test(a.$('nearLabel').textContent)) throw new Error(a.$('nearLabel').textContent);
  });
  check('where: never asks for a cached position', () => {
    if(a.state.fixOpts.some(o => o.maximumAge !== 0)) throw new Error(JSON.stringify(a.state.fixOpts));
  });
  a.click(a.$('nearFix'));
  a.$('nearInput').value = 'healdsburg';
  a.enter(a.$('nearInput')); await a.settle(200);
  check('where: a typed place replaces the browser\'s and re-sorts', () => {
    if(!/^from Healdsburg/.test(a.$('nearLabel').textContent)) throw new Error(a.$('nearLabel').textContent);
    const first = a.$('list').querySelector('.item .item-name').textContent;
    if(!/Spoonbar/.test(first)) throw new Error('first is ' + first);
    if(a.$('nearSet').style.display !== 'none') throw new Error('box left open');
  });
  a.click(a.$('nearChip')); await a.settle(50);
  check('where: turning nearest off hides the line', () => {
    if(a.$('nearWhere').style.display !== 'none') throw new Error('still shown');
  });
}


/* ---------- To Order: no strikethrough anywhere (2026.10.29-1) ---------- */
{
  const a = app({
    venues: [
      ['Zuni Cafe','American','SF','roast chicken; caesar','','','x','','','','','',''],
      ['Arsicault Bakery','Bakery','SF','Almond Croissant, Chocolate chip Cookie','','','','','','','','',''],
    ],
    menu: [['Zuni Cafe','SF','','Mains','roast chicken'], ['Q Specialty Coffee','SF','','Pastry','Almond croissant']],
  });
  await a.settle(400);
  const detail = name => { a.w.eval('openDetail(venues.find(v => v.name === ' + JSON.stringify(name) + '))'); return a.$('detailContent'); };
  const strikes = el => Array.from(el.querySelectorAll('*')).filter(x => /line-through/.test(a.w.getComputedStyle(x).textDecoration || '') || /line-through/.test(x.getAttribute('style') || '') || x.matches('s,del,strike,.hadx'));
  check('to order: an item had at another place is not struck (detail sheet)', () => {
    const c = detail('Arsicault Bakery');
    const items = Array.from(c.querySelectorAll('.wantrow span')).map(x => x.textContent);
    if(items.join('|') !== 'Almond Croissant|Chocolate chip Cookie') throw new Error(items.join('|'));
    if(strikes(c).length) throw new Error(strikes(c).length + ' struck');
  });
  check('to order: an item had at this same place is not struck either', () => {
    const c = detail('Zuni Cafe');
    if(strikes(c).length) throw new Error('struck');
    if(/already ordered/.test(c.innerHTML)) throw new Error('caption left');
  });
  a.w.eval('closeDetail()');
  a.$('searchBox').value = 'arsicault'; a.fire(a.$('searchBox'), 'input');
  check('to order: the list line shows the items plain', () => {
    const line = a.$('list').querySelector('.toorder');
    if(!/Almond Croissant, Chocolate chip Cookie/.test(line.textContent)) throw new Error(line.textContent);
    if(strikes(line).length) throw new Error('struck');
  });
  a.$('searchBox').value = 'croissant'; a.fire(a.$('searchBox'), 'input');
  check('to order: search finds a place by an item on its To Order', () => {
    if(!/Arsicault/.test(a.$('list').textContent)) throw new Error(a.$('list').textContent.slice(0, 200));
  });
  check('to order: nothing was written', () => {
    if(a.state.calls.length) throw new Error(a.state.calls.length + ' writes');
  });
}

/* ---------- data check: addresses (a category is never written) ---------- */
{
  const osm = (name, extra) => Object.assign({ name, lat: '37.7733', lon: '-122.4390', category: 'amenity', type: 'restaurant',
    place_rank: 30, address: { house_number: '306', road: 'Broderick Street', city: 'San Francisco', 'ISO3166-2-lvl4': 'US-CA', postcode: '94117' } }, extra);
  const a = app({
    venues: [
      ['Nopalito','Restaurant','San Francisco, CA','','','','','','37.77','-122.44','zip','',''],
      ['Lers Ros','Thai','San Francisco, CA','','','','','','','','','',''],
      ['Mystery Spot','Thai','San Francisco, CA','','','','','','','','','',''],
      ['Ghost Kitchen','Restaurants','San Francisco, CA','','','','','12 Main St','37.77','-122.44','exact','',''],
      ['In-N-Out','Restaurant','','','','','','','','','','x',''],
      ['Rose Street Diner','Diner','San Francisco, CA','','','','','','','','','',''],
      ['Mexican Place','Mexican','San Francisco, CA','','','','','1 A St','37.77','-122.44','exact','',''],
    ],
    geocode: q => /^Nopalito/.test(q) ? osm('Nopalito', { extratags: { cuisine: 'mexican' } })
               : /^Lers Ros/.test(q) ? osm('Lers Ros Thai')
               : /^Rose Street Diner/.test(q) ? Object.assign(osm('Rose Street'), { category: 'highway', type: 'residential' })
               : /^Ghost Kitchen/.test(q) ? osm('Some Other Place', { extratags: { cuisine: 'pizza' } })
               : null,
  });
  await a.settle(400);
  a.w.eval('openDiag()');
  check('data check: lists places with no address', () => {
    const t = a.$('diagBody').textContent;
    if(!/No address\s*4/.test(t)) throw new Error(t);
    if(/restaurant/i.test(t)) throw new Error(t);
    if(!/Look up 4 places/.test(a.$('diagLookup').textContent)) throw new Error(a.$('diagLookup').textContent);
  });
  a.click(a.$('diagLookup')); await a.settle(9000);
  const row = n => a.state.venues.find(r => r[0] === n);
  check('lookup: address and point written for a found place, category untouched', () => {
    const r = row('Nopalito');
    if(r[7] !== '306 Broderick Street, San Francisco, CA 94117') throw new Error(r[7]);
    if(r[8] !== 37.7733 || r[10] !== 'name') throw new Error(r.slice(8, 11).join(','));
    if(r[1] !== 'Restaurant') throw new Error('category ' + r[1]);
  });
  check('lookup: a found place keeps its real category', () => {
    const r = row('Lers Ros');
    if(r[1] !== 'Thai' || !r[7]) throw new Error(r.join('|'));
  });
  check('lookup: a street that shares the name is not taken', () => {
    if(row('Rose Street Diner')[7]) throw new Error('took the street');
  });
  check('lookup: a row that has an address is not looked up and its category is untouched', () => {
    const r = row('Ghost Kitchen');
    if(a.state.geocoded.some(q => /Ghost Kitchen/.test(q))) throw new Error('looked up');
    if(r[1] !== 'Restaurants' || r[7] !== '12 Main St') throw new Error(r.join('|'));
  });
  check('lookup: a chain is never looked up and its category is untouched', () => {
    if(a.state.geocoded.some(q => /In-N-Out/.test(q))) throw new Error('chain looked up');
    if(row('In-N-Out')[1] !== 'Restaurant') throw new Error(row('In-N-Out')[1]);
  });
  check('lookup: rows with an address and a real category are left alone', () => {
    if(a.state.geocoded.some(q => /Mexican Place/.test(q))) throw new Error('looked up');
  });
  check('lookup: what was not found is listed and not asked again', () => {
    const t = a.$('diagBody').textContent;
    if(!/not found online — type them in Edit: .*Mystery Spot/.test(t)) throw new Error(t);
    if(!/Nothing to look up/.test(a.$('diagLookup').textContent)) throw new Error(a.$('diagLookup').textContent);
  });
  check('form: a typed address replaces the point, so nearest looks it up', () => {
    a.w.eval('closeDiag(); openForm(venues.find(v => v.name === "Mystery Spot"))');
    a.$('f_address').value = '1 Market St, San Francisco, CA';
  });
  a.click(a.$('saveForm')); await a.settle(300);
  check('form: the address is saved', () => {
    const r = row('Mystery Spot');
    if(r[7] !== '1 Market St, San Francisco, CA' || r[8] !== '' || r[10] !== '') throw new Error(r.join('|'));
  });
}

/* ---------- opening hours (2026.11.12-1) ---------- */
{
  const a = app({ venues: [V.zuni] });
  await a.settle(300);
  const label = (hours, iso, closed = '') =>
    a.w.eval(`hoursLabel({ hours: ${JSON.stringify(hours)}, closed: ${JSON.stringify(closed)} }, new Date(${JSON.stringify(iso)}))`);
  const expect = (name, got, want) => check(name, () => { if(got !== want) throw new Error(JSON.stringify(got)); });
  // 2026-11-12 is a Thursday; Pacific is UTC-8 in November
  expect('hours: open, with the closing time', label('Mo-Fr 11:30-22:00; Sa,Su 10:00-22:00', '2026-11-12T23:00:00Z'), 'open until 10 pm');
  expect('hours: shut for the night, opens tomorrow', label('Mo-Fr 11:30-22:00; Sa,Su 10:00-22:00', '2026-11-13T07:00:00Z'), 'opens tomorrow 11:30 am');
  expect('hours: later today', label('Mo-Fr 11:30-14:00,17:00-22:00', '2026-11-12T23:00:00Z'), 'opens 5 pm');
  expect('hours: a span past midnight is still open after it', label('Mo-Su 17:00-02:00', '2026-11-13T09:30:00Z'), 'open until 2 am');
  expect('hours: a day off is skipped to the next opening', label('Tu-Sa 17:30-22:00', '2026-11-15T20:00:00Z'), 'opens Tue 5:30 pm');
  expect('hours: a later rule replaces an earlier one', label('Mo-Su 10:00-22:00; Th off', '2026-11-12T23:00:00Z'), 'opens tomorrow 10 am');
  expect('hours: 24/7', label('24/7', '2026-11-12T23:00:00Z'), 'open 24 hours');
  expect('hours: the clock is Pacific wherever the device is', label('Mo-Su 08:00-15:30', '2026-11-12T23:20:00Z'), 'open until 3:30 pm');
  expect('hours: syntax it does not understand makes no claim', label('Apr-Oct: Mo-Su 10:00-18:00', '2026-11-12T23:00:00Z'), '');
  expect('hours: nothing stored, nothing said', label('', '2026-11-12T23:00:00Z'), '');
  expect('hours: a closed place has no hours line', label('24/7', '2026-11-12T23:00:00Z', 'closed 2026'), '');
}

/* ---------- Excel asleep: no saved copy ---------- */
{
  const a = app({ venues: [V.zuni], asleep: true });
  await a.settle(400);
  check('asleep: the gate says what is wrong, not a raw Graph error', () => {
    const t = a.$('gateInner').textContent;
    if(a.$('gate').style.display === 'none') throw new Error('gate gone');
    if(!/Excel service has stopped answering/.test(t)) throw new Error(t);
    if(/FileOpenUserUnauthorized|Graph GET/.test(t)) throw new Error('raw error shown: ' + t);
  });
  check('asleep: the link opens the workbook itself', () => {
    if(a.$('wakeLink').href !== 'https://onedrive.live.com/the-workbook') throw new Error(a.$('wakeLink').href);
    if(a.$('wakeLink').target !== '_blank') throw new Error('would navigate away from the app');
  });
  a.state.asleep = false;
  Object.defineProperty(a.w.document, 'visibilityState', { get: () => 'visible', configurable: true });
  Object.defineProperty(a.w.document, 'hidden', { get: () => false, configurable: true });
  a.w.dispatchEvent(new a.w.Event('focus'));      // coming back from Excel
  await a.settle(400);
  check('asleep: coming back to the page loads the list by itself', () => {
    if(a.$('gate').style.display !== 'none') throw new Error('still gated: ' + a.$('gateInner').textContent);
    if(a.w.eval('venues.length') !== 1) throw new Error('no rows');
    if(a.$('addBtn').style.display === 'none') throw new Error('add button hidden');
  });
  await a.settle(500);
  check('snapshot: a read that worked is kept on the device', () => {
    const snap = JSON.parse(a.w.localStorage.getItem('eatdrink.snapshot') || 'null');
    if(!snap || snap.main.length !== 1 || snap.main[0][0] !== 'Zuni Cafe') throw new Error(JSON.stringify(snap));
  });
}

/* ---------- Excel asleep: the saved copy, read-only ---------- */
{
  const a = app({ venues: [V.zuni], menu: [M.zuni], asleep: true,
    snapshot: { at: Date.now() - 3600e3, main: [V.zuni, V.philzSF], menu: [M.zuni] } });
  await a.settle(400);
  check('snapshot: the list as last saved is shown when Excel will not answer', () => {
    if(a.$('gate').style.display !== 'none') throw new Error('gated: ' + a.$('gateInner').textContent);
    if(a.w.eval('venues.length') !== 2) throw new Error(a.w.eval('venues.length') + ' rows');
    if(a.w.eval('menuRows.length') !== 1) throw new Error('log not restored');
  });
  check('snapshot: the banner says it is a saved copy and how to wake Excel', () => {
    const b = a.$('staleBanner');
    if(b.style.display === 'none') throw new Error('no banner');
    if(!/showing the list as saved today .* read-only/.test(b.textContent)) throw new Error(b.textContent);
    if(!b.querySelector('#wakeLink')) throw new Error('no link');
  });
  check('snapshot: the add button is put away', () => {
    if(a.$('addBtn').style.display !== 'none') throw new Error('add button shown');
  });
  let refused = null;
  try{ await a.w.eval("graph('PATCH', '/worksheets(x)/tables(y)/rows/itemAt(index=0)', { values: [[]] })"); }
  catch(e){ refused = e; }
  check('snapshot: a write is refused before it reaches Graph', () => {
    if(!refused || !refused.userFacing) throw new Error('not refused');
    if(a.state.calls.length) throw new Error('reached Graph: ' + JSON.stringify(a.state.calls));
  });
  a.state.asleep = false;
  Object.defineProperty(a.w.document, 'visibilityState', { get: () => 'visible', configurable: true });
  Object.defineProperty(a.w.document, 'hidden', { get: () => false, configurable: true });
  a.w.dispatchEvent(new a.w.Event('focus'));
  await a.settle(400);
  check('snapshot: the live list replaces it as soon as a read works', () => {
    if(a.w.eval('venues.length') !== 1) throw new Error(a.w.eval('venues.length') + ' rows');
    if(a.w.eval('readOnly') !== null) throw new Error('still read-only');
    if(a.$('staleBanner').style.display !== 'none') throw new Error('banner left up');
    if(a.$('addBtn').style.display === 'none') throw new Error('add button still hidden');
  });
}

/* ---------- a write that hits a sleeping Excel ---------- */
{
  const a = app({ venues: [V.zuni] });
  await a.settle(300);
  a.state.asleep = true;
  a.w.eval("withBusy(() => graph('PATCH', '/worksheets(x)/tables(y)/rows/itemAt(index=0)', { values: [[]] }))");
  await a.settle(200);
  check('asleep: a failed write says to open Excel, and keeps the list', () => {
    if(!/Excel service has stopped answering/.test(a.state.alerted || '')) throw new Error(a.state.alerted);
    if(a.$('staleBanner').style.display === 'none' || !a.$('staleBanner').querySelector('#wakeLink')) throw new Error('no link offered');
    if(a.w.eval('venues.length') !== 1 || a.w.eval('readOnly') !== null) throw new Error('list changed');
  });
}

/* ---------- hours lookup (Overpass) ---------- */
{
  const row = (name, addr, lat, lng, geo, chain = '', closed = '') =>
    [name, 'Food', 'San Francisco, CA', '', '', '', '', addr, lat, lng, geo, chain, closed];
  const el = (name, lat, lon, hours, extra = {}) =>
    ({ type: 'node', lat, lon, tags: Object.assign({ name, amenity: 'restaurant' }, hours ? { opening_hours: hours } : {}, extra) });
  const a = app({
    venues: [
      row('Nopa', '560 Divisadero St', '37.7749', '-122.4374', 'exact'),
      row('Zuni Cafe', '1658 Market St', '37.7735', '-122.4222', 'exact'),
      row('Tartine', '600 Guerrero St', '37.7614', '-122.4241', 'exact'),
      row('Thai House', '1 Castro St', '37.7620', '-122.4350', 'exact'),
      row('Che Fico', '', '', '', ''),                                   // entered through the app: no point
      row('Souvla', '', '', '', ''),                                     // two of them on the map
      row('20 Spot', '', '', '', ''),                                    // "Spot" is someone else
      row('Fava', '', '', '', ''),                                       // filed under SF, stands in Daly City
      row('Gone', '9 Main St', '37.7700', '-122.4200', 'exact', '', 'closed'),
    ],
    osm: [
      el('Nopa', 37.7749, -122.4374, 'Mo-Su 17:00-23:00'),
      el('Zuni Café', 37.77352, -122.42222, 'Tu-Su 11:30-22:00'),
      el('Tartine Bakery', 37.76145, -122.42412, 'Mo-Su 08:00-17:00'),
      el('Lers Ros Thai', 37.76202, -122.43502, 'Mo-Su 11:00-22:00'),
      el('Che Fico', 37.7765, -122.4381, 'Tu-Sa 17:30-22:00', { 'addr:housenumber': '838', 'addr:street': 'Divisadero Street', 'addr:postcode': '94117' }),
      el('Souvla', 37.7765, -122.4263, 'Mo-Su 11:00-21:00'),
      el('Souvla', 37.7609, -122.4215, 'Mo-Su 11:00-21:00'),
      el('Spot', 37.7700, -122.4300, 'Mo-Su 09:00-17:00'),
      el('Fava', 37.7060, -122.4610, '', { 'addr:housenumber': '2114', 'addr:street': 'Vine Street', 'addr:city': 'Daly City', 'addr:postcode': '94014' }),
      el('Gone', 37.7700, -122.4200, 'Mo-Su 09:00-17:00'),
    ],
  });
  await a.settle(400);
  a.w.eval('openDiag()');
  a.click(a.$('diagHours'));
  await a.settle(900);
  const r = name => a.state.venues.find(x => x[0] === name);
  check('hours lookup: one question to OpenStreetMap, not one per place', () => {
    if(a.state.osmAsked !== 1) throw new Error(a.state.osmAsked + ' requests');
    if(a.state.geocoded.length) throw new Error('per-place lookups: ' + a.state.geocoded.join('; '));
  });
  check('hours lookup: the run holds one workbook session and closes it', () => {
    if(a.state.sessions !== 1 || a.state.sessionsClosed !== 1) throw new Error(a.state.sessions + ' opened, ' + a.state.sessionsClosed + ' closed');
    if(!a.state.inSession) throw new Error('writes went outside the session');
    if(a.w.eval('wbSession') !== null) throw new Error('session left set for everyday calls');
  });
  check('hours lookup: the Hours column is added once', () => {
    const cols = a.state.calls.filter(c => c.kind === 'addcol');
    if(cols.length !== 1 || cols[0].body.name !== 'Hours') throw new Error(JSON.stringify(cols));
  });
  check('hours lookup: same name at the same spot gets its hours', () => {
    if(r('Nopa')[13] !== 'Mo-Su 17:00-23:00') throw new Error(String(r('Nopa')[13]));
    if(r('Zuni Cafe')[13] !== 'Tu-Su 11:30-22:00') throw new Error('accent not matched: ' + r('Zuni Cafe')[13]);
  });
  check('hours lookup: a longer form of the name on the same doorstep counts', () => {
    if(r('Tartine')[13] !== 'Mo-Su 08:00-17:00') throw new Error(String(r('Tartine')[13]));
  });
  check('hours lookup: a neighbour that only shares a word does not', () => {
    if(r('Thai House')[13]) throw new Error(r('Thai House')[13]);
  });
  check('hours lookup: a place with no point gets point, address and hours from a name that exists once', () => {
    const c = r('Che Fico');
    if(c[8] !== 37.7765 || c[9] !== -122.4381 || c[10] !== 'name') throw new Error(c.slice(8, 11).join('|'));
    if(c[7] !== '838 Divisadero Street, San Francisco, CA 94117') throw new Error(c[7]);
    if(c[13] !== 'Tu-Sa 17:30-22:00') throw new Error(String(c[13]));
  });
  check('hours lookup: the address carries the town the map gives, not the row\'s Location', () => {
    if(r('Fava')[7] !== '2114 Vine Street, Daly City, CA 94014') throw new Error(r('Fava')[7]);
    if(r('Fava')[2] !== 'San Francisco, CA') throw new Error('Location rewritten: ' + r('Fava')[2]);
  });
  check('hours lookup: a name that exists twice places nothing', () => {
    if(r('Souvla')[8] !== '' || r('Souvla')[13]) throw new Error(r('Souvla').join('|'));
  });
  check('hours lookup: "20 Spot" is not "Spot"', () => {
    if(r('20 Spot')[8] !== '' || r('20 Spot')[13]) throw new Error(r('20 Spot').join('|'));
  });
  check('hours lookup: a closed place is left alone', () => {
    if(r('Gone')[13]) throw new Error(String(r('Gone')[13]));
  });
  check('hours lookup: the data check counts them', () => {
    if(!/Have hours\s*4/.test(a.$('diagBody').textContent)) throw new Error(a.$('diagBody').textContent);
  });
  a.w.eval("closeDiag(); state.query = 'nopa'; render()");
  check('hours: the list line carries open or opens', () => {
    const sub = a.$('list').querySelector('.item-sub').textContent;
    if(!/open until 11 pm|opens (tomorrow )?5 pm/.test(sub)) throw new Error(sub);
  });
  a.w.eval("openDetail(venues.find(v => v.name === 'Nopa'))");
  check('hours: the detail sheet shows the week', () => {
    const t = a.$('detailContent').textContent;
    if(!/Hours/.test(t) || !/Mon\s*5 pm . 11 pm/.test(t)) throw new Error(t);
  });
  a.w.eval("closeDetail(); openForm(venues.find(v => v.name === 'Nopa'))");
  a.$('f_notes').value = 'get the pork chop';
  a.click(a.$('saveForm')); await a.settle(300);
  check('form: an edit keeps the hours', () => {
    if(r('Nopa')[5] !== 'get the pork chop') throw new Error('edit not saved');
    if(r('Nopa')[13] !== 'Mo-Su 17:00-23:00') throw new Error('hours blanked: ' + r('Nopa')[13]);
  });
}
{
  // Graph keeps serving the 13-column copy after the Hours column is added
  const a = app({ stale: 20,
    venues: [['Nopa','Food','San Francisco, CA','','','','','560 Divisadero St','37.7749','-122.4374','exact','','']],
    osm: [{ type: 'node', lat: 37.7749, lon: -122.4374, tags: { name: 'Nopa', amenity: 'restaurant', opening_hours: 'Mo-Su 17:00-23:00' } }] });
  await a.settle(300);
  a.w.eval('openDiag()'); a.click(a.$('diagHours')); await a.settle(900);
  check('hours lookup: a stale read after the column is added does not shrink the rows written', () => {
    const patches = a.state.calls.filter(c => c.method === 'PATCH');
    if(!patches.length) throw new Error('nothing written: ' + a.state.alerted);
    const bad = patches.filter(c => c.body.values[0].length !== 14);
    if(bad.length) throw new Error('row written at width ' + bad[0].body.values[0].length);
    if(patches[0].body.values[0][13] !== 'Mo-Su 17:00-23:00') throw new Error(String(patches[0].body.values[0][13]));
  });
}
{
  const a = app({ venues: [['Nopa','Food','San Francisco, CA','','','','','560 Divisadero St','37.7749','-122.4374','exact','','']], osm: null });
  await a.settle(300);
  a.w.eval('openDiag()'); a.click(a.$('diagHours')); await a.settle(400);
  check('hours lookup: no answer from OpenStreetMap changes nothing and says so', () => {
    if(a.state.calls.length) throw new Error(JSON.stringify(a.state.calls));
    if(!/OpenStreetMap did not answer/.test(a.state.alerted || '')) throw new Error(a.state.alerted);
    if(a.$('diagHours').textContent !== 'Look up hours') throw new Error('button stuck: ' + a.$('diagHours').textContent);
  });
}

/* ---------- Google Maps (2026.11.12-5) ---------- */
{
  const a = app({ venues: [V.zuni] });
  await a.settle(300);
  const week = periods => a.w.eval(`JSON.stringify(googleWeek(${JSON.stringify({ periods })}))`);
  const P = (od, oh, om, cd, ch, cm) => ({ open: { day: od, hour: oh, minute: om }, close: { day: cd, hour: ch, minute: cm } });
  check('google hours: a weekday span lands on its day (0 is Sunday)', () => {
    const w = JSON.parse(week([P(1, 11, 30, 1, 22, 0)]));
    if(JSON.stringify(w[1]) !== '[[690,1320]]' || w[0].length || w[2].length) throw new Error(JSON.stringify(w));
  });
  check('google hours: a span that closes the next day runs past midnight', () => {
    const w = JSON.parse(week([P(5, 17, 0, 6, 2, 0)]));
    if(JSON.stringify(w[5]) !== '[[1020,1560]]') throw new Error(JSON.stringify(w));
  });
  check('google hours: a place that never shuts is open every day', () => {
    const w = JSON.parse(week([{ open: { day: 0, hour: 0, minute: 0 } }]));
    if(!w.every(d => JSON.stringify(d) === '[[0,1440]]')) throw new Error(JSON.stringify(w));
  });
  check('google hours: no periods, no claim', () => {
    if(week([]) !== 'null') throw new Error(week([]));
  });
  a.w.eval('openDiag()');
  check('google: without a key the check is off and says why', () => {
    if(!a.$('diagGoogle').disabled || !/needs a key/.test(a.$('diagGoogle').textContent)) throw new Error(a.$('diagGoogle').textContent);
    if(a.$('diagGoogleKey').textContent !== 'Set Google key') throw new Error(a.$('diagGoogleKey').textContent);
  });
  a.w.eval("closeDiag(); openDetail(venues[0])"); await a.settle(100);
  check('google: without a key nothing is asked', () => {
    if(a.state.google.length) throw new Error(a.state.google.length + ' requests');
  });
}
{
  const row = (name, addr, lat, lng, geo, closed = '', hours = '') =>
    [name, 'Food', 'San Francisco, CA', '', '', '', '', addr, lat, lng, geo, '', closed, hours];
  const g = (id, name, lat, lng, status = 'OPERATIONAL') =>
    ({ id, displayName: { text: name }, location: { latitude: lat, longitude: lng }, businessStatus: status });
  const a = app({ googleKey: 'key-on-this-device',
    venues: [
      row('Nopa', '560 Divisadero St', '37.7749', '-122.4374', 'exact', '', 'Mo-Su 17:00-23:00'),
      row('Bellota', '888 Brannan St', '37.7716', '-122.4052', 'exact'),
      row('Thai House', '1 Castro St', '37.7620', '-122.4350', 'exact'),
      row('Far Away', '1 Main St', '37.7900', '-122.3930', 'exact'),
      row('Che Fico', '', '', '', ''),                                   // no point: found by its town
      row('Moved Cafe', '5 Hayes St', '37.7770', '-122.4200', 'exact'),
      row('Gone', '9 Main St', '37.7700', '-122.4200', 'exact', 'closed'),
    ],
    google: (kind, arg) => {
      if(kind !== 'search') return 400;
      const q = arg.textQuery;
      if(/^Nopa/.test(q)) return { places: [g('g-nopa', 'Nopa', 37.7749, -122.4374)] };
      if(/^Bellota/.test(q)) return { places: [g('g-bellota', 'Bellota', 37.7716, -122.4052, 'CLOSED_PERMANENTLY')] };
      if(/^Thai House/.test(q)) return { places: [g('g-lers', 'Lers Ros Thai', 37.7620, -122.4350)] };
      if(/^Far Away/.test(q)) return { places: [g('g-far', 'Far Away', 37.8300, -122.3930)] };
      if(/^Che Fico/.test(q)) return { places: [g('g-chefico', 'Che Fico', 37.7765, -122.4381)] };
      if(/^Moved Cafe/.test(q)) return { places: [g('g-old', 'Moved Cafe', 37.7770, -122.4200, 'CLOSED_PERMANENTLY'), g('g-new', 'Moved Cafe', 37.7772, -122.4203)] };
      return { places: [] };
    } });
  await a.settle(400);
  a.w.eval('openDiag()');
  check('google: with a key the button says how many', () => {
    if(a.$('diagGoogle').disabled || a.$('diagGoogle').textContent !== 'Check 6 with Google') throw new Error(a.$('diagGoogle').textContent);
  });
  a.click(a.$('diagGoogle')); await a.settle(900);
  const r = name => a.state.venues.find(x => x[0] === name);
  check('google: one search per open place, with this device\'s key', () => {
    if(a.state.google.length !== 6) throw new Error(a.state.google.length + ' requests');
    if(a.state.google.some(q => q.key !== 'key-on-this-device')) throw new Error('wrong key sent');
    if(a.state.google.some(q => /^Gone/.test(q.arg.textQuery))) throw new Error('asked about a closed place');
  });
  check('google: the check asks for nothing beyond the Pro fields (hours are never fetched ahead)', () => {
    const masks = [...new Set(a.state.google.map(q => q.mask))];
    if(masks.join('|') !== 'places.id,places.displayName,places.location,places.businessStatus') throw new Error(masks.join('|'));
  });
  check('google: Place ID and Checked are added as the 15th and 16th columns', () => {
    const cols = a.state.calls.filter(c => c.kind === 'addcol').map(c => c.body.name);
    if(cols.join('|') !== 'Place ID|Checked') throw new Error(cols.join('|'));
  });
  check('google: a found place keeps Google\'s ID and the day it was asked', () => {
    if(r('Nopa')[14] !== 'g-nopa') throw new Error(String(r('Nopa')[14]));
    if(!/^checked \d{4}-\d{2}-\d{2}$/.test(r('Nopa')[15])) throw new Error(String(r('Nopa')[15]));
  });
  check('google: nothing else of Google\'s is written - no address, point or hours', () => {
    const c = r('Che Fico');
    if(c[14] !== 'g-chefico') throw new Error('not found: ' + c[14]);
    if(c[7] !== '' || c[8] !== '' || c[9] !== '' || c[10] !== '' || c[13] !== '') throw new Error(c.slice(7, 14).join('|'));
    if(r('Nopa')[13] !== 'Mo-Su 17:00-23:00') throw new Error('stored hours touched');
  });
  check('google: a neighbour of another name, or the same name miles off, is not a match', () => {
    if(r('Thai House')[14] !== '' || r('Far Away')[14] !== '') throw new Error(r('Thai House')[14] + ' / ' + r('Far Away')[14]);
    if(!/^checked /.test(r('Thai House')[15])) throw new Error('not stamped, would be asked again first');
  });
  check('google: a place Google lists as permanently closed is marked, after being shown', () => {
    if(!/^closed \(Google Maps, \w+ \d{4}\)$/.test(r('Bellota')[12])) throw new Error(String(r('Bellota')[12]));
  });
  check('google: a place that moved - old listing closed, new one open - stays open', () => {
    if(r('Moved Cafe')[12] !== '' || r('Moved Cafe')[14] !== 'g-new') throw new Error(r('Moved Cafe')[12] + ' / ' + r('Moved Cafe')[14]);
  });
  check('google: the run says what it did', () => {
    if(!/^4 found on Google Maps, 2 not matched, 1 marked closed\.$/.test(a.state.alerted || '')) throw new Error(a.state.alerted);
  });
  check('google: the key is not written anywhere but this device', () => {
    if(JSON.stringify(a.state.venues).includes('key-on-this-device') || JSON.stringify(a.state.calls).includes('key-on-this-device')) throw new Error('key in the workbook');
    if((a.w.localStorage.getItem('eatdrink.snapshot') || '').includes('key-on-this-device')) throw new Error('key in the saved copy');
  });
  a.w.eval("closeDiag(); openForm(venues.find(v => v.name === 'Nopa'))");
  a.$('f_notes').value = 'pork chop';
  a.click(a.$('saveForm')); await a.settle(300);
  check('form: an edit keeps the Google ID', () => {
    if(r('Nopa')[14] !== 'g-nopa' || !/^checked /.test(r('Nopa')[15])) throw new Error(r('Nopa').slice(14).join('|'));
  });
  a.w.eval("openForm(venues.find(v => v.name === 'Nopa'))");
  a.$('f_name').value = 'Nopalito';
  a.click(a.$('saveForm')); await a.settle(300);
  check('form: a renamed place is asked about again', () => {
    if(r('Nopalito')[14] !== '' || r('Nopalito')[15] !== '') throw new Error(r('Nopalito').slice(14).join('|'));
  });
}
{
  // the daily cap on the key runs out part-way
  const row = (name, lat) => [name, 'Food', 'San Francisco, CA', '', '', '', '', '1 St', lat, '-122.4374', 'exact', '', '', '', '', ''];
  const a = app({ googleKey: 'k', venues: [row('A One', '37.7700'), row('B Two', '37.7710'), row('C Three', '37.7720')],
    google: (kind, arg, n) => n >= 3 ? 429 : { places: [{ id: 'g' + n, displayName: { text: arg.textQuery.split(',')[0] },
      location: { latitude: n === 1 ? 37.7700 : 37.7710, longitude: -122.4374 }, businessStatus: 'OPERATIONAL' }] } });
  await a.settle(300);
  a.w.eval('openDiag()'); a.click(a.$('diagGoogle')); await a.settle(700);
  check('google: when the day\'s allowance runs out the run stops, keeps what it has, and says so', () => {
    const got = a.state.venues.filter(x => x[14]).length;
    if(got !== 2) throw new Error(got + ' written');
    if(!/allowance for today is used up/.test(a.state.alerted || '')) throw new Error(a.state.alerted);
    if(a.state.google.length !== 3) throw new Error('kept asking: ' + a.state.google.length);
  });
  a.w.eval("closeDiag()"); a.w.eval('openDiag()');
  check('google: the place not reached is first in line next time', () => {
    if(a.w.eval('googleRows()[0].name') !== 'C Three') throw new Error(a.w.eval('googleRows()[0].name'));
  });
}
{
  // hours read live on the detail sheet
  const P = (d, oh, om, ch, cm) => ({ open: { day: d, hour: oh, minute: om }, close: { day: d, hour: ch, minute: cm } });
  let status = 'OPERATIONAL', fail = 0;
  const a = app({ googleKey: 'k',
    venues: [['Nopa','Food','San Francisco, CA','','','','','560 Divisadero St','37.7749','-122.4374','exact','','','Mo-Su 09:00-10:00','g-nopa','checked 2026-11-12'],
             ['Zuni Cafe','Food','San Francisco, CA','','','','','1658 Market St','37.7735','-122.4222','exact','','','Tu-Su 11:30-22:00','','']],
    google: (kind, id) => fail || (kind === 'details' && id === 'g-nopa'
      ? { businessStatus: status, regularOpeningHours: { periods: [0,1,2,3,4,5,6].map(d => P(d, 17, 30, 23, 0)) } } : 400) });
  await a.settle(300);
  a.w.eval("openDetail(venues[0])"); await a.settle(150);
  check('google: opening a place reads its hours there and then', () => {
    const t = a.$('d_hours').textContent;
    if(!/Google Maps/.test(t) || !/Mon\s*5:30 pm . 11 pm/.test(t)) throw new Error(t);
    if(/9 am/.test(t)) throw new Error('stored hours still shown: ' + t);
  });
  check('google: that is one Place Details request for hours and status only', () => {
    if(a.state.google.length !== 1 || a.state.google[0].kind !== 'details' || a.state.google[0].arg !== 'g-nopa') throw new Error(JSON.stringify(a.state.google));
    if(a.state.google[0].mask !== 'regularOpeningHours,businessStatus') throw new Error(a.state.google[0].mask);
  });
  a.w.eval("closeDetail(); openDetail(venues[0])"); await a.settle(150);
  check('google: opening it again in the same visit does not ask twice', () => {
    if(a.state.google.length !== 1) throw new Error(a.state.google.length + ' requests');
    if(!/Google Maps/.test(a.$('d_hours').textContent)) throw new Error('live hours lost');
  });
  check('google: what it said is kept nowhere - not the workbook, not the device', () => {
    if(a.state.calls.length) throw new Error('wrote: ' + JSON.stringify(a.state.calls));
    const stored = Object.keys(a.w.localStorage).map(k => a.w.localStorage.getItem(k)).join(' ');
    if(/17:30|"periods"|1050/.test(stored)) throw new Error('hours on the device');
  });
  a.w.eval("closeDetail(); openDetail(venues[1])"); await a.settle(150);
  check('google: a place not found there shows the stored hours and asks nothing', () => {
    if(a.state.google.length !== 1) throw new Error('asked');
    if(!/OpenStreetMap/.test(a.$('d_hours').textContent)) throw new Error(a.$('d_hours').textContent);
  });
  a.w.eval("closeDetail(); liveCache.clear()"); fail = 429;
  a.w.eval("openDetail(venues[0])"); await a.settle(150);
  check('google: out of allowance, the stored hours stay on the sheet', () => {
    const t = a.$('d_hours').textContent;
    if(!/OpenStreetMap/.test(t) || !/9 am/.test(t)) throw new Error(t);
    if(a.state.alerted) throw new Error('nagged: ' + a.state.alerted);
  });
  a.w.eval("closeDetail(); liveCache.clear()"); fail = 0; status = 'CLOSED_PERMANENTLY';
  a.w.eval("openDetail(venues[0])"); await a.settle(150);
  check('google: a place it lists as permanently closed says so on the sheet', () => {
    if(!/Google Maps lists this place as permanently closed/.test(a.$('d_hours').textContent)) throw new Error(a.$('d_hours').textContent);
    if(a.state.venues[0][12] !== '') throw new Error('marked without being asked');
  });
}

console.log(results.join('\n'));
console.log(failures ? `\n${failures} failing` : `\nall ${results.length} checks passed`);
process.exit(failures ? 1 : 0);
