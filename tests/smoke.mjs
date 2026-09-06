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

function app({ venues = [], menu = [] } = {}){
  const state = { venues: venues.map(r => r.slice()), menu: menu.map(r => r.slice()), calls: [] };
  const rows = v => ({ value: v.map((r, i) => ({ index: i, values: [r] })) });
  const dom = new JSDOM(html, {
    runScripts: 'dangerously',
    url: 'https://nameisredacted.github.io/food-drink-app/',
    beforeParse(w){
      w.msal = { PublicClientApplication: class {
        async initialize(){} async handleRedirectPromise(){ return null; }
        getAllAccounts(){ return [{ username: 'test' }]; }
        async acquireTokenSilent(){ return { accessToken: 'token' }; }
      } };
      w.fetch = async (url, opts = {}) => {
        const u = String(url), m = (opts.method || 'GET').toUpperCase();
        if(u.includes('version.txt')) return { ok: true, status: 200, async text(){ return VERSION; } };
        const isMenu = u.includes('MenuDetails');
        if(m !== 'GET'){
          const idx = (/itemAt\(index=(\d+)\)/.exec(u) || [])[1];
          const kind = u.includes('/columns') ? 'addcol' : (isMenu ? 'menu' : 'venue');
          state.calls.push({ method: m, kind, body: opts.body && JSON.parse(opts.body), index: idx && +idx });
          const table = isMenu ? state.menu : state.venues;
          if(kind === 'addcol') state.venues = state.venues.map(r => r.concat(['']));
          else if(m === 'PATCH') table[+idx] = JSON.parse(opts.body).values[0];
          else if(m === 'DELETE') table.splice(+idx, 1);
          else if(m === 'POST') table.push(JSON.parse(opts.body).values[0]);
          return { ok: true, status: 201, async json(){ return {}; }, async text(){ return ''; } };
        }
        const body = isMenu ? rows(state.menu) : rows(state.venues);
        return { ok: true, status: 200, async json(){ return body; }, async text(){ return JSON.stringify(body); } };
      };
      w.alert = () => {}; w.confirm = () => true; w.prompt = () => 'Other';
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

/* ---------- marking the known closures from the data check ---------- */
{
  const eleven = r => r.slice(0, 11);
  const a = app({ venues: [
    eleven(['Bellota','Spanish','San Francisco, CA','','y','','','','','','']),
    eleven(['The Wurst','German','Healdsburg, CA','','y','','','','','','']),
    eleven(['Zuni Cafe','American','San Francisco, CA','','y','','','','','','']),
  ], menu: [] });
  await a.settle(400);
  a.w.eval('openDiag()');
  check('the data check counts the known closures', () => {
    if(!/Mark 2 closures/.test(a.$('diagMarkClosed').textContent)) throw new Error(a.$('diagMarkClosed').textContent);
  });
  a.click(a.$('diagMarkClosed')); await a.settle(500);
  check('both columns get added and only the closed places are marked', () => {
    const adds = a.state.calls.filter(c => c.kind === 'addcol').map(c => c.body.name);
    if(adds.join(',') !== 'Chain,Closed') throw new Error('columns added: ' + adds.join(','));
    const bell = a.state.venues.find(r => r[0] === 'Bellota');
    const wurst = a.state.venues.find(r => r[0] === 'The Wurst');
    const zuni = a.state.venues.find(r => r[0] === 'Zuni Cafe');
    if(bell[12] !== 'closed 2026') throw new Error('Bellota: ' + JSON.stringify(bell));
    if(!wurst[12]) throw new Error('The Wurst not marked');
    if(zuni[12]) throw new Error('an open place was marked closed');
    if(a.state.venues.length !== 3) throw new Error('a row was removed');
  });
  check('nothing is left to mark', () => {
    if(!a.$('diagMarkClosed').disabled) throw new Error(a.$('diagMarkClosed').textContent);
  });
}

console.log(results.join('\n'));
console.log(failures ? `\n${failures} failing` : `\nall ${results.length} checks passed`);
process.exit(failures ? 1 : 0);
