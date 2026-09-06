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

console.log(results.join('\n'));
console.log(failures ? `\n${failures} failing` : `\nall ${results.length} checks passed`);
process.exit(failures ? 1 : 0);
