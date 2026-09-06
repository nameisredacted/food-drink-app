# eat + drink — working notes

Single-file web app for a personal food/drink list. `index.html` is the whole app;
`version.txt` is what the running copy polls to know a new build exists.

## Where things live

| Thing | Path |
|---|---|
| App source | `index.html` (one file: markup, CSS, JS) |
| Build stamp | `version.txt`, and `APP_VERSION` near the top of the script |
| Repo on the Mac | `~/Desktop/Invisible Hand/Apps/eat + drink/food-drink-app` |
| Live site | https://nameisredacted.github.io/food-drink-app/ (GitHub Pages, `main`) |
| Workbook | OneDrive `the list/Food + Drink/Claude/Food + Drink.xlsx` |
| Old workbooks | `…/Food + Drink/Claude/archive/` (per-region, `to eat.xlsx`, old dashboard) |
| Tests | `tests/smoke.mjs` (jsdom, no network, no OneDrive) |

## Data model

`Food + Drink.xlsx`, table **FoodDrinkTable** on sheet **Food + Drink**:

`Name | Category | Location | To Order | Rating | Notes | Has Menu Detail | Address | Lat | Lng | Geo Precision | Chain`

- `Chain` is the 12th column and may be **absent** in an older copy of the workbook.
  `mainWidth` is read from the first row on load; `objToRow()` pads/truncates to it, so
  an 11-column sheet keeps working. `ensureChainColumn()` adds the column via Graph
  the first time a chain is marked.
- Table **MenuDetailsTable** on sheet **Menu Details**: `Venue | Location | Vendor | Section | Item`.
  It is a flat log of what has been ordered.

**Venue identity is name + location** (`vKey`, `venueKey`, `normTxt`). Names repeat —
28 of them, 11 within the same location — so nothing may key on name alone. `resolveVenue()`
returns null rather than guessing.

Auth: MSAL (`MSAL_CLIENT_ID`), Graph Excel table API, redirect URI is the Pages URL.

## Conventions

- Bump `APP_VERSION` **and** `version.txt` together in every change; the app compares them
  and offers a reload.
- Commit message: `2026.09.06-N: what changed`.
- **Pushing**: Terminal on this Mac can only be granted click-only access, and the cloud
  sandbox's git proxy will not issue credentials for this repo, so pushes go through
  **GitHub Desktop** (Repository ▸ Fetch, type the summary, Commit, Push). `.DS_Store` is
  tracked but should be left unchecked in commits.
- Editing from a Cowork session: stage `index.html`, edit, commit it back to the same path
  with `expectedMtimeMs`, then push from GitHub Desktop.

## UI decisions worth keeping

- The **ordered** sheet lists nothing until something is typed; `#logSheet` has a
  `min-height` so an empty list does not collapse the sheet into a strip.
- Placeholders are lowercase (`search…`, `where`, `what`) with `autocapitalize="none"`.
- Suggestions on `where` / `what` / `location` / `cuisine` are `typedOnly` — nothing on focus.
- One copy line only, beside the `+`: `nothing called X yet — add it?` /
  `X is already on the list in SF, Palo Alto` / (with a location typed)
  `X in SF is already on the list — this just adds what is new`. No quotes around names.
  It refreshes deferred (`setTimeout(…, 0)`) because the autocomplete completes the word
  in its own `input` handler.
- The `+` is an inline-block SVG (crossbar at the box centre), `vertical-align: middle`,
  margins zeroed — `.bulklink` brings a `margin-top` that otherwise drops it off the line.
- Venue ordered page: add fields stay behind `+ add items`.
- Sheets are dismissed with the `×`; panels have no cancel buttons.

## Data check (repairs, each one-tap with a confirm)

`findDataIssues()` + the buttons in `#diagSheet`:

- **Repair menu flags** — `Has Menu Detail` that disagrees with the log.
- **Mark chains** — every row whose name exists in more than one location.
- **Merge duplicates** — same name+location twice: keeps the lowest row, first non-empty
  field wins, `To Order` is the union with a shorter entry absorbed by a longer one.
- **Fix chain rows** — rows filed under a made-up location (`Multiple Locations`,
  `Local Chain`, `National Chain`, `Chain`, `International Chain`): sets `Chain = x`,
  clears the location, and moves the logged items that pointed at the fake label.
- **Attach logged items** — log rows whose name matches exactly one venue but whose
  Location text is stale (mostly `North Bay, CA`).

## Open items (as of the 2026-09-04 workbook snapshot)

- ~13 log names are spelling drift and are **not** auto-attached: Black Oak Coffee
  Roastery→Roasters (21 rows), Healdsburg Bagel Co. + Drewish→& Drewish Deli (12),
  Underdog→Underdogs Tres, Dry Creek General Store bar, Guiso's→Guiso, Taco el charro,
  Daeho kalbijjin, Thourough Bread, troubador, Brecks, Equator Coffee(s),
  Dumpling House Mongolian Cuisin(e), Everett & Jones BBQ.
- ~120 logged rows name places that are not on the list at all: Laureate (26),
  flying goat (18), `as quoted` (13, looks like a placeholder), plank coffee + roastery,
  Roof 106, SFCFC, Lo + Behold, Espressiosos, Levi's, SFO.
- 54 rows flagged `Has Menu Detail` with nothing logged anywhere.
- 48 rows with no location — fine by design; that group should only ever hold
  single-location places or chains.

## Testing

`node tests/smoke.mjs` (needs `npm i jsdom`). It loads `index.html` in jsdom with MSAL and
Graph stubbed, drives the real handlers, and asserts on the rows the app would have written.
No network, no OneDrive, nothing touches the live workbook.
