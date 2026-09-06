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

## Bugs found in review (fixed 2026.09.06-20) — worth not reintroducing

- `saveForm` called `closeForm()` (which nulls `editingVenue`) **before** the save branch
  read it, so every edit fell through to the add branch and wrote a **second row**. This is
  the most likely source of the duplicate rows in the workbook. Capture the venue first.
- The same handler rebuilt the record literally and omitted `chain`, blanking the flag on
  save. Any literal record must list every column in `COLS`.
- The global log opened a venue by name, which is the wrong branch for a chain. It now
  carries the log row's index and opens the branch that row resolved to.
- `markChains` called `resolveVenue()` (a full `loadAll`) per row. PATCH does not shift row
  indexes, so one read per pass is enough — the 43-row backfill was 43 full table reads.

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

## Closure sweep (started 2026-09-06)

Method: cross-match published closure round-ups against the workbook names first (cheap,
broad), then verify individual venues. Confirmed closed so far — these rows should be
retired or marked:

| Row | Location | Rating | Evidence |
|---|---|---|---|
| Bellota | San Francisco | y | SF Chronicle; Yelp shows CLOSED; Absinthe Group took the space |
| The Wurst | Healdsburg | y | Healdsburg Tribune / Sonoma Magazine: closed, replacement named |
| Funky Elephant | Berkeley | - | Infatuation closings list, June 2026 |
| Bar Brucato | San Francisco | - | Infatuation + Eddie's List, June 2026 |
| Decant | San Francisco | - | Infatuation (DecantSF bottleshop, June 2026) |
| Super Mensch | San Francisco | - | Infatuation, June 2026 |
| Ama | San Francisco | - | Infatuation, May 2026 (Transamerica Pyramid) |
| Café Sebastian | San Francisco | - | Infatuation, May 2026 (Transamerica Pyramid) |
| Madlab | San Francisco | - | Infatuation, May 2026 (Transamerica Pyramid) |
| Oken | Oakland | - | Infatuation, May 2026 (Rockridge) |
| Gold Palm | Oakland | - | Infatuation, May 2026 |
| International Smoke | San Francisco | - | Infatuation, May 2026 |
| Casa Borinqueña | San Francisco | - | Eddie's List, Jan 2026 (Saluhall) |
| Tiger's Taproom | Oakland | - | Eddie's List, Mar 2026 |
| Del Popolo | San Francisco | - | Eddie's List, May 2026 — verify, name spelled "Del Poppolo" there |
| Hamburger Project | San Francisco | - | Eddie's List, May 2026 (Mission) |
| Noodle in a haystack | San Francisco | - | Eddie's List, June 2026 |

Verified still open: Spoonbar, Willi's Seafood, Elephant in the Room (Healdsburg);
Sociale, Pacific Cocktail Haven, Yank Sing (101 Spear), Osha Thai (4 Embarcadero Ctr),
PPQ Dungeness Island (5821 Geary — the Balboa St branch is the closed one) (SF);
Hang Ah, Bollywood Kitchen, Anna's Seafood, Arandas, El Coyote, Tokyo Central,
Centurión Lounge SEA.

Note: Amy's Drive Thru matched a closure list, but only the Rohnert Park branch closed and
the row is a Local Chain — do not retire it. No Sonoma County closure from the 2025 and
2026 round-ups matches any other row.

Still to sweep: the remaining rated-y rows, then Healdsburg, then the rest of SF.

## Missing locations, verified 2026-09-06

Nine of the 48 no-location rows have logged items; all nine are open and their locations are:
Hang Ah Dim Sum → San Francisco, CA (1 Pagoda Pl); Bollywood Kitchen → Healdsburg, CA;
Anna's Seafood → Petaluma, CA; Arandas → Healdsburg, CA; El Coyote → Sonoma, CA;
Tokyo Central → Emeryville, CA; Centurión Lounge SEA → Seattle, WA.
Raimondo Park and Mezclá could not be identified — ask before filling them in.

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

## Working on this across sessions

Chat threads are disposable; this file is not. The routine:

1. **Knowledge lives here, not in the thread.** Anything worth surviving — a data-model
   fact, a convention, a UI decision, an open item — gets written into this file in the
   same commit as the change that prompted it.
2. **Start each work block in a fresh Cowork session**, linked to the Mac, with both
   folders connected: this repo, and the OneDrive `the list/Food + Drink/Claude` folder.
   Opening line: *"read CLAUDE.md in food-drink-app, then …"*. That is the whole handoff.
3. **Run the tests before and after any change**: `node tests/smoke.mjs` (`npm i jsdom`
   once). They stub MSAL and Graph, so the live workbook is never touched.
4. **Verify every write.** A `device_commit_files` call has reported success while the
   Mac still held the old bytes. After writing a file to the Mac, stage it back and compare
   (`md5sum`) before committing it in git; only then push.
5. **Push**: GitHub Desktop, per the convention above. The cloud sandbox's git proxy
   refuses credentials for this repo unless `nameisredacted/food-drink-app` is added to
   the session's authorized sources — do that and pushes can happen straight from the
   session instead.

## Testing

`node tests/smoke.mjs` (needs `npm i jsdom`). It loads `index.html` in jsdom with MSAL and
Graph stubbed, drives the real handlers, and asserts on the rows the app would have written.
No network, no OneDrive, nothing touches the live workbook.
