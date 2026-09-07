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

The workbook now carries all 13 columns — `Chain` (12th) and `Closed` (13th) were written
into it directly on 2026-09-07, so neither is added lazily any more.

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
- Editing from a Cowork session: with the folder connected, edit `index.html` in place
  through `device_bash` (no staging round-trip, so none of the cached-bytes trouble below),
  then push from GitHub Desktop.
- **Pushing from a Cowork session now works end to end** (2026-09-07): ask for delete
  permission on the repo folder (`device_request_delete_permission`) so git can clear its own
  locks, `git commit` in the mounted repo, then drive **GitHub Desktop** with computer use —
  `computer_resolve_access` / `computer_request_access` for `com.github.GitHubClient`, then
  Fetch origin and press the toolbar **Push origin** button. The in-page Push button does not
  always take a background press; the toolbar one does. The cloud sandbox's git proxy still
  refuses credentials for this repo (`not in this session's authorized repository set`), so
  the sandbox cannot push directly.
- **Git from `device_bash` leaves stale locks.** The shell cannot unlink, so every git
  command leaves `.git/index.lock`, `.git/HEAD.lock` and `tmp_obj_*` files behind, and a
  stale `index.lock` blocks GitHub Desktop. Ask for delete permission on the repo folder
  once per session (`device_request_delete_permission`), and clear the locks after the
  last git command.

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
- The **closed label is not a list decoration** — it renders wherever a place's name is
  shown (list, detail heading, ordered-page title, global-log heading) through `closedTag(v)`.
- The **primary search bar has no dropdown** (removed 2026.09.07-1). The list below it
  already filters live and shows the whole result set; the box only ever showed a top-8
  slice of the same thing. Enter dismisses the keyboard, Escape clears the box. The
  `typedOnly` autocompletes on the form and `+` panel fields are unrelated and stay.

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

## Duplicate rows — prevention and the 2026-09-07 cleanup

The workbook held 19 duplicate rows. All were removed on 2026-09-07; backup at
`archive/Food + Drink (pre-dedupe 20260907-021050).xlsx`. Two kinds:

- **11 same name + same location** (Cinderella, Dumpling Home, Ensarro, Fiestabowls,
  Frank Grizzly's, Grand Opening, Sofiya, Studio Estepan, Taqueria Los Mayas, Z&Y,
  Zona Rosa) — the pre-`-20` edit bug. The newer row sat at the end of the table and
  differed only by a refined `To Order` ("agua fresca" -> "agua fresca mango").
- **8 same name + identical street address, one filed under `Multiple Locations`**
  (Ariscault, Chuy's Fiestas, Cinderella, Ensarro, Fentons, Outta Sight Pizza, Tartine,
  Yonsei) — the fake-location problem the data check's **Fix chain rows** button targets.

Merge rule used, same as the app's: lowest row kept, first non-empty field wins,
`To Order` unioned with a shorter entry absorbed by a longer one.

Guards in `saveForm` (2026.09.07-1) so neither kind can come back:

- the **add** branch re-reads and looks for the name+location before writing. If it is
  already there it offers to update that row (typed values win, blanks keep what the row
  had) instead of appending a twin; a closed row is refused outright.
- the **edit** branch refuses a rename that would land on another row's name+location.

Note a duplicate pair used to be self-perpetuating: two field-identical rows make
`matchVenue()` return null, `resolveVenue()` throws the ambiguity error, the edit fails,
and the obvious next move is to add the place again.

Four same-name groups were **left alone** — no matching address to prove they are one
venue, so each is either one chain row or per-branch rows and needs a call:
Andytown Coffee Roasters (Local Chain + SF), Boudin Bakery (Millbrae + SF + Multiple
Locations), Matcha Cafe Maiko (SF + Multiple Locations), Philz Coffee (Local Chain + SF).

The live workbook is now **13 columns**: `Chain` and `Closed` were written into it on
2026-09-07 rather than waiting for the app to add them lazily. `Chain` is still empty —
nothing has been marked as a chain yet.

## Closed places (2026.09.06-21)

`Closed` is the 13th column. Any non-empty value means closed ("closed Jun 2026" reads
better than an x, and shows on the detail sheet). Closed rows are **never deleted**:

- the list keeps them, name in `--bad` red with a CLOSED tag, sub-line muted red;
- the detail sheet shows everything but offers no rating, no To Order, no Edit, no Delete —
  only the ordered history and a `reopen this place` link;
- the venue's ordered page keeps its items but hides `+ add items` and the row removers;
- logging against a closed place is refused wherever it is attempted (ordered screen,
  location picker, the + panel), and the dice never picks one.

`mark as closed` / `reopen this place` on the detail sheet is the only mutation a closed
row accepts, and it is now the **only** way a row gets marked: the bulk **Mark closures**
button, `KNOWN_CLOSURES` and `knownClosureRows()` were removed in 2026.09.07-2. The sixteen
verified closures were written straight into the workbook instead, so the list they held is
in the data, not in the code.

The closed label follows the place **everywhere its name is shown** (2026.09.07-2) — the
list, the detail heading, the title of its ordered page, and the venue heading in the global
log — via one `closedTag(v)` helper. A record should never read as a live entry, whichever
screen it turns up on.

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
broad), then verify individual venues. All sixteen below are now **written into the
workbook's `Closed` column** (2026-09-07), so they carry their label in the app; the bulk
button that used to apply them is gone. Closed rows are never deleted:

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

## Missing locations — verified 2026-09-06, written in 2026-09-07

All applied directly to the workbook (not through the app), with the matching log rows
relocated so venue identity resolves:

| Row | Location written |
|---|---|
| Hang Ah Dim Sum | **Santa Rosa, CA** |
| Bollywood Kitchen | Healdsburg, CA |
| Arandas | Healdsburg, CA |
| Anna's Seafood | Petaluma, CA |
| El Coyote | Sonoma, CA |
| Tokyo Central | Emeryville, CA |
| Centurión Lounge SEA | Seattle, WA |
| Raimondo Park | Oakland, CA (1800 Wood St) |

**Hang Ah** is two different places and the list holds both: `Hang Ah Tea Room` is the SF
Chinatown one at 1 Pagoda Pl, and `Hang Ah Dim Sum` is **Santa Rosa** — corrected on
2026-09-07 after it was first filed under the Tea Room's SF address. Its 25 logged rows
moved with it. Do not merge the two.

**Raimondo Park** was the other open question: the log's vendors are "Ballers Grill" and
"Cocktail Cart", which identify it as **Raimondi Park**, 1800 Wood St, West Oakland — the
Oakland Ballers' ballpark. The row keeps the workbook's spelling; only the location moved.

**Mezclá** is still blank on purpose — the row's own note says *Food truck*, and its one
logged item (Wisconsin Cheese Curds) turns up again as a **vendor at an SFCFC match**.
A truck has no fixed location, so it stays in the no-location group.

53 log rows carried stale location text for these eight (`North Bay, CA`,
`Prior repository`, `International Chain`) and now carry the venue's real location.

## Location spelling normalised (2026-09-07)

21 rows were filed under `san francisco` / `sf, ca` rather than `San Francisco, CA`,
which splits venue identity because identity is name + location. 20 were rewritten to the
canonical form. The 21st, **Lucania** (`sf, ca`), was the same venue as an existing
`San Francisco, CA` row, so it merged into it under the standard rule (lowest row kept,
first non-empty wins) and the twin was deleted — the table is now 2214 rows.

Worth a guard in the app: the location field should canonicalise on save, or the data check
should list rows whose location differs from a known location only by case or abbreviation.

## Open items (workbook state after the 2026-09-07 pass)

**Every logged row now resolves to a venue.** Unmatched log rows went 208 -> 38 -> **0**, and
`Has Menu Detail` rows with nothing logged went 54 -> 29 -> **0** (the orphan flags were
cleared outright — if a flag cannot be tied to an ordered item it is noise). 2219 venues,
990 logged items, 41 rows with no location.

Two things had been happening in the log: real spelling drift, and names simply shorter than
the venue row ("Laureate" for The Laureate, "Roof 106" for Roof 106 at The Matheson, "Levi's"
for Levi's Stadium). About 190 log rows were repointed at their venue and given its location.

Five venues were **added** for names that had been logged against nothing:

| Added | Location | What it is |
|---|---|---|
| As Quoted | San Francisco, CA (3613 Sacramento St) | restaurant on Sacramento St |
| SFCFC | San Francisco, CA | San Francisco City FC — match catering; the log's vendors (Adelita's Antojitos, Chairman Bao, Chidos Pizza, El Fuego, Mezclá, Saltwater bakeshop, Tacos King Maya) rotate by venue |
| SFO | San Francisco, CA | the airport — many vendors under it, "The Club" so far |
| El Yucatero | San Francisco, CA | |
| Maison | *unconfirmed* | logged under `North Bay, CA`; its items (Burrata, Troubadour Grilled Cheese) point at Healdsburg but nothing confirms it — **ask before filling it in** |

SFCFC and SFO behave like Levi's Stadium and Raimondo Park: one venue row, many vendors
underneath it in the log. The vendor column is what carries the detail.

**Jane** and **Réveille** are separate chains, not duplicates — their several rows are real
branches. The log rows that named only the chain were attached to the chain-level rows
(`Jane the Bakery`, `Réveille Coffee Co.`, both `Local Chain`), which is where they were
already filed.

Remaining:

- **41 rows with no location** — fine by design; single-location places, chains, and the
  one food truck.
- The **Repair menu flags** and **Attach logged items** buttons should now report nothing.
  If either starts finding rows again, something upstream is writing names that do not match.

## Working on this across sessions

Chat threads are disposable; this file is not. The routine:

1. **Knowledge lives here, not in the thread.** Anything worth surviving — a data-model
   fact, a convention, a UI decision, an open item — gets written into this file in the
   same commit as the change that prompted it.
2. **Start each work block in a fresh Cowork session**, linked to the Mac, with exactly
   two folders connected — no more:
   - `~/Desktop/Invisible Hand/Apps/eat + drink/food-drink-app`
   - `~/Library/CloudStorage/OneDrive-Personal/the list/Food + Drink/Claude`

   Do **not** also connect the parent `eat + drink`: it holds nothing but the repo, so the
   same files arrive under two mount paths and edits can be made through the wrong one.
   Opening line: *"read CLAUDE.md in food-drink-app, then …"*. That is the whole handoff.
3. **Run the tests before and after any change**: `node tests/smoke.mjs` (`npm i jsdom`
   once). They stub MSAL and Graph, so the live workbook is never touched.
4. **Verify every write, and give each write a fresh staged filename.** The bridge
   appears to cache by staged path: committing a second, different file from the *same*
   staged path silently delivered the OLD bytes (seen twice on 2026-09-06 — index.html and
   CLAUDE.md). Write to a new name (`index-v22.html`) each time, then stage the file back
   from the Mac and compare `md5sum` before committing in git.
5. **Verify every write.** A `device_commit_files` call has reported success while the
   Mac still held the old bytes. After writing a file to the Mac, stage it back and compare
   (`md5sum`) before committing it in git; only then push.
6. **Push**: GitHub Desktop, per the convention above. If its change list shows 0 files
   while the tree really has changes, or its menus come back disabled, it has lost track —
   click it to the front once and it re-syncs. The cloud sandbox's git proxy
   refuses credentials for this repo unless `nameisredacted/food-drink-app` is added to
   the session's authorized sources — do that and pushes can happen straight from the
   session instead.

## Testing

`node tests/smoke.mjs` (needs `npm i jsdom`). It loads `index.html` in jsdom with MSAL and
Graph stubbed, drives the real handlers, and asserts on the rows the app would have written.
No network, no OneDrive, nothing touches the live workbook.
