# CLAUDE.md — Report Template Project
# A Pathway to You — report-template

Read the parent folder's CLAUDE.md first (shared GAS conventions). This file covers
everything specific to this project: what it does, file inventory, key conventions,
and current state.

---

## What This Is

A Google Doc–bound GAS script that automates report production for assessment clients.
Bound to a report document template. Three main functions:

1. **Pronoun replacement** — highlights and replaces gendered pronouns throughout the doc
2. **Score entry and insertion** — enters measure scores via sidebar, inserts formatted
   score tables and data-source lists into the doc using `{{placeholders}}`
3. **Resource insertion** — filters and inserts recommended resources from the resource
   database into the doc at the `{{RESOURCES}}` placeholder

---

## File Inventory

| File | Purpose |
|------|---------|
| `Code.js` | Main entry point, `doGet`, `onOpen`, menu setup, placeholder clearing, score table logic, `pushSnippet` / `pushAllSnippets`, `lookupCase`, `getActiveMeasures`, `SCORE_MEASURES` array |
| `ReportGenerator.js` | `insertResources` — filters and inserts resources at `{{RESOURCES}}` placeholder in the report doc. Reads from resource Sheet by ID. Tracks inserted links via ScriptProperties. |
| `ResoListGenerator.js` | `insertResources` variant for generating resource lists in a document (different from ReportGenerator — uses DocumentApp.getActiveDocument() rather than opening by ID) |
| `SetupSidebar.html` | Step 1 — link case (ehrId), set pronouns |
| `ScoreSidebar.html` | Step 3 — score entry UI, grouped by measure category |
| `ReportSidebar.html` | Resource filter UI (population, format, cost) |
| `Cleanup.js` | One-time taxonomy normalization for the resource sheet |

*(Add new files here when created)*

---

## Key IDs

- **Report doc template ID:** `1HCD-VOk1V1G4-qCGLXGK7rZl-IGQIRmfImlp7OP9748`
- **Resource Sheet ID:** `173Gsuzo6RdzGSjUPlWDg3JwXe8PWY6XblyMzlNJK3zc`
- **Resource Sheet tab:** `full_list`

---

## Placeholder System

The report doc uses `{{placeholder}}` markers for all automated insertions:

| Placeholder | Inserted by | What it becomes |
|-------------|-------------|-----------------|
| `{{RESOURCES}}` | `insertResources` | Filtered resource list (H2/H3 headings + bullet links) |
| `{{DATA-SOURCES}}` | `generateDataSources` in Code.js | Formatted list of scored measures |
| `{{slot-id}}` (e.g. `{{ASD-A1}}`) | `pushSnippet` / `pushAllSnippets` | Bullet-point content snippets |

**Rule:** Always clear placeholders before finalizing a report. Run "Remove Placeholders"
from the menu. The `clearPlaceholders()` function also strips `^` skip markers.

---

## SCORE_MEASURES Array

Defined in `Code.js`. This is the **single source of truth** for all scored measures.
To add a measure, append one object to the correct group in this array — nothing else
needs to change. The score table and sidebar both generate themselves from it automatically.

Groups: `General`, `ADHD`, `Autism`, `Sensory`, `Other-Report`

---

## Resource Data — Column Map (0-based)

From `full_list` tab of the resource Sheet:

| Index | Field |
|-------|-------|
| 0 | title |
| 1 | desc |
| 2 | link |
| 3 | format |
| 4 | population |
| 5 | cost |
| 6 | pages |

**`splitT_(str)`** — shared helper, splits comma-separated tag strings. Defined locally
in both `ReportGenerator.js` and `ResoListGenerator.js`.

---

## Link Tracking

`ReportGenerator.js` tracks which resources have already been inserted using
`ScriptProperties` under the key `insertedLinks_v1`. This prevents duplicate insertions
across multiple runs. To reset (allow re-insertion): run the diagnostic reset function
or call `PropertiesService.getScriptProperties().deleteProperty('insertedLinks_v1')`.

---

## Case Linking

`lookupCase(ehrId)` in `Code.js` connects this doc to the Assessment Control System's
Cases sheet. Used by `getActiveMeasures()` to filter the score sidebar to only the
measures assigned to the linked client (`ActiveForms` + NovoPsych forms).

If no case is linked, the score sidebar falls back to showing all measures.

---

## What Is Already Working — Do Not Break

- Three-step sidebar flow (Setup → Pronouns → Score Entry)
- Pronoun highlight, verb-fix, and replace workflow
- Score entry and `{{DATA-SOURCES}}` insertion
- Resource insertion at `{{RESOURCES}}` with population/format/cost filtering
- Link tracking (insertedLinks_v1) preventing duplicate resource insertions
- `clearPlaceholders()` removing all `{{...}}` and `^` markers
- `pushSnippet` / `pushAllSnippets` for slot-based content insertion

---

## Known Issues

None documented yet. Add here as discovered.

---

## Current To-Do

Nothing active. Add items here as work is planned.

---

## Sync Workflow

This repo does not use clasp + sync.sh (no `.clasp.json` / GAS project sync configured
at time of writing). Changes are currently made directly in GAS and backed up to GitHub
manually. Update this section if clasp sync is set up for this repo.
