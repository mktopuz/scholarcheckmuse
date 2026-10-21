# Thesis & Dissertation Format Checker

A browser-only tool that checks a Word `.docx` thesis or dissertation against
a school's formatting guidelines: page size, margins, line spacing, body font
and size, heading styles, page numbers, and the required parts for the
document type (master's thesis or doctoral dissertation).

## Privacy by design

- The document is opened and checked **inside the browser tab**. It is never
  uploaded, stored, or copied anywhere — there is no server.
- Only the report exists, and only on the page. The user can download it as
  CSV or print it; closing the tab discards everything.
- An optional AI review is the only network feature. It activates only when
  the user enters their own LLM API key (kept in the browser's local
  storage). It sends *text excerpts* — never the file — to the provider the
  user configured.

## How it works

- `index.html`, `styles.css` — the page.
- `js/profiles.js` — guideline profiles. Adding a school means adding one
  entry; every value carries a `verified` flag until checked against the
  school's current official guide. First verified school: **University of
  Tulsa — Graduate School**, read from its official Thesis and Dissertation
  Word template (version 2022), including its direct-format heading spec
  (1.1 centered bold, 1.1.1 left italic, 1.1.1.1 left italic double-spaced).
  Users can also add and save their own school's guidelines in the page
  (stored in the browser, with JSON export/import).
- `js/docx.js` — minimal .docx reader (sections, styles, theme fonts,
  headers/footers).
- `js/checker.js` — the rule engine. Each result reports expected vs. found,
  where it occurs (section / paragraph number / first words — a .docx stores
  no page numbers), and the Word menu path that fixes it.
- `js/ai.js` — optional AI review against any OpenAI-compatible endpoint.
- `js/app.js` — UI wiring and report rendering.

Open `index.html` in a browser (or serve the folder statically, e.g. GitHub
Pages) and drop in a `.docx`.

## Checks in v1

Page size per section · margins per section · body font (dominant, weighted
by text volume) · body size · line spacing (resolved like Word: paragraph,
then style, then defaults) · Heading 1/2/3/4 style usage (or an info note
where a school's spec uses direct formatting instead) · per-level heading
formats (alignment, bold, italic, spacing) where a profile defines them ·
bold pseudo-heading detection · heading font consistency · page-number field
presence, location and alignment · required front-matter parts per document
type · abstract word count against the school's per-document-type limit where
one is defined (e.g. TU: 150 thesis / 350 dissertation) · page-numbering
scheme where a profile defines one (e.g. TU: lowercase roman preliminaries,
body restarting at Arabic 1), judged from the sections' number-format
settings · conditional parts (a List of Tables / List of Figures is required
only when the document actually contains tables / figures) · table text size
(informational). Sections containing no text (e.g. empty transition sections
in official templates) are skipped for page setup checks and reported as an
info line. Every report also carries an **overall score (0–100)**: checks
are grouped by rule, the worst verdict in a group counts, core formatting
rules weigh 3, structure and required parts weigh 2, the advisory
pseudo-heading warning weighs 1, and informational rows never score
(pass = full credit, review = half). Bands: 90+ Excellent — ready to
submit · 75+ Good — minor fixes · 50+ Needs work · below 50 Major rework
needed.

## Honest limits

A .docx does not store where pages break, so the report never invents page
numbers. The tool checks formatting only — not content or citation style —
and it never modifies the document.
