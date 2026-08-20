# Nicholas D. Gray — personal site

A small static site: a landing page, a full resume page, and downloadable Word
and PDF copies of the resume. No build tooling, no dependencies, no framework —
plain HTML and one stylesheet, plus a Python script that keeps the resume page
and the two downloads in sync.

| Page | File |
| --- | --- |
| Landing | [`index.html`](index.html) |
| Resume | [`resume.html`](resume.html) — **generated, see below** |
| Outside work | [`hobbies.html`](hobbies.html) |
| Contact | [`contact-me.html`](contact-me.html) |

## Updating the resume

[`data/resume.json`](data/resume.json) is the single source of truth. Edit it,
then run:

```bash
python3 tools/build_resume.py
```

That regenerates all three outputs from the same data, so the page and the
downloads can never disagree:

- `resume.html` — the styled web page
- `assets/Nicholas_Gray_Resume.pdf` — print version, 2 pages
- `assets/Nicholas_Gray_Resume.docx` — editable Word version

The script is standard library only. The PDF step shells out to a headless
Chrome or Chromium (it looks in the usual places, or set `CHROME_PATH`), and
falls back to LibreOffice converting the `.docx`. If neither is installed the
HTML and `.docx` are still written, and you can open `build/resume-print.html`
in a browser and print to PDF by hand.

**Do not hand-edit `resume.html`** — the next build overwrites it. Everything
about that page comes from `data/resume.json` (content) and
[`tools/build_resume.py`](tools/build_resume.py) (layout).

### Variants

A variant is a targeted version of the resume for a different kind of role. Each
one is a file in `data/variants/` holding **only what differs** from the base;
everything it does not mention — dates, employers, education, publications,
contact details — comes from `data/resume.json`, so the two can never disagree
on a fact neither meant to change.

```bash
python3 tools/build_resume.py --variant research   # one variant
python3 tools/build_resume.py --all                # base + every variant
```

A variant may override any top-level field, plus two of its own:

- `section_order` — which sections appear and in what order. The research
  variant puts publications ahead of projects, which is what that audience
  reads first.
- `experience_overrides` — keyed by employer, normally supplying just
  `bullets`. The same job is described with different emphasis without
  restating its title or dates. Naming an employer that is not in the base
  file is an error rather than a silent no-op.

Variants produce `assets/Nicholas_Gray_Resume_<Name>.pdf` and `.docx` only. The
website always shows the default version — variants exist to be attached to an
application, not published.

Two variants exist so far:

| Variant | For | Leads with | Drops |
| --- | --- | --- | --- |
| *(default)* | Data science roles | XGBoost, pipelines, AI & automation | survey methods, IRB, SPSS |
| `research` | Academic and nonprofit research, evaluation | mixed methods, IRB, Qualtrics, publications above projects | machine learning framing |
| `survey` | Survey research science, consumer insights | survey design *and* modeling, on insurance domain experience | nothing much — it is the widest of the three |

`survey` exists because some roles want both halves at once: survey
methodology and statistical modeling, in Python and SQL, against a commercial
domain. Neither of the other two covers that on its own — the default buries
the survey work, and `research` drops the machine learning.

Keep skills honest per variant rather than listing everything everywhere. SPSS
belongs on the research-facing versions and would read as legacy on the data
science one; XGBoost is the reverse.

The other three pages *are* hand-written, so edit them directly. They each carry
their own copy of the masthead and footer; if you change the navigation, change
it in all four places — the build script's copy lives in the `masthead()` and
`footer()` functions.

### One deliberate difference between the page and the downloads

The phone number appears only in the `.pdf` and `.docx`, never on the web page,
so it does not get scraped. `web_contact` in the JSON controls which contact
details reach the page; the documents always carry the full set.

The `summary` field is written in the conventional third-person resume voice and
is used verbatim on `resume.html` and in both documents. The first-person intro
on the landing page is separate prose that lives in `index.html`, alongside the
rest of that page's hand-written copy.

## Design

One stylesheet, [`assets/style.css`](assets/style.css), with CSS custom
properties for the palette. It follows the reader's system light/dark setting by
default, and the toggle in the masthead overrides it and remembers the choice in
`localStorage`.

Both themes are defined with tokens on `:root` — light on the bare selector,
dark under both `@media (prefers-color-scheme: dark)` and `[data-theme="dark"]`
so the toggle works in both directions. Adding a colour anywhere else will break
one of the two themes.

`@media print` is set up so `resume.html` prints sensibly straight from the
browser, independently of the generated PDF.

## Previewing locally

Any static file server, or just open the files. For example:

```bash
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Publishing

The site is plain static files at the repository root, so GitHub Pages can serve
it as-is: **Settings → Pages → Deploy from a branch → `main` / `(root)`**.

Note that this repository is currently **private**. GitHub Pages will not serve a
private repository on the Free plan — the repository has to be made public (or
the account upgraded) before a Pages site goes live.
