#!/usr/bin/env python3
"""
Build the resume page and its downloadable versions from data/resume.json.

    python3 tools/build_resume.py

Writes:
    resume.html                          the styled web page
    assets/Nicholas_Gray_Resume.docx     editable Word version (ATS-friendly)
    assets/Nicholas_Gray_Resume.pdf      print version
    build/resume-print.html              intermediate the PDF is rendered from

Standard library only, except that the PDF step shells out to a headless
Chrome/Chromium if one can be found, and falls back to LibreOffice converting
the .docx. If neither exists the .docx and .html are still written and the PDF
step is skipped with a warning.
"""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data" / "resume.json"
BUILD = ROOT / "build"

BASENAME = "Nicholas_Gray_Resume"

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

TAG_RE = re.compile(r"<[^>]+>")


def strip_tags(s: str) -> str:
    """Inline HTML (<em>) out, plain text in — for the Word version."""
    return html.unescape(TAG_RE.sub("", s))


def e(s: str) -> str:
    """Escape text for HTML."""
    return html.escape(s, quote=True)


def xe(s: str) -> str:
    """Escape text for XML (OOXML)."""
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def load() -> dict:
    with DATA.open(encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------------------
# shared markup fragments
# ---------------------------------------------------------------------------

ICONS = {
    "mail": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m2 7 10 6 10-6"/></svg>',
    "pin": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/></svg>',
    "github": '<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 .5C5.73.5.98 5.24.98 11.52c0 4.86 3.15 8.98 7.52 10.44.55.1.75-.24.75-.53 0-.26-.01-1.13-.02-2.05-3.06.66-3.71-1.3-3.71-1.3-.5-1.27-1.22-1.61-1.22-1.61-1-.68.08-.67.08-.67 1.1.08 1.68 1.13 1.68 1.13.98 1.68 2.57 1.2 3.2.92.1-.71.38-1.2.7-1.47-2.44-.28-5.01-1.22-5.01-5.44 0-1.2.43-2.18 1.13-2.95-.11-.28-.49-1.4.11-2.92 0 0 .92-.3 3.02 1.13a10.4 10.4 0 0 1 5.5 0c2.1-1.43 3.02-1.13 3.02-1.13.6 1.52.22 2.64.11 2.92.7.77 1.13 1.75 1.13 2.95 0 4.23-2.58 5.16-5.03 5.43.4.34.75 1 .75 2.03 0 1.47-.02 2.65-.02 3.01 0 .29.2.64.76.53a10.55 10.55 0 0 0 7.51-10.44C23.02 5.24 18.27.5 12 .5Z"/></svg>',
    "phone": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2Z"/></svg>',
    "download": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></svg>',
    "print": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8" rx="1"/></svg>',
    "external": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg>',
    "code": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m16 18 6-6-6-6"/><path d="m8 6-6 6 6 6"/></svg>',
}

THEME_TOGGLE = """      <button class="theme-toggle" type="button" id="theme-toggle"
              aria-label="Switch between light and dark theme">
        <svg class="icon-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8Z"/></svg>
        <svg class="icon-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>
      </button>"""

# Applied before first paint so the chosen theme never flashes.
THEME_SCRIPT = """  <script>
    (function () {
      try {
        var saved = localStorage.getItem('theme');
        if (saved === 'dark' || saved === 'light') {
          document.documentElement.setAttribute('data-theme', saved);
        }
      } catch (err) { /* private mode: fall back to the system theme */ }
    })();
  </script>"""

THEME_INIT = """  <script>
    (function () {
      var btn = document.getElementById('theme-toggle');
      if (!btn) return;
      btn.addEventListener('click', function () {
        var root = document.documentElement;
        var explicit = root.getAttribute('data-theme');
        var dark = explicit
          ? explicit === 'dark'
          : window.matchMedia('(prefers-color-scheme: dark)').matches;
        var next = dark ? 'light' : 'dark';
        root.setAttribute('data-theme', next);
        try { localStorage.setItem('theme', next); } catch (err) { /* ignore */ }
      });
    })();
  </script>"""


def masthead(active: str) -> str:
    def link(href: str, label: str, key: str) -> str:
        cur = ' aria-current="page"' if key == active else ""
        return f'<a href="{href}"{cur}>{label}</a>'

    return f"""  <header class="masthead">
    <div class="wrap masthead__inner">
      <a class="brand" href="index.html">
        <span class="brand__mark" aria-hidden="true">NG</span>
        <span>Nick Gray</span>
      </a>
      <nav class="nav" aria-label="Main">
        {link("index.html", "Home", "home")}
        {link("resume.html", "Resume", "resume")}
        {link("hobbies.html", "Outside Work", "hobbies")}
        {link("contact-me.html", "Contact", "contact")}
      </nav>
{THEME_TOGGLE}
    </div>
  </header>"""


def footer(d: dict) -> str:
    gh = d["contact"]["github"]
    return f"""  <footer class="footer">
    <div class="wrap footer__inner">
      <p>&copy; {date.today().year} Nicholas D. Gray</p>
      <div class="footer__links">
        <a href="mailto:{e(d['contact']['email'])}">Email</a>
        <a href="{e(gh)}" rel="noopener">GitHub</a>
        <a href="{e(d['downloads']['pdf'])}">Resume&nbsp;PDF</a>
      </div>
    </div>
  </footer>"""


def head(title: str, description: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{e(title)}</title>
  <meta name="description" content="{e(description)}">
  <meta name="color-scheme" content="light dark">
  <link rel="stylesheet" href="assets/style.css">
  <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='7' fill='%230d7361'/><text x='16' y='22' font-family='monospace' font-size='15' font-weight='700' fill='white' text-anchor='middle'>NG</text></svg>">
{THEME_SCRIPT}
</head>"""


def contact_rail(d: dict) -> str:
    c = d["contact"]
    parts = []
    for key in d["web_contact"]:
        if key == "email":
            parts.append(
                f'<li>{ICONS["mail"]}<a href="mailto:{e(c["email"])}">{e(c["email"])}</a></li>'
            )
        elif key == "location":
            parts.append(f'<li>{ICONS["pin"]}<span>{e(c["location_long"])}</span></li>')
        elif key == "github":
            parts.append(
                f'<li>{ICONS["github"]}<a href="{e(c["github"])}" rel="noopener">'
                f'{e(c["github_label"])}</a></li>'
            )
        elif key == "phone":
            digits = re.sub(r"\D", "", c["phone"])
            parts.append(
                f'<li>{ICONS["phone"]}<a href="tel:+1{digits}">{e(c["phone"])}</a></li>'
            )
    return '<ul class="contact-rail">\n        ' + "\n        ".join(parts) + "\n      </ul>"


def project_cards(d: dict) -> str:
    cards = []
    for p in d["projects"]:
        tags = "".join(f"<li>{e(t)}</li>" for t in p["tags"])
        links = [
            f'<a href="{e(p["repo"])}" rel="noopener">{ICONS["code"]}Source</a>'
        ]
        if p.get("demo"):
            links.append(
                f'<a href="{e(p["demo"])}" rel="noopener">{ICONS["external"]}'
                f'{e(p["demo_label"])}</a>'
            )
        cards.append(f"""        <article class="project">
          <div class="project__head">
            <h3 class="project__name"><a href="{e(p["repo"])}" rel="noopener">{e(p["name"])}</a></h3>
            <span class="lang" data-lang="{e(p["language"])}"><span class="dot"></span>{e(p["language"])}</span>
          </div>
          <p class="project__blurb">{p["blurb"]}</p>
          <ul class="project__tags">{tags}</ul>
          <div class="project__links">{"".join(links)}</div>
        </article>""")
    return '<div class="project-grid">\n' + "\n".join(cards) + "\n      </div>"


# ---------------------------------------------------------------------------
# resume.html
# ---------------------------------------------------------------------------


def build_site_html(d: dict) -> str:
    c = d["contact"]

    stats = "".join(
        f"""        <div class="stat">
          <div class="stat__value">{e(s["value"])}</div>
          <div class="stat__label">{e(s["label"])}</div>
        </div>"""
        for s in d["stats"]
    )

    jobs = []
    for job in d["experience"]:
        cur = " job--current" if job.get("current") else ""
        bullets = "".join(f"\n            <li>{b}</li>" for b in job["bullets"])
        jobs.append(f"""        <article class="job{cur}">
          <div class="job__year" aria-hidden="true">{e(job["year_label"])}</div>
          <div class="job__head">
            <h3 class="job__role">{e(job["role"])}</h3>
            <span class="job__dates">{e(job["start"])} &ndash; {e(job["end"])}</span>
          </div>
          <div class="job__org">{e(job["org"])}<span class="sep">&middot;</span><span class="place">{e(job["place"])}</span></div>
          <ul>{bullets}
          </ul>
        </article>""")

    skills = []
    for g in d["skills"]:
        chips = []
        for item in g["items"]:
            if item.get("level"):
                chips.append(
                    f'<li class="chip chip--leveled">{e(item["name"])}'
                    f'<span class="chip__level">{e(item["level"])}</span></li>'
                )
            else:
                chips.append(f'<li class="chip">{e(item["name"])}</li>')
        skills.append(f"""        <div class="skill-group">
          <h3>{e(g["group"])}</h3>
          <ul class="chips">{"".join(chips)}</ul>
        </div>""")

    edu = "".join(
        f"""        <li class="edu">
          <div class="edu__year">{e(x["year"])}</div>
          <div>
            <div class="edu__degree">{e(x["degree"])}</div>
            <div class="edu__school">{e(x["school"])}</div>
            {f'<div class="edu__note">{e(x["note"])}</div>' if x.get("note") else ""}
          </div>
        </li>"""
        for x in d["education"]
    )

    pubs = "".join(f"\n        <li>{p}</li>" for p in d["publications"])

    return f"""{head(f"{d['name']} — Resume", f"{d['title']} in {c['location']}. {strip_tags(d['tagline'])}")}

<body>
  <a class="skip-link" href="#main">Skip to content</a>
{masthead("resume")}

  <main id="main">
    <section class="wrap hero">
      <div class="hero__grid">
        <div class="monogram" aria-hidden="true">{e(d["initials"])}</div>
        <div>
          <p class="eyebrow">Resume</p>
          <h1>{e(d["name"])}</h1>
          <p class="hero__role"><strong>{e(d["title"])}</strong> &middot; {e(d["tagline"])}</p>
          <p class="hero__summary">{d["summary"]}</p>
          {contact_rail(d)}
        </div>
      </div>

      <div class="stats">
{stats}
      </div>

      <div class="downloads no-print">
        <div class="downloads__label">Take it with you
          <span>Same content, formatted for print and for editing.</span>
        </div>
        <div class="btn-row">
          <a class="btn btn--primary" href="{e(d["downloads"]["pdf"])}" download>
            {ICONS["download"]}PDF<span class="btn__meta">.pdf</span>
          </a>
          <a class="btn" href="{e(d["downloads"]["docx"])}" download>
            {ICONS["download"]}Word<span class="btn__meta">.docx</span>
          </a>
          <button class="btn" type="button" onclick="window.print()">
            {ICONS["print"]}Print
          </button>
        </div>
      </div>
    </section>

    <section class="wrap section" id="experience">
      <div class="section__head">
        <span class="section__num">01</span>
        <h2>Experience</h2>
        <span class="section__aside">2013 &rarr; present</span>
      </div>
      <div class="timeline">
{chr(10).join(jobs)}
      </div>
    </section>

    <section class="wrap section" id="skills">
      <div class="section__head">
        <span class="section__num">02</span>
        <h2>Technical Skills</h2>
      </div>
      <div class="skill-grid">
{chr(10).join(skills)}
      </div>
    </section>

    <section class="wrap section" id="projects">
      <div class="section__head">
        <span class="section__num">03</span>
        <h2>Public Projects</h2>
        <span class="section__aside"><a href="{e(c["github"])}" rel="noopener">{e(c["github_label"])}</a></span>
      </div>
      {project_cards(d)}
    </section>

    <section class="wrap section" id="education">
      <div class="section__head">
        <span class="section__num">04</span>
        <h2>Education</h2>
      </div>
      <ul class="edu-list">
{edu}
      </ul>
      <div class="coursework">
        <strong>Graduate methods training</strong>
        {e(d["methods_training"])}
      </div>
      <div class="coursework">
        <strong>Additional coursework</strong>
        {e(d["coursework"])}
      </div>
    </section>

    <section class="wrap section" id="publications">
      <div class="section__head">
        <span class="section__num">05</span>
        <h2>Selected Publications</h2>
      </div>
      <ol class="pubs">{pubs}
      </ol>
      <p class="pubs-note">{e(d["publications_note"])}</p>
    </section>
  </main>

{footer(d)}
{THEME_INIT}
</body>

</html>
"""


# ---------------------------------------------------------------------------
# print HTML (source for the PDF)
# ---------------------------------------------------------------------------

PRINT_CSS = """
  @page { size: letter; margin: 0.6in 0.65in; }
  * { box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body {
    margin: 0;
    font-family: "Source Sans Pro", "Segoe UI", Calibri, Carlito, Helvetica, Arial, sans-serif;
    font-size: 9.9pt;
    line-height: 1.36;
    color: #1a2422;
  }
  .mono {
    font-family: "DejaVu Sans Mono", "Liberation Mono", Consolas, monospace;
  }
  h1, h2, h3 { margin: 0; font-weight: 700; }

  header.top { border-bottom: 2px solid #0d7361; padding-bottom: 7pt; margin-bottom: 9pt; }
  h1.name {
    font-size: 21pt; letter-spacing: -0.4pt; line-height: 1.05; color: #10201d;
  }
  .role {
    font-size: 9.6pt; font-weight: 700; letter-spacing: 1.5pt;
    text-transform: uppercase; color: #0d7361; margin-top: 3pt;
  }
  .contact { font-size: 8.9pt; color: #475653; margin-top: 5pt; }
  .contact .sep { color: #0d7361; margin: 0 5pt; }

  .summary { font-size: 9.7pt; color: #263533; margin-bottom: 11pt; }

  h2.sec {
    font-size: 8.6pt; letter-spacing: 1.5pt; text-transform: uppercase;
    color: #0d7361; border-bottom: 0.75pt solid #cfd8d5;
    padding-bottom: 2.5pt; margin: 0 0 6pt;
    /* never let a section heading be the last thing on a page */
    break-after: avoid; page-break-after: avoid;
  }
  section { margin-bottom: 10pt; }
  p, .proj-desc, .job li, .pub { orphans: 2; widows: 2; }

  .job { margin-bottom: 8pt; page-break-inside: avoid; }
  .job:last-child { margin-bottom: 0; }
  .job-head { display: flex; justify-content: space-between; align-items: baseline; gap: 10pt; }
  .job-role { font-size: 10.4pt; font-weight: 700; color: #10201d; }
  .job-dates { font-size: 8.4pt; color: #5c6b68; white-space: nowrap; }
  .job-org { font-size: 9.2pt; color: #3d4d4a; font-style: italic; margin-bottom: 2.5pt; }
  ul { margin: 0; padding: 0; list-style: none; }
  .job li {
    position: relative; padding-left: 9pt; margin-bottom: 1.6pt; color: #263533;
  }
  .job li::before {
    content: ""; position: absolute; left: 0; top: 5.2pt;
    width: 3.6pt; height: 0.9pt; background: #0d7361;
  }

  .skills-table { width: 100%; border-collapse: collapse; }
  .skills-table td { padding: 1.6pt 0; vertical-align: top; }
  .skills-table td.k {
    width: 96pt; font-weight: 700; font-size: 8.8pt; color: #0d7361;
    padding-right: 8pt; white-space: nowrap;
  }
  .skills-table td.v { font-size: 9.4pt; color: #263533; }

  .proj { margin-bottom: 5.5pt; page-break-inside: avoid; }
  .proj:last-child { margin-bottom: 0; }
  .proj-head { font-size: 9.6pt; }
  .proj-name { font-weight: 700; color: #10201d; }
  .proj-url { font-size: 8.3pt; color: #5c6b68; }
  .proj-desc { font-size: 9.3pt; color: #263533; }

  .edu { display: flex; justify-content: space-between; align-items: baseline;
         gap: 10pt; margin-bottom: 1pt; }
  .edu-block { margin-bottom: 5.5pt; page-break-inside: avoid; }
  .edu-block:last-of-type { margin-bottom: 0; }
  .edu-degree { font-size: 9.9pt; font-weight: 700; color: #10201d; }
  .edu-year { font-size: 8.4pt; color: #5c6b68; white-space: nowrap; }
  .edu-school { font-size: 9.2pt; color: #3d4d4a; }
  .edu-note { font-size: 8.8pt; color: #5c6b68; font-style: italic; }
  .coursework { font-size: 9pt; color: #3d4d4a; margin-top: 6pt; }
  .coursework b { color: #0d7361; }

  .pub { font-size: 8.9pt; color: #263533; margin-bottom: 3pt; padding-left: 10pt;
         text-indent: -10pt; }
  .pub-note { font-size: 8.7pt; color: #5c6b68; font-style: italic; margin-top: 4pt; }
"""


def build_print_html(d: dict) -> str:
    c = d["contact"]

    contact_bits = [
        e(c["email"]),
        e(c["phone"]),
        e(c["location"]),
        e(c["github_label"]),
    ]
    contact_line = '<span class="sep">|</span>'.join(contact_bits)

    jobs = ""
    for job in d["experience"]:
        bullets = "".join(f"<li>{strip_tags(b)}</li>" for b in job["bullets"])
        jobs += f"""      <div class="job">
        <div class="job-head">
          <div class="job-role">{e(job["role"])}</div>
          <div class="job-dates mono">{e(job["start"])} &ndash; {e(job["end"])}</div>
        </div>
        <div class="job-org">{e(job["org"])} &middot; {e(job["place"])}</div>
        <ul>{bullets}</ul>
      </div>
"""

    skills = ""
    for g in d["skills"]:
        vals = ", ".join(
            f'{i["name"]} ({i["level"]})' if i.get("level") else i["name"]
            for i in g["items"]
        )
        skills += (
            f'        <tr><td class="k">{e(g["group"])}</td>'
            f'<td class="v">{e(vals)}</td></tr>\n'
        )

    projects = ""
    for p in d["projects"]:
        url = p["demo"] or p["repo"]
        label = p["demo_label"] or p["repo"].replace("https://", "")
        projects += f"""      <div class="proj">
        <div class="proj-head"><span class="proj-name">{e(p["name"])}</span>
          <span class="proj-url mono">&nbsp;&mdash;&nbsp;{e(p["language"])} &middot; {e(label)}</span>
        </div>
        <div class="proj-desc">{strip_tags(p["blurb_plain"])}</div>
      </div>
"""

    edu = ""
    for x in d["education"]:
        note = f'<div class="edu-note">{e(x["note"])}</div>' if x.get("note") else ""
        edu += f"""      <div class="edu-block">
        <div class="edu">
          <div class="edu-degree">{e(x["degree"])}</div>
          <div class="edu-year mono">{e(x["year"])}</div>
        </div>
        <div class="edu-school">{e(x["school"])}</div>
        {note}
      </div>
"""

    pubs = "".join(f'      <div class="pub">{strip_tags(p)}</div>\n' for p in d["publications"])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{e(d["name"])} — Resume</title>
<style>{PRINT_CSS}</style>
</head>
<body>
  <header class="top">
    <h1 class="name">{e(d["name"])}</h1>
    <div class="role">{e(d["title"])}</div>
    <div class="contact mono">{contact_line}</div>
  </header>

  <div class="summary">{strip_tags(d["summary"])}</div>

  <section>
    <h2 class="sec">Technical Skills</h2>
    <table class="skills-table">
{skills}    </table>
  </section>

  <section>
    <h2 class="sec">Professional Experience</h2>
{jobs}  </section>

  <section>
    <h2 class="sec">Selected Projects</h2>
{projects}  </section>

  <section>
    <h2 class="sec">Education</h2>
{edu}    <div class="coursework"><b>Graduate methods training:</b> {e(d["methods_training"])}</div>
    <div class="coursework"><b>Additional coursework:</b> {e(d["coursework"])}</div>
  </section>

  <section>
    <h2 class="sec">Selected Publications</h2>
{pubs}    <div class="pub-note">{e(d["publications_note"])}</div>
  </section>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# .docx
# ---------------------------------------------------------------------------

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

TEAL = "0D7361"
INK = "10201D"
BODY = "263533"
SOFT = "3D4D4A"
MUTED = "5C6B68"
RULE = "CFD8D5"

CONTENT_W = 10080  # twips inside 0.7" margins on letter


def r(text, *, b=False, i=False, color=None, sz=None, caps=False, spacing=None):
    """One OOXML run. `sz` is in half-points (20 == 10pt)."""
    props = []
    if b:
        props.append("<w:b/>")
    if i:
        props.append("<w:i/>")
    if caps:
        props.append("<w:caps/>")
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    if spacing:
        props.append(f'<w:spacing w:val="{spacing}"/>')
    if sz:
        props.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xe(text)}</w:t></w:r>'


def tab_run():
    return "<w:r><w:tab/></w:r>"


def para(runs, *, before=0, after=0, tabs=False, border=False,
         line=None, indent=None, numbered=False, keep_next=False):
    """
    One OOXML paragraph.

    The children of <w:pPr> are emitted in the order the wordprocessingml
    schema requires (keepNext, numPr, pBdr, tabs, spacing, ind). Word refuses
    to open the file if this sequence is wrong, so keep the appends in order.
    """
    props = []
    if keep_next:
        props.append("<w:keepNext/>")
    if numbered:
        props.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    if border:
        props.append(
            f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" '
            f'w:color="{RULE}"/></w:pBdr>'
        )
    if tabs:
        props.append(f'<w:tabs><w:tab w:val="right" w:pos="{CONTENT_W}"/></w:tabs>')
    sp = f'<w:spacing w:before="{before}" w:after="{after}"'
    if line:
        sp += f' w:line="{line}" w:lineRule="auto"'
    sp += "/>"
    props.append(sp)
    if indent:
        props.append(f'<w:ind w:left="{indent[0]}" w:hanging="{indent[1]}"/>')
    return f"<w:p><w:pPr>{''.join(props)}</w:pPr>{''.join(runs)}</w:p>"


def heading(text):
    return para(
        [r(text, b=True, color=TEAL, sz=17, caps=True, spacing=30)],
        before=220, after=90, border=True, keep_next=True,
    )


def build_docx_document(d: dict) -> str:
    c = d["contact"]
    body = []

    # --- header block
    body.append(para([r(d["name"], b=True, color=INK, sz=42, spacing=-8)], after=20))
    body.append(
        para([r(d["title"], b=True, color=TEAL, sz=19, caps=True, spacing=30)], after=60)
    )
    contact = f'{c["email"]}   |   {c["phone"]}   |   {c["location"]}   |   {c["github_label"]}'
    body.append(para([r(contact, color=MUTED, sz=17)], after=140, border=True))

    # --- summary
    body.append(para([r(strip_tags(d["summary"]), color=BODY, sz=20)], before=100, after=40))

    # --- skills
    body.append(heading("Technical Skills"))
    for g in d["skills"]:
        vals = ", ".join(
            f'{i["name"]} ({i["level"]})' if i.get("level") else i["name"]
            for i in g["items"]
        )
        body.append(
            para(
                [
                    r(f'{g["group"]}:  ', b=True, color=TEAL, sz=18),
                    r(vals, color=BODY, sz=19),
                ],
                after=40, indent=(1620, 1620),
            )
        )

    # --- experience
    body.append(heading("Professional Experience"))
    for job in d["experience"]:
        body.append(
            para(
                [
                    r(job["role"], b=True, color=INK, sz=21),
                    tab_run(),
                    r(f'{job["start"]} – {job["end"]}', color=MUTED, sz=17),
                ],
                before=110, after=0, tabs=True, keep_next=True,
            )
        )
        body.append(
            para(
                [r(f'{job["org"]}  ·  {job["place"]}', i=True, color=SOFT, sz=19)],
                after=50, keep_next=True,
            )
        )
        for bullet in job["bullets"]:
            body.append(
                para([r(strip_tags(bullet), color=BODY, sz=19)], after=30, numbered=True)
            )

    # --- projects
    body.append(heading("Selected Projects"))
    for p in d["projects"]:
        label = p["demo_label"] or p["repo"].replace("https://", "")
        body.append(
            para(
                [
                    r(p["name"], b=True, color=INK, sz=20),
                    r(f'  —  {p["language"]}  ·  {label}', color=MUTED, sz=17),
                ],
                before=90, after=10, keep_next=True,
            )
        )
        body.append(
            para([r(strip_tags(p["blurb_plain"]), color=BODY, sz=19)], after=30)
        )

    # --- education
    body.append(heading("Education"))
    for x in d["education"]:
        body.append(
            para(
                [
                    r(x["degree"], b=True, color=INK, sz=20),
                    tab_run(),
                    r(x["year"], color=MUTED, sz=17),
                ],
                before=90, after=0, tabs=True, keep_next=True,
            )
        )
        runs = [r(x["school"], color=SOFT, sz=19)]
        body.append(para(runs, after=0 if x.get("note") else 30))
        if x.get("note"):
            body.append(para([r(x["note"], i=True, color=MUTED, sz=18)], after=30))
    body.append(
        para(
            [
                r("Graduate methods training:  ", b=True, color=TEAL, sz=18),
                r(d["methods_training"], color=SOFT, sz=19),
            ],
            before=90, after=30,
        )
    )
    body.append(
        para(
            [
                r("Additional coursework:  ", b=True, color=TEAL, sz=18),
                r(d["coursework"], color=SOFT, sz=19),
            ],
            after=40,
        )
    )

    # --- publications
    body.append(heading("Selected Publications"))
    for p in d["publications"]:
        body.append(
            para([r(strip_tags(p), color=BODY, sz=18)], after=50, indent=(340, 340))
        )
    body.append(para([r(d["publications_note"], i=True, color=MUTED, sz=18)], before=60))

    sect = (
        "<w:sectPr>"
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1008" w:right="1080" w:bottom="1008" w:left="1080" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {W}><w:body>{''.join(body)}{sect}</w:body></w:document>"
    )


DOCX_STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}>
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:color w:val="{BODY}"/>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
      <w:lang w:val="en-US"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="252" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
  </w:style>
</w:styles>"""

DOCX_NUMBERING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {W}>
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#8211;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="360" w:hanging="220"/></w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:hint="default"/>
        <w:color w:val="{TEAL}"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""

DOCX_SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}>
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>"""

DOCX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""

DOCX_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

DOCX_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""


def build_docx(d: dict, out: Path) -> None:
    stamp = date.today().isoformat()
    core = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{xe(d["name"])} — Resume</dc:title>
  <dc:creator>{xe(d["name"])}</dc:creator>
  <cp:lastModifiedBy>{xe(d["name"])}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{stamp}T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""

    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
</Properties>"""

    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", DOCX_CONTENT_TYPES)
        z.writestr("_rels/.rels", DOCX_ROOT_RELS)
        z.writestr("word/document.xml", build_docx_document(d))
        z.writestr("word/_rels/document.xml.rels", DOCX_DOC_RELS)
        z.writestr("word/styles.xml", DOCX_STYLES)
        z.writestr("word/numbering.xml", DOCX_NUMBERING)
        z.writestr("word/settings.xml", DOCX_SETTINGS)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

CHROME_CANDIDATES = [
    "google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
    "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]


def find_chrome() -> str | None:
    env = os.environ.get("CHROME_PATH")
    if env and Path(env).exists():
        return env
    for cand in CHROME_CANDIDATES:
        found = shutil.which(cand) if "/" not in cand else (cand if Path(cand).exists() else None)
        if found:
            return found
    matches = sorted(Path("/opt/pw-browsers").glob("chromium-*/chrome-linux/chrome"))
    return str(matches[-1]) if matches else None


def render_pdf(html_path: Path, pdf_path: Path) -> bool:
    chrome = find_chrome()
    if chrome:
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        cmd = [
            chrome, "--headless", "--disable-gpu", "--no-sandbox",
            "--no-pdf-header-footer", "--run-all-compositor-stages-before-draw",
            "--virtual-time-budget=4000",
            f"--print-to-pdf={pdf_path}", html_path.as_uri(),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            return True
        print(f"  ! chrome print failed: {proc.stderr.strip()[:400]}", file=sys.stderr)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        print("  · falling back to LibreOffice (converting the .docx)")
        docx = ROOT / "assets" / f"{BASENAME}.docx"
        proc = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir",
             str(pdf_path.parent), str(docx)],
            capture_output=True, text=True,
        )
        produced = pdf_path.parent / f"{BASENAME}.pdf"
        if produced.exists():
            return True
        print(f"  ! libreoffice failed: {proc.stderr.strip()[:400]}", file=sys.stderr)

    return False


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main() -> int:
    d = load()
    BUILD.mkdir(exist_ok=True)

    site = ROOT / "resume.html"
    site.write_text(build_site_html(d), encoding="utf-8")
    print(f"  ✓ {site.relative_to(ROOT)}")

    print_html = BUILD / "resume-print.html"
    print_html.write_text(build_print_html(d), encoding="utf-8")
    print(f"  ✓ {print_html.relative_to(ROOT)}")

    docx = ROOT / "assets" / f"{BASENAME}.docx"
    build_docx(d, docx)
    print(f"  ✓ {docx.relative_to(ROOT)}  ({docx.stat().st_size // 1024} KB)")

    pdf = ROOT / "assets" / f"{BASENAME}.pdf"
    if render_pdf(print_html, pdf):
        print(f"  ✓ {pdf.relative_to(ROOT)}  ({pdf.stat().st_size // 1024} KB)")
    else:
        print("  ! PDF not generated — install Chrome/Chromium or LibreOffice,",
              "or open build/resume-print.html and print to PDF.", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
