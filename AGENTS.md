# AGENTS.md

## Cursor Cloud specific instructions

This repository is **static content only**: French-language e-learning training courses
(CHC hospital IT onboarding) built as self-contained HTML pages with inline CSS/JS.
There is **no package manager, no build step, no backend, and no database**. Nothing
needs to be installed — `python3` (already present) is enough to serve the files.

### Services

There is a single "service": a static file server used only so relative links, the
`?mode=elearning` redirect, and `Images/` backgrounds resolve cleanly.

Run it from the repo root (do not background this in the update script — start it in a
terminal/tmux session):

```bash
python3 -m http.server 8000
```

### Entry points (open in a browser after starting the server)

- IT environment course launcher: `http://localhost:8000/Environnement%20Informatique/index.html`
  (has a "Formation guidée" / "E-Learning libre" mode toggle; module self-assessment
  quizzes live at `Module N/2. Auto-évaluation.html`).
- Excel: `Excel/Module {1,2,3}/WEB_INTERACTIVE/*.html`
- Word: `Word/Module {1,2}/WEB_INTERACTIVE/*.html`
- Windows: `Windows/Module 1/WEB_INTERACTIVE/Module1_Windows{10,11}_Formation_Interactive.html`

### Notes / gotchas

- Interactivity (quizzes, slide navigation, mode toggle) is pure client-side vanilla JS;
  test it by loading a page and clicking through, not via a build/test runner.
- Paths contain spaces and non-ASCII characters (e.g. `Auto-évaluation`); URL-encode them
  when using `curl`/`http.server` (space → `%20`, `é` → `%C3%A9`).
- There is no lint/test/build tooling; "testing" means opening a course page and exercising
  the UI (e.g. answering a quiz and clicking "Valider mes réponses").
