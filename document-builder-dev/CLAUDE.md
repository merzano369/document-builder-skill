# document-builder: instructions for an agent changing the skill

Goal: a skill for building Word reports that is as simple in architecture and as functional as possible. The user agent writes only content via `doc.add_*`; all formatting is set by `styles_config.json`.

Before a change, read `docs/ARCHITECTURE.md`: file map, mechanisms, invariants, python-docx and Word pitfalls.

## 1. Think before coding

- State your assumptions. If a request can be read in different ways, show the options and ask instead of choosing silently.
- If there is a simpler way (a parameter to an existing method, a config key), say so. A new public method is a last resort.
- If you do not understand something in the code, stop and find out. Do not guess Word's behaviour: check it on a built document.

## 2. Simplicity above all

- The minimum code that solves the task. Nothing "for the future": no abstractions for a single call, no flexibility nobody asked for, no handling of impossible situations.
- Do not add template engines, DSLs, class hierarchies, plugins, two-pass builds. These ideas have already been considered and rejected.
- If a solution came out at 200 lines and could be 50, rewrite it.

## 3. Surgical changes

- Change only what the task needs. Do not "improve" neighbouring code, comments or formatting, and do not refactor what works.
- Follow the existing style: comments, docstrings and console messages in English, no emoji, console prefixes `[OK]`, `[WARNING]`, `[ERROR]`.
- Keep everything that ends up in the document in Ukrainian: default labels, title page strings, placeholders, example text in `SKILL.md` and `chapters/`. Do not translate the Ukrainian patterns in `tools/check_report.py`: they match the PDF text.
- Remove only what became unnecessary because of your change. Do not touch someone else's dead code; mention it in the report.
- The public API is backward compatible: new behaviour is enabled by a parameter or a config key whose default keeps the old behaviour.

## 4. Work from a verifiable goal

- Restate the task as a check: "fix a bug" means "write a test that reproduces it and make it pass".
- For a multi-step task, first briefly describe the plan and how you will verify each step.
- A change is done when:
  1. `python tests/test_build.py` passes completely;
  2. for layout changes, `python ../document-builder-demo/run_all.py` passed and you looked through `output/compare/`;
  3. a new config key is in `DEFAULT_CONFIG`, `styles_config.json` and `reference/config.md` (checked by `test_config_sync`);
  4. `SKILL.md` or `reference/` is updated if the API changed;
  5. `__version__` in `DocumentStyles.py` and the `SKILL.md` heading are bumped.
- Report honestly: what was checked, what was not, which tests failed.

## Working environment

- Source of truth: this folder. The owner updates installed copies of the skill personally; there is no need to sync them.
- Do not use Git unless the owner explicitly asked.
- Make large edits with the Edit tool or a script in a file, not a heredoc in Bash: backslashes and quotes break there.
