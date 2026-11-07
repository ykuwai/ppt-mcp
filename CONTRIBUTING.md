# Contributing to ppt-mcp

Thank you for helping. ppt-mcp drives a live PowerPoint on two platforms, and
most of what goes wrong only shows on one of them, so a pull request is easiest
to take in when it says plainly what it changes and where it was tried.

## Where a pull request goes

- **Base it on `develop`.** `main` only takes releases.
- **One topic per pull request.** A fix and an unrelated cleanup are two pull
  requests. Several commits are fine, one per step, each with a message that
  says what changed and why.
- **Do not stack pull requests on each other's branches.** Base each one on
  `develop`. When one needs another, say so in its description and keep it as
  a draft until the first is merged. A pull request whose base branch is
  deleted is closed by GitHub.
- **Open a draft while it is still moving**, and mark it ready when the tests
  pass and the description is complete.

## Before you open it

1. **Run the tests.** `uv run --frozen pytest -q`. They need no PowerPoint: on
   Windows the COM connection is made on the first tool call, and on macOS
   `tests/conftest.py` keeps every test off the real PowerPoint. CI runs them on
   Windows (with mcp 1.x and 2.x) and on macOS, and a pull request is merged
   only when all three pass.
2. **Add a test for what you changed.** A fix comes with a test that fails
   without it. Use the fakes already in `tests/` for COM objects and Apple
   Events rather than a live PowerPoint.
3. **Try it on a real PowerPoint, if you can,** on the platform it touches, and
   say which (version and OS) in the description. Close only the decks you
   opened.
4. **Keep the docs in step.** A change to what a tool does or takes goes into
   `TOOLS.md`, which is in English only. `README.md` is the English
   original, and `README.<language>.md` are its translations (ja, zh-CN, es,
   pt-BR, he). A change to how the server is set up or used goes into all of
   them; a coding agent can do the translations. Never type a tool count by hand: run
   `uv run python scripts/sync_readmes.py`, which writes every count from the
   code. `tests/test_readmes.py` checks the counts, the settings and the
   sections.

## Things that have bitten before

- **Read text files with `encoding="utf-8"`.** On a Japanese Windows the default
  is cp932, and a bare `read_text()` or `open()` fails there and nowhere else.
- **Call PowerPoint one call at a time.** Concurrent calls answer with errors
  that are not true while the change still lands.
- **On macOS, do not trust a reference PowerPoint hands back.** Count, then
  address by index; see `MACOS_PORT.md`.
- **A tool that cannot do something on one platform says so** with the shared
  refusal helpers, naming what to use instead, rather than doing nothing.

## The description

The pull request template asks for what changed, why, and how it was tested.
Numbers help: a measurement before and after, the test counts, the PowerPoint
version a live check ran on. A reviewer reading only the description should
know what to look at.

## Review

Every pull request gets an automatic review from Claude Code when it is
opened or marked ready, not on every push. After larger changes, ask for
another with a comment that mentions @claude. For a pull request from a fork it runs once the author has had a pull request merged
here; before that a maintainer starts it. Its findings are a starting point,
not a verdict: answer them in the thread, with a fix or a reason.
