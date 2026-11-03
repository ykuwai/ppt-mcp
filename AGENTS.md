# AGENTS.md

Notes for coding agents (Claude Code, Codex and others) working in this
repository. People are welcome to read it too. The full rules for pull
requests are in [CONTRIBUTING.md](CONTRIBUTING.md); this file is the short
version, plus what an agent needs to know about the code.

## What this is

An MCP server that drives a live Microsoft PowerPoint. On Windows it goes
through COM (pywin32), on macOS through Apple Events (appscript). Code,
comments, issues and pull requests are in English.

- `src/server.py` registers every tool and handles the MCP side.
- `src/ppt_com/` holds the tools, one module per area, each with a
  `register_tools(mcp)`. These are the Windows implementations.
- `src/ppt_mac/` holds the macOS implementations of the same tools, swapped in
  by name through `backend.use_mac_impls`. A tool a platform cannot do refuses
  with the shared refusal helpers and names what to use instead.
- `src/backend/` and `src/utils/com_wrapper.py` own the connection to
  PowerPoint. Every call goes through one worker thread (`ppt.execute`).
- `src/gvml/` rewrites the clipboard package the Mac uses for charts,
  SmartArt and groups.
- `MACOS_PORT.md` explains how the Mac side works and what it cannot do.

## Commands

```bash
uv run --frozen pytest -q        # the tests; no PowerPoint needed
uv run ppt-mcp                   # run the server over stdio
uv run mcp dev src/server.py     # the MCP inspector
```

The tests run on Windows and on macOS, and CI runs both. They never touch a
real PowerPoint: use the fakes in `tests/` for COM objects and Apple Events.

## Conventions that matter

- **Call PowerPoint one call at a time.** Concurrent calls answer with errors
  that are not true while the change still lands.
- **Read and write text files with `encoding="utf-8"`.** A Japanese Windows
  defaults to cp932.
- **COM** takes colours as BGR integers (`utils.color`), sizes in points
  (`utils.units`), collection indexes from 1, absolute paths, and positional
  arguments only, since late binding drops keywords.
- **On macOS, do not trust a reference PowerPoint hands back.** Count, then
  address by index.
- **Moving the editor view.** A tool that changes a slide calls
  `utils.navigation.goto_slide` first. It does nothing when the user set
  `PPT_FOLLOW_EDITS=off`; pass `force=True` only when the tool's job is to show
  something, as `ppt_select_shapes` does.
- **Keep the tool count low.** Every tool's schema is sent to the model on each
  connection. Prefer a new parameter on an existing tool, and add a tool only
  when folding it in would make the existing one confusing.
- **Keep the three READMEs in step** (`README.md`, `README_ja.md`,
  `README_he.md`) when a tool's behaviour, arguments or count change.
  `tests/test_readmes.py` checks them.

## Before you open a pull request

1. Base the branch on `develop`, one topic per pull request, named like
   `fix/issue-123-short-name`. Do not stack it on another pull request's branch.
2. `uv run --frozen pytest -q` passes, and a test covers what you changed (a
   fix comes with a test that fails without it).
3. If you can, try it on a real PowerPoint on the platform it touches, and say
   which version and OS. Close only the decks you opened.
4. Fill in the pull request template: what changed, why, how it was tested,
   with numbers where there are any. Write `Closes #N` for the issue it fixes.
5. Answer the automatic review in the thread, with a fix or a reason.

Do not bump the version or tag a release. Releases are cut by the maintainer
from `develop`.
