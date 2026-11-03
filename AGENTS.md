# AGENTS.md

Instructions for coding agents working on ppt-mcp, and the rules the pull
request review checks a change against. How to open a pull request (base
branch, one topic, the template) is in [CONTRIBUTING.md](CONTRIBUTING.md);
this file is about the code.

ppt-mcp drives a live PowerPoint on two platforms: over COM on Windows
(`src/ppt_com/`) and over Apple Events on macOS (`src/ppt_mac/`, `src/backend/`).
Most of what goes wrong shows on one of them only.

## Tools

- A tool's input model (pydantic) lives in `src/ppt_com/` and is shared by
  both platforms. Each platform has its own `_name_impl` function, and the two
  take the same arguments in the same order
  (`tests/test_impl_signature_parity.py`).
- Check an argument in the input model, with a validator, so a wrong value is
  refused before PowerPoint is touched. A check made after the first write
  leaves a half-done change behind.
- A failure is answered, never swallowed: a tool returns `{"error": ...}` with
  a message that says what to do instead. Do not fall back silently to
  something close to what was asked.
- When a platform cannot do something, refuse with the shared helper
  (`backend.unsupported.refusal`), naming the alternatives, rather than doing
  nothing or something else. MACOS_PORT section 6.1 lists the tools that
  always refuse on macOS, and a test keeps that list and its count true.
- Read back what was written where PowerPoint can say, and report what did
  not take. A call that raised nothing is not evidence that it worked.

## PowerPoint

- One call at a time. Every PowerPoint call goes through the worker
  (`ppt.execute`); concurrent calls answer with errors that are not true while
  the change still lands.
- On macOS, do not trust a reference PowerPoint hands back. Count, then
  address by index; see MACOS_PORT.md.
- On macOS, what Apple Events cannot reach (charts, freeforms, groups and
  their members) goes through the clipboard package (`ppt_mac/gvml_paste.py`):
  the user's clipboard is taken first and put back after, and a paste is
  checked by reading it back before the original is deleted.

## Files and text

- Read and write text files with `encoding="utf-8"`. On a Japanese Windows the
  default is cp932, and a bare `open()` or `read_text()` fails there only.
- A change to what a tool does or takes goes into `README.md`, `README_ja.md`
  and `README_he.md`; `tests/test_readmes.py` checks that they agree.

## Tests

- `uv run --frozen pytest -q` needs no PowerPoint. Use the fakes in `tests/`
  for COM objects and Apple Events; on macOS, `tests/conftest.py` keeps every
  test off the real PowerPoint.
- A fix comes with a test that fails without it.
- A test that drives a COM implementation is marked to run on Windows only:
  on macOS the `ppt_com` modules carry the Mac implementations.
