"""Every tool wrapper must pass every argument its implementation takes.

Twice in one day a new argument was added to a model, threaded into the
implementation and then dropped by the one function in between. The call still
succeeded, the implementation used its default, and the answer reported the
default back as though it were what was asked for. `ppt_add_textbox` accepted
`zorder` and placed nothing; `ppt_set_shadow` accepted `target='text'` and set
the shape shadow.

Nothing else catches it. The signature parity test compares the two platforms
with each other, not a wrapper with its implementation, and the unit tests call
the implementations directly.

Read with ast, so it runs on Windows without importing appscript.
"""

import ast
import pathlib

import pytest

SRC = pathlib.Path(__file__).resolve().parents[1] / "src"

# A wrapper that deliberately passes fewer arguments than its implementation
# takes, with the reason. Keep this empty if you can.
ALLOWED_GAPS: dict[str, str] = {}


def _impl_arity(package):
    """{impl name: how many arguments it takes} for one package."""
    arity = {}
    for path in sorted((SRC / package).glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in tree.body:
            if (isinstance(node, ast.FunctionDef)
                    and node.name.startswith("_")
                    and node.name.endswith("_impl")):
                args = node.args
                if args.vararg or args.kwarg:
                    continue  # takes anything, nothing to pin
                arity[node.name] = len(args.posonlyargs) + len(args.args)
    return arity


def _execute_calls(package):
    """Every ppt.execute(_x_impl, ...) call, as (where, impl name, how many)."""
    calls = []
    for path in sorted((SRC / package).glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if not (isinstance(func, ast.Attribute) and func.attr == "execute"):
                continue
            if not node.args:
                continue
            first = node.args[0]
            if not (isinstance(first, ast.Name) and first.id.endswith("_impl")):
                continue
            if any(isinstance(a, ast.Starred) for a in node.args):
                continue
            calls.append((f"{path.name}:{node.lineno}", first.id,
                          len(node.args) - 1))
    return calls


ARITY = _impl_arity("ppt_com")
CALLS = _execute_calls("ppt_com")


def test_there_are_wrappers_to_check():
    # A guard on the guard: if the walk stops finding calls, everything below
    # passes by having nothing to compare.
    assert len(CALLS) > 50


@pytest.mark.parametrize("where,impl,passed", CALLS,
                         ids=[f"{w} {i}" for w, i, _ in CALLS])
def test_the_wrapper_passes_every_argument(where, impl, passed):
    expected = ARITY.get(impl)
    if expected is None:
        pytest.skip(f"{impl} is not defined in ppt_com")
    if impl in ALLOWED_GAPS:
        pytest.skip(ALLOWED_GAPS[impl])
    assert passed == expected, (
        f"{where} calls {impl} with {passed} arguments, and it takes "
        f"{expected}.\n"
        "An argument the model accepts and the wrapper drops is a silent "
        "default: the call succeeds and the answer reports something the "
        "caller did not ask for."
    )
