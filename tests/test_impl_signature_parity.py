"""Every _impl that exists on both platforms must take the same arguments.

`use_mac_impls` swaps implementations by name, and every public function calls
its impl positionally, so a macOS impl whose signature drifted from the Windows
one is a TypeError the first time a mac user calls that tool. Nothing else in
the suite catches it: the only arity check that existed covers batch_apply
dispatch and is macOS only, and this file has to run on Windows, where the
change is usually made.

Read with ast rather than imported, so appscript is not needed.
"""

import ast
import pathlib
import sys

sys.path.insert(0, "src")

import pytest

SRC = pathlib.Path(__file__).resolve().parents[1] / "src"


def _impls(package):
    """{impl name: (module name, [argument names])} for one package."""
    found = {}
    for path in sorted((SRC / package).glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in tree.body:
            if not isinstance(node, ast.FunctionDef):
                continue
            if not (node.name.startswith("_") and node.name.endswith("_impl")):
                continue
            args = node.args
            names = [a.arg for a in args.posonlyargs + args.args]
            if args.vararg:
                names.append("*" + args.vararg.arg)
            names += [a.arg for a in args.kwonlyargs]
            if args.kwarg:
                names.append("**" + args.kwarg.arg)
            found[node.name] = (path.name, names)
    return found


WINDOWS = _impls("ppt_com")
MAC = _impls("ppt_mac")
SHARED = sorted(set(WINDOWS) & set(MAC))


def test_the_two_packages_really_do_share_implementations():
    # A guard on the guard: if the globs stop matching, everything below
    # passes by finding nothing to compare.
    assert len(SHARED) > 100


@pytest.mark.parametrize("name", SHARED)
def test_the_arguments_are_the_same_and_in_the_same_order(name):
    win_module, win_args = WINDOWS[name]
    mac_module, mac_args = MAC[name]
    if any(arg.startswith("*") for arg in mac_args):
        # A mac impl that refuses whatever it is handed takes (*args,
        # **kwargs) on purpose. It cannot drift, so there is nothing to pin.
        pytest.skip(f"{name} accepts anything on macOS")
    assert win_args == mac_args, (
        f"{name} takes different arguments on the two platforms.\n"
        f"  ppt_com/{win_module}: {win_args}\n"
        f"  ppt_mac/{mac_module}: {mac_args}\n"
        "The call is positional and the swap is by name, so this is a "
        "TypeError on macOS."
    )
