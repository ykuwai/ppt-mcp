"""ppt_add_comment answers what PowerPoint made of the comment: a modern
comment is signed by the Office user whatever author it is given, and on the
Mac a new one belongs to the slide, not to a point on it (measured
2026-11-08). The Windows implementation, which the Mac runs through VBA."""

import datetime
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ppt_com import comments as mod  # noqa: E402

# On the Mac the name holds the routed implementation; the Windows one is
# what it runs through VBA.
add = getattr(mod._add_comment_impl, "vba_route", (mod._add_comment_impl,))[0]


class Comments:
    def __init__(self, signed=("Alon", "AL"), placed=None, has_add2=True, existing=()):
        self.items = [SimpleNamespace(Text=t, Author="X", AuthorInitials="X", Left=0, Top=0) for t in existing]
        self.signed, self.placed, self.has_add2 = signed, placed, has_add2

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Add2(self, left, top, author, initials, text, provider, user):
        if not self.has_add2:
            raise AttributeError("Add2")
        self.Add(left, top, author, initials, text)

    def Add(self, left, top, author, initials, text):
        left, top = self.placed or (left, top)
        author, initials = self.signed or (author, initials)
        self.items.append(SimpleNamespace(Text=text, Author=author, AuthorInitials=initials, Left=left, Top=top))


@pytest.fixture
def slide(monkeypatch):
    holder = {}

    def use(comments):
        deck = SimpleNamespace(Slides=lambda i: SimpleNamespace(Comments=comments))
        monkeypatch.setattr(mod.ppt, "_get_app_impl", lambda *a, **k: object())
        monkeypatch.setattr(mod.ppt, "_get_pres_impl", lambda: deck)
        monkeypatch.setattr(mod, "goto_slide", lambda app, i: None)
        holder["c"] = comments
        return comments
    return use


def test_the_office_user_who_signed_it_is_answered_and_said(slide):
    slide(Comments(existing=["Earlier"]))
    out = add(1, "Check this", "Tester", "TT", 200, 150)
    assert (out["author"], out["author_initials"], out["requested_author"]) == ("Alon", "AL", "Tester")
    assert (out["left"], out["top"]) == (200, 150)
    assert "signed the comment as 'Alon'" in out["note"] and "placed" not in out["note"]


def test_a_comment_placed_elsewhere_is_said(slide):
    slide(Comments(signed=None, placed=(960.0, 0.0), has_add2=False))
    out = add(1, "Check this", "Tester", "TT", 200, 150)
    assert (out["author"], out["left"], out["top"]) == ("Tester", 960.0, 0.0)
    assert "requested_author" not in out
    assert "placed it at 960.0, 0.0, not at 200, 150" in out["note"]


def test_a_comment_made_as_asked_has_no_note(slide):
    slide(Comments(signed=None))
    out = add(1, "Check this", "Tester", "TT", 200, 150)
    assert out == {"success": True, "text": "Check this", "author": "Tester", "author_initials": "TT",
                   "left": 200, "top": 150}


def test_a_comment_powerpoint_did_not_add_is_an_error(slide):
    comments = slide(Comments())
    comments.Add = lambda *a: None  # PowerPoint answered, and added nothing
    comments.Add2 = lambda *a: None
    with pytest.raises(RuntimeError, match="did not add the comment"):
        add(1, "Check this", "Tester", "TT", 0, 0)


def test_a_comment_that_cannot_be_found_again_answers_as_before(slide):
    comments = slide(Comments())
    real_add = comments.Add

    def add_changed(left, top, author, initials, text):
        real_add(left, top, author, initials, "something else")
    comments.Add = add_changed
    comments.Add2 = lambda *a: add_changed(*a[:5])
    assert add(1, "Check this", "Tester", "TT", 0, 0) == {"success": True, "text": "Check this", "author": "Tester"}


def test_a_read_back_that_fails_does_not_fail_a_comment_that_was_added(slide):
    # The caller would retry, and leave two.
    comments = slide(Comments())
    real_add = comments.Add

    class Unreadable:
        Text = "Check this"

        @property
        def Author(self):
            raise RuntimeError("busy")

    def add_unreadable(*a):
        real_add(*a)
        comments.items[-1] = Unreadable()
    comments.Add = add_unreadable
    comments.Add2 = lambda *a: add_unreadable(*a[:5])
    out = add(1, "Check this", "Tester", "TT", 0, 0)
    assert out["success"] is True and "would not say" in out["note"]
    assert comments.Count == 1


def test_a_comment_with_a_line_break_is_found_again(slide):
    comments = slide(Comments())
    real_add = comments.Add

    def stored_with_cr(left, top, author, initials, text):
        real_add(left, top, author, initials, text.replace("\n", "\r"))
    comments.Add = stored_with_cr
    comments.Add2 = lambda *a: stored_with_cr(*a[:5])
    assert add(1, "One\nTwo", "Tester", "TT", 0, 0)["author"] == "Alon"


def test_a_listed_comment_is_dated_in_local_time_with_no_zone(slide):
    # pywin32 and the bridge both hand over local time labelled UTC (measured
    # 2026-11-09: 22:15 in Japan came back as 22:15:04+00:00).
    when = datetime.datetime(2026, 11, 9, 22, 15, 4, tzinfo=datetime.timezone.utc)
    comments = Comments(existing=["Earlier"])
    comments.items[0].DateTime = when
    slide(comments)
    listed = getattr(mod._list_comments_impl, "vba_route", (mod._list_comments_impl,))[0]
    assert listed(1)["comments"][0]["datetime"] == "2026-11-09 22:15:04"


def test_a_comment_with_no_date_lists_an_empty_one(slide):
    slide(Comments(existing=["Earlier"]))
    listed = getattr(mod._list_comments_impl, "vba_route", (mod._list_comments_impl,))[0]
    assert listed(1)["comments"][0]["datetime"] == ""
