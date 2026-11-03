"""fill_type left out of a fill, as ppt_add_shape allows (#194, item 4).

ppt_add_shape takes a solid fill from fill_color alone, while ppt_set_fill,
and ppt_batch_apply_formatting through the same model, required fill_type.
The batch is reached for after the single shape tool, so that was a surprise.
Both now take the type from the colours when it is left out.
"""

import sys

import pytest
from pydantic import ValidationError

sys.path.insert(0, "src")

from ppt_com.batch_apply import _checked  # noqa: E402
from ppt_com.formatting import SetFillInput  # noqa: E402

SHAPE = {"slide_index": 1, "shape_name_or_index": "Box"}


class TestInferred:
    def test_a_color_alone_is_a_solid_fill(self):
        assert SetFillInput(**SHAPE, color="#1F2937").fill_type == "solid"

    def test_gradient_colors_are_a_gradient(self):
        inp = SetFillInput(**SHAPE, gradient_color1="#FFFFFF", gradient_color2="#000000")
        assert inp.fill_type == "gradient"

    def test_one_gradient_color_is_still_a_gradient(self):
        assert SetFillInput(**SHAPE, gradient_color2="#000000").fill_type == "gradient"

    def test_a_type_given_is_kept(self):
        assert SetFillInput(**SHAPE, fill_type="none").fill_type == "none"
        assert SetFillInput(**SHAPE, fill_type="solid", color="#FFF").fill_type == "solid"

    def test_nothing_to_go_on_says_what_to_pass(self):
        with pytest.raises(ValidationError, match="Pass fill_type"):
            SetFillInput(**SHAPE)

    def test_the_description_says_so(self):
        assert "Defaults to 'solid'" in SetFillInput.model_fields["fill_type"].description


class TestInTheBatch:
    def test_a_batch_fill_takes_a_color_alone(self):
        m = _checked(SetFillInput, "set_fill", {"color": "#2563EB"}, **SHAPE)
        assert (m.fill_type, m.color) == ("solid", "#2563EB")

    def test_a_batch_fill_with_nothing_is_refused_before_anything_is_applied(self):
        with pytest.raises(ValidationError, match="Pass fill_type"):
            _checked(SetFillInput, "set_fill", {"transparency": 0.5}, **SHAPE)
