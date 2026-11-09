import json
import os
import subprocess
import sys

import pytest
from pptx import Presentation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from deckops.compiler import CompileError  # noqa: E402
from deckops.executor import execute  # noqa: E402
from deckops.patterns import compile_from_kit, example_texts, layout_items, load_kit  # noqa: E402
from deckops.probe import iter_paragraphs, probe  # noqa: E402

sys.path.insert(0, os.path.join(ROOT, "examples"))
from create_from_style_kit import PLAN  # noqa: E402


@pytest.fixture(scope="module")
def kit_dir(tmp_path_factory):
    d = tmp_path_factory.mktemp("kit")
    subprocess.run([sys.executable, os.path.join(ROOT, "samples/make_style_template.py"), str(d)], check=True,
                   capture_output=True)
    return d


def g(x, y, w, h):
    return {"x": x, "y": y, "w": w, "h": h}


def test_single_row_example_stays_one_row():
    groups = [[g(0, 0, 100, 50)], [g(120, 0, 100, 50)], [g(240, 0, 100, 50)]]
    placed = layout_items(groups, 4)
    assert len({geoms[0]["y"] for _, geoms in placed}) == 1          # all on one row
    assert [k for k, _ in placed] == [0, 1, 2, 0]                     # example groups cycle (keeps colour order)


def test_grid_reflows_and_small_shapes_keep_size():
    # 4x2 example grid of cards: a small circle + a full-width body
    groups = [[g(x, y, 30, 30), g(x, y + 40, 200, 100)] for y in (0, 200) for x in (0, 220, 440, 660)]
    placed = layout_items(groups, 6)
    xs = sorted({geoms[0]["x"] for _, geoms in placed})
    assert len(xs) == 3                                               # 6 items -> 3 x 2
    assert all(geoms[0]["w"] == 30 for _, geoms in placed)            # circle not stretched
    assert all(geoms[1]["w"] > 200 for _, geoms in placed)            # body widened to the new column


def test_unknown_pattern_and_bad_item_count(kit_dir):
    kit = load_kit(str(kit_dir / "example_style_kit.json"))
    index = probe(str(kit_dir / "example_style.pptx"))
    with pytest.raises(CompileError):
        compile_from_kit({"slides": [{"pattern": "nope"}]}, index, kit)
    with pytest.raises(CompileError):
        compile_from_kit({"slides": [{"pattern": "columns", "items": [{"heading": "only one"}]}]}, index, kit)


def test_create_from_kit_end_to_end(kit_dir, tmp_path):
    kit = load_kit(str(kit_dir / "example_style_kit.json"))
    src = str(kit_dir / "example_style.pptx")
    index = probe(src)
    out = str(tmp_path / "new.pptx")
    execute(src, compile_from_kit(PLAN, index, kit), out)
    prs = Presentation(out)
    assert len(prs.slides) == len(PLAN["slides"])
    texts = {p["text"].strip() for p in iter_paragraphs(probe(out))}
    assert not (example_texts(index, kit) & texts)                   # no example text leaked
    assert "Support assistant" in texts and "4 / 7" in texts         # content + automatic page numbers
    assert prs.core_properties.title == "Support assistant"           # file properties reset
