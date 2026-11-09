"""Create a new deck in the look of an existing one: "copy an example slide, swap the content".

Why this exists: many real corporate decks have no usable placeholders. Every title and box
is a hand-placed shape with fixed colours, all on one blank layout. Their design lives in the
example slides themselves, not in the template's layouts. So instead of "fill the slots of a
layout", we copy a good example slide and change only its words (and, for repeated items such
as cards or columns, how many there are and where they sit).

A STYLE KIT (JSON, written once per template, by a person) says which example slide is good for
what, and which shape on it holds which piece of content:

    {
      "name": "Sample Co",
      "page_format": "{n} / {total}",
      "patterns": {
        "columns": {
          "slide": 2,                                   # example slide (1-based)
          "use_for": "2-4 parallel themes, each with a heading and bullet points",
          "fields": {"title": 3, "subtitle": 4, "footer": 6, "page": 19},   # name -> shape id
          "items": {                                    # optional repeated block
            "groups": [[7, 8, 9, 10], [11, 12, 13, 14], [15, 16, 17, 18]], # shape ids per item
            "fields": {"heading": 1, "points": 3},      # name -> position inside a group
            "min": 2, "max": 4, "max_columns": 4
          }
        },
        "table": {"slide": 24, "fields": {...}, "table": {"shape": 25, "max_rows": 9}}
      }
    }

The MODEL fills a plan that names patterns and supplies text only:

    {"footer": "...", "slides": [
       {"pattern": "columns", "fields": {"title": "...", "subtitle": "..."},
        "items": [{"heading": "Today", "points": ["...", "..."]}, ...]}]}

compile_from_kit (pure, no model) turns that into ops: copy slide N, duplicate/delete item
shapes to match the item count, re-flow the items in a grid, write the text, number the pages,
then delete every original example slide.

Safety rule: any text shape on the copied slide that the plan does not fill is DELETED, so no
example content can leak into the new deck.
"""
from __future__ import annotations

import json
import math

from .compiler import CompileError

AUTO_FIELDS = {"page", "footer"}     # filled by the compiler, not the model


def load_kit(path: str) -> dict:
    kit = json.load(open(path))
    for name, p in kit["patterns"].items():
        if "slide" not in p or "fields" not in p:
            raise CompileError([f"pattern {name!r}: needs 'slide' and 'fields'"])
    return kit


def describe_kit(kit: dict) -> str:
    """What the model is told about the available slide patterns (no geometry, no colours)."""
    lines = []
    for name, p in kit["patterns"].items():
        fields = [f for f in p["fields"] if f not in AUTO_FIELDS]
        s = f"- {name}: {p.get('use_for', '')}. fields: {fields}"
        if "items" in p:
            it = p["items"]
            s += f"; items: {it.get('min', 1)}-{it.get('max', 12)} of {list(it['fields'])}"
        if "table" in p:
            s += f"; table: header + up to {p['table'].get('max_rows', 10)} rows"
        lines.append(s)
    return "\n".join(lines)


# ---- helpers over the DeckIndex ----------------------------------------------------------

def _shapes_by_id(slide: dict) -> dict[int, dict]:
    return {int(s["ref"].rsplit("shape:", 1)[1]): s for s in slide["shapes"]}


def _text_of(shape: dict) -> str:
    t = "".join(p["text"] for p in shape.get("paras", []))
    t += "".join(p["text"] for c in shape.get("cells", []) for p in c["paras"])
    return t.strip()


def _as_paragraphs(value) -> list[str]:
    if isinstance(value, list):
        return [str(v) for v in value] or [""]
    return str(value).split("\n")


def _bbox(geoms):
    x0 = min(g["x"] for g in geoms); y0 = min(g["y"] for g in geoms)
    x1 = max(g["x"] + g["w"] for g in geoms); y1 = max(g["y"] + g["h"] for g in geoms)
    return x0, y0, x1 - x0, y1 - y0


def _stretch(offset, size, margin, group_size, new_group_size):
    """Shapes that reach (nearly) to the group's far edge grow with the group; small ones keep size."""
    if size > 0.3 * group_size and margin < 0.25 * group_size:
        return max(1, int(new_group_size - offset - margin))
    return size


def layout_items(groups_geom: list[list[dict]], n: int, max_columns: int | None = None):
    """Grid for n items inside the area the example's items occupied.

    groups_geom: per example group, the geometry of each of its shapes.
    Returns, per new item i: (template_group_index, [new geom per shape]).
    """
    gboxes = [_bbox(g) for g in groups_geom]
    span = _bbox([{"x": b[0], "y": b[1], "w": b[2], "h": b[3]} for b in gboxes])
    xs = sorted({b[0] for b in gboxes}); ys = sorted({b[1] for b in gboxes})
    w0, h0 = gboxes[0][2], gboxes[0][3]
    gap_x = (xs[1] - xs[0] - w0) if len(xs) > 1 else 228600          # default 0.25in
    gap_y = (ys[1] - ys[0] - h0) if len(ys) > 1 else 182880
    # default: a single-row example stays a single row; a grid keeps its column count
    max_cols = max_columns or (len(xs) if len(ys) > 1 else n)
    rows = math.ceil(n / max_cols)
    cols = math.ceil(n / rows)
    item_w = (span[2] - gap_x * (cols - 1)) / cols
    item_h = h0 if rows == len(ys) else (span[3] - gap_y * (rows - 1)) / rows
    out = []
    for i in range(n):
        k = i % len(groups_geom)                                   # cycle example groups (keeps colour order)
        gx, gy, gw, gh = gboxes[k]
        ix = span[0] + (i % cols) * (item_w + gap_x)
        iy = span[1] + (i // cols) * (item_h + gap_y)
        new = []
        for g in groups_geom[k]:
            dx, dy = g["x"] - gx, g["y"] - gy
            w = _stretch(dx, g["w"], gx + gw - (g["x"] + g["w"]), gw, item_w)
            h = _stretch(dy, g["h"], gy + gh - (g["y"] + g["h"]), gh, item_h)
            new.append({"x": int(ix + dx), "y": int(iy + dy), "w": int(w), "h": int(h)})
        out.append((k, new))
    return out


# ---- the compiler ---------------------------------------------------------------------------

def compile_from_kit(plan: dict, index: dict, kit: dict) -> list[dict]:
    errors, ops = [], []
    total = len(plan["slides"])
    for n, s in enumerate(plan["slides"], start=1):
        where = f"slide {n} ({s.get('pattern')})"
        pat = kit["patterns"].get(s.get("pattern"))
        if pat is None:
            errors.append(f"{where}: unknown pattern; available: {sorted(kit['patterns'])}")
            continue
        ex = index["slides"][pat["slide"] - 1]
        shapes = _shapes_by_id(ex)
        steps, used = [], set()
        given = dict(s.get("fields", {}))
        unknown = set(given) - set(pat["fields"])
        if unknown:
            errors.append(f"{where}: unknown fields {sorted(unknown)}")
        auto = {"page": kit.get("page_format", "{n} / {total}").format(n=n, total=total),
                "footer": plan.get("footer")}
        for name, sid in pat["fields"].items():
            used.add(sid)
            value = auto.get(name) if name in AUTO_FIELDS else given.get(name)
            if value in (None, "", []):
                steps.append({"op": "delete_shape", "target": f"shape:{sid}"})
            else:
                steps.append({"op": "set_shape_text", "target": f"shape:{sid}", "paragraphs": _as_paragraphs(value)})

        # repeated items (cards, columns, stat tiles, panels)
        if "items" in pat:
            it = pat["items"]
            items = s.get("items", [])
            lo, hi = it.get("min", 1), it.get("max", 12)
            if not lo <= len(items) <= hi:
                errors.append(f"{where}: needs {lo}-{hi} items, got {len(items)}")
                items = items[:hi] or [{}]
            groups = it["groups"]
            used.update(sid for g in groups for sid in g)
            geoms = [[shapes[sid]["geom"] for sid in g] for g in groups]
            next_id = max(shapes) + 1
            dup_steps, del_steps, geo_steps, text_steps = [], [], [], []
            placed = layout_items(geoms, len(items), it.get("max_columns")) if items else []
            for i, (k, new_geoms) in enumerate(placed):
                if i < len(groups):
                    ids = groups[i]
                else:                                            # more items than the example has
                    ids = []
                    for sid in groups[k]:
                        dup_steps.append({"op": "duplicate_shape", "source": f"shape:{sid}", "new_id": next_id})
                        ids.append(next_id)
                        next_id += 1
                for sid, g in zip(ids, new_geoms):
                    geo_steps.append({"op": "set_geometry", "target": f"shape:{sid}", **g})
                item = dict(items[i])
                bad = set(item) - set(it["fields"])
                if bad:
                    errors.append(f"{where} item {i + 1}: unknown fields {sorted(bad)}")
                if "number" in it["fields"] and "number" not in item:
                    item["number"] = str(i + 1)
                for fname, pos in it["fields"].items():
                    sid = ids[pos]
                    v = item.get(fname)
                    if v in (None, "", []):
                        del_steps.append({"op": "delete_shape", "target": f"shape:{sid}"})
                    else:
                        text_steps.append({"op": "set_shape_text", "target": f"shape:{sid}",
                                           "paragraphs": _as_paragraphs(v)})
            for g in groups[len(items):]:                        # fewer items than the example has
                del_steps += [{"op": "delete_shape", "target": f"shape:{sid}"} for sid in g]
            steps = dup_steps + steps + geo_steps + text_steps + del_steps

        if "table" in pat:
            t, tsid = s.get("table"), pat["table"]["shape"]
            used.add(tsid)
            ncols = 1 + max(c["rc"][1] for c in shapes[tsid]["cells"])
            if not t:
                errors.append(f"{where}: needs a table")
            elif len(t["header"]) != ncols or any(len(r) != ncols for r in t["rows"]):
                errors.append(f"{where}: table must have exactly {ncols} columns")
            elif len(t["rows"]) > pat["table"].get("max_rows", 10):
                errors.append(f"{where}: at most {pat['table'].get('max_rows', 10)} table rows")
            else:
                steps.append({"op": "set_table", "target": f"shape:{tsid}", "header": t["header"], "rows": t["rows"]})

        # leak guard: any example text the plan didn't claim is removed
        for sid, shp in shapes.items():
            if sid not in used and _text_of(shp) and shp.get("kind") != "group":
                steps.append({"op": "delete_shape", "target": f"shape:{sid}"})

        ops.append({"op": "add_slide_from_pattern", "source": ex["ref"], "steps": steps, "notes": s.get("notes")})

    if errors:
        raise CompileError(errors)
    ops += [{"op": "delete_slide", "target": sl["ref"]} for sl in index["slides"]]   # drop all examples
    first = plan["slides"][0].get("fields", {}).get("title", "") if plan["slides"] else ""
    ops.append({"op": "set_doc_properties", "title": plan.get("title", first), "author": plan.get("author", "")})
    return ops


def example_texts(index: dict, kit: dict) -> set[str]:
    """Text from the example slides, used to prove none of it leaked into the new deck."""
    out = set()
    for p in kit["patterns"].values():
        for shp in index["slides"][p["slide"] - 1]["shapes"]:
            for para in shp.get("paras", []) + [q for c in shp.get("cells", []) for q in c["paras"]]:
                if len(para["text"].strip()) >= 12:
                    out.add(para["text"].strip())
    return out
