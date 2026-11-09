"""Execute an OpList against a source deck.

Two properties the design asserts, enforced here rather than checked afterwards:

1. Bind-then-apply. Every target is resolved and its `expect` fingerprint checked
   against the *source* before anything is mutated. One stale fingerprint fails
   the whole batch, and nothing is written. Positional indices can't drift
   mid-batch because elements are bound up front.

2. Byte-preservation. python-pptx re-serializes every XML part it parsed, so a
   PowerPoint-authored slide nobody touched still comes out with different bytes.
   `save_preserving` compares each part canonically and writes the ORIGINAL bytes
   for any part whose content didn't change. "Untouched slide = identical bytes"
   becomes a structural property of the executor.
"""
from __future__ import annotations

import copy
import os
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement

from .ops import AddSlide, DeleteShape, DeleteSlide, OpList, SetFontSize, SetGeometry, SetParagraphText, SetRunText
from .refs import Ref, RefError, TagError, para_runs, para_text, para_tokens, parse_tagged, resolve, sha

PHASE = {SetParagraphText: 0, SetRunText: 0, SetFontSize: 1, SetGeometry: 1, DeleteShape: 2, AddSlide: 3, DeleteSlide: 4}


class ExecError(Exception):
    def __init__(self, errors):
        super().__init__("; ".join(f"op[{i}] {e}" for i, e in errors))
        self.errors = errors


# ---- bind ---------------------------------------------------------------------

def _bind(prs, op):
    if isinstance(op, (SetParagraphText, SetRunText)):
        ref = Ref.parse(op.target)
        kind, el, _ = resolve(prs, ref)
        want = "para" if isinstance(op, SetParagraphText) else "run"
        if kind != want:
            raise RefError(f"{op.op} needs a {want} ref, got {kind}")
        current = para_text(el) if want == "para" else (el.find(qn("a:t")).text or "")
        if sha(current) != op.expect.text_sha:
            raise RefError(f"expect mismatch at {op.target}: deck has changed since it was planned")
        if want == "para":
            _check_tags(el, op.value)
        return {"el": el}
    if isinstance(op, (SetFontSize, SetGeometry, DeleteShape)):
        kind, el, _ = resolve(prs, Ref.parse(op.target))
        if kind != "shape":
            raise RefError(f"{op.op} needs a shape ref")
        if isinstance(op, SetGeometry) and el.find(".//" + qn("a:xfrm")) is None:
            raise RefError(f"{op.target}: shape has no own geometry (inherits from layout)")
        return {"el": el}
    if isinstance(op, DeleteSlide):
        ref = Ref.parse(op.target)
        resolve(prs, ref)
        return {"sldId": _sldId_el(prs, ref.slide)}
    if isinstance(op, AddSlide):
        layout = _layout(prs, op.layout)
        after = _sldId_el(prs, Ref.parse(op.after).slide) if op.after else None
        idxs = {ph.placeholder_format.idx for ph in layout.placeholders}
        missing = set(op.fill) - idxs
        if missing:
            raise RefError(f"layout {op.layout} has no placeholder idx {sorted(missing)}")
        return {"layout": layout, "after": after}
    raise RefError(f"unknown op {op}")


def _check_tags(p, value):
    tokens = parse_tagged(value)
    n_runs = len(para_runs(p))
    n_flds = sum(1 for c in p if c.tag == qn("a:fld"))
    for t in tokens:
        if t[0] == "r" and t[1] >= n_runs:
            raise TagError(f"<r{t[1]}> but paragraph has {n_runs} runs")
        if t[0] == "f" and t[1] >= n_flds:
            raise TagError(f"<f{t[1]}/> but paragraph has {n_flds} fields")


def _sldId_el(prs, slide_id):
    for el in prs.slides._sldIdLst:
        if int(el.get("id")) == slide_id:
            return el
    raise RefError(f"no slide id {slide_id}")


def _layout(prs, layout_id):
    try:
        mi, li = layout_id[1:].split("L")
        return prs.slide_masters[int(mi)].slide_layouts[int(li)]
    except Exception:
        raise RefError(f"unknown layout {layout_id}")


# ---- apply --------------------------------------------------------------------

def _new_run(text):
    r = OxmlElement("a:r")
    t = OxmlElement("a:t")
    t.text = text
    r.append(t)
    return r


def _apply_paragraph(p, value, warnings, target):
    tokens = parse_tagged(value)
    runs = para_runs(p)
    flds = [c for c in p if c.tag == qn("a:fld")]
    brs = [c for c in p if c.tag == qn("a:br")]
    if tokens[0][0] == "plain":
        text = tokens[0][1]
        if len(runs) > 1:
            warnings.append(f"{target}: untagged value collapsed {len(runs)} runs into one")
        tokens = [("r", 0, text)] if runs else [("new", text)]
    for c in para_tokens(p):
        p.remove(c)
    used, new = set(), []
    for t in tokens:
        if t[0] == "r":
            el = runs[t[1]] if t[1] not in used else copy.deepcopy(runs[t[1]])
            used.add(t[1])
            el.find(qn("a:t")).text = t[2]
        elif t[0] == "new":
            el = _new_run(t[1])
        elif t[0] == "br":
            el = copy.deepcopy(brs[0]) if brs else OxmlElement("a:br")
        else:
            el = flds[t[1]]
        new.append(el)
    end = p.find(qn("a:endParaRPr"))
    for el in new:
        if end is not None:
            end.addprevious(el)
        else:
            p.append(el)


def _apply(prs, op, b, warnings):
    if isinstance(op, SetParagraphText):
        _apply_paragraph(b["el"], op.value, warnings, op.target)
    elif isinstance(op, SetRunText):
        b["el"].find(qn("a:t")).text = op.value
    elif isinstance(op, SetFontSize):
        sz = str(int(round(op.size_pt * 100)))
        for r in b["el"].iter(qn("a:r")):
            r.get_or_add_rPr().set("sz", sz)
        for e in b["el"].iter(qn("a:endParaRPr")):
            e.set("sz", sz)
    elif isinstance(op, SetGeometry):
        xfrm = b["el"].find(".//" + qn("a:xfrm"))
        xfrm.find(qn("a:off")).set("x", str(op.x)); xfrm.find(qn("a:off")).set("y", str(op.y))
        xfrm.find(qn("a:ext")).set("cx", str(op.w)); xfrm.find(qn("a:ext")).set("cy", str(op.h))
    elif isinstance(op, DeleteShape):
        b["el"].getparent().remove(b["el"])
    elif isinstance(op, DeleteSlide):
        el = b["sldId"]
        prs.part.drop_rel(el.get(qn("r:id")))
        el.getparent().remove(el)
    elif isinstance(op, AddSlide):
        slide = prs.slides.add_slide(b["layout"])
        for ph in list(slide.placeholders):
            content = op.fill.get(ph.placeholder_format.idx)
            if content is None:
                if op.drop_empty_placeholders:
                    ph._element.getparent().remove(ph._element)
                continue
            tf = ph.text_frame
            items = [content.text] if content.kind == "text" else content.items
            tf.text = items[0]
            for it in items[1:]:
                tf.add_paragraph().text = it
        if op.notes:
            slide.notes_slide.notes_text_frame.text = op.notes
        if b["after"] is not None:
            new_el = prs.slides._sldIdLst[-1]
            b["after"].addnext(new_el)


# ---- save ---------------------------------------------------------------------

_PARSER = etree.XMLParser(remove_blank_text=True, resolve_entities=False, no_network=True)


def _canon(data: bytes):
    try:
        return etree.tostring(etree.fromstring(data, _PARSER), method="c14n")
    except etree.XMLSyntaxError:
        return None


_CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
_RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _rels_meaning(name, data):
    """Relationship set with targets resolved, so '/ppt/x.xml' == '../x.xml'."""
    import posixpath
    base = posixpath.dirname(posixpath.dirname(name))          # ppt/slides/_rels/a.rels -> ppt/slides
    out = set()
    for r in etree.fromstring(data, _PARSER).iter(_RELS + "Relationship"):
        t, mode = r.get("Target"), r.get("TargetMode", "Internal")
        if mode == "Internal":
            t = posixpath.normpath(t.lstrip("/") if t.startswith("/") else posixpath.join(base, t))
        out.add((r.get("Id"), r.get("Type"), t, mode))
    return out


def _ct_meaning(data, part_names):
    """Content type each *existing* part resolves to; dangling entries don't count."""
    root = etree.fromstring(data, _PARSER)
    defaults = {e.get("Extension").lower(): e.get("ContentType") for e in root.iter(_CT + "Default")}
    overrides = {e.get("PartName"): e.get("ContentType") for e in root.iter(_CT + "Override")}
    return {n: overrides.get("/" + n) or defaults.get(n.rsplit(".", 1)[-1].lower()) for n in part_names}


def _equivalent(name, old, new, part_names):
    if old == new:
        return True
    try:
        if name == "[Content_Types].xml":
            return _ct_meaning(old, part_names) == _ct_meaning(new, part_names)
        if name.endswith(".rels"):
            return _rels_meaning(name, old) == _rels_meaning(name, new)
        if name.endswith(".xml"):
            return _canon(old) == _canon(new)
    except etree.XMLSyntaxError:
        pass
    return False


def save_preserving(prs, src: str, dst: str) -> dict:
    tmp = dst + ".tmp"
    prs.save(tmp)
    stats = {"preserved": [], "changed": [], "added": [], "removed": []}
    with zipfile.ZipFile(src) as zs, zipfile.ZipFile(tmp) as zn, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        old_names = {n for n in zs.namelist() if not n.endswith("/")}   # ignore folder entries
        new_parts = [n for n in zn.namelist() if not n.endswith("/") and n != "[Content_Types].xml"]
        for name in zn.namelist():
            new = zn.read(name)
            if name not in old_names:
                stats["added"].append(name)
                out.writestr(name, new)
                continue
            old = zs.read(name)
            same = _equivalent(name, old, new, new_parts)
            stats["preserved" if same else "changed"].append(name)
            out.writestr(name, old if same else new)
        stats["removed"] = sorted(old_names - set(zn.namelist()))
    os.remove(tmp)
    return stats


# ---- entry point ----------------------------------------------------------------

def execute(src: str, ops_json: list, dst: str) -> dict:
    ops = OpList.validate_python(ops_json)
    prs = Presentation(src)
    bound, errors = [], []
    for i, op in enumerate(ops):
        try:
            bound.append((i, op, _bind(prs, op)))
        except (RefError, TagError) as e:
            errors.append((i, str(e)))
    if errors:
        raise ExecError(errors)
    warnings = []
    for i, op, b in sorted(bound, key=lambda t: (PHASE[type(t[1])], t[0])):
        _apply(prs, op, b, warnings)
    stats = save_preserving(prs, src, dst)
    return {"ops_applied": len(ops), "warnings": warnings, "parts": {k: (v if k != "preserved" else len(v)) for k, v in stats.items()}}
