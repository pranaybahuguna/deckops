"""NodeRef addressing scheme.

Grammar (all ids are the deck's own stable ids, not positions, except para/run/cell):

    slide:<sldId>                                    p:sldId/@id  - stable across reorder
    slide:<sldId>/shape:<id>                         p:cNvPr/@id  - unique within a slide (groups included)
    slide:<sldId>/shape:<id>/para:<i>
    slide:<sldId>/shape:<id>/para:<i>/run:<j>
    slide:<sldId>/shape:<id>/cell:<r>,<c>/para:<i>   table cell
    slide:<sldId>/notes/para:<i>                     speaker notes

Paragraph/run indices are positional, so every op that targets one carries an
`expect` fingerprint (see `sha`). Execution binds all targets *before* mutating
anything, so ops in one list can never shift each other's indices.
"""
from __future__ import annotations

import hashlib
import html
import re
from dataclasses import dataclass

from pptx.oxml.ns import qn

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSMAP = {"a": A_NS, "p": P_NS}

_REF_RE = re.compile(
    r"^slide:(?P<slide>\d+)"
    r"(?:/(?:shape:(?P<shape>\d+)(?:/cell:(?P<row>\d+),(?P<col>\d+))?|(?P<notes>notes)))?"
    r"(?:/para:(?P<para>\d+)(?:/run:(?P<run>\d+))?)?$"
)


def sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:10]


class RefError(Exception):
    pass


@dataclass
class Ref:
    slide: int
    shape: int | None = None
    row: int | None = None
    col: int | None = None
    notes: bool = False
    para: int | None = None
    run: int | None = None

    @classmethod
    def parse(cls, s: str) -> "Ref":
        m = _REF_RE.match(s)
        if not m:
            raise RefError(f"malformed ref: {s!r}")
        g = m.groupdict()
        as_int = lambda k: int(g[k]) if g[k] is not None else None
        return cls(as_int("slide"), as_int("shape"), as_int("row"), as_int("col"),
                   bool(g["notes"]), as_int("para"), as_int("run"))

    def __str__(self) -> str:
        s = f"slide:{self.slide}"
        if self.notes:
            s += "/notes"
        elif self.shape is not None:
            s += f"/shape:{self.shape}"
            if self.row is not None:
                s += f"/cell:{self.row},{self.col}"
        if self.para is not None:
            s += f"/para:{self.para}"
            if self.run is not None:
                s += f"/run:{self.run}"
        return s


# ---- paragraph text helpers -------------------------------------------------

def para_tokens(p_el):
    """Ordered content children of an a:p: runs, line breaks, fields."""
    return [c for c in p_el if c.tag in (qn("a:r"), qn("a:br"), qn("a:fld"))]


def para_runs(p_el):
    return [c for c in p_el if c.tag == qn("a:r")]


def para_text(p_el) -> str:
    out = []
    for c in para_tokens(p_el):
        if c.tag == qn("a:br"):
            out.append("\n")
        else:
            t = c.find(qn("a:t"))
            out.append(t.text or "" if t is not None else "")
    return "".join(out)


def para_tagged(p_el) -> str:
    """Paragraph as tagged text the model can translate while keeping formatting.

    'Revenue grew <b>12%</b> YoY' with three runs becomes
    '<r0>Revenue grew </r0><r1>12%</r1><r2> YoY</r2>'.  Line breaks are <br/>,
    fields (slide numbers, dates) are opaque <f0/> tokens that must be kept.
    """
    out, ri, fi = [], 0, 0
    for c in para_tokens(p_el):
        if c.tag == qn("a:r"):
            t = c.find(qn("a:t"))
            txt = html.escape((t.text or "") if t is not None else "", quote=False)
            out.append(f"<r{ri}>{txt}</r{ri}>")
            ri += 1
        elif c.tag == qn("a:br"):
            out.append("<br/>")
        else:
            out.append(f"<f{fi}/>")
            fi += 1
    return "".join(out)


_TOKEN_RE = re.compile(r"<r(\d+)>(.*?)</r\1>|<br/>|<f(\d+)/>", re.S)


class TagError(ValueError):
    pass


def parse_tagged(value: str):
    """Inverse of para_tagged. Returns [('r', i, text) | ('br',) | ('f', i)].

    A value with no tags at all is returned as [('plain', text)].
    """
    if "<r" not in value and "<br/>" not in value and "<f" not in value:
        return [("plain", html.unescape(value))]
    tokens, pos = [], 0
    for m in _TOKEN_RE.finditer(value):
        if value[pos:m.start()].strip():
            raise TagError(f"text outside run tags: {value[pos:m.start()]!r}")
        pos = m.end()
        if m.group(1) is not None:
            tokens.append(("r", int(m.group(1)), html.unescape(m.group(2))))
        elif m.group(3) is not None:
            tokens.append(("f", int(m.group(3))))
        else:
            tokens.append(("br",))
    if value[pos:].strip():
        raise TagError(f"text outside run tags: {value[pos:]!r}")
    return tokens


def text_body_paragraphs(txBody):
    return txBody.findall(qn("a:p")) if txBody is not None else []


# ---- resolution against a live python-pptx Presentation ---------------------

def find_slide(prs, slide_id: int):
    s = prs.slides.get(slide_id)
    if s is None:
        raise RefError(f"no slide with id {slide_id}")
    return s


def find_shape_el(slide, shape_id: int):
    hits = slide.shapes._spTree.xpath(f".//p:cNvPr[@id='{shape_id}']")
    if not hits:
        raise RefError(f"slide {slide.slide_id}: no shape id {shape_id}")
    # p:cNvPr -> p:nvSpPr|p:nvGraphicFramePr|... -> shape element
    return hits[0].getparent().getparent()


def shape_txBody(shape_el, row=None, col=None):
    if row is not None:
        trs = shape_el.findall(".//" + qn("a:tr"))
        if row >= len(trs):
            raise RefError(f"table has no row {row}")
        tcs = trs[row].findall(qn("a:tc"))
        if col >= len(tcs):
            raise RefError(f"table has no col {col}")
        return tcs[col].find(qn("a:txBody"))
    return shape_el.find(qn("p:txBody"))


def resolve(prs, ref: Ref):
    """Return (kind, element, slide). kind in slide|shape|para|run."""
    slide = find_slide(prs, ref.slide)
    if ref.shape is None and not ref.notes:
        if ref.para is not None:
            raise RefError(f"{ref}: para needs a shape or notes")
        return "slide", slide._element, slide
    if ref.notes:
        if not slide.has_notes_slide:
            raise RefError(f"{ref}: slide has no notes")
        txBody = slide.notes_slide.notes_text_frame._txBody
        shape_el = None
    else:
        shape_el = find_shape_el(slide, ref.shape)
        if ref.para is None:
            return "shape", shape_el, slide
        txBody = shape_txBody(shape_el, ref.row, ref.col)
    paras = text_body_paragraphs(txBody)
    if ref.para is None or ref.para >= len(paras):
        raise RefError(f"{ref}: paragraph out of range ({len(paras)} paragraphs)")
    p = paras[ref.para]
    if ref.run is None:
        return "para", p, slide
    runs = para_runs(p)
    if ref.run >= len(runs):
        raise RefError(f"{ref}: run out of range")
    return "run", runs[ref.run], slide
