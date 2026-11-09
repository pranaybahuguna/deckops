"""Build an example "house style" deck + its style kit, for demos and tests.

It is deliberately built the way many real corporate decks are: no placeholders, every
title and box a hand-placed shape with fixed colours, one blank layout. Seven example
slides, one per pattern: cover, columns, steps, stats, table, panels, closing.

    python3 samples/make_style_template.py out/
    -> out/example_style.pptx, out/example_style_kit.json

To use YOUR template instead, write a kit for it by hand (see README, "Style kits"),
using `python3 tools/inspect_deck.py your_template.pptx` to find the shape ids.
"""
import json
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

OUT = sys.argv[1] if len(sys.argv) > 1 else "."
INK, MUTED, BODY = "1F2937", "6B7280", "374151"
DARK, ACCENT, LIGHT = "1E1B4B", "6366F1", "F3F4F6"
GROUP_COLOURS = ["4F46E5", "D97706", "059669"]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]


def rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid()
    r.fill.fore_color.rgb = RGBColor.from_string(fill)
    r.line.fill.background()
    return r


def text(s, x, y, w, h, value, size, colour, bold=False, italic=False, align=None, bullets=False, middle=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    if middle:
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = value if isinstance(value, list) else [value]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size, r.font.bold, r.font.italic = Pt(size), bold, italic
        r.font.color.rgb = RGBColor.from_string(colour)
        if align:
            p.alignment = align
        if bullets:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "228600")
            pPr.set("indent", "-228600")
            spc = etree.SubElement(pPr, qn("a:spcAft"))
            etree.SubElement(spc, qn("a:spcPts")).set("val", "600")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
    return tb


def dark_background(s):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor.from_string(DARK)


def chrome(s, title, subtitle):
    """Header bar, title, subtitle, footer rule, footer text, page number. Returns their ids."""
    rect(s, 0.5, 0.42, 0.1, 0.7, ACCENT)
    t = text(s, 0.75, 0.38, 12, 0.58, title, 26, INK, bold=True)
    st = text(s, 0.75, 0.97, 12, 0.32, subtitle, 12, MUTED, italic=True)
    rect(s, 0.5, 7.08, 12.33, 0.015, "E5E7EB")
    ft = text(s, 0.5, 7.13, 8, 0.25, "Example footer text for the deck", 9, MUTED)
    pg = text(s, 11.8, 7.13, 1.03, 0.25, "1 / 7", 9, MUTED, align=PP_ALIGN.RIGHT)
    return {"title": t.shape_id, "subtitle": st.shape_id, "footer": ft.shape_id, "page": pg.shape_id}


kit = {"name": "Example house style", "template": "example_style.pptx", "page_format": "{n} / {total}", "patterns": {}}

# 1 cover
s = prs.slides.add_slide(BLANK); dark_background(s)
rect(s, 0.7, 1.4, 0.08, 4.6, ACCENT)
f = {"title": text(s, 1.1, 1.3, 11, 1.1, "Example cover title", 60, "FFFFFF", bold=True).shape_id,
     "subtitle": text(s, 1.1, 2.5, 11, 0.65, "Example cover subtitle goes here", 26, "C7D2FE").shape_id,
     "tagline": text(s, 1.1, 3.3, 11, 0.45, "An example tagline sentence for the cover", 15, "A5B4FC", italic=True).shape_id}
stats = []
for i in range(4):
    v = text(s, 1.1 + 3 * i, 5.5, 2.7, 0.7, "42", 36, "A5B4FC", bold=True)
    lb = text(s, 1.1 + 3 * i, 6.2, 2.7, 0.35, "Example stat label", 11, "C7D2FE")
    stats.append([v.shape_id, lb.shape_id])
f["footer"] = text(s, 0.5, 7.0, 12.3, 0.3, "Example cover footer", 10, "C7D2FE", align=PP_ALIGN.RIGHT).shape_id
kit["patterns"]["cover"] = {"slide": 1, "use_for": "opening slide: title, subtitle, tagline, up to 4 headline numbers",
                            "fields": f, "items": {"groups": stats, "fields": {"value": 0, "label": 1}, "min": 0, "max": 4}}

# 2 columns
s = prs.slides.add_slide(BLANK)
f = chrome(s, "Example columns slide", "Example subtitle for the columns slide")
groups = []
for i in range(3):
    x = 0.5 + 4.2 * i
    hr = rect(s, x, 1.5, 4.0, 0.55, GROUP_COLOURS[i])
    ht = text(s, x, 1.5, 4.0, 0.55, f"Example heading {i + 1}", 18, "FFFFFF", bold=True, align=PP_ALIGN.CENTER, middle=True)
    br = rect(s, x, 2.05, 4.0, 4.85, LIGHT)
    bt = text(s, x + 0.2, 2.2, 3.6, 4.6, ["Example point one for this column", "Example point two for this column"], 13, BODY, bullets=True)
    groups.append([hr.shape_id, ht.shape_id, br.shape_id, bt.shape_id])
kit["patterns"]["columns"] = {"slide": 2, "use_for": "2-4 parallel themes side by side, each a heading plus bullet points",
                              "fields": f, "items": {"groups": groups, "fields": {"heading": 1, "points": 3},
                                                     "min": 2, "max": 4, "max_columns": 4}}

# 3 steps (numbered cards, 4 x 2)
s = prs.slides.add_slide(BLANK)
f = chrome(s, "Example process slide", "Example subtitle for the process slide")
groups = []
for i in range(8):
    x, y = 0.5 + 3.1 * (i % 4), 1.55 if i < 4 else 4.2
    c = rect(s, x, y, 0.55, 0.55, ACCENT, MSO_SHAPE.OVAL)
    n = text(s, x, y, 0.55, 0.55, str(i + 1), 16, "FFFFFF", bold=True, align=PP_ALIGN.CENTER, middle=True)
    h = text(s, x + 0.7, y, 2.25, 0.55, "Example step", 16, INK, bold=True, middle=True)
    b = text(s, x, y + 0.7, 2.95, 1.7, "Example description of what happens in this step.", 12, BODY)
    groups.append([c.shape_id, n.shape_id, h.shape_id, b.shape_id])
kit["patterns"]["steps"] = {"slide": 3, "use_for": "a numbered process of 3-8 steps, each a short heading and one or two sentences",
                            "fields": f, "items": {"groups": groups, "fields": {"number": 1, "heading": 2, "body": 3},
                                                   "min": 3, "max": 8, "max_columns": 4}}

# 4 stats
s = prs.slides.add_slide(BLANK)
f = chrome(s, "Example numbers slide", "Example subtitle for the numbers slide")
f["intro"] = text(s, 0.6, 1.5, 12.1, 0.9, "An example introductory sentence that frames the numbers below.", 14, BODY).shape_id
rect(s, 0.6, 2.6, 12.1, 4.0, LIGHT)
f["panel_title"] = text(s, 0.8, 2.75, 11.7, 0.4, "Example panel title", 14, INK, bold=True).shape_id
groups = []
for i in range(3):
    x = 1.0 + 3.9 * i
    tile = rect(s, x, 3.35, 3.5, 2.9, "FFFFFF")
    lb = text(s, x, 3.5, 3.5, 0.4, "Example label", 12, INK, bold=True, align=PP_ALIGN.CENTER)
    v = text(s, x, 4.05, 3.5, 1.1, "99", 48, GROUP_COLOURS[i], bold=True, align=PP_ALIGN.CENTER, middle=True)
    cp = text(s, x, 5.3, 3.5, 0.4, "example caption", 11, MUTED, align=PP_ALIGN.CENTER)
    groups.append([tile.shape_id, lb.shape_id, v.shape_id, cp.shape_id])
f["takeaway"] = text(s, 0.6, 6.7, 12.1, 0.3, "Example takeaway line for the numbers", 12, ACCENT, bold=True, italic=True).shape_id
kit["patterns"]["stats"] = {"slide": 4, "use_for": "2-4 headline numbers with label and caption, an intro sentence and a takeaway",
                            "fields": f, "items": {"groups": groups, "fields": {"label": 1, "value": 2, "caption": 3},
                                                   "min": 2, "max": 4}}

# 5 table
s = prs.slides.add_slide(BLANK)
f = chrome(s, "Example table slide", "Example subtitle for the table slide")
gf = s.shapes.add_table(3, 2, Inches(0.6), Inches(1.55), Inches(12.1), Inches(1.5))
tbl = gf.table
tbl.columns[0].width, tbl.columns[1].width = Inches(4.5), Inches(7.6)
for r, row in enumerate([["Example column A", "Example column B"], ["Example row label", "Example row text"],
                         ["Example row label", "Example row text"]]):
    for c, v in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = v
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold = Pt(12), (r == 0 or c == 0)
        run.font.color.rgb = RGBColor.from_string("FFFFFF" if r == 0 else INK)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(DARK if r == 0 else ("FFFFFF" if r % 2 else LIGHT))
kit["patterns"]["table"] = {"slide": 5, "use_for": "a two-column table, e.g. item and description, up to 9 rows",
                            "fields": f, "table": {"shape": gf.shape_id, "max_rows": 9}}

# 6 panels
s = prs.slides.add_slide(BLANK)
f = chrome(s, "Example summary slide", "Example subtitle for the summary slide")
groups = []
for i in range(3):
    y = 1.55 + 1.65 * i
    pr = rect(s, 0.5, y, 12.33, 1.55, LIGHT)
    lb = text(s, 0.7, y + 0.1, 3, 0.38, "Example label", 14, GROUP_COLOURS[i], bold=True)
    tx = text(s, 0.7, y + 0.5, 11.9, 0.95, "Example paragraph for this panel.", 13, BODY)
    groups.append([pr.shape_id, lb.shape_id, tx.shape_id])
f["takeaway"] = text(s, 0.5, 6.6, 12.33, 0.35, "Example closing line for the summary", 12, INK, bold=True, italic=True).shape_id
kit["patterns"]["panels"] = {"slide": 6, "use_for": "2-4 stacked labelled paragraphs (e.g. problem / change / outcome) and a closing line",
                             "fields": f, "items": {"groups": groups, "fields": {"label": 1, "text": 2},
                                                    "min": 1, "max": 4, "max_columns": 1}}

# 7 closing
s = prs.slides.add_slide(BLANK); dark_background(s)
rect(s, 0.7, 2.0, 0.08, 3.5, ACCENT)
f = {"title": text(s, 1.0, 1.9, 12, 1.1, "Example closing", 54, "FFFFFF", bold=True).shape_id,
     "subtitle": text(s, 1.0, 3.1, 12, 0.65, "Example closing subtitle", 22, "C7D2FE", italic=True).shape_id,
     "note": text(s, 1.0, 4.3, 12, 0.5, "Example next step sentence", 14, "A5B4FC").shape_id,
     "page": text(s, 11.8, 7.0, 1.03, 0.3, "7 / 7", 10, "A5B4FC", align=PP_ALIGN.RIGHT).shape_id}
kit["patterns"]["closing"] = {"slide": 7, "use_for": "final slide: a big line, a subtitle and a next step", "fields": f}

prs.save(f"{OUT}/example_style.pptx")
json.dump(kit, open(f"{OUT}/example_style_kit.json", "w"), indent=1)
print("wrote example_style.pptx, example_style_kit.json")
