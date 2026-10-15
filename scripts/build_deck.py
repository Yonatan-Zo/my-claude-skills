"""Build the green roof talk, version 4: click-by-click animations and slide numbers.

Run from the project folder:
    python scripts/prep_images.py
    python scripts/build_deck.py
Output: output/green_roof_talk_v4.pptx   (earlier versions in output/ are never touched)
"""
import itertools
import json
import os
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
BUILD = Path(os.environ.get("TALK_BUILD", Path(os.environ["TEMP"]) / "green_roof_talk_build"))
IMG = BUILD / "img"
OUT = ROOT / "output" / "green_roof_talk_v4.pptx"
CREDITS = json.loads((IMG / "credits.json").read_text(encoding="utf-8"))

# ---------- design tokens ----------
HEAD, BODY = "Segoe UI Semibold", "Segoe UI"
INK, MUTED = "1E2B33", "5E6E76"
GREEN, LEAF, MINT = "2F7D5B", "8CC5A3", "EAF5EE"
BLUE, SKY, WATER = "2F7FB5", "E6F1F8", "8EC5E8"
SUN, SAND = "E0912A", "FCF3E3"
WHITE, LINE, SOFT = "FFFFFF", "D5E3DA", "F4F8F5"
ROSE = "FBEDEC"
SOIL, DECK, MEMBRANE = "C99A6E", "C9CDD0", "39444C"
TREAT = {"wet": "3A76B5", "medium": "4E9A6C", "min": "E8902F", "dry": "D9534F"}
SW, SH, MX = 13.333, 7.5, 0.6

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
TITLE_ONLY = prs.slide_layouts[5]


def rgb(h):
    return RGBColor.from_string(h)


def credit(key):
    c = CREDITS[key]
    return f"{c['artist']} ({c['license']})"


# ---------- drawing helpers (c = slide or group) ----------
def _style_run(run, size, color, bold, font, italic=False):
    run.font.size = Pt(size)
    run.font.color.rgb = rgb(color)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font


def _bullet(p, indent=0.26, color=GREEN):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(indent))))
    pPr.set("indent", str(-Emu(Inches(indent))))
    buClr = pPr.makeelement(qn("a:buClr"), {})
    buClr.append(buClr.makeelement(qn("a:srgbClr"), {"val": color}))
    pPr.append(buClr)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def text(c, x, y, w, h, paras, size=16, color=INK, bold=False, font=BODY, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, space_after=6, name=None):
    box = c.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(paras, str):
        paras = [paras]
    for i, spec in enumerate(paras):
        if isinstance(spec, str):
            spec = {"text": spec}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", align)
        p.space_after = Pt(spec.get("space_after", space_after))
        for rtext, ropt in spec.get("runs") or [(spec["text"], {})]:
            r = p.add_run()
            r.text = rtext
            _style_run(r, ropt.get("size", spec.get("size", size)), ropt.get("color", spec.get("color", color)),
                       ropt.get("bold", spec.get("bold", bold)), ropt.get("font", spec.get("font", font)),
                       ropt.get("italic", spec.get("italic", False)))
        if spec.get("bullet"):
            _bullet(p, color=spec.get("bullet_color", GREEN))
    return box


def rect(c, x, y, w, h, fill, radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line=None, alpha=None, name=None,
         rotation=0):
    s = c.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            clr = s._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 100000))}))
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if rotation:
        s.rotation = rotation
    return s


def label_in(s, txt, size=12, color=WHITE, font=HEAD, align=PP_ALIGN.CENTER):
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = txt
    _style_run(r, size, color, False, font)
    return s


def badge(c, x, y, d, txt, fill=GREEN, size=16):
    return label_in(rect(c, x, y, d, d, fill, shape=MSO_SHAPE.OVAL), txt, size=size)


def pill(c, x, y, w, h, txt, fill, color=WHITE, size=11):
    return label_in(rect(c, x, y, w, h, fill, radius=0.5), txt, size=size, color=color)


def line(c, x1, y1, x2, y2, color=MUTED, width=1.25):
    ln = c.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


def picture(c, path, x, y, w, h, rounded=True, cover=True, oval=False, name=None, focus_y=0.5):
    iw, ih = Image.open(path).size
    if cover:
        pic = c.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
        box_r, img_r = w / h, iw / ih
        if img_r > box_r:
            pic.crop_left = pic.crop_right = (1 - box_r / img_r) / 2
        else:
            excess = 1 - img_r / box_r
            pic.crop_top, pic.crop_bottom = excess * focus_y, excess * (1 - focus_y)
    else:
        sc = min(w / iw, h / ih)
        pic = c.shapes.add_picture(str(path), Inches(x + (w - iw * sc) / 2), Inches(y + (h - ih * sc) / 2),
                                   Inches(iw * sc), Inches(ih * sc))
    if oval:
        pic.auto_shape_type = MSO_SHAPE.OVAL
    elif rounded:
        pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = geom.makeelement(qn("a:avLst"), {})
            geom.append(av)
        for gd in list(av):
            av.remove(gd)
        av.append(av.makeelement(qn("a:gd"), {"name": "adj", "fmla": "val 5000"}))
    if name:
        pic.name = name
    return pic


def group(slide, name=None):
    g = slide.shapes.add_group_shape()
    if name:
        g.name = name
    return g


LOGO_W = 1.25


def logo(c, w=LOGO_W, x=None, y=None):
    lw, lh = Image.open(IMG / "technion_logo.png").size
    h = w * lh / lw
    x = SW - 0.5 - 0.55 - w if x is None else x
    y = SH - 0.36 - h if y is None else y
    p = c.shapes.add_picture(str(IMG / "technion_logo.png"), Inches(x), Inches(y), Inches(w), Inches(h))
    p.name = "Technion logo"
    return p


def slide_number(s):
    """Real slide-number field (updates if slides are reordered)."""
    n = len(prs.slides._sldIdLst)
    box = s.shapes.add_textbox(Inches(SW - 0.5 - 0.42), Inches(SH - 0.62), Inches(0.42), Inches(0.3))
    box.name = "Slide number"
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    p._p.append(parse_xml(
        f'<a:fld {nsdecls("a")} id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
        f'<a:rPr lang="en-US" sz="1200" dirty="0"><a:solidFill><a:srgbClr val="{MUTED}"/></a:solidFill>'
        f'<a:latin typeface="{HEAD}"/></a:rPr><a:t>{n}</a:t></a:fld>'))


def base(section, title, source=None, notes=""):
    s = prs.slides.add_slide(TITLE_ONLY)
    t = s.shapes.title
    t.left, t.top, t.width, t.height = Inches(MX), Inches(0.62), Inches(SW - 2 * MX), Inches(0.9)
    tf = t.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.text = title
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    _style_run(tf.paragraphs[0].runs[0], 32, INK, False, HEAD)
    if section:
        text(s, MX, 0.3, 8, 0.3, section.upper(), size=11, color=GREEN, bold=True, font=HEAD, name="Section label")
    if source:
        text(s, MX, 6.9, 10.1, 0.45, source, size=9.5, color=MUTED, anchor=MSO_ANCHOR.BOTTOM, name="Source")
    logo(s)
    slide_number(s)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def stat_card(c, x, y, w, h, big, label, sub=None, fill=MINT, big_color=GREEN, big_size=30, label_size=14):
    rect(c, x, y, w, h, fill)
    paras = [{"text": big, "size": big_size, "color": big_color, "font": HEAD, "space_after": 2},
             {"text": label, "size": label_size, "color": INK, "space_after": 4}]
    if sub:
        paras.append({"text": sub, "size": 11, "color": MUTED})
    text(c, x + 0.25, y + 0.15, w - 0.45, h - 0.25, paras, anchor=MSO_ANCHOR.MIDDLE)


# ---------- animation ----------
def animate(slide, steps, fade_ms=500):
    """Each step = list of shapes fading in together on one click. Also adds a soft fade transition."""
    sld = slide._element
    for tag in ("p:transition", "p:timing"):
        for el in sld.findall(qn(tag)):
            sld.remove(el)
    ids = itertools.count(3)
    pars, bld = [], []
    for step in steps:
        a, b = next(ids), next(ids)
        inner = []
        for k, shp in enumerate(step):
            spid = shp.shape_id
            c1, c2, c3 = next(ids), next(ids), next(ids)
            node = "clickEffect" if k == 0 else "withEffect"
            inner.append(
                f'<p:par><p:cTn id="{c1}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" '
                f'grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{c2}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                f'</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst>'
                f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to>'
                f'<p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade">'
                f'<p:cBhvr><p:cTn id="{c3}" dur="{fade_ms}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                f'</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
            if shp._element.tag == qn("p:sp"):
                bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        pars.append(
            f'<p:par><p:cTn id="{a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    timing = (
        f'<p:timing {nsdecls("p")}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
        f'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" '
        f'nodeType="mainSeq"><p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn><p:prevCondLst>'
        f'<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst>'
        f'<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
        f'</p:childTnLst></p:cTn></p:par></p:tnLst>'
        + (f'<p:bldLst>{"".join(bld)}</p:bldLst>' if bld else "") + '</p:timing>')
    transition = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    anchor.addnext(transition)
    if steps:
        transition.addnext(parse_xml(timing))


# =====================================================================
# 1  Title  (v1 design: text left, full-height roof photo right)
# =====================================================================
s = prs.slides.add_slide(TITLE_ONLY)
picture(s, IMG / "title_roof.jpg", 6.9, 0, SW - 6.9, SH, rounded=False, name="Roof photo")
t = s.shapes.title
t.left, t.top, t.width, t.height = Inches(MX), Inches(1.7), Inches(5.9), Inches(2.4)
tf = t.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.BOTTOM
tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
tf.text = "Can a green roof be monitored from space?"
_style_run(tf.paragraphs[0].runs[0], 40, INK, False, HEAD)
tf.paragraphs[0].alignment = PP_ALIGN.LEFT
text(s, MX, 1.2, 5.9, 0.35, "GREEN INFRASTRUCTURE  ·  WATER  ·  REMOTE SENSING", size=11, color=GREEN, bold=True,
     font=HEAD)
text(s, MX, 4.3, 5.9, 0.9, "A water-efficient blue-green roof in a Mediterranean city", size=20, color=GREEN)
text(s, MX, 5.35, 5.9, 0.8, [{"text": "Yonatan Zohar, Fadi Kizel and Eran Friedler", "size": 15, "color": INK},
                             {"text": "Technion – Israel Institute of Technology", "size": 13, "color": MUTED}])
logo(s, w=2.1, x=MX, y=6.4)
s.notes_slide.notes_text_frame.text = (
    "Who I am. The promise of the talk: why cities need green roofs, why we do not see them here yet, and what our "
    "test roof and a free satellite told us this summer.")
animate(s, [])

# =====================================================================
# 2  Urban challenges: two drivers, many examples
# =====================================================================
s = base("Urban challenges", "Cities under pressure",
         "Urbanization: UN DESA, World Urbanization Prospects 2018. Climate change: IPCC AR6 WGI (2021). Heat: US EPA (urban "
         "areas 1–7 °F warmer by day). Valencia: AEMET via Moody's (491 mm in 8 h, Chiva); Fox Weather (200+ deaths). "
         "Photo: Valencia, October 2024 (provided by Y. Zohar).",
         "Two forces push on every city. Urbanization: 55% of people live in cities today, 68% by 2050. Climate change: "
         "hot extremes and heavy rain are becoming more frequent and intense. Together they show up as many problems at "
         "once. Heat islands, more energy for air-conditioning, air pollution. Floods, like Valencia in October 2024: up to "
         "491 mm of rain in 8 hours and more than 200 deaths. Drought and water scarcity, polluted runoff, and loss of "
         "nature.")
drivers = [("1", "Urbanization", "55% of people live in cities today, 68% by 2050", INK),
           ("2", "Climate change", "More frequent and intense heat extremes and heavy rain", TREAT["dry"])]
dg = []
for i, (num, head, sub, col) in enumerate(drivers):
    g = group(s, head)
    x = MX + i * 3.2
    rect(g, x, 1.7, 3.05, 1.75, SOFT)
    badge(g, x + 0.22, 1.9, 0.5, num, fill=col, size=14)
    text(g, x + 0.22, 2.5, 2.65, 0.95, [{"text": head, "size": 19, "font": HEAD, "space_after": 2},
                                       {"text": sub, "size": 12.5, "color": MUTED}])
    dg.append(g)
lead = text(s, MX, 3.6, 6.25, 0.35, "… which show up as:", size=13, color=MUTED, name="Lead-in")
examples = [("Heat islands", SUN), ("Energy for cooling", SUN), ("Air pollution", "8A6D57"),
            ("Floods", BLUE), ("Drought & water scarcity", BLUE), ("Polluted runoff", BLUE), ("Loss of nature", GREEN)]
chips = {}
for i, (name, col) in enumerate(examples):
    x, y = MX + (i % 2) * 3.2, 4.0 + (i // 2) * 0.66
    chips[name] = pill(s, x, y, 3.05, 0.54, name, col, size=13)
flood = group(s, "Valencia flood")
picture(flood, IMG / "valencia_flood_street.png", 7.1, 1.7, 5.63, 3.75)
text(flood, 7.1, 5.6, 5.63, 1.0, [{"runs": [("Valencia, October 2024  ", {"font": HEAD}),
                                            ("up to 491 mm of rain in 8 hours, more than 200 deaths",
                                             {"color": MUTED})]}], size=14)
animate(s, [[dg[0]], [dg[1]], [lead, chips["Heat islands"], chips["Energy for cooling"], chips["Air pollution"]],
            [chips["Floods"], chips["Drought & water scarcity"], chips["Polluted runoff"], flood],
            [chips["Loss of nature"]]])

# =====================================================================
# 3  Nature-based solutions: green infrastructure people build
# =====================================================================
CITY = dict(X=5.35, Y=1.75, W=7.4, H=4.4)
BUILDINGS = [(0.3, 1.5, 2.6), (2.0, 1.2, 1.4), (3.4, 1.8, 3.1), (5.4, 1.3, 1.9), (6.9, 1.7, 2.9)]
B_COL = ["D5DCE1", "C4CDD4", "DCE2E6", "CBD3D9", "D2D9DE"]
GROUND = 5.2


def cx(v):
    return CITY["X"] + v * CITY["W"] / 10


def cy(v):
    return CITY["Y"] + v * CITY["H"] / 6


def cw(v):
    return v * CITY["W"] / 10


def ch_(v):
    return v * CITY["H"] / 6


def city_base(slide):
    g = group(slide, "City")
    rect(g, cx(0), cy(GROUND), cw(10), ch_(0.8), "C5CDD3", radius=0.04)
    for k in range(4):
        rect(g, cx(0.4 + k * 1.1), cy(GROUND + 0.5), cw(0.5), ch_(0.05), WHITE, shape=MSO_SHAPE.RECTANGLE)
    for (x, w, top), col in zip(BUILDINGS, B_COL):
        rect(g, cx(x), cy(top), cw(w), ch_(GROUND - top), col, shape=MSO_SHAPE.RECTANGLE)
        ncol = max(2, int((w - 0.2) / 0.38))
        gap = (w - ncol * 0.18) / (ncol + 1)
        yy = top + 0.35
        while yy < GROUND - 0.45:
            for k in range(ncol):
                rect(g, cx(x + gap + k * (0.18 + gap)), cy(yy), cw(0.18), ch_(0.22), "EEF2F5",
                     shape=MSO_SHAPE.RECTANGLE)
            yy += 0.5
    return g


def tag(c, x, y, txt, fill, w=1.6, h=0.42):
    return pill(c, cx(x), cy(y), cw(w), ch_(h), txt, fill, size=11)


gi = [("roof_lush.jpg", "Green roofs", "our roof, Technion", None),
      ("wall_quadronno.jpg", "Green walls", "Milan", "wall_quadronno"),
      ("nbs_raingarden.jpg", "Rain gardens", "planted runoff basin", "nbs_raingarden"),
      ("gi_bioswale.jpg", "Bioswales", "planted kerb strip", "gi_bioswale"),
      ("gi_tram.jpg", "Green tram tracks", "Germany", "gi_tram"),
      ("gi_watersquare.jpg", "Water squares", "Rotterdam", "gi_watersquare")]
s = base("Nature-based solutions", "Green infrastructure: let nature do the work",
         "EEA (2024): 91% of 19,000 local climate action plans include nature-based solutions. Photos, Wikimedia Commons: "
         + "; ".join(credit(k) for *_, k in gi if k) + ".",
         "Nature-based solutions in the city are mostly green infrastructure that people design and build. Click by "
         "click: green roofs, like ours. Green walls on facades. Rain gardens, planted basins that collect runoff. "
         "Bioswales along the kerb. Green tram tracks, grass instead of concrete between the rails. And water squares, "
         "like Benthemplein in Rotterdam, a sunken plaza designed to fill with storm water and empty slowly. Finally: 91% "
         "of local climate plans in Europe already include nature-based solutions.")
text(s, MX, 1.38, SW - 2 * MX, 0.35, "Built features that cool the city, hold and clean storm water, and bring nature "
     "back", size=14, color=MUTED)
city_base(s)
photos = []
for i, (f, head, where, _) in enumerate(gi):
    g = group(s, head + " photo")
    x, y = MX + (i % 2) * 2.27, 1.95 + (i // 2) * 1.57
    picture(g, IMG / f, x, y, 2.17, 1.2)
    text(g, x, y + 1.24, 2.17, 0.3, [{"runs": [(head, {"font": HEAD}), (" · " + where, {"color": MUTED})]}], size=9)
    photos.append(g)
els = []
g = group(s, "Green roofs")
for i in (0, 2, 4):
    x, w, top = BUILDINGS[i]
    rect(g, cx(x - 0.05), cy(top - 0.22), cw(w + 0.1), ch_(0.24), LEAF, radius=0.3)
tag(g, 0.3, 1.7, "Green roofs", GREEN, w=1.5)
els.append(g)
g = group(s, "Green walls")
x, w, top = BUILDINGS[1]
rect(g, cx(x + w - 0.28), cy(top + 0.25), cw(0.28), ch_(GROUND - top - 0.25), GREEN, radius=0.2)
tag(g, 1.85, 0.75, "Green walls", GREEN, w=1.5)
els.append(g)
g = group(s, "Rain garden")
rect(g, cx(0.4), cy(GROUND + 0.08), cw(1.3), ch_(0.36), "7FB7A0", shape=MSO_SHAPE.OVAL)
tag(g, 0.0, 6.1, "Rain gardens", BLUE, w=1.75)
els.append(g)
g = group(s, "Bioswale")
rect(g, cx(2.0), cy(GROUND + 0.06), cw(2.6), ch_(0.18), LEAF, radius=0.5)
tag(g, 2.2, 6.1, "Bioswales", GREEN, w=1.6)
els.append(g)
g = group(s, "Green tram tracks")
rect(g, cx(4.9), cy(GROUND + 0.38), cw(3.4), ch_(0.3), LEAF, radius=0.2)
for yy in (GROUND + 0.44, GROUND + 0.6):
    rect(g, cx(4.9), cy(yy), cw(3.4), ch_(0.04), "6B7780", shape=MSO_SHAPE.RECTANGLE)
tag(g, 5.0, 6.1, "Green tram tracks", GREEN, w=2.1)
els.append(g)
g = group(s, "Water square")
rect(g, cx(8.65), cy(GROUND - 0.6), cw(1.3), ch_(0.6), "B7C3CB", radius=0.05)
rect(g, cx(8.8), cy(GROUND - 0.42), cw(1.0), ch_(0.42), WATER, radius=0.05)
tag(g, 7.85, 6.1, "Water squares", BLUE, w=1.9)
els.append(g)
stat = group(s, "91 percent")
rect(stat, 8.95, 1.8, 3.78, 1.0, SOFT)
text(stat, 9.15, 1.8, 3.45, 1.0, [{"runs": [("91%  ", {"size": 28, "font": HEAD, "color": GREEN}),
                                            ("of 19,000 European local climate plans include nature-based solutions",
                                             {"size": 11.5})]}], anchor=MSO_ANCHOR.MIDDLE)
animate(s, [[els[i], photos[i]] for i in range(6)] + [[stat]])

# =====================================================================
# 4  Benefits (your illustration as the base)
# =====================================================================
s = base("Green roofs", "One roof, many benefits",
         "Runoff: Zheng et al. (2021), Resources, Conservation & Recycling 170, 105577 (2,375 rain events, 75 studies, 21 "
         "countries). Other benefits: Berardi et al. (2014), Applied Energy 115, 411–428. Illustration: Zohar.",
         "One roof, six jobs. Start with water: on average a green roof keeps 62% of the rain of each storm, across 75 "
         "studies in 21 countries. Then heat: plants and moist soil cool the roof and the street. The roof insulates the "
         "rooms below, which saves energy for air-conditioning. Plants capture pollutants and produce oxygen, the roof pairs well with solar panels, and it becomes a home "
         "for plants and insects.")
iw, ih = Image.open(IMG / "benefits_illustration.png").size
PH = 4.95
PW = PH * iw / ih
picture(s, IMG / "benefits_illustration.png", (SW - PW) / 2, 1.7, PW, PH, cover=False, name="Benefits illustration")
benefits = {
    "Runoff": ("62% of each storm's rain stays on the roof, on average", BLUE, SKY, "right", 2),
    "Urban heat": ("Plants and moist soil cool the roof and the street", TREAT["dry"], ROSE, "right", 0),
    "Insulation": ("A buffer against heat and cold for the rooms below", "8A6D57", "F3EEE9", "right", 1),
    "Energy": ("Less air-conditioning; a natural partner for solar panels (PV)", "C98A12", "FBF4DF", "left", 1),
    "Air quality": ("Plants capture pollutants and produce oxygen", "1E9E5A", MINT, "left", 2),
    "Biodiversity": ("Habitat for plants and insects", "7B3FBF", "F1EAFA", "left", 0),
}
bg = {}
for name, (sub, col, fill, side, row) in benefits.items():
    g = group(s, name)
    x = MX if side == "left" else SW - MX - 3.95
    y = 1.7 + row * 1.72
    rect(g, x, y, 3.95, 1.52, fill)
    big = "62%" if name == "Runoff" else name
    paras = [{"text": big, "size": 30 if name == "Runoff" else 20, "font": HEAD, "color": col, "space_after": 2}]
    if name == "Runoff":
        paras.append({"text": "Runoff", "size": 14, "font": HEAD, "color": col, "space_after": 0})
    paras.append({"text": sub, "size": 13})
    text(g, x + 0.25, y + 0.1, 3.5, 1.32, paras, anchor=MSO_ANCHOR.MIDDLE)
    bg[name] = g
animate(s, [[bg[k]] for k in ("Runoff", "Urban heat", "Insulation", "Energy", "Air quality", "Biodiversity")])

# =====================================================================
# 5  Green roof types (your green-blue roof illustration + six roof types)
# =====================================================================
s = base("Green roofs", "Green roof types",
         "Blue and blue-green roofs: Philadelphia Water Department, Stormwater Management Guidance Manual §4.6 and §4.14. "
         "Illustration: Zohar.",
         "On the left, a green-blue roof: a water storage layer under the plants; irrigation from above drains down into "
         "the storage. On the right, six common roof types: extensive green roof, intensive roof garden, blue roof, "
         "blue-green roof, biosolar roof with solar panels over plants, and rooftop farm.")
picture(s, IMG / "green_blue_roof.png", MX, 1.75, 5.9, 3.95, rounded=False, cover=False, name="Green-blue roof")
text(s, MX + 0.2, 5.85, 5.5, 0.8, [{"text": "Green-blue roof", "size": 18, "font": HEAD, "color": BLUE,
                                    "space_after": 0, "align": PP_ALIGN.CENTER},
                                   {"text": "a water storage layer under the plants", "size": 13, "color": MUTED,
                                    "align": PP_ALIGN.CENTER}])


def mini_roof(c, x, y_bottom, w, kind):
    """Small cross-section; y_bottom = bottom of the roof deck."""
    y = y_bottom - 0.12
    rect(c, x, y, w, 0.12, DECK, shape=MSO_SHAPE.RECTANGLE)
    y -= 0.04
    rect(c, x, y, w, 0.04, MEMBRANE, shape=MSO_SHAPE.RECTANGLE)

    def layer(h, col):
        nonlocal y
        y -= h
        rect(c, x, y, w, h, col, shape=MSO_SHAPE.RECTANGLE)

    def plants(n=7, h=0.12):
        for k in range(n):
            rect(c, x + k * w / n, y - h, w / n + 0.01, h, LEAF if k % 2 else GREEN, shape=MSO_SHAPE.OVAL)

    if kind == "green":
        layer(0.16, SOIL)
        plants()
    elif kind == "garden":
        layer(0.3, SOIL)
        plants(5, 0.14)
        rect(c, x + w * 0.66, y - 0.38, 0.04, 0.38, "8D6E4E", shape=MSO_SHAPE.RECTANGLE)
        rect(c, x + w * 0.66 - 0.17, y - 0.62, 0.38, 0.32, GREEN, shape=MSO_SHAPE.OVAL)
    elif kind == "blue":
        layer(0.24, WATER)
        rect(c, x + w - 0.1, y + 0.06, 0.1, 0.18, MEMBRANE, radius=0.2)
    elif kind == "bluegreen":
        layer(0.16, WATER)
        top_water = y
        layer(0.16, SOIL)
        for k in range(3):
            line(c, x + 0.2 + k * 0.3, top_water + 0.14, x + 0.2 + k * 0.3, top_water - 0.12, color="E8C547", width=2)
        plants()
    elif kind == "biosolar":
        layer(0.14, SOIL)
        plants()
        rect(c, x + w * 0.45, y - 0.36, 0.04, 0.3, "6B7780", shape=MSO_SHAPE.RECTANGLE)
        rect(c, x + w * 0.12, y - 0.5, w * 0.78, 0.1, "1F3A5F", shape=MSO_SHAPE.RECTANGLE, rotation=-12)
    elif kind == "farm":
        layer(0.06, "B7BEC4")
        for k in range(3):
            bx = x + 0.05 + k * (w - 0.1) / 3
            rect(c, bx, y - 0.18, (w - 0.1) / 3 - 0.06, 0.18, "A0724A", shape=MSO_SHAPE.RECTANGLE)
            for j in range(2):
                rect(c, bx + 0.04 + j * 0.13, y - 0.3, 0.11, 0.11, ["E8902F", "D9534F", "4E9A6C"][(k + j) % 3],
                     shape=MSO_SHAPE.OVAL)


types = [("green", "Green roof", "Light, low plants on thin soil"),
         ("garden", "Roof garden", "Deep soil, shrubs and trees, space for people"),
         ("blue", "Blue roof", "Holds rain and releases it slowly"),
         ("bluegreen", "Blue-green roof", "Stored rain feeds the plants"),
         ("biosolar", "Biosolar roof", "Solar panels above the plants"),
         ("farm", "Rooftop farm", "Vegetables and herbs in raised beds")]
tg = []
for i, (kind, name, sub) in enumerate(types):
    g = group(s, name)
    x, y = 6.8 + (i % 2) * 3.0, 1.8 + (i // 2) * 1.62
    rect(g, x, y, 2.9, 1.47, SKY if "blue" in kind else MINT)
    mini_roof(g, x + 0.15, y + 1.25, 1.0, kind)
    text(g, x + 1.3, y + 0.12, 1.5, 1.25, [{"text": name, "size": 13, "font": HEAD, "space_after": 2},
                                          {"text": sub, "size": 10.5, "color": MUTED}], anchor=MSO_ANCHOR.MIDDLE)
    tg.append(g)
animate(s, [[g] for g in tg])

# =====================================================================
# 6  Why not here?
# =====================================================================
s = base("Green roofs", "Why not here?",
         "Wu & Biljecki (2021), Landscape and Urban Planning 214, 104167 (share of roof area on buildings with green roofs, "
         "sampled areas; Zurich rule since 1991) · Challenges: Cascone (2019), Sustainability 11, 3020; Zambrano-Prado et "
         "al. (2021), Barcelona",
         "Satellite mapping of 17 cities: in Zurich, buildings with green roofs make up 42% of roof area, helped by a rule "
         "since 1991 that new buildings get green roofs. In drier areas green roofs are harder to maintain. Here the first "
         "challenge is the long dry summer: irrigation and maintenance. Then cost and roof load, and missing local rules. "
         "Which leaves three questions: who monitors, how much water, and what is the impact at city scale.")
cd = CategoryChartData()
cities = [("Zurich", 41.6), ("Berlin", 24.8), ("New York", 17.2), ("Copenhagen", 13.1), ("Paris", 11.2),
          ("Melbourne", 8.4), ("Los Angeles", 4.7), ("Vancouver", 1.0)]
cd.categories = [c for c, _ in cities]
cd.add_series("Share of roof area with green roofs (%)", [v for _, v in cities])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(MX), Inches(2.15), Inches(6.0), Inches(4.5), cd)
chart = gf.chart
chart.has_title = False
chart.has_legend = False
chart.font.name = BODY
chart.font.size = Pt(13)
chart.font.color.rgb = rgb(INK)
pl = chart.plots[0]
pl.gap_width = 45
pl.has_data_labels = True
dl = pl.data_labels
dl.number_format = '0.0"%"'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(12)
ser = pl.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = rgb(LEAF)
pt = ser.points[0]
pt.format.fill.solid()
pt.format.fill.fore_color.rgb = rgb(GREEN)
chart.value_axis.visible = False
chart.value_axis.has_major_gridlines = False
chart.value_axis.maximum_scale = 70
chart.category_axis.reverse_order = True
chart.category_axis.format.line.fill.background()
text(s, MX, 1.7, 6.0, 0.4, "Share of roof area with green roofs, mapped from satellite images", size=13, color=MUTED)
zur = group(s, "Zurich note")
pill(zur, 5.25, 2.33, 1.65, 0.36, "required since 1991", SUN, size=10.5)
text(s, 7.1, 1.75, 5.6, 0.45, "Challenges in Mediterranean cities", size=18, font=HEAD)
chal = [("Long dry summers", "irrigation & maintenance", TREAT["dry"], ROSE),
        ("Cost and roof load", "investment & structure", SUN, SAND),
        ("Missing rules", "regulations & guidelines", BLUE, SKY)]
chg = []
for i, (head, sub, col, fill) in enumerate(chal):
    g = group(s, head)
    rect(g, 7.1, 2.35 + i * 1.02, 5.63, 0.88, fill)
    badge(g, 7.3, 2.52 + i * 1.02, 0.54, str(i + 1), fill=col, size=14)
    text(g, 8.05, 2.42 + i * 1.02, 4.5, 0.75,
         [{"runs": [(head, {"font": HEAD}), ("  ·  " + sub, {"color": MUTED, "size": 13})]}], size=16,
         anchor=MSO_ANCHOR.MIDDLE)
    chg.append(g)
qs = group(s, "Questions")
pill(qs, 7.1, 5.6, 1.8, 0.95, "Who will monitor all these roofs?", BLUE, size=12)
pill(qs, 9.02, 5.6, 1.8, 0.95, "How much water will they need?", SUN, size=12)
pill(qs, 10.94, 5.6, 1.79, 0.95, "What is the impact at city scale?", GREEN, size=12)
animate(s, [[zur]] + [[g] for g in chg] + [[qs]])

# =====================================================================
# 7  From roof to city: objectives
# =====================================================================
s = base("Our research", "From roof to city: our objectives",
         "Research plan. The city-scale step starts after two years of measurements.",
         "Three objectives, from one roof to the whole city. Smart monitoring: follow the roof with free satellite images "
         "instead of sending people up. Operational efficiency: the least water, energy and maintenance that keep the roof "
         "healthy. And next, after two years of measurements: scale up to the city, for runoff and water quality.")
picture(s, IMG / "drone_clean.jpg", 0.95, 2.05, 3.9, 3.9, oval=True, name="Roof")
objs = [("sentinel2.jpg", "Smart monitoring", "free satellite images, fewer site visits", "NOW", GREEN),
        ("aptenia.jpg", "Operational efficiency", "less water, energy and maintenance", "NOW", GREEN),
        ("aerial_city.jpg", "Scale up to the city", "runoff and water quality for a whole city", "NEXT", BLUE)]
og = []
for i, (img, head, sub, status, col) in enumerate(objs):
    g = group(s, head)
    y = 1.75 + i * 1.65
    line(g, 4.85, 4.0, 6.0, y + 0.7, color=LINE, width=2)
    picture(g, IMG / img, 6.0, y, 1.4, 1.4, oval=True)
    text(g, 7.65, y + 0.2, 4.0, 1.1, [{"text": head, "size": 22, "font": HEAD, "space_after": 2},
                                     {"text": sub, "size": 14, "color": MUTED}])
    pill(g, 11.75, y + 0.45, 0.95, 0.42, status, col, size=11)
    og.append(g)
animate(s, [[g] for g in og])

# =====================================================================
# 8  Site
# =====================================================================
s = base("Our research · setup", "A 200 m² living lab on a Technion roof",
         "Drone photo and dimensions: draft manuscript (Zohar et al., in preparation); system: NovaTech 2026 presentation.",
         "Real scale, real rooftop conditions. Four plots of about 50 square metres each, A to D, each with its own "
         "irrigation. A conventional control roof nearby lets us compare.")
picture(s, IMG / "drone_plots.png", MX, 1.75, 7.9, 4.43)
st = []
for i, (big, lab) in enumerate([("200 m²", "blue-green roof in real rooftop conditions"),
                                ("4 plots", "A–D, each with its own irrigation"),
                                ("Control roof", "nearby, for comparison")]):
    g = group(s, big)
    stat_card(g, 8.85, 1.75 + i * 1.53, 3.88, 1.37, big, lab, big_size=26)
    st.append(g)
animate(s, [[g] for g in st])

# =====================================================================
# 9  Water budgets, L/m² per month (July–August)
# =====================================================================
LITERS = {"wet": 1694, "medium": 757, "min": 356, "dry": 96}   # plot totals, draft manuscript irrigation table
PLOT_M2, MONTHS = 50, 2                                          # your assumptions: 50 m² plots, July + August
PER_MONTH = {k: v / PLOT_M2 / MONTHS for k, v in LITERS.items()}
s = base("Our research · setup", "Same roof, four very different water budgets",
         "Plot irrigation totals (draft manuscript, Zohar et al., in preparation) ÷ 50 m² per plot ÷ 2 months (July and "
         "August); every run 10 minutes. 1 L/m² = 1 mm of water.",
         "Under the plants: 15 cm of lightweight substrate over an 85 mm water reservoir, with wicks that lift stored water "
         "to the roots. In July and August each plot got its own schedule. Per square metre per month: about 17 litres on "
         "the wet plot, 1 litre on the dry plot. That's about 18 times less water. The wet and medium plots are our wet "
         "treatments; minimum and dry are the dry treatments.")
picture(s, IMG / "construction.jpg", MX, 1.75, 5.0, 3.25, name="Construction photo")
text(s, MX, 5.15, 5.0, 1.5, [
    {"runs": [("15 cm lightweight substrate ", {"font": HEAD}), ("— 70% perlite, 30% peat", {})], "bullet": True,
     "space_after": 4},
    {"runs": [("85 mm water reservoir ", {"font": HEAD}), ("under the plants", {})], "bullet": True, "space_after": 4},
    {"runs": [("Capillary wicks ", {"font": HEAD}), ("lift stored water to the roots", {})], "bullet": True}], size=14.5)
text(s, 6.1, 1.7, 6.63, 0.75, [{"text": "Irrigation in July and August", "size": 17, "font": HEAD, "space_after": 0},
                               {"text": "litres per m² per month (L/m²/month)", "size": 13, "color": MUTED}])
cd = CategoryChartData()
cd.categories = ["Wet (B)\n2× a day", "Medium (A)\n5× a week", "Minimum (C)\n2× a week", "Dry (D)\n6× a summer"]
cd.add_series("Irrigation (L/m²/month)", [round(PER_MONTH[k], 1) for k in ("wet", "medium", "min", "dry")])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.1), Inches(2.45), Inches(6.63), Inches(3.85), cd)
chart = gf.chart
chart.has_title = False
chart.has_legend = False
chart.font.name = BODY
chart.font.size = Pt(13)
chart.font.color.rgb = rgb(INK)
pl = chart.plots[0]
pl.gap_width = 55
pl.has_data_labels = True
dl = pl.data_labels
dl.number_format = '0.0'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(16)
dl.font.bold = True
for i, key in enumerate(["wet", "medium", "min", "dry"]):
    p = pl.series[0].points[i]
    p.format.fill.solid()
    p.format.fill.fore_color.rgb = rgb(TREAT[key])
chart.value_axis.visible = False
chart.value_axis.has_major_gridlines = False
chart.value_axis.maximum_scale = 20
chart.category_axis.format.line.color.rgb = rgb(LINE)
chart.category_axis.tick_labels.font.size = Pt(12)
groups9 = group(s, "Treatment groups")
pill(groups9, 6.35, 6.38, 3.0, 0.34, "wet treatments", TREAT["wet"], size=11)
pill(groups9, 9.6, 6.38, 3.0, 0.34, "dry treatments", TREAT["dry"], size=11)
x18 = group(s, "Ratio")
pill(x18, 10.15, 2.75, 2.45, 0.5, f"wet ≈ {LITERS['wet'] / LITERS['dry']:.0f}× dry", INK, size=14)
animate(s, [[groups9], [x18]])

# =====================================================================
# 10  Sensors
# =====================================================================
s = base("Our research · setup", "Every plot is measured and controlled",
         "Photos and components: draft manuscript, Fig. 4 (Zohar et al., in preparation); NovaTech 2026 presentation.",
         "The valve decides how much water stays on the roof. The weather station gives what we need to estimate "
         "evaporation. The field spectroradiometer is the ground truth for the satellite.")
pg = []
for i, (f, head, sub) in enumerate([("valve.jpg", "Outlet valve", "with water level sensor, in every plot"),
                                    ("storage_module.jpg", "Storage module", "the blue layer under the plants"),
                                    ("station.jpg", "Weather station", "rain, wind, sun, temperature, humidity"),
                                    ("aptenia.jpg", "Aptenia cordifolia", "a drought-resilient succulent")]):
    g = group(s, head)
    x = MX + i * 3.08
    picture(g, IMG / f, x, 1.75, 2.88, 3.2)
    text(g, x, 5.1, 2.88, 1.0, [{"text": head, "size": 16, "font": HEAD, "space_after": 2,
                                 "italic": head.startswith("Aptenia")},
                                {"text": sub, "size": 13, "color": MUTED}])
    pg.append(g)
plus = group(s, "Plus")
rect(plus, MX, 6.05, SW - 2 * MX, 0.6, SOFT)
text(plus, MX + 0.3, 6.05, SW - 2 * MX - 0.6, 0.6, [{"runs": [("Plus: ", {"font": HEAD, "color": GREEN}), (
    "soil moisture sensors in every plot, and field spectroradiometer campaigns as ground truth", {})]}],
     size=14, anchor=MSO_ANCHOR.MIDDLE)
animate(s, [[g] for g in pg] + [[plus]])

# =====================================================================
# 11  Monitoring from space
# =====================================================================
s = base("Our research · setup", "Monitoring from space",
         "Sentinel-2 image: ESA. Pixel layout: draft manuscript (Zohar et al., in preparation).",
         "Sentinel-2 is free and covers the world every few days at 10 m. A vegetation index turns reflected light into a "
         "greenness number. Our roof falls into two pixels: one over the two wet treatments, one over the two dry "
         "treatments. In the figures they are called P6 and P7. Each pixel also sees some of the surroundings.")
picture(s, IMG / "sentinel2.jpg", MX, 1.75, 3.6, 4.85)
px = group(s, "Pixels")
picture(px, IMG / "pixel_grid.png", 4.45, 1.75, 4.3, 2.4)
text(px, 4.45, 4.3, 4.3, 2.3, [
    {"runs": [("Wet-treatment pixel ", {"font": HEAD, "color": TREAT["wet"]}), ("(P6): wet + medium plots", {})],
     "space_after": 6},
    {"runs": [("Dry-treatment pixel ", {"font": HEAD, "color": TREAT["dry"]}), ("(P7): minimum + dry plots", {})],
     "space_after": 6},
    {"text": "Each 10 m pixel also sees some of the roof's surroundings.", "size": 13, "color": MUTED}], size=14.5)
card = group(s, "Sentinel card")
rect(card, 9.05, 1.75, 3.68, 4.85, MINT)
text(card, 9.3, 2.0, 3.2, 4.4, [
    {"text": "Sentinel-2", "size": 22, "font": HEAD, "color": GREEN, "space_after": 2},
    {"text": "Free · global · 10 m pixels", "size": 14, "color": MUTED, "space_after": 14},
    {"text": "Greenness", "size": 16, "font": HEAD, "space_after": 0},
    {"text": "NDVI, GNDVI", "size": 14, "color": MUTED, "space_after": 10},
    {"text": "Plant moisture", "size": 16, "font": HEAD, "space_after": 0},
    {"text": "NDMI", "size": 14, "color": MUTED, "space_after": 10},
    {"text": "Ground truth", "size": 16, "font": HEAD, "space_after": 0},
    {"text": "Field spectroradiometer on the plots", "size": 14, "color": MUTED}])
animate(s, [[px], [card]])

# =====================================================================
# 12  Results: field measurements (full four-panel figure + statistics)
# =====================================================================
s = base("Results · field measurements", "On the ground: drier soil and stressed plants within weeks",
         "Field measurements 7 Jul – 23 Sep 2026. Change = 23 Sep minus 7 Jul; paired t-test plus trend tests (linear, "
         "Spearman, Kendall, Theil–Sen). Draft manuscript, Tables XA and A2 (Zohar et al., in preparation).",
         "These are field measurements on the plots: soil moisture from the sensors, and three vegetation indices from the "
         "field spectroradiometer. The wet plot stayed stable on every measure. Medium, minimum and dry lost 11 to 14 "
         "percentage points of soil moisture, and their greenness fell, most strongly in minimum and dry. Plant moisture, "
         "NDMI, dropped only in the two driest plots. We do not claim one index is the most sensitive yet.")
picture(s, IMG / "field_four_panels.png", MX - 0.25, 1.5, 6.2, 5.35, rounded=False, cover=False, name="Field figure")
cards12 = [("a", "Soil moisture (VWC)", "Medium −12.9, minimum −14.5, dry −11.4 percentage points (p < 0.001). "
                                        "Wet: no decline.", TREAT["dry"], ROSE),
           ("b", "NDVI · greenness", "Minimum −0.22, dry −0.16 (all tests significant). Medium −0.05, no consistent "
                                     "trend. Wet: no decline.", TREAT["medium"], MINT),
           ("c", "GNDVI · greenness", "Medium −0.12, minimum −0.18, dry −0.14 (p < 0.001, all tests). "
                                      "Wet: no consistent trend.", TREAT["min"], SAND),
           ("d", "NDMI · plant moisture", "Falling trend in minimum and dry only; end-point drop significant in dry "
                                          "(−0.12, p = 0.03).", TREAT["wet"], SKY)]
cg = []
for i, (letter, head, sub, col, fill) in enumerate(cards12):
    g = group(s, head)
    y = 1.65 + i * 1.27
    rect(g, 6.6, y, 6.13, 1.15, fill)
    badge(g, 6.78, y + 0.3, 0.52, letter, fill=col, size=14)
    text(g, 7.5, y + 0.08, 5.1, 1.0, [{"text": head, "size": 15, "font": HEAD, "space_after": 1},
                                     {"text": sub, "size": 12}], anchor=MSO_ANCHOR.MIDDLE)
    cg.append(g)
animate(s, [[g] for g in cg])

# =====================================================================
# 13  Results: from space (new figure)
# =====================================================================
s = base("Results · satellite", "From space: establishment, then the dry treatments faded",
         "Sentinel-2 GNDVI and NDVI; linear slope within each period, per 30 days. Draft manuscript, Table X and Table S1 "
         "(Zohar et al., in preparation).",
         "Left to right: bare roof before installation, establishment in spring 2026, then the summer irrigation "
         "treatments. Both roof pixels greened during establishment. In summer, GNDVI over the dry treatments declined; "
         "over the wet treatments it did not, and neither did the control roofs. NDVI did not pick up the summer decline.")
picture(s, IMG / "sentinel_segmented_trim.png", MX - 0.25, 1.5, 6.6, 5.35, rounded=False, cover=False, name="Satellite figure")
cards13 = [("Establishment", "GNDVI rose +0.033 (wet-treatment pixel, p = 0.003) and +0.056 per 30 days "
                             "(dry-treatment pixel, p < 0.001)", GREEN, MINT),
           ("Summer · dry treatments", "GNDVI fell −0.023 per 30 days (p = 0.004)", TREAT["dry"], ROSE),
           ("Summer · wet treatments and control", "No significant change in GNDVI or NDVI", TREAT["wet"], SKY),
           ("NDVI vs. GNDVI", "NDVI missed the summer decline over the dry treatments (−0.003, p = 0.77)",
            SUN, SAND)]
cg = []
for i, (head, sub, col, fill) in enumerate(cards13):
    g = group(s, head)
    y = 1.65 + i * 1.27
    rect(g, 7.0, y, 5.73, 1.15, fill)
    text(g, 7.25, y + 0.08, 5.3, 1.0, [{"text": head, "size": 15, "font": HEAD, "color": col, "space_after": 1},
                                      {"text": sub, "size": 12.5}], anchor=MSO_ANCHOR.MIDDLE)
    cg.append(g)
animate(s, [[g] for g in cg])

# =====================================================================
# 14  Results: satellite vs ground
# =====================================================================
s = base("Results · satellite vs. ground", "Satellite vs. ground: a good match over the dry treatments",
         "Sentinel-2 vs. field spectroradiometer scaled to the pixel, 11 matched dates (±5 days). Draft manuscript, Fig. JJ "
         "(Zohar et al., in preparation). P6 = wet-treatment pixel, P7 = dry-treatment pixel.",
         "Each point is one date: field on the x axis, satellite on the y axis; the dashed line is perfect agreement. Over "
         "the dry treatments the match is reasonable. Over the wet treatments it is weak. Mixed pixels are the main "
         "suspect: each pixel also sees the surroundings, including 4 square metres of eucalyptus.")
picture(s, IMG / "fig_field_vs_sat_gndvi.png", MX, 1.7, 8.0, 4.95, rounded=False, cover=False)
rc = []
for i, (big, lab, col, fill) in enumerate([
        ("R² 0.41–0.65", "Dry treatments (P7): good agreement, error 0.03–0.05", GREEN, MINT),
        ("R² ≈ 0", "Wet treatments (P6): weak agreement", TREAT["wet"], SKY),
        ("Mixed pixels", "each pixel also sees surroundings, incl. 4 m² of eucalyptus", SUN, SAND)]):
    g = group(s, big)
    stat_card(g, 8.85, 1.75 + i * 1.65, 3.88, 1.5, big, lab, fill=fill, big_color=col, big_size=24)
    rc.append(g)
animate(s, [[g] for g in rc])

# =====================================================================
# 15  Future research
# =====================================================================
s = base("Future research", "Next: two years of data, then the whole city",
         "Research plan.",
         "Now: irrigation, satellite monitoring, and biodiversity on the roof, both microbial and insects. Next: two years "
         "of measurements including winter runoff, plus a pot experiment on plant water use, stress and recovery. Then the "
         "city: a runoff and water-quality simulation, with an economic calculation and potentially a life-cycle "
         "assessment.")
steps = [("Now", "2026", "Irrigation, satellite monitoring and biodiversity: microbes and insects", GREEN, MINT),
         ("Next", "two years", "Measure winter runoff too. Pot experiment: plant water use, stress and recovery", BLUE, SKY),
         ("Then", "city scale", "Runoff and water-quality simulation, economic calculation, potential LCA", SUN, SAND)]
fg = []
for i, (head, when, sub, col, fill) in enumerate(steps):
    g = group(s, head)
    y = 1.75 + i * 1.65
    rect(g, MX, y, 7.4, 1.5, fill)
    badge(g, MX + 0.3, y + 0.38, 0.75, str(i + 1), fill=col, size=20)
    text(g, MX + 1.35, y + 0.15, 5.85, 1.25, [
        {"runs": [(head, {"font": HEAD, "color": col, "size": 18}), (f"  ·  {when}", {"color": MUTED, "size": 14})],
         "space_after": 3}, {"text": sub, "size": 14}], anchor=MSO_ANCHOR.MIDDLE)
    fg.append(g)
picture(s, IMG / "roof_lush.jpg", 8.35, 1.75, 4.38, 4.8, name="Our roof", focus_y=0.6)
animate(s, [[g] for g in fg])

# =====================================================================
# 16  Take-home
# =====================================================================
s = base("Take-home messages", "Three things to take home", None,
         "Repeat the three messages. Invite questions, and invite people to visit the roof.")
msgs = ["Green and blue-green roofs are a proven nature-based solution. In dry Mediterranean cities, water and "
        "monitoring effort hold them back.",
        "Water limits show fast: with little irrigation, soil moisture fell by 11 to 14 points and plants faded within "
        "weeks. The wet plot stayed stable.",
        "A free 10 m satellite tracked the roof's establishment and the summer decline of the dry treatments, but small "
        "roofs share pixels, so the method needs care."]
mg = []
for i, m in enumerate(msgs):
    g = group(s, f"Message {i + 1}")
    y = 1.75 + i * 1.45
    badge(g, MX, y + 0.12, 0.65, str(i + 1), fill=[GREEN, TREAT["min"], BLUE][i], size=18)
    text(g, MX + 0.9, y, 6.3, 1.35, m, size=16, anchor=MSO_ANCHOR.MIDDLE)
    mg.append(g)
picture(s, IMG / "team.jpg", 8.1, 1.75, 4.63, 3.55)
text(s, 8.1, 5.4, 4.63, 0.9, [{"text": "Thank you — questions welcome", "size": 18, "font": HEAD, "color": GREEN,
                              "space_after": 2},
                             {"text": "Yonatan Zohar · yonatanz@campus.technion.ac.il", "size": 13, "color": MUTED}])
lx = MX
for f, w in [("logo_water4all.png", 1.5), ("logo_rainwater.png", 0.75), ("logo_smartflow.png", 1.45),
             ("logo_gwri.png", 0.6)]:
    iw, ih = Image.open(IMG / f).size
    h = w * ih / iw
    s.shapes.add_picture(str(IMG / f), Inches(lx), Inches(6.95 - h / 2), Inches(w), Inches(h))
    lx += w + 0.45
animate(s, [[g] for g in mg])

OUT.parent.mkdir(parents=True, exist_ok=True)
if OUT.exists():
    raise SystemExit(f"{OUT} already exists; bump the version instead of overwriting")
prs.save(OUT)
print("saved", OUT, len(prs.slides), "slides")
