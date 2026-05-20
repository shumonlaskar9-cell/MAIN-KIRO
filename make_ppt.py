"""Generate a polished PowerPoint presentation on Repeat Breeding in cattle.

Theme : navy + amber/teal (modern, professional dairy)
Polish: section divider slides, cow-vs-buffalo comparison table,
        key-takeaway callouts, soft drop-shadows on cards,
        speaker notes on every slide, and embedded photographs
        from Wikimedia Commons.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
IMG  = os.path.join(HERE, "images")

# ---------- Polished theme ----------
NAVY      = RGBColor(0x1B, 0x3A, 0x57)   # primary deep navy
NAVY_DARK = RGBColor(0x0F, 0x24, 0x3A)   # darker navy (shadows)
AMBER     = RGBColor(0xE0, 0x7A, 0x18)   # warm amber (accent)
TEAL      = RGBColor(0x2E, 0x86, 0xAB)   # ocean teal (accent 2)
CREAM     = RGBColor(0xF8, 0xF4, 0xED)   # warm cream background
TINT      = RGBColor(0xEC, 0xE3, 0xD2)   # darker cream / soft beige
DARK      = RGBColor(0x1F, 0x25, 0x33)   # text body
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GRAY      = RGBColor(0x55, 0x5C, 0x66)
LIGHTGRAY = RGBColor(0xD7, 0xDA, 0xDF)
LIGHTNAVY = RGBColor(0xDC, 0xE6, 0xF0)   # soft navy tint

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# images
COW_HOLSTEIN = os.path.join(IMG, "cow_holstein.jpg")
COW_DAIRY    = os.path.join(IMG, "cow_dairy.jpg")
COW_FARM     = os.path.join(IMG, "cow_farm.jpg")
SAHIWAL      = os.path.join(IMG, "sahiwal.jpg")
BUFFALO      = os.path.join(IMG, "buffalo_murrah.jpg")

TOTAL = 24  # final slide count


# =================================================================
# Layout / drawing helpers
# =================================================================
def add_bg(slide, color=CREAM):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.line.fill.background(); bg.fill.solid()
    bg.fill.fore_color.rgb = color
    return bg


def add_side_bar(slide):
    """Thin navy bar with amber stripe on the far left of every content slide."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.30), SH)
    bar.line.fill.background(); bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(0.30), 0, Inches(0.06), SH)
    accent.line.fill.background(); accent.fill.solid()
    accent.fill.fore_color.rgb = AMBER


def add_corner_badge(slide, page_num, total):
    """Small navy disc with page number in top-right corner."""
    disc = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                  Inches(12.55), Inches(0.32),
                                  Inches(0.55), Inches(0.55))
    disc.line.color.rgb = AMBER; disc.line.width = Pt(1.25)
    disc.fill.solid(); disc.fill.fore_color.rgb = NAVY
    tf = disc.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f"{page_num:02d}"
    r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"


def add_footer(slide, page_num, total):
    # left footer text
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.1),
                                  Inches(8), Inches(0.3))
    tf = tb.text_frame; tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "Repeat Breeding in Cattle  |  Veterinary Gynaecology & Obstetrics"
    r.font.size = Pt(9); r.font.color.rgb = GRAY; r.font.name = "Calibri"

    # right footer - page  x / y
    tb2 = slide.shapes.add_textbox(Inches(11.0), Inches(7.1),
                                   Inches(2.0), Inches(0.3))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run(); r2.text = f"Page {page_num} of {total}"
    r2.font.size = Pt(9); r2.font.color.rgb = GRAY; r2.font.name = "Calibri"

    # thin separator line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0.6), Inches(7.05),
                                  Inches(12.13), Emu(6000))
    line.line.fill.background(); line.fill.solid()
    line.fill.fore_color.rgb = LIGHTGRAY


def add_title(slide, title, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(0.65), Inches(0.30),
                                  Inches(11.5), Inches(0.85))
    tf = tb.text_frame; tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size = Pt(32); r.font.bold = True
    r.font.color.rgb = NAVY; r.font.name = "Calibri"

    # underline (amber + teal segments)
    u1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(0.65), Inches(1.10),
                                Inches(0.9), Inches(0.07))
    u1.line.fill.background(); u1.fill.solid()
    u1.fill.fore_color.rgb = AMBER
    u2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(1.55), Inches(1.10),
                                Inches(0.4), Inches(0.07))
    u2.line.fill.background(); u2.fill.solid()
    u2.fill.fore_color.rgb = TEAL

    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.65), Inches(1.22),
                                      Inches(12.0), Inches(0.42))
        sp = sb.text_frame.paragraphs[0]
        sr = sp.add_run(); sr.text = subtitle
        sr.font.size = Pt(13); sr.font.italic = True
        sr.font.color.rgb = GRAY; sr.font.name = "Calibri"


def add_bullets(slide, bullets, left=Inches(0.85), top=Inches(1.85),
                width=Inches(12), height=Inches(5.0), size=17):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    for i, item in enumerate(bullets):
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = level
        p.space_after = Pt(6)
        bullet = "▸  " if level == 0 else "•  "
        r = p.add_run()
        r.text = bullet + text
        r.font.size = Pt(size - level * 2)
        r.font.color.rgb = DARK
        r.font.name = "Calibri"


def add_card(slide, left, top, width, height, title, body,
             head_color=NAVY, shadow=True):
    # subtle drop shadow (offset darker rectangle)
    if shadow:
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    left + Inches(0.05),
                                    top + Inches(0.05),
                                    width, height)
        sh.line.fill.background(); sh.fill.solid()
        sh.fill.fore_color.rgb = LIGHTGRAY

    head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top,
                                  width, Inches(0.55))
    head.line.fill.background(); head.fill.solid()
    head.fill.fore_color.rgb = head_color
    htf = head.text_frame
    htf.margin_left = Inches(0.18); htf.margin_top = Inches(0.05)
    htf.margin_bottom = Inches(0.05)
    hp = htf.paragraphs[0]; hp.alignment = PP_ALIGN.LEFT
    hp.space_before = Pt(0)
    hr = hp.add_run(); hr.text = title
    hr.font.size = Pt(15); hr.font.bold = True
    hr.font.color.rgb = WHITE; hr.font.name = "Calibri"

    body_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      left, top + Inches(0.55),
                                      width, height - Inches(0.55))
    body_box.line.color.rgb = head_color
    body_box.line.width = Pt(0.75)
    body_box.fill.solid(); body_box.fill.fore_color.rgb = WHITE
    btf = body_box.text_frame; btf.word_wrap = True
    btf.margin_left = Inches(0.18); btf.margin_right = Inches(0.18)
    btf.margin_top = Inches(0.12)
    for i, line in enumerate(body):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        p.space_after = Pt(4)
        r = p.add_run(); r.text = "•  " + line
        r.font.size = Pt(12); r.font.color.rgb = DARK
        r.font.name = "Calibri"


def add_takeaway(slide, text, left=Inches(0.85), top=Inches(6.0),
                 width=Inches(11.6), height=Inches(0.95)):
    """Amber-bordered 'Key Takeaway' callout box."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 left, top, width, height)
    box.line.color.rgb = AMBER; box.line.width = Pt(1.5)
    box.fill.solid(); box.fill.fore_color.rgb = WHITE

    # left amber tab
    tab = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 left, top, Inches(0.18), height)
    tab.line.fill.background(); tab.fill.solid()
    tab.fill.fore_color.rgb = AMBER

    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.45); tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.08)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = "KEY TAKEAWAY  "
    r.font.size = Pt(11); r.font.bold = True
    r.font.color.rgb = AMBER; r.font.name = "Calibri"
    r2 = p.add_run(); r2.text = text
    r2.font.size = Pt(13); r2.font.color.rgb = DARK
    r2.font.italic = True; r2.font.name = "Calibri"


def add_image_with_frame(slide, path, left, top, width, height,
                         caption=None, frame_color=NAVY):
    # outer thin frame (acts like a photo border)
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   left - Inches(0.04), top - Inches(0.04),
                                   width + Inches(0.08), height + Inches(0.08))
    frame.line.fill.background(); frame.fill.solid()
    frame.fill.fore_color.rgb = frame_color

    pic = slide.shapes.add_picture(path, left, top, width=width, height=height)

    if caption:
        cb = slide.shapes.add_textbox(left, top + height + Inches(0.04),
                                      width, Inches(0.3))
        cp = cb.text_frame.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
        cr = cp.add_run(); cr.text = caption
        cr.font.size = Pt(10); cr.font.italic = True
        cr.font.color.rgb = GRAY; cr.font.name = "Calibri"
    return pic


def add_section_divider(slide, num, title, sub=None, accent=AMBER):
    """Big full-bleed divider page."""
    add_bg(slide, NAVY)
    # diagonal accent stripe (thin)
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    0, Inches(2.4),
                                    SW, Inches(0.12))
    stripe.line.fill.background(); stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent

    # giant section number
    nb = slide.shapes.add_textbox(Inches(0.7), Inches(2.7),
                                  Inches(5.0), Inches(2.5))
    np_ = nb.text_frame.paragraphs[0]
    nr = np_.add_run(); nr.text = f"0{num}"
    nr.font.size = Pt(140); nr.font.bold = True
    nr.font.color.rgb = accent; nr.font.name = "Calibri"

    # tag above title
    tag = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(5.6), Inches(3.0),
                                 Inches(1.2), Inches(0.32))
    tag.line.fill.background(); tag.fill.solid()
    tag.fill.fore_color.rgb = accent
    ttf = tag.text_frame; ttf.margin_left = Inches(0.1)
    ttf.margin_top = Inches(0.02); ttf.margin_bottom = Inches(0.02)
    tp = ttf.paragraphs[0]
    tr = tp.add_run(); tr.text = "SECTION"
    tr.font.size = Pt(11); tr.font.bold = True
    tr.font.color.rgb = NAVY; tr.font.name = "Calibri"

    # title text
    tb = slide.shapes.add_textbox(Inches(5.55), Inches(3.4),
                                  Inches(7.5), Inches(1.2))
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size = Pt(40); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"

    if sub:
        sb = slide.shapes.add_textbox(Inches(5.55), Inches(4.6),
                                      Inches(7.5), Inches(0.8))
        sp = sb.text_frame.paragraphs[0]
        sr = sp.add_run(); sr.text = sub
        sr.font.size = Pt(15); sr.font.italic = True
        sr.font.color.rgb = LIGHTNAVY; sr.font.name = "Calibri"

    # bottom-right brand
    br = slide.shapes.add_textbox(Inches(10.5), Inches(6.85),
                                  Inches(2.5), Inches(0.35))
    bp = br.text_frame.paragraphs[0]; bp.alignment = PP_ALIGN.RIGHT
    brun = bp.add_run(); brun.text = "REPEAT BREEDING"
    brun.font.size = Pt(10); brun.font.bold = True
    brun.font.color.rgb = accent; brun.font.name = "Calibri"


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# =================================================================
# SLIDES
# =================================================================

# -----------------------------------------------------------
# 1. TITLE SLIDE
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_bg(s, NAVY)

# soft amber band
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.95),
                          SW, Inches(1.55))
band.line.fill.background(); band.fill.solid()
band.fill.fore_color.rgb = AMBER

# tag
tag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(0.7),
                        Inches(3.4), Inches(0.45))
tag.line.fill.background(); tag.fill.solid()
tag.fill.fore_color.rgb = AMBER
ttf = tag.text_frame; ttf.margin_left = Inches(0.18)
ttf.margin_top = Inches(0.04); ttf.margin_bottom = Inches(0.04)
tp = ttf.paragraphs[0]
tr = tp.add_run(); tr.text = "ANIMAL REPRODUCTION"
tr.font.size = Pt(13); tr.font.bold = True
tr.font.color.rgb = NAVY; tr.font.name = "Calibri"

# small horizontal lines decoration above title
for i, x in enumerate([0.65, 1.05, 1.45]):
    d = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                           Inches(x), Inches(1.45),
                           Inches(0.30), Inches(0.04))
    d.line.fill.background(); d.fill.solid()
    d.fill.fore_color.rgb = AMBER if i == 0 else WHITE

# title
tb = s.shapes.add_textbox(Inches(0.65), Inches(1.7), Inches(12), Inches(1.2))
p = tb.text_frame.paragraphs[0]
r = p.add_run(); r.text = "Repeat Breeding"
r.font.size = Pt(72); r.font.bold = True
r.font.color.rgb = WHITE; r.font.name = "Calibri"

tb2 = s.shapes.add_textbox(Inches(0.65), Inches(3.05), Inches(12), Inches(0.7))
p2 = tb2.text_frame.paragraphs[0]
r2 = p2.add_run()
r2.text = "Causes  ·  Diagnosis  ·  Treatment  ·  Prevention"
r2.font.size = Pt(24); r2.font.bold = True
r2.font.color.rgb = NAVY; r2.font.name = "Calibri"

tb3 = s.shapes.add_textbox(Inches(0.65), Inches(3.65), Inches(12), Inches(0.5))
p3 = tb3.text_frame.paragraphs[0]
r3 = p3.add_run(); r3.text = "in Cows and Buffaloes"
r3.font.size = Pt(18); r3.font.italic = True
r3.font.color.rgb = NAVY_DARK; r3.font.name = "Calibri"

# Hero photos
add_image_with_frame(s, COW_HOLSTEIN, Inches(0.65), Inches(5.0),
                     Inches(4.0), Inches(1.85),
                     "Holstein Friesian — high-yielding dairy cow", WHITE)
add_image_with_frame(s, BUFFALO, Inches(8.7), Inches(5.0),
                     Inches(4.0), Inches(1.85),
                     "Murrah buffalo — leading Indian dairy breed", WHITE)

# Center plate between hero photos
plate = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                           Inches(5.0), Inches(5.05),
                           Inches(3.4), Inches(1.75))
plate.line.color.rgb = AMBER; plate.line.width = Pt(1.5)
plate.fill.solid(); plate.fill.fore_color.rgb = NAVY_DARK
ptf = plate.text_frame; ptf.word_wrap = True
ptf.margin_left = Inches(0.15); ptf.margin_right = Inches(0.15)
ptf.margin_top = Inches(0.18)
pp1 = ptf.paragraphs[0]; pp1.alignment = PP_ALIGN.CENTER
prr = pp1.add_run(); prr.text = "Veterinary Gynaecology"
prr.font.size = Pt(15); prr.font.bold = True
prr.font.color.rgb = WHITE; prr.font.name = "Calibri"
pp2 = ptf.add_paragraph(); pp2.alignment = PP_ALIGN.CENTER
prr2 = pp2.add_run(); prr2.text = "& Obstetrics"
prr2.font.size = Pt(15); prr2.font.bold = True
prr2.font.color.rgb = WHITE; prr2.font.name = "Calibri"
pp3 = ptf.add_paragraph(); pp3.alignment = PP_ALIGN.CENTER
pp3.space_before = Pt(8)
prr3 = pp3.add_run(); prr3.text = "Seminar Presentation"
prr3.font.size = Pt(12); prr3.font.italic = True
prr3.font.color.rgb = AMBER; prr3.font.name = "Calibri"

set_notes(s, ("Welcome the audience. Introduce the topic of repeat breeding "
              "as a major reproductive disorder in dairy cattle and buffaloes. "
              "Mention that the seminar will cover etiology, diagnosis, "
              "treatment, prevention and a special focus on buffaloes."))

# -----------------------------------------------------------
# 2. OUTLINE
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_bg(s); add_side_bar(s)
add_title(s, "Presentation Roadmap",
          "Twenty-four slides organised into four sections")
add_corner_badge(s, 2, TOTAL)

sections = [
    ("I",   "Foundations",            ["Introduction & Definition",
                                       "Incidence & Economic Impact"], TEAL),
    ("II",  "Etiology",               ["Overview of causes",
                                       "Infectious / Genital",
                                       "Nutritional & Metabolic",
                                       "Hormonal & Endocrine",
                                       "Management & AI errors",
                                       "Pathophysiology flow"], AMBER),
    ("III", "Diagnosis & Treatment",  ["Clinical approach",
                                       "Treatment principles",
                                       "Hormonal protocols",
                                       "Intra-uterine therapy"], NAVY),
    ("IV",  "Buffalo Focus & Prevention",
            ["Buffalo-specific repeat breeding",
             "Cow vs Buffalo comparison",
             "Prevention strategies",
             "Recent advances",
             "Conclusion"], TEAL),
]

x_positions = [0.85, 4.0, 7.15, 10.3]
for col, (num, name, items, color) in enumerate(sections):
    L = Inches(x_positions[col])
    # number badge
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, L, Inches(1.85),
                               Inches(0.65), Inches(0.65))
    badge.line.color.rgb = color; badge.line.width = Pt(1.5)
    badge.fill.solid(); badge.fill.fore_color.rgb = NAVY
    btf = badge.text_frame
    btf.margin_left = btf.margin_right = 0
    btf.margin_top = btf.margin_bottom = Inches(0.02)
    bp = btf.paragraphs[0]; bp.alignment = PP_ALIGN.CENTER
    br = bp.add_run(); br.text = num
    br.font.size = Pt(14); br.font.bold = True
    br.font.color.rgb = WHITE; br.font.name = "Calibri"

    # section name
    nb = s.shapes.add_textbox(L + Inches(0.75), Inches(1.85),
                              Inches(2.5), Inches(0.7))
    np_ = nb.text_frame.paragraphs[0]
    nr = np_.add_run(); nr.text = name
    nr.font.size = Pt(15); nr.font.bold = True
    nr.font.color.rgb = NAVY; nr.font.name = "Calibri"
    np2 = nb.text_frame.add_paragraph()
    nr2 = np2.add_run(); nr2.text = "Section " + num
    nr2.font.size = Pt(10); nr2.font.color.rgb = color; nr2.font.name = "Calibri"

    # items
    yb = s.shapes.add_textbox(L, Inches(2.75),
                              Inches(2.95), Inches(4.0))
    ytf = yb.text_frame; ytf.word_wrap = True
    for i, it in enumerate(items):
        p = ytf.paragraphs[0] if i == 0 else ytf.add_paragraph()
        p.space_after = Pt(5)
        rr = p.add_run(); rr.text = "•  " + it
        rr.font.size = Pt(12); rr.font.color.rgb = DARK; rr.font.name = "Calibri"

add_footer(s, 2, TOTAL)
set_notes(s, ("Walk through the four-section structure. Tell the audience "
              "they will see section divider slides between major parts."))

# -----------------------------------------------------------
# 3. SECTION DIVIDER I
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_section_divider(s, 1, "Foundations",
                    "Setting the stage — what repeat breeding is and why it matters",
                    accent=AMBER)
set_notes(s, "Introduce Section I — the foundational concepts of repeat breeding.")

# -----------------------------------------------------------
# 4. INTRODUCTION
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Introduction",
          "Reproductive efficiency is the backbone of profitable dairying")
add_corner_badge(s, 4, TOTAL)

add_bullets(s, [
    "Reproduction is the most important factor governing the economic "
    "success of any dairy enterprise.",
    "A normal fertile cow should conceive within 60–90 days post-partum and "
    "deliver one calf every 12–13 months.",
    "Failure to conceive after repeated inseminations — despite normal "
    "estrous cycles and apparently healthy genitalia — is termed "
    "REPEAT BREEDING.",
    "One of the most frustrating and economically damaging reproductive "
    "disorders in cattle and buffaloes worldwide.",
    "Multifactorial — involving the cow, the bull / semen, the inseminator "
    "and the environment.",
], left=Inches(0.85), top=Inches(1.85), width=Inches(8.0), size=15)

add_image_with_frame(s, COW_DAIRY, Inches(9.1), Inches(1.95),
                     Inches(3.7), Inches(2.55),
                     "Holstein dairy cow on pasture")
add_image_with_frame(s, SAHIWAL, Inches(9.1), Inches(4.85),
                     Inches(3.7), Inches(1.5),
                     "Sahiwal — indigenous Indian dairy breed")

add_takeaway(s, "Reproductive failure of even a few cows can erase the "
                "entire margin of a dairy unit.")
add_footer(s, 4, TOTAL)
set_notes(s, ("Reproduction drives milk production, calf supply and culling "
              "decisions. Repeat breeding directly hits all three. "
              "Stress that the cow is otherwise apparently healthy."))

# -----------------------------------------------------------
# 5. DEFINITION
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Definition")
add_corner_badge(s, 5, TOTAL)

# definition shadow + box
sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(0.90), Inches(1.90),
                        Inches(11.6), Inches(1.85))
sh.line.fill.background(); sh.fill.solid()
sh.fill.fore_color.rgb = LIGHTGRAY
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.85), Inches(1.85),
                         Inches(11.6), Inches(1.85))
box.line.color.rgb = AMBER; box.line.width = Pt(2)
box.fill.solid(); box.fill.fore_color.rgb = WHITE
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = Inches(0.4); tf.margin_right = Inches(0.4)
tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]
r = p.add_run(); r.text = "“"
r.font.size = Pt(36); r.font.bold = True
r.font.color.rgb = AMBER; r.font.name = "Calibri"
r2 = p.add_run()
r2.text = (" A repeat breeder is a cow / buffalo of normal breeding age and "
          "parity, cycling regularly at intervals of 18–24 days, with no "
          "clinically detectable abnormality of the genital tract, that has "
          "failed to conceive after three or more successive services with "
          "fertile semen. ")
r2.font.size = Pt(17); r2.font.italic = True
r2.font.color.rgb = DARK; r2.font.name = "Calibri"
r3 = p.add_run(); r3.text = "”"
r3.font.size = Pt(36); r3.font.bold = True
r3.font.color.rgb = AMBER; r3.font.name = "Calibri"

# 5 key criteria as small chips
chips = [
    ("Regular cycles", TEAL),
    ("≥ 3 services",   AMBER),
    ("No abnormality", NAVY),
    ("Fertile semen",  TEAL),
    ("Competent AI",   AMBER),
]
chip_w = Inches(2.25); chip_h = Inches(0.55)
for i, (lab, col) in enumerate(chips):
    L = Inches(0.85 + i * (2.25 + 0.13))
    c = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                           L, Inches(4.1), chip_w, chip_h)
    c.line.fill.background(); c.fill.solid()
    c.fill.fore_color.rgb = col
    ctf = c.text_frame
    cp = ctf.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run(); cr.text = lab
    cr.font.size = Pt(13); cr.font.bold = True
    cr.font.color.rgb = WHITE; cr.font.name = "Calibri"

add_bullets(s, [
    "Also known as: Repeat Breeder Syndrome / Repeat Breeder Cow (RBC).",
    "Older synonym: 'Conception Failure'.",
    "Differentiate from anoestrus, sub-oestrus, true infertility and sterility.",
], top=Inches(5.0), size=15)

add_takeaway(s, "Five criteria together define a repeat breeder — missing "
                "any one means a different reproductive problem.")
add_footer(s, 5, TOTAL)
set_notes(s, ("Emphasise all five criteria. Many cows mis-classified as "
              "repeat breeders are actually anoestrus or sub-oestrus cases."))

# -----------------------------------------------------------
# 6. INCIDENCE & ECONOMICS
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Incidence & Economic Importance")
add_corner_badge(s, 6, TOTAL)

# stat tiles row
stats = [
    ("10 – 15 %", "Global incidence in normal herds", TEAL),
    ("20 – 30 %", "In poorly managed herds", AMBER),
    ("12 – 24 %", "Indian dairy herds (cross-breds higher)", NAVY),
    ("8 – 18 %",  "Buffaloes (often masked by silent oestrus)", TEAL),
]
for i, (val, lab, col) in enumerate(stats):
    L = Inches(0.85 + i * 3.075)
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                            L + Inches(0.05), Inches(1.95),
                            Inches(2.93), Inches(1.6))
    sh.line.fill.background(); sh.fill.solid()
    sh.fill.fore_color.rgb = LIGHTGRAY
    tile = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              L, Inches(1.9), Inches(2.93), Inches(1.6))
    tile.line.color.rgb = col; tile.line.width = Pt(1.5)
    tile.fill.solid(); tile.fill.fore_color.rgb = WHITE
    ttf = tile.text_frame; ttf.word_wrap = True
    ttf.margin_left = Inches(0.15); ttf.margin_right = Inches(0.15)
    ttf.margin_top = Inches(0.18)
    tp = ttf.paragraphs[0]; tp.alignment = PP_ALIGN.CENTER
    tr = tp.add_run(); tr.text = val
    tr.font.size = Pt(28); tr.font.bold = True
    tr.font.color.rgb = col; tr.font.name = "Calibri"
    tp2 = ttf.add_paragraph(); tp2.alignment = PP_ALIGN.CENTER
    tp2.space_before = Pt(4)
    tr2 = tp2.add_run(); tr2.text = lab
    tr2.font.size = Pt(11); tr2.font.color.rgb = DARK; tr2.font.name = "Calibri"

# Economic impact cards
add_card(s, Inches(0.85), Inches(3.85), Inches(3.85), Inches(2.8),
         "Direct Losses",
         ["Extra AI doses & semen cost",
          "Veterinary & hormone bills",
          "Extended dry / open period",
          "Reduced lifetime calves"], head_color=NAVY)
add_card(s, Inches(4.95), Inches(3.85), Inches(3.85), Inches(2.8),
         "Indirect Losses",
         ["Lower lactation yield",
          "Increased calving interval",
          "Premature culling",
          "Loss of genetic progress"], head_color=AMBER)
add_card(s, Inches(9.05), Inches(3.85), Inches(3.85), Inches(2.8),
         "Estimated Cost",
         ["Each extra open day ≈ ₹150–250",
          "Per repeater ≈ ₹8 000–15 000 / lactation",
          "National losses run into thousands of crores",
          "Major drag on smallholder economy"], head_color=TEAL)

add_footer(s, 6, TOTAL)
set_notes(s, ("Place the problem in numerical terms. The audience must "
              "appreciate that a single repeat breeder can wipe out a "
              "smallholder's annual profit."))

# -----------------------------------------------------------
# 7. SECTION DIVIDER II
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_section_divider(s, 2, "Etiology",
                    "A multifactorial syndrome — the cow, the bull, the inseminator & the environment",
                    accent=TEAL)
set_notes(s, "Section II covers all major causes of repeat breeding.")

# -----------------------------------------------------------
# 8. ETIOLOGY OVERVIEW
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Etiology — An Overview",
          "Causes are grouped into four broad heads")
add_corner_badge(s, 8, TOTAL)

add_card(s, Inches(0.85), Inches(1.85), Inches(2.95), Inches(4.7),
         "1. Fertilization Failure",
         ["Improper AI timing",
          "Poor semen quality",
          "Anovulation / delayed ovulation",
          "Tract obstruction",
          "Anti-sperm antibodies"], head_color=NAVY)
add_card(s, Inches(4.0), Inches(1.85), Inches(2.95), Inches(4.7),
         "2. Early Embryonic Death",
         ["Hormonal imbalance",
          "Uterine infections",
          "Heat stress",
          "Chromosomal defects",
          "Nutritional deficiency"], head_color=AMBER)
add_card(s, Inches(7.15), Inches(1.85), Inches(2.95), Inches(4.7),
         "3. Management Errors",
         ["Missed / mis-detected oestrus",
          "Wrong AI technique",
          "Faulty semen handling",
          "Inadequate records",
          "Stress at AI"], head_color=TEAL)
add_card(s, Inches(10.3), Inches(1.85), Inches(2.6), Inches(4.7),
         "4. Genetic / Other",
         ["Inbreeding",
          "Lethal genes",
          "Freemartinism",
          "White heifer disease",
          "Age & parity"], head_color=NAVY)

add_takeaway(s, "Roughly 60 % of repeat breeding is due to fertilization "
                "failure or early embryonic death.")
add_footer(s, 8, TOTAL)
set_notes(s, ("Categorising causes helps direct diagnosis. Most refractory "
              "cases involve more than one category at a time."))

# -----------------------------------------------------------
# 9. INFECTIOUS CAUSES
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Infectious & Genital Tract Causes")
add_corner_badge(s, 9, TOTAL)

add_bullets(s, [
    "Subclinical / chronic endometritis — most common cause (40–60 %).",
    ("Specific infections:", 0),
    ("Brucellosis (Brucella abortus)", 1),
    ("Trichomoniasis (Tritrichomonas foetus)", 1),
    ("Vibriosis / Campylobacteriosis (Campylobacter fetus)", 1),
    ("IBR–IPV (Bovine Herpesvirus-1)", 1),
    ("BVD (Bovine Viral Diarrhoea virus)", 1),
    ("Leptospirosis, Mycoplasma, Ureaplasma spp.", 1),
    "Anatomical defects — cervical stenosis, persistent hymen, segmental aplasia.",
    "Cystic ovarian disease (follicular & luteal cysts).",
    "Salpingitis & hydrosalpinx — block fertilization.",
], size=15)

add_takeaway(s, "Always rule out subclinical endometritis before any "
                "hormonal therapy.")
add_footer(s, 9, TOTAL)
set_notes(s, ("Discuss white-side test, endometrial cytology and "
              "USG findings of intra-uterine fluid as bedside diagnostics "
              "for subclinical endometritis."))

# -----------------------------------------------------------
# 10. NUTRITIONAL CAUSES
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Nutritional & Metabolic Causes")
add_corner_badge(s, 10, TOTAL)

add_bullets(s, [
    "Negative energy balance in early lactation → delayed cyclicity & poor "
    "embryo quality.",
    "Excess dietary protein → high blood urea N (>20 mg/dL) is embryotoxic.",
    ("Mineral & trace-element deficiencies:", 0),
    ("Phosphorus, Calcium — anoestrus, weak heat", 1),
    ("Copper, Cobalt, Zinc, Manganese — silent heat, low conception", 1),
    ("Selenium / Vitamin E — early embryonic death, retained placenta", 1),
    ("Iodine — irregular cycles", 1),
    ("Vitamin A, D, E — endometrial & embryo defects", 1),
    "Mycotoxins (zearalenone, aflatoxin) → ovarian & embryonic disturbance.",
    "Body Condition Score < 2.5 or > 4 at AI is strongly linked to repeat breeding.",
], size=15)

add_takeaway(s, "Aim for BCS 3.0–3.5 at AI; correct minerals 30–45 days "
                "before breeding.")
add_footer(s, 10, TOTAL)
set_notes(s, ("Mention the 'transition cow' concept and how energy balance "
              "in the first 3 weeks post-calving affects fertility weeks "
              "later through follicle quality."))

# -----------------------------------------------------------
# 11. HORMONAL CAUSES
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Hormonal & Endocrine Causes")
add_corner_badge(s, 11, TOTAL)

add_bullets(s, [
    "Inadequate pre-ovulatory LH surge → delayed or failed ovulation.",
    "Sub-luteal progesterone (< 1 ng/mL after Day-5) → embryonic loss.",
    "Persistent corpus luteum (CL) — pseudo-pregnancy effect.",
    "Hyper-prolactinaemia & thyroid dysfunction.",
    "Premature luteolysis — failure of maternal recognition (low Interferon-τ).",
    "Stress-induced cortisol rise → suppression of GnRH / LH.",
    "Heat stress in summer — reduced oestradiol, weak signs of oestrus, "
    "poor oocyte quality.",
], size=17)

add_takeaway(s, "Day-7 progesterone < 4 ng/mL is a strong predictor of "
                "embryonic loss in repeat breeders.")
add_footer(s, 11, TOTAL)
set_notes(s, ("Hormonal causes are mostly diagnosed indirectly through "
              "treatment response; serum P4 assay where available is the "
              "single most useful test."))

# -----------------------------------------------------------
# 12. MANAGEMENT / AI ERRORS
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Management & Inseminator-Related Causes")
add_corner_badge(s, 12, TOTAL)

add_bullets(s, [
    "Improper heat detection — single twice-daily check misses 25–50 % of heats.",
    "Wrong time of AI — outside the optimum window (mid- to late-oestrus).",
    "Faulty semen thawing (water temperature, time) and storage practice.",
    "Repeated use of same sub-fertile bull / semen batch.",
    "Trauma to cervix or uterus during AI; deep-horn deposition errors.",
    "Hygiene lapses → introduction of pathogens at the time of AI.",
    "Lack of accurate breeding records and follow-up.",
], top=Inches(1.85), size=17)

add_takeaway(s, "Up to 40 % of 'repeat breeders' are simply mis-timed AI — "
                "fix detection first.")
add_footer(s, 12, TOTAL)
set_notes(s, ("Discuss the AM-PM rule, role of pedometers / activity "
              "collars, and importance of culture-tested semen batches."))

# -----------------------------------------------------------
# 13. PATHOPHYSIOLOGY
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Pathophysiology — Where Things Go Wrong")
add_corner_badge(s, 13, TOTAL)

labels = ["Oestrus &\nOvulation", "Sperm\nTransport",
          "Fertilization", "Embryo\nDevelopment",
          "Maternal\nRecognition", "Pregnancy"]
left = Inches(0.7)
w = Inches(1.95); h = Inches(1.0)
top = Inches(2.15)
for i, lab in enumerate(labels):
    color = NAVY if i % 2 == 0 else AMBER
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             left + i * (w + Inches(0.07)), top, w, h)
    box.line.fill.background(); box.fill.solid()
    box.fill.fore_color.rgb = color
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_top = Inches(0.1); tf.margin_bottom = Inches(0.1)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = lab
    r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"

for i in range(len(labels) - 1):
    arr_left = left + (i + 1) * w + i * Inches(0.07) - Inches(0.02)
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arr_left,
                           top + Inches(0.42), Inches(0.13), Inches(0.18))
    a.line.fill.background(); a.fill.solid()
    a.fill.fore_color.rgb = DARK

add_bullets(s, [
    "Failure of ovulation or asynchronous ovulation → no fertilization.",
    "Sperm transport disturbance / hostile cervico-uterine environment.",
    "Fertilization failure — poor oocyte / sperm quality, anti-sperm antibodies.",
    "Early embryonic death (Day 8–16) — most common in repeat breeders.",
    "Failure of maternal recognition of pregnancy → CL regresses → return to oestrus.",
], top=Inches(3.55), size=15)

add_takeaway(s, "Day 8–16 is the danger zone — most repeat-breeder embryos "
                "die before maternal recognition of pregnancy.")
add_footer(s, 13, TOTAL)
set_notes(s, ("Explain that repeat breeders typically return to heat at a "
              "normal interval because their embryo died before the "
              "Interferon-tau signal was strong enough to maintain the CL."))

# -----------------------------------------------------------
# 14. SECTION DIVIDER III
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_section_divider(s, 3, "Diagnosis & Treatment",
                    "Systematic clinical approach and evidence-based therapeutics",
                    accent=AMBER)
set_notes(s, "Section III — practical diagnosis and treatment.")

# -----------------------------------------------------------
# 15. DIAGNOSIS
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Diagnosis — Systematic Clinical Approach")
add_corner_badge(s, 15, TOTAL)

add_bullets(s, [
    ("History:", 0),
    ("Age, parity, calving date, services given, AI dates, semen used", 1),
    ("Bull fertility, nutrition, deworming, vaccination, prior treatments", 1),
    ("General Examination — BCS, anaemia, lameness, mastitis, systemic disease.", 0),
    ("Gynaeco-clinical Examination:", 0),
    ("Vaginoscopy — discharge, cervicitis, adhesions", 1),
    ("Per-rectal palpation — uterus, ovaries, cervix, follicles, CL", 1),
    ("Trans-rectal ultrasonography — follicles, CL, uterine fluid, embryo", 1),
    ("Laboratory Tests:", 0),
    ("Cervico-vaginal mucus culture & sensitivity, cytology", 1),
    ("Serum progesterone, blood urea, minerals, liver / thyroid profile", 1),
    ("Serology / PCR — Brucella, IBR, BVD, Leptospira, Trichomonas", 1),
], size=14)

add_takeaway(s, "Always proceed in order: history → general exam → "
                "gynaeco exam → laboratory tests.")
add_footer(s, 15, TOTAL)
set_notes(s, ("Stress the diagnostic discipline. Most field failures come "
              "from skipping history-taking and going straight to hormones."))

# -----------------------------------------------------------
# 16. TREATMENT GENERAL
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Treatment — General Principles",
          "Treat the cause, not just the symptom")
add_corner_badge(s, 16, TOTAL)

add_bullets(s, [
    "Correct underlying nutrition & body condition (BCS 3.0–3.5 at AI).",
    "Treat genital tract infection (intra-uterine antibiotics, antiseptics).",
    "Correct hormonal imbalance with appropriate protocols.",
    "Improve heat detection and AI timing (AM-PM rule / activity meters).",
    "Use proven, high-quality semen; ensure proper thawing & deposition.",
    "Reduce stress — provide shade, water, comfortable housing, fly control.",
    "Maintain proper records; cull chronic non-responders after 5–6 services.",
], left=Inches(0.85), top=Inches(1.85), width=Inches(8.0), size=15)

add_image_with_frame(s, COW_FARM, Inches(9.1), Inches(1.95),
                     Inches(3.7), Inches(2.74),
                     "Well-managed dairy cow on farm")

# elegant quote box
qb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(9.1), Inches(5.05),
                        Inches(3.7), Inches(1.5))
qb.line.color.rgb = AMBER; qb.line.width = Pt(1.5)
qb.fill.solid(); qb.fill.fore_color.rgb = WHITE
qtf = qb.text_frame; qtf.word_wrap = True
qtf.margin_left = Inches(0.18); qtf.margin_right = Inches(0.18)
qtf.margin_top = Inches(0.15)
qp = qtf.paragraphs[0]; qp.alignment = PP_ALIGN.CENTER
qr = qp.add_run(); qr.text = "“Prevention is better, "
qr.font.size = Pt(13); qr.font.italic = True; qr.font.bold = True
qr.font.color.rgb = NAVY
qr2 = qp.add_run(); qr2.text = "but timely treatment saves the cow.”"
qr2.font.size = Pt(13); qr2.font.italic = True
qr2.font.color.rgb = DARK
qp2 = qtf.add_paragraph(); qp2.alignment = PP_ALIGN.CENTER
qp2.space_before = Pt(6)
qr3 = qp2.add_run(); qr3.text = "— Field motto in dairy practice"
qr3.font.size = Pt(10); qr3.font.color.rgb = GRAY

add_takeaway(s, "Hormones are the last step, not the first. Fix nutrition, "
                "infection and AI errors before reaching for the syringe.")
add_footer(s, 16, TOTAL)
set_notes(s, ("Empower the audience to think causally. Hormonal therapy "
              "should be the last step after correcting husbandry."))

# -----------------------------------------------------------
# 17. HORMONAL PROTOCOLS
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Specific Therapeutic Protocols")
add_corner_badge(s, 17, TOTAL)

add_card(s, Inches(0.85), Inches(1.85), Inches(3.85), Inches(4.6),
         "GnRH at AI",
         ["Buserelin 10 µg or Gonadorelin 100 µg IM at AI",
          "Improves LH surge & ovulation",
          "Increases conception by 10–20 %",
          "Best for delayed ovulators"], head_color=NAVY)
add_card(s, Inches(4.95), Inches(1.85), Inches(3.85), Inches(4.6),
         "Progesterone Support",
         ["CIDR / PRID for 7 days post-AI",
          "Long-acting P4 injection on Day 5",
          "Prevents premature luteolysis & embryo loss",
          "Best for cows with short luteal phase"], head_color=AMBER)
add_card(s, Inches(9.05), Inches(1.85), Inches(3.85), Inches(4.6),
         "Ovsynch / Double-AI",
         ["GnRH (D0) → PGF2α (D7) → GnRH (D9) → AI 16 h later",
          "Synchronises ovulation precisely",
          "Useful in herds with poor heat detection",
          "Modified / Co-Synch / Double-Ovsynch for repeaters"], head_color=TEAL)

add_takeaway(s, "Combining GnRH at AI + Day-5 progesterone is a simple, "
                "field-friendly upgrade for repeat breeders.")
add_footer(s, 17, TOTAL)
set_notes(s, ("Discuss dose, timing, contra-indications. Ovsynch needs "
              "good cow handling and timely injections — limited in "
              "smallholder settings."))

# -----------------------------------------------------------
# 18. INTRA-UTERINE THERAPY
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Intra-Uterine & Adjunct Therapy")
add_corner_badge(s, 18, TOTAL)

add_bullets(s, [
    ("Intra-uterine antibiotics (after culture / sensitivity):", 0),
    ("Cephapirin, Ceftiofur, Oxytetracycline, Enrofloxacin", 1),
    ("Avoid penicillin — inactivated by uterine exudate", 1),
    ("Lugol's iodine (1–2 %) flushing — mild irritant, stimulates leucocytosis.", 0),
    ("Oyster glycogen / E. coli LPS — non-specific immuno-stimulation.", 0),
    ("Intra-uterine PGF2α & oxytocin — promote uterine clearance.", 0),
    ("Mineral & vitamin supplementation:", 0),
    ("AD3E, Vitamin E + Selenium, Chelated Zn-Cu-Mn-Co", 1),
    ("Bypass protein & energy correction during AI period", 1),
    ("Herbal / Ayurvedic preparations (Aloes compound, Janova, Prajana) "
     "as adjuncts.", 0),
], size=15)

add_takeaway(s, "Intra-uterine therapy must follow culture & sensitivity "
                "— blind antibiotic infusion damages the endometrium.")
add_footer(s, 18, TOTAL)
set_notes(s, ("Caution: repeated antibiotic infusions can themselves "
              "induce chronic endometritis. Mineral therapy is often "
              "under-used and gives the best return on investment."))

# -----------------------------------------------------------
# 19. SECTION DIVIDER IV
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_section_divider(s, 4, "Buffalo Focus & Prevention",
                    "Special considerations and the road ahead",
                    accent=TEAL)
set_notes(s, "Section IV — buffalo-specific issues, prevention and the future.")

# -----------------------------------------------------------
# 20. BUFFALO-SPECIFIC
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Buffalo-Specific Repeat Breeding",
          "Why buffaloes pose unique reproductive challenges")
add_corner_badge(s, 20, TOTAL)

add_image_with_frame(s, BUFFALO, Inches(0.85), Inches(1.9),
                     Inches(4.5), Inches(3.05),
                     "Murrah buffalo (Bubalus bubalis)")

add_bullets(s, [
    "Silent / sub-oestrus — up to 50–60 % of heats are missed without close "
    "observation.",
    "Late maturity (36–42 months) and longer post-partum anoestrus "
    "(90–150 days) than cows.",
    "Strongly seasonal — short photoperiod (Oct–Feb) is the peak breeding "
    "season.",
    "Summer infertility — heat stress depresses LH surge, oocyte quality "
    "and embryo survival.",
    "Smaller, deeper-seated ovaries with fewer antral follicles than cows.",
    "Higher incidence of true anoestrus, ovarian inactivity and "
    "anovulatory cysts.",
], left=Inches(5.6), top=Inches(1.9), width=Inches(7.3),
   height=Inches(3.5), size=13)

add_card(s, Inches(0.85), Inches(5.2), Inches(3.95), Inches(1.55),
         "Heat Detection",
         ["Observe 4–5 ×/day; early morning & late evening",
          "Teaser bull / pedometer / activity meter",
          "Watch vulvar oedema, mucus, bellowing"], head_color=NAVY)
add_card(s, Inches(4.95), Inches(5.2), Inches(3.95), Inches(1.55),
         "Hormonal / Synchrony",
         ["Ovsynch / Heatsynch / CIDR work well",
          "Pre-synch / Double-Ovsynch in summer",
          "GnRH at AI for late-cycling buffaloes"], head_color=AMBER)
add_card(s, Inches(9.05), Inches(5.2), Inches(3.85), Inches(1.55),
         "Environment & Nutrition",
         ["Wallowing / showers when > 32 °C",
          "Shade, fans, fly control",
          "Bypass fat & mineral supplementation"], head_color=TEAL)

add_footer(s, 20, TOTAL)
set_notes(s, ("Buffaloes are not just 'larger cows' — emphasise silent "
              "oestrus, seasonality and heat-stress management as the "
              "three biggest practical issues."))

# -----------------------------------------------------------
# 21. COW vs BUFFALO COMPARISON TABLE
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Cow vs Buffalo — At a Glance",
          "Comparative reproductive features relevant to repeat breeding")
add_corner_badge(s, 21, TOTAL)

rows = [
    ("Parameter",                      "Cow (Bos taurus / indicus)",  "Buffalo (Bubalus bubalis)"),
    ("Age at puberty",                 "12–18 months",                "24–36 months"),
    ("Age at first calving",           "24–30 months",                "36–48 months"),
    ("Oestrous cycle length",          "21 days (18–24)",             "21 days (18–22)"),
    ("Duration of oestrus",            "12–18 hours",                 "18–30 hours"),
    ("Signs of oestrus",               "Overt — clear mucus, mounting",
                                        "Often silent / sub-oestrus"),
    ("Seasonality",                    "Year-round breeder",
                                        "Strongly seasonal (Oct–Feb)"),
    ("Post-partum anoestrus",          "45–90 days",                  "90–150 days"),
    ("Antral follicle count",          "Higher",                      "Lower"),
    ("Heat-stress sensitivity",        "Moderate",                    "High — needs wallow / shower"),
    ("Common cause of repeat breeding","Subclinical endometritis",
                                        "Silent oestrus + summer infertility"),
]

# build table
left = Inches(0.85); top = Inches(1.85)
total_w = Inches(11.6); total_h = Inches(4.65)
n_rows = len(rows); n_cols = 3
tbl_shape = s.shapes.add_table(n_rows, n_cols, left, top, total_w, total_h)
tbl = tbl_shape.table

# column widths
tbl.columns[0].width = Inches(3.0)
tbl.columns[1].width = Inches(4.3)
tbl.columns[2].width = Inches(4.3)

for r_idx, row in enumerate(rows):
    for c_idx, val in enumerate(row):
        cell = tbl.cell(r_idx, c_idx)
        cell.text = ""
        tf = cell.text_frame
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
        tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        run = p.add_run(); run.text = val
        run.font.name = "Calibri"
        if r_idx == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
            run.font.size = Pt(13); run.font.bold = True
            run.font.color.rgb = WHITE
        else:
            # zebra rows
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r_idx % 2 else LIGHTNAVY
            run.font.size = Pt(11.5)
            if c_idx == 0:
                run.font.bold = True
                run.font.color.rgb = NAVY
            else:
                run.font.color.rgb = DARK

add_takeaway(s, "Silent oestrus and summer infertility are the dominant "
                "drivers of repeat breeding in buffaloes.")
add_footer(s, 21, TOTAL)
set_notes(s, ("Use this table to highlight why generic 'cow protocols' "
              "often under-perform in buffaloes — different physiology "
              "demands different management."))

# -----------------------------------------------------------
# 22. PREVENTION
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Prevention — Better Than Cure")
add_corner_badge(s, 22, TOTAL)

add_card(s, Inches(0.85), Inches(1.85), Inches(5.9), Inches(2.3),
         "Herd-Level Practices",
         ["Routine post-partum exam on Day 30–45",
          "Maintain calving interval 12–13 months",
          "Vaccinate against Brucella, IBR, BVD, Leptospira",
          "Bio-security; quarantine new animals"], head_color=NAVY)
add_card(s, Inches(7.0), Inches(1.85), Inches(5.9), Inches(2.3),
         "Cow / Buffalo-Level Practices",
         ["Balanced ration — energy, protein & minerals",
          "Body Condition Scoring at dry-off, calving, AI",
          "Heat detection 3× daily / pedometers / activity collars",
          "AI by trained technicians using fertile, properly thawed semen"],
         head_color=AMBER)
add_card(s, Inches(0.85), Inches(4.35), Inches(5.9), Inches(2.3),
         "Environment",
         ["Shade, fans, sprinklers / wallows in summer",
          "Comfortable, non-slippery flooring",
          "Adequate water and clean housing",
          "Reduce overcrowding & social stress"], head_color=TEAL)
add_card(s, Inches(7.0), Inches(4.35), Inches(5.9), Inches(2.3),
         "Records & Decisions",
         ["Maintain individual breeding records / software",
          "Identify chronic repeaters early",
          "Selective culling after 5–6 unsuccessful services",
          "Periodic herd fertility audit"], head_color=NAVY)

add_footer(s, 22, TOTAL)
set_notes(s, ("Prevention is a herd-level discipline. Single-cow "
              "interventions rarely work without supportive management."))

# -----------------------------------------------------------
# 23. RECENT ADVANCES
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Recent Advances")
add_corner_badge(s, 23, TOTAL)

add_bullets(s, [
    "Doppler ultrasonography of CL — early detection of luteal insufficiency.",
    "Embryo transfer (ET) in chronic repeat breeders — bypasses fertilization "
    "and early embryo problems.",
    "In-vitro fertilization (IVF / OPU-IVP) for valuable persistent repeaters.",
    "Sexed semen with optimised AI timing for high-genetic-merit cows.",
    "Genomic selection — identifying cows predisposed to early embryonic death.",
    "Activity-monitoring collars & AI-based heat detection systems.",
    "Anti-oxidant therapy (Vitamin E, Selenium, Astaxanthin) to improve "
    "oocyte and embryo quality.",
    "Targeted reproductive ultrasonography & uterine biopsy for refractory cases.",
], size=16)

add_takeaway(s, "ET and OPU-IVP can rescue genetically valuable "
                "repeat breeders that fail all conventional therapy.")
add_footer(s, 23, TOTAL)
set_notes(s, ("Recent advances are largely accessibility-limited in "
              "smallholder settings, but Doppler USG and activity meters "
              "are increasingly affordable."))

# -----------------------------------------------------------
# 24. CONCLUSION + REFERENCES + THANK YOU
# -----------------------------------------------------------
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Conclusion & References")
add_corner_badge(s, 24, TOTAL)

# Conclusion box
sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(0.90), Inches(1.90),
                        Inches(11.6), Inches(2.3))
sh.line.fill.background(); sh.fill.solid()
sh.fill.fore_color.rgb = LIGHTGRAY
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.85), Inches(1.85),
                         Inches(11.6), Inches(2.3))
box.line.color.rgb = AMBER; box.line.width = Pt(2)
box.fill.solid(); box.fill.fore_color.rgb = WHITE
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = Inches(0.3); tf.margin_right = Inches(0.3)
tf.margin_top = Inches(0.2)

p1 = tf.paragraphs[0]
r1 = p1.add_run(); r1.text = "Conclusion"
r1.font.size = Pt(18); r1.font.bold = True
r1.font.color.rgb = NAVY; r1.font.name = "Calibri"

for line in [
    "Repeat breeding is a multifactorial syndrome — a symptom rather than a "
    "disease.",
    "Successful management requires accurate diagnosis, correction of "
    "nutrition, infection and hormonal imbalance, plus excellent AI practices.",
    "Buffaloes need extra attention to silent oestrus, seasonality and "
    "summer heat-stress mitigation.",
    "Prevention through good husbandry and herd health is far more "
    "economical than treatment.",
]:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "▸  " + line
    r.font.size = Pt(13); r.font.color.rgb = DARK; r.font.name = "Calibri"
    p.space_after = Pt(2)

# References
ref_box = s.shapes.add_textbox(Inches(0.85), Inches(4.35),
                               Inches(11.6), Inches(2.4))
rtf = ref_box.text_frame; rtf.word_wrap = True
p = rtf.paragraphs[0]
r = p.add_run(); r.text = "References (selected)"
r.font.size = Pt(16); r.font.bold = True
r.font.color.rgb = NAVY; r.font.name = "Calibri"

refs = [
    "Roberts S.J. — Veterinary Obstetrics & Genital Diseases (Theriogenology), 3rd Ed.",
    "Arthur, Noakes, Pearson & Parkinson — Veterinary Reproduction & Obstetrics, 11th Ed.",
    "Hafez E.S.E. & Hafez B. — Reproduction in Farm Animals, 7th Ed.",
    "Peters A.R. & Ball P.J.H. — Reproduction in Cattle, 3rd Ed.",
    "Bartlett P.C. et al. — Repeat breeder syndrome in dairy cattle, JDS.",
    "Gustafsson H. & Emanuelson U. — Repeat breeding syndrome in Swedish dairy cattle, Acta Vet. Scand.",
    "Perera B.M.A.O. — Reproductive cycles of buffalo, Anim. Reprod. Sci.",
]
for ref in refs:
    p = rtf.add_paragraph()
    rr = p.add_run(); rr.text = "•  " + ref
    rr.font.size = Pt(11); rr.font.color.rgb = DARK; rr.font.name = "Calibri"
    p.space_after = Pt(2)

# Image credits
cred = s.shapes.add_textbox(Inches(0.85), Inches(6.75),
                            Inches(8.5), Inches(0.3))
cp = cred.text_frame.paragraphs[0]
cr = cp.add_run()
cr.text = ("Image credits: cow & buffalo photographs from Wikimedia Commons "
           "(CC BY-SA / public domain).")
cr.font.size = Pt(9); cr.font.italic = True
cr.font.color.rgb = GRAY; cr.font.name = "Calibri"

# Thank you ribbon
ty = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                        Inches(9.6), Inches(6.7),
                        Inches(3.13), Inches(0.45))
ty.line.fill.background(); ty.fill.solid()
ty.fill.fore_color.rgb = AMBER
ttf = ty.text_frame
ttf.margin_top = Inches(0.04); ttf.margin_bottom = Inches(0.04)
tp = ttf.paragraphs[0]; tp.alignment = PP_ALIGN.CENTER
tr = tp.add_run(); tr.text = "THANK YOU"
tr.font.size = Pt(16); tr.font.bold = True
tr.font.color.rgb = WHITE; tr.font.name = "Calibri"

set_notes(s, ("Wrap up. Restate that repeat breeding is preventable. "
              "Invite questions and discussion."))

# Save
out = os.path.join(HERE, "Repeat_Breeding.pptx")
prs.save(out)
print(f"Saved: {out}")
print(f"Total slides: {len(prs.slides)}")
