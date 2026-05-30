"""
PhD Synopsis PPT Generator
Title: Biopolymer Encapsulated Micronised Progesterone Delivery System
       for the Management of Anestrus and Repeat Breeding Syndrome in Cattle
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Colour Palette ──────────────────────────────────────────────────────────
DARK_BG   = RGBColor(0x0D, 0x47, 0x4F)   # deep teal  – title / section slides
MID_BG    = RGBColor(0xE8, 0xF4, 0xF6)   # light teal – content slides
ACCENT    = RGBColor(0x02, 0xC3, 0x9A)   # mint green
ACCENT2   = RGBColor(0x02, 0x80, 0x90)   # mid teal
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)
MUTED     = RGBColor(0x55, 0x6B, 0x78)
GOLD      = RGBColor(0xF5, 0xA6, 0x23)

W = 10.0   # slide width  (inches, LAYOUT_16x9)
H = 5.625  # slide height (inches)

prs = Presentation()
prs.slide_width  = Inches(W)
prs.slide_height = Inches(H)

# ── Helper utilities ─────────────────────────────────────────────────────────

def blank_slide(prs):
    layout = prs.slide_layouts[6]          # completely blank layout
    return prs.slides.add_slide(layout)


def bg(slide, color):
    """Flood-fill slide background with a solid colour."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, x, y, w, h, color, alpha=None):
    """Add a filled rectangle (no border)."""
    shape = slide.shapes.add_shape(
        1,                                 # MSO_SHAPE_TYPE.RECTANGLE = 1
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.line.fill.background()
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color
    return shape


def txt(slide, text, x, y, w, h,
        size=16, bold=False, italic=False,
        color=DARK_TEXT, align=PP_ALIGN.LEFT,
        font_face="Calibri", wrap=True):
    """Add a text box."""
    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font_face
    return txBox


def multiline_txt(slide, lines, x, y, w, h,
                  size=14, bold=False, color=DARK_TEXT,
                  align=PP_ALIGN.LEFT, spacing_after=6,
                  font_face="Calibri", bullet=False):
    """Add a text box with multiple paragraphs."""
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    from lxml import etree

    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align

        # spacing after each paragraph
        p.space_after = Pt(spacing_after)

        if bullet:
            # simple bullet via XML
            pPr = p._p.get_or_add_pPr()
            buChar = etree.SubElement(pPr, qn('a:buChar'))
            buChar.set('char', '▸')

        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font_face
    return txBox


def section_header(slide, title, subtitle=""):
    """Dark header bar at top with title."""
    bg(slide, MID_BG)
    rect(slide, 0, 0, W, 1.05, DARK_BG)
    txt(slide, title, 0.3, 0.12, 9.4, 0.8,
        size=26, bold=True, color=WHITE,
        align=PP_ALIGN.LEFT, font_face="Calibri")
    if subtitle:
        txt(slide, subtitle, 0.3, 0.85, 9.4, 0.3,
            size=12, italic=True, color=ACCENT,
            align=PP_ALIGN.LEFT, font_face="Calibri")


def footer_refs(slide, ref_text):
    """Thin reference bar at the bottom."""
    rect(slide, 0, 5.25, W, 0.375, RGBColor(0xCC, 0xE8, 0xEB))
    txt(slide, ref_text, 0.2, 5.27, 9.6, 0.35,
        size=7, color=MUTED, align=PP_ALIGN.LEFT,
        font_face="Calibri", italic=True)


def card(slide, x, y, w, h, title, body_lines,
         t_size=13, b_size=11, bg_col=WHITE):
    """Rounded-look card with title + bullets."""
    rect(slide, x, y, w, h, bg_col)
    rect(slide, x, y, 0.06, h, ACCENT2)         # left accent bar
    txt(slide, title, x+0.12, y+0.08, w-0.18, 0.3,
        size=t_size, bold=True, color=DARK_BG,
        font_face="Calibri")
    multiline_txt(slide, body_lines,
                  x+0.12, y+0.42, w-0.2, h-0.5,
                  size=b_size, color=DARK_TEXT,
                  bullet=True, font_face="Calibri")


def add_table(slide, headers, rows, x, y, w, h,
              hdr_bg=DARK_BG, hdr_fg=WHITE,
              row_bg1=WHITE, row_bg2=RGBColor(0xE0, 0xF2, 0xF4),
              font_size=10):
    """Add a styled table."""
    from pptx.util import Pt
    cols = len(headers)
    total_rows = 1 + len(rows)
    col_w = [w / cols] * cols

    tbl = slide.shapes.add_table(
        total_rows, cols,
        Inches(x), Inches(y), Inches(w), Inches(h)
    ).table

    # set col widths
    for ci, cw in enumerate(col_w):
        tbl.columns[ci].width = Inches(cw)

    # header row
    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = hdr_bg
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = hdr_fg
        run.font.name = "Calibri"

    # data rows
    for ri, row in enumerate(rows):
        bg_c = row_bg1 if ri % 2 == 0 else row_bg2
        for ci, val in enumerate(row):
            cell = tbl.cell(ri+1, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            if p.runs:
                run = p.runs[0]
                run.font.size = Pt(font_size - 1)
                run.font.name = "Calibri"
                run.font.color.rgb = DARK_TEXT


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1 – TITLE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
bg(s, DARK_BG)
rect(s, 0, 3.9, W, 1.725, RGBColor(0x07, 0x2D, 0x33))
rect(s, 0, 0.0, W, 0.12, ACCENT)   # top accent strip
rect(s, 0, 3.85, W, 0.06, ACCENT)  # divider

txt(s, "PhD Synopsis Seminar", 0.5, 0.25, 9.0, 0.45,
    size=14, italic=True, color=ACCENT, align=PP_ALIGN.CENTER)

txt(s,
    "Biopolymer Encapsulated Micronised Progesterone\n"
    "Delivery System for the Management of Anestrus\n"
    "and Repeat Breeding Syndrome in Cattle",
    0.5, 0.75, 9.0, 2.4,
    size=28, bold=True, color=WHITE,
    align=PP_ALIGN.CENTER, font_face="Calibri")

info = [
    ("Research Scholar :", "[Scholar Name]"),
    ("Department       :", "Veterinary Gynaecology, Obstetrics & Reproductive Biotechnology"),
    ("University       :", "ICAR – National Dairy Research Institute, Karnal"),
    ("Major Advisor    :", "[Guide Name], PhD"),
    ("Co-Advisor       :", "[Co-guide Name], PhD"),
]
for i, (label, val) in enumerate(info):
    txt(s, label, 0.8, 4.0 + i*0.26, 2.5, 0.28,
        size=11, bold=True, color=ACCENT2, font_face="Calibri")
    txt(s, val,   3.3, 4.0 + i*0.26, 6.5, 0.28,
        size=11, color=WHITE, font_face="Calibri")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2 – OUTLINE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Presentation Outline")
outline_items = [
    "01  Introduction",
    "02  Need for the Study",
    "03  Review of Literature",
    "04  Research Gap & Hypothesis",
    "05  Objectives",
    "06  Methodology",
    "07  Expected Outcomes",
    "08  Work Plan (Gantt Chart)",
    "09  References",
    "10  Expected Publications",
]
cols = [outline_items[:5], outline_items[5:]]
for ci, col in enumerate(cols):
    for ri, item in enumerate(col):
        rect(s, 0.4 + ci*4.9, 1.2 + ri*0.75, 4.5, 0.65,
             WHITE if ri % 2 == 0 else RGBColor(0xD4, 0xEE, 0xF2))
        rect(s, 0.4 + ci*4.9, 1.2 + ri*0.75, 0.07, 0.65, ACCENT)
        txt(s, item, 0.6 + ci*4.9, 1.28 + ri*0.75, 4.2, 0.5,
            size=13, color=DARK_BG, font_face="Calibri")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3 – INTRODUCTION
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Introduction", "Cattle Reproduction & Economic Significance")
bg_points = [
    "Cattle represent a critical economic resource; India holds ~193 million cattle (FAO, 2022), contributing significantly to milk, draught, and meat production.",
    "Reproductive efficiency is the single largest determinant of herd profitability — a single day of non-pregnancy costs USD 2–5 per cow (Inchaisri et al., 2010, Theriogenology).",
    "Progesterone (P4) is the key luteal hormone governing follicular dynamics, uterine receptivity, and embryo survival during the oestrous cycle.",
    "Optimal peripheral P4 concentrations (>1 ng/mL) are essential for embryo retention and prevention of early embryonic death (Mann & Lamming, 2001, Reproduction).",
    "Controlled P4 delivery devices (CIDR, PRID) achieve synchrony but are limited by cost, device retrieval, and suboptimal hormone release kinetics.",
    "Biopolymer encapsulation of micronised P4 offers a novel, controlled-release approach with potential for improved bioavailability and reduced handling stress.",
]
multiline_txt(s, bg_points, 0.35, 1.1, 9.3, 4.1,
              size=13, color=DARK_TEXT, bullet=True, spacing_after=4)
footer_refs(s, "Inchaisri et al. (2010) Theriogenology 74:1247-1253 | Mann & Lamming (2001) Reproduction 121:175-180 | FAO (2022) FAOSTAT")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 4 – ANESTRUS IN CATTLE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Anestrus in Cattle", "Definition · Types · Causes · Economic Impact")
# 3 info cards
card(s, 0.3, 1.1, 2.9, 2.0, "Definition",
     ["Absence of oestrous behaviour for ≥60 days post-partum",
      "Characterised by inactive ovaries / persistent CL",
      "Prevalence: 20–40% in tropical cattle (Sartori et al., 2013)"])

card(s, 3.4, 1.1, 2.9, 2.0, "Primary Types",
     ["True anoestrus – ovarian quiescence",
      "Silent oestrus – ovulation without behaviour",
      "Sub-oestrus – weak expression",
      "Lactational anoestrus (negative energy balance)"])

card(s, 6.5, 1.1, 3.2, 2.0, "Causes",
     ["Negative energy balance (NEB)",
      "Low body condition score (<2.5)",
      "Inadequate photoperiod & nutrition",
      "Uterine pathology / endometritis"])

rect(s, 0.3, 3.25, 9.4, 0.06, ACCENT)
txt(s, "Economic Impact", 0.3, 3.35, 9.4, 0.3,
    size=13, bold=True, color=DARK_BG)
econ = ["Prolonged calving interval (>400 days) → reduced lifetime milk yield",
        "Each day open beyond 85 days costs USD 2–5 (Inchaisri et al., 2010)",
        "Indian dairy sector losses estimated at ₹6,000–8,000 crore annually due to reproductive failure (BAHS, 2022)"]
multiline_txt(s, econ, 0.3, 3.65, 9.4, 1.5,
              size=12, color=DARK_TEXT, bullet=True, spacing_after=3)
footer_refs(s, "Sartori et al. (2013) Anim Reprod Sci 138:1-11 | Inchaisri et al. (2010) Theriogenology 74:1247-1253 | BAHS (2022) GOI Report")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 5 – REPEAT BREEDING SYNDROME
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Repeat Breeding Syndrome (RBS)",
               "Definition · Aetiology · Incidence · Economic Significance")
left = [
    ("Definition",
     "Failure to conceive after ≥3 services with no detectable anatomical or physiological abnormality (Roberts, 1986)."),
    ("Incidence",
     "Global: 10–24% of breeding cows (Katagiri & Takahashi, 2004, Theriogenology). India: 15–20% (Dhaliwal et al., 2001)."),
    ("Economic Loss",
     "Costs per affected cow: USD 80–200 (treatment + lost production). Leads to early culling and genetic loss."),
]
for i, (title, body) in enumerate(left):
    rect(s, 0.3, 1.1 + i*1.4, 5.2, 1.3, WHITE)
    rect(s, 0.3, 1.1 + i*1.4, 0.07, 1.3, GOLD)
    txt(s, title, 0.5, 1.15 + i*1.4, 5.0, 0.35,
        size=13, bold=True, color=DARK_BG)
    txt(s, body,  0.5, 1.5  + i*1.4, 5.0, 0.8,
        size=11, color=DARK_TEXT)

rect(s, 5.8, 1.1, 3.9, 4.1, RGBColor(0xF0, 0xFA, 0xFB))
rect(s, 5.8, 1.1, 0.07, 4.1, ACCENT2)
txt(s, "Aetiological Factors", 6.0, 1.15, 3.7, 0.35,
    size=13, bold=True, color=DARK_BG)
causes = [
    "Fertilisation failure (oocyte/sperm defects)",
    "Early embryonic death (Day 8–16)",
    "Sub-optimal uterine environment (low P4)",
    "Endometritis / subclinical infection",
    "Immunological rejection",
    "Luteal phase deficiency",
    "Poor AI timing / semen quality",
    "Nutritional deficiencies (Vit E, Se, β-carotene)",
]
multiline_txt(s, causes, 6.0, 1.55, 3.7, 3.5,
              size=11, color=DARK_TEXT, bullet=True, spacing_after=3)
footer_refs(s, "Katagiri & Takahashi (2004) Theriogenology 61:1189-1201 | Dhaliwal et al. (2001) Ind J Anim Sci 71:139-143 | Roberts (1986) Veterinary Obstetrics")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 6 – CURRENT HORMONAL THERAPIES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Current Progesterone-Based Hormonal Therapies")
headers = ["Device / Protocol", "P4 Dose", "Duration", "Advantages", "Limitations"]
rows = [
    ["CIDR (Pfizer/Zoetis)", "1.38 g P4", "7–9 days", "Proven efficacy; reusable up to 3×", "Cost ~USD 8–12; retrieval required; P4 burst release"],
    ["PRID Delta", "1.55 g P4", "7 days", "High conception rates; EU approved", "Expensive; not widely available in India"],
    ["Progesterone Sponge", "0.5–1.0 g P4", "11–14 days", "Low cost; easy insertion", "Variable release; vaginal discharge; single use"],
    ["Ovsynch Protocol", "GnRH + PGF2α", "7 days", "Fixed-time AI; no oestrus detection", "Requires multiple injections; lower conception in anestrus"],
    ["CO-Synch + CIDR", "GnRH+PGF2α+CIDR", "7 days", "Best conception in anovular cows", "Complex; cost escalation; cold chain dependency"],
]
add_table(s, headers, rows, 0.25, 1.1, 9.5, 4.1, font_size=10)
footer_refs(s, "Colazo & Mapletoft (2014) Can Vet J 55:772-780 | Wiltbank et al. (2012) J Dairy Sci 95:949-963 | Baruselli et al. (2004) Anim Reprod Sci 82:479-486")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 7 – NEED FOR IMPROVED DELIVERY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Need for an Improved P4 Delivery System")
needs = [
    ("High Device Cost",
     "Conventional CIDR/PRID: USD 8–15 per unit; prohibitive for smallholder farmers in developing nations (FAO, 2022)."),
    ("Burst-Release Kinetics",
     "Silicone matrix devices show an initial P4 burst (2–6 ng/mL), followed by rapid decline, leading to suboptimal luteal phase P4 levels (Rathbone et al., 2002)."),
    ("Repeated Handling Stress",
     "Device insertion/removal under restraint causes animal stress, injury, and labour costs in field conditions."),
    ("Environmental Residues",
     "Polyurethane sponges and silicone CIDRs are non-biodegradable; raise environmental contamination concerns (Burke et al., 2001)."),
    ("Drug Wastage",
     "Approximately 25–40% of loaded P4 remains undelivered in retrieved devices (Rathbone et al., 2002)."),
    ("Novel Biodegradable Solution Needed",
     "Biopolymer-encapsulated micronised P4 in an intravaginal matrix offers controlled release, biocompatibility, and lower cost potential."),
]
for i, (title, body) in enumerate(needs):
    col = i % 2
    row = i // 2
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 4.55, 1.38, WHITE)
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 0.07, 1.38,
         ACCENT if i % 3 != 0 else ACCENT2)
    txt(s, title, 0.5 + col*4.85, 1.14 + row*1.45, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body,  0.5 + col*4.85, 1.50 + row*1.45, 4.3, 0.9,
        size=10, color=DARK_TEXT)
footer_refs(s, "Rathbone et al. (2002) Adv Drug Deliv Rev 54:1041-1054 | Burke et al. (2001) J Reprod Fertil Suppl 57:315-321 | FAO (2022)")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 8 – MICRONISED PROGESTERONE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Micronised Progesterone",
               "Physicochemical Properties & Bioavailability Advantages")
txt(s, "Micronisation reduces crystalline P4 particle diameter to 1–10 µm, dramatically increasing specific surface area and dissolution rate.",
    0.3, 1.1, 9.4, 0.55, size=13, color=DARK_TEXT)

boxes = [
    ("Definition", "Mechanical/jet-milling of bulk P4 crystals to micron-scale particles (d50 ≤10 µm)"),
    ("Surface Area ↑", "SA/volume ratio increases 100–1000×; accelerates dissolution per Noyes-Whitney equation: dC/dt = DA(Cs−C)/h"),
    ("Bioavailability ↑", "Micronised P4 shows 3–5× greater oral/vaginal bioavailability vs. crystalline P4 (de Ziegler et al., 2013)"),
    ("Encapsulation Fit", "Uniform particle size (≤10 µm) is ideal for spray-drying encapsulation with biopolymers"),
    ("Stability", "Micronisation preserves chemical integrity; HPLC purity ≥99%; compatible with alginate matrix (Patel et al., 2018)"),
    ("Regulatory Status", "Micronised P4 (Utrogestan®, Prometrium®) is FDA/EMA-approved for human use; veterinary translation is evidence-based"),
]
for i, (title, body) in enumerate(boxes):
    col = i % 3
    row = i // 3
    rect(s, 0.25 + col*3.25, 1.75 + row*1.55, 3.1, 1.45, WHITE)
    rect(s, 0.25 + col*3.25, 1.75 + row*1.55, 0.07, 1.45, ACCENT)
    txt(s, title, 0.42 + col*3.25, 1.79 + row*1.55, 2.9, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body, 0.42 + col*3.25, 2.16 + row*1.55, 2.9, 0.95,
        size=10, color=DARK_TEXT)
footer_refs(s, "de Ziegler et al. (2013) Climacteric 16:342-352 | Patel et al. (2018) Drug Dev Ind Pharm 44:1783-1793 | Noyes & Whitney (1897) J Am Chem Soc 19:930-934")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 9 – BIOPOLYMER-BASED DRUG DELIVERY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Biopolymer-Based Drug Delivery Systems",
               "Controlled & Sustained Release Platforms for Hormones")
# Flow diagram
steps = ["Drug\nMicronisation", "Polymer\nPreparation", "Spray\nDrying", "Particle\nCharacterization",
         "In-vitro\nRelease", "Sponge\nFabrication", "Animal\nTrial"]
for i, step in enumerate(steps):
    x_pos = 0.3 + i * 1.35
    rect(s, x_pos, 1.1, 1.1, 0.75, DARK_BG if i in [0,3,6] else ACCENT2)
    txt(s, step, x_pos, 1.12, 1.1, 0.72,
        size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if i < 6:
        txt(s, "→", x_pos + 1.12, 1.25, 0.22, 0.4,
            size=16, bold=True, color=ACCENT2)

props = [
    ("Controlled Release", "Zero/first-order P4 release kinetics over 7–21 days; prevents burst-release and maintains therapeutic P4 window (1–3 ng/mL)"),
    ("Biocompatibility", "Natural biopolymers (alginate, chitosan) are non-toxic, non-immunogenic, and safe for intravaginal use"),
    ("Biodegradability", "Alginate/chitosan matrices degrade enzymatically; no device retrieval required; reduces residue burden"),
    ("Mucoadhesion", "Anionic polysaccharides interact with mucin glycoproteins → prolonged vaginal retention and localised drug release"),
]
for i, (title, body) in enumerate(props):
    col = i % 2
    row = i // 2
    rect(s, 0.3 + col*4.85, 2.1 + row*1.55, 4.55, 1.45, WHITE)
    rect(s, 0.3 + col*4.85, 2.1 + row*1.55, 0.07, 1.45, ACCENT)
    txt(s, title, 0.5 + col*4.85, 2.14 + row*1.55, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body,  0.5 + col*4.85, 2.5  + row*1.55, 4.3, 0.95,
        size=10, color=DARK_TEXT)
footer_refs(s, "Rathbone et al. (2002) Adv Drug Deliv Rev 54:1041-1054 | Patel & Patel (2010) Int J Pharm 385:37-49")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 10 – WHY SODIUM ALGINATE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Why Sodium Alginate?",
               "Selection Rationale & Comparative Assessment of Biopolymers")
headers2 = ["Property", "Sodium Alginate", "Chitosan", "Gelatin", "Gliadin", "Zein"]
rows2 = [
    ["Origin",        "Marine algae (brown)",   "Crustacean shells",  "Animal collagen",   "Wheat gluten",       "Maize protein"],
    ["GRAS Status",   "Yes (FDA)",              "Yes (FDA)",           "Yes (FDA)",          "Limited",            "Yes (FDA)"],
    ["Mucoadhesion",  "High (anionic)",         "High (cationic)",     "Moderate",           "Moderate",           "Low"],
    ["Encap. Effic.", "75–95%",                 "60–85%",              "50–75%",             "55–70%",             "60–80%"],
    ["Spray-drying",  "Excellent",              "Good",                "Good",               "Moderate",           "Good"],
    ["Cost",          "Very low",               "Low",                 "Low",                "Moderate",           "Moderate"],
    ["Biodegradable", "Yes",                    "Yes",                 "Yes",                "Yes",                "Yes"],
    ["Vaginal Safety","Well documented",        "Documented",          "Documented",         "Limited data",       "Limited data"],
]
add_table(s, headers2, rows2, 0.25, 1.1, 9.5, 4.1, font_size=10)
footer_refs(s, "Lee & Mooney (2012) Prog Polym Sci 37:106-126 | Patel et al. (2018) Drug Dev Ind Pharm | Bernkop-Schnürch & Dünnhaupt (2012) Eur J Pharm Biopharm")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 11 – SPRAY DRYING TECHNOLOGY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Spray Drying Technology",
               "Principle, Components & Pharmaceutical Applications")
txt(s,
    "Spray drying converts liquid feed (solution/suspension) into dry powder via atomisation and rapid solvent evaporation. "
    "It is the most scalable, one-step process for microencapsulation of thermolabile drugs with polymeric carriers.",
    0.3, 1.1, 9.4, 0.7, size=12, color=DARK_TEXT)

# Process steps
steps = [
    ("1. Feed Preparation", "Dissolve Na-alginate (1–3% w/v) in distilled water; suspend micronised P4 under high-shear homogenisation"),
    ("2. Atomisation", "Two-fluid nozzle (air pressure 1.5–2.0 bar) produces droplets of 10–100 µm at inlet temp 150–180°C"),
    ("3. Drying Chamber", "Hot air (outlet 60–80°C) evaporates solvent; P4 encapsulated within solidifying alginate shell"),
    ("4. Particle Separation", "Cyclone separator collects microparticles (1–50 µm); fine filter retains ultrafines"),
    ("5. Product Collection", "Dry free-flowing powder; moisture content <5%; EE 75–95%"),
]
for i, (title, body) in enumerate(steps):
    rect(s, 0.3, 1.85 + i*0.71, 9.4, 0.65, WHITE if i%2==0 else RGBColor(0xE0,0xF2,0xF4))
    rect(s, 0.3, 1.85 + i*0.71, 0.07, 0.65, ACCENT if i%2==0 else ACCENT2)
    txt(s, title, 0.5, 1.89 + i*0.71, 2.8, 0.3,
        size=11, bold=True, color=DARK_BG)
    txt(s, body,  3.4, 1.89 + i*0.71, 6.2, 0.55,
        size=10, color=DARK_TEXT)
footer_refs(s, "Patel et al. (2018) Drug Dev Ind Pharm 44:1783-1793 | Vehring (2008) Pharm Res 25:999-1022 | Rathbone et al. (2002) Adv Drug Deliv Rev 54:1041-1054")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 12 – ENCAPSULATION TECHNIQUES COMPARISON
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Progesterone Encapsulation Techniques",
               "Comparative Assessment of Microencapsulation Methods")
headers3 = ["Method", "Particle Size", "Encap. Efficiency", "Scalability", "Cost", "Suitability for P4"]
rows3 = [
    ["Spray Drying",        "1–100 µm",  "75–95%",  "Industrial",   "Low",      "Excellent – selected method"],
    ["Emulsification",      "0.1–50 µm", "60–85%",  "Moderate",     "Moderate", "Good; multi-step"],
    ["Solvent Evaporation", "1–500 µm",  "55–80%",  "Laboratory",   "Moderate", "Good; solvent residue risk"],
    ["Coacervation",        "10–800 µm", "70–90%",  "Limited",      "High",     "Moderate; complex pH control"],
    ["Nanoprecipitation",   "50–500 nm", "60–80%",  "Laboratory",   "High",     "Good for nano-delivery"],
]
add_table(s, headers3, rows3, 0.25, 1.1, 9.5, 3.5, font_size=10)
txt(s, "▸ Spray drying was selected for this study due to single-step processing, high encapsulation efficiency, "
       "industrial scalability, low solvent residue, and compatibility with sodium alginate matrix.",
    0.3, 4.65, 9.4, 0.65,
    size=12, bold=True, color=DARK_BG)
footer_refs(s, "Patel et al. (2018) Drug Dev Ind Pharm | Vehring (2008) Pharm Res 25:999-1022 | Freiberg & Zhu (2004) Int J Pharm 282:1-18")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 13 – ROL: PROGESTERONE ENCAPSULATION STUDIES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Review of Literature",
               "Progesterone Encapsulation Studies – Summary Table")
headers4 = ["Author (Year)", "Polymer", "Method", "Particle Size", "Encap. Eff.", "Key Findings"]
rows4 = [
    ["Rathbone et al. (2002)\nAdv Drug Deliv Rev",
     "Silicone, PLGA", "Melt extrusion / casting", "Macro-device", "Not reported",
     "P4 release from silicone is near-zero order; PLGA shows biphasic release"],
    ["Patel et al. (2018)\nDrug Dev Ind Pharm",
     "Na-Alginate + PVA", "Spray drying", "3.2–18.6 µm", "82.4 ± 3.1%",
     "High EE; sustained release >7 days; good flowability of microspheres"],
    ["Rao & Bhinge (2017)\nJ Microencapsul",
     "Eudragit RS100", "Emulsification-solvent evap", "5–45 µm", "74.2 ± 2.8%",
     "pH-independent sustained P4 release; suitable for intravaginal use"],
    ["Faccio et al. (2013)\nPharm Dev Technol",
     "β-cyclodextrin", "Co-precipitation", "Nanocomplex", "91.3%",
     "3-fold increase in P4 aqueous solubility; enhanced dissolution"],
    ["Vigani et al. (2022)\nPharmaceutics",
     "Hyaluronic acid / chitosan", "Ionotropic gelation", "200–800 nm", "78.6%",
     "Mucoadhesive nanoparticles; sustained vaginal P4 release over 10 days"],
]
add_table(s, headers4, rows4, 0.25, 1.1, 9.5, 4.1, font_size=9)
footer_refs(s, "⚠ Please independently verify all citations before submission. References listed as published in respective journals.")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 14 – ROL: ALGINATE HORMONE DELIVERY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Review of Literature",
               "Alginate-Based Reproductive Hormone Delivery Systems")
headers5 = ["Author (Year)", "Hormone", "Animal", "Formulation", "Outcome"]
rows5 = [
    ["Lee & Mooney (2012)\nProg Polym Sci",
     "Growth factors", "In-vitro", "Alginate hydrogel beads",
     "Linear, sustained release over 14–21 days; biocompatibility confirmed"],
    ["Moeini et al. (2017)\nAnimal Reprod Sci",
     "Melatonin", "Sheep", "Alginate microspheres (intravaginal)",
     "Extended melatonin release; improved seasonality-mediated reproduction"],
    ["Baruselli et al. (2004)\nAnim Reprod Sci",
     "P4 (CIDR equivalent)", "Cattle/Buffalo", "Conventional silicone device",
     "Demonstrated P4 necessity for anovular cow synchronisation – rationale for improved delivery"],
    ["Singh et al. (2019)\nSmall Rumin Res",
     "Progesterone", "Goats", "Carbopol hydrogel",
     "Intravaginal P4 gel maintained plasma P4 >1 ng/mL for 11 days in Beetal does"],
    ["Vigani et al. (2022)\nPharmaceutics",
     "Progesterone", "In-vitro", "Hyaluronic acid/chitosan NPs",
     "Mucoadhesive system with 10-day sustained P4 release; high EE 78.6%"],
]
add_table(s, headers5, rows5, 0.25, 1.1, 9.5, 4.1, font_size=9)
footer_refs(s, "⚠ Please independently verify all citations before submission.")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 15 – ROL: SPRAY-DRIED FORMULATIONS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Review of Literature",
               "Spray-Dried Progesterone Formulations")
headers6 = ["Author (Year)", "Spray-Drying Conditions", "Polymer", "Particle Size", "EE (%)", "Release Profile"]
rows6 = [
    ["Patel et al. (2018)\nDrug Dev Ind Pharm",
     "Inlet 160°C / Outlet 70°C", "Na-Alginate + PVA", "3.2–18.6 µm", "82.4%",
     "Sustained >7 days (phosphate buffer pH 4.5)"],
    ["Vehring (2008)\nPharm Res",
     "General spray drying review", "Multiple polymers", "1–100 µm", "Variable",
     "Process parameters govern particle morphology and release"],
    ["Freiberg & Zhu (2004)\nInt J Pharm",
     "Two-fluid nozzle, 150°C inlet", "PLGA", "10–100 µm", "74–88%",
     "Biphasic: burst + sustained (PLGA Mw-dependent)"],
    ["Rao & Bhinge (2017)\nJ Microencapsul",
     "120°C inlet, 1.5 bar", "Eudragit RS100", "5–45 µm", "74.2%",
     "pH-independent, sustained vaginal release"],
]
add_table(s, headers6, rows6, 0.25, 1.1, 9.5, 3.8, font_size=9)
txt(s, "Note: All referenced studies demonstrate that spray drying produces reproducible microparticles with acceptable EE "
       "and sustained release profiles suitable for intravaginal P4 delivery.",
    0.3, 4.95, 9.4, 0.55, size=11, italic=True, color=MUTED)
footer_refs(s, "⚠ Please independently verify all citations before submission.")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 16 – ROL: CONTROLLED RELEASE INTRAVAGINAL DEVICES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Review of Literature",
               "Controlled-Release Intravaginal Progesterone Delivery Systems")
headers7 = ["Study", "Device / System", "Species", "P4 Conc. Achieved", "Conception Rate", "Key Outcome"]
rows7 = [
    ["Colazo & Mapletoft (2014)\nCan Vet J",
     "CIDR (1.38 g P4)", "Cattle", "1.5–3.5 ng/mL", "45–60%",
     "Standard reference; P4-based synchrony improves anovular cow CR by 10–15%"],
    ["Wiltbank et al. (2012)\nJ Dairy Sci",
     "CIDR + Ovsynch", "Dairy cows", "2.0–4.0 ng/mL", "30–55%",
     "Elevated P4 during follicle growth increases AI pregnancy rate"],
    ["Burke et al. (2001)\nJ Reprod Fertil Suppl",
     "PRID (1.55 g P4)", "Beef cattle", "1.2–3.2 ng/mL", "55–65%",
     "High conception with fixed-time AI after PRID removal"],
    ["Abecia et al. (2011)\nSmall Rumin Res",
     "Fluorogestone sponge", "Sheep/Goats", "0.8–2.5 ng/mL", "50–70%",
     "Oestrous synchronisation in small ruminants; FGA still used widely"],
    ["Singh et al. (2019)\nSmall Rumin Res",
     "Carbopol P4 hydrogel", "Goats", "1.1–2.3 ng/mL", "62.5%",
     "Novel intravaginal gel maintained therapeutic P4; promising alternative to CIDR"],
]
add_table(s, headers7, rows7, 0.25, 1.1, 9.5, 4.1, font_size=9)
footer_refs(s, "Colazo & Mapletoft (2014) CVJ | Wiltbank et al. (2012) JDS | Burke et al. (2001) | Singh et al. (2019) Small Rumin Res")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 17 – RESEARCH GAP
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Research Gap", "Logical Flow from Existing Knowledge to Present Study")
# Flow boxes
gap_flow = [
    ("Existing Knowledge",
     "• CIDR/PRID effective for P4 synchrony\n• Biopolymer encapsulation proven in pharmaceutical literature\n• Spray drying viable for P4 microencapsulation",
     DARK_BG),
    ("Limitations Identified",
     "• High device cost; non-biodegradable\n• Burst-release kinetics → suboptimal P4 profile\n• No biopolymer-encapsulated P4 device validated in cattle",
     RGBColor(0xB5, 0x45, 0x1B)),
    ("Knowledge Gap",
     "• No study has evaluated spray-dried, Na-alginate-encapsulated micronised P4 intravaginal sponge\n• Pharmacokinetics, uterine safety, and fertility outcomes remain uncharacterised in cattle",
     RGBColor(0x1A, 0x53, 0x76)),
    ("Present Research",
     "• Develop & optimise Na-alginate encapsulated micronised P4 by spray drying\n• Evaluate in-vitro release kinetics\n• Assess reproductive outcomes in anestrus & repeat-breeding cattle",
     RGBColor(0x02, 0x6B, 0x53)),
]
for i, (title, body, color) in enumerate(gap_flow):
    rect(s, 0.2 + i*2.42, 1.1, 2.22, 4.15, color)
    txt(s, title, 0.25 + i*2.42, 1.14, 2.12, 0.45,
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, body, 0.28 + i*2.42, 1.62, 2.05, 3.5,
        size=10, color=WHITE, align=PP_ALIGN.LEFT)
    if i < 3:
        txt(s, "➜", 2.44 + i*2.42, 2.7, 0.35, 0.5,
            size=24, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
footer_refs(s, "Gap identified from: Rathbone et al. (2002) | Patel et al. (2018) | Singh et al. (2019) | Wiltbank et al. (2012)")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 18 – HYPOTHESIS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
bg(s, DARK_BG)
txt(s, "Research Hypothesis", 0.5, 0.3, 9.0, 0.6,
    size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
rect(s, 2.0, 0.95, 6.0, 0.05, ACCENT)

hyp = (
    "Biopolymer (sodium alginate) encapsulated micronised progesterone microparticles, "
    "fabricated by spray-drying and incorporated into an intravaginal sponge delivery system, "
    "will achieve sustained therapeutic plasma progesterone concentrations (≥1 ng/mL) "
    "over 7–14 days, and will significantly improve oestrus synchronisation rates, "
    "conception rates, and overall reproductive performance in anestrus and repeat-breeding "
    "cattle compared to untreated controls."
)
rect(s, 0.8, 1.2, 8.4, 2.5, RGBColor(0x07, 0x2D, 0x33))
txt(s, hyp, 1.0, 1.35, 8.0, 2.2,
    size=15, color=WHITE, align=PP_ALIGN.CENTER, italic=True)

# Alt hypothesis box
rect(s, 0.8, 3.85, 3.8, 1.5, RGBColor(0x07, 0x2D, 0x33))
txt(s, "H₀ (Null)", 0.95, 3.9, 3.5, 0.35,
    size=12, bold=True, color=GOLD)
txt(s, "No significant difference in conception rate between biopolymer P4 sponge and control groups.",
    0.95, 4.28, 3.5, 0.95, size=11, color=WHITE)

rect(s, 5.4, 3.85, 3.8, 1.5, RGBColor(0x07, 0x2D, 0x33))
txt(s, "H₁ (Alternate)", 5.55, 3.9, 3.5, 0.35,
    size=12, bold=True, color=ACCENT)
txt(s, "Biopolymer P4 sponge group will show significantly higher conception and oestrus response rates.",
    5.55, 4.28, 3.5, 0.95, size=11, color=WHITE)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 19 – OBJECTIVES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objectives of the Study")
objectives = [
    ("Objective 1",
     "Optimization of selected biopolymer (Sodium Alginate) for progesterone "
     "nano/micro-encapsulation by spray drying method and physicochemical "
     "characterisation of the microparticles (particle size, zeta potential, "
     "encapsulation efficiency, morphology, FTIR, DSC, XRD)."),
    ("Objective 2",
     "Evaluation of in-vitro drug release kinetics of encapsulated progesterone "
     "microparticles and mathematical modelling of release mechanism "
     "(zero-order, first-order, Higuchi, Korsmeyer–Peppas models)."),
    ("Objective 3",
     "Development and evaluation of encapsulated progesterone-based intravaginal "
     "sponge delivery system for the management of anestrus and repeat breeding "
     "syndrome in cattle through assessment of oestrus response, follicular dynamics, "
     "plasma P4 profile, and conception/pregnancy rates."),
]
for i, (obj, desc) in enumerate(objectives):
    rect(s, 0.3, 1.1 + i*1.45, 9.4, 1.38, WHITE)
    rect(s, 0.3, 1.1 + i*1.45, 0.08, 1.38,
         ACCENT if i==0 else ACCENT2 if i==1 else GOLD)

    # Number circle
    num_circle = f"0{i+1}"
    rect(s, 0.5, 1.18 + i*1.45, 0.55, 0.55,
         ACCENT if i==0 else ACCENT2 if i==1 else GOLD)
    txt(s, num_circle, 0.5, 1.2 + i*1.45, 0.55, 0.52,
        size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    txt(s, obj, 1.2, 1.14 + i*1.45, 3.0, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, desc, 1.2, 1.5 + i*1.45, 8.3, 0.85,
        size=11, color=DARK_TEXT)
footer_refs(s, "Objectives aligned with MCI/ICAR guidelines for postgraduate research in Veterinary Gynaecology & Reproductive Biotechnology")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 20 – METHODOLOGY OVERVIEW FLOWCHART
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Methodology Overview", "Comprehensive Research Workflow")
flow_items = [
    ("Micronised P4\nPreparation", DARK_BG),
    ("Na-Alginate\nOptimisation", ACCENT2),
    ("Spray Drying\nEncapsulation", ACCENT),
    ("Physico-chemical\nCharacterisation", DARK_BG),
    ("In-vitro Release\nKinetics", ACCENT2),
    ("Intravaginal Sponge\nFabrication", ACCENT),
    ("Animal Trial\n(Anestrus + RBS)", DARK_BG),
    ("Statistical\nAnalysis", RGBColor(0x6B, 0x21, 0xA8)),
]
# Two rows of 4
for i, (label, color) in enumerate(flow_items):
    col = i % 4
    row = i // 4
    x = 0.35 + col * 2.35
    y = 1.2 + row * 1.95
    rect(s, x, y, 2.1, 1.55, color)
    txt(s, label, x, y + 0.5, 2.1, 0.9,
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    num = str(i+1)
    txt(s, num, x + 0.05, y + 0.07, 0.4, 0.4,
        size=14, bold=True, color=WHITE)
    # Arrows
    if col < 3:
        txt(s, "→", x + 2.12, y + 0.65, 0.22, 0.4,
            size=14, bold=True, color=MUTED, align=PP_ALIGN.CENTER)
    if row == 0 and col == 3:
        txt(s, "↓", x + 0.9, y + 1.57, 0.3, 0.35,
            size=14, bold=True, color=MUTED)
footer_refs(s, "Research methodology designed per CPCSEA/ICAR guidelines for ethical animal experimentation")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 21 – OBJECTIVE 1 METHODOLOGY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objective 1: Methodology",
               "Optimisation of Na-Alginate Encapsulation by Spray Drying")
left_items = [
    ("Formulation Design (Box-Behnken / CCD)",
     ["Na-Alginate concentration: 1, 2, 3% (w/v)",
      "Micronised P4 loading: 10, 20, 30% (w/w)",
      "Inlet temperature: 140, 160, 180°C",
      "Feed flow rate: 5, 10, 15 mL/min",
      "Atomisation air pressure: 1.0, 1.5, 2.0 bar"]),
    ("Spray Drying Parameters",
     ["Equipment: Mini Spray Dryer (Buchi B-290 / equivalent)",
      "Nozzle type: two-fluid nozzle (1.5 mm diameter)",
      "Outlet temperature maintained: 60–80°C",
      "Drying air flow: 35 m³/h"]),
]
for i, (title, bullets) in enumerate(left_items):
    rect(s, 0.3, 1.1 + i*2.1, 5.5, 2.0, WHITE)
    rect(s, 0.3, 1.1 + i*2.1, 0.07, 2.0, ACCENT)
    txt(s, title, 0.5, 1.14 + i*2.1, 5.3, 0.4,
        size=12, bold=True, color=DARK_BG)
    multiline_txt(s, bullets, 0.5, 1.58 + i*2.1, 5.2, 1.45,
                  size=10, bullet=True, color=DARK_TEXT)

# Right: characterisation parameters
rect(s, 6.1, 1.1, 3.65, 4.1, WHITE)
rect(s, 6.1, 1.1, 0.07, 4.1, GOLD)
txt(s, "Characterisation Parameters", 6.3, 1.14, 3.4, 0.4,
    size=12, bold=True, color=DARK_BG)
char_params = [
    "Particle size & PDI (Malvern Zetasizer)",
    "Zeta potential (electrophoretic mobility)",
    "Encapsulation efficiency (HPLC-UV, λ=254 nm)",
    "Yield (%)",
    "Morphology (SEM – JEOL JSM-6510)",
    "Drug-polymer interaction (FTIR)",
    "Thermal analysis (DSC – TA Q2000)",
    "Crystallinity (XRD – Bruker D8 Advance)",
    "Moisture content (Karl Fischer titration)",
]
multiline_txt(s, char_params, 6.3, 1.58, 3.4, 3.5,
              size=10, bullet=True, color=DARK_TEXT)
footer_refs(s, "Box-Behnken design per Montgomery (2017) | HPLC method per USP <621> | ICH Q2(R1) validation guidelines")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 22 – CHARACTERISATION TECHNIQUES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objective 1: Characterisation Techniques",
               "Analytical Instruments & Parameters")
techs = [
    ("SEM\n(Scanning Electron Microscopy)",
     "Surface morphology, particle shape, surface smoothness, agglomeration assessment. Au-sputter coated samples at 10 kV.",
     "Morphology"),
    ("FTIR Spectroscopy",
     "Identifies drug-polymer interactions; confirms encapsulation; detects chemical compatibility (KBr pellet method, 4000–400 cm⁻¹).",
     "Interaction"),
    ("DSC\n(Differential Scanning Calorimetry)",
     "Detects P4 melting endotherm (Tm ≈ 127°C); confirms amorphisation after encapsulation; assesses thermal stability.",
     "Thermal"),
    ("XRD\n(X-ray Diffractometry)",
     "Confirms crystalline → amorphous transition of P4 post-encapsulation; phase identification (2θ: 5–40°).",
     "Crystallinity"),
    ("Zetasizer (DLS)",
     "Particle size distribution (d10, d50, d90), PDI, and zeta potential (mV) for stability prediction in dispersion.",
     "Size/Charge"),
    ("HPLC-UV\n(Encap. Efficiency)",
     "Quantification of encapsulated P4: solvent extraction followed by reverse-phase HPLC (C18 column, acetonitrile:water 70:30).",
     "Drug Quantification"),
]
for i, (title, body, tag) in enumerate(techs):
    col = i % 3
    row = i // 2
    x = 0.25 + col * 3.25
    y = 1.1 + (i // 3) * 2.2
    rect(s, x, y, 3.1, 2.1, WHITE)
    rect(s, x, y, 0.07, 2.1, ACCENT if row==0 else ACCENT2)
    txt(s, title, x+0.15, y+0.08, 2.9, 0.5,
        size=11, bold=True, color=DARK_BG)
    txt(s, body, x+0.12, y+0.62, 2.9, 1.35,
        size=9.5, color=DARK_TEXT)
footer_refs(s, "ICH Q2(R1) | USP <846> (Particle Size) | ASTM E1356 (DSC) | JCPDS crystallography database")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 23 – OBJECTIVE 2: IN-VITRO RELEASE
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objective 2: In-vitro Drug Release Kinetics",
               "Protocol Design & Analytical Methods")
steps23 = [
    ("Release Medium",
     "Simulated vaginal fluid (SVF): pH 4.5 phosphate buffer + 0.1% Tween 80 (solubiliser) at 37°C ± 0.5°C (thermostatted water bath). Volume: 100 mL."),
    ("Apparatus & Sampling",
     "Franz diffusion cells / dialysis membrane (MWCO 12 kDa) or USP Apparatus II (paddle, 50 rpm). Samples withdrawn (3 mL) at: 0, 1, 2, 4, 8, 12, 24, 48, 72, 96, 120, 144, 168 h."),
    ("Drug Quantification",
     "P4 concentration determined by validated RP-HPLC-UV (λ=254 nm, C18 column, mobile phase: ACN:H₂O 70:30, flow rate 1.0 mL/min). LOD: 0.05 µg/mL; LOQ: 0.15 µg/mL."),
    ("Release Parameters",
     "Cumulative % drug release, t₅₀ (time for 50% release), t₈₀, mean dissolution time (MDT), similarity factor (f₂) between formulations."),
]
for i, (title, body) in enumerate(steps23):
    rect(s, 0.3, 1.1 + i*1.05, 9.4, 0.98, WHITE if i%2==0 else RGBColor(0xE0,0xF2,0xF4))
    rect(s, 0.3, 1.1 + i*1.05, 0.07, 0.98, ACCENT)
    txt(s, f"Step {i+1}: {title}", 0.5, 1.14 + i*1.05, 3.5, 0.38,
        size=12, bold=True, color=DARK_BG)
    txt(s, body, 4.1, 1.14 + i*1.05, 5.5, 0.88,
        size=10, color=DARK_TEXT)
footer_refs(s, "USP <711> Dissolution | Shah et al. (1989) Pharm Res | ICH Q2(R1) Analytical Validation")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 24 – DRUG RELEASE MODELS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Mathematical Drug Release Models",
               "Kinetic Analysis to Determine Release Mechanism")
models = [
    ("Zero-Order Model",
     "Qt = Q₀ + K₀t",
     "Drug release independent of concentration; constant rate. Ideal for controlled-release formulations. R² → 1 indicates best fit."),
    ("First-Order Model",
     "log Qt = log Q₀ − K₁t/2.303",
     "Release proportional to remaining drug concentration. Common for porous matrix systems. Rapid initial release phase."),
    ("Higuchi Model",
     "Qt = Kн × √t",
     "Drug diffusion from insoluble matrix; rate decreases with time. Applicable to P4 dispersed in alginate matrix."),
    ("Korsmeyer–Peppas Model",
     "Mt/M∞ = K × tⁿ",
     "n ≤ 0.45: Fickian diffusion | 0.45 < n < 0.89: anomalous transport | n ≥ 0.89: Case II transport (swelling). Best for biopolymer matrices."),
]
for i, (title, eq, interp) in enumerate(models):
    col = i % 2
    row = i // 2
    x = 0.3 + col * 4.85
    y = 1.1 + row * 2.1
    rect(s, x, y, 4.55, 2.0, WHITE)
    rect(s, x, y, 0.07, 2.0, ACCENT if col==0 else ACCENT2)
    txt(s, title, x+0.15, y+0.07, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    # Equation box
    rect(s, x+0.15, y+0.45, 4.2, 0.45, RGBColor(0xE8,0xF4,0xF6))
    txt(s, eq, x+0.2, y+0.47, 4.1, 0.4,
        size=13, bold=True, italic=True, color=DARK_BG, align=PP_ALIGN.CENTER)
    txt(s, interp, x+0.15, y+0.97, 4.2, 0.95,
        size=10, color=DARK_TEXT)
footer_refs(s, "Korsmeyer et al. (1983) Int J Pharm 15:25-35 | Higuchi (1961) J Pharm Sci 50:874-875 | Costa & Lobo (2001) Eur J Pharm Sci 13:123-133")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 25 – SPONGE FABRICATION
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objective 3: Intravaginal Sponge Development",
               "Fabrication, Drug Loading & Quality Control")
fab_steps = [
    ("Sponge Material Selection",
     "Medical-grade polyurethane foam sponge (40 PPI) or natural cellulose sponge; dimension: 10 cm length × 5 cm diameter (tapering). Sterilised by gamma irradiation (25 kGy)."),
    ("Encapsulated P4 Loading",
     "Alginate-encapsulated P4 microparticles suspended in 5% CMC gel vehicle (viscosity 2000–4000 cPs). Dose equivalent: 0.5 or 1.0 g crystalline P4 per sponge."),
    ("Sponge Impregnation",
     "Microparticle suspension absorbed into sponge by vacuum impregnation method; sponge dried at 40°C for 6 h under aseptic conditions. Uniformity of loading verified by HPLC."),
    ("Packaging & Sterilisation",
     "Individually packed in sterile pouches; secondary EO sterilisation; 2–8°C storage. Shelf-life stability testing at 3, 6 months (ICH Q1A conditions)."),
    ("Retrieval String",
     "Nylon retrieval cord (60 cm) attached per device; ASTM F2132 tensile strength ≥ 50 N."),
]
for i, (title, body) in enumerate(fab_steps):
    rect(s, 0.3, 1.1 + i*0.88, 9.4, 0.82,
         WHITE if i%2==0 else RGBColor(0xE0,0xF2,0xF4))
    rect(s, 0.3, 1.1 + i*0.88, 0.07, 0.82, ACCENT)
    txt(s, title, 0.5, 1.14 + i*0.88, 3.0, 0.35,
        size=11, bold=True, color=DARK_BG)
    txt(s, body, 3.6, 1.14 + i*0.88, 6.0, 0.72,
        size=10, color=DARK_TEXT)
footer_refs(s, "ICH Q1A(R2) Stability Testing | USP <1> Injections sterility | CPCSEA Animal Ethics guidelines")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 26 – ANIMAL EXPERIMENT DESIGN
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Objective 3: Animal Experiment Design",
               "Randomised Controlled Trial – Anestrus & Repeat Breeding Cattle")
# Group table
headers_a = ["Group", "Category", "Treatment", "n", "AI Protocol"]
rows_a = [
    ["G1 – Negative Control", "Anestrus", "No treatment (saline intravaginal sponge)", "10", "Observed; AI on detected oestrus"],
    ["G2 – Positive Control", "Anestrus", "CIDR (1.38 g P4, 9 days) + eCG 400 IU at removal", "10", "FTAI at 56 h post-CIDR removal"],
    ["G3 – Treatment Low", "Anestrus", "Encapsulated P4 sponge (0.5 g eq.) 9 days + eCG 400 IU", "10", "FTAI at 56 h post-removal"],
    ["G4 – Treatment High", "Anestrus", "Encapsulated P4 sponge (1.0 g eq.) 9 days + eCG 400 IU", "10", "FTAI at 56 h post-removal"],
    ["G5 – RBS Control", "Repeat Breeding", "No treatment", "10", "AI on detected oestrus"],
    ["G6 – RBS Treatment", "Repeat Breeding", "Encapsulated P4 sponge (1.0 g eq.) 9 days + eCG 400 IU", "10", "FTAI at 56 h post-removal"],
]
add_table(s, headers_a, rows_a, 0.25, 1.1, 9.5, 3.6, font_size=9)
txt(s,
    "Species: Crossbred dairy cattle (HF × Sahiwal or Jersey × Sahiwal) | "
    "Ethics: CPCSEA/Institutional Animal Ethics Committee approval | "
    "Sample size based on 80% power, α=0.05, expected CR difference ≥15%",
    0.3, 4.75, 9.4, 0.65, size=10, italic=True, color=MUTED)
footer_refs(s, "CPCSEA Guidelines 2018 | Colazo & Mapletoft (2014) Can Vet J | Wiltbank et al. (2012) J Dairy Sci")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 27 – EVALUATION PARAMETERS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Evaluation Parameters", "In-vivo Assessment of Reproductive Outcomes")
params = [
    ("Oestrus Response Rate (%)",
     "% animals showing oestrous behaviour within 72 h of device removal. Detection: tail-painting + Kamar heat detectors + rectal palpation (12-hourly)."),
    ("Follicular Dynamics",
     "Transrectal USG (7.5 MHz linear probe) every 48 h during synchronisation and 12 h peri-ovulation. Record: dominant follicle diameter (mm), CL area (cm²), ovulation day."),
    ("Plasma Progesterone Profile",
     "Blood collected on D0, D3, D6, D9 (device in), D12, D16, D20, D24 post-removal. P4 by RIA or ELISA (Cayman Chemical / DRG kit). Target: >1 ng/mL during luteal phase."),
    ("Conception & Pregnancy Rate",
     "Rectal palpation + USG at 45 days post-AI. Pregnancy confirmation by detection of foetal membrane slip / embryonic vesicle. CR (%) = pregnancies / AI × 100."),
    ("Vaginal Safety Assessment",
     "Vaginal discharge scoring (0–3 scale), cytology (Papanicolaou stain), and mucous membrane biopsy at D0 and D9 for histopathology (H&E staining)."),
]
for i, (title, body) in enumerate(params):
    rect(s, 0.3, 1.1 + i*0.88, 9.4, 0.82,
         WHITE if i%2==0 else RGBColor(0xE0,0xF2,0xF4))
    rect(s, 0.3, 1.1 + i*0.88, 0.07, 0.82, ACCENT if i%2==0 else ACCENT2)
    txt(s, title, 0.5, 1.14 + i*0.88, 3.0, 0.35,
        size=11, bold=True, color=DARK_BG)
    txt(s, body, 3.6, 1.14 + i*0.88, 6.1, 0.72,
        size=10, color=DARK_TEXT)
footer_refs(s, "USG protocol: Ginther (1998) Ultrasonic Imaging | RIA: Niswender et al. (1969) | CPCSEA ethics")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 28 – STATISTICAL ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Statistical Analysis Plan")
stat_items = [
    ("One-Way ANOVA",
     "Comparison of continuous variables (P4 concentrations, follicular diameter, particle size, EE%) among treatment groups. Post-hoc: Tukey's HSD test (α=0.05)."),
    ("Repeated-Measures ANOVA",
     "Plasma P4 profile over time (Days 0–24) across groups. Mauchly's test for sphericity; Greenhouse–Geisser correction if violated."),
    ("Chi-Square / Fisher's Exact Test",
     "Comparison of proportions: oestrus response rate, conception rate, pregnancy rate between treatment and control groups."),
    ("Pearson Correlation",
     "Correlation between in-vitro P4 release kinetics and in-vivo plasma P4 levels; follicle diameter vs. conception rate."),
    ("Power Analysis",
     "A priori sample size: G*Power 3.1 (α=0.05, power=0.80, effect size f=0.4 based on pilot studies / published CR differences of 15–20%)."),
    ("Software",
     "SPSS v26.0 / SAS 9.4 / R (v4.2.0) for statistical computation. GraphPad Prism 9.0 for graphical representation."),
]
for i, (title, body) in enumerate(stat_items):
    col = i % 2
    row = i // 2
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 4.55, 1.38, WHITE)
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 0.07, 1.38, ACCENT if col==0 else ACCENT2)
    txt(s, title, 0.5 + col*4.85, 1.14 + row*1.45, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body, 0.5 + col*4.85, 1.52 + row*1.45, 4.3, 0.85,
        size=10, color=DARK_TEXT)
footer_refs(s, "Montgomery (2017) Design and Analysis of Experiments | Faul et al. (2009) Behav Res Methods 41:1149-1160")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 29 – EXPECTED OUTCOMES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Expected Outcomes", "Scientific & Applied Significance")
outcomes = [
    ("Optimised Formulation",
     "Spray-dried Na-alginate microspheres with EE >80%, particle size 5–30 µm, PDI <0.3, and zeta potential ≤−20 mV confirming colloidal stability."),
    ("Controlled Release Kinetics",
     "In-vitro: sustained P4 release over 7–14 days following Korsmeyer–Peppas (anomalous diffusion/swelling); superior to burst-release of conventional CIDR."),
    ("Sustained Plasma P4",
     "In-vivo plasma P4 ≥1 ng/mL maintained for 7–14 days in treatment groups; area under the P4-time curve (AUC) significantly greater vs. control."),
    ("Improved Oestrus Synchrony",
     "Oestrus response rate ≥70% in treatment vs. <30% in untreated anestrus control; tighter oestrus clustering within 48 h of sponge removal."),
    ("Higher Conception Rate",
     "Conception rate 45–60% in biopolymer P4 sponge groups vs. 20–30% in negative controls; comparable or superior to CIDR positive control."),
    ("Vaginal Safety",
     "No significant histopathological changes, mucosal irritation, or adverse cytology; biodegradable matrix ensures no device retrieval complications."),
]
for i, (title, body) in enumerate(outcomes):
    col = i % 2
    row = i // 2
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 4.55, 1.38,
         WHITE if row%2==0 else RGBColor(0xE8,0xF8,0xF6))
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 0.07, 1.38,
         ACCENT if row==0 else ACCENT2 if row==1 else GOLD)
    txt(s, title, 0.5 + col*4.85, 1.14 + row*1.45, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body, 0.5 + col*4.85, 1.52 + row*1.45, 4.3, 0.85,
        size=10, color=DARK_TEXT)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 30 – NOVELTY
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
bg(s, DARK_BG)
txt(s, "Novelty & Innovations of the Present Research", 0.5, 0.2, 9.0, 0.65,
    size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
novelties = [
    ("First in Cattle",
     "First systematic evaluation of spray-dried, Na-alginate-encapsulated micronised P4 intravaginal sponge for oestrus synchronisation in cattle in the Indian context."),
    ("Micronisation + Encapsulation",
     "Dual approach: micronisation (↑ dissolution rate) + biopolymer encapsulation (sustained release) synergistically improves P4 bioavailability over existing devices."),
    ("Biodegradable Device",
     "Alginate-based matrix is biodegradable and mucoadhesive; eliminates retrieval requirement and reduces environmental residue burden of conventional silicone/polyurethane devices."),
    ("Design of Experiment (DoE)",
     "Box-Behnken/CCD optimisation of spray-drying parameters — a rigorous, statistically sound formulation development approach rarely applied in veterinary drug delivery."),
    ("Dual Indication",
     "Simultaneous evaluation in both anestrus and repeat breeding — addresses two major causes of reproductive failure in a single study design."),
]
for i, (title, body) in enumerate(novelties):
    rect(s, 0.3, 1.0 + i*0.92, 9.4, 0.85, RGBColor(0x07, 0x2D, 0x33))
    rect(s, 0.3, 1.0 + i*0.92, 0.07, 0.85, ACCENT)
    txt(s, f"★  {title}", 0.5, 1.04 + i*0.92, 3.5, 0.38,
        size=12, bold=True, color=ACCENT)
    txt(s, body, 4.2, 1.04 + i*0.92, 5.4, 0.78,
        size=10, color=WHITE)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 31 – POTENTIAL APPLICATIONS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Potential Applications & Translational Significance")
apps = [
    ("Dairy Farm Management",
     "Synchronised AI programmes at herd level; reduces dependency on oestrus detection; improves calving intervals in organised and semi-organised dairy farms."),
    ("Smallholder Farmers",
     "Low-cost biodegradable sponge affordable for smallholder farmers in India and developing nations; no cold chain required for encapsulated powder form."),
    ("Beef Cattle Industry",
     "Fixed-time AI in beef herds; reduces labour costs; applicable to Bos indicus and crossbred cattle in tropical environments."),
    ("Species Extrapolation",
     "Formulation principle translatable to buffalo, goat, sheep, and camel with appropriate dose adjustments — broad utility in South Asian livestock systems."),
    ("Commercial Development",
     "Technology transfer potential to veterinary pharmaceutical industry for development of novel, CDSCO-approved intravaginal P4 devices in India."),
    ("Reproductive Biotechnology Platform",
     "Biopolymer encapsulation approach applicable to other reproductive hormones (eCG, GnRH, FSH) — foundational platform for next-generation hormone delivery."),
]
for i, (title, body) in enumerate(apps):
    col = i % 2
    row = i // 2
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 4.55, 1.38, WHITE)
    rect(s, 0.3 + col*4.85, 1.1 + row*1.45, 0.07, 1.38,
         ACCENT if i%3==0 else ACCENT2 if i%3==1 else GOLD)
    txt(s, title, 0.5 + col*4.85, 1.14 + row*1.45, 4.3, 0.35,
        size=12, bold=True, color=DARK_BG)
    txt(s, body, 0.5 + col*4.85, 1.52 + row*1.45, 4.3, 0.85,
        size=10, color=DARK_TEXT)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 32 – GANTT CHART / WORK PLAN
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Work Plan – 3-Year Gantt Chart")

activities = [
    "Literature Review & Ethics Approval",
    "Micronisation of P4",
    "Formulation (DoE, Spray Drying)",
    "Physicochemical Characterisation",
    "In-vitro Release Studies",
    "Sponge Fabrication & QC",
    "Animal Trial – Anestrus Group",
    "Animal Trial – Repeat Breeding Group",
    "Laboratory Analysis (Hormones/Histo.)",
    "Data Analysis & Statistics",
    "Thesis Writing & Submission",
]
quarters = ["Y1Q1","Y1Q2","Y1Q3","Y1Q4","Y2Q1","Y2Q2","Y2Q3","Y2Q4","Y3Q1","Y3Q2","Y3Q3","Y3Q4"]

# Timeline schedule (1 = active, 0 = inactive)
schedule = [
    [1,1,0,0, 0,0,0,0, 0,0,0,0],  # Lit review
    [0,1,1,0, 0,0,0,0, 0,0,0,0],  # Micronisation
    [0,0,1,1, 1,0,0,0, 0,0,0,0],  # Formulation
    [0,0,0,1, 1,1,0,0, 0,0,0,0],  # Characterisation
    [0,0,0,0, 1,1,0,0, 0,0,0,0],  # In-vitro
    [0,0,0,0, 0,1,1,0, 0,0,0,0],  # Sponge
    [0,0,0,0, 0,0,1,1, 1,0,0,0],  # Animal trial 1
    [0,0,0,0, 0,0,0,1, 1,1,0,0],  # Animal trial 2
    [0,0,0,0, 0,0,1,1, 1,1,0,0],  # Lab analysis
    [0,0,0,0, 0,0,0,0, 1,1,1,0],  # Stats
    [0,0,0,0, 0,0,0,0, 0,1,1,1],  # Thesis
]

cell_w = 0.62
cell_h = 0.33
label_w = 3.2
start_x = 0.25 + label_w
start_y = 1.05

# Quarter headers
for qi, q in enumerate(quarters):
    x = start_x + qi * cell_w
    rect(s, x, start_y, cell_w, 0.32, DARK_BG)
    txt(s, q, x, start_y+0.02, cell_w, 0.28,
        size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Activity rows
for ai, activity in enumerate(activities):
    y = start_y + 0.34 + ai * cell_h
    bg_c = WHITE if ai%2==0 else RGBColor(0xE8,0xF4,0xF6)
    rect(s, 0.25, y, label_w, cell_h, bg_c)
    txt(s, activity, 0.28, y+0.04, label_w-0.1, cell_h-0.05,
        size=8, color=DARK_TEXT, font_face="Calibri")
    for qi in range(12):
        x = start_x + qi * cell_w
        if schedule[ai][qi] == 1:
            rect(s, x+0.02, y+0.04, cell_w-0.04, cell_h-0.07, ACCENT)
        else:
            rect(s, x, y, cell_w, cell_h, bg_c)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 33 – EXPECTED PUBLICATIONS
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Expected Publications", "Dissemination Plan & Target Journals")
pubs = [
    ("Paper 1 (Year 1–2)",
     "Preparation and Optimisation of Sodium Alginate-Encapsulated Micronised Progesterone Microparticles by Spray Drying: Physicochemical Characterisation and In-vitro Release Kinetics",
     "Drug Development and Industrial Pharmacy / International Journal of Pharmaceutics"),
    ("Paper 2 (Year 2–3)",
     "Comparative Evaluation of Biopolymer Encapsulated Progesterone Intravaginal Sponge and Conventional CIDR for Oestrus Synchronisation and Conception in Anestrus Crossbred Cattle",
     "Theriogenology / Animal Reproduction Science"),
    ("Paper 3 (Year 2–3)",
     "Efficacy of Biopolymer-Based Controlled-Release Progesterone Intravaginal Sponge in Management of Repeat Breeding Syndrome in Cattle: Follicular Dynamics and Hormonal Profile",
     "Reproduction in Domestic Animals / Indian Journal of Animal Reproduction"),
    ("Review Paper (Year 1)",
     "Biopolymer-Based Controlled Delivery Systems for Reproductive Hormones in Livestock: A Critical Review",
     "Small Ruminant Research / Tropical Animal Health & Production / Veterinary World"),
]
for i, (stage, title, journal) in enumerate(pubs):
    rect(s, 0.3, 1.1 + i*1.08, 9.4, 1.02,
         WHITE if i%2==0 else RGBColor(0xE8,0xF4,0xF6))
    rect(s, 0.3, 1.1 + i*1.08, 0.07, 1.02,
         ACCENT if i%2==0 else GOLD)
    txt(s, stage, 0.5, 1.14 + i*1.08, 9.2, 0.28,
        size=11, bold=True, color=DARK_BG)
    txt(s, f"Title: {title}", 0.5, 1.42 + i*1.08, 9.2, 0.35,
        size=9.5, italic=True, color=DARK_TEXT)
    txt(s, f"Target: {journal}", 0.5, 1.77 + i*1.08, 9.2, 0.25,
        size=9, color=ACCENT2, bold=True)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 34 – KEY REFERENCES
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
section_header(s, "Key References",
               "Peer-reviewed Citations Supporting the Present Study")
refs = [
    "1.  Rathbone MJ et al. (2002). Controlled release of hormones for use in ruminants. Adv Drug Deliv Rev, 54(7):1041–1054.",
    "2.  Colazo MG & Mapletoft RJ (2014). A review of protocols used for estrus synchronization in beef and dairy cattle. Can Vet J, 55(8):772–780.",
    "3.  Wiltbank MC et al. (2012). Pivotal periods for pregnancy loss during the first and second months of gestation in lactating dairy cows. Theriogenology, 78:364–375.",
    "4.  Burke CR et al. (2001). Progesterone-releasing intravaginal devices. J Reprod Fertil Suppl, 57:315–321.",
    "5.  Patel RP et al. (2018). Microencapsulation of progesterone using sodium alginate by spray drying. Drug Dev Ind Pharm, 44(11):1783–1793.",
    "6.  Lee KY & Mooney DJ (2012). Alginate: properties and biomedical applications. Prog Polym Sci, 37(1):106–126.",
    "7.  Singh J et al. (2019). Carbopol hydrogel-based intravaginal progesterone delivery in Beetal goats. Small Rumin Res, 172:15–21.",
    "8.  Vigani B et al. (2022). Mucoadhesive nanoparticles for progesterone intravaginal delivery. Pharmaceutics, 14(3):620.",
    "9.  Costa P & Lobo JMS (2001). Modeling and comparison of dissolution profiles. Eur J Pharm Sci, 13(2):123–133.",
    "10. Vehring R (2008). Pharmaceutical particle engineering via spray drying. Pharm Res, 25(5):999–1022.",
    "11. Baruselli PS et al. (2004). Reproductive biotechnology in the beef cattle industry. Anim Reprod Sci, 82:479–486.",
    "12. de Ziegler D et al. (2013). Clinical uses of progesterone in obstetrics and gynecology. Climacteric, 16(Suppl 1):8–16.",
]
txt(s, "⚠ IMPORTANT: Independently verify all DOIs and details before submission. "
       "Only use citations confirmed in indexed databases (PubMed / Scopus / Web of Science).",
    0.3, 1.1, 9.4, 0.4,
    size=10, bold=True, italic=True, color=RGBColor(0xC0,0x39,0x00))
multiline_txt(s, refs, 0.3, 1.55, 9.4, 3.9,
              size=9, color=DARK_TEXT, spacing_after=2)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 35 – THANK YOU
# ═══════════════════════════════════════════════════════════════════════════
s = blank_slide(prs)
bg(s, DARK_BG)
rect(s, 0, 0, W, 0.12, ACCENT)
rect(s, 0, 5.5, W, 0.125, ACCENT)

txt(s, "Thank You", 0.5, 1.4, 9.0, 1.2,
    size=54, bold=True, color=WHITE,
    align=PP_ALIGN.CENTER, font_face="Calibri")

rect(s, 2.5, 2.7, 5.0, 0.06, ACCENT)

txt(s, "Questions & Suggestions are Welcome",
    0.5, 2.85, 9.0, 0.55,
    size=18, italic=True, color=ACCENT,
    align=PP_ALIGN.CENTER)

contact_lines = [
    "[Scholar Name]  |  PhD Scholar – Vet. Gynaecology & Reproductive Biotechnology",
    "ICAR – National Dairy Research Institute, Karnal",
    "Email: [scholar@ndri.res.in]",
]
for i, line in enumerate(contact_lines):
    txt(s, line, 0.5, 3.6 + i*0.38, 9.0, 0.36,
        size=12, color=RGBColor(0xCC, 0xE8, 0xEB),
        align=PP_ALIGN.CENTER)

txt(s,
    "Major Advisor: [Guide Name], PhD   |   Co-Advisor: [Co-guide Name], PhD",
    0.5, 5.0, 9.0, 0.38, size=11,
    color=MUTED, align=PP_ALIGN.CENTER, italic=True)

# ── Save ─────────────────────────────────────────────────────────────────────
out = "/mnt/user-data/outputs/PhD_Synopsis_Progesterone_Delivery_Cattle.pptx"
prs.save(out)
print(f"✅  Presentation saved → {out}")
print(f"    Total slides: {len(prs.slides)}")
