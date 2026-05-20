"""Generate a PowerPoint presentation on Repeat Breeding in cattle.

Theme: navy + amber/orange (modern dairy)
Includes embedded cow & buffalo photographs from Wikimedia Commons.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

HERE = os.path.dirname(os.path.abspath(__file__))
IMG  = os.path.join(HERE, "images")

# ---------- New theme colors (navy + amber) ----------
PRIMARY = RGBColor(0x1B, 0x3A, 0x57)      # deep navy
ACCENT  = RGBColor(0xE0, 0x7A, 0x18)      # warm amber/orange
ACCENT2 = RGBColor(0x2E, 0x86, 0xAB)      # ocean teal
LIGHT   = RGBColor(0xF7, 0xF3, 0xEB)      # warm cream
DARK    = RGBColor(0x1F, 0x25, 0x33)      # near-black navy
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GRAY    = RGBColor(0x55, 0x5C, 0x66)
SOFT    = RGBColor(0xE6, 0xDF, 0xD0)      # soft warm gray

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


# ---------- helpers ----------
def add_bg(slide, color=LIGHT):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    return bg


def add_side_bar(slide):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.32), SH)
    bar.line.fill.background()
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(0.32), 0, Inches(0.07), SH)
    accent.line.fill.background()
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT


def add_footer(slide, page_num, total):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.05),
                                  Inches(12), Inches(0.3))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "Repeat Breeding in Cattle"
    r.font.size = Pt(10); r.font.color.rgb = GRAY; r.font.name = "Calibri"

    tb2 = slide.shapes.add_textbox(Inches(11.7), Inches(7.05),
                                   Inches(1.3), Inches(0.3))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"{page_num} / {total}"
    r2.font.size = Pt(10); r2.font.color.rgb = GRAY; r2.font.name = "Calibri"


def add_title(slide, title, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(0.65), Inches(0.32),
                                  Inches(12.2), Inches(0.9))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size = Pt(34); r.font.bold = True
    r.font.color.rgb = PRIMARY; r.font.name = "Calibri"

    u = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(0.65), Inches(1.18),
                               Inches(1.4), Inches(0.07))
    u.line.fill.background(); u.fill.solid()
    u.fill.fore_color.rgb = ACCENT

    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.65), Inches(1.28),
                                      Inches(12.2), Inches(0.4))
        sp = sb.text_frame.paragraphs[0]
        sr = sp.add_run(); sr.text = subtitle
        sr.font.size = Pt(14); sr.font.italic = True
        sr.font.color.rgb = GRAY; sr.font.name = "Calibri"


def add_bullets(slide, bullets, left=Inches(0.85), top=Inches(1.85),
                width=Inches(12), height=Inches(5.0), size=18):
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
        bullet = "•  " if level == 0 else "–  "
        r = p.add_run()
        r.text = bullet + text
        r.font.size = Pt(size - level * 2)
        r.font.color.rgb = DARK
        r.font.name = "Calibri"


def add_card(slide, left, top, width, height, title, body,
             head_color=PRIMARY):
    head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top,
                                  width, Inches(0.55))
    head.line.fill.background(); head.fill.solid()
    head.fill.fore_color.rgb = head_color
    htf = head.text_frame; htf.margin_left = Inches(0.15)
    htf.margin_top = Inches(0.05); htf.margin_bottom = Inches(0.05)
    hp = htf.paragraphs[0]; hp.alignment = PP_ALIGN.LEFT
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
    btf.margin_left = Inches(0.15); btf.margin_right = Inches(0.15)
    btf.margin_top = Inches(0.1)
    for i, line in enumerate(body):
        p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
        p.space_after = Pt(4)
        r = p.add_run(); r.text = "• " + line
        r.font.size = Pt(12); r.font.color.rgb = DARK
        r.font.name = "Calibri"


def add_caption(slide, left, top, width, text):
    cb = slide.shapes.add_textbox(left, top, width, Inches(0.3))
    p = cb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(10); r.font.italic = True
    r.font.color.rgb = GRAY; r.font.name = "Calibri"


def add_image_with_frame(slide, path, left, top, width, height,
                         caption=None, frame_color=PRIMARY):
    # outer thin frame
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   left - Inches(0.05), top - Inches(0.05),
                                   width + Inches(0.1), height + Inches(0.1))
    frame.line.fill.background(); frame.fill.solid()
    frame.fill.fore_color.rgb = frame_color
    pic = slide.shapes.add_picture(path, left, top, width=width, height=height)
    if caption:
        add_caption(slide, left, top + height + Inches(0.05),
                    width, caption)
    return pic


# ============================================================
# SLIDE 1 — TITLE (with hero buffalo + cow images)
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, PRIMARY)

# soft amber band
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.0),
                          SW, Inches(1.5))
band.line.fill.background(); band.fill.solid()
band.fill.fore_color.rgb = ACCENT

# tag
tag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(0.65),
                        Inches(3.2), Inches(0.45))
tag.line.fill.background(); tag.fill.solid()
tag.fill.fore_color.rgb = ACCENT
ttf = tag.text_frame; ttf.margin_left = Inches(0.15)
tp = ttf.paragraphs[0]
tr = tp.add_run(); tr.text = "ANIMAL REPRODUCTION"
tr.font.size = Pt(13); tr.font.bold = True
tr.font.color.rgb = PRIMARY; tr.font.name = "Calibri"

# title
tb = s.shapes.add_textbox(Inches(0.65), Inches(3.1), Inches(12), Inches(0.9))
p = tb.text_frame.paragraphs[0]
r = p.add_run(); r.text = "REPEAT BREEDING"
r.font.size = Pt(56); r.font.bold = True
r.font.color.rgb = PRIMARY; r.font.name = "Calibri"

tb2 = s.shapes.add_textbox(Inches(0.65), Inches(3.95), Inches(12), Inches(0.6))
p2 = tb2.text_frame.paragraphs[0]
r2 = p2.add_run()
r2.text = "Causes, Diagnosis, Treatment & Prevention in Cows and Buffaloes"
r2.font.size = Pt(22); r2.font.color.rgb = DARK
r2.font.italic = True; r2.font.name = "Calibri"

# Hero photos in lower area
add_image_with_frame(s, COW_HOLSTEIN, Inches(0.65), Inches(5.05),
                     Inches(4.0), Inches(1.7),
                     "Holstein Friesian (dairy cow)", WHITE)
add_image_with_frame(s, BUFFALO, Inches(8.7), Inches(5.05),
                     Inches(4.0), Inches(1.7),
                     "Murrah buffalo", WHITE)

info = s.shapes.add_textbox(Inches(4.8), Inches(5.4),
                            Inches(3.8), Inches(1.4))
itf = info.text_frame
ip = itf.paragraphs[0]; ip.alignment = PP_ALIGN.CENTER
ir = ip.add_run()
ir.text = "Veterinary Gynaecology\n& Obstetrics"
ir.font.size = Pt(18); ir.font.bold = True
ir.font.color.rgb = WHITE; ir.font.name = "Calibri"
p2 = itf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run(); r2.text = "Seminar Presentation"
r2.font.size = Pt(13); r2.font.italic = True
r2.font.color.rgb = WHITE; r2.font.name = "Calibri"

# ============================================================
# SLIDE 2 — OUTLINE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s); add_side_bar(s)
add_title(s, "Presentation Outline")

items = [
    "1.  Introduction & Definition",
    "2.  Incidence & Economic Impact",
    "3.  Etiology / Causes",
    "4.  Pathophysiology",
    "5.  Diagnosis & Clinical Approach",
    "6.  Treatment & Management",
    "7.  Hormonal Therapy Protocols",
    "8.  Buffalo-Specific Repeat Breeding",
    "9.  Prevention Strategies",
    "10. Recent Advances",
    "11. Conclusion & References",
]
for i, text in enumerate(items):
    col = i // 6
    row = i % 6
    l = Inches(0.85 + col * 6.0)
    t = Inches(1.85 + row * 0.7)
    bullet = s.shapes.add_shape(MSO_SHAPE.OVAL, l, t + Inches(0.12),
                                Inches(0.18), Inches(0.18))
    bullet.line.fill.background(); bullet.fill.solid()
    bullet.fill.fore_color.rgb = ACCENT
    tx = s.shapes.add_textbox(l + Inches(0.35), t, Inches(5.5), Inches(0.5))
    p = tx.text_frame.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(18); r.font.color.rgb = DARK; r.font.name = "Calibri"

add_footer(s, 2, 19)

# ============================================================
# SLIDE 3 — INTRODUCTION (with cow image)
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Introduction",
          "Reproductive efficiency is the backbone of profitable dairy farming")

add_bullets(s, [
    "Reproduction is the most important factor governing the economic "
    "success of any dairy enterprise.",
    "A normal fertile cow should conceive within 60–90 days post-partum and "
    "deliver one calf every 12–13 months.",
    "Failure to conceive after repeated inseminations — despite normal "
    "estrous cycles and apparently healthy genitalia — is termed "
    "REPEAT BREEDING.",
    "It is one of the most frustrating and economically damaging "
    "reproductive disorders in cattle and buffaloes worldwide.",
    "Often multifactorial — involving the cow, the bull/semen, the "
    "inseminator and the environment.",
], left=Inches(0.85), top=Inches(1.85), width=Inches(8.0), size=16)

add_image_with_frame(s, COW_DAIRY, Inches(9.1), Inches(1.95),
                     Inches(3.7), Inches(2.78),
                     "Holstein dairy cow on pasture")
add_image_with_frame(s, SAHIWAL, Inches(9.1), Inches(5.0),
                     Inches(3.7), Inches(1.85),
                     "Sahiwal — indigenous Indian dairy breed")
add_footer(s, 3, 19)

# ============================================================
# SLIDE 4 — DEFINITION
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Definition")

box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.85), Inches(1.85),
                         Inches(11.6), Inches(1.8))
box.line.color.rgb = ACCENT; box.line.width = Pt(2)
box.fill.solid(); box.fill.fore_color.rgb = WHITE
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = Inches(0.3); tf.margin_right = Inches(0.3)
tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]
r = p.add_run()
r.text = ("A repeat breeder is a cow / buffalo of normal breeding age and "
         "parity, cycling regularly at intervals of 18–24 days, with no "
         "clinically detectable abnormality of the genital tract, that has "
         "failed to conceive after three or more successive services with "
         "fertile semen.")
r.font.size = Pt(17); r.font.italic = True
r.font.color.rgb = DARK; r.font.name = "Calibri"

add_bullets(s, [
    "Also known as: Repeat Breeder Syndrome / Repeat Breeder Cow (RBC).",
    "Synonym in older literature: 'Conception Failure'.",
    ("Key criteria — (a) regular oestrous cycles, (b) ≥3 services, "
     "(c) no palpable abnormality, (d) fertile semen, (e) competent AI."),
    "Differentiate from anoestrus, sub-oestrus, true infertility and sterility.",
], top=Inches(3.95), size=17)
add_footer(s, 4, 19)

# ============================================================
# SLIDE 5 — INCIDENCE
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Incidence & Economic Importance")

add_bullets(s, [
    "Global incidence: 10 – 15 % of all breedable cows; "
    "may exceed 20 – 30 % in poorly managed herds.",
    "Indian dairy herds: reported range 12 – 24 % "
    "(higher in cross-bred and high-yielding cows).",
    "Buffaloes: 8 – 18 %, often masked by silent oestrus.",
], top=Inches(1.85), size=17)

add_card(s, Inches(0.85), Inches(4.0), Inches(3.85), Inches(2.9),
         "Direct Losses",
         ["Extra AI doses & semen cost",
          "Veterinary & hormone bills",
          "Extended dry / open period",
          "Reduced lifetime calves"])
add_card(s, Inches(4.95), Inches(4.0), Inches(3.85), Inches(2.9),
         "Indirect Losses",
         ["Lower lactation yield",
          "Increased calving interval",
          "Premature culling",
          "Loss of genetic progress"], head_color=ACCENT)
add_card(s, Inches(9.05), Inches(4.0), Inches(3.85), Inches(2.9),
         "Estimated Cost",
         ["Each extra open day ≈ ₹150–250",
          "Per repeat breeder: ₹8 000–15 000 / lactation",
          "National-level losses run into thousands of crores"],
         head_color=ACCENT2)

add_footer(s, 5, 19)

# ============================================================
# SLIDE 6 — CAUSES OVERVIEW
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Etiology — An Overview",
          "Repeat breeding is multifactorial; causes are grouped into four broad heads")

add_card(s, Inches(0.85), Inches(1.95), Inches(2.95), Inches(4.9),
         "1. Fertilization Failure",
         ["Improper AI timing",
          "Poor semen quality",
          "Anovulation / delayed ovulation",
          "Reproductive tract obstruction",
          "Anti-sperm antibodies"])
add_card(s, Inches(4.0), Inches(1.95), Inches(2.95), Inches(4.9),
         "2. Early Embryonic Death",
         ["Hormonal imbalance",
          "Uterine infections",
          "Heat stress",
          "Chromosomal defects",
          "Nutritional deficiency"], head_color=ACCENT)
add_card(s, Inches(7.15), Inches(1.95), Inches(2.95), Inches(4.9),
         "3. Management Errors",
         ["Missed / mis-detected oestrus",
          "Wrong AI technique",
          "Faulty semen handling",
          "Inadequate records",
          "Stress at AI"], head_color=ACCENT2)
add_card(s, Inches(10.3), Inches(1.95), Inches(2.6), Inches(4.9),
         "4. Genetic / Other",
         ["Inbreeding",
          "Lethal genes",
          "Freemartinism",
          "White heifer disease",
          "Age & parity"], head_color=PRIMARY)

add_footer(s, 6, 19)

# ============================================================
# SLIDE 7 — INFECTIOUS / GENITAL CAUSES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Infectious & Genital Tract Causes")
add_bullets(s, [
    "Subclinical / chronic endometritis — most common cause (40–60 %).",
    ("Specific infections:", 0),
    ("Brucellosis (Brucella abortus)", 1),
    ("Trichomoniasis (Tritrichomonas foetus)", 1),
    ("Vibriosis / Campylobacteriosis (Campylobacter fetus)", 1),
    ("IBR–IPV (Bovine Herpesvirus-1)", 1),
    ("BVD (Bovine Viral Diarrhoea virus)", 1),
    ("Leptospirosis, Mycoplasma, Ureaplasma spp.", 1),
    "Anatomical defects: cervical stenosis, persistent hymen, segmental aplasia.",
    "Cystic ovarian disease (follicular & luteal cysts).",
    "Salpingitis & hydrosalpinx — block fertilization.",
], size=17)
add_footer(s, 7, 19)

# ============================================================
# SLIDE 8 — NUTRITIONAL CAUSES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Nutritional & Metabolic Causes")
add_bullets(s, [
    "Negative energy balance in early lactation → delayed cyclicity & poor "
    "embryo quality.",
    "Excess dietary protein → high blood urea N (>20 mg/dL) is embryotoxic.",
    ("Mineral & trace-element deficiencies:", 0),
    ("Phosphorus, Calcium — anoestrus, weak heat", 1),
    ("Copper, Cobalt, Zinc, Manganese — silent heat, low conception", 1),
    ("Selenium / Vitamin E — early embryonic death, retained placenta", 1),
    ("Iodine — irregular cycles", 1),
    ("Vitamin A, D, E deficiencies — endometrial & embryo defects", 1),
    "Mycotoxin contamination of feed (zearalenone, aflatoxin) → ovarian "
    "& embryonic disturbance.",
    "Body Condition Score < 2.5 or > 4 at AI is strongly associated with "
    "repeat breeding.",
], size=17)
add_footer(s, 8, 19)

# ============================================================
# SLIDE 9 — HORMONAL CAUSES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Hormonal & Endocrine Causes")
add_bullets(s, [
    "Inadequate pre-ovulatory LH surge → delayed or failed ovulation.",
    "Sub-luteal progesterone (< 1 ng/mL after Day-5) → embryonic loss.",
    "Persistent corpus luteum (CL) — pseudo-pregnancy effect.",
    "Hyper-prolactinaemia & thyroid dysfunction.",
    "Premature luteolysis — failure of maternal recognition (low Interferon-τ).",
    "Stress-induced cortisol rise → suppression of GnRH/LH.",
    "Heat stress in summer — reduced oestradiol, weak signs of oestrus, "
    "poor oocyte quality.",
], size=18)
add_footer(s, 9, 19)

# ============================================================
# SLIDE 10 — MANAGEMENT / AI ERRORS
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Management & Inseminator-Related Causes")
add_bullets(s, [
    "Improper heat detection — single twice-daily check misses 25–50 % of heats.",
    "Wrong time of AI — outside the optimum window (mid- to late-oestrus).",
    "Faulty semen thawing (water temp, time) and storage practices.",
    "Repeated use of same sub-fertile bull / semen batch.",
    "Trauma to cervix or uterus during AI; deep horn deposition errors.",
    "Hygiene lapses → introduction of pathogens at the time of AI.",
    "Lack of accurate breeding records and follow-up.",
], top=Inches(1.85), size=17)
add_footer(s, 10, 19)

# ============================================================
# SLIDE 11 — PATHOPHYSIOLOGY
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Pathophysiology — Where Things Go Wrong")

labels = ["Oestrus &\nOvulation", "Sperm\nTransport",
          "Fertilization", "Embryo\nDevelopment",
          "Maternal\nRecognition", "Pregnancy"]
left = Inches(0.7)
w = Inches(1.95); h = Inches(1.0)
top = Inches(2.2)
for i, lab in enumerate(labels):
    color = PRIMARY if i % 2 == 0 else ACCENT
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             left + i * (w + Inches(0.07)), top, w, h)
    box.line.fill.background(); box.fill.solid()
    box.fill.fore_color.rgb = color
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = lab
    r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = "Calibri"

for i in range(len(labels) - 1):
    arr_left = left + (i + 1) * w + i * Inches(0.07) - Inches(0.02)
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arr_left, top + Inches(0.42),
                           Inches(0.13), Inches(0.18))
    a.line.fill.background(); a.fill.solid()
    a.fill.fore_color.rgb = DARK

add_bullets(s, [
    "Failure of ovulation or asynchronous ovulation → no fertilization.",
    "Sperm transport disturbance / hostile cervico-uterine environment.",
    "Fertilization failure — poor oocyte / sperm quality, anti-sperm antibodies.",
    "Early embryonic death (Day 8–16) — most common in repeat breeders.",
    "Failure of maternal recognition of pregnancy → CL regresses → return to oestrus.",
], top=Inches(3.6), size=16)
add_footer(s, 11, 19)

# ============================================================
# SLIDE 12 — DIAGNOSIS
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Diagnosis — Systematic Clinical Approach")
add_bullets(s, [
    ("History:", 0),
    ("Age, parity, calving date, services given, AI dates, semen used, bull fertility", 1),
    ("Nutrition, deworming, vaccination, previous treatments", 1),
    ("General Examination: Body condition, anaemia, lameness, mastitis, systemic disease.", 0),
    ("Gynaeco-clinical Examination:", 0),
    ("Vaginoscopy — discharge, cervicitis, adhesions", 1),
    ("Per-rectal palpation — uterus, ovaries, cervix, follicles, CL", 1),
    ("Trans-rectal ultrasonography — follicles, CL, uterine fluid, embryo", 1),
    ("Laboratory Tests:", 0),
    ("Cervico-vaginal mucus culture & sensitivity, cytology", 1),
    ("Serum progesterone, blood urea, minerals, liver / thyroid profile", 1),
    ("Serology / PCR — Brucella, IBR, BVD, Leptospira, Trichomonas", 1),
], size=15)
add_footer(s, 12, 19)

# ============================================================
# SLIDE 13 — TREATMENT GENERAL (with image)
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Treatment — General Principles",
          "Treat the cause, not just the symptom")

add_bullets(s, [
    "Correct underlying nutrition & body condition (BCS 3.0–3.5 at AI).",
    "Treat genital tract infection (intra-uterine antibiotics, antiseptics).",
    "Correct hormonal imbalance with appropriate protocols.",
    "Improve heat detection and AI timing (AM-PM rule / activity meters).",
    "Use proven, high-quality semen; ensure proper thawing & deposition.",
    "Reduce stress — provide shade, water, comfortable housing, fly control.",
    "Maintain proper records; cull chronic non-responders after 5–6 services.",
], left=Inches(0.85), top=Inches(1.85), width=Inches(8.0), size=16)

add_image_with_frame(s, COW_FARM, Inches(9.1), Inches(2.0),
                     Inches(3.7), Inches(2.74),
                     "Well-managed dairy cow on farm")
# small accent quote box
qb = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(9.1), Inches(5.1),
                        Inches(3.7), Inches(1.6))
qb.line.color.rgb = ACCENT; qb.line.width = Pt(1.5)
qb.fill.solid(); qb.fill.fore_color.rgb = WHITE
qtf = qb.text_frame; qtf.word_wrap = True
qtf.margin_left = Inches(0.15); qtf.margin_right = Inches(0.15)
qtf.margin_top = Inches(0.1)
qp = qtf.paragraphs[0]
qr = qp.add_run(); qr.text = "“Prevention is better, "
qr.font.size = Pt(14); qr.font.italic = True; qr.font.bold = True
qr.font.color.rgb = PRIMARY
qr2 = qp.add_run(); qr2.text = "but timely treatment saves the cow.”"
qr2.font.size = Pt(14); qr2.font.italic = True
qr2.font.color.rgb = DARK
qp2 = qtf.add_paragraph()
qr3 = qp2.add_run(); qr3.text = "— Field motto in dairy practice"
qr3.font.size = Pt(11); qr3.font.color.rgb = GRAY

add_footer(s, 13, 19)

# ============================================================
# SLIDE 14 — HORMONAL PROTOCOLS
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Specific Therapeutic Protocols")

add_card(s, Inches(0.85), Inches(1.95), Inches(3.85), Inches(5.0),
         "GnRH at AI",
         ["Buserelin 10 µg or Gonadorelin 100 µg IM at the time of AI",
          "Improves LH surge & ovulation",
          "Increases conception by 10–20 %",
          "Specially useful in delayed ovulators"])
add_card(s, Inches(4.95), Inches(1.95), Inches(3.85), Inches(5.0),
         "Progesterone Support",
         ["CIDR / PRID for 7 days post-AI",
          "Long-acting progesterone injection on Day 5",
          "Prevents premature luteolysis & embryo loss",
          "Best for cows with short luteal phase"], head_color=ACCENT)
add_card(s, Inches(9.05), Inches(1.95), Inches(3.85), Inches(5.0),
         "Ovsynch / Double-AI",
         ["GnRH (D0) → PGF2α (D7) → GnRH (D9) → AI 16h later",
          "Synchronises ovulation precisely",
          "Useful in herds with poor heat detection",
          "Modified Ovsynch / Co-Synch / Double-Ovsynch for repeaters"],
         head_color=ACCENT2)
add_footer(s, 14, 19)

# ============================================================
# SLIDE 15 — INTRA-UTERINE THERAPY
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Intra-Uterine & Adjunct Therapy")
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
    ("Herbal / Ayurvedic preparations (e.g. Aloes compound, Janova, "
     "Prajana) — adjunct support.", 0),
], size=16)
add_footer(s, 15, 19)

# ============================================================
# SLIDE 16 — BUFFALO-SPECIFIC REPEAT BREEDING (NEW)
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Buffalo-Specific Repeat Breeding",
          "Why buffaloes pose unique reproductive challenges")

# image (large, left side)
add_image_with_frame(s, BUFFALO, Inches(0.85), Inches(1.9),
                     Inches(4.5), Inches(3.1),
                     "Murrah buffalo (Bubalus bubalis)")

# right side bullets
add_bullets(s, [
    "Silent / sub-oestrus is common — up to 50–60 % of heats are missed "
    "without close observation.",
    "Late maturity (36–42 months) and longer post-partum anoestrus "
    "(90–150 days) than cows.",
    "Strongly seasonal breeder — short photoperiod (Oct–Feb) is the "
    "peak breeding season.",
    "Summer infertility: heat stress depresses LH surge, oocyte quality "
    "and embryo survival.",
    "Smaller, deeper-seated ovaries with fewer antral follicles than cows.",
    "Higher incidence of true anoestrus, ovarian inactivity and "
    "anovulatory cysts.",
], left=Inches(5.6), top=Inches(1.9), width=Inches(7.3),
   height=Inches(3.5), size=14)

# bottom — practical management cards
add_card(s, Inches(0.85), Inches(5.2), Inches(3.95), Inches(1.85),
         "Heat Detection",
         ["Observe 4–5 times/day, esp. early morning & late evening",
          "Use teaser bull / pedometer / activity meter",
          "Watch for vulvar oedema, mucus, bellowing"])
add_card(s, Inches(4.95), Inches(5.2), Inches(3.95), Inches(1.85),
         "Hormonal / Synchrony",
         ["Ovsynch / Heatsynch / CIDR-based protocols work well",
          "Double-Ovsynch & Pre-synch improve summer conception",
          "GnRH at AI improves ovulation in late-cycling buffaloes"],
         head_color=ACCENT)
add_card(s, Inches(9.05), Inches(5.2), Inches(3.85), Inches(1.85),
         "Environment & Nutrition",
         ["Wallowing tank / showers in summer (>32 °C)",
          "Shade, fans, fly control",
          "Bypass fat & mineral supplementation"],
         head_color=ACCENT2)

add_footer(s, 16, 19)

# ============================================================
# SLIDE 17 — PREVENTION
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Prevention — Better Than Cure")

add_card(s, Inches(0.85), Inches(1.95), Inches(5.9), Inches(2.4),
         "Herd-Level Practices",
         ["Routine post-partum examination on Day 30–45",
          "Maintain calving interval 12–13 months",
          "Vaccinate against Brucella, IBR, BVD, Leptospira",
          "Bio-security; quarantine new animals"])
add_card(s, Inches(7.0), Inches(1.95), Inches(5.9), Inches(2.4),
         "Cow-Level Practices",
         ["Balanced ration with adequate energy, protein & minerals",
          "Body condition scoring at dry-off, calving and AI",
          "Heat detection 3× daily / use of pedometers / activity collars",
          "AI by trained technicians using fertile, properly thawed semen"],
         head_color=ACCENT)
add_card(s, Inches(0.85), Inches(4.5), Inches(5.9), Inches(2.4),
         "Environment",
         ["Provide shade, fans, sprinklers in summer",
          "Comfortable, non-slippery flooring",
          "Adequate water and clean housing",
          "Reduce overcrowding & social stress"],
         head_color=ACCENT2)
add_card(s, Inches(7.0), Inches(4.5), Inches(5.9), Inches(2.4),
         "Records & Decisions",
         ["Maintain individual breeding records / software",
          "Identify chronic repeaters early",
          "Selective culling of non-responders after 5–6 services",
          "Periodic herd fertility audit"], head_color=PRIMARY)

add_footer(s, 17, 19)

# ============================================================
# SLIDE 18 — RECENT ADVANCES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Recent Advances")
add_bullets(s, [
    "Doppler ultrasonography of CL — early detection of luteal insufficiency.",
    "Embryo transfer (ET) in chronic repeat breeders bypasses fertilization "
    "and early embryo problems.",
    "In-vitro fertilization (IVF / OPU-IVP) for valuable but persistent repeaters.",
    "Sexed semen with optimised AI timing for high-genetic-merit cows.",
    "Genomic selection — identifying cows with genetic predisposition to "
    "early embryonic death.",
    "Use of activity-monitoring collars & AI-based heat detection systems.",
    "Anti-oxidant therapy (Vitamin E, Selenium, Astaxanthin) to improve "
    "oocyte and embryo quality.",
    "Targeted reproductive ultrasonography & uterine biopsy for refractory cases.",
], size=17)
add_footer(s, 18, 19)

# ============================================================
# SLIDE 19 — CONCLUSION & REFERENCES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Conclusion & References")

box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.85), Inches(1.85),
                         Inches(11.6), Inches(2.2))
box.line.color.rgb = ACCENT; box.line.width = Pt(2)
box.fill.solid(); box.fill.fore_color.rgb = WHITE
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = Inches(0.25); tf.margin_right = Inches(0.25)
tf.margin_top = Inches(0.15)

p1 = tf.paragraphs[0]
r1 = p1.add_run(); r1.text = "Conclusion"
r1.font.size = Pt(18); r1.font.bold = True
r1.font.color.rgb = PRIMARY; r1.font.name = "Calibri"

for line in [
    "Repeat breeding is a multifactorial syndrome — a symptom rather than a "
    "disease.",
    "Successful management requires accurate diagnosis, correction of "
    "nutrition, infection and hormonal imbalance, plus excellent AI practices.",
    "Buffaloes need extra attention to silent oestrus, seasonality and "
    "summer heat-stress mitigation.",
    "Prevention — through good husbandry, heat detection and herd health "
    "programs — is far more economical than treatment.",
]:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "• " + line
    r.font.size = Pt(13); r.font.color.rgb = DARK; r.font.name = "Calibri"
    p.space_after = Pt(2)

ref_box = s.shapes.add_textbox(Inches(0.85), Inches(4.2),
                               Inches(11.6), Inches(2.7))
rtf = ref_box.text_frame; rtf.word_wrap = True
p = rtf.paragraphs[0]
r = p.add_run(); r.text = "References (selected)"
r.font.size = Pt(18); r.font.bold = True
r.font.color.rgb = PRIMARY; r.font.name = "Calibri"

refs = [
    "Roberts, S.J. — Veterinary Obstetrics & Genital Diseases (Theriogenology), 3rd Ed.",
    "Arthur, Noakes, Pearson & Parkinson — Veterinary Reproduction & Obstetrics, 11th Ed.",
    "Hafez E.S.E. & Hafez B. — Reproduction in Farm Animals, 7th Ed.",
    "Peters A.R. & Ball P.J.H. — Reproduction in Cattle, 3rd Ed.",
    "Bartlett P.C. et al. — Repeat breeder syndrome in dairy cattle, JDS.",
    "Gustafsson H. & Emanuelson U. — Characterisation of the repeat breeding "
    "syndrome in Swedish dairy cattle, Acta Vet. Scand.",
    "Perera B.M.A.O. — Reproductive cycles of buffalo, Anim. Reprod. Sci.",
]
for ref in refs:
    p = rtf.add_paragraph()
    rr = p.add_run(); rr.text = "• " + ref
    rr.font.size = Pt(12); rr.font.color.rgb = DARK; rr.font.name = "Calibri"
    p.space_after = Pt(2)

# image-credit footnote
cred = s.shapes.add_textbox(Inches(0.85), Inches(6.8),
                            Inches(11.6), Inches(0.3))
cp = cred.text_frame.paragraphs[0]
cr = cp.add_run()
cr.text = ("Image credits: photographs of cows and buffaloes are from "
           "Wikimedia Commons (CC BY-SA / public domain).")
cr.font.size = Pt(9); cr.font.italic = True
cr.font.color.rgb = GRAY; cr.font.name = "Calibri"

# Thank you ribbon
ty = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                        Inches(4.7), Inches(7.05),
                        Inches(4.0), Inches(0.4))
ty.line.fill.background(); ty.fill.solid()
ty.fill.fore_color.rgb = ACCENT
ttf = ty.text_frame
ttf.margin_top = Inches(0.02); ttf.margin_bottom = Inches(0.02)
tp = ttf.paragraphs[0]; tp.alignment = PP_ALIGN.CENTER
tr = tp.add_run(); tr.text = "THANK YOU"
tr.font.size = Pt(16); tr.font.bold = True
tr.font.color.rgb = WHITE; tr.font.name = "Calibri"

# Save
out = os.path.join(HERE, "Repeat_Breeding.pptx")
prs.save(out)
print(f"Saved: {out}")
print(f"Total slides: {len(prs.slides)}")
