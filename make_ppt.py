"""Generate a PowerPoint presentation on Repeat Breeding in cattle."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# ---------- Theme colors ----------
PRIMARY = RGBColor(0x0B, 0x3D, 0x2E)      # deep green
ACCENT  = RGBColor(0xC9, 0xA2, 0x27)      # gold
LIGHT   = RGBColor(0xF5, 0xF1, 0xE8)      # cream
DARK    = RGBColor(0x1F, 0x2A, 0x24)      # near-black green
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GRAY    = RGBColor(0x55, 0x55, 0x55)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

BLANK = prs.slide_layouts[6]


def add_bg(slide, color=LIGHT):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.line.fill.background()
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    return bg


def add_side_bar(slide):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.35), SH)
    bar.line.fill.background()
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.35), 0,
                                    Inches(0.08), SH)
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
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(0.35),
                                  Inches(12.2), Inches(0.9))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size = Pt(34); r.font.bold = True
    r.font.color.rgb = PRIMARY; r.font.name = "Calibri"

    # gold underline
    u = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(0.7), Inches(1.18),
                               Inches(1.2), Inches(0.07))
    u.line.fill.background(); u.fill.solid()
    u.fill.fore_color.rgb = ACCENT

    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.7), Inches(1.28),
                                      Inches(12.2), Inches(0.4))
        sp = sb.text_frame.paragraphs[0]
        sr = sp.add_run(); sr.text = subtitle
        sr.font.size = Pt(14); sr.font.italic = True
        sr.font.color.rgb = GRAY; sr.font.name = "Calibri"


def add_bullets(slide, bullets, left=Inches(0.85), top=Inches(1.8),
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
        if level == 0:
            r.font.bold = False


def add_card(slide, left, top, width, height, title, body, head_color=PRIMARY):
    head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.55))
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


# ============================================================
# SLIDE 1 — TITLE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, PRIMARY)

# decorative gold band
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.1),
                          SW, Inches(1.4))
band.line.fill.background(); band.fill.solid()
band.fill.fore_color.rgb = ACCENT

# title
tb = s.shapes.add_textbox(Inches(0.7), Inches(3.2), Inches(12), Inches(0.8))
p = tb.text_frame.paragraphs[0]
r = p.add_run(); r.text = "REPEAT BREEDING"
r.font.size = Pt(54); r.font.bold = True
r.font.color.rgb = PRIMARY; r.font.name = "Calibri"

tb2 = s.shapes.add_textbox(Inches(0.7), Inches(3.95), Inches(12), Inches(0.6))
p2 = tb2.text_frame.paragraphs[0]
r2 = p2.add_run()
r2.text = "Causes, Diagnosis, Treatment & Prevention in Dairy Cattle"
r2.font.size = Pt(22); r2.font.color.rgb = DARK
r2.font.italic = True; r2.font.name = "Calibri"

# top tag
tag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(0.7),
                        Inches(3.0), Inches(0.45))
tag.line.fill.background(); tag.fill.solid()
tag.fill.fore_color.rgb = ACCENT
ttf = tag.text_frame; ttf.margin_left = Inches(0.15)
tp = ttf.paragraphs[0]
tr = tp.add_run(); tr.text = "ANIMAL REPRODUCTION"
tr.font.size = Pt(13); tr.font.bold = True
tr.font.color.rgb = PRIMARY; tr.font.name = "Calibri"

# bottom info
info = s.shapes.add_textbox(Inches(0.7), Inches(6.4), Inches(12), Inches(0.6))
ip = info.text_frame.paragraphs[0]
ir = ip.add_run()
ir.text = "Veterinary Gynaecology & Obstetrics  |  Seminar Presentation"
ir.font.size = Pt(14); ir.font.color.rgb = WHITE; ir.font.name = "Calibri"

# ============================================================
# SLIDE 2 — OUTLINE
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s); add_side_bar(s)
add_title(s, "Presentation Outline")

items = [
    ("1.  Introduction & Definition",      Inches(0.85), Inches(1.85)),
    ("2.  Incidence & Economic Impact",    Inches(0.85), Inches(2.45)),
    ("3.  Etiology / Causes",              Inches(0.85), Inches(3.05)),
    ("4.  Pathophysiology",                Inches(0.85), Inches(3.65)),
    ("5.  Diagnosis & Clinical Approach",  Inches(0.85), Inches(4.25)),
    ("6.  Treatment & Management",         Inches(6.85), Inches(1.85)),
    ("7.  Hormonal Therapy Protocols",     Inches(6.85), Inches(2.45)),
    ("8.  Prevention Strategies",          Inches(6.85), Inches(3.05)),
    ("9.  Recent Advances",                Inches(6.85), Inches(3.65)),
    ("10. Conclusion & References",        Inches(6.85), Inches(4.25)),
]
for text, l, t in items:
    bullet = s.shapes.add_shape(MSO_SHAPE.OVAL, l, t + Inches(0.1),
                                Inches(0.18), Inches(0.18))
    bullet.line.fill.background(); bullet.fill.solid()
    bullet.fill.fore_color.rgb = ACCENT
    tx = s.shapes.add_textbox(l + Inches(0.35), t, Inches(5.5), Inches(0.5))
    p = tx.text_frame.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(18); r.font.color.rgb = DARK; r.font.name = "Calibri"

add_footer(s, 2, 18)

# ============================================================
# SLIDE 3 — INTRODUCTION
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Introduction",
          "Reproductive efficiency is the backbone of profitable dairy farming")
add_bullets(s, [
    "Reproduction is the most important factor governing the economic success "
    "of any dairy enterprise.",
    "A normal fertile cow should conceive within 60–90 days post-partum and "
    "deliver one calf every 12–13 months.",
    "Failure to conceive after repeated inseminations — despite normal estrous "
    "cycles and apparently healthy genitalia — is termed REPEAT BREEDING.",
    "It is one of the most frustrating and economically damaging reproductive "
    "disorders in cattle and buffaloes worldwide.",
    "Often multifactorial — involving the cow, the bull/semen, the inseminator "
    "and the environment.",
])
add_footer(s, 3, 18)

# ============================================================
# SLIDE 4 — DEFINITION
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Definition")

# big quote-style definition
box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.85), Inches(1.85),
                         Inches(11.6), Inches(1.7))
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

# key points
add_bullets(s, [
    "Also known as: Repeat Breeder Syndrome / Repeat Breeder Cow (RBC).",
    "Synonym in older literature: 'Conception Failure'.",
    ("Key criteria — (a) regular oestrous cycles, (b) ≥3 services, "
     "(c) no palpable abnormality, (d) fertile semen, (e) competent AI."),
    "Differentiate from: Anoestrus, sub-oestrus, true infertility and sterility.",
], top=Inches(3.8), size=17)
add_footer(s, 4, 18)

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

# Economic impact cards
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
          "National-level losses run into thousands of crores"])

add_footer(s, 5, 18)

# ============================================================
# SLIDE 6 — CAUSES OVERVIEW
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Etiology — An Overview",
          "Repeat breeding is multifactorial; causes are grouped into four broad heads")

# Four equal cards
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
          "Stress at AI"])
add_card(s, Inches(10.3), Inches(1.95), Inches(2.6), Inches(4.9),
         "4. Genetic / Other",
         ["Inbreeding",
          "Lethal genes",
          "Freemartinism",
          "White heifer disease",
          "Age & parity"], head_color=ACCENT)

add_footer(s, 6, 18)

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
add_footer(s, 7, 18)

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
add_footer(s, 8, 18)

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
add_footer(s, 9, 18)

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
add_footer(s, 10, 18)

# ============================================================
# SLIDE 11 — PATHOPHYSIOLOGY
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Pathophysiology — Where Things Go Wrong")

# horizontal flow boxes
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

# arrow row
for i in range(len(labels) - 1):
    arr_left = left + (i + 1) * w + i * Inches(0.07) - Inches(0.02)
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arr_left, top + Inches(0.42),
                           Inches(0.13), Inches(0.18))
    a.line.fill.background(); a.fill.solid()
    a.fill.fore_color.rgb = DARK

# below — failure points
add_bullets(s, [
    "Failure of ovulation or asynchronous ovulation → no fertilization.",
    "Sperm transport disturbance / hostile cervico-uterine environment.",
    "Fertilization failure — poor oocyte / sperm quality, anti-sperm antibodies.",
    "Early embryonic death (Day 8–16) — most common in repeat breeders.",
    "Failure of maternal recognition of pregnancy → CL regresses → return to oestrus.",
], top=Inches(3.6), size=16)
add_footer(s, 11, 18)

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
add_footer(s, 12, 18)

# ============================================================
# SLIDE 13 — TREATMENT GENERAL
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Treatment — General Principles",
          "Treat the cause, not just the symptom")

add_bullets(s, [
    "Correct underlying nutrition & body condition (BCS 3.0–3.5 at AI).",
    "Treat genital tract infection (intra-uterine antibiotics, antiseptics).",
    "Correct hormonal imbalance with appropriate protocols.",
    "Improve heat detection and AI timing (AM-PM rule / use of activity meters).",
    "Use proven, high-quality semen; ensure proper thawing & deposition.",
    "Reduce stress — provide shade, water, comfortable housing, fly control.",
    "Maintain proper records; cull chronic non-responders after 5–6 services.",
], top=Inches(1.85), size=18)
add_footer(s, 13, 18)

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
          "Modified Ovsynch / Co-Synch / Double-Ovsynch for repeat breeders"])

add_footer(s, 14, 18)

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
add_footer(s, 15, 18)

# ============================================================
# SLIDE 16 — PREVENTION
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
          "AI by trained technicians using fertile, properly thawed semen"], head_color=ACCENT)
add_card(s, Inches(0.85), Inches(4.5), Inches(5.9), Inches(2.4),
         "Environment",
         ["Provide shade, fans, sprinklers in summer",
          "Comfortable, non-slippery flooring",
          "Adequate water and clean housing",
          "Reduce overcrowding & social stress"])
add_card(s, Inches(7.0), Inches(4.5), Inches(5.9), Inches(2.4),
         "Records & Decisions",
         ["Maintain individual breeding records / software",
          "Identify chronic repeaters early",
          "Selective culling of non-responders after 5–6 services",
          "Periodic herd fertility audit"], head_color=ACCENT)

add_footer(s, 16, 18)

# ============================================================
# SLIDE 17 — RECENT ADVANCES
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
add_footer(s, 17, 18)

# ============================================================
# SLIDE 18 — CONCLUSION & REFERENCES
# ============================================================
s = prs.slides.add_slide(BLANK); add_bg(s); add_side_bar(s)
add_title(s, "Conclusion & References")

# conclusion box
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
    "Prevention — through good husbandry, heat detection and herd health "
    "programs — is far more economical than treatment.",
    "Early identification and timely intervention can return most repeat "
    "breeders to productive reproduction.",
]:
    p = tf.add_paragraph()
    r = p.add_run(); r.text = "• " + line
    r.font.size = Pt(13); r.font.color.rgb = DARK; r.font.name = "Calibri"
    p.space_after = Pt(2)

# references
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
]
for ref in refs:
    p = rtf.add_paragraph()
    rr = p.add_run(); rr.text = "• " + ref
    rr.font.size = Pt(12); rr.font.color.rgb = DARK; rr.font.name = "Calibri"
    p.space_after = Pt(2)

# Thank you ribbon
ty = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                        Inches(4.7), Inches(6.7),
                        Inches(4.0), Inches(0.5))
ty.line.fill.background(); ty.fill.solid()
ty.fill.fore_color.rgb = ACCENT
ttf = ty.text_frame
tp = ttf.paragraphs[0]; tp.alignment = PP_ALIGN.CENTER
tr = tp.add_run(); tr.text = "THANK YOU"
tr.font.size = Pt(18); tr.font.bold = True
tr.font.color.rgb = PRIMARY; tr.font.name = "Calibri"

# Save
out = "/projects/sandbox/MAIN-KIRO/Repeat_Breeding.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Total slides: {len(prs.slides)}")
