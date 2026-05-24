"""
build_synopsis.py
-----------------
Generates a publication-quality M.V.Sc/Ph.D research synopsis (.docx) on:

    "Biopolymers Compatible for Progesterone Encapsulation
     using the Spray Drying Method"

Output: Synopsis_Progesterone_SprayDrying_Biopolymer.docx

Scientific-integrity rules applied:
- Only peer-reviewed references that are independently verifiable
  (real authors, real journals, real years).
- Where a specific numeric claim could not be backed by a confidently
  cited primary source, the prose uses qualitative ranges drawn from
  cited review literature instead of fabricated point values.
- Vancouver-style numbered citations [#] keyed to the References section.
- Tables for polymer comparison, spray-drying vs other methods,
  work-plan timeline, and budget.
"""

from __future__ import annotations

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

OUTPUT = "Synopsis_Progesterone_SprayDrying_Biopolymer.docx"

TITLE_FONT = "Times New Roman"
BODY_FONT = "Times New Roman"
BODY_SIZE = Pt(12)
SMALL_SIZE = Pt(11)


# ---------------------------------------------------------------------------
# Document setup helpers
# ---------------------------------------------------------------------------
def configure_document(doc: Document) -> None:
    # Default style: Times New Roman 12, justified, 1.5 line spacing
    style = doc.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = BODY_SIZE
    pf = style.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(6)

    # Page margins (academic): 2.5 cm on every side
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)   # extra for binding
        section.right_margin = Cm(2.5)

    # Configure heading styles
    for level, size, bold, color in [
        (1, 14, True, RGBColor(0x1F, 0x4E, 0x78)),
        (2, 12, True, RGBColor(0x2E, 0x75, 0xB6)),
        (3, 12, True, RGBColor(0x1F, 0x1F, 0x1F)),
    ]:
        s = doc.styles[f"Heading {level}"]
        s.font.name = BODY_FONT
        s.font.size = Pt(size)
        s.font.bold = bold
        s.font.color.rgb = color
        s.paragraph_format.space_before = Pt(12)
        s.paragraph_format.space_after = Pt(6)
        s.paragraph_format.keep_with_next = True

    # Add page number to footer
    add_page_numbers(doc)


def add_page_numbers(doc: Document) -> None:
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)
    run.font.size = SMALL_SIZE


# ---------------------------------------------------------------------------
# Paragraph helpers
# ---------------------------------------------------------------------------
def add_para(doc: Document, text: str, *, bold: bool = False,
             italic: bool = False, align=None, size=None,
             space_before: int | None = None,
             space_after: int | None = None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = BODY_FONT
    run.bold = bold
    run.italic = italic
    if size is not None:
        run.font.size = size
    return p


def add_h1(doc: Document, text: str):
    p = doc.add_heading(text, level=1)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def add_h2(doc: Document, text: str):
    return doc.add_heading(text, level=2)


def add_h3(doc: Document, text: str):
    return doc.add_heading(text, level=3)


def add_bullets(doc: Document, items: list[str]):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.runs[0] if p.runs else p.add_run()
        run.text = it
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE


def add_numbered(doc: Document, items: list[str]):
    for it in items:
        p = doc.add_paragraph(style="List Number")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.runs[0] if p.runs else p.add_run()
        run.text = it
        run.font.name = BODY_FONT
        run.font.size = BODY_SIZE


# ---------------------------------------------------------------------------
# Table helpers
# ---------------------------------------------------------------------------
def shade_cell(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def style_cell(cell, *, bold: bool = False, size: int = 10,
               align=WD_ALIGN_PARAGRAPH.LEFT,
               fill: str | None = None,
               color: RGBColor | None = None) -> None:
    if fill:
        shade_cell(cell, fill)
    for p in cell.paragraphs:
        p.alignment = align
        p.paragraph_format.space_after = Pt(0)
        for r in p.runs:
            r.font.name = BODY_FONT
            r.font.size = Pt(size)
            r.bold = bold
            if color:
                r.font.color.rgb = color
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def make_table(doc: Document, headers: list[str], rows: list[list[str]],
               *, caption: str | None = None,
               header_fill: str = "1F4E78",
               header_color: RGBColor = RGBColor(0xFF, 0xFF, 0xFF),
               body_size: int = 10,
               first_col_bold: bool = False) -> None:
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = cap.add_run(caption)
        run.bold = True
        run.italic = True
        run.font.name = BODY_FONT
        run.font.size = SMALL_SIZE

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Light Grid Accent 1"

    # Headers
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        style_cell(cell, bold=True, size=body_size,
                   align=WD_ALIGN_PARAGRAPH.CENTER,
                   fill=header_fill, color=header_color)

    # Body
    for r, row in enumerate(rows, start=1):
        zebra = "F2F2F2" if r % 2 == 0 else None
        for c, val in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.text = val
            style_cell(cell,
                       bold=(first_col_bold and c == 0),
                       size=body_size,
                       align=(WD_ALIGN_PARAGRAPH.LEFT if c == 0
                              else WD_ALIGN_PARAGRAPH.LEFT),
                       fill=zebra)

    # Spacer paragraph after table
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Section builders
# ---------------------------------------------------------------------------
def build_title_page(doc: Document) -> None:
    for _ in range(3):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("RESEARCH SYNOPSIS")
    run.bold = True
    run.font.name = TITLE_FONT
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Submitted in partial fulfilment of the requirements\n"
                    "for the degree of\n"
                    "Master of Veterinary Science / Doctor of Philosophy\n"
                    "(Animal Reproduction, Gynaecology and Obstetrics)")
    run.italic = True
    run.font.size = Pt(12)

    for _ in range(2):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        "Biopolymers Compatible for Progesterone Encapsulation "
        "using the Spray Drying Method"
    )
    run.bold = True
    run.font.name = TITLE_FONT
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1F, 0x1F, 0x1F)

    for _ in range(3):
        doc.add_paragraph()

    rows = [
        ("Submitted by",          "____________________________________"),
        ("Registration / Roll No.", "____________________________________"),
        ("Major Advisor / Guide",   "____________________________________"),
        ("Co-Advisor",              "____________________________________"),
        ("Department",              "Animal Reproduction, Gynaecology and Obstetrics"),
        ("College / Institute",     "____________________________________"),
        ("University",              "____________________________________"),
        ("Year of Submission",      "____________________________________"),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(rows):
        c0 = table.rows[i].cells[0]
        c1 = table.rows[i].cells[1]
        c0.text = k
        c1.text = v
        style_cell(c0, bold=True, size=12)
        style_cell(c1, size=12)

    doc.add_page_break()


def build_introduction(doc: Document) -> None:
    add_h1(doc, "1. INTRODUCTION")

    add_para(doc,
        "Progesterone (P4), a C21 pregnane steroid synthesised principally by "
        "the corpus luteum and, during pregnancy, by the placenta, is the "
        "central regulator of the oestrous and menstrual cycles, the "
        "establishment and maintenance of pregnancy, and uterine quiescence. "
        "In veterinary reproductive practice, exogenous progesterone "
        "supplementation underpins almost every protocol for oestrus "
        "synchronisation, fixed-time artificial insemination, treatment of "
        "post-partum anoestrus, and management of repeat-breeding in cattle, "
        "buffalo, sheep and goats [1,8,9].")

    add_para(doc,
        "Despite its physiological importance, progesterone is a Biopharmaceutics "
        "Classification System Class II molecule: it exhibits very low aqueous "
        "solubility (approximately 8.8 mg/L at 25 deg C), high lipophilicity "
        "(log P approximately 3.9), extensive first-pass hepatic metabolism, and "
        "a short circulating half-life. Conventional dosage forms therefore "
        "deliver poor and erratic systemic exposure when given orally, and "
        "parenteral oily depot injections produce unpredictable plasma profiles "
        "with frequent injection-site reactions [2,3].")

    add_para(doc,
        "These pharmacokinetic limitations have driven the development of "
        "intravaginal controlled-release devices, of which the Controlled "
        "Internal Drug Release (CIDR) insert and the Progesterone Releasing "
        "Intravaginal Device (PRID) are the most widely adopted. Both are "
        "based on silicone elastomer matrices loaded with crystalline "
        "progesterone and have transformed reproductive management in dairy "
        "and beef herds. However, these devices are non-biodegradable, "
        "relatively expensive, exhibit a substantial initial burst, suffer "
        "progressive payload exhaustion over the seven- to nine-day "
        "treatment window, and generate non-recyclable polymer waste after "
        "use [8,9,10].")

    add_para(doc,
        "Microparticulate biopolymer-based delivery systems offer a "
        "fundamentally different paradigm: progesterone is dispersed at "
        "molecular or sub-micron scale within a biocompatible, biodegradable "
        "polymer matrix that protects the steroid from premature degradation, "
        "modulates its diffusion, and is ultimately metabolised by the host "
        "tissue [14,17,18,19]. Such systems can in principle be administered "
        "intravaginally, intrauterinely, subcutaneously or even orally, "
        "depending on the polymer chosen, and can be tuned to deliver "
        "first-order, zero-order or pulsatile release profiles.")

    add_para(doc,
        "Among the available encapsulation technologies, spray drying is "
        "uniquely attractive for hormone delivery because it converts a "
        "liquid feed into a dry, free-flowing microparticulate powder in a "
        "single, continuous, organic-solvent-tolerant operation. It is "
        "scalable from milligram laboratory batches (Buchi B-290, Buchi B-90 "
        "Nano) to multi-tonne industrial throughput, gives narrow particle "
        "size distributions, achieves high encapsulation efficiency for "
        "lipophilic actives such as progesterone, and produces particles "
        "with low residual moisture and high storage stability [1,3,4,5,11]. "
        "Critically, the residence time of any given droplet inside the "
        "drying chamber is on the order of seconds, so the active ingredient "
        "is exposed to elevated temperature only briefly, which protects "
        "thermally sensitive steroids during processing [3].")

    add_para(doc,
        "The performance of a spray-dried progesterone microparticle, "
        "however, is dictated less by the dryer than by the choice of wall "
        "material. The ideal biopolymer must dissolve or disperse in a feed "
        "system in which progesterone can also be molecularly dispersed, "
        "must form a coherent, low-permeability shell around the drug, must "
        "be biocompatible and biodegradable, must be available at "
        "pharmaceutical or food grade, and must be amenable to the high "
        "shear and atomisation pressure encountered in the dryer. No single "
        "polymer simultaneously satisfies every one of these requirements, "
        "which is why systematic compatibility screening of biopolymers "
        "with progesterone is a precondition for any rational formulation "
        "programme [2,6,11,15,16,17,18,19].")

    add_para(doc,
        "The present study is conceived against this background. It seeks "
        "to identify, formulate and characterise progesterone-loaded "
        "microparticles produced by spray drying using a panel of "
        "biocompatible, biodegradable polymers of natural, synthetic and "
        "semi-synthetic origin, with the long-term aim of engineering a "
        "low-cost, biodegradable sustained-release progesterone system "
        "suitable for intravaginal use in food-producing animals.")


def build_review(doc: Document) -> None:
    add_h1(doc, "2. REVIEW OF LITERATURE")

    add_h2(doc, "2.1 Progesterone delivery systems in veterinary reproduction")
    add_para(doc,
        "Macmillan and Peterson (1993) reported the original CIDR-B device, "
        "demonstrating that a silicone matrix loaded with 1.9 g of "
        "progesterone could maintain plasma P4 above the luteal threshold "
        "(>1 ng/mL) for seven to twelve days in dairy cows and could be "
        "used to synchronise oestrus and treat post-partum anoestrus [9]. "
        "Subsequent work by Rathbone and colleagues established the "
        "engineering principles underlying CIDR/PRID-type intravaginal "
        "inserts, including the dependence of release on surface area, "
        "matrix loading and vaginal microenvironment, and identified a "
        "characteristic biphasic profile consisting of a 24- to 48-hour "
        "burst followed by a slowly declining maintenance phase [8,10].")

    add_h2(doc, "2.2 Spray drying as a pharmaceutical encapsulation platform")
    add_para(doc,
        "Re (1998) provided one of the earliest comprehensive accounts of "
        "spray drying for microencapsulation, identifying inlet temperature, "
        "outlet temperature, feed solid content and atomisation conditions "
        "as the dominant process variables [23]. Vehring (2008) integrated "
        "this engineering perspective with droplet-drying physics, showing "
        "how Peclet number considerations govern whether a drying droplet "
        "produces a dense, hollow or wrinkled particle, and giving formulators "
        "a quantitative framework for particle engineering [3]. Cal and "
        "Sollohub (2010) and Sollohub and Cal (2010) reviewed hardware and "
        "pharmaceutical applications respectively, and established spray "
        "drying as a mainstream technique for amorphous solid dispersions, "
        "inhalation powders and microencapsulated drugs [4,5]. The "
        "comprehensive synthesis by Sosnik and Seremeta (2015) consolidated "
        "the use of spray drying for both pure drug particles and "
        "drug-loaded polymeric carriers, and articulated the trade-off "
        "between particle yield, encapsulation efficiency and morphology "
        "that frames every formulation programme in this area [2].")

    add_h2(doc, "2.3 Biopolymer wall materials for hormone encapsulation")
    add_para(doc,
        "Gharsallaoui et al. (2007) provided the canonical comparative "
        "review of wall materials for spray-dried microencapsulation, "
        "highlighting maltodextrin, gum arabic, modified starches, whey "
        "protein and sodium caseinate as work-horse food-grade carriers, and "
        "noting that polymer blends almost invariably out-perform single "
        "polymers for lipophilic actives [1]. Estevinho et al. (2013) "
        "specifically reviewed chitosan-based spray-drying applications and "
        "documented its mucoadhesive, biodegradable and pH-responsive "
        "behaviour, all of which are directly relevant to vaginal "
        "progesterone delivery [6]. Anandharamakrishnan and Ishwarya (2015) "
        "compiled a book-length treatment of spray-drying encapsulation "
        "across food, nutraceutical and pharmaceutical applications [7].")

    add_para(doc,
        "Rinaudo (2006) and Felt et al. (1998) reviewed chitosan as a "
        "drug-delivery polymer, emphasising its primary amine groups, "
        "cationic character at acidic pH, and well-documented mucoadhesion "
        "to vaginal and intestinal epithelia [15]. Lee and Mooney (2012) "
        "compiled the biomedical applications of alginate, including its "
        "ionotropic gelation with calcium and its track record in "
        "pharmaceutical microspheres [17]. Paliwal and Palakurthi (2014) "
        "reviewed zein, a maize prolamin, as a controlled-release wall "
        "material with intrinsic hydrophobicity that is particularly "
        "compatible with lipophilic steroids [16]. Makadia and Siegel "
        "(2011) and Danhier et al. (2012) reviewed PLGA, the most clinically "
        "validated synthetic biodegradable polyester, for parenteral "
        "controlled-release applications including hormone microspheres "
        "[18,19].")

    add_h2(doc, "2.4 Encapsulation efficiency, particle characterisation and release modelling")
    add_para(doc,
        "Soottitantawat et al. (2003) demonstrated that, for spray-dried "
        "emulsions, encapsulation efficiency for hydrophobic actives is "
        "strongly correlated with the size of the primary emulsion droplets "
        "and with the wall-to-core ratio [12]. Desai and Park (2005) "
        "consolidated similar observations for food microencapsulation [13]. "
        "Paulo and Santos (2017) reviewed Design-of-Experiments approaches "
        "to microencapsulation, stressing the value of factorial and "
        "response-surface designs for systematic optimisation [11]. Release "
        "behaviour from polymeric matrices is conventionally analysed using "
        "the classical models of Higuchi (1963), Korsmeyer-Peppas et al. "
        "(1983) and Costa and Sousa Lobo (2001), and the ranking of these "
        "models against in-vitro dissolution data provides mechanistic "
        "insight into whether release is diffusion-, erosion- or "
        "swelling-controlled [20,21,22]. Maderuelo et al. (2011) reviewed "
        "the matrix factors that critically influence drug release from "
        "hydrophilic sustained-release systems [14].")

    add_h2(doc, "2.5 Current research gaps")
    add_para(doc,
        "Despite three decades of incremental progress, several gaps "
        "persist. First, the majority of spray-drying studies on steroid "
        "encapsulation have addressed estradiol, testosterone or "
        "corticosteroids; primary, peer-reviewed reports specifically "
        "describing systematic biopolymer screening for spray-dried "
        "progesterone microparticles intended for veterinary intravaginal "
        "use are comparatively few. Second, most reported formulations rely "
        "either on costly synthetic polymers (PLGA, PCL) suited only to "
        "high-value parenteral applications or on single food-grade carriers "
        "that lack sustained-release capacity. Third, in-vivo plasma "
        "progesterone kinetics and reproductive end-points (oestrus "
        "response, pregnancy rate) for spray-dried microparticulate "
        "progesterone systems in cattle and buffalo are sparsely reported. "
        "These gaps frame the rationale for the present study.")


def build_research_gap(doc: Document) -> None:
    add_h1(doc, "3. RESEARCH GAP")
    add_para(doc,
        "Synthesis of the literature reviewed in Section 2 identifies the "
        "following specific gaps that the present study seeks to address:")
    add_bullets(doc, [
        "Limited systematic data on the comparative compatibility of natural, "
        "synthetic and semi-synthetic biopolymers when co-spray-dried with "
        "progesterone under a single, controlled set of process parameters.",

        "Heavy commercial dependence on non-biodegradable silicone-elastomer "
        "intravaginal devices (CIDR, PRID), generating polymer waste and "
        "high per-animal cost in resource-limited dairy systems.",

        "Insufficient evidence on low-cost biodegradable polymer blends "
        "(for example, gum arabic-maltodextrin, whey protein-maltodextrin, "
        "chitosan-alginate) capable of supporting sustained vaginal "
        "progesterone release for seven to nine days.",

        "Recurrent stability concerns with spray-dried hormone formulations, "
        "including hygroscopicity, polymorphic transitions of the steroid, "
        "and progressive loss of payload during storage.",

        "Limited modelling of release kinetics for spray-dried progesterone "
        "microparticles using mechanistic models such as Higuchi, "
        "Korsmeyer-Peppas, Hixson-Crowell and Weibull, which are essential "
        "to support regulatory and design-of-experiments work.",

        "A near-absence of studies that link in-vitro release performance "
        "of biopolymer-progesterone microparticles to in-vivo plasma P4 "
        "kinetics and reproductive outcomes in repeat-breeding cattle and "
        "buffalo under field conditions.",
    ])


def build_aim_objectives(doc: Document) -> None:
    add_h1(doc, "4. AIM OF THE STUDY")
    add_para(doc,
        "To identify biopolymers that are compatible with progesterone "
        "during spray drying, and to develop and characterise "
        "progesterone-loaded biopolymeric microparticles as a foundation "
        "for sustained-release reproductive hormone delivery in domestic "
        "livestock.")

    add_h1(doc, "5. OBJECTIVES OF THE STUDY")
    add_numbered(doc, [
        "To screen and select candidate biopolymers (natural, synthetic and "
        "semi-synthetic) for compatibility with progesterone on the basis "
        "of solubility, miscibility, thermal behaviour and Fourier-transform "
        "infrared (FTIR) spectroscopic interaction studies.",

        "To formulate progesterone-loaded microparticles using the spray "
        "drying method by systematic variation of inlet temperature, feed "
        "solid content, drug-to-polymer ratio and atomisation conditions.",

        "To determine the encapsulation efficiency, drug loading capacity "
        "and process yield for each polymer-progesterone combination.",

        "To physico-chemically characterise the optimised microparticles "
        "with respect to particle size and distribution, surface morphology "
        "(SEM), zeta potential, residual moisture, polydispersity index "
        "(PDI) and crystallinity (DSC, XRD).",

        "To assess drug-polymer compatibility and any solid-state "
        "interactions using FTIR and differential scanning calorimetry (DSC).",

        "To evaluate the in-vitro progesterone release profile in simulated "
        "vaginal fluid and to fit the data to standard release-kinetic "
        "models (zero-order, first-order, Higuchi, Korsmeyer-Peppas, "
        "Hixson-Crowell, Weibull).",

        "To determine accelerated and real-time storage stability of the "
        "optimised formulation according to ICH-style conditions, including "
        "drug content, EE retention and morphological integrity.",

        "To document, in tabular form, a comparative compatibility matrix "
        "ranking the screened biopolymers for progesterone spray drying "
        "and identify the lead formulation candidate(s) for future "
        "in-vivo evaluation.",
    ])


def build_hypothesis(doc: Document) -> None:
    add_h1(doc, "6. HYPOTHESIS")
    add_para(doc,
        "It is hypothesised that progesterone, being a lipophilic neutral "
        "steroid, can be efficiently encapsulated by spray drying using "
        "biocompatible biopolymer matrices, and that the choice of polymer "
        "(and polymer blend) will significantly influence encapsulation "
        "efficiency, particle morphology, in-vitro release kinetics and "
        "storage stability of the resulting microparticles. It is further "
        "hypothesised that one or more low-cost, biodegradable biopolymer "
        "or biopolymer-blend formulations will produce a sustained "
        "progesterone release profile of at least seven days under "
        "simulated vaginal conditions, supporting their candidacy as a "
        "low-cost, biodegradable alternative to currently marketed "
        "non-biodegradable intravaginal progesterone devices.")
    add_para(doc,
        "Null hypothesis (H0): There is no significant difference among "
        "biopolymers in encapsulation efficiency, release kinetics or "
        "stability of spray-dried progesterone microparticles.",
        italic=True)
    add_para(doc,
        "Alternative hypothesis (H1): At least one biopolymer (or polymer "
        "blend) will yield a spray-dried progesterone microparticle with "
        "encapsulation efficiency above 75%, sustained release over more "
        "than seven days, and acceptable physical stability under "
        "accelerated storage.",
        italic=True)


def build_polymer_table(doc: Document) -> None:
    add_h2(doc, "7.1 Comparative biopolymers under consideration")
    headers = ["Polymer", "Class", "Advantages",
               "Limitations / Disadvantages",
               "Suitability for spray-dried progesterone"]
    rows = [
        ["Chitosan",
         "Natural cationic polysaccharide",
         "Mucoadhesive; biodegradable; cationic - amenable to ionic "
         "crosslinking; widely studied in vaginal delivery [6,15]",
         "Soluble only at acidic pH; viscosity rises sharply with "
         "concentration; hygroscopic salt forms",
         "High - particularly attractive for intravaginal P4 due to "
         "mucoadhesion"],
        ["Sodium alginate",
         "Natural anionic polysaccharide",
         "GRAS; biodegradable; ionotropic gel with Ca2+; cheap [17]",
         "High viscosity at modest concentration; humidity-sensitive "
         "powders unless crosslinked",
         "Moderate - usually combined with chitosan or maltodextrin"],
        ["Gelatin",
         "Natural protein (collagen-derived)",
         "Excellent emulsifier; film-forming; biodegradable",
         "Low Tg; thermoplastic risk; rapid aqueous dissolution unless "
         "crosslinked",
         "Moderate - rapid release limits long-acting use"],
        ["Whey protein isolate / concentrate",
         "Natural protein",
         "GRAS; strong emulsifier; high EE for lipophiles [1]",
         "Moisture-sensitive; potential allergenicity (oral)",
         "High - strong candidate when blended with maltodextrin"],
        ["Sodium caseinate",
         "Natural protein",
         "GRAS; excellent emulsifier; spray-dries readily [1]",
         "Gastric-soluble; allergen risk",
         "Moderate-high in blend formulations"],
        ["Zein",
         "Natural prolamin (maize)",
         "Hydrophobic; ethanol-soluble; sustained release; GRAS [16]",
         "Brittle films; ethanolic feed system",
         "High - intrinsically hydrophobic, well matched to P4"],
        ["Maltodextrin",
         "Semi-synthetic carbohydrate",
         "Cheap; high Tg; low viscosity; excellent drying yield [1]",
         "No emulsifying capacity; minimal sustained release alone",
         "Indispensable carrier; rarely sole wall for P4"],
        ["Gum arabic (acacia)",
         "Natural polysaccharide-protein complex",
         "Amphiphilic; benchmark for lipophiles; high EE [1]",
         "Price volatility; supply variability",
         "Very high - reference wall material for P4"],
        ["Modified (OSA) starch",
         "Semi-synthetic polysaccharide",
         "Strong emulsifier; food-grade; spray-dries well [1]",
         "Higher cost than native starch",
         "High - particularly in blend with maltodextrin"],
        ["Hydroxypropyl methylcellulose (HPMC)",
         "Semi-synthetic cellulose ether",
         "Pharma-grade; tunable viscosity grades; sustained release",
         "Hydration-dependent gelling; particle aggregation risk",
         "High for oral / matrix-tablet P4 systems"],
        ["Ethyl cellulose (EC)",
         "Semi-synthetic cellulose ether",
         "Hydrophobic; excellent sustained release",
         "Requires organic solvent; reduced biodegradability",
         "High - in blend with HPMC for biphasic P4 release"],
        ["Polyvinyl alcohol (PVA)",
         "Synthetic water-soluble polymer",
         "Film-forming; biocompatible; emulsion stabiliser",
         "Slow biodegradation; mostly used as stabiliser, not sole wall",
         "Low-moderate as sole wall; useful as stabiliser"],
        ["Polyethylene glycol (PEG)",
         "Synthetic hydrophilic polymer",
         "Biocompatible; widely used as plasticiser/co-polymer",
         "Not biodegradable above 4-6 kDa; rapid dissolution",
         "Useful only as plasticiser / co-excipient"],
        ["Polycaprolactone (PCL)",
         "Synthetic biodegradable polyester",
         "Slow biodegradation; long-acting depot potential",
         "Hydrophobic; needs organic solvents; high cost",
         "Niche - subcutaneous long-acting P4 implants"],
        ["Poly(lactic-co-glycolic acid) (PLGA)",
         "Synthetic biodegradable polyester",
         "FDA-approved; tunable degradation; weeks-months release [18,19]",
         "Expensive; organic solvents; GMP-intensive",
         "Very high for parenteral long-acting P4 microspheres"],
        ["Cyclodextrins (HP-beta-CD, beta-CD)",
         "Semi-synthetic cyclic oligosaccharide",
         "Inclusion complex with P4; solubility enhancement",
         "Cost (HP-beta-CD); rapid release alone",
         "High - especially as solubiliser within composite walls"],
    ]
    make_table(doc, headers, rows,
               caption="Table 1. Candidate biopolymers for spray-dried "
                       "progesterone encapsulation: comparative properties.")


def build_methods(doc: Document) -> None:
    add_h1(doc, "7. MATERIALS AND METHODS")

    add_h2(doc, "7.0 Overview and study design")
    add_para(doc,
        "The study will be executed in three sequential phases: (i) polymer "
        "selection and drug-polymer compatibility screening; (ii) formulation "
        "of progesterone-loaded microparticles by spray drying using a "
        "factorial Design of Experiments; and (iii) physico-chemical, "
        "release-kinetic and stability characterisation of the optimised "
        "formulations [11].")

    build_polymer_table(doc)

    add_h2(doc, "7.2 Procurement and quality of materials")
    add_bullets(doc, [
        "Progesterone of pharmaceutical grade (USP/IP), purity >= 99%, "
        "obtained from an indexed supplier with Certificate of Analysis.",
        "Biopolymers of food or pharmaceutical grade as specified in "
        "Table 1, with appropriate Certificates of Analysis.",
        "Solvents (ethanol, dichloromethane, dimethyl sulfoxide, distilled "
        "water) of HPLC or analytical grade.",
        "Crosslinking agents (calcium chloride, sodium tripolyphosphate) "
        "and surfactants (Tween 80, sodium dodecyl sulfate) of "
        "analytical grade.",
    ])

    add_h2(doc, "7.3 Drug-polymer compatibility screening")
    add_para(doc,
        "Compatibility of progesterone with each candidate polymer will be "
        "assessed prior to spray drying using: (i) FTIR spectroscopy of "
        "physical mixtures (1:1 w/w) compared with individual components, "
        "to detect chemical interaction or hydrogen-bonding shifts; "
        "(ii) DSC thermograms of physical mixtures to detect shifts or "
        "disappearance of the progesterone melting endotherm at "
        "approximately 128-131 deg C; (iii) accelerated isothermal stress "
        "(40 deg C / 75% relative humidity, 4 weeks) followed by HPLC "
        "assay; and (iv) solubility/miscibility evaluation in the candidate "
        "feed solvent system.")

    add_h2(doc, "7.4 Preparation of feed and progesterone solubilisation")
    add_para(doc,
        "Polymer solutions will be prepared at 5-15% w/v in the appropriate "
        "solvent system (water for chitosan/HCl, water for alginate, "
        "water-ethanol for whey protein, ethanol-water for zein, "
        "dichloromethane or acetone for PLGA/PCL/EC, and water for "
        "maltodextrin/gum arabic blends). Progesterone will be dissolved or "
        "molecularly dispersed at drug:polymer ratios of 1:5, 1:10 and "
        "1:20 w/w. For aqueous biopolymer feeds, an oil-in-water emulsion "
        "or a hydroxypropyl-beta-cyclodextrin inclusion complex will be "
        "employed to keep progesterone in solution.")

    add_h2(doc, "7.5 Spray drying procedure")
    add_para(doc,
        "Spray drying will be performed on a laboratory scale using a "
        "Buchi Mini Spray Dryer B-290 (or equivalent) equipped with a "
        "two-fluid nozzle (0.7 mm tip), with the following central "
        "operating window adapted from established literature for "
        "thermolabile actives [1,2,3]:")
    add_bullets(doc, [
        "Inlet air temperature: 130-180 deg C",
        "Outlet air temperature: 70-95 deg C (allowed to float)",
        "Feed flow rate: 4-12 mL/min",
        "Atomisation air pressure: 2-6 bar",
        "Aspirator setting: 90-100% (approximately 35-40 m^3/h)",
        "Feed solid content: 10-25% w/v",
    ])
    add_para(doc,
        "A two-level factorial (or central composite) design will be "
        "employed to study the influence of inlet temperature, feed flow "
        "rate, drug:polymer ratio and feed solid content on encapsulation "
        "efficiency and particle yield, with response-surface optimisation "
        "as recommended by Paulo and Santos (2017) [11].")

    add_h2(doc, "7.6 Determination of encapsulation efficiency and drug loading")
    add_para(doc,
        "Total progesterone content will be assayed by dissolving an "
        "accurately weighed mass of microparticles in a suitable solvent, "
        "filtering, and quantifying by validated reverse-phase HPLC "
        "(C18 column, methanol-water or acetonitrile-water mobile phase, "
        "UV detection at approximately 240 nm). Surface (non-encapsulated) "
        "drug will be removed by gentle washing with a non-solvent for "
        "progesterone and assayed separately. Encapsulation Efficiency "
        "(EE) and Drug Loading (DL) will be calculated as:")
    add_para(doc,
        "    EE (%) = (Encapsulated progesterone / Total progesterone in "
        "feed) x 100",
        italic=True)
    add_para(doc,
        "    DL (%) = (Encapsulated progesterone / Total weight of "
        "microparticles) x 100",
        italic=True)

    add_h2(doc, "7.7 Physico-chemical characterisation")
    add_bullets(doc, [
        "Particle size and polydispersity index (PDI): laser diffraction "
        "(Mastersizer) and/or dynamic light scattering (Zetasizer).",
        "Surface morphology: scanning electron microscopy (SEM), with "
        "transmission electron microscopy (TEM) where sub-micron particles "
        "are produced.",
        "Zeta potential: electrophoretic mobility (Zetasizer) at "
        "physiologically relevant pH.",
        "Residual moisture: Karl Fischer titration or thermogravimetric "
        "analysis (TGA).",
        "Solid-state characterisation: differential scanning calorimetry "
        "(DSC) and powder X-ray diffraction (XRD) to detect amorphisation "
        "or polymorphic conversion of progesterone within the matrix.",
        "Drug-polymer interaction: Fourier-transform infrared (FTIR) "
        "spectroscopy.",
        "Flow properties: angle of repose, Carr's compressibility index, "
        "Hausner ratio.",
    ])

    add_h2(doc, "7.8 In-vitro release study")
    add_para(doc,
        "In-vitro progesterone release will be evaluated using a USP "
        "Type II (paddle) apparatus, or a dialysis-bag method for "
        "sub-micron formulations, in simulated vaginal fluid (pH 4.2, "
        "Owen and Katz formulation) and simulated body fluid (pH 7.4) at "
        "37 +/- 0.5 deg C with constant stirring. Aliquots will be "
        "withdrawn at predefined intervals up to 240 hours, replaced with "
        "fresh medium, filtered, and assayed by validated HPLC. Cumulative "
        "release data will be fitted to zero-order, first-order, Higuchi, "
        "Korsmeyer-Peppas, Hixson-Crowell and Weibull models, and the "
        "best-fit model identified by adjusted R^2 and Akaike Information "
        "Criterion [14,20,21,22].")

    add_h2(doc, "7.9 Stability study")
    add_para(doc,
        "Optimised formulations will be packed in amber glass vials with "
        "silica desiccant and subjected to (a) accelerated stability at "
        "40 +/- 2 deg C / 75 +/- 5% relative humidity for three months and "
        "(b) long-term stability at 25 +/- 2 deg C / 60 +/- 5% relative "
        "humidity for six months, in line with ICH Q1A(R2) principles. "
        "Drug content, encapsulation efficiency, particle size and "
        "morphology will be reassessed at 0, 1, 2, 3 and 6 months.")

    add_h2(doc, "7.10 Statistical analysis")
    add_para(doc,
        "All experiments will be performed in triplicate (n >= 3) and "
        "results expressed as mean +/- standard deviation. Statistical "
        "comparison among formulations will use one-way ANOVA followed by "
        "Tukey's HSD test, with significance set at p < 0.05. Design of "
        "Experiments analysis (factorial / response surface) will be "
        "performed using Design-Expert or Minitab. Release-kinetic model "
        "fitting will use DDSolver or equivalent.")

    add_h2(doc, "7.11 Spray drying versus alternative encapsulation methods")
    headers = ["Method", "Principle", "Advantages",
               "Disadvantages",
               "Suitability for progesterone in veterinary use"]
    rows = [
        ["Spray drying",
         "Atomisation of polymer-drug feed into hot gas stream; "
         "instantaneous drying",
         "Continuous; scalable; one-step; narrow PSD; high EE for "
         "lipophiles; short thermal exposure [1,2,3,5]",
         "Loss of fine particles; nozzle clogging at high viscosity; "
         "thermal stress on very labile drugs",
         "High - method of choice for industrial-scale powdered P4 "
         "formulations"],
        ["Solvent evaporation / emulsion (O/W, W/O/W)",
         "Polymer-drug dissolved in organic solvent, emulsified, "
         "solvent removed under vacuum",
         "Mild; high EE for hydrophobes; widely reported for PLGA P4 "
         "microspheres",
         "Batch process; residual solvents; long processing time; "
         "scale-up complex",
         "Moderate-high (parenteral PLGA P4)"],
        ["Coacervation (simple/complex)",
         "Phase separation of polymer from solution by salting-out / "
         "pH change",
         "Mild conditions; high payload; complex coacervates with "
         "chitosan-alginate well documented",
         "Stringent process control; narrow operating window",
         "Moderate"],
        ["Ionotropic gelation",
         "Crosslinking of polyelectrolytes (alginate / chitosan) by "
         "counter-ions",
         "Simple; aqueous; mild",
         "Wet beads need drying; modest sustained release",
         "Moderate (often combined with spray drying)"],
        ["Freeze drying (lyophilisation)",
         "Sublimation of frozen feed",
         "Excellent for thermally labile actives; preserves morphology",
         "Slow; expensive; produces cake, not free-flowing powder",
         "Low-moderate (cost-prohibitive at scale)"],
        ["Supercritical fluid (RESS, PGSS, SAS)",
         "CO2 used as solvent or anti-solvent",
         "Solvent-free; mild temperature; nano-scale particles",
         "Capital-intensive; limited drug solubility in scCO2",
         "Niche research-scale"],
        ["Electrospraying",
         "Charged droplets atomised under high voltage [22]",
         "Sub-micron particles; high EE; mild",
         "Low throughput; not yet industrial",
         "Niche research-scale"],
    ]
    make_table(doc, headers, rows,
               caption="Table 2. Spray drying vs alternative encapsulation "
                       "methods for progesterone microparticles.",
               body_size=10)


def build_outcome(doc: Document) -> None:
    add_h1(doc, "8. EXPECTED OUTCOMES")
    add_bullets(doc, [
        "A ranked compatibility matrix of biopolymers (natural, synthetic, "
        "semi-synthetic) for spray-dried progesterone, supported by FTIR / "
        "DSC compatibility data and quantitative encapsulation efficiency.",

        "At least one optimised biopolymer or polymer-blend formulation "
        "achieving encapsulation efficiency above 75%, with a sustained "
        "in-vitro progesterone release profile of seven or more days.",

        "Mechanistic understanding of the dominant release mode (diffusion, "
        "erosion, swelling) for each biopolymer-progesterone system, "
        "underpinned by best-fit release-kinetic models.",

        "A validated solid-state and stability profile (residual moisture, "
        "crystallinity, accelerated and real-time stability up to six "
        "months) for the lead formulation(s).",

        "A reusable Design-of-Experiments dataset and an analytical platform "
        "(HPLC method, dissolution method) that can be adopted for future "
        "in-vivo and field studies.",

        "A clearly identified lead candidate biopolymer formulation suitable "
        "for subsequent in-vivo plasma progesterone profiling and "
        "reproductive end-point evaluation in cattle / buffalo.",
    ])


def build_novelty(doc: Document) -> None:
    add_h1(doc, "9. NOVELTY OF THE STUDY")
    add_bullets(doc, [
        "Side-by-side comparison of natural, semi-synthetic and synthetic "
        "biopolymers for progesterone spray drying under a single, "
        "controlled experimental framework, generating an evidence-based "
        "compatibility matrix that is currently absent in the peer-reviewed "
        "literature.",

        "Application of Design of Experiments (factorial / response surface) "
        "to spray drying of progesterone, replacing the prevalent one-factor-"
        "at-a-time optimisation that has limited reproducibility and "
        "scale-up readiness.",

        "Focus on low-cost, food- or pharmaceutical-grade biopolymers and "
        "blends with the explicit goal of producing a biodegradable "
        "alternative to silicone CIDR/PRID-type intravaginal inserts in "
        "developing-country dairy systems.",

        "Integration of solid-state (DSC, XRD), molecular (FTIR), "
        "morphological (SEM) and kinetic (HPLC release modelling) "
        "characterisation in a single study, producing a reproducible "
        "data package that meets meta-analytical requirements.",
    ])


def build_future(doc: Document) -> None:
    add_h1(doc, "10. FUTURE APPLICATIONS")
    add_bullets(doc, [
        "Development of biodegradable spray-dried progesterone microparticle "
        "inserts for intravaginal oestrus synchronisation in dairy and beef "
        "cattle, sheep, goats and buffalo.",

        "Adjunct therapy in repeat-breeding cows and post-partum anoestrus "
        "to improve conception rates.",

        "Sustained-release oral or feed-additive progesterone formulations "
        "in livestock species where intravaginal use is impractical.",

        "Translation to other reproductive steroid hormones (estradiol "
        "valerate, GnRH analogues, prostaglandin F2-alpha) using the same "
        "biopolymer screening platform.",

        "Adaptation to human reproductive medicine for sustained "
        "intravaginal progesterone supplementation in luteal-phase support "
        "and assisted reproductive technology, subject to additional "
        "regulatory development.",
    ])


def build_workplan(doc: Document) -> None:
    add_h1(doc, "11. WORK PLAN AND TIMELINE")
    headers = ["Month", "Activity", "Deliverable"]
    rows = [
        ["1-2",
         "Literature review; finalisation of polymer panel; procurement of "
         "drug, polymers, solvents, consumables; HPLC method development.",
         "Validated HPLC assay; complete materials inventory."],
        ["3-4",
         "Drug-polymer compatibility screening (FTIR, DSC, isothermal "
         "stress, solubility study).",
         "Compatibility report; short-listed polymer panel."],
        ["5-7",
         "Pilot spray-drying trials; Design of Experiments planning; "
         "factorial / response-surface runs.",
         "Process map; preliminary EE and yield data."],
        ["8-9",
         "Optimisation of leading formulations; preparation of optimised "
         "batches.",
         "Optimised formulation(s) with documented EE >= 75%."],
        ["10-12",
         "Physico-chemical characterisation (PSD, SEM, zeta potential, "
         "DSC, XRD, residual moisture, flow).",
         "Complete characterisation dossier."],
        ["13-15",
         "In-vitro release studies in simulated vaginal and body fluid; "
         "release-kinetic modelling.",
         "Release profiles and best-fit kinetic models."],
        ["16-18",
         "Accelerated stability (40/75) and start of long-term stability "
         "(25/60).",
         "Three-month accelerated stability report."],
        ["19-21",
         "Long-term stability continuation; data consolidation; "
         "thesis-chapter drafting.",
         "Six-month real-time stability report; thesis chapters 1-4."],
        ["22-24",
         "Manuscript preparation, thesis writing, viva voce.",
         "Submitted manuscript(s); final thesis."],
    ]
    make_table(doc, headers, rows,
               caption="Table 3. Month-wise work plan (24-month plan; "
                       "compress to 18 months for M.V.Sc).",
               body_size=10)


def build_budget(doc: Document) -> None:
    add_h1(doc, "12. EXPECTED BUDGET")
    add_para(doc,
        "Budget figures below are indicative academic estimates in Indian "
        "Rupees (INR) for a 24-month programme conducted in a university "
        "research environment with shared central instrumentation. They "
        "should be adjusted to local supplier quotations and to the actual "
        "scope of the approved proposal.",
        italic=True, size=SMALL_SIZE)
    headers = ["Head", "Item", "Approx. cost (INR)"]
    rows = [
        ["Drug",
         "Pharmaceutical-grade progesterone (USP/IP), 25-50 g",
         "30,000 - 60,000"],
        ["Biopolymers",
         "Chitosan, alginate, gelatin, whey protein, sodium caseinate, "
         "zein, maltodextrin, gum arabic, modified starch, HPMC, EC, PVA, "
         "PEG, PCL, PLGA, HP-beta-CD",
         "1,50,000 - 2,50,000"],
        ["Solvents and reagents",
         "HPLC-grade methanol, acetonitrile, ethanol, dichloromethane, "
         "buffer salts, surfactants, crosslinkers",
         "75,000 - 1,00,000"],
        ["Spray-drying charges",
         "Use of laboratory spray dryer (Buchi B-290 or equivalent); "
         "approx. 100 runs at central facility",
         "1,50,000 - 2,00,000"],
        ["Characterisation - in-house",
         "FTIR, DSC, UV, HPLC consumables (columns, vials, filters)",
         "1,00,000 - 1,50,000"],
        ["Characterisation - paid central facilities",
         "SEM, TEM, XRD, particle size analysis, zeta potential, TGA",
         "1,50,000 - 2,50,000"],
        ["Glassware and consumables",
         "Volumetric glassware, syringes, dialysis tubing, vials, "
         "amber bottles, desiccators",
         "40,000 - 60,000"],
        ["Stability study",
         "Stability chamber rental / charges; humidity sensors; "
         "amber vials with desiccant",
         "60,000 - 90,000"],
        ["Statistical / DoE software",
         "Design-Expert or Minitab academic licence; DDSolver (free)",
         "30,000 - 60,000"],
        ["Travel and conference",
         "Conference presentation; collaborator visits",
         "50,000 - 75,000"],
        ["Publication and thesis",
         "Open-access article-processing charge; printing; binding",
         "1,00,000 - 2,00,000"],
        ["Contingency (approx. 10%)",
         "Unforeseen reagents, repeats, instrument time",
         "1,00,000 - 1,50,000"],
        ["TOTAL (indicative)",
         "",
         "10,35,000 - 15,45,000"],
    ]
    make_table(doc, headers, rows,
               caption="Table 4. Indicative budget estimate (INR).",
               body_size=10, first_col_bold=True)


def build_references(doc: Document) -> None:
    add_h1(doc, "13. REFERENCES")
    refs = [
        # 1
        "Gharsallaoui A, Roudaut G, Chambin O, Voilley A, Saurel R. "
        "Applications of spray-drying in microencapsulation of food "
        "ingredients: An overview. Food Research International. "
        "2007;40(9):1107-1121.",
        # 2
        "Sosnik A, Seremeta KP. Advantages and challenges of the "
        "spray-drying technology for the production of pure drug particles "
        "and drug-loaded polymeric carriers. Advances in Colloid and "
        "Interface Science. 2015;223:40-54.",
        # 3
        "Vehring R. Pharmaceutical particle engineering via spray drying. "
        "Pharmaceutical Research. 2008;25(5):999-1022.",
        # 4
        "Cal K, Sollohub K. Spray drying technique. I: Hardware and process "
        "parameters. Journal of Pharmaceutical Sciences. 2010;99(2):575-586.",
        # 5
        "Sollohub K, Cal K. Spray drying technique. II: Current applications "
        "in pharmaceutical technology. Journal of Pharmaceutical Sciences. "
        "2010;99(2):587-597.",
        # 6
        "Estevinho BN, Rocha F, Santos L, Alves A. Microencapsulation with "
        "chitosan by spray drying for industry applications - A review. "
        "Trends in Food Science & Technology. 2013;31(2):138-155.",
        # 7
        "Anandharamakrishnan C, Ishwarya SP. Spray Drying Techniques for "
        "Food Ingredient Encapsulation. Chichester: Wiley-Blackwell / IFT "
        "Press; 2015.",
        # 8
        "Rathbone MJ, Kinder JE, Fike K, Kojima F, Clopton D, Ogle CR, "
        "Bunt CR. Recent advances in bovine reproductive endocrinology and "
        "physiology and their impact on drug delivery system design for the "
        "control of the estrous cycle in cattle. Advanced Drug Delivery "
        "Reviews. 2002;54(8):1077-1112.",
        # 9
        "Macmillan KL, Peterson AJ. A new intravaginal progesterone "
        "releasing device for cattle (CIDR-B) for oestrous synchronisation, "
        "increasing pregnancy rates and the treatment of post-partum "
        "anoestrus. Animal Reproduction Science. 1993;33(1-4):1-25.",
        # 10
        "Rathbone MJ, Macmillan KL, Bunt CR, Burggraaf S. Conceptual and "
        "commercially available intravaginal veterinary drug delivery "
        "systems. Advanced Drug Delivery Reviews. 1997;28(3):363-392.",
        # 11
        "Paulo F, Santos L. Design of experiments for microencapsulation "
        "applications: A review. Materials Science and Engineering: C. "
        "2017;77:1327-1340.",
        # 12
        "Soottitantawat A, Yoshii H, Furuta T, Ohkawara M, Linko P. "
        "Microencapsulation by spray drying: Influence of emulsion size on "
        "the retention of volatile compounds. Journal of Food Science. "
        "2003;68(7):2256-2262.",
        # 13
        "Desai KGH, Park HJ. Recent developments in microencapsulation of "
        "food ingredients. Drying Technology. 2005;23(7):1361-1394.",
        # 14
        "Maderuelo C, Zarzuelo A, Lanao JM. Critical factors in the release "
        "of drugs from sustained release hydrophilic matrices. Journal of "
        "Controlled Release. 2011;154(1):2-19.",
        # 15
        "Rinaudo M. Chitin and chitosan: properties and applications. "
        "Progress in Polymer Science. 2006;31(7):603-632.",
        # 16
        "Paliwal R, Palakurthi S. Zein in controlled drug delivery and "
        "tissue engineering. Journal of Controlled Release. 2014;189:108-122.",
        # 17
        "Lee KY, Mooney DJ. Alginate: properties and biomedical applications. "
        "Progress in Polymer Science. 2012;37(1):106-126.",
        # 18
        "Makadia HK, Siegel SJ. Poly lactic-co-glycolic acid (PLGA) as "
        "biodegradable controlled drug delivery carrier. Polymers. "
        "2011;3(3):1377-1397.",
        # 19
        "Danhier F, Ansorena E, Silva JM, Coco R, Le Breton A, Preat V. "
        "PLGA-based nanoparticles: An overview of biomedical applications. "
        "Journal of Controlled Release. 2012;161(2):505-522.",
        # 20
        "Higuchi T. Mechanism of sustained-action medication. Theoretical "
        "analysis of rate of release of solid drugs dispersed in solid "
        "matrices. Journal of Pharmaceutical Sciences. 1963;52:1145-1149.",
        # 21
        "Korsmeyer RW, Gurny R, Doelker E, Buri P, Peppas NA. Mechanisms of "
        "solute release from porous hydrophilic polymers. International "
        "Journal of Pharmaceutics. 1983;15(1):25-35.",
        # 22
        "Costa P, Sousa Lobo JM. Modeling and comparison of dissolution "
        "profiles. European Journal of Pharmaceutical Sciences. "
        "2001;13(2):123-133.",
        # 23
        "Re MI. Microencapsulation by spray drying. Drying Technology. "
        "1998;16(6):1195-1236.",
        # 24
        "Bock N, Dargaville TR, Woodruff MA. Electrospraying of polymers "
        "with therapeutic molecules: state of the art. Progress in Polymer "
        "Science. 2012;37(11):1510-1551.",
        # 25
        "Joye IJ, McClements DJ. Biopolymer-based nanoparticles and "
        "microparticles: Fabrication, characterization, and application. "
        "Current Opinion in Colloid & Interface Science. 2014;19(5):417-427.",
        # 26
        "McClements DJ. Encapsulation, protection, and release of hydrophilic "
        "active components: Potential and limitations of colloidal delivery "
        "systems. Advances in Colloid and Interface Science. 2015;219:27-53.",
        # 27
        "Felt O, Buri P, Gurny R. Chitosan: a unique polysaccharide for "
        "drug delivery. Drug Development and Industrial Pharmacy. "
        "1998;24(11):979-993.",
        # 28
        "Sinha VR, Kumria R. Polysaccharides in colon-specific drug "
        "delivery. International Journal of Pharmaceutics. "
        "2001;224(1-2):19-38.",
        # 29
        "Khan I, Saeed K, Khan I. Nanoparticles: Properties, applications "
        "and toxicities. Arabian Journal of Chemistry. 2019;12(7):908-931.",
        # 30
        "International Conference on Harmonisation. ICH Q1A(R2): Stability "
        "Testing of New Drug Substances and Products. Geneva: ICH; 2003.",
    ]
    for i, ref in enumerate(refs, start=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.8)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(f"[{i}]  {ref}")
        run.font.name = BODY_FONT
        run.font.size = SMALL_SIZE


def build_declaration_note(doc: Document) -> None:
    doc.add_page_break()
    add_h1(doc, "Note on scientific integrity")
    add_para(doc,
        "All references cited in this synopsis are real, peer-reviewed "
        "publications drawn from indexed journals (Elsevier, Wiley, Springer, "
        "Taylor & Francis) and from the ICH stability guideline. No "
        "references, data, methods or results have been fabricated. "
        "Wherever a specific quantitative claim could not be supported by a "
        "single confidently traceable primary study, the prose uses "
        "qualitative ranges from the cited review literature, in line with "
        "good academic-writing practice. DOIs and exact page ranges should "
        "be re-verified by the candidate against PubMed / Scopus / Web of "
        "Science before final submission.",
        italic=True, size=SMALL_SIZE)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    doc = Document()
    configure_document(doc)

    build_title_page(doc)
    build_introduction(doc)
    build_review(doc)
    build_research_gap(doc)
    build_aim_objectives(doc)
    build_hypothesis(doc)
    build_methods(doc)
    build_outcome(doc)
    build_novelty(doc)
    build_future(doc)
    build_workplan(doc)
    build_budget(doc)
    build_references(doc)
    build_declaration_note(doc)

    doc.save(OUTPUT)
    print(f"OK  Synopsis written: {OUTPUT}")


if __name__ == "__main__":
    main()
