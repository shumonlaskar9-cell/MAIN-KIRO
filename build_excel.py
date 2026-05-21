#!/usr/bin/env python3
"""
Build verified peer-reviewed Excel workbook on:
Estrus induction/synchronization using intravaginal P4 (CIDR/PRID) in dairy animals (2015-2026)
Sheet 1: India Studies | Sheet 2: Global Studies
"""
import sys
sys.path.insert(0, '/root/.local/lib/python3.9/site-packages')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Column headers
HEADERS = [
    "Serial No.", "Article Title", "First Author", "Year", "Journal Source",
    "DOI/Article Link", "Country", "State/Region/Institution",
    "Dairy Animal Type/Breed", "Sample Size", "AIM/OBJECTIVES OF STUDY",
    "Research Gap", "P4 Sponge/Device Type", "Hormones Used", "Methodology",
    "Day-wise Protocol", "Timing Relative to AI/Mating",
    "Estrus Induction Rate (%)", "Conception Rate (%)", "Pregnancy Rate (%)",
    "Key Findings", "Major Results", "Statistical Significance",
    "Limitations", "Notes"
]


# ============ INDIA STUDIES (Verified from PubMed) ============
india_studies = [
    {
        "sno": 1,
        "title": "5d CIDR-Heatsynch improves the circulatory estradiol levels, estrus expression and conception rate in anestrus buffalo (Bubalus bubalis)",
        "author": "Singh H",
        "year": 2023,
        "journal": "Animal Biotechnology",
        "doi": "https://doi.org/10.1080/10495398.2022.2158337",
        "country": "India",
        "institution": "LUVAS, Hisar, Haryana / ICAR-IVRI, Bareilly",
        "breed": "Murrah Buffalo (anestrus)",
        "sample": "n=156 (78 per group)",
        "aim": "Investigate whether increasing estradiol during preovulatory period improves estrus expression and conception rate in 5d CIDR-based TAI protocol",
        "gap": "Limited data on estradiol supplementation in short-duration CIDR protocols for anestrus buffalo",
        "device": "CIDR (1.38 g P4, 5-day protocol)",
        "hormones": "GnRH, PGF2alpha, Estradiol Benzoate",
        "methodology": "RCT comparing 5d CIDR-Cosynch vs 5d CIDR-Heatsynch; subset (n=58) for follicle and luteal monitoring; AI at 72 and 84 h post-CIDR removal",
        "protocol": "Day -5: CIDR inserted + GnRH\nDay 0: PGF2alpha + CIDR removal\n24h post-CIDR removal: EB (Heatsynch) or 72h: GnRH (Cosynch)\n72 & 84h post-CIDR removal: AI",
        "timing": "AI at 72 h and 84 h after CIDR removal",
        "estrus_rate": "Higher in Heatsynch (p<0.05)",
        "conception_rate": "57.7% (Heatsynch) vs 43.6% (Cosynch)",
        "pregnancy_rate": "Not Reported",
        "findings": "5d CIDR-Heatsynch had greater E2, estrus induction, and trending higher conception rate than Cosynch",
        "results": "Estrus expression and P4 on d5 & d12 post-AI are indicators of pregnancy probability",
        "stats": "p<0.05 for E2 and estrus; p<0.08 trend for conception rate",
        "limitations": "Trending but non-significant conception difference; single location study",
        "notes": "PMID: 36576041"
    },

    {
        "sno": 2,
        "title": "Effect of lowering blood cortisol level along with progesterone priming on intensifying the secondary symptoms of oestrus in Murrah buffaloes during the hot summer months",
        "author": "Sengupta D",
        "year": 2022,
        "journal": "Reproduction in Domestic Animals",
        "doi": "https://doi.org/10.1111/rda.14218",
        "country": "India",
        "institution": "Bihar Veterinary College, Patna, Bihar",
        "breed": "Murrah Buffalo (pre-pubertal heifers)",
        "sample": "n=80 (Expt 2; 20 per group)",
        "aim": "Evaluate effect of cortisol reduction + progesterone priming on oestrus symptoms in heat-stressed pre-pubertal Murrah buffalo heifers",
        "gap": "Limited strategies to intensify oestrus expression in heat-stressed anestrus buffalo",
        "device": "CIDR (intravaginal progesterone, 5-day insertion)",
        "hormones": "Estradiol Benzoate, Glycerol, Vitamin E + Selenium",
        "methodology": "4 groups (P4, Gly-E, P4+Gly-E, Control); CIDR 5 days; EB on day 6; grading of oestrus symptoms (mucus, vulvar redness, ferning, cellularity)",
        "protocol": "Day 0: CIDR inserted (P4 group)\nDay 0-6: Gly-E treatment (Gly-E group)\nDay 5: CIDR removed\nDay 6: 2 mg EB (all groups)\nOestrus grading post-EB",
        "timing": "Oestrus graded post-EB injection on day 6",
        "estrus_rate": "Improved oestrus symptoms in P4+Gly-E group",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported",
        "findings": "Combination of cortisol reduction + P4 priming synergistically improved oestrus symptoms (mucus, ferning, cellularity)",
        "results": "P4+Gly-E group had significantly better mucus discharge, ferning, and lower Type B cells than P4 or Gly-E alone",
        "stats": "p<0.05 for mucus discharge, ferning, cellularity between groups",
        "limitations": "No conception/pregnancy data reported; focused on oestrus signs only",
        "notes": "PMID: 36317482"
    },

    {
        "sno": 3,
        "title": "Manipulation of reproductive performance of lactating buffaloes using melatonin and controlled internal drug release device treatment during out-of-breeding season under tropical conditions",
        "author": "Ramadan TA",
        "year": 2016,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2016.03.034",
        "country": "India",
        "institution": "CIRB, Hisar, Haryana / GADVASU, Ludhiana, Punjab",
        "breed": "Murrah Buffalo (lactating)",
        "sample": "n=12 (6 per group)",
        "aim": "Study effect of melatonin + CIDR on resumption of ovarian activity during out-of-breeding season (summer solstice)",
        "gap": "Need for strategies to overcome seasonal anestrus in buffalo during summer",
        "device": "CIDR (9-day insertion)",
        "hormones": "Melatonin implant (18 mg/50 kg BW), eCG (500 IU), GnRH",
        "methodology": "Melatonin implant 45 days + CIDR 9 days; eCG day before CIDR removal; GnRH day after CIDR removal; transrectal USG + blood sampling weekly",
        "protocol": "Day -45: Melatonin implant (treatment group)\nDay 0: CIDR inserted (both groups)\nDay 8: eCG 500 IU IM\nDay 9: CIDR removed\nDay 10: GnRH IM\nAI at detected estrus",
        "timing": "AI at detected estrus post-CIDR removal",
        "estrus_rate": "Not Reported",
        "conception_rate": "Higher in melatonin+CIDR vs CIDR alone",
        "pregnancy_rate": "Higher in treated group (CL maintained at d21 and d30 post-AI)",
        "findings": "Melatonin + CIDR improved CL diameter, P4 concentration, and maintained CL at d21 & d30 post-AI",
        "results": "Treatment group achieved higher conception than control; improved SOD activity with melatonin",
        "stats": "p<0.01 for CL diameter; p<0.05 for P4 concentration",
        "limitations": "Small sample size (n=12); single season; no exact CR percentages reported",
        "notes": "PMID: 27125696. Joint India-Egypt study conducted at CIRB Hisar India"
    },

    {
        "sno": 4,
        "title": "Comparative efficacy of E-17beta and GnRH administration on day 0 of a controlled internal drug release (CIDR) based protocol on synchrony of wave emergence, ovulation and conception rates in Murrah buffalos (Bubalus bubalis)",
        "author": "Bhat GR",
        "year": 2015,
        "journal": "Iranian Journal of Veterinary Research",
        "doi": "PMCID: PMC4789240",
        "country": "India",
        "institution": "GADVASU, Ludhiana, Punjab",
        "breed": "Murrah Buffalo",
        "sample": "n=60 (25+25+10)",
        "aim": "Compare efficacy of E-17beta+CIDR vs Ovsynch+CIDR for synchrony of wave emergence and ovulation in Murrah buffaloes",
        "gap": "Limited comparison between estradiol-based and GnRH-based CIDR protocols in buffalo",
        "device": "CIDR (1.38 g P4)",
        "hormones": "Estradiol-17beta (1.5 mg), GnRH (20 mcg), PGF2alpha (500 mcg)",
        "methodology": "Group I: CIDR+E-17beta; Group II: Ovsynch+CIDR 7d; Group III: Control; double AI at 12h and 24h after GnRH; USG for follicular dynamics",
        "protocol": "Group I:\nDay 0: CIDR + 1.5 mg E-17beta\nDay 9: CIDR removed + PGF2alpha\nDay 11: GnRH + AI (12h & 24h)\n\nGroup II:\nDay 0: GnRH + CIDR\nDay 7: CIDR removed + PGF2alpha\nDay 9: GnRH + AI at detected estrus",
        "timing": "Double AI at 12 h and 24 h after GnRH injection",
        "estrus_rate": "Not Reported",
        "conception_rate": "Higher (p<0.05) in Group I (E-17beta+CIDR) vs Group II",
        "pregnancy_rate": "Not Reported",
        "findings": "E-17beta+CIDR showed more synchronous wave emergence, larger pre-ovulatory follicles, higher FSCR",
        "results": "E-17beta+CIDR protocol superior for synchrony and conception in Murrah buffalo",
        "stats": "p<0.05 for follicle diameter and FSCR between Group I and II",
        "limitations": "Exact conception rate percentages not provided in abstract; single center",
        "notes": "PMID: 27175151"
    },

    {
        "sno": 5,
        "title": "Comparative efficacy of different estrus synchronization protocols on estrus induction response, fertility and plasma progesterone and biochemical profile in crossbred anestrus cows",
        "author": "Dhami AJ",
        "year": 2015,
        "journal": "Veterinary World",
        "doi": "https://doi.org/10.14202/vetworld.2015.1310-1316",
        "country": "India",
        "institution": "Anand Agricultural University, Anand, Gujarat",
        "breed": "Crossbred dairy cows (anestrus)",
        "sample": "n=50 (10 per group: CIDR, Ovsynch, Norgestomet, anestrus control, cyclic control)",
        "aim": "Evaluate estrus induction response and fertility following CIDR, Ovsynch, and Norgestomet protocols in anestrus crossbred cows",
        "gap": "Need to compare standard protocols with plasma biochemical monitoring in anestrus crossbred cows",
        "device": "CIDR (standard intravaginal device)",
        "hormones": "GnRH, PGF2alpha (Ovsynch); Norgestomet ear implant",
        "methodology": "CIDR/Ovsynch/Norgestomet with FTAI; blood sampling on day 0, 7, 9, and 21 post-AI for P4 and biochemistry",
        "protocol": "CIDR Protocol:\nDay 0: CIDR inserted\nDay 7: PGF2alpha + CIDR removed\nDay 9: FTAI\nDay 21 post-AI: blood for P4",
        "timing": "FTAI on Day 9 of CIDR protocol",
        "estrus_rate": "100% in all three treatment groups",
        "conception_rate": "CIDR: 60% (1st service), 80% (overall 3 cycles); Ovsynch: 50%, 80%; Norgestomet: 50%, 70%",
        "pregnancy_rate": "Not Reported separately",
        "findings": "All protocols induced 100% estrus; CIDR and Ovsynch achieved 80% overall CR vs 30% in untreated anestrus control",
        "results": "P4 on day 21 significantly higher in conceived vs non-conceived (CIDR: 4.36 vs 1.65 ng/ml); cholesterol and protein higher in cyclic vs anestrus",
        "stats": "p<0.01 for P4 on day 7 and 21; p<0.05 conceived vs non-conceived",
        "limitations": "Small sample (n=10/group); single institution; no randomization details",
        "notes": "PMID: 27047035; PMCID: PMC4774743"
    },

    {
        "sno": 6,
        "title": "Controlled breeding and reproductive management in water buffaloes (Bubalus bubalis) using Eazi Breed controlled internal drug release",
        "author": "Hiremath S",
        "year": 2015,
        "journal": "Journal of the South African Veterinary Association",
        "doi": "https://doi.org/10.4102/jsava.v86i1.1064",
        "country": "India",
        "institution": "ICAR-NDRI Southern Campus, Bengaluru, Karnataka",
        "breed": "Water Buffalo (anoestrus, 4-6 years)",
        "sample": "n=500",
        "aim": "Evaluate efficiency of Eazi Breed CIDR intravaginal device for estrus induction and fertility in true anoestrus buffaloes",
        "gap": "Large-scale field evaluation of CIDR for anoestrus buffalo management under Indian village conditions",
        "device": "Eazi Breed CIDR (intravaginal progesterone device)",
        "hormones": "Oestradiol Benzoate (Cidirol, 1 mg IM)",
        "methodology": "CIDR for 9 days in 500 anoestrus buffalo; Cidirol on day 10; repeat CIDR for non-responders; double AI 12h apart; pregnancy diagnosis at 45-60 days",
        "protocol": "Day 0: CIDR inserted\nDay 9: CIDR removed\nDay 10: 1 mg Oestradiol Benzoate IM\nEstrus detection\nDouble AI at 12h intervals\nNon-responders: repeat CIDR + AI",
        "timing": "Double AI starting 12 h after onset of oestrus signs, 12 h apart",
        "estrus_rate": "91.6% after first treatment; 100% after second treatment",
        "conception_rate": "Intense oestrus: 85.16%; Intermediate: 60.47%; Weak: 44.11%",
        "pregnancy_rate": "76% overall (380/500)",
        "findings": "CIDR+Cidirol highly effective in field conditions; pregnancy rate 76% in anoestrus buffalo; intensity of estrus correlated with conception",
        "results": "Intense oestrus 67.4%, intermediate 25.8%, weak 6.8%; 86 buffalo needed 3rd AI for prolonged oestrus",
        "stats": "Significant correlation between oestrus intensity and conception rate",
        "limitations": "Non-randomized field study; no untreated control group; variable inseminator",
        "notes": "PMID: 26244580; PMCID: PMC6137963. Study in Dharwad district, Karnataka"
    },

    {
        "sno": 7,
        "title": "p-cresol and oleic acid as reliable biomarkers of estrus: evidence from synchronized Murrah buffaloes",
        "author": "Muniasamy S",
        "year": 2017,
        "journal": "Iranian Journal of Veterinary Research",
        "doi": "PMCID: PMC5534256",
        "country": "India",
        "institution": "Bharathidasan University, Tiruchirappalli, Tamil Nadu",
        "breed": "Murrah Buffalo",
        "sample": "n=6 synchronized buffaloes",
        "aim": "Investigate presence of estrus-specific urinary pheromones (p-cresol, oleic acid) in CIDR-synchronized Murrah buffaloes",
        "gap": "Estrus biomarkers not validated in hormonally synchronized buffalo",
        "device": "CIDR (for estrus synchronization)",
        "hormones": "Progesterone (CIDR-based synchronization)",
        "methodology": "CIDR synchronization; mid-stream urine collected during estrus phases; GC-MS volatile compound analysis; bull behavioral response testing",
        "protocol": "CIDR-based synchronization protocol\nUrine collection at different phases\nGC-MS analysis of volatiles\nBull flehmen response test",
        "timing": "Urine sampled at synchronized estrus",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported",
        "findings": "p-cresol and oleic acid identified in synchronized estrus urine, same as natural estrus; bulls showed persistent flehmen to estrus urine",
        "results": "42 volatile compounds identified; p-cresol and oleic acid confirmed as estrus biomarkers in synchronized animals",
        "stats": "Behavioral significance confirmed via bull response",
        "limitations": "Very small sample (n=6); no fertility data; biomarker identification only",
        "notes": "PMID: 28775753. Included as it uses CIDR synchronization in Indian Murrah buffalo"
    },
]


# ============ GLOBAL STUDIES (Outside India - Verified from PubMed) ============
global_studies = [
    {
        "sno": 1,
        "title": "Fertility of Crossbred Dairy Cattle After Progesterone-Supplemented Co-Synch in Southwestern Ethiopia",
        "author": "Dodicho AA",
        "year": 2026,
        "journal": "Veterinary Medicine International",
        "doi": "https://doi.org/10.1155/vmi/1243812",
        "country": "Ethiopia",
        "institution": "Jimma University, College of Agriculture and Veterinary Medicine",
        "breed": "Crossbred dairy cattle (heifers and cows)",
        "sample": "n=120",
        "aim": "Evaluate effect of P4 supplementation during Co-Synch on estrus response and pregnancy per AI in crossbred dairy cattle",
        "gap": "Low reproductive efficiency of AI in Ethiopian crossbred dairy cattle",
        "device": "PRID (progesterone-releasing intravaginal device)",
        "hormones": "GnRH (100 mcg), PGF2alpha (25 mg)",
        "methodology": "All animals: GnRH+PRID on Day 0; PGF2alpha+PRID removal Day 7; GnRH+AI Day 9; estrus monitored AM/PM; pregnancy by USG at 60 days",
        "protocol": "Day 0: GnRH 100 mcg + PRID inserted\nDay 7: PGF2alpha 25 mg + PRID removed\nDay 9: GnRH 100 mcg + Timed AI",
        "timing": "Timed AI concurrent with 2nd GnRH on Day 9",
        "estrus_rate": "80% overall; primiparous 93.33%, heifers 83.33%",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "59.17% overall P/AI",
        "findings": "P4 supplementation improved pregnancy outcomes above national average; parity affected P/AI",
        "results": "Primiparous cows had highest P/AI; no effect of bull or BCS on pregnancy",
        "stats": "Logistic regression; parity significant predictor of P/AI",
        "limitations": "Single protocol (no comparison group without P4); limited to southwestern Ethiopia",
        "notes": "PMID: 41704301; PMCID: PMC12907567"
    },

    {
        "sno": 2,
        "title": "Evaluation of progesterone device reuse in shortened ovulation synchronization protocol in buffaloes raised in Amazon",
        "author": "Perdigao HH",
        "year": 2025,
        "journal": "Tropical Animal Health and Production",
        "doi": "https://doi.org/10.1007/s11250-025-04314-w",
        "country": "Brazil",
        "institution": "Universidade Federal do Amazonas, Manaus/Parintins",
        "breed": "Buffalo (Bubalus bubalis, 4-7 years)",
        "sample": "n=15 buffaloes, 45 experimental units (switch-back design)",
        "aim": "Assess viability of reusing P4 intravaginal devices in shortened ovulation synchronization protocol in buffaloes",
        "gap": "Cost of new P4 devices limits adoption in Amazon buffalo production",
        "device": "Intravaginal P4 device (2 g); New, used 7d, 14d, 21d, 28d",
        "hormones": "Estradiol Benzoate (2 mg), Sodium Cloprostenol (150 mcg), Estradiol Cypionate (1 mg)",
        "methodology": "Switch-back design, 3 periods; 5 treatments (new vs used devices); blood P4 on d0,2,4,7; ovarian USG alternate days",
        "protocol": "Day 0: P4 device inserted + 2 mg EB + 150 mcg cloprostenol IM\nDay 7: P4 device removed + 150 mcg cloprostenol + 1 mg estradiol cypionate IM",
        "timing": "Ovulation assessed post-device removal",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported (follicular dynamics study)",
        "findings": "Devices used up to 28 days maintained P4 >1 ng/mL; P4 >1.5 ng/mL linked to follicular cysts; low-content devices more efficient",
        "results": "GNew, G7, G14 tended to form persistent follicles/cysts; G21, G28 did not",
        "stats": "P4 concentrations compared across groups; tendency for cyst formation with higher P4",
        "limitations": "Small sample; no pregnancy/conception data; Amazonian conditions may not generalize",
        "notes": "PMID: 39932634"
    },

    {
        "sno": 3,
        "title": "Effect of dose and frequency of prostaglandin F2alpha treatments during a 7-day Ovsynch protocol with an intravaginal progesterone releasing device on luteal regression and pregnancy outcomes in lactating Holstein cows",
        "author": "Holper M",
        "year": 2023,
        "journal": "Journal of Dairy Science",
        "doi": "https://doi.org/10.3168/jds.2022-22245",
        "country": "Germany",
        "institution": "Freie Universitat Berlin, Clinic for Animal Reproduction",
        "breed": "Holstein (lactating dairy cows)",
        "sample": "n=1056 (356 Control, 353 DoubleDose, 347 2PGF)",
        "aim": "Evaluate effect of dose and frequency of PGF2alpha during 7-d Ovsynch+PRID on luteal regression and P/AI",
        "gap": "Incomplete luteal regression with single PGF in PRID-supplemented Ovsynch reduces fertility",
        "device": "PRID (progesterone-releasing intravaginal device, Day 0-8)",
        "hormones": "GnRH (100 mcg), Dinoprost (PGF2alpha 25 or 50 mg)",
        "methodology": "3-arm RCT; Control (single PGF), DoubleDose (50 mg PGF), 2PGF (25 mg d7+d8); PRID d0-d8; blood P4 on d0,7,9; pregnancy USG d38 and d80",
        "protocol": "Day 0: GnRH 100 mcg + PRID inserted\nDay 7: PGF2alpha (25 or 50 mg)\nDay 8: PGF2alpha (2PGF group) + PRID removed\nDay 9: GnRH 100 mcg\n16h later: Timed AI",
        "timing": "TAI approximately 16 h after 2nd GnRH (Day 9)",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "2PGF: 47.9% (no CL cows); Control 36.5%; DoubleDose: 32.7% (no CL); Overall ~40%",
        "findings": "2PGF improved P/AI in cows lacking CL; double dose did not improve over single; PRID vaginal discharge positively associated with P/AI",
        "results": "P4 at G2 lower in 2PGF and DoubleDose vs Control; 2PGF best for no-CL cows",
        "stats": "Treatment x CL interaction significant; P4 difference p<0.05",
        "limitations": "Single herd; vaginal discharge scoring subjective",
        "notes": "PMID: 36460496; Open access CC-BY"
    },

    {
        "sno": 4,
        "title": "The effect of GnRH administration/insemination time on follicular growth rate, ovulation intervals, and conception rate of Nili Ravi buffalo heifers in 7-day-CIDR Co-synch",
        "author": "Haider S",
        "year": 2021,
        "journal": "Tropical Animal Health and Production",
        "doi": "https://doi.org/10.1007/s11250-021-03003-8",
        "country": "Pakistan",
        "institution": "University of Veterinary and Animal Sciences, Lahore",
        "breed": "Nili Ravi Buffalo (heifers)",
        "sample": "n=40 (20 per group)",
        "aim": "Compare AI timing (72 vs 84 h post-CIDR removal) with GnRH in 7-day CIDR Co-synch on conception rate",
        "gap": "Optimal insemination timing after CIDR removal in buffalo heifers not established",
        "device": "CIDR (1.38 g progesterone, 7-day insertion)",
        "hormones": "GnRH (100 mcg), PGF2alpha (150 mcg)",
        "methodology": "RCT; CIDR 7 days; PGF at removal; GnRH+AI at 72h or 84h post-removal; USG for follicle and ovulation; pregnancy diagnosis day 40",
        "protocol": "Day 0: CIDR inserted\nDay 7: CIDR removed + PGF2alpha 150 mcg IM\n72h or 84h post-removal: GnRH 100 mcg + Timed AI\nDay 40: Pregnancy diagnosis (USG)",
        "timing": "TAI at 72 h or 84 h after CIDR removal concurrent with GnRH",
        "estrus_rate": "Not Reported",
        "conception_rate": "65% at 84 h vs 25% at 72 h",
        "pregnancy_rate": "Not Reported separately",
        "findings": "AI at 84h had significantly higher conception (65% vs 25%); greater follicular growth rate at 84h; shorter AI-to-ovulation interval",
        "results": "Follicular growth rate higher at 84h (0.102 vs 0.079 mm); GnRH/AI-ovulation interval shorter at 84h (15.1 vs 26.8 h)",
        "stats": "p=0.01 for follicular growth and ovulation interval; significant for conception rate",
        "limitations": "Small sample (n=20/group); single breed/location",
        "notes": "PMID: 34860311"
    },

    {
        "sno": 5,
        "title": "Effect of resynchronization with GnRH or progesterone (P4) intravaginal device (CIDR) on Day 23 after timed artificial insemination on cumulative pregnancy and embryonic losses in CIDR-GnRH synchronized Nili-Ravi buffaloes",
        "author": "Arshad U",
        "year": 2017,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2017.07.054",
        "country": "Pakistan",
        "institution": "University of Veterinary and Animal Sciences, Lahore",
        "breed": "Nili-Ravi Buffalo (lactating, mixed parity)",
        "sample": "n=181",
        "aim": "Determine effect of resynchronization on Day 23 with GnRH or CIDR on pregnancy rate, cumulative pregnancy, and embryonic losses",
        "gap": "Optimal resynchronization strategy after first TAI in buffalo not established",
        "device": "CIDR (for initial synchronization and resynchronization)",
        "hormones": "GnRH, PGF2alpha",
        "methodology": "CIDR 9.5 days for sync; GnRH 36h after removal; TAI 18h later; Day 23 resynch: CON vs P4 (CIDR) vs GnRH; serial USG days 30,45,60,90",
        "protocol": "Day -9.5: CIDR inserted\nDay 0: CIDR removed\n36h post-removal: GnRH\n18h later: TAI (Day 0)\nDay 23: Resynchronization (CON/CIDR/GnRH)\nDay 30: Pregnancy diagnosis",
        "timing": "1st TAI: 18 h after GnRH (54h post-CIDR removal)",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "GnRH resynch: consistently higher than CON (p<0.05) at d30,45,60,90; Cumulative: GnRH+OVS 81% vs CON+AIDE 59%",
        "findings": "GnRH on Day 23 maintained higher pregnancy and reduced embryonic loss; cumulative pregnancy 81% with GnRH+Ovsynch resynch",
        "results": "Embryonic loss: GnRH 18% vs CON 42% at Day 45 (p<0.05); Fetal losses similar across groups after Day 60",
        "stats": "p<0.05 for pregnancy rate and embryonic loss between GnRH and CON",
        "limitations": "Non-uniform initial synchronization timing; mixed parity",
        "notes": "PMID: 28779607"
    },

    {
        "sno": 6,
        "title": "Synchronization with controlled internal drug release (CIDR) and prostaglandin F2alpha (PGF2alpha) influences oxidant/antioxidant biomarkers and mineral profile in summer-stressed anoestrous buffalo (Bubalus bubalis)",
        "author": "Amin YA",
        "year": 2019,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2019.05.014",
        "country": "Egypt",
        "institution": "Faculty of Veterinary Medicine, Aswan University",
        "breed": "Egyptian Buffalo (anoestrus, >120 days postpartum)",
        "sample": "n=50 (25 treated, 25 control)",
        "aim": "Detect role of oxidative stress in summer anoestrus and effects of CIDR treatment on oxidant/antioxidant and mineral profile",
        "gap": "Mechanism linking heat stress, oxidative stress, and anoestrus in buffalo not well characterized",
        "device": "CIDR (controlled internal drug release)",
        "hormones": "PGF2alpha (on Day 6 of CIDR)",
        "methodology": "CIDR group vs untreated control; blood collected before, after treatment, and 45 days post-oestrus; TAC, MDA, NO, vitamin C, P, Cu, Zn measured",
        "protocol": "Day 0: CIDR inserted\nDay 6: PGF2alpha IM\nDay 7: CIDR removed\nEstrus detection\nAI at detected estrus",
        "timing": "AI at detected estrus post-CIDR removal",
        "estrus_rate": "80% in treated group",
        "conception_rate": "75% in treated group",
        "pregnancy_rate": "Not Reported separately",
        "findings": "80% estrus induction and 75% conception with CIDR; TAC higher in treated; MDA and NO decreased at oestrus; P, Cu, Zn increased in oestrus",
        "results": "Significant improvement in antioxidant status and mineral profile in buffalo that conceived",
        "stats": "p<0.05 for TAC, MDA, P, Cu, Zn between treated and control",
        "limitations": "No randomization details; hot climate specific; single season",
        "notes": "PMID: 31129479"
    },

    {
        "sno": 7,
        "title": "Equine chorionic gonadotropin (eCG) enhances reproductive responses in CIDR-EB treated lactating anovular Nili-Ravi buffalo during the breeding season",
        "author": "Khan AS",
        "year": 2018,
        "journal": "Animal Reproduction Science",
        "doi": "https://doi.org/10.1016/j.anireprosci.2018.06.012",
        "country": "Pakistan",
        "institution": "University of Veterinary and Animal Sciences, Lahore",
        "breed": "Nili-Ravi Buffalo (lactating, anovular)",
        "sample": "n=87 (44 +eCG, 43 -eCG)",
        "aim": "Evaluate effect of eCG addition on estrus response, ovulation, CL development, P4, and P/AI in CIDR-EB protocol for anovular buffalo",
        "gap": "Role of eCG in improving CIDR-EB protocol outcomes in anovular buffalo not established",
        "device": "CIDR (7-day insertion, random stage)",
        "hormones": "Estradiol Benzoate (2 mg d0, 1 mg d8), eCG, PGF2alpha",
        "methodology": "CIDR d0; EB 2mg d0 + 1mg d8; PGF2alpha + eCG d6; CIDR removed d7; FTAI at 48 and 60h; pregnancy diagnosis 35 days post-AI",
        "protocol": "Day 0: CIDR inserted + 2 mg EB\nDay 6: PGF2alpha + eCG (treatment group)\nDay 7: CIDR removed\nDay 8: 1 mg EB\n48 & 60h post-CIDR removal: FTAI\nDay 35 post-AI: Pregnancy diagnosis",
        "timing": "FTAI at 48 h and 60 h after CIDR removal",
        "estrus_rate": "Not Reported separately",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "+eCG: 56% vs -eCG: 31% (p<0.05)",
        "findings": "eCG enhanced ovulation rate (93% vs 74%), CL diameter, P4 on d18 and d21, and P/AI",
        "results": "eCG group had larger CL on d15,18,21; higher P4; nearly doubled P/AI",
        "stats": "p<0.05 for ovulation rate, CL diameter, P4, and P/AI",
        "limitations": "Breeding season only; anovular buffalo specific",
        "notes": "PMID: 30149872"
    },

    {
        "sno": 8,
        "title": "Pregnancy rate in water buffalo following fixed-time artificial insemination using new or used intravaginal devices with two progesterone concentrations",
        "author": "Gutierrez-Anez JC",
        "year": 2018,
        "journal": "Tropical Animal Health and Production",
        "doi": "https://doi.org/10.1007/s11250-017-1479-1",
        "country": "Venezuela",
        "institution": "Universidad del Zulia, Maracaibo",
        "breed": "Water Buffalo (dairy, Bubalus bubalis)",
        "sample": "n=247",
        "aim": "Evaluate pregnancy rate using new vs reused intravaginal devices (DIB 1.0g vs CIDR 1.38g P4) in buffalo during breeding and non-breeding season",
        "gap": "Effect of P4 concentration and device reuse on buffalo fertility not clearly defined",
        "device": "DIB (1.0 g P4) and CIDR (1.38 g P4); new, used 1x (9d), used 2x (18d)",
        "hormones": "GnRH (Buserelin 10.5 mcg), PGF2alpha (25 mg), eCG (500 IU)",
        "methodology": "2x3 factorial design; 4 replicates; IVD+GnRH d0; IVD removal+PGF+eCG d9; GnRH+AI d11 (48h post-removal)",
        "protocol": "Day 0: IVD (new or used) inserted + GnRH 10.5 mcg IM\nDay 9: IVD removed + PGF2alpha 25 mg + eCG 500 IU IM\nDay 11 (48h post-removal): GnRH 10.5 mcg + AI (8-12h later)",
        "timing": "AI 8-12 h after GnRH on Day 11 (48 h post-IVD removal)",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "DIB-New: 62.7% vs CIDR-New: 40% (p=0.019); Overall DIB: 62.7% vs CIDR: 45% (p=0.006)",
        "findings": "Lower P4 concentration (DIB 1.0g) resulted in higher PR than higher P4 (CIDR 1.38g); reuse did not negatively affect PR",
        "results": "P4 levels can be reduced without impairing buffalo fertility; device reuse is viable",
        "stats": "p=0.019 DIB-New vs CIDR-New; p=0.006 overall DIB vs CIDR",
        "limitations": "Breed-specific (water buffalo); tropical conditions; season interaction not fully explored",
        "notes": "PMID: 29274053"
    },

    {
        "sno": 9,
        "title": "Hormonal profile and follicular dynamics concurrent with CIDR and insulin modified Ovsync TAI programs and their impacts on the fertility response in buffaloes",
        "author": "Ramoun AA",
        "year": 2017,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2017.08.018",
        "country": "Egypt",
        "institution": "Kafrelsheikh University / Mahallet-Mousa Buffalo Research Station",
        "breed": "Egyptian Buffalo (cyclic)",
        "sample": "n=51 (13 Ovsync, 20 CIDR-sync, 18 Insulin-sync)",
        "aim": "Study hormonal profile and follicular dynamics with CIDR and insulin-modified Ovsync and their impacts on fertility in buffaloes",
        "gap": "Limited data on CIDR and insulin modification of Ovsync in Egyptian buffaloes",
        "device": "CIDR (1.38 g progesterone, Day 0-7)",
        "hormones": "GnRH (Buserelin 20 mcg), PGF2alpha (Cloprostenol 500 mcg), Insulin (0.25 IU/kg)",
        "methodology": "3 groups: Ovsync-alone, CIDR-sync (Ovsync+CIDR d0-7), Insulin-sync (Ovsync+insulin d7-9); TAI 16h after GnRH2; USG + blood for P4, E2, insulin, IGF-1",
        "protocol": "Day 0: GnRH1 (all groups) + CIDR (CIDR-sync)\nDay 7: PGF2alpha + CIDR removed (CIDR-sync)\nDay 7,8,9: Insulin SC (Insulin-sync)\nDay 9: GnRH2\n16h later: TAI\nDay 30: Pregnancy USG",
        "timing": "TAI 16 h after GnRH2 on Day 9",
        "estrus_rate": "Not Reported",
        "conception_rate": "Improved in CIDR-sync and Insulin-sync vs Ovsync-alone",
        "pregnancy_rate": "Not Reported as exact %",
        "findings": "CIDR-sync and Insulin-sync improved fertility through modulating hormonal profile and follicular dynamics",
        "results": "P4 higher in pregnant vs non-pregnant (Ovsync, Insulin-sync); E2 on d9 higher in pregnant in all groups; Insulin-sync had largest LF in pregnant",
        "stats": "p<0.01 for P4 and E2 between pregnant and non-pregnant",
        "limitations": "Small sample; low breeding season; exact conception rates not in abstract",
        "notes": "PMID: 28888896"
    },

    {
        "sno": 10,
        "title": "Effect of timing of insemination after CIDR removal with or without GnRH on pregnancy rates in Nili-Ravi buffalo",
        "author": "Haider MS",
        "year": 2015,
        "journal": "Animal Reproduction Science",
        "doi": "https://doi.org/10.1016/j.anireprosci.2015.09.010",
        "country": "Pakistan",
        "institution": "University of Veterinary and Animal Sciences, Lahore",
        "breed": "Nili-Ravi Buffalo",
        "sample": "n=201 (151 CIDR, 50 CIDR-GnRH)",
        "aim": "Determine optimum AI time in relation to CIDR removal with or without GnRH in buffalo",
        "gap": "Optimal AI timing post-CIDR for buffalo fertility not established",
        "device": "CIDR (7-day protocol)",
        "hormones": "PGF2alpha (Day 6), GnRH (36h post-CIDR removal in GnRH group)",
        "methodology": "CIDR d0; PGF d6; CIDR removed d7; CIDR-GnRH group: GnRH 36h post-removal; TAI at 48, 60, or 72h post-removal; USG for follicle and pregnancy",
        "protocol": "Day 0: CIDR inserted\nDay 6: PGF2alpha\nDay 7: CIDR removed\n(CIDR-GnRH group: 36h post-removal GnRH)\nTAI at 48h, 60h, or 72h post-CIDR removal",
        "timing": "TAI at 48, 60, or 72 h post-CIDR removal",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "CIDR-GnRH: 48h=50%, 60h=59%, 72h=18%; CIDR: 48h=10%, 60h=37%, 72h=40%",
        "findings": "Optimal AI: 48-60h for CIDR-GnRH; 60-72h for CIDR-only protocol; mean ovulation earlier in CIDR-GnRH (68.4h vs 76.1h)",
        "results": "AI timing critically affects pregnancy; GnRH advances ovulation and narrows AI window",
        "stats": "p<0.05 for pregnancy rate differences between AI timings within each protocol",
        "limitations": "Field conditions; mixed parity; seasonal variation not controlled",
        "notes": "PMID: 26518606"
    },

    {
        "sno": 11,
        "title": "Adding a second prostaglandin F2alpha treatment to but not reducing the duration of a PRID-Synch protocol increases fertility after resynchronization of ovulation in lactating Holstein cows",
        "author": "Santos VG",
        "year": 2016,
        "journal": "Journal of Dairy Science",
        "doi": "https://doi.org/10.3168/jds.2015-10557",
        "country": "USA / Portugal",
        "institution": "University of Wisconsin-Madison / CEVA Sante Animale",
        "breed": "Holstein (lactating dairy cows)",
        "sample": "n=821",
        "aim": "Evaluate effect of second PGF2alpha and protocol duration in PRID-Synch resynchronization on P4 and P/AI",
        "gap": "Incomplete luteolysis in PRID-supplemented resynchronization protocols reduces fertility",
        "device": "PRID (progesterone-releasing intravaginal device)",
        "hormones": "GnRH, PGF2alpha (single or double)",
        "methodology": "3-arm RCT at nonpregnancy diagnosis: 7D1PGF, 7D2PGF, 5D2PGF; all with PRID; blood P4; pregnancy USG d32 and d60",
        "protocol": "Day 0: GnRH + PRID inserted (7D groups) or Day 2: GnRH + PRID (5D)\nDay 7: 1st PGF2alpha + PRID removed\nDay 8: 2nd PGF2alpha (2PGF groups)\nDay 9.5: GnRH\n~16h later: TAI",
        "timing": "TAI approximately 16 h after 2nd GnRH",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "7D2PGF+5D2PGF: 42.6% vs 7D1PGF: 35.7% at d32 (p<0.05)",
        "findings": "Adding 2nd PGF reduced incomplete luteal regression and improved P/AI at d32; shortening protocol did not improve P/AI",
        "results": "Incomplete luteal regression: 1.9% (7D2PGF) vs 11% (7D1PGF); no difference in pregnancy loss",
        "stats": "p<0.05 for P/AI d32 (2PGF vs 1PGF); trend for luteal regression",
        "limitations": "Single study; breed-specific (Holstein); US/Portugal farm systems",
        "notes": "PMID: 26971149"
    },

    {
        "sno": 12,
        "title": "Ovarian response and conception rate following oestrus synchronization using three protocols in Egyptian buffalo heifers",
        "author": "Hussein HA",
        "year": 2016,
        "journal": "Tierarztliche Praxis Ausgabe G Grosstiere/Nutztiere",
        "doi": "https://doi.org/10.15653/TPG-160214",
        "country": "Egypt",
        "institution": "Assiut University, Faculty of Veterinary Medicine",
        "breed": "Egyptian Buffalo (heifers, cyclic)",
        "sample": "n=80 (20 per group: CIDR, Ovsynch, PGF, Control)",
        "aim": "Monitor ovarian response and conception rate following CIDR, Ovsynch, and double-PGF protocols in Egyptian buffalo heifers",
        "gap": "Limited comparison of synchronization protocols in Egyptian buffalo heifers",
        "device": "EAZI-BREED CIDR (intravaginal progesterone device)",
        "hormones": "GnRH, PGF2alpha",
        "methodology": "4-group comparison with TAI; ultrasound for follicles; blood P4; pregnancy diagnosis",
        "protocol": "CIDR Group:\nDay 0: CIDR inserted\nDay 7-9: CIDR removed + PGF2alpha\nTAI at detected oestrus\n\nOvsynch: GnRH-7d-PGF-56h-GnRH-TAI\nPGF: Double PGF 11 days apart",
        "timing": "TAI at detected oestrus (CIDR) or timed (Ovsynch)",
        "estrus_rate": "Not Reported separately",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "CIDR: 35%; Ovsynch: 40%; PGF: 35%; Control: 20%",
        "findings": "CIDR achieved 100% ovulation rate (vs 75% Ovsynch, 70% PGF); smaller ovulatory follicle in CIDR; all protocols superior to control",
        "results": "Ovulation rate highest with CIDR; pregnancy rates similar across treatment groups but all higher than control",
        "stats": "p<0.05 for ovulatory follicle diameter; non-significant P4 differences",
        "limitations": "Relatively small groups (n=20); single herd; P/AI relatively low across all",
        "notes": "PMID: 27595673"
    },

    {
        "sno": 13,
        "title": "Ovarian follicular changes and hemodynamics in Egyptian buffaloes under CIDR-PGF2alpha and Ovsynch-CIDR estrus synchronization treatments",
        "author": "Samir H",
        "year": 2019,
        "journal": "Journal of Reproduction and Development",
        "doi": "https://doi.org/10.1262/jrd.2019-035",
        "country": "Egypt / Japan",
        "institution": "Cairo University / Tokyo University of Agriculture and Technology",
        "breed": "Egyptian Buffalo (pluriparous)",
        "sample": "n=36 (18 per group)",
        "aim": "Investigate efficacy of CIDR-PGF2alpha vs Ovsynch-CIDR on follicular changes and hemodynamics in Egyptian buffaloes",
        "gap": "Doppler-based hemodynamic comparison of CIDR protocols in buffalo not previously done",
        "device": "CIDR (7-day insertion, both groups)",
        "hormones": "GnRH (Ovsynch-CIDR group), PGF2alpha",
        "methodology": "2 groups: CIDR-PGF (CIDR 7d + PGF d6) vs Ovsynch-CIDR (Ovsynch + CIDR 7d); blood sampling; Grayscale + Color/Power Doppler USG at removal, estrus, luteal phase",
        "protocol": "CIDR-PGF2alpha:\nDay 0: CIDR inserted\nDay 6: PGF2alpha\nDay 7: CIDR removed\n\nOvsynch-CIDR:\nDay 0: GnRH + CIDR inserted\nDay 7: PGF2alpha + CIDR removed\nDay 9: GnRH",
        "timing": "AI at detected estrus (both protocols)",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported (hemodynamic study)",
        "findings": "Ovsynch-CIDR superior: higher follicle population and MFD at estrus; larger CL volume; greater blood flow to DF and CL",
        "results": "TBF to DF 7x higher in Ovsynch-CIDR at estrus; P4 higher in luteal phase; E2 higher at estrus in Ovsynch-CIDR",
        "stats": "p<0.01 for TBF; p<0.05 for E2 and PDP",
        "limitations": "No pregnancy data; hemodynamic endpoints only; small sample",
        "notes": "PMID: 31406064; PMCID: PMC6815744; Open access"
    },

    {
        "sno": 14,
        "title": "Ovarian dynamics of buffalo (Bubalus bubalis) synchronized with different hormonal protocols",
        "author": "Peralta-Torres JA",
        "year": 2020,
        "journal": "Tropical Animal Health and Production",
        "doi": "https://doi.org/10.1007/s11250-020-02381-9",
        "country": "Mexico",
        "institution": "Universidad Juarez Autonoma de Tabasco / Universidad Autonoma de Yucatan",
        "breed": "Murrah Buffalo (pluriparous)",
        "sample": "n=29 (10 Ovsynch, 10 CIDR+EB, 9 CIDR+eCG)",
        "aim": "Evaluate effect of three hormonal protocols (Ovsynch, CIDR+EB, CIDR+eCG) on ovarian dynamics and P4 secretion in buffalo",
        "gap": "Comparative ovarian dynamic data for CIDR protocols in Mexican buffalo lacking",
        "device": "CIDR (used in CIDR+EB and CIDR+eCG groups)",
        "hormones": "GnRH (Gonadorelin 100 mcg), Cloprostenol (500 mcg), EB (2 mg/1 mg), eCG (400 IU)",
        "methodology": "3 groups; follicle count and measurement with USG at 0, 24, 54h; ovulation at 70, 80, 94h post-CIDR removal; blood P4 on d0,7,10,15,22; estrus detection 1h 3x daily",
        "protocol": "CIDR+EB:\nDay 0: CIDR + 2 mg EB\nDay 7: CIDR removed + 500 mcg cloprostenol\nDay 8: 1 mg EB\n\nCIDR+eCG:\nDay 0: CIDR + 2 mg EB\nDay 7: CIDR removed + cloprostenol + 400 IU eCG",
        "timing": "AI based on estrus detection (3x daily observation)",
        "estrus_rate": "CIDR+EB: 50% showed estrus at 69.6h",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported (ovarian dynamics study)",
        "findings": "All buffaloes in all groups ovulated; CIDR+eCG had shortest ovulation time (69h); P4 normal in Ovsynch and CIDR+eCG on d15 and d22 (>1 ng/ml)",
        "results": "No treatment differences for follicular population or max follicle diameter; P4 secretion normal in Ovsynch and CIDR+eCG but not CIDR+EB",
        "stats": "P>0.05 for follicle variables; p<0.05 for P4 on d15 and d22 (Ovsynch/CIDR+eCG > CIDR+EB)",
        "limitations": "Small sample; no pregnancy outcome; single farm in tropical Mexico",
        "notes": "PMID: 32949356"
    },

    {
        "sno": 15,
        "title": "LH peak and ovulation after two different estrus synchronization treatments in buffalo cows in the daylight-lengthening period",
        "author": "Barile VL",
        "year": 2015,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2015.03.019",
        "country": "Italy",
        "institution": "CRA-PCM, Monterotondo, Rome / University of Perugia",
        "breed": "Italian Mediterranean Buffalo (lactating)",
        "sample": "n=48 (12 PRID Feb-Mar, 12 Ovsynch Feb-Mar, 12 PRID May-Jun, 12 Ovsynch May-Jun)",
        "aim": "Determine timing of ovulation relative to LH peak after PRID vs Ovsynch and assess seasonal effects on these parameters",
        "gap": "Timing of LH-ovulation interval in PRID-synchronized buffalo during daylight-lengthening period not characterized",
        "device": "PRID (progesterone-releasing intravaginal device)",
        "hormones": "GnRH, PGF2alpha",
        "methodology": "2x2 design (PRID vs Ovsynch x Feb-Mar vs May-Jun); blood every 4h for LH from 24h post-PRID removal or 12h post-PGF; transrectal USG for ovulation",
        "protocol": "PRID Group:\nDay 0: PRID inserted\nDay 10: PRID removed\nLH/ovulation monitoring q4h from 24h post-removal\n\nOvsynch:\nDay 0: GnRH\nDay 7: PGF2alpha\nDay 9: GnRH\nLH monitoring from 12h post-PGF",
        "timing": "PRID: Ovulation ~76.8h post-removal; Ovsynch: ~35.7h post-last GnRH",
        "estrus_rate": "Feb-Mar: 95.8% LH peak, 83.3% ovulation; May-Jun: 75% LH peak, 54.1% ovulation (p<0.05)",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Not Reported (endocrine/ovulation study)",
        "findings": "Ovsynch showed better synchronization of LH peaks and ovulations; PRID had high variability in ovulation timing; seasonal decline May-Jun",
        "results": "LH-ovulation interval: PRID 30.1h, Ovsynch 32.8h (similar); ovulation timing post-removal highly variable in PRID (76.8 +/- 3.65h)",
        "stats": "p<0.05 for seasonal effect on LH peak and ovulation percentage",
        "limitations": "No pregnancy data; small groups (n=12); Italian conditions",
        "notes": "PMID: 25958084"
    },

    {
        "sno": 16,
        "title": "Effect of progesterone supplementation on fertility responses of lactating dairy cows with corpus luteum at the initiation of the Ovsynch protocol",
        "author": "Bisinotto RS",
        "year": 2015,
        "journal": "Theriogenology",
        "doi": "https://doi.org/10.1016/j.theriogenology.2014.09.021",
        "country": "USA",
        "institution": "University of Florida, Gainesville",
        "breed": "Holstein (lactating dairy cows with CL)",
        "sample": "n=1725 (863 control, 862 1CIDR)",
        "aim": "Determine effects of supplemental progesterone (CIDR) on fertility in cows with CL at initiation of Ovsynch",
        "gap": "Unclear whether P4 supplementation benefits cows that already have CL at Ovsynch start",
        "device": "CIDR (1.38 g P4, Day -10 to -3 of Ovsynch)",
        "hormones": "GnRH (Day -10 and Day -0.7), PGF2alpha (Day -3)",
        "methodology": "RCT; Ovsynch with or without CIDR (d-10 to d-3); estrus detection from d-9; subset blood P4; pregnancy USG d32 and d60",
        "protocol": "Day -10: GnRH + CIDR inserted (treatment) or no CIDR (control)\nDay -3: PGF2alpha + CIDR removed\nDay -0.7: GnRH\nDay 0: Timed AI\nDay 32 & 60: Pregnancy diagnosis",
        "timing": "Timed AI on Day 0 (after Ovsynch protocol)",
        "estrus_rate": "CIDR prevented premature estrus (0% vs 4.7% from d-9 to d-3)",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Control: 45.0% vs CIDR: 40.5% at d32 (p=0.06 trend); d60: 39.7% vs 36.6% (NS)",
        "findings": "CIDR did not benefit P/AI in cows with CL at Ovsynch; trend toward reduced P/AI; CIDR increased P4 by 1.3 ng/mL but did not improve fertility",
        "results": "CIDR prevented premature estrus; interaction: CIDR helped cows that lost CL by d-3 (38.1% vs 27.7%) but reduced P/AI in CL-maintained cows (40.3% vs 46.7%)",
        "stats": "p=0.06 tendency for reduced P/AI with CIDR; p<0.01 for P4 increase; p=0.09 interaction",
        "limitations": "Cows with CL only; CIDR may be contraindicated in CL-positive cows",
        "notes": "PMID: 25442385"
    },

    {
        "sno": 17,
        "title": "Reproductive outcomes of lactating dairy cows submitted to first timed artificial insemination protocols with different strategies to induce final ovulation",
        "author": "Consentini CEC",
        "year": 2025,
        "journal": "Journal of Dairy Science",
        "doi": "https://doi.org/10.3168/jds.2024-25188",
        "country": "Brazil",
        "institution": "University of Sao Paulo / University of Wisconsin-Madison",
        "breed": "Holstein (lactating dairy cows)",
        "sample": "n=909 (6 dairy herds)",
        "aim": "Evaluate strategies for induction of ovulation at end of TAI protocols using EC, GnRH, or EC+GnRH after novel presynchronization",
        "gap": "Optimal ovulation induction strategy after presynchronization with P4 implant not determined",
        "device": "P4 implant (2.0 g, intravaginal; used in presync and breeding protocol)",
        "hormones": "GnRH (Buserelin 16.8/8.4 mcg), Estradiol Cypionate (1 mg), PGF2alpha (Cloprostenol 0.530 mg)",
        "methodology": "Presync: P4 implant d-15 to d-8 + EC + PGF; Breeding protocol: GnRH + P4 implant d0; PGF d6+d7 + P4 removal; 3 groups (EC, GnRH, EC+GnRH); TAI d9",
        "protocol": "Presync:\nDay -15: P4 implant inserted\nDay -8: EC + PGF + P4 removed\n\nBreeding:\nDay 0: GnRH 16.8 mcg + P4 implant\nDay 6: PGF\nDay 7: PGF + P4 removed (+ EC in EC groups)\nDay 9 (56h post-1st PGF): GnRH 8.4 mcg (G and EC/G groups)\nDay 9: TAI (48h post-P4 removal)",
        "timing": "TAI on Day 9 (48 h after P4 withdrawal)",
        "estrus_rate": "81% of cows had CL at d0",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "Overall P/AI d31: 40.4% (no treatment difference); Pregnancy loss: G 19.8% vs EC 12.2% vs EC/G 10.1%",
        "findings": "No P/AI difference among treatments; EC reduced pregnancy loss vs GnRH alone; estrus expression positively affected P/AI in EC groups but not G group",
        "results": "Combined EC groups: 11.2% pregnancy loss vs 19.8% GnRH only (p=0.05); cows with CL at PGF had higher fertility (45.9% vs 17.7%)",
        "stats": "p=0.05 for pregnancy loss EC vs G; p<0.05 for CL presence effect",
        "limitations": "Multi-herd; no control without P4 in breeding protocol",
        "notes": "PMID: 39694244; Open access CC-BY"
    },

    {
        "sno": 18,
        "title": "Pregnancy outcomes are not improved by administering gonadotropin-releasing hormone at initiation of a 5-day CIDR-Cosynch resynchronization protocol for lactating dairy cows",
        "author": "Spencer JA",
        "year": 2018,
        "journal": "Journal of Dairy Science",
        "doi": "https://doi.org/10.3168/jds.2017-13491",
        "country": "USA",
        "institution": "University of Idaho, Moscow",
        "breed": "Holstein (lactating dairy cows)",
        "sample": "n=429 (226 GnRH, 203 no-GnRH)",
        "aim": "Determine effect of initial GnRH injection on P/AI in 5-day CIDR-Cosynch resynchronization protocol",
        "gap": "Unclear whether initial GnRH is necessary in short 5d CIDR-Cosynch resynch",
        "device": "CIDR (5-day insertion, Day 0-5)",
        "hormones": "GnRH (initial +/- Day 0), PGF2alpha (Day 5)",
        "methodology": "RCT at NPD (d37 post-1st AI); CIDR d0-5; +/- GnRH d0; PGF d5; estrus detection d6-7; EDAI or TAI d8; palpation d37 post-2nd AI",
        "protocol": "Day 0 (NPD): CIDR inserted +/- GnRH\nDay 5: CIDR removed + PGF2alpha\nDay 6-7: Estrus detection (AI if detected)\nDay 8: GnRH + TAI (if no estrus)\nDay 37 post-AI: Pregnancy palpation",
        "timing": "EDAI on Day 6-7 or TAI on Day 8 (with GnRH)",
        "estrus_rate": "Not Reported",
        "conception_rate": "Not Reported",
        "pregnancy_rate": "No-GnRH: 27% vs GnRH: 21% (NS); Primiparous: 31% vs Multiparous: 21%",
        "findings": "Eliminating initial GnRH did not reduce P/AI; primiparous had better fertility; high initial P4 tended toward better P/AI (26% vs 16%)",
        "results": "Initial GnRH is unnecessary in 5d CIDR-Cosynch resynch; can simplify protocol",
        "stats": "NS for treatment effect; tendency (p=trend) for high P4 vs low P4",
        "limitations": "Single farm; relatively low overall P/AI (21-27%); resynch only (2nd AI)",
        "notes": "PMID: 29885889"
    },
]



# ============ BUILD EXCEL WORKBOOK ============
def build_workbook():
    wb = Workbook()
    
    # Formatting styles
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    data_font = Font(name='Calibri', size=10)
    wrap_align = Alignment(wrap_text=True, vertical='top', horizontal='left')
    header_align = Alignment(wrap_text=True, vertical='center', horizontal='center')
    light_fill = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    def write_sheet(ws, studies, sheet_title):
        ws.title = sheet_title
        # Write headers
        for col_idx, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        
        # Write data
        keys = ["sno", "title", "author", "year", "journal", "doi", "country",
                "institution", "breed", "sample", "aim", "gap", "device",
                "hormones", "methodology", "protocol", "timing", "estrus_rate",
                "conception_rate", "pregnancy_rate", "findings", "results",
                "stats", "limitations", "notes"]
        
        for row_idx, study in enumerate(studies, 2):
            for col_idx, key in enumerate(keys, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=study.get(key, ""))
                cell.font = data_font
                cell.alignment = wrap_align
                cell.border = thin_border
            # Alternating row color
            if row_idx % 2 == 0:
                for col_idx in range(1, len(HEADERS) + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = light_fill
        
        # Auto-adjust column widths (cap at 40)
        for col_idx in range(1, len(HEADERS) + 1):
            max_len = len(HEADERS[col_idx - 1])
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if cell.value:
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            max_len = max(max_len, min(len(line), 40))
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2
        
        # Freeze top row
        ws.freeze_panes = 'A2'
        # Enable auto-filter
        ws.auto_filter.ref = ws.dimensions
    
    # Sheet 1: India
    ws1 = wb.active
    write_sheet(ws1, india_studies, "INDIA STUDIES")
    
    # Sheet 2: Global
    ws2 = wb.create_sheet()
    write_sheet(ws2, global_studies, "GLOBAL STUDIES")
    
    # Save
    output_path = "/projects/sandbox/MAIN-KIRO/Estrus_Synchronization_P4_Dairy_2015_2026.xlsx"
    wb.save(output_path)
    print(f"Workbook saved to: {output_path}")
    print(f"India studies: {len(india_studies)}")
    print(f"Global studies: {len(global_studies)}")

if __name__ == "__main__":
    build_workbook()
