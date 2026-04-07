"""
Generate SETUP.pdf — XLSForm Quality Reviewer
Includes a tool introduction section followed by the technical setup manual.
Font: Open Sans (Regular, SemiBold, Bold, Italic, BoldItalic)
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

OUTPUT    = "SETUP.pdf"
FONT_DIR  = "/Users/nicolewu/Library/Fonts/OpenSans-Static"
PAGE_W, PAGE_H = A4
MARGIN    = 22 * mm

# ── Register Open Sans ────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont("OpenSans",         f"{FONT_DIR}/OpenSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("OpenSans-Bold",    f"{FONT_DIR}/OpenSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("OpenSans-Semi",    f"{FONT_DIR}/OpenSans-SemiBold.ttf"))
pdfmetrics.registerFont(TTFont("OpenSans-Italic",  f"{FONT_DIR}/OpenSans-Italic.ttf"))
pdfmetrics.registerFont(TTFont("OpenSans-BoldIt",  f"{FONT_DIR}/OpenSans-BoldItalic.ttf"))
pdfmetrics.registerFontFamily(
    "OpenSans",
    normal    = "OpenSans",
    bold      = "OpenSans-Bold",
    italic    = "OpenSans-Italic",
    boldItalic= "OpenSans-BoldIt",
)

# ── Colour palette ────────────────────────────────────────────────────────────
C_PRIMARY    = colors.HexColor("#3a6ea8")
C_PRIMARY_LT = colors.HexColor("#d8eaf8")
C_PRIMARY_MD = colors.HexColor("#b0cfe8")
C_ACCENT     = colors.HexColor("#4a7c7e")   # teal accent
C_MAC        = colors.HexColor("#4a7c7e")
C_WIN        = colors.HexColor("#6a5aaa")
C_LIN        = colors.HexColor("#5a8f6a")
C_CODE_BG    = colors.HexColor("#f1f5f9")
C_CODE_FG    = colors.HexColor("#1e293b")
C_MUTED      = colors.HexColor("#64748b")
C_WARN_BG    = colors.HexColor("#fdf5ee")
C_WARN_FG    = colors.HexColor("#b87050")
C_TBL_HEAD   = colors.HexColor("#e8f2fc")
C_TBL_ALT    = colors.HexColor("#f8fafc")
C_WHITE      = colors.white
C_BLACK      = colors.HexColor("#1e293b")
C_BORDER     = colors.HexColor("#e2e8f0")
C_HIGHLIGHT  = colors.HexColor("#fdf8e8")


# ── Styles ────────────────────────────────────────────────────────────────────
def build_styles():
    def S(name, **kw):
        kw.setdefault("fontName", "OpenSans")
        kw.setdefault("textColor", C_BLACK)
        return ParagraphStyle(name, **kw)

    return {
        # Cover
        "cover_title":  S("cover_title",
                           fontName="OpenSans-Bold", fontSize=28,
                           textColor=C_WHITE, leading=34, spaceAfter=6),
        "cover_sub":    S("cover_sub",
                           fontName="OpenSans-Italic", fontSize=13,
                           textColor=colors.HexColor("#c8dff2"), leading=18, spaceAfter=0),
        "cover_tag":    S("cover_tag",
                           fontName="OpenSans", fontSize=9,
                           textColor=colors.HexColor("#a8c8e8"), leading=13),

        # Document headings
        "doc_title":    S("doc_title",
                           fontName="OpenSans-Bold", fontSize=20,
                           textColor=C_PRIMARY,
                           spaceBefore=0, spaceAfter=4, leading=26),
        "doc_sub":      S("doc_sub",
                           fontName="OpenSans-Italic", fontSize=11,
                           textColor=C_MUTED,
                           spaceBefore=0, spaceAfter=8, leading=16),
        "h1":           S("h1",
                           fontName="OpenSans-Bold", fontSize=15,
                           textColor=C_PRIMARY,
                           spaceBefore=16, spaceAfter=6, leading=20),
        "h2":           S("h2",
                           fontName="OpenSans-Semi", fontSize=12,
                           textColor=C_PRIMARY,
                           spaceBefore=14, spaceAfter=5, leading=17),
        "h3":           S("h3",
                           fontName="OpenSans-Semi", fontSize=10.5,
                           textColor=C_BLACK,
                           spaceBefore=10, spaceAfter=4, leading=15),
        "h_step":       S("h_step",
                           fontName="OpenSans-Semi", fontSize=10.5,
                           textColor=C_BLACK,
                           spaceBefore=0, spaceAfter=3, leading=15),
        "h_os":         S("h_os",
                           fontName="OpenSans-Bold", fontSize=13,
                           textColor=C_WHITE,
                           spaceBefore=0, spaceAfter=0, leading=18),

        # Body
        "body":         S("body",
                           fontName="OpenSans", fontSize=9.5,
                           textColor=C_BLACK,
                           spaceBefore=3, spaceAfter=5, leading=15),
        "body_j":       S("body_j",
                           fontName="OpenSans", fontSize=9.5,
                           textColor=C_BLACK, alignment=TA_JUSTIFY,
                           spaceBefore=3, spaceAfter=6, leading=15),
        "bullet":       S("bullet",
                           fontName="OpenSans", fontSize=9.5,
                           textColor=C_BLACK,
                           spaceBefore=2, spaceAfter=2, leading=14,
                           leftIndent=14, firstLineIndent=-8),
        "note":         S("note",
                           fontName="OpenSans-Italic", fontSize=9,
                           textColor=C_WARN_FG,
                           spaceBefore=3, spaceAfter=3, leading=13,
                           leftIndent=10),
        "caption":      S("caption",
                           fontName="OpenSans-Italic", fontSize=8.5,
                           textColor=C_MUTED,
                           spaceBefore=2, spaceAfter=4, leading=12),

        # Code
        "code_line":    S("code_line",
                           fontName="Courier", fontSize=8.5,
                           textColor=C_CODE_FG,
                           spaceBefore=0, spaceAfter=0, leading=12),

        # Feature card
        "feat_title":   S("feat_title",
                           fontName="OpenSans-Semi", fontSize=9.5,
                           textColor=C_PRIMARY,
                           spaceBefore=0, spaceAfter=2, leading=13),
        "feat_body":    S("feat_body",
                           fontName="OpenSans", fontSize=8.8,
                           textColor=C_BLACK,
                           spaceBefore=0, spaceAfter=0, leading=13),

        # FAQ
        "faq_q":        S("faq_q",
                           fontName="OpenSans-Semi", fontSize=9.5,
                           textColor=C_PRIMARY,
                           spaceBefore=9, spaceAfter=2, leading=14),
        "faq_a":        S("faq_a",
                           fontName="OpenSans", fontSize=9.5,
                           textColor=C_BLACK,
                           spaceBefore=0, spaceAfter=3, leading=14),

        # Table cells
        "tbl_hdr":      S("tbl_hdr",
                           fontName="OpenSans-Semi", fontSize=8.5,
                           textColor=C_PRIMARY,
                           spaceBefore=0, spaceAfter=0, leading=12),
        "tbl_cell":     S("tbl_cell",
                           fontName="OpenSans", fontSize=8.5,
                           textColor=C_BLACK,
                           spaceBefore=0, spaceAfter=0, leading=12),

        # Footer
        "footer":       S("footer",
                           fontName="OpenSans", fontSize=7.5,
                           textColor=C_MUTED, alignment=TA_CENTER,
                           spaceBefore=0, spaceAfter=0, leading=11),
        "safety":       S("safety",
                           fontName="OpenSans", fontSize=9.5,
                           textColor=C_PRIMARY,
                           spaceBefore=0, spaceAfter=0, leading=14),
    }


# ── Helpers ───────────────────────────────────────────────────────────────────
def hr(color=C_BORDER, thickness=0.5, sb=6, sa=10):
    return HRFlowable(width="100%", thickness=thickness,
                      color=color, spaceBefore=sb, spaceAfter=sa)

def sp(h=6):
    return Spacer(1, h)


def code_block(lines, styles):
    rows = [[Paragraph(l, styles["code_line"])] for l in lines]
    tbl  = Table(rows, colWidths=[PAGE_W - 2 * MARGIN - 10 * mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_CODE_BG),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("LINEABOVE",     (0, 0), (-1, 0),  0.5, C_BORDER),
        ("LINEBELOW",     (0, -1),(-1, -1), 0.5, C_BORDER),
    ]))
    return tbl


def os_banner(text, color, styles):
    p   = Paragraph(text, styles["h_os"])
    tbl = Table([[p]], colWidths=[PAGE_W - 2 * MARGIN])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), color),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
    ]))
    return tbl


def step_row(num, title, color, styles):
    badge = Paragraph(
        f'<font color="white"><b>Step {num}</b></font>',
        ParagraphStyle("sb", fontName="OpenSans-Bold", fontSize=8.5,
                       textColor=C_WHITE, alignment=TA_CENTER, leading=12)
    )
    heading  = Paragraph(title, styles["h_step"])
    badge_tbl = Table([[badge]], colWidths=[16 * mm])
    badge_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), color),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 3),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    row = Table([[badge_tbl, heading]],
                colWidths=[18 * mm, PAGE_W - 2 * MARGIN - 18 * mm])
    row.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return row


def shortcut_box(lines, styles):
    header    = Paragraph("Returning users — quick launch",
                           ParagraphStyle("shh", fontName="OpenSans-Semi", fontSize=9,
                                          textColor=C_PRIMARY, leading=13))
    code_rows = [[Paragraph(l, styles["code_line"])] for l in lines]
    inner     = Table(code_rows, colWidths=[PAGE_W - 2 * MARGIN - 28 * mm])
    inner.setStyle(TableStyle([
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
    ]))
    tbl = Table([[header], [inner]], colWidths=[PAGE_W - 2 * MARGIN])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_PRIMARY_LT),
        ("TOPPADDING",    (0, 0), (0, 0),  9),
        ("BOTTOMPADDING", (0, -1),(-1, -1), 9),
        ("LEFTPADDING",   (0, 0), (-1, -1), 13),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 13),
        ("TOPPADDING",    (0, 1), (-1, -1), 5),
    ]))
    return tbl


def data_table(headers, rows, col_widths, styles, alt=True):
    import re
    def fmt(txt, hdr=False):
        s = styles["tbl_hdr"] if hdr else styles["tbl_cell"]
        txt = re.sub(r"`([^`]+)`",
                     lambda m: f'<font name="Courier" size="7.8">{m.group(1)}</font>', txt)
        return Paragraph(txt, s)

    tdata = [[fmt(h, True) for h in headers]]
    for row in rows:
        tdata.append([fmt(c) for c in row])

    tbl = Table(tdata, colWidths=col_widths, repeatRows=1)
    ts  = [
        ("BACKGROUND",    (0, 0), (-1, 0),  C_TBL_HEAD),
        ("LINEBELOW",     (0, 0), (-1, 0),  1,   C_PRIMARY),
        ("LINEBELOW",     (0, 1), (-1, -1), 0.4, C_BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]
    if alt:
        for i in range(2, len(tdata), 2):
            ts.append(("BACKGROUND", (0, i), (-1, i), C_TBL_ALT))
    tbl.setStyle(TableStyle(ts))
    return tbl


def feature_card(icon, title, body, styles, color=C_PRIMARY_LT):
    title_p = Paragraph(f"{icon}  {title}", styles["feat_title"])
    body_p  = Paragraph(body, styles["feat_body"])
    tbl = Table([[title_p], [body_p]], colWidths=[PAGE_W / 2 - MARGIN - 4 * mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), color),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LEFTPADDING",   (0, 0), (-1, -1), 11),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 11),
        ("TOPPADDING",    (0, 1), (-1, -1), 3),
        ("LINEABOVE",     (0, 0), (-1, 0),  2, color),
    ]))
    return tbl


# ── Page header / footer ──────────────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(C_BORDER)
    canvas.setLineWidth(0.5)
    # top rule
    canvas.line(MARGIN, PAGE_H - 13 * mm, PAGE_W - MARGIN, PAGE_H - 13 * mm)
    # bottom rule + page number
    canvas.line(MARGIN, 12 * mm, PAGE_W - MARGIN, 12 * mm)
    canvas.setFont("OpenSans", 7.5)
    canvas.setFillColor(C_MUTED)
    canvas.drawCentredString(
        PAGE_W / 2, 8 * mm,
        f"XLSForm Quality Reviewer  ·  Setup & Introduction Guide  ·  Page {doc.page}"
    )
    canvas.restoreState()

def on_cover(canvas, doc):
    """Cover page: full bleed solid background, no header/footer rules."""
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#2a5298"))
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.restoreState()


# ── Build document ────────────────────────────────────────────────────────────
def build_pdf():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=18 * mm, bottomMargin=18 * mm,
        title="XLSForm Quality Reviewer — Setup & Introduction Guide",
        author="XLSForm Quality Reviewer",
        subject="Tool introduction and local installation instructions",
    )

    styles = build_styles()
    story  = []

    # ══════════════════════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════════════════════
    # Top half: white space (background drawn by on_cover canvas callback)
    story.append(Spacer(1, 52 * mm))

    cover_title = Table(
        [[Paragraph("XLSForm Quality Reviewer", styles["cover_title"])],
         [Paragraph("Setup &amp; Introduction Guide", styles["cover_sub"])],
         [sp(8)],
         [Paragraph(
             "Structural and simulation-based quality assurance for household survey instruments",
             styles["cover_tag"]
         )]],
        colWidths=[PAGE_W - 2 * MARGIN]
    )
    cover_title.setStyle(TableStyle([
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
    ]))
    story.append(cover_title)
    story.append(Spacer(1, 60 * mm))

    # Bottom half: meta info on white background
    meta_rows = [
        ["Standards", "World Bank DIME ietestform  ·  IPA ipacheckscto"],
        ["Mode",      "Local installation — data never leaves your machine"],
        ["Platforms", "macOS  ·  Windows  ·  Linux"],
        ["Requires",  "Python 3.10 or higher"],
    ]
    C_META_KEY = colors.HexColor("#8ab4d8")   # soft sky-blue — readable on dark navy
    C_META_VAL = colors.HexColor("#e8f2fc")   # near-white — full contrast on dark navy
    C_META_DIV = colors.HexColor("#3a5a80")   # muted navy divider line

    meta_tbl = Table(
        [[Paragraph(r[0], ParagraphStyle("mk", fontName="OpenSans-Semi", fontSize=8.5,
                                          textColor=C_META_KEY, leading=12)),
          Paragraph(r[1], ParagraphStyle("mv", fontName="OpenSans", fontSize=9,
                                          textColor=C_META_VAL, leading=13))]
         for r in meta_rows],
        colWidths=[32 * mm, PAGE_W - 2 * MARGIN - 32 * mm]
    )
    meta_tbl.setStyle(TableStyle([
        ("LINEABOVE",     (0, 0), (-1, 0),  0.5, C_META_DIV),
        ("LINEBELOW",     (0, -1),(-1, -1), 0.5, C_META_DIV),
        ("LINEBELOW",     (0, 0), (-1, -2), 0.4, C_META_DIV),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(meta_tbl)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1 — INTRODUCTION
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("Introduction", styles["h1"]))
    story.append(hr(C_PRIMARY, thickness=1.2, sb=2, sa=10))

    # — Overview
    story.append(Paragraph("What is XLSForm Quality Reviewer?", styles["h2"]))
    story.append(Paragraph(
        "XLSForm Quality Reviewer is a locally-run quality assurance tool designed to help "
        "researchers, data managers, and survey designers identify structural and logical errors "
        "in XLSForm-based household survey instruments before deployment to the field.",
        styles["body_j"]
    ))
    story.append(Paragraph(
        "The tool performs three complementary types of analysis: a <b>structural inspection</b> "
        "that examines the form's design without executing it; a <b>simulation-based analysis</b> "
        "that runs hundreds of synthetic respondents through the form to expose logic errors that "
        "structural inspection alone cannot detect; and a <b>design advisor</b> that evaluates the "
        "form against applied economics and social science survey methodology best practices, "
        "returning prioritised improvement suggestions grounded in the XLSForm specification and "
        "SurveyCTO documentation. Results are presented in an interactive web interface and "
        "can be exported as a formatted HTML report.",
        styles["body_j"]
    ))
    story.append(sp(4))

    # — Why it matters
    story.append(Paragraph("Why Quality Assurance Matters", styles["h2"]))
    story.append(Paragraph(
        "Errors in survey instruments — even minor ones — can have significant consequences for "
        "data quality and research validity. Common problems such as misconfigured skip logic, "
        "duplicate variable names, missing validation rules, or broken group structures may go "
        "undetected during manual review, only surfacing during data collection when corrections "
        "are costly or impossible.",
        styles["body_j"]
    ))
    story.append(Paragraph(
        "XLSForm Quality Reviewer systematically checks for these issues before the form reaches "
        "enumerators, reducing the risk of data loss, respondent burden, and the need for "
        "costly re-surveys.",
        styles["body_j"]
    ))
    story.append(sp(4))

    # — Key features grid
    story.append(Paragraph("Key Features", styles["h2"]))
    story.append(sp(4))

    CARD_W  = (PAGE_W - 2 * MARGIN - 6 * mm) / 2

    features = [
        ("🔬", "Structural Analysis",
         "Inspects the form design against World Bank DIME ietestform and IPA ipacheckscto "
         "specifications — without executing the form. Detects broken group structures, "
         "duplicate field names, missing labels, outdated syntax, and more."),
        ("🎲", "Respondent Path Simulation",
         "Generates realistic synthetic respondents with correlated demographic profiles and "
         "response style archetypes (neutral, agreeable, cautious, extreme). Runs them through "
         "the form to detect unreachable questions, broken skip chains, and logic conditions "
         "that never fire across diverse respondent profiles."),
        ("💡", "Design Improvement Suggestions",
         "Evaluates the form against applied economics and social science survey methodology "
         "best practices. Returns 24 prioritised checks across 13 categories: reference integrity, "
         "choice logic, validation bounds, skip logic, performance, multilingual completeness, "
         "GPS accuracy, form metadata, and more — grounded in the XLSForm spec and SurveyCTO docs."),
        ("📋", "Exportable Quality Report",
         "Generates a self-contained HTML report with a metrics dashboard, issue breakdown "
         "charts, and grouped issue tables with remediation guidance for each problem type. "
         "Suitable for sharing with colleagues or archiving as pre-registration documentation."),
        ("🔒", "Local & Offline Operation",
         "The application runs entirely on the user's machine. No data is transmitted to any "
         "external server at any point. Safe for sensitive survey instruments and restricted data."),
        ("📊", "Severity & Priority Classification",
         "Structural and simulation issues are classified as Critical / High / Medium / Low. "
         "Design suggestions carry High / Medium / Low priority ratings. Every finding includes "
         "a plain-language explanation and a concrete, copy-ready fix."),
    ]

    for i in range(0, len(features), 2):
        left  = feature_card(*features[i],   styles, C_PRIMARY_LT)
        right = feature_card(*features[i+1], styles, colors.HexColor("#f5f2fd"))
        row_tbl = Table([[left, sp(6), right]],
                        colWidths=[CARD_W, 6 * mm, CARD_W])
        row_tbl.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
            ("TOPPADDING",    (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(row_tbl)
        story.append(sp(6))

    story.append(sp(2))

    # — Who is it for
    story.append(Paragraph("Who Is This Tool For?", styles["h2"]))
    for item in [
        "<b>Survey Designers</b> — Validate XLSForm structure and logic before pilot testing or deployment.",
        "<b>Data Managers</b> — Ensure all field names conform to platform limits and naming conventions, "
        "and that metadata fields required for tracking are present.",
        "<b>Principal Investigators</b> — Review a concise quality report to confirm the instrument "
        "meets methodological standards prior to field work.",
        "<b>Survey Platform Teams</b> — Identify SurveyCTO or ODK compatibility issues that would cause "
        "form upload or rendering failures.",
    ]:
        story.append(Paragraph(f"&#8226;  {item}", styles["bullet"]))
    story.append(sp(4))

    # — Standards
    story.append(Paragraph("Quality Standards", styles["h2"]))
    story.append(Paragraph(
        "Structural checks are derived from two widely adopted quality assurance frameworks "
        "for household survey instruments. Design suggestions are grounded in the XLSForm "
        "specification and SurveyCTO form design documentation.",
        styles["body"]
    ))
    story.append(sp(4))
    story.append(data_table(
        ["Source", "Institution / Publisher", "Scope"],
        [
            ["ietestform.ado",
             "World Bank DIME Analytics",
             "Field name length, group structure, metadata, outdated syntax, "
             "constraint messages, choice list hygiene"],
            ["ipacheckscto.ado",
             "Innovations for Poverty Action (IPA)",
             "Metadata fields, disabled/read-only flags, required field logic, "
             "or_other syntax, integer constraints, duplicate codes"],
            ["XLSForm Specification",
             "XLSForm.org (ODK / KoBoToolbox)",
             "Field types, column definitions, skip logic syntax, choice list rules, "
             "appearance options, repeat groups, multilingual forms"],
            ["SurveyCTO Form Design Docs",
             "Dobility / SurveyCTO",
             "Performance limits, pulldata() best practices, GPS accuracy parameters, "
             "repeat group sizing, calculate field optimisation, form version management"],
        ],
        [42 * mm, 46 * mm, PAGE_W - 2 * MARGIN - 88 * mm],
        styles,
    ))
    story.append(sp(6))

    # — How to use it
    story.append(Paragraph("How to Use the Tool", styles["h2"]))
    steps_intro = [
        ("1", "Install the application",
         "Follow the platform-specific setup instructions in this guide (Section 2). "
         "Installation is required only once; subsequent launches use the provided "
         "one-click launcher script."),
        ("2", "Launch the application",
         "Start the app using the provided launcher script or the terminal command. "
         "The tool opens in your web browser at http://localhost:8501."),
        ("3", "Upload your XLSForm",
         "Click the upload area and select your .xlsx file. The tool will immediately "
         "parse the survey and choices sheets, detect languages, and display summary statistics."),
        ("4", "Review structural issues",
         "Structural checks run automatically on upload. Review issues grouped by "
         "severity — Critical and High issues should be resolved before proceeding to "
         "simulation or field testing."),
        ("5", "Run respondent simulations",
         "Configure the number of simulations (50–100 recommended) and click Run. "
         "The tool generates synthetic respondents with realistic demographic profiles and "
         "response styles, then reports unreachable questions and broken skip chains."),
        ("6", "Review design suggestions",
         "Scroll to the Design Improvement Suggestions section for prioritised best-practice "
         "recommendations grounded in the XLSForm specification and SurveyCTO documentation. "
         "Each suggestion includes a ready-to-apply fix example."),
        ("7", "Export the quality report",
         "Download the HTML report for documentation, sharing with colleagues, "
         "or archiving as part of the study's pre-registration materials."),
    ]
    for num, title, desc in steps_intro:
        row = Table(
            [[Paragraph(num, ParagraphStyle("sn", fontName="OpenSans-Bold", fontSize=11,
                                             textColor=C_WHITE, alignment=TA_CENTER, leading=14)),
              Paragraph(f"<b>{title}</b>", styles["h3"]),
            ]],
            colWidths=[9 * mm, PAGE_W - 2 * MARGIN - 9 * mm]
        )
        row.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (0, 0), C_PRIMARY),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING",   (0, 0), (0, 0),  0),
            ("RIGHTPADDING",  (0, 0), (0, 0),  0),
            ("LEFTPADDING",   (1, 0), (1, 0),  10),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(KeepTogether([
            row,
            Paragraph(desc, ParagraphStyle("idesc", fontName="OpenSans", fontSize=9.5,
                                            textColor=C_MUTED, leading=14,
                                            leftIndent=19 * mm, spaceAfter=6, spaceBefore=2)),
        ]))

    story.append(sp(4))

    # — Data safety
    safety_tbl = Table(
        [[Paragraph(
            "&#128274;  <b>Data Privacy Statement</b><br/>"
            "XLSForm Quality Reviewer processes all data locally on the user's machine. "
            "No XLSForm content, analysis results, or any other information is transmitted "
            "to any external server, cloud service, or third party at any point during use. "
            "The tool is safe for use with sensitive survey instruments.",
            ParagraphStyle("sp", fontName="OpenSans", fontSize=9.5,
                           textColor=C_PRIMARY, leading=15)
        )]],
        colWidths=[PAGE_W - 2 * MARGIN]
    )
    safety_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_PRIMARY_LT),
        ("LINEABOVE",     (0, 0), (-1, 0),  2, C_PRIMARY),
        ("TOPPADDING",    (0, 0), (-1, -1), 11),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 11),
        ("LEFTPADDING",   (0, 0), (-1, -1), 13),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 13),
    ]))
    story.append(safety_tbl)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2 — SETUP GUIDE
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("Local Installation Guide", styles["h1"]))
    story.append(hr(C_PRIMARY, thickness=1.2, sb=2, sa=10))

    story.append(Paragraph("What You Need", styles["h2"]))
    story.append(Paragraph(
        "Confirm you have received the following files and place them together in the same "
        "folder (for example, a folder named <i>XLSForm Reviewer</i> on your Desktop).",
        styles["body"]
    ))
    story.append(sp(4))
    story.append(data_table(
        ["File", "Purpose"],
        [
            ["`app.py`",             "The application"],
            ["`requirements.txt`",   "List of software dependencies"],
            ["`launch_mac.command`", "One-click launcher for macOS (optional)"],
            ["`launch_windows.bat`", "One-click launcher for Windows (optional)"],
        ],
        [52 * mm, PAGE_W - 2 * MARGIN - 52 * mm],
        styles,
    ))
    story.append(sp(6))
    story.append(hr())

    # ── macOS ─────────────────────────────────────────────────────────────────
    def render_os(banner_text, color, steps, shortcut_lines, stop_terminal):
        out = [sp(4), os_banner(banner_text, color, styles), sp(8)]
        for num, heading, code_lines, body_paras, note in steps:
            block = [step_row(num, heading, color, styles)]
            for p in body_paras:
                block.append(Paragraph(p, styles["body"]))
            if code_lines:
                block.append(sp(3))
                block.append(code_block(code_lines, styles))
            if note:
                block.append(sp(3))
                block.append(Paragraph(f"<i>{note}</i>", styles["note"]))
            block.append(sp(4))
            out.append(KeepTogether(block))
        out += [sp(4), shortcut_box(shortcut_lines, styles), sp(6),
                Paragraph(
                    f"To <b>stop the app</b>, press "
                    f'<font name="Courier" size="8.5">Ctrl + C</font> in {stop_terminal}.',
                    styles["body"]
                )]
        return out

    mac_steps = [
        (1, "Check Python is installed",
         ["python3 --version"],
         ["Open <b>Terminal</b> (&#8984; + Space, type <i>Terminal</i>, press Enter) "
          "and run the command above. You should see <b>Python 3.10.x</b> or <b>3.11.x</b>.",
          "<b>If Python is not found or below 3.10:</b> download the latest 3.11.x installer "
          "from <i>python.org/downloads</i> and follow the prompts."], None),
        (2, "Navigate to the app folder",
         ['cd ~/Desktop/"XLSForm Reviewer"'],
         ['Run <font name="Courier" size="8.5">ls</font> to confirm '
          '<font name="Courier" size="8.5">app.py</font> and '
          '<font name="Courier" size="8.5">requirements.txt</font> are listed.'], None),
        (3, "Create a virtual environment  (first time only)",
         ["python3 -m venv venv"],
         ["Creates an isolated environment so the app's dependencies do not affect "
          "other software on your machine."], None),
        (4, "Activate the virtual environment",
         ["source venv/bin/activate"],
         ['The terminal prompt will change to show '
          '<font name="Courier" size="8.5">(venv)</font>, confirming it is active.'], None),
        (5, "Install dependencies  (first time only)",
         ["pip install -r requirements.txt"],
         ["Downloads and installs all required packages. This may take 1–2 minutes."], None),
        (6, "Launch the application",
         ["streamlit run app.py"],
         ['A browser window will open at '
          '<font name="Courier" size="8.5">http://localhost:8501</font>. '
          "If it does not open, paste that address into any browser."], None),
    ]
    story.extend(render_os(
        "macOS Instructions", C_MAC, mac_steps,
        ['cd ~/Desktop/"XLSForm Reviewer"', "source venv/bin/activate", "streamlit run app.py"],
        "Terminal"
    ))
    story.append(hr())

    win_steps = [
        (1, "Check Python is installed",
         ["python --version"],
         ["Open <b>Command Prompt</b> (Win + R, type <i>cmd</i>, Enter). "
          "You should see <b>Python 3.10.x</b> or <b>3.11.x</b>.",
          "<b>If not found or below 3.10:</b> download the 3.11.x 64-bit installer from "
          "<i>python.org/downloads/windows</i>. On the first screen, tick "
          "<b>\"Add Python to PATH\"</b> before clicking Install Now."], None),
        (2, "Navigate to the app folder",
         [r'cd %USERPROFILE%\Desktop\XLSForm Reviewer'],
         ['Run <font name="Courier" size="8.5">dir</font> to confirm the files are listed.'], None),
        (3, "Create a virtual environment  (first time only)",
         ["python -m venv venv"], [], None),
        (4, "Activate the virtual environment",
         [r"venv\Scripts\activate"],
         ['The prompt will show <font name="Courier" size="8.5">(venv)</font>.'],
         "If activation is blocked by an execution policy error, open PowerShell as "
         "Administrator and run:  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned "
         "-Scope CurrentUser  Then return to Command Prompt."),
        (5, "Install dependencies  (first time only)",
         ["pip install -r requirements.txt"], [], None),
        (6, "Launch the application",
         ["streamlit run app.py"],
         ['A browser window will open at '
          '<font name="Courier" size="8.5">http://localhost:8501</font>.'], None),
    ]
    story.extend(render_os(
        "Windows Instructions", C_WIN, win_steps,
        [r'cd %USERPROFILE%\Desktop\XLSForm Reviewer',
         r"venv\Scripts\activate", "streamlit run app.py"],
        "Command Prompt"
    ))
    story.append(hr())

    # Linux — Step 1 has an extra apt block inserted after the body
    story.append(sp(4))
    story.append(os_banner("Linux Instructions  (Ubuntu / Debian)", C_LIN, styles))
    story.append(sp(8))

    lin_steps = [
        (1, "Check Python is installed",
         ["python3 --version"],
         ["Open a terminal and run the command above. "
          "<b>If Python 3.10 or higher is not available</b>, install it:"], None),
        (2, "Navigate to the app folder",
         ['cd ~/Desktop/"XLSForm Reviewer"'],
         ['Run <font name="Courier" size="8.5">ls</font> to confirm the files are present.'], None),
        (3, "Create a virtual environment  (first time only)",
         ["python3 -m venv venv"], [], None),
        (4, "Activate the virtual environment",
         ["source venv/bin/activate"],
         ['The prompt will show <font name="Courier" size="8.5">(venv)</font>.'], None),
        (5, "Install dependencies  (first time only)",
         ["pip install -r requirements.txt"], [], None),
        (6, "Launch the application",
         ["streamlit run app.py"],
         ['The app will open at <font name="Courier" size="8.5">http://localhost:8501</font>. '
          "Navigate there manually if the browser does not open automatically."], None),
    ]
    for num, heading, code_lines, body_paras, note in lin_steps:
        block = [step_row(num, heading, C_LIN, styles)]
        for p in body_paras:
            block.append(Paragraph(p, styles["body"]))
        if code_lines:
            block.append(sp(3))
            block.append(code_block(code_lines, styles))
        if num == 1:
            block.append(sp(3))
            block.append(code_block(
                ["sudo apt update",
                 "sudo apt install python3 python3-pip python3-venv -y"],
                styles
            ))
        if note:
            block.append(sp(3))
            block.append(Paragraph(f"<i>{note}</i>", styles["note"]))
        block.append(sp(4))
        story.append(KeepTogether(block))

    story.append(sp(4))
    story.append(shortcut_box(
        ['cd ~/Desktop/"XLSForm Reviewer"',
         "source venv/bin/activate", "streamlit run app.py"],
        styles
    ))
    story.append(sp(6))
    story.append(Paragraph(
        'To <b>stop the app</b>, press <font name="Courier" size="8.5">Ctrl + C</font> '
        'in the terminal.',
        styles["body"]
    ))
    story.append(hr())

    # ── Troubleshooting ────────────────────────────────────────────────────────
    story.append(Paragraph("Troubleshooting", styles["h2"]))
    story.append(data_table(
        ["Problem", "Likely Cause", "Solution"],
        [
            ["`python3: command not found`",
             "Python is not installed",
             "Follow Step 1 for your OS"],
            ["`pip: command not found`",
             "pip not on PATH",
             "Use `python3 -m pip install -r requirements.txt`"],
            ["`ModuleNotFoundError: No module named 'streamlit'`",
             "Virtual environment not active",
             "Run the activate command (Step 4) before launching"],
            ["Browser does not open",
             "Streamlit cannot detect browser",
             "Navigate manually to `http://localhost:8501`"],
            ["Port 8501 already in use",
             "Another instance is running",
             "Stop it with Ctrl + C, or append `--server.port 8502`"],
            ["Windows: `activate` execution policy error",
             "PowerShell security setting",
             "See the PowerShell note in Windows Step 4"],
            ["`openpyxl` error on upload",
             "Missing dependency",
             "Re-run Step 5 with the virtual environment active"],
            ["App shows 0 questions",
             "Incorrect sheet structure",
             "File must have a `survey` sheet with `type` and `name` columns"],
        ],
        [52 * mm, 44 * mm, PAGE_W - 2 * MARGIN - 96 * mm],
        styles,
    ))
    story.append(hr())

    # ── FAQ ────────────────────────────────────────────────────────────────────
    story.append(Paragraph("Frequently Asked Questions", styles["h2"]))
    faqs = [
        ("Is my data safe?",
         "Yes. The application runs entirely on your local machine. No XLSForm data, results, "
         "or personal information is transmitted to any external server at any point."),
        ("Can I run this without an internet connection?",
         "Yes. Once the dependencies are installed (Step 5), the app operates fully offline. "
         "An internet connection is only required for the initial installation."),
        ("Can multiple people use the same installation?",
         "The app runs on a single local port. For independent concurrent use, each user "
         "should install the tool separately in their own folder."),
        ("How do I update the app when a new version is provided?",
         "Replace app.py with the new file. If a new requirements.txt is provided, "
         "re-run pip install -r requirements.txt with the virtual environment active."),
        ("Where are downloaded HTML reports saved?",
         "Reports are saved to your browser's default download folder."),
        ("What forms does this tool support?",
         "Any XLSForm-compatible .xlsx file with a survey sheet and type/name columns. "
         "The tool is optimised for SurveyCTO and ODK Collect forms."),
    ]
    for q, a in faqs:
        story.append(KeepTogether([
            Paragraph(q, styles["faq_q"]),
            Paragraph(a, styles["faq_a"]),
        ]))

    story.append(sp(12))
    story.append(hr(C_PRIMARY, thickness=1, sb=4, sa=6))
    story.append(Paragraph(
        "XLSForm Quality Reviewer  ·  Standards: World Bank DIME ietestform &amp; IPA ipacheckscto",
        styles["footer"]
    ))

    doc.build(story, onFirstPage=on_cover, onLaterPages=on_page)
    print(f"PDF saved: {OUTPUT}")


if __name__ == "__main__":
    build_pdf()
