"""
utils/pdf_generator.py
───────────────────────
Premium PDF report generator — Editorial Dark design.
Generates a fully styled PDF from report data and sentiment results.

Public API:
    from utils.pdf_generator import generate_pdf
    pdf_path = generate_pdf(data_dict)

    data_dict keys:
        run_code   (str)   e.g. "IITM-AAIA-25-09 · #1_1"
        title      (str)   session title
        date       (str)   e.g. "19 April 2026"
        pl_name    (str)   facilitator name
        lob        (str)   e.g. "SEP-2025"
        rows       (list)  list of (sr_no, best_part, rating, improvement) tuples
        avg_score  (float) e.g. 4.89
        sentiments (list)  output of analyse_from_excel_rows()
        output_dir (str)   directory to save PDF (optional, defaults to /tmp)
"""

import os
import platform
from collections import Counter, defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus.flowables import Flowable


# ── Font registration ──────────────────────────────────────────
def _register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    if platform.system() == "Windows":
        font_paths = {
            'Sans':    'C:/Windows/Fonts/arial.ttf',
            'Sans-B':  'C:/Windows/Fonts/arialbd.ttf',
            'Serif-B': 'C:/Windows/Fonts/timesbd.ttf',
            'Mono':    'C:/Windows/Fonts/cour.ttf',
        }
    else:
        font_paths = {
            'Sans':    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            'Sans-B':  '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
            'Serif-B': '/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf',
            'Mono':    '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf',
        }

    registered = []
    for name, path in font_paths.items():
        try:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont(name, path))
                registered.append(name)
            else:
                print(f'⚠️ Font not found: {path} — using Helvetica fallback')
        except Exception as e:
            print(f'⚠️ Font registration failed for {name}: {e}')

    print(f'✅ Fonts registered: {registered}')

_register_fonts()

# ── Colours ────────────────────────────────────────────────────
GREEN      = colors.HexColor("#00B050")
GREEN_DARK = colors.HexColor("#007A38")
GREEN_L    = colors.HexColor("#E8F7EE")
GREEN_M    = colors.HexColor("#C2EACF")
AMBER      = colors.HexColor("#F59E0B")
AMBER_BG   = colors.HexColor("#FEF9C3")
AMBER_TX   = colors.HexColor("#92400E")
RED        = colors.HexColor("#EF4444")
RED_BG     = colors.HexColor("#FEE2E2")
RED_TX     = colors.HexColor("#B91C1C")
POS_BG     = colors.HexColor("#DCFCE7")
POS_TX     = colors.HexColor("#15803D")
INDIGO_BG  = colors.HexColor("#EEF2FF")
INDIGO     = colors.HexColor("#6366F1")
CARD       = colors.white
BORDER     = colors.HexColor("#E4E9E1")
TEXT       = colors.HexColor("#1A2418")
MUTED      = colors.HexColor("#6B7A66")
ROW_ALT    = colors.HexColor("#FAFCF9")
WHITE      = colors.white

# ── Editorial Dark constants ───────────────────────────────────
HEADER_BG  = colors.HexColor("#0F1C0F")
SCORE_BG   = colors.HexColor("#0A140A")
SCORE_BDR  = colors.HexColor("#1E3A1E")
TITLE_CLR  = colors.HexColor("#F5F3EE")
TBL_HDR    = colors.HexColor("#0F1C0F")
TBL_ACC    = colors.HexColor("#4A8C4A")
FOOTER_BG  = colors.HexColor("#F8FAF8")
CARD_BG    = colors.HexColor("#F8FAF8")

# ── Layout constants ───────────────────────────────────────────
PAGE_W, PAGE_H = A4
MARGIN = 1.6 * cm
USABLE = PAGE_W - 2 * MARGIN
COL2   = (USABLE - 10) / 2
KPI_W  = (USABLE - 3 * 6) / 4

_NO_FLAGS = {"no comments from the learner","na","none","n/a","nil","-",""}


# ── Style helper ───────────────────────────────────────────────
def S(name, **kw):
    base = dict(fontName='Sans', fontSize=9, textColor=TEXT, leading=13,
                spaceAfter=0, spaceBefore=0)
    base.update(kw)
    return ParagraphStyle(name, **base)


# ── Custom Flowables ───────────────────────────────────────────

class HeaderBlock(Flowable):
    H = 108

    def __init__(self, width, run_code, title, date, pl, lob, score):
        super().__init__()
        self.width = width
        self.run_code = run_code
        self.title = title
        self.date = date
        self.pl = pl
        self.lob = lob
        self.score = score

    def wrap(self, *_): return self.width, self.H

    def draw(self):
        c = self.canv
        w, h = self.width, self.H

        c.saveState()
        c.setFillColor(HEADER_BG)
        c.roundRect(0, 0, w, h, 12, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#1A3A1A"))
        c.circle(w - 20, h + 20, 62, fill=1, stroke=0)
        c.circle(w - 10, -14, 38, fill=1, stroke=0)
        c.restoreState()

        text_right = w - 130
        badge = self.run_code
        badge_w = min(pdfmetrics.stringWidth(badge, 'Mono', 7.5) + 18, text_right - 14)
        c.saveState()
        c.setFillColor(colors.HexColor("#00000025"))
        c.roundRect(14, h - 26, badge_w, 16, 8, fill=1, stroke=0)
        c.setFont('Mono', 7.5)
        c.setFillColor(GREEN)
        c.drawString(23, h - 19.5, badge)
        c.restoreState()

        c.saveState()
        c.setFont('Serif-B', 13)
        c.setFillColor(TITLE_CLR)
        max_w = text_right - 28
        words = self.title.split()
        line, lines = "", []
        for word in words:
            test = (line + " " + word).strip()
            if pdfmetrics.stringWidth(test, 'Serif-B', 13) <= max_w:
                line = test
            else:
                if line: lines.append(line)
                line = word
        if line: lines.append(line)
        y0 = h - 48
        for i, ln in enumerate(lines[:2]):
            c.drawString(14, y0 - i * 17, ln)
        c.restoreState()

        c.saveState()
        c.setFont('Sans', 8)
        c.setFillColor(GREEN)
        c.drawString(14, 14, f"{self.date}    {self.pl}")
        c.restoreState()

        bx, by, bw, bh = w - 122, 11, 110, 86
        c.saveState()
        c.setFillColor(SCORE_BG)
        c.setStrokeColor(SCORE_BDR)
        c.setLineWidth(1.2)
        c.roundRect(bx, by, bw, bh, 10, fill=1, stroke=1)
        c.setFont('Serif-B', 30)
        c.setFillColor(GREEN)
        c.drawCentredString(bx + bw / 2, by + 46, str(self.score))
        c.setFont('Sans', 7.5)
        c.setFillColor(colors.HexColor("#CCEAD8"))
        c.drawCentredString(bx + bw / 2, by + 32, "out of 5")
        c.setFont('Sans-B', 7)
        c.setFillColor(colors.HexColor("#CCEAD8"))
        c.drawCentredString(bx + bw / 2, by + 16, "AVG RATING")
        c.restoreState()


class SentimentBar(Flowable):
    def __init__(self, label, pct, bar_color, width):
        super().__init__()
        self.label = label
        self.pct = pct
        self.bar_color = bar_color
        self.width = width
        self.height = 17

    def wrap(self, *_): return self.width, self.height

    def draw(self):
        c = self.canv
        lbl_w, pct_w = 82, 30
        bar_w = self.width - lbl_w - pct_w - 8
        bar_h = 6
        bar_y = (self.height - bar_h) / 2
        c.setFont('Sans', 8); c.setFillColor(TEXT)
        c.drawString(0, bar_y + 0.5, self.label)
        c.setFillColor(BORDER)
        c.roundRect(lbl_w, bar_y, bar_w, bar_h, 3, fill=1, stroke=0)
        if self.pct > 0:
            c.setFillColor(self.bar_color)
            c.roundRect(lbl_w, bar_y, max(bar_w * self.pct, 0), bar_h, 3, fill=1, stroke=0)
        c.setFont('Mono', 8); c.setFillColor(MUTED)
        c.drawRightString(self.width, bar_y + 1, f"{int(self.pct * 100)}%")


class SentimentPill(Flowable):
    _cfg = {
        'Positive': (POS_BG, POS_TX, "Positive"),
        'Neutral':  (AMBER_BG, AMBER_TX, "Neutral"),
        'Negative': (RED_BG, RED_TX, "Negative"),
    }

    def __init__(self, sentiment):
        super().__init__()
        self.bg, self.fg, self.label = self._cfg.get(
            sentiment, (BORDER, MUTED, sentiment))
        self.width = 52; self.height = 15

    def wrap(self, *_): return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg)
        c.roundRect(0, 0, self.width, self.height, 7, fill=1, stroke=0)
        c.setFillColor(self.fg)
        c.circle(9, 7.5, 3, fill=1, stroke=0)
        c.setFont('Sans-B', 7); c.drawString(15, 4.5, self.label)


class RatingBadge(Flowable):
    _clr = {5: (POS_BG, POS_TX), 4: (AMBER_BG, AMBER_TX),
            3: (RED_BG, RED_TX),  2: (RED_BG, RED_TX), 1: (RED_BG, RED_TX)}

    def __init__(self, rating):
        super().__init__()
        self.rating = rating
        self.bg, self.fg = self._clr.get(rating, (BORDER, MUTED))
        self.width = self.height = 22

    def wrap(self, *_): return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.bg)
        c.roundRect(0, 0, 22, 22, 5, fill=1, stroke=0)
        c.setFont('Sans-B', 10); c.setFillColor(self.fg)
        c.drawCentredString(11, 6, str(self.rating))


class SectionHeading(Flowable):
    def __init__(self, text, width):
        super().__init__()
        self.text = text; self.width = width; self.height = 15

    def wrap(self, *_): return self.width, self.height

    def draw(self):
        c = self.canv
        lbl = self.text.upper()
        c.setFont('Sans-B', 7); c.setFillColor(MUTED)
        tw = pdfmetrics.stringWidth(lbl, 'Sans-B', 7)
        c.drawString(0, 4, lbl)
        c.setStrokeColor(BORDER); c.setLineWidth(0.5)
        c.line(tw + 8, 7, self.width, 7)


class ColorDot(Flowable):
    def __init__(self, color):
        super().__init__()
        self.color = color
        self.width = self.height = 10

    def wrap(self, *_): return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.circle(5, 5, 3, fill=1, stroke=0)


# ── Layout helpers ─────────────────────────────────────────────

def sp(n): return Spacer(1, n)


def wrap_in_card(content_table, width, pad_lr=14, pad_tb=12):
    outer = Table([[content_table]], colWidths=[width])
    outer.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), CARD),
        ('BOX',           (0, 0), (-1, -1), 0.5, BORDER),
        ('LEFTPADDING',   (0, 0), (-1, -1), pad_lr),
        ('RIGHTPADDING',  (0, 0), (-1, -1), pad_lr),
        ('TOPPADDING',    (0, 0), (-1, -1), pad_tb),
        ('BOTTOMPADDING', (0, 0), (-1, -1), pad_tb),
        ('ROUNDEDCORNERS',(0, 0), (-1, -1), [10]),
    ]))
    return outer


def two_col(left, right):
    t = Table([[left, right]], colWidths=[COL2, COL2])
    t.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('COLPADDING',    (0, 0), (-1, -1), 5),
    ]))
    return t


def dot_row(text_para, dot_color, inner_w):
    dot = Table([[ColorDot(dot_color)]], colWidths=[12])
    dot.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('TOPPADDING',    (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    r = Table([[dot, text_para]], colWidths=[14, inner_w - 14])
    r.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    return r


# ── Main public function ───────────────────────────────────────

def generate_pdf(data_dict: dict) -> str:
    """
    Generate premium Editorial Dark PDF from report data.

    Parameters
    ----------
    data_dict : dict with keys:
        run_code, title, date, pl_name, lob,
        rows (list of tuples: sr_no, best_part, rating, improvement),
        avg_score (float),
        sentiments (list of dicts from analyse_from_excel_rows),
        output_dir (str, optional)

    Returns
    -------
    str — absolute path to the saved PDF file
    """
    run_code   = data_dict["run_code"]
    title      = data_dict["title"]
    date       = data_dict["date"]
    pl_name    = data_dict["pl_name"]
    lob        = data_dict["lob"]
    raw_rows   = data_dict["rows"]
    avg_score  = data_dict["avg_score"]
    sentiments = data_dict.get("sentiments", [])
    output_dir = data_dict.get("output_dir", "/tmp")

    os.makedirs(output_dir, exist_ok=True)

    # Build rows with sentiment attached
    sent_map = {s["row"]: s["sentiment"] for s in sentiments}
    rows = [
        (sr, bp, rt, imp, sent_map.get(sr, "Positive"))
        for sr, bp, rt, imp in raw_rows
    ]

    total   = len(rows)
    perfect = sum(1 for r in rows if r[2] == 5)
    engage  = sum(1 for r in rows if r[3].lower().strip() not in _NO_FLAGS)

    sent_ct = Counter(r[4] for r in rows)

    theme_ct = defaultdict(int)
    for _, bp, *_ in rows:
        for t in bp.split(";"): theme_ct[t.strip()] += 1
    top3 = sorted(theme_ct.items(), key=lambda x: -x[1])[:3]

    meaningful = [r[3] for r in rows if r[3].lower().strip() not in _NO_FLAGS]
    no_comment_count = sum(
        1 for r in rows if "no comments from the learner" in r[3].lower())

    filename = data_dict.get("filename", f"Feedback_Report_{run_code}")
    output_path = os.path.join(output_dir, f"{filename}.pdf")

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
        title=f"Feedback Report — {run_code}",
        author="EMERITUS Feedback Automation Tool",
    )
    story = []

    # ── 1. Header ─────────────────────────────────────────────
    story.append(HeaderBlock(USABLE, run_code, title, date, pl_name, lob, avg_score))
    story.append(sp(14))

    # ── 2. KPI row ─────────────────────────────────────────────
    st_kpi_lbl = S('kl', fontName='Sans-B', fontSize=7, textColor=MUTED, leading=10)
    st_kpi_sub = S('ks', fontName='Sans',   fontSize=8, textColor=MUTED, leading=11)

    top_theme_name = top3[0][0] if top3 else "N/A"
    top_theme_words = top_theme_name.split()
    if len(top_theme_name) > 18:
        mid = len(top_theme_words) // 2
        top_theme_name = " ".join(top_theme_words[:mid]) + "\n" + " ".join(top_theme_words[mid:])

    kpi_cells = [
        (CARD_BG, GREEN,  "RESPONSES",      str(total),          "Learners participated"),
        (CARD_BG, GREEN,  "PERFECT SCORES", str(perfect),        "Rated 5 / 5"),
        (CARD_BG, AMBER,  "ENGAGEMENT",     f"{engage}/{total}", "Gave real feedback"),
        (CARD_BG, INDIGO, "TOP THEME",      top_theme_name,      "Most cited strength"),
    ]

    kpi_row_data = []
    for bg, accent, lbl, val, sub in kpi_cells:
        val_fs = 13 if lbl == "TOP THEME" else 17
        val_para = Paragraph(
            val.replace("\n", "<br/>"),
            S('kv', fontName='Sans-B', fontSize=val_fs, textColor=TEXT,
              leading=val_fs + 3, alignment=TA_LEFT)
        )
        inner = Table(
            [[Paragraph(lbl, st_kpi_lbl)], [val_para], [Paragraph(sub, st_kpi_sub)]],
            colWidths=[KPI_W - 20],
        )
        inner.setStyle(TableStyle([
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        cell = Table([[""], [inner]], colWidths=[KPI_W], rowHeights=[4, None])
        cell.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, 0), accent),
            ('BACKGROUND',    (0, 1), (-1, -1), bg),
            ('BOX',           (0, 0), (-1, -1), 0.5, BORDER),
            ('ROUNDEDCORNERS',(0, 0), (-1, -1), [8]),
            ('LEFTPADDING',   (0, 1), (-1, -1), 12),
            ('RIGHTPADDING',  (0, 1), (-1, -1), 8),
            ('TOPPADDING',    (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('LEFTPADDING',   (0, 0), (-1, 0), 0),
            ('RIGHTPADDING',  (0, 0), (-1, 0), 0),
            ('TOPPADDING',    (0, 0), (-1, 0), 0),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
        ]))
        kpi_row_data.append(cell)

    kpi_table = Table([kpi_row_data], colWidths=[KPI_W] * 4)
    kpi_table.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('COLPADDING',    (0, 0), (-1, -1), 6),
    ]))
    story.append(kpi_table)
    story.append(sp(12))

    # ── 3. Sentiment + Insights ────────────────────────────────
    inner_w = COL2 - 28

    no_comment_pct = no_comment_count / total if total else 0
    positive_imp   = sum(1 for r in rows
                         if r[3].lower().strip() not in _NO_FLAGS
                         and sent_ct.get(r[4], 0) >= 0) / total if total else 0
    na_only_pct    = sum(1 for r in rows
                         if r[3].lower().strip() in {"na","n/a","none","-",""}) / total if total else 0

    sent_items = [
        [SectionHeading("Sentiment Analysis", inner_w)], [sp(4)],
        [SentimentBar("Positive", sent_ct.get('Positive', 0) / total, GREEN, inner_w)], [sp(5)],
        [SentimentBar("Neutral",  sent_ct.get('Neutral',  0) / total, AMBER, inner_w)], [sp(5)],
        [SentimentBar("Negative", sent_ct.get('Negative', 0) / total, RED,   inner_w)], [sp(10)],
        [SectionHeading("Improvement Column Breakdown", inner_w)], [sp(4)],
        [SentimentBar("No Comment", no_comment_count / total, RED,   inner_w)], [sp(5)],
        [SentimentBar("Substantive", len(meaningful) / total, GREEN, inner_w)], [sp(5)],
        [SentimentBar("N/A only",   na_only_pct,              AMBER, inner_w)],
    ]
    sent_inner = Table(sent_items, colWidths=[inner_w])
    sent_inner.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    sent_card = wrap_in_card(sent_inner, COL2)

    st_ih = S('ih', fontName='Sans-B', fontSize=8, textColor=TEXT,  leading=12)
    st_ib = S('ib', fontName='Sans',   fontSize=8, textColor=MUTED, leading=12)

    def insight_row(icon_bg, emoji, hd, bd):
        icon_t = Table(
            [[Paragraph(emoji, S('em', fontSize=11, alignment=TA_CENTER))]],
            colWidths=[26], rowHeights=[26],
        )
        icon_t.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), icon_bg),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('ROUNDEDCORNERS',(0, 0), (-1, -1), [6]),
        ]))
        txt_t = Table(
            [[Paragraph(hd, st_ih)], [Paragraph(bd, st_ib)]],
            colWidths=[inner_w - 32],
        )
        txt_t.setStyle(TableStyle([
            ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        row_t = Table([[icon_t, txt_t]], colWidths=[30, inner_w - 30])
        row_t.setStyle(TableStyle([
            ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ]))
        return row_t

    top_theme_full = top3[0][0] if top3 else "N/A"
    top_theme_count = top3[0][1] if top3 else 0
    neg_count  = sent_ct.get('Negative', 0)
    neg_phrase = "Zero negative sentiment detected." if neg_count == 0 \
        else f"{neg_count} response(s) flagged as negative — review recommended."

    ins_items = [
        [SectionHeading("AI Insights", inner_w)], [sp(6)],
        [insight_row(POS_BG,   "!", f"Top theme: {top_theme_full[:30]}",
                     f"{top_theme_count}/{total} learners cited this as the session standout.")],
        [sp(6)],
        [insight_row(POS_BG,   ">", "Session engagement",
                     f"{engage}/{total} learners provided substantive improvement feedback.")],
        [sp(6)],
        [insight_row(AMBER_BG, "~", f"{no_comment_count} auto-filled 'No comments' entries",
                     f"Real improvement column engagement: {engage}/{total} ({int(engage/total*100) if total else 0}%).")],
        [sp(6)],
        [insight_row(POS_BG if neg_count == 0 else RED_BG,
                     "+" if neg_count == 0 else "!", "Sentiment summary", neg_phrase)],
    ]
    ins_inner = Table(ins_items, colWidths=[inner_w])
    ins_inner.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    ins_card = wrap_in_card(ins_inner, COL2)

    story.append(two_col(sent_card, ins_card))
    story.append(sp(12))

    # ── 4. Response table ──────────────────────────────────────
    story.append(SectionHeading("Individual Responses", USABLE))
    story.append(sp(6))

    CW = [USABLE * p for p in [0.05, 0.29, 0.10, 0.13, 0.43]]
    st_th = S('th', fontName='Sans-B', fontSize=7,  textColor=TBL_ACC, wordWrap='LTR')
    st_td = S('td', fontName='Sans',   fontSize=8,  textColor=TEXT,   leading=11)
    st_tr = S('tr', fontName='Sans-B', fontSize=8,  textColor=RED_TX, leading=11)
    st_tm = S('tm', fontName='Sans',   fontSize=8,  textColor=MUTED,  leading=11)
    st_sr = S('sr', fontName='Mono',   fontSize=8,  textColor=MUTED,  alignment=TA_CENTER)

    tbl_data = [[Paragraph(t, st_th) for t in
                 ["#", "Best Part of Session", "Rating", "Sentiment", "Improvement Feedback"]]]

    for sr, bp, rt, imp, sent in rows:
        il = imp.lower().strip()
        if "no comments from the learner" in il:
            imp_cell = Paragraph("No comments from the Learner", st_tr)
        elif il in {"na", "none", "n/a", "nil", "-", ""}:
            imp_cell = Paragraph(imp if imp.strip() else "-", st_tm)
        else:
            imp_cell = Paragraph(imp, st_td)

        tbl_data.append([
            Paragraph(f"{sr:02d}", st_sr),
            Paragraph(bp.replace(";", ";\n"), st_td),
            RatingBadge(rt),
            SentimentPill(sent),
            imp_cell,
        ])

    avg_str = f"{avg_score:.2f} / 5"
    tbl_data.append([
        Paragraph("Average Rating", S('al', fontName='Sans-B', fontSize=9,
                                      textColor=WHITE, alignment=TA_RIGHT, leading=12)),
        "", "", "",
        Paragraph(avg_str, S('av', fontName='Serif-B', fontSize=11,
                              textColor=GREEN, alignment=TA_RIGHT, leading=12)),
    ])

    rh = [24] + [None] * len(rows) + [26]
    row_bgs = [('BACKGROUND', (0, i), (-1, i), ROW_ALT if i % 2 == 0 else CARD)
               for i in range(1, len(rows) + 1)]

    resp = Table(tbl_data, colWidths=CW, rowHeights=rh)
    resp.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0),  (-1, 0),  TBL_HDR),
        ('LINEBELOW',     (0, 0),  (-1, 0),  0.8, TBL_ACC),
        *row_bgs,
        ('BACKGROUND',    (0, -1), (-1, -1), TBL_HDR),
        ('BOX',           (0, 0),  (-1, -1), 0.5, BORDER),
        ('LINEBELOW',     (0, 1),  (-1, -2), 0.3, BORDER),
        ('ALIGN',         (0, 0),  (-1, -1), 'LEFT'),
        ('ALIGN',         (2, 1),  (3, -2),  'CENTER'),
        ('VALIGN',        (0, 0),  (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0),  (-1, -1), 7),
        ('RIGHTPADDING',  (0, 0),  (-1, -1), 6),
        ('TOPPADDING',    (0, 0),  (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0),  (-1, -1), 5),
        ('ROUNDEDCORNERS',(0, 0),  (-1, -1), [10]),
        # Avg row: span cols 0-3 for label, col 4 for value; wider padding
        ('SPAN',          (0, -1), (3, -1)),
        ('LEFTPADDING',   (0, -1), (-1, -1), 10),
        ('RIGHTPADDING',  (0, -1), (-1, -1), 10),
    ]))
    story.append(resp)
    story.append(sp(12))

    # ── 5. Themes + Improvements ──────────────────────────────
    inner_w2 = COL2 - 28
    st_tb = S('tbd', fontName='Sans',   fontSize=8, textColor=MUTED, leading=12)

    theme_items = [[SectionHeading('Themes in "Best Part"', inner_w2)]]
    for name, count in top3:
        theme_items += [[sp(4)], [dot_row(
            Paragraph(f"<b>{name}</b> - cited by {count}/{total}", st_tb),
            GREEN, inner_w2,
        )]]
    th_inner = Table(theme_items, colWidths=[inner_w2])
    th_inner.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    theme_card = wrap_in_card(th_inner, COL2)

    imp_items = [[SectionHeading("Meaningful Improvement Feedback", inner_w2)]]
    if meaningful:
        for txt in meaningful[:4]:
            imp_items += [[sp(4)], [dot_row(Paragraph(txt, st_tb), AMBER, inner_w2)]]
    if no_comment_count > 0:
        imp_items += [[sp(4)], [dot_row(
            Paragraph(f'{no_comment_count} response(s) auto-flagged "No comments"', st_tb),
            RED, inner_w2,
        )]]

    imp_inner = Table(imp_items, colWidths=[inner_w2])
    imp_inner.setStyle(TableStyle([
        ('LEFTPADDING',   (0, 0), (-1, -1), 0), ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
        ('TOPPADDING',    (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    imp_card = wrap_in_card(imp_inner, COL2)

    story.append(two_col(theme_card, imp_card))
    story.append(sp(14))

    # ── 6. Footer ──────────────────────────────────────────────
    from config.settings import FOOTER_TEXT
    footer_para = Paragraph(
        f"{FOOTER_TEXT}  -  Generated {date}",
        S('ft', fontName='Sans', fontSize=7, textColor=MUTED, alignment=TA_CENTER),
    )
    footer_tbl = Table([[footer_para]], colWidths=[USABLE])
    footer_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), FOOTER_BG),
        ('LEFTPADDING',   (0, 0), (-1, -1), 10),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 10),
        ('TOPPADDING',    (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('ROUNDEDCORNERS',(0, 0), (-1, -1), [6]),
    ]))
    story.append(footer_tbl)

    doc.build(story)
    return output_path
