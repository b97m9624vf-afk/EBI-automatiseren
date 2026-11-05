import io
import os
from xml.sax.saxutils import escape

from reportlab.graphics.shapes import Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    KeepTogether,
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(BASE_DIR, "fonts")

_FONTS_REGISTERED = False


def register_fonts():
    global _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return
    font_map = {
        "OpenSans": "OpenSans-normal-400.ttf",
        "OpenSans-Italic": "OpenSans-italic-400.ttf",
        "OpenSans-SemiBold": "OpenSans-normal-600.ttf",
        "OpenSans-Bold": "OpenSans-normal-700.ttf",
        "OpenSans-ExtraBold": "OpenSans-normal-800.ttf",
    }
    for name, filename in font_map.items():
        path = os.path.join(FONTS_DIR, filename)
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont(name, path))
    pdfmetrics.registerFontFamily(
        "OpenSans",
        normal="OpenSans",
        bold="OpenSans-Bold",
        italic="OpenSans-Italic",
        boldItalic="OpenSans-Bold",
    )
    _FONTS_REGISTERED = True


def get_styles():
    register_fonts()
    c_body = colors.HexColor("#333333")
    c_h1 = colors.HexColor("#1E3764")
    c_h2 = colors.HexColor("#F18A00")
    c_h3 = colors.HexColor("#7DB8E2")
    c_black = colors.HexColor("#111111")

    return {
        "h1": ParagraphStyle(
            "EBI_H1",
            fontName="OpenSans-ExtraBold",
            fontSize=23,
            leading=28,
            textColor=c_h1,
            spaceBefore=6,
            spaceAfter=18,
        ),
        "h2": ParagraphStyle(
            "EBI_H2",
            fontName="OpenSans-SemiBold",
            fontSize=15.5,
            leading=20,
            textColor=c_h2,
            spaceBefore=12,
            spaceAfter=4,
            keepWithNext=True,
        ),
        "h3": ParagraphStyle(
            "EBI_H3",
            fontName="OpenSans-SemiBold",
            fontSize=11.5,
            leading=16.5,
            textColor=c_h3,
            spaceBefore=12,
            spaceAfter=2,
            keepWithNext=True,
        ),
        "section_bold": ParagraphStyle(
            "EBI_SecBold",
            fontName="OpenSans-Bold",
            fontSize=11.5,
            leading=16.5,
            textColor=c_black,
            spaceBefore=14,
            spaceAfter=6,
            keepWithNext=True,
        ),
        "sub_risk": ParagraphStyle(
            "EBI_SubRisk",
            fontName="OpenSans-Bold",
            fontSize=11.5,
            leading=16.5,
            textColor=c_black,
            spaceBefore=10,
            spaceAfter=2,
            keepWithNext=True,
        ),
        "body": ParagraphStyle(
            "EBI_Body",
            fontName="OpenSans",
            fontSize=11.2,
            leading=17.2,
            textColor=c_body,
            spaceBefore=0,
            spaceAfter=8,
        ),
        "body_tight": ParagraphStyle(
            "EBI_BodyTight",
            fontName="OpenSans",
            fontSize=11.2,
            leading=16.5,
            textColor=c_body,
            spaceBefore=0,
            spaceAfter=3,
        ),
        "bullet": ParagraphStyle(
            "EBI_Bullet",
            fontName="OpenSans",
            fontSize=11.2,
            leading=16.8,
            textColor=c_body,
            leftIndent=36,
            firstLineIndent=0,
            bulletIndent=18,
            spaceBefore=0,
            spaceAfter=3,
        ),
        "table_cell": ParagraphStyle(
            "EBI_TableCell",
            fontName="OpenSans",
            fontSize=11.0,
            leading=15.0,
            textColor=c_body,
        ),
        "table_cell_bold": ParagraphStyle(
            "EBI_TableCellBold",
            fontName="OpenSans-Bold",
            fontSize=11.0,
            leading=15.0,
            textColor=c_black,
        ),
        "fin_title": ParagraphStyle(
            "EBI_FinTitle",
            fontName="OpenSans-Bold",
            fontSize=9.2,
            leading=13.5,
            textColor=c_body,
            spaceBefore=8,
            spaceAfter=1,
            keepWithNext=True,
        ),
        "fin_subtitle": ParagraphStyle(
            "EBI_FinSubtitle",
            fontName="OpenSans-Italic",
            fontSize=9.2,
            leading=13.5,
            textColor=c_body,
            spaceBefore=0,
            spaceAfter=6,
            keepWithNext=True,
        ),
        "fin_cell": ParagraphStyle(
            "EBI_FinCell",
            fontName="OpenSans",
            fontSize=9.2,
            leading=12.0,
            textColor=c_body,
        ),
        "fin_cell_bold": ParagraphStyle(
            "EBI_FinCellBold",
            fontName="OpenSans-Bold",
            fontSize=9.2,
            leading=12.0,
            textColor=c_black,
        ),
    }


def _draw_box_with_text(d: Drawing, x: float, y: float, w: float, h: float, text: str, is_geldnemer: bool = True, font_size: float = 9.5):
    red_color = colors.HexColor("#B80000")
    red_border = colors.HexColor("#8A4B38")
    white_border = colors.HexColor("#1E3764")

    fill_col = red_color if is_geldnemer else colors.white
    stroke_col = red_border if is_geldnemer else white_border
    text_col = colors.white if is_geldnemer else colors.HexColor("#1E3764")

    d.add(
        Rect(
            x,
            y,
            w,
            h,
            fillColor=fill_col,
            strokeColor=stroke_col,
            strokeWidth=0.75,
        )
    )

    clean = (text or "").strip()
    max_chars = int(w / (font_size * 0.52))
    if len(clean) > max_chars and " " in clean:
        words = clean.split()
        mid = len(words) // 2
        line1 = " ".join(words[:mid])
        line2 = " ".join(words[mid:])
        d.add(
            String(
                x + w / 2.0,
                y + h / 2.0 + 2.2 * mm,
                line1,
                fontName="OpenSans-SemiBold" if not is_geldnemer else "OpenSans",
                fontSize=font_size,
                fillColor=text_col,
                textAnchor="middle",
            )
        )
        d.add(
            String(
                x + w / 2.0,
                y + h / 2.0 - 2.8 * mm,
                line2,
                fontName="OpenSans-SemiBold" if not is_geldnemer else "OpenSans",
                fontSize=font_size,
                fillColor=text_col,
                textAnchor="middle",
            )
        )
    else:
        d.add(
            String(
                x + w / 2.0,
                y + h / 2.0 - 1.2 * mm,
                clean,
                fontName="OpenSans-SemiBold" if not is_geldnemer else "OpenSans",
                fontSize=font_size,
                fillColor=text_col,
                textAnchor="middle",
            )
        )


def make_organigram_drawing(text: str, nodes: list[dict] | None = None) -> Drawing:
    """Create either the single-block or multi-layer Collin Crowdfund organigram with '= geldnemer' legend."""
    valid_nodes = [n for n in (nodes or []) if (n.get("naam") or "").strip()]
    red_color = colors.HexColor("#B80000")
    border_color = colors.HexColor("#8A4B38")

    # Single-node classic organigram (matches reference PDF 100%)
    if len(valid_nodes) <= 1:
        single_text = valid_nodes[0]["naam"] if valid_nodes else text
        is_gn = bool(valid_nodes[0].get("is_geldnemer", True)) if valid_nodes else True
        d = Drawing(160 * mm, 30 * mm)
        _draw_box_with_text(d, 0, 2 * mm, 84 * mm, 26 * mm, single_text, is_geldnemer=is_gn, font_size=10.5)
        d.add(
            Rect(
                107 * mm,
                2 * mm,
                30 * mm,
                9.5 * mm,
                fillColor=red_color,
                strokeColor=border_color,
                strokeWidth=0.6,
            )
        )
        d.add(
            String(
                122 * mm,
                5.6 * mm,
                "= geldnemer",
                fontName="OpenSans",
                fontSize=8.5,
                fillColor=colors.white,
                textAnchor="middle",
            )
        )
        return d

    # Multi-layer organigram (1..3 layers, 1..4 boxes per layer)
    layers_map: dict[int, list[dict]] = {}
    for n in valid_nodes:
        lvl = max(1, min(4, int(n.get("laag", 1) or 1)))
        layers_map.setdefault(lvl, []).append(n)

    sorted_levels = sorted(layers_map.keys())
    num_layers = len(sorted_levels)
    box_h = 17 * mm
    v_gap = 11 * mm
    total_h = (num_layers * box_h) + ((num_layers - 1) * v_gap) + 10 * mm

    d = Drawing(160 * mm, total_h)
    chart_area_w = 118 * mm
    line_col = colors.HexColor("#333333")

    layer_positions: dict[int, list[tuple[float, float, dict]]] = {}

    max_row_w = max(
        len(layers_map[lvl]) * min(56 * mm, (chart_area_w - (len(layers_map[lvl]) - 1) * 6 * mm) / len(layers_map[lvl]))
        + (len(layers_map[lvl]) - 1) * 6 * mm
        for lvl in sorted_levels
    )

    for l_idx, lvl in enumerate(sorted_levels):
        row_nodes = layers_map[lvl]
        k = len(row_nodes)
        box_w = min(56 * mm, (chart_area_w - (k - 1) * 6 * mm) / k)
        row_total_w = k * box_w + (k - 1) * 6 * mm
        # Left-align the organigram block at x=0, centering narrower layers relative to the widest layer
        start_x = max(0, (max_row_w - row_total_w) / 2.0)
        # Top layer has highest Y in ReportLab coordinate system
        y = total_h - 4 * mm - (l_idx + 1) * box_h - l_idx * v_gap

        coords = []
        for i, node in enumerate(row_nodes):
            bx = start_x + i * (box_w + 6 * mm)
            _draw_box_with_text(
                d,
                bx,
                y,
                box_w,
                box_h,
                node.get("naam", ""),
                is_geldnemer=bool(node.get("is_geldnemer", True)),
                font_size=9.0,
            )
            coords.append((bx + box_w / 2.0, y, node))
        layer_positions[l_idx] = coords

    # Draw orthogonal connecting lines between consecutive layers
    for l_idx in range(num_layers - 1):
        upper = layer_positions[l_idx]
        lower = layer_positions[l_idx + 1]
        upper_bottom_y = upper[0][1]
        lower_top_y = lower[0][1] + box_h
        mid_y = (upper_bottom_y + lower_top_y) / 2.0

        # Vertical stubs down from upper boxes to mid_y
        for ux, uy, _ in upper:
            d.add(Line(ux, upper_bottom_y, ux, mid_y, strokeColor=line_col, strokeWidth=0.9))

        # Horizontal bus line if multiple boxes in upper or lower layer
        all_x = [ux for ux, _, _ in upper] + [lx for lx, _, _ in lower]
        min_x, max_x = min(all_x), max(all_x)
        if max_x > min_x:
            d.add(Line(min_x, mid_y, max_x, mid_y, strokeColor=line_col, strokeWidth=0.9))

        # Vertical stubs down from mid_y to lower boxes + optional percentage label
        for lx, ly, lnode in lower:
            d.add(Line(lx, mid_y, lx, lower_top_y, strokeColor=line_col, strokeWidth=0.9))
            lbl = (lnode.get("label") or "").strip()
            if lbl:
                d.add(
                    String(
                        lx + 2.2 * mm,
                        lower_top_y + 1.8 * mm,
                        lbl,
                        fontName="OpenSans-SemiBold",
                        fontSize=7.8,
                        fillColor=colors.HexColor("#333333"),
                        textAnchor="start",
                    )
                )

    # Legend box '= geldnemer' at bottom-right
    d.add(
        Rect(
            124 * mm,
            4 * mm,
            30 * mm,
            9.5 * mm,
            fillColor=red_color,
            strokeColor=border_color,
            strokeWidth=0.6,
        )
    )
    d.add(
        String(
            139 * mm,
            7.6 * mm,
            "= geldnemer",
            fontName="OpenSans",
            fontSize=8.5,
            fillColor=colors.white,
            textAnchor="middle",
        )
    )
    return d


class NumberedCanvasWithFootnotes:
    """Custom page decorator to render footnotes on Page 1 and last page, plus page numbers from Page 2."""

    def __init__(self):
        self.total_pages = 0

    def draw_page(self, canvas, doc, is_last_page=False):
        canvas.saveState()
        left_x = 25 * mm

        if doc.page == 1:
            # Draw separator line and Footnotes 1 & 2 at the bottom of page 1
            line_y = 33 * mm
            canvas.setStrokeColor(colors.HexColor("#555555"))
            canvas.setLineWidth(0.5)
            canvas.line(left_x, line_y, left_x + 50 * mm, line_y)

            st = ParagraphStyle(
                "FN",
                fontName="OpenSans",
                fontSize=9.0,
                leading=12.5,
                textColor=colors.HexColor("#333333"),
            )
            fn1 = Paragraph(
                "<super>1</super> Richtlijn 2014/49/EU van het Europees Parlement en de Raad van 16 april 2014 inzake de<br/>"
                "depositogarantiestelsels (PB L 173 van 12.6.2014, blz. 149)",
                st,
            )
            w1, h1 = fn1.wrap(160 * mm, 20 * mm)
            fn1.drawOn(canvas, left_x, line_y - h1 - 2 * mm)

            fn2 = Paragraph(
                "<super>2</super> Richtlijn 97/9/EG van het Europees Parlement en de Raad van 3 maart 1997 inzake de<br/>"
                "beleggerscompensatiestelsels (PB L 84 van 26.3.1997, blz. 22)",
                st,
            )
            w2, h2 = fn2.wrap(160 * mm, 20 * mm)
            fn2.drawOn(canvas, left_x, line_y - h1 - h2 - 3.5 * mm)
        else:
            # Page number at bottom left starting from Page 2
            canvas.setFont("OpenSans", 11.2)
            canvas.setFillColor(colors.HexColor("#333333"))
            canvas.drawString(left_x, 14 * mm, str(doc.page))

        if is_last_page:
            line_y = 40 * mm
            canvas.setStrokeColor(colors.HexColor("#555555"))
            canvas.setLineWidth(0.5)
            canvas.line(left_x, line_y, left_x + 52 * mm, line_y)

            st = ParagraphStyle(
                "FN3",
                fontName="OpenSans",
                fontSize=9.0,
                leading=12.5,
                textColor=colors.HexColor("#333333"),
            )
            fn3 = Paragraph(
                "<super>3</super> Bij elke investering wordt maandelijks een beheervergoeding berekend over het op dat moment "
                "nog openstaande investeringsbedrag. Dit bedrag wordt ingehouden bij de maandelijkse betaling "
                "van rente en aflossing. De beheervergoeding is een bruto bedrag, dus exclusief een eventuele "
                'korting (zie het <u><font color="#0563C1">Collin Loyaliteitsprogramma</font></u>).',
                st,
            )
            w3, h3 = fn3.wrap(160 * mm, 25 * mm)
            fn3.drawOn(canvas, left_x, line_y - h3 - 2 * mm)

        canvas.restoreState()


def generate_ebi_pdf(data: dict, organigram_image_bytes: bytes | None = None) -> bytes:
    """Build the complete EBI PDF document matching the Collin Crowdfund template."""
    styles = get_styles()
    buf = io.BytesIO()

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=25 * mm,
        rightMargin=25 * mm,
        topMargin=25 * mm,
        bottomMargin=36 * mm,
        title=f"Blad met essentiële informatie - {data.get('investeringskans', '')}",
    )

    story = []

    # --- PAGE 1: Header & Algehele mededeling vooraf ---
    story.append(Paragraph("Blad met essentiële informatie", styles["h1"]))
    story.append(Paragraph(f"Lei-code: {escape(data.get('lei_code', '724500VFNB9F43XDI745'))}", styles["body_tight"]))
    story.append(Paragraph(f"Leningnummer: {escape(data.get('leningnummer', ''))}", styles["body_tight"]))
    story.append(Paragraph(f"Investeringskans: {escape(data.get('investeringskans', ''))}", styles["body_tight"]))
    story.append(Spacer(1, 8))

    story.append(Paragraph("Algehele mededeling vooraf", styles["h2"]))
    story.append(
        Paragraph(
            "“Dit crowdfundingaanbod is niet geverifieerd noch goedgekeurd door de "
            "bevoegde autoriteiten of de Europese Autoriteit voor effecten en markten "
            "(ESMA).",
            styles["body"],
        )
    )
    story.append(
        Paragraph(
            "De geschiktheid van uw ervaring en kennis is niet noodzakelijk beoordeeld "
            "voordat u toegang hebt gekregen tot deze belegging. Door deze belegging te "
            "doen, neemt u het volledige risico op u van deze belegging, waaronder het risico "
            "van gedeeltelijk of volledig verlies van het belegde geld.”",
            styles["body"],
        )
    )
    story.append(
        Paragraph(
            "“Beleggen in dit crowdfundingproject brengt risico's met zich mee, waaronder "
            "het risico van geheel of gedeeltelijk verlies van het belegde geld. Uw belegging "
            "valt niet onder de depositogarantieregelingen die zijn ingesteld overeenkomstig "
            "Richtlijn 2014/49/EU van het Europees Parlement en de Raad<super>1</super>. Evenmin valt uw "
            "belegging onder de beleggerscompensatiestelsels die zijn ingesteld "
            "overeenkomstig Richtlijn 97/9/EG van het Europees Parlement en de Raad<super>2</super>.",
            styles["body"],
        )
    )
    story.append(Paragraph("U ontvangt mogelijk geen rendement op uw belegging.", styles["body"]))
    story.append(
        Paragraph(
            "Dit is geen spaarproduct en wij raden u aan niet meer dan 10% van uw netto "
            "vermogen in crowdfundingprojecten te beleggen.",
            styles["body"],
        )
    )
    story.append(
        Paragraph(
            "U kunt de beleggingsinstrumenten mogelijk niet verkopen wanneer u dat wenst. "
            "Zelfs als u ze wel kunt verkopen, zult u mogelijk verlies lijden.",
            styles["body"],
        )
    )
    story.append(Spacer(1, 6))

    story.append(Paragraph("Mogelijkheid tot wijzigen reeds verrichte investering", styles["h3"]))
    story.append(
        Paragraph(
            "Indien u in deze investeringskans heeft geïnvesteerd en de gedane investering "
            "wenst te wijzigen of te annuleren is dit mogelijk zolang de bedenktermijn van vier "
            "dagen nog niet is verstreken. Deze bedenktermijn gaat in op het moment van de "
            "gedane investering. Deze wijziging kunt u verrichten via uw ‘mijn Collin’ waar u "
            "gaat naar uw ‘investeringen’ vervolgens selecteert u de lening en klikt u op "
            "‘details’. U komt dan bij de pagina uit waar u een tweetal knoppen heeft "
            "waarmee u de investering kunt wijzigen of verwijderen.",
            styles["body"],
        )
    )
    story.append(Spacer(1, 8))

    # --- TABLE 1: Leningnummer / Projecteigenaar / Projectnaam / Streefbedrag / Deadline ---
    t1_rows = [
        [
            Paragraph("Leningnummer", styles["table_cell"]),
            Paragraph(escape(data.get("leningnummer", "")), styles["table_cell"]),
        ],
        [
            Paragraph(
                "Geldnemer(s) en tevens<br/>Projecteigen(a)r(en), hierna te noemen<br/>“Projecteigenaar”",
                styles["table_cell"],
            ),
            Paragraph(escape(data.get("projecteigenaar", "")), styles["table_cell"]),
        ],
        [
            Paragraph("Projectnaam", styles["table_cell"]),
            Paragraph(escape(data.get("projectnaam", "")), styles["table_cell"]),
        ],
        [
            Paragraph("Streefbedrag / Doelbedrag", styles["table_cell"]),
            Paragraph(escape(data.get("streefbedrag", "")), styles["table_cell"]),
        ],
        [
            Paragraph("Deadline (wanneer sluit het project)", styles["table_cell"]),
            Paragraph(escape(data.get("deadline", "30 dagen na publicatiedatum")), styles["table_cell"]),
        ],
    ]
    t1 = Table(t1_rows, colWidths=[80 * mm, 80 * mm], hAlign="LEFT")
    t1.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t1)

    # --- SECTION: Informatie over de projecteigena(a)r(en) en het crowdfundingproject ---
    story.append(
        Paragraph(
            "Informatie over de projecteigena(a)r(en) en het crowdfundingproject",
            styles["section_bold"],
        )
    )

    geldnemers = data.get("geldnemers") or [{"naam": data.get("projecteigenaar", ""), "kvk": "n.v.t."}]
    t2_rows = [
        [
            "",
            "",
            Paragraph("KvK-nummer", styles["table_cell"]),
        ]
    ]
    for gn in geldnemers:
        t2_rows.append(
            [
                Paragraph("Geldnemer", styles["table_cell"]),
                Paragraph(escape(gn.get("naam", "")), styles["table_cell"]),
                Paragraph(escape(gn.get("kvk", "n.v.t.")), styles["table_cell"]),
            ]
        )
    n_gn = len(geldnemers)
    t2_rows.extend(
        [
            [
                Paragraph("Website", styles["table_cell"]),
                Paragraph(escape(data.get("website", "n.v.t.")), styles["table_cell"]),
                "",
            ],
            [
                Paragraph("Vestigingsplaats", styles["table_cell"]),
                Paragraph(escape(data.get("vestigingsplaats", "")), styles["table_cell"]),
                "",
            ],
            [
                Paragraph("Bestuurder(s),<br/>eigenaren en/of<br/>vennoten", styles["table_cell"]),
                Paragraph("Zie onderstaand<br/>organigram.", styles["table_cell"]),
                "",
            ],
        ]
    )

    t2 = Table(t2_rows, colWidths=[40 * mm, 40 * mm, 40 * mm], hAlign="LEFT")
    t2.setStyle(
        TableStyle(
            [
                # Header cell for KvK-nummer only in col 2, row 0
                ("GRID", (2, 0), (2, 0), 0.5, colors.HexColor("#333333")),
                # Geldnemer rows across all 3 columns
                ("GRID", (0, 1), (2, n_gn), 0.5, colors.HexColor("#333333")),
                # Bottom 3 rows across first 2 columns only
                ("GRID", (0, n_gn + 1), (1, n_gn + 3), 0.5, colors.HexColor("#333333")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t2)

    # --- SECTION: Juridische structuur ---
    story.append(Paragraph("Juridische structuur", styles["section_bold"]))
    if organigram_image_bytes:
        img_buf = io.BytesIO(organigram_image_bytes)
        img = Image(img_buf)
        max_w = 145 * mm
        max_h = 65 * mm
        iw, ih = img.imageWidth, img.imageHeight
        if iw > 0 and ih > 0:
            scale = min(max_w / iw, max_h / ih, 1.0)
            img.drawWidth = iw * scale
            img.drawHeight = ih * scale
        img.hAlign = "LEFT"
        story.append(img)
    else:
        story.append(
            make_organigram_drawing(
                data.get("organigram_text") or data.get("projecteigenaar", ""),
                nodes=data.get("organigram_nodes"),
            )
        )
    story.append(Spacer(1, 4))

    # --- SECTION: Verantwoordelijkheidsverklaring Projecteigenaar ---
    story.append(Paragraph("Verantwoordelijkheidsverklaring Projecteigenaar", styles["h3"]))
    story.append(
        Paragraph(
            "“De Projecteigenaar verklaart dat er, voor zover hem bekend, geen informatie is "
            "weggelaten en geen wezenlijk misleidende of onnauwkeurige informatie wordt "
            "verstrekt. De Projecteigenaar is verantwoordelijk voor het opstellen van dit blad "
            "met essentiële beleggingsinformatie.”.",
            styles["body"],
        )
    )

    # --- SECTION: Belangrijkste activiteiten ---
    story.append(
        Paragraph(
            "Belangrijkste activiteiten van de projecteigena(a)r(en); door de<br/>"
            "Projecteigenaar aangeboden producten of diensten.",
            styles["h3"],
        )
    )
    for para in (data.get("onderneming_text") or "").split("\n\n"):
        if para.strip():
            story.append(Paragraph(escape(para.strip()).replace("\n", "<br/>"), styles["body"]))

    # --- SECTION: Presentatie van de belangrijkste financiële cijfers ---
    story.append(
        Paragraph(
            "Presentatie van de belangrijkste financiële cijfers vanuit de Projecteigenaar",
            styles["h3"],
        )
    )
    if data.get("presentatie_intro"):
        story.append(Paragraph(escape(data["presentatie_intro"].strip()), styles["body"]))
        story.append(Spacer(1, 6))

    for para in (data.get("kasstroom_text") or "").split("\n\n"):
        if para.strip():
            story.append(Paragraph(escape(para.strip()).replace("\n", "<br/>"), styles["body"]))
            story.append(Spacer(1, 4))

    # Render included Jaarcijfers tables
    for j_tbl in data.get("jaarcijfers_tables", []):
        if not j_tbl.get("included", True):
            continue
        rows = j_tbl.get("rows", [])
        if not rows:
            continue
        num_cols = max(len(r) for r in rows)
        if num_cols == 0:
            continue

        # Determine appropriate column widths based on column count
        if num_cols == 6:
            col_widths = [32 * mm, 21 * mm, 20 * mm, 41 * mm, 21 * mm, 21 * mm]
        elif num_cols == 5:
            col_widths = [32 * mm, 21 * mm, 6 * mm, 41 * mm, 24 * mm]
        elif num_cols == 4:
            col_widths = [32 * mm, 21 * mm, 36 * mm, 21 * mm]
        elif num_cols == 2:
            col_widths = [48 * mm, 24 * mm]
        elif num_cols == 3:
            col_widths = [48 * mm, 24 * mm, 24 * mm]
        else:
            col_widths = [(150 * mm) / num_cols] * num_cols

        rl_rows = []
        for r in rows:
            rl_row = []
            for c_idx in range(num_cols):
                cell = r[c_idx] if c_idx < len(r) else {"text": "", "bold": False}
                st = styles["fin_cell_bold"] if cell.get("bold") else styles["fin_cell"]
                rl_row.append(Paragraph(escape(cell.get("text", "")), st))
            rl_rows.append(rl_row)

        fin_table = Table(rl_rows, colWidths=col_widths, hAlign="LEFT")
        fin_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#444444")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                ]
            )
        )
        block = [
            Paragraph(escape(j_tbl.get("title", "")), styles["fin_title"]),
            Paragraph(escape(j_tbl.get("subtitle", "In duizendtallen (x 1.000)")), styles["fin_subtitle"]),
            fin_table,
            Spacer(1, 8),
        ]
        story.append(KeepTogether(block))

    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "Hieronder treft u de gedeponeerde jaarcijfers bij de Kamer van Koophandel "
            "zoals bekend bij Dun &amp; Bradstreet:",
            styles["body_tight"],
        )
    )
    story.append(Paragraph(escape(data.get("kvk_jaarcijfers_text", "Niet van toepassing.")), styles["body"]))
    story.append(Spacer(1, 8))

    # --- SECTION: Beschrijving van het crowdfundingproject ---
    story.append(Paragraph("Beschrijving van het crowdfundingproject", styles["h3"]))
    inv_rows = data.get("investeringsoverzicht", [])
    if inv_rows:
        inv_rl_rows = []
        for r in inv_rows:
            c0 = r[0] if len(r) > 0 else {"text": "", "bold": False}
            c1 = r[1] if len(r) > 1 else {"text": "", "bold": False}
            st0 = styles["table_cell_bold"] if c0.get("bold") else styles["table_cell"]
            st1_left = styles["table_cell_bold"] if c1.get("bold") else styles["table_cell"]
            st1_right = ParagraphStyle(
                "InvRight",
                parent=st1_left,
                alignment=TA_RIGHT,
            )
            raw_amt = (c1.get("text") or "").strip()
            if raw_amt.startswith("€"):
                amt_num = raw_amt[1:].strip()
                inv_rl_rows.append(
                    [
                        Paragraph(escape(c0.get("text", "")), st0),
                        Paragraph("€", st1_left),
                        Paragraph(escape(amt_num), st1_right),
                    ]
                )
            else:
                inv_rl_rows.append(
                    [
                        Paragraph(escape(c0.get("text", "")), st0),
                        Paragraph("", st1_left),
                        Paragraph(escape(raw_amt), st1_right),
                    ]
                )
        t_inv = Table(inv_rl_rows, colWidths=[110 * mm, 6 * mm, 26 * mm], hAlign="LEFT")
        t_inv.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#444444")),
                    ("INNERGRID", (0, 0), (0, -1), 0.5, colors.HexColor("#444444")),
                    ("LINEBEFORE", (1, 0), (1, -1), 0.5, colors.HexColor("#444444")),
                    ("LINEABOVE", (0, 1), (-1, -1), 0.5, colors.HexColor("#444444")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (0, -1), 5),
                    ("RIGHTPADDING", (0, 0), (0, -1), 5),
                    ("LEFTPADDING", (1, 0), (1, -1), 4),
                    ("RIGHTPADDING", (1, 0), (1, -1), 0),
                    ("LEFTPADDING", (2, 0), (2, -1), 0),
                    ("RIGHTPADDING", (2, 0), (2, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(t_inv)
    story.append(Spacer(1, 8))

    # --- SECTION: Belangrijkste kenmerken van het crowdfundingproces ---
    story.append(
        Paragraph(
            "Belangrijkste kenmerken van het crowdfundingproces en, in voorkomend "
            "geval, de voorwaarden voor het aantrekken van kapitaal of het lenen van "
            "gelden, naargelang het geval.",
            styles["section_bold"],
        )
    )
    story.append(Spacer(1, 4))

    situatie_html = (
        "Indien gedurende de publicatieperiode het doelbedrag niet wordt gerealiseerd zal het project geen doorgang vinden.<br/>"
        "In overleg met de projecteigenaar kan ervoor gekozen worden om de inschrijfperiode te verlengen.<br/>"
        "In overleg met de projecteigenaar kan ervoor gekozen worden om het project onder gewijzigde condities en "
        "voorwaarden opnieuw voor te leggen aan de investeerders. De investeerders zullen hierover worden geïnformeerd."
    )
    t_kenmerken_rows = [
        [
            Paragraph("Streefbedrag financiering", styles["table_cell"]),
            Paragraph(escape(data.get("streefbedrag", "")), styles["table_cell"]),
        ],
        [
            Paragraph("Bodembedrag financiering", styles["table_cell"]),
            Paragraph(escape(data.get("bodembedrag", "n.v.t.")), styles["table_cell"]),
        ],
        [
            Paragraph("Inschrijfperiode investeringskans", styles["table_cell"]),
            Paragraph(escape(data.get("inschrijfperiode", "Maximaal 30 dagen na publicatie")), styles["table_cell"]),
        ],
        [
            Paragraph("Situatie bij niet tijdige volschrijving<br/>investeringskans", styles["table_cell"]),
            Paragraph(situatie_html, styles["table_cell"]),
        ],
        [
            Paragraph("Eigen inbreng door Projecteigenaar", styles["table_cell"]),
            Paragraph(
                'Zie “<u><font color="#0563C1">Beschrijving van het<br/>crowdfundingproject</font></u>”',
                styles["table_cell"],
            ),
        ],
        [
            Paragraph("Nevenfinanciering", styles["table_cell"]),
            Paragraph(
                'Zie “<u><font color="#0563C1">Presentatie van de belangrijkste<br/>financiële cijfers vanuit de<br/>Projecteigenaar</font></u>”.',
                styles["table_cell"],
            ),
        ],
        [
            Paragraph("Bedenktermijn", styles["table_cell"]),
            Paragraph(escape(data.get("bedenktermijn", "Vier kalenderdagen")), styles["table_cell"]),
        ],
    ]
    t_kenmerken = Table(t_kenmerken_rows, colWidths=[80 * mm, 80 * mm], hAlign="LEFT")
    t_kenmerken.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t_kenmerken)
    story.append(Spacer(1, 10))

    if data.get("collin_invested_note"):
        story.append(Paragraph(escape(data["collin_invested_note"]), styles["body"]))

    # --- SECTION: Risicofactoren ---
    story.append(Paragraph("Risicofactoren", styles["section_bold"]))
    story.append(
        Paragraph(
            "Onderstaand volgt een omschrijving van de diverse risico’s van het crowdfundingproject.",
            styles["body"],
        )
    )

    story.append(Paragraph("<u><i>Kasstroom</i></u>", styles["sub_risk"]))
    for para in (data.get("kasstroom_text") or "").split("\n\n"):
        if para.strip():
            story.append(Paragraph(escape(para.strip()).replace("\n", "<br/>"), styles["body"]))
            story.append(Spacer(1, 4))

    if data.get("kasstroom_risico_afsluiting"):
        story.append(Paragraph(escape(data["kasstroom_risico_afsluiting"]), styles["body"]))

    story.append(Paragraph("<u><i>Looptijd</i></u>", styles["sub_risk"]))
    story.append(
        Paragraph(
            "Er is een risico dat de lening niet gedurende de gepresenteerde looptijd afgelost "
            "kan worden. Dit kan allerlei oorzaken hebben. Waardoor u er rekening mee dient "
            "te houden dat de betalingen mogelijk afwijken van het gepresenteerde "
            "betaalschema.",
            styles["body"],
        )
    )

    if data.get("include_herfinanciering") and data.get("herfinanciering_text"):
        story.append(Paragraph("<u><i>Herfinanciering</i></u>", styles["sub_risk"]))
        story.append(Paragraph(escape(data["herfinanciering_text"]), styles["body"]))

    story.append(Paragraph("<u><i>Take-out</i></u>", styles["sub_risk"]))
    story.append(
        Paragraph(
            'Zie ‘<u><font color="#0563C1">Risico(scores) en voorwaarden</font></u>’.',
            styles["body"],
        )
    )

    story.append(Paragraph("<u><i>Andere risico’s</i></u>", styles["sub_risk"]))
    story.append(
        Paragraph(
            "Bij elke aanvraag zijn er meer risicofactoren van toepassing. Denk hierbij onder "
            "andere aan het ondernemers- en managementrisico, bedrijfsrisico’s, "
            "structuurrisico, operationeel risico, risico van tegenvallende omzet, risico op "
            "tegenvallend rendement, verlies van uw (gedeeltelijke) inleg en risico op "
            "platformstoring.",
            styles["body"],
        )
    )

    # --- SECTION: Openbaarmaking in verband met leningen ---
    story.append(Paragraph("Openbaarmaking in verband met leningen", styles["section_bold"]))
    story.append(Paragraph("Financieringsbehoefte", styles["h3"]))
    for para in (data.get("financieringsbehoefte_text") or "").split("\n\n"):
        if para.strip():
            story.append(Paragraph(escape(para.strip()).replace("\n", "<br/>"), styles["body"]))

    story.append(Paragraph("Risico(scores) en voorwaarden", styles["h3"]))
    if data.get("risico_scores_text"):
        story.append(
            Paragraph(
                escape(data["risico_scores_text"].strip()).replace("\n", "<br/>"),
                styles["body"],
            )
        )

    bullets = data.get("zekerheden_bullets", [])
    if bullets:
        story.append(Paragraph("De gestelde voorwaarden en zekerheden zijn:", styles["body_tight"]))
        for b in bullets:
            if b.strip():
                story.append(Paragraph(f"<bullet>&bull;</bullet>{escape(b.strip())}", styles["bullet"]))

    # --- PAGE 8: Betaalschema Projecteigenaar & Historie betaalverplichtingen ---
    story.append(PageBreak())
    story.append(Paragraph("Betaalschema Projecteigenaar", styles["h3"]))

    schema_rows = data.get("betaalschema", [])
    if schema_rows:
        n_rows = len(schema_rows)
        # Automatically scale row font/padding so even 96-120 months fits cleanly on 1 page with Historie
        font_sz = 5.1 if n_rows > 65 else 6.4
        row_h = 5.4 if n_rows > 65 else 8.2

        def _split_euro_cell(val_str: str) -> tuple[str, str]:
            s = (val_str or "").strip()
            if not s:
                return ("", "")
            if s.startswith("€"):
                num_part = s[1:].strip()
                if num_part == "-":
                    return ("€", "-   ")
                return ("€", num_part)
            return ("", s)

        # 9-column Excel Accounting layout: [Termijn, €, Termijnbedrag, €, Terugbetaling, €, Rente en kosten, €, Saldo]
        tbl_data = [
            ["Termijn", "Termijnbedrag", "", "Terugbetaling", "", "Rente en kosten", "", "Saldo", ""]
        ]
        for r in schema_rows:
            tb_e, tb_v = _split_euro_cell(str(r.get("termijnbedrag", "")))
            tr_e, tr_v = _split_euro_cell(str(r.get("terugbetaling", "")))
            rk_e, rk_v = _split_euro_cell(str(r.get("rente_en_kosten", "")))
            sa_e, sa_v = _split_euro_cell(str(r.get("saldo", "")))
            tbl_data.append(
                [
                    str(r.get("termijn", "")),
                    tb_e,
                    tb_v,
                    tr_e,
                    tr_v,
                    rk_e,
                    rk_v,
                    sa_e,
                    sa_v,
                ]
            )

        t_schema = Table(
            tbl_data,
            colWidths=[
                12 * mm,
                3.5 * mm,
                20.5 * mm,
                3.5 * mm,
                19.5 * mm,
                3.5 * mm,
                21.5 * mm,
                3.5 * mm,
                20.5 * mm,
            ],
            rowHeights=[row_h + 1.5] + [row_h] * len(schema_rows),
            hAlign="LEFT",
        )
        t_schema.setStyle(
            TableStyle(
                [
                    # Span header cells across (€, amount) pairs
                    ("SPAN", (1, 0), (2, 0)),
                    ("SPAN", (3, 0), (4, 0)),
                    ("SPAN", (5, 0), (6, 0)),
                    ("SPAN", (7, 0), (8, 0)),
                    ("FONTNAME", (0, 0), (-1, 0), "OpenSans-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "OpenSans"),
                    ("FONTSIZE", (0, 0), (-1, -1), font_sz),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#111111")),
                    # Header alignment
                    ("ALIGN", (0, 0), (-1, 0), "LEFT"),
                    # Termijn numbers right-aligned
                    ("ALIGN", (0, 1), (0, -1), "RIGHT"),
                    # '€' signs left-aligned in cols 1, 3, 5, 7
                    ("ALIGN", (1, 1), (1, -1), "LEFT"),
                    ("ALIGN", (3, 1), (3, -1), "LEFT"),
                    ("ALIGN", (5, 1), (5, -1), "LEFT"),
                    ("ALIGN", (7, 1), (7, -1), "LEFT"),
                    # Currency values right-aligned in cols 2, 4, 6, 8
                    ("ALIGN", (2, 1), (2, -1), "RIGHT"),
                    ("ALIGN", (4, 1), (4, -1), "RIGHT"),
                    ("ALIGN", (6, 1), (6, -1), "RIGHT"),
                    ("ALIGN", (8, 1), (8, -1), "RIGHT"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    # Paddings
                    ("LEFTPADDING", (0, 0), (0, -1), 2.0),
                    ("RIGHTPADDING", (0, 0), (0, -1), 2.5),
                    ("LEFTPADDING", (1, 0), (-1, 0), 2.5),
                    ("RIGHTPADDING", (1, 0), (-1, 0), 2.0),
                    ("LEFTPADDING", (1, 1), (1, -1), 2.2),
                    ("RIGHTPADDING", (1, 1), (1, -1), 0),
                    ("LEFTPADDING", (3, 1), (3, -1), 2.2),
                    ("RIGHTPADDING", (3, 1), (3, -1), 0),
                    ("LEFTPADDING", (5, 1), (5, -1), 2.2),
                    ("RIGHTPADDING", (5, 1), (5, -1), 0),
                    ("LEFTPADDING", (7, 1), (7, -1), 2.2),
                    ("RIGHTPADDING", (7, 1), (7, -1), 0),
                    ("LEFTPADDING", (2, 1), (2, -1), 0),
                    ("RIGHTPADDING", (2, 1), (2, -1), 2.2),
                    ("LEFTPADDING", (4, 1), (4, -1), 0),
                    ("RIGHTPADDING", (4, 1), (4, -1), 2.2),
                    ("LEFTPADDING", (6, 1), (6, -1), 0),
                    ("RIGHTPADDING", (6, 1), (6, -1), 2.2),
                    ("LEFTPADDING", (8, 1), (8, -1), 0),
                    ("RIGHTPADDING", (8, 1), (8, -1), 2.2),
                    ("TOPPADDING", (0, 0), (-1, -1), 0.1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0.1),
                    # Outer border + horizontal row borders + vertical borders only between the 5 logical columns
                    ("BOX", (0, 0), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("LINEABOVE", (0, 1), (-1, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("LINEBEFORE", (1, 0), (1, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("LINEBEFORE", (3, 0), (3, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("LINEBEFORE", (5, 0), (5, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("LINEBEFORE", (7, 0), (7, -1), 0.25, colors.HexColor("#BFBFBF")),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F5F7FA")),
                ]
            )
        )
        story.append(t_schema)

    story.append(Spacer(1, 3))
    story.append(Paragraph("Historie betaalverplichtingen", styles["h3"]))
    story.append(
        Paragraph(
            escape(
                data.get(
                    "historie_betaalverplichtingen",
                    "Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen of meer gehad? "
                    "Conform recent inzicht vanuit BKR zijn deze niet bekend.",
                )
            ),
            styles["body_tight"],
        )
    )

    # --- PAGE 9: Non-betaling, Vergoedingen, Klachtenprocedure ---
    story.append(PageBreak())
    story.append(
        Paragraph(
            "Wat gebeurt er indien de Projecteigenaar niet aan zijn verplichtingen richting<br/>"
            "de Collin investeerders voldoet?",
            styles["h3"],
        )
    )
    story.append(
        Paragraph(
            "Het kan voor komen dat de Projecteigenaar problemen ondervindt met het "
            "terugbetalen van de lening die zij van onze investeerders hebben ontvangen.",
            styles["body"],
        )
    )
    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "Wat doet Collin in geval van een non-betaling? Signalen kunnen komen van de "
            "Projecteigenaar zelf, uit de markt, de accountant, de investeerder of via de "
            "Crowdfund Coach. Ook kan het signaal komen door problemen rond de "
            "maandelijkse incasso van rente (en aflossing). Wanneer er zich een mogelijk "
            "probleem aandient, neemt Collin contact op met de Projecteigenaar. Dit met als "
            "doel de status in beeld te brengen.",
            styles["body"],
        )
    )
    story.append(
        Paragraph(
            "Voor verdere informatievoorziening verwijzen we u graag naar:<br/>"
            '<u><font color="#0563C1">https://www.collincrowdfund.nl/beheer/</font></u>.',
            styles["body"],
        )
    )
    story.append(Spacer(1, 8))

    story.append(Paragraph("Vergoedingen, informatie en verhaal", styles["section_bold"]))
    story.append(Paragraph("Overzicht tarieven investeerder(s)", styles["h3"]))
    story.append(Spacer(1, 4))

    t_tarieven_rows = [
        [
            Paragraph("Investeren", styles["table_cell_bold"]),
            Paragraph("Kosten (incl. BTW)", styles["table_cell_bold"]),
        ],
        [
            Paragraph("Inschrijven", styles["table_cell"]),
            Paragraph("Gratis", styles["table_cell"]),
        ],
        [
            Paragraph("Contractkosten", styles["table_cell"]),
            Paragraph("Gratis", styles["table_cell"]),
        ],
        [
            Paragraph("Beheervergoeding<super>3</super>", styles["table_cell"]),
            Paragraph(escape(data.get("beheervergoeding", "0,1%")), styles["table_cell"]),
        ],
    ]
    t_tarieven = Table(t_tarieven_rows, colWidths=[80 * mm, 80 * mm], hAlign="LEFT")
    t_tarieven.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(t_tarieven)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Mogelijkheid tot opvragen additionele informatie", styles["h3"]))
    story.append(
        Paragraph(
            "Aanvullende vragen kunnen worden gesteld onder de pitch of via<br/>"
            '<u><font color="#0563C1">info@collincrowdfund.nl</font></u>.',
            styles["body"],
        )
    )
    story.append(Spacer(1, 6))

    story.append(Paragraph("Klachtenprocedure", styles["h3"]))
    story.append(
        Paragraph(
            "Collin doet haar uiterste best om haar klantgroepen optimaal te bedienen. "
            "Uiteraard kan het voor komen dat een Projecteigenaar of investeerder toch niet "
            "volledig tevreden is. Hiervoor heeft Collin het klachtenformulier opgesteld. Heeft "
            "u onverhoopt een klacht, dan kunt u de procedure hiervoor vinden via de "
            '<u><font color="#0563C1">klachtenprocedure</font></u> zoals weergegeven op de Collin website.',
            styles["body"],
        )
    )

    # First pass to count total pages so we know which page gets Footnote 3
    temp_buf = io.BytesIO()
    temp_doc = SimpleDocTemplate(
        temp_buf,
        pagesize=A4,
        leftMargin=25 * mm,
        rightMargin=25 * mm,
        topMargin=25 * mm,
        bottomMargin=36 * mm,
    )
    import copy

    temp_doc.build(copy.deepcopy(story))
    total_pages = temp_doc.page

    decorator = NumberedCanvasWithFootnotes()
    decorator.total_pages = total_pages

    def on_page(canvas, doc_obj):
        decorator.draw_page(canvas, doc_obj, is_last_page=(doc_obj.page == total_pages))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()
