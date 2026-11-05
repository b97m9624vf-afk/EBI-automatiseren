import io
import os
import zipfile
from xml.sax.saxutils import escape

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DOCX = os.path.join(BASE_DIR, "Blad met essentiele informatie (format leeg) EBI.docx")


def _w_p(text: str = "", style: str | None = None, bold: bool = False, italic: bool = False, underline: bool = False, bullet: bool = False, page_break_before: bool = False) -> str:
    """Create a WordprocessingML <w:p> element."""
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{escape(style)}"/>')
    if page_break_before:
        ppr_parts.append("<w:pageBreakBefore/>")
    if bullet:
        ppr_parts.append('<w:ind w:left="360" w:hanging="180"/>')
    ppr = f"<w:pPr>{''.join(ppr_parts)}</w:pPr>" if ppr_parts else ""

    rpr_parts = []
    if bold:
        rpr_parts.append("<w:b/>")
    if italic:
        rpr_parts.append("<w:i/>")
    if underline:
        rpr_parts.append('<w:u w:val="single"/>')
    rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>" if rpr_parts else ""

    prefix = "•  " if bullet else ""
    lines = (prefix + (text or "")).split("\n")
    runs = []
    for idx, line in enumerate(lines):
        br = "<w:br/>" if idx > 0 else ""
        runs.append(f'{br}<w:t xml:space="preserve">{escape(line)}</w:t>')
    run_xml = f"<w:r>{rpr}{''.join(runs)}</w:r>" if text or bullet else ""
    return f"<w:p>{ppr}{run_xml}</w:p>"


def _w_table(rows: list[list[dict]], col_widths_dxa: list[int] | None = None, compact: bool = False) -> str:
    """Create a WordprocessingML <w:tbl> element with grid borders."""
    border_sz = "2" if compact else "4"
    border_color = "BFBFBF" if compact else "333333"
    borders = (
        f'<w:tblBorders>'
        f'<w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'<w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'<w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'<w:insideH w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'<w:insideV w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>'
        f'</w:tblBorders>'
    )
    tbl_pr = f"<w:tblPr>{borders}</w:tblPr>"
    tr_list = []
    for r in rows:
        tc_list = []
        for idx, cell in enumerate(r):
            w_dxa = col_widths_dxa[idx] if col_widths_dxa and idx < len(col_widths_dxa) else 2500
            align = cell.get("align", "left")
            jc = f'<w:jc w:val="{align}"/>' if align in ("right", "center") else ""
            rpr_parts = []
            if cell.get("bold"):
                rpr_parts.append("<w:b/>")
            if compact:
                rpr_parts.append('<w:sz w:val="12"/><w:szCs w:val="12"/>')
            rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>" if rpr_parts else ""
            ppr = f"<w:pPr>{jc}<w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>"
            text = escape(str(cell.get("text", "")))
            fn_id = cell.get("footnote_id")
            fn_xml = f'<w:r><w:rPr><w:rStyle w:val="Voetnootmarkering"/></w:rPr><w:footnoteReference w:id="{fn_id}"/></w:r>' if fn_id else ""
            tc_list.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{w_dxa}" w:type="dxa"/></w:tcPr>'
                f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>{fn_xml}</w:p></w:tc>'
            )
        tr_list.append(f"<w:tr>{''.join(tc_list)}</w:tr>")
    return f"<w:tbl>{tbl_pr}{''.join(tr_list)}</w:tbl>"


def generate_ebi_docx(data: dict) -> bytes:
    """Generate a populated Word (.docx) document based on the Collin Crowdfund EBI template."""
    with zipfile.ZipFile(TEMPLATE_DOCX, "r") as zin:
        orig_doc_xml = zin.read("word/document.xml").decode("utf-8")
        import re

        sect_m = re.search(r"<w:sectPr.*?</w:sectPr>", orig_doc_xml, re.S)
        sect_pr = sect_m.group(0) if sect_m else ""

        body_parts = []
        body_parts.append(_w_p("Blad met essentiële informatie", style="Kop1"))
        body_parts.append(_w_p(""))
        body_parts.append(_w_p(f"Lei-code: {data.get('lei_code', '724500VFNB9F43XDI745')}"))
        body_parts.append(_w_p(f"Leningnummer: {data.get('leningnummer', '')}"))
        body_parts.append(_w_p(f"Investeringskans: {data.get('investeringskans', '')}"))
        body_parts.append(_w_p(""))
        body_parts.append(_w_p("Algehele mededeling vooraf", style="Kop2"))
        body_parts.append(
            _w_p(
                "“Dit crowdfundingaanbod is niet geverifieerd noch goedgekeurd door de "
                "bevoegde autoriteiten of de Europese Autoriteit voor effecten en markten (ESMA)."
            )
        )
        body_parts.append(
            _w_p(
                "De geschiktheid van uw ervaring en kennis is niet noodzakelijk beoordeeld "
                "voordat u toegang hebt gekregen tot deze belegging. Door deze belegging te "
                "doen, neemt u het volledige risico op u van deze belegging, waaronder het risico "
                "van gedeeltelijk of volledig verlies van het belegde geld.”"
            )
        )
        # Paragraph with Footnotes 2 and 3 (which render as 1 and 2 in the template's footnotes.xml)
        body_parts.append(
            '<w:p><w:r><w:t xml:space="preserve">“Beleggen in dit crowdfundingproject brengt risico\'s met zich mee, waaronder het risico van geheel of gedeeltelijk verlies van het belegde geld. Uw belegging valt niet onder de depositogarantieregelingen die zijn ingesteld overeenkomstig Richtlijn 2014/49/EU van het Europees Parlement en de Raad</w:t></w:r>'
            '<w:r><w:rPr><w:rStyle w:val="Voetnootmarkering"/></w:rPr><w:footnoteReference w:id="2"/></w:r>'
            '<w:r><w:t xml:space="preserve">. Evenmin valt uw belegging onder de beleggerscompensatiestelsels die zijn ingesteld overeenkomstig Richtlijn 97/9/EG van het Europees Parlement en de Raad</w:t></w:r>'
            '<w:r><w:rPr><w:rStyle w:val="Voetnootmarkering"/></w:rPr><w:footnoteReference w:id="3"/></w:r>'
            '<w:r><w:t>.</w:t></w:r></w:p>'
        )
        body_parts.append(_w_p("U ontvangt mogelijk geen rendement op uw belegging."))
        body_parts.append(
            _w_p(
                "Dit is geen spaarproduct en wij raden u aan niet meer dan 10% van uw netto "
                "vermogen in crowdfundingprojecten te beleggen."
            )
        )
        body_parts.append(
            _w_p(
                "U kunt de beleggingsinstrumenten mogelijk niet verkopen wanneer u dat wenst. "
                "Zelfs als u ze wel kunt verkopen, zult u mogelijk verlies lijden."
            )
        )
        body_parts.append(_w_p(""))
        body_parts.append(_w_p("Mogelijkheid tot wijzigen reeds verrichte investering", style="Kop3"))
        body_parts.append(
            _w_p(
                "Indien u in deze investeringskans heeft geïnvesteerd en de gedane investering "
                "wenst te wijzigen of te annuleren is dit mogelijk zolang de bedenktermijn van vier "
                "dagen nog niet is verstreken. Deze bedenktermijn gaat in op het moment van de "
                "gedane investering. Deze wijziging kunt u verrichten via uw ‘mijn Collin’ waar u "
                "gaat naar uw ‘investeringen’ vervolgens selecteert u de lening en klikt u op "
                "‘details’. U komt dan bij de pagina uit waar u een tweetal knoppen heeft "
                "waarmee u de investering kunt wijzigen of verwijderen."
            )
        )
        body_parts.append(_w_p(""))

        # Table 1
        t1_rows = [
            [{"text": "Leningnummer"}, {"text": data.get("leningnummer", "")}],
            [
                {"text": "Geldnemer(s) en tevens Projecteigen(a)r(en), hierna te noemen “Projecteigenaar”"},
                {"text": data.get("projecteigenaar", "")},
            ],
            [{"text": "Projectnaam"}, {"text": data.get("projectnaam", "")}],
            [{"text": "Streefbedrag / Doelbedrag"}, {"text": data.get("streefbedrag", "")}],
            [
                {"text": "Deadline (wanneer sluit het project)"},
                {"text": data.get("deadline", "30 dagen na publicatiedatum")},
            ],
        ]
        body_parts.append(_w_table(t1_rows, [4500, 4500]))
        body_parts.append(_w_p(""))

        # Section: Informatie over de projecteigena(a)r(en)
        body_parts.append(
            _w_p("Informatie over de projecteigena(a)r(en) en het crowdfundingproject", bold=True)
        )
        t2_rows = [[{"text": ""}, {"text": ""}, {"text": "KvK-nummer"}]]
        for gn in data.get("geldnemers") or [{"naam": data.get("projecteigenaar", ""), "kvk": "n.v.t."}]:
            t2_rows.append(
                [
                    {"text": "Geldnemer"},
                    {"text": gn.get("naam", "")},
                    {"text": gn.get("kvk", "n.v.t.")},
                ]
            )
        t2_rows.extend(
            [
                [{"text": "Website"}, {"text": data.get("website", "n.v.t.")}, {"text": ""}],
                [{"text": "Vestigingsplaats"}, {"text": data.get("vestigingsplaats", "")}, {"text": ""}],
                [
                    {"text": "Bestuurder(s), eigenaren en/of vennoten"},
                    {"text": "Zie onderstaand organigram."},
                    {"text": ""},
                ],
            ]
        )
        body_parts.append(_w_table(t2_rows, [2500, 2500, 2200]))
        body_parts.append(_w_p(""))

        body_parts.append(_w_p("Juridische structuur", bold=True))
        body_parts.append(_w_p(f"{data.get('organigram_text') or data.get('projecteigenaar', '')} (= geldnemer)", bold=True))
        body_parts.append(_w_p(""))

        body_parts.append(_w_p("Verantwoordelijkheidsverklaring Projecteigenaar", style="Kop3"))
        body_parts.append(
            _w_p(
                "“De Projecteigenaar verklaart dat er, voor zover hem bekend, geen informatie is "
                "weggelaten en geen wezenlijk misleidende of onnauwkeurige informatie wordt "
                "verstrekt. De Projecteigenaar is verantwoordelijk voor het opstellen van dit blad "
                "met essentiële beleggingsinformatie.”."
            )
        )
        body_parts.append(_w_p(""))

        body_parts.append(
            _w_p(
                "Belangrijkste activiteiten van de projecteigena(a)r(en); door de Projecteigenaar aangeboden producten of diensten.",
                style="Kop3",
            )
        )
        for para in (data.get("onderneming_text") or "").split("\n\n"):
            if para.strip():
                body_parts.append(_w_p(para.strip()))

        body_parts.append(
            _w_p("Presentatie van de belangrijkste financiële cijfers vanuit de Projecteigenaar", style="Kop3")
        )
        if data.get("presentatie_intro"):
            body_parts.append(_w_p(data["presentatie_intro"].strip()))
            body_parts.append(_w_p(""))
        for para in (data.get("kasstroom_text") or "").split("\n\n"):
            if para.strip():
                body_parts.append(_w_p(para.strip()))
                body_parts.append(_w_p(""))

        for j_tbl in data.get("jaarcijfers_tables", []):
            if not j_tbl.get("included", True):
                continue
            body_parts.append(_w_p(j_tbl.get("title", ""), bold=True))
            body_parts.append(_w_p(j_tbl.get("subtitle", "In duizendtallen (x 1.000)"), italic=True))
            rows = j_tbl.get("rows", [])
            if rows:
                ncols = max(len(r) for r in rows)
                body_parts.append(_w_table(rows, [int(8500 / max(1, ncols))] * ncols))
                body_parts.append(_w_p(""))

        body_parts.append(
            _w_p(
                "Hieronder treft u de gedeponeerde jaarcijfers bij de Kamer van Koophandel zoals bekend bij Dun & Bradstreet:"
            )
        )
        body_parts.append(_w_p(data.get("kvk_jaarcijfers_text", "Niet van toepassing.")))
        body_parts.append(_w_p(""))

        body_parts.append(_w_p("Beschrijving van het crowdfundingproject", style="Kop3"))
        if data.get("investeringsoverzicht"):
            body_parts.append(_w_table(data["investeringsoverzicht"], [6200, 2200]))
        body_parts.append(_w_p(""))

        body_parts.append(
            _w_p(
                "Belangrijkste kenmerken van het crowdfundingproces en, in voorkomend geval, de voorwaarden voor het aantrekken van kapitaal of het lenen van gelden, naargelang het geval.",
                bold=True,
            )
        )
        t_kenmerken = [
            [{"text": "Streefbedrag financiering"}, {"text": data.get("streefbedrag", "")}],
            [{"text": "Bodembedrag financiering"}, {"text": data.get("bodembedrag", "n.v.t.")}],
            [
                {"text": "Inschrijfperiode investeringskans"},
                {"text": data.get("inschrijfperiode", "Maximaal 30 dagen na publicatie")},
            ],
            [
                {"text": "Situatie bij niet tijdige volschrijving investeringskans"},
                {
                    "text": (
                        "Indien gedurende de publicatieperiode het doelbedrag niet wordt gerealiseerd zal het project geen doorgang vinden.\n"
                        "In overleg met de projecteigenaar kan ervoor gekozen worden om de inschrijfperiode te verlengen.\n"
                        "In overleg met de projecteigenaar kan ervoor gekozen worden om het project onder gewijzigde condities en voorwaarden opnieuw voor te leggen aan de investeerders. De investeerders zullen hierover worden geïnformeerd."
                    )
                },
            ],
            [
                {"text": "Eigen inbreng door Projecteigenaar"},
                {"text": "Zie “Beschrijving van het crowdfundingproject”"},
            ],
            [
                {"text": "Nevenfinanciering"},
                {"text": "Zie “Presentatie van de belangrijkste financiële cijfers vanuit de Projecteigenaar”."},
            ],
            [{"text": "Bedenktermijn"}, {"text": data.get("bedenktermijn", "Vier kalenderdagen")}],
        ]
        body_parts.append(_w_table(t_kenmerken, [4500, 4500]))
        body_parts.append(_w_p(""))
        if data.get("collin_invested_note"):
            body_parts.append(_w_p(data["collin_invested_note"]))
        body_parts.append(_w_p(""))

        body_parts.append(_w_p("Risicofactoren", bold=True))
        body_parts.append(_w_p("Onderstaand volgt een omschrijving van de diverse risico’s van het crowdfundingproject."))
        body_parts.append(_w_p("Kasstroom", bold=True, italic=True, underline=True))
        for para in (data.get("kasstroom_text") or "").split("\n\n"):
            if para.strip():
                body_parts.append(_w_p(para.strip()))
        if data.get("kasstroom_risico_afsluiting"):
            body_parts.append(_w_p(data["kasstroom_risico_afsluiting"]))

        body_parts.append(_w_p("Looptijd", bold=True, italic=True, underline=True))
        body_parts.append(
            _w_p(
                "Er is een risico dat de lening niet gedurende de gepresenteerde looptijd afgelost kan worden. "
                "Dit kan allerlei oorzaken hebben. Waardoor u er rekening mee dient te houden dat de betalingen "
                "mogelijk afwijken van het gepresenteerde betaalschema."
            )
        )
        if data.get("include_herfinanciering") and data.get("herfinanciering_text"):
            body_parts.append(_w_p("Herfinanciering", bold=True, italic=True, underline=True))
            body_parts.append(_w_p(data["herfinanciering_text"]))

        body_parts.append(_w_p("Take-out", bold=True, italic=True, underline=True))
        body_parts.append(_w_p("Zie ‘Risico(scores) en voorwaarden’."))

        body_parts.append(_w_p("Andere risico’s", bold=True, italic=True, underline=True))
        body_parts.append(
            _w_p(
                "Bij elke aanvraag zijn er meer risicofactoren van toepassing. Denk hierbij onder andere aan het "
                "ondernemers- en managementrisico, bedrijfsrisico’s, structuurrisico, operationeel risico, risico van "
                "tegenvallende omzet, risico op tegenvallend rendement, verlies van uw (gedeeltelijke) inleg en risico "
                "op platformstoring."
            )
        )
        body_parts.append(_w_p(""))

        body_parts.append(_w_p("Openbaarmaking in verband met leningen", bold=True))
        body_parts.append(_w_p("Financieringsbehoefte", style="Kop3"))
        for para in (data.get("financieringsbehoefte_text") or "").split("\n\n"):
            if para.strip():
                body_parts.append(_w_p(para.strip()))

        body_parts.append(_w_p("Risico(scores) en voorwaarden", style="Kop3"))
        if data.get("risico_scores_text"):
            body_parts.append(_w_p(data["risico_scores_text"].strip()))
        bullets = data.get("zekerheden_bullets", [])
        if bullets:
            body_parts.append(_w_p("De gestelde voorwaarden en zekerheden zijn:"))
            for b in bullets:
                if b.strip():
                    body_parts.append(_w_p(b.strip(), bullet=True))

        body_parts.append(_w_p("Betaalschema Projecteigenaar", style="Kop3", page_break_before=True))
        schema_rows = [
            [
                {"text": "Termijn", "bold": True},
                {"text": "Termijnbedrag", "bold": True},
                {"text": "Terugbetaling", "bold": True},
                {"text": "Rente en kosten", "bold": True},
                {"text": "Saldo", "bold": True},
            ]
        ]
        for r in data.get("betaalschema", []):
            schema_rows.append(
                [
                    {"text": r.get("termijn", ""), "align": "right"},
                    {"text": r.get("termijnbedrag", ""), "align": "right"},
                    {"text": r.get("terugbetaling", ""), "align": "right"},
                    {"text": r.get("rente_en_kosten", ""), "align": "right"},
                    {"text": r.get("saldo", ""), "align": "right"},
                ]
            )
        body_parts.append(_w_table(schema_rows, [800, 1300, 1300, 1400, 1300], compact=True))
        body_parts.append(_w_p("Historie betaalverplichtingen", style="Kop3"))
        body_parts.append(_w_p(data.get("historie_betaalverplichtingen", "")))

        body_parts.append(
            _w_p(
                "Wat gebeurt er indien de Projecteigenaar niet aan zijn verplichtingen richting de Collin investeerders voldoet?",
                style="Kop3",
                page_break_before=True,
            )
        )
        body_parts.append(
            _w_p(
                "Het kan voor komen dat de Projecteigenaar problemen ondervindt met het terugbetalen van de lening die zij van onze investeerders hebben ontvangen."
            )
        )
        body_parts.append(
            _w_p(
                "Wat doet Collin in geval van een non-betaling? Signalen kunnen komen van de Projecteigenaar zelf, uit de markt, de accountant, de investeerder of via de Crowdfund Coach. Ook kan het signaal komen door problemen rond de maandelijkse incasso van rente (en aflossing). Wanneer er zich een mogelijk probleem aandient, neemt Collin contact op met de Projecteigenaar. Dit met als doel de status in beeld te brengen."
            )
        )
        body_parts.append(_w_p("Voor verdere informatievoorziening verwijzen we u graag naar: https://www.collincrowdfund.nl/beheer/."))
        body_parts.append(_w_p("Vergoedingen, informatie en verhaal", bold=True))
        body_parts.append(_w_p("Overzicht tarieven investeerder(s)", style="Kop3"))

        t_tarieven = [
            [{"text": "Investeren", "bold": True}, {"text": "Kosten (incl. BTW)", "bold": True}],
            [{"text": "Inschrijven"}, {"text": "Gratis"}],
            [{"text": "Contractkosten"}, {"text": "Gratis"}],
            [{"text": "Beheervergoeding", "footnote_id": "4"}, {"text": data.get("beheervergoeding", "0,1%")}],
        ]
        body_parts.append(_w_table(t_tarieven, [4500, 4500]))
        body_parts.append(_w_p("Mogelijkheid tot opvragen additionele informatie", style="Kop3"))
        body_parts.append(_w_p("Aanvullende vragen kunnen worden gesteld onder de pitch of via info@collincrowdfund.nl."))
        body_parts.append(_w_p("Klachtenprocedure", style="Kop3"))
        body_parts.append(
            _w_p(
                "Collin doet haar uiterste best om haar klantgroepen optimaal te bedienen. Uiteraard kan het voor komen dat een Projecteigenaar of investeerder toch niet volledig tevreden is. Hiervoor heeft Collin het klachtenformulier opgesteld. Heeft u onverhoopt een klacht, dan kunt u de procedure hiervoor vinden via de klachtenprocedure zoals weergegeven op de Collin website."
            )
        )

        new_body_xml = f"<w:body>{''.join(body_parts)}{sect_pr}</w:body>"
        new_doc_xml = re.sub(r"<w:body>.*</w:body>", lambda _: new_body_xml, orig_doc_xml, flags=re.S)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
        return out_buf.getvalue()
