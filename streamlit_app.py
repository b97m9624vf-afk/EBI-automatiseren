import base64
import json
import os
import re
import streamlit as st

from app import parse_pasted_schedule
from docx_generator import generate_ebi_docx
from pdf_generator import generate_ebi_pdf
from scraper import (
    calculate_loan_cost_summary,
    fetch_html,
    format_euro,
    generate_amortization_schedule,
    lookup_kvk,
    parse_pitch_html,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "collin_logo.svg")

st.set_page_config(
    page_title="Collin Crowdfund — EBI Generator",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def get_logo_b64() -> str:
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""


logo_b64 = get_logo_b64()
logo_img_html = (
    f'<img src="data:image/svg+xml;base64,{logo_b64}" style="height:46px;" alt="Collin Crowdfund" />'
    if logo_b64
    else '<span style="font-weight:800;font-size:22px;color:#1E3764;">COLLIN CROWDFUND</span>'
)

# Authentic Collin Crowdfund styling
st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,600;0,700;0,800;1,400&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Open Sans', sans-serif !important;
    }}
    .block-container {{
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 96% !important;
    }}
    /* Top Collin Navigation Bar */
    .collin-navbar {{
        background: #FFFFFF;
        border-bottom: 3px solid #EF8200;
        border-radius: 10px 10px 6px 6px;
        padding: 14px 24px;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 3px 12px rgba(30, 55, 100, 0.08);
    }}
    .collin-navbar-right {{
        display: flex;
        align-items: center;
        gap: 12px;
    }}
    .collin-pill {{
        background: #1E3764;
        color: #FFFFFF;
        font-size: 12.5px;
        font-weight: 700;
        padding: 5px 12px;
        border-radius: 20px;
        letter-spacing: 0.3px;
    }}
    .collin-pill-orange {{
        background: #FFF4E5;
        color: #D97200;
        border: 1px solid #FCD39B;
        font-size: 12px;
        font-weight: 700;
        padding: 4px 11px;
        border-radius: 20px;
    }}
    /* Collin KPI Summary Bar */
    .collin-kpi-grid {{
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        margin-bottom: 16px;
    }}
    .collin-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #E1E8F0;
        border-top: 3px solid #1E3764;
        border-radius: 8px;
        padding: 10px 14px;
        box-shadow: 0 1px 4px rgba(30, 55, 100, 0.04);
    }}
    .collin-kpi-card.orange-accent {{
        border-top-color: #EF8200;
    }}
    .collin-kpi-card.sky-accent {{
        border-top-color: #7DB8E2;
    }}
    .collin-kpi-label {{
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        color: #64748B;
        letter-spacing: 0.4px;
        margin-bottom: 3px;
    }}
    .collin-kpi-val {{
        font-size: 15px;
        font-weight: 800;
        color: #1E3764;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}
    /* Validation chips */
    .val-bar {{
        background: #FFFFFF;
        border: 1px solid #E1E8F0;
        border-radius: 8px;
        padding: 10px 16px;
        margin-bottom: 16px;
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 10px;
    }}
    .val-chip {{
        font-size: 12px;
        font-weight: 600;
        padding: 3px 10px;
        border-radius: 14px;
        display: inline-flex;
        align-items: center;
        gap: 5px;
    }}
    .val-ok {{
        background: #ECFDF5;
        color: #047857;
        border: 1px solid #A7F3D0;
    }}
    .val-warn {{
        background: #FFFBEB;
        color: #B45309;
        border: 1px solid #FDE68A;
    }}
    /* Tabs styling */
    button[data-baseweb="tab"] {{
        font-weight: 700 !important;
        color: #1E3764 !important;
    }}
    button[data-baseweb="tab"][aria-selected="true"] {{
        color: #EF8200 !important;
    }}
    </style>

    <div class="collin-navbar">
        <div style="display:flex;align-items:center;gap:18px;">
            {logo_img_html}
            <div style="border-left: 2px solid #E2E8F0; padding-left: 16px;">
                <div style="font-size:17px;font-weight:800;color:#1E3764;line-height:1.2;">
                    EBI Document Generator
                </div>
                <div style="font-size:12px;color:#64748B;font-weight:600;">
                    Automatische generatie van het Blad met essentiële informatie (ECSPR / ESMA)
                </div>
            </div>
        </div>
        <div class="collin-navbar-right">
            <span class="collin-pill-orange">Comfortabel investeren</span>
            <span class="collin-pill">EBI Automatisatie</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

DEFAULT_URL = (
    "https://www.collincrowdfund.nl/a-f-j-simons-en-a-m-m-kurver/"
    "?auth=48135dbcd3c64ea6a50cae86ef14d3bb1cfc160d4699e9a8713c7010d3db57eb"
)


def populate_session_from_dict(data: dict):
    st.session_state["ebi_data"] = data
    st.session_state["w_leningnummer"] = data.get("leningnummer", "")
    st.session_state["w_streefbedrag"] = data.get("streefbedrag", "")
    st.session_state["w_lei_code"] = data.get("lei_code", "724500VFNB9F43XDI745")
    st.session_state["w_investeringskans"] = data.get("investeringskans", "")
    st.session_state["w_projecteigenaar"] = data.get("projecteigenaar", "")
    st.session_state["w_deadline"] = data.get("deadline", "30 dagen na publicatiedatum")
    st.session_state["w_rechtsvorm"] = data.get("rechtsvorm", "")
    st.session_state["w_website"] = data.get("website", "n.v.t.")
    st.session_state["w_vestigingsplaats"] = data.get("vestigingsplaats", "")
    st.session_state["w_organigram_text"] = data.get("organigram_text", "")
    st.session_state["w_onderneming_text"] = data.get("onderneming_text", "")
    st.session_state["w_presentatie_intro"] = data.get("presentatie_intro", "")
    st.session_state["w_kasstroom_text"] = data.get("kasstroom_text", "")
    st.session_state["w_kvk_jaarcijfers_text"] = data.get("kvk_jaarcijfers_text", "Niet van toepassing.")
    st.session_state["w_bodembedrag"] = data.get("bodembedrag", "n.v.t.")
    st.session_state["w_inschrijfperiode"] = data.get("inschrijfperiode", "Maximaal 30 dagen na publicatie")
    st.session_state["w_bedenktermijn"] = data.get("bedenktermijn", "Vier kalenderdagen")
    st.session_state["w_collin_invested_note"] = data.get("collin_invested_note", "")
    st.session_state["w_kasstroom_risico_afsluiting"] = data.get("kasstroom_risico_afsluiting", "")
    st.session_state["w_include_herfinanciering"] = bool(data.get("include_herfinanciering", False))
    st.session_state["w_herfinanciering_text"] = data.get("herfinanciering_text", "")
    st.session_state["w_financieringsbehoefte_text"] = data.get("financieringsbehoefte_text", "")
    st.session_state["w_risico_scores_text"] = data.get("risico_scores_text", "")
    st.session_state["w_zekerheden_bullets"] = "\n\n".join(data.get("zekerheden_bullets", []))
    st.session_state["w_historie_betaalverplichtingen"] = data.get("historie_betaalverplichtingen", "")

    lp = data.get("loan_params", {})
    st.session_state["lp_principal"] = float(lp.get("principal", 100000.0))
    st.session_state["lp_rate"] = float(lp.get("annual_rate_pct", 7.5))
    st.session_state["lp_months"] = int(lp.get("months", 96))
    st.session_state["lp_grace"] = int(lp.get("grace_months", 12))
    st.session_state["lp_balloon"] = float(lp.get("balloon_amount", 0.0))
    st.session_state["lp_succes_fee"] = float(lp.get("succes_fee_pct", 2.0))
    st.session_state["lp_admin_pct"] = float(lp.get("admin_fee_pct", 0.07))
    st.session_state["lp_type"] = lp.get("loan_type", "Lineair")


def load_pitch_into_state(url: str):
    html = fetch_html(url)
    data = parse_pitch_html(html, url)
    populate_session_from_dict(data)
    # Save in history of recent pitches
    history = st.session_state.get("pitch_history", {})
    label = f"{data.get('leningnummer', '—')} | {data.get('investeringskans', url)}"
    history[label] = data
    st.session_state["pitch_history"] = history


def run_ebi_quality_checks(payload: dict) -> list[tuple[bool, str]]:
    """Run automated EBI compliance and consistency checks."""
    checks = []
    ln = (payload.get("leningnummer") or "").strip()
    checks.append((len(ln) == 8 and ln.isdigit(), f"Leningnummer ({ln or 'ontbreekt'})"))

    vp = (payload.get("vestigingsplaats") or "").strip()
    checks.append((bool(vp), f"Vestigingsplaats ({vp or 'ontbreekt!'})"))

    gns = payload.get("geldnemers") or []
    gn_ok = len(gns) > 0 and all((g.get("naam") or "").strip() for g in gns)
    checks.append((gn_ok, f"{len(gns)} Geldnemer(s)"))

    schema = payload.get("betaalschema") or []
    lp_months = int(st.session_state.get("lp_months", 96))
    last_saldo = schema[-1].get("saldo", "") if schema else ""
    schema_ok = len(schema) == lp_months + 1 and last_saldo in ("€ 0,00", "€ 0,-", "€ 0")
    checks.append(
        (
            schema_ok,
            f"Betaalschema ({max(0, len(schema)-1)} mnd, eindsaldo {last_saldo or '?'})",
        )
    )

    fb_txt = (payload.get("financieringsbehoefte_text") or "").lower()
    has_slot_word = "slottermijn" in fb_txt or "restschuld" in fb_txt
    herfin_on = bool(payload.get("include_herfinanciering"))
    slot_ok = (has_slot_word == herfin_on) or not has_slot_word
    checks.append(
        (
            slot_ok,
            "Slottermijn / Herfinanciering consistent"
            if slot_ok
            else "Let op: 'slottermijn' in tekst maar kopje Herfinanciering staat uit!",
        )
    )
    return checks


# Sidebar for History & JSON Draft Import/Export
with st.sidebar:
    st.markdown("### 📂 Sessie & Concepten")
    history = st.session_state.get("pitch_history", {})
    if len(history) > 1:
        selected_hist = st.selectbox("Recent opgehaalde pitches", options=list(history.keys()))
        if st.button("Wissel naar geselecteerde pitch", use_container_width=True):
            populate_session_from_dict(history[selected_hist])
            st.rerun()

    st.divider()
    st.markdown("**Concept opslaan / inladen (.json)**")
    uploaded_json = st.file_uploader("Laad eerder opgeslagen EBI (.json)", type=["json"])
    if uploaded_json is not None:
        if st.button("Activeer ingeladen JSON", use_container_width=True):
            loaded_dict = json.loads(uploaded_json.read().decode("utf-8"))
            populate_session_from_dict(loaded_dict)
            st.success("EBI concept ingeladen!")
            st.rerun()

# Top URL Bar
url_col, btn_col = st.columns([4.2, 1.0])
with url_col:
    pitch_url = st.text_input(
        "Collin Crowdfund Pitch URL",
        value=DEFAULT_URL,
        placeholder="Plak hier de Collin Crowdfund pitch-link (bijv. https://www.collincrowdfund.nl/.../?auth=...)",
        label_visibility="collapsed",
    )
with btn_col:
    fetch_clicked = st.button("🚀 Haal Pitch op & Vul EBI", type="primary", use_container_width=True)

if "ebi_data" not in st.session_state or fetch_clicked:
    with st.spinner("Pitch-pagina van Collin Crowdfund wordt opgehaald en geanalyseerd..."):
        try:
            load_pitch_into_state(pitch_url)
            if fetch_clicked:
                st.toast(
                    f"Pitch '{st.session_state['ebi_data'].get('investeringskans')}' succesvol ingeladen!",
                    icon="✅",
                )
        except Exception as e:
            st.error(f"Kon pitch niet ophalen: {e}")

data = st.session_state.get("ebi_data", {})

# Collin KPI Summary Cards
lp_cur = data.get("loan_params", {})
risico_lines = (st.session_state.get("w_risico_scores_text") or "").splitlines()
cc_display = "Ruim voldoende"
for ln_txt in risico_lines:
    if "collin credit score is" in ln_txt.lower():
        cc_display = ln_txt.split("is")[-1].strip().rstrip(".")

st.markdown(
    f"""
    <div class="collin-kpi-grid">
        <div class="collin-kpi-card">
            <div class="collin-kpi-label">Leningnummer</div>
            <div class="collin-kpi-val">{st.session_state.get('w_leningnummer', '—')}</div>
        </div>
        <div class="collin-kpi-card orange-accent">
            <div class="collin-kpi-label">Doelbedrag &amp; Rechtsvorm</div>
            <div class="collin-kpi-val">{st.session_state.get('w_streefbedrag', '—')} <span style="font-size:12px;font-weight:600;color:#64748B;">({st.session_state.get('w_rechtsvorm', '—')})</span></div>
        </div>
        <div class="collin-kpi-card sky-accent">
            <div class="collin-kpi-label">Rente &amp; Looptijd</div>
            <div class="collin-kpi-val">{str(st.session_state.get('lp_rate', 7.5)).replace('.', ',')}% | {st.session_state.get('lp_months', 96)} mnd</div>
        </div>
        <div class="collin-kpi-card">
            <div class="collin-kpi-label">Vestigingsplaats</div>
            <div class="collin-kpi-val">{st.session_state.get('w_vestigingsplaats') or '—'}</div>
        </div>
        <div class="collin-kpi-card orange-accent">
            <div class="collin-kpi-label">Collin Credit Score</div>
            <div class="collin-kpi-val" style="color:#EF8200;">{cc_display}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

left_col, right_col = st.columns([1.12, 0.88], gap="large")

with left_col:
    tabs = st.tabs(
        [
            "1. Basis & KvK",
            "2. Activiteiten & Cijfers",
            "3. Project & Kenmerken",
            "4. Risico & Zekerheden",
            "5. Betaalschema",
        ]
    )

    # TAB 1: Basisgegevens, Geldnemers, KvK & Organigram
    with tabs[0]:
        st.markdown("#### Basisgegevens (Pagina 1 & 2)")
        c1, c2, c3 = st.columns(3)
        with c1:
            leningnummer = st.text_input("Leningnummer (8 cijfers)", key="w_leningnummer")
        with c2:
            streefbedrag = st.text_input("Streefbedrag / Doelbedrag", key="w_streefbedrag")
        with c3:
            lei_code = st.text_input("Lei-code", key="w_lei_code")

        c4, c5 = st.columns(2)
        with c4:
            investeringskans = st.text_input("Investeringskans / Projectnaam", key="w_investeringskans")
        with c5:
            projecteigenaar = st.text_input("Geldnemer(s) / Projecteigenaar", key="w_projecteigenaar")

        c6, c7 = st.columns(2)
        with c6:
            deadline = st.text_input("Deadline", key="w_deadline")
        with c7:
            rechtsvorm = st.text_input("Rechtsvorm", key="w_rechtsvorm")

        st.divider()
        st.markdown("#### Geldnemer(s) & Automatische KvK-opzoeker (Pagina 2)")

        geldnemers_list = data.get("geldnemers", [{"naam": projecteigenaar, "kvk": "n.v.t."}])
        updated_geldnemers = []
        for idx, gn in enumerate(geldnemers_list):
            gc1, gc2, gc_kvk, gc3 = st.columns([2.0, 1.1, 0.7, 0.35])
            with gc1:
                naam_val = st.text_input(
                    f"Geldnemer {idx + 1}",
                    value=gn.get("naam", ""),
                    key=f"gn_naam_{idx}_{leningnummer}",
                )
            with gc2:
                kvk_val = st.text_input(
                    f"KvK-nummer {idx + 1}",
                    value=gn.get("kvk", "n.v.t."),
                    key=f"gn_kvk_{idx}_{leningnummer}",
                )
            with gc_kvk:
                st.write("")
                st.write("")
                if st.button("🔍 Zoek KvK", key=f"lookup_gn_{idx}_{leningnummer}", use_container_width=True):
                    hits = lookup_kvk(naam_val, st.session_state.get("w_vestigingsplaats", ""))
                    if hits:
                        geldnemers_list[idx]["kvk"] = hits[0]["kvk"]
                        geldnemers_list[idx]["naam"] = hits[0]["naam"]
                        st.session_state[f"gn_kvk_{idx}_{leningnummer}"] = hits[0]["kvk"]
                        st.session_state[f"gn_naam_{idx}_{leningnummer}"] = hits[0]["naam"]
                        if hits[0].get("plaats") and not st.session_state.get("w_vestigingsplaats"):
                            st.session_state["w_vestigingsplaats"] = hits[0]["plaats"]
                        st.toast(f"KvK gevonden: {hits[0]['naam']} → {hits[0]['kvk']} ({hits[0].get('plaats', '')})", icon="✅")
                        st.rerun()
                    else:
                        st.warning(f"Geen actieve KvK-inschrijving gevonden voor '{naam_val}'.")
            with gc3:
                st.write("")
                st.write("")
                if st.button("🗑️", key=f"del_gn_{idx}_{leningnummer}") and len(geldnemers_list) > 1:
                    geldnemers_list.pop(idx)
                    st.rerun()
            updated_geldnemers.append({"naam": naam_val, "kvk": kvk_val})

        btn_add_gn, btn_kvk_all = st.columns(2)
        with btn_add_gn:
            if st.button("➕ Geldnemer toevoegen", use_container_width=True):
                geldnemers_list.append({"naam": "", "kvk": "n.v.t."})
                st.rerun()
        with btn_kvk_all:
            if st.button("🔍 Zoek KvK voor alle zakelijke geldnemers", use_container_width=True):
                found_count = 0
                for idx, gn in enumerate(updated_geldnemers):
                    hits = lookup_kvk(gn["naam"], st.session_state.get("w_vestigingsplaats", ""))
                    if hits:
                        updated_geldnemers[idx]["kvk"] = hits[0]["kvk"]
                        st.session_state[f"gn_kvk_{idx}_{leningnummer}"] = hits[0]["kvk"]
                        found_count += 1
                data["geldnemers"] = updated_geldnemers
                if found_count:
                    st.toast(f"KvK-nummer gevonden voor {found_count} geldnemer(s)!", icon="✅")
                    st.rerun()
                else:
                    st.info("Geen KvK-inschrijvingen gevonden voor de opgegeven namen.")
        data["geldnemers"] = updated_geldnemers

        with st.expander("🔎 Vrij zoeken in het KVK Handelsregister (op bedrijfsnaam of plaats)"):
            kvk_q_col, kvk_btn_col = st.columns([3, 1])
            with kvk_q_col:
                kvk_search_query = st.text_input(
                    "Zoek bedrijfsnaam in KVK",
                    placeholder="Bijv. Roboland Beheer B.V. of SolarPark Galecop B.V.",
                    label_visibility="collapsed",
                )
            with kvk_btn_col:
                do_kvk_search = st.button("Zoek in KVK", use_container_width=True)
            if do_kvk_search and kvk_search_query.strip():
                st.session_state["kvk_search_results"] = lookup_kvk(kvk_search_query.strip())
            for r_i, hit in enumerate(st.session_state.get("kvk_search_results", [])):
                hc1, hc2, hc3 = st.columns([2.6, 1.1, 1.1])
                with hc1:
                    st.markdown(f"**{hit['naam']}** — KvK `{hit['kvk']}` ({hit.get('plaats') or '—'})")
                with hc2:
                    if st.button("Als Geldnemer +", key=f"kvk_add_gn_{r_i}", use_container_width=True):
                        data["geldnemers"].append({"naam": hit["naam"], "kvk": hit["kvk"]})
                        if hit.get("plaats") and not st.session_state.get("w_vestigingsplaats"):
                            st.session_state["w_vestigingsplaats"] = hit["plaats"]
                        st.rerun()
                with hc3:
                    if st.button("In Organigram +", key=f"kvk_add_org_{r_i}", use_container_width=True):
                        nodes = data.get("organigram_nodes") or []
                        nodes.append({"naam": hit["naam"], "laag": 1, "is_geldnemer": False, "label": "100%"})
                        data["organigram_nodes"] = nodes
                        st.session_state["org_mode_sel"] = "Meerlaags Organigram (Holdings / Werkmij)"
                        st.rerun()

        w_col, v_col = st.columns(2)
        with w_col:
            website = st.text_input("Website", key="w_website")
        with v_col:
            vestigingsplaats = st.text_input("Vestigingsplaats", key="w_vestigingsplaats")

        st.divider()
        st.markdown("#### Juridische Structuur / Organigram (Pagina 2)")

        default_nodes = data.get("organigram_nodes") or [
            {"naam": st.session_state.get("w_organigram_text") or projecteigenaar, "laag": 1, "is_geldnemer": True, "label": ""}
        ]
        default_mode_idx = 1 if len(default_nodes) > 1 else 0

        org_mode = st.radio(
            "Kies weergave Juridische Structuur:",
            options=[
                "Enkel blok (Standaard)",
                "Meerlaags Organigram (Holdings / Werkmij)",
                "Eigen afbeelding uploaden",
            ],
            index=default_mode_idx,
            horizontal=True,
            key="org_mode_sel",
        )

        uploaded_org_file = None
        if org_mode == "Enkel blok (Standaard)":
            organigram_text = st.text_input(
                "Tekst in het rode blok ('= geldnemer')",
                key="w_organigram_text",
            )
            data["organigram_nodes"] = [
                {"naam": organigram_text, "laag": 1, "is_geldnemer": True, "label": ""}
            ]
        elif org_mode == "Meerlaags Organigram (Holdings / Werkmij)":
            organigram_text = st.session_state.get("w_organigram_text", projecteigenaar)
            st.caption(
                "Stel hieronder de blokken per laag in (Laag 1 = Boven/Holding, Laag 2 = Midden/Werkmaatschappij, Laag 3 = Onder). "
                "Vink **Is Geldnemer** aan voor een rood blok, of laat uit voor een wit blok."
            )
            tp1, tp2, tp3 = st.columns(3)
            with tp1:
                if st.button("Sjabloon: Holding (100%) → Werkmij", use_container_width=True):
                    data["organigram_nodes"] = [
                        {"naam": "Holding B.V.", "laag": 1, "is_geldnemer": False, "label": ""},
                        {"naam": projecteigenaar, "laag": 2, "is_geldnemer": True, "label": "100%"},
                    ]
                    st.rerun()
            with tp2:
                if st.button("Sjabloon: 2 Holdings (50/50%) → Werkmij", use_container_width=True):
                    data["organigram_nodes"] = [
                        {"naam": "Holding A B.V.", "laag": 1, "is_geldnemer": False, "label": ""},
                        {"naam": "Holding B B.V.", "laag": 1, "is_geldnemer": False, "label": ""},
                        {"naam": projecteigenaar, "laag": 2, "is_geldnemer": True, "label": "50% / 50%"},
                    ]
                    st.rerun()
            with tp3:
                if st.button("Sjabloon: Holding + 2 Werkmaatschappijen", use_container_width=True):
                    data["organigram_nodes"] = [
                        {"naam": "Holding B.V.", "laag": 1, "is_geldnemer": True, "label": ""},
                        {"naam": "Werkmaatschappij 1 B.V.", "laag": 2, "is_geldnemer": True, "label": "100%"},
                        {"naam": "Werkmaatschappij 2 B.V.", "laag": 2, "is_geldnemer": True, "label": "100%"},
                    ]
                    st.rerun()

            edited_org_nodes = st.data_editor(
                data.get("organigram_nodes") or default_nodes,
                column_config={
                    "naam": st.column_config.TextColumn("Naam Entiteit / Persoon", required=True),
                    "laag": st.column_config.NumberColumn("Laag (1=Boven, 2=Midden, 3=Onder)", min_value=1, max_value=3, step=1, default=1),
                    "is_geldnemer": st.column_config.CheckboxColumn("Is Geldnemer? (Rood)", default=True),
                    "label": st.column_config.TextColumn("Aandelen % (bijv. 100%)", default=""),
                },
                num_rows="dynamic",
                use_container_width=True,
                key=f"org_nodes_editor_{leningnummer}",
            )
            data["organigram_nodes"] = edited_org_nodes
        else:
            organigram_text = st.session_state.get("w_organigram_text", projecteigenaar)
            uploaded_org_file = st.file_uploader(
                "Upload een eigen organigram-screenshot/afbeelding (PNG/JPG)",
                type=["png", "jpg", "jpeg"],
            )

    # TAB 2: Onderneming & Financiële Cijfers
    with tabs[1]:
        st.markdown("#### Activiteiten & Financiële Cijfers (Pagina 2–4)")
        onderneming_text = st.text_area(
            "Belangrijkste activiteiten van de projecteigena(a)r(en) ('Onderneming')",
            height=110,
            key="w_onderneming_text",
        )
        presentatie_intro = st.text_area(
            "Inleiding presentatie financiële cijfers",
            height=80,
            key="w_presentatie_intro",
        )
        kasstroom_text = st.text_area(
            "Toelichting jaarcijfers & Kasstroom (komt bij Financiële cijfers én bij Risicofactoren → Kasstroom)",
            height=135,
            key="w_kasstroom_text",
        )

        st.markdown("#### Jaarcijfers Tabellen uit de Pitch (aan/uit te vinken & bewerkbaar)")
        for t_idx, j_tbl in enumerate(data.get("jaarcijfers_tables", [])):
            with st.expander(
                f"{'✅' if j_tbl.get('included', True) else '⬜'} {j_tbl.get('title', f'Tabel {t_idx + 1}')}",
                expanded=j_tbl.get("included", True),
            ):
                inc = st.checkbox(
                    "Deze tabel opnemen in het EBI document",
                    value=j_tbl.get("included", True),
                    key=f"jc_inc_{t_idx}_{leningnummer}",
                )
                j_tbl["included"] = inc
                new_title = st.text_input(
                    "Tabeltitel",
                    value=j_tbl.get("title", ""),
                    key=f"jc_title_{t_idx}_{leningnummer}",
                )
                j_tbl["title"] = new_title

                rows = j_tbl.get("rows", [])
                if rows:
                    ncols = max(len(r) for r in rows)
                    grid_dicts = []
                    for r in rows:
                        row_d = {}
                        for c_i in range(ncols):
                            row_d[f"Kolom {c_i + 1}"] = r[c_i].get("text", "") if c_i < len(r) else ""
                        grid_dicts.append(row_d)
                    edited_grid = st.data_editor(
                        grid_dicts,
                        key=f"jc_ed_{t_idx}_{leningnummer}",
                        use_container_width=True,
                        num_rows="dynamic",
                    )
                    new_rows = []
                    for r_i, ed_row in enumerate(edited_grid):
                        orig_r = rows[r_i] if r_i < len(rows) else []
                        nr = []
                        for c_i in range(ncols):
                            orig_cell = orig_r[c_i] if c_i < len(orig_r) else {"bold": False, "align": "left"}
                            nr.append(
                                {
                                    "text": str(ed_row.get(f"Kolom {c_i + 1}", "") or ""),
                                    "bold": orig_cell.get("bold", False),
                                    "align": orig_cell.get("align", "left"),
                                }
                            )
                        new_rows.append(nr)
                    j_tbl["rows"] = new_rows

        kvk_jaarcijfers_text = st.text_input(
            "Gedeponeerde jaarcijfers bij de Kamer van Koophandel (D&B)",
            key="w_kvk_jaarcijfers_text",
        )

    # TAB 3: Investeringsoverzicht & Kenmerken
    with tabs[2]:
        st.markdown("#### Investeringsoverzicht & Kenmerken (Pagina 4–5)")
        st.markdown("**Tabel: Beschrijving van het crowdfundingproject**")
        inv_rows = data.get("investeringsoverzicht", [])
        updated_inv_rows = []
        for r_idx, r in enumerate(inv_rows):
            c0 = r[0] if len(r) > 0 else {"text": "", "bold": False}
            c1 = r[1] if len(r) > 1 else {"text": "", "bold": False}
            ic1, ic2, ic3 = st.columns([2.2, 1.2, 0.6])
            with ic1:
                t0 = st.text_input(
                    f"Omschrijving {r_idx + 1}",
                    value=c0.get("text", ""),
                    key=f"inv_c0_{r_idx}_{leningnummer}",
                    label_visibility="collapsed",
                )
            with ic2:
                t1 = st.text_input(
                    f"Bedrag {r_idx + 1}",
                    value=c1.get("text", ""),
                    key=f"inv_c1_{r_idx}_{leningnummer}",
                    label_visibility="collapsed",
                )
            with ic3:
                is_b = st.checkbox(
                    "Vet",
                    value=bool(c0.get("bold")),
                    key=f"inv_b_{r_idx}_{leningnummer}",
                )
            updated_inv_rows.append(
                [
                    {"text": t0, "bold": is_b, "align": "left"},
                    {"text": t1, "bold": is_b, "align": "right"},
                ]
            )
        data["investeringsoverzicht"] = updated_inv_rows

        kc1, kc2, kc3 = st.columns(3)
        with kc1:
            bodembedrag = st.text_input("Bodembedrag financiering", key="w_bodembedrag")
        with kc2:
            inschrijfperiode = st.text_input("Inschrijfperiode", key="w_inschrijfperiode")
        with kc3:
            bedenktermijn = st.text_input("Bedenktermijn", key="w_bedenktermijn")

        collin_invested_note = st.text_input(
            "Opmerking neven-investering Collin medewerkers/bestuurders",
            key="w_collin_invested_note",
        )

    # TAB 4: Risicofactoren, Voorwaarden & Zekerheden
    with tabs[3]:
        st.markdown("#### Risicofactoren, Voorwaarden & Zekerheden (Pagina 5–7)")

        preset_col1, preset_col2, preset_col3 = st.columns(3)
        with preset_col1:
            if st.button("Preset: Historisch + Toekomstig", use_container_width=True):
                st.session_state["w_kasstroom_risico_afsluiting"] = (
                    "De aangeleverde en beoordeelde (cijfer)informatie is in belangrijke mate "
                    "gebaseerd op historische, en toekomstige (cijfer)informatie. Het risico bestaat "
                    "dat de toekomstige kasstroom niet toereikend is om aan de betalingsverplichting(en) "
                    "te kunnen voldoen."
                )
                st.rerun()
        with preset_col2:
            if st.button("Preset: Alleen Historisch", use_container_width=True):
                st.session_state["w_kasstroom_risico_afsluiting"] = (
                    "De aangeleverde en beoordeelde (cijfer)informatie is in belangrijke mate "
                    "gebaseerd op historische (cijfer)informatie. Het risico bestaat "
                    "dat de toekomstige kasstroom niet toereikend is om aan de betalingsverplichting(en) "
                    "te kunnen voldoen."
                )
                st.rerun()
        with preset_col3:
            if st.button("Preset: Box 3 Huuropbrengsten", use_container_width=True):
                st.session_state["w_kasstroom_risico_afsluiting"] = (
                    "Het risico bestaat dat de verwachte huuropbrengsten niet gerealiseerd worden."
                )
                st.rerun()

        kasstroom_risico_afsluiting = st.text_area(
            "Afsluitende risico-zin onder Kasstroom",
            height=80,
            key="w_kasstroom_risico_afsluiting",
        )
        include_herfinanciering = st.checkbox(
            "Kopje 'Herfinanciering' opnemen (bij slottermijn / restschuld)",
            key="w_include_herfinanciering",
        )
        herfinanciering_text = st.text_area(
            "Tekst Herfinanciering",
            height=90,
            key="w_herfinanciering_text",
            disabled=not include_herfinanciering,
        )
        financieringsbehoefte_text = st.text_area(
            "Tekst Financieringsbehoefte",
            height=80,
            key="w_financieringsbehoefte_text",
        )
        risico_scores_text = st.text_area(
            "Tekst Risico(scores) (Dun & Bradstreet + Collin Credit Score)",
            height=68,
            key="w_risico_scores_text",
        )
        zekerheden_raw = st.text_area(
            "De gestelde voorwaarden en zekerheden (scheid bullets met een lege regel)",
            height=220,
            key="w_zekerheden_bullets",
        )

        hb_col1, hb_col2, hb_col3 = st.columns(3)
        with hb_col1:
            if st.button("BKR — Niet bekend (Box 3)", use_container_width=True):
                st.session_state["w_historie_betaalverplichtingen"] = (
                    "Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen "
                    "of meer gehad? Conform recent inzicht vanuit BKR zijn deze niet bekend."
                )
                st.rerun()
        with hb_col2:
            if st.button("D&B — Niet bekend (Zakelijk)", use_container_width=True):
                st.session_state["w_historie_betaalverplichtingen"] = (
                    "Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen "
                    "of meer gehad? Conform recent inzicht vanuit Dun & Bradstreet zijn deze niet bekend."
                )
                st.rerun()
        with hb_col3:
            if st.button("D&B — Wel sprake van geweest", use_container_width=True):
                st.session_state["w_historie_betaalverplichtingen"] = (
                    "Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen "
                    "of meer gehad? Conform recent inzicht vanuit Dun & Bradstreet is hier sprake van geweest."
                )
                st.rerun()

        historie_betaalverplichtingen = st.text_area(
            "Historie betaalverplichtingen (Pagina 8)",
            height=68,
            key="w_historie_betaalverplichtingen",
        )

    # TAB 5: Betaalschema Projecteigenaar (Exacte formules uit Amo EBI.xlsm -> 'Amortisatie EBI')
    with tabs[4]:
        st.markdown("#### Betaalschema Projecteigenaar (conform `Amo EBI.xlsm` → *Amortisatie EBI*)")
        sc1, sc2, sc3, sc4 = st.columns(4)
        with sc1:
            lp_principal = st.number_input("Leenbedrag (€)", step=1000.0, min_value=0.0, key="lp_principal")
            lp_balloon = st.number_input("Slottermijn (€)", step=1000.0, min_value=0.0, key="lp_balloon")
        with sc2:
            lp_months = st.number_input("Looptijd (maanden)", step=1, min_value=1, key="lp_months")
            lp_grace = st.number_input("Aflosvrije periode (mnd)", step=1, min_value=0, key="lp_grace")
        with sc3:
            lp_rate = st.number_input("Rente (% per jaar)", step=0.1, min_value=0.0, key="lp_rate")
            lp_type = st.selectbox("Aflosvorm", options=["Lineair", "Annuïtair"], key="lp_type")
        with sc4:
            lp_succes_fee = st.number_input("Succesfee (%)", step=0.1, min_value=0.0, key="lp_succes_fee")
            lp_admin_pct = st.number_input(
                "Mnd. admin vergoeding (%)",
                step=0.01,
                min_value=0.0,
                format="%.4f",
                key="lp_admin_pct",
            )

        cost_info = calculate_loan_cost_summary(
            principal=lp_principal,
            months=int(lp_months),
            succes_fee_pct=lp_succes_fee,
            admin_fee_pct=lp_admin_pct,
        )

        st.markdown(
            f"""
            <div style="background:#F8FAFC;border:1px solid #DCE2EC;border-left:4px solid #1E3764;border-radius:6px;padding:10px 14px;margin-bottom:12px;font-size:12.5px;">
                <strong style="color:#1E3764;">Berekening Kosten &amp; Vergoedingen (conform <code>Amo EBI.xlsm</code>):</strong><br/>
                • <b>Maandelijkse admin vergoeding ({lp_admin_pct:.2f}% van {format_euro(lp_principal, 0)}):</b>
                  {format_euro(cost_info['monthly_admin_excl'], 2)} excl. BTW + {format_euro(cost_info['monthly_admin_btw'], 2)} (21% BTW) =
                  <b style="color:#EF8200;">{format_euro(cost_info['monthly_admin_incl'], 2)} incl. BTW / mnd</b>
                  <i>(Totaal over {int(lp_months)} mnd: {format_euro(cost_info['total_admin_incl'], 2)} incl. BTW)</i><br/>
                • <b>Eenmalige Succesfee ({lp_succes_fee:.1f}%) + Publicatiekosten ({format_euro(cost_info['publicatiekosten_excl'], 0)}):</b>
                  {format_euro(cost_info['succes_plus_pub_excl'], 2)} excl. BTW + {format_euro(cost_info['succes_plus_pub_btw'], 2)} (21% BTW) =
                  <b>{format_euro(cost_info['succes_plus_pub_incl'], 2)} incl. BTW</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🔄 Genereer / Herbereken Amortisatieschema uit bovenstaande variabelen", type="primary", use_container_width=True):
            data["betaalschema"] = generate_amortization_schedule(
                principal=lp_principal,
                annual_rate_pct=lp_rate,
                months=int(lp_months),
                grace_months=int(lp_grace),
                balloon_amount=lp_balloon,
                succes_fee_pct=lp_succes_fee,
                admin_fee_pct=lp_admin_pct,
                loan_type=lp_type,
            )
            st.success(f"Amortisatieschema herberekend volgens 'Amortisatie EBI' formules ({len(data['betaalschema'])} rijen).")

        st.divider()
        pasted_sched = st.text_area(
            "Of plak hier optioneel een afwijkend schema uit Excel / CCS",
            height=75,
            placeholder="Kopieer en plak hier de rijen uit Excel (Termijn | Termijnbedrag | Terugbetaling | Rente en kosten | Saldo)...",
        )
        if st.button("📋 Laad geplakt schema in vanuit Excel") and pasted_sched.strip():
            parsed_rows = parse_pasted_schedule(pasted_sched)
            if parsed_rows:
                data["betaalschema"] = parsed_rows
                st.success(f"{len(parsed_rows)} rijen uit Excel/CCS ingeladen!")
            else:
                st.error("Kon geen geldig schema herkennen uit de geplakte tekst.")

        with st.expander(f"Bekijk / bewerk volledig betaalschema ({len(data.get('betaalschema', []))} rijen)", expanded=True):
            edited_schema = st.data_editor(
                data.get("betaalschema", []),
                use_container_width=True,
                height=320,
                key=f"schema_ed_{leningnummer}",
            )
            data["betaalschema"] = edited_schema

# Collect latest values from Streamlit state into `current_payload`
bullets_list = [
    re.sub(r"^[•\-]\s*", "", b.strip())
    for b in re.split(r"\n{2,}|\n(?=[•\-])", zekerheden_raw or "")
    if b.strip()
]

current_payload = {
    **data,
    "lei_code": lei_code,
    "leningnummer": leningnummer,
    "investeringskans": investeringskans,
    "projectnaam": investeringskans,
    "projecteigenaar": projecteigenaar,
    "streefbedrag": streefbedrag,
    "deadline": deadline,
    "rechtsvorm": rechtsvorm,
    "website": website,
    "vestigingsplaats": vestigingsplaats,
    "organigram_text": organigram_text,
    "onderneming_text": onderneming_text,
    "presentatie_intro": presentatie_intro,
    "kasstroom_text": kasstroom_text,
    "kvk_jaarcijfers_text": kvk_jaarcijfers_text,
    "bodembedrag": bodembedrag,
    "inschrijfperiode": inschrijfperiode,
    "bedenktermijn": bedenktermijn,
    "collin_invested_note": collin_invested_note,
    "kasstroom_risico_afsluiting": kasstroom_risico_afsluiting,
    "include_herfinanciering": include_herfinanciering,
    "herfinanciering_text": herfinanciering_text,
    "financieringsbehoefte_text": financieringsbehoefte_text,
    "risico_scores_text": risico_scores_text,
    "zekerheden_bullets": bullets_list,
    "historie_betaalverplichtingen": historie_betaalverplichtingen,
}

# Automated EBI Quality Checklist
checks = run_ebi_quality_checks(current_payload)
chips_html = "".join(
    f'<span class="val-chip {"val-ok" if ok else "val-warn"}">{"✓" if ok else "⚠️"} {label}</span>'
    for ok, label in checks
)
st.markdown(
    f'<div class="val-bar"><strong style="font-size:12.5px;color:#1E3764;">EBI Kwaliteitscheck:</strong>{chips_html}</div>',
    unsafe_allow_html=True,
)

org_bytes = uploaded_org_file.read() if uploaded_org_file is not None else None
pdf_bytes = generate_ebi_pdf(current_payload, organigram_image_bytes=org_bytes)
docx_bytes = generate_ebi_docx(current_payload)

safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", investeringskans or "EBI").strip("_")
file_stem = f"EBI_{leningnummer}_{safe_name}".replace("__", "_")

with right_col:
    dl1, dl2, dl3 = st.columns([1.2, 1.1, 0.9])
    with dl1:
        st.download_button(
            label="⬇️ Download PDF (.pdf)",
            data=pdf_bytes,
            file_name=f"{file_stem}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            label="📝 Download Word (.docx)",
            data=docx_bytes,
            file_name=f"{file_stem}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl3:
        st.download_button(
            label="💾 Concept (.json)",
            data=json.dumps(current_payload, indent=2, ensure_ascii=False).encode("utf-8"),
            file_name=f"{file_stem}.json",
            mime="application/json",
            use_container_width=True,
        )

    b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="780px" style="border: 1px solid #DCE2EC; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);"></iframe>',
        unsafe_allow_html=True,
    )
