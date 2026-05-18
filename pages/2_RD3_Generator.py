"""
pages/2_RD3_Generator.py  — RD3 Waterproofness Final Report Generator
Standalone Streamlit page. Entry point remains rd6_app.py.
"""
import os, io, tempfile
from pathlib import Path
from datetime import date
import streamlit as st

# ── Paths (pages/ is one level below repo root) ───────────────────────────────
BASE       = Path(__file__).parent.parent
_tpl_underscore = BASE / "Template_RD3.docx"
_tpl_space      = BASE / "Template RD3.docx"
TPL = _tpl_underscore if _tpl_underscore.exists() else _tpl_space
EXCEL      = BASE / "malath_log.xlsx"
TEAM_EXCEL = BASE / "IDI_Team.xlsx"

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="RD3 Generator · SOCOTEC Arabia",
                   page_icon="🌊", layout="wide")

st.markdown("""
<style>
.step-title {
    font-size: 1.15rem; font-weight: 700; color: #1f4e79;
    border-left: 5px solid #0072BB; padding-left: 10px; margin-bottom: 1rem;
}
</style>""", unsafe_allow_html=True)

# ── Import helpers (same repo) ────────────────────────────────────────────────
try:
    from rd6_extractor import (extract_from_policy_pdf, lookup_from_excel,
                                extract_date_from_cert, load_engineer_team)
    from rd3_generator import generate_rd3, build_rd3_reference
    IMPORTS_OK = True
except ImportError as _ie:
    IMPORTS_OK = False
    st.error(f"Import error: {_ie}. Make sure rd6_extractor.py and rd3_generator.py are in the repo root.")
    st.stop()

# ── Load engineer team once ───────────────────────────────────────────────────
@st.cache_data
def get_team():
    if TEAM_EXCEL.exists():
        return load_engineer_team(str(TEAM_EXCEL))
    return {}

TEAM = get_team()
TEAM_NAMES = sorted(TEAM.keys())

# ── Session initialisation ────────────────────────────────────────────────────
_defaults = [
    ('rd3_step', 1),
    ('rd3_data', {}),
    ('rd3_visits', []),
    ('rd3_annex_data', {}),
]
for k, v in _defaults:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = BASE / "socotec_logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=160)
    else:
        st.markdown("## 🌊 RD3 Generator")
    st.markdown("**SOCOTEC Arabia · TIS Division**")
    st.markdown("---")
    labels = ["Engineer & Signature", "Policy Upload", "Project Info",
              "Site Visits", "WP Works & Annexes", "Generate"]
    cur = st.session_state.rd3_step
    for i, lbl in enumerate(labels, 1):
        icon = "✅" if i < cur else ("🔵" if i == cur else "⬜")
        md = "**{} {}. {}**".format(icon, i, lbl) if i == cur else "{} {}. {}".format(icon, i, lbl)
        st.markdown(md)
    st.markdown("---")
    st.markdown(
        '<div style="position:fixed;bottom:18px;left:12px;width:255px;'
        'font-size:0.72rem;color:#4a90a4;border-top:1px solid #2a4a5a;'
        'padding-top:8px;line-height:1.6">'
        '⚙️ Built by<br>'
        '<strong style="color:#5ba8c4">Eng. Mohamed Mossad</strong><br>'
        '<span style="color:#888">SOCOTEC Arabia · TIS Division</span>'
        '</div>',
        unsafe_allow_html=True
    )
    if st.button("🔄 Restart", key="rd3_restart"):
        for k, v in _defaults:
            st.session_state[k] = v
        st.session_state.rd3_step = 1
        st.rerun()

step = st.session_state.rd3_step

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Engineer Details
# ═══════════════════════════════════════════════════════════════════════════════
if step == 1:
    st.markdown('<div class="step-title">Step 1 — Engineer Details</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Report Author (Engineer)**")
        if TEAM_NAMES:
            options = [''] + TEAM_NAMES
            selected = st.selectbox("Full Name *", options=options, index=0,
                                    help="Select name — phone/email auto-fill",
                                    key="rd3_eng_name")

            # ── Auto-fill: directly write into widget session state + rerun ──
            if selected != st.session_state.get('_rd3_prev_eng', '__UNSET__'):
                st.session_state['_rd3_prev_eng'] = selected
                if selected and selected in TEAM:
                    st.session_state['rd3_phone'] = TEAM[selected]['phone']
                    st.session_state['rd3_email'] = TEAM[selected]['email']
                else:
                    st.session_state['rd3_phone'] = ''
                    st.session_state['rd3_email'] = ''
                st.rerun()

            name = selected
            # ── Save to dedicated backup key — survives all reruns ───────────
            if name and name.strip():
                st.session_state['_rd3_eng_backup'] = name.strip()
                st.session_state.rd3_data['eng_full'] = name.strip()
        else:
            name = st.text_input("Full Name (First Last) *", placeholder="Mohamed Mossad",
                                 key="rd3_eng_name_text")
            if name and name.strip():
                st.session_state['_rd3_eng_backup'] = name.strip()
                st.session_state.rd3_data['eng_full'] = name.strip()

        if name:
            parts = name.strip().split()
            pfx = (parts[0][0] + parts[1][:2]).upper() if len(parts) >= 2 else name[:3].upper()
            st.caption("Reference prefix: **{}**".format(pfx))

        phone = st.text_input("📞 Phone Number *",
                              placeholder="+966 xxxxxxxxx", key="rd3_phone")
        email = st.text_input("✉️ Email *",
                              placeholder="xxxx@socotec.com", key="rd3_email")
        phase  = st.selectbox("Phase / Speciality",
                              ["Waterproofing", "Senior", "Mid-Level", "Junior"],
                              key="rd3_phase")
        degree = st.text_input("Degree", value="Bachelor", key="rd3_degree")
        spec   = st.text_input("Speciality", value="Civil Engineer", key="rd3_spec")

    with c2:
        st.markdown("**Reviewer & Manager**")
        reviewer_options = [''] + TEAM_NAMES
        reviewer = st.selectbox("Reviewer (Engineer in Charge) *",
                                options=reviewer_options, index=0, key="rd3_reviewer")

        reviewer_custom = st.text_input("Or type reviewer name manually",
                                        placeholder="Leave blank if selected above",
                                        key="rd3_reviewer_custom")
        final_reviewer = reviewer_custom.strip() if reviewer_custom.strip() else reviewer

        manager = st.text_input("Manager / Head of Department",
                                value="Nizar Lazreg", key="rd3_manager")

        st.markdown("---")
        issue_dt  = st.date_input("Report Issue Date", value=date.today(), key="rd3_issue_dt")
        issue_str = "{}/{}/{}".format(issue_dt.day, issue_dt.month, issue_dt.year)
        city      = st.text_input("Report City", value="Riyadh", key="rd3_city")

        st.markdown("---")
        st.markdown("**Engineer Signature Image** *(optional)*")
        sig_file = st.file_uploader("Upload PNG/JPG signature (transparent preferred)",
                                     type=["png", "jpg", "jpeg"], key="rd3_sig_upload")
        if sig_file:
            st.image(sig_file, width=150, caption="Signature preview")
            st.session_state['_rd3_sig_bytes'] = sig_file.read()
        elif '_rd3_sig_bytes' not in st.session_state:
            st.session_state['_rd3_sig_bytes'] = None

    ready = bool(name and name.strip() and phone.strip() and email.strip() and final_reviewer)
    if not ready:
        st.info("Fill all required fields (*) including Reviewer to continue.")
    if st.button("Next →", type="primary", disabled=not ready, key="rd3_s1_next"):
        st.session_state.rd3_data.update({
            'eng_full':       name.strip(),
            'eng_phase':      phase,
            'eng_degree':     degree,
            'eng_speciality': spec,
            'eng_phone':      phone.strip(),
            'eng_email':      email.strip(),
            'reviewer_name':  final_reviewer,
            'manager_name':   manager.strip() or 'Nizar Lazreg',
            'issue_date':     issue_str,
            'city':           city.strip() or 'Riyadh',
        })
        st.session_state.rd3_step = 2
        st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Policy Upload
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 2:
    st.markdown('<div class="step-title">Step 2 — Policy Upload</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown("**Upload Malath or Tawuniya Policy PDF**")
        pdf_file = st.file_uploader("Policy PDF", type=["pdf"], key="rd3_policy_pdf")
        if pdf_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                tmp.write(pdf_file.read())
                tmp_pdf = tmp.name
            try:
                extracted = extract_from_policy_pdf(tmp_pdf)
                st.session_state.rd3_data.update(extracted)
                st.success("✅ Extracted: **{}** policy · IDI No: **{}**".format(
                    extracted.get('ins_type','?'), extracted.get('idi_no','?')))
                ins_type = extracted.get('ins_type', 'Malath')

                # If Malath and IDI no found → look up Excel
                idi_no = extracted.get('idi_no', '')
                if idi_no and EXCEL.exists():
                    xl_data = lookup_from_excel(str(EXCEL), idi_no)
                    if xl_data:
                        st.session_state.rd3_data.update(xl_data)
                        st.success("✅ Pulled extra fields from malath_log.xlsx")
                    else:
                        st.info("IDI No. **{}** not found in malath_log.xlsx — fill fields manually in Step 3.".format(idi_no))
            except Exception as e:
                st.error("PDF extraction failed: {}".format(e))
            finally:
                try: os.unlink(tmp_pdf)
                except: pass

    with c2:
        st.markdown("**Tawuniya Visit ID**")
        taw_visit_id = st.text_input(
            "Tawuniya Visit ID *",
            value=st.session_state.rd3_data.get('taw_visit_id', ''),
            placeholder="e.g. 1153490",
            help="Found on the Tawuniya inspection portal",
            key="rd3_taw_visit_id"
        )
        st.markdown("---")
        st.info("💡 If the policy PDF is not available, skip it and fill all fields manually in Step 3.")

        # Manual IDI / Tawuniya policy override
        st.markdown("**Manual Reference Override**")
        idi_manual = st.text_input("IDI / Reference No.",
                                   value=st.session_state.rd3_data.get('idi_no',''),
                                   key="rd3_idi_manual")
        taw_manual = st.text_input("Tawuniya Policy No. (TAW_POL)",
                                   value=st.session_state.rd3_data.get('taw_pol',''),
                                   key="rd3_taw_manual")
        nt_ft_manual = st.selectbox("NT / FT",
                                    options=['NT','FT'],
                                    index=0 if st.session_state.rd3_data.get('nt_ft','NT')=='NT' else 1,
                                    key="rd3_nt_ft")
        ins_type_manual = st.selectbox("Insurance Type",
                                       options=['Malath','Tawuniya'],
                                       index=0 if st.session_state.rd3_data.get('ins_type','Malath')=='Malath' else 1,
                                       key="rd3_ins_type")

    c1b, c2b = st.columns(2)
    with c1b:
        if st.button("← Back", key="rd3_s2_back"):
            st.session_state.rd3_step = 1; st.rerun()
    with c2b:
        if st.button("Next →", type="primary", key="rd3_s2_next"):
            if idi_manual:  st.session_state.rd3_data['idi_no']   = idi_manual
            if taw_manual:  st.session_state.rd3_data['taw_pol']  = taw_manual
            st.session_state.rd3_data['nt_ft']       = nt_ft_manual
            st.session_state.rd3_data['ins_type']    = ins_type_manual
            st.session_state.rd3_data['taw_visit_id']= taw_visit_id.strip()
            # Auto-generate reference
            rd3_ref, short_ref = build_rd3_reference(
                st.session_state.rd3_data.get('eng_full',''),
                st.session_state.rd3_data.get('nt_ft','NT'),
                st.session_state.rd3_data.get('idi_no',''),
                st.session_state.rd3_data.get('taw_pol',''),
                st.session_state.rd3_data.get('ins_type','Malath'),
            )
            st.session_state.rd3_data['rd3_ref']   = rd3_ref
            st.session_state.rd3_data['short_ref'] = short_ref
            # Pre-populate visits from Excel if available
            if not st.session_state.rd3_visits and st.session_state.rd3_data.get('visits'):
                st.session_state.rd3_visits = list(st.session_state.rd3_data['visits'])
            st.session_state.rd3_step = 3
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Project Info Review
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 3:
    st.markdown('<div class="step-title">Step 3 — Project Information</div>', unsafe_allow_html=True)
    d = st.session_state.rd3_data

    # Recover engineer name if lost
    if not d.get('eng_full','').strip():
        backup = st.session_state.get('_rd3_eng_backup', '')
        if backup:
            d['eng_full'] = backup

    tab1, tab2, tab3 = st.tabs(["📋 Core Info", "📅 Dates & Reference", "🔖 Document Ref"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            d['project_title'] = st.text_input("Project Title (Name)", value=d.get('project_title',''), key="rd3_proj_title")
            d['owner']         = st.text_input("Principal / Owner",     value=d.get('owner',''),         key="rd3_owner")
            d['address']       = st.text_area("Address of Premises",    value=d.get('address',''),       height=80, key="rd3_address")
        with c2:
            d['buildings']     = st.text_input("Buildings & Use",
                                               value=d.get('buildings','1 residential building'),
                                               key="rd3_buildings")
            d['building_type'] = st.selectbox("Building Type",
                                              ["Residential","Commercial","Mixed"],
                                              index=["Residential","Commercial","Mixed"].index(
                                                  d.get('building_type','Residential')
                                                  if d.get('building_type','Residential') in ["Residential","Commercial","Mixed"]
                                                  else "Residential"),
                                              key="rd3_btype")
            d['rd0_ref']       = st.text_input("Reference RD0",         value=d.get('rd0_ref',''),       key="rd3_rd0_ref")

    with tab2:
        c1, c2 = st.columns(2)
        with c1:
            d['occ_date']   = st.text_input("Occupancy Certificate Date (d/m/yyyy)",
                                             value=d.get('occ_date',''), key="rd3_occ_date")
            d['occ_status'] = st.radio("Occupancy Status", ["Expected", "Confirmed"],
                                        horizontal=True, key="rd3_occ_status")
        with c2:
            d['taw_visit_id'] = st.text_input("Tawuniya Visit ID",
                                               value=d.get('taw_visit_id',''), key="rd3_taw_id_s3")
            d['idi_no']     = st.text_input("IDI / Reference No.",
                                             value=d.get('idi_no',''), key="rd3_idi_s3")
            d['taw_pol']    = st.text_input("Tawuniya Policy No.",
                                             value=d.get('taw_pol',''), key="rd3_taw_s3")

    with tab3:
        c1, c2 = st.columns(2)
        with c1:
            # Always auto-compute short_ref — never let user corrupt it
            auto_ref, auto_short = build_rd3_reference(
                d.get('eng_full',''), d.get('nt_ft','NT'),
                d.get('idi_no',''), d.get('taw_pol',''),
                d.get('ins_type','Malath')
            )
            d['rd3_ref']   = st.text_input("RD3 Document Reference",
                                            value=d.get('rd3_ref', auto_ref), key="rd3_ref_edit")
            # short_ref is ALWAYS derived from the full ref — not user-editable
            d['short_ref'] = auto_short
            st.info("**Short Ref (Ref.: line):** `{}`".format(auto_short))
        with c2:
            st.info(
                "**Full ref**: `{}`\n\n"
                "**Body line** `Ref.:  {}`\n\n"
                "**File name**: `{}.docx`".format(
                    d.get('rd3_ref', auto_ref), auto_short, d.get('rd3_ref', auto_ref)
                )
            )

    st.markdown("---")
    # Defects and conclusion — default N/A, required
    c1, c2 = st.columns(2)
    with c1:
        d['defects_text']      = st.text_area("III – Defects / Disorders *",
                                               value=d.get('defects_text','None'),
                                               height=70, key="rd3_defects",
                                               help="Write 'None' if no defects")
        d['reservations_text'] = st.text_area("IV – Technical Reservations Not Closed *",
                                               value=d.get('reservations_text','None'),
                                               height=70, key="rd3_reservations",
                                               help="Write 'None' if no open reservations")
    with c2:
        # Auto-fill last visit date into conclusion text
        last_date = st.session_state.rd3_visits[-1].get('date','___') if st.session_state.rd3_visits else '___'
        default_conc = 'At the final visit done on {}, no defects in the waterproofing were noticed or observed.'.format(last_date)
        d['conclusion_text'] = st.text_area(
            "IV – Final Conclusion Text *",
            value=d.get('conclusion_text', default_conc),
            height=80, key="rd3_conclusion"
        )
        d['conclusion_yn'] = st.radio("Conclusion: Works adapted to project?",
                                       ["YES", "NO"], horizontal=True, key="rd3_conc_yn")

    # Validate required fields
    s3_errors = []
    if not d.get('project_title','').strip(): s3_errors.append("Project Title")
    if not d.get('owner','').strip():         s3_errors.append("Owner")
    if not d.get('address','').strip():       s3_errors.append("Address")
    if not d.get('occ_date','').strip():      s3_errors.append("Occupancy Date")
    if not d.get('defects_text','').strip():  s3_errors.append("Defects text (write 'None')")
    if not d.get('reservations_text','').strip(): s3_errors.append("Reservations (write 'None')")
    if not d.get('conclusion_text','').strip(): s3_errors.append("Conclusion Text")

    if s3_errors:
        st.warning("⚠️ Required: " + ", ".join(s3_errors))

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back", key="rd3_s3_back"): st.session_state.rd3_step = 2; st.rerun()
    with c2:
        if st.button("Next →", type="primary", disabled=bool(s3_errors), key="rd3_s3_next"):
            st.session_state.rd3_data = d
            st.session_state.rd3_step = 4; st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Site Visits
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 4:
    st.markdown('<div class="step-title">Step 4 — Site Visits</div>', unsafe_allow_html=True)
    st.info("Add all waterproofing inspection visits. Reference format: e.g. **YYO-V01-NT358273-750580**")

    eng = st.session_state.rd3_data.get('eng_full','')
    visits = list(st.session_state.rd3_visits)

    hc = st.columns([3, 2, 3, 3, 1])
    for col, lbl in zip(hc, ["Visit Reference", "Date (d/m/yyyy)", "Site Inspector", "Part Inspected", ""]):
        col.markdown("**{}**".format(lbl))

    updated, to_del = [], []
    for i, v in enumerate(visits):
        c1,c2,c3,c4,c5 = st.columns([3,2,3,3,1])
        ref = c1.text_input("", v.get('ref',''),       key="rd3_vr{}".format(i), label_visibility="collapsed")
        dat = c2.text_input("", v.get('date',''),      key="rd3_vd{}".format(i), label_visibility="collapsed")
        isp = c3.text_input("", v.get('inspector',''), key="rd3_vi{}".format(i), label_visibility="collapsed")
        prt = c4.text_input("", v.get('part',''),      key="rd3_vp{}".format(i), label_visibility="collapsed")
        if c5.button("✕", key="rd3_vx{}".format(i)):
            to_del.append(i)
        else:
            updated.append({'ref':ref,'date':dat,'inspector':isp,'part':prt})

    st.session_state.rd3_visits = [v for i,v in enumerate(updated) if i not in to_del]

    if len(st.session_state.rd3_visits) < 10:
        if st.button("➕ Add Row", key="rd3_add_visit"):
            st.session_state.rd3_visits.append({'ref':'','date':'','inspector':eng,'part':'Waterproofing'})
            st.rerun()
    else:
        st.caption("Maximum 10 visit rows")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back", key="rd3_s4_back"): st.session_state.rd3_step = 3; st.rerun()
    with c2:
        if st.button("Next →", type="primary", key="rd3_s4_next"):
            st.session_state.rd3_step = 5; st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — WP Works & Annex Details
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 5:
    st.markdown('<div class="step-title">Step 5 — Waterproofing Works & Annex Details</div>',
                unsafe_allow_html=True)

    # Which annexes are applicable?
    st.markdown("**Select applicable waterproofing works:**")
    col_r, col_f, col_b = st.columns(3)
    has_roofs     = col_r.checkbox("🏠 Roofs (Annex 1)",     value=True,  key="rd3_has_roofs")
    has_facades   = col_f.checkbox("🧱 Façades (Annex 2)",   value=False, key="rd3_has_facades")
    has_basements = col_b.checkbox("🏚️ Basements (Annex 3)", value=False, key="rd3_has_basements")

    st.markdown("---")
    annex_data = {}

    # Description template library
    ROOF_TEMPLATES = {
        "Bituminous membrane — residential roof + wet areas": {
            "i1": "The waterproofing had been executed to the roof top and the uncovered terraces",
            "i2": "The waterproofing system consists of a bituminous primer and two layers of 4mm bituminous membranes (SABIT), protected by a sand bed with ceramic tiles.",
            "i3": "The waterproofing covered 7 cm height as plinths above the final level at the junctions with walls and parapets.",
            "other": "Wet areas (Kitchen and toilets)", "types": ["ROOF", "OTHER"],
        },
        "Bituminous membrane — single layer": {
            "i1": "The waterproofing had been executed to the roof top and the uncovered terraces",
            "i2": "The waterproofing system consists of a bituminous primer and one layer of 4mm bituminous membranes (SABIT), protected by a sand bed with ceramic tiles.",
            "i3": "The waterproofing covered 7 cm height as plinths above the final level at the junctions with walls and parapets.",
            "other": "Wet areas (Kitchen and toilets)", "types": ["ROOF", "OTHER"],
        },
        "Cement-based waterproofing (Fosam)": {
            "i1": "The waterproofing had been executed to the roof top and the uncovered terraces",
            "i2": "The waterproofing materials consists of two cement based waterproofing (Fosam) layers protected by one bed of sand with ceramic tiles.",
            "i3": "The waterproofing covered 7 cm height as plinths above the final level at the junctions with walls and parapets.",
            "other": "Wet areas (Kitchen and toilets)", "types": ["ROOF", "OTHER"],
        },
        "Metal panel roof (industrial/commercial)": {
            "i1": "The roof consists of sloped metal panels with rockwool layer laid on metal ribs fixed to a metal structure supported by metal columns.",
            "i2": "The waterproofing system consists of sloped metal panels with rockwool layer fixed by screws with silicone around the screws.",
            "i3": "The junctions are treated with metal flashing and silicone sealant around all penetrations and edges.",
            "other": "", "types": ["ROOF"],
        },
        "Custom (fill manually)": {"i1":"","i2":"","i3":"","other":"","types":[]},
    }
    FACADE_TEMPLATES = {
        "Block walls + aluminum windows (standard residential)": {
            "i1": "The facades consist of plastered painted block walls facades and aluminum glazed windows and openings.",
            "i2": "Aluminium glazed windows are treated with silicone at the junctions with the walls.",
            "i3": "Aluminium glazed windows are treated with silicone at the junctions with the walls.",
            "other": "Stone cladding on the main façade only, with aluminum glazed windows and openings",
        },
        "Block walls + metal sheets (industrial)": {
            "i1": "The facades consist of block walls with paint and metal sheets on the upper part fixed on metal ribs and columns with Metal opening (Doors).",
            "i2": "The Facades metal sheets cover the painted block walls on the upper level of the facades.",
            "i3": "Metal flashing is used at the junction between metal panels and block walls.",
            "other": "Upper part with Metal sheets fixed on metal ribs and columns with Metal opening (Doors)",
        },
        "Plastered block walls only": {
            "i1": "The facades consist of plastered and painted block walls.",
            "i2": "The waterproofing system consists of waterproofing paint applied to the exterior facades.",
            "i3": "The junctions with the roof and ground level are sealed with flexible sealant.",
            "other": "",
        },
        "Custom (fill manually)": {"i1":"","i2":"","i3":"","other":""},
    }

    # ── ANNEX 1 — ROOFS ────────────────────────────────────────────────────────
    if has_roofs:
        with st.expander("📋 Annex 1 — Roofs", expanded=True):
            # ── Template selector ─────────────────────────────────────────────
            st.markdown("**🏗️ Quick-fill: Select standard description**")
            tpl_r = st.selectbox("Roof description template:", options=list(ROOF_TEMPLATES.keys()),
                                  key="rd3_r_tpl")
            tpl_r_data = ROOF_TEMPLATES[tpl_r]

            if st.button("✅ Apply template", key="rd3_r_apply_tpl"):
                for key, val in [('rd3_r_desc_i1', tpl_r_data['i1']),
                                  ('rd3_r_desc_i2', tpl_r_data['i2']),
                                  ('rd3_r_desc_i3', tpl_r_data['i3']),
                                  ('rd3_rt_other_text', tpl_r_data['other'])]:
                    st.session_state[key] = val
                if tpl_r_data.get('types'):
                    for k in ['rd3_rt_roof','rd3_rt_rtt','rd3_rt_it','rd3_rt_pat','rd3_rt_other']:
                        st.session_state[k] = False
                    if 'ROOF'  in tpl_r_data['types']: st.session_state['rd3_rt_roof'] = True
                    if 'OTHER' in tpl_r_data['types']: st.session_state['rd3_rt_other'] = True
                st.rerun()

            st.markdown("---")
            st.markdown("**Type of Roof(s)** — select all that apply:")
            col1, col2, col3 = st.columns(3)
            col4, col5       = st.columns([1, 2])
            rt_roof  = col1.checkbox("☐ ROOF",                key="rd3_rt_roof")
            rt_rtt   = col2.checkbox("☐ ROOFTOP TERRACE",     key="rd3_rt_rtt")
            rt_it    = col3.checkbox("☐ INTERMEDIATE TERRACE",key="rd3_rt_it")
            rt_pat   = col4.checkbox("☐ PATIOS",              key="rd3_rt_pat")
            rt_other = col5.checkbox("☐ OTHER",               key="rd3_rt_other")
            roof_types = []
            if rt_roof:  roof_types.append('ROOF')
            if rt_rtt:   roof_types.append('ROOFTOP TERRACE')
            if rt_it:    roof_types.append('INTERMEDIATE TERRACE')
            if rt_pat:   roof_types.append('PATIOS')
            if rt_other: roof_types.append('OTHER')
            other_type_text = ''
            if rt_other:
                other_type_text = st.text_input("Other roof type description:",
                                                 value=st.session_state.get('rd3_rt_other_text',''),
                                                 key="rd3_rt_other_text")

            st.markdown("**Works Description**")
            c1, c2 = st.columns(2)
            desc_i1 = c1.text_area("I.1. Describe the roof (type, materials, layers, slope, location):",
                                    value=st.session_state.get('rd3_r_desc_i1',''),
                                    height=100, key="rd3_r_desc_i1")
            desc_i2 = c2.text_area("I.2. Describe the waterproofing system layers (material, manufacturer, thickness):",
                                    value=st.session_state.get('rd3_r_desc_i2',''),
                                    height=100, key="rd3_r_desc_i2")
            desc_i3 = c1.text_area("I.3. Describe junctions (façade, vertical surfaces, etc.):",
                                    value=st.session_state.get('rd3_r_desc_i3',''),
                                    height=80, key="rd3_r_desc_i3")
            innovative_yn = c2.radio("I.4. Does WP include innovative technique/materials?",
                                      ["NO", "YES"], horizontal=True, key="rd3_r_inn_yn")

            st.markdown("**Materials**")
            c1, c2, c3 = st.columns(3)
            delivery_yn  = c1.radio("Delivery orders / certificates available?", ["YES","NO"], key="rd3_r_del_yn")
            compliant_yn = c2.radio("Materials compliant with design?",           ["YES","NO"], key="rd3_r_comp_yn")
            ponding_yn   = c3.radio("Water ponding test report available?",       ["YES","NO"], key="rd3_r_pond_yn")

            st.markdown("**Reviewed Documents** (leave blank if not applicable)")
            c1, c2 = st.columns(2)
            docs_mat = [
                c1.text_input("Material Doc 1:", key="rd3_r_mdc1"),
                c1.text_input("Material Doc 2:", key="rd3_r_mdc2"),
                c2.text_input("Material Doc 3:", key="rd3_r_mdc3"),
                c2.text_input("Material Doc 4:", key="rd3_r_mdc4"),
            ]
            docs_mat = [d for d in docs_mat if d.strip()]
            docs_test = [
                c1.text_input("Test Doc 1:", key="rd3_r_tdc1"),
                c1.text_input("Test Doc 2:", key="rd3_r_tdc2"),
                c2.text_input("Test Doc 3:", key="rd3_r_tdc3"),
                c2.text_input("Test Doc 4:", key="rd3_r_tdc4"),
            ]
            docs_test = [d for d in docs_test if d.strip()]

            last_visit_date = (st.session_state.rd3_visits[-1].get('date','')
                               if st.session_state.rd3_visits else '')
            default_conc = ('At the final visit done on {}, no defects in the waterproofing '
                            'were noticed or observed.'.format(last_visit_date))

            st.markdown("---")
            st.markdown("**Sections III / IV / V — Modifications, Defects, Reservations**")
            c1, c2, c3 = st.columns(3)
            mods_yn_r    = c1.radio("III. Any modifications from RD0?", ["NO","YES"], horizontal=True, key="rd3_r_mods_yn")
            def_yn_r     = c2.radio("IV. Defects / water infiltration observed?", ["NO","YES"], horizontal=True, key="rd3_r_def_yn")
            res_closed_r = c3.radio("V. Reservations closed?", ["NO","YES"], horizontal=True, key="rd3_r_res_closed")
            def_ref_r, def_brief_r, res_open_r = '', '', 'NO'
            if def_yn_r == 'YES':
                def_ref_r   = st.text_input("Reservation ref (if defect found):", key="rd3_r_def_ref")
                def_brief_r = st.text_input("Brief description of defect:", key="rd3_r_def_brief")
            res_open_r = st.radio("V. Any open (unclosed) reservations?", ["NO","YES"], horizontal=True, key="rd3_r_res_open")

            conc_yn_r   = st.radio("VI. Execution adapted to project?", ["YES","NO"],
                                    horizontal=True, key="rd3_r_conc_yn")
            conc_text_r = st.text_area("VI. Conclusion details:", value=default_conc,
                                        height=80, key="rd3_r_conc_text")

            annex_data['roofs'] = {
                'roof_types':              roof_types,
                'other_type_text':         other_type_text,
                'description_i1':          desc_i1,
                'description_i2':          desc_i2,
                'description_i3':          desc_i3,
                'innovative_yn':           innovative_yn,
                'delivery_yn':             delivery_yn,
                'compliant_yn':            compliant_yn,
                'ponding_yn':              ponding_yn,
                'reviewed_docs_materials': docs_mat,
                'reviewed_docs_tests':     docs_test,
                'reviewed_docs_other':     [],
                'modifications_yn':        mods_yn_r,
                'defects_observed_yn':     def_yn_r,
                'defects_ref':             def_ref_r,
                'defects_brief':           def_brief_r,
                'reservations_closed_yn':  res_closed_r,
                'reservations_open_yn':    res_open_r,
                'conclusion_yn':           conc_yn_r,
                'conclusion_text':         conc_text_r,
            }

    # ── ANNEX 2 — FAÇADES ─────────────────────────────────────────────────────
    if has_facades:
        with st.expander("📋 Annex 2 — Façades", expanded=True):
            # Template selector
            st.markdown("**🏗️ Quick-fill: Select standard description**")
            tpl_f = st.selectbox("Façade description template:", options=list(FACADE_TEMPLATES.keys()),
                                  key="rd3_f_tpl")
            tpl_f_data = FACADE_TEMPLATES[tpl_f]
            if st.button("✅ Apply template", key="rd3_f_apply_tpl"):
                for key, val in [('rd3_f_desc_i1', tpl_f_data['i1']),
                                  ('rd3_f_desc_i2', tpl_f_data['i2']),
                                  ('rd3_f_desc_i3', tpl_f_data['i3']),
                                  ('rd3_ft_other_text', tpl_f_data.get('other',''))]:
                    st.session_state[key] = val
                st.rerun()

            st.markdown("---")
            st.markdown("**Type of Façade(s)** — select all that apply:")
            col1, col2, col3 = st.columns(3)
            col4, _          = st.columns([1, 2])
            ft_conc   = col1.checkbox("☐ CONCRETE OR MASONRY", key="rd3_ft_conc")
            ft_clad   = col2.checkbox("☐ CLADDING",            key="rd3_ft_clad")
            ft_cw     = col3.checkbox("☐ CURTAIN WALL",        key="rd3_ft_cw")
            ft_other  = col4.checkbox("☐ OTHER",               key="rd3_ft_other")
            fac_types = []
            if ft_conc:  fac_types.append('CONCRETE OR MASONRY')
            if ft_clad:  fac_types.append('CLADDING')
            if ft_cw:    fac_types.append('CURTAIN WALL')
            if ft_other: fac_types.append('OTHER')
            fac_other_text = ''
            if ft_other:
                fac_other_text = st.text_input("Other façade description:", key="rd3_ft_other_text")

            st.markdown("**Works Description**")
            c1, c2 = st.columns(2)
            fdesc_i1 = c1.text_area("I.1. Describe the façade (type, materials, layers, location):",
                                     value=st.session_state.get('rd3_f_desc_i1','N/A'), height=90, key="rd3_f_desc_i1")
            fdesc_i2 = c2.text_area("I.2. Identify waterproofing junctions (location, material, manufacturer):",
                                     value=st.session_state.get('rd3_f_desc_i2','N/A'), height=90, key="rd3_f_desc_i2")
            fdesc_i3 = c1.text_area("I.3. Describe junctions with other elements (roofs, floors, balconies):",
                                     value=st.session_state.get('rd3_f_desc_i3','N/A'), height=90, key="rd3_f_desc_i3")

            st.markdown("**Materials**")
            c1, c2, c3 = st.columns(3)
            fdel_yn  = c1.radio("Delivery orders / certificates?", ["YES","NO"], key="rd3_f_del_yn")
            fcomp_yn = c2.radio("Materials compliant with design?", ["YES","NO"], key="rd3_f_comp_yn")
            fpond_yn = c3.radio("Ponding test report available?",   ["YES","NO"], key="rd3_f_pond_yn")

            c1, c2 = st.columns(2)
            fdocs_mat = [
                c1.text_input("Material Doc 1:", key="rd3_f_mdc1"),
                c1.text_input("Material Doc 2:", key="rd3_f_mdc2"),
                c2.text_input("Material Doc 3:", key="rd3_f_mdc3"),
                c2.text_input("Material Doc 4:", key="rd3_f_mdc4"),
            ]
            fdocs_mat = [d for d in fdocs_mat if d.strip()]
            fdocs_test = [
                c1.text_input("Test Doc 1:", key="rd3_f_tdc1"),
                c1.text_input("Test Doc 2:", key="rd3_f_tdc2"),
                c2.text_input("Test Doc 3:", key="rd3_f_tdc3"),
                c2.text_input("Test Doc 4:", key="rd3_f_tdc4"),
            ]
            fdocs_test = [d for d in fdocs_test if d.strip()]

            last_visit_date = (st.session_state.rd3_visits[-1].get('date','')
                               if st.session_state.rd3_visits else '')
            default_conc_f = ('At the final visit done on {}, no defects in the waterproofing '
                              'were noticed or observed.'.format(last_visit_date))

            st.markdown("---")
            st.markdown("**Sections III / IV / V — Modifications, Defects, Reservations**")
            c1, c2, c3 = st.columns(3)
            fmods_yn    = c1.radio("III. Any modifications from RD0?", ["NO","YES"], horizontal=True, key="rd3_f_mods_yn")
            fdef_yn     = c2.radio("IV. Defects observed?",             ["NO","YES"], horizontal=True, key="rd3_f_def_yn")
            fres_closed = c3.radio("V. Reservations closed?",           ["NO","YES"], horizontal=True, key="rd3_f_res_closed")
            fdef_ref, fdef_brief, fres_open = '', '', 'NO'
            if fdef_yn == 'YES':
                fdef_ref   = st.text_input("Reservation ref:", key="rd3_f_def_ref")
                fdef_brief = st.text_input("Brief description:", key="rd3_f_def_brief")
            fres_open = st.radio("V. Any open reservations?", ["NO","YES"], horizontal=True, key="rd3_f_res_open")

            fconc_yn   = st.radio("VI. Execution adapted to project?", ["YES","NO"],
                                   horizontal=True, key="rd3_f_conc_yn")
            fconc_text = st.text_area("VI. Conclusion details:", value=default_conc_f,
                                       height=80, key="rd3_f_conc_text")

            annex_data['facades'] = {
                'facade_types':            fac_types,
                'other_type_text':         fac_other_text,
                'description_i1':          fdesc_i1,
                'description_i2':          fdesc_i2,
                'description_i3':          fdesc_i3,
                'delivery_yn':             fdel_yn,
                'compliant_yn':            fcomp_yn,
                'ponding_yn':              fpond_yn,
                'reviewed_docs_materials': fdocs_mat,
                'reviewed_docs_tests':     fdocs_test,
                'reviewed_docs_other':     [],
                'modifications_yn':        fmods_yn,
                'defects_observed_yn':     fdef_yn,
                'defects_ref':             fdef_ref,
                'defects_brief':           fdef_brief,
                'reservations_closed_yn':  fres_closed,
                'reservations_open_yn':    fres_open,
                'conclusion_yn':           fconc_yn,
                'conclusion_text':         fconc_text,
            }

    # ── ANNEX 3 — BASEMENTS ────────────────────────────────────────────────────
    if has_basements:
        with st.expander("📋 Annex 3 — Basements", expanded=True):
            st.markdown("**Works Description**")
            c1, c2 = st.columns(2)
            bdesc_i1 = c1.text_area("I.1. Describe the basement (type, materials, layers, location):",
                                     height=100, key="rd3_b_desc_i1")
            bdesc_i2 = c2.text_area("I.2. Describe the waterproofing layers:",
                                     height=100, key="rd3_b_desc_i2")

            st.markdown("**Materials**")
            c1, c2, c3 = st.columns(3)
            bdel_yn  = c1.radio("Delivery orders / certificates?", ["YES","NO"], key="rd3_b_del_yn")
            bcomp_yn = c2.radio("Materials compliant with design?", ["YES","NO"], key="rd3_b_comp_yn")
            bpond_yn = c3.radio("Ponding test report available?",   ["YES","NO"], key="rd3_b_pond_yn")

            c1, c2 = st.columns(2)
            bdocs_mat = [
                c1.text_input("Material Doc 1:", key="rd3_b_mdc1"),
                c1.text_input("Material Doc 2:", key="rd3_b_mdc2"),
                c2.text_input("Material Doc 3:", key="rd3_b_mdc3"),
                c2.text_input("Material Doc 4:", key="rd3_b_mdc4"),
            ]
            bdocs_mat = [d for d in bdocs_mat if d.strip()]
            bdocs_test = [
                c1.text_input("Test Doc 1:", key="rd3_b_tdc1"),
                c1.text_input("Test Doc 2:", key="rd3_b_tdc2"),
                c2.text_input("Test Doc 3:", key="rd3_b_tdc3"),
                c2.text_input("Test Doc 4:", key="rd3_b_tdc4"),
            ]
            bdocs_test = [d for d in bdocs_test if d.strip()]

            last_visit_date = (st.session_state.rd3_visits[-1].get('date','')
                               if st.session_state.rd3_visits else '')
            default_conc_b = ('At the final visit done on {}, no defects in the waterproofing '
                              'were noticed or observed.'.format(last_visit_date))

            st.markdown("---")
            st.markdown("**Sections III / IV / V — Modifications, Defects, Reservations**")
            c1, c2, c3 = st.columns(3)
            bmods_yn    = c1.radio("III. Any modifications from RD0?", ["NO","YES"], horizontal=True, key="rd3_b_mods_yn")
            bdef_yn     = c2.radio("IV. Defects observed?",             ["NO","YES"], horizontal=True, key="rd3_b_def_yn")
            bres_closed = c3.radio("V. Reservations closed?",           ["NO","YES"], horizontal=True, key="rd3_b_res_closed")
            bdef_ref, bdef_brief, bres_open = '', '', 'NO'
            if bdef_yn == 'YES':
                bdef_ref   = st.text_input("Reservation ref:", key="rd3_b_def_ref")
                bdef_brief = st.text_input("Brief description:", key="rd3_b_def_brief")
            bres_open = st.radio("V. Any open reservations?", ["NO","YES"], horizontal=True, key="rd3_b_res_open")

            bconc_yn   = st.radio("VI. Execution adapted to project?", ["YES","NO"],
                                   horizontal=True, key="rd3_b_conc_yn")
            bconc_text = st.text_area("VI. Conclusion details:", value=default_conc_b,
                                       height=80, key="rd3_b_conc_text")

            annex_data['basements'] = {
                'description_i1':          bdesc_i1,
                'description_i2':          bdesc_i2,
                'delivery_yn':             bdel_yn,
                'compliant_yn':            bcomp_yn,
                'ponding_yn':              bpond_yn,
                'reviewed_docs_materials': bdocs_mat,
                'reviewed_docs_tests':     bdocs_test,
                'reviewed_docs_other':     [],
                'modifications_yn':        bmods_yn,
                'defects_observed_yn':     bdef_yn,
                'defects_ref':             bdef_ref,
                'defects_brief':           bdef_brief,
                'reservations_closed_yn':  bres_closed,
                'reservations_open_yn':    bres_open,
                'conclusion_yn':           bconc_yn,
                'conclusion_text':         bconc_text,
            }

    if not any([has_roofs, has_facades, has_basements]):
        st.warning("⚠️ Select at least one waterproofing work type to continue.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back", key="rd3_s5_back"): st.session_state.rd3_step = 4; st.rerun()
    with c2:
        ready5 = any([has_roofs, has_facades, has_basements])
        if st.button("Next →", type="primary", disabled=not ready5, key="rd3_s5_next"):
            st.session_state.rd3_annex_data = annex_data
            st.session_state.rd3_step = 6; st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — Generate
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 6:
    st.markdown('<div class="step-title">Step 6 — Generate RD3 Report</div>', unsafe_allow_html=True)
    d      = st.session_state.rd3_data
    visits = st.session_state.rd3_visits
    ann    = st.session_state.rd3_annex_data

    # ── Recover engineer name from backup if somehow lost ─────────────────
    if not d.get('eng_full','').strip():
        backup = st.session_state.get('_rd3_eng_backup', '')
        if backup:
            d['eng_full'] = backup
            st.session_state.rd3_data['eng_full'] = backup

    # Summary metrics
    c1,c2,c3 = st.columns(3)
    c1.metric("RD3 Reference",    d.get('rd3_ref','—'))
    c2.metric("Engineer",         d.get('eng_full','⚠️ MISSING'))
    c3.metric("Issue Date",       d.get('issue_date','—'))
    c1.metric("IDI No.",          d.get('idi_no','—') or '—')
    c2.metric("Tawuniya Policy",  d.get('taw_pol','—') or 'N/A')
    c3.metric("Site Visits",      len(visits))
    c1.metric("Annexes",          ", ".join(ann.keys()) or "None")
    c2.metric("Reviewer",         d.get('reviewer_name','—'))
    c3.metric("Tawuniya Visit ID",d.get('taw_visit_id','—') or '—')

    st.markdown("---")

    # Attachments
    st.markdown("**📎 Attachments** *(optional — added as pages at end of report)*")
    att_files = st.file_uploader("Upload images to attach (PNG, JPG, PDF)",
                                  type=["png","jpg","jpeg","pdf"],
                                  accept_multiple_files=True, key="rd3_attachments")
    if att_files:
        st.caption("{} attachment(s) will be added to the report.".format(len(att_files)))

    if not TPL.exists():
        st.error("Template not found: `{}`  —  Upload Template_RD3.docx to the repo root.".format(TPL))
        st.info("📁 Expected path: repo root → `Template_RD3.docx`")
    elif not d.get('eng_full','').strip():
        st.error("⚠️ Engineer name is missing — go back to Step 1 and select an engineer.")
    else:
        if st.button("🚀 Generate RD3 Report", type="primary", key="rd3_gen_btn"):
            with st.spinner("Building RD3 report…"):
                try:
                    # Fix conclusion date if it contains placeholder ___
                    last_v_date = visits[-1].get('date','') if visits else ''
                    if last_v_date:
                        if '___' in d.get('conclusion_text',''):
                            d['conclusion_text'] = d['conclusion_text'].replace('___', last_v_date)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                        out = tmp.name
                    # Prepare signature bytes
                    sig_b = st.session_state.get('_rd3_sig_bytes', None)
                    # Prepare attachments
                    attach_list = None
                    if att_files:
                        attach_list = [(f.name, f.read()) for f in att_files]
                    generate_rd3(
                        template_path = str(TPL),
                        output_path   = out,
                        data          = d,
                        visits        = visits,
                        annex_data    = ann,
                        sig_bytes     = sig_b,
                        attachments   = attach_list,
                    )
                    with open(out,'rb') as f:
                        docx_bytes = f.read()
                    os.unlink(out)
                    fname = "{}.docx".format(d.get('rd3_ref','RD3_Report'))
                    st.download_button(
                        "⬇️ Download {}".format(fname),
                        docx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="rd3_download"
                    )
                    st.success("✅ Report generated. Click above to download.")
                except Exception as e:
                    st.error("Generation failed: {}".format(e))
                    import traceback; st.code(traceback.format_exc())

    if st.button("← Back", key="rd3_s6_back"):
        st.session_state.rd3_step = 5; st.rerun()
