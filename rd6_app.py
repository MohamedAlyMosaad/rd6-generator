"""
rd6_app.py — RD6 Completion of Works Report Generator
Run: streamlit run rd6_app.py
"""
import os, io, tempfile
from pathlib import Path
from datetime import date
import streamlit as st
from rd6_extractor import (extract_from_policy_pdf, lookup_from_excel,
                            build_rd6_reference, extract_date_from_cert,
                            load_engineer_team)
from rd6_generator import generate_rd6, DOC_KEYS, STANDARD_MISSING_DOCS

BASE       = Path(__file__).parent
TPL        = BASE / "RD6-AutoTemplate.docx"
EXCEL      = BASE / "malath_log.xlsx"
TEAM_EXCEL = BASE / "IDI_Team.xlsx"

st.set_page_config(page_title="RD6 Generator · SOCOTEC Arabia",
                   page_icon="🏗️", layout="wide")

st.markdown("""
<style>
.step-title {
    font-size: 1.15rem; font-weight: 700; color: #1f4e79;
    border-left: 5px solid #2e75b6; padding-left: 10px; margin-bottom: 1rem;
}
</style>""", unsafe_allow_html=True)

# ── Load engineer team once ───────────────────────────────────────────────────
@st.cache_data
def get_team():
    if TEAM_EXCEL.exists():
        return load_engineer_team(str(TEAM_EXCEL))
    return {}

TEAM = get_team()
TEAM_NAMES = sorted(TEAM.keys())

# ── Session init ──────────────────────────────────────────────────────────────
for k, v in [('step', 1), ('data', {}), ('visits', []),
             ('sig_bytes', None), ('sig_ext', 'png'),
             ('ins_bytes', None), ('ins_name', '')]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # SOCOTEC Logo
    logo_path = BASE / "socotec_logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=160)
    else:
        st.markdown("## 🏗️ RD6 Generator")
    st.markdown("**SOCOTEC Arabia · TIS Division**")
    st.markdown("RD6 Report Generator")
    st.markdown("---")
    labels = ["Engineer & Signature", "Policy Upload", "Project Info",
              "Site Visits", "Requirements", "Generate"]
    cur = st.session_state.step
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
    if st.button("🔄 Restart"):
        for k in ['step','data','visits','sig_bytes','sig_ext','ins_bytes','ins_name']:
            st.session_state[k] = {'step':1,'data':{},'visits':[]}.get(
                k, None if ('bytes' in k or k=='ins_name') else ([] if k=='visits' else {}))
        st.session_state.step = 1
        st.rerun()

step = st.session_state.step

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Engineer Details & Signature
# ═══════════════════════════════════════════════════════════════════════════════
if step == 1:
    st.markdown('<div class="step-title">Step 1 — Engineer Details & Signature</div>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Engineer Info**")

        # Name — searchable dropdown from IDI_Team + free text fallback
        if TEAM_NAMES:
            options = [''] + TEAM_NAMES
            selected = st.selectbox(
                "Full Name *",
                options=options,
                index=0,
                help="Select your name — phone and email auto-fill from the IDI team list"
            )
            # Auto-fill fields when a name is selected
            if selected and selected in TEAM:
                info = TEAM[selected]
                default_phone = info['phone']
                default_email = info['email']
                default_phase = info['phase']
                default_degree = info['degree']
            else:
                default_phone = ''
                default_email = ''
                default_phase = 'Senior'
                default_degree = 'Civil Engineering Bachelor'
            name = selected
        else:
            name = st.text_input("Full Name (First Last) *", placeholder="Mohamed Mossad")
            default_phone = ''
            default_email = ''
            default_phase = 'Senior'
            default_degree = 'Civil Engineering Bachelor'

        if name:
            parts = name.strip().split()
            pfx = (parts[0][0] + parts[1][:2]).upper() if len(parts) >= 2 else name[:3].upper()
            st.caption("Reference prefix: **{}**".format(pfx))

        phone  = st.text_input("📞 Phone Number *", value=default_phone,
                               placeholder="+966 xxxxxxxxx")
        email  = st.text_input("✉️ Email *", value=default_email,
                               placeholder="xxxx@socotec.com")
        phase  = st.selectbox("Phase / Level", ["Senior", "Mid-Level", "Junior"],
                              index=["Senior","Mid-Level","Junior"].index(default_phase)
                              if default_phase in ["Senior","Mid-Level","Junior"] else 0)
        degree = st.text_input("Degree", value=default_degree)
        spec   = st.selectbox("Speciality",
                              ["Civil", "Structural", "Geotechnical", "Architecture", "MEP"])

    with c2:
        st.markdown("**Signature Image**")
        st.info("Upload a transparent PNG of your signature. It will appear under the Author column only.")
        sig_file = st.file_uploader("Signature image (PNG/JPG)",
                                     type=["png","jpg","jpeg"], key="sig_upload")
        if sig_file:
            sig_bytes = sig_file.read()
            st.image(sig_bytes, caption="Preview", width=250)
            sig_ext = Path(sig_file.name).suffix.lstrip('.').lower()
            st.session_state.sig_bytes = sig_bytes
            st.session_state.sig_ext   = sig_ext
            st.success("✅ Signature uploaded")
        elif st.session_state.sig_bytes:
            st.info("Signature already uploaded.")
        st.markdown("---")
        issue_dt  = st.date_input("Report Issue Date", value=date.today())
        issue_str = "{}/{}/{}".format(issue_dt.day, issue_dt.month, issue_dt.year)

    if not st.session_state.sig_bytes:
        st.warning("No signature uploaded — a blank placeholder will remain in the report.")

    ready = bool(name and name.strip() and phone.strip() and email.strip())
    if st.button("Next →", type="primary", disabled=not ready):
        st.session_state.data.update({
            'eng_full':    name.strip(),
            'eng_phase':   phase,
            'eng_degree':  degree,
            'eng_speciality': spec,
            'eng_phone':   phone.strip(),
            'eng_email':   email.strip(),
            'issue_date':  issue_str,
        })
        st.session_state.step = 2
        st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Policy Upload
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 2:
    st.markdown('<div class="step-title">Step 2 — Policy Upload & Data Extraction</div>',
                unsafe_allow_html=True)
    ins = st.radio("Insurance Company", ["Malath","Tawuniya"], horizontal=True)
    if ins == 'Tawuniya':
        st.info(
            "**Tawuniya policy:** Fields are extracted automatically. "
            "The policy number is **not printed inside the PDF** — it is the filename "
            "(e.g. file named `5xxxxxxxx.pdf` → policy number `5xxxxxxxx`). "
            "You will enter it in the next step."
        )
    pdf = st.file_uploader("Upload {} IDI Policy PDF *".format(ins), type=["pdf"])
    if not EXCEL.exists():
        st.warning("malath_log.xlsx not found — Excel lookup will be skipped.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back"):
            st.session_state.step = 1; st.rerun()
    with c2:
        if st.button("Extract & Continue →", type="primary", disabled=pdf is None):
            with st.spinner("Extracting from PDF…"):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(pdf.read()); tp = tmp.name
                pdf_data   = extract_from_policy_pdf(tp)
                os.unlink(tp)
                idi        = pdf_data.get('idi_no', '')
                excel_data = lookup_from_excel(str(EXCEL), idi) if idi and EXCEL.exists() else {}
                merged     = {**pdf_data, **{k:v for k,v in excel_data.items() if v}}
                merged['ins_type'] = ins
                merged['idi_no']   = idi
                eng = st.session_state.data.get('eng_full','')
                taw = merged.get('taw_pol','')
                merged['rd6_ref'] = build_rd6_reference(
                    eng, merged.get('nt_ft','NT'), idi, taw, ins_type=ins)
                if excel_data.get('visits'):
                    st.session_state.visits = excel_data['visits']
                st.session_state.data.update(merged)
            st.session_state.step = 3; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Project Info
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 3:
    st.markdown('<div class="step-title">Step 3 — Review & Edit Project Information</div>',
                unsafe_allow_html=True)
    st.info("All fields pre-filled from PDF/Excel. Review and correct anything before proceeding.")
    d        = st.session_state.data
    ins_type = d.get('ins_type','Malath')
    t1, t2, t3 = st.tabs(["📌 Core","📅 Dates","🔢 References"])

    with t1:
        c1, c2 = st.columns(2)
        with c1:
            d['project_title'] = st.text_input("Project Title",     value=d.get('project_title',''))
            d['owner']         = st.text_input("Owner",             value=d.get('owner',''))
            d['address']       = st.text_area("Address",            value=d.get('address',''), height=80)
            d['building_type'] = st.text_input("Building Type",     value=d.get('building_type','Residential'))
        with c2:
            d['sum_insured']   = st.text_input("Sum Insured (SR)",  value=d.get('sum_insured',''))
            d['no_buildings']  = st.text_input("No. of Buildings",  value=d.get('no_buildings','1'))
            if ins_type == 'Tawuniya':
                st.markdown("---")
                st.markdown("**🔑 Tawuniya Policy Number**")
                st.info("Enter the number from the **PDF filename** (e.g. `5xxxxxxxx`). "
                        "It is not printed inside the document.")
                d['taw_pol'] = st.text_input("Tawuniya Policy No. *", value=d.get('taw_pol',''))
                new_ref = build_rd6_reference(
                    d.get('eng_full',''), '', d.get('idi_no',''), d['taw_pol'], ins_type='Tawuniya')
                d['rd6_ref'] = new_ref
                st.caption("RD6 Reference: **{}**".format(new_ref))
            else:
                d['nt_ft']   = st.selectbox("NT / FT", ['NT','FT'],
                                             index=0 if d.get('nt_ft','NT')=='NT' else 1)
                d['taw_pol'] = st.text_input("Tawuniya Policy No.", value=d.get('taw_pol',''))
                new_ref = build_rd6_reference(
                    d.get('eng_full',''), d['nt_ft'], d.get('idi_no',''),
                    d['taw_pol'], ins_type='Malath')
                d['rd6_ref'] = new_ref
                st.caption("RD6 Reference: **{}**".format(new_ref))

    with t2:
        c1, c2 = st.columns(2)
        with c1:
            d['start_date']      = st.text_input("Works Start Date (d/m/yyyy)", value=d.get('start_date',''))
            d['finish_date']     = st.text_input("Works Completion Date",        value=d.get('finish_date',''))
            d['last_visit_date'] = st.text_input("Last Site Visit Date",         value=d.get('last_visit_date',''))
        with c2:
            d['occ_date']       = st.text_input("OCC Date",            value=d.get('occ_date',''))
            d['occ_confirmed']  = st.checkbox("OCC Confirmed",         value=d.get('occ_confirmed',False))
            d['roof_test_date'] = st.text_input("Roof Ponding Test Date",
                                                 value=d.get('roof_test_date',''),
                                                 help="Auto-filled from insulation cert in Step 5")

    with t3:
        c1, c2 = st.columns(2)
        with c1:
            d['rd0_ref']  = st.text_input("RD0 Reference *",          value=d.get('rd0_ref',''),
                                           help="e.g. HOS-RD0-NT358273-1")
        with c2:
            d['rd0_date'] = st.text_input("RD0 Issue Date * (d/m/yyyy)", value=d.get('rd0_date',''),
                                           help="Date the RD0 was issued — appears in Section IV")
        if not d.get('rd0_date',''):
            st.warning("RD0 Issue Date is empty. Enter it above — it appears in Section IV.")
        d['reservations_note'] = st.text_area("Reservations Note (optional)",
                                               value=d.get('reservations_note',''), height=60)

    if ins_type == 'Tawuniya' and not d.get('taw_pol','').strip():
        st.warning("Tawuniya Policy No. is required. Enter it in the Core tab.")

    st.session_state.data = d
    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back"):
            st.session_state.step = 2; st.rerun()
    with c2:
        if st.button("Next →", type="primary"):
            st.session_state.step = 4; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Site Visits
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 4:
    st.markdown('<div class="step-title">Step 4 — Site Visits</div>', unsafe_allow_html=True)
    st.info("Supports up to 10 visit rows. Pre-filled from Excel where available.")
    eng    = st.session_state.data.get('eng_full','')
    visits = list(st.session_state.visits)

    hc = st.columns([3,2,3,3,1])
    for col, lbl in zip(hc, ["Visit Reference","Date (d/m/yyyy)","Site Inspector","Inspected Part",""]):
        col.markdown("**{}**".format(lbl))

    updated, to_del = [], []
    for i, v in enumerate(visits):
        c1,c2,c3,c4,c5 = st.columns([3,2,3,3,1])
        ref = c1.text_input("Ref",       v.get('ref',''),       key="vr{}".format(i), label_visibility="collapsed")
        dat = c2.text_input("Date",      v.get('date',''),      key="vd{}".format(i), label_visibility="collapsed")
        isp = c3.text_input("Inspector", v.get('inspector',''), key="vi{}".format(i), label_visibility="collapsed")
        prt = c4.text_input("Part",      v.get('part',''),      key="vp{}".format(i), label_visibility="collapsed")
        if c5.button("✕", key="vx{}".format(i)):
            to_del.append(i)
        else:
            updated.append({'ref':ref,'date':dat,'inspector':isp,'part':prt})

    st.session_state.visits = [v for i,v in enumerate(updated) if i not in to_del]

    if len(st.session_state.visits) < 10:
        if st.button("➕ Add Row"):
            st.session_state.visits.append({'ref':'','date':'','inspector':eng,'part':''})
            st.rerun()
    else:
        st.caption("Maximum 10 visit rows")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back"):
            st.session_state.step = 3; st.rerun()
    with c2:
        if st.button("Next →", type="primary"):
            st.session_state.step = 5; st.rerun()
# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — Client Requirements
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 5:
    st.markdown('<div class="step-title">Step 5 — Client Requirements</div>', unsafe_allow_html=True)
    st.info("Upload documents the client provided. **Missing documents will be listed in the report.** "
            "The Insulation Certificate is mandatory and will be appended as final page(s).")

    slots = [
        ("insulation_cert",   "⭐ Waterproofing / Insulation Certificate",
         "Certified by Commerce Chamber — MANDATORY — appended to report", True),
        ("cost_letter",       "1. Cost Letter  (خطاب التكلفة)",
         "Contractor letter stating actual project cost", False),
        ("contractor_letter", "2. Contractor Letter  (خطاب المقاول)",
         "Letter confirming design remarks were incorporated", False),
        ("supervision_letter","3. Engineering Supervision Letter  (خطاب الإشراف الهندسي)",
         "Stamped certificate from supervising engineering office", False),
        ("calc_notes",        "4. Calculation Notes  (النوتة الحسابية)",
         "Approved structural design calculation notes", False),
        ("soil_tests",        "5. Soil / Compaction Tests  (اختبارات التربة)",
         "Compaction test results above and below foundations", False),
        ("concrete_tests",    "6. Concrete Strength Tests  (اختبارات الخرسانة)",
         "Compressive strength results from accredited lab", False),
        ("steel_invoices",    "7. Steel Invoices / Warranty  (فواتير الحديد)",
         "All rebar delivery invoices or SASO-certified warranty", False),
        ("material_warranty", "8. Material Warranty Certificates  (شهادات ضمان المواد)",
         "Warranty certs for structural / façade elements", False),
    ]

    provided  = []
    ins_bytes = None
    ins_name  = ''

    for key, label, sub, mandatory in slots:
        st.markdown("---")
        la, ra = st.columns([2,3])
        with la:
            color = "#b00" if mandatory else "inherit"
            st.markdown("**<span style='color:{}'>{}</span>**".format(color, label),
                        unsafe_allow_html=True)
            st.caption(sub)
        with ra:
            f = st.file_uploader(label, type=["pdf","jpg","jpeg","png"],
                                  key="req_{}".format(key), label_visibility="collapsed")
            if f:
                provided.append(key)
                if key == 'insulation_cert':
                    ins_bytes = f.read()
                    ins_name  = f.name
                    cert_date = extract_date_from_cert(ins_bytes, ins_name)
                    if cert_date and not st.session_state.data.get('roof_test_date',''):
                        st.session_state.data['roof_test_date'] = cert_date
                        st.caption("📅 Roof test date auto-filled: **{}**".format(cert_date))
                    elif cert_date:
                        st.caption("📅 Date found in cert: {}".format(cert_date))
                elif key in ['cost_letter', 'contractor_letter', 'supervision_letter']:
                    st.session_state[f'cert_bytes_{key}'] = f.read()
                    st.session_state[f'cert_ext_{key}'] = Path(f.name).suffix.lstrip('.').lower()
                st.success("✅ {}".format(f.name))
            elif not mandatory:
                st.caption("Not uploaded → listed as missing in report")

    st.session_state.data['provided_doc_keys'] = provided
    if ins_bytes:
        st.session_state.ins_bytes = ins_bytes
        st.session_state.ins_name  = ins_name

    has_ins = 'insulation_cert' in provided
    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Back"):
            st.session_state.step = 4; st.rerun()
    with c2:
        if not has_ins:
            st.warning("Upload the Insulation Certificate to continue.")
        if st.button("Next →", type="primary", disabled=not has_ins):
            st.session_state.step = 6; st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — Generate
# ═══════════════════════════════════════════════════════════════════════════════
elif step == 6:
    st.markdown('<div class="step-title">Step 6 — Generate RD6 Report</div>', unsafe_allow_html=True)
    d      = st.session_state.data
    visits = st.session_state.visits
    pkeys  = d.get('provided_doc_keys',[])
    miss   = sum(1 for k in DOC_KEYS if k not in pkeys)

    c1,c2,c3 = st.columns(3)
    c1.metric("RD6 Reference",     d.get('rd6_ref','—'))
    c2.metric("Engineer",          d.get('eng_full','—'))
    c3.metric("Issue Date",        d.get('issue_date','—'))
    c1.metric("IDI No.",           d.get('idi_no','—') or '—')
    c2.metric("Policy / Tawuniya", d.get('taw_pol','—') or 'N/A')
    c3.metric("Site Visits",       len(visits))
    c1.metric("Docs Provided",     len(pkeys))
    c2.metric("Missing in Report", miss)
    c3.metric("Signature",         "✅ Uploaded" if st.session_state.sig_bytes else "⚠️ None")

    st.markdown("---")

    if not TPL.exists():
        st.error("Template not found: {}. Place RD6-AutoTemplate.docx in the app folder.".format(TPL))
    else:
        if st.button("🚀 Generate Report", type="primary"):
            with st.spinner("Building report… (this may take 15–30 seconds for the final repack)"):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                        out = tmp.name
                    extra_cert_bytes = []
                    for ck in ['cost_letter', 'contractor_letter', 'supervision_letter']:
                        cb = st.session_state.get(f'cert_bytes_{ck}')
                        ce = st.session_state.get(f'cert_ext_{ck}', 'pdf')
                        if cb:
                            extra_cert_bytes.append((cb, ce))
                    st.write("DEBUG visits:", st.session_state.visits)
                    generate_rd6(
                        template_path      = str(TPL),
                        output_path        = out,
                        data               = d,
                        visits             = visits,
                        provided_doc_keys  = pkeys,
                        signature_bytes    = st.session_state.sig_bytes,
                        signature_ext      = st.session_state.sig_ext or 'png',
                        insulation_bytes   = st.session_state.ins_bytes,
                        insulation_filename= st.session_state.ins_name or 'cert.pdf',
                        extra_cert_bytes   = extra_cert_bytes,
                    )
                    with open(out,'rb') as f:
                        docx_bytes = f.read()
                    os.unlink(out)
                    fname = "{}.docx".format(d.get('rd6_ref','RD6_Report'))
                    st.download_button(
                        "⬇️ Download {}".format(fname), docx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                    st.success("✅ Report generated successfully. Click above to download.")
                    if miss:
                        st.markdown("**Documents listed as missing:**")
                        for k,(short,_) in zip(DOC_KEYS, STANDARD_MISSING_DOCS):
                            if k not in pkeys:
                                st.markdown("  - {}".format(short))
                except Exception as e:
                    st.error("Generation failed: {}".format(e))
                    import traceback; st.code(traceback.format_exc())

    if st.button("← Back"):
        st.session_state.step = 5; st.rerun()
