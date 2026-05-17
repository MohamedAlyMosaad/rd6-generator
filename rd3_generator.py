"""
rd3_generator.py — Fills Template_RD3.docx with project data.

Template structure:
  Table 0: Header (logo/company name)
  Table 1: Document reference + expert info (6 rows × 4 cols)
  Table 2: Main report body (12 rows × 3 cols)
  Table 3: Annex 1 – Roofs (13 rows × 1 col)
  Table 4: Annex 2 – Façades (13 rows × 1 col)
  Table 5: Annex 3 – Basements (12 rows × 1 col)

30 SDT checkboxes (no tags, position-based):
  CB00=ROOFS  CB01=FACADES  CB02=BASEMENTS
  CB03-07  = Annex1 roof type
  CB08-13  = Annex1 material Y/N
  CB14-17  = Annex2 facade type
  CB18-23  = Annex2 material Y/N
  CB24-29  = Annex3 material Y/N
"""
import io, os, shutil, tempfile
from pathlib import Path
from datetime import datetime
from lxml import etree
from docx import Document

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# ── Helpers ────────────────────────────────────────────────────────────────────

def _get_text(elem):
    return ''.join(t.text or '' for t in elem.iter(f'{{{W}}}t'))

def _set_para_text(para_elem, new_text):
    """Replace text in a paragraph element, preserving the first run's formatting."""
    t_elems = list(para_elem.iter(f'{{{W}}}t'))
    if not t_elems:
        # Create a bare run
        r = etree.SubElement(para_elem, f'{{{W}}}r')
        t = etree.SubElement(r, f'{{{W}}}t')
        t.text = new_text
        if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
            t.set(XML_SPACE, 'preserve')
    else:
        t_elems[0].text = new_text
        for t in t_elems[1:]:
            t.text = ''
        if new_text and (new_text[0] == ' ' or new_text[-1] == ' '):
            t_elems[0].set(XML_SPACE, 'preserve')


def _replace_global(body, old_text, new_text):
    """Replace all paragraphs whose full text contains old_text (global search)."""
    for p in body.iter(f'{{{W}}}p'):
        full = _get_text(p)
        if old_text in full:
            _set_para_text(p, full.replace(old_text, new_text))


def _toggle_sdt(body, index, checked):
    """Toggle the nth SDT checkbox (0-based) to ☒ (checked) or ☐ (unchecked)."""
    sdts = list(body.iter(f'{{{W}}}sdt'))
    if index >= len(sdts):
        return
    sdt = sdts[index]
    checked_elem = sdt.find(f'.//{{{W14}}}checked')
    if checked_elem is not None:
        checked_elem.set(f'{{{W14}}}val', '1' if checked else '0')
    content = sdt.find(f'{{{W}}}sdtContent')
    if content is not None:
        for t in content.iter(f'{{{W}}}t'):
            t.text = '\u2612' if checked else '\u2610'   # ☒ or ☐


def _append_to_label_para(cell_elem, label, value):
    """
    Find a paragraph in cell_elem whose text starts with label,
    and append value to it.
    Handles multi-run paragraphs (e.g. 'DOCUMENT REFERENCE' + ':' as separate runs).
    """
    paras = cell_elem.findall(f'.//{{{W}}}p')
    for p in paras:
        # Get FULL text across all runs/t-elements
        full_txt = _get_text(p).strip()
        if full_txt.startswith(label.strip()):
            if value:
                combined = full_txt.rstrip() + ' ' + str(value)
            else:
                combined = full_txt
            # Write combined text into first t-element, clear the rest
            _set_para_text(p, combined)
            return


def _fill_cell_text(cell, text):
    """Clear a cell and fill it with plain text."""
    paras = cell._tc.findall(f'.//{{{W}}}p')
    if not paras:
        return
    # Clear all text from first paragraph
    for t in paras[0].iter(f'{{{W}}}t'):
        t.text = ''
    # Add new run to first paragraph
    r = etree.SubElement(paras[0], f'{{{W}}}r')
    t = etree.SubElement(r, f'{{{W}}}t')
    t.text = str(text)
    if text and (str(text)[0] == ' ' or str(text)[-1] == ' '):
        t.set(XML_SPACE, 'preserve')
    # Clear remaining paragraphs
    for p in paras[1:]:
        for t in p.iter(f'{{{W}}}t'):
            t.text = ''


# ── Reference builder ──────────────────────────────────────────────────────────

def build_rd3_reference(eng_full, nt_ft, idi_no, taw_pol, ins_type='Malath'):
    """
    Malath:   {INITIALS}-RD3-{NT/FT}{IDI_NO}-{TAW_POL}-1
    Tawuniya: {INITIALS}-RD3-{TAW_POL}-01
    Also returns the short ref used in the Ref.: body line.
    """
    parts = eng_full.strip().split()
    if len(parts) >= 2:
        initials = (parts[0][0] + parts[1][:2]).upper()
    elif parts:
        initials = parts[0][:3].upper()
    else:
        initials = 'ENG'

    idi = str(idi_no).strip().replace('.0', '')
    taw = str(taw_pol).strip().replace('.0', '')
    nt  = str(nt_ft).strip()

    if ins_type == 'Tawuniya':
        full_ref  = '{}-RD3-{}-01'.format(initials, taw) if taw else '{}-RD3-{}-01'.format(initials, idi)
        short_ref = taw if taw else idi
    else:
        if taw:
            full_ref  = '{}-RD3-{}{}-{}-1'.format(initials, nt, idi, taw)
            short_ref = '{}{}-{}'.format(nt, idi, taw)
        else:
            full_ref  = '{}-RD3-{}{}-1'.format(initials, nt, idi)
            short_ref = '{}{}'.format(nt, idi)

    return full_ref, short_ref


# ── Table 1 — Document header + expert info ───────────────────────────────────

def _fill_table1(table, data):
    """Fill Table 1: doc ref, date, expert name/phase/degree/speciality, author block."""
    rd3_ref    = data.get('rd3_ref', '')
    issue_date = data.get('issue_date', '')
    eng_full   = data.get('eng_full', '')
    eng_phase  = data.get('eng_phase', 'Waterproofing')
    eng_degree = data.get('eng_degree', 'Bachelor')
    eng_spec   = data.get('eng_speciality', 'Civil Engineer')
    eng_phone  = data.get('eng_phone', '')
    eng_email  = data.get('eng_email', '')

    # Row 0 col 0: DOCUMENT REFERENCE
    _append_to_label_para(table.rows[0].cells[0]._tc, 'DOCUMENT REFERENCE', rd3_ref)
    # Row 0 col 2: TIS AGENCY (always SOCOTEC ARABIA)
    _append_to_label_para(table.rows[0].cells[2]._tc, 'TIS AGENCY', 'SOCOTEC ARABIA')
    # Row 1 col 0: VERSION
    _append_to_label_para(table.rows[1].cells[0]._tc, 'VERSION', '1')
    # Row 1 col 2: DATE OF ISSUE
    _append_to_label_para(table.rows[1].cells[2]._tc, 'DATE OF ISSUE', issue_date)

    # Row 4: Expert Name / Phase / Degree / Speciality
    _fill_cell_text(table.rows[4].cells[0], eng_full)
    _fill_cell_text(table.rows[4].cells[1], eng_phase)
    _fill_cell_text(table.rows[4].cells[2], eng_degree)
    _fill_cell_text(table.rows[4].cells[3], eng_spec)

    # Row 5: AUTOR block (single merged cell)
    cell5 = table.rows[5].cells[0]
    _append_to_label_para(cell5._tc, 'AUTOR of THIS REPORT', eng_full)
    _append_to_label_para(cell5._tc, 'PHONE NUMBER', eng_phone)
    _append_to_label_para(cell5._tc, 'EMAIL', eng_email)


# ── Table 2 — Main report body ────────────────────────────────────────────────

def _fill_table2(table, data, visits):
    """Fill Table 2: project info, occupancy, site visits, conclusion."""
    # Row 1: project information cell
    proj_cell = table.rows[1].cells[0]._tc
    _append_to_label_para(proj_cell, 'PROJECT TITLE (NAME)',          data.get('project_title', ''))
    _append_to_label_para(proj_cell, 'ADDRESS OF THE PREMISES',       data.get('address', ''))
    _append_to_label_para(proj_cell, 'REFERENCE RD0',                 data.get('rd0_ref', ''))
    _append_to_label_para(proj_cell, 'PRINCIPAL/OWNER',               data.get('owner', ''))
    _append_to_label_para(proj_cell, 'BUILDINGS INCLUDED IN THE PROJECT AND ITS USE',
                                                                       data.get('buildings', '1 residential building'))

    # Row 5 col 2: occupancy date + Expected/Confirmed
    occ_cell = table.rows[5].cells[2]._tc
    _append_to_label_para(occ_cell, 'Date of Occupancy Certificate', data.get('occ_date', ''))

    occ_status = data.get('occ_status', 'Expected')   # 'Expected' or 'Confirmed'
    # These are plain ☐ in a run — replace the whole paragraph text
    for p in occ_cell.findall(f'.//{{{W}}}p'):
        txt = _get_text(p)
        if 'Expected' in txt and 'Confirmed' in txt:
            if occ_status == 'Expected':
                new = txt.replace('\u2610 Expected', '\u2612 Expected').replace('\u2610 Confirmed', '\u2610 Confirmed')
            else:
                new = txt.replace('\u2610 Confirmed', '\u2612 Confirmed').replace('\u2610 Expected', '\u2610 Expected')
            _set_para_text(p, new)
            break

    # Row 6 col 2: visits nested table
    visit_cell = table.rows[6].cells[2]._tc
    nested_tbls = visit_cell.findall(f'.//{{{W}}}tbl')
    if nested_tbls and visits:
        nested_tbl = nested_tbls[0]
        rows = nested_tbl.findall(f'{{{W}}}tr')
        for i, v in enumerate(visits[:10]):   # max 10 rows in template
            row_idx = i + 1   # row 0 is the header
            if row_idx >= len(rows):
                break
            tr = rows[row_idx]
            cells = tr.findall(f'.//{{{W}}}tc')
            values = [
                v.get('ref', ''),
                v.get('date', ''),
                v.get('inspector', ''),
                v.get('part', ''),
            ]
            for ci, (tc, val) in enumerate(zip(cells, values)):
                _fill_cell_text_tc(tc, val)

    # Row 7: defects text
    defects = data.get('defects_text', '')
    if defects:
        _append_text_to_cell_last_para(table.rows[7].cells[2]._tc, defects)

    # Row 9: reservations (YES/NO + details) — plain ☐ toggle
    reservations = data.get('reservations', 'NO')
    for p in table.rows[9].cells[0]._tc.findall(f'.//{{{W}}}p') if len(table.rows) > 9 else []:
        txt = _get_text(p)
        if '☐ YES' in txt and '☐ NO' in txt:
            if reservations == 'YES':
                new = txt.replace('☐ YES', '☑ YES').replace('☐ NO', '☐ NO')
            else:
                new = txt.replace('☐ YES', '☐ YES').replace('☐ NO', '☑ NO')
            _set_para_text(p, new)
            break

    # Row 11: main conclusion (nested table + paragraph)
    if len(table.rows) > 11:
        conc_cell = table.rows[11].cells[0]._tc
        # Toggle YES/NO in nested table
        conc_yn = data.get('conclusion_yn', 'YES')
        for p in conc_cell.findall(f'.//{{{W}}}p'):
            txt = _get_text(p)
            if '☐ YES' in txt and '☐ NO' in txt:
                if conc_yn == 'YES':
                    new = txt.replace('☐ YES', '☑ YES')
                else:
                    new = txt.replace('☐ NO', '☑ NO')
                _set_para_text(p, new)
                break
        # Add conclusion text
        conc_text = data.get('conclusion_text', '')
        if conc_text:
            _append_text_to_cell_last_para(conc_cell, conc_text)


def _fill_cell_text_tc(tc_elem, text):
    """Fill a table cell element (w:tc) with text."""
    paras = tc_elem.findall(f'.//{{{W}}}p')
    if not paras:
        return
    for t in paras[0].iter(f'{{{W}}}t'):
        t.text = ''
    r = etree.SubElement(paras[0], f'{{{W}}}r')
    t = etree.SubElement(r, f'{{{W}}}t')
    t.text = str(text) if text else ''
    if text and (str(text)[0] == ' ' or str(text)[-1] == ' '):
        t.set(XML_SPACE, 'preserve')


def _append_text_to_cell_last_para(tc_elem, text):
    """Append text as a new run to the last meaningful paragraph in a cell."""
    paras = tc_elem.findall(f'.//{{{W}}}p')
    if not paras:
        return
    # Find the 'Please expose...' paragraph and add text after it
    for p in paras:
        txt = _get_text(p)
        if 'Please expose' in txt or 'develop' in txt.lower():
            # Add text to next empty paragraph
            idx = paras.index(p)
            if idx + 1 < len(paras):
                target = paras[idx + 1]
            else:
                target = paras[-1]
            r = etree.SubElement(target, f'{{{W}}}r')
            t_elem = etree.SubElement(r, f'{{{W}}}t')
            t_elem.text = text
            t_elem.set(XML_SPACE, 'preserve')
            return
    # Fallback: add to last paragraph
    r = etree.SubElement(paras[-1], f'{{{W}}}r')
    t_elem = etree.SubElement(r, f'{{{W}}}t')
    t_elem.text = text
    t_elem.set(XML_SPACE, 'preserve')


# ── Annex table filler ────────────────────────────────────────────────────────

def _fill_annex_table(table, ann, annex_type):
    """
    Fill an Annex table (T3=roofs, T4=facades, T5=basements).
    ann dict keys: description_i1, description_i2, description_i3,
                   innovative_yn, reviewed_docs_materials, reviewed_docs_tests,
                   reviewed_docs_other, conclusion_text, conclusion_yn
    """
    cell0 = table.rows[2].cells[0]._tc   # Works Description

    # I.1 description
    if ann.get('description_i1'):
        _append_after_label(cell0, 'I.1.', ann['description_i1'])

    # I.2 layers description
    if ann.get('description_i2'):
        _append_after_label(cell0, 'I.2.', ann['description_i2'])

    # I.3 junctions (only for roofs)
    if ann.get('description_i3') and annex_type == 'roof':
        _append_after_label(cell0, 'I.3.', ann['description_i3'])

    # I.4 Innovative technique YES/NO (plain ☐ in paragraph)
    if annex_type == 'roof':
        inn = ann.get('innovative_yn', 'NO')
        for p in cell0.findall(f'.//{{{W}}}p'):
            txt = _get_text(p)
            if 'innovative' in txt.lower() and '☐ YES' in txt:
                if inn == 'YES':
                    new = txt.replace('☐ YES', '☒ YES')
                else:
                    new = txt.replace('☐ NO', '☒ NO')
                _set_para_text(p, new)
                break

    # Row 4: Materials section — fill reviewed docs
    mat_cell = table.rows[4].cells[0]._tc
    docs_mat  = ann.get('reviewed_docs_materials', [])
    docs_test = ann.get('reviewed_docs_tests', [])
    docs_other= ann.get('reviewed_docs_other', [])

    # Fill delivery docs 1-4
    _fill_numbered_docs(mat_cell, docs_mat, start_after='Delivery Orders - Material certificates', count=4)
    # Fill test docs 1-4
    _fill_numbered_docs(mat_cell, docs_test, start_after='Tests and quality control reports', count=4)
    # Fill other docs 1-4
    _fill_numbered_docs(mat_cell, docs_other, start_after='Oher Reviewed documents', count=4)

    # Conclusion row (last row of annex)
    last_row = table.rows[-1]
    conc_cell = last_row.cells[0]._tc

    # YES/NO toggle in nested table (plain ☐)
    conc_yn = ann.get('conclusion_yn', 'YES')
    for p in conc_cell.findall(f'.//{{{W}}}p'):
        txt = _get_text(p)
        if '☐ YES' in txt and '☐ NO' in txt:
            if conc_yn == 'YES':
                new = txt.replace('☐ YES', '☒ YES')
            else:
                new = txt.replace('☐ NO', '☒ NO')
            _set_para_text(p, new)
            break

    # Conclusion text
    if ann.get('conclusion_text'):
        _append_text_to_cell_last_para(conc_cell, ann['conclusion_text'])


def _append_after_label(tc_elem, label_prefix, text):
    """Find a paragraph starting with label_prefix and append text to the next empty paragraph."""
    paras = tc_elem.findall(f'.//{{{W}}}p')
    for i, p in enumerate(paras):
        txt = _get_text(p)
        if txt.strip().startswith(label_prefix):
            # Look for next empty paragraph to put text in
            for j in range(i + 1, min(i + 5, len(paras))):
                next_txt = _get_text(paras[j]).strip()
                if not next_txt:
                    r = etree.SubElement(paras[j], f'{{{W}}}r')
                    t_elem = etree.SubElement(r, f'{{{W}}}t')
                    t_elem.text = text
                    t_elem.set(XML_SPACE, 'preserve')
                    return
            # No empty para found, append to same paragraph
            r = etree.SubElement(p, f'{{{W}}}r')
            t_elem = etree.SubElement(r, f'{{{W}}}t')
            t_elem.text = ' ' + text
            t_elem.set(XML_SPACE, 'preserve')
            return


def _fill_numbered_docs(tc_elem, docs, start_after, count=4):
    """Fill numbered list items (\t1., \t2., ...) after a section header."""
    if not docs:
        return
    paras = tc_elem.findall(f'.//{{{W}}}p')
    in_section = False
    doc_idx = 0
    for p in paras:
        txt = _get_text(p)
        if start_after in txt:
            in_section = True
            continue
        if in_section and doc_idx < len(docs):
            # Match numbered items like '\t1.' or '\t2.'
            stripped = txt.strip()
            if stripped and stripped[0].isdigit() and '.' in stripped[:3]:
                # Fill this numbered item
                num = stripped.split('.')[0]
                t_elems = list(p.iter(f'{{{W}}}t'))
                if t_elems:
                    t_elems[0].text = '\t{}. {}'.format(num, docs[doc_idx])
                    t_elems[0].set(XML_SPACE, 'preserve')
                    for t in t_elems[1:]:
                        t.text = ''
                doc_idx += 1
            elif doc_idx > 0 and not stripped.isdigit():
                # We've left the section
                if stripped and not stripped[0].isdigit():
                    break


# ── Signature block replacement ───────────────────────────────────────────────

def _fill_signature_blocks(body, data):
    """
    Replace the 4 signature placeholder blocks in the document body:
    - 'Made in (CITY), (DATE)'       →  'Issued in Riyadh on {date}'
    - '(RESPONSIBLE EXPERTS NAMES)'  →  'Eng. {name}  Reviewer:  HEAD OF LOCAL DEPT...'
    - '(RESPONSIBLE EXPERTS SIGNATURES + INK PAD)'  →  leave blank (manual)
    - '(RESPONSIBLE EXPERTS POSITIONS)'  →  keep for manual
    """
    eng_full  = data.get('eng_full', '')
    reviewer  = data.get('reviewer_name', '')
    manager   = data.get('manager_name', 'Nizar Lazreg')
    issue_str = data.get('issue_date', '')

    # Format date for "Issued in Riyadh on" line
    # Convert d/m/yyyy → dd-mm-yyyy
    try:
        parts = issue_str.split('/')
        issue_fmt = '{:02d}-{:02d}-{}'.format(int(parts[0]), int(parts[1]), parts[2])
    except Exception:
        issue_fmt = issue_str

    # City
    city = data.get('city', 'Riyadh')

    # Replace all occurrences
    _replace_global(body, 'Made in (CITY), (DATE)',
                    'Issued in {} on {}'.format(city, issue_fmt))

    _replace_global(body, '(RESPONSIBLE EXPERTS NAMES)',
                    'Eng. {}\t\t\t\tReviewer:\t\t\t\tHEAD OF LOCAL DEPARTMENT OR MANAGER'.format(eng_full))

    if reviewer:
        _replace_global(body, '(RESPONSIBLE EXPERTS SIGNATURES + INK PAD)',
                        '\t\t\t\t\tEng. {}\t\t\t\tEng. {}'.format(reviewer, manager.upper()))

    # Positions line — leave as placeholder if no data, or fill positions
    _replace_global(body, '(RESPONSIBLE EXPERTS POSITIONS)', '')

    # Body Ref. paragraph (exact match)
    for p in body.iter(f'{{{W}}}p'):
        txt = _get_text(p)
        if txt.strip() == 'Ref.:':
            short_ref = data.get('short_ref', '')
            _set_para_text(p, 'Ref.:  {}'.format(short_ref))
            break

    # Tawuniya Visit ID paragraph
    for p in body.iter(f'{{{W}}}p'):
        txt = _get_text(p)
        if txt.strip() == 'Tawuniya visit ID:':
            taw_id = data.get('taw_visit_id', '')
            _set_para_text(p, 'TAWUNIYA Visit ID:  {}'.format(taw_id))
            break


# ── SDT checkbox logic ────────────────────────────────────────────────────────

def _toggle_all_sdts(body, data, annex_data):
    """Toggle all 30 SDT checkboxes based on data."""
    has_roofs     = 'roofs'     in annex_data
    has_facades   = 'facades'   in annex_data
    has_basements = 'basements' in annex_data

    # CB00-02: WP works
    _toggle_sdt(body, 0,  has_roofs)
    _toggle_sdt(body, 1,  has_facades)
    _toggle_sdt(body, 2,  has_basements)

    # CB03-07: Annex 1 roof types
    if has_roofs:
        roof_types = annex_data['roofs'].get('roof_types', [])
        _toggle_sdt(body, 3,  'ROOF'                in roof_types)
        _toggle_sdt(body, 4,  'ROOFTOP TERRACE'     in roof_types)
        _toggle_sdt(body, 5,  'INTERMEDIATE TERRACE'in roof_types)
        _toggle_sdt(body, 6,  'PATIOS'              in roof_types)
        _toggle_sdt(body, 7,  'OTHER'               in roof_types)

    # CB08-13: Annex 1 materials Y/N
    if has_roofs:
        rd = annex_data['roofs']
        _toggle_sdt(body, 8,  rd.get('delivery_yn',  'YES') == 'YES')
        _toggle_sdt(body, 9,  rd.get('delivery_yn',  'YES') == 'NO')
        _toggle_sdt(body, 10, rd.get('compliant_yn', 'YES') == 'YES')
        _toggle_sdt(body, 11, rd.get('compliant_yn', 'YES') == 'NO')
        _toggle_sdt(body, 12, rd.get('ponding_yn',   'YES') == 'YES')
        _toggle_sdt(body, 13, rd.get('ponding_yn',   'YES') == 'NO')

    # CB14-17: Annex 2 facade types
    if has_facades:
        fac_types = annex_data['facades'].get('facade_types', [])
        _toggle_sdt(body, 14, 'CONCRETE OR MASONRY' in fac_types)
        _toggle_sdt(body, 15, 'CLADDING'            in fac_types)
        _toggle_sdt(body, 16, 'CURTAIN WALL'        in fac_types)
        _toggle_sdt(body, 17, 'OTHER'               in fac_types)

    # CB18-23: Annex 2 materials Y/N
    if has_facades:
        fd = annex_data['facades']
        _toggle_sdt(body, 18, fd.get('delivery_yn',  'YES') == 'YES')
        _toggle_sdt(body, 19, fd.get('delivery_yn',  'YES') == 'NO')
        _toggle_sdt(body, 20, fd.get('compliant_yn', 'YES') == 'YES')
        _toggle_sdt(body, 21, fd.get('compliant_yn', 'YES') == 'NO')
        _toggle_sdt(body, 22, fd.get('ponding_yn',   'YES') == 'YES')
        _toggle_sdt(body, 23, fd.get('ponding_yn',   'YES') == 'NO')

    # CB24-29: Annex 3 materials Y/N
    if has_basements:
        bd = annex_data['basements']
        _toggle_sdt(body, 24, bd.get('delivery_yn',  'YES') == 'YES')
        _toggle_sdt(body, 25, bd.get('delivery_yn',  'YES') == 'NO')
        _toggle_sdt(body, 26, bd.get('compliant_yn', 'YES') == 'YES')
        _toggle_sdt(body, 27, bd.get('compliant_yn', 'YES') == 'NO')
        _toggle_sdt(body, 28, bd.get('ponding_yn',   'YES') == 'YES')
        _toggle_sdt(body, 29, bd.get('ponding_yn',   'YES') == 'NO')


# ── Main entry point ──────────────────────────────────────────────────────────

def generate_rd3(template_path, output_path, data, visits, annex_data):
    """
    Fill Template_RD3.docx and save to output_path.

    Parameters
    ----------
    template_path : str – path to Template_RD3.docx
    output_path   : str – path for output .docx
    data          : dict with keys:
        rd3_ref, short_ref, issue_date, city,
        eng_full, eng_phase, eng_degree, eng_speciality, eng_phone, eng_email,
        reviewer_name, manager_name,
        project_title, address, owner, rd0_ref, buildings,
        occ_date, occ_status (Expected/Confirmed),
        taw_visit_id,
        defects_text, reservations (YES/NO),
        conclusion_yn (YES/NO), conclusion_text
    visits        : list of {ref, date, inspector, part}
    annex_data    : dict — keys are 'roofs', 'facades', 'basements' (only present if selected)
        Each value is a dict with:
        For roofs:     roof_types (list), description_i1, description_i2, description_i3,
                       innovative_yn, delivery_yn, compliant_yn, ponding_yn,
                       reviewed_docs_materials, reviewed_docs_tests, reviewed_docs_other,
                       conclusion_yn, conclusion_text, other_type_text
        For facades:   facade_types (list), description_i1, description_i2,
                       delivery_yn, compliant_yn, ponding_yn,
                       reviewed_docs_materials, reviewed_docs_tests, reviewed_docs_other,
                       conclusion_yn, conclusion_text, other_type_text
        For basements: description_i1, description_i2,
                       delivery_yn, compliant_yn, ponding_yn,
                       reviewed_docs_materials, reviewed_docs_tests, reviewed_docs_other,
                       conclusion_yn, conclusion_text
    """
    doc = Document(template_path)
    body = doc.element.body

    # 1. Toggle SDT checkboxes
    _toggle_all_sdts(body, data, annex_data)

    # 2. Fill signature / date blocks (body paragraphs)
    _fill_signature_blocks(body, data)

    # 3. Table 1: document header + expert info
    _fill_table1(doc.tables[1], data)

    # 4. Table 2: main report
    _fill_table2(doc.tables[2], data, visits)

    # 5. Annex tables (only for selected annexes)
    if 'roofs' in annex_data and len(doc.tables) > 3:
        _fill_annex_table(doc.tables[3], annex_data['roofs'], 'roof')
        # Handle OTHER type text in roof type row
        if 'OTHER' in annex_data['roofs'].get('roof_types', []):
            other_text = annex_data['roofs'].get('other_type_text', '')
            if other_text:
                _replace_global(body, 'OTHER: \t', 'OTHER: {}\t'.format(other_text))

    if 'facades' in annex_data and len(doc.tables) > 4:
        _fill_annex_table(doc.tables[4], annex_data['facades'], 'facade')
        if 'OTHER' in annex_data['facades'].get('facade_types', []):
            other_text = annex_data['facades'].get('other_type_text', '')
            if other_text:
                _replace_global(body, 'OTHER: \t', 'OTHER: {}\t'.format(other_text))

    if 'basements' in annex_data and len(doc.tables) > 5:
        _fill_annex_table(doc.tables[5], annex_data['basements'], 'basement')

    doc.save(output_path)
    return output_path
