"""
rd6_generator.py — Fills RD6-AutoTemplate.docx (v4 reference).

Template layout (v4):
  - rId9  (image1.jpeg): reviewer+manager sig — FIXED, never replaced
  - rId10 (image2.png):  second reviewer/manager sig — FIXED, never replaced
  - rId11/rId12: insulation cert pages (re-generated each time)
  - Author sig: injected as NEW rId (rId13+) anchor at column/200000 (left)

Engineer name appears on:
  - Para "Eng. {name}  Reviewer:  MANAGER"  (already hardcoded in template as Mohamed Mossad)
  - Para "  [spaces]  Eng. {name}  Eng. NIZAR LAZREG"
Both replaced via hardcoded text replacement.
"""
import copy, io, os, shutil, struct, subprocess, sys, tempfile, zipfile
from pathlib import Path
from datetime import datetime
from lxml import etree

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'

BLUE = '2F5496'

# ── Missing docs ───────────────────────────────────────────────────────────────
DOC_KEYS = [
    'cost_letter','contractor_letter','supervision_letter','calc_notes',
    'soil_tests','concrete_tests','steel_invoices','material_warranty',
]
STANDARD_MISSING_DOCS = [
    ('Cost letter',
     'Cost letter — A letter from the contractor stating the actual cost of the project '
     'after completion, which the owner did not provide to us'),
    ('Contractor letter',
     'Contractor letter — A letter from the contractor confirming that design remarks '
     'sent by email or noted during visits were taken into consideration during execution.'),
    ('Engineering supervision letter',
     'Engineering supervision letter — A stamped certificate from the supervising '
     'engineering office covering all project items'),
    ('Calculation notes',
     'Calculation notes — Approved structural design calculation notes & design criteria'),
    ('Soil / compaction tests',
     'Soil / compaction tests — All results for compaction tests under foundations, '
     'between foundations, column necks, and grade beams'),
    ('Concrete strength tests',
     'Concrete strength tests — Compressive strength results for foundations, column '
     'necks, grade beams, ground floor columns/slab, first floor columns/slab, and annex floor.'),
    ('Steel invoices / warranty',
     'Steel invoices / warranty — All rebar delivery invoices or SASO-certified warranty'),
    ('Material warranty certificates',
     'Material warranty certificates — Warranty certificates for structural and façade elements'),
]


# ── Blue color helper ──────────────────────────────────────────────────────────
def _apply_blue(t_elem):
    r = t_elem.getparent()
    if r is None or r.tag != f'{{{W}}}r':
        return
    rPr = r.find(f'{{{W}}}rPr')
    if rPr is None:
        rPr = etree.Element(f'{{{W}}}rPr')
        r.insert(0, rPr)
    color = rPr.find(f'{{{W}}}color')
    if color is None:
        color = etree.SubElement(rPr, f'{{{W}}}color')
    color.set(f'{{{W}}}val', BLUE)


# ── Content control filler ─────────────────────────────────────────────────────
def _fill_sdt(tree, tag_val, new_text):
    for sdt in tree.iter(f'{{{W}}}sdt'):
        tag_elem = sdt.find(f'.//{{{W}}}tag')
        if tag_elem is None or tag_elem.get(f'{{{W}}}val','') != tag_val:
            continue
        sdt_content = sdt.find(f'{{{W}}}sdtContent')
        if sdt_content is None:
            continue
        t_elems = list(sdt_content.iter(f'{{{W}}}t'))
        if t_elems:
            # Normal case: t elements exist, update them
            t_elems[0].text = str(new_text)
            for t in t_elems[1:]:
                t.text = ''
            for t in t_elems:
                _apply_blue(t)
        else:
            # Empty control — check if sdt is inline (parent=w:p) or block (parent=w:tc/body)
            sdt_parent = sdt.getparent()
            sdt_parent_tag = sdt_parent.tag.split('}')[1] if sdt_parent is not None else ''
            
            if sdt_parent_tag == 'p':
                # INLINE control (sdt inside w:p): add w:r directly to sdtContent
                r = etree.SubElement(sdt_content, f'{{{W}}}r')
                rPr = etree.SubElement(r, f'{{{W}}}rPr')
                color = etree.SubElement(rPr, f'{{{W}}}color')
                color.set(f'{{{W}}}val', BLUE)
                t = etree.SubElement(r, f'{{{W}}}t')
                t.text = str(new_text)
            else:
                # BLOCK control (sdt inside w:tc or w:body): add w:p -> w:r
                para = sdt_content.find(f'.//{{{W}}}p')
                if para is None:
                    para = etree.SubElement(sdt_content, f'{{{W}}}p')
                r = etree.SubElement(para, f'{{{W}}}r')
                rPr = etree.SubElement(r, f'{{{W}}}rPr')
                color = etree.SubElement(rPr, f'{{{W}}}color')
                color.set(f'{{{W}}}val', BLUE)
                t = etree.SubElement(r, f'{{{W}}}t')
                t.text = str(new_text)
            if new_text and (str(new_text)[0] == ' ' or str(new_text)[-1] == ' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def fill_all_controls(tree, mapping):
    for tag, value in mapping.items():
        _fill_sdt(tree, tag, str(value) if value is not None else '')


# ── Hardcoded text replacement with blue color ─────────────────────────────────
def _replace_in_tree(tree, old, new):
    for para in tree.iter(f'{{{W}}}p'):
        t_elems = list(para.iter(f'{{{W}}}t'))
        full = ''.join(t.text or '' for t in t_elems)
        if old not in full:
            continue
        new_full = full.replace(old, new)
        if t_elems:
            t_elems[0].text = new_full
            for t in t_elems[1:]:
                t.text = ''
            _apply_blue(t_elems[0])

def replace_hardcoded(tree, replacements):
    for old, new in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
        _replace_in_tree(tree, old, new)


# ── Missing docs text ──────────────────────────────────────────────────────────
def build_missing_doc_text(provided_keys):
    missing = [
        (i+1, long_desc)
        for i, (key, (_, long_desc)) in enumerate(zip(DOC_KEYS, STANDARD_MISSING_DOCS))
        if key not in provided_keys
    ]
    if not missing:
        return 'All required documents were provided by the client.'
    return '\n'.join(f"{n}. {desc}" for n, desc in missing)


# ── Author signature injection ─────────────────────────────────────────────────
def _build_anchor_xml(rid, cx, cy, h_offset, v_offset, img_id):
    """Build a floating anchor XML element for the author signature."""
    return f'''<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               distT="0" distB="0" distL="114300" distR="114300"
               simplePos="0" relativeHeight="251659264" behindDoc="0"
               locked="0" layoutInCell="1" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="column">
    <wp:posOffset>{h_offset}</wp:posOffset>
  </wp:positionH>
  <wp:positionV relativeFrom="paragraph">
    <wp:posOffset>{v_offset}</wp:posOffset>
  </wp:positionV>
  <wp:extent cx="{cx}" cy="{cy}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="{img_id}" name="AuthorSig{img_id}"/>
  <wp:cNvGraphicFramePr/>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr>
          <pic:cNvPr id="{img_id+100}" name="AuthorSig{img_id}"/>
          <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="{rid}"/>
          <a:stretch><a:fillRect/></a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:anchor>'''


def _inject_author_sig_into_zip(docx_path, out_path, sig_bytes, sig_ext):
    """
    Add author signature as a new floating anchor in both sig paragraphs.
    Strategy: copy the existing rId9 anchor, change rId and h_offset to 200000.
    This avoids namespace issues from building anchor XML from scratch.
    """
    with zipfile.ZipFile(docx_path, 'r') as zin:
        doc_xml  = zin.read('word/document.xml')
        rels_xml = zin.read('word/_rels/document.xml.rels')
        ct_xml   = zin.read('[Content_Types].xml')

    doc_tree  = etree.fromstring(doc_xml)
    rels_tree = etree.fromstring(rels_xml)
    ct_tree   = etree.fromstring(ct_xml)

    # Find next available rId
    existing = [int(r.get('Id','rId0').replace('rId',''))
                for r in rels_tree if r.get('Id','').startswith('rId')]
    next_rid_num = max(existing, default=20) + 1
    sig_rid = f'rId{next_rid_num}'
    clean_ext = sig_ext.lstrip('.').lower()
    media_name = f'media/author_sig.{clean_ext}'

    # Add relationship
    etree.SubElement(rels_tree, f'{{{PKG}}}Relationship', {
        'Id': sig_rid,
        'Type': f'{R}/image',
        'Target': media_name,
    })

    # Add content type if needed
    ct_map = {'jpeg':'image/jpeg','jpg':'image/jpeg','png':'image/png'}
    existing_exts = {e.get('Extension','').lower() for e in ct_tree}
    if clean_ext not in existing_exts:
        etree.SubElement(ct_tree, f'{{{CT}}}Default',
                         {'Extension': clean_ext,
                          'ContentType': ct_map.get(clean_ext,'image/png')})

    # Calculate sig dimensions
    try:
        from PIL import Image as PILImage
        pil = PILImage.open(io.BytesIO(sig_bytes))
        w_px, h_px = pil.size
        target_cx = 990600
        target_cy = int(target_cx * h_px / w_px)
    except Exception:
        target_cx, target_cy = 990600, 400050

    # Find image paragraphs (those containing rId9 anchor)
    all_paras = list(doc_tree.iter(f'{{{W}}}p'))
    v_offsets = [10160, 6985]
    img_para_indices = []
    for i, p in enumerate(all_paras):
        for anchor in p.iter(f'{{{WP}}}anchor'):
            for b in anchor.iter(f'{{{A}}}blip'):
                if b.get(f'{{{R}}}embed') == 'rId9':
                    img_para_indices.append(i)
                    break

    for idx, para_i in enumerate(img_para_indices):
        para = all_paras[para_i]

        # Find the existing rId9 anchor and DEEP COPY it as template
        src_anchor = None
        for anchor in para.iter(f'{{{WP}}}anchor'):
            for b in anchor.iter(f'{{{A}}}blip'):
                if b.get(f'{{{R}}}embed') == 'rId9':
                    src_anchor = anchor
                    break
            if src_anchor is not None:
                break

        if src_anchor is None:
            continue

        # Deep copy the anchor (preserves all namespaces correctly)
        new_anchor = copy.deepcopy(src_anchor)

        # Change blip rId to our new signature
        for b in new_anchor.iter(f'{{{A}}}blip'):
            b.set(f'{{{R}}}embed', sig_rid)

        # Change positionH to column/200000 (left, under Author)
        posH = new_anchor.find(f'{{{WP}}}positionH')
        if posH is not None:
            posH.set('relativeFrom', 'column')
            align = posH.find(f'{{{WP}}}align')
            if align is not None:
                posH.remove(align)
            offset = posH.find(f'{{{WP}}}posOffset')
            if offset is None:
                offset = etree.SubElement(posH, f'{{{WP}}}posOffset')
            offset.text = '200000'

        # Change positionV offset
        posV = new_anchor.find(f'{{{WP}}}positionV')
        if posV is not None:
            v_off = posV.find(f'{{{WP}}}posOffset')
            if v_off is not None:
                v_off.text = str(v_offsets[idx] if idx < len(v_offsets) else 10160)

        # Change wp:extent (inline size hint)
        ext = new_anchor.find(f'{{{WP}}}extent')
        if ext is not None:
            ext.set('cx', str(target_cx))
            ext.set('cy', str(target_cy))
        # Change a:ext inside a:xfrm ONLY (not extLst - those must not have cx/cy)
        XFRM = '{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm'
        for xfrm in new_anchor.iter(XFRM):
            for a_ext in xfrm.iter(f'{{{A}}}ext'):
                a_ext.set('cx', str(target_cx))
                a_ext.set('cy', str(target_cy))

        # Change docPr id and name to avoid conflicts
        doc_pr = new_anchor.find(f'{{{WP}}}docPr')
        if doc_pr is not None:
            doc_pr.set('id', str(500 + idx))
            doc_pr.set('name', f'AuthorSig{idx+1}')

        # Change cNvPr id
        for cNvPr in new_anchor.iter('{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr'):
            cNvPr.set('id', str(600 + idx))
            cNvPr.set('name', f'AuthorSig{idx+1}')

        # Wrap in w:r and insert at start of paragraph
        r_elem = etree.Element(f'{{{W}}}r')
        rPr = etree.SubElement(r_elem, f'{{{W}}}rPr')
        etree.SubElement(rPr, f'{{{W}}}noProof')
        draw = etree.SubElement(r_elem, f'{{{W}}}drawing')
        draw.append(new_anchor)
        para.insert(0, r_elem)

    # Serialize
    new_doc  = etree.tostring(doc_tree,  xml_declaration=True, encoding='UTF-8', standalone=True)
    new_rels = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    new_ct   = etree.tostring(ct_tree,   xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name == 'word/document.xml':       zout.writestr(item, new_doc)
                elif name == 'word/_rels/document.xml.rels': zout.writestr(item, new_rels)
                elif name == '[Content_Types].xml':   zout.writestr(item, new_ct)
                else:                                 zout.writestr(item, zin.read(name))
            zout.writestr(f'word/{media_name}', sig_bytes)


def _inject_footer_reference(footer_bytes, rd6_ref):
    """
    Replace the first paragraph in the footer that looks like an RD6 reference
    (contains 'RD6' or matches our reference pattern).
    If none found, prepend a new paragraph.
    This prevents duplicate references when the template already has one.
    """
    import re
    tree = etree.fromstring(footer_bytes)

    # Find existing reference paragraph to REPLACE (not add on top)
    ref_para = None
    for para in tree.iter(f'{{{W}}}p'):
        t_elems = list(para.iter(f'{{{W}}}t'))
        text = ''.join(t.text or '' for t in t_elems)
        if re.search(r'[A-Z]{2,4}-RD[0-9]-', text):
            ref_para = para
            break

    if ref_para is not None:
        # Replace all runs in existing paragraph with new reference
        for r in list(ref_para.findall(f'{{{W}}}r')):
            ref_para.remove(r)
        r    = etree.SubElement(ref_para, f'{{{W}}}r')
        rPr  = etree.SubElement(r, f'{{{W}}}rPr')
        col  = etree.SubElement(rPr, f'{{{W}}}color')
        col.set(f'{{{W}}}val', BLUE)
        sz   = etree.SubElement(rPr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '16')
        t    = etree.SubElement(r, f'{{{W}}}t')
        t.text = str(rd6_ref)
    else:
        # No existing reference — prepend new paragraph
        new_para = etree.Element(f'{{{W}}}p')
        pPr  = etree.SubElement(new_para, f'{{{W}}}pPr')
        jc   = etree.SubElement(pPr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', 'left')
        r    = etree.SubElement(new_para, f'{{{W}}}r')
        rPr  = etree.SubElement(r, f'{{{W}}}rPr')
        col  = etree.SubElement(rPr, f'{{{W}}}color')
        col.set(f'{{{W}}}val', BLUE)
        sz   = etree.SubElement(rPr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '16')
        t    = etree.SubElement(new_para, f'{{{W}}}t')
        t.text = str(rd6_ref)
        tree.insert(0, new_para)

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)

def _inject_footers(docx_path, out_path, rd6_ref):
    footers = ['word/footer1.xml','word/footer2.xml','word/footer3.xml']
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name in footers:
                    try:
                        zout.writestr(item, _inject_footer_reference(zin.read(name), rd6_ref))
                    except Exception:
                        zout.writestr(item, zin.read(name))
                else:
                    zout.writestr(item, zin.read(name))


# ── Insulation cert appender ───────────────────────────────────────────────────
def _pdf_to_images(pdf_bytes):
    try:
        import pdfplumber
        images = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                img = page.to_image(resolution=150)
                buf = io.BytesIO()
                img.save(buf, format='PNG')
                images.append(buf.getvalue())
        return images
    except Exception:
        return []

def append_insulation_cert(docx_path, out_path, cert_bytes, cert_filename):
    """
    Replace the existing cert page images (rId11=image3.png, rId12=image4.png)
    with pages from the new insulation certificate.
    This keeps the ZIP structure 100% identical to the template — no new entries,
    no rels changes — so Word on Windows always accepts the result.
    If the cert has only 1 page, image4.png is replaced with a blank/same image.
    If the cert has more than 2 pages, only the first 2 are used.
    """
    # Convert cert to image(s)
    import pathlib
    ext = pathlib.Path(cert_filename).suffix.lower()
    if ext == '.pdf':
        img_list = _pdf_to_images(cert_bytes)  # list of PNG bytes
    elif ext in ('.jpg', '.jpeg'):
        img_list = [cert_bytes]
    else:
        img_list = [cert_bytes]

    # Map the two cert slots: rId11->image3.png, rId12->image4.png
    # Read from the template/current docx to find the actual media paths
    with zipfile.ZipFile(docx_path, 'r') as z:
        rels_tree = etree.fromstring(z.read('word/_rels/document.xml.rels'))

    PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    slot_map = {}   # {rId: media_path_in_zip}
    for r in rels_tree:
        if r.get('Id') in ('rId11', 'rId12'):
            slot_map[r.get('Id')] = 'word/' + r.get('Target', '').lstrip('/')

    if not slot_map:
        # Template has no cert slots — fall back to legacy append
        shutil.copy(docx_path, out_path)
        return

    # Assign images to slots
    slot_images = {}
    rids_ordered = sorted(slot_map.keys())  # rId11 first, rId12 second
    for i, rid in enumerate(rids_ordered):
        if i < len(img_list):
            slot_images[slot_map[rid]] = img_list[i]
        else:
            # Repeat last page for empty slots
            slot_images[slot_map[rid]] = img_list[-1] if img_list else b''

    # Rebuild ZIP replacing the cert image files
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename in slot_images:
                    zout.writestr(item, slot_images[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))


def _append_images(docx_path, out_path, img_list):
    with zipfile.ZipFile(docx_path, 'r') as zin:
        doc_xml  = zin.read('word/document.xml')
        rels_xml = zin.read('word/_rels/document.xml.rels')
        ct_xml   = zin.read('[Content_Types].xml')

    doc_tree  = etree.fromstring(doc_xml)
    rels_tree = etree.fromstring(rels_xml)
    ct_tree   = etree.fromstring(ct_xml)

    existing = [int(r.get('Id','rId0').replace('rId',''))
                for r in rels_tree if r.get('Id','').startswith('rId')]
    next_rid = max(existing, default=30) + 1
    body     = doc_tree.find(f'{{{W}}}body')
    sect_pr  = body.find(f'{{{W}}}sectPr')
    store    = {}

    for i, (img_bytes, img_ext) in enumerate(img_list):
        rid        = f'rId{next_rid+i}'
        media_name = f'media/cert_{i+1}.{img_ext}'
        store[media_name] = img_bytes

        etree.SubElement(rels_tree, f'{{{PKG}}}Relationship', {
            'Id': rid,
            'Type': f'{R}/image',
            'Target': media_name,
        })
        clean_ext = img_ext.lower()
        existing_exts = {e.get('Extension','').lower() for e in ct_tree}
        if clean_ext not in existing_exts:
            etree.SubElement(ct_tree, f'{{{CT}}}Default', {
                'Extension': clean_ext,
                'ContentType': 'image/jpeg' if clean_ext in ('jpg','jpeg') else 'image/png',
            })

        cx = 5624322
        try:
            from PIL import Image as PILImage
            pil = PILImage.open(io.BytesIO(img_bytes))
            cy  = int(cx * pil.size[1] / pil.size[0])
        except Exception:
            cy = 7955694

        para_xml = f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:pPr><w:pageBreakBefore/><w:jc w:val="center"/></w:pPr>
  <w:r><w:drawing><wp:inline>
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:docPr id="{300+i}" name="Cert{i+1}"/>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr>
          <pic:cNvPr id="{400+i}" name="Cert{i+1}"/>
          <pic:cNvPicPr/>
        </pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData></a:graphic>
  </wp:inline></w:drawing></w:r>
</w:p>'''
        para_elem = etree.fromstring(para_xml)
        idx_in_body = list(body).index(sect_pr) if sect_pr is not None else len(list(body))
        body.insert(idx_in_body, para_elem)

    new_doc  = etree.tostring(doc_tree,  xml_declaration=True, encoding='UTF-8', standalone=True)
    new_rels = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    new_ct   = etree.tostring(ct_tree,   xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name == 'word/document.xml':       zout.writestr(item, new_doc)
                elif name == 'word/_rels/document.xml.rels': zout.writestr(item, new_rels)
                elif name == '[Content_Types].xml':   zout.writestr(item, new_ct)
                else:                                 zout.writestr(item, zin.read(name))
            for media_name, img_bytes in store.items():
                zout.writestr(f'word/{media_name}', img_bytes)


# ── Word-safe repack using unpack/pack round-trip ─────────────────────────────

# Scripts are bundled in docx_scripts/ subfolder
_HERE     = Path(__file__).parent
UNPACK_PY = str(_HERE / 'docx_scripts' / 'unpack.py')
PACK_PY   = str(_HERE / 'docx_scripts' / 'pack.py')

def _fix_xml_declaration(data):
    """
    Ensure the XML declaration matches exactly what Word writes:
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n
    lxml omits standalone="yes" and the trailing \r\n — Word on Windows
    requires both. This replaces the declaration in the raw bytes.
    """
    # Remove existing declaration (with or without standalone, with or without newline)
    import re as _re
    data = _re.sub(
        rb'<\?xml[^?]*\?>[\r\n]*',
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        data, count=1
    )
    return data


def _normalize_xml_inplace(docx_path):
    """
    Pure-Python XML normalization — no external scripts required.
    Fixes every issue that causes Microsoft Word on Windows to reject the file:
      1. Missing standalone="yes" and \r\n in XML declaration (critical for Word).
      2. Missing xml:space="preserve" on <w:t> with leading/trailing whitespace.
      3. <a:ext> inside <a:extLst> that incorrectly carry cx/cy attributes.
    Works by reading every XML file in the ZIP, applying fixes, and writing back.
    """
    XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
    W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    A  = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    import io as _io
    buf = _io.BytesIO()
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                    try:
                        tree = etree.fromstring(data)
                        changed = False

                        # Fix 1: xml:space="preserve" on w:t with leading/trailing space
                        for t in tree.iter(f'{{{W}}}t'):
                            txt = t.text or ''
                            if txt != txt.strip():
                                if t.get(XML_SPACE) != 'preserve':
                                    t.set(XML_SPACE, 'preserve')
                                    changed = True

                        # Fix 2: remove cx/cy from a:ext inside a:extLst (not inside a:xfrm)
                        XFRM    = f'{{{A}}}xfrm'
                        EXTLST  = f'{{{A}}}extLst'
                        for ext_elem in tree.iter(f'{{{A}}}ext'):
                            parent = ext_elem.getparent()
                            if parent is not None and parent.tag == EXTLST:
                                if 'cx' in ext_elem.attrib or 'cy' in ext_elem.attrib:
                                    ext_elem.attrib.pop('cx', None)
                                    ext_elem.attrib.pop('cy', None)
                                    changed = True

                        if changed:
                            data = etree.tostring(
                                tree, xml_declaration=True,
                                encoding='UTF-8', standalone=True
                            )
                    except Exception:
                        pass  # leave unchanged if parse fails
                # Always fix XML declaration for ALL xml/rels files we touch
                if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                    data = _fix_xml_declaration(data)
                zout.writestr(item, data)

    # Atomically replace the file
    with open(docx_path, 'wb') as f:
        f.write(buf.getvalue())


def _fix_zip_metadata(docx_path):
    """
    Patch ZIP raw bytes so every entry has:
      create_system = 0  (Windows/MS-DOS)
      flag_bits     = 6  (0b110 - matches exactly what Microsoft Word writes)
    Python's zipfile ignores ZipInfo.flag_bits on write and always uses its own
    value; and sets create_system=3 (Unix) on Linux/Mac. Word on Windows requires
    create_system=0. We patch the raw bytes directly rather than rebuilding the ZIP.
    """
    with open(docx_path, 'rb') as f:
        data = bytearray(f.read())

    i = 0
    while i < len(data) - 4:
        sig = data[i:i+4]
        if sig == b'PK\x03\x04':        # local file header
            data[i+6] = 6; data[i+7] = 0  # flag_bits = 6
            i += 4
        elif sig == b'PK\x01\x02':      # central directory header
            data[i+5] = 0                  # create_system = 0 (Windows)
            data[i+8] = 6; data[i+9] = 0  # flag_bits = 6
            i += 4
        else:
            i += 1

    with open(docx_path, 'wb') as f:
        f.write(bytes(data))


def _repack_for_word(docx_path):
    """
    Produce a Word-safe DOCX in two stages:
      Stage 1 (always runs): pure-Python XML normalization — fixes xml:space and a:ext.
      Stage 2 (best-effort): unpack/pack round-trip via bundled scripts for deep repair.
    Stage 2 uses sys.executable so it works on Windows, Mac, and Linux alike.
    """
    # Stage 1 — subprocess round-trip for structural repair (best-effort)
    scripts_dir = str(_HERE / 'docx_scripts')
    tmp_dir  = tempfile.mkdtemp()
    repacked = docx_path + '.repacked.docx'
    try:
        r1 = subprocess.run(
            [sys.executable, 'unpack.py', docx_path, tmp_dir],
            capture_output=True, timeout=60,
            cwd=scripts_dir
        )
        if r1.returncode == 0:
            r2 = subprocess.run(
                [sys.executable, 'pack.py', tmp_dir, repacked,
                 '--original', docx_path],
                capture_output=True, timeout=60,
                cwd=scripts_dir
            )
            if r2.returncode == 0 and os.path.exists(repacked):
                os.replace(repacked, docx_path)
    except Exception:
        pass
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        if os.path.exists(repacked):
            try: os.unlink(repacked)
            except: pass

    # Stage 2 — pure Python normalizer runs after pack.py
    # Fixes XML declarations (standalone="yes" + CRLF) that pack.py strips out.
    try:
        _normalize_xml_inplace(docx_path)
    except Exception:
        pass

    # Stage 3 — ZIP metadata fix, always runs LAST
    # Sets create_system=0 (Windows) on every ZIP entry.
    # Word on Windows rejects files with create_system=3 (Unix/Linux).
    try:
        _fix_zip_metadata(docx_path)
    except Exception:
        pass


# ── In-memory preparation helpers ──────────────────────────────────────────────

def _prepare_author_sig(template_path, sig_bytes, sig_ext):
    """
    Prepare author signature additions as in-memory bytes dict.
    Returns {filename: bytes} to merge into file_contents.
    """
    result = {}
    clean_ext = sig_ext.lstrip('.').lower()
    media_name = 'word/media/author_sig.{}'.format(clean_ext)
    sig_rid = 'rId21'

    # Compute display size
    try:
        from PIL import Image as PILImage
        pil = PILImage.open(io.BytesIO(sig_bytes))
        w_px, h_px = pil.size
        target_cx = 990600
        target_cy = int(target_cx * h_px / w_px)
    except Exception:
        target_cx, target_cy = 990600, 400050

    # Add media file
    result[media_name] = sig_bytes

    # Update rels to add rId21
    with zipfile.ZipFile(template_path, 'r') as z:
        rels_tree = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        doc_xml   = z.read('word/document.xml')
        ct_tree   = etree.fromstring(z.read('[Content_Types].xml'))

    # Add relationship
    existing_ids = {r.get('Id','') for r in rels_tree}
    if sig_rid not in existing_ids:
        etree.SubElement(rels_tree, f'{{{PKG}}}Relationship', {
            'Id': sig_rid,
            'Type': f'{R}/image',
            'Target': media_name.replace('word/',''),
        })
    new_rels = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    result['word/_rels/document.xml.rels'] = _fix_xml_declaration(new_rels)

    # Add content type
    ct_map = {'jpeg':'image/jpeg','jpg':'image/jpeg','png':'image/png'}
    existing_exts = {e.get('Extension','').lower() for e in ct_tree}
    if clean_ext not in existing_exts:
        etree.SubElement(ct_tree, f'{{{CT}}}Default', {
            'Extension': clean_ext,
            'ContentType': ct_map.get(clean_ext,'image/png'),
        })
    new_ct = etree.tostring(ct_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    result['[Content_Types].xml'] = _fix_xml_declaration(new_ct)

    # Inject anchor into document.xml (already modified, in file_contents)
    # We do this separately since doc_xml may already be modified
    # Return sig info so caller can inject into the already-modified doc_xml
    result['__sig_rid__']    = sig_rid
    result['__sig_cx__']     = str(target_cx)
    result['__sig_cy__']     = str(target_cy)
    return result


def _inject_sig_anchors(tree, sig_rid, target_cx, target_cy, template_path):
    """Inject author sig as floating anchors into the document tree in-place."""
    # Find image paragraphs (those containing rId9 anchor)
    v_offsets = [10160, 6985]
    img_para_indices = []
    all_paras = list(tree.iter(f'{{{W}}}p'))
    for i, p in enumerate(all_paras):
        for anchor in p.iter(f'{{{WP}}}anchor'):
            for b in anchor.iter(f'{{{A}}}blip'):
                if b.get(f'{{{R}}}embed') == 'rId9':
                    img_para_indices.append(i)
                    break

    for idx, para_i in enumerate(img_para_indices[:2]):
        para = all_paras[para_i]
        src_anchor = None
        for anchor in para.iter(f'{{{WP}}}anchor'):
            for b in anchor.iter(f'{{{A}}}blip'):
                if b.get(f'{{{R}}}embed') == 'rId9':
                    src_anchor = anchor; break
            if src_anchor is not None: break
        if src_anchor is None: continue

        new_anchor = copy.deepcopy(src_anchor)
        for b in new_anchor.iter(f'{{{A}}}blip'):
            b.set(f'{{{R}}}embed', sig_rid)
        posH = new_anchor.find(f'{{{WP}}}positionH')
        if posH is not None:
            posH.set('relativeFrom', 'column')
            for child in list(posH): posH.remove(child)
            off = etree.SubElement(posH, f'{{{WP}}}posOffset')
            off.text = '200000'
        posV = new_anchor.find(f'{{{WP}}}positionV')
        if posV is not None:
            v_off = posV.find(f'{{{WP}}}posOffset')
            if v_off is not None:
                v_off.text = str(v_offsets[idx] if idx < len(v_offsets) else 10160)
        ext = new_anchor.find(f'{{{WP}}}extent')
        if ext is not None:
            ext.set('cx', str(target_cx)); ext.set('cy', str(target_cy))
        XFRM = f'{{{A}}}xfrm'
        for xfrm in new_anchor.iter(XFRM):
            for a_ext in xfrm.iter(f'{{{A}}}ext'):
                a_ext.set('cx', str(target_cx)); a_ext.set('cy', str(target_cy))
        doc_pr = new_anchor.find(f'{{{WP}}}docPr')
        if doc_pr is not None:
            doc_pr.set('id', str(500+idx)); doc_pr.set('name', f'AuthorSig{idx+1}')
        r_elem = etree.Element(f'{{{W}}}r')
        rPr = etree.SubElement(r_elem, f'{{{W}}}rPr')
        etree.SubElement(rPr, f'{{{W}}}noProof')
        draw = etree.SubElement(r_elem, f'{{{W}}}drawing')
        draw.append(new_anchor)
        para.insert(0, r_elem)


def _prepare_cert_images(cert_bytes, cert_filename):
    """
    Prepare cert page image replacements as in-memory bytes dict.
    Replaces image3.png (rId11) and image4.png (rId12) — the two cert slots in the template.
    """
    import pathlib
    ext = pathlib.Path(cert_filename).suffix.lower()
    if ext == '.pdf':
        img_list = _pdf_to_images(cert_bytes)
    elif ext in ('.jpg', '.jpeg'):
        img_list = [cert_bytes]
    else:
        img_list = [cert_bytes]

    result = {}
    slots = ['word/media/image3.png', 'word/media/image4.png']
    for i, slot in enumerate(slots):
        if i < len(img_list):
            result[slot] = img_list[i]
        else:
            result[slot] = img_list[-1] if img_list else b''
    return result


# ── Master ZIP writer — preserves ALL template ZIP metadata ────────────────────

def _write_docx_preserving_metadata(template_path, file_contents, output_path):
    """
    Write output DOCX preserving ALL ZIP metadata from the template:
    - extra fields (required by Windows Word for fast memory-mapped access)
    - create_system = 0 (Windows)
    - flag_bits = 6
    - timestamps
    file_contents: dict {filename: bytes} — only files to ADD or REPLACE.
    All other files are copied from template unchanged.
    """
    # Extract extra fields from template local file headers (raw bytes)
    tmpl_raw = open(template_path, 'rb').read()
    tmpl_extra = {}
    off = 0
    while off < len(tmpl_raw) - 4:
        if tmpl_raw[off:off+4] == b'PK\x03\x04':
            fl = struct.unpack_from('<H', tmpl_raw, off+26)[0]
            el = struct.unpack_from('<H', tmpl_raw, off+28)[0]
            fn = tmpl_raw[off+30:off+30+fl].decode('utf-8', 'replace')
            tmpl_extra[fn] = tmpl_raw[off+30+fl:off+30+fl+el]
            off += 30 + fl + el
        else:
            off += 1

    # Build ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        existing = {i.filename for i in zin.infolist()}
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                item.extra = tmpl_extra.get(item.filename, b'')
                data = file_contents.get(item.filename, zin.read(item.filename))
                zout.writestr(item, data)
            # Add new files (not in template)
            for fname, data in file_contents.items():
                if fname not in existing:
                    zi = zipfile.ZipInfo(fname)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zout.writestr(zi, data)

    # Patch raw ZIP headers: flags=6, create_system=0 on every entry
    raw = bytearray(buf.getvalue())
    off = 0
    while off < len(raw) - 4:
        sig = bytes(raw[off:off+4])
        if sig == b'PK\x03\x04':
            raw[off+6] = 6;  raw[off+7] = 0
            fl = struct.unpack_from('<H', raw, off+26)[0]
            el = struct.unpack_from('<H', raw, off+28)[0]
            off += 30 + fl + el
        elif sig == b'PK\x01\x02':
            raw[off+5] = 0
            raw[off+8] = 6;  raw[off+9] = 0
            fl  = struct.unpack_from('<H', raw, off+28)[0]
            el  = struct.unpack_from('<H', raw, off+30)[0]
            cl  = struct.unpack_from('<H', raw, off+32)[0]
            off += 46 + fl + el + cl
        else:
            off += 1

    with open(output_path, 'wb') as f:
        f.write(bytes(raw))
def _add_extra_visit_rows(tree, visits):
    """Appends new rows to the visits table for visits beyond the 4 template slots."""
    if len(visits) <= 4:
        return

    visit_field_tags = {
        '1stVisit_Ref','1stVisit_date','1stVisit_isp','1stVisit_part',
        '2ndVisit_Ref','2ndVisit_Date','2ndVisit_ins','2ndVisit_part',
        '3rdVisit_Ref','3rdVisit_date','3rdVisit_ins','3rdVisit_part',
        '4thVisit_Ref','4thVisit_date','4thVisit_ins','4thVisit_part',
        '5thVisit_Ref','5thVisit_date','5thVisit_ins','5thVisit_part',
        '6thVisit_Ref','6thVisit_date','6thVisit_ins','6thVisit_part',
        '7thVisit_Ref','7thVisit_date','7thVisit_ins','7thVisit_part',
        '8thVisit_Ref','8thVisit_date','8thVisit_ins','8thVisit_part',
        '9thVisit_Ref','9thVisit_date','9thVisit_ins','9thVisit_part',
        '10thVisit_Ref','10thVisit_date','10thVisit_ins','10thVisit_part',
    }
    # Find the visit table
    visit_table = None
    for tbl in tree.iter(f'{{{W}}}tbl'):
        for sdt in tbl.iter(f'{{{W}}}sdt'):
            tag_elem = sdt.find(f'.//{{{W}}}tag')
            if tag_elem is not None and tag_elem.get(f'{{{W}}}val', '') == '1stVisit_Ref':
                visit_table = tbl
                break
        if visit_table is not None:
            break

    if visit_table is None:
        return

    # Find the 4th data row
    data_rows = []
    for tr in visit_table.iter(f'{{{W}}}tr'):
        row_tags = set()
        for sdt in tr.iter(f'{{{W}}}sdt'):
            tag_elem = sdt.find(f'.//{{{W}}}tag')
            if tag_elem is not None:
                row_tags.add(tag_elem.get(f'{{{W}}}val', ''))
        if row_tags & visit_field_tags:
            data_rows.append(tr)

    if not data_rows:
        return

    template_row = data_rows[-1]
    actual_parent = template_row.getparent()
    last_idx = list(actual_parent).index(template_row)

    # ── Extract structural properties only (no content) ──────────────────────
    # Use iter() to find w:tc regardless of SDT wrapping — this is what
    # caused findall() to silently return [] and the old wipe to never run.
    trPr = template_row.find(f'{{{W}}}trPr')
    template_tcPrs = []
    for tc in template_row.iter(f'{{{W}}}tc'):
        tcPr = tc.find(f'{{{W}}}tcPr')
        template_tcPrs.append(copy.deepcopy(tcPr) if tcPr is not None else None)
        if len(template_tcPrs) == 4:
            break

    # ── Build each extra row from scratch — never copies visit content ────────
    for i, visit in enumerate(visits[4:], start=4):
        cell_values = [
            visit.get('ref', ''),
            visit.get('date', ''),
            visit.get('inspector', ''),
            visit.get('part', ''),
        ]

        new_row = etree.Element(f'{{{W}}}tr')
        if trPr is not None:
            new_row.append(copy.deepcopy(trPr))

        for ci in range(4):
            val = cell_values[ci] if ci < len(cell_values) else ''
            tc = etree.SubElement(new_row, f'{{{W}}}tc')

            # Restore cell formatting (width, borders, shading)
            if ci < len(template_tcPrs) and template_tcPrs[ci] is not None:
                tc.append(copy.deepcopy(template_tcPrs[ci]))

            # Fresh paragraph — no old content to bleed through
            p  = etree.SubElement(tc, f'{{{W}}}p')
            r  = etree.SubElement(p,  f'{{{W}}}r')
            rPr = etree.SubElement(r, f'{{{W}}}rPr')
            color_el = etree.SubElement(rPr, f'{{{W}}}color')
            color_el.set(f'{{{W}}}val', BLUE)
            t  = etree.SubElement(r,  f'{{{W}}}t')
            t.text = val
            if val and (val[0] == ' ' or val[-1] == ' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        actual_parent.insert(last_idx + 1 + (i - 4), new_row)

# ── Main entry point ───────────────────────────────────────────────────────────
def generate_rd6(template_path, output_path, data, visits,
                 provided_doc_keys,
                 signature_bytes=None, signature_ext='png',
                 insulation_bytes=None, insulation_filename='cert.pdf',
                 extra_cert_bytes=None):

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        tmp_path = tmp.name
    shutil.copy(template_path, tmp_path)

    # Parse
    with zipfile.ZipFile(tmp_path, 'r') as z:
        doc_xml = z.read('word/document.xml')
    tree = etree.fromstring(doc_xml)

    # Derived values
    eng_full  = data.get('eng_full', '')
    eng_upper = eng_full.upper()
    idi_no    = str(data.get('idi_no','')).replace('.0','')
    taw_pol   = str(data.get('taw_pol','')).replace('.0','')
    today     = datetime.today()
    issue_date = data.get('issue_date', f"{today.day}/{today.month}/{today.year}")

    sum_raw = str(data.get('sum_insured','')).replace('.0','').replace(',','')
    try:    sum_fmt = f"{int(sum_raw):,}"
    except: sum_fmt = sum_raw

    rd6_ref  = data.get('rd6_ref','')
    rd0_ref  = data.get('rd0_ref','')
    rd0_date = data.get('rd0_date','')

    # Visit controls (4 slots)
    visit_controls = {}
    visit_key_map = [
        ('1stVisit_Ref','1stVisit_date','1stVisit_isp','1stVisit_part'),
        ('2ndVisit_Ref','2ndVisit_Date','2ndVisit_ins','2ndVisit_part'),
        ('3rdVisit_Ref','3rdVisit_date','3rdVisit_ins','3rdVisit_part'),
        ('4thVisit_Ref','4thVisit_date','4thVisit_ins','4thVisit_part'),
    ]
    for i, (rk, dk, ik, pk) in enumerate(visit_key_map):
        v = visits[i] if i < len(visits) else {}
        visit_controls[rk] = v.get('ref','')
        visit_controls[dk] = v.get('date','')
        visit_controls[ik] = v.get('inspector','')
        visit_controls[pk] = v.get('part','')

    missing_text = build_missing_doc_text(provided_doc_keys)

    # Fill content controls
    fill_all_controls(tree, {
        'RD6Reference':     rd6_ref,
        'Reference':        rd6_ref,
        'ReportDate':       issue_date,
        'No_Buildings':     data.get('no_buildings','1'),
        'ProjectTitle':     data.get('project_title',''),
        'Address':          data.get('address',''),
        'Owner':            data.get('owner',''),
        'ProjectType':      data.get('building_type','Residential'),
        'StartDate':        data.get('start_date',''),
        'FinishDate':       data.get('finish_date',''),
        'Last_VisitDate':   data.get('last_visit_date',''),
        'OCCDate':          data.get('occ_date',''),
        'TotCostRD0':       sum_fmt,
        'Act_Cost':         sum_fmt,
        'RD0_Ref':          rd0_ref,
        'RD0Date':          rd0_date,
        'RD0_Date':         rd0_date,
        'RoofTestDate':     data.get('roof_test_date',''),
        'IDI_No':           idi_no,
        'ReservationsNote': data.get('reservations_note',''),
        'MissingDoc':       missing_text,
        **visit_controls,
    })

    # Replace hardcoded engineer fields
    # Template has 'Mohamed Mossad' as default — replace with actual engineer
    eng_phone = data.get('eng_phone','')
    eng_email = data.get('eng_email','')
    replace_hardcoded(tree, {
        'Mohamed Mossad':             eng_full,
        'MOHAMED MOSSAD':             eng_upper,
        '00966-546380314':            eng_phone,
        '+966546380314':              eng_phone,
        'Mohamed.mossad@socotec.com': eng_email,
        'mohamed.mossad@socotec.com': eng_email,
        "Master\u2019s Degree":       data.get('eng_degree','Civil Engineering Bachelor'),
        'Civil Engineering Bachelor': data.get('eng_degree','Civil Engineering Bachelor'),
        'Senior':                     data.get('eng_phase','Senior'),
        'Civil':                      data.get('eng_speciality','Civil'),
    })

    # Serialize modified document XML
    new_doc_xml = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    new_doc_xml = _fix_xml_declaration(new_doc_xml)

    # Collect all file changes in memory
    file_contents = {'word/document.xml': new_doc_xml}

    # Inject modified footers
    with zipfile.ZipFile(tmp_path, 'r') as z:
        for fname in ['word/footer1.xml','word/footer2.xml','word/footer3.xml']:
            try:
                footer_bytes = z.read(fname)
                file_contents[fname] = _fix_xml_declaration(
                    _inject_footer_reference(footer_bytes, rd6_ref))
            except Exception:
                pass

    # Inject author signature
    if signature_bytes:
        sig_info = _prepare_author_sig(tmp_path, signature_bytes, signature_ext)
        # Extract metadata, don't add internal keys to file_contents
        sig_rid = sig_info.pop('__sig_rid__', 'rId21')
        sig_cx  = int(sig_info.pop('__sig_cx__', 990600))
        sig_cy  = int(sig_info.pop('__sig_cy__', 400050))
        file_contents.update(sig_info)  # adds rels, content_types, media file
        # Inject the anchor into the already-modified document XML
        doc_tree = etree.fromstring(file_contents['word/document.xml'])
        _inject_sig_anchors(doc_tree, sig_rid, sig_cx, sig_cy, tmp_path)
        new_with_sig = etree.tostring(doc_tree, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)
        file_contents['word/document.xml'] = _fix_xml_declaration(new_with_sig)

# Replace cert page images (swap image3.png and image4.png)
    if insulation_bytes:
        cert_images = _prepare_cert_images(insulation_bytes, insulation_filename)
        file_contents.update(cert_images)

    # Expand visit table for visits beyond 4
    doc_tree_final = etree.fromstring(file_contents['word/document.xml'])
    _add_extra_visit_rows(doc_tree_final, visits)
    final_doc = etree.tostring(doc_tree_final, xml_declaration=True,
                               encoding='UTF-8', standalone=True)
    file_contents['word/document.xml'] = _fix_xml_declaration(final_doc)

    # Write everything at once using the template as ZIP skeleton
    _write_docx_preserving_metadata(tmp_path, file_contents, output_path)
    os.unlink(tmp_path)

    # Append extra certificates (cost letter, contractor letter, supervision letter)
    if extra_cert_bytes:
        img_list = []
        for cert_b, cert_ext in extra_cert_bytes:
            if cert_ext.lower() == 'pdf':
                pages = _pdf_to_images(cert_b)
                img_list.extend((p, 'png') for p in pages)
            else:
                img_list.append((cert_b, cert_ext.lstrip('.')))
        if img_list:
            tmp_extra = output_path + '.extra.docx'
            _append_images(output_path, tmp_extra, img_list)
            os.replace(tmp_extra, output_path)

    return output_path
