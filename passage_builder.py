import json
import re
import os
import tempfile
import zipfile
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def xml_escape(s):
    if s is None:
        return ""
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;'))

def generate_passage_docx(passage_num, title, subtitle, topic_focus, paragraphs, vocab_chunk, output_paths):
    """
    Builds a beautifully styled Word document (.docx) with:
    - Title, Subtitle, Metadata table
    - Interactive Guide (single click)
    - Passage paragraphs with target words colored, single-click hyperlinked with ScreenTips, and Word Comments
    - Vocabulary Reference & Study Guide table at the end with [↩ Return] links
    """
    # Build vocabulary map
    vocab_map = {}
    for i, item in enumerate(vocab_chunk):
        w = item['word'].strip()
        clean_k = re.sub(r'[^a-zA-Z0-9_]', '_', w.lower())
        vocab_map[w.lower()] = {
            'id': i,
            'word': w,
            'explanation': item.get('explanation', '').strip(),
            'clean_key': clean_k
        }

    # Verify all words exist in text
    full_text = " ".join(paragraphs)
    missing = []
    for k, v in vocab_map.items():
        pattern = r'\b' + re.escape(v['word']) + r'\b'
        if not re.search(pattern, full_text, re.IGNORECASE):
            missing.append(v['word'])

    if missing:
        raise ValueError(f"Passage {passage_num} is missing {len(missing)} vocabulary words: {missing}")

    # Extract in-context sentences
    sentences = []
    for p_idx, p_text in enumerate(paragraphs):
        raw_sents = re.split(r'(?<=[.!?])\s+', p_text)
        for s in raw_sents:
            sentences.append((p_idx, s.strip()))

    for k, v in vocab_map.items():
        v['sentence'] = ""
        for p_idx, s in sentences:
            if re.search(r'\b' + re.escape(k) + r'\b', s, re.IGNORECASE):
                v['sentence'] = s
                break

    sorted_words = sorted(vocab_map.keys(), key=len, reverse=True)
    pattern = re.compile(r'\b(' + '|'.join(re.escape(w) for w in sorted_words) + r')\b', re.IGNORECASE)

    doc = docx.Document()

    # 1 inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base font
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Calibri'
    normal_font.size = Pt(11)
    normal_font.color.rgb = RGBColor(0x24, 0x29, 0x2F)

    # 1. Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(f"Taiwan GRE Master Reader: Passage {passage_num:03d}")
    r_title.bold = True
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    # 2. Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(f"{title}: {subtitle}")
    r_sub.italic = True
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

    # 3. Metadata Table
    meta_table = doc.add_table(rows=1, cols=3)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    meta_cells = meta_table.rows[0].cells
    widths = [Inches(2.1), Inches(2.3), Inches(2.1)]
    for i, cell in enumerate(meta_cells):
        cell.width = widths[i]
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F6F8FA"/>')
        tcPr.append(shd)

    def style_meta_cell(cell, label, val):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        r1 = p.add_run(label + "\n")
        r1.font.size = Pt(9)
        r1.font.color.rgb = RGBColor(0x6E, 0x77, 0x81)
        r1.bold = True
        r2 = p.add_run(val)
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
        r2.bold = True

    style_meta_cell(meta_cells[0], "TARGET VOCABULARY", f"{len(vocab_chunk)} High-Freq GRE Words")
    style_meta_cell(meta_cells[1], "THEMATIC FOCUS", topic_focus)
    style_meta_cell(meta_cells[2], "PERSPECTIVE", "Democratic Sovereignty & Resilience")

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 4. Instructions Callout Box (Single Click)
    callout_table = doc.add_table(rows=1, cols=1)
    callout_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    callout_cell = callout_table.rows[0].cells[0]
    callout_cell.width = Inches(6.5)
    tcPr = callout_cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EBF5FF"/>')
    tcPr.append(shd)

    borders = parse_xml(f'''
    <w:tcBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="6" w:space="0" w:color="0969DA"/>
        <w:left w:val="single" w:sz="24" w:space="0" w:color="0969DA"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0969DA"/>
        <w:right w:val="single" w:sz="6" w:space="0" w:color="0969DA"/>
    </w:tcBorders>
    ''')
    tcPr.append(borders)

    p_inst = callout_cell.paragraphs[0]
    p_inst.paragraph_format.space_before = Pt(6)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst_h = p_inst.add_run("💡 Interactive Vocabulary Guide:")
    r_inst_h.bold = True
    r_inst_h.font.size = Pt(10.5)
    r_inst_h.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)

    instructions = [
        ("• Single Click any highlighted word: ", "Instantly jumps to the word's detailed definition and contextual sentence in the Glossary at the end of the document."),
        ("• Hover over any highlighted word: ", "A ScreenTip popup immediately displays the GRE definition directly on screen."),
        ("• Word Review Pane: ", "Native Word Comments are also available in the right margin for side-by-side reviewing."),
        ("• Return to Passage: ", "In the Glossary, simply click '[↩ Return]' to jump right back to where you left off.")
    ]

    for label, desc in instructions:
        p_item = callout_cell.add_paragraph()
        p_item.paragraph_format.space_before = Pt(2)
        p_item.paragraph_format.space_after = Pt(2)
        r_l = p_item.add_run(label)
        r_l.bold = True
        r_l.font.size = Pt(9.5)
        r_l.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
        r_d = p_item.add_run(desc)
        r_d.font.size = Pt(9.5)
        r_d.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5. Heading for Passage
    h_pass = doc.add_paragraph()
    h_pass.paragraph_format.space_before = Pt(8)
    h_pass.paragraph_format.space_after = Pt(8)
    r_hp = h_pass.add_run(f"Reading Passage {passage_num:03d}: {title}")
    r_hp.bold = True
    r_hp.font.size = Pt(15)
    r_hp.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    # 6. Render Paragraphs
    comment_entries = []

    for p_num, p_text in enumerate(paragraphs):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(9)
        p.paragraph_format.line_spacing = 1.2
        
        last_end = 0
        for match in pattern.finditer(p_text):
            start, end = match.span()
            matched_text = match.group(0)
            word_lower = matched_text.lower()
            info = vocab_map[word_lower]
            
            if start > last_end:
                p.add_run(p_text[last_end:start])
            
            w_id = info['id']
            w_word = info['word']
            w_clean_key = info['clean_key']
            w_expl = info['explanation']
            
            # Bookmark for return: text_<clean_key>
            bm_id_text = str(1000 + w_id)
            p._p.append(parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id_text}" w:name="text_{w_clean_key}"/>'))
            p._p.append(parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id_text}"/>'))
            
            # Comment Range Start
            c_id_str = str(w_id)
            p._p.append(parse_xml(f'<w:commentRangeStart {nsdecls("w")} w:id="{c_id_str}"/>'))
            
            # Hyperlink with ScreenTip and Anchor
            tooltip_text = f"GRE Definition: {w_expl}"
            safe_tooltip = xml_escape(tooltip_text)
            safe_matched = xml_escape(matched_text)
            
            hl_xml = f'''
            <w:hyperlink {nsdecls("w")} w:anchor="vocab_{w_clean_key}" w:tooltip="{safe_tooltip}">
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:color w:val="0969DA"/>
                        <w:u w:val="single"/>
                    </w:rPr>
                    <w:t xml:space="preserve">{safe_matched}</w:t>
                </w:r>
            </w:hyperlink>
            '''
            p._p.append(parse_xml(hl_xml))
            
            # Comment Range End & Reference
            p._p.append(parse_xml(f'<w:commentRangeEnd {nsdecls("w")} w:id="{c_id_str}"/>'))
            p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{c_id_str}"/></w:r>'))
            
            comment_entries.append((c_id_str, w_word, w_expl))
            last_end = end
        
        if last_end < len(p_text):
            p.add_run(p_text[last_end:])

    # 7. Page Break
    doc.add_page_break()

    # 8. Vocabulary Reference Guide Heading
    h_vocab = doc.add_paragraph()
    h_vocab.paragraph_format.space_before = Pt(8)
    h_vocab.paragraph_format.space_after = Pt(4)
    r_hv = h_vocab.add_run(f"Vocabulary Reference & Study Guide (Passage {passage_num:03d})")
    r_hv.bold = True
    r_hv.font.size = Pt(18)
    r_hv.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_vdesc = doc.add_paragraph()
    p_vdesc.paragraph_format.space_after = Pt(14)
    r_vd = p_vdesc.add_run(f"Alphabetical listing of all {len(vocab_chunk)} high-frequency GRE words featured in Passage {passage_num:03d}, including authoritative GRE definitions, in-context usage, and return navigation links.")
    r_vd.italic = True
    r_vd.font.size = Pt(10)
    r_vd.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

    # 9. Vocabulary Table
    vocab_table = doc.add_table(rows=1, cols=4)
    vocab_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    vocab_table.autofit = False

    col_widths = [Inches(1.3), Inches(2.2), Inches(2.0), Inches(1.0)]

    hdr_cells = vocab_table.rows[0].cells
    hdr_titles = ["GRE Word", "GRE Definition", "In-Context Passage Sentence", "Action"]
    for i, cell in enumerate(hdr_cells):
        cell.width = col_widths[i]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(5)
        p.paragraph_format.space_after = Pt(5)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 3) else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(hdr_titles[i])
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1B365D"/>')
        tcPr.append(shd)

    sorted_vocab_items = sorted(vocab_map.values(), key=lambda x: x['word'].lower())

    for row_idx, item in enumerate(sorted_vocab_items):
        row_cells = vocab_table.add_row().cells
        bg_color = "F6F8FA" if row_idx % 2 == 1 else "FFFFFF"
        
        for i, cell in enumerate(row_cells):
            cell.width = col_widths[i]
            tcPr = cell._tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            tcPr.append(shd)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Cell 0: Word + Bookmark
        p0 = row_cells[0].paragraphs[0]
        p0.paragraph_format.space_before = Pt(4)
        p0.paragraph_format.space_after = Pt(4)
        bm_id_vocab = str(2000 + item['id'])
        p0._p.append(parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id_vocab}" w:name="vocab_{item["clean_key"]}"/>'))
        r_w = p0.add_run(item['word'])
        r_w.bold = True
        r_w.font.size = Pt(10)
        r_w.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
        p0._p.append(parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id_vocab}"/>'))
        
        # Cell 1: Definition
        p1 = row_cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(4)
        p1.paragraph_format.space_after = Pt(4)
        r_def = p1.add_run(item['explanation'])
        r_def.font.size = Pt(9.5)
        r_def.font.color.rgb = RGBColor(0x24, 0x29, 0x2F)
        
        # Cell 2: Sentence
        p2 = row_cells[2].paragraphs[0]
        p2.paragraph_format.space_before = Pt(4)
        p2.paragraph_format.space_after = Pt(4)
        sent_text = item['sentence']
        w_match = re.search(r'\b(' + re.escape(item['word']) + r')\b', sent_text, re.IGNORECASE)
        if w_match:
            s_start, s_end = w_match.span()
            p2.add_run(sent_text[:s_start])
            r_hw = p2.add_run(sent_text[s_start:s_end])
            r_hw.bold = True
            r_hw.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
            p2.add_run(sent_text[s_end:])
        else:
            p2.add_run(sent_text)
        for r in p2.runs:
            r.font.size = Pt(8.5)
            r.italic = True
        
        # Cell 3: Action Link
        p3 = row_cells[3].paragraphs[0]
        p3.paragraph_format.space_before = Pt(4)
        p3.paragraph_format.space_after = Pt(4)
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return_hl_xml = f'''
        <w:hyperlink {nsdecls("w")} w:anchor="text_{item["clean_key"]}" w:tooltip="Return to Passage">
            <w:r>
                <w:rPr>
                    <w:color w:val="0969DA"/>
                    <w:u w:val="single"/>
                </w:rPr>
                <w:t>↩ Return</w:t>
            </w:r>
        </w:hyperlink>
        '''
        p3._p.append(parse_xml(return_hl_xml))

    temp_doc = tempfile.mktemp(suffix=".docx")
    doc.save(temp_doc)

    comments_xml_parts = ['''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">''']

    for c_id, c_word, c_expl in comment_entries:
        safe_word = xml_escape(c_word)
        safe_expl = xml_escape(c_expl)
        comment_item = f'''  <w:comment w:id="{c_id}" w:author="GRE Prep" w:date="2026-10-05T22:00:00Z" w:initials="GRE">
        <w:p>
          <w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
          <w:r><w:rPr><w:b/></w:rPr><w:t>{safe_word}: </w:t></w:r>
          <w:r><w:t>{safe_expl}</w:t></w:r>
        </w:p>
      </w:comment>'''
        comments_xml_parts.append(comment_item)

    comments_xml_parts.append('</w:comments>')
    comments_xml_content = "\n".join(comments_xml_parts)

    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(temp_doc, 'r') as zin:
        zin.extractall(temp_dir)

    with open(os.path.join(temp_dir, 'word', 'comments.xml'), 'w', encoding='utf-8') as f:
        f.write(comments_xml_content)

    rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels_content = f.read()

    rel_entry = '<Relationship Id="rIdCommentList01" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
    rels_content = rels_content.replace('</Relationships>', rel_entry + '</Relationships>')
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(rels_content)

    ct_path = os.path.join(temp_dir, '[Content_Types].xml')
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct_content = f.read()
    ct_entry = '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    ct_content = ct_content.replace('</Types>', ct_entry + '</Types>')
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct_content)

    for p in output_paths:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, temp_dir)
                    zout.write(full_path, rel_path)
        print(f"Generated: {p}")

print("passage_builder module successfully installed!")
