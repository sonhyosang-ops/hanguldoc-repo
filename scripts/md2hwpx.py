#!/usr/bin/env python3
"""Convert a Markdown file to HWPX (Hangul Word Processor Open XML) format.

HWPX is a ZIP-based open document format (OWPML, KS X 6101) used by
Hancom's Hangul Word Processor. This script creates a valid HWPX file
by building on the Skeleton.hwpx template structure.

Native HWPX table structure (derived from real HWPX file analysis):
  hp:p > hp:run > hp:tbl > hp:tr > hp:tc > hp:subList > hp:p > hp:run > hp:t

Usage:
    python md2hwpx.py <input.md> <output.hwpx> [--skeleton <Skeleton.hwpx>]
    python md2hwpx.py <input.md> <output.hwpx> --title "My Title" --author "Author"
"""

import zipfile
import re
import os
import sys
import argparse
import tempfile
import shutil
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

# Page geometry (HWPUNIT = 1/7200 inch)
TABLE_WIDTH = 42000   # slightly less than text area (42520) for outer margins
ROW_HEIGHT = 850      # default row height (~3mm)

# HWPX namespace declarations
NS_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

# ID counters
_pid = [1000000000]
_tbl_z = [0]

def next_pid():
    v = _pid[0]; _pid[0] += 1; return v

def next_z():
    v = _tbl_z[0]; _tbl_z[0] += 1; return v

def read_file(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


# ── Character Property Builder ──────────────────────────────────────

def make_char_pr(cid, height, color="#000000", bold=False, italic=False, font_id=0):
    # Bold/italic are child elements (<hh:bold/>, <hh:italic/>) placed between offset and underline
    style_elements = ''
    if bold:   style_elements += '<hh:bold/>'
    if italic: style_elements += '<hh:italic/>'
    return (
        f'<hh:charPr id="{cid}" height="{height}" textColor="{color}" '
        f'shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">'
        f'<hh:fontRef hangul="{font_id}" latin="{font_id}" hanja="{font_id}" '
        f'japanese="{font_id}" other="{font_id}" symbol="{font_id}" user="{font_id}"/>'
        f'<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" '
        f'other="100" symbol="100" user="100"/>'
        f'<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" '
        f'other="0" symbol="0" user="0"/>'
        f'<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" '
        f'other="100" symbol="100" user="100"/>'
        f'<hh:offset hangul="0" latin="0" hanja="0" japanese="0" '
        f'other="0" symbol="0" user="0"/>'
        f'{style_elements}'
        f'<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        f'<hh:strikeout shape="NONE" color="#000000"/>'
        f'<hh:outline type="NONE"/>'
        f'<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
        f'</hh:charPr>'
    )


# ── BorderFill for tables ───────────────────────────────────────────

BORDER_FILL_TABLE_CELL = (
    '<hh:borderFill id="3" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
    '</hh:borderFill>'
)

BORDER_FILL_TABLE_HEADER = (
    '<hh:borderFill id="4" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
    '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
    '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
    '<hc:fillBrush><hc:winBrush faceColor="#E8E8E8" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
    '</hh:borderFill>'
)


# ── Font definitions ────────────────────────────────────────────────

# Font IDs: 0=함초롬돋움, 1=함초롬바탕, 2=나눔고딕, 3=나눔고딕 ExtraBold, 4=나눔고딕 Bold
FONT_NANUM_GOTHIC = (
    '<hh:font id="2" face="\ub098\ub214\uace0\ub515" type="TTF" isEmbedded="0">'
    '<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" '
    'contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>'
    '</hh:font>'
)

FONT_NANUM_GOTHIC_EB = (
    '<hh:font id="3" face="\ub098\ub214\uace0\ub515 ExtraBold" type="TTF" isEmbedded="0">'
    '<hh:typeInfo familyType="FCAT_GOTHIC" weight="9" proportion="4" '
    'contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>'
    '</hh:font>'
)

FONT_NANUM_GOTHIC_BOLD = (
    '<hh:font id="4" face="\ub098\ub214\uace0\ub515 Bold" type="TTF" isEmbedded="0">'
    '<hh:typeInfo familyType="FCAT_GOTHIC" weight="7" proportion="4" '
    'contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>'
    '</hh:font>'
)


# ── Header XML Builder ──────────────────────────────────────────────

def build_header_xml(skeleton_dir):
    header = read_file(f'{skeleton_dir}/Contents/header.xml')

    # 1. Add 나눔고딕, 나눔고딕 ExtraBold, 나눔고딕 Bold fonts (id=2,3,4) to each fontface
    header = header.replace('fontCnt="2"', 'fontCnt="5"')
    header = header.replace(
        '</hh:fontface>',
        f'{FONT_NANUM_GOTHIC}{FONT_NANUM_GOTHIC_EB}{FONT_NANUM_GOTHIC_BOLD}</hh:fontface>'
    )

    # 2. Modify default charPr id=0: 11pt body with 나눔고딕 (font_id=2)
    header = re.sub(
        r'(<hh:charPr id="0" height=")1000(")',
        r'\g<1>1100\2',
        header
    )
    header = re.sub(
        r'(<hh:charPr id="0"[^>]*>)\s*<hh:fontRef hangul="1" latin="1" hanja="1" '
        r'japanese="1" other="1" symbol="1" user="1"/>',
        r'\1<hh:fontRef hangul="2" latin="2" hanja="2" '
        r'japanese="2" other="2" symbol="2" user="2"/>',
        header
    )

    # 3. Set paragraph spacing: 6pt (600 HWPUNIT) space after all paragraphs
    header = header.replace(
        '<hc:next value="0" unit="HWPUNIT"/>',
        '<hc:next value="600" unit="HWPUNIT"/>'
    )

    # 3. Add borderFills for table (id=3 data cells, id=4 header cells with gray bg)
    header = header.replace(
        '<hh:borderFills itemCnt="2">',
        '<hh:borderFills itemCnt="4">'
    )
    header = header.replace(
        '</hh:borderFills>',
        f'{BORDER_FILL_TABLE_CELL}{BORDER_FILL_TABLE_HEADER}</hh:borderFills>'
    )

    # 4. Add character properties (id 7-14)
    # Skeleton has 7 charPrs (id 0-6), we add 8 more (id 7-14)
    # font_id=2: 나눔고딕, font_id=3: 나눔고딕 ExtraBold
    new_cps = ''.join([
        make_char_pr(7,  2400, "#000000", bold=True,  font_id=3),   # H1 (24pt, 나눔고딕 ExtraBold)
        make_char_pr(8,  1300, "#000000", bold=True,  font_id=3),   # H2 (13pt, 나눔고딕 ExtraBold)
        make_char_pr(9,  1200, "#000000", bold=True,  font_id=3),   # H3 (12pt, 나눔고딕 ExtraBold)
        make_char_pr(10, 1100, "#000000", bold=True,  font_id=4),   # Bold body (11pt, 나눔고딕 Bold)
        make_char_pr(11, 900,  "#000000", bold=False, font_id=2),   # Code (9pt, 나눔고딕)
        make_char_pr(12, 1000, "#000000", bold=True,  font_id=4),   # Table header (10pt, 나눔고딕 Bold)
        make_char_pr(13, 1050, "#000000", bold=False, italic=True, font_id=2),  # Blockquote (10.5pt, 나눔고딕)
        make_char_pr(14, 1100, "#000000", bold=True,  font_id=3),   # H4 (11pt, 나눔고딕 ExtraBold)
        make_char_pr(15, 1000, "#000000", bold=False, font_id=2),  # Table data (10pt, 나눔고딕)
        make_char_pr(16, 1200, "#000000", bold=False, font_id=2),  # Cover page info (12pt, 나눔고딕)
        make_char_pr(17, 800,  "#000000", bold=False, font_id=2),  # Source/출처 (8pt, 나눔고딕)
        make_char_pr(18, 900,  "#666666", bold=False, italic=True, font_id=2),  # Page header (9pt, 진한회색, 이탤릭)
        make_char_pr(19, 900,  "#000000", bold=False, font_id=2),  # Page footer (9pt, 나눔고딕)
    ])
    header = header.replace(
        '<hh:charProperties itemCnt="7">',
        '<hh:charProperties itemCnt="20">'
    )
    header = header.replace(
        '</hh:charProperties>',
        f'{new_cps}</hh:charProperties>'
    )

    # 6. Add paraPr id=20 for code blocks (0pt after spacing)
    code_para_pr = (
        '<hh:paraPr id="20" tabPrIDRef="0" condense="0" fontLineHeight="0" '
        'snapToGrid="1" suppressLineNumbers="0" checked="0" textDir="LTR">'
        '<hh:align horizontal="LEFT" vertical="BASELINE"/>'
        '<hh:heading type="NONE" idRef="0" level="0"/>'
        '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" '
        'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        '<hp:switch>'
        '<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:case><hp:default>'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:default></hp:switch>'
        '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
        'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        '</hh:paraPr>'
    )
    # 7. Add paraPr id=21 for bullet lists (justify + hanging indent 15pt)
    bullet_para_pr = (
        '<hh:paraPr id="21" tabPrIDRef="0" condense="0" fontLineHeight="0" '
        'snapToGrid="1" suppressLineNumbers="0" checked="0" textDir="LTR">'
        '<hh:align horizontal="JUSTIFY" vertical="BASELINE"/>'
        '<hh:heading type="NONE" idRef="0" level="0"/>'
        '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" '
        'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        '<hp:switch>'
        '<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        '<hh:margin><hc:intent value="-2000" unit="HWPUNIT"/>'
        '<hc:left value="2000" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:case><hp:default>'
        '<hh:margin><hc:intent value="-2000" unit="HWPUNIT"/>'
        '<hc:left value="2000" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:default></hp:switch>'
        '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
        'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        '</hh:paraPr>'
    )

    # 8. Add paraPr id=22 for page header (right-aligned, 0pt spacing)
    header_para_pr = (
        '<hh:paraPr id="22" tabPrIDRef="0" condense="0" fontLineHeight="0" '
        'snapToGrid="1" suppressLineNumbers="0" checked="0" textDir="LTR">'
        '<hh:align horizontal="RIGHT" vertical="BASELINE"/>'
        '<hh:heading type="NONE" idRef="0" level="0"/>'
        '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" '
        'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        '<hp:switch>'
        '<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:case><hp:default>'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:default></hp:switch>'
        '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
        'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        '</hh:paraPr>'
    )

    # 9. Add paraPr id=23 for page footer (center-aligned, 0pt spacing)
    footer_para_pr = (
        '<hh:paraPr id="23" tabPrIDRef="0" condense="0" fontLineHeight="0" '
        'snapToGrid="1" suppressLineNumbers="0" checked="0" textDir="LTR">'
        '<hh:align horizontal="CENTER" vertical="BASELINE"/>'
        '<hh:heading type="NONE" idRef="0" level="0"/>'
        '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" '
        'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        '<hp:switch>'
        '<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:case><hp:default>'
        '<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
        '<hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>'
        '<hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/>'
        '</hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
        '</hp:default></hp:switch>'
        '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
        'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        '</hh:paraPr>'
    )

    header = header.replace(
        '<hh:paraProperties itemCnt="20">',
        '<hh:paraProperties itemCnt="24">'
    )
    header = header.replace(
        '</hh:paraProperties>',
        f'{code_para_pr}{bullet_para_pr}{header_para_pr}{footer_para_pr}</hh:paraProperties>'
    )

    return header


# ── Heading Numbering ──────────────────────────────────────────────

def strip_heading_number(text):
    """Strip existing numbering prefix from heading text."""
    text = re.sub(r'^[\d]+\.[\d]+\.[\d]+\s+', '', text)   # 1.1.1
    text = re.sub(r'^[\d]+\.[\d]+\s+', '', text)           # 1.1
    text = re.sub(r'^[\d]+\.\s+', '', text)                # 1.
    text = re.sub(r'^[A-Z]+-\d+\.\s*', '', text)           # CRIT-001.
    return text.strip()


def preprocess_markdown(text):
    """Preprocess markdown text: number headings, convert numbered lists, remove excess blanks.

    Modifies the raw markdown text directly:
    1. Strips existing heading numbers and applies hierarchical numbering
    2. Converts numbered lists (1., 2., ...) in body to bullet lists (- )
    3. Removes horizontal rules (---)
    4. Removes blank lines immediately before/after headings
    5. Removes consecutive blank lines (keeps at most one)

    Code blocks (``` ... ```) are preserved without modification.

    Returns the modified markdown text.
    """
    lines = text.split('\n')

    # Pass 1: Number headings, convert numbered lists, skip code blocks
    h2 = h3 = h4 = 0
    numbered = []
    in_code = False
    for line in lines:
        if line.strip().startswith('```'):
            in_code = not in_code
            numbered.append(line)
            continue
        if in_code:
            numbered.append(line)
            continue

        m = re.match(r'^(#{1,4})\s+(.+)$', line)
        if m:
            level = len(m.group(1))
            title = strip_heading_number(m.group(2).strip())
            if level == 2:
                h2 += 1; h3 = 0; h4 = 0
                title = f'{h2}. {title}'
            elif level == 3:
                h3 += 1; h4 = 0
                title = f'{h2}.{h3} {title}'
            elif level == 4:
                h4 += 1
                title = f'{h2}.{h3}.{h4} {title}'
            numbered.append(f'{"#" * level} {title}')
        elif re.match(r'^---+\s*$', line):
            continue  # Remove horizontal rules
        elif re.match(r'^\d+\.\s', line):
            numbered.append(re.sub(r'^\d+\.\s+', '- ', line))
        else:
            numbered.append(line)

    # Pass 2: Remove blank lines before/after headings and consecutive blanks (skip code blocks)
    result = []
    in_code = False
    for line in numbered:
        if line.strip().startswith('```'):
            in_code = not in_code
            result.append(line)
            continue
        if in_code:
            result.append(line)
            continue

        is_blank = not line.strip()
        is_heading = bool(re.match(r'^#{1,4}\s+', line))
        if is_blank and result:
            prev = result[-1]
            if not prev.strip():
                continue  # skip consecutive blank
            if re.match(r'^#{1,4}\s+', prev):
                continue  # skip blank after heading
        if is_heading and result and not result[-1].strip():
            result.pop()  # remove blank line before heading
        result.append(line)

    return '\n'.join(result)


# ── Markdown Parser ─────────────────────────────────────────────────

def parse_markdown(text):
    """Parse markdown text into a list of (type, data) tuples.

    Supported elements:
    - ('heading', (level, text))     - # to #### headings
    - ('paragraph', text)            - plain text
    - ('blockquote', [lines])        - > quoted text
    - ('bullet', [items])            - - or * list items
    - ('numbered', [items])          - 1. 2. 3. list items
    - ('code', [lines])              - ```code blocks```
    - ('table', [[cells], ...])      - | col | col | tables
    - ('hr', None)                   - --- horizontal rules
    """
    elements = []
    lines = text.split('\n')
    i, n = 0, len(lines)

    while i < n:
        line = lines[i]

        if not line.strip():
            elements.append(('empty', None))
            i += 1; continue

        # Code block
        if line.strip().startswith('```'):
            code_lines = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i]); i += 1
            if i < n: i += 1
            elements.append(('code', code_lines)); continue

        # Heading
        m = re.match(r'^(#{1,4})\s+(.+)$', line)
        if m:
            elements.append(('heading', (len(m.group(1)), m.group(2).strip())))
            i += 1; continue

        # Horizontal rule
        if re.match(r'^---+\s*$', line):
            elements.append(('hr', None)); i += 1; continue

        # Table
        if line.strip().startswith('|'):
            tlines = []
            while i < n and lines[i].strip().startswith('|'):
                tlines.append(lines[i]); i += 1
            rows = []
            for tl in tlines:
                cells = [c.strip() for c in tl.strip().strip('|').split('|')]
                if not all(re.match(r'^[-:]+$', c.strip()) for c in cells if c.strip()):
                    rows.append(cells)
            if rows:
                elements.append(('table', rows))
            continue

        # Blockquote
        if line.startswith('>'):
            ql = []
            while i < n and lines[i].startswith('>'):
                ql.append(lines[i].lstrip('>').strip()); i += 1
            elements.append(('blockquote', ql)); continue

        # Bullet list
        if re.match(r'^\s*[-*]\s', line):
            items = []
            while i < n and re.match(r'^\s*[-*]\s', lines[i]):
                items.append(re.sub(r'^\s*[-*]\s+', '', lines[i])); i += 1
            elements.append(('bullet', items)); continue

        # Numbered list
        if re.match(r'^\s*\d+\.\s', line):
            items = []
            while i < n and re.match(r'^\s*\d+\.\s', lines[i]):
                items.append(re.sub(r'^\s*\d+\.\s+', '', lines[i])); i += 1
            elements.append(('numbered', items)); continue

        # Regular paragraph
        elements.append(('paragraph', line)); i += 1

    return elements


# ── HWPX XML Builders ───────────────────────────────────────────────

def text_runs(text, default_cpr="0"):
    """Convert text with **bold** and `code` into <hp:run> elements."""
    runs = []
    parts = re.split(r'(\*\*[^*]+\*\*|`[^`]+`)', text)
    for part in parts:
        if not part: continue
        if part.startswith('**') and part.endswith('**'):
            runs.append(f'<hp:run charPrIDRef="10"><hp:t>{xml_escape(part[2:-2])}</hp:t></hp:run>')
        elif part.startswith('`') and part.endswith('`'):
            runs.append(f'<hp:run charPrIDRef="{default_cpr}"><hp:t>{xml_escape(part[1:-1])}</hp:t></hp:run>')
        else:
            runs.append(f'<hp:run charPrIDRef="{default_cpr}"><hp:t>{xml_escape(part)}</hp:t></hp:run>')
    return ''.join(runs) if runs else f'<hp:run charPrIDRef="{default_cpr}"><hp:t/></hp:run>'


def make_para(content_runs, para_pr="0", style_id="0"):
    pid = next_pid()
    return (
        f'<hp:p id="{pid}" paraPrIDRef="{para_pr}" styleIDRef="{style_id}" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'{content_runs}'
        f'</hp:p>'
    )


def empty_para():
    return make_para('<hp:run charPrIDRef="0"><hp:t/></hp:run>')


# ── Code Block Table Builder ───────────────────────────────────────

def make_code_table_xml(code_lines):
    """Build a 1x1 table containing code lines (9pt, 0pt after spacing)."""
    tbl_id = next_pid()
    z_order = next_z()
    total_height = max(len(code_lines), 1) * ROW_HEIGHT

    # Build paragraphs for each code line inside the cell
    cell_paras = []
    for line in code_lines:
        pid = next_pid()
        escaped = xml_escape(line) if line else ''
        cell_paras.append(
            f'<hp:p id="{pid}" paraPrIDRef="20" styleIDRef="0" '
            f'pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="11"><hp:t>{escaped}</hp:t></hp:run>'
            f'</hp:p>'
        )

    cell = (
        f'<hp:tc name="" header="0" hasMargin="1" '
        f'protect="0" editable="0" dirty="0" borderFillIDRef="3">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'{"".join(cell_paras)}'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="0" rowAddr="0"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{TABLE_WIDTH}" height="{total_height}"/>'
        f'<hp:cellMargin left="340" right="340" top="340" bottom="340"/>'
        f'</hp:tc>'
    )

    tbl = (
        f'<hp:tbl id="{tbl_id}" zOrder="{z_order}" numberingType="TABLE" '
        f'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
        f'dropcapstyle="None" pageBreak="CELL" repeatHeader="0" '
        f'rowCnt="1" colCnt="1" cellSpacing="0" '
        f'borderFillIDRef="3" noAdjust="0">'
        f'<hp:sz width="{TABLE_WIDTH}" widthRelTo="ABSOLUTE" '
        f'height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
        f'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
        f'horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" '
        f'vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="141" right="141" top="141" bottom="141"/>'
        f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
        f'<hp:tr>{cell}</hp:tr>'
        f'</hp:tbl>'
    )

    pid = next_pid()
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">{tbl}</hp:run>'
        f'</hp:p>'
    )


# ── Native HWPX Table Builder ───────────────────────────────────────

def calc_col_widths(rows, total_width=TABLE_WIDTH):
    """Calculate column widths proportional to max content length."""
    if not rows:
        return []
    col_count = max(len(r) for r in rows)
    max_lens = [0] * col_count
    for row in rows:
        for ci, cell in enumerate(row):
            if ci < col_count:
                max_lens[ci] = max(max_lens[ci], len(cell))
    max_lens = [max(l, 3) for l in max_lens]
    total_len = sum(max_lens)
    widths = [int(total_width * l / total_len) for l in max_lens]
    widths[-1] = total_width - sum(widths[:-1])
    return widths


def make_cell(text, col_addr, row_addr, width, is_header=False):
    """Build a <hp:tc> element."""
    bf_id = "4" if is_header else "3"
    cpr = "12" if is_header else "15"  # 10pt: header=Bold, data=Regular
    pid = next_pid()

    cell_runs = text_runs(text, cpr) if not is_header else (
        f'<hp:run charPrIDRef="12"><hp:t>{xml_escape(text)}</hp:t></hp:run>'
    )

    # Cell margin: 1.2mm = 340 HWPUNIT (1.2/25.4*7200)
    return (
        f'<hp:tc name="" header="{"1" if is_header else "0"}" hasMargin="1" '
        f'protect="0" editable="0" dirty="0" borderFillIDRef="{bf_id}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'{cell_runs}'
        f'</hp:p>'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="{col_addr}" rowAddr="{row_addr}"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{width}" height="{ROW_HEIGHT}"/>'
        f'<hp:cellMargin left="340" right="340" top="340" bottom="340"/>'
        f'</hp:tc>'
    )


def make_table_xml(rows, caption_text=""):
    """Build a complete <hp:p> containing a native HWPX table with caption."""
    if not rows:
        return ''

    col_count = max(len(r) for r in rows)
    row_count = len(rows)
    col_widths = calc_col_widths(rows)
    total_height = row_count * ROW_HEIGHT
    tbl_id = next_pid()
    z_order = next_z()

    tr_list = []
    for ri, row in enumerate(rows):
        is_header = (ri == 0)
        cells = []
        for ci in range(col_count):
            cell_text = row[ci] if ci < len(row) else ''
            cells.append(make_cell(cell_text, ci, ri, col_widths[ci], is_header))
        tr_list.append(f'<hp:tr>{"".join(cells)}</hp:tr>')

    # Build caption element
    cap_pid = next_pid()
    caption_xml = (
        f'<hp:caption side="BOTTOM" fullSz="0" width="{TABLE_WIDTH}" '
        f'gap="850" lastWidth="{TABLE_WIDTH}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{cap_pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="15"><hp:t>{xml_escape(caption_text)}</hp:t></hp:run>'
        f'</hp:p>'
        f'</hp:subList>'
        f'</hp:caption>'
    )

    tbl = (
        f'<hp:tbl id="{tbl_id}" zOrder="{z_order}" numberingType="TABLE" '
        f'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
        f'dropcapstyle="None" pageBreak="CELL" repeatHeader="1" '
        f'rowCnt="{row_count}" colCnt="{col_count}" cellSpacing="0" '
        f'borderFillIDRef="3" noAdjust="0">'
        f'<hp:sz width="{TABLE_WIDTH}" widthRelTo="ABSOLUTE" '
        f'height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
        f'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
        f'horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" '
        f'vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="141" right="141" top="141" bottom="141"/>'
        f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
        f'{caption_xml}'
        f'{"".join(tr_list)}'
        f'</hp:tbl>'
    )

    pid = next_pid()
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">{tbl}</hp:run>'
        f'</hp:p>'
    )


# ── Section Properties (from skeleton) ──────────────────────────────

SEC_PR = (
    '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" '
    'tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" '
    'outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" '
    'masterPageCnt="0">'
    '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
    '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
    '<hp:visibility hideFirstHeader="1" hideFirstFooter="1" '
    'hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" '
    'hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
    '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
    '<hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">'
    '<hp:margin header="4252" footer="4252" gutter="0" '
    'left="8504" right="8504" top="5668" bottom="4252"/>'
    '</hp:pagePr>'
    '<hp:footNotePr>'
    '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
    '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
    '<hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="EACH_COLUMN" beneathText="0"/>'
    '</hp:footNotePr>'
    '<hp:endNotePr>'
    '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
    '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
    '<hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>'
    '</hp:endNotePr>'
    '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '</hp:secPr>'
)


# ── Section XML Builder ─────────────────────────────────────────────

def build_section_xml(elements, title=""):
    paras = []

    # First paragraph: secPr + colPr in run 1, header/footer as ctrl in run 2
    # textWidth=42520 (page width 59528 - left 8504 - right 8504)
    # textHeight=4252 (header/footer margin from pagePr)

    # Header (머릿글): document title, 9pt gray italic, right-aligned
    header_ctrl = (
        f'<hp:ctrl>'
        f'<hp:header id="1" applyPageType="BOTH">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="42520" textHeight="4252" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{next_pid()}" paraPrIDRef="22" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="18"><hp:t>{xml_escape(title)}</hp:t></hp:run>'
        f'</hp:p>'
        f'</hp:subList>'
        f'</hp:header>'
        f'</hp:ctrl>'
    )

    # Footer (바닥글): page number centered, "- N -" format
    footer_ctrl = (
        f'<hp:ctrl>'
        f'<hp:footer id="2" applyPageType="BOTH">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="42520" textHeight="4252" hasTextRef="0" hasNumRef="0">'
        f'<hp:p id="{next_pid()}" paraPrIDRef="23" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="19"><hp:t>- </hp:t></hp:run>'
        f'<hp:run charPrIDRef="19">'
        f'<hp:ctrl><hp:autoNum num="1" numType="PAGE"/></hp:ctrl>'
        f'</hp:run>'
        f'<hp:run charPrIDRef="19"><hp:t> -</hp:t></hp:run>'
        f'</hp:p>'
        f'</hp:subList>'
        f'</hp:footer>'
        f'</hp:ctrl>'
    )

    pid = next_pid()
    paras.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">{SEC_PR}'
        f'<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" '
        f'colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
        f'</hp:run>'
        f'<hp:run charPrIDRef="0">{header_ctrl}{footer_ctrl}<hp:t/></hp:run>'
        f'</hp:p>'
    )

    # ── Cover page ─────────────────────────────────────────────────
    # 4 empty lines
    for _ in range(4):
        paras.append(empty_para())

    # Document title (H1 style: 24pt bold)
    pid = next_pid()
    paras.append(
        f'<hp:p id="{pid}" paraPrIDRef="12" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="7"><hp:t>{xml_escape(title)}</hp:t></hp:run>'
        f'</hp:p>'
    )

    # Author (12pt, 0pt paragraph spacing)
    pid = next_pid()
    paras.append(
        f'<hp:p id="{pid}" paraPrIDRef="20" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="16"><hp:t>'
        f'{xml_escape("문서 작성자: 주식회사 파워솔루션")}'
        f'</hp:t></hp:run></hp:p>'
    )

    # Date (12pt, 0pt paragraph spacing)
    now = datetime.now()
    date_str = f"{now.year}년 {now.month:02d}월 {now.day:02d}일 {now.hour:02d}시 {now.minute:02d}분"
    pid = next_pid()
    paras.append(
        f'<hp:p id="{pid}" paraPrIDRef="20" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="16"><hp:t>'
        f'{xml_escape("문서 작성일: " + date_str)}'
        f'</hp:t></hp:run></hp:p>'
    )

    # Merge consecutive bullet groups separated by empty elements
    merged = []
    for elem in elements:
        if (elem[0] == 'bullet' and len(merged) >= 2
                and merged[-1][0] == 'empty' and merged[-2][0] == 'bullet'):
            merged.pop()  # remove empty between bullets
            merged[-1][1].extend(elem[1])  # merge items into previous bullet
        else:
            merged.append(elem)
    elements = merged

    tbl_count = 0
    last_heading = ""
    h1_skipped = False

    for etype, data in elements:

        if etype == 'heading':
            level, title = data
            # Skip the first H1 (already on cover page)
            if level == 1 and not h1_skipped:
                h1_skipped = True
                last_heading = title
                continue
            cpr = {1: "7", 2: "8", 3: "9", 4: "14"}.get(level, "10")
            pb = "1" if level == 2 else "0"  # H2 always starts new page
            last_heading = title
            pid = next_pid()
            paras.append(
                f'<hp:p id="{pid}" paraPrIDRef="12" styleIDRef="0" '
                f'pageBreak="{pb}" columnBreak="0" merged="0">'
                f'{text_runs(title, cpr)}'
                f'</hp:p>'
            )

        elif etype == 'paragraph':
            paras.append(make_para(text_runs(data, "0")))

        elif etype == 'empty':
            paras.append(empty_para())

        elif etype == 'blockquote':
            for line in data:
                if line:
                    # Source lines (출처:) use 8pt charPr 17
                    cpr = "17" if line.lstrip().startswith("출처") else "13"
                    paras.append(make_para(text_runs(line, cpr), para_pr="1"))

        elif etype == 'bullet':
            for item in data:
                prefix = '<hp:run charPrIDRef="0"><hp:t>\u2022 </hp:t></hp:run>'
                paras.append(make_para(prefix + text_runs(item, "0"), para_pr="21"))

        elif etype == 'numbered':
            for idx, item in enumerate(data, 1):
                prefix = f'<hp:run charPrIDRef="0"><hp:t>{idx}. </hp:t></hp:run>'
                paras.append(make_para(prefix + text_runs(item, "0"), para_pr="1"))

        elif etype == 'code':
            paras.append(make_code_table_xml(data))

        elif etype == 'table':
            # Remove empty paragraph immediately before table
            if paras and '<hp:t/>' in paras[-1] and 'hp:tbl' not in paras[-1]:
                paras.pop()
            tbl_count += 1
            caption = f'[\ud45c {tbl_count}] {last_heading}'
            tbl_xml = make_table_xml(data, caption_text=caption)
            paras.append(tbl_xml)

        elif etype == 'hr':
            pass  # Horizontal rules are ignored

    print(f"  Generated {tbl_count} native tables")

    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        f'<hs:sec {NS_DECL}>'
        f'{"".join(paras)}'
        f'</hs:sec>'
    )


# ── Content HPF (metadata) ──────────────────────────────────────────

def build_content_hpf(title="Untitled", author="Claude Code", subject="", description="",
                      language="ko", keywords=""):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        f'<opf:package {NS_DECL} version="" unique-identifier="" id="">'
        f'<opf:metadata>'
        f'<opf:title>{xml_escape(title)}</opf:title>'
        f'<opf:language>{language}</opf:language>'
        f'<opf:meta name="creator" content="text">{xml_escape(author)}</opf:meta>'
        f'<opf:meta name="subject" content="text">{xml_escape(subject)}</opf:meta>'
        f'<opf:meta name="description" content="text">{xml_escape(description)}</opf:meta>'
        f'<opf:meta name="lastsaveby" content="text">{xml_escape(author)}</opf:meta>'
        f'<opf:meta name="keyword" content="text">{xml_escape(keywords)}</opf:meta>'
        f'</opf:metadata>'
        f'<opf:manifest>'
        f'<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
        f'<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
        f'<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
        f'</opf:manifest>'
        f'<opf:spine>'
        f'<opf:itemref idref="header" linear="yes"/>'
        f'<opf:itemref idref="section0" linear="yes"/>'
        f'</opf:spine>'
        f'</opf:package>'
    )


# ── Main ─────────────────────────────────────────────────────────────

def convert(input_md, output_hwpx, skeleton_path=None, title=None, author="Claude Code",
            subject="", description="", language="ko", keywords=""):
    """Convert a markdown file to HWPX format.

    Args:
        input_md: Path to input markdown file
        output_hwpx: Path to output HWPX file
        skeleton_path: Path to Skeleton.hwpx template (auto-detected if None)
        title: Document title (defaults to filename)
        author: Document author
        subject: Document subject
        description: Document description
        language: Document language code
        keywords: Document keywords
    """
    # Resolve skeleton path
    if skeleton_path is None:
        # Try common locations
        candidates = [
            os.path.join(os.path.dirname(__file__), '..', 'assets', 'Skeleton.hwpx'),
            os.path.expanduser('~/Library/Python/3.12/lib/python/site-packages/hwpx/data/Skeleton.hwpx'),
            os.path.expanduser('~/Library/Python/3.13/lib/python/site-packages/hwpx/data/Skeleton.hwpx'),
        ]
        for c in candidates:
            if os.path.exists(c):
                skeleton_path = c
                break
        if skeleton_path is None:
            # Try pip show
            try:
                import hwpx
                skeleton_path = os.path.join(os.path.dirname(hwpx.__file__), 'data', 'Skeleton.hwpx')
            except ImportError:
                pass
        if skeleton_path is None or not os.path.exists(skeleton_path):
            print("ERROR: Skeleton.hwpx not found. Install hwpx package or specify --skeleton path.")
            sys.exit(1)

    if title is None:
        title = os.path.splitext(os.path.basename(input_md))[0]

    # Extract skeleton to temp directory
    skeleton_dir = tempfile.mkdtemp(prefix='hwpx_skeleton_')
    try:
        with zipfile.ZipFile(skeleton_path, 'r') as zf:
            zf.extractall(skeleton_dir)

        print(f"Reading {input_md}...")
        report = read_file(input_md)

        print("Preprocessing markdown (numbering headings, removing excess blanks)...")
        report = preprocess_markdown(report)
        with open(input_md, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"  Saved preprocessed markdown to {input_md}")

        print("Parsing markdown...")
        elements = parse_markdown(report)
        print(f"  Parsed {len(elements)} elements")

        print("Building header.xml...")
        header_xml = build_header_xml(skeleton_dir)

        print("Building section0.xml...")
        section_xml = build_section_xml(elements, title=title)

        print("Building content.hpf...")
        content_hpf = build_content_hpf(title, author, subject, description, language, keywords)

        # Read unchanged files from skeleton
        mimetype = read_file(f'{skeleton_dir}/mimetype').strip()
        version_xml = read_file(f'{skeleton_dir}/version.xml')
        settings_xml = read_file(f'{skeleton_dir}/settings.xml')
        container_xml = read_file(f'{skeleton_dir}/META-INF/container.xml')
        manifest_xml = read_file(f'{skeleton_dir}/META-INF/manifest.xml')
        container_rdf = read_file(f'{skeleton_dir}/META-INF/container.rdf')

        preview = f"{title}\r\n{subject}\r\n{description}"

        print(f"Writing HWPX to {output_hwpx}...")
        os.makedirs(os.path.dirname(os.path.abspath(output_hwpx)), exist_ok=True)

        with zipfile.ZipFile(output_hwpx, 'w', zipfile.ZIP_DEFLATED) as zf:
            # mimetype MUST be first entry, stored uncompressed
            zf.writestr('mimetype', mimetype, compress_type=zipfile.ZIP_STORED)
            zf.writestr('version.xml', version_xml)
            zf.writestr('settings.xml', settings_xml)
            zf.writestr('Contents/content.hpf', content_hpf)
            zf.writestr('Contents/header.xml', header_xml)
            zf.writestr('Contents/section0.xml', section_xml)
            zf.writestr('META-INF/container.xml', container_xml)
            zf.writestr('META-INF/manifest.xml', manifest_xml)
            zf.writestr('META-INF/container.rdf', container_rdf)
            zf.writestr('Preview/PrvText.txt', preview)

        size = os.path.getsize(output_hwpx)
        print(f"\nDone! HWPX saved: {output_hwpx}")
        print(f"File size: {size:,} bytes ({size/1024:.1f} KB)")

    finally:
        shutil.rmtree(skeleton_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to HWPX format')
    parser.add_argument('input', help='Input markdown file')
    parser.add_argument('output', help='Output HWPX file')
    parser.add_argument('--skeleton', help='Path to Skeleton.hwpx template')
    parser.add_argument('--title', help='Document title')
    parser.add_argument('--author', default='Claude Code', help='Document author')
    parser.add_argument('--subject', default='', help='Document subject')
    parser.add_argument('--description', default='', help='Document description')
    parser.add_argument('--language', default='ko', help='Document language (default: ko)')
    parser.add_argument('--keywords', default='', help='Document keywords')
    args = parser.parse_args()

    convert(
        args.input, args.output,
        skeleton_path=args.skeleton,
        title=args.title, author=args.author,
        subject=args.subject, description=args.description,
        language=args.language, keywords=args.keywords,
    )


if __name__ == '__main__':
    main()
