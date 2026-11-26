#!/usr/bin/env python3
"""Build an editable deck from TEMPLATE + plan.json with REAL master/layout/placeholder
inheritance, a documented font ladder, two-column & side-by-side figure layouts, native
theme-styled tables, inline **bold** lead-ins, aspect-safe figure fitting, and structured
evidence in speaker notes + build_report.json.

Usage:
  python3 build_deck.py PLAN.json --out report.json

No-script fallback: build manually in PowerPoint from the template's own layouts; keep
placeholders (do not delete used placeholders) so inheritance stays verifiable.
"""
import argparse, json, math, os, re, sys, tempfile, zipfile
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu
from PIL import Image, ImageChops

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

FOOTERISH = ("DATE", "FOOTER", "SLIDE_NUMBER", "HEADER", "BILLBOARD")
TITLEISH = ("CENTER_TITLE", "TITLE")
SUBISH = ("SUBTITLE",)
TEXT_BODYISH = ("BODY", "OBJECT", "TEXT", "VERTICAL_BODY", "CONTENT")
PIC_PH = ("PICTURE", "CLIP_ART")
ALL_CONTENT_PH = TEXT_BODYISH + PIC_PH

KEYWORDS = {
    "title": ["title slide", "标题幻灯片", "标题和副标题", "title and subtitle", "presentation title"],
    "section": ["section header", "section title", "节标题", "小节", "章节"],
    "content": ["title and content", "标题和内容", "title and text", "content with caption",
                "bullet", "标题和文本", "one content", "title content"],
    "two_col": ["two content", "comparison", "两栏内容", "比较", "two column", "2 content",
                "title and content", "标题和内容"],
    "figure": ["picture with caption", "title only", "图片与标题", "picture", "two content",
               "title and content", "标题和内容", "content with caption", "仅标题"],
}

DEFAULT_LADDER = {
    "title": 36, "subtitle": 20, "section": 30, "slide_title": 28,
    "body": [18, 16, 14, 13], "caption": 11.5, "table": [13.5, 12.5, 11.5, 10.5],
}
MIN_BODY_PT = 13.0
BOTTOM_MARGIN_IN = 0.45
GUTTER_IN = 0.30


def as_in(v, default, positive=False):
    """EMU Length -> inches, guarding inherited-None (and degenerate <=0.05in sizes when positive=True)."""
    try:
        val = v / 914400
        if positive and val <= 0.05:
            return default
        return val
    except TypeError:
        return default


def hex_to_rgb(h, fallback=(31, 78, 121)):
    if not h or not isinstance(h, str):
        return RGBColor(*fallback)
    s = h.lstrip("#")
    if len(s) != 6:
        return RGBColor(*fallback)
    try:
        return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        return RGBColor(*fallback)


def tint_rgb(rgb: RGBColor, factor=0.88):
    """Blend RGB toward white by factor (0=original, 1=white)."""
    r = int(rgb[0] + (255 - rgb[0]) * factor)
    g = int(rgb[1] + (255 - rgb[1]) * factor)
    b = int(rgb[2] + (255 - rgb[2]) * factor)
    return RGBColor(min(255, r), min(255, g), min(255, b))


def read_theme_colors(pptx_path: Path):
    colors = {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml", n))
            if not names:
                return colors
            xml = z.read(names[0]).decode("utf-8", "ignore")
        for cname in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            m = re.search(rf"<a:{cname}>(.*?)</a:{cname}>", xml, re.S)
            if m:
                rgb = re.search(r'(?:lastClr|val)="([0-9A-Fa-f]{6})"', m.group(1))
                if rgb:
                    colors[cname] = "#" + rgb.group(1)
    except Exception:
        pass
    return colors


def _slide_clr_map(slide):
    """Resolve effective OOXML color map (bg1/bg2/tx1/tx2 -> dk1/lt1/dk2/lt2) from master & layout override."""
    cmap = {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}
    try:
        master = slide.slide_layout.slide_master
        cm = master._element.find(qn("p:clrMap"))
        if cm is not None:
            for k in ("bg1", "bg2", "tx1", "tx2"):
                if cm.get(k):
                    cmap[k] = cm.get(k)
        ovr = slide.slide_layout._element.find(qn("p:clrMapOvr"))
        if ovr is not None:
            ocm = ovr.find(qn("a:overrideClrMapping"))
            if ocm is not None:
                for k in ("bg1", "bg2", "tx1", "tx2"):
                    if ocm.get(k):
                        cmap[k] = ocm.get(k)
    except Exception:
        pass
    return cmap


def _extract_fill_hex(container, theme_colors=None, alias_map=None):
    """Extract 6-hex RGB from an XML container's solidFill, gradFill (mean stop), or direct color node."""
    if container is None:
        return None
    amap = alias_map or {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}

    def _node_hex(node):
        if node is None:
            return None
        srgb = node.find(qn("a:srgbClr"))
        if srgb is not None and srgb.get("val"):
            return srgb.get("val").lstrip("#")
        sch = node.find(qn("a:schemeClr"))
        if sch is not None and sch.get("val") and theme_colors:
            k = sch.get("val")
            hx = theme_colors.get(amap.get(k, k)) or theme_colors.get(k)
            if hx:
                return hx.lstrip("#")
        return None

    sf = container.find(qn("a:solidFill"))
    hx = _node_hex(sf)
    if hx and len(hx) == 6:
        return hx
    gf = container.find(qn("a:gradFill"))
    if gf is not None:
        gs_lst = gf.find(qn("a:gsLst"))
        if gs_lst is not None:
            stops = []
            for gs in gs_lst.findall(qn("a:gs")):
                shx = _node_hex(gs)
                if shx and len(shx) == 6:
                    stops.append((int(shx[0:2], 16), int(shx[2:4], 16), int(shx[4:6], 16)))
            if stops:
                r = sum(c[0] for c in stops) // len(stops)
                g = sum(c[1] for c in stops) // len(stops)
                b = sum(c[2] for c in stops) // len(stops)
                return f"{r:02X}{g:02X}{b:02X}"
    return _node_hex(container)


def _bg_hex_readonly(obj, theme_colors=None, clr_map=None, slide_w=None, slide_h=None):
    """Read solid/gradient background or full-slide backdrop shape hex from XML WITHOUT mutating XML."""
    try:
        csld = obj._element.find(qn("p:cSld"))
        if csld is None:
            return None
        alias_map = clr_map or {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}
        bg = csld.find(qn("p:bg"))
        if bg is not None:
            for container_tag in ("p:bgPr", "p:bgRef"):
                cont = bg.find(qn(container_tag))
                hx = _extract_fill_hex(cont, theme_colors, alias_map)
                if hx and len(hx) == 6:
                    return hx
        # Also detect full-slide decorative backdrop rectangle (<p:sp> covering >=75% slide area)
        if slide_w and slide_h:
            slide_area = max(slide_w * slide_h, 1.0)
            for sh in getattr(obj, "shapes", ()):
                if getattr(sh, "is_placeholder", False) or getattr(sh, "has_table", False):
                    continue
                if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                    continue
                bw, bh = as_in(sh.width, 0.0), as_in(sh.height, 0.0)
                if bw * bh >= 0.75 * slide_area:
                    spPr = sh._element.find(qn("p:spPr"))
                    hx = _extract_fill_hex(spPr, theme_colors, alias_map)
                    if hx and len(hx) == 6:
                        return hx
    except Exception:
        pass
    return None


def is_dark_slide(slide, theme_colors=None):
    """Detect whether a slide, its layout, or its master has a dark solid/gradient/backdrop background (read-only)."""
    cmap = _slide_clr_map(slide)
    sw = sh = None
    try:
        prs_el = slide.part.package.presentation_part.presentation
        sw, sh = Emu(prs_el.slide_width).inches, Emu(prs_el.slide_height).inches
    except Exception:
        pass
    for obj in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        hx = _bg_hex_readonly(obj, theme_colors, clr_map=cmap, slide_w=sw, slide_h=sh)
        if hx and len(hx) == 6:
            r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
            return (0.299 * r + 0.587 * g + 0.114 * b) < 128
    # Fallback: if master/layout clrMap explicitly maps bg1 to a dark scheme color
    bg1_slot = cmap.get("bg1", "lt1")
    hx = (theme_colors or {}).get(bg1_slot)
    if hx and len(hx) == 7:
        r, g, b = int(hx[1:3], 16), int(hx[3:5], 16), int(hx[5:7], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) < 128
    return False


def pin_ph_geometry(ph, content_box):
    """Materialize full explicit geometry on a cloned placeholder BEFORE resizing it,
    healing degenerate 0-height/0-width partial-<a:xfrm> template placeholders."""
    raw_h = as_in(ph.height, 0.0)
    raw_w = as_in(ph.width, 0.0)
    l = as_in(ph.left, content_box[0])
    t = as_in(ph.top, content_box[1])
    if raw_h <= 0.05 and t <= 0.05:
        t = content_box[1]
    if raw_w <= 0.05 and l <= 0.05:
        l = content_box[0]
    w = raw_w if raw_w > 0.05 else content_box[2]
    h = raw_h if raw_h > 0.05 else content_box[3]
    ph.left, ph.top, ph.width, ph.height = Inches(l), Inches(t), Inches(w), Inches(h)
    return l, t, w, h


def ph_type(p):
    return str(p.placeholder_format.type).split(" ")[0]


def is_footerish(p):
    return any(t in ph_type(p) for t in FOOTERISH)


def find_phs(shapes, kinds, exclude_footer=True):
    out = []
    for p in shapes:
        if not p.is_placeholder:
            continue
        if exclude_footer and is_footerish(p):
            continue
        t = ph_type(p)
        if any(t == k or t.startswith(k) for k in kinds):
            out.append(p)
    out.sort(key=lambda s: (round(as_in(s.left, 0), 2), round(as_in(s.top, 0), 2)))
    return out


def find_ph(shapes, kinds, exclude_footer=True):
    found = find_phs(shapes, kinds, exclude_footer=exclude_footer)
    return found[0] if found else None


def remove_shape(sh):
    try:
        el = sh._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    except Exception:
        pass


def re_split(s):
    return [t.strip().lower() for t in s.replace("，", ",").replace("、", ",").replace("|", ",").split(",") if t.strip()]


def select_layout(prs, kind, hint, exact, n_images=0, has_bullets=False, has_table=False):
    layouts = [lo for m in prs.slide_masters for lo in m.slide_layouts]
    if exact:
        for lo in layouts:
            if lo.name.strip().lower() == exact.strip().lower():
                return lo, "exact-name"
    kws = KEYWORDS.get(kind, KEYWORDS["content"])
    hint_toks = [t for t in re_split(hint)] if hint else []
    want_vertical = any(v in " ".join(hint_toks) for v in ("vertical", "竖排"))

    def struct_info(lo):
        phs = [p for p in lo.placeholders if not is_footerish(p)]
        has_title = any(ph_type(p) in TITLEISH for p in phs)
        t_phs = [p for p in phs if ph_type(p) in TEXT_BODYISH]
        n_text = len(t_phs)
        max_text_h = max((as_in(p.height, 0) for p in t_phs), default=0.0)
        n_pic = sum(1 for p in phs if ph_type(p) in PIC_PH)
        has_sub = any(ph_type(p) in SUBISH for p in phs)
        return has_title, n_text, max_text_h, n_pic, has_sub, len(phs)

    best, best_score, how = None, -1e9, ""
    for lo in layouts:
        has_title, n_text, max_text_h, n_pic, has_sub, n_phs = struct_info(lo)
        name = (lo.name or "").lower()
        score = 0.0
        if not want_vertical and ("vertical" in name or "竖排" in name):
            score -= 35
        hits = [kw for kw in kws if kw in name]
        if hits:
            score += 20 + 5 * (len(hits) - 1)
        for tok in hint_toks:
            if tok and tok in name:
                score += 8
        tph = next((p for p in lo.placeholders if ph_type(p) in TITLEISH), None)
        title_top = as_in(tph.top, 99) if tph is not None else 99
        if has_title:
            score += 5
            if kind in ("content", "figure", "two_col"):
                if title_top < 2.5:
                    score += 8
                elif title_top > 3.5:
                    score -= 14

        if kind == "title":
            score += 8 if has_sub else 0
            score -= 4 if n_text >= 1 else 0
        elif kind == "section":
            score -= 4 if n_text >= 2 else 0
        elif kind == "content":
            score += 10 if n_text >= 1 else -4
            if n_text == 1:
                score += 3
        elif kind == "two_col":
            if n_text >= 2 or (n_text >= 1 and n_pic >= 1):
                score += 18
            elif n_text == 1:
                score += 6
        elif kind == "figure":
            if n_images == 1 and has_bullets:
                if n_pic >= 1 and n_text >= 1 and max_text_h >= 1.8 and title_top < 2.5:
                    score += 14
                elif n_pic >= 1 and n_text >= 1 and (max_text_h < 1.8 or title_top >= 3.0):
                    # Avoid PowerPoint's default Picture with Caption (tiny 0.88in body + bottom title)
                    score -= 14
                elif n_text >= 2:
                    score += 12
                elif n_text >= 1:
                    score += 9
            elif n_images >= 2 or not has_bullets:
                if n_text == 0 and n_pic == 0 and has_title and title_top < 2.5:
                    score += 12   # clean Title Only canvas
                elif n_pic >= 1 and title_top < 2.5:
                    score += 6

        score -= 0.15 * n_phs
        if score > best_score:
            best, best_score, how = lo, score, "keyword-score"
    if best is None and layouts:
        best, how = layouts[0], "first-layout-fallback"
    return best, how


def norm_bullets(raw):
    out = []
    for b in (raw or []):
        if isinstance(b, str):
            out.append({"text": b, "level": 0})
        elif isinstance(b, dict) and b.get("text"):
            out.append({"text": str(b["text"]), "level": int(b.get("level", 0) or 0)})
    return out


def strip_md_bold(text):
    return text.replace("**", "")


def est_text_w_pt(text, size):
    clean = strip_md_bold(text)
    return sum(size * (1.0 if ord(c) > 0x2E7F else 0.52) for c in clean)


def est_body_height_in(bullets, width_in, size, line_spacing=1.18, para_gap=6):
    lines = 0
    for b in bullets:
        lvl = b.get("level", 0)
        w_pt = max((width_in - 0.4 * lvl) * 72 - 24, 40)
        lines += max(1, math.ceil(est_text_w_pt(b["text"], size) / w_pt))
    return (lines * size * line_spacing + max(0, len(bullets) - 1) * para_gap) / 72


def write_rich_runs(para, text, size, base_bold=False, italic=False, color=None):
    """Parse **bold** spans cleanly without setting font.name (preserves theme inheritance)."""
    tokens = re.split(r"(\*\*[^*]+\*\*)", text)
    wrote = False
    for tok in tokens:
        if not tok:
            continue
        is_b = base_bold
        chunk = tok
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            is_b = True
            chunk = tok[2:-2]
        r = para.add_run()
        r.text = chunk
        r.font.size = Pt(size)
        if is_b:
            r.font.bold = True
        if italic:
            r.font.italic = True
        if color is not None:
            r.font.color.rgb = color
        wrote = True
    if not wrote:
        r = para.add_run()
        r.text = text
        r.font.size = Pt(size)


def set_tf(tf, size, bold=False, italic=False, color=None, align=None):
    tf.word_wrap = True
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size)
            if bold:
                r.font.bold = True
            if italic:
                r.font.italic = True
            if color is not None:
                r.font.color.rgb = color
        if align is not None:
            p.alignment = align


def _dark_slide_text_color(slide, dark_bg, theme_colors):
    """On dark slides whose effective tx1/txStyles still resolve to dark ink, return a crisp light
    foreground RGBColor so text never becomes invisible black-on-navy. Returns None when template
    already provides light text or on light slides (preserving pure theme inheritance)."""
    if not dark_bg:
        return None
    cmap = _slide_clr_map(slide)
    tx1_slot = cmap.get("tx1", "dk1")
    tx1_hex = (theme_colors or {}).get(tx1_slot) or (theme_colors or {}).get("dk1")
    tx1_dark = True
    if tx1_hex and len(tx1_hex) == 7:
        r, g, b = int(tx1_hex[1:3], 16), int(tx1_hex[3:5], 16), int(tx1_hex[5:7], 16)
        tx1_dark = (0.299 * r + 0.587 * g + 0.114 * b) < 140
    if not tx1_dark:
        return None
    # Check if slide master txStyles explicitly maps title/body text to lt1/bg1/light srgb
    try:
        master = slide.slide_layout.slide_master
        txs = master._element.find(qn("p:txStyles"))
        if txs is not None:
            xml_str = txs.xml
            if 'val="lt1"' in xml_str or ('val="bg1"' in xml_str and cmap.get("bg1") == "lt1"):
                return None
            for m in re.finditer(r'srgbClr\s+val="([0-9A-Fa-f]{6})"', xml_str):
                hx = m.group(1)
                if (0.299 * int(hx[0:2], 16) + 0.587 * int(hx[2:4], 16) + 0.114 * int(hx[4:6], 16)) > 170:
                    return None
    except Exception:
        pass
    return RGBColor(242, 246, 252)


def fill_title(slide, text, size, color=None, slide_w=13.33):
    ph = find_ph(slide.shapes, TITLEISH)
    if ph is None:
        return None, None
    raw_w = as_in(ph.width, 0.0)
    raw_h = as_in(ph.height, 0.0)
    if raw_w <= 0.1 or raw_h <= 0.1:
        tl = as_in(ph.left, 0.55)
        tt = as_in(ph.top, 0.35)
        if raw_h <= 0.1 and tt <= 0.05:
            tt = 0.35
        tw = raw_w if raw_w > 0.1 else max(slide_w - 1.1, 4.0)
        th = raw_h if raw_h > 0.1 else 1.05
        ph.left, ph.top, ph.width, ph.height = Inches(tl), Inches(tt), Inches(tw), Inches(th)
    ph_w = max(as_in(ph.width, 8.5), 2.0)
    ph_h = max(as_in(ph.height, 1.05), 0.55)
    sz = float(size)
    n_eff = len(text) + sum(0.6 for c in text if ord(c) > 0x2E7F)
    if n_eff > 46:
        sz = max(sz - 4, 20.0)
    if n_eff > 70:
        sz = max(sz - 6, 18.0)
    # Geometry-aware shrink so wrapped titles never overflow narrow/short title placeholders
    while sz > 18.0:
        w_pt = max(ph_w * 72 - 16, 80)
        lines = max(1, math.ceil(est_text_w_pt(text, sz) / w_pt))
        need_h = lines * sz * 1.18 / 72
        if need_h <= ph_h * 1.06:
            break
        sz = max(18.0, sz - 2.0)
    tf = ph.text_frame
    tf.word_wrap = True
    tf.clear()
    para = tf.paragraphs[0]
    write_rich_runs(para, text, sz, base_bold=True, color=color)
    return ph, dict(idx=ph.placeholder_format.idx, type=ph_type(ph), pt=sz)


def _strip_bullets(para):
    pPr = para._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
        old = pPr.find(qn(tag))
        if old is not None:
            pPr.remove(old)
    pPr.append(parse_xml(f'<a:buNone xmlns:a="{A_NS}"/>'))


def fill_bullets(target_shape, bullets, ladder, width_in, height_in, report_slide,
                 heading=None, accent_rgb=None, text_rgb=None):
    sizes = sorted({float(x) for x in ladder["body"]}, reverse=True)
    extra_h = 0.34 if heading else 0.0
    avail_h = max(height_in - extra_h, 0.5)
    size = sizes[0]
    for cand in sizes[1:]:
        if est_body_height_in(bullets, width_in, size) <= avail_h:
            break
        size = max(cand, MIN_BODY_PT)
    line_sp = 1.15
    para_gap_pt = 5 if size <= 15 else 6
    # Graceful tight-fit pass when strict_slide_count prevents splitting (keeps >=12.5pt floor)
    if est_body_height_in(bullets, width_in, size, line_sp, para_gap_pt) > avail_h * 1.02:
        line_sp = 1.06
        para_gap_pt = 2.5
        if est_body_height_in(bullets, width_in, size, line_sp, para_gap_pt) > avail_h * 1.05:
            size = max(12.5, size - 0.5)
    tf = target_shape.text_frame
    tf.word_wrap = True
    tf.clear()
    first = True
    if heading:
        para = tf.paragraphs[0]
        first = False
        para.level = 0
        para.space_after = Pt(5 if line_sp < 1.1 else 6)
        para.line_spacing = line_sp
        _strip_bullets(para)
        write_rich_runs(para, heading, min(size + 2, 20), base_bold=True, color=accent_rgb)
    for b in bullets:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.level = min(b.get("level", 0), 4)
        para.space_after = Pt(para_gap_pt)
        para.line_spacing = line_sp
        write_rich_runs(para, b["text"], size, base_bold=False, color=text_rgb)
    prev = report_slide.get("body_pt")
    report_slide["body_pt"] = min(prev, size) if prev else size
    return size


def _blend_rgb(c1: RGBColor, c2: RGBColor, t: float):
    """Blend c1 -> c2 by factor t in [0, 1]."""
    return RGBColor(
        int(c1[0] + (c2[0] - c1[0]) * t),
        int(c1[1] + (c2[1] - c1[1]) * t),
        int(c1[2] + (c2[2] - c1[2]) * t),
    )


def place_table(slide, tbl_spec, box, ladder, accent_rgb, dark_bg, report_slide):
    """Render a native editable PPTX table styled with the template's accent1 color,
    computing real wrapped row heights so rows never collide with caption or slide bottom."""
    headers = tbl_spec.get("headers") or tbl_spec.get("columns") or []
    rows = tbl_spec.get("rows") or []
    if not headers and rows:
        headers, rows = rows[0], rows[1:]
    ncols = max([len(headers)] + [len(r) for r in rows] + [0])
    if ncols == 0:
        return
    all_rows = ([headers] if headers else []) + rows
    nrows = len(all_rows)
    left, top, w, h = box

    cap_txt = tbl_spec.get("caption", "")
    src_txt = tbl_spec.get("source", "")
    full_cap = (f"{cap_txt}" + (f"  [{src_txt}]" if src_txt else "")).strip()
    cap_h = 0.0
    if full_cap:
        cap_lines = max(1, math.ceil(est_text_w_pt(full_cap, ladder["caption"]) / max(w * 72 - 12, 40)))
        cap_h = cap_lines * ladder["caption"] * 1.25 / 72 + 0.14
    tbl_h = max(h - cap_h - 0.08, 0.8)

    # 1) Proportional column widths computed FIRST so we can estimate true wrapped row heights
    col_lens = []
    for c in range(ncols):
        mx = 4.0
        for r in all_rows:
            val = strip_md_bold(str(r[c])) if c < len(r) else ""
            mx = max(mx, min(len(val) + sum(0.65 for ch in val if ord(ch) > 0x2E7F), 32.0))
        col_lens.append(mx)
    tot = sum(col_lens) or 1.0
    col_widths_in = [max(w * cl / tot, 0.55) for cl in col_lens]
    scale_w = w / sum(col_widths_in)
    col_widths_in = [cw * scale_w for cw in col_widths_in]

    def est_row_heights(pt_val):
        rhs = []
        for r_idx, row_vals in enumerate(all_rows):
            max_lines = 1
            for c_idx in range(ncols):
                val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
                avail_pt = max(col_widths_in[c_idx] * 72 - 14, 24)
                # header & bold cells are slightly wider
                eff_pt = pt_val * (1.06 if (r_idx == 0 and headers) else 1.02)
                lines = max(1, math.ceil(est_text_w_pt(val, eff_pt) / avail_pt)) if val else 1
                max_lines = max(max_lines, lines)
            rhs.append(max_lines * pt_val * 1.28 / 72 + 0.13)
        return rhs

    tsizes = sorted({float(x) for x in ladder.get("table", [13.5, 12.5, 11.5, 10.5])}, reverse=True)
    for extra_pt in (9.5, 9.0):
        if extra_pt not in tsizes:
            tsizes.append(extra_pt)
    pt = tsizes[-1]
    row_hs = est_row_heights(pt)
    for cand in tsizes:
        rhs = est_row_heights(cand)
        if sum(rhs) <= tbl_h * 1.02:
            pt, row_hs = cand, rhs
            break
    # Stretch gently if table is compact (up to 1.30x) without exceeding tbl_h
    sum_rh = sum(row_hs)
    if sum_rh < min(tbl_h, 2.2) and sum_rh > 0:
        grow = min(1.30, min(tbl_h, 2.2) / sum_rh)
        row_hs = [rh * grow for rh in row_hs]
        sum_rh = sum(row_hs)
    actual_h = min(sum_rh, tbl_h)
    dense_tbl = sum_rh > tbl_h
    if dense_tbl and sum_rh > 0:
        shrink = tbl_h / sum_rh
        row_hs = [rh * shrink for rh in row_hs]

    shape = slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top), Inches(w), Inches(actual_h))
    table = shape.table

    for c in range(ncols):
        table.columns[c].width = Inches(col_widths_in[c])
    for r_idx, rh in enumerate(row_hs):
        table.rows[r_idx].height = Inches(rh)

    # Highlight rows (0-indexed among data rows; negative index supported)
    hl_set = set()
    for hr in (tbl_spec.get("highlight_rows") or []):
        try:
            idx = int(hr)
            if idx < 0:
                idx = len(rows) + idx
            if 0 <= idx < len(rows):
                hl_set.add(idx + (1 if headers else 0))
        except (TypeError, ValueError):
            pass

    hdr_luma = 0.299 * accent_rgb[0] + 0.587 * accent_rgb[1] + 0.114 * accent_rgb[2]
    hdr_txt_rgb = RGBColor(24, 30, 40) if hdr_luma > 168 else RGBColor(255, 255, 255)
    if dark_bg:
        even_rgb = RGBColor(28, 42, 64)
        odd_rgb = RGBColor(36, 53, 80)
        hl_rgb = _blend_rgb(even_rgb, accent_rgb, 0.35)
        body_txt_rgb = RGBColor(238, 243, 250)
    else:
        even_rgb = tint_rgb(accent_rgb, 0.92)
        odd_rgb = RGBColor(255, 255, 255)
        hl_rgb = tint_rgb(accent_rgb, 0.82)
        body_txt_rgb = RGBColor(28, 34, 43)

    num_re = re.compile(r"^[\s+\-−~<>≤≥]*\d[\d.,%\s×x±()\-−+]*(?:[eE][+\-−]?\d+)?[a-zA-Z%]*$")
    for r_idx, row_vals in enumerate(all_rows):
        is_hdr = bool(headers) and r_idx == 0
        is_hl = r_idx in hl_set
        for c_idx in range(ncols):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.06 if dense_tbl else 0.07)
            cell.margin_right = Inches(0.06 if dense_tbl else 0.07)
            cell.margin_top = Inches(0.02 if dense_tbl else 0.03)
            cell.margin_bottom = Inches(0.02 if dense_tbl else 0.03)
            cell.fill.solid()
            if is_hdr:
                cell.fill.fore_color.rgb = accent_rgb
            elif is_hl:
                cell.fill.fore_color.rgb = hl_rgb
            elif r_idx % 2 == 0:
                cell.fill.fore_color.rgb = even_rgb
            else:
                cell.fill.fore_color.rgb = odd_rgb

            val = str(row_vals[c_idx]) if c_idx < len(row_vals) else ""
            tf = cell.text_frame
            tf.word_wrap = True
            tf.clear()
            p = tf.paragraphs[0]
            if is_hdr:
                p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            elif c_idx > 0 and num_re.match(strip_md_bold(val).strip()):
                p.alignment = PP_ALIGN.RIGHT
            else:
                p.alignment = PP_ALIGN.LEFT
            write_rich_runs(
                p, val, pt,
                base_bold=(is_hdr or is_hl),
                color=(hdr_txt_rgb if is_hdr else body_txt_rgb),
            )

    if full_cap:
        cap_color = RGBColor(215, 222, 232) if dark_bg else RGBColor(90, 98, 110)
        cap_top = min(top + actual_h + 0.06, top + h - cap_h)
        cap_box = slide.shapes.add_textbox(
            Inches(left), Inches(cap_top), Inches(w), Inches(cap_h))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        write_rich_runs(p, full_cap, ladder["caption"], italic=True, color=cap_color)

    report_slide["table"] = dict(rows=nrows, cols=ncols, pt=pt, est_height_in=round(actual_h, 2))


CHART_FALLBACK_HEX = ["#1F4E79", "#D95F02", "#2B8CBE", "#31A354", "#756BB1", "#E6AB02"]


def _theme_palette(theme_colors):
    pal = []
    for k in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        if theme_colors and theme_colors.get(k):
            pal.append(hex_to_rgb(theme_colors[k]))
    for h in CHART_FALLBACK_HEX:
        if len(pal) >= 6:
            break
        pal.append(hex_to_rgb(h))
    return pal


def _parse_chart_num(v):
    """Parse numeric chart value robustly (handles %, commas, units, unicode minus, None/N/A)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and (math.isnan(v) or math.isinf(v))) else float(v)
    s = str(v).strip().replace("−", "-").replace("–", "-")
    if not s or s.lower() in ("-", "—", "n/a", "na", "none", "null", "nan"):
        return None
    s_clean = s.replace(",", "")
    try:
        return float(s_clean)
    except ValueError:
        m = re.search(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", s_clean)
        return float(m.group(0)) if m else None


def place_chart(slide, chart_spec, box, ladder, theme_colors, dark_bg, report_slide):
    """Render an editable native OOXML chart (column/bar/line) styled with template accents."""
    cats = [str(c) for c in (chart_spec.get("categories") or [])]
    series_list = chart_spec.get("series") or []
    if not cats or not series_list:
        return
    left, top, w, h = box
    cap = chart_spec.get("caption") or ""
    src = chart_spec.get("source") or ""
    full_cap = (f"{cap}  [{src}]" if (cap and src) else (cap or (f"[{src}]" if src else ""))).strip()
    cap_h = 0.0
    if full_cap:
        cap_lines = max(1, math.ceil(est_text_w_pt(full_cap, ladder["caption"]) / max(w * 72 - 12, 40)))
        cap_h = cap_lines * ladder["caption"] * 1.25 / 72 + 0.12
    ch_h = max(h - cap_h - 0.04, 1.4)

    ctype_str = (chart_spec.get("type") or "column").lower()
    ctype = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
    }.get(ctype_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

    cdata = CategoryChartData()
    cdata.categories = cats
    clean_series = []
    for ser in series_list:
        name = str(ser.get("name") or f"Series {len(clean_series) + 1}")
        raw_vals = ser.get("values") or []
        vals = []
        for idx in range(len(cats)):
            v = raw_vals[idx] if idx < len(raw_vals) else None
            vals.append(_parse_chart_num(v))
        cdata.add_series(name, vals)
        clean_series.append((name, vals))

    gframe = slide.shapes.add_chart(ctype, Inches(left), Inches(top), Inches(w), Inches(ch_h), cdata)
    chart = gframe.chart
    chart.has_legend = len(clean_series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(11)

    fg_rgb = RGBColor(235, 240, 248) if dark_bg else RGBColor(40, 46, 56)
    grid_rgb = RGBColor(80, 96, 120) if dark_bg else RGBColor(215, 220, 228)
    pal = _theme_palette(theme_colors)
    if dark_bg:
        pal = [tint_rgb(c, 0.28) if (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) < 115 else c for c in pal]

    if chart_spec.get("title"):
        chart.has_title = True
        chart.chart_title.has_text_frame = True
        tf = chart.chart_title.text_frame
        tf.clear()
        write_rich_runs(tf.paragraphs[0], str(chart_spec["title"]), 13.5, base_bold=True, color=fg_rgb)
    else:
        chart.has_title = False

    if chart.has_legend:
        chart.legend.font.color.rgb = fg_rgb

    try:
        cat_ax = chart.category_axis
        cat_ax.tick_labels.font.size = Pt(11)
        cat_ax.tick_labels.font.color.rgb = fg_rgb
        cat_ax.has_major_gridlines = False
    except Exception:
        pass
    try:
        val_ax = chart.value_axis
        val_ax.tick_labels.font.size = Pt(10.5)
        val_ax.tick_labels.font.color.rgb = fg_rgb
        val_ax.has_major_gridlines = True
        val_ax.major_gridlines.format.line.color.rgb = grid_rgb
        if chart_spec.get("unit"):
            val_ax.has_title = True
            val_ax.axis_title.text_frame.text = str(chart_spec["unit"])
            val_ax.axis_title.text_frame.paragraphs[0].font.size = Pt(10.5)
            val_ax.axis_title.text_frame.paragraphs[0].font.color.rgb = fg_rgb
    except Exception:
        pass

    plot = chart.plots[0]
    show_vals = chart_spec.get("show_values", True)
    if show_vals and len(cats) * len(clean_series) <= 18:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(10)
        dl.font.color.rgb = fg_rgb
        if ctype != XL_CHART_TYPE.LINE_MARKERS:
            try:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass
    for idx, ser in enumerate(plot.series):
        c = pal[idx % len(pal)]
        if ctype == XL_CHART_TYPE.LINE_MARKERS:
            ser.format.line.color.rgb = c
            ser.format.line.width = Pt(2.25)
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = c

    if full_cap:
        cap_color = RGBColor(215, 222, 232) if dark_bg else RGBColor(90, 98, 110)
        cap_box = slide.shapes.add_textbox(Inches(left), Inches(top + ch_h + 0.04), Inches(w), Inches(cap_h))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        write_rich_runs(p, full_cap, ladder["caption"], italic=True, color=cap_color)

    report_slide["chart"] = dict(
        type=ctype_str, categories=len(cats), series=len(clean_series),
        title=chart_spec.get("title") or "",
    )


def _prepare_image(p: Path, dark_bg: bool, trim: bool, tmp_files: list):
    """Optionally composite dark-ink transparent PNG onto white card for dark slides, and
    conservatively trim excessive pure-white margins so academic plots fill more usable area.
    Never modifies the original file; returns (path_to_insert, width_px, height_px)."""
    with Image.open(p) as im0:
        im = im0.copy()
    modified = False
    # 1) Dark-slide readability guard for transparent-background scientific figures
    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
    if dark_bg and has_alpha:
        rgba = im.convert("RGBA")
        alpha = rgba.getchannel("A")
        a_min, a_max = alpha.getextrema()
        if a_min < 240 and a_max > 20:
            # Sample visible pixels' luminance via raw RGBA bytes (Pillow 14+ safe)
            small = rgba.resize((min(120, rgba.width), min(120, rgba.height)))
            raw = small.tobytes()
            vis = [
                0.299 * raw[k] + 0.587 * raw[k + 1] + 0.114 * raw[k + 2]
                for k in range(0, len(raw), 4) if raw[k + 3] > 40
            ]
            if vis and (sum(vis) / len(vis)) < 150:
                pad = 10
                card = Image.new("RGB", (rgba.width + 2 * pad, rgba.height + 2 * pad), (255, 255, 255))
                card.paste(rgba, (pad, pad), rgba)
                im = card
                modified = True

    # 2) Conservative pure-white border trim (only when saving >8% width or height)
    if trim and im.width >= 80 and im.height >= 80:
        rgb = im.convert("RGB")
        bg = Image.new("RGB", rgb.size, (255, 255, 255))
        diff = ImageChops.difference(rgb, bg).convert("L")
        mask = diff.point(lambda v: 255 if v > 12 else 0)
        bbox = mask.getbbox()
        if bbox:
            x0, y0, x1, y1 = bbox
            bw, bh = x1 - x0, y1 - y0
            if bw >= 40 and bh >= 40 and (bw < im.width * 0.92 or bh < im.height * 0.92):
                pad = 8
                cx0, cy0 = max(0, x0 - pad), max(0, y0 - pad)
                cx1, cy1 = min(im.width, x1 + pad), min(im.height, y1 + pad)
                im = im.crop((cx0, cy0, cx1, cy1))
                modified = True

    if not modified:
        return p, im.width, im.height
    fd, tmp_p = tempfile.mkstemp(prefix="adaptor_img_", suffix=".png")
    os.close(fd)
    im.save(tmp_p, format="PNG")
    if tmp_files is not None:
        tmp_files.append(tmp_p)
    return Path(tmp_p), im.width, im.height


def place_images(slide, images, box, ladder, dark_bg, report_slide, warnings, tmp_files=None):
    """Contain-fit images into box (left, top, w, h inches); captions below in one strip."""
    left, top, w, h = box
    caps = [f"{im.get('caption', '')}" + (f"  [{im.get('source', '')}]" if im.get("source") else "")
            for im in images]
    non_empty_caps = [c.strip() for c in caps if c.strip()]
    cap_h = 0.0
    if non_empty_caps:
        cap_lines = sum(
            max(1, math.ceil(est_text_w_pt(c, ladder["caption"]) / max(w * 72 - 12, 40)))
            for c in non_empty_caps)
        cap_h = cap_lines * ladder["caption"] * 1.25 / 72 + 0.14
    img_h = max(h - cap_h, 0.7)
    placed = []
    n = len(images)
    if n == 1:
        boxes = [(left, top, w, img_h)]
    else:
        col_w = (w - GUTTER_IN * (n - 1)) / n
        boxes = [(left + i * (col_w + GUTTER_IN), top, col_w, img_h) for i in range(n)]

    max_bottom = top
    for im, b in zip(images, boxes):
        p = Path(im["path"]) if im.get("path") else None
        if p is None or not p.exists():
            warnings.append(f"slide{report_slide['n']}: missing image {im.get('path')!r}")
            tb = slide.shapes.add_textbox(Inches(b[0]), Inches(b[1]), Inches(b[2]), Inches(min(b[3], 1.2)))
            tf = tb.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            write_rich_runs(p0, f"[Image unavailable: {im.get('caption') or im.get('path')}]",
                            13, italic=True, color=RGBColor(180, 60, 60))
            max_bottom = max(max_bottom, b[1] + min(b[3], 1.2))
            placed.append(dict(path=str(im.get("path")), missing=True))
            continue
        use_p, iw, ih = _prepare_image(p, dark_bg=dark_bg, trim=im.get("trim", True), tmp_files=tmp_files)
        bw, bh = b[2], b[3]
        scale = min(bw / iw if iw else 1, bh / ih if ih else 1)
        disp_w, disp_h = iw * scale, ih * scale
        x = b[0] + (bw - disp_w) / 2
        y = b[1] + (bh - disp_h) / 2
        slide.shapes.add_picture(str(use_p), Inches(x), Inches(y), Inches(disp_w), Inches(disp_h))
        max_bottom = max(max_bottom, y + disp_h)
        placed.append(dict(
            path=str(p), native=f"{iw}x{ih}", placed_in=f"{disp_w:.2f}x{disp_h:.2f}",
            eff_dpi=round(iw / disp_w, 1) if disp_w else 0,
            aspect_kept=abs((iw / ih) - (disp_w / disp_h)) < 0.01 if ih and disp_h else True,
            preprocessed=(use_p != p),
        ))

    if non_empty_caps:
        cap_top = min(top + img_h + 0.04, max_bottom + 0.06)
        cap_color = RGBColor(215, 222, 232) if dark_bg else RGBColor(90, 98, 110)
        cap_box = slide.shapes.add_textbox(Inches(left), Inches(cap_top), Inches(w), Inches(cap_h))
        tf = cap_box.text_frame
        tf.word_wrap = True
        tf.clear()
        for i, c in enumerate(non_empty_caps):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.CENTER
            para.space_after = Pt(2)
            write_rich_runs(para, c, ladder["caption"], italic=True, color=cap_color)
    report_slide["images"] = placed
    report_slide["caption_pt"] = ladder["caption"]


def needs_split(bullets, width_in, height_in):
    return len(bullets) >= 4 and est_body_height_in(bullets, width_in, MIN_BODY_PT) > height_in


def expand_splits(plan_slides, ladder, slide_w_in, slide_h_in, allow_split=True):
    """Pre-pass: split over-long content slides into (cont.) continuations unless slide count is locked."""
    if not allow_split:
        return list(plan_slides)
    default_content_h = max(slide_h_in - 2.15, 3.8)
    out = []
    for s in plan_slides:
        out.append(s)
        if s.get("kind", "content") != "content" or s.get("table") or s.get("chart"):
            continue
        bullets = norm_bullets(s.get("bullets"))
        if not bullets:
            continue
        w = slide_w_in - 1.1
        guard = 0
        while needs_split(bullets, w, default_content_h) and guard < 4:
            guard += 1
            half = max(2, len(bullets) // 2)
            head, tail = bullets[:half], bullets[half:]
            s2 = dict(s)
            s2["bullets"] = head
            s = dict(s)
            s["title"] = (s.get("title") or "") + " (cont.)"
            s["bullets"] = tail
            s["notes"] = (s.get("notes") or "") + " [continuation of previous slide]"
            out[-1] = s2
            out.append(s)
            bullets = tail
    return out


def format_evidence_for_notes(ev_list):
    parts = []
    for e in (ev_list or []):
        if isinstance(e, dict):
            loc = str(e.get("locator") or e.get("source") or "").strip()
            q = str(e.get("quote") or "").strip()
            if loc and q:
                parts.append(f'{loc} ("{q}")')
            elif loc:
                parts.append(loc)
        elif e:
            parts.append(str(e).strip())
    return "; ".join(p for p in parts if p)


def load_prs(path):
    """Open a pptx/potx; normalize potx content-type via a temp copy when needed."""
    try:
        return Presentation(str(path)), False, None
    except ValueError as e:
        if "template.main" not in str(e):
            raise
        tmp_fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
        os.close(tmp_fd)
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "[Content_Types].xml":
                    data = data.replace(
                        b"presentationml.template.main+xml",
                        b"presentationml.presentation.main+xml")
                zout.writestr(item, data)
        return Presentation(tmp_name), True, tmp_name


def _decide_figure_mode(s, SW, SH, has_two_cols):
    mode = (s.get("layout_mode") or "auto").lower()
    if mode in ("stacked", "side_by_side"):
        return mode
    imgs = s.get("images") or []
    bullets = norm_bullets(s.get("bullets"))
    if len(imgs) != 1 or not bullets:
        return "stacked"
    if has_two_cols:
        return "side_by_side"
    p = Path(imgs[0]["path"]) if imgs[0].get("path") else None
    if p and p.exists() and SW >= 11.0:
        try:
            with Image.open(p) as im:
                aspect = im.width / max(im.height, 1)
            if aspect < 1.65:
                return "side_by_side"
        except Exception:
            pass
    return "stacked"


def apply_footer_chrome(slide, slide_num, footer_text, show_snum, slide_w=13.33, slide_h=7.5,
                        dark_bg=False, has_bottom_decor=False):
    """Populate existing FOOTER / SLIDE_NUMBER placeholders on non-title slides; only add
    fallback subtle textboxes when the template lacks a static bottom logo/banner strip."""
    if not footer_text and not show_snum:
        return
    did_ftr = False
    did_snum = False
    for sh in slide.shapes:
        if not sh.is_placeholder:
            continue
        pt = ph_type(sh)
        if pt == "FOOTER" and footer_text and sh.has_text_frame:
            tf = sh.text_frame
            tf.clear()
            write_rich_runs(tf.paragraphs[0], footer_text, 10)
            did_ftr = True
        elif pt == "SLIDE_NUMBER" and show_snum and sh.has_text_frame:
            tf = sh.text_frame
            tf.clear()
            write_rich_runs(tf.paragraphs[0], str(slide_num), 10)
            did_snum = True
    if has_bottom_decor:
        return
    muted = RGBColor(190, 200, 215) if dark_bg else RGBColor(115, 122, 134)
    y_ftr = max(slide_h - 0.36, 0.5)
    if footer_text and not did_ftr:
        tb = slide.shapes.add_textbox(Inches(0.55), Inches(y_ftr), Inches(slide_w - 1.8), Inches(0.28))
        tb.name = "__adaptor_footer__"
        tf = tb.text_frame
        tf.word_wrap = True
        tf.clear()
        write_rich_runs(tf.paragraphs[0], footer_text, 9.5, color=muted)
    if show_snum and not did_snum:
        tb = slide.shapes.add_textbox(Inches(slide_w - 1.15), Inches(y_ftr), Inches(0.65), Inches(0.28))
        tb.name = "__adaptor_snum__"
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        write_rich_runs(p, str(slide_num), 9.5, color=muted)


def _pair_comparison_phs(text_phs):
    """Detect PowerPoint's 4-placeholder Comparison layout:
    2 short column-header boxes (h<=1.25in) above 2 tall body boxes (h>=2.0in).
    Returns ((l_hdr, l_body), (r_hdr, r_body)) or None."""
    if len(text_phs) < 4:
        return None
    hdrs = [p for p in text_phs if as_in(p.height, 9) <= 1.25]
    bodies = [p for p in text_phs if as_in(p.height, 0) >= 2.0]
    if len(hdrs) >= 2 and len(bodies) >= 2:
        hdrs.sort(key=lambda p: as_in(p.left, 0))
        bodies.sort(key=lambda p: as_in(p.left, 0))
        return (hdrs[0], bodies[0]), (hdrs[1], bodies[1])
    return None


def _slide_decor_margins(slide, sw, sh):
    """Inspect static non-placeholder edge bands/logos on the slide's layout & master."""
    m_top, m_bot, m_left, m_right = 0.0, 0.0, 0.0, 0.0
    containers = []
    if slide.slide_layout is not None:
        containers.append(slide.slide_layout)
        if slide.slide_layout.slide_master is not None:
            containers.append(slide.slide_layout.slide_master)
    for cont in containers:
        for sp in cont.shapes:
            if sp.is_placeholder:
                continue
            l = as_in(sp.left, -1)
            t = as_in(sp.top, -1)
            w = as_in(sp.width, 0)
            h = as_in(sp.height, 0)
            if l < 0 or t < 0 or w <= 0.05 or h <= 0.05:
                continue
            if w >= sw * 0.85 and h >= sh * 0.85:
                continue
            r, b = l + w, t + h
            if t < sh * 0.16 and h <= sh * 0.18 and b <= sh * 0.24 and w >= sw * 0.25:
                m_top = max(m_top, b)
            if b > sh * 0.84 and h <= sh * 0.18 and t >= sh * 0.78 and w >= sw * 0.12:
                m_bot = max(m_bot, sh - t)
            if l < sw * 0.12 and w <= sw * 0.20 and r <= sw * 0.24 and h >= sh * 0.45:
                m_left = max(m_left, r)
            if r > sw * 0.88 and w <= sw * 0.20 and l >= sw * 0.76 and h >= sh * 0.45:
                m_right = max(m_right, sw - l)
    return min(m_top, 1.4), min(m_bot, 1.1), min(m_left, 1.4), min(m_right, 1.0)


def build(plan_path: Path, report_path: Path):
    plan = json.loads(Path(plan_path).read_text(encoding="utf-8"))
    tpl_raw = plan.get("template")
    out = Path(plan["output"])
    out.parent.mkdir(parents=True, exist_ok=True)
    ladder = dict(DEFAULT_LADDER)
    ladder.update(plan.get("ladder") or {})

    tmp_name = None
    tmp_img_files = []
    if tpl_raw and str(tpl_raw).lower() != "none" and Path(tpl_raw).exists():
        tpl = Path(tpl_raw)
        prs, normalized, tmp_name = load_prs(tpl)
        theme_colors = read_theme_colors(tpl)
    else:
        tpl = None
        prs, normalized = Presentation(), False
        theme_colors = {}

    try:
        SW, SH = Emu(prs.slide_width).inches, Emu(prs.slide_height).inches
        accent_rgb = hex_to_rgb(theme_colors.get("accent1"), (31, 78, 121))
        report = dict(
            template=str(tpl) if tpl else "none",
            output=str(out),
            slide_size=f"{SW:.2f}x{SH:.2f}in",
            theme_colors=theme_colors,
            ladder=ladder,
            slides=[],
            fallbacks=[],
            warnings=[],
            auto_splits=[],
            normalizations=[],
        )
        if tpl is None:
            report["normalizations"].append("no template provided; built with python-pptx default presentation")
        if normalized:
            report["normalizations"].append(f"template {tpl} is potx; content-type normalized to pptx")

        # clear pre-existing template slides unless explicitly kept
        if not plan.get("keep_template_slides", False) and len(prs.slides) > 0:
            n0 = len(prs.slides)
            sldIdLst = prs.slides._sldIdLst
            rid_key = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            for sldId in list(sldIdLst):
                rId = sldId.get(rid_key)
                try:
                    prs.part.drop_rel(rId)
                except Exception:
                    pass
                sldIdLst.remove(sldId)
            report["normalizations"].append(f"removed {n0} pre-existing template slides")

        # Auto-split guard: never violate explicit strict/expected/max slide count
        raw_slides = plan.get("slides", [])
        expected = plan.get("expected_slides")
        max_slides = plan.get("max_slides")
        strict_count = bool(plan.get("strict_slide_count") or (expected and len(raw_slides) == expected))
        cand_splits = expand_splits(raw_slides, ladder, SW, SH, allow_split=not strict_count)
        if max_slides and len(cand_splits) > max_slides and len(raw_slides) <= max_slides:
            slides_plan = list(raw_slides)
            report["warnings"].append(
                f"auto-split suppressed to respect max_slides={max_slides}; tighten long bullet lists if tight")
        else:
            slides_plan = cand_splits
        if len(slides_plan) != len(raw_slides):
            report["auto_splits"].append(
                f"plan had {len(raw_slides)} slides, expanded to {len(slides_plan)} via (cont.) splits")
        if expected and len(slides_plan) != expected:
            report["warnings"].append(
                f"slide-count mismatch: plan has {len(slides_plan)} slides but case expects "
                f"{expected} — adjust plan before delivering")

        # per-slide speaking-time hints when a duration is declared
        time_hints = None
        dur_min = plan.get("duration_min")
        if isinstance(dur_min, (int, float)) and dur_min > 0 and slides_plan:
            weights = {"title": 0.5, "section": 0.25, "figure": 1.8, "two_col": 1.4, "content": 1.2}
            wsum = sum(weights.get(s.get("kind", "content"), 1.0) for s in slides_plan) or 1.0
            total_s = dur_min * 60 * 0.92
            time_hints = [
                max(15, int(round(total_s * weights.get(s.get("kind", "content"), 1.0) / wsum / 5.0) * 5))
                for s in slides_plan
            ]

        footer_text = plan.get("footer")
        show_snum = bool(plan.get("slide_numbers", False))

        for i, s in enumerate(slides_plan, 1):
            kind = s.get("kind", "content")
            bullets = norm_bullets(s.get("bullets"))
            raw_imgs = s.get("images") or []
            images = [{"path": im} if isinstance(im, str) else im for im in raw_imgs if im]
            table_spec = s.get("table")
            chart_spec = s.get("chart")
            # Defensive: a content slide with images (and no table/chart) routes cleanly to figure layout
            if kind == "content" and images and not table_spec and not chart_spec:
                kind = "figure"
            rep = dict(n=i, kind=kind)

            lo, how = select_layout(
                prs, kind, s.get("layout_hint"), s.get("layout_name"),
                n_images=len(images), has_bullets=bool(bullets), has_table=bool(table_spec or chart_spec))
            rep["layout_used"], rep["layout_select"] = lo.name, how
            if how in ("first-layout-fallback", "exact-name-missing"):
                report["fallbacks"].append(f"slide{i}:{how}:{lo.name}")
            slide = prs.slides.add_slide(lo)
            dark_bg = is_dark_slide(slide, theme_colors)
            ph_text_rgb = _dark_slide_text_color(slide, dark_bg, theme_colors)
            heading_rgb = accent_rgb
            if dark_bg and accent_rgb is not None:
                if (0.299 * accent_rgb[0] + 0.587 * accent_rgb[1] + 0.114 * accent_rgb[2]) < 160:
                    heading_rgb = tint_rgb(accent_rgb, 0.58)
            used_ph_ids = set()

            # title
            tph, tinfo = None, None
            if s.get("title"):
                t_sz = (ladder["section"] if kind == "section" else
                        ladder["title"] if kind == "title" else ladder["slide_title"])
                tph, tinfo = fill_title(slide, s["title"], t_sz, color=ph_text_rgb, slide_w=SW)
            if tinfo:
                rep["title_ph"] = tinfo
                used_ph_ids.add(tph.shape_id)
            elif s.get("title"):
                report["warnings"].append(f"slide{i}: no title placeholder; title put in notes only")

            # compute safe content box (respecting template placeholder envelope & edge decor bands)
            d_top, d_bot, d_left, d_right = _slide_decor_margins(slide, SW, SH)
            eff_bot_margin = max(BOTTOM_MARGIN_IN, d_bot + 0.08)
            eff_left = max(0.55, d_left + 0.14)
            eff_right = max(0.55, d_right + 0.14)
            tph_ref = tph or find_ph(slide.shapes, TITLEISH)
            content_top = (as_in(tph_ref.height, 1.15) + as_in(tph_ref.top, 0.3) + 0.14
                           if tph_ref is not None else max(1.35, d_top + 0.14))
            content_top = min(max(content_top, d_top + 0.10), SH * 0.35)
            content_box = (eff_left, content_top,
                           max(SW - eff_left - eff_right, 2.5),
                           max(SH - content_top - eff_bot_margin - 0.15, 2.0))

            text_phs = find_phs(slide.shapes, TEXT_BODYISH)
            pic_phs = find_phs(slide.shapes, PIC_PH)
            sub_ph = find_ph(slide.shapes, SUBISH)
            all_content_phs = text_phs + pic_phs
            if kind in ("content", "two_col", "figure") and all_content_phs:
                xs0 = [as_in(p.left, eff_left) for p in all_content_phs]
                ys0 = [as_in(p.top, content_top) for p in all_content_phs]
                xs1 = [as_in(p.left, eff_left) + as_in(p.width, 0) for p in all_content_phs]
                ys1 = [as_in(p.top, content_top) + as_in(p.height, 0) for p in all_content_phs]
                env_l = max(0.3, eff_left if d_left > 0.2 else 0.3, min(xs0))
                env_t = max(content_top, min(ys0))
                env_r = min(SW - 0.3, SW - eff_right if d_right > 0.2 else SW - 0.3, max(xs1))
                env_b = min(SH - eff_bot_margin, max(ys1))
                # Only adopt placeholder envelope when it spans a genuine full content canvas (>=65% W)
                if (env_r - env_l) >= SW * 0.65 and (env_b - env_t) >= SH * 0.35:
                    content_box = (env_l, env_t, env_r - env_l, env_b - env_t)
            body_phs_used = []

            # --- TITLE SLIDE ---
            if kind == "title":
                if s.get("subtitle"):
                    target_sub = sub_ph or (text_phs[0] if text_phs else None)
                    if target_sub is not None:
                        if as_in(target_sub.width, 0.0) <= 0.1 or as_in(target_sub.height, 0.0) <= 0.1:
                            pin_ph_geometry(target_sub, (0.8, SH * 0.52, max(SW - 1.6, 4.0), 1.6))
                        tf = target_sub.text_frame
                        tf.word_wrap = True
                        tf.clear()
                        sub_lines = [ln.strip() for ln in str(s["subtitle"]).splitlines() if ln.strip()]
                        meta_color = RGBColor(200, 212, 228) if dark_bg else RGBColor(92, 102, 118)
                        for idx_ln, ln in enumerate(sub_lines):
                            para = tf.paragraphs[0] if idx_ln == 0 else tf.add_paragraph()
                            if idx_ln == 0:
                                write_rich_runs(para, ln, ladder["subtitle"], color=ph_text_rgb)
                            else:
                                para.space_before = Pt(6)
                                write_rich_runs(para, ln, max(14, ladder["subtitle"] - 4), color=meta_color)
                        rep["subtitle_ph"] = dict(idx=target_sub.placeholder_format.idx, type=ph_type(target_sub))
                        used_ph_ids.add(target_sub.shape_id)

            # --- SECTION SLIDE ---
            elif kind == "section":
                if s.get("subtitle") or bullets:
                    target = sub_ph or (text_phs[0] if text_phs else None)
                    if target is not None:
                        if s.get("subtitle"):
                            tf = target.text_frame
                            tf.clear()
                            write_rich_runs(tf.paragraphs[0], s["subtitle"], ladder["subtitle"], color=ph_text_rgb)
                        else:
                            fill_bullets(target, bullets, ladder, as_in(target.width, content_box[2]),
                                         as_in(target.height, content_box[3]), rep, text_rgb=ph_text_rgb)
                        used_ph_ids.add(target.shape_id)
                        body_phs_used.append(dict(idx=target.placeholder_format.idx, type=ph_type(target)))

            # --- TWO-COLUMN SLIDE (supports Comparison 4-ph, Two Content 2-ph, or 1/0-ph fallback;
            #     symmetric bullets / images / table on EITHER left or right column) ---
            elif kind == "two_col":
                cols = s.get("columns") or []
                left_col = s.get("left") or (cols[0] if len(cols) >= 1 else {})
                right_col = s.get("right") or (cols[1] if len(cols) >= 2 else {})
                col_w = (content_box[2] - GUTTER_IN) / 2
                l_box = (content_box[0], content_box[1], col_w, content_box[3])
                r_box = (content_box[0] + col_w + GUTTER_IN, content_box[1], col_w, content_box[3])

                comp_pairs = _pair_comparison_phs(text_phs)
                if comp_pairs is not None:
                    (l_hdr_ph, l_body_ph), (r_hdr_ph, r_body_ph) = comp_pairs
                    col_slots = [
                        (left_col, l_hdr_ph, l_body_ph,
                         (as_in(l_body_ph.left, l_box[0]), as_in(l_body_ph.top, l_box[1]),
                          as_in(l_body_ph.width, l_box[2]), as_in(l_body_ph.height, l_box[3]))),
                        (right_col, r_hdr_ph, r_body_ph,
                         (as_in(r_body_ph.left, r_box[0]), as_in(r_body_ph.top, r_box[1]),
                          as_in(r_body_ph.width, r_box[2]), as_in(r_body_ph.height, r_box[3]))),
                    ]
                elif len(text_phs) >= 2:
                    lph, rph = text_phs[0], text_phs[1]
                    col_slots = [
                        (left_col, None, lph,
                         (as_in(lph.left, l_box[0]), as_in(lph.top, l_box[1]),
                          as_in(lph.width, l_box[2]), as_in(lph.height, l_box[3]))),
                        (right_col, None, rph,
                         (as_in(rph.left, r_box[0]), as_in(rph.top, r_box[1]),
                          as_in(rph.width, r_box[2]), as_in(rph.height, r_box[3]))),
                    ]
                else:
                    single_ph = text_phs[0] if len(text_phs) == 1 else None
                    l_wants_text = bool(norm_bullets(left_col.get("bullets")) or left_col.get("heading")) and not (left_col.get("images") or left_col.get("table") or left_col.get("chart"))
                    r_wants_text = bool(norm_bullets(right_col.get("bullets")) or right_col.get("heading")) and not (right_col.get("images") or right_col.get("table") or right_col.get("chart"))
                    l_ph_use = single_ph if l_wants_text else None
                    r_ph_use = single_ph if (not l_wants_text and r_wants_text) else None
                    col_slots = [
                        (left_col, None, l_ph_use, l_box),
                        (right_col, None, r_ph_use, r_box),
                    ]

                for col_spec, hdr_ph, bph, cbox in col_slots:
                    c_bullets = norm_bullets(col_spec.get("bullets"))
                    c_images = col_spec.get("images") or []
                    c_table = col_spec.get("table")
                    c_chart = col_spec.get("chart")
                    c_heading = col_spec.get("heading")
                    if hdr_ph is not None and c_heading:
                        tf = hdr_ph.text_frame
                        tf.word_wrap = True
                        tf.clear()
                        p0 = tf.paragraphs[0]
                        _strip_bullets(p0)
                        write_rich_runs(p0, c_heading, min(ladder["body"][0] + 1, 20),
                                        base_bold=True, color=heading_rgb)
                        used_ph_ids.add(hdr_ph.shape_id)
                        body_phs_used.append(dict(idx=hdr_ph.placeholder_format.idx, type=ph_type(hdr_ph)))
                        c_heading = None  # already placed in dedicated Comparison header placeholder
                    if c_images or c_table or c_chart:
                        vis_box = cbox
                        if c_heading:
                            # Render column subheading cleanly above figure/table/chart when no Comparison header ph exists
                            hh = 0.36
                            htb = slide.shapes.add_textbox(Inches(cbox[0]), Inches(cbox[1]), Inches(cbox[2]), Inches(hh))
                            htf = htb.text_frame
                            htf.word_wrap = True
                            htf.clear()
                            hp0 = htf.paragraphs[0]
                            hp0.alignment = PP_ALIGN.CENTER
                            write_rich_runs(hp0, c_heading, min(ladder["body"][0] - 1, 16),
                                            base_bold=True, color=heading_rgb)
                            vis_box = (cbox[0], cbox[1] + hh + 0.06, cbox[2], max(cbox[3] - hh - 0.06, 1.2))
                        if c_images:
                            place_images(slide, c_images, vis_box, ladder, dark_bg, rep, report["warnings"], tmp_files=tmp_img_files)
                        elif c_table:
                            place_table(slide, c_table, vis_box, ladder, accent_rgb, dark_bg, rep)
                        elif c_chart:
                            place_chart(slide, c_chart, vis_box, ladder, theme_colors, dark_bg, rep)
                    elif c_bullets or c_heading:
                        if bph is not None:
                            if len(text_phs) == 1:
                                pin_ph_geometry(bph, cbox)
                                bph.left, bph.top, bph.width, bph.height = (
                                    Inches(cbox[0]), Inches(cbox[1]), Inches(cbox[2]), Inches(cbox[3]))
                            fill_bullets(bph, c_bullets, ladder, cbox[2], cbox[3], rep,
                                         heading=c_heading, accent_rgb=heading_rgb, text_rgb=ph_text_rgb)
                            used_ph_ids.add(bph.shape_id)
                            body_phs_used.append(dict(idx=bph.placeholder_format.idx, type=ph_type(bph)))
                        else:
                            tb = slide.shapes.add_textbox(Inches(cbox[0]), Inches(cbox[1]), Inches(cbox[2]), Inches(cbox[3]))
                            fill_bullets(tb, c_bullets, ladder, cbox[2], cbox[3], rep,
                                         heading=c_heading, accent_rgb=heading_rgb,
                                         text_rgb=(RGBColor(240, 244, 248) if dark_bg else None))
                            rep["fallbacks_manual_textbox"] = True

            # --- FIGURE SLIDE ---
            elif kind == "figure":
                has_two_cols = (len(text_phs) >= 2) or (len(text_phs) >= 1 and len(pic_phs) >= 1)
                fig_mode = _decide_figure_mode(s, SW, SH, has_two_cols)
                rep["figure_layout_mode"] = fig_mode

                if fig_mode == "side_by_side" and bullets and images:
                    # Only trust native (bph, pph) geometry if genuinely side-by-side & tall enough
                    native_sbs = False
                    if len(text_phs) >= 1 and len(pic_phs) >= 1:
                        bph0, pph0 = text_phs[0], pic_phs[0]
                        if (abs(as_in(bph0.left, 0) - as_in(pph0.left, 0)) > 1.5
                                and as_in(bph0.height, 0) >= 2.2 and as_in(pph0.height, 0) >= 2.2):
                            native_sbs = True
                    if native_sbs:
                        bph, pph = text_phs[0], pic_phs[0]
                        fill_bullets(bph, bullets, ladder, as_in(bph.width, content_box[2] * 0.4),
                                     as_in(bph.height, content_box[3]), rep, text_rgb=ph_text_rgb)
                        used_ph_ids.add(bph.shape_id)
                        body_phs_used.append(dict(idx=bph.placeholder_format.idx, type=ph_type(bph)))
                        img_box = (as_in(pph.left, content_box[0] + content_box[2] * 0.42),
                                   as_in(pph.top, content_box[1]),
                                   as_in(pph.width, content_box[2] * 0.56),
                                   as_in(pph.height, content_box[3]))
                        place_images(slide, images, img_box, ladder, dark_bg, rep, report["warnings"], tmp_files=tmp_img_files)
                    else:
                        left_w = content_box[2] * 0.40
                        right_w = content_box[2] - left_w - GUTTER_IN
                        l_box = (content_box[0], content_box[1], left_w, content_box[3])
                        r_box = (content_box[0] + left_w + GUTTER_IN, content_box[1], right_w, content_box[3])
                        if text_phs:
                            bph = text_phs[0]
                            pin_ph_geometry(bph, l_box)
                            bph.left, bph.top, bph.width, bph.height = (
                                Inches(l_box[0]), Inches(l_box[1]), Inches(l_box[2]), Inches(l_box[3]))
                            fill_bullets(bph, bullets, ladder, left_w, content_box[3], rep, text_rgb=ph_text_rgb)
                            used_ph_ids.add(bph.shape_id)
                            body_phs_used.append(dict(idx=bph.placeholder_format.idx, type=ph_type(bph)))
                        else:
                            tb = slide.shapes.add_textbox(Inches(l_box[0]), Inches(l_box[1]), Inches(l_box[2]), Inches(l_box[3]))
                            fill_bullets(tb, bullets, ladder, left_w, content_box[3], rep,
                                         text_rgb=(RGBColor(240, 244, 248) if dark_bg else None))
                            rep["fallbacks_manual_textbox"] = True
                        place_images(slide, images, r_box, ladder, dark_bg, rep, report["warnings"], tmp_files=tmp_img_files)
                else:
                    # stacked mode: optional takeaway strip on top, images below
                    img_box = content_box
                    if bullets:
                        if text_phs:
                            bph = text_phs[0]
                            gl, gt, gw, gh0 = pin_ph_geometry(bph, content_box)
                            sz = ladder["body"][0]
                            est = est_body_height_in(bullets, gw, sz)
                            body_h = min(gh0, 1.85, max(est + 0.16, 0.45))
                            bph.height = Inches(body_h)
                            fill_bullets(bph, bullets, ladder, gw, body_h, rep, text_rgb=ph_text_rgb)
                            used_ph_ids.add(bph.shape_id)
                            body_phs_used.append(dict(idx=bph.placeholder_format.idx, type=ph_type(bph)))
                            img_top = max(gt + body_h + 0.14, content_box[1] + body_h + 0.14)
                            img_box = (content_box[0], img_top, content_box[2],
                                       max(SH - img_top - eff_bot_margin - 0.10, 1.2))
                        else:
                            est = est_body_height_in(bullets, content_box[2], ladder["body"][0])
                            body_h = min(1.85, max(est + 0.16, 0.45))
                            tb = slide.shapes.add_textbox(Inches(content_box[0]), Inches(content_box[1]),
                                                          Inches(content_box[2]), Inches(body_h))
                            fill_bullets(tb, bullets, ladder, content_box[2], body_h, rep,
                                         text_rgb=(RGBColor(240, 244, 248) if dark_bg else None))
                            rep["fallbacks_manual_textbox"] = True
                            img_top = content_box[1] + body_h + 0.14
                            img_box = (content_box[0], img_top, content_box[2],
                                       max(SH - img_top - eff_bot_margin - 0.10, 1.2))
                    elif len(images) == 1 and pic_phs:
                        pph = pic_phs[0]
                        img_box = (as_in(pph.left, content_box[0]), as_in(pph.top, content_box[1]),
                                   as_in(pph.width, content_box[2]), as_in(pph.height, content_box[3]))
                    if images:
                        place_images(slide, images, img_box, ladder, dark_bg, rep, report["warnings"], tmp_files=tmp_img_files)
                    elif chart_spec:
                        place_chart(slide, chart_spec, img_box, ladder, theme_colors, dark_bg, rep)

            # --- CONTENT SLIDE (bullets and/or native table or native chart) ---
            else:
                body_ph = text_phs[0] if text_phs else None
                visual_spec = table_spec or chart_spec
                if bullets and visual_spec:
                    # Top concise takeaway strip + native table/chart below
                    if body_ph is not None:
                        gl, gt, gw, gh0 = pin_ph_geometry(body_ph, content_box)
                        est = est_body_height_in(bullets, gw, ladder["body"][0])
                        body_h = min(gh0 * 0.44, 1.85, max(est + 0.24, 0.55))
                        body_ph.height = Inches(body_h)
                        fill_bullets(body_ph, bullets, ladder, gw, body_h, rep, text_rgb=ph_text_rgb)
                        used_ph_ids.add(body_ph.shape_id)
                        body_phs_used.append(dict(idx=body_ph.placeholder_format.idx, type=ph_type(body_ph)))
                        t_top = gt + body_h + 0.22
                        t_box = (content_box[0], t_top, content_box[2],
                                 max(SH - t_top - eff_bot_margin - 0.10, 1.5))
                    else:
                        est = est_body_height_in(bullets, content_box[2], ladder["body"][0])
                        body_h = min(1.85, max(est + 0.24, 0.55))
                        tb = slide.shapes.add_textbox(Inches(content_box[0]), Inches(content_box[1]),
                                                      Inches(content_box[2]), Inches(body_h))
                        fill_bullets(tb, bullets, ladder, content_box[2], body_h, rep,
                                     text_rgb=(RGBColor(240, 244, 248) if dark_bg else None))
                        rep["fallbacks_manual_textbox"] = True
                        t_top = content_box[1] + body_h + 0.22
                        t_box = (content_box[0], t_top, content_box[2],
                                 max(SH - t_top - eff_bot_margin - 0.10, 1.5))
                    if table_spec:
                        place_table(slide, table_spec, t_box, ladder, accent_rgb, dark_bg, rep)
                    elif chart_spec:
                        place_chart(slide, chart_spec, t_box, ladder, theme_colors, dark_bg, rep)
                elif table_spec:
                    place_table(slide, table_spec, content_box, ladder, accent_rgb, dark_bg, rep)
                elif chart_spec:
                    place_chart(slide, chart_spec, content_box, ladder, theme_colors, dark_bg, rep)
                elif bullets and body_ph is not None:
                    bw = as_in(body_ph.width, 0.0)
                    bh = as_in(body_ph.height, 0.0)
                    if bw <= 0.1 or bh <= 0.1:
                        _, _, bw, bh = pin_ph_geometry(body_ph, content_box)
                    fill_bullets(body_ph, bullets, ladder, bw, bh, rep, text_rgb=ph_text_rgb)
                    used_ph_ids.add(body_ph.shape_id)
                    body_phs_used.append(dict(idx=body_ph.placeholder_format.idx, type=ph_type(body_ph)))
                elif bullets:
                    report["warnings"].append(f"slide{i}: no body placeholder; manual textbox used")
                    tb = slide.shapes.add_textbox(Inches(content_box[0]), Inches(content_box[1]),
                                                  Inches(content_box[2]), Inches(content_box[3]))
                    fill_bullets(tb, bullets, ladder, content_box[2], content_box[3], rep,
                                 text_rgb=(RGBColor(240, 244, 248) if dark_bg else None))
                    rep["fallbacks_manual_textbox"] = True
                if images:
                    # defensive: if a content slide also passed images without kind=figure
                    place_images(slide, images, content_box, ladder, dark_bg, rep, report["warnings"], tmp_files=tmp_img_files)

            if body_phs_used:
                rep["body_ph"] = body_phs_used[0] if len(body_phs_used) == 1 else body_phs_used

            # Prune unused EMPTY content/picture placeholders so no ghost box overlaps figures/tables
            for sh in list(slide.shapes):
                if not sh.is_placeholder or sh.shape_id in used_ph_ids:
                    continue
                pt = ph_type(sh)
                if pt in ALL_CONTENT_PH or (kind != "title" and pt in SUBISH):
                    if not (sh.has_text_frame and sh.text_frame.text.strip()):
                        remove_shape(sh)

            # Consistent footer/slide number on non-title slides when requested
            if kind != "title":
                apply_footer_chrome(slide, i, footer_text, show_snum, slide_w=SW, slide_h=SH,
                                    dark_bg=dark_bg, has_bottom_decor=(d_bot >= 0.25))

            # notes = speaker script + structured evidence + time hint
            raw_ev = s.get("evidence") or []
            ev_list = raw_ev if isinstance(raw_ev, list) else [raw_ev]
            ev_str = format_evidence_for_notes(ev_list)
            if s.get("notes") or ev_str:
                ns = slide.notes_slide.notes_text_frame
                parts = []
                if s.get("notes"):
                    parts.append(s["notes"])
                if ev_str:
                    parts.append("Sources | 证据: " + ev_str)
                if time_hints:
                    parts.append(f"[建议用时 ~{time_hints[i - 1]}s]")
                ns.text = "\n".join(parts)
                rep["notes"] = True
                if time_hints:
                    rep["time_hint_s"] = time_hints[i - 1]

            # authoritative per-slide record for write_result.py
            rep["claim"] = s.get("title")
            rep["evidence"] = ev_list
            rep["notes_present"] = bool(s.get("notes"))
            all_imgs = list(images)
            if kind == "two_col":
                cols = s.get("columns") or []
                lc = s.get("left") or (cols[0] if len(cols) >= 1 else {})
                rc = s.get("right") or (cols[1] if len(cols) >= 2 else {})
                all_imgs.extend(lc.get("images") or [])
                all_imgs.extend(rc.get("images") or [])
            if all_imgs:
                rep["figure_refs"] = [
                    {"caption": im.get("caption"), "source": im.get("source"), "asset": im.get("path")}
                    for im in all_imgs
                ]
            report["slides"].append(rep)

        prs.save(str(out))
        report["slide_count"] = len(prs.slides)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"OK {report['slide_count']} slides -> {out}; report -> {report_path}")
        if report["warnings"]:
            print("WARNINGS:", *report["warnings"], sep="\n  ")
        return 0
    finally:
        if tmp_name and os.path.exists(tmp_name):
            try:
                os.remove(tmp_name)
            except OSError:
                pass
        for tfp in tmp_img_files:
            if os.path.exists(tfp):
                try:
                    os.remove(tfp)
                except OSError:
                    pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("plan")
    ap.add_argument("--out", default="build_report.json")
    a = ap.parse_args()
    return build(Path(a.plan), Path(a.out))


if __name__ == "__main__":
    sys.exit(main())
