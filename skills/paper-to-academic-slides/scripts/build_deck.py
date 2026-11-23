#!/usr/bin/env python3
"""Build an editable academic deck from slides.json using a template's layouts
or by adapting an existing starting deck (--reuse).

Usage:
  python3 build_deck.py slides.json --template T.pptx --out deck.pptx [options]

Core principles:
  * Non-reuse with --template: clears template sample slides while preserving all
    masters, layouts, theme colors/fonts, and footers.
  * Reuse (--reuse): edits existing slides in-place and removes stale non-footer
    content shapes before inserting updated figures/tables so nothing collides.
  * Respects template placeholders; prefers Two Content layout when both bullets
    and a figure/table are present; never crops scientific figures (aspect-fit).
  * Renders inline **bold** markdown spans as real bold runs; auto-scales text
    (18->11pt) and tables (12->10pt) dynamically against actual title/footer bounds.
  * Clean 16:9 placeholder geometry when no template is supplied (blank mode).
"""
import argparse
import json
import os
import re
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

FONT_SIZES = [18, 16, 14, 12, 11, 10.5]
GRAY = (0x59, 0x59, 0x59)
LIGHT_GRAY = (0xD4, 0xD8, 0xE0)
BLANK_NAVY = (0x1B, 0x36, 0x5D)
BLANK_ACCENT = (0x2B, 0x6C, 0xB0)
# PP_PLACEHOLDER enum values (python-pptx)
PH_TITLE, PH_CTITLE, PH_BODY, PH_SUB, PH_VBODY, PH_OBJ = 1, 3, 2, 4, 6, 7
PH_DATE, PH_SLIDENUM, PH_FOOTER, PH_PIC = 16, 13, 15, 18
TITLE_TYPES = (PH_TITLE, PH_CTITLE)
BODY_TYPES = (PH_BODY, PH_SUB, PH_VBODY, PH_OBJ)
FOOTER_TYPES = (PH_DATE, PH_SLIDENUM, PH_FOOTER)
GHOST_CLEANABLE_TYPES = (PH_BODY, PH_SUB, PH_VBODY, PH_OBJ, PH_PIC)
EMU_IN = 914400.0
BOLD_SPLIT_RE = re.compile(r"(\*\*[^*]+\*\*)")
NUM_UNIT_RE = re.compile(
    r"(\d(?:\.\d+)?)\s+(meV/A|meV/Å|kcal/mol|kJ/mol|GFLOPs|TFLOPs|MFLOPs|FLOPs|"
    r"tok/s|tokens/s|fps|MHz|GHz|Hz|GB|MB|KB|TB|ms|ns|us|μs|eV|mW|nm|μm|mm|cm|km|dB|"
    r"deg|degrees|%|×|x\b|[mMsWAK]\b)"
)
_WIDTH_SAFETY = 1.0
_BLANK_LATIN_FONT = None
_BLANK_CJK_FONT = None


def discover_installed_fonts():
    """Query fontconfig (fc-list) once to discover installed font families on this host."""
    import shutil
    import subprocess
    if not shutil.which("fc-list"):
        return set(), {}
    try:
        r = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True, timeout=8)
        fams_low, fams_orig = set(), {}
        for ln in (r.stdout or "").splitlines():
            for part in ln.split(","):
                p = part.strip()
                if p:
                    lp = p.lower()
                    fams_low.add(lp)
                    fams_orig.setdefault(lp, p)
        return fams_low, fams_orig
    except Exception:
        return set(), {}


def _read_template_theme_fonts(prs):
    fonts = {}
    try:
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if "theme" in str(rel.reltype).lower():
                xml = rel.target_part.blob.decode("utf-8", "ignore")
                for key, pat in (
                    ("major_latin", r"<a:majorFont>.*?<a:latin[^>]*typeface=\"([^\"]+)\""),
                    ("minor_latin", r"<a:minorFont>.*?<a:latin[^>]*typeface=\"([^\"]+)\""),
                    ("major_ea", r"<a:majorFont>.*?<a:ea[^>]*typeface=\"([^\"]+)\""),
                    ("minor_ea", r"<a:minorFont>.*?<a:ea[^>]*typeface=\"([^\"]+)\""),
                ):
                    m = re.search(pat, xml, re.S)
                    if m and m.group(1).strip():
                        fonts[key] = m.group(1).strip()
                break
    except Exception:
        pass
    return fonts


def init_font_strategy(prs, is_blank):
    """Configure metric-safe font selection (blank mode) or width safety compensation
    (template mode when template theme fonts are absent from the host OS)."""
    global _WIDTH_SAFETY, _BLANK_LATIN_FONT, _BLANK_CJK_FONT
    _WIDTH_SAFETY = 1.0
    _BLANK_LATIN_FONT = None
    _BLANK_CJK_FONT = None
    fams_low, fams_orig = discover_installed_fonts()
    audit = {
        "installed_font_count": len(fams_low),
        "template_theme_fonts": {},
        "missing_template_fonts": [],
        "metric_safety_factor": 1.0,
        "blank_latin_font": None,
        "blank_cjk_font": None,
    }
    if not fams_low:
        return audit
    if is_blank:
        for cand in ("carlito", "liberation sans", "noto sans", "dejavu sans"):
            if cand in fams_low:
                _BLANK_LATIN_FONT = fams_orig[cand]
                break
        for cand in ("noto sans cjk sc", "noto sans sc", "wenquanyi micro hei", "droid sans fallback"):
            if cand in fams_low:
                _BLANK_CJK_FONT = fams_orig[cand]
                break
        audit["blank_latin_font"] = _BLANK_LATIN_FONT
        audit["blank_cjk_font"] = _BLANK_CJK_FONT
    else:
        tfonts = _read_template_theme_fonts(prs)
        audit["template_theme_fonts"] = tfonts
        missing = []
        for k in ("major_latin", "minor_latin", "major_ea", "minor_ea"):
            fv = (tfonts.get(k) or "").strip()
            if fv and fv.lower() not in fams_low:
                if not any(fl.startswith(fv.lower()) for fl in fams_low):
                    missing.append(fv)
        missing = sorted(set(missing))
        audit["missing_template_fonts"] = missing
        if missing:
            _WIDTH_SAFETY = 1.08
            audit["metric_safety_factor"] = 1.08
    return audit


def _apply_blank_font(run, text, is_blank):
    """In blank mode only, bind installed Latin/CJK font so LibreOffice PDF export has zero drift."""
    if not is_blank:
        return
    has_cjk = any(ord(ch) > 0x2E7F for ch in (text or ""))
    chosen = _BLANK_CJK_FONT if (has_cjk and _BLANK_CJK_FONT) else _BLANK_LATIN_FONT
    if chosen:
        try:
            run.font.name = chosen
        except Exception:
            pass


def glue_num_units(text):
    """Replace breaking space between a number and its scientific unit with NBSP
    so '53.8 m' or '14.2 meV/A' never splits across lines."""
    if not text:
        return text
    return NUM_UNIT_RE.sub("\\1\u00A0\\2", text)


def est_text_width(text, pt):
    w = 0.0
    for ch in text:
        o = ord(ch)
        if o > 0x2E7F:
            w += 1.02
        elif ch in " ilj.,:;'|!":
            w += 0.30
        elif ch.isupper() or ch.isdigit():
            w += 0.62
        else:
            w += 0.52
    return w * _WIDTH_SAFETY * pt / 72.0


def est_lines(text, pt, width_in):
    if width_in <= 0.2:
        width_in = 6.0
    total = 0
    for para in text.split("\n"):
        clean = para.replace("**", "")
        if not clean.strip():
            total += 1
            continue
        # Handle CJK (no spaces) and Latin (space-separated) cleanly
        tokens = re.findall(r"\S+\s*", clean)
        cur, lines = 0.0, 1
        for tok in tokens:
            wl = est_text_width(tok, pt)
            if wl > width_in:
                # long CJK run or unbroken token: break by width
                if cur > 0:
                    lines += 1
                    cur = 0.0
                extra = max(1, int(wl / width_in))
                lines += extra
                cur = wl - extra * width_in
            elif cur + wl > width_in and cur > 0:
                lines += 1
                cur = wl
            else:
                cur += wl
        total += lines
    return total


def remove_shape(shape):
    try:
        el = shape._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    except Exception:
        pass


def drop_all_slides(prs):
    """Remove sample slides from a template while keeping masters/layouts intact."""
    xml_slides = prs.slides._sldIdLst
    ids = list(xml_slides)
    for sldId in ids:
        rId = sldId.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass
        xml_slides.remove(sldId)
    return len(ids)


def body_placeholders_sorted(slide, is_cover=False):
    """Return body placeholders sorted left-to-right, prioritizing main content-sized
    BODY/OBJECT placeholders over tiny subtitle/header boxes on non-cover slides."""
    phs = [ph for ph in slide.placeholders
           if ph.placeholder_format.idx != 0 and ph.placeholder_format.type in BODY_TYPES]
    def key_fn(p):
        try:
            h_in = (p.height or 0) / EMU_IN
            small_pen = 0 if (is_cover or h_in >= 1.4) else 1
            sub_pen = 0 if (is_cover or p.placeholder_format.type != PH_SUB) else 1
            return (small_pen, sub_pen, p.left or 0, p.top or 0, p.placeholder_format.idx)
        except Exception:
            return (0, 0, 0, 0, p.placeholder_format.idx)
    phs.sort(key=key_fn)
    return phs


def placeholder_by_role(slide, role):
    if role == "title":
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 0 or ph.placeholder_format.type in TITLE_TYPES:
                return ph
        return None
    if role == "body":
        b = body_placeholders_sorted(slide)
        return b[0] if b else None
    if role == "pic":
        for ph in slide.placeholders:
            if ph.placeholder_format.type == PH_PIC:
                return ph
        return None
    return None


def pick_layout(prs, need_two_col=False, need_picture=False, is_title=False):
    if is_title:
        for i, layout in enumerate(prs.slide_layouts):
            types = [p.placeholder_format.type for p in layout.placeholders]
            if PH_CTITLE in types:
                return i
    scored = []
    for i, layout in enumerate(prs.slide_layouts):
        phs = layout.placeholders
        has_title = any(p.placeholder_format.idx == 0 or
                        p.placeholder_format.type in TITLE_TYPES or
                        "title" in p.name.lower() for p in phs)
        bodies = [p for p in phs if p.placeholder_format.idx != 0 and
                  p.placeholder_format.type in BODY_TYPES]
        pics = [p for p in phs if p.placeholder_format.type == PH_PIC]

        def area(ph):
            try:
                return (ph.width / EMU_IN) * (ph.height / EMU_IN)
            except Exception:
                return 0.0

        def is_two_side_by_side(bs):
            if len(bs) < 2:
                return False
            try:
                pair = sorted(bs[:2], key=lambda b: b.left)
                x0, x1 = pair[0].left / EMU_IN, pair[1].left / EMU_IN
                y0, y1 = pair[0].top / EMU_IN, pair[1].top / EMU_IN
                w0, w1 = pair[0].width / EMU_IN, pair[1].width / EMU_IN
                ratio = min(w0, w1) / max(w0, w1, 0.1)
                return (x1 - x0) >= 2.5 and min(w0, w1) >= 2.8 and abs(y1 - y0) <= 0.6 and ratio >= 0.65
            except Exception:
                return False

        score = 0.0
        lname = (layout.name or "").lower()
        if "vertical" in lname:
            score -= 5.0
        if "blank" in lname:
            score -= 2.0
        if "caption" in lname:
            score -= 2.5
        if is_title:
            score += 5.0 if has_title else 0.0
            score += 2.0 if bodies else 0.0
            score -= 3.0 if pics else 0.0
            if bodies:
                score += min(area(bodies[0]), 30.0) * 0.1
        else:
            types_l = [p.placeholder_format.type for p in phs]
            if PH_CTITLE in types_l or "title slide" in lname or "section header" in lname:
                score -= 6.0  # never use cover/section-divider layouts for regular content slides
            score += (2.5 if has_title else 0.0) + (4.0 if bodies else 0.0)
            if need_two_col and is_two_side_by_side(bodies):
                score += 5.0  # native Two Content layout fits left-text + right-fig/table
                if "two content" in lname:
                    score += 1.5
            elif need_picture and pics and bodies and "caption" not in lname:
                score += 3.5
            elif len(bodies) == 1 and not need_two_col:
                score += 1.5
                if bodies[0].placeholder_format.type == PH_BODY:
                    score += 0.5
            if bodies:
                score += min(area(bodies[0]), 30.0) * 0.15
            if has_title:
                tps = [p for p in phs if p.placeholder_format.idx == 0 or
                       p.placeholder_format.type in TITLE_TYPES]
                if tps:
                    score += min(area(tps[0]), 25.0) * 0.08
        scored.append((score, i))
    scored.sort(key=lambda t: (-t[0], t[1]))
    return scored[0][1] if scored else 0


def set_placeholder_box(ph, left_in, top_in, width_in, height_in):
    """Assign all 4 xfrm attributes together so inherited placeholder geometry never zeroes out."""
    ph.left = Emu(int(left_in * EMU_IN))
    ph.top = Emu(int(top_in * EMU_IN))
    ph.width = Emu(int(width_in * EMU_IN))
    ph.height = Emu(int(height_in * EMU_IN))


def normalize_blank_16x9_placeholders(slide, prs, is_title=False):
    """Expand default 4:3 (10in) blank-template placeholders cleanly onto 16:9 (13.333in)."""
    SW = prs.slide_width / EMU_IN
    SH = prs.slide_height / EMU_IN
    tph = placeholder_by_role(slide, "title")
    bphs = body_placeholders_sorted(slide)
    if is_title:
        if tph is not None:
            set_placeholder_box(tph, 0.8, SH * 0.28, SW - 1.6, 1.8)
            for p in tph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
        if bphs:
            set_placeholder_box(bphs[0], 1.2, SH * 0.56, SW - 2.4, 1.4)
            for p in bphs[0].text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
    else:
        if tph is not None:
            set_placeholder_box(tph, 0.55, 0.30, SW - 1.1, 1.05)
        if len(bphs) == 1:
            set_placeholder_box(bphs[0], 0.55, 1.52, SW - 1.1, SH - 2.15)
        elif len(bphs) >= 2:
            col_w = (SW - 1.45) / 2.0
            set_placeholder_box(bphs[0], 0.55, 1.52, col_w, SH - 2.15)
            set_placeholder_box(bphs[1], 0.55 + col_w + 0.35, 1.52, col_w, SH - 2.15)


def scale_4x3_template_placeholders_if_needed(slide, prs):
    """If a 16:9 template (SW>=12.5in) inherited raw 4:3 (10in) python-pptx master
    placeholders (all content placeholders end <=9.65in and no master/layout
    chrome exists in the right strip x>=9.65in), scale left/width cleanly to SW."""
    SW = prs.slide_width / EMU_IN
    if SW < 12.5:
        return False
    # Check if master or layout has intentional right-side non-placeholder shapes
    for container in (getattr(slide, "slide_layout", None),
                      getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
        if container is None:
            continue
        for shp in container.shapes:
            if getattr(shp, "is_placeholder", False):
                continue
            try:
                sx = shp.left / EMU_IN
                sw = shp.width / EMU_IN
                # Ignore full-width header/footer/background bars (sw >= SW*0.8)
                if sx >= 9.6 and sw < SW * 0.75:
                    return False
            except Exception:
                pass
    content_phs = [ph for ph in slide.placeholders
                   if ph.placeholder_format.type not in FOOTER_TYPES]
    if not content_phs:
        return False
    rights = []
    for ph in content_phs:
        try:
            rights.append((ph.left + ph.width) / EMU_IN)
        except Exception:
            pass
    if not rights or max(rights) > 9.65 or max(rights) < 7.0:
        return False
    sx = SW / 10.0
    for ph in content_phs:
        try:
            set_placeholder_box(ph, (ph.left / EMU_IN) * sx, ph.top / EMU_IN,
                                (ph.width / EMU_IN) * sx, ph.height / EMU_IN)
        except Exception:
            pass
    return True


def _rgb_luminance(rgb):
    try:
        r, g, b = rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    except Exception:
        return 1.0


def _solid_fill_rgb(fill):
    try:
        if fill is not None and fill.type == 1 and fill.fore_color is not None:
            rgb = fill.fore_color.rgb
            if rgb is not None:
                return (rgb[0], rgb[1], rgb[2])
    except Exception:
        pass
    return None


P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _theme_color_map(prs):
    """Read theme1.xml clrScheme into a dict (dk1/lt1/bg1/bg2/tx1/tx2/accent1..6 -> RGB tuple)
    strictly read-only without mutating any slide/layout/master XML."""
    cmap = {
        "dk1": (0, 0, 0), "tx1": (0, 0, 0),
        "lt1": (255, 255, 255), "bg1": (255, 255, 255),
        "dk2": (0x1B, 0x36, 0x5D), "tx2": (0x1B, 0x36, 0x5D),
        "lt2": (0xF2, 0xF4, 0xF8), "bg2": (0xF2, 0xF4, 0xF8),
    }
    try:
        import xml.etree.ElementTree as ET
        for master in prs.slide_masters:
            for rel in master.part.rels.values():
                if "theme" in (rel.reltype or "").lower():
                    root = ET.fromstring(rel.target_part.blob)
                    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                    scheme = root.find(".//a:clrScheme", ns)
                    if scheme is not None:
                        for child in scheme:
                            tag = child.tag.split("}", 1)[-1]
                            srgb = child.find("a:srgbClr", ns)
                            sys_c = child.find("a:sysClr", ns)
                            hx = (srgb.get("val") if srgb is not None else
                                  (sys_c.get("lastClr") if sys_c is not None else None))
                            if hx and len(hx) == 6:
                                cmap[tag] = (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
                    # Map PowerPoint bg1/bg2/tx1/tx2 aliases via master clrMap if present
                    clr_map = master._element.find(f"{P_NS}clrMap")
                    if clr_map is not None:
                        for k in ("bg1", "bg2", "tx1", "tx2"):
                            target = clr_map.get(k)
                            if target and target in cmap:
                                cmap[k] = cmap[target]
                    break
    except Exception:
        pass
    return cmap


def _xml_fill_rgb(container_el, cmap):
    """Extract RGB tuple from <a:solidFill> or <p:bgRef> strictly read-only."""
    if container_el is None:
        return None
    sf = container_el.find(f"{A_NS}solidFill")
    if sf is None and container_el.tag == f"{P_NS}bgRef":
        sf = container_el
    if sf is None:
        return None
    srgb = sf.find(f"{A_NS}srgbClr")
    if srgb is not None and srgb.get("val") and len(srgb.get("val")) == 6:
        hx = srgb.get("val")
        return (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
    sc = sf.find(f"{A_NS}schemeClr")
    if sc is not None and sc.get("val"):
        val = sc.get("val")
        if val in cmap:
            return cmap[val]
    return None


def is_dark_slide_bg(slide, prs):
    """Detect whether the slide/layout/master has a dark background fill using strictly
    read-only XML inspection (NEVER touch obj.background, which mutates XML with <a:noFill/>!)."""
    cmap = _theme_color_map(prs) if prs is not None else {}
    # 1) Check explicit <p:bg> from slide -> layout -> master (first explicit fill wins)
    for obj in (slide, getattr(slide, "slide_layout", None),
                getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
        if obj is None:
            continue
        try:
            bg = obj._element.find(f"{P_NS}cSld/{P_NS}bg")
            if bg is None:
                continue
            bg_pr = bg.find(f"{P_NS}bgPr")
            if bg_pr is not None:
                rgb = _xml_fill_rgb(bg_pr, cmap)
                if rgb is not None:
                    return _rgb_luminance(rgb) < 0.45
                if bg_pr.find(f"{A_NS}gradFill") is not None:
                    stops = bg_pr.findall(f".//{A_NS}gs")
                    rgbs = [_xml_fill_rgb(st, cmap) for st in stops]
                    rgbs = [c for c in rgbs if c is not None]
                    if rgbs:
                        avg_lum = sum(_rgb_luminance(c) for c in rgbs) / len(rgbs)
                        return avg_lum < 0.45
            bg_ref = bg.find(f"{P_NS}bgRef")
            if bg_ref is not None:
                rgb = _xml_fill_rgb(bg_ref, cmap)
                if rgb is not None:
                    return _rgb_luminance(rgb) < 0.45
        except Exception:
            pass
    # 2) Check large background cover shapes (>60% of slide area) on slide/layout/master (read-only XML)
    slide_area = (prs.slide_width / EMU_IN) * (prs.slide_height / EMU_IN)
    for container in (slide, getattr(slide, "slide_layout", None),
                      getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
        if container is None:
            continue
        for shp in container.shapes:
            if getattr(shp, "is_placeholder", False):
                continue
            try:
                area = (shp.width / EMU_IN) * (shp.height / EMU_IN)
                if area >= 0.60 * slide_area:
                    sp_pr = shp._element.find(f"{P_NS}spPr")
                    rgb = _xml_fill_rgb(sp_pr, cmap)
                    if rgb is not None:
                        return _rgb_luminance(rgb) < 0.45
            except Exception:
                pass
    return False


def capture_placeholder_color(slide, ph):
    """Capture explicit RGB font color from a placeholder (or its layout counterpart)
    using read-only XML so layout XML is never mutated."""
    def _xml_text_rgb(el):
        try:
            for tag in (f".//{A_NS}rPr/{A_NS}solidFill/{A_NS}srgbClr",
                        f".//{A_NS}defRPr/{A_NS}solidFill/{A_NS}srgbClr"):
                node = el.find(tag)
                if node is not None and node.get("val") and len(node.get("val")) == 6:
                    hx = node.get("val")
                    return RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
        except Exception:
            pass
        return None

    c = _xml_text_rgb(ph._element)
    if c is not None:
        return c
    try:
        idx = ph.placeholder_format.idx
        lay = getattr(slide, "slide_layout", None)
        if lay is not None:
            for lph in lay.placeholders:
                if lph.placeholder_format.idx == idx:
                    c = _xml_text_rgb(lph._element)
                    if c is not None:
                        return c
    except Exception:
        pass
    return None


def set_title(slide, text, prs=None, max_pt=30, is_blank=False, is_cover=False, warnings=None):
    ph = placeholder_by_role(slide, "title")
    if ph is None:
        if not text:
            return None
        # Fallback: never silently drop slide title when custom template lacks Title placeholder
        SW = prs.slide_width / EMU_IN if prs else 13.333
        ph = slide.shapes.add_textbox(Inches(0.55), Inches(0.35), Inches(SW - 1.10), Inches(0.95))
        ph.name = "FallbackTitle"
        if warnings is not None:
            warnings.append("layout lacks title placeholder; fallback title textbox used")
    # Avoid overlapping any top-right corner logo/badge on slide, layout, or master
    if not is_cover and prs is not None and ph is not None:
        try:
            SW = prs.slide_width / EMU_IN
            SH = prs.slide_height / EMU_IN
            tl, tt = ph.left / EMU_IN, ph.top / EMU_IN
            tw, th = ph.width / EMU_IN, ph.height / EMU_IN
            tr = tl + tw
            tb = tt + th
            right_cap = tr
            for container in (slide, getattr(slide, "slide_layout", None),
                              getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
                if container is None:
                    continue
                for shp in container.shapes:
                    if shp._element is ph._element or getattr(shp, "is_placeholder", False):
                        continue
                    if getattr(shp, "name", "") in ("TitleRule", "CoverAccent", "EvidenceLine", "SlideNumber"):
                        continue
                    try:
                        sl, st = shp.left / EMU_IN, shp.top / EMU_IN
                        sw, sh = shp.width / EMU_IN, shp.height / EMU_IN
                    except Exception:
                        continue
                    if sw * sh >= 0.35 * SW * SH:
                        continue
                    # Vertical overlap with title band & sits in right half
                    if sl >= SW * 0.60 and (st < tb - 0.05) and (st + sh > tt + 0.05):
                        right_cap = min(right_cap, sl - 0.18)
            if right_cap < tr - 0.05 and (right_cap - tl) >= 4.0:
                set_placeholder_box(ph, tl, tt, right_cap - tl, th)
        except Exception:
            pass

    orig_rgb = None if is_blank else capture_placeholder_color(slide, ph)
    tf = ph.text_frame
    tf.word_wrap = True
    w = ph.width / EMU_IN if ph.width else 11.5
    h = ph.height / EMU_IN if ph.height else 1.1
    eff_w = max(w - 0.25, 2.0)
    pt = None
    # Prefer clean single-line fit on content slides (0.84 factor guards LibreOffice CJK/Latin width)
    disp_text = glue_num_units((text or "").replace("**", ""))
    if not is_cover:
        for cand in (28, 26, 24, 22, 20, 19):
            if est_lines(disp_text, cand, eff_w * 0.84) == 1 and 1.25 * cand / 72.0 <= h * 0.96:
                pt = cand
                break
    if pt is None:
        pt = max_pt
        for cand in (32 if is_cover else 28, 26, 24, 22, 20, 18):
            pt = cand
            if est_lines(disp_text, cand, eff_w) * 1.25 * cand / 72.0 <= h * 0.96:
                break
    tf.clear()
    para = tf.paragraphs[0]
    if is_blank and not is_cover:
        para.alignment = PP_ALIGN.LEFT
    run = para.add_run()
    run.text = disp_text
    run.font.size = Pt(pt)
    _apply_blank_font(run, disp_text, is_blank)
    if is_blank:
        run.font.bold = True
        run.font.color.rgb = RGBColor(*BLANK_NAVY)
        if not is_cover:
            try:
                from pptx.enum.shapes import MSO_SHAPE
                line_top = (ph.top + ph.height) / EMU_IN + 0.03
                bar = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(line_top),
                    Inches(max(w, 4.0)), Inches(0.025))
                bar.name = "TitleRule"
                bar.fill.solid()
                bar.fill.fore_color.rgb = RGBColor(*BLANK_ACCENT)
                bar.line.fill.background()
            except Exception:
                pass
    elif orig_rgb is not None:
        run.font.color.rgb = orig_rgb
    elif prs is not None and is_dark_slide_bg(slide, prs):
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return ph


def add_formatted_runs(para, text, pt, color_rgb=None, force_bold=False, is_blank=False):
    """Render inline **bold** spans as native PowerPoint bold runs."""
    parts = BOLD_SPLIT_RE.split(text or "")
    any_run = False
    for part in parts:
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**") and len(part) > 4
        clean = part[2:-2] if is_bold else part
        if not clean:
            continue
        run = para.add_run()
        run.text = clean
        run.font.size = Pt(pt)
        _apply_blank_font(run, clean, is_blank)
        if is_bold or force_bold:
            run.font.bold = True
        if color_rgb is not None:
            run.font.color.rgb = color_rgb
        any_run = True
    if not any_run:
        run = para.add_run()
        run.text = (text or "").replace("**", "")
        run.font.size = Pt(pt)
        _apply_blank_font(run, run.text, is_blank)
        if force_bold:
            run.font.bold = True
        if color_rgb is not None:
            run.font.color.rgb = color_rgb


def disable_paragraph_bullet(para):
    """Add <a:buNone/> to suppress bullet markers (e.g. on cover subtitles)."""
    try:
        from pptx.oxml.ns import qn
        from pptx.oxml import parse_xml
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        pPr.append(parse_xml('<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
    except Exception:
        pass


def fill_body(ph, bullets, base_pt=18, slide=None, prs=None, is_blank=False, no_bullets=False):
    """bullets: list of (level, text). Step-shrink until fits placeholder height."""
    orig_rgb = None if (is_blank or slide is None) else capture_placeholder_color(slide, ph)
    if orig_rgb is None and not is_blank and slide is not None and prs is not None and is_dark_slide_bg(slide, prs):
        orig_rgb = RGBColor(0xEA, 0xEE, 0xF5)
    tf = ph.text_frame
    tf.word_wrap = True
    try:
        tf.auto_size = None
    except Exception:
        pass
    w = ph.width / EMU_IN if ph.width else 6.0
    h = ph.height / EMU_IN if ph.height else 4.5
    inner_w = max(w - (tf.margin_left + tf.margin_right) / EMU_IN, 1.5)
    inner_h = max(h - (tf.margin_top + tf.margin_bottom) / EMU_IN, 0.8)
    chosen = base_pt
    chosen_est_h = inner_h
    for pt in [p for p in FONT_SIZES if p <= base_pt] or [base_pt]:
        gap_in = (5.0 if pt >= 14 else (2.5 if pt >= 11 else 1.5)) / 72.0
        line_mult = 1.28 if pt >= 12 else 1.20
        est_h = 0.0
        for lv, t in bullets:
            lpt = max(pt - 2 * lv, 10.5)
            nl = est_lines(t, lpt, max(inner_w - 0.35 * lv, 1.2))
            est_h += nl * line_mult * lpt / 72.0 + gap_in
        chosen = pt
        chosen_est_h = est_h
        if est_h <= inner_h * 0.96:
            break

    # Tighten text-frame vertical padding when bullets are dense to prevent overflow
    if chosen <= 11.5 or chosen_est_h > inner_h * 0.92:
        try:
            tf.margin_top = Inches(0.02)
            tf.margin_bottom = Inches(0.02)
        except Exception:
            pass

    # Adaptive vertical rhythm: generous breathing room when bullets have vertical slack
    relaxed = (not no_bullets) and (chosen >= 14) and (chosen_est_h < inner_h * 0.65)
    tf.clear()
    first = True
    for lv, text in bullets:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.level = min(max(int(lv), 0), 4)
        if relaxed:
            para.space_after = Pt(8 if para.level == 0 else 4)
        elif chosen <= 11:
            para.space_after = Pt(1.5)
        else:
            para.space_after = Pt(5 if chosen >= 14 else 3)
        if no_bullets:
            disable_paragraph_bullet(para)
        add_formatted_runs(para, glue_num_units(text), max(chosen - 2 * lv, 10.5), color_rgb=orig_rgb, is_blank=is_blank)


def extract_theme_accent(prs):
    """Read template's theme1.xml accent1 RGB (fallback to BLANK_NAVY) so native tables
    harmonize with the template palette while maintaining high contrast."""
    try:
        from pptx.oxml.ns import qn
        import xml.etree.ElementTree as ET
        for master in prs.slide_masters:
            for rel in master.part.rels.values():
                if "theme" in (rel.reltype or "").lower():
                    root = ET.fromstring(rel.target_part.blob)
                    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                    for tag in ("a:accent1", "a:dk2"):
                        el = root.find(f".//a:clrScheme/{tag}/a:srgbClr", ns)
                        if el is not None and el.get("val") and len(el.get("val")) == 6:
                            hx = el.get("val")
                            rgb = (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
                            # Avoid near-white or near-black unreadable header fills
                            lum = _rgb_luminance(rgb)
                            if 0.04 <= lum <= 0.65:
                                return rgb
                            if lum > 0.65:
                                return (max(0, int(rgb[0] * 0.55)),
                                        max(0, int(rgb[1] * 0.55)),
                                        max(0, int(rgb[2] * 0.55)))
    except Exception:
        pass
    return BLANK_NAVY


def _tint_rgb(rgb, factor=0.88):
    """Blend rgb towards white by factor (0..1) for subtle Ours-row highlighting."""
    return (
        min(255, int(rgb[0] + (255 - rgb[0]) * factor)),
        min(255, int(rgb[1] + (255 - rgb[1]) * factor)),
        min(255, int(rgb[2] + (255 - rgb[2]) * factor)),
    )


def footer_layout_info(slide, prs):
    """Scan slide placeholders AND non-placeholder bottom bands on slide/layout/master.
    Returns (ev_top_in, ev_max_w_in, content_safe_bottom_in) so EvidenceLine never
    collides with template footers, page numbers, or bottom accent bars."""
    SW = prs.slide_width / EMU_IN
    SH = prs.slide_height / EMU_IN
    slide_area = SW * SH
    ev_top = SH - 0.40
    ev_left = 0.35
    ev_right_limit = SW - 0.45

    def _inspect_box(left, top, width, height, has_content=True):
        nonlocal ev_top, ev_right_limit
        if not has_content:
            return
        if width * height >= 0.55 * slide_area:
            return
        # Only care about bottom zone shapes
        if top < SH - 0.90 or height > 0.72:
            return
        # If shape sits strictly in right half, shrink EvidenceLine width instead of lifting it
        if left >= SW * 0.42:
            ev_right_limit = min(ev_right_limit, left - 0.20)
        else:
            ev_top = min(ev_top, top - 0.30)

    for ph in slide.placeholders:
        if ph.placeholder_format.type in FOOTER_TYPES:
            try:
                has_txt = ph.has_text_frame and bool(ph.text_frame.text.strip())
                _inspect_box(ph.left / EMU_IN, ph.top / EMU_IN,
                             ph.width / EMU_IN, ph.height / EMU_IN, has_txt)
            except Exception:
                pass

    for container in (slide, getattr(slide, "slide_layout", None),
                      getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
        if container is None:
            continue
        for shp in container.shapes:
            if getattr(shp, "is_placeholder", False):
                continue
            if getattr(shp, "name", "") in ("EvidenceLine", "ImageCaption", "TableNote",
                                            "TitleRule", "FallbackTitle", "SlideNumber", "CoverAccent") or \
                    str(getattr(shp, "name", "")).startswith(("KPICard", "KPIBar")):
                continue
            try:
                l_in, t_in = shp.left / EMU_IN, shp.top / EMU_IN
                w_in, h_in = shp.width / EMU_IN, shp.height / EMU_IN
                # Count visible footer textboxes or decorative bottom bars
                has_vis = bool(getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip())
                if not has_vis and w_in >= SW * 0.25 and h_in >= 0.04 and t_in >= SH - 0.75:
                    has_vis = True
                _inspect_box(l_in, t_in, w_in, h_in, has_vis)
            except Exception:
                pass

    ev_top = max(ev_top, SH - 1.05)
    ev_max_w = max(ev_right_limit - ev_left, 3.2)
    # Content (body/media/caption/note) must stay cleanly above EvidenceLine
    content_safe_bottom = ev_top - 0.12
    return ev_top, ev_max_w, content_safe_bottom


def footer_safe_top(slide, prs):
    """Backward-compatible helper returning Y top of EvidenceLine."""
    ev_top, _w, _cb = footer_layout_info(slide, prs)
    return ev_top


def add_evidence_line(slide, prs, text, reserve_right_in=0.0):
    if not text:
        return
    ev_top, ev_max_w, _cb = footer_layout_info(slide, prs)
    if reserve_right_in > 0:
        SW = prs.slide_width / EMU_IN
        ev_max_w = min(ev_max_w, max(3.2, SW - 0.35 - reserve_right_in))
    # Trim gracefully to fit available horizontal width without clipping/overlapping right footer
    trimmed = text[:160].strip()
    if est_text_width(trimmed, 10) > ev_max_w - 0.08:
        while len(trimmed) > 12 and est_text_width(trimmed + "…", 10) > ev_max_w - 0.08:
            trimmed = trimmed[:-1]
        trimmed = trimmed.rstrip(" ;,，") + "…"
    box = slide.shapes.add_textbox(Inches(0.35), Inches(ev_top),
                                   Inches(ev_max_w), Inches(0.28))
    box.name = "EvidenceLine"
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    run = tf.paragraphs[0].add_run()
    run.text = trimmed
    run.font.size = Pt(10)
    fg = LIGHT_GRAY if is_dark_slide_bg(slide, prs) else GRAY
    run.font.color.rgb = RGBColor(*fg)


def add_blank_slide_number(slide, prs, idx, total):
    """In template-free blank mode, add a clean bottom-right slide counter (i / N)."""
    SW = prs.slide_width / EMU_IN
    ev_top, _w, _cb = footer_layout_info(slide, prs)
    box = slide.shapes.add_textbox(Inches(SW - 1.15), Inches(ev_top),
                                   Inches(0.80), Inches(0.28))
    box.name = "SlideNumber"
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = f"{idx} / {total}"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(*GRAY)


def synthesize_slide_line(paper_refs):
    """Auto-synthesize a compact evidence footer line from paper_refs when slide_line is omitted."""
    if not paper_refs:
        return ""
    parts = []
    for r in paper_refs[:3]:
        rid = str(r.get("id") or "").strip()
        pg = r.get("page")
        if rid and pg:
            parts.append(f"{rid} (p.{pg})")
        elif rid:
            parts.append(rid)
        elif pg:
            parts.append(f"p.{pg}")
    return ("证据: " + "; ".join(parts)) if parts else ""


def _style_cell_booktabs(cell, bot_hex=None, bot_pt=0.75, top_hex=None, top_pt=0.75):
    """Apply clean publication-grade booktabs borders (no vertical gridlines, crisp horizontal rules)."""
    try:
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import qn
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            old = tcPr.find(qn(tag))
            if old is not None:
                tcPr.remove(old)
        ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        tcPr.append(parse_xml(f'<a:lnL {ns} w="0"><a:noFill/></a:lnL>'))
        tcPr.append(parse_xml(f'<a:lnR {ns} w="0"><a:noFill/></a:lnR>'))
        if top_hex:
            w_emu = int(top_pt * 12700)
            tcPr.append(parse_xml(f'<a:lnT {ns} w="{w_emu}"><a:solidFill><a:srgbClr val="{top_hex}"/></a:solidFill></a:lnT>'))
        else:
            tcPr.append(parse_xml(f'<a:lnT {ns} w="0"><a:noFill/></a:lnT>'))
        if bot_hex:
            w_emu = int(bot_pt * 12700)
            tcPr.append(parse_xml(f'<a:lnB {ns} w="{w_emu}"><a:solidFill><a:srgbClr val="{bot_hex}"/></a:solidFill></a:lnB>'))
        else:
            tcPr.append(parse_xml(f'<a:lnB {ns} w="0"><a:noFill/></a:lnB>'))
    except Exception:
        pass


def add_kpi_strip(slide, prs, kpis, left_in, top_in, width_in, theme_accent=None):
    """Render 2-4 crisp headline metric cards (value + short label) right below title.
    Returns bottom Y (in inches) of the strip."""
    items = [k for k in (kpis or []) if isinstance(k, dict) and str(k.get("value", "")).strip()][:4]
    if not items:
        return top_in
    from pptx.enum.shapes import MSO_SHAPE
    dark_bg = is_dark_slide_bg(slide, prs)
    accent = theme_accent or BLANK_NAVY
    card_bg = (0x1E, 0x26, 0x36) if dark_bg else _tint_rgb(accent, 0.93)
    card_border = (0x33, 0x41, 0x55) if dark_bg else _tint_rgb(accent, 0.78)
    val_fg = (0x93, 0xC5, 0xFD) if dark_bg else accent
    lbl_fg = LIGHT_GRAY if dark_bg else GRAY
    bar_rgb = (min(255, accent[0] + 60), min(255, accent[1] + 70), min(255, accent[2] + 85)) if dark_bg else accent

    n = len(items)
    gap = 0.20
    cw = (width_in - gap * (n - 1)) / float(n)
    ch = 0.88
    inner_w = max(cw - 0.24, 0.8)
    for idx, item in enumerate(items):
        cx = left_in + idx * (cw + gap)
        val_txt = glue_num_units(str(item.get("value", "")).replace("**", "").strip())[:24]
        lbl_txt = glue_num_units(str(item.get("label", "")).replace("**", "").strip())[:48]
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(top_in), Inches(cw), Inches(ch))
        card.name = f"KPICard_{idx + 1}"
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(*card_bg)
        card.line.color.rgb = RGBColor(*card_border)
        card.line.width = Pt(0.75)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.10)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.04)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        p0.space_after = Pt(1.5)
        v_pt = 14
        for cand_v in (20, 18, 16, 14):
            if est_lines(val_txt, cand_v, inner_w * 0.86) == 1:
                v_pt = cand_v
                break
        r0 = p0.add_run()
        r0.text = val_txt
        r0.font.size = Pt(v_pt)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(*val_fg)
        if lbl_txt:
            l_pt = 9.5
            v_h = 1.25 * v_pt / 72.0
            for cand_l in (10.5, 10, 9.5):
                nl_l = est_lines(lbl_txt, cand_l, inner_w * 0.88)
                if v_h + nl_l * 1.25 * cand_l / 72.0 <= 0.70:
                    l_pt = cand_l
                    break
            p1 = tf.add_paragraph()
            p1.alignment = PP_ALIGN.LEFT
            r1 = p1.add_run()
            r1.text = lbl_txt
            r1.font.size = Pt(l_pt)
            r1.font.color.rgb = RGBColor(*lbl_fg)
        # Sleek thin accent bar on top edge of each card
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(top_in), Inches(cw), Inches(0.045))
        bar.name = f"KPIBar_{idx + 1}"
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(*bar_rgb)
        bar.line.fill.background()
    return top_in + ch


def add_takeaway_bar(slide, prs, text, left_in, top_in, width_in, theme_accent=None, is_blank=False):
    """Render a sleek single-line left-accented Key Takeaway callout bar cleanly above the footer."""
    clean_t = glue_num_units((text or "").strip())[:160]
    if not clean_t:
        return top_in
    from pptx.enum.shapes import MSO_SHAPE
    dark_bg = is_dark_slide_bg(slide, prs)
    accent = theme_accent or BLANK_NAVY
    bg_rgb = (0x1E, 0x26, 0x36) if dark_bg else _tint_rgb(accent, 0.92)
    bdr_rgb = (0x33, 0x41, 0x55) if dark_bg else _tint_rgb(accent, 0.76)
    bar_rgb = (min(255, accent[0] + 60), min(255, accent[1] + 70), min(255, accent[2] + 85)) if dark_bg else accent
    txt_rgb = RGBColor(0xEA, 0xEE, 0xF5) if dark_bg else RGBColor(0x1E, 0x29, 0x3B)
    has_cjk = any(ord(ch) > 0x2E7F for ch in clean_t)
    prefix = "要点：" if has_cjk else "Takeaway: "
    full_t = clean_t if re.match(r"^(?:takeaway|key\s+takeaway|要点|核心结论)\s*[:：]", clean_t, re.I) else (prefix + clean_t)

    bar_h = 0.42
    inner_w = max(width_in - 0.30, 3.0)
    pt = 10.0
    for cand in (11.5, 11.0, 10.5, 10.0):
        if est_lines(full_t, cand, inner_w * 0.92) == 1:
            pt = cand
            break

    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left_in), Inches(top_in), Inches(width_in), Inches(bar_h))
    box.name = "TakeawayBar"
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(*bg_rgb)
    box.line.color.rgb = RGBColor(*bdr_rgb)
    box.line.width = Pt(0.65)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.16)
    tf.margin_right = Inches(0.10)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if ":" in full_t[:16] or "：" in full_t[:8]:
        sep = "：" if "：" in full_t[:8] else ":"
        head, rest = full_t.split(sep, 1)
        r_h = p.add_run()
        r_h.text = head + sep
        r_h.font.size = Pt(pt)
        r_h.font.bold = True
        r_h.font.color.rgb = RGBColor(*bar_rgb)
        _apply_blank_font(r_h, r_h.text, is_blank)
        add_formatted_runs(p, rest, pt, color_rgb=txt_rgb, is_blank=is_blank)
    else:
        add_formatted_runs(p, full_t, pt, color_rgb=txt_rgb, is_blank=is_blank)

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left_in), Inches(top_in), Inches(0.06), Inches(bar_h))
    stripe.name = "TakeawayAccent"
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = RGBColor(*bar_rgb)
    stripe.line.fill.background()
    return top_in + bar_h


def add_table(slide, prs, table_spec, left_in, top_in, max_w_in, max_h_in,
              base_pt=12, is_blank=False, theme_accent=None, safe_bottom_override=None):
    """Render an editable native pptx table (verbatim rows, hard-capped at 6 data rows),
    with CJK-aware column sizing and template-harmonized high-contrast styling."""
    headers = [str(x).strip() for x in table_spec.get("headers", [])]
    orig_rows = table_spec.get("rows", []) or []
    rows = [[str(c).strip() for c in r] for r in orig_rows][:6]
    n_more = len(orig_rows) - len(rows)
    note = (table_spec.get("note") or "").strip()
    if n_more > 0:
        rows.append(["…"] * max(len(headers), 1))
        note = (note + f" (另有 {n_more} 行见论文原文)").strip()
    n_cols = max(len(headers), max((len(r) for r in rows), default=1), 1)
    n_rows = len(rows) + (1 if headers else 0)
    if n_rows == 0:
        return None

    # CJK-aware column width weighting using est_text_width (strip inline **bold** markers)
    def cweight(ci):
        vals = ([headers[ci] if ci < len(headers) else ""] +
                [r[ci] if ci < len(r) else "" for r in rows])
        ws = [est_text_width(v.replace("**", ""), 12) for v in vals if v]
        return max(0.55, min(max(ws or [0.6]), 2.8))

    tot = sum(cweight(ci) for ci in range(n_cols))
    widths = [max(0.60, max_w_in * cweight(ci) / tot) for ci in range(n_cols)]
    scale = max_w_in / sum(widths) if sum(widths) > max_w_in else 1.0
    widths = [w * scale for w in widths]

    # Choose cell font size & exact per-row heights so LibreOffice expansion never overlaps TableNote
    _ev_top, _ev_w, safe_bot = footer_layout_info(slide, prs)
    if safe_bottom_override is not None:
        safe_bot = min(safe_bot, safe_bottom_override)
    max_h_clamped = min(max_h_in, max(safe_bot - top_in, 1.4))
    note_disp = glue_num_units(note[:150]) if note else ""
    note_lines = est_lines(note_disp, 10, max(max_w_in - 0.04, 1.5)) if note_disp else 0
    note_pt = 9.5 if note_lines >= 2 else 10
    note_h = (0.40 if note_lines >= 2 else 0.26) if note_disp else 0.0
    avail_h = max(max_h_clamped - (note_h + 0.08 if note_disp else 0.04), 1.1)
    cell_pt = base_pt
    row_heights = [0.34] * n_rows
    all_r = ([headers] if headers else []) + rows
    for cand_pt in (base_pt, 11, 10):
        cell_pt = cand_pt
        rhs = []
        for r in all_r:
            rh = 0.30
            for ci in range(min(len(r), n_cols)):
                nl = est_lines(r[ci], cand_pt, max(widths[ci] - 0.14, 0.42))
                rh = max(rh, nl * 1.30 * cand_pt / 72.0 + 0.12)
            rhs.append(rh)
        row_heights = rhs
        if sum(rhs) <= avail_h:
            break

    tot_rh = sum(row_heights)
    if tot_rh > avail_h and tot_rh > 0:
        shrink = avail_h / tot_rh
        row_heights = [max(0.26, rh * shrink) for rh in row_heights]
    tbl_h = min(avail_h, sum(row_heights))
    gfx = slide.shapes.add_table(n_rows, n_cols, Inches(left_in), Inches(top_in),
                                 Inches(sum(widths)), Inches(tbl_h))
    tbl = gfx.table
    for ci, w in enumerate(widths):
        tbl.columns[ci].width = Inches(w)
    for r_idx, rh in enumerate(row_heights):
        tbl.rows[r_idx].height = Inches(rh)

    dark_bg = is_dark_slide_bg(slide, prs)
    accent = theme_accent or BLANK_NAVY
    if dark_bg:
        hdr_bg = (min(255, int(accent[0] * 0.85) + 35),
                  min(255, int(accent[1] * 0.85) + 45),
                  min(255, int(accent[2] * 0.85) + 60))
        row_even_bg = (0x1E, 0x26, 0x36)
        row_odd_bg = (0x27, 0x31, 0x44)
        row_ours_bg = (0x2D, 0x45, 0x63)
        body_fg = (0xF2, 0xF5, 0xFA)
    else:
        hdr_bg = accent
        row_even_bg = (0xFF, 0xFF, 0xFF)
        row_odd_bg = (0xF5, 0xF7, 0xFA)
        row_ours_bg = _tint_rgb(accent, 0.86)
        body_fg = (0x1A, 0x20, 0x2C)

    ri = 0
    num_re = re.compile(
        r"^[+-]?\d[\d.,%\s±x×/–-]*(?:meV/A|meV/Å|kcal/mol|GFLOPs|TFLOPs|tok/s|fps|MHz|GHz|Hz|GB|MB|KB|ms|ns|μs|eV|mW|nm|mm|cm|km|dB|[sMWK])?$",
        re.I,
    )
    na_vals = {"—", "–", "-", "--", "n/a", "na", "…", "..."}
    col_numeric = []
    for ci in range(n_cols):
        vals = [r[ci].replace("**", "").strip() for r in rows if ci < len(r) and r[ci].replace("**", "").strip()]
        non_na = [v for v in vals if v.lower() not in na_vals]
        col_numeric.append(bool(non_na) and all(num_re.match(v) for v in non_na))

    hdr_rgb = RGBColor(0xFF, 0xFF, 0xFF)
    body_rgb = RGBColor(*body_fg)
    rule_hex = "3B475E" if dark_bg else "D8DEE9"
    end_hex = "64748B" if dark_bg else "{:02X}{:02X}{:02X}".format(*accent)

    if headers:
        for ci, txt in enumerate(headers[:n_cols]):
            cell = tbl.cell(0, ci)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(*hdr_bg)
            _style_cell_booktabs(cell, top_hex=end_hex, top_pt=1.0, bot_hex=end_hex, bot_pt=1.25)
            ctf = cell.text_frame
            ctf.word_wrap = True
            ctf.clear()
            p = ctf.paragraphs[0]
            if col_numeric[ci]:
                p.alignment = PP_ALIGN.CENTER
            add_formatted_runs(p, glue_num_units(txt), cell_pt, color_rgb=hdr_rgb, force_bold=True)
        ri = 1
    ours_re = re.compile(r"(?:\bours\b|\bproposed\b|\bthis\s+work\b|本文|本方法)", re.I)
    for r_i, row in enumerate(rows):
        first_clean = str(row[0]).replace("**", "") if row else ""
        is_ours = bool(first_clean and ours_re.search(first_clean))
        is_last = (r_i == len(rows) - 1)
        for ci in range(n_cols):
            cell = tbl.cell(ri + r_i, ci)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.fill.solid()
            if is_ours:
                cell.fill.fore_color.rgb = RGBColor(*row_ours_bg)
            elif r_i % 2 == 1:
                cell.fill.fore_color.rgb = RGBColor(*row_odd_bg)
            else:
                cell.fill.fore_color.rgb = RGBColor(*row_even_bg)
            _style_cell_booktabs(
                cell,
                bot_hex=(end_hex if is_last else rule_hex),
                bot_pt=(1.25 if is_last else 0.5),
            )
            ctf = cell.text_frame
            ctf.word_wrap = True
            ctf.clear()
            p = ctf.paragraphs[0]
            if col_numeric[ci]:
                p.alignment = PP_ALIGN.CENTER
            cell_txt = row[ci] if ci < len(row) else ""
            add_formatted_runs(p, glue_num_units(cell_txt), cell_pt, color_rgb=body_rgb, force_bold=is_ours)

    if note_disp:
        note_top = min(top_in + tbl_h + 0.05, safe_bot - note_h)
        box = slide.shapes.add_textbox(Inches(left_in), Inches(note_top),
                                       Inches(max_w_in), Inches(note_h))
        box.name = "TableNote"
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        run = tf.paragraphs[0].add_run()
        run.text = note_disp
        run.font.size = Pt(note_pt)
        fg = LIGHT_GRAY if dark_bg else GRAY
        run.font.color.rgb = RGBColor(*fg)
    return gfx


def add_picture_fit(slide, img_path, left_in, top_in, max_w_in, max_h_in,
                    v_align="top_bias", prs=None, add_frame=True):
    """Insert picture aspect-fitted inside target box without cropping or distortion,
    with a crisp subtle hairline frame so white-background PDF plots stand out cleanly."""
    try:
        from PIL import Image
        iw, ih = Image.open(img_path).size
    except Exception:
        iw, ih = 800, 600
    # Prevent excessive upscaling of tiny rasters (>2.2x at 96dpi)
    max_nat_w = (iw / 96.0) * 2.2
    max_nat_h = (ih / 96.0) * 2.2
    eff_w = min(max_w_in, max(max_nat_w, min(max_w_in, 2.2)))
    eff_h = min(max_h_in, max(max_nat_h, min(max_h_in, 1.8)))
    scale = min(eff_w / max(iw, 1), eff_h / max(ih, 1))
    w, h = iw * scale, ih * scale
    left = left_in + (max_w_in - w) / 2.0
    if v_align == "top_bias":
        top = top_in + min((max_h_in - h) * 0.25, 0.38)
    else:
        top = top_in + (max_h_in - h) / 2.0
    pic = slide.shapes.add_picture(img_path, Inches(left), Inches(top),
                                   Inches(w), Inches(h))
    if add_frame:
        try:
            dark_bg = is_dark_slide_bg(slide, prs) if prs is not None else False
            pic.line.color.rgb = RGBColor(*( (0x33, 0x41, 0x55) if dark_bg else (0xCB, 0xD5, 0xE1) ))
            pic.line.width = Pt(0.75)
        except Exception:
            pass
    return pic


def stack_images(slide, imgs, left_in, top_in, max_w_in, max_h_in, prs=None):
    """Place 1..3 figures vertically stacked (or 2 side-by-side if wide box).
    Returns bottom Y (in inches) of the lowest placed picture."""
    imgs = imgs[:3]
    n = len(imgs)
    if n == 0:
        return top_in
    gap = 0.14
    bottoms = []
    if n == 1:
        p_shp = add_picture_fit(slide, imgs[0], left_in, top_in, max_w_in, max_h_in, v_align="top_bias", prs=prs)
        bottoms.append((p_shp.top + p_shp.height) / EMU_IN)
    elif n == 2 and (max_w_in >= 8.5 or (max_w_in >= 5.6 and max_w_in / max(max_h_in, 1.0) >= 1.15)):
        # Side-by-side layout when width allows two crisp figures
        avg_ar = 1.4
        try:
            from PIL import Image
            ars = []
            for p in imgs:
                iw, ih = Image.open(p).size
                ars.append(iw / max(ih, 1))
            avg_ar = sum(ars) / len(ars)
        except Exception:
            pass
        if max_w_in >= 8.5 or avg_ar <= 1.75:
            w_each = (max_w_in - gap) / 2.0
            for k, p in enumerate(imgs):
                p_shp = add_picture_fit(slide, p, left_in + k * (w_each + gap), top_in, w_each, max_h_in, v_align="top_bias", prs=prs)
                bottoms.append((p_shp.top + p_shp.height) / EMU_IN)
        else:
            h_each = (max_h_in - gap) / 2.0
            for k, p in enumerate(imgs):
                p_shp = add_picture_fit(slide, p, left_in, top_in + k * (h_each + gap), max_w_in, h_each, v_align="center", prs=prs)
                bottoms.append((p_shp.top + p_shp.height) / EMU_IN)
    else:
        h_each = (max_h_in - gap * (n - 1)) / n
        for k, p in enumerate(imgs):
            p_shp = add_picture_fit(slide, p, left_in, top_in + k * (h_each + gap), max_w_in, h_each, v_align="center", prs=prs)
            bottoms.append((p_shp.top + p_shp.height) / EMU_IN)
    return max(bottoms) if bottoms else top_in


def _placeholder_has_media(shp):
    """Return True if a placeholder actually contains a picture, table, chart, or blip."""
    try:
        if shp.shape_type == 13 or getattr(shp, "has_table", False) or getattr(shp, "has_chart", False):
            return True
        if shp._element.find(f".//{A_NS}blip") is not None:
            return True
    except Exception:
        pass
    return False


def prune_empty_placeholders(slide, filled_els, keep_existing_media=False):
    """Remove unfilled non-footer ghost placeholders (BODY/OBJECT/PICTURE/SUBTITLE)
    so neither newly added multi-placeholder layouts nor --reuse slides leave empty
    <p:ph> boxes behind."""
    removed = 0
    for shp in list(slide.shapes):
        if not getattr(shp, "is_placeholder", False):
            continue
        if shp._element in filled_els:
            continue
        try:
            pt = shp.placeholder_format.type
        except Exception:
            continue
        if pt in FOOTER_TYPES or pt not in GHOST_CLEANABLE_TYPES:
            continue
        if keep_existing_media and _placeholder_has_media(shp):
            continue
        has_text = bool(shp.has_text_frame and shp.text_frame.text.strip())
        if not has_text and not _placeholder_has_media(shp):
            remove_shape(shp)
            removed += 1
    return removed


def clean_reused_slide(slide, keep_shapes, prs=None, keep_existing_media=False):
    """In --reuse mode, clear stale pictures/tables/charts/groups/extra textboxes while
    protecting kept placeholders (matched by XML _element identity), template header/footer
    ribbons, small corner logos, and full-bleed background panels. Honors
    keep_existing_media=True to preserve native charts/diagrams/pictures from the starting deck."""
    keep_els = {s._element for s in keep_shapes if s is not None}
    SW = prs.slide_width / EMU_IN if prs else 13.333
    SH = prs.slide_height / EMU_IN if prs else 7.5
    slide_area = SW * SH
    for shp in list(slide.shapes):
        if shp._element in keep_els:
            continue
        if shp.is_placeholder:
            pt = shp.placeholder_format.type
            if pt in FOOTER_TYPES:
                continue
            if _placeholder_has_media(shp):
                if not keep_existing_media:
                    remove_shape(shp)
            elif pt in GHOST_CLEANABLE_TYPES:
                remove_shape(shp)
            elif shp.has_text_frame:
                shp.text_frame.clear()
        else:
            try:
                w_in = shp.width / EMU_IN
                h_in = shp.height / EMU_IN
                t_in = shp.top / EMU_IN
            except Exception:
                w_in = h_in = t_in = 0.0
            is_bg_panel = (w_in * h_in >= 0.65 * slide_area)
            if is_bg_panel:
                continue
            is_corner_logo = (shp.shape_type == 13 and w_in <= 2.5 and h_in <= 1.3 and
                              (t_in <= 0.85 or t_in + h_in >= SH - 0.85))
            if is_corner_logo:
                continue
            is_media_shp = (shp.shape_type == 13 or getattr(shp, "has_table", False) or
                            getattr(shp, "has_chart", False) or
                            (shp.shape_type == 6 and w_in >= 1.5 and h_in >= 1.0))
            if is_media_shp:
                if not keep_existing_media:
                    remove_shape(shp)
            elif shp.has_text_frame and shp.text_frame.text.strip():
                # Protect template header/footer bands
                if (t_in >= SH - 0.65 and h_in <= 0.60) or (t_in + h_in <= 0.55 and h_in <= 0.45):
                    continue
                if w_in >= 2.0 and h_in >= 0.25:
                    remove_shape(shp)


def should_use_wide_stacked_layout(imgs, table_spec, bullets, is_cover, avail_h):
    """Switch to Top-Bullets + Full-Width Media for wide tables (>=5 cols, or 4 wide cols)
    or wide banner architecture diagrams (aspect >= 2.2) when bullets are concise."""
    if is_cover or not bullets or avail_h < 3.6:
        return False
    if table_spec:
        headers = table_spec.get("headers") or []
        rows = table_spec.get("rows") or []
        n_cols = max(len(headers), max((len(r) for r in rows), default=0))
        if n_cols >= 5 and len(bullets) <= 4:
            return True
        if n_cols == 4 and len(bullets) <= 3 and avail_h >= 3.8:
            hdr_w = sum(est_text_width(str(h).replace("**", ""), 12) for h in headers)
            if hdr_w >= 4.2:
                return True
    if len(bullets) <= 3 and len(imgs) == 1 and os.path.isfile(imgs[0]):
        try:
            from PIL import Image
            iw, ih = Image.open(imgs[0]).size
            if ih > 0 and (iw / float(ih)) >= 2.2:
                return True
        except Exception:
            pass
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("slides_json")
    ap.add_argument("--template", default=None)
    ap.add_argument("--out", default="/app/output/slides.pptx")
    ap.add_argument("--base-pt", type=int, default=18)
    ap.add_argument("--evidence-line", action="store_true", default=True)
    ap.add_argument("--no-evidence-line", dest="evidence_line", action="store_false")
    ap.add_argument("--reuse", action="store_true",
                    help="edit an existing starting deck in-place instead of rebuilding from layouts")
    a = ap.parse_args()

    with open(a.slides_json, "r", encoding="utf-8") as f:
        spec = json.load(f)
    slides = spec.get("slides", [])
    meta = spec.get("meta", {})
    problems, warnings = [], []
    max_slides = meta.get("max_slides")
    if max_slides and len(slides) > int(max_slides):
        warnings.append(f"plan has {len(slides)} slides > max_slides={max_slides}; "
                        "trim before delivery (cover/results/conclusion are mandatory)")

    prs = None
    is_blank = False
    cleared_template_slides = 0
    if a.template and os.path.isfile(a.template):
        try:
            prs = Presentation(a.template)
        except Exception as e:
            warnings.append(f"template load failed ({e}); using blank 16:9")
    if prs is None:
        is_blank = True
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    # Non-reuse with template: drop any existing sample slides while keeping masters/layouts!
    if a.template and not is_blank and not a.reuse and len(prs.slides) > 0:
        cleared_template_slides = drop_all_slides(prs)

    SW_in = prs.slide_width / EMU_IN
    SH_in = prs.slide_height / EMU_IN
    theme_accent = BLANK_NAVY if is_blank else extract_theme_accent(prs)
    font_audit = init_font_strategy(prs, is_blank=is_blank)

    def new_slide(sld):
        has_media = bool(sld.get("image") or sld.get("images") or sld.get("table"))
        has_bullets = bool(sld.get("bullets"))
        has_two_col = bool((has_media and has_bullets) or (sld.get("right_bullets") and not has_media))
        is_title = sld.get("kind", "") in ("title", "cover") or sld is slides[0]
        li = pick_layout(prs, need_two_col=has_two_col,
                         need_picture=bool(sld.get("image") or sld.get("images")),
                         is_title=is_title)
        slide = prs.slides.add_slide(prs.slide_layouts[li])
        if is_blank:
            normalize_blank_16x9_placeholders(slide, prs, is_title=is_title)
        else:
            scale_4x3_template_placeholders_if_needed(slide, prs)
        return slide, li

    # Reuse mode: support explicit source_slide selection/reordering or positional trimming.
    # CRITICAL: create any newly added slides BEFORE dropping unselected starting slides so
    # python-pptx's `slide{len+1}.xml` partname allocator can never collide with kept slides!
    reuse_plan = None
    reuse_mapping = []
    if a.reuse and a.template and not is_blank:
        xml_slides = prs.slides._sldIdLst
        orig_ids = list(xml_slides)
        n_orig = len(orig_ids)
        has_explicit_src = any(
            isinstance(s.get("source_slide"), int) and 1 <= s["source_slide"] <= n_orig
            for s in slides
        )
        if has_explicit_src:
            used_src = set()
            target_slots = []
            for si_p, s in enumerate(slides):
                ss = s.get("source_slide")
                if isinstance(ss, int) and 1 <= ss <= n_orig and ss not in used_src:
                    used_src.add(ss)
                    target_slots.append((True, ss, orig_ids[ss - 1]))
                else:
                    _ns, _li = new_slide(s)
                    target_slots.append((False, None, xml_slides[-1]))
            # Now safely drop unselected original slides
            for idx1, sldId in enumerate(orig_ids, 1):
                if idx1 not in used_src:
                    rId = sldId.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    try:
                        prs.part.drop_rel(rId)
                    except Exception:
                        pass
                    xml_slides.remove(sldId)
            # Reorder all kept + newly created slides into exact target order
            for _reused, _ss, el_k in target_slots:
                xml_slides.remove(el_k)
                xml_slides.append(el_k)
            reuse_plan = target_slots
        else:
            n_keep = len(slides)
            for sldId in orig_ids[n_keep:]:
                rId = sldId.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                try:
                    prs.part.drop_rel(rId)
                except Exception:
                    pass
                xml_slides.remove(sldId)
            if len(list(xml_slides)) < n_keep:
                warnings.append("starting deck has fewer slides than plan; extras appended")

    used_layouts = []
    for si, sld in enumerate(slides):
        if reuse_plan is not None:
            reused_this, src_no, _el = reuse_plan[si]
            slide = prs.slides[si]
            li = getattr(slide.slide_layout, "name", "?")
            reuse_mapping.append({
                "slide": si + 1,
                "source_slide": src_no,
                "action": ("reused" if reused_this else "added"),
            })
        else:
            reused_this = bool(a.reuse and not is_blank and si < len(prs.slides))
            if reused_this:
                slide = prs.slides[si]
                li = getattr(slide.slide_layout, "name", "?")
                reuse_mapping.append({"slide": si + 1, "source_slide": si + 1, "action": "reused"})
            else:
                slide, li = new_slide(sld)
                if a.reuse:
                    reuse_mapping.append({"slide": si + 1, "source_slide": None, "action": "added"})
        used_layouts.append(getattr(slide.slide_layout, "name", str(li)))

        is_cover = sld.get("kind", "") in ("title", "cover") or si == 0
        title = sld.get("title", "")
        bullets = [(b.get("level", 0) if isinstance(b, dict) else 0,
                    b.get("text", "") if isinstance(b, dict) else str(b))
                   for b in sld.get("bullets", [])]
        right_bullets = [(b.get("level", 0) if isinstance(b, dict) else 0,
                          b.get("text", "") if isinstance(b, dict) else str(b))
                         for b in sld.get("right_bullets", [])]
        left_heading = (sld.get("left_heading") or "").strip()
        right_heading = (sld.get("right_heading") or "").strip()
        if is_cover and not bullets:
            seen_cov = []
            for ck in ("subtitle", "authors", "venue", "date"):
                cv = (sld.get(ck) or (meta.get(ck) if ck == "authors" else "") or "").strip()
                if cv and not any(cv.lower() in prev.lower() for prev in seen_cov):
                    bullets.append((0, cv))
                    seen_cov.append(cv)
        notes = sld.get("notes", "")
        ev = sld.get("evidence") or {}
        ev_text = (ev.get("slide_line") or "").strip()
        if not ev_text and not is_cover and ev.get("paper_refs"):
            ev_text = synthesize_slide_line(ev.get("paper_refs"))

        keep_existing_media = bool(sld.get("keep_existing_media") and reused_this)
        img = sld.get("image")
        if img and not os.path.isfile(img):
            warnings.append(f"slide {si+1}: image not found: {img}")
            img = None
        imgs = sld.get("images") or ([img] if img else [])
        imgs = [p for p in imgs if os.path.isfile(p)]
        if sld.get("images") and len(imgs) < len(sld.get("images")):
            warnings.append(f"slide {si+1}: some images missing ({len(imgs)}/{len(sld['images'])})")
        table_spec = sld.get("table")
        has_media = bool(imgs or table_spec)

        tph = placeholder_by_role(slide, "title")
        bodies = body_placeholders_sorted(slide, is_cover=is_cover)
        body = bodies[0] if bodies else None
        right_ph = bodies[1] if len(bodies) >= 2 else None
        pic_ph = placeholder_by_role(slide, "pic")
        init_content_phs = [p for p in (body, right_ph, pic_ph) if p is not None]
        filled_els = set()

        if reused_this:
            clean_reused_slide(slide, keep_shapes=[tph, body, right_ph, pic_ph],
                               prs=prs, keep_existing_media=keep_existing_media)
            scale_4x3_template_placeholders_if_needed(slide, prs)
            if keep_existing_media and not has_media and body is not None:
                # Prevent body text from overlapping any preserved right-side chart/picture/group
                try:
                    bx, by = body.left / EMU_IN, body.top / EMU_IN
                    bw, bh = body.width / EMU_IN, body.height / EMU_IN
                    kept_right_min = None
                    skip_els = {s._element for s in (body, tph) if s is not None}
                    for shp in slide.shapes:
                        if shp._element in skip_els:
                            continue
                        if getattr(shp, "is_placeholder", False) and not _placeholder_has_media(shp):
                            continue
                        if shp.shape_type == 13 or getattr(shp, "has_table", False) or getattr(shp, "has_chart", False) or shp.shape_type == 6 or _placeholder_has_media(shp):
                            sl = shp.left / EMU_IN
                            sw = shp.width / EMU_IN
                            sh = shp.height / EMU_IN
                            if sl >= SW_in * 0.42 and sw >= 1.5 and sh >= 1.2 and (sw * sh < 0.65 * SW_in * SH_in):
                                kept_right_min = sl if kept_right_min is None else min(kept_right_min, sl)
                    if kept_right_min is not None and (bx + bw) > kept_right_min - 0.15:
                        set_placeholder_box(body, bx, by, max(kept_right_min - bx - 0.25, 3.0), bh)
                except Exception:
                    pass

        if is_blank and is_cover:
            try:
                from pptx.enum.shapes import MSO_SHAPE
                ribbon = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0),
                    Inches(SW_in), Inches(0.10))
                ribbon.name = "CoverAccent"
                ribbon.fill.solid()
                ribbon.fill.fore_color.rgb = RGBColor(*BLANK_NAVY)
                ribbon.line.fill.background()
            except Exception:
                pass

        tph_used = set_title(slide, title, prs=prs, is_blank=is_blank, is_cover=is_cover, warnings=warnings)
        if tph is None and tph_used is not None:
            tph = tph_used
        if tph_used is not None and getattr(tph_used, "is_placeholder", False) and (title or "").strip():
            filled_els.add(tph_used._element)

        # Compute dynamic vertical bounds from actual title bottom & footer top
        title_bottom = 1.45
        if tph is not None:
            try:
                title_bottom = max(0.95, (tph.top + tph.height) / EMU_IN + 0.12)
            except Exception:
                pass
        _ev_top, _ev_w, safe_bottom = footer_layout_info(slide, prs)

        # Optional bottom Key Takeaway callout bar (reserve clean vertical slot above footer)
        takeaway_txt = (sld.get("takeaway") or "").strip()
        takeaway_top = None
        if (not is_cover and not keep_existing_media and takeaway_txt
                and (safe_bottom - title_bottom) >= 3.2):
            takeaway_top = safe_bottom - 0.44
            safe_bottom = takeaway_top - 0.10

        # Optional executive KPI callout strip on content slides (2-4 headline metrics)
        kpi_items = sld.get("kpis") or sld.get("metrics") or []
        if (not is_cover and not keep_existing_media and isinstance(kpi_items, list) and
                len(kpi_items) >= 2 and (safe_bottom - title_bottom) >= 3.4):
            kpi_bot = add_kpi_strip(slide, prs, kpi_items, 0.55, title_bottom, SW_in - 1.10, theme_accent=theme_accent)
            title_bottom = kpi_bot + 0.16
            # Push native placeholders cleanly below the KPI strip
            for ph_s in (body, right_ph, pic_ph):
                if ph_s is not None:
                    try:
                        px, py = ph_s.left / EMU_IN, ph_s.top / EMU_IN
                        pw, ph_h = ph_s.width / EMU_IN, ph_s.height / EMU_IN
                        if py < title_bottom:
                            new_h = max(safe_bottom - title_bottom, 1.5)
                            set_placeholder_box(ph_s, px, title_bottom, pw, min(ph_h, new_h))
                    except Exception:
                        pass

        avail_h = max(safe_bottom - title_bottom, 2.2)

        # Clamp non-cover body placeholder bottom so it never overlaps EvidenceLine
        if not is_cover and body is not None:
            try:
                bx, by = body.left / EMU_IN, body.top / EMU_IN
                bw, bh = body.width / EMU_IN, body.height / EMU_IN
                if by + bh > safe_bottom:
                    set_placeholder_box(body, bx, by, bw, max(safe_bottom - by, 1.5))
            except Exception:
                pass

        # Determine target media box (prefer native right_ph/pic_ph, or top-bullets + wide media for wide tables/banners)
        media_box = None
        two_col_text = (not has_media and not is_cover and not keep_existing_media and bool(right_bullets))
        wide_stacked = has_media and should_use_wide_stacked_layout(imgs, table_spec, bullets, is_cover, avail_h)
        if wide_stacked:
            if right_ph is not None:
                remove_shape(right_ph)
                right_ph = None
            if pic_ph is not None:
                remove_shape(pic_ph)
                pic_ph = None
            full_w = SW_in - 1.10
            # Estimate compact top bullets height at 15pt
            top_est = 0.0
            for lv, t in bullets:
                nl = est_lines(t, 15, max(full_w - 0.45 - 0.35 * lv, 3.0))
                top_est += nl * 1.28 * 15.0 / 72.0 + 0.07
            top_h = min(max(top_est + 0.20, 0.95), min(1.85, avail_h * 0.38))
            if body is not None:
                set_placeholder_box(body, 0.55, title_bottom, full_w, top_h)
            media_top = title_bottom + top_h + 0.14
            media_box = (0.55, media_top, full_w, max(safe_bottom - media_top, 1.8))
        elif two_col_text:
            if pic_ph is not None:
                remove_shape(pic_ph)
                pic_ph = None
            col_gap = 0.40
            col_w = (SW_in - 1.10 - col_gap) / 2.0
            left_x = 0.55
            right_x = left_x + col_w + col_gap
            col_top = title_bottom
            if left_heading or right_heading:
                dark_bg = is_dark_slide_bg(slide, prs)
                hdr_col_rgb = RGBColor(0x93, 0xC5, 0xFD) if dark_bg else RGBColor(*theme_accent)
                for h_name, hx, htxt in (("LeftColHeading", left_x, left_heading),
                                         ("RightColHeading", right_x, right_heading)):
                    if not htxt:
                        continue
                    hbox = slide.shapes.add_textbox(Inches(hx), Inches(title_bottom), Inches(col_w), Inches(0.36))
                    hbox.name = h_name
                    htf = hbox.text_frame
                    htf.word_wrap = True
                    htf.margin_top = htf.margin_bottom = htf.margin_left = htf.margin_right = 0
                    hrun = htf.paragraphs[0].add_run()
                    hrun.text = glue_num_units(htxt.replace("**", ""))[:60]
                    hrun.font.size = Pt(15)
                    hrun.font.bold = True
                    hrun.font.color.rgb = hdr_col_rgb
                col_top = title_bottom + 0.40
            col_h = max(safe_bottom - col_top, 1.8)
            # Subtle vertical hairline divider between columns
            try:
                from pptx.enum.shapes import MSO_SHAPE
                div_x = left_x + col_w + (col_gap / 2.0) - 0.008
                div = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(div_x), Inches(title_bottom + 0.05),
                                             Inches(0.016), Inches(max(safe_bottom - title_bottom - 0.15, 1.5)))
                div.name = "ColDivider"
                div.fill.solid()
                div.fill.fore_color.rgb = RGBColor(*( (0x33, 0x41, 0x55) if is_dark_slide_bg(slide, prs) else (0xDC, 0xE1, 0xE8) ))
                div.line.fill.background()
            except Exception:
                pass
            if body is not None:
                set_placeholder_box(body, left_x, col_top, col_w, col_h)
            if right_ph is not None:
                set_placeholder_box(right_ph, right_x, col_top, col_w, col_h)
            else:
                right_ph = slide.shapes.add_textbox(Inches(right_x), Inches(col_top), Inches(col_w), Inches(col_h))
                right_ph.name = "RightColBody"
                right_ph.text_frame.word_wrap = True
            if bullets and body is not None:
                fill_body(body, bullets, min(a.base_pt, 16), slide=slide, prs=prs, is_blank=is_blank, no_bullets=False)
                filled_els.add(body._element)
            elif bullets:
                ltb = slide.shapes.add_textbox(Inches(left_x), Inches(col_top), Inches(col_w), Inches(col_h))
                ltb.name = "LeftColBody"
                ltb.text_frame.word_wrap = True
                fill_body(ltb, bullets, min(a.base_pt, 16), slide=slide, prs=prs, is_blank=is_blank, no_bullets=False)
            elif body is not None:
                remove_shape(body)
            fill_body(right_ph, right_bullets, min(a.base_pt, 16), slide=slide, prs=prs, is_blank=is_blank, no_bullets=False)
            if getattr(right_ph, "is_placeholder", False):
                filled_els.add(right_ph._element)
            body = None
            bullets = []
        elif has_media:
            if pic_ph is not None:
                try:
                    media_box = (pic_ph.left / EMU_IN, pic_ph.top / EMU_IN,
                                 pic_ph.width / EMU_IN, min(pic_ph.height / EMU_IN, safe_bottom - pic_ph.top / EMU_IN))
                    remove_shape(pic_ph)  # replace with aspect-fit picture (zero axis cropping)
                except Exception:
                    media_box = None
            elif right_ph is not None and bullets:
                try:
                    media_box = (right_ph.left / EMU_IN, right_ph.top / EMU_IN,
                                 right_ph.width / EMU_IN, min(right_ph.height / EMU_IN, safe_bottom - right_ph.top / EMU_IN))
                    remove_shape(right_ph)
                except Exception:
                    media_box = None
            elif not bullets:
                # Full-width showcase when no left bullets
                if body is not None:
                    remove_shape(body)
                    body = None
                if right_ph is not None:
                    remove_shape(right_ph)
                media_box = (SW_in * 0.07, title_bottom, SW_in * 0.86, avail_h)
            else:
                # Single body placeholder: safely capture all 4 xfrm dims, then shrink to left half
                if body is not None:
                    try:
                        bx = body.left / EMU_IN
                        by = body.top / EMU_IN
                        bw = body.width / EMU_IN
                        bh = min(body.height / EMU_IN, safe_bottom - by)
                        if bx < SW_in * 0.40:
                            new_w = min(bw, SW_in * 0.45)
                            set_placeholder_box(body, bx, by, new_w, max(bh, 2.0))
                            mx = max(bx + new_w + 0.25, SW_in * 0.52)
                            media_box = (mx, title_bottom, SW_in - mx - 0.45, avail_h)
                        else:
                            # Body lives on right; place media on left
                            media_box = (0.45, title_bottom, max(bx - 0.70, SW_in * 0.42), avail_h)
                    except Exception:
                        media_box = (SW_in * 0.52, title_bottom, SW_in * 0.44, avail_h)
                else:
                    media_box = (SW_in * 0.52, title_bottom, SW_in * 0.44, avail_h)
        else:
            # No new media: avoid expanding body across preserved existing media when keep_existing_media=True
            if right_ph is not None:
                if not is_cover and not keep_existing_media and body is not None:
                    try:
                        bx, by = body.left / EMU_IN, body.top / EMU_IN
                        bw, bh = body.width / EMU_IN, body.height / EMU_IN
                        rx, ry, rw = right_ph.left / EMU_IN, right_ph.top / EMU_IN, right_ph.width / EMU_IN
                        if abs(by - ry) <= 0.6 and (rx + rw) > (bx + bw) + 1.0:
                            set_placeholder_box(body, bx, by, (rx + rw) - bx, bh)
                    except Exception:
                        pass
                if not (keep_existing_media and _placeholder_has_media(right_ph)):
                    remove_shape(right_ph)
            if pic_ph is not None:
                if not (keep_existing_media and _placeholder_has_media(pic_ph)):
                    remove_shape(pic_ph)

        b_base_pt = min(a.base_pt, 15) if wide_stacked else a.base_pt
        if bullets and body is not None:
            fill_body(body, bullets, b_base_pt, slide=slide, prs=prs, is_blank=is_blank, no_bullets=is_cover)
            filled_els.add(body._element)
        elif bullets:
            tb_w = (SW_in - 1.10) if wide_stacked else (SW_in * 0.45 if has_media else SW_in * 0.88)
            tb_h = top_h if wide_stacked else avail_h
            tb = slide.shapes.add_textbox(Inches(0.55), Inches(title_bottom),
                                          Inches(tb_w), Inches(tb_h))
            tb.text_frame.word_wrap = True
            fill_body(tb, bullets, b_base_pt, slide=slide, prs=prs, is_blank=is_blank, no_bullets=is_cover)
            warnings.append(f"slide {si+1}: layout lacks body placeholder; textbox used")
        elif body is not None and not is_cover:
            # Clear unused body placeholder on media-only slides
            if not (keep_existing_media and _placeholder_has_media(body)):
                remove_shape(body)

        if has_media and media_box:
            ml, mt, mw, mh = media_box
            mw = max(mw, 2.2)
            mh = max(mh, 1.8)
            if imgs:
                img_cap = (sld.get("image_caption") or sld.get("caption") or "").strip()
                if not img_cap and isinstance(sld.get("image_captions"), list):
                    img_cap = "  |  ".join(str(x).strip() for x in sld["image_captions"] if str(x).strip())
                cap_disp = glue_num_units(img_cap[:135]) if img_cap else ""
                cap_lines = est_lines(cap_disp, 10, max(mw - 0.06, 1.5)) if cap_disp else 0
                cap_pt = 9.5 if cap_lines >= 2 else 10
                cap_h = (0.40 if cap_lines >= 2 else 0.26) if cap_disp else 0.0
                cap_reserve = (cap_h + 0.08) if cap_disp else 0.0
                mh_clamped = min(mh, max(safe_bottom - mt, 1.4))
                pic_bottom = stack_images(slide, imgs, ml, mt, mw, max(mh_clamped - cap_reserve, 1.3), prs=prs)
                if cap_disp:
                    cap_top = min(pic_bottom + 0.04, safe_bottom - cap_h)
                    cbox = slide.shapes.add_textbox(Inches(ml), Inches(cap_top),
                                                    Inches(mw), Inches(cap_h))
                    cbox.name = "ImageCaption"
                    ctf = cbox.text_frame
                    ctf.word_wrap = True
                    ctf.margin_top = ctf.margin_bottom = ctf.margin_left = ctf.margin_right = 0
                    cpara = ctf.paragraphs[0]
                    cpara.alignment = PP_ALIGN.CENTER
                    crun = cpara.add_run()
                    crun.text = cap_disp
                    crun.font.size = Pt(cap_pt)
                    _apply_blank_font(crun, cap_disp, is_blank)
                    cfg = LIGHT_GRAY if is_dark_slide_bg(slide, prs) else GRAY
                    crun.font.color.rgb = RGBColor(*cfg)
            elif table_spec:
                add_table(slide, prs, table_spec, ml, mt, mw, mh,
                          is_blank=is_blank, theme_accent=theme_accent,
                          safe_bottom_override=safe_bottom)

        if takeaway_top is not None and takeaway_txt:
            add_takeaway_bar(slide, prs, takeaway_txt, 0.55, takeaway_top, SW_in - 1.10,
                             theme_accent=theme_accent, is_blank=is_blank)

        # End-of-slide sweep: remove any unfilled non-footer ghost placeholders (both newly added and --reuse)
        for ph_unused in init_content_phs:
            if ph_unused._element not in filled_els and ph_unused._element.getparent() is not None:
                if not (keep_existing_media and _placeholder_has_media(ph_unused)):
                    remove_shape(ph_unused)
        prune_empty_placeholders(slide, filled_els, keep_existing_media=keep_existing_media)

        if a.evidence_line and ev_text:
            add_evidence_line(slide, prs, ev_text, reserve_right_in=(1.35 if (is_blank and not is_cover) else 0.0))
        if is_blank and not is_cover:
            add_blank_slide_number(slide, prs, si + 1, len(slides))

        if notes:
            try:
                slide.notes_slide.notes_text_frame.text = notes
            except Exception as e:
                warnings.append(f"slide {si+1}: notes failed ({e})")

    os.makedirs(os.path.dirname(os.path.abspath(a.out)) or ".", exist_ok=True)
    prs.save(a.out)

    actual_n = len(prs.slides)
    report = {
        "ok": os.path.isfile(a.out) and os.path.getsize(a.out) > 0,
        "out": os.path.abspath(a.out),
        "n_slides": actual_n,
        "planned_slides": len(slides),
        "cleared_template_slides": cleared_template_slides,
        "template": os.path.abspath(a.template) if a.template else None,
        "template_mode": "reuse" if a.reuse else ("layouts" if a.template and not is_blank else "blank"),
        "theme_accent_hex": "#{:02X}{:02X}{:02X}".format(*theme_accent),
        "font_audit": font_audit,
        "reuse_mapping": reuse_mapping,
        "layouts_used": used_layouts,
        "slide_size_in": [round(SW_in, 3), round(SH_in, 3)],
        "warnings": warnings,
        "problems": problems,
    }
    rp = os.path.join(os.path.dirname(os.path.abspath(a.out)), "deck_build_report.json")
    with open(rp, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
