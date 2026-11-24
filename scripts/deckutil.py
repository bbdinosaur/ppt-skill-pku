#!/usr/bin/env python3
"""deckutil.py — shared helpers for the academic-slides skill scripts.

Pure python-pptx based helpers: text width metrics, layout picking,
placeholder filling, geometry utilities, change-log. No network, no caches.
"""
from __future__ import annotations

import json
import math
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.util import Emu, Inches, Pt
    HAS_PPTX = True
except ImportError:
    Presentation = None  # type: ignore[assignment]
    PP_PLACEHOLDER = None  # type: ignore[assignment]
    Emu = Inches = Pt = None  # type: ignore[assignment]
    HAS_PPTX = False

SCRIPT_DIR = Path(__file__).resolve().parent

if HAS_PPTX:
    TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
    BODY_TYPES = {
        PP_PLACEHOLDER.BODY,
        PP_PLACEHOLDER.OBJECT,
        PP_PLACEHOLDER.SUBTITLE,
        PP_PLACEHOLDER.VERTICAL_BODY,
        PP_PLACEHOLDER.VERTICAL_OBJECT,
    }
    SKIP_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}
    PIC_TYPES = {PP_PLACEHOLDER.PICTURE}
    SUBTITLE_TYPES = {PP_PLACEHOLDER.SUBTITLE}
else:
    TITLE_TYPES = set()
    BODY_TYPES = set()
    SKIP_TYPES = set()
    PIC_TYPES = set()
    SUBTITLE_TYPES = set()

DEFAULT_TITLE_PT = 28.0
DEFAULT_BODY_PT = 18.0
LINE_FACTOR = 1.28       # conservative line-height factor
PARA_GAP_EM = 0.45       # extra gap between paragraphs, in em units


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def log_change(work: Path, action: str, detail: str) -> None:
    """Append one JSON line to work/change_log.jsonl (time/change record)."""
    try:
        work = Path(work)
        work.mkdir(parents=True, exist_ok=True)
        entry = {"ts": now_iso(), "action": action, "detail": str(detail)[:400]}
        with (work / "change_log.jsonl").open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


def load_inventory(work: Path) -> dict:
    p = Path(work) / "inventory.json"
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


# ---------------------------------------------------------------- text metrics
def em_width(text: str) -> float:
    """Approximate rendered width of text in em units (works for CJK+latin mix)."""
    w = 0.0
    for ch in text:
        o = ord(ch)
        if 0x2E80 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF or 0x3000 <= o <= 0x303F or 0x2000 <= o <= 0x206F:
            w += 1.08  # CJK glyphs render slightly wider than 1em in most fonts
        elif ch in " .,:;'|!ilI()[]{}":
            w += 0.33
        elif ch.isdigit():
            w += 0.52
        elif ch.isupper():
            w += 0.62
        elif ch.isalpha():
            w += 0.48
        else:
            w += 0.50
    return w


def est_lines(text: str, size_pt: float, avail_w_pt: float) -> int:
    """Estimated wrapped line count for one paragraph."""
    if not text:
        return 1
    w_pt = em_width(text) * size_pt
    if avail_w_pt <= 0:
        return 99
    return max(1, math.ceil(w_pt / avail_w_pt))


def tf_text_height_pt(tf, shape_w_pt: float, default_pt: float) -> float:
    """Estimate rendered text height (pt) of a text frame at given shape width."""
    try:
        ml = tf.margin_left.pt if tf.margin_left is not None else 7.2
        mr = tf.margin_right.pt if tf.margin_right is not None else 7.2
    except Exception:
        ml = mr = 7.2
    base_avail = shape_w_pt - ml - mr
    total = 0.0
    for p in tf.paragraphs:
        text = "".join(r.text for r in p.runs) or ""
        if not text and not p.runs:
            text = ""
        size = default_pt
        for r in p.runs:
            if r.font.size is not None:
                size = r.font.size.pt
                break
        else:
            if p.font.size is not None:
                size = p.font.size.pt
        lvl = getattr(p, "level", 0) or 0
        avail = max(24.0, base_avail - lvl * 18.0)
        lines = est_lines(text, size, avail) if text else 1
        gap = size * 0.14 if (p.space_after is not None and p.space_after == 0) else size * PARA_GAP_EM
        total += lines * size * LINE_FACTOR + gap
    return total


def load_master_styles(prs, template_path=None):
    """Parse master txStyles (title/body lvl1 sizes, pt) and inherited text brightness
    (clrMap tx1 / titleStyle / bodyStyle srgbClr) into prs._master_styles (read-only XML)."""
    styles = {}
    path = template_path
    try:
        if not path:
            return styles
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(path)) as z:
            names = [n for n in z.namelist()
                     if _re.match(r"ppt/slideMasters/slideMaster\d+\.xml", n)]
            if not names:
                return styles
            xml = z.read(sorted(names)[0]).decode("utf-8", "replace")
        m = _re.search(r"<a:titleStyle>.*?<a:lvl1pPr>.*?<a:defRPr[^>]*sz=\"(\d{3,5})\"", xml, _re.S)
        if m:
            styles["title"] = int(m.group(1)) / 100.0
        m = _re.search(r"<a:bodyStyle>\s*<a:lvl1pPr>.*?<a:defRPr[^>]*sz=\"(\d{3,5})\"", xml, _re.S)
        if m:
            styles["body"] = int(m.group(1)) / 100.0
        # Detect whether master explicitly maps text (tx1) to light ink (lt1/lt2) or light srgbClr
        text_light = False
        cm = _re.search(r"<p:clrMap\b[^>]*\btx1=\"(lt1|lt2|dk1|dk2)\"", xml)
        if cm and cm.group(1) in ("lt1", "lt2"):
            text_light = True
        if not text_light:
            for tag in ("titleStyle", "bodyStyle"):
                mb = _re.search(rf"<a:{tag}>(.*?)</a:{tag}>", xml, _re.S)
                if mb:
                    cs = _re.search(r"<a:solidFill>\s*<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", mb.group(1))
                    if cs and _rel_luminance(_parse_hex(cs.group(1))) >= 0.45:
                        text_light = True
                        break
                    sc = _re.search(r"<a:solidFill>\s*<a:schemeClr val=\"(lt1|lt2|bg1)\"", mb.group(1))
                    if sc:
                        text_light = True
                        break
        styles["text_is_light"] = text_light
    except Exception:
        pass
    try:
        prs._master_styles = styles
    except Exception:
        pass
    return styles


def dark_slide_needs_light_ink(prs, slide) -> bool:
    """True when a slide has a dark background AND the master did not map default text to light ink
    (prevents unreadable black-on-navy text without ever modifying master or font.name)."""
    if not slide_is_dark(slide):
        return False
    ms = getattr(prs, "_master_styles", None) or {}
    return not bool(ms.get("text_is_light"))


def ensure_readable_tf_color(tf, rgb_tuple: tuple[int, int, int] = (0xF1, 0xF5, 0xF9)) -> None:
    """Apply readable foreground color to runs lacking explicit color (never sets font.name)."""
    try:
        from pptx.dml.color import RGBColor
        c_rgb = RGBColor(*rgb_tuple)
        for p in tf.paragraphs:
            for r in p.runs:
                has_explicit = False
                try:
                    if r.font.color is not None and r.font.color.type is not None:
                        has_explicit = True
                except Exception:
                    pass
                if not has_explicit:
                    r.font.color.rgb = c_rgb
    except Exception:
        pass


def ph_inherited_pt(slide_ph, fallback: float, master_styles=None) -> float:
    """Best-effort inherited font size: slide ph XML sz, then master txStyles, then fallback."""
    import re
    try:
        xml = slide_ph._element.xml
        m = re.search(r'sz="(\d{3,5})"', xml)
        if m:
            raw = int(m.group(1))
            if 900 <= raw <= 9900:
                return raw / 100.0
    except Exception:
        pass
    if master_styles and isinstance(master_styles, dict):
        try:
            t = slide_ph.placeholder_format.type
            key = "title" if t in TITLE_TYPES else "body"
            if key in master_styles:
                return float(master_styles[key])
        except Exception:
            pass
    # style-less template: renderers use their own (large) defaults; assume conservative
    # renderer defaults so that fit estimation and shrinking actually bite.
    try:
        t = slide_ph.placeholder_format.type
        if t in TITLE_TYPES:
            return 40.0
        if t in BODY_TYPES:
            return 28.0
    except Exception:
        pass
    return fallback


def effective_pt(shape, fallback: float) -> float:
    """Effective font size used for fit estimation (explicit run size wins, else inherited)."""
    try:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is not None:
                        return r.font.size.pt
                if p.font.size is not None:
                    return p.font.size.pt
    except Exception:
        pass
    return ph_inherited_pt(shape, fallback)


# ---------------------------------------------------------------- layout picking
HINT_KEYWORDS = {
    "title": ("title slide", "title and subtitle", "标题幻灯片", "标题和副标题"),
    "section": ("section header", "section", "节标题", "章节"),
    "content": ("title and content", "标题和内容", "content"),
    "two_content": ("two content", "两栏内容", "two column", "two-col"),
    "picture": ("picture with caption", "picture", "图片"),
    "quote": ("quote", "引言"),
    "blank": ("blank", "空白", "empty"),
}


def _layout_stats(layout):
    has_title = False
    has_pic = False
    nbody = 0
    for ph in layout.placeholders:
        t = ph.placeholder_format.type
        if t in TITLE_TYPES:
            has_title = True
        elif t in PIC_TYPES:
            has_pic = True
        elif t in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT,
                   PP_PLACEHOLDER.VERTICAL_BODY, PP_PLACEHOLDER.VERTICAL_OBJECT):
            nbody += 1  # SUBTITLE deliberately excluded: it is not a content column
    return has_title, has_pic, nbody


def pick_layout(prs, hint: str):
    """Pick best-matching layout for a plan layout_hint. Returns (layout, name)."""
    hint = (hint or "auto").lower().strip()
    layouts = list(prs.slide_layouts)
    best, best_score = None, -10e9
    for layout in layouts:
        name = (layout.name or "").lower()
        has_title, has_pic, nbody = _layout_stats(layout)
        s = 0.0
        for kw in HINT_KEYWORDS.get(hint, ()):
            if kw in name:
                s += 5.0
        if hint == "title":
            s += 4.0 * has_title - 2.0 * min(nbody, 2)
        elif hint == "section":
            s += 3.0 * has_title - 1.0 * min(nbody, 2)
        elif hint == "two_content":
            s += 2.0 * has_title + 3.0 * min(nbody, 2)
        elif hint == "picture":
            s += 2.0 * has_title + 1.0 * min(nbody, 1) + 2.0 * has_pic
        elif hint == "blank":
            s -= 1.0 * len(list(layout.placeholders))
        else:  # content / auto / unknown
            s += 2.0 * has_title + 2.0 * min(nbody, 2)
        s -= 0.05 * len(list(layout.placeholders))
        if s > best_score:
            best, best_score = layout, s
    if best is None:
        best = prs.slide_layouts[0]
    return best, best.name


# ---------------------------------------------------------------- shape helpers
def get_ph(slide, types):
    for ph in slide.placeholders:
        if ph.placeholder_format.type in types:
            return ph
    return None


def body_placeholders(slide):
    phs = [ph for ph in slide.placeholders if ph.placeholder_format.type in BODY_TYPES]
    phs.sort(key=lambda p: ((p.left or 0), (p.top or 0)))
    return phs


def slide_title_ph(slide):
    return get_ph(slide, TITLE_TYPES)


def ph_ok(v) -> bool:
    return v is not None and int(v) > 0


def ph_geom(prs, slide, ph):
    """Resolve placeholder geometry via inheritance chain slide→layout→master.
    Returns (left, top, width, height) with None/0 replaced by inherited values."""
    def get(p):
        try:
            return (p.left, p.top, p.width, p.height)
        except Exception:
            return (None, None, None, None)
    l, t, w, h = get(ph)
    idx = ph.placeholder_format.idx
    if not (ph_ok(w) and ph_ok(h)):
        try:
            containers = [slide.slide_layout, slide.slide_layout.slide_master]
        except Exception:
            containers = []
        for container in containers:
            for p in container.placeholders:
                if p.placeholder_format.idx == idx:
                    l2, t2, w2, h2 = get(p)
                    l = l if ph_ok(l) else l2
                    t = t if ph_ok(t) else t2
                    w = w if ph_ok(w) else w2
                    h = h if ph_ok(h) else h2
                    if ph_ok(w) and ph_ok(h):
                        break
            if ph_ok(w) and ph_ok(h):
                break
    return l, t, w, h


def shape_geom(prs, slide, sh):
    """Geometry for any shape; placeholders resolved via inheritance chain."""
    if getattr(sh, "is_placeholder", False):
        return ph_geom(prs, slide, sh)
    return sh.left, sh.top, sh.width, sh.height


def content_region(prs, slide):
    """Content area below title, above footer zone. Returns (l, t, w, h) as Emu ints."""
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    tph = slide_title_ph(slide)
    top = int(SH * 0.17)
    if tph is not None:
        _, tph_t, _, tph_h = ph_geom(prs, slide, tph)
        if ph_ok(tph_h):
            top = int(tph_t or 0) + int(tph_h) + Inches(0.22)
    top = min(int(SH * 0.28), max(int(SH * 0.14), top))
    bottom = int(SH * 0.905)
    left = int(SW * 0.065)
    right = int(SW * 0.935)
    return left, top, max(Inches(0.5), right - left), max(Inches(0.4), bottom - top)


def _flat_shape(sh, corner_adj: float | None = None) -> None:
    """Strip PowerPoint/LibreOffice default theme drop-shadows and optionally tighten corner radius."""
    try:
        sh.shadow.inherit = False
    except Exception:
        pass
    try:
        if corner_adj is not None and len(sh.adjustments) > 0:
            sh.adjustments[0] = corner_adj
    except Exception:
        pass
    try:
        from lxml import etree
        A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
        el = sh._element
        spPr = el.find(P + "spPr")
        if spPr is not None:
            for old in list(spPr.findall(A + "effectLst")):
                spPr.remove(old)
            etree.SubElement(spPr, A + "effectLst")
        style = el.find(P + "style")
        if style is not None:
            eref = style.find(A + "effectRef")
            if eref is not None:
                eref.set("idx", "0")
    except Exception:
        pass


def set_ph_rect(prs, slide, ph, left=None, top=None, width=None, height=None):
    """Set placeholder (left, top, width, height) atomically via spPr/xfrm.
    Resolves inherited values first so modifying top/width never zeroes left/top!"""
    l0, t0, w0, h0 = ph_geom(prs, slide, ph)
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    nl = int(left) if left is not None else (int(l0) if ph_ok(l0) else int(SW * 0.065))
    nt = int(top) if top is not None else (int(t0) if ph_ok(t0) else int(SH * 0.18))
    nw = int(width) if width is not None else (int(w0) if ph_ok(w0) else int(SW * 0.87))
    nh = int(height) if height is not None else (int(h0) if ph_ok(h0) else int(SH * 0.65))
    try:
        from lxml import etree
        A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        sppr = ph._element.spPr
        xfrm = sppr.xfrm
        if xfrm is None:
            xfrm = etree.SubElement(sppr, A + "xfrm")
        off = xfrm.find(A + "off")
        if off is None:
            off = etree.SubElement(xfrm, A + "off")
        off.set("x", str(nl))
        off.set("y", str(nt))
        ext = xfrm.find(A + "ext")
        if ext is None:
            ext = etree.SubElement(xfrm, A + "ext")
        ext.set("cx", str(nw))
        ext.set("cy", str(nh))
    except Exception:
        try:
            ph.left = Emu(nl)
            ph.top = Emu(nt)
            ph.width = Emu(nw)
            ph.height = Emu(nh)
        except Exception:
            pass


def set_ph_size(ph, w, h):
    """Set placeholder width & height atomically via xfrm ext attributes.
    (python-pptx's width setter alone can clobber height on cloned placeholder sp.)"""
    try:
        sppr = ph._element.spPr
        xfrm = sppr.xfrm
        if xfrm is None:
            ph.width = Emu(int(w))
            ph.height = Emu(int(h))
            return
        ext = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
        if ext is None:
            from lxml import etree
            ext = etree.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
        ext.set('cx', str(int(w)))
        ext.set('cy', str(int(h)))
    except Exception:
        try:
            ph.width = Emu(int(w))
            ph.height = Emu(int(h))
        except Exception:
            pass


def theme_palette(prs, template_path=None):
    """Read accent1..accent6 srgbClr from template theme; return list of '#RRGGBB'."""
    fallback = ["#1F4E79", "#2E75B6", "#ED7D31", "#70AD47", "#5B9BD5", "#A5A5A5"]
    if not template_path:
        return fallback
    try:
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(template_path)) as z:
            names = [n for n in z.namelist() if _re.match(r"ppt/theme/theme\d+\.xml", n)]
            if not names:
                return fallback
            xml = z.read(sorted(names)[0]).decode("utf-8", "replace")
        cols = []
        for idx in range(1, 7):
            m = _re.search(rf'<a:accent{idx}>\s*<a:srgbClr val="([0-9A-Fa-f]{{6}})"', xml)
            if m:
                cols.append("#" + m.group(1).upper())
        return cols if cols else fallback
    except Exception:
        return fallback


def theme_accent_color(prs, template_path=None):
    """Read accent1 srgbClr from the template theme; return '#RRGGBB' or None."""
    if not template_path:
        return None
    try:
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(template_path)) as z:
            names = [n for n in z.namelist() if _re.match(r"ppt/theme/theme\d+\.xml", n)]
            if not names:
                return None
            xml = z.read(sorted(names)[0]).decode("utf-8", "replace")
        m = _re.search(r'<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', xml)
        if m:
            return "#" + m.group(1).upper()
    except Exception:
        pass
    return None


def slide_is_dark(slide) -> bool:
    """Read-only XML check for dark solid background (luminance < 0.30) across
    slide -> layout -> master. NEVER touches python-pptx's mutating .background property!"""
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for owner in (slide, getattr(slide, "slide_layout", None),
                  getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
        if owner is None:
            continue
        try:
            el = owner._element
            bg = el.find(f"{P}cSld/{P}bg")
            if bg is None:
                continue
            srgb = bg.find(f"{P}bgPr/{A}solidFill/{A}srgbClr")
            if srgb is not None:
                h = srgb.get("val") or ""
                if len(h) == 6:
                    lum = _rel_luminance((int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))
                    return lum < 0.30
            # Explicit non-solid / scheme background at this level stops inheritance lookup
            if bg.find(f"{P}bgPr") is not None or bg.find(f"{P}bgRef") is not None:
                return False
        except Exception:
            continue
    return False


def _safe_accent_for_dark_bg(color: str | None) -> tuple[int, int, int]:
    """Brighten very dark accents so hairlines remain visible on dark slide canvases."""
    r, g, b = _parse_hex(color)
    if _rel_luminance((r, g, b)) >= 0.22:
        return (r, g, b)
    return (min(255, int(r + (255 - r) * 0.55)),
            min(255, int(g + (255 - g) * 0.55)),
            min(255, int(b + (255 - b) * 0.55)))


def accent_bar(slide, prs, color, left=None, top=None, width_emu=None, height_in=0.022):
    """Subtle full-width hairline divider below the title (works cleanly with both
    centered and left-aligned template titles, light and dark canvases)."""
    if not color:
        return None
    try:
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Emu
        tph = slide_title_ph(slide)
        SW = int(prs.slide_width)
        rl = int(SW * 0.065)
        rw = int(SW * 0.87)
        if left is None:
            left = rl
        if width_emu is None:
            width_emu = rw
        if top is None:
            if tph is not None:
                _, tp, _, th = ph_geom(prs, slide, tph)
                if ph_ok(th) and ph_ok(tp):
                    top = int(tp) + int(th) + int(Inches(0.02))
                else:
                    top = int(prs.slide_height * 0.17)
            else:
                top = int(prs.slide_height * 0.17)
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Emu(int(left)), Emu(int(top)),
                                      Emu(int(width_emu)), Emu(int(Inches(height_in))))
        rect.fill.solid()
        if slide_is_dark(slide):
            rect.fill.fore_color.rgb = RGBColor(*_safe_accent_for_dark_bg(color))
        else:
            rect.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
        rect.line.fill.background()
        _flat_shape(rect)
        return rect
    except Exception:
        return None


def fit_textbox(slide, rect, text, size_pt=None, bold=False, italic=False, align=None, color=None):
    """Add a plain textbox inside rect. No font-name override (theme font inherited)."""
    from pptx.dml.color import RGBColor
    from pptx.util import Emu

    l, t, w, h = rect
    tb = slide.shapes.add_textbox(Emu(int(l)), Emu(int(t)), Emu(int(w)), Emu(int(h)))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = glue_num_units(str(text))
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    if size_pt is not None:
        for r in p.runs:
            r.font.size = Pt(size_pt)
        p.font.size = Pt(size_pt)
    if bold:
        for r in p.runs:
            r.font.bold = True
    if italic:
        for r in p.runs:
            r.font.italic = True
    if color:
        rgb = RGBColor.from_string(color.lstrip("#"))
        for r in p.runs:
            r.font.color.rgb = rgb
    return tb


_MD_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_MD_CODE_RE = re.compile(r"`([^`\n]+)`")
_SUP_CH = str.maketrans("0123456789+-()nijk", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾ⁿⁱʲᵏ")
_SUB_CH = str.maketrans("0123456789+-()aehklmnopstx", "₀₁₂₃₄₅₆₇₈₉₊₋₍₎ₐₑₕₖₗₘₙₒₚₛₜₓ")
_LATEX_SYMS = (
    ("\\times", "×"), ("\\pm", "±"), ("\\mp", "∓"), ("\\leq", "≤"), ("\\le", "≤"),
    ("\\geq", "≥"), ("\\ge", "≥"), ("\\approx", "≈"), ("\\neq", "≠"), ("\\ne", "≠"),
    ("\\cdot", "·"), ("\\rightarrow", "→"), ("\\to", "→"), ("\\leftarrow", "←"),
    ("\\leftrightarrow", "↔"), ("\\infty", "∞"), ("\\sim", "~"), ("\\equiv", "≡"),
    ("\\Delta", "Δ"), ("\\Sigma", "Σ"), ("\\Omega", "Ω"), ("\\Lambda", "Λ"),
    ("\\alpha", "α"), ("\\beta", "β"), ("\\gamma", "γ"), ("\\delta", "δ"),
    ("\\varepsilon", "ε"), ("\\epsilon", "ε"), ("\\zeta", "ζ"), ("\\eta", "η"),
    ("\\theta", "θ"), ("\\lambda", "λ"), ("\\mu", "μ"), ("\\nu", "ν"),
    ("\\pi", "π"), ("\\rho", "ρ"), ("\\sigma", "σ"), ("\\tau", "τ"),
    ("\\phi", "φ"), ("\\psi", "ψ"), ("\\omega", "ω"),
)
_LATEX_CMD_RE = re.compile(r"\\(?:times|pm|mp|leq|le|geq|ge|approx|neq|ne|cdot|rightarrow|to|leftarrow|leftrightarrow|infty|sim|equiv|Delta|Sigma|Omega|Lambda|alpha|beta|gamma|delta|varepsilon|epsilon|zeta|eta|theta|lambda|mu|nu|pi|rho|sigma|tau|phi|psi|omega)\b")
_INLINE_MATH_RE = re.compile(r"\$([^$\n]{1,64})\$")
_SUP_BRACE_RE = re.compile(r"(?<=[A-Za-z0-9\)\]])\^\{([0-9+\-()nijk]{1,6})\}")
_SUP_BARE_RE = re.compile(r"(?<=[A-Za-z0-9\)\]])\^([+-]?\d{1,3})(?![A-Za-z0-9])")
_SUB_BRACE_RE = re.compile(r"(?<=[A-Za-z])_\{([0-9]{1,3}|[aehklmnopstx]{1,2})\}")
_SUB_BARE_RE = re.compile(r"\b(CO|H|O|N|L|F|R|T|P|S|C)_([0-9]{1,2})\b")


def clean_academic_math(text: str) -> str:
    """Convert raw LaTeX inline math ($...$, \\times, \\alpha, 10^{-6}, R^2, CO_2) into
    clean publication Unicode typography while preserving currency ($100) and plain underscores."""
    s = str(text or "")
    if not any(ch in s for ch in ("$", "\\", "^", "_")):
        return s

    def _unwrap_math(m):
        inner = m.group(1)
        # Keep currency intact (e.g., '$5 to $10') — only unwrap if inner has LaTeX/math cues
        if re.match(r"^\s*\d[\d,.]*(?:\s*[kKMB])?\s*$", inner):
            return m.group(0)
        if any(c in inner for c in ("\\", "^", "_", "=", "<", ">", "×", "±")) or re.fullmatch(r"[A-Za-z](?:_[A-Za-z0-9]+)?", inner.strip()):
            return inner.strip()
        return m.group(0)

    s = _INLINE_MATH_RE.sub(_unwrap_math, s)
    if "\\" in s:
        sym_map = dict(_LATEX_SYMS)
        s = _LATEX_CMD_RE.sub(lambda m: sym_map.get(m.group(0), m.group(0)), s)
    s = _SUP_BRACE_RE.sub(lambda m: m.group(1).translate(_SUP_CH), s)
    s = _SUP_BARE_RE.sub(lambda m: m.group(1).translate(_SUP_CH), s)
    s = _SUB_BRACE_RE.sub(lambda m: m.group(1).translate(_SUB_CH), s)
    s = _SUB_BARE_RE.sub(lambda m: m.group(1) + m.group(2).translate(_SUB_CH), s)
    return s


_NUM_UNIT_RE = re.compile(
    r"(?<=[\d⁰¹²³⁴⁵⁶⁷⁸⁹]) (?=(?:ms|us|μs|µs|ns|ps|fs|s|Hz|kHz|MHz|GHz|THz|KB|MB|GB|TB|PB|"
    r"nm|μm|µm|mm|cm|km|m|K|mK|μK|dB|W|mW|kW|MW|V|mV|kV|A|mA|eV|keV|MeV|GeV|TeV|"
    r"Pa|kPa|MPa|GPa|mol|mmol|mg|kg|g|L|mL|万|亿|倍|%)(?![A-Za-z0-9]))"
)


def glue_num_units(text: str) -> str:
    """Bind numeric values to short following units with non-breaking space (\\u00a0)
    so rendered lines never orphan units like '8.4\\nms' or '1.2\\n万'."""
    s = str(text or "")
    if " " not in s:
        return s
    return _NUM_UNIT_RE.sub("\u00a0", s)


def strip_md(text: str) -> str:
    """Strip inline Markdown **bold** and `code` ticks and format academic LaTeX math cleanly."""
    s = str(text or "")
    s = _MD_BOLD_RE.sub(r"\1", s)
    s = _MD_CODE_RE.sub(r"\1", s)
    return clean_academic_math(s)


def no_bullet(paragraph) -> None:
    """Suppress bullet symbol on a native placeholder paragraph via <a:buNone/>."""
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        pPr = paragraph._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buBlip", "a:buNone"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        etree.SubElement(pPr, qn("a:buNone"))
    except Exception:
        pass


def add_rich_runs(p, text: str, base_bold: bool = False, size_pt: float | None = None,
                  color_rgb=None) -> None:
    """Append runs to paragraph `p`, converting Markdown **bold** into native bold runs,
    stripping `code` ticks, and formatting academic math while NEVER overriding font.name."""
    raw = glue_num_units(clean_academic_math(_MD_CODE_RE.sub(r"\1", str(text or ""))))
    pos = 0
    emitted = False
    for m in _MD_BOLD_RE.finditer(raw):
        if m.start() > pos:
            r = p.add_run()
            r.text = raw[pos : m.start()]
            if base_bold:
                r.font.bold = True
            if size_pt is not None:
                r.font.size = Pt(float(size_pt))
            if color_rgb is not None:
                r.font.color.rgb = color_rgb
            emitted = True
        r = p.add_run()
        r.text = m.group(1)
        r.font.bold = True
        if size_pt is not None:
            r.font.size = Pt(float(size_pt))
        if color_rgb is not None:
            r.font.color.rgb = color_rgb
        emitted = True
        pos = m.end()
    if pos < len(raw) or not emitted:
        r = p.add_run()
        r.text = raw[pos:]
        if base_bold:
            r.font.bold = True
        if size_pt is not None:
            r.font.size = Pt(float(size_pt))
        if color_rgb is not None:
            r.font.color.rgb = color_rgb


def _norm_bullets_list(raw):
    if isinstance(raw, str):
        return [{"t": raw, "lvl": 0}] if raw.strip() else []
    if not isinstance(raw, list):
        return []
    out = []
    for b in raw:
        if isinstance(b, dict):
            nb = dict(b)
            if "t" not in nb:
                for alt in ("text", "content", "point", "body", "label"):
                    if nb.get(alt) is not None:
                        nb["t"] = str(nb[alt])
                        break
            out.append(nb)
        elif b is not None and str(b).strip():
            out.append({"t": str(b), "lvl": 0})
    return out


def _norm_sources_list(raw):
    if isinstance(raw, dict):
        return [raw]
    if isinstance(raw, list):
        return [x for x in raw if isinstance(x, dict)]
    return []


def _norm_slide_spec(spec: dict) -> dict:
    if not isinstance(spec, dict):
        return {}
    s = dict(spec)
    if "bullets" in s:
        s["bullets"] = _norm_bullets_list(s.get("bullets"))
    if "sources" in s:
        s["sources"] = _norm_sources_list(s.get("sources"))
    tc = s.get("two_col")
    if isinstance(tc, dict):
        ntc = dict(tc)
        if "left_bullets" not in ntc and "left" in ntc:
            ntc["left_bullets"] = ntc.get("left")
        if "right_bullets" not in ntc and "right" in ntc:
            ntc["right_bullets"] = ntc.get("right")
        if "left_bullets" in ntc:
            ntc["left_bullets"] = _norm_bullets_list(ntc.get("left_bullets"))
        if "right_bullets" in ntc:
            ntc["right_bullets"] = _norm_bullets_list(ntc.get("right_bullets"))
        s["two_col"] = ntc
    hi = s.get("highlights")
    if isinstance(hi, list):
        nhi = []
        for it in hi:
            if isinstance(it, dict):
                nit = dict(it)
                if not nit.get("value"):
                    for alt in ("metric", "num", "stat", "number", "val"):
                        if nit.get(alt) is not None:
                            nit["value"] = str(nit[alt])
                            break
                if not nit.get("label"):
                    for alt in ("title", "name", "desc", "caption"):
                        if nit.get(alt) is not None:
                            nit["label"] = str(nit[alt])
                            break
                nhi.append(nit)
        s["highlights"] = nhi
    st = s.get("steps")
    if isinstance(st, list):
        nst = []
        for it in st:
            if isinstance(it, dict):
                nit = dict(it)
                if not nit.get("title"):
                    for alt in ("name", "label", "stage", "step"):
                        if nit.get(alt) is not None:
                            nit["title"] = str(nit[alt])
                            break
                if not nit.get("desc") and not nit.get("detail"):
                    for alt in ("text", "summary", "body", "description"):
                        if nit.get(alt) is not None:
                            nit["desc"] = str(nit[alt])
                            break
                nst.append(nit)
            elif it is not None and str(it).strip():
                nst.append(str(it))
        s["steps"] = nst
    tbl = s.get("table")
    if isinstance(tbl, dict):
        ntbl = dict(tbl)
        if not ntbl.get("rows") and isinstance(ntbl.get("data"), list):
            ntbl["rows"] = ntbl.get("data")
        hdrs = ntbl.get("headers") or ntbl.get("columns")
        rows = list(ntbl.get("rows") or [])
        if isinstance(hdrs, list) and hdrs:
            hdr_strs = [str(c) for c in hdrs]
            if not rows or [str(c) for c in rows[0]] != hdr_strs:
                rows = [hdr_strs] + rows
            ntbl["rows"] = rows
        s["table"] = ntbl
    ch = s.get("chart")
    if isinstance(ch, dict):
        nch = dict(ch)
        if not nch.get("categories") and isinstance(nch.get("labels"), list):
            nch["categories"] = nch.get("labels")
        if not nch.get("series") and isinstance(nch.get("data"), list):
            nch["series"] = nch.get("data")
        s["chart"] = nch
    return s


def normalize_plan(plan: dict) -> dict:
    """Deterministically normalize common LLM schema aliases (single-dict sources,
    bullets[].text, two_col.left/right, table.headers/data, chart.labels/data,
    highlights[].metric, steps[].label) without altering facts."""
    if not isinstance(plan, dict):
        return {}
    out = dict(plan)
    if isinstance(out.get("slides"), list):
        out["slides"] = [_norm_slide_spec(s) if isinstance(s, dict) else s for s in out["slides"]]
    if isinstance(out.get("edits"), list):
        nedits = []
        for e in out["edits"]:
            if not isinstance(e, dict):
                continue
            ne = dict(e)
            if "sources" in ne:
                ne["sources"] = _norm_sources_list(ne.get("sources"))
            if "bullets" in ne:
                ne["bullets"] = _norm_bullets_list(ne.get("bullets"))
            if isinstance(ne.get("two_col"), dict):
                tc_e = _norm_slide_spec({"two_col": ne["two_col"]}).get("two_col") or {}
                for k_tc in ("left_title", "right_title", "left_bullets", "right_bullets"):
                    if k_tc not in ne and k_tc in tc_e:
                        ne[k_tc] = tc_e[k_tc]
            if "left_bullets" not in ne and "left" in ne:
                ne["left_bullets"] = ne.get("left")
            if "right_bullets" not in ne and "right" in ne:
                ne["right_bullets"] = ne.get("right")
            if "left_bullets" in ne:
                ne["left_bullets"] = _norm_bullets_list(ne.get("left_bullets"))
            if "right_bullets" in ne:
                ne["right_bullets"] = _norm_bullets_list(ne.get("right_bullets"))
            if isinstance(ne.get("table"), dict):
                ne["table"] = _norm_slide_spec({"table": ne["table"]}).get("table")
            elif isinstance(ne.get("headers"), list) and ne.get("headers"):
                hdr_strs = [str(c) for c in ne["headers"]]
                rows = list(ne.get("rows") or ne.get("data") or [])
                if not rows or [str(c) for c in rows[0]] != hdr_strs:
                    ne["rows"] = [hdr_strs] + rows
            if isinstance(ne.get("chart"), dict):
                ne["chart"] = _norm_slide_spec({"chart": ne["chart"]}).get("chart")
            elif not ne.get("categories") and isinstance(ne.get("labels"), list):
                ne["categories"] = ne.get("labels")
            if isinstance(ne.get("slide_spec"), dict):
                ne["slide_spec"] = _norm_slide_spec(ne["slide_spec"])
            nedits.append(ne)
        out["edits"] = nedits
    return out


def fill_bullets_tf(tf, bullets, default_pt=None, accent_rgb=None):
    """Fill a text frame with planned bullets (replaces existing content).
    Renders inline Markdown **bold** as native bold runs and inherits template fonts."""
    try:
        tf.clear()  # drop pre-existing paragraphs (edit mode must not merge with old text)
    except Exception:
        pass
    tf.word_wrap = True
    first = True
    for b in bullets:
        if isinstance(b, dict):
            text = str(b.get("t", b.get("text", "")))
            lvl = int(b.get("lvl", 0) or 0)
            bold = bool(b.get("bold", False))
            size = b.get("size")
            is_col_hdr = bool(b.get("_col_header", False))
        else:
            text, lvl, bold, size, is_col_hdr = str(b), 0, False, None, False
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        try:
            p.level = max(0, min(4, lvl))
        except Exception:
            pass
        if is_col_hdr:
            no_bullet(p)
            try:
                p.space_after = Pt(6)
            except Exception:
                pass
            add_rich_runs(p, text, base_bold=True, size_pt=size, color_rgb=accent_rgb)
            continue
        # Auto-bold crisp lead-in label before first ':'/'：' (2..22 chars, non-URL)
        # for executive academic scannability — never sets font.name!
        split_done = False
        has_md_bold = "**" in text
        if not bold and text:
            plain = strip_md(text)
            for sep in ("：", ": "):
                idx = plain.find(sep)
                if 2 <= idx <= 22:
                    head = plain[: idx + len(sep)]
                    if not head.lower().startswith(("http", "https", "ftp")):
                        if has_md_bold:
                            ridx = text.find(sep)
                            if 2 <= ridx <= 28:
                                rhead = strip_md(text[: ridx + len(sep)])
                                rtail = text[ridx + len(sep) :]
                                if rtail.strip():
                                    add_rich_runs(p, rhead, base_bold=True, size_pt=size)
                                    add_rich_runs(p, rtail, base_bold=False, size_pt=size)
                                    split_done = True
                                    break
                        else:
                            tail = text[idx + len(sep) :]
                            if tail.strip():
                                add_rich_runs(p, head, base_bold=True, size_pt=size)
                                add_rich_runs(p, tail, base_bold=False, size_pt=size)
                                split_done = True
                                break
        if not split_done:
            add_rich_runs(p, text, base_bold=bold, size_pt=size)
    return tf


def shrink_body_to_fit(tf, w_pt, h_pt, body_pt, min_pt=11.0, compact_spacing_pt=None):
    """Step down body font size until estimated height fits, or floor reached.
    Applies compact paragraph spacing/margins BEFORE measuring so larger fonts fit cleanly."""
    if w_pt <= 0 or h_pt <= 0:
        return body_pt
    if compact_spacing_pt is not None:
        try:
            for p in tf.paragraphs:
                p.space_after = Pt(float(compact_spacing_pt))
                p.space_before = Pt(0)
            if float(compact_spacing_pt) <= 1.5:
                tf.margin_top = Pt(2)
                tf.margin_bottom = Pt(2)
        except Exception:
            pass
    size = body_pt
    while size > min_pt:
        est = tf_text_height_pt(tf, w_pt, size)
        if est <= h_pt * 0.98:
            break
        step = 0.5 if size <= 13.0 else 1.0
        size = max(min_pt, size - step)
    # apply explicit size when we had to shrink below inherited size
    if size < body_pt - 0.3:
        for p in tf.paragraphs:
            p.font.size = Pt(size)
            for r in p.runs:
                if r.font.size is None or r.font.size.pt > size:
                    r.font.size = Pt(size)
    return size


def clamp_slide_shapes(prs, slide) -> None:
    """Keep non-background shapes safely within [0, SW] x [0, SH] to eliminate minor bounds slips."""
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    pad = int(Inches(0.05))
    for sh in slide.shapes:
        try:
            l, t, w, h = shape_geom(prs, slide, sh)
            if l is None or t is None or w is None or h is None:
                continue
            li, ti, wi, hi = int(l), int(t), int(w), int(h)
            if wi <= 0 or hi <= 0 or wi * hi > 0.80 * SW * SH:
                continue
            nl, nt, nw, nh = li, ti, wi, hi
            if nl < 0:
                nl = pad
            if nt < 0:
                nt = pad
            if nl + nw > SW:
                nw = max(int(Inches(0.4)), SW - nl - pad)
            if nt + nh > SH:
                if nt > pad and hi <= SH - 2 * pad:
                    nt = max(pad, SH - hi - pad)
                else:
                    nh = max(int(Inches(0.25)), SH - nt - pad)
            if (nl, nt, nw, nh) != (li, ti, wi, hi):
                if getattr(sh, "is_placeholder", False):
                    set_ph_rect(prs, slide, sh, left=nl, top=nt, width=nw, height=nh)
                else:
                    sh.left, sh.top, sh.width, sh.height = Emu(nl), Emu(nt), Emu(nw), Emu(nh)
        except Exception:
            continue


def set_notes(slide, notes: str) -> None:
    if not notes:
        return
    try:
        slide.notes_slide.notes_text_frame.text = str(notes)
    except Exception:
        pass


def delete_slide(prs, index0: int) -> None:
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    sid = ids[index0]
    rId = sid.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    try:
        prs.part.drop_rel(rId)
    except Exception:
        pass
    sldIdLst.remove(sid)


def move_slide(prs, src0: int, dst0: int) -> None:
    lst = prs.slides._sldIdLst
    ids = list(lst)
    el = ids[src0]
    lst.remove(el)
    lst.insert(dst0, el)


def parse_chart_num(v) -> tuple[float, bool]:
    """Parse numeric chart value robustly (accepts '85.4%', '1,240', '+58.4%', '−1.5').
    Returns (float_value, had_percent_sign)."""
    if isinstance(v, (int, float)):
        return float(v), False
    s = str(v or "").strip()
    has_pct = "%" in s
    s = s.replace(",", "").replace("%", "").replace("−", "-").replace("–", "-").strip()
    if s.startswith("+"):
        s = s[1:].strip()
    return float(s), has_pct


def add_native_chart(slide, rect, chart_spec, palette=None):
    """Add a fully-editable native PPT chart bound to plan data (创新点：图表数据绑定).
    chart_spec: {type: bar|hbar|line|pie, categories: [...], series: [{name, values}],
                 ref, caption, title, data_labels}. Returns the graphic frame."""
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.util import Emu

    ctype = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "stacked_bar": XL_CHART_TYPE.COLUMN_STACKED,
             "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
             "stacked_hbar": XL_CHART_TYPE.BAR_STACKED,
             "line": XL_CHART_TYPE.LINE_MARKERS,
             "pie": XL_CHART_TYPE.PIE}.get(str(chart_spec.get("type", "bar")).lower(),
                                           XL_CHART_TYPE.COLUMN_CLUSTERED)
    cats = [strip_md(str(c)) for c in (chart_spec.get("categories") or [])]
    series_list = chart_spec.get("series") or []
    cd = CategoryChartData()
    cd.categories = cats
    any_pct = str(chart_spec.get("unit") or "").strip() == "%"
    for s in series_list:
        raw_vals = list(s.get("values") or [])
        parsed_vals = []
        for rv in raw_vals[:len(cats)]:
            fv, is_p = parse_chart_num(rv)
            if is_p:
                any_pct = True
            parsed_vals.append(fv)
        while len(parsed_vals) < len(cats):
            parsed_vals.append(0.0)
        cd.add_series(strip_md(str(s.get("name", "series"))), tuple(parsed_vals))
    l, t, w, h = [int(v) for v in rect]
    gf = slide.shapes.add_chart(ctype, Emu(l), Emu(t), Emu(w), Emu(h), cd)
    try:
        chart = gf.chart
        n_series = len(series_list)
        chart.has_legend = (n_series > 1) or (ctype == XL_CHART_TYPE.PIE)
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(10.5)
        if chart_spec.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = strip_md(str(chart_spec["title"]))
            for p in chart.chart_title.text_frame.paragraphs:
                p.font.size = Pt(13.5)
                for r in p.runs:
                    r.font.size = Pt(13.5)
                    r.font.bold = True
        else:
            chart.has_title = False
        # Data labels for crisp academic numbers (auto on for <=16 points, <=12 for line)
        total_pts = len(cats) * max(1, n_series)
        default_dl = (total_pts <= 12) if (ctype == XL_CHART_TYPE.LINE_MARKERS) else (total_pts <= 16)
        show_labels = chart_spec.get("data_labels", default_dl)
        dark_bg = slide_is_dark(slide)
        any_frac = False
        for sr_d in (chart_spec.get("series") or []):
            for rv in (sr_d.get("values") or []):
                fv, _ = parse_chart_num(rv)
                if abs(fv - round(fv)) > 1e-6:
                    any_frac = True
                    break
        if show_labels and chart.plots:
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.font.size = Pt(10.5 if n_series > 1 else 11.0)
            dl.font.bold = True
            try:
                if ctype == XL_CHART_TYPE.LINE_MARKERS:
                    dl.position = XL_LABEL_POSITION.ABOVE
                elif ctype in (XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED):
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END
                elif ctype in (XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED):
                    dl.position = XL_LABEL_POSITION.CENTER
                elif ctype == XL_CHART_TYPE.PIE:
                    dl.show_percentage = True
                    dl.show_category_name = True
                    dl.show_value = bool(chart_spec.get("show_value", False))
                    dl.position = XL_LABEL_POSITION.BEST_FIT
                    # Prevent PowerPoint/LibreOffice default '0%' integer rounding on fractional shares
                    if any_pct or any_frac:
                        dl.number_format = "0.0%"
            except Exception:
                pass
            if any_pct and ctype != XL_CHART_TYPE.PIE:
                try:
                    dl.number_format = '0.0"%"'
                except Exception:
                    pass
            if dark_bg and ctype not in (XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED):
                dl.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
        # Apply theme palette to series (never sets font.name)
        cols = palette or ["#1F4E79", "#2E75B6", "#ED7D31", "#70AD47", "#5B9BD5"]
        if ctype == XL_CHART_TYPE.PIE:
            try:
                if chart.series:
                    sr0 = chart.series[0]
                    lums = []
                    for pi in range(len(cats)):
                        c_hex = cols[pi % len(cols)].lstrip("#")
                        rgb_t = _safe_accent_for_dark_bg(c_hex) if dark_bg else _parse_hex(c_hex)
                        c_rgb = RGBColor(*rgb_t)
                        lums.append(_rel_luminance(rgb_t))
                        pt = sr0.points[pi]
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.rgb = c_rgb
                    if show_labels and chart.plots and lums:
                        avg_lum = sum(lums) / len(lums)
                        chart.plots[0].data_labels.font.color.rgb = (
                            RGBColor(0xFF, 0xFF, 0xFF) if avg_lum < 0.30 else RGBColor(0x1A, 0x20, 0x2C)
                        )
            except Exception:
                pass
        else:
            for idx, sr in enumerate(chart.series):
                c_hex = cols[idx % len(cols)].lstrip("#")
                try:
                    rgb_t = _safe_accent_for_dark_bg(c_hex) if dark_bg else _parse_hex(c_hex)
                    c_rgb = RGBColor(*rgb_t)
                    if ctype == XL_CHART_TYPE.LINE_MARKERS:
                        sr.format.line.color.rgb = c_rgb
                        sr.format.line.width = Pt(2.4)
                        try:
                            sr.marker.size = 7
                            sr.marker.format.fill.solid()
                            sr.marker.format.fill.fore_color.rgb = c_rgb
                        except Exception:
                            pass
                    else:
                        sr.format.fill.solid()
                        sr.format.fill.fore_color.rgb = c_rgb
                        # Contrast-adaptive CENTER data label ink per series on stacked bars
                        if show_labels and ctype in (XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED):
                            try:
                                sr.has_data_labels = True
                                sdl = sr.data_labels
                                sdl.show_value = True
                                sdl.position = XL_LABEL_POSITION.CENTER
                                sdl.font.size = Pt(10.5 if n_series > 1 else 11.0)
                                sdl.font.bold = True
                                if any_pct:
                                    sdl.number_format = '0.0"%"'
                                sdl.font.color.rgb = (
                                    RGBColor(0xFF, 0xFF, 0xFF) if _rel_luminance(rgb_t) < 0.42
                                    else RGBColor(0x1A, 0x20, 0x2C)
                                )
                            except Exception:
                                pass
                        # Optional highlight_last on single-series bar/hbar: mute baselines, pop final bar
                        if n_series == 1 and chart_spec.get("highlight_last") and len(cats) >= 2 and ctype in (
                            XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED
                        ):
                            muted = RGBColor(0x94, 0xA3, 0xB8) if dark_bg else RGBColor(0xB8, 0xC4, 0xD4)
                            for pi in range(len(cats)):
                                try:
                                    pt = sr.points[pi]
                                    pt.format.fill.solid()
                                    pt.format.fill.fore_color.rgb = c_rgb if (pi == len(cats) - 1) else muted
                                except Exception:
                                    pass
                except Exception:
                    pass
            for axis_attr, title_key in (("category_axis", "x_title"), ("value_axis", "y_title")):
                try:
                    ax = getattr(chart, axis_attr)
                    ax.tick_labels.font.size = Pt(11.5)
                    if dark_bg:
                        ax.tick_labels.font.color.rgb = RGBColor(0xD8, 0xE2, 0xEC)
                    if chart_spec.get(title_key):
                        ax.has_title = True
                        ax.axis_title.text_frame.text = strip_md(str(chart_spec[title_key]))
                        for p in ax.axis_title.text_frame.paragraphs:
                            p.font.size = Pt(11.5)
                            p.font.bold = True
                            if dark_bg:
                                p.font.color.rgb = RGBColor(0xD8, 0xE2, 0xEC)
                    if axis_attr == "value_axis":
                        if chart_spec.get("y_min") is not None:
                            ax.minimum_scale = float(chart_spec["y_min"])
                        if chart_spec.get("y_max") is not None:
                            ax.maximum_scale = float(chart_spec["y_max"])
                        if getattr(ax, "has_major_gridlines", False):
                            ax.major_gridlines.format.line.color.rgb = (
                                RGBColor(0x47, 0x55, 0x69) if dark_bg else RGBColor(0xE2, 0xE8, 0xF0)
                            )
                except Exception:
                    pass
        if dark_bg and chart.has_legend:
            try:
                chart.legend.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
            except Exception:
                pass
        if dark_bg and chart.has_title:
            try:
                for p in chart.chart_title.text_frame.paragraphs:
                    p.font.color.rgb = RGBColor(0xF1, 0xF5, 0xF9)
            except Exception:
                pass
    except Exception:
        pass
    return gf


def format_caption(ref: str | None, caption: str | None) -> str:
    """Join figure/table ref and caption cleanly with academic separator without duplicating prefix."""
    r = strip_md(str(ref or "")).strip()
    c = strip_md(str(caption or "")).strip()
    if not r:
        return c
    if not c:
        return r
    r_low = r.rstrip(":：. ").lower()
    c_low = c.lower()
    if r_low and (c_low.startswith(r_low + " ") or c_low.startswith(r_low + ":") or c_low.startswith(r_low + "：") or c_low.startswith(r_low + ".") or c_low == r_low):
        return c
    if r[-1] in (":", "：", ".", "—", "–", "-"):
        return f"{r} {c}"
    sep = "：" if any("\u4e00" <= ch <= "\u9fff" for ch in r) else ": "
    return f"{r}{sep}{c}"


def caption_height_emu(cap_text: str | None, width_emu: int, cap_pt: float = 11.0) -> int:
    """Compute snug caption strip height (1 or 2 lines) so wrapped captions never overflow."""
    if not cap_text:
        return 0
    w_pt = max(width_emu / 12700.0 - 10.0, 60.0)
    n_ln = min(2, max(1, est_lines(cap_text, cap_pt, w_pt)))
    return int(Inches(0.32 if n_ln == 1 else 0.50))


def add_caption_box(slide, rect, cap_text: str, cap_pt: float = 11.0, color: str | None = None):
    """Add an auto-fitted centered caption textbox (never triggers text_fit overflow)."""
    from pptx.enum.text import PP_ALIGN
    if not cap_text:
        return None
    l, t, w, h = [int(v) for v in rect]
    cap_col = color or ("#D8E2EC" if slide_is_dark(slide) else "#444444")
    tb = fit_textbox(slide, (l, t, w, h), cap_text, size_pt=cap_pt, align=PP_ALIGN.CENTER, color=cap_col)
    w_pt = max(w / 12700.0 - 8.0, 40.0)
    h_pt = max(h / 12700.0, 14.0)
    if tf_text_height_pt(tb.text_frame, w_pt, cap_pt) > h_pt * 0.96:
        shrink_body_to_fit(tb.text_frame, w_pt, h_pt, cap_pt, min_pt=10.0, compact_spacing_pt=0.5)
    return tb


def add_picture_fit(slide, img_path, rect, cap_text=None, cap_pt=11.0):
    """Insert picture aspect-preserved, centered in rect; optional snug caption strip below picture.
    Returns (picture_shape, actual_rect)."""
    from pptx.util import Emu

    l, t, w, h = [int(v) for v in rect]
    cap_box_h = caption_height_emu(cap_text, w, cap_pt)
    cap_h = (cap_box_h + int(Inches(0.06))) if cap_text else 0
    avail_w, avail_h = w, max(h - cap_h, int(Inches(0.5)))
    iw_ih = None
    try:
        from PIL import Image as PILImage

        with PILImage.open(str(img_path)) as im:
            iw_ih = im.size
    except Exception:
        iw_ih = None
    if iw_ih:
        iw, ih = iw_ih
        scale = min(avail_w / iw, avail_h / ih) * 0.97
        pw, ph = int(iw * scale), int(ih * scale)
        pic = slide.shapes.add_picture(
            str(img_path), Emu(l + (w - pw) // 2), Emu(t + (avail_h - ph) // 2),
            width=Emu(pw), height=Emu(ph),
        )
    else:  # PIL missing: width-only keeps aspect via pptx native size
        pic = slide.shapes.add_picture(str(img_path), Emu(l), Emu(t), width=Emu(int(avail_w * 0.97)))
    if cap_text:
        pic_bot = int(pic.top) + int(pic.height)
        cap_y = min(t + avail_h + int(Inches(0.04)), pic_bot + int(Inches(0.08)))
        cap_y = min(cap_y, t + h - cap_box_h)
        add_caption_box(slide, (l, max(t, cap_y), w, cap_box_h), cap_text, cap_pt=cap_pt)
    return pic, (pic.left, pic.top, pic.width, pic.height)


def _is_numeric_cell(txt: str) -> bool:
    import re as _re
    s = (txt or "").strip()
    if not s:
        return False
    if s.upper() in ("N/A", "NA", "—", "–", "-", "--"):
        return True
    return bool(_re.match(
        r"^[+-±≈~<>≤≥$€¥]?\s*\d[\d,\.]*"
        r"(?:\s*(?:[×xX]\s*10[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻\^\-+0-9]+|[eE][+-]?\d+|±\s*\d[\d,\.]*))?"
        r"\s*(%|ms|us|μs|µs|ns|s|Hz|kHz|MHz|GHz|GB|MB|KB|TB|K|mK|dB|x|×|h|min|k|M|B|万|亿|倍)?"
        r"(\s*\([^)]*\))?$",
        s,
    ))


def _parse_hex(color: str | None, default: str = "1F4E79") -> tuple[int, int, int]:
    h = (color or default).lstrip("#")
    if len(h) != 6:
        h = default
    try:
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except Exception:
        return (0x1F, 0x4E, 0x79)


def _rel_luminance(rgb: tuple[int, int, int]) -> float:
    def _lin(c: int) -> float:
        s = c / 255.0
        return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def _safe_accent_for_light_bg(color: str | None) -> tuple[int, int, int]:
    """Ensure accent text/border has strong contrast (>=4.5:1) on light backgrounds."""
    r, g, b = _parse_hex(color)
    lum = _rel_luminance((r, g, b))
    if lum <= 0.30:
        return (r, g, b)
    # Darken proportionally while preserving hue; fall back to deep slate if near-grey
    if max(r, g, b) - min(r, g, b) < 18 and lum > 0.65:
        return (0x1F, 0x4E, 0x79)
    scale = 0.42 if lum > 0.60 else 0.62
    return (max(12, int(r * scale)), max(12, int(g * scale)), max(12, int(b * scale)))


def _tint_rgb(rgb: tuple[int, int, int], alpha: float = 0.12) -> tuple[int, int, int]:
    return tuple(int(255 * (1.0 - alpha) + c * alpha) for c in rgb)


def compute_table_col_widths(rows, total_w: int) -> list[int]:
    """Content-weighted column widths (sqrt-softened) summing to total_w."""
    ncols = max((len(r) for r in rows), default=1)
    col_weights = []
    for ci in range(ncols):
        mx = 3.0
        for row in rows:
            if ci < len(row):
                mx = max(mx, em_width(str(row[ci])))
        col_weights.append(math.sqrt(mx))
    w_sum = sum(col_weights) or 1.0
    col_widths = [max(int(total_w * 0.10), int(total_w * (cw / w_sum))) for cw in col_weights]
    diff = total_w - sum(col_widths)
    if col_widths:
        col_widths[0] = max(int(total_w * 0.10), col_widths[0] + diff)
    return col_widths


def add_table(slide, rect, rows, header=True, font_pt=12.0, accent=None,
              highlight_row=None, max_h=None):
    """Add an academic-styled table with content-weighted column widths, contrast-safe
    theme header, zebra rows, optional emphasized row, and clamped height."""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Pt as _Pt

    l, t, w, h = [int(v) for v in rect]
    nrows, ncols = max(1, len(rows)), max((len(r) for r in rows), default=1)
    col_widths = compute_table_col_widths(rows, w)

    # Estimate required height accurately so frame matches LibreOffice expansion
    est_total_pt = 0.0
    for row in rows:
        max_lines = 1
        for ci in range(ncols):
            txt = str(row[ci]) if ci < len(row) else ""
            cw_pt = max(col_widths[ci] / 12700.0 - 10.0, 24.0)
            max_lines = max(max_lines, est_lines(txt, font_pt, cw_pt))
        est_total_pt += max_lines * font_pt * 1.32 + 8.0
    eff_h = max(h, int(est_total_pt * 12700))
    if max_h is not None:
        eff_h = min(int(max_h), eff_h)

    gf = slide.shapes.add_table(nrows, ncols, Emu(l), Emu(t), Emu(w), Emu(eff_h))
    tbl = gf.table
    for c in range(ncols):
        tbl.columns[c].width = Emu(col_widths[c])

    hdr_t = _parse_hex(accent)
    hdr_rgb = RGBColor(*hdr_t)
    hdr_text_rgb = RGBColor(0x1A, 0x1A, 0x1A) if _rel_luminance(hdr_t) > 0.52 else RGBColor(0xFF, 0xFF, 0xFF)
    acc_safe = _safe_accent_for_light_bg(accent)
    hi_bg_rgb = RGBColor(*_tint_rgb(acc_safe, 0.14))
    zebra_rgb = RGBColor(0xF4, 0xF6, 0xF9)
    white_rgb = RGBColor(0xFF, 0xFF, 0xFF)
    dark_rgb = RGBColor(0x22, 0x22, 0x22)

    hi_idx = None
    if highlight_row is not None:
        try:
            hi_idx = int(highlight_row)
            if hi_idx < 0:
                hi_idx = nrows + hi_idx
            if not (0 <= hi_idx < nrows) or (header and hi_idx == 0):
                hi_idx = None
        except Exception:
            hi_idx = None

    for ri, row in enumerate(rows):
        is_hdr = header and ri == 0
        is_hi = (hi_idx is not None and ri == hi_idx)
        for ci in range(ncols):
            cell = tbl.cell(ri, ci)
            txt = str(row[ci]) if ci < len(row) else ""
            try:
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
                cell.fill.solid()
                if is_hdr:
                    cell.fill.fore_color.rgb = hdr_rgb
                elif is_hi:
                    cell.fill.fore_color.rgb = hi_bg_rgb
                else:
                    cell.fill.fore_color.rgb = zebra_rgb if (ri % 2 == 0) else white_rgb
            except Exception:
                pass
            tf = cell.text_frame
            tf.word_wrap = True
            tf.clear()
            p = tf.paragraphs[0]
            plain_txt = strip_md(txt)
            if is_hdr:
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            elif ci > 0 and _is_numeric_cell(plain_txt):
                p.alignment = PP_ALIGN.CENTER
            cell_color = hdr_text_rgb if is_hdr else (RGBColor(*acc_safe) if is_hi else dark_rgb)
            add_rich_runs(p, txt, base_bold=(is_hdr or is_hi), size_pt=font_pt, color_rgb=cell_color)
    return gf


def add_step_flow(slide, rect, steps, accent=None):
    """Add 2–5 native editable horizontal workflow/pipeline stage cards with vector arrows.
    steps: [{title, desc}] or [str]. Never sets font.name; contrast-safe on light/dark slides."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu

    clean = []
    for idx, st in enumerate((steps or [])[:5], 1):
        if isinstance(st, dict):
            t = strip_md(str(st.get("title") or f"Step {idx}")).strip()
            d = strip_md(str(st.get("desc") or st.get("detail") or "")).strip()
        else:
            t, d = f"Step {idx}", strip_md(str(st)).strip()
        if t or d:
            clean.append((t, d))
    if not clean:
        return []
    l, t, w, h = [int(v) for v in rect]
    n = len(clean)
    arrow_w = int(Inches(0.24)) if n > 1 else 0
    gap = int(Inches(0.06)) if n > 1 else 0
    step_span = arrow_w + 2 * gap
    card_w = max(int(Inches(1.1)), (w - step_span * (n - 1)) // n)
    card_w_pt = max(card_w / 12700.0 - 12.0, 40.0)
    card_h_pt = max(h / 12700.0 - 10.0, 36.0)
    dark = slide_is_dark(slide)
    acc_rgb = RGBColor(*(_safe_accent_for_dark_bg(accent) if dark else _safe_accent_for_light_bg(accent)))
    arrow_rgb = acc_rgb
    bg_rgb = RGBColor(0x1B, 0x2A, 0x44) if dark else RGBColor(0xF5, 0xF7, 0xFA)
    bdr_rgb = RGBColor(0x33, 0x4E, 0x73) if dark else RGBColor(0xD8, 0xE1, 0xEC)
    body_rgb = RGBColor(0xE2, 0xE8, 0xF0) if dark else RGBColor(0x22, 0x22, 0x22)
    stripe_h = int(Inches(0.055))
    created = []
    for i, (st_title, st_desc) in enumerate(clean):
        cl = l + i * (card_w + step_span)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(cl), Emu(t), Emu(card_w), Emu(h))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        card.line.color.rgb = bdr_rgb
        card.line.width = Pt(0.85)
        _flat_shape(card, corner_adj=0.04)
        # Subtle top accent stripe (height 0.055" < 0.08", pure decorative vector shape)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(cl + int(Inches(0.05))), Emu(t),
                                     Emu(max(int(Inches(0.4)), card_w - int(Inches(0.10)))), Emu(stripe_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = acc_rgb
        bar.line.fill.background()
        _flat_shape(bar)
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.10)
        tf.margin_right = Inches(0.10)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.06)
        t_pt = 15.5 if n <= 3 else (14.0 if n == 4 else 12.5)
        d_pt = 13.0 if n <= 3 else (12.0 if n == 4 else 11.0)
        # Auto-step down font size cleanly if text would exceed card height
        while d_pt > 10.0:
            est_h = (est_lines(st_title, t_pt, card_w_pt) * t_pt * 1.25 + 4.0
                     + (est_lines(st_desc, d_pt, card_w_pt) * d_pt * 1.25 if st_desc else 0.0))
            if est_h <= card_h_pt * 0.94:
                break
            d_pt = max(10.0, d_pt - 0.5)
            t_pt = max(10.5, t_pt - 0.5)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.space_before = Pt(0)
        p0.space_after = Pt(3 if st_desc else 0)
        r0 = p0.add_run()
        r0.text = st_title
        r0.font.size = Pt(t_pt)
        r0.font.bold = True
        r0.font.color.rgb = acc_rgb
        if st_desc:
            p1 = tf.add_paragraph()
            p1.alignment = PP_ALIGN.CENTER
            p1.space_before = Pt(0)
            p1.space_after = Pt(0)
            r1 = p1.add_run()
            r1.text = st_desc
            r1.font.size = Pt(d_pt)
            r1.font.color.rgb = body_rgb
        created.append(card)
        if i < n - 1:
            ar_h = int(Inches(0.24))
            ar_l = cl + card_w + gap
            ar_t = t + (h - ar_h) // 2
            arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(ar_l), Emu(ar_t), Emu(arrow_w), Emu(ar_h))
            arr.fill.solid()
            arr.fill.fore_color.rgb = arrow_rgb
            arr.line.fill.background()
            _flat_shape(arr)
            created.append(arr)
    return created


def add_highlights(slide, rect, items, accent=None):
    """Add 2–4 native KPI highlight cards in a horizontal row (contrast-safe, editable)."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu

    items = [it for it in (items or []) if isinstance(it, dict)][:4]
    if not items:
        return []
    l, t, w, h = [int(v) for v in rect]
    gap = int(Inches(0.16))
    n = len(items)
    card_w = (w - gap * (n - 1)) // n
    card_w_pt = card_w / 12700.0
    card_h_pt = h / 12700.0
    dark = slide_is_dark(slide)
    acc_rgb = RGBColor(*(_safe_accent_for_dark_bg(accent) if dark else _safe_accent_for_light_bg(accent)))
    bg_rgb = RGBColor(0x1B, 0x2A, 0x44) if dark else RGBColor(0xF5, 0xF7, 0xFA)
    bdr_rgb = RGBColor(0x33, 0x4E, 0x73) if dark else RGBColor(0xD8, 0xE1, 0xEC)
    lbl_rgb = RGBColor(0xF0, 0xF4, 0xF8) if dark else RGBColor(0x22, 0x22, 0x22)
    sub_rgb = RGBColor(0xCB, 0xD5, 0xE1) if dark else RGBColor(0x55, 0x55, 0x55)
    stripe_h = int(Inches(0.055))
    created = []
    for i, it in enumerate(items):
        cl = l + i * (card_w + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(cl), Emu(t), Emu(card_w), Emu(h))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        card.line.color.rgb = bdr_rgb
        card.line.width = Pt(0.85)
        _flat_shape(card, corner_adj=0.04)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(cl + int(Inches(0.05))), Emu(t),
                                     Emu(max(int(Inches(0.4)), card_w - int(Inches(0.10)))), Emu(stripe_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = acc_rgb
        bar.line.fill.background()
        _flat_shape(bar)
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.08)
        tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.05)
        tf.margin_bottom = Inches(0.05)
        val_txt = strip_md(str(it.get("value", "")))
        val_pt = 20.0 if (n <= 3 and em_width(val_txt) <= 8) else 16.5
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.space_before = Pt(0)
        p0.space_after = Pt(0)
        r_val = p0.add_run()
        r_val.text = val_txt
        r_val.font.size = Pt(val_pt)
        r_val.font.bold = True
        r_val.font.color.rgb = acc_rgb
        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        p1.space_before = Pt(0)
        p1.space_after = Pt(0)
        r_lbl = p1.add_run()
        r_lbl.text = strip_md(str(it.get("label", "")))
        r_lbl.font.size = Pt(11.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = lbl_rgb
        if it.get("sub"):
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(0)
            p2.space_after = Pt(0)
            r_sub = p2.add_run()
            r_sub.text = strip_md(str(it.get("sub")))
            r_sub.font.size = Pt(10.5)
            r_sub.font.color.rgb = sub_rgb
        # Safety check: shrink card fonts proportionally if long label/sub wraps
        if tf_text_height_pt(tf, card_w_pt, 11.5) > card_h_pt * 0.96:
            shrink_body_to_fit(tf, card_w_pt, card_h_pt, val_pt, min_pt=10.5)
        created.append(card)
    return created


def add_takeaway_bar(slide, rect, text, accent=None):
    """Single-line bottom takeaway callout bar (native shape, theme accent, auto-fit)."""
    if not text:
        return None
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu

    l, t, w, h = [int(v) for v in rect]
    dark = slide_is_dark(slide)
    acc_rgb = RGBColor(*(_safe_accent_for_dark_bg(accent) if dark else _safe_accent_for_light_bg(accent)))
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(l), Emu(t), Emu(w), Emu(h))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x1B, 0x2A, 0x44) if dark else RGBColor(0xF0, 0xF4, 0xF8)
    box.line.color.rgb = RGBColor(0x33, 0x4E, 0x73) if dark else RGBColor(0xD8, 0xE1, 0xEC)
    box.line.width = Pt(0.75)
    _flat_shape(box)
    # Crisp vertical left accent ribbon (width 0.065" < 0.08", pure decorative vector shape)
    rib = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(l), Emu(t), Emu(int(Inches(0.065))), Emu(h))
    rib.fill.solid()
    rib.fill.fore_color.rgb = acc_rgb
    rib.line.fill.background()
    _flat_shape(rib)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    add_rich_runs(p, str(text), base_bold=True, size_pt=12.5,
                  color_rgb=(RGBColor(0xF0, 0xF4, 0xF8) if dark else acc_rgb))
    w_pt, h_pt = max(w / 12700.0 - 20.0, 40.0), h / 12700.0
    if tf_text_height_pt(tf, w_pt, 12.5) > h_pt * 0.96:
        shrink_body_to_fit(tf, w_pt, h_pt, 12.5, min_pt=10.5)
    return box


def add_keypoint_cards(slide, rect, items, accent=None):
    """Render 2–4 structured (heading, body) points as crisp stacked horizontal accent cards.
    Eliminates empty-canvas void on pure-text slides while inheriting template fonts."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu

    clean = [(str(h).strip(), str(b).strip()) for h, b in (items or []) if str(h).strip() or str(b).strip()][:4]
    if not clean:
        return []
    l, t, w, h = [int(v) for v in rect]
    n = len(clean)
    gap = int(Inches(0.16))
    max_card_h = int(Inches(1.22 if n <= 3 else 1.02))
    avail_each = (h - gap * (n - 1)) // n
    card_h = max(int(Inches(0.68)), min(max_card_h, avail_each))
    total_h = card_h * n + gap * (n - 1)
    start_t = t + max(0, (h - total_h) // 4)
    dark = slide_is_dark(slide)
    acc_rgb = RGBColor(*(_safe_accent_for_dark_bg(accent) if dark else _safe_accent_for_light_bg(accent)))
    bg_rgb = RGBColor(0x1B, 0x2A, 0x44) if dark else RGBColor(0xF6, 0xF8, 0xFB)
    bdr_rgb = RGBColor(0x33, 0x4E, 0x73) if dark else RGBColor(0xD8, 0xE1, 0xEC)
    body_rgb = RGBColor(0xE2, 0xE8, 0xF0) if dark else RGBColor(0x26, 0x26, 0x26)
    w_pt = max(w / 12700.0 - 26.0, 60.0)
    h_pt = max(card_h / 12700.0 - 10.0, 32.0)
    created = []
    for i, (head, body) in enumerate(clean):
        ct = start_t + i * (card_h + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(l), Emu(ct), Emu(w), Emu(card_h))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        card.line.color.rgb = bdr_rgb
        card.line.width = Pt(0.85)
        _flat_shape(card, corner_adj=0.04)
        rib = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(l), Emu(ct + int(Inches(0.04))),
                                     Emu(int(Inches(0.065))), Emu(max(int(Inches(0.3)), card_h - int(Inches(0.08)))))
        rib.fill.solid()
        rib.fill.fore_color.rgb = acc_rgb
        rib.line.fill.background()
        _flat_shape(rib)
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.22)
        tf.margin_right = Inches(0.16)
        tf.margin_top = Inches(0.06)
        tf.margin_bottom = Inches(0.06)
        h_sz = 15.5 if n <= 3 else 14.0
        b_sz = 14.0 if n <= 3 else 13.0
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        p0.space_before = Pt(0)
        p0.space_after = Pt(2 if body else 0)
        r0 = p0.add_run()
        r0.text = strip_md(head)
        r0.font.size = Pt(h_sz)
        r0.font.bold = True
        r0.font.color.rgb = acc_rgb
        if body:
            p1 = tf.add_paragraph()
            p1.alignment = PP_ALIGN.LEFT
            p1.space_before = Pt(0)
            p1.space_after = Pt(0)
            add_rich_runs(p1, body, base_bold=False, size_pt=b_sz, color_rgb=body_rgb)
        if tf_text_height_pt(tf, w_pt, h_sz) > h_pt * 0.96:
            shrink_body_to_fit(tf, w_pt, h_pt, h_sz, min_pt=10.5, compact_spacing_pt=1.0)
        created.append(card)
    return created


def add_agenda_rows(slide, rect, bullets, accent=None):
    """Render 3–6 agenda/outline items as sleek numbered horizontal strips (01, 02, ...)."""
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu

    clean = [str(b.get("t") if isinstance(b, dict) else b).strip() for b in (bullets or []) if str(b.get("t") if isinstance(b, dict) else b).strip()][:6]
    if not clean:
        return []
    l, t, w, h = [int(v) for v in rect]
    n = len(clean)
    gap = int(Inches(0.12))
    max_row_h = int(Inches(0.84 if n <= 4 else 0.72))
    avail_each = (h - gap * (n - 1)) // n
    row_h = max(int(Inches(0.50)), min(max_row_h, avail_each))
    total_h = row_h * n + gap * (n - 1)
    start_t = t + max(0, (h - total_h) // 4)
    dark = slide_is_dark(slide)
    acc_rgb = RGBColor(*(_safe_accent_for_dark_bg(accent) if dark else _safe_accent_for_light_bg(accent)))
    bg_rgb = RGBColor(0x1B, 0x2A, 0x44) if dark else RGBColor(0xF6, 0xF8, 0xFB)
    bdr_rgb = RGBColor(0x33, 0x4E, 0x73) if dark else RGBColor(0xD8, 0xE1, 0xEC)
    body_rgb = RGBColor(0xF0, 0xF4, 0xF8) if dark else RGBColor(0x22, 0x22, 0x22)
    w_pt = max(w / 12700.0 - 28.0, 60.0)
    h_pt = max(row_h / 12700.0 - 8.0, 24.0)
    sz = 16.0 if n <= 4 else 14.5
    created = []
    for i, txt in enumerate(clean, 1):
        ct = start_t + (i - 1) * (row_h + gap)
        row = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(l), Emu(ct), Emu(w), Emu(row_h))
        row.fill.solid()
        row.fill.fore_color.rgb = bg_rgb
        row.line.color.rgb = bdr_rgb
        row.line.width = Pt(0.8)
        _flat_shape(row, corner_adj=0.04)
        rib = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(l), Emu(ct + int(Inches(0.04))),
                                     Emu(int(Inches(0.06))), Emu(max(int(Inches(0.24)), row_h - int(Inches(0.08)))))
        rib.fill.solid()
        rib.fill.fore_color.rgb = acc_rgb
        rib.line.fill.background()
        _flat_shape(rib)
        tf = row.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.20)
        tf.margin_right = Inches(0.14)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r_num = p.add_run()
        r_num.text = f"{i:02d}   "
        r_num.font.size = Pt(sz + 1.0)
        r_num.font.bold = True
        r_num.font.color.rgb = acc_rgb
        split_done = False
        plain_txt = strip_md(txt)
        for sep in ("：", ": "):
            if sep in plain_txt:
                hd, rest = plain_txt.split(sep, 1)
                if 0 < em_width(hd) <= 16 and rest.strip():
                    r1 = p.add_run()
                    r1.text = hd + sep
                    r1.font.size = Pt(sz)
                    r1.font.bold = True
                    r1.font.color.rgb = body_rgb
                    r2 = p.add_run()
                    r2.text = rest
                    r2.font.size = Pt(sz)
                    r2.font.color.rgb = body_rgb
                    split_done = True
                    break
        if not split_done:
            add_rich_runs(p, txt, base_bold=False, size_pt=sz, color_rgb=body_rgb)
        if tf_text_height_pt(tf, w_pt, sz) > h_pt * 0.96:
            shrink_body_to_fit(tf, w_pt, h_pt, sz, min_pt=10.5)
        created.append(row)
    return created


def set_text_preserving_style(tf, text: str) -> None:
    """Update a text frame's text while preserving the first run's formatting (edit mode)."""
    clean_t = glue_num_units(strip_md(str(text)))
    try:
        if tf.paragraphs and tf.paragraphs[0].runs:
            p0 = tf.paragraphs[0]
            p0.runs[0].text = clean_t
            for r in p0.runs[1:]:
                r.text = ""
            # Remove extra paragraphs if any
            txBody = tf._txBody
            for p in tf.paragraphs[1:]:
                txBody.remove(p._p)
            return
    except Exception:
        pass
    tf.text = clean_t


def find_slide_title_shape(slide, prs=None):
    """Return title placeholder if present; else top-most text shape (edit-mode fallback)."""
    tph = slide_title_ph(slide)
    if tph is not None:
        return tph
    cands = []
    SH = int(prs.slide_height) if prs is not None else int(7.5 * 914400)
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            tp = int(sh.top) if sh.top is not None else SH
            if tp < int(SH * 0.28):
                cands.append((tp, -effective_pt(sh, 18.0), sh))
    cands.sort(key=lambda x: (x[0], x[1]))
    return cands[0][2] if cands else None


def find_slide_body_shapes(slide, prs=None):
    """Return body placeholders if present; else non-title/non-footer/non-source text shapes."""
    bphs = body_placeholders(slide)
    if bphs:
        return bphs
    tsh = find_slide_title_shape(slide, prs)
    t_el = tsh._element if tsh is not None else None
    SH = int(prs.slide_height) if prs is not None else int(7.5 * 914400)
    cands = []
    for sh in slide.shapes:
        if t_el is not None and sh._element is t_el:
            continue
        if sh.is_placeholder and sh.placeholder_format.type in (SKIP_TYPES | TITLE_TYPES):
            continue
        if sh.has_text_frame:
            tp = int(sh.top) if sh.top is not None else 0
            if tp > int(SH * 0.84):
                continue  # likely footer/source strip
            t_cur = (sh.text_frame.text or "").strip()
            if t_cur.startswith(("来源：", "来源:", "Source:", "Source：")):
                continue
            cands.append((int(sh.left or 0), tp, sh))
    cands.sort(key=lambda x: (x[0], x[1]))
    return [c[2] for c in cands]


_PUNCT_NORM = str.maketrans({
    "\u2013": "-", "\u2014": "-", "\u2212": "-",
    "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'", "\u00a0": " ",
})


def _iter_all_shapes(shapes):
    """Yield shapes recursively, descending into GROUP shapes."""
    for sh in shapes:
        yield sh
        try:
            if hasattr(sh, "shapes") and "GROUP" in str(sh.shape_type):
                yield from _iter_all_shapes(sh.shapes)
        except Exception:
            pass


def replace_text_in_slide(slide, old: str, new: str) -> int:
    """In-place substring replacement across text frames, grouped shapes, tables,
    chart titles/axis titles & notes, preserving formatting (with smart-punctuation & whitespace fallbacks)."""
    if not old:
        return 0
    new_clean = glue_num_units(strip_md(new)) if any(ch in new for ch in ("**", "`", "$", "\\", "^", "_")) else glue_num_units(new)
    old_norm = old.translate(_PUNCT_NORM)
    old_ws = re.sub(r"\s+", " ", old_norm).strip()
    count = 0

    def _rep_tf(tf):
        nonlocal count
        for p in tf.paragraphs:
            # 1. Exact per-run replacement (preserves every run's formatting)
            hit_run = False
            for r in p.runs:
                if old in r.text:
                    c = r.text.count(old)
                    r.text = r.text.replace(old, new_clean)
                    count += c
                    hit_run = True
            if hit_run:
                continue
            # 2. Exact cross-run match: collapse into first run preserving style
            full = "".join(r.text for r in p.runs)
            if old in full and p.runs:
                count += full.count(old)
                p.runs[0].text = full.replace(old, new_clean)
                for r in p.runs[1:]:
                    r.text = ""
                continue
            # 3. Smart-punctuation/dash fallback (e.g. en-dash '–' vs '-')
            full_n = full.translate(_PUNCT_NORM)
            if old_norm and old_norm in full_n and p.runs:
                count += full_n.count(old_norm)
                p.runs[0].text = full_n.replace(old_norm, new_clean)
                for r in p.runs[1:]:
                    r.text = ""
                continue
            # 4. Whitespace-normalized fallback (handles \xa0, multi-space, soft line breaks)
            full_ws = re.sub(r"\s+", " ", full_n)
            if old_ws and old_ws in full_ws and p.runs:
                count += full_ws.count(old_ws)
                p.runs[0].text = full_ws.replace(old_ws, new_clean)
                for r in p.runs[1:]:
                    r.text = ""

    for sh in _iter_all_shapes(slide.shapes):
        if sh.has_text_frame:
            _rep_tf(sh.text_frame)
        try:
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        _rep_tf(cell.text_frame)
        except Exception:
            pass
        try:
            if sh.has_chart:
                ch = sh.chart
                if ch.has_title:
                    _rep_tf(ch.chart_title.text_frame)
                for ax_attr in ("category_axis", "value_axis"):
                    try:
                        ax = getattr(ch, ax_attr, None)
                        if ax is not None and ax.has_title:
                            _rep_tf(ax.axis_title.text_frame)
                    except Exception:
                        pass
        except Exception:
            pass
    if slide.has_notes_slide:
        _rep_tf(slide.notes_slide.notes_text_frame)
    return count


def extract_slide_snapshot(slide, prs=None) -> dict:
    """Extract structured text snapshot of a slide (including grouped shapes, table rows & chart data)."""
    tsh = find_slide_title_shape(slide, prs)
    title = tsh.text_frame.text.strip() if (tsh is not None and tsh.has_text_frame) else ""
    bullets = []
    tables = []
    charts = []
    images = 0
    image_sigs = []
    for sh in _iter_all_shapes(slide.shapes):
        if sh is tsh:
            continue
        if sh.is_placeholder and sh.placeholder_format.type in SKIP_TYPES:
            continue
        try:
            if "PICTURE" in str(sh.shape_type):
                images += 1
                try:
                    import hashlib
                    sig = hashlib.sha1(sh.image.blob).hexdigest()[:10]
                    image_sigs.append(f"{sig}:{sh.image.size[0]}x{sh.image.size[1]}")
                except Exception:
                    image_sigs.append(str(sh.name))
                continue
        except Exception:
            pass
        try:
            if sh.has_table:
                rows = [[c.text.strip() for c in row.cells] for row in sh.table.rows]
                tables.append({"rows": len(rows), "preview": rows[:6], "all_rows": rows[:30]})
                continue
        except Exception:
            pass
        try:
            if sh.has_chart:
                ch = sh.chart
                ct = ch.chart_title.text_frame.text.strip() if ch.has_title else ""
                cats, sers = [], []
                try:
                    if ch.plots:
                        cats = [str(c.label) for c in list(ch.plots[0].categories)[:10]]
                    for sr in list(ch.series)[:4]:
                        vals = [round(v, 4) if isinstance(v, float) else v for v in list(sr.values)[:10]]
                        sers.append({"name": str(sr.name), "values": vals})
                except Exception:
                    pass
                charts.append({"title": ct, "categories": cats, "series": sers})
                continue
        except Exception:
            pass
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                t = p.text.strip()
                if t:
                    bullets.append(t)
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
    layout_name = ""
    try:
        layout_name = slide.slide_layout.name or ""
    except Exception:
        pass
    return {"layout": layout_name, "title": title, "bullets": bullets,
            "tables": tables, "charts": charts, "images": images,
            "image_sigs": image_sigs, "notes": notes}


def fix_degenerate_placeholders(prs, slide):
    """Placeholders cloned with explicit zero geometry (some templates) break rendering:
    explicit zeros override master inheritance. Assign sane defaults only when the slide
    placeholder carries explicit but degenerate values; inherited (None) stays untouched."""
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    for ph in slide.placeholders:
        try:
            t = ph.placeholder_format.type
        except Exception:
            continue
        if t in SKIP_TYPES:
            continue
        vals = (ph.left, ph.top, ph.width, ph.height)
        if all(v is None for v in vals):
            continue  # inherited geometry — fine
        if all(v is not None and int(v) > 0 for v in vals):
            continue  # explicit sane geometry — fine
        if t in TITLE_TYPES:
            geom = (0.065, 0.05, 0.87, 0.13)
        elif t in BODY_TYPES:
            geom = (0.065, 0.20, 0.87, 0.68)
        elif t in PIC_TYPES:
            geom = (0.10, 0.20, 0.80, 0.62)
        else:  # subtitle / misc
            geom = (0.065, 0.19, 0.87, 0.12)
        ph.left = Emu(int(SW * geom[0]))
        ph.top = Emu(int(SH * geom[1]))
        ph.width = Emu(int(SW * geom[2]))
        ph.height = Emu(int(SH * geom[3]))


def fill_subtitle(slide, subtitle, prs):
    if not subtitle:
        return
    clean_sub = strip_md(str(subtitle))
    sub_ph = None
    for ph in slide.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
            sub_ph = ph
            break
    if sub_ph is not None:
        sub_ph.text_frame.text = clean_sub
        _, _, sw, sh = ph_geom(prs, slide, sub_ph)
        if sw and sh and sw.pt > 20 and sh.pt > 12:
            inh = ph_inherited_pt(sub_ph, 20.0, getattr(prs, "_master_styles", None))
            shrink_body_to_fit(sub_ph.text_frame, sw.pt, sh.pt, inh, min_pt=12.0)
        return
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    tb = fit_textbox(slide, (int(SW * 0.10), int(SH * 0.30), int(SW * 0.80), int(SH * 0.16)),
                     clean_sub, size_pt=20)
    shrink_body_to_fit(tb.text_frame, tb.width.pt, tb.height.pt, 20.0, min_pt=12.0)


def source_line(slide, rect_bottom, sources, sw, lang="zh"):
    """Small bottom source/citation line built from slide sources."""
    if not sources:
        return None
    parts = []
    for s in sources:
        ref = str(s.get("ref", "")).replace("§", "Sec.")
        loc = s.get("locator", "")
        parts.append(f"{ref} {loc}".strip())
    label = "Source: " if str(lang).lower().startswith("en") else "来源: "
    text = label + "; ".join(dict.fromkeys(parts))
    l, t, w, h = rect_bottom
    col = "#CBD5E1" if slide_is_dark(slide) else "#595959"
    return fit_textbox(slide, (l, t, w, h), text, size_pt=10.5, italic=True, color=col)
