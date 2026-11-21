#!/usr/bin/env python3
"""deckutil.py — shared helpers for the academic-slides skill scripts.

Pure python-pptx based helpers: text width metrics, layout picking,
placeholder filling, geometry utilities, change-log. No network, no caches.
"""
from __future__ import annotations

import json
import math
import sys
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu, Inches, Pt

SCRIPT_DIR = Path(__file__).resolve().parent

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
    avail = shape_w_pt - ml - mr
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
        lines = est_lines(text, size, avail) if text else 1
        total += lines * size * LINE_FACTOR + size * PARA_GAP_EM
    return total


def load_master_styles(prs, template_path=None):
    """Parse master txStyles (title/body lvl1 sizes, pt) into prs._master_styles."""
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
    except Exception:
        pass
    try:
        prs._master_styles = styles
    except Exception:
        pass
    return styles


def ph_inherited_pt(slide_ph, fallback: float, master_styles=None) -> float:
    """Best-effort inherited font size: slide ph XML sz, then master txStyles, then fallback."""
    import re
    try:
        xml = slide_ph._element.xml
        m = re.search(r'sz="(\d{3,5})"', xml)
        if m:
            val = int(m.group(1)) / 100.0
            if 900 <= val <= 9900:
                return val
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
    top = int(SH * 0.16)
    if tph is not None:
        _, tph_t, _, tph_h = ph_geom(prs, slide, tph)
        if ph_ok(tph_h):
            top = int(tph_t or 0) + int(tph_h) + Inches(0.12)
    bottom = int(SH * 0.93)
    left = int(SW * 0.065)
    right = int(SW * 0.935)
    return left, top, max(Inches(0.5), right - left), max(Inches(0.4), bottom - top)


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


def theme_accent_color(prs, template_path=None):
    """Read accent1 srgbClr from the template theme; return '#RRGGBB' or None."""
    path = template_path
    try:
        if not path:
            return None
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(path)) as z:
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


def accent_bar(slide, prs, color, left=None, top=None, width_in=1.4, height_in=0.035):
    """Thin accent underline below the title. Pure shape (no text), theme-colored."""
    if not color:
        return None
    try:
        from pptx.dml.color import RGBColor
        from pptx.util import Emu
        tph = slide_title_ph(slide)
        SW = int(prs.slide_width)
        if left is None:
            left = int(SW * 0.065)
            if tph is not None:
                _, tp, _, _ = ph_geom(prs, slide, tph)
                if ph_ok(tp):
                    left = int(tp or left)
        if top is None:
            if tph is not None:
                _, tp, tw, th = ph_geom(prs, slide, tph)
                if ph_ok(th) and ph_ok(tp):
                    top = int(tp) + int(th) + Inches(0.03)
                    # hug the actual text: use estimated text height when title exists
                    if tph.has_text_frame and tph.text_frame.text.strip():
                        size = effective_pt(tph, DEFAULT_TITLE_PT)
                        est_emu = int(tf_text_height_pt(tph.text_frame, tw.pt if tw else 0, size) * 12700)
                        top = int(tp) + int(min(est_emu + Inches(0.05), max(int(th) - Inches(0.03), Inches(0.12))))
                else:
                    top = int(prs.slide_height * 0.185)
            else:
                top = int(prs.slide_height * 0.185)
        from pptx.enum.shapes import MSO_SHAPE
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Emu(int(left)), Emu(int(top)),
                                      Emu(int(Inches(width_in))), Emu(int(Inches(height_in))))
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
        rect.line.fill.background()
        rect.shadow.inherit = False
        return rect
    except Exception:
        return None


def fit_textbox(slide, rect, text, size_pt=None, bold=False, italic=False, align=None):
    """Add a plain textbox inside rect. No font-name override (theme font inherited)."""
    from pptx.util import Emu

    l, t, w, h = rect
    tb = slide.shapes.add_textbox(Emu(int(l)), Emu(int(t)), Emu(int(w)), Emu(int(h)))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = str(text)
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
    return tb


def fill_bullets_tf(tf, bullets, default_pt=None):
    """Fill a text frame with planned bullets (replaces existing content).
    Inherits template fonts unless bullet sets size."""
    try:
        tf.clear()  # drop pre-existing paragraphs (edit mode must not merge with old text)
    except Exception:
        pass
    tf.word_wrap = True
    first = True
    for b in bullets:
        if isinstance(b, dict):
            text = str(b.get("t", ""))
            lvl = int(b.get("lvl", 0) or 0)
            bold = bool(b.get("bold", False))
            size = b.get("size")
        else:
            text, lvl, bold, size = str(b), 0, False, None
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        try:
            p.level = max(0, min(4, lvl))
        except Exception:
            pass
        run = p.add_run()
        run.text = text
        if bold:
            run.font.bold = True
        if size:
            run.font.size = Pt(float(size))
    return tf


def shrink_body_to_fit(tf, w_pt, h_pt, body_pt, min_pt=12.0):
    """Step down body font size until estimated height fits, or floor reached.
    w_pt/h_pt are the resolved shape size in points. Returns final pt used."""
    if w_pt <= 0 or h_pt <= 0:
        return body_pt
    size = body_pt
    while size > min_pt:
        est = tf_text_height_pt(tf, w_pt, size)
        if est <= h_pt * 0.98:
            break
        size = max(min_pt, size - 1.0)
    # apply explicit size when we had to shrink below inherited size
    if size < body_pt - 0.5:
        for p in tf.paragraphs:
            p.font.size = Pt(size)
            for r in p.runs:
                if r.font.size is None or r.font.size.pt > size:
                    r.font.size = Pt(size)
    return size


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


def add_native_chart(slide, rect, chart_spec):
    """Add a fully-editable native PPT chart bound to plan data (创新点：图表数据绑定).
    chart_spec: {type: bar|line|pie, categories: [...], series: [{name, values}],
                 ref, caption}. Returns the graphic frame."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Emu

    ctype = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS,
             "pie": XL_CHART_TYPE.PIE}.get(str(chart_spec.get("type", "bar")).lower(),
                                           XL_CHART_TYPE.COLUMN_CLUSTERED)
    cats = [str(c) for c in (chart_spec.get("categories") or [])]
    cd = CategoryChartData()
    cd.categories = cats
    for s in chart_spec.get("series") or []:
        cd.add_series(str(s.get("name", "series")),
                      tuple(float(v) for v in (s.get("values") or [])))
    l, t, w, h = [int(v) for v in rect]
    gf = slide.shapes.add_chart(ctype, Emu(l), Emu(t), Emu(w), Emu(h), cd)
    try:
        chart = gf.chart
        chart.has_legend = len(chart_spec.get("series") or []) > 1
        if chart_spec.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = str(chart_spec["title"])
        else:
            chart.has_title = False
    except Exception:
        pass
    return gf


def add_picture_fit(slide, img_path, rect, cap_text=None, cap_pt=11.0):
    """Insert picture aspect-preserved, centered in rect; optional caption strip at rect bottom.
    Returns (picture_shape, actual_rect)."""
    from pptx.util import Emu

    l, t, w, h = [int(v) for v in rect]
    cap_h = int(Inches(0.32)) if cap_text else 0
    avail_w, avail_h = w, max(h - cap_h, Inches(0.5))
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
        fit_textbox(
            slide,
            (l, t + avail_h + Inches(0.02), w, Inches(0.3)),
            cap_text, size_pt=cap_pt,
        )
    return pic, (pic.left, pic.top, pic.width, pic.height)


def add_table(slide, rect, rows, header=True, font_pt=12.0):
    """Add a table with equal column widths; header row bold. rows = list[list[str]]."""
    from pptx.util import Emu, Pt as _Pt

    l, t, w, h = [int(v) for v in rect]
    nrows, ncols = len(rows), max(len(r) for r in rows)
    nrows = max(1, nrows)
    ncols = max(1, ncols)
    gf = slide.shapes.add_table(nrows, ncols, Emu(l), Emu(t), Emu(w), Emu(h))
    tbl = gf.table
    col_w = int(w / ncols)
    for c in range(ncols):
        tbl.columns[c].width = Emu(col_w)
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = tbl.cell(ri, ci)
            txt = str(row[ci]) if ci < len(row) else ""
            tf = cell.text_frame
            tf.word_wrap = True
            tf.text = txt
            p = tf.paragraphs[0]
            size = _Pt(font_pt)
            for r in p.runs:
                r.font.size = size
                if header and ri == 0:
                    r.font.bold = True
    return gf


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
    sub_ph = None
    for ph in slide.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
            sub_ph = ph
            break
    if sub_ph is not None:
        sub_ph.text_frame.text = str(subtitle)
        return
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    fit_textbox(slide, (int(SW * 0.10), int(SH * 0.30), int(SW * 0.80), int(SH * 0.14)),
                subtitle, size_pt=20)


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
    return fit_textbox(slide, (l, t, w, h), text, size_pt=10.5, italic=True)
