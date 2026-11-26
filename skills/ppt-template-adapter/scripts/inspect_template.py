#!/usr/bin/env python3
"""Inventory a PPTX/POTX template: slide size, masters, layouts, placeholders, theme fonts/colors.

Usage:
  python3 inspect_template.py TEMPLATE.pptx [--json out.json]

The inventory drives layout selection in build_deck.py — never guess layout indices.
No-script fallback snippet (python-pptx):
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation("t.pptx")
    for m in prs.slide_masters:
        for lo in m.slide_layouts:
            print(lo.name, [(p.placeholder_format.idx, str(p.placeholder_format.type),
                             Emu(p.left).inches, Emu(p.top).inches) for p in lo.placeholders])
"""
import argparse, json, os, re, sys, tempfile, zipfile
from pathlib import Path

FOOTERISH = ("DATE", "FOOTER", "SLIDE_NUMBER", "HEADER", "BILLBOARD")
TEXT_BODYISH = ("BODY", "OBJECT", "TEXT", "VERTICAL_BODY", "CONTENT")


def theme_info(pptx_path: Path):
    """Read major/minor latin + East-Asian fonts and scheme colors from the first theme part."""
    fonts, colors = {}, {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml", n))
            if not names:
                return fonts, colors
            xml = z.read(names[0]).decode("utf-8", "ignore")
    except Exception:
        return fonts, colors
    for tag in ("majorFont", "minorFont"):
        block = re.search(rf"<a:{tag}>(.*?)</a:{tag}>", xml, re.S)
        if not block:
            continue
        inner = block.group(1)
        m_lat = re.search(r'<a:latin\s+typeface="([^"]*)"', inner)
        if m_lat and m_lat.group(1):
            fonts[tag] = m_lat.group(1)
        m_ea = re.search(r'<a:ea\s+typeface="([^"]*)"', inner)
        if m_ea and m_ea.group(1):
            fonts[f"{tag}_ea"] = m_ea.group(1)
    for cname in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        m = re.search(rf"<a:{cname}>(.*?)</a:{cname}>", xml, re.S)
        if not m:
            continue
        rgb = re.search(r'(?:lastClr|val)="([0-9A-Fa-f]{6})"', m.group(1))
        if rgb:
            colors[cname] = "#" + rgb.group(1)
    return fonts, colors


def _layout_clr_map(lo):
    cmap = {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}
    try:
        from pptx.oxml.ns import qn
        cm = lo.slide_master._element.find(qn("p:clrMap"))
        if cm is not None:
            for k in ("bg1", "bg2", "tx1", "tx2"):
                if cm.get(k):
                    cmap[k] = cm.get(k)
        ovr = lo._element.find(qn("p:clrMapOvr"))
        if ovr is not None:
            ocm = ovr.find(qn("a:overrideClrMapping"))
            if ocm is not None:
                for k in ("bg1", "bg2", "tx1", "tx2"):
                    if ocm.get(k):
                        cmap[k] = ocm.get(k)
    except Exception:
        pass
    return cmap


def _fill_hex_from_el(cont, theme_colors, alias_map, qn):
    if cont is None:
        return None

    def _node_hex(node):
        if node is None:
            return None
        srgb = node.find(qn("a:srgbClr"))
        if srgb is not None and srgb.get("val"):
            return "#" + srgb.get("val").lstrip("#")
        sch = node.find(qn("a:schemeClr"))
        if sch is not None and sch.get("val") and theme_colors:
            k = sch.get("val")
            return theme_colors.get(alias_map.get(k, k)) or theme_colors.get(k)
        return None

    hx = _node_hex(cont.find(qn("a:solidFill")))
    if hx and len(hx) == 7:
        return hx
    gf = cont.find(qn("a:gradFill"))
    if gf is not None:
        gs_lst = gf.find(qn("a:gsLst"))
        if gs_lst is not None:
            stops = []
            for gs in gs_lst.findall(qn("a:gs")):
                shx = _node_hex(gs)
                if shx and len(shx) == 7:
                    stops.append((int(shx[1:3], 16), int(shx[3:5], 16), int(shx[5:7], 16)))
            if stops:
                r = sum(c[0] for c in stops) // len(stops)
                g = sum(c[1] for c in stops) // len(stops)
                b = sum(c[2] for c in stops) // len(stops)
                return f"#{r:02X}{g:02X}{b:02X}"
    return _node_hex(cont)


def _bg_hex(obj, theme_colors=None, clr_map=None, slide_w=None, slide_h=None):
    """Read solid/gradient background or full-slide backdrop hex from XML without mutating XML."""
    try:
        from pptx.oxml.ns import qn
        csld = obj._element.find(qn("p:cSld"))
        if csld is None:
            return None
        alias_map = clr_map or {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}
        bg = csld.find(qn("p:bg"))
        if bg is not None:
            for container_tag in ("p:bgPr", "p:bgRef"):
                hx = _fill_hex_from_el(bg.find(qn(container_tag)), theme_colors, alias_map, qn)
                if hx and len(hx) == 7:
                    return hx
        if slide_w and slide_h:
            slide_area = max(slide_w * slide_h, 1.0)
            for sh in getattr(obj, "shapes", ()):
                if getattr(sh, "is_placeholder", False) or getattr(sh, "has_table", False):
                    continue
                if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                    continue
                bw, bh = inches(sh.width) or 0.0, inches(sh.height) or 0.0
                if bw * bh >= 0.75 * slide_area:
                    hx = _fill_hex_from_el(sh._element.find(qn("p:spPr")), theme_colors, alias_map, qn)
                    if hx and len(hx) == 7:
                        return hx
    except Exception:
        pass
    return None


def _is_dark_hex(hex_str):
    if not hex_str or len(hex_str) != 7:
        return False
    r, g, b = int(hex_str[1:3], 16), int(hex_str[3:5], 16), int(hex_str[5:7], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 128


def inches(v):
    try:
        return round(v / 914400, 2)
    except TypeError:
        return None


def open_prs_any(path: Path):
    """Open .pptx or .potx cleanly; returns (Presentation, tmp_path_or_None)."""
    from pptx import Presentation
    try:
        return Presentation(str(path)), None
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
        return Presentation(tmp_name), tmp_name


def _edge_decor_margins(container, sw, sh):
    """Detect non-placeholder static decorative bands/logos along template edges (top/bottom/left/right)."""
    m_top, m_bot, m_left, m_right = 0.0, 0.0, 0.0, 0.0
    if not sw or not sh:
        return dict(top=0.0, bottom=0.0, left=0.0, right=0.0)
    for sp in container.shapes:
        if sp.is_placeholder:
            continue
        l, t = inches(sp.left), inches(sp.top)
        w, h = inches(sp.width), inches(sp.height)
        if None in (l, t, w, h) or w <= 0.05 or h <= 0.05:
            continue
        # Ignore full-slide background shapes
        if w >= sw * 0.85 and h >= sh * 0.85:
            continue
        r, b = l + w, t + h
        # Top banner/bar (occupies upper 22% strip, thin height)
        if t < sh * 0.16 and h <= sh * 0.18 and b <= sh * 0.24 and w >= sw * 0.25:
            m_top = max(m_top, round(b, 2))
        # Bottom footer banner/logo (occupies lower 18% strip)
        if b > sh * 0.84 and h <= sh * 0.18 and t >= sh * 0.78 and w >= sw * 0.12:
            m_bot = max(m_bot, round(sh - t, 2))
        # Left sidebar accent
        if l < sw * 0.12 and w <= sw * 0.20 and r <= sw * 0.24 and h >= sh * 0.45:
            m_left = max(m_left, round(r, 2))
        # Right sidebar accent
        if r > sw * 0.88 and w <= sw * 0.20 and l >= sw * 0.76 and h >= sh * 0.45:
            m_right = max(m_right, round(sw - l, 2))
    return dict(top=m_top, bottom=m_bot, left=m_left, right=m_right)


def inspect(path: Path):
    prs, tmp_name = open_prs_any(path)
    try:
        t_fonts, t_colors = theme_info(path)
        sw, sh = inches(prs.slide_width), inches(prs.slide_height)
        masters = []
        for master in prs.slide_masters:
            m_bg = _bg_hex(master, t_colors, slide_w=sw, slide_h=sh)
            m_decor = _edge_decor_margins(master, sw, sh)
            layouts = []
            for lo in master.slide_layouts:
                cmap = _layout_clr_map(lo)
                lo_bg = (
                    _bg_hex(lo, t_colors, cmap, slide_w=sw, slide_h=sh)
                    or _bg_hex(master, t_colors, cmap, slide_w=sw, slide_h=sh)
                    or t_colors.get(cmap.get("bg1", "lt1"))
                )
                lo_decor = _edge_decor_margins(lo, sw, sh)
                decor = {k: max(m_decor[k], lo_decor[k]) for k in ("top", "bottom", "left", "right")}
                phs = []
                n_text_body, n_pic = 0, 0
                has_title = has_sub = has_footer = has_snum = False
                for p in lo.placeholders:
                    pf = p.placeholder_format
                    pt = str(pf.type).split(" ")[0]
                    phs.append(dict(idx=pf.idx, type=pt, name=p.name,
                                    left=inches(p.left), top=inches(p.top),
                                    width=inches(p.width), height=inches(p.height)))
                    if pt in ("TITLE", "CENTER_TITLE"):
                        has_title = True
                    elif pt == "SUBTITLE":
                        has_sub = True
                    elif pt in ("PICTURE", "CLIP_ART"):
                        n_pic += 1
                    elif pt in TEXT_BODYISH:
                        n_text_body += 1
                    elif pt == "FOOTER":
                        has_footer = True
                    elif pt == "SLIDE_NUMBER":
                        has_snum = True
                layouts.append(dict(
                    name=lo.name,
                    placeholders=phs,
                    has_title=has_title,
                    has_subtitle=has_sub,
                    text_body_count=n_text_body,
                    picture_count=n_pic,
                    has_footer=has_footer,
                    has_slide_number=has_snum,
                    bg_color=lo_bg,
                    dark_bg=_is_dark_hex(lo_bg),
                    decor_margins=decor,
                ))
            masters.append(dict(bg_color=m_bg, decor_margins=m_decor, layouts=layouts))
        inv = dict(
            template=str(path),
            is_potx=tmp_name is not None,
            existing_slides=len(prs.slides),
            slide_width_in=sw,
            slide_height_in=sh,
            n_layouts_total=sum(len(m["layouts"]) for m in masters),
            masters=masters,
        )
        inv["theme_fonts"], inv["theme_colors"] = t_fonts, t_colors
        return inv
    finally:
        if tmp_name and os.path.exists(tmp_name):
            try:
                os.remove(tmp_name)
            except OSError:
                pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("--json", default=None, help="write inventory JSON here (default: stdout)")
    a = ap.parse_args()
    p = Path(a.template)
    if not p.exists():
        print(f"FATAL: {p} not found", file=sys.stderr)
        return 2
    try:
        inv = inspect(p)
    except Exception as e:
        print(f"FATAL: cannot parse template: {e}", file=sys.stderr)
        return 2
    js = json.dumps(inv, ensure_ascii=False, indent=1)
    if a.json:
        Path(a.json).parent.mkdir(parents=True, exist_ok=True)
        Path(a.json).write_text(js, encoding="utf-8")
        print(f"OK {inv['n_layouts_total']} layouts, {inv['slide_width_in']}x{inv['slide_height_in']}in "
              f"(existing_slides={inv['existing_slides']}) -> {a.json}")
        print("theme fonts:", inv["theme_fonts"], "| accents:", {k: v for k, v in inv["theme_colors"].items() if k.startswith("accent")})
    else:
        print(js)
    return 0


if __name__ == "__main__":
    sys.exit(main())
