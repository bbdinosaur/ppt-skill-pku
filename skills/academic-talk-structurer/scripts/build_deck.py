#!/usr/bin/env python3
"""S3 deck builder: deck_spec.json -> editable PPTX on top of provided template/starting deck.

Usage:
  python3 build_deck.py --spec deck_spec.json --out /app/output/talk.pptx \
      [--template T.pptx|.potx] [--deck starting.pptx] [--work /app/work]

Spec schema (see SKILL.md S2):
  meta{title,duration_minutes,claim_on_slide,evidence_footer,time_annotations,
       drop_existing_slides,keep_only_spec_slides,goal,audience,tradeoffs},
  slides[{role,layout_hint,title,claim,bullets[str|{text,level}],
          steps[{label,title,detail}],kpis[{value,label,delta}],
          two_col{left_title,left_bullets,right_title,right_bullets},
          figure{path,region,caption},table{headers,rows,region,caption,highlight_row},
          evidence[{type,locator,strength}],notes,time_s,
          reuse{slide_index,action,clear_visuals}}].
No-script fallback: python-pptx snippet in SKILL.md S3.
Exit: 0 ok, 3 python-pptx missing, 4 spec invalid.
"""
import argparse
import colorsys
import json
import math
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

EMU_IN = 914400.0
DEFAULT_ACCENT = (0x1F, 0x4E, 0x79)
CJK_FONT = "Noto Sans CJK SC"


def char_units(s):
    u = 0.0
    for c in s:
        u += 1.0 if "\u2e80" <= c <= "\uffef" or "\u3000" <= c <= "\u303f" else 0.55
    return u


def speaking_units(s):
    """Language-aware speaking length (1 CJK char ~= 1 unit; 1 Latin word ~= 1.5 units)."""
    if not s:
        return 0.0
    cjk = sum(1 for c in s if "\u2e80" <= c <= "\uffef")
    latin_words = len(re.findall(r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", s))
    return float(cjk) + latin_words * 1.5


def is_cjk(s):
    return any("\u4e00" <= c <= "\uffef" for c in (s or ""))


_SUB_MAP = str.maketrans("0123456789+-=()aeoxhklmnpstij", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜᵢⱼ")
_SUP_MAP = str.maketrans("0123456789+-=()nidkmtrT", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵈᵏᵐᵗʳᵀ")


_LATEX_SYM_MAP = {
    "pm": "±", "mp": "∓", "times": "×", "cdot": "·", "leq": "≤", "le": "≤",
    "geq": "≥", "ge": "≥", "approx": "≈", "neq": "≠", "ne": "≠", "sim": "~",
    "equiv": "≡", "propto": "∝", "perp": "⊥", "parallel": "∥", "star": "★", "dagger": "†",
    "nabla": "∇", "partial": "∂", "in": "∈", "notin": "∉", "subset": "⊂", "subseteq": "⊆",
    "forall": "∀", "exists": "∃", "ell": "ℓ", "dots": "…", "ldots": "…", "cdots": "⋯",
    "rightarrow": "→", "to": "→", "leftarrow": "←", "leftrightarrow": "↔", "infty": "∞",
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "varepsilon": "ε",
    "eta": "η", "theta": "θ", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "phi": "φ", "psi": "ψ", "omega": "ω",
    "Delta": "Δ", "Sigma": "Σ", "Omega": "Ω", "Lambda": "Λ", "Gamma": "Γ", "Theta": "Θ",
}
_LATEX_SYM_RE = re.compile(r"\\(" + "|".join(sorted(_LATEX_SYM_MAP.keys(), key=len, reverse=True)) + r")(?![A-Za-z])")
_NBSP_UNIT_RE = re.compile(r"(\d)\s+(pt|pp|ms|μs|us|ns|GB|MB|KB|TB|Hz|kHz|MHz|GHz|×)(?![A-Za-z0-9_])")
_HAT_MAP = {"y": "ŷ", "x": "x̂", "w": "ŵ", "h": "ĥ", "p": "p̂", "z": "ẑ", "Y": "Ŷ"}


def polish_math_text(s):
    """Convert unambiguous inline math/LaTeX symbols (\\pm, \\alpha, x_t, ℝ^d, 10^{-4}, >=) to crisp Unicode & prevent orphan unit wraps."""
    if not s:
        return s
    out = s
    if "\\" in out:
        out = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"\1/\2", out)
        out = re.sub(r"\\(?:text|mathrm|mathbf|mathit|operatorname|boldsymbol)\{([^{}]+)\}", r"\1", out)
        out = re.sub(r"\\hat\{([A-Za-z])\}", lambda m: _HAT_MAP.get(m.group(1), m.group(1) + "\u0302"), out)
        out = re.sub(r"\\%", "%", out)
        out = re.sub(r"\\mathbb\{R\}", "ℝ", out)
        out = re.sub(r"\\mathbb\{E\}", "𝔼", out)
        out = re.sub(r"\\mathbb\{N\}", "ℕ", out)
        out = re.sub(r"\\mathbb\{Z\}", "ℤ", out)
        out = re.sub(r"\\mathcal\{([A-Z])\}", r"\1", out)
        out = _LATEX_SYM_RE.sub(lambda m: _LATEX_SYM_MAP.get(m.group(1), m.group(0)), out)
    # Clean ASCII math comparison/arrow shorthands outside URLs/paths
    if not re.search(r"https?://", out):
        out = re.sub(r"(?<![<>=!])>=(?!=)", "≥", out)
        out = re.sub(r"(?<![<>=!])<=(?!=)", "≤", out)
        out = re.sub(r"(?<![<>=!])!=(?!=)", "≠", out)
        out = re.sub(r"(?<![A-Za-z0-9/])\+/-(?=\s*\d|\s|$)", "±", out)
        out = re.sub(r"(?<=\s)->(?=\s)", "→", out)
    # Avoid touching URLs, file paths, or multi-char snake_case identifiers
    if "_" in out or "^" in out:
        def _sub_repl(m):
            base, sub = m.group(1), m.group(2)
            return base + sub.translate(_SUB_MAP)

        def _sup_repl(m):
            base, sup = m.group(1), m.group(2)
            return base + sup.translate(_SUP_MAP)

        out = re.sub(r"(?<![A-Za-z0-9_/.])([α-ωΑ-ΩλσμθβγδA-Za-z])_\{?([0-9ijlmnkt])\}?(?![A-Za-z0-9_/.])", _sub_repl, out)
        out = re.sub(r"(?<![A-Za-z_/.])([A-Za-z0-9α-ωΑ-Ωℝ𝔼ℕℤ)])\^\{([+-]?[0-9nidkmtrT]{1,3})\}(?![A-Za-z0-9_])", _sup_repl, out)
        out = re.sub(r"(?<![A-Za-z_/.])([A-Za-z0-9α-ωΑ-Ωℝ𝔼ℕℤ)])\^([+-]?[0-9nidkmtrT]{1,3})(?![A-Za-z0-9_])", _sup_repl, out)
    # Bind numbers to short units with non-breaking space so PowerPoint/LibreOffice never orphan 'pt'/'GB'/'ms' on a new line
    out = _NBSP_UNIT_RE.sub(lambda m: f"{m.group(1)}\u00A0{m.group(2)}", out)
    out = re.sub(r"(\d)\s+%", r"\1%", out)
    return out


def blend_with_white(rgb, white_ratio=0.90):
    return tuple(int(round(ch * (1.0 - white_ratio) + 255 * white_ratio)) for ch in rgb)


def blend_with_black(rgb, dark_ratio=0.65):
    return tuple(int(round(ch * (1.0 - dark_ratio))) for ch in rgb)


def luminance(rgb):
    r, g, b = rgb
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def theme_palette(pptx_path):
    """Parse dk1/lt1/dk2/lt2/accent1..6 RGB tuples from ppt/theme/theme1.xml."""
    pal = {}
    if not pptx_path or not Path(pptx_path).exists():
        return pal
    try:
        import zipfile
        with zipfile.ZipFile(str(pptx_path)) as z:
            xml = z.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
        for key in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            m = re.search(rf'<a:{key}>.*?(?:srgbClr val|lastClr)="([0-9A-Fa-f]{{6}})"', xml, re.S)
            if m:
                v = m.group(1)
                pal[key] = tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        pass
    return pal


def _xml_local(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) and "}" in el.tag else str(el.tag)


def _resolve_color_el(color_el, scheme_map):
    """Resolve an OOXML color element (srgbClr/schemeClr/sysClr) + shade/tint/lumMod/lumOff/alpha modifiers."""
    if color_el is None:
        return None, 1.0
    tag = _xml_local(color_el)
    base = None
    if tag == "srgbClr":
        val = color_el.get("val", "")
        if re.fullmatch(r"[0-9A-Fa-f]{6}", val):
            base = tuple(float(int(val[i:i + 2], 16)) for i in (0, 2, 4))
    elif tag == "schemeClr":
        val = color_el.get("val", "")
        if val in scheme_map:
            base = tuple(float(c) for c in scheme_map[val])
    elif tag == "sysClr":
        val = color_el.get("lastClr", "")
        if re.fullmatch(r"[0-9A-Fa-f]{6}", val):
            base = tuple(float(int(val[i:i + 2], 16)) for i in (0, 2, 4))
    if base is None:
        return None, 1.0
    r, g, b = base
    alpha = 1.0
    for mod in color_el:
        mt = _xml_local(mod)
        mv = mod.get("val")
        if mv is None:
            continue
        try:
            f = int(mv) / 100000.0
        except (TypeError, ValueError):
            continue
        if mt == "shade":
            r, g, b = r * f, g * f, b * f
        elif mt == "tint":
            r = r * f + 255.0 * (1.0 - f)
            g = g * f + 255.0 * (1.0 - f)
            b = b * f + 255.0 * (1.0 - f)
        elif mt in ("lumMod", "lumOff"):
            h, l, s = colorsys.rgb_to_hls(
                max(0.0, min(1.0, r / 255.0)),
                max(0.0, min(1.0, g / 255.0)),
                max(0.0, min(1.0, b / 255.0)),
            )
            l = max(0.0, min(1.0, l * f)) if mt == "lumMod" else max(0.0, min(1.0, l + f))
            rr, gg, bb = colorsys.hls_to_rgb(h, l, s)
            r, g, b = rr * 255.0, gg * 255.0, bb * 255.0
        elif mt == "alpha":
            alpha = max(0.0, min(1.0, f))
    rgb = (
        int(round(max(0.0, min(255.0, r)))),
        int(round(max(0.0, min(255.0, g)))),
        int(round(max(0.0, min(255.0, b)))),
    )
    return rgb, alpha


def _image_blob_is_dark(blob, bright_offset=0.0):
    """Sample perceptual mean luminance of an image blob via PIL (True if dark < 95, False if light, None on error)."""
    if not blob:
        return None
    try:
        from io import BytesIO
        from PIL import Image as _PILImage
        with _PILImage.open(BytesIO(blob)) as im:
            if im.mode in ("RGBA", "LA") or ("transparency" in im.info):
                rgba = im.convert("RGBA").resize((48, 36))
                a_vals = list(rgba.getdata(3))
                if a_vals and (sum(1 for a in a_vals if a < 64) / float(len(a_vals))) > 0.65:
                    return None
                white_bg = _PILImage.new("RGBA", rgba.size, (255, 255, 255, 255))
                rgb_im = _PILImage.alpha_composite(white_bg, rgba).convert("RGB")
            else:
                rgb_im = im.convert("RGB").resize((48, 36))
            w, h = rgb_im.size
            crop = rgb_im.crop((int(w * 0.10), int(h * 0.10), max(int(w * 0.90), 1), max(int(h * 0.90), 1)))
            pxs = list(crop.getdata())
            if not pxs:
                return None
            mean_lum = sum(0.299 * p[0] + 0.587 * p[1] + 0.114 * p[2] for p in pxs) / float(len(pxs))
            if bright_offset:
                mean_lum = max(0.0, min(255.0, mean_lum + bright_offset * 255.0))
            return mean_lum < 95.0
    except Exception:
        return None


def _blip_el_is_dark(blip_fill_el, part):
    """Inspect <a:blipFill> or <p:blipFill> image luminance via related part blob."""
    if blip_fill_el is None or part is None:
        return None
    blip_el = None
    for el in blip_fill_el.iter():
        if _xml_local(el) == "blip":
            blip_el = el
            break
    if blip_el is None:
        return None
    rid = None
    for ak, av in blip_el.attrib.items():
        if ak.endswith("}embed") or ak == "embed":
            rid = av
            break
    if not rid:
        return None
    bright_off = 0.0
    for ch in blip_el:
        if _xml_local(ch) == "lum" and ch.get("bright"):
            try:
                bright_off = int(ch.get("bright")) / 100000.0
            except (TypeError, ValueError):
                pass
    try:
        rel_part = part.related_part(rid)
        return _image_blob_is_dark(rel_part.blob, bright_offset=bright_off)
    except Exception:
        return None


def _fill_container_is_dark(container_el, part, scheme_map, treat_nofill_light=False):
    """Evaluate solidFill / gradFill (mean stop luminance) / blipFill / noFill inside <p:bgPr> or <p:spPr>."""
    if container_el is None:
        return None
    for ch in container_el:
        tag = _xml_local(ch)
        if tag == "noFill":
            return False if treat_nofill_light else None
        if tag == "solidFill":
            for c_el in ch:
                if _xml_local(c_el) in ("srgbClr", "schemeClr", "sysClr"):
                    rgb, alpha = _resolve_color_el(c_el, scheme_map)
                    if rgb is None or alpha < 0.25:
                        return None
                    return luminance(rgb) < 95
        if tag == "gradFill":
            lums, alphas = [], []
            for el in ch.iter():
                if _xml_local(el) == "gs":
                    for c_el in el:
                        if _xml_local(c_el) in ("srgbClr", "schemeClr", "sysClr"):
                            rgb, alpha = _resolve_color_el(c_el, scheme_map)
                            if rgb is not None:
                                lums.append(luminance(rgb))
                                alphas.append(alpha)
                            break
            if lums and (sum(alphas) / float(len(alphas))) >= 0.25:
                return (sum(lums) / float(len(lums))) < 95.0
        if tag == "blipFill":
            res = _blip_el_is_dark(ch, part)
            if res is not None:
                return res
    return None


def _bg_xml_is_dark(el_xml, scheme_map, part=None):
    """Return True (dark), False (light/noFill), or None (inherit from parent) for one <p:bg> block."""
    if not el_xml or "<p:bg" not in el_xml:
        return None
    try:
        root = ET.fromstring(el_xml)
    except Exception:
        m_bg = re.search(r"<p:bg\b.*?</p:bg>", el_xml, re.S)
        if not m_bg:
            return None
        blk = m_bg.group(0)
        if "<a:noFill" in blk:
            return False
        ms = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', blk)
        if ms:
            v = ms.group(1)
            return luminance(tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))) < 95
        return None
    bg_el = None
    for el in root.iter():
        if _xml_local(el) == "bg":
            bg_el = el
            break
    if bg_el is None:
        return None
    for ch in bg_el:
        ct = _xml_local(ch)
        if ct == "bgPr":
            res = _fill_container_is_dark(ch, part, scheme_map, treat_nofill_light=True)
            if res is not None:
                return res
        elif ct == "bgRef":
            for c_el in ch:
                if _xml_local(c_el) in ("srgbClr", "schemeClr", "sysClr"):
                    rgb, alpha = _resolve_color_el(c_el, scheme_map)
                    if rgb is not None and alpha >= 0.25:
                        return luminance(rgb) < 95
    return None


def detect_slide_dark_bg(slide, pal=None, sw=None, sh=None):
    """Per-slide dark background detection following true Z-order: slide -> layout -> master (shapes + <p:bg>)."""
    pal = pal or {}
    scheme_map = {
        "bg1": pal.get("lt1", (255, 255, 255)),
        "bg2": pal.get("lt2", (245, 245, 245)),
        "lt1": pal.get("lt1", (255, 255, 255)),
        "lt2": pal.get("lt2", (245, 245, 245)),
        "dk1": pal.get("dk1", (0, 0, 0)),
        "dk2": pal.get("dk2", (31, 41, 55)),
        "tx1": pal.get("dk1", (0, 0, 0)),
        "tx2": pal.get("dk2", (31, 41, 55)),
        "phClr": pal.get("accent1", DEFAULT_ACCENT),
        "accent1": pal.get("accent1", DEFAULT_ACCENT),
        "accent2": pal.get("accent2", DEFAULT_ACCENT),
        "accent3": pal.get("accent3", DEFAULT_ACCENT),
        "accent4": pal.get("accent4", DEFAULT_ACCENT),
        "accent5": pal.get("accent5", DEFAULT_ACCENT),
        "accent6": pal.get("accent6", DEFAULT_ACCENT),
    }
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
    except Exception:
        _MST = None
    try:
        layers = [slide, slide.slide_layout, slide.slide_layout.slide_master]
    except Exception:
        layers = [slide]
    slide_area = float(sw * sh) if (sw and sh) else 0.0
    try:
        for obj in layers:
            part = getattr(obj, "part", None)
            # 1) Full-bleed background shapes on this layer (reverse XML order = topmost first)
            if slide_area > 0:
                try:
                    sh_list = list(obj.shapes)
                except Exception:
                    sh_list = []
                for sh_el in reversed(sh_list):
                    w, h = getattr(sh_el, "width", None), getattr(sh_el, "height", None)
                    if not (w and h):
                        continue
                    if (float(w) * float(h)) / slide_area < 0.65:
                        continue
                    if getattr(sh_el, "has_table", False) or getattr(sh_el, "has_chart", False):
                        continue
                    if getattr(sh_el, "has_text_frame", False) and len((sh_el.text_frame.text or "").strip()) > 30:
                        continue
                    # Picture shape -> PIL luminance sampling on image blob
                    if _MST is not None and sh_el.shape_type == _MST.PICTURE:
                        try:
                            res = _image_blob_is_dark(sh_el.image.blob)
                            if res is not None:
                                return res
                        except Exception:
                            pass
                        try:
                            el_root = ET.fromstring(sh_el._element.xml)
                            for bf in el_root.iter():
                                if _xml_local(bf) == "blipFill":
                                    res = _blip_el_is_dark(bf, part)
                                    if res is not None:
                                        return res
                        except Exception:
                            pass
                        continue
                    # AutoShape / background rectangle -> XML fill (solidFill/gradFill/blipFill + shade/lumMod/lumOff)
                    try:
                        el_root = ET.fromstring(sh_el._element.xml)
                        for sp_pr in el_root.iter():
                            if _xml_local(sp_pr) == "spPr":
                                res = _fill_container_is_dark(sp_pr, part, scheme_map, treat_nofill_light=False)
                                if res is not None:
                                    return res
                                break
                    except Exception:
                        pass
            # 2) <p:bg> on this layer (solidFill / gradFill / blipFill / bgRef + modifiers)
            try:
                res = _bg_xml_is_dark(obj._element.xml, scheme_map, part=part)
                if res is not None:
                    return res
            except Exception:
                pass
    except Exception:
        pass
    return False


def master_font_sizes(prs):
    """(title_pt, body_pt) parsed from master txStyles; fallback (36, 20)."""
    try:
        from pptx.oxml.ns import qn
        m = prs.slide_masters[0].element
        ts = m.find(qn("p:txStyles"))

        def sz_of(tag):
            if ts is None:
                return None
            el = ts.find(qn(tag))
            if el is None:
                return None
            lvl1 = el.find(qn("a:lvl1pPr"))
            if lvl1 is None:
                return None
            dr = lvl1.find(qn("a:defRPr"))
            if dr is None:
                return None
            sz = dr.get("sz")
            return float(sz) / 100.0 if sz else None

        return (sz_of("p:titleStyle") or 36.0, sz_of("p:bodyStyle") or 20.0)
    except Exception:
        return 36.0, 20.0


ROLE_KEYS = {
    "title": ["title slide", "标题幻灯片", "标题和副标题", "主标题"],
    "roadmap": ["title and content", "标题和内容", "标题和文本", "section header", "节标题"],
    "motivation": ["title and content", "标题和内容", "标题和文本"],
    "related": ["two content", "两块内容", "两栏内容", "title and content", "标题和文本"],
    "method": ["title and content", "标题和内容", "标题和文本", "two content", "两块内容", "两栏内容"],
    "setup": ["title and content", "标题和内容", "标题和文本"],
    "result": ["title and content", "标题和内容", "标题和文本", "two content", "两块内容", "两栏内容"],
    "analysis": ["title and content", "标题和内容", "标题和文本", "two content", "两块内容", "两栏内容"],
    "limitation": ["title and content", "标题和内容", "标题和文本"],
    "conclusion": ["title and content", "标题和内容", "标题和文本", "section header", "节标题"],
    "qa_backup": ["title and content", "标题和内容", "title only", "仅标题", "blank", "空白"],
    "thanks": ["title slide", "标题幻灯片", "section header", "节标题"],
    "default": ["title and content", "标题和内容", "标题和文本", "blank"],
}


def theme_accent(pptx_path, idx=1):
    """(r,g,b) of theme accent{idx}; None on failure."""
    if not pptx_path or not Path(pptx_path).exists():
        return None
    try:
        import zipfile
        with zipfile.ZipFile(str(pptx_path)) as z:
            xml = z.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
        m = re.search(rf'<a:accent{idx}>.*?<a:srgbClr val="([0-9A-Fa-f]{{6}})"', xml, re.S)
        if not m:
            m = re.search(rf'<a:accent{idx}>.*?lastClr="([0-9A-Fa-f]{{6}})"', xml, re.S)
        if not m:
            return None
        v = m.group(1)
        return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return None


def theme_fonts(pptx_path):
    """(major_latin, minor_latin) from ppt/theme/theme1.xml; ('','') on any failure."""
    if not pptx_path or not Path(pptx_path).exists():
        return ("", "")
    try:
        import zipfile
        with zipfile.ZipFile(str(pptx_path)) as z:
            xml = z.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
        faces = re.findall(r'<a:latin typeface="([^"]+)"', xml)
        return (faces[0] if faces else "", faces[1] if len(faces) > 1 else "")
    except Exception:
        return ("", "")


def resolve_figure_path(raw_path, work: Path):
    if not raw_path:
        return None
    p = Path(raw_path)
    if p.exists():
        return p
    for cand in (
        work / raw_path, work / "figs" / p.name, work / p.name,
        work.parent / raw_path, work.parent / "input" / raw_path, work.parent.parent / raw_path,
    ):
        if cand.exists():
            return cand
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--template")
    ap.add_argument("--deck")
    ap.add_argument("--work", default="/app/work")
    a = ap.parse_args()
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    slides = spec.get("slides") or []
    if not slides:
        print("FATAL: spec.slides empty", file=sys.stderr)
        sys.exit(4)
    meta = spec.get("meta", {})
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)
    log = {"mode": None, "warnings": [], "slide_plan": []}
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE, MSO_SHAPE_TYPE
        from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    except ImportError:
        print("FATAL: python-pptx not installed; use SKILL.md S3 fallback snippet", file=sys.stderr)
        sys.exit(3)

    TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
    BODY_TYPES = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE}
    FOOTER_TYPES = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER}

    base, mode = None, "no_template"
    if a.deck and Path(a.deck).exists():
        base, mode = Path(a.deck), "edit_deck"
    elif a.template and Path(a.template).exists():
        base, mode = Path(a.template), "from_template"
    if mode == "edit_deck":
        shutil.copy(base, out)
        prs = Presentation(str(out))
    elif mode == "from_template":
        prs = Presentation(str(base))
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        log["warnings"].append("no template/deck -> python-pptx default 16:9 template (record in result.uncertainty)")
    log["mode"] = mode

    def drop_all_slides():
        xml = prs.slides._sldIdLst
        n0 = len(list(xml))
        for sld in list(xml):
            try:
                prs.part.drop_rel(sld.rId)
            except Exception:
                pass
            xml.remove(sld)
        return n0

    if mode == "from_template" and meta.get("drop_existing_slides", True):
        n0 = drop_all_slides()
        if n0:
            log["warnings"].append(f"from_template: dropped {n0} pre-existing slide(s)")
    if mode == "edit_deck" and meta.get("drop_existing_slides"):
        drop_all_slides()

    SW, SH = prs.slide_width, prs.slide_height
    MARGIN = Inches(0.5)
    pal = theme_palette(str(base)) if base else {}
    title_sz, body_sz = master_font_sizes(prs)
    raw_major_f, raw_minor_f = theme_fonts(str(base)) if base else ("", "")
    tpl_font = raw_minor_f if (raw_minor_f and not raw_minor_f.startswith("+")) else (
        raw_major_f if (raw_major_f and not raw_major_f.startswith("+")) else ""
    )
    tpl_font_is_cjk = bool(re.search(r"noto|source\s*han|pingfang|yahei|heiti|songti|kaiti|simhei|simsun|wenquanyi|cjk|ya\s*hei|微软|思源|黑体|宋体", tpl_font, re.I))

    def apply_font(run, text_sample=""):
        if tpl_font and (tpl_font_is_cjk or not is_cjk(text_sample or getattr(run, "text", ""))):
            try:
                run.font.name = tpl_font
            except Exception:
                pass
    accent = pal.get("accent1") or theme_accent(str(base)) or DEFAULT_ACCENT
    accent_dark = tuple(max(0, int(ch * 0.65)) for ch in accent)
    zebra_tint = blend_with_white(accent, 0.94)

    # Active per-slide theme tokens (updated per slide in the main loop for mixed dark/light templates!)
    dark_bg = False
    accent_tint = blend_with_white(accent, 0.91)
    card_bg = blend_with_white(accent, 0.95)
    card_border = blend_with_white(accent, 0.72)
    text_primary = (0x1E, 0x29, 0x3B)
    text_muted = (0x52, 0x5E, 0x6E)
    claim_prefix_rgb = accent_dark

    def set_slide_theme(slide):
        nonlocal dark_bg, accent_tint, card_bg, card_border, text_primary, text_muted, claim_prefix_rgb
        dark_bg = detect_slide_dark_bg(slide, pal, SW, SH)
        accent_tint = blend_with_black(accent, 0.72) if dark_bg else blend_with_white(accent, 0.91)
        card_bg = (0x1E, 0x29, 0x3B) if dark_bg else blend_with_white(accent, 0.95)
        card_border = blend_with_white(accent, 0.45) if dark_bg else blend_with_white(accent, 0.72)
        tp = pal.get("lt1", (0xF1, 0xF5, 0xF9)) if dark_bg else pal.get("dk1", (0x1E, 0x29, 0x3B))
        if not dark_bg and luminance(tp) > 120:
            tp = (0x1E, 0x29, 0x3B)
        elif dark_bg and luminance(tp) < 160:
            tp = (0xF1, 0xF5, 0xF9)
        text_primary = tp
        text_muted = (0xCB, 0xD5, 0xE1) if dark_bg else (0x52, 0x5E, 0x6E)
        claim_prefix_rgb = blend_with_white(accent, 0.35) if dark_bg else (
            tuple(max(0, int(ch * 0.45)) for ch in accent) if luminance(accent) > 155 else accent_dark
        )

    layouts = list(prs.slide_layouts)
    names = [(lo.name or "").strip().lower() for lo in layouts]

    def strip_shadow(sh):
        sh.shadow.inherit = False
        try:
            from pptx.oxml.ns import qn
            style = sh._element.find(qn("p:style"))
            if style is not None:
                ef = style.find(qn("a:effectRef"))
                if ef is not None:
                    ef.set("idx", "0")
        except Exception:
            pass

    def pick_layout(role, hint=None):
        if hint:
            hl = hint.strip().lower()
            for i, n in enumerate(names):
                if n == hl:
                    return layouts[i]
        keys = [k.lower() for k in ROLE_KEYS.get(role, ROLE_KEYS["default"])]
        for k in keys:
            for i, n in enumerate(names):
                if k in n:
                    return layouts[i]
        for i, n in enumerate(names):
            if "content" in n or "内容" in n:
                return layouts[i]
        return layouts[min(1, len(layouts) - 1)]

    def phs(slide, allow_fallback=False):
        t = b = None
        for ph in slide.placeholders:
            ty = ph.placeholder_format.type
            if ty in TITLE_TYPES and t is None:
                t = ph
            elif ty in BODY_TYPES and b is None:
                b = ph
        if allow_fallback and (t is None or b is None):
            text_shapes = [
                sh for sh in slide.shapes
                if sh.has_text_frame and sh.text_frame.text.strip()
                and sh is not t and sh is not b
                and sh.left is not None and sh.top is not None and sh.width and sh.height
                and sh.top < SH - Inches(0.65)
            ]
            if t is None:
                top_cands = [sh for sh in text_shapes if sh.top < int(SH * 0.25) and sh.width > int(SW * 0.30)]
                if top_cands:
                    t = min(top_cands, key=lambda sh: sh.top)
            if b is None:
                body_cands = [sh for sh in text_shapes if sh is not t and sh.top >= int(SH * 0.16) and sh.width > int(SW * 0.25)]
                if body_cands:
                    b = max(body_cands, key=lambda sh: (sh.width * sh.height))
        return t, b

    def detect_footer_y(slide):
        """Find a safe Y coordinate for the evidence footer without overlapping template/master chrome."""
        foot_y = int(SH - Inches(0.36))
        containers = [slide.shapes]
        try:
            containers.append(slide.slide_layout.shapes)
            containers.append(slide.slide_layout.slide_master.shapes)
        except Exception:
            pass
        for shapes in containers:
            for sh in shapes:
                l, top_v, w, h = sh.left, sh.top, sh.width, sh.height
                if None in (l, top_v, w, h):
                    continue
                if top_v + h < int(SH - Inches(0.62)):
                    continue
                # Ignore full-slide background fills and full-bleed background pictures (>=65% area)
                if (w > int(SW * 0.92) and h > int(SH * 0.70)) or ((w * h) >= 0.65 * SW * SH):
                    continue
                has_txt = bool(sh.has_text_frame and sh.text_frame.text.strip())
                is_footer_ph = bool(sh.is_placeholder and sh.placeholder_format.type in FOOTER_TYPES and has_txt)
                is_bottom_bar = (
                    sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                    and w > int(SW * 0.50) and h <= int(Inches(0.48))
                    and top_v >= int(SH - Inches(0.55))
                )
                is_bottom_chrome = (
                    (sh.shape_type == MSO_SHAPE_TYPE.PICTURE and h < int(SH * 0.18))
                    or is_footer_ph
                    or is_bottom_bar
                    or (has_txt and top_v > int(SH - Inches(0.55)) and l < int(SW * 0.60))
                )
                if is_bottom_chrome:
                    foot_y = min(foot_y, int(top_v - Inches(0.28)))
        return max(int(SH - Inches(0.70)), foot_y)

    def detect_top_right_logo_pad(slide):
        """Return horizontal right padding (EMU) if master/layout has a top-right institutional logo (or slide has a small top-right picture logo)."""
        pad = 0
        master_containers = []
        try:
            master_containers.append(slide.slide_layout.shapes)
            master_containers.append(slide.slide_layout.slide_master.shapes)
        except Exception:
            pass
        for shapes in master_containers:
            for sh in shapes:
                l, top_v, w, h = sh.left, sh.top, sh.width, sh.height
                if None in (l, top_v, w, h) or sh.is_placeholder:
                    continue
                if top_v < int(Inches(1.10)) and l >= int(SW * 0.68) and w < int(SW * 0.34) and h < int(SH * 0.20):
                    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or (sh.has_text_frame and sh.text_frame.text.strip()):
                        pad = max(pad, int(SW - l + Inches(0.18)))
        for sh in slide.shapes:
            l, top_v, w, h = sh.left, sh.top, sh.width, sh.height
            if None in (l, top_v, w, h) or sh.is_placeholder:
                continue
            is_small_pic = (sh.shape_type == MSO_SHAPE_TYPE.PICTURE and top_v < int(Inches(0.95)) and l >= int(SW * 0.72) and w < int(SW * 0.28) and h < int(Inches(0.95)))
            is_badge_txt = (sh.has_text_frame and 0 < len(sh.text_frame.text.strip()) <= 22 and top_v < int(Inches(0.80)) and l >= int(SW * 0.75) and w <= int(Inches(2.6)) and h <= int(Inches(0.80)))
            if is_small_pic or is_badge_txt:
                pad = max(pad, int(SW - l + Inches(0.18)))
        return min(pad, int(SW * 0.28))

    def slide_frame(slide, t, b):
        """Derive template-respecting (left, width, foot_y) from actual placeholder geometry."""
        fl, fw = int(MARGIN), int(SW - 2 * MARGIN)
        ref = None
        if t is not None and t.left is not None and t.width and t.width >= int(SW * 0.55) and int(Inches(0.20)) <= t.left <= int(SW * 0.28):
            ref = t
        elif b is not None and b.left is not None and b.width and b.width >= int(SW * 0.55) and int(Inches(0.20)) <= b.left <= int(SW * 0.28):
            ref = b
        if ref is not None:
            fl = int(ref.left)
            right_margin = int(Inches(0.50))
            max_fw = int(SW - fl - right_margin)
            # If template has 4:3 placeholders stretched on 16:9 canvas (leaving >20% dead space on right), use full width
            if ref.left + ref.width < int(SW * 0.80):
                fw = max_fw
            else:
                fw = int(min(ref.width, max_fw))
        foot_y = detect_footer_y(slide)
        return fl, fw, foot_y

    def remove_shape(sh):
        try:
            el = sh._element
            el.getparent().remove(el)
        except Exception:
            pass

    def cleanup_unused_placeholders(slide, keep_shapes):
        """Remove empty non-footer placeholders (e.g., 2nd body box on Two Content) to prevent ghost frames."""
        keep_els = {s._element for s in keep_shapes if s is not None}
        for sh in list(slide.placeholders):
            if sh._element in keep_els:
                continue
            ty = sh.placeholder_format.type
            if ty in TITLE_TYPES or ty in FOOTER_TYPES:
                continue
            txt = sh.text_frame.text.strip() if sh.has_text_frame else ""
            if not txt:
                remove_shape(sh)

    def clear_overlapping_visuals(slide, region_box, keep_shapes=None, clear_orig_els=None, only_orig_els=None):
        """Remove existing pictures/tables/charts (and stale original non-placeholder shapes when clear_orig_els set) colliding with region_box."""
        rl, rt, rw, rh = region_box
        keep_els = {s._element for s in (keep_shapes or []) if s is not None}
        to_del = []
        for sh in list(slide.shapes):
            if sh._element in keep_els:
                continue
            if only_orig_els is not None and sh._element not in only_orig_els:
                continue
            l, t, w, h = sh.left, sh.top, sh.width, sh.height
            if None in (l, t, w, h):
                continue
            # Never remove full-slide backgrounds, full-bleed >=65% bg pictures, top/bottom banner chrome, or tall sidebar strips
            if (w > int(SW * 0.88) and h > int(SH * 0.78)) or (sh.shape_type == MSO_SHAPE_TYPE.PICTURE and (w * h) >= 0.65 * SW * SH):
                continue
            if t <= int(Inches(0.35)) and w >= int(SW * 0.80) and h <= int(Inches(1.35)):
                continue
            if t >= int(SH - Inches(0.85)) and w >= int(SW * 0.80) and h <= int(Inches(0.90)):
                continue
            if (l <= int(Inches(0.35)) or l + w >= int(SW - Inches(0.35))) and w <= int(Inches(0.70)) and h >= int(SH * 0.65):
                continue
            if t >= int(SH - Inches(0.60)) and h <= int(Inches(0.55)):
                continue
            if sh.is_placeholder and sh.placeholder_format.type in (TITLE_TYPES | FOOTER_TYPES):
                continue
            is_vis = (sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP) or getattr(sh, "has_table", False) or getattr(sh, "has_chart", False))
            has_txt_sh = bool(sh.has_text_frame and sh.text_frame.text.strip())
            is_thin_decor = (not is_vis) and (not has_txt_sh) and (w < int(Inches(0.22)) or h < int(Inches(0.18)))
            is_stale_orig = bool(clear_orig_els and sh._element in clear_orig_els and not sh.is_placeholder and not is_thin_decor)
            if not (is_vis or is_stale_orig):
                continue
            ix = max(0, min(l + w, rl + rw) - max(l, rl))
            iy = max(0, min(t + h, rt + rh) - max(t, rt))
            if ix * iy > 0:
                to_del.append(sh)
        for sh in to_del:
            remove_shape(sh)

    def set_geom(sh, left=None, top=None, width=None, height=None):
        """Set placeholder/shape geometry atomically so python-pptx never zeroes inherited coords."""
        l0, t0, w0, h0 = sh.left, sh.top, sh.width, sh.height
        nl = int(l0 if left is None else left) if (left is not None or l0 is not None) else None
        nt = int(t0 if top is None else top) if (top is not None or t0 is not None) else None
        nw = int(w0 if width is None else width) if (width is not None or w0 is not None) else None
        nh = int(h0 if height is None else height) if (height is not None or h0 is not None) else None
        if nl is not None:
            sh.left = nl
        if nt is not None:
            sh.top = nt
        if nw is not None:
            sh.width = nw
        if nh is not None:
            sh.height = nh

    def normalize_cover_placeholders(t, b):
        """Center cover placeholders cleanly when 4:3 placeholders sit on a 16:9 canvas."""
        for ph in (t, b):
            if ph is None or ph.width is None:
                continue
            if ph.width < int(SW * 0.85):
                center_x = (ph.left or 0) + ph.width / 2.0
                if abs(center_x - SW / 2.0) > Inches(0.4) or ph.width < int(SW * 0.78):
                    new_w = int(min(SW - Inches(1.2), max(ph.width, SW * 0.84)))
                    set_geom(ph, left=(SW - new_w) // 2, width=new_w)

    def set_title(slide, text, t, role="default", is_reuse=False, frame=None):
        if not text:
            return None
        text = polish_math_text(text)
        fl, fw = (frame[0], frame[1]) if frame else (int(MARGIN), int(SW - 2 * MARGIN))
        logo_pad = detect_top_right_logo_pad(slide) if role not in ("title", "thanks") else 0
        eff_fw = max(int(SW * 0.45), int(min(fw, SW - fl - logo_pad))) if logo_pad else fw
        if t is not None:
            if role in ("title", "thanks"):
                if not is_reuse:
                    normalize_cover_placeholders(t, None)
            else:
                # Fix broken width, 4:3-on-16:9 dead right margin (also on edited reused slides), or top-right logo collision
                if t.width is None or t.width < int(SW * 0.25) or t.left is None:
                    nt = int(Inches(0.30)) if (t.top is None or t.top > Inches(1.2)) else t.top
                    set_geom(t, left=fl, top=nt, width=eff_fw)
                elif (t.left + t.width) < int(SW * 0.80) and abs(t.left - fl) <= Inches(0.25):
                    set_geom(t, left=fl, width=eff_fw)
                elif logo_pad and (t.left + t.width) > int(SW - logo_pad):
                    set_geom(t, width=max(int(SW * 0.45), int(SW - logo_pad - t.left)))
            t.text_frame.text = text
            try:
                w_pt = max(60.0, (t.width - 182880) / EMU_IN * 72.0) if t.width else 600.0
                base_sz = title_sz
                # Normalize giant default master titles (>34pt) on content slides for crisp hierarchy
                if role not in ("title", "thanks") and base_sz > 34.0:
                    base_sz = 30.0
                est_lines = max(1, math.ceil(char_units(text) * base_sz / w_pt))
                eff_sz = base_sz
                if role in ("title", "thanks"):
                    # Cover titles naturally wrap to 2 lines; shrink gently if >2 lines OR if line 2 would be a short 1-word orphan
                    if est_lines > 2 and base_sz > 28:
                        eff_sz = max(28.0, min(base_sz, (2.0 * w_pt) / max(1.0, char_units(text)) * 0.94))
                    elif est_lines == 2 and not is_reuse and base_sz >= 32:
                        ratio_1l = (char_units(text) * base_sz) / w_pt
                        if ratio_1l <= 1.24:
                            eff_sz = max(28.0, min(base_sz, w_pt / max(1.0, char_units(text)) * 0.95))
                elif est_lines > 1 and base_sz > 26:
                    eff_sz = max(22.0, min(base_sz, w_pt / max(1.0, char_units(text)) * 0.96))
                for p0 in t.text_frame.paragraphs:
                    p0.alignment = PP_ALIGN.CENTER if role in ("title", "thanks") else PP_ALIGN.LEFT
                    for r0 in p0.runs:
                        if eff_sz != title_sz:
                            r0.font.size = Pt(round(eff_sz))
                        if dark_bg:
                            needs_light = not is_reuse
                            if is_reuse:
                                try:
                                    c_rgb = r0.font.color.rgb if r0.font.color and r0.font.color.type else None
                                    needs_light = (c_rgb is None) or (luminance((c_rgb[0], c_rgb[1], c_rgb[2])) < 120)
                                except Exception:
                                    needs_light = True
                            if needs_light:
                                r0.font.color.rgb = RGBColor(*text_primary)
                need_emu = int(max(1, math.ceil(char_units(text) * eff_sz / w_pt)) * eff_sz * 1.24 / 72.0 * EMU_IN + 140000)
                if t.height is not None and (need_emu > t.height or (role not in ("title", "thanks") and t.height > need_emu + Inches(0.35))):
                    set_geom(t, height=max(need_emu, int(Inches(0.62))))
                # Ensure full-width top header banners and full-bleed >=65% background pictures stay BEHIND the title in z-order
                sp_tree = slide.shapes._spTree
                for _bsh in list(slide.shapes):
                    if _bsh.shape_type == MSO_SHAPE_TYPE.PICTURE and None not in (_bsh.top, _bsh.width, _bsh.height):
                        if _bsh.top <= int(Inches(0.35)) and _bsh.width >= int(SW * 0.80) and _bsh.height <= int(Inches(1.35)):
                            sp_tree.remove(_bsh._element)
                            sp_tree.insert(2, _bsh._element)
                for _bsh in list(slide.shapes):
                    if _bsh.shape_type == MSO_SHAPE_TYPE.PICTURE and None not in (_bsh.width, _bsh.height):
                        if (_bsh.width * _bsh.height) >= 0.65 * SW * SH:
                            sp_tree.remove(_bsh._element)
                            sp_tree.insert(2, _bsh._element)
            except Exception:
                pass
            return None
        tb = slide.shapes.add_textbox(fl, Inches(0.30), eff_fw, Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER if role in ("title", "thanks") else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = text
        r.font.size = Pt(28 if role not in ("title", "thanks") else 36)
        r.font.bold = True
        r.font.color.rgb = RGBColor(*text_primary)
        apply_font(r, text)
        return tb

    def fix_body_geom(b, t, role="default", is_reuse=False, frame=None):
        """Repair ONLY truly broken placeholder geometry (zero/narrow width, off-center cover, runaway height)."""
        if b is None:
            return
        if not is_reuse and role in ("title", "thanks"):
            normalize_cover_placeholders(None, b)
            return
        fl, fw = (frame[0], frame[1]) if frame else (int(MARGIN), int(SW - 2 * MARGIN))
        nl, nt, nw, nh = b.left, b.top, b.width, b.height
        changed = False
        if nw is None or nw < int(SW * 0.25) or nl is None:
            nl, nw, changed = fl, fw, True
        elif not is_reuse and nl == fl and nw < fw and (nl + nw) < int(SW * 0.80) and nw >= int(SW * 0.55):
            nw, changed = fw, True
        if nh is None or nh > int(SH * 0.90):
            nt = nt if (nt is not None and nt > Inches(0.8)) else int(Inches(1.70))
            nh = int(SH - nt - Inches(0.50))
            changed = True
        if changed:
            set_geom(b, left=nl, top=nt, width=nw, height=nh)

    def remove_existing_claim(slide, title_shape=None):
        """Remove any stale claim bar (+ left accent stripe) before inserting updated claim."""
        bars = []
        CLAIM_PREFIXES = (
            "本页主张：", "本页主张 ·", "主张：", "Claim:", "TAKEAWAY ·",
            "动机：", "定位：", "主线：", "方法：", "设置：", "结果：", "分析：", "局限：", "进展：", "总结：", "备用：",
            "问题动机 ·", "研究定位 ·", "汇报主线 ·", "核心方法 ·", "实验设置 ·",
            "核心结果 ·", "深度分析 ·", "局限与边界 ·", "阶段进展 ·", "总结展望 ·", "追问备用 ·",
            "Motivation:", "Context:", "Roadmap:", "Method:", "Setup:", "Result:", "Analysis:", "Limits:", "Progress:", "Takeaway:", "Q&A:",
            "MOTIVATION ·", "POSITIONING ·", "ROADMAP ·", "METHOD ·", "SETUP ·",
            "KEY RESULT ·", "ANALYSIS ·", "SCOPE & LIMITS ·", "PROGRESS ·", "Q&A BACKUP ·",
        )
        t_el = title_shape._element if title_shape is not None else None
        for sh in list(slide.shapes):
            if sh.is_placeholder or (t_el is not None and sh._element == t_el):
                continue
            if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
                continue
            if sh.top is not None and sh.top < int(Inches(0.85)):
                continue
            if sh.has_text_frame:
                tx = sh.text_frame.text.strip()
                if tx.startswith(CLAIM_PREFIXES):
                    bars.append((sh, sh.left, sh.top, sh.height))
        for sh, bl, bt, bh in bars:
            remove_shape(sh)
            for other in list(slide.shapes):
                if (not other.has_text_frame or not other.text_frame.text.strip()) and other.left is not None and other.top is not None:
                    if bl is not None and bt is not None and abs(other.left - bl) < Inches(0.08) and abs(other.top - bt) < Inches(0.08) and other.width and other.width <= Inches(0.14):
                        remove_shape(other)

    def add_claim_strip(slide, claim, t, title_text="", role="default", frame=None):
        if role in ("title", "thanks"):
            return None
        remove_existing_claim(slide, title_shape=t)
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        if t is not None and t.top is not None and t.height:
            top = t.top + t.height
            sz = title_sz
            try:
                for p in t.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size:
                            sz = r.font.size.pt
                            break
            except Exception:
                pass
            w_pt = max(60.0, (t.width - 182880) / EMU_IN * 72.0) if t.width else 600.0
            need_emu = int(max(1, math.ceil(char_units(title_text) * sz / w_pt)) * sz * 1.24 / 72.0 * EMU_IN + 140000)
            if need_emu > t.height:
                top = t.top + need_emu
            elif t.height > need_emu + Inches(0.35):
                try:
                    l0, w0, t0 = t.left, t.width, t.top
                    new_h = int(need_emu + Inches(0.06))
                    t.left = int(l0 if l0 is not None else fl)
                    t.top = int(t0 if t0 is not None else Inches(0.28))
                    t.width = int(w0 if w0 else fw)
                    t.height = new_h
                    top = t.top + new_h
                except Exception:
                    top = t.top + t.height
        else:
            top = Inches(1.06)
        top = max(int(top) + int(Inches(0.04)), int(Inches(0.95)))
        left = int(fl)
        width = int(fw)
        ROLE_TAG_ZH = {
            "motivation": "动机：", "related": "定位：", "roadmap": "主线：",
            "method": "方法：", "setup": "设置：", "result": "结果：",
            "analysis": "分析：", "limitation": "局限：", "progress": "进展：",
            "conclusion": "总结：", "qa_backup": "备用：",
        }
        ROLE_TAG_EN = {
            "motivation": "Motivation: ", "related": "Context: ", "roadmap": "Roadmap: ",
            "method": "Method: ", "setup": "Setup: ", "result": "Result: ",
            "analysis": "Analysis: ", "limitation": "Limits: ", "progress": "Progress: ",
            "conclusion": "Takeaway: ", "qa_backup": "Q&A: ",
        }
        role_l = str(role or "default").lower()
        prefix = (ROLE_TAG_ZH.get(role_l, "主张：") if is_cjk(claim)
                  else ROLE_TAG_EN.get(role_l, "Claim: "))
        full_txt = prefix + claim
        claim_pt = 13.5
        usable_w_pt = max(120.0, (width - Inches(0.30)) / EMU_IN * 72.0)
        n_lines = max(1, math.ceil(char_units(full_txt) * claim_pt / usable_w_pt))
        h = int(max(Inches(0.42), Inches(n_lines * claim_pt * 1.32 / 72.0 + 0.14)))
        if top + h >= foot_y - Inches(0.30):
            return None

        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(*accent_tint)
        bar.line.fill.background()
        strip_shadow(bar)
        tf = bar.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.16)
        tf.margin_right = Inches(0.12)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r1 = p.add_run()
        r1.text = prefix
        r1.font.size = Pt(claim_pt)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(*claim_prefix_rgb)
        apply_font(r1, prefix)
        r2 = p.add_run()
        r2.text = polish_math_text(claim)
        r2.font.size = Pt(claim_pt)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(*text_primary)
        apply_font(r2, claim)

        stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.06), h)
        stripe.fill.solid()
        stripe.fill.fore_color.rgb = RGBColor(*accent)
        stripe.line.fill.background()
        strip_shadow(stripe)
        return bar, top + h

    def add_evidence_footer(slide, evidence, role, lang_cjk, frame=None):
        if role in ("title", "thanks") or not meta.get("evidence_footer", True):
            return
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), detect_footer_y(slide))
        # Replace any stale evidence footer on reused/re-edited slides so evidence stays current
        for sh in list(slide.shapes):
            if sh.has_text_frame and sh.top is not None and sh.top > SH - Inches(0.85):
                if sh.text_frame.text.strip().startswith(("来源：", "Source:", "备用问答页", "Q&A Backup")):
                    remove_shape(sh)
        parts = []
        seen = set()
        for e in (evidence or []):
            loc = (e.get("locator") or "").strip()
            st = (e.get("strength") or "").strip()
            if not loc:
                continue
            tag = f"{loc} [{st}]" if st else loc
            if tag not in seen:
                seen.add(tag)
                parts.append(tag)
        if parts:
            lead = "来源：" if lang_cjk else "Source: "
            text = lead + " · ".join(parts)
        elif role == "qa_backup":
            text = "备用问答页（不计入正文时长）" if lang_cjk else "Q&A Backup Slide (excluded from main talk time)"
        else:
            return
        if len(text) > 105:
            text = text[:102] + "..."
        fb = slide.shapes.add_textbox(fl, foot_y, min(fw - Inches(0.95), SW - fl - Inches(1.2)), Inches(0.26))
        ftf = fb.text_frame
        ftf.word_wrap = False
        ftf.margin_left = ftf.margin_top = ftf.margin_bottom = 0
        fp = ftf.paragraphs[0]
        fr = fp.add_run()
        fr.text = text
        fr.font.size = Pt(9.0 if len(text) > 88 else 10.0)
        fr.font.color.rgb = RGBColor(*text_muted)
        apply_font(fr, text)
        if mode == "no_template":
            try:
                fdiv = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, fl, int(foot_y - Inches(0.05)), fw, Inches(0.012))
                fdiv.fill.solid()
                fdiv.fill.fore_color.rgb = RGBColor(*card_border)
                fdiv.line.fill.background()
                strip_shadow(fdiv)
            except Exception:
                pass

    def add_kpis(slide, kpis, top_y, frame=None, clear_old=False, only_orig_els=None):
        """Render 2-4 compact metric highlight cards across top band; returns bottom EMU."""
        items = [k for k in (kpis or []) if isinstance(k, dict) and k.get("value")]
        if not items:
            return top_y
        fl, fw, _ = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        items = items[:4]
        n = len(items)
        gap = Inches(0.18)
        total_w = fw
        card_w = int((total_w - gap * (n - 1)) / n)
        w_pt_est = max(72.0, (card_w - Inches(0.28)) / EMU_IN * 72.0)
        # Fit value & label font sizes cleanly so KPI cards NEVER overflow even with 4 cards on 4:3
        val_sz = 22.5 if n <= 3 else 18.5
        lbl_sz = 11.2 if n <= 3 else 10.0
        need_h_in = 1.04
        def _delta_color(d_s: str):
            dl = (d_s or "").strip().lower()
            if re.search(r"^(?:[-−]\s*\d)|bottleneck|drop|loss|fail|risk|limit|oom|degrad|slower|cost|wall|瓶颈|下降|损耗|风险|局限", dl):
                return (0xFB, 0xBF, 0x24) if dark_bg else (0xB4, 0x53, 0x09)
            if re.search(r"^(?:\+\s*\d|↑|↓\s*\d)|faster|speedup|gain|win|best|save|reduc|提升|加速|优于|降低", dl):
                return (0x5E, 0xEA, 0xD4) if dark_bg else (0x0F, 0x76, 0x6E)
            return text_muted

        inline_delta = True
        for k in items:
            vtxt = str(k.get("value", ""))
            dtxt = str(k.get("delta", "") or "")
            ltxt = str(k.get("label", "") or "")
            if dtxt and (char_units(vtxt) * val_sz + char_units(f" · {dtxt}") * 10.5) > w_pt_est * 0.96:
                inline_delta = False
            best_k = 13.5
            for cand in (val_sz, 20.0, 17.5, 15.5, 13.5):
                if cand > val_sz:
                    continue
                used = char_units(vtxt) * cand + (char_units(f" · {dtxt}") * 10.5 if (dtxt and inline_delta) else 0.0)
                if used <= w_pt_est * 0.96:
                    best_k = cand
                    break
            val_sz = min(val_sz, best_k)
            v_lines = max(1, math.ceil((char_units(vtxt) * val_sz + (char_units(f" · {dtxt}") * 10.5 if (dtxt and inline_delta) else 0.0)) / w_pt_est))
            d_extra_lines = (1 if (dtxt and not inline_delta) else 0)
            for lc in (lbl_sz, 9.5, 8.8):
                if lc > lbl_sz:
                    continue
                l_lines = max(1, math.ceil(char_units(ltxt) * lc / w_pt_est)) if ltxt else 0
                est_in = (v_lines * val_sz * 1.28 + d_extra_lines * 13.5 + l_lines * lc * 1.30 + 18.0) / 72.0 + 0.18
                if est_in <= 1.34 or lc == 8.8:
                    lbl_sz = min(lbl_sz, lc)
                    need_h_in = max(need_h_in, est_in)
                    break
        card_h = int(Inches(min(1.36, max(1.04, need_h_in))))
        y = int(top_y + Inches(0.10))
        if clear_old:
            clear_overlapping_visuals(slide, (fl, y, fw, card_h), only_orig_els=only_orig_els)
        for idx, k in enumerate(items):
            x = int(fl + idx * (card_w + gap))
            bar_h = Inches(0.055)
            sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y + bar_h, card_w, card_h - bar_h)
            sh.fill.solid()
            sh.fill.fore_color.rgb = RGBColor(*card_bg)
            sh.line.color.rgb = RGBColor(*card_border)
            sh.line.width = Pt(1.0)
            strip_shadow(sh)
            kbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, bar_h)
            kbar.fill.solid()
            kbar.fill.fore_color.rgb = RGBColor(*accent)
            kbar.line.fill.background()
            strip_shadow(kbar)
            tf = sh.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.14)
            tf.margin_right = Inches(0.12)
            tf.margin_top = Inches(0.06)
            tf.margin_bottom = Inches(0.06)

            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.LEFT
            rv = p0.add_run()
            rv.text = polish_math_text(str(k.get("value", "")))
            rv.font.size = Pt(val_sz)
            rv.font.bold = True
            rv.font.color.rgb = RGBColor(*claim_prefix_rgb)
            apply_font(rv, rv.text)
            if k.get("delta") and inline_delta:
                rd = p0.add_run()
                rd.text = f"  ·  {polish_math_text(str(k['delta']))}"
                rd.font.size = Pt(10.5)
                rd.font.bold = True
                rd.font.color.rgb = RGBColor(*_delta_color(str(k["delta"])))
                apply_font(rd, rd.text)

            if k.get("label"):
                p1 = tf.add_paragraph()
                p1.alignment = PP_ALIGN.LEFT
                p1.space_before = Pt(3)
                rl = p1.add_run()
                rl.text = polish_math_text(str(k["label"]))
                rl.font.size = Pt(lbl_sz)
                rl.font.color.rgb = RGBColor(*text_muted)
                apply_font(rl, rl.text)

            if k.get("delta") and not inline_delta:
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.LEFT
                p2.space_before = Pt(2)
                rd2 = p2.add_run()
                rd2.text = polish_math_text(str(k["delta"]))
                rd2.font.size = Pt(10.0)
                rd2.font.bold = True
                rd2.font.color.rgb = RGBColor(*_delta_color(str(k["delta"])))
                apply_font(rd2, rd2.text)
        return y + card_h

    def add_steps(slide, steps, top_y, has_below=False, frame=None, clear_old=False, only_orig_els=None):
        """Render 2-5 horizontal numbered step cards (for roadmap / method pipeline); returns bottom EMU."""
        items = [st for st in (steps or []) if isinstance(st, dict) and (st.get("title") or st.get("detail"))]
        if not items:
            return top_y
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        items = items[:5]
        n = len(items)
        gap = Inches(0.26) if fw >= int(Inches(10.5)) else Inches(0.18)
        total_w = fw
        card_w = int((total_w - gap * (n - 1)) / n)
        y = int(top_y + Inches(0.14))
        max_bot = foot_y - Inches(0.16)
        avail_h = int(max(Inches(1.42), (max_bot - y) if not has_below else min(Inches(1.92), (max_bot - y) * 0.42)))
        w_pt_est = max(72.0, (card_w - Inches(0.32)) / EMU_IN * 72.0)

        def est_step_in(ts, ds):
            mx = 1.55
            for st in items:
                tl = max(1, math.ceil(char_units(str(st.get("title", ""))) * ts / w_pt_est)) if st.get("title") else 0
                dl = max(1, math.ceil(char_units(str(st.get("detail", ""))) * ds / w_pt_est)) if st.get("detail") else 0
                est = (18.0 + tl * ts * 1.32 + dl * ds * 1.35 + 32.0) / 72.0 + 0.44
                mx = max(mx, est)
            return mx

        # Auto-fit step font pair so text NEVER overflows card_h (including 4:3 and has_below=True)
        cand_pairs = [(16.5, 13.8), (15.0, 12.5), (13.5, 11.2), (12.0, 10.0)] if n <= 4 else [(14.5, 11.8), (13.0, 10.5), (11.5, 9.2)]
        t_sz, d_sz = cand_pairs[-1]
        need_max_in = est_step_in(t_sz, d_sz)
        avail_in = avail_h / EMU_IN
        for ts_c, ds_c in cand_pairs:
            ni = est_step_in(ts_c, ds_c)
            if ni <= avail_in * 0.96:
                t_sz, d_sz, need_max_in = ts_c, ds_c, ni
                break
        min_stand_in = 1.90 if not has_below else 1.40
        cap_h = int(Inches(min(3.30 if not has_below else 1.92, max(min_stand_in, need_max_in + (0.20 if not has_below else 0.0)))))
        card_h = int(min(avail_h, cap_h))
        # Optical vertical balance when steps are the standalone body of the slide
        if not has_below and avail_h > card_h + Inches(0.50):
            y = int(y + min(Inches(0.75), (avail_h - card_h) * 0.28))
        if clear_old:
            clear_overlapping_visuals(slide, (fl, y, fw, card_h), only_orig_els=only_orig_els)
        for idx, st in enumerate(items):
            x = int(fl + idx * (card_w + gap))
            # Accent top bar (no text -> zero overlap)
            bar_h = Inches(0.08)
            tbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, bar_h)
            tbar.fill.solid()
            tbar.fill.fore_color.rgb = RGBColor(*accent)
            tbar.line.fill.background()
            strip_shadow(tbar)

            # Single card shape holding badge + title + detail (zero self-overlap)
            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y + bar_h, card_w, card_h - bar_h)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(*card_bg)
            card.line.color.rgb = RGBColor(*card_border)
            card.line.width = Pt(1.0)
            strip_shadow(card)

            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.margin_left = Inches(0.18)
            tf.margin_right = Inches(0.14)
            tf.margin_top = Inches(0.16)
            tf.margin_bottom = Inches(0.12)

            lbl = str(st.get("label") or f"0{idx + 1}")
            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.LEFT
            r_lbl = p0.add_run()
            r_lbl.text = lbl
            r_lbl.font.size = Pt(13)
            r_lbl.font.bold = True
            r_lbl.font.color.rgb = RGBColor(*claim_prefix_rgb)
            apply_font(r_lbl, lbl)

            if st.get("title"):
                p1 = tf.add_paragraph()
                p1.alignment = PP_ALIGN.LEFT
                p1.space_before = Pt(5)
                rt = p1.add_run()
                rt.text = polish_math_text(str(st["title"]))
                rt.font.size = Pt(t_sz)
                rt.font.bold = True
                rt.font.color.rgb = RGBColor(*text_primary)
                apply_font(rt, rt.text)

            if st.get("detail"):
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.LEFT
                p2.space_before = Pt(7)
                rd = p2.add_run()
                rd.text = polish_math_text(str(st["detail"]))
                rd.font.size = Pt(d_sz)
                rd.font.color.rgb = RGBColor(*text_muted)
                apply_font(rd, rd.text)

            # Thin arrow chevron between cards (no text -> zero overlap)
            if idx < n - 1:
                aw = Inches(0.18)
                ah = Inches(0.24)
                ax = x + card_w + (gap - aw) // 2
                ay = y + (card_h - ah) // 2
                arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, int(ax), int(ay), int(aw), int(ah))
                arr.fill.solid()
                arr.fill.fore_color.rgb = RGBColor(*accent)
                arr.line.fill.background()
                strip_shadow(arr)

        return y + card_h

    def add_bullet_runs(p, text, sz_pt=None, color_rgb=None, bold_lead=True):
        """Render a bullet paragraph, optionally auto-bolding a short lead-in label before ':' / '：'."""
        text = polish_math_text(text)
        lead = rest = None
        if bold_lead and text:
            m = re.match(r"^([^:：\n]{2,26})([：:])(\s*\S.*)$", text)
            if m and char_units(m.group(1)) <= 14.0 and not re.match(r"^(https?|ftp)$", m.group(1), re.I):
                lead = m.group(1) + m.group(2)
                rest = m.group(3)
        if lead is not None:
            r1 = p.add_run()
            r1.text = lead
            r1.font.bold = True
            if sz_pt is not None:
                r1.font.size = Pt(sz_pt)
            if color_rgb is not None:
                r1.font.color.rgb = RGBColor(*color_rgb)
            apply_font(r1, lead)
            r2 = p.add_run()
            r2.text = rest
            if sz_pt is not None:
                r2.font.size = Pt(sz_pt)
            if color_rgb is not None:
                r2.font.color.rgb = RGBColor(*color_rgb)
            apply_font(r2, rest)
        else:
            r = p.add_run()
            r.text = text
            if sz_pt is not None:
                r.font.size = Pt(sz_pt)
            if color_rgb is not None:
                r.font.color.rgb = RGBColor(*color_rgb)
            apply_font(r, text)

    def add_two_col(slide, tc, top_y, frame=None, clear_old=False, only_orig_els=None):
        """Render balanced 2-column structured comparison cards (each card is 1 text shape)."""
        if not isinstance(tc, dict):
            return top_y
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        cols = [
            (tc.get("left_title", ""), tc.get("left_bullets") or []),
            (tc.get("right_title", ""), tc.get("right_bullets") or []),
        ]
        if not any(c[0] or c[1] for c in cols):
            return top_y
        gap = Inches(0.28)
        col_w = int((fw - gap) / 2)
        y = int(top_y + Inches(0.14))
        max_col_h = int(max(Inches(2.0), foot_y - y - Inches(0.18)))
        w_pt_est = max(120.0, (col_w - Inches(0.36)) / EMU_IN * 72.0)
        need_col_in = 2.15
        for ctitle, cbullets in cols:
            pts = (28.0 if ctitle else 0.0)
            for item in cbullets:
                txt = item.get("text", "") if isinstance(item, dict) else str(item)
                lines = max(1, math.ceil(char_units(f"•  {txt}") * 16.0 / w_pt_est))
                pts += lines * 16.0 * 1.32 + 9.0
            need_col_in = max(need_col_in, pts / 72.0 + 0.65)
        col_h = int(min(max_col_h, max(Inches(2.85), Inches(need_col_in + 0.25))))
        if max_col_h > col_h + Inches(0.40):
            y = int(y + min(Inches(0.55), (max_col_h - col_h) * 0.28))
        if clear_old:
            clear_overlapping_visuals(slide, (fl, y, fw, col_h), only_orig_els=only_orig_els)
        bar_h = Inches(0.08)

        for idx, (ctitle, cbullets) in enumerate(cols):
            x = int(fl + idx * (col_w + gap))
            tbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, col_w, bar_h)
            tbar.fill.solid()
            tbar.fill.fore_color.rgb = RGBColor(*(accent if idx == 0 else accent_dark))
            tbar.line.fill.background()
            strip_shadow(tbar)

            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y + bar_h, col_w, col_h - bar_h)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(*card_bg)
            card.line.color.rgb = RGBColor(*card_border)
            card.line.width = Pt(1.0)
            strip_shadow(card)

            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.margin_left = Inches(0.20)
            tf.margin_right = Inches(0.16)
            tf.margin_top = Inches(0.16)
            tf.margin_bottom = Inches(0.14)

            # Adaptive bullet font size inside card (default to smallest if all candidates exceed height)
            w_pt = max(120.0, (col_w - Inches(0.36)) / EMU_IN * 72.0)
            h_pt = max(100.0, (col_h - Inches(0.75)) / EMU_IN * 72.0)
            norm_bullets = [item.get("text", "") if isinstance(item, dict) else str(item) for item in cbullets]
            b_sz = 11.5
            for cand in (16.5, 15.5, 14.5, 13.5, 12.5, 11.5):
                need = sum(max(1, math.ceil(char_units(f"•  {tx}") * cand / w_pt)) * cand * 1.28 + 8 for tx in norm_bullets)
                if need <= h_pt * 0.92:
                    b_sz = cand
                    break

            first = True
            if ctitle:
                p0 = tf.paragraphs[0]
                first = False
                p0.alignment = PP_ALIGN.LEFT
                p0.space_after = Pt(10)
                rt = p0.add_run()
                rt.text = polish_math_text(str(ctitle))
                rt.font.size = Pt(17)
                rt.font.bold = True
                rt.font.color.rgb = RGBColor(*claim_prefix_rgb)
                apply_font(rt, rt.text)

            for txt in norm_bullets:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(8)
                rb_dot = p.add_run()
                rb_dot.text = "•  "
                rb_dot.font.size = Pt(b_sz)
                rb_dot.font.bold = True
                rb_dot.font.color.rgb = RGBColor(*(claim_prefix_rgb if dark_bg else accent))
                apply_font(rb_dot, txt)
                add_bullet_runs(p, txt, sz_pt=b_sz, color_rgb=text_primary, bold_lead=True)
        return y + col_h

    def fit_bullet_font_size(b_shape, bullets, default_sz, has_split, role="default"):
        """Normalize oversized master sizes (>20pt) and shrink only when needed to prevent overflow."""
        if not bullets or b_shape is None or not b_shape.width or not b_shape.height:
            return None
        if role in ("title", "thanks"):
            return None
        w_pt = max(80.0, (b_shape.width - 182880) / EMU_IN * 72.0)
        h_pt = max(60.0, (b_shape.height - 91440) / EMU_IN * 72.0)

        def est_need(sz):
            tot = 0.0
            for item in bullets:
                tx = item.get("text", "") if isinstance(item, dict) else str(item)
                if not tx:
                    continue
                lines = max(1, math.ceil(char_units(tx) * sz / w_pt))
                tot += lines * sz * 1.26 + 6.0
            return tot

        # Cap oversized default master body font across ALL content slides (consistent academic hierarchy)
        cap = 18.5 if has_split else 20.0
        start_sz = min(default_sz, cap) if default_sz > cap else default_sz
        if start_sz == default_sz and est_need(default_sz) <= h_pt * 0.94:
            return None
        for cand in (start_sz, 18.0, 16.5, 15.0, 13.5, 12.5, 11.5):
            if cand <= start_sz and est_need(cand) <= h_pt * 0.94:
                return cand
        return 11.0

    def fill_body(slide, b, bullets, top_override=None, is_reuse=False, has_right=False, has_left=False, has_bottom=False, bottom_cap_y=None, role="default", frame=None):
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        safe_bot = foot_y - Inches(0.10)
        eff_bot_cap = int(min(safe_bot - Inches(1.2), bottom_cap_y)) if (has_bottom and bottom_cap_y) else int(SH * 0.51)
        if is_reuse and bullets is None:
            return b
        if not bullets:
            # Remove unused empty body placeholder when replaced by steps/two_col/full visual
            if b is not None and not is_reuse and role not in ("title", "thanks"):
                remove_shape(b)
            return None
        created_textbox = False
        has_split = bool(has_right or has_left)
        if b is None:
            top = int(top_override + Inches(0.10)) if top_override else Inches(1.68)
            if has_left and not has_right:
                bl_x = int(fl + fw * 0.51)
                w = int(fl + fw - bl_x)
            else:
                bl_x = fl
                w = int(fw * 0.47) if has_right else fw
            max_bot = eff_bot_cap if has_bottom else safe_bot
            h = max(Inches(0.9), max_bot - top)
            tb = slide.shapes.add_textbox(bl_x, top, w, h)
            tf = tb.text_frame
            tf.word_wrap = True
            b = tb
            created_textbox = True
        else:
            if top_override is not None and role not in ("title", "thanks"):
                new_top = int(top_override + Inches(0.10))
                if b.top is None or b.top < new_top:
                    set_geom(b, top=new_top, height=max(int(Inches(0.8)), int(safe_bot - new_top)))
            if has_left and not has_right and b.left is not None and b.width:
                rl_x = int(fl + fw * 0.51)
                set_geom(b, left=rl_x, width=max(int(Inches(2.2)), int(fl + fw - rl_x)))
            elif has_right and b.left is not None and b.width:
                keep = int(fl + fw * 0.49)
                if b.left < keep and b.width > keep - b.left:
                    set_geom(b, width=int(keep - b.left))
            elif not has_split and not is_reuse and role not in ("title", "thanks") and b.left is not None and b.width:
                # If a two-column layout placeholder sits alone without a right visual, expand cleanly to full width
                if abs(b.left - fl) <= Inches(0.25) and b.width < int(fw * 0.72):
                    set_geom(b, left=fl, width=fw)
            if has_bottom and b.top is not None:
                if b.height is None or b.top + b.height > eff_bot_cap:
                    set_geom(b, height=int(max(Inches(0.8), eff_bot_cap - b.top)))
            if b.top is not None and b.height is not None and b.top + b.height > safe_bot:
                set_geom(b, height=int(max(Inches(0.8), safe_bot - b.top)))
            tf = b.text_frame
            tf.clear()

        target_sz = fit_bullet_font_size(b, bullets, body_sz, has_split, role=role)
        sparse_plain = (not has_split and not has_bottom and role not in ("title", "thanks") and len(bullets or []) <= 4)
        # Optical vertical balance on split slides (e.g. 2-3 bullets beside a tall right figure/table)
        if has_split and not has_bottom and (created_textbox or not is_reuse) and role not in ("title", "thanks") and 1 <= len(bullets or []) <= 4:
            try:
                w_pt_s = max(80.0, (b.width - 182880) / EMU_IN * 72.0) if b.width else 320.0
                h_pt_s = max(60.0, (b.height - 91440) / EMU_IN * 72.0) if b.height else 300.0
                sz_s = target_sz if target_sz is not None else min(body_sz, 18.0)
                est_s = sum(
                    max(1, math.ceil(char_units(it.get("text", "") if isinstance(it, dict) else str(it)) * sz_s / w_pt_s)) * sz_s * 1.32 + 10.0
                    for it in (bullets or [])
                )
                if est_s < h_pt_s * 0.65:
                    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
        first = True
        for item in bullets or []:
            if isinstance(item, dict):
                text, level = item.get("text", ""), int(item.get("level", 0))
            else:
                text, level = str(item), 0
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = level
            if has_bottom and len(bullets or []) >= 3:
                p.space_after = Pt(5)
            else:
                p.space_after = Pt(12 if (sparse_plain or (has_split and not has_bottom and len(bullets or []) <= 3)) else 8)
            if role in ("title", "thanks") and not is_reuse:
                p.alignment = PP_ALIGN.CENTER
            eff_sz = max(11.0, target_sz - level * 1.5) if target_sz is not None else None
            eff_col = (text_muted if role in ("title", "thanks") else text_primary) if dark_bg else None
            if created_textbox and role not in ("title", "thanks") and text and not text.lstrip().startswith(("•", "-", "▪", "·", "▸")):
                rb_dot = p.add_run()
                rb_dot.text = ("  " * level) + ("▸  " if level == 0 else "•  ")
                if eff_sz is not None:
                    rb_dot.font.size = Pt(eff_sz)
                rb_dot.font.bold = True
                rb_dot.font.color.rgb = RGBColor(*(claim_prefix_rgb if dark_bg else accent))
            add_bullet_runs(p, text, sz_pt=eff_sz, color_rgb=eff_col, bold_lead=(role not in ("title", "thanks")))
        return b

    def visual_box(region, top_anchor=None, frame=None, split_y=None):
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        top0 = int(top_anchor + Inches(0.12)) if top_anchor else int(Inches(1.68))
        # Reserve space for caption (0.32") + gap above evidence footer so caption & footer NEVER collide
        max_vis_bot = foot_y - Inches(0.40)
        left_w = max(Inches(2.2), int(fw * 0.49))
        right_l = int(fl + fw * 0.51)
        right_w = max(Inches(2.2), int(fl + fw - right_l))
        full_h = max(Inches(1.2), max_vis_bot - top0)
        if region == "left":
            return fl, top0, left_w, full_h
        if region == "right":
            return right_l, top0, right_w, full_h
        if region == "bottom":
            bt = int(split_y) if split_y is not None else max(int(SH * 0.54), top0)
            return fl, bt, fw, max(Inches(1.1), max_vis_bot - bt)
        if region == "bottom_left":
            bt = int(split_y) if split_y is not None else max(int(SH * 0.54), top0)
            return fl, bt, left_w, max(Inches(1.1), max_vis_bot - bt)
        if region == "bottom_right":
            bt = int(split_y) if split_y is not None else max(int(SH * 0.54), top0)
            return right_l, bt, right_w, max(Inches(1.1), max_vis_bot - bt)
        if region in ("bottom_third_0", "bottom_third_1", "bottom_third_2"):
            k_idx = int(region[-1])
            bt = int(split_y) if split_y is not None else max(int(SH * 0.54), top0)
            g3 = Inches(0.18)
            w3 = max(Inches(1.8), int((fw - 2 * g3) / 3))
            return int(fl + k_idx * (w3 + g3)), bt, w3, max(Inches(1.1), max_vis_bot - bt)
        if region == "top":
            mid_gap = Inches(0.42)
            hh = max(Inches(1.1), int((max_vis_bot - top0 - mid_gap) / 2))
            return fl, top0, fw, hh
        if region == "bottom_half":
            mid_gap = Inches(0.42)
            hh = max(Inches(1.1), int((max_vis_bot - top0 - mid_gap) / 2))
            bt = int(top0 + hh + mid_gap)
            return fl, bt, fw, max(Inches(1.1), max_vis_bot - bt)
        if region == "right_top":
            mid_gap = Inches(0.42)
            hh = max(Inches(1.0), int((max_vis_bot - top0 - mid_gap) / 2))
            return right_l, top0, right_w, hh
        if region == "right_bottom":
            mid_gap = Inches(0.42)
            hh = max(Inches(1.0), int((max_vis_bot - top0 - mid_gap) / 2))
            bt = int(top0 + hh + mid_gap)
            return right_l, bt, right_w, max(Inches(1.0), max_vis_bot - bt)
        return fl, top0, fw, full_h

    def bind_caption(raw_cap, evidence, kind="figure"):
        cap = (raw_cap or "").strip()
        locs = [e.get("locator", "") for e in (evidence or []) if e.get("locator")]
        kws = ("Table", "Tab", "表", "Fig", "图") if kind == "table" else ("Fig", "fig", "图", "Table", "表")
        hit = next((l for l in locs if any(k in l for k in kws)), "")
        key = ""
        if hit:
            mm = re.search(r"(Fig(?:ure)?\.?\s*[S]?\d+[a-z]?|图\s*\d+|Tab(?:le)?\.?\s*(?:[S]?\d+|[IVX]+)|表\s*\d+)", hit, re.I)
            key = mm.group(1) if mm else hit
        already_cited = bool(re.search(r"(?:Fig(?:ure)?\.?\s*[S]?\d+|图\s*\d+|Tab(?:le)?\.?\s*(?:[S]?\d+|[IVX]+)|表\s*\d+|§\s*\d+|Sec\.?\s*\d+)", cap, re.I))
        if cap and hit and not already_cited and key.lower() not in cap.lower():
            cap = f"{cap}（{hit}）" if is_cjk(cap) else f"{cap} ({hit})"
        elif not cap and hit:
            cap = hit
        return cap

    def add_caption_box(slide, left, top, width, cap, frame=None):
        if not cap:
            return
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        cw = int(min(fw, width if width >= Inches(2.0) else max(width, Inches(2.2))))
        cl = int(max(fl, min(left, fl + fw - cw)))
        cap_txt = polish_math_text(str(cap).strip())
        if len(cap_txt) > 135:
            cap_txt = cap_txt[:132].rsplit(" ", 1)[0] + "…"
        usable_w_pt = max(120.0, (cw - 182880) / EMU_IN * 72.0)
        cap_pt = 10.0
        for cand_pt in (10.0, 9.0, 8.5):
            n_ln = max(1, math.ceil(char_units(cap_txt) * cand_pt / usable_w_pt))
            if n_ln <= 2 and (n_ln * cand_pt * 1.22) <= 22.0:
                cap_pt = cand_pt
                break
            cap_pt = cand_pt
        n_ln_f = max(1, math.ceil(char_units(cap_txt) * cap_pt / usable_w_pt))
        box_h = int(min(Inches(0.36), max(Inches(0.28), Inches((n_ln_f * cap_pt * 1.28 + 6.0) / 72.0))))
        y = min(int(top + Inches(0.03)), int(foot_y - box_h - Inches(0.03)))
        cb = slide.shapes.add_textbox(cl, y, cw, box_h)
        ctf = cb.text_frame
        ctf.word_wrap = True
        ctf.margin_top = ctf.margin_bottom = 0
        cr = ctf.paragraphs[0].add_run()
        cr.text = cap_txt
        cr.font.size = Pt(cap_pt)
        cr.font.italic = True
        cr.font.color.rgb = RGBColor(*text_muted)
        apply_font(cr, cr.text)

    def add_figure(slide, fig, region, evidence, top_anchor=None, clear_old=False, frame=None, split_y=None, only_orig_els=None):
        p = resolve_figure_path(fig.get("path", ""), work)
        if not p or not p.exists():
            log["warnings"].append(f"figure missing on disk, skipped: {fig.get('path')}")
            return
        box = visual_box(region, top_anchor, frame=frame, split_y=split_y)
        if clear_old:
            clear_overlapping_visuals(slide, box, only_orig_els=only_orig_els)
        l, t, w, h = box
        # On dark backgrounds, transparent PNGs with dark axes become invisible unless backed by a light mat
        needs_light_mat = False
        if dark_bg:
            try:
                from PIL import Image as _PILImage
                with _PILImage.open(str(p)) as _im:
                    if _im.mode in ("RGBA", "LA") or ("transparency" in _im.info):
                        _rgba = _im.convert("RGBA")
                        if _rgba.getextrema()[3][0] < 245:
                            needs_light_mat = True
            except Exception:
                pass
        pic = slide.shapes.add_picture(str(p), l, t, width=w)
        if pic.height > h:
            r = h / pic.height
            pic.height = int(pic.height * r)
            pic.width = int(pic.width * r)
        pic.left = l + (w - pic.width) // 2
        pic.top = t + (h - pic.height) // 2
        if needs_light_mat:
            pad_m = int(Inches(0.06))
            mat_l = max(l, pic.left - pad_m)
            mat_t = max(t, pic.top - pad_m)
            mat_w = min(w, pic.width + 2 * pad_m)
            mat_h = min(h, pic.height + 2 * pad_m)
            mat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, mat_l, mat_t, mat_w, mat_h)
            mat.fill.solid()
            mat.fill.fore_color.rgb = RGBColor(0xF8, 0xFA, 0xFC)
            mat.line.color.rgb = RGBColor(*card_border)
            mat.line.width = Pt(0.75)
            strip_shadow(mat)
            # Move mat behind picture in XML z-order
            try:
                pic._element.addprevious(mat._element)
            except Exception:
                pass
        cap = bind_caption(fig.get("caption"), evidence, kind="figure")
        cap_top = (pic.top + pic.height + pad_m) if needs_light_mat else (pic.top + pic.height)
        add_caption_box(slide, l, cap_top, w, cap, frame=frame)

    def add_dual_figures(slide, figs_list, region, evidence, top_anchor=None, clear_old=False, frame=None, split_y=None, only_orig_els=None):
        """Render 2 figures side-by-side in equal aspect-ratio-locked panels with individual captions."""
        valid = []
        for fobj in (figs_list or [])[:2]:
            if not isinstance(fobj, dict):
                continue
            p = resolve_figure_path(fobj.get("path", ""), work)
            if p and p.exists():
                valid.append((p, fobj))
            else:
                log["warnings"].append(f"figure missing on disk, skipped: {fobj.get('path')}")
        if not valid:
            return
        if len(valid) == 1:
            add_figure(slide, valid[0][1], region or "full", evidence, top_anchor=top_anchor, clear_old=clear_old, frame=frame, split_y=split_y, only_orig_els=only_orig_els)
            return
        box = visual_box(region or "full", top_anchor, frame=frame, split_y=split_y)
        if clear_old:
            clear_overlapping_visuals(slide, box, only_orig_els=only_orig_els)
        l, t, w, h = box
        gap = Inches(0.24)
        panel_w = int((w - gap) / 2)
        for pi, (p, fobj) in enumerate(valid):
            pl = int(l + pi * (panel_w + gap))
            needs_light_mat = False
            if dark_bg:
                try:
                    from PIL import Image as _PILImage
                    with _PILImage.open(str(p)) as _im:
                        if _im.mode in ("RGBA", "LA") or ("transparency" in _im.info):
                            if _im.convert("RGBA").getextrema()[3][0] < 245:
                                needs_light_mat = True
                except Exception:
                    pass
            pic = slide.shapes.add_picture(str(p), pl, t, width=panel_w)
            if pic.height > h:
                r = h / pic.height
                pic.height = int(pic.height * r)
                pic.width = int(pic.width * r)
            pic.left = pl + (panel_w - pic.width) // 2
            pic.top = t + (h - pic.height) // 2
            pad_m = int(Inches(0.05))
            if needs_light_mat:
                mat = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    max(pl, pic.left - pad_m), max(t, pic.top - pad_m),
                    min(panel_w, pic.width + 2 * pad_m), min(h, pic.height + 2 * pad_m),
                )
                mat.fill.solid()
                mat.fill.fore_color.rgb = RGBColor(0xF8, 0xFA, 0xFC)
                mat.line.color.rgb = RGBColor(*card_border)
                mat.line.width = Pt(0.75)
                strip_shadow(mat)
                try:
                    pic._element.addprevious(mat._element)
                except Exception:
                    pass
            cap = str(fobj.get("caption") or "").strip() or bind_caption("", evidence, kind="figure")
            cap_top = (pic.top + pic.height + pad_m) if needs_light_mat else (pic.top + pic.height)
            add_caption_box(slide, pl, cap_top, panel_w, cap, frame=frame)

    def add_table(slide, tbl, region, evidence, top_anchor=None, clear_old=False, frame=None, split_y=None, only_orig_els=None):
        headers, rows = tbl.get("headers") or [], tbl.get("rows") or []
        if not rows:
            log["warnings"].append("table spec has no rows; skipped")
            return
        box = visual_box(region, top_anchor, frame=frame, split_y=split_y)
        if clear_old:
            clear_overlapping_visuals(slide, box, only_orig_els=only_orig_els)
        l, tp, w, h = box
        ncols = len(headers) or max((len(r) for r in rows), default=1) or 1
        nrows = len(rows) + (1 if headers else 0)
        all_rows = ([headers] if headers else []) + list(rows)
        # Proportional column widths with unbroken-word floor so headers/metrics never break mid-word
        def _max_word_units(s_val):
            ws = re.findall(r"[A-Za-z0-9%./+\-−]+", str(s_val or ""))
            return max((char_units(tok) for tok in ws if len(tok) <= 14), default=0.0)

        col_weights = []
        col_word_floors = []
        for ci in range(ncols):
            mx = max((char_units(str(r[ci])) if ci < len(r) else 0.0) for r in all_rows)
            mw = max((_max_word_units(r[ci]) if ci < len(r) else 0.0) for r in all_rows)
            col_word_floors.append(mw)
            col_weights.append(max(3.8, mw * 1.35 + 1.6, min(18.0, mx + 1.8)))
        w_sum = sum(col_weights) or 1.0
        col_widths = [int(w * (cw / w_sum)) for cw in col_weights]
        if col_widths:
            col_widths[0] += int(w - sum(col_widths))

        # Auto-shrink table fonts so total wrapped height NEVER overflows box h AND header words never break mid-word
        def est_rows(hp, cp):
            rhs = []
            word_breaks = False
            for r_idx, r_vals in enumerate(all_rows):
                is_h = bool(headers and r_idx == 0)
                pt_sz = hp if is_h else cp
                max_lines = 1
                for ci in range(ncols):
                    txt = str(r_vals[ci]) if ci < len(r_vals) else ""
                    cw_pt = max(34.0, (col_widths[ci] - Inches(0.18)) / EMU_IN * 72.0)
                    if is_h and _max_word_units(txt) * (pt_sz * 1.08) > cw_pt:
                        word_breaks = True
                    lines = max(1, math.ceil(char_units(txt) * pt_sz / cw_pt))
                    max_lines = max(max_lines, lines)
                rhs.append(int(max(Inches(0.36), Inches(max_lines * (pt_sz * 1.38 / 72.0) + 0.14))))
            return rhs, word_breaks

        hdr_pt, cell_pt = 10.0, 9.0
        row_heights, _ = est_rows(hdr_pt, cell_pt)
        for hp, cp in ((14.0, 13.0), (13.0, 12.0), (12.0, 11.0), (11.0, 10.0), (10.0, 9.2), (9.5, 8.8)):
            cand_rh, wb = est_rows(hp, cp)
            if sum(cand_rh) <= h and (not wb or hp <= 9.5):
                hdr_pt, cell_pt, row_heights = hp, cp, cand_rh
                break
        if sum(row_heights) > h and sum(row_heights) > 0:
            scale = h / float(sum(row_heights))
            row_heights = [max(int(Inches(0.26)), int(rh * scale)) for rh in row_heights]
        elif nrows <= 6 and sum(row_heights) < int(h * 0.62) and sum(row_heights) > 0:
            # Give compact academic tables generous breathing room so rows look crisp and legible
            target_tot = min(int(h * 0.70), nrows * int(Inches(0.56)))
            if target_tot > sum(row_heights):
                up_s = target_tot / float(sum(row_heights))
                row_heights = [min(int(Inches(0.62)), int(rh * up_s)) for rh in row_heights]
        hdr_font, cell_font = Pt(hdr_pt), Pt(cell_pt)
        tbl_h = min(h, sum(row_heights))
        # Optical vertical balance when a compact table sits in a tall region (avoid top-heavy dead void)
        tp_eff = tp
        if region in ("full", "left", "right", "bottom_left", "bottom_right") and tbl_h < int(h * 0.72):
            tp_eff = int(tp + min(Inches(0.55), int((h - tbl_h) * 0.26)))
        gshape = slide.shapes.add_table(nrows, ncols, l, tp_eff, w, tbl_h)
        gt = gshape.table
        for ci, cw in enumerate(col_widths):
            gt.columns[ci].width = cw
        for ri_h, rh in enumerate(row_heights):
            gt.rows[ri_h].height = rh
        hl_row = tbl.get("highlight_row")
        if hl_row is None:
            for r_i, r_vals in enumerate(rows):
                c0 = str(r_vals[0]).strip() if r_vals else ""
                if re.search(r"^(?:ours?|proposed|本方法|本文)\b|\(ours?\)|（本方法）|（本文）", c0, re.I):
                    hl_row = r_i
                    break

        hdr_text_rgb = (0x0F, 0x17, 0x2A) if luminance(accent) > 155 else (0xFF, 0xFF, 0xFF)
        row_even_bg = (0x16, 0x22, 0x34) if dark_bg else (0xFF, 0xFF, 0xFF)
        row_odd_bg = (0x1F, 0x2E, 0x45) if dark_bg else zebra_tint
        hl_bg = blend_with_black(accent, 0.58) if dark_bg else blend_with_white(accent, 0.84)
        delta_rgb = blend_with_white(accent, 0.35) if dark_bg else accent_dark
        body_cell_rgb = text_primary if dark_bg else (0x1E, 0x29, 0x3B)
        rule_hex = "%02X%02X%02X" % card_border
        strong_rule_hex = "%02X%02X%02X" % (blend_with_white(accent, 0.30) if dark_bg else accent)

        def _set_booktabs_borders(cell, is_hdr, is_last_row):
            """Apply crisp publication-grade horizontal-only (booktabs) cell borders."""
            try:
                from pptx.oxml import parse_xml
                from pptx.oxml.ns import nsdecls
                tcPr = cell._tc.get_or_add_tcPr()
                for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    for old in list(tcPr.findall(qn(tag))):
                        tcPr.remove(old)
                tcPr.append(parse_xml(f'<a:lnL {nsdecls("a")} w="0"><a:noFill/></a:lnL>'))
                tcPr.append(parse_xml(f'<a:lnR {nsdecls("a")} w="0"><a:noFill/></a:lnR>'))
                if is_hdr:
                    tcPr.append(parse_xml(f'<a:lnT {nsdecls("a")} w="15875"><a:solidFill><a:srgbClr val="{strong_rule_hex}"/></a:solidFill></a:lnT>'))
                    tcPr.append(parse_xml(f'<a:lnB {nsdecls("a")} w="15875"><a:solidFill><a:srgbClr val="{strong_rule_hex}"/></a:solidFill></a:lnB>'))
                else:
                    tcPr.append(parse_xml(f'<a:lnT {nsdecls("a")} w="0"><a:noFill/></a:lnT>'))
                    bw, bhex = ("15875", strong_rule_hex) if is_last_row else ("6350", rule_hex)
                    tcPr.append(parse_xml(f'<a:lnB {nsdecls("a")} w="{bw}"><a:solidFill><a:srgbClr val="{bhex}"/></a:solidFill></a:lnB>'))
            except Exception:
                pass

        num_cols = set()
        for ci in range(1, ncols):
            vals_c = [str(r[ci]).strip() for r in rows if ci < len(r) and str(r[ci]).strip()]
            if vals_c and sum(1 for v in vals_c if re.match(r"^[+\-−~<>≤≥]?\d", v)) * 2 >= len(vals_c):
                num_cols.add(ci)

        def _fill(cell, text, is_hdr, row_idx, col_idx):
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.03)
            cell.margin_bottom = Inches(0.03)
            cell.fill.solid()
            data_idx = row_idx - (1 if headers else 0)
            is_hl = (not is_hdr) and (hl_row is not None and data_idx == hl_row)
            if is_hdr:
                cell.fill.fore_color.rgb = RGBColor(*accent)
            elif is_hl:
                cell.fill.fore_color.rgb = RGBColor(*hl_bg)
            else:
                cell.fill.fore_color.rgb = RGBColor(*(row_odd_bg if row_idx % 2 == 1 else row_even_bg))
            _set_booktabs_borders(cell, is_hdr, row_idx == nrows - 1)
            tf = cell.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            txt = str(text)
            if col_idx in num_cols:
                p0.alignment = PP_ALIGN.RIGHT
            elif not is_hdr and col_idx > 0 and re.match(r"^[+\-−]?\d", txt.strip()):
                p0.alignment = PP_ALIGN.RIGHT
            elif is_hdr and col_idx > 0:
                p0.alignment = PP_ALIGN.CENTER
            r0 = p0.add_run()
            r0.text = polish_math_text(txt)
            r0.font.size = hdr_font if is_hdr else cell_font
            r0.font.bold = is_hdr or is_hl or txt.strip().startswith(("+", "ours", "Ours", "本方法", "本文"))
            if is_hdr:
                r0.font.color.rgb = RGBColor(*hdr_text_rgb)
            elif txt.strip().startswith("+"):
                r0.font.color.rgb = RGBColor(*delta_rgb)
            else:
                r0.font.color.rgb = RGBColor(*body_cell_rgb)
            apply_font(r0, r0.text)

        ri = 0
        if headers:
            for ci in range(ncols):
                _fill(gt.cell(0, ci), headers[ci] if ci < len(headers) else "", True, 0, ci)
            ri = 1
        for r in rows:
            for ci in range(ncols):
                _fill(gt.cell(ri, ci), r[ci] if ci < len(r) else "", False, ri, ci)
            ri += 1
        cap = bind_caption(tbl.get("caption"), evidence, kind="table")
        actual_h = sum(r.height for r in gt.rows) if gt.rows else tbl_h
        add_caption_box(slide, l, tp_eff + max(tbl_h, actual_h) + Inches(0.06), w, cap, frame=frame)

    def add_chart(slide, ch_spec, region, evidence, top_anchor=None, clear_old=False, frame=None, split_y=None, only_orig_els=None):
        """Render a native editable PowerPoint chart (bar/barh/line) from source-verified data."""
        if not isinstance(ch_spec, dict):
            return
        cats = [str(c) for c in (ch_spec.get("categories") or [])]
        series_in = [s for s in (ch_spec.get("series") or []) if isinstance(s, dict) and s.get("values")]
        if not cats or not series_in:
            log["warnings"].append("chart spec missing categories or series; skipped")
            return
        try:
            from pptx.chart.data import CategoryChartData
            from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
        except Exception as e:
            log["warnings"].append(f"chart skipped (pptx.chart unavailable): {e!r}")
            return
        box = visual_box(region, top_anchor, frame=frame, split_y=split_y)
        if clear_old:
            clear_overlapping_visuals(slide, box, only_orig_els=only_orig_els)
        l, tp, w, h = box
        if ch_spec.get("caption"):
            h = max(Inches(1.1), int(h - Inches(0.06)))
        ctype_str = str(ch_spec.get("type", "bar")).lower()
        ctype = {
            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "grouped_bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "barh": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
        }.get(ctype_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

        cd = CategoryChartData()
        cd.categories = cats
        all_vals = []
        valid_series = []
        for idx_s, s_item in enumerate(series_in[:4]):
            vals = []
            for v in s_item.get("values", [])[: len(cats)]:
                fv = 0.0
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    fv = float(v)
                else:
                    mv = re.search(r"[+\-−]?\d[\d,]*(?:\.\d+)?(?:[eE][+\-]?\d+)?", str(v or "").strip())
                    if mv:
                        try:
                            fv = float(mv.group(0).replace("−", "-").replace(",", ""))
                        except ValueError:
                            fv = 0.0
                vals.append(fv)
                all_vals.append(fv)
            while len(vals) < len(cats):
                vals.append(0.0)
            sname = str(s_item.get("name") or f"Series {idx_s + 1}")
            cd.add_series(sname, tuple(vals))
            valid_series.append(sname)

        try:
            ch_sh = slide.shapes.add_chart(ctype, l, tp, w, h, cd)
            ch = ch_sh.chart
            ch.has_title = bool(ch_spec.get("title"))
            if ch.has_title:
                ch.chart_title.text_frame.text = polish_math_text(str(ch_spec["title"]))
                for p0 in ch.chart_title.text_frame.paragraphs:
                    p0.font.size = Pt(12)
                    p0.font.bold = True
                    p0.font.color.rgb = RGBColor(*text_primary)
                    for r0 in p0.runs:
                        apply_font(r0, r0.text)
            ch.has_legend = len(valid_series) > 1
            if ch.has_legend:
                ch.legend.position = XL_LEGEND_POSITION.TOP
                ch.legend.include_in_layout = True
                ch.legend.font.size = Pt(10.5)
                ch.legend.font.color.rgb = RGBColor(*text_primary)
                leg_txt = " ".join(valid_series)
                leg_font = tpl_font if (tpl_font and (tpl_font_is_cjk or not is_cjk(leg_txt))) else (CJK_FONT if is_cjk(leg_txt) else (tpl_font or "Calibri"))
                if leg_font:
                    try:
                        ch.legend.font.name = leg_font
                    except Exception:
                        pass

            # Harmonious series palette derived from template accent1 (protagonist/Ours highlighted, baselines muted)
            muted_s = (0x64, 0x74, 0x8B) if dark_bg else (0x94, 0xA3, 0xB8)
            alt_s = blend_with_white(accent, 0.35) if dark_bg else blend_with_white(accent, 0.48)
            deep_s = blend_with_white(accent, 0.58) if dark_bg else accent_dark
            ours_idx = next((idx_n for idx_n, nm in enumerate(valid_series)
                             if re.search(r"\b(ours?|proposed)\b|本方法|本文", nm, re.I)), None)
            if len(valid_series) == 1:
                s_colors = [accent]
            elif ours_idx is not None:
                others_pal = [muted_s, alt_s, deep_s]
                s_colors, oi = [], 0
                for idx_n in range(len(valid_series)):
                    if idx_n == ours_idx:
                        s_colors.append(accent)
                    else:
                        s_colors.append(others_pal[oi % len(others_pal)])
                        oi += 1
            elif len(valid_series) == 2:
                s_colors = [muted_s, accent]
            else:
                s_colors = [muted_s, alt_s, accent, deep_s]

            plot = ch.plots[0]
            if ctype != XL_CHART_TYPE.LINE_MARKERS:
                try:
                    plot.gap_width = 75
                except Exception:
                    pass
            show_labels = ch_spec.get("data_labels", len(cats) * len(valid_series) <= 16)
            plot.has_data_labels = bool(show_labels)
            if plot.has_data_labels:
                plot.data_labels.font.size = Pt(9.5)
                plot.data_labels.font.color.rgb = RGBColor(*text_primary)
            for s_obj, rgb in zip(ch.series, s_colors):
                if ctype == XL_CHART_TYPE.LINE_MARKERS:
                    s_obj.format.line.color.rgb = RGBColor(*rgb)
                    s_obj.format.line.width = Pt(2.2)
                else:
                    s_obj.format.fill.solid()
                    s_obj.format.fill.fore_color.rgb = RGBColor(*rgb)
            # Single-series bar/barh: highlight 'Ours'/protagonist category in accent & mute baseline bars
            if len(valid_series) == 1 and ctype != XL_CHART_TYPE.LINE_MARKERS:
                hl_cat = ch_spec.get("highlight_category")
                ours_cat_i = next(
                    (ci for ci, cn in enumerate(cats)
                     if (hl_cat and str(hl_cat).lower() in cn.lower())
                     or re.search(r"\b(ours?|proposed)\b|本方法|本文", cn, re.I)),
                    None,
                )
                if ours_cat_i is not None and len(cats) >= 2:
                    s0 = ch.series[0]
                    for ci in range(len(cats)):
                        pt = s0.points[ci]
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.rgb = RGBColor(*(accent if ci == ours_cat_i else muted_s))
        except Exception as e_ch:
            log["warnings"].append(f"native chart fallback->table ({e_ch!r})")
            fb_rows = [[str(s_item.get("name", f"S{i_s+1}"))] + [str(v) for v in (s_item.get("values") or [])[:len(cats)]]
                       for i_s, s_item in enumerate(series_in[:4])]
            add_table(slide, {"headers": ["Series"] + cats, "rows": fb_rows, "caption": ch_spec.get("caption")},
                      region, evidence, top_anchor=top_anchor, clear_old=clear_old, frame=frame)
            return

        try:
            chart_all_txt = " ".join(cats + valid_series + [str(ch_spec.get("y_label") or ""), str(ch_spec.get("x_label") or "")])
            ch_font_name = (tpl_font if (tpl_font and (tpl_font_is_cjk or not is_cjk(chart_all_txt)))
                            else (CJK_FONT if is_cjk(chart_all_txt) else (tpl_font or "Calibri")))
            ca = ch.category_axis
            ca.tick_labels.font.size = Pt(10)
            ca.tick_labels.font.color.rgb = RGBColor(*text_primary)
            if ch_font_name:
                ca.tick_labels.font.name = ch_font_name
            va = ch.value_axis
            va.tick_labels.font.size = Pt(9.5)
            va.tick_labels.font.color.rgb = RGBColor(*text_muted)
            if ch_font_name:
                va.tick_labels.font.name = ch_font_name
            if plot.has_data_labels and ch_font_name:
                plot.data_labels.font.name = ch_font_name
            va.has_major_gridlines = True
            grid_rgb = (0x33, 0x41, 0x55) if dark_bg else (0xE2, 0xE8, 0xF0)
            axis_rgb = (0x47, 0x55, 0x69) if dark_bg else (0xCB, 0xD5, 0xE1)
            va.major_gridlines.format.line.color.rgb = RGBColor(*grid_rgb)
            try:
                ca.format.line.color.rgb = RGBColor(*axis_rgb)
                va.format.line.color.rgb = RGBColor(*axis_rgb)
            except Exception:
                pass
            if ch_spec.get("y_min") is not None:
                va.minimum_scale = float(ch_spec["y_min"])
            elif ctype != XL_CHART_TYPE.LINE_MARKERS and all_vals and min(all_vals) >= 0:
                if max(all_vals) > 0 and min(all_vals) <= 0.72 * max(all_vals):
                    va.minimum_scale = 0.0
            if ch_spec.get("y_max") is not None:
                va.maximum_scale = float(ch_spec["y_max"])
            elif plot.has_data_labels and all_vals and max(all_vals) > 0 and ctype != XL_CHART_TYPE.BAR_CLUSTERED:
                mx_v = max(all_vals)
                base_min = float(ch_spec["y_min"]) if ch_spec.get("y_min") is not None else (0.0 if min(all_vals) >= 0 else min(all_vals))
                span_v = max(1e-6, mx_v - base_min)
                raw_top = mx_v + (span_v * 0.26 if ch_spec.get("y_min") is not None else mx_v * 0.18)
                mag = 10 ** math.floor(math.log10(max(1e-6, span_v)))
                step = mag * (0.2 if span_v / mag <= 2 else (0.5 if span_v / mag <= 5 else 1.0))
                nice_max = math.ceil(raw_top / step) * step
                if nice_max > mx_v + span_v * 0.10:
                    va.maximum_scale = float(round(nice_max, 6))
            if ch_spec.get("y_label"):
                va.has_title = True
                va.axis_title.text_frame.text = polish_math_text(str(ch_spec["y_label"]))
                va.axis_title.text_frame.paragraphs[0].font.size = Pt(9.5)
                va.axis_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(*text_muted)
                if ch_font_name:
                    va.axis_title.text_frame.paragraphs[0].font.name = ch_font_name
            if ch_spec.get("x_label"):
                ca.has_title = True
                ca.axis_title.text_frame.text = polish_math_text(str(ch_spec["x_label"]))
                ca.axis_title.text_frame.paragraphs[0].font.size = Pt(9.5)
                ca.axis_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(*text_muted)
                if ch_font_name:
                    ca.axis_title.text_frame.paragraphs[0].font.name = ch_font_name
        except Exception:
            pass

        cap = bind_caption(ch_spec.get("caption"), evidence, kind="table")
        add_caption_box(slide, l, tp + h + Inches(0.03), w, cap, frame=frame)

    def add_callout(slide, callout_spec, frame=None, clear_old=False, only_orig_els=None):
        """Render a sleek bottom-anchored highlight box (key formula, core insight, or decision needed); returns top EMU."""
        if not callout_spec:
            return None
        if isinstance(callout_spec, str):
            ctitle, ctext = "", callout_spec.strip()
        elif isinstance(callout_spec, dict):
            ctitle = str(callout_spec.get("title") or callout_spec.get("label") or "").strip()
            ctext = str(callout_spec.get("text") or callout_spec.get("body") or callout_spec.get("detail") or "").strip()
        else:
            return None
        if not (ctitle or ctext):
            return None
        fl, fw, foot_y = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        w_pt = max(160.0, (fw - Inches(0.46)) / EMU_IN * 72.0)
        full_line = (f"{ctitle}: {ctext}" if (ctitle and ctext) else (ctitle or ctext))
        pt_sz = 13.5
        n_ln = max(1, math.ceil(char_units(full_line) * pt_sz / w_pt))
        if n_ln > 2:
            pt_sz = 12.0
            n_ln = max(1, math.ceil(char_units(full_line) * pt_sz / w_pt))
        box_h = int(min(Inches(1.08), max(Inches(0.56), Inches((n_ln * pt_sz * 1.32 + 16.0) / 72.0))))
        y = int(foot_y - box_h - Inches(0.14))
        if clear_old:
            clear_overlapping_visuals(slide, (fl, y, fw, box_h), only_orig_els=only_orig_els)
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, fl, y, fw, box_h)
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(*accent_tint)
        box.line.color.rgb = RGBColor(*card_border)
        box.line.width = Pt(1.0)
        strip_shadow(box)
        stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, fl, y, Inches(0.07), box_h)
        stripe.fill.solid()
        stripe.fill.fore_color.rgb = RGBColor(*accent)
        stripe.line.fill.background()
        strip_shadow(stripe)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.20)
        tf.margin_right = Inches(0.16)
        tf.margin_top = Inches(0.06)
        tf.margin_bottom = Inches(0.06)
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        if ctitle:
            rt = p0.add_run()
            rt.text = polish_math_text(ctitle) + ("：" if is_cjk(ctitle + ctext) else ": ") if ctext else polish_math_text(ctitle)
            rt.font.size = Pt(pt_sz)
            rt.font.bold = True
            rt.font.color.rgb = RGBColor(*claim_prefix_rgb)
            apply_font(rt, rt.text)
        if ctext:
            rx = p0.add_run()
            rx.text = polish_math_text(ctext)
            rx.font.size = Pt(pt_sz)
            rx.font.bold = not bool(ctitle)
            rx.font.color.rgb = RGBColor(*text_primary)
            apply_font(rx, rx.text)
        return y

    def add_bullet_cards(slide, bullets, top_y, bottom_y, frame=None):
        """Render 2-4 standalone level-0 bullets as balanced horizontal left-accent insight cards (eliminates dead whitespace)."""
        items = [(it.get("text", "") if isinstance(it, dict) else str(it)).strip() for it in (bullets or [])]
        items = [x for x in items if x]
        if not (2 <= len(items) <= 4):
            return False
        fl, fw, _ = frame if frame else (int(MARGIN), int(SW - 2 * MARGIN), int(SH - Inches(0.36)))
        y0 = int(top_y + Inches(0.14))
        avail_h = int(bottom_y - y0 - Inches(0.10))
        if avail_h < int(Inches(2.15)):
            return False
        n = len(items)
        gap = Inches(0.14)
        w_pt = max(160.0, (fw - Inches(0.46)) / EMU_IN * 72.0)
        chosen_sz, card_hs = None, None
        for cand_sz in (17.0, 15.5, 14.5, 13.5, 12.5):
            hs = []
            for tx in items:
                n_ln = max(1, math.ceil(char_units(tx) * cand_sz / w_pt))
                hs.append(int(max(Inches(0.68), Inches((n_ln * cand_sz * 1.38 + 22.0) / 72.0))))
            tot = sum(hs) + int(gap) * (n - 1)
            if tot <= avail_h:
                chosen_sz, card_hs = cand_sz, hs
                break
        if chosen_sz is None:
            return False
        tot_h = sum(card_hs) + int(gap) * (n - 1)
        if tot_h < int(avail_h * 0.72):
            extra_each = min(int(Inches(0.30)), int((avail_h * 0.74 - tot_h) / n))
            if extra_each > 0:
                card_hs = [h_i + extra_each for h_i in card_hs]
                tot_h = sum(card_hs) + int(gap) * (n - 1)
        y_cur = y0 + (min(int(Inches(0.42)), int((avail_h - tot_h) * 0.30)) if avail_h > tot_h + Inches(0.35) else 0)
        for tx, ch_i in zip(items, card_hs):
            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, fl, y_cur, fw, ch_i)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(*card_bg)
            card.line.color.rgb = RGBColor(*card_border)
            card.line.width = Pt(0.85)
            strip_shadow(card)
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, fl, y_cur, Inches(0.06), ch_i)
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(*accent)
            bar.line.fill.background()
            strip_shadow(bar)
            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.20)
            tf.margin_right = Inches(0.16)
            tf.margin_top = Inches(0.06)
            tf.margin_bottom = Inches(0.06)
            p0 = tf.paragraphs[0]
            p0.alignment = PP_ALIGN.LEFT
            add_bullet_runs(p0, tx, sz_pt=chosen_sz, color_rgb=text_primary, bold_lead=True)
            y_cur += ch_i + int(gap)
        return True

    def set_notes(slide, notes, claim, time_s):
        txt = notes or ""
        if not txt and claim:
            txt = (f"讲解：本页主张——{claim}。指向页面证据展开说明；避免读幻灯片原文。" if is_cjk(claim)
                   else f"Talk track: page claim — {claim}. Point to the on-page evidence; do not read the slide verbatim.")
            log["warnings"].append("notes missing in spec; generated from claim (fill before delivery)")
        if meta.get("time_annotations", True) and time_s:
            prefix = f"⏱ 约{time_s}秒 | " if is_cjk(txt or claim) else f"[{time_s}s] "
            txt = prefix + txt
        slide.notes_slide.notes_text_frame.text = txt

    timed = [s for s in slides if s.get("role") not in ("title", "thanks", "qa_backup")]
    dur_s = float(meta.get("duration_minutes", 10)) * 60.0
    per = max(30.0, round((dur_s - 25.0) / max(1, len(timed))))
    orig_slides_count = len(prs.slides)
    orig_sld_ids = list(prs.slides._sldIdLst)
    spec_sld_ids = []
    touched_orig_indices = set()

    for i, s in enumerate(slides, 1):
        role = s.get("role", "default")
        reuse = s.get("reuse") or {}
        is_reuse = False
        reuse_action = str(reuse.get("action", "edit")).lower()
        clear_visuals = bool(reuse.get("clear_visuals", False))

        if mode == "edit_deck" and isinstance(reuse.get("slide_index"), int):
            si = reuse["slide_index"]
            if not (0 <= si < orig_slides_count):
                log["warnings"].append(f"reuse slide_index {si} out of range; appending instead")
                slide = prs.slides.add_slide(pick_layout(role, s.get("layout_hint")))
                sld_id_el = prs.slides._sldIdLst[-1]
            elif si in touched_orig_indices:
                log["warnings"].append(f"s{i}: duplicate reuse slide_index {si}; cloning layout into new slide to avoid dropping content")
                slide = prs.slides.add_slide(prs.slides[si].slide_layout)
                sld_id_el = prs.slides._sldIdLst[-1]
            else:
                slide = prs.slides[si]
                sld_id_el = orig_sld_ids[si]
                touched_orig_indices.add(si)
                is_reuse = True
        else:
            slide = prs.slides.add_slide(pick_layout(role, s.get("layout_hint")))
            sld_id_el = prs.slides._sldIdLst[-1]

        spec_sld_ids.append(sld_id_el)
        orig_shape_els = {sh._element for sh in slide.shapes} if is_reuse else set()
        claim = s.get("claim", "")
        fig = s.get("figure")
        figs_list = s.get("figures") if isinstance(s.get("figures"), list) and len(s.get("figures")) >= 2 else None
        if not fig and isinstance(s.get("figures"), list) and len(s.get("figures")) == 1:
            fig = s["figures"][0]
        tbl = s.get("table")
        ch_spec = s.get("chart")
        steps = s.get("steps")
        kpis = s.get("kpis")
        two_col = s.get("two_col")
        callout = s.get("callout")

        set_slide_theme(slide)
        if not (is_reuse and reuse_action == "keep"):
            t, b = phs(slide, allow_fallback=is_reuse)
            frame = slide_frame(slide, t, b)
            # Subtle top accent hairline ONLY in no_template mode for crisp editorial finish
            if mode == "no_template" and not is_reuse:
                hbar_h = Inches(0.075) if role == "title" else Inches(0.055)
                hbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, hbar_h)
                hbar.fill.solid()
                hbar.fill.fore_color.rgb = RGBColor(*accent)
                hbar.line.fill.background()
                strip_shadow(hbar)
            fix_body_geom(b, t, role=role, is_reuse=is_reuse, frame=frame)
            if not (is_reuse and not s.get("title")):
                set_title(slide, s.get("title"), t, role=role, is_reuse=is_reuse, frame=frame)
            # Subtle accent divider on template-less default cover slide
            if mode == "no_template" and role == "title" and not is_reuse and t is not None and t.top is not None and t.height:
                div_w = Inches(2.4)
                div_y = int(t.top + t.height + Inches(0.06))
                if b is None or b.top is None or div_y + Inches(0.12) < b.top:
                    dv = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, (SW - div_w) // 2, div_y, div_w, Inches(0.05))
                    dv.fill.solid()
                    dv.fill.fore_color.rgb = RGBColor(*accent)
                    dv.line.fill.background()
                    strip_shadow(dv)

            content_top = None
            claim_bar = None
            if claim and meta.get("claim_on_slide", True):
                strip = add_claim_strip(slide, claim, t, s.get("title", ""), role=role, frame=frame)
                if strip:
                    claim_bar, content_top = strip
            if content_top is None and t is not None and t.top is not None and t.height:
                content_top = int(t.top + t.height)

            # On edited reused slides, only clear full body zone when clear_visuals=True (or full-canvas two_col/steps replacement);
            # when clear_visuals=False, preserve untouched original visuals and let each newly added component clear only its own bbox.
            if is_reuse and (clear_visuals or bool(two_col)):
                fl_c, fw_c, fy_c = frame
                ct_c = int(content_top or Inches(1.48))
                clear_overlapping_visuals(
                    slide,
                    (fl_c, ct_c, fw_c, max(Inches(1.0), fy_c - ct_c - Inches(0.05))),
                    keep_shapes=[t, b, claim_bar],
                    clear_orig_els=orig_shape_els,
                    only_orig_els=orig_shape_els,
                )

            callout_top = None
            if callout and role not in ("title", "thanks"):
                callout_top = add_callout(slide, callout, frame=frame, clear_old=(is_reuse and not clear_visuals), only_orig_els=orig_shape_els)
            eff_frame = (frame[0], frame[1], min(frame[2], int(callout_top - Inches(0.08)))) if callout_top else frame

            if kpis and role not in ("title", "thanks"):
                content_top = add_kpis(slide, kpis, content_top or Inches(1.48), frame=eff_frame, clear_old=(is_reuse and not clear_visuals), only_orig_els=orig_shape_els)

            has_bullets = bool(s.get("bullets"))
            if is_reuse and b is not None and not has_bullets and any([kpis, steps, two_col, fig, figs_list, tbl, ch_spec, callout]):
                remove_shape(b)
                b = None
            if steps and role not in ("title", "thanks"):
                content_top = add_steps(slide, steps, content_top or Inches(1.48),
                                        has_below=has_bullets or bool(fig or figs_list or tbl or ch_spec),
                                        frame=eff_frame, clear_old=(is_reuse and not clear_visuals), only_orig_els=orig_shape_els)

            if two_col and role not in ("title", "thanks"):
                # two_col takes over main body area cleanly
                if b is not None:
                    remove_shape(b)
                    b = None
                add_two_col(slide, two_col, content_top or Inches(1.48), frame=eff_frame, clear_old=(is_reuse and not clear_visuals), only_orig_els=orig_shape_els)
            else:
                def norm_reg(obj):
                    if not obj:
                        return None
                    r = obj.get("region") or ("right" if has_bullets else "full")
                    return ("right" if has_bullets else "full") if r == "auto" else r

                media_slots = []
                if figs_list:
                    media_slots.append(("dual", figs_list))
                elif fig:
                    media_slots.append(("fig", fig))
                if tbl:
                    media_slots.append(("tbl", tbl))
                if ch_spec:
                    media_slots.append(("chart", ch_spec))

                reg_map = {}
                media_split_y = None
                bottom_cap_y = None
                if len(media_slots) == 1:
                    mk, mobj = media_slots[0]
                    reg_map[mk] = ("bottom" if has_bullets else "full") if mk == "dual" else norm_reg(mobj)
                elif len(media_slots) == 2:
                    (k0, _), (k1, _) = media_slots
                    if not has_bullets:
                        if k0 == "dual":
                            reg_map[k0], reg_map[k1] = "top", "bottom_half"
                        else:
                            reg_map[k0], reg_map[k1] = "left", "right"
                        log["warnings"].append(f"s{i}: 2 media items without bullets -> auto-placed side-by-side ('{reg_map[k0]}' + '{reg_map[k1]}')")
                    else:
                        _, _, fy_m = eff_frame
                        top_m = int((content_top or Inches(1.48)) + Inches(0.10))
                        max_vis_m = int(fy_m - Inches(0.40))
                        avail_m = max(Inches(2.4), max_vis_m - top_m)
                        band_h = int(max(Inches(1.05), min(Inches(1.85), avail_m * 0.33)))
                        bottom_cap_y = int(top_m + band_h)
                        media_split_y = int(bottom_cap_y + Inches(0.12))
                        reg_map[k0], reg_map[k1] = "bottom_left", "bottom_right"
                        log["warnings"].append(f"s{i}: 2 media items with bullets -> top bullet band + side-by-side ('bottom_left' + 'bottom_right')")
                elif len(media_slots) >= 3:
                    if not has_bullets:
                        for idx_m, (mk, _) in enumerate(media_slots[:3]):
                            reg_map[mk] = ("left", "right_top", "right_bottom")[idx_m]
                        log["warnings"].append(f"s{i}: 3 media items without bullets -> auto-placed ('left' + 'right_top' + 'right_bottom')")
                    else:
                        _, _, fy_m = eff_frame
                        top_m = int((content_top or Inches(1.48)) + Inches(0.10))
                        max_vis_m = int(fy_m - Inches(0.40))
                        avail_m = max(Inches(2.4), max_vis_m - top_m)
                        band_h = int(max(Inches(1.00), min(Inches(1.70), avail_m * 0.31)))
                        bottom_cap_y = int(top_m + band_h)
                        media_split_y = int(bottom_cap_y + Inches(0.12))
                        for idx_m, (mk, _) in enumerate(media_slots[:3]):
                            reg_map[mk] = f"bottom_third_{idx_m}"
                        log["warnings"].append(f"s{i}: 3 media items with bullets -> top bullet band + 3-col bottom strip")

                fig_reg = reg_map.get("fig")
                dual_reg = reg_map.get("dual")
                tbl_reg = reg_map.get("tbl")
                ch_reg = reg_map.get("chart")
                used_regs = list(reg_map.values())

                has_right = any(r in ("right", "right_top", "right_bottom") for r in used_regs)
                has_left = any(r == "left" for r in used_regs)
                has_bottom = any(r in ("bottom", "bottom_left", "bottom_right", "bottom_half", "bottom_third_0", "bottom_third_1", "bottom_third_2") for r in used_regs)
                # In edit_deck mode without clear_visuals, detect preserved original visual (right/left/bottom) so bullets & Source footer never overlap it
                if is_reuse and not clear_visuals and not used_regs:
                    fl_p, fw_p, fy_p = frame
                    ct_p = int(content_top or Inches(1.48))
                    for _sh in list(slide.shapes):
                        if _sh._element not in orig_shape_els:
                            continue
                        is_pvis = (
                            _sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP)
                            or getattr(_sh, "has_table", False)
                            or getattr(_sh, "has_chart", False)
                        )
                        if not is_pvis or None in (_sh.left, _sh.top, _sh.width, _sh.height):
                            continue
                        if (_sh.width > int(SW * 0.88) and _sh.height > int(SH * 0.78)) or (_sh.shape_type == MSO_SHAPE_TYPE.PICTURE and (_sh.width * _sh.height) >= 0.65 * SW * SH):
                            continue
                        if _sh.top >= int(fy_p - Inches(0.45)):
                            continue
                        # Gently fit any preserved picture above evidence footer (aspect-ratio-locked) so Source bar never overlaps it
                        max_pic_bot = int(fy_p - Inches(0.08))
                        if _sh.shape_type == MSO_SHAPE_TYPE.PICTURE and (_sh.top + _sh.height) > max_pic_bot and _sh.height > Inches(0.6):
                            new_h = max(int(Inches(0.8)), int(max_pic_bot - _sh.top))
                            sc_r = new_h / float(_sh.height)
                            cx_old = _sh.left + _sh.width // 2
                            _sh.height = new_h
                            _sh.width = int(_sh.width * sc_r)
                            _sh.left = max(int(MARGIN), int(cx_old - _sh.width // 2))
                        if _sh.left >= int(fl_p + fw_p * 0.43):
                            has_right = True
                        elif (_sh.left + _sh.width) <= int(fl_p + fw_p * 0.56) and _sh.height >= int(Inches(1.4)):
                            has_left = True
                        elif _sh.top >= int(ct_p + Inches(0.20 if not has_bullets else 1.15)):
                            has_bottom = True
                            bottom_cap_y = min(bottom_cap_y, int(_sh.top - Inches(0.12))) if bottom_cap_y else int(_sh.top - Inches(0.12))
                can_use_cards = (
                    has_bullets
                    and not (has_right or has_left or has_bottom or fig or figs_list or tbl or ch_spec or steps or two_col)
                    and role not in ("title", "thanks", "qa_backup")
                    and (not is_reuse or clear_visuals or bool(kpis))
                    and 2 <= len(s.get("bullets") or []) <= 4
                    and all((not isinstance(it, dict) or int(it.get("level", 0)) == 0) for it in (s.get("bullets") or []))
                )
                used_cards = False
                if can_use_cards:
                    used_cards = add_bullet_cards(slide, s.get("bullets"), content_top or Inches(1.48), eff_frame[2], frame=eff_frame)
                    if used_cards and b is not None:
                        remove_shape(b)
                        b = None
                if not used_cards:
                    b = fill_body(slide, b, s.get("bullets"), content_top, is_reuse=is_reuse,
                                  has_right=has_right, has_left=has_left, has_bottom=has_bottom, bottom_cap_y=bottom_cap_y, role=role, frame=eff_frame)
                if media_split_y is not None and b is not None and b.top is not None and b.height is not None:
                    media_split_y = max(media_split_y, int(b.top + b.height + Inches(0.10)))
                # Subtle centered takeaway badge on template-less cover when cover claim is distinct from title & subtitle
                if mode == "no_template" and role == "title" and not is_reuse and claim and len(claim.strip()) <= 95:
                    sub_txt = " ".join((x.get("text", "") if isinstance(x, dict) else str(x)) for x in (s.get("bullets") or []))
                    if claim.strip().lower() not in (str(s.get("title") or "").lower() + " " + sub_txt.lower()):
                        anchor_bot = (b.top + b.height) if (b is not None and b.top is not None and b.height) else int(SH * 0.66)
                        pill_y = int(max(anchor_bot + Inches(0.16), SH * 0.72))
                        pill_h = Inches(0.46)
                        if pill_y + pill_h < int(SH - Inches(0.55)):
                            pill_w = int(min(SW - Inches(2.2), Inches(max(5.4, min(9.8, char_units(claim) * 13.0 / 72.0 + 0.9)))))
                            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, (SW - pill_w) // 2, pill_y, pill_w, pill_h)
                            pill.fill.solid()
                            pill.fill.fore_color.rgb = RGBColor(*accent_tint)
                            pill.line.color.rgb = RGBColor(*card_border)
                            pill.line.width = Pt(0.85)
                            strip_shadow(pill)
                            ptf = pill.text_frame
                            ptf.word_wrap = True
                            ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
                            ptf.margin_left = Inches(0.16)
                            ptf.margin_right = Inches(0.16)
                            pp0 = ptf.paragraphs[0]
                            pp0.alignment = PP_ALIGN.CENTER
                            pr0 = pp0.add_run()
                            pr0.text = polish_math_text(claim.strip())
                            pr0.font.size = Pt(12.5)
                            pr0.font.bold = True
                            pr0.font.color.rgb = RGBColor(*claim_prefix_rgb)
                            apply_font(pr0, pr0.text)
                # Clear stale original visuals (and their stale caption textboxes in the replaced zone) ONCE before inserting new media
                if is_reuse and not clear_visuals and media_slots:
                    _, _, fy_clr = eff_frame
                    for _mk, _ in media_slots:
                        _mreg = reg_map.get(_mk)
                        if _mreg:
                            _ml, _mt, _mw, _mh = visual_box(_mreg, content_top, frame=eff_frame, split_y=media_split_y)
                            _clr_h = max(_mh, int(min(fy_clr - Inches(0.02), _mt + _mh + Inches(0.42)) - _mt))
                            clear_overlapping_visuals(
                                slide, (_ml, _mt, _mw, _clr_h),
                                keep_shapes=[t, b, claim_bar],
                                clear_orig_els=orig_shape_els,
                                only_orig_els=orig_shape_els,
                            )
                if figs_list:
                    add_dual_figures(slide, figs_list, dual_reg, s.get("evidence"), content_top, clear_old=False, frame=eff_frame, split_y=media_split_y)
                elif fig:
                    add_figure(slide, fig, fig_reg, s.get("evidence"), content_top, clear_old=False, frame=eff_frame, split_y=media_split_y)
                if tbl:
                    add_table(slide, tbl, tbl_reg, s.get("evidence"), content_top, clear_old=False, frame=eff_frame, split_y=media_split_y)
                if ch_spec:
                    add_chart(slide, ch_spec, ch_reg, s.get("evidence"), content_top, clear_old=False, frame=eff_frame, split_y=media_split_y)

            if not is_reuse:
                cleanup_unused_placeholders(slide, [t, b])
            lang_cjk = is_cjk((s.get("title") or "") + (claim or ""))
            add_evidence_footer(slide, s.get("evidence"), role, lang_cjk, frame=frame)
        else:
            # keep mode: preserve slide visuals intact; add footer only if evidence explicitly supplied
            if s.get("evidence"):
                t_k, b_k = phs(slide, allow_fallback=True)
                frame_k = slide_frame(slide, t_k, b_k)
                lang_cjk = is_cjk((s.get("title") or "") + (claim or ""))
                add_evidence_footer(slide, s.get("evidence"), role, lang_cjk, frame=frame_k)

        time_s = s.get("time_s") or (None if role in ("title", "thanks", "qa_backup") else per)
        raw_notes = s.get("notes") or ""
        if not raw_notes and is_reuse:
            try:
                if slide.has_notes_slide:
                    raw_notes = (slide.notes_slide.notes_text_frame.text or "").strip()
            except Exception:
                raw_notes = ""
        if s.get("notes") or not (is_reuse and reuse_action == "keep"):
            set_notes(slide, raw_notes, claim, time_s)

        vis_tags = []
        if figs_list:
            for fl_i in figs_list[:2]:
                if isinstance(fl_i, dict):
                    vis_tags.append(f"figure:{fl_i.get('caption') or Path(str(fl_i.get('path',''))).name}")
        elif fig:
            vis_tags.append(f"figure:{fig.get('caption') or Path(str(fig.get('path',''))).name}")
        if tbl:
            vis_tags.append(f"table:{tbl.get('caption') or 'table'}")
        if ch_spec:
            vis_tags.append(f"chart:{ch_spec.get('type', 'bar')}:{ch_spec.get('caption') or 'chart'}")
        if steps:
            vis_tags.append(f"steps:{len(steps)}")
        if kpis:
            vis_tags.append(f"kpis:{len(kpis)}")
        if two_col:
            vis_tags.append("two_col")
        if callout:
            vis_tags.append("callout")

        log["slide_plan"].append({
            "_sld_id": id(sld_id_el),
            "pptx_slide": i, "role": role, "title": s.get("title", ""), "claim": claim,
            "orig_slide": (reuse["slide_index"] + 1 if is_reuse else None),
            "edit_action": (reuse_action if is_reuse else "new"),
            "bullets": s.get("bullets", []),
            "figure": (fig or (figs_list[0] if figs_list else {}) or {}).get("path"),
            "figures": [fl_i.get("path") for fl_i in (figs_list or [])[:2] if isinstance(fl_i, dict) and fl_i.get("path")],
            "figure_label": (fig or (figs_list[0] if figs_list else {}) or {}).get("caption"), "has_table": bool(tbl),
            "visuals": vis_tags,
            "evidence": s.get("evidence", []),
            "time_s": time_s, "notes_chars": len(raw_notes),
            "notes_units": round(speaking_units(raw_notes), 1),
            "notes_tail": raw_notes[-45:],
            "notes_full": raw_notes,
        })

    # auto-fuse: guarantee QA-prep closure if spec has no backup slide (prefer structured qa_plan over raw hooks)
    if not any(s.get("role") == "qa_backup" for s in slides):
        qa_plan_items = []
        qa_file = work / "qa_plan.json"
        if qa_file.exists():
            try:
                raw_q = json.loads(qa_file.read_text(encoding="utf-8"))
                if isinstance(raw_q, list):
                    qa_plan_items = [q for q in raw_q if isinstance(q, dict) and q.get("question")]
            except Exception:
                qa_plan_items = []
        hooks = []
        if not qa_plan_items:
            hooks_file = work / "qa_hooks.json"
            if hooks_file.exists():
                try:
                    hooks = json.loads(hooks_file.read_text(encoding="utf-8")).get("hooks", [])[:3]
                except Exception:
                    hooks = []
        if qa_plan_items or hooks:
            slide = prs.slides.add_slide(pick_layout("qa_backup"))
            sld_id_el = prs.slides._sldIdLst[-1]
            spec_sld_ids.append(sld_id_el)
            set_slide_theme(slide)
            tt, bb = phs(slide)
            qa_frame = slide_frame(slide, tt, bb)
            if mode == "no_template":
                hbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.055))
                hbar.fill.solid()
                hbar.fill.fore_color.rgb = RGBColor(*accent)
                hbar.line.fill.background()
                strip_shadow(hbar)
            fix_body_geom(bb, tt, role="qa_backup", frame=qa_frame)
            is_en_lang = str(meta.get("language", "")).lower().startswith("en")
            if qa_plan_items:
                t_qa = "Backup: Anticipated Q&A Reference" if is_en_lang else "备用：高频追问应答速查"
                CAT_TAG_EN = {"mechanism": "[Mech] ", "data": "[Data] ", "baseline": "[Ctrl] ", "limitation": "[Limit] "}
                CAT_TAG_ZH = {"mechanism": "[机制] ", "data": "[数据] ", "baseline": "[对照] ", "limitation": "[局限] "}
                left_q, right_q, bullets = [], [], []
                for q in qa_plan_items[:4]:
                    cite = f" [P{q['cite_slide']}]" if q.get("cite_slide") else ""
                    cat_l = str(q.get("category") or "").strip().lower()
                    tag = (CAT_TAG_EN if is_en_lang else CAT_TAG_ZH).get(cat_l, "")
                    q_raw = str(q.get("question", "")).strip()
                    a_raw = str(q.get("answer_key", "")).strip()
                    q_short = (q_raw[:68].rsplit(" ", 1)[0] + "…" if len(q_raw) > 68 and " " in q_raw[:68] else (q_raw[:68] + "…" if len(q_raw) > 68 else q_raw))
                    a_short = (a_raw[:140].rsplit(" ", 1)[0] + "…" if len(a_raw) > 140 and " " in a_raw[:140] else (a_raw[:140] + "…" if len(a_raw) > 140 else a_raw))
                    sep = " → " if q_short.endswith(("?", "？")) else (": " if is_en_lang else "：")
                    entry_s = f"{tag}{q_short}{sep}{a_short}{cite}"
                    bullets.append(entry_s)
                    if cat_l in ("mechanism", "data"):
                        left_q.append(entry_s)
                    elif cat_l in ("baseline", "limitation"):
                        right_q.append(entry_s)
                    else:
                        (left_q if len(left_q) <= len(right_q) else right_q).append(entry_s)
                if not left_q or not right_q:
                    mid_q = (len(bullets) + 1) // 2
                    left_q, right_q = bullets[:mid_q], bullets[mid_q:]
                claim = "Source-grounded answer keys aligned with main-deck locators (uncounted backup)" if is_en_lang else "与正文证据口径一致的答辩/追问速查（不占正文时长）"
                nt = ("Backup Q&A slide synthesized from qa_plan.json; point directly to cited main-deck slides during Q&A."
                      if is_en_lang else
                      "由 qa_plan.json 自动汇总的追问备用页；回答追问时直接跳转引用页与原表/图定位。")
            else:
                t_qa = "Backup: Key Q&A Prompts (Auto-Drafted)" if is_en_lang else "备用：追问要点（自动汇总，请人工复核）"
                bullets = [f"{h.get('type')}: ...{str(h.get('snippet'))[:70]}... (p{h.get('page')})" for h in hooks]
                left_q, right_q = [], []
                claim = "Source-grounded Q&A quick reference (uncounted backup)" if is_en_lang else "源材料高频追问点速查（不占正文时间）"
                nt = ("Auto-generated Q&A backup slide drafted from source hooks; review and align answer keys before presenting."
                      if is_en_lang else
                      "自动生成的备用页：由 qa_draft 从源材料扫描得到的高频追问钩子汇总；上页前请核对口径并补充 answer key。")
            set_title(slide, t_qa, tt, role="qa_backup", frame=qa_frame)
            st = add_claim_strip(slide, claim, tt, t_qa, role="qa_backup", frame=qa_frame)
            sb = st[1] if st else None
            if qa_plan_items and len(bullets) >= 2:
                # Render crisp category-grouped 2-column Q&A Reference Board (Left: Mechanism & Data; Right: Baseline & Scope)
                if bb is not None:
                    remove_shape(bb)
                    bb = None
                qa_two_col = {
                    "left_title": ("Mechanism & Data Q&A" if is_en_lang else "机制与数据追问"),
                    "left_bullets": left_q,
                    "right_title": ("Control, Scope & Risk Q&A" if is_en_lang else "对照、边界与局限追问"),
                    "right_bullets": right_q,
                }
                add_two_col(slide, qa_two_col, sb or Inches(1.48), frame=qa_frame)
                vis_qa = ["two_col"]
            else:
                bb = fill_body(slide, bb, bullets, sb, role="qa_backup", frame=qa_frame)
                vis_qa = []
            cleanup_unused_placeholders(slide, [tt, bb])
            add_evidence_footer(slide, [], "qa_backup", not is_en_lang, frame=qa_frame)
            if meta.get("time_annotations", True):
                nt = ("⏱ Backup | " if is_en_lang else "⏱ 备用 | ") + nt
            slide.notes_slide.notes_text_frame.text = nt
            log["warnings"].append("auto-appended qa_backup slide from " + ("qa_plan" if qa_plan_items else "qa_hooks"))
            log["slide_plan"].append({
                "_sld_id": id(sld_id_el),
                "pptx_slide": len(prs.slides), "role": "qa_backup",
                "orig_slide": None, "edit_action": "auto_qa_backup",
                "title": t_qa, "claim": claim,
                "bullets": bullets, "figure": None, "figure_label": None, "has_table": False,
                "visuals": vis_qa, "evidence": [], "time_s": None, "notes_chars": len(nt),
                "notes_units": round(speaking_units(nt), 1),
            })

    # Reorder / prune in edit_deck mode so slide order & count match spec cleanly
    pruned_orig_indices = []
    if mode == "edit_deck":
        sldIdLst = prs.slides._sldIdLst
        keep_only_raw = meta.get("keep_only_spec_slides", None)
        built_roles = {str(p.get("role") or "") for p in log.get("slide_plan") or []}
        is_full_reorg = (
            keep_only_raw is None
            and len(spec_sld_ids) >= 4
            and (
                (orig_slides_count > len(slides) and ("title" in built_roles or bool(built_roles & {"conclusion", "thanks", "qa_backup"})))
                or (len(touched_orig_indices) < orig_slides_count and "title" in built_roles and bool(built_roles & {"conclusion", "thanks", "qa_backup"}))
            )
        )
        keep_only = bool(keep_only_raw) if keep_only_raw is not None else is_full_reorg
        if keep_only:
            keep_set = set(id(el) for el in spec_sld_ids)
            orig_id_to_idx = {id(el): idx + 1 for idx, el in enumerate(orig_sld_ids)}
            dropped = 0
            for el in list(sldIdLst):
                if id(el) not in keep_set:
                    if id(el) in orig_id_to_idx:
                        pruned_orig_indices.append(orig_id_to_idx[id(el)])
                    try:
                        prs.part.drop_rel(el.rId)
                    except Exception:
                        pass
                    sldIdLst.remove(el)
                    dropped += 1
            for el in spec_sld_ids:
                if el.getparent() is sldIdLst:
                    sldIdLst.remove(el)
                sldIdLst.append(el)
            if dropped:
                log["warnings"].append(f"edit_deck(prune): pruned {dropped} unlisted slide(s) {pruned_orig_indices} and reordered to spec")
        else:
            # Include preserved untouched slides in slide_plan so rendering.consistency stays exact
            for si in range(orig_slides_count):
                if si not in touched_orig_indices:
                    sl = prs.slides[si]
                    t0, _ = phs(sl)
                    ttxt = t0.text_frame.text.strip()[:80] if (t0 is not None and t0.has_text_frame) else f"(original slide {si + 1})"
                    log["slide_plan"].append({
                        "_sld_id": id(orig_sld_ids[si]),
                        "pptx_slide": si + 1, "role": "preserved", "title": ttxt, "claim": "",
                        "bullets": [], "figure": None, "figure_label": None, "has_table": False,
                        "visuals": [], "evidence": [], "time_s": None, "notes_chars": 0, "notes_units": 0.0,
                    })

    # Re-index pptx_slide positions from final XML order
    final_order = {id(el): idx + 1 for idx, el in enumerate(list(prs.slides._sldIdLst))}
    for entry in log["slide_plan"]:
        sid = entry.pop("_sld_id", None)
        if sid in final_order:
            entry["pptx_slide"] = final_order[sid]
    log["slide_plan"].sort(key=lambda e: e.get("pptx_slide", 999))

    # Add subtle right-aligned slide counter ONLY in no_template mode (or when explicitly enabled) so we never collide with organizer template footers/master slide numbers
    total_final = len(prs.slides)
    want_slide_nums = meta.get("slide_numbers") if ("slide_numbers" in meta) else (mode == "no_template")
    if want_slide_nums and total_final >= 2:
        plan_role_by_no = {e.get("pptx_slide"): e.get("role") for e in log["slide_plan"]}
        for idx_f, sl in enumerate(prs.slides, 1):
            if idx_f == 1 or plan_role_by_no.get(idx_f) in ("title", "thanks", "preserved"):
                continue
            has_native_num = False
            for sh in list(sl.shapes):
                try:
                    if sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                        if sh.has_text_frame and sh.text_frame.text.strip():
                            has_native_num = True
                            break
                    if sh.has_text_frame and sh.left is not None and sh.top is not None:
                        if sh.left > SW - Inches(1.8) and sh.top > SH - Inches(0.85) and sh.text_frame.text.strip():
                            has_native_num = True
                            break
                except Exception:
                    continue
            if has_native_num:
                continue
            set_slide_theme(sl)
            t_f, b_f = phs(sl, allow_fallback=True)
            fl_f, fw_f, fy_f = slide_frame(sl, t_f, b_f)
            nw = Inches(0.85)
            nx = int(min(SW - MARGIN - nw, fl_f + fw_f - nw))
            nb = sl.shapes.add_textbox(nx, fy_f, nw, Inches(0.26))
            ntf = nb.text_frame
            ntf.word_wrap = False
            ntf.margin_left = ntf.margin_right = ntf.margin_top = ntf.margin_bottom = 0
            np0 = ntf.paragraphs[0]
            np0.alignment = PP_ALIGN.RIGHT
            nr = np0.add_run()
            nr.text = f"{idx_f:02d} / {total_final:02d}"
            nr.font.size = Pt(9.0)
            nr.font.bold = True
            nr.font.color.rgb = RGBColor(*text_muted)

    prs.save(str(out))
    (work / "build_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
    fonts = theme_fonts(out) if out.exists() else ("", "")
    plan = {
        "meta": {
            "mode": mode, "deck": str(out), "duration_minutes": meta.get("duration_minutes"),
            "fonts": {"major": fonts[0], "minor": fonts[1]},
            "seconds_per_timed_slide": per,
            "template": a.template, "starting_deck": a.deck,
            "goal": meta.get("goal"), "audience": meta.get("audience"),
            "tradeoffs": meta.get("tradeoffs", []),
            "orig_slides_count": (orig_slides_count if mode == "edit_deck" else 0),
            "pruned_orig_slides": pruned_orig_indices,
        },
        "slides": log["slide_plan"],
    }
    (work / "slide_plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"OK: {out} slides={len(prs.slides)} mode={mode} per_timed_slide={per}s")
    for w in log["warnings"]:
        print("  warn:", w)


if __name__ == "__main__":
    main()
