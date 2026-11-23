#!/usr/bin/env python3
"""Static QA on a built deck: bounds, 2D shape overlap, text-overflow estimate,
minimum font size, image aspect distortion & effective DPI, per-slide notes presence,
same-layout title geometry consistency, and verifiable master/placeholder inheritance stats.

Usage:
  python3 check_deck.py DECK.pptx [--json report.json] [--margin 0.06]
Exit codes: 0 = pass (critical=0), 1 = critical violations found.

No-script fallback: open the deck, verify every slide has a title + notes, no shape
crosses the slide edge, no text overlaps figures, and no figure is stretched.
"""
import argparse, json, math, re, sys
from collections import defaultdict
from pathlib import Path

FOOTERISH = ("DATE", "FOOTER", "SLIDE_NUMBER", "HEADER", "BILLBOARD")
CHROME_NAMES = ("__adaptor_footer__", "__adaptor_snum__")
BOILERPLATE_CRIT = ("click to add", "单击此处添加", "双击此处添加", "lorem ipsum", "[image unavailable")
BOILERPLATE_WARN = re.compile(r"\b(TBD|TODO|FIXME|XXX)\b")
FIG_REF_RE = re.compile(r"\b(?:Figure|Fig\.?)\s*([SA]?\d+[a-z]?|[IVX]{1,4})\b|图\s*([SA]?\d+[a-z]?)", re.I)
NUM_TOKEN_RE = re.compile(
    r"(?<![A-Za-z0-9_])("
    r"\d+\.\d+\s*%|"
    r"\d+\.\d+\s*[x×]\b|"
    r"\d{2,3}\s*%|"
    r"\d+\.\d{1,4}|"
    r"\d{1,3}(?:,\d{3})+"
    r")(?![A-Za-z0-9_])"
)


def est_table_height_in(table, total_w_in):
    """Estimate realistic rendered height of a PPTX table accounting for wrapped cells."""
    from pptx.util import Emu
    col_ws = []
    for col in table.columns:
        try:
            col_ws.append(max(Emu(col.width).inches, 0.3))
        except Exception:
            col_ws.append(max(total_w_in / max(len(table.columns), 1), 0.4))
    tot = 0.0
    for row in table.rows:
        try:
            explicit_h = Emu(row.height).inches if row.height else 0.0
        except Exception:
            explicit_h = 0.0
        max_lines = 1
        row_pt = 12.0
        for c_idx, cell in enumerate(row.cells):
            txt = cell.text_frame.text.strip()
            if not txt:
                continue
            pt = 12.0
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        pt = r.font.size.pt
                        break
            row_pt = max(row_pt, pt)
            cw = col_ws[c_idx] if c_idx < len(col_ws) else 1.2
            avail_pt = max(cw * 72 - 12, 24)
            lines = sum(max(1, math.ceil(est_text_w_pt(ln, pt) / avail_pt)) for ln in txt.splitlines() if ln.strip()) or 1
            max_lines = max(max_lines, lines)
        need_h = max_lines * row_pt * 1.24 / 72 + 0.11
        tot += max(explicit_h, need_h)
    return tot


def _norm_num_variants(tok: str):
    """Generate canonical number string variants for matching against paper_text.txt."""
    core = re.sub(r"[%\s×x,]", "", tok).strip()
    out = {core}
    try:
        v = float(core)
        out.add(f"{v:g}")
        if abs(v - round(v)) < 1e-6:
            out.add(str(int(round(v))))
        out.add(f"{v:.1f}")
        out.add(f"{v:.2f}")
    except ValueError:
        pass
    return out


def _build_paper_num_index(paper_text: str):
    """Extract normalized numeric tokens from paper_text.txt (incl. suffixed values like 1.47x, 32ms)."""
    nums = set()
    clean = paper_text.replace(",", "")
    for m in re.finditer(r"(?<![A-Za-z0-9_.])\d+(?:\.\d+)?", clean):
        tok = m.group(0)
        nums.add(tok)
        try:
            v = float(tok)
            nums.add(f"{v:g}")
            nums.add(f"{v:.1f}")
            nums.add(f"{v:.2f}")
        except ValueError:
            pass
    return nums


def est_text_w_pt(text, size):
    return sum(size * (1.0 if ord(c) > 0x2E7F else 0.52) for c in text)


def est_tf_height_in(tf, width_in):
    """Rough required height for a text frame at its run font sizes."""
    total = 0.0
    for p in tf.paragraphs:
        size = 18.0
        txt = "".join(r.text for r in p.runs)
        for r in p.runs:
            if r.font.size:
                size = r.font.size.pt
                break
        lvl_indent = 0.4 * (p.level or 0)
        w_pt = max((width_in - lvl_indent) * 72 - 20, 40)
        lines = max(1, math.ceil(est_text_w_pt(txt, size) / w_pt)) if txt else 1
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.15
        total += lines * size * ls / 72
        if p.space_after:
            total += p.space_after.pt / 72
    return total


def _rect_overlap(a, b):
    """Intersection area of two boxes (l, t, w, h) in sq inches."""
    ix0 = max(a[0], b[0])
    iy0 = max(a[1], b[1])
    ix1 = min(a[0] + a[2], b[0] + b[2])
    iy1 = min(a[1] + a[3], b[1] + b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)


def check(deck_path: Path, margin_in: float, paper_text_path: Path = None, task_text_path: Path = None):
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation(str(deck_path))
    SW, SH = Emu(prs.slide_width).inches, Emu(prs.slide_height).inches
    slide_area = max(SW * SH, 1.0)
    issues = []  # dict(slide, shape, kind, severity, detail)

    paper_nums = None
    if paper_text_path and paper_text_path.exists():
        paper_nums = _build_paper_num_index(paper_text_path.read_text(encoding="utf-8", errors="ignore"))
        if task_text_path and task_text_path.exists():
            paper_nums |= _build_paper_num_index(task_text_path.read_text(encoding="utf-8", errors="ignore"))

    slides_with_title_ph = 0
    slides_with_body_ph = 0
    manual_textboxes = 0
    layouts_used = set()
    num_checked = 0
    num_grounded = 0
    unverified_nums = []

    for i, slide in enumerate(prs.slides, 1):
        lo_name = slide.slide_layout.name if slide.slide_layout is not None else "?"
        layouts_used.add(lo_name)
        has_tph = False
        has_bph = False
        content_boxes = []  # (name, kind_tag, (l, t, w, h))
        slide_body_texts = []
        slide_caption_texts = []
        has_pic_or_tbl = False
        title_pt_max = 0.0
        body_pt_max = 0.0

        for sh in slide.shapes:
            name = f"{sh.shape_type}/{sh.name}"
            pt = str(sh.placeholder_format.type).split(" ")[0] if sh.is_placeholder else ""
            is_ftr = any(f in pt for f in FOOTERISH) or (sh.name in CHROME_NAMES)
            if pt in ("TITLE", "CENTER_TITLE"):
                has_tph = True
            elif pt in ("BODY", "OBJECT", "TEXT", "SUBTITLE", "VERTICAL_BODY", "CONTENT"):
                has_bph = True

            try:
                l, t = Emu(sh.left).inches, Emu(sh.top).inches
                w, h = Emu(sh.width).inches, Emu(sh.height).inches
            except (TypeError, AttributeError):
                continue

            # skip full-slide decorative background shapes
            if (w * h) > 0.78 * slide_area and not sh.has_text_frame and not sh.has_table:
                continue

            if l < -margin_in or t < -margin_in or l + w > SW + margin_in or t + h > SH + margin_in:
                issues.append(dict(slide=i, shape=name, kind="out-of-bounds", severity="critical",
                                   detail=f"box=({l:.2f},{t:.2f},{w:.2f},{h:.2f}) slide=({SW:.2f}x{SH:.2f})"))
            if sh.is_placeholder and (w <= 0.01 or h <= 0.01):
                issues.append(dict(slide=i, shape=name, kind="zero-size-placeholder", severity="critical",
                                   detail=f"placeholder collapsed to {w:.2f}x{h:.2f}in — inheritance broken"))

            is_pic = sh.shape_type is not None and "PICTURE" in str(sh.shape_type)
            has_txt = sh.has_text_frame and bool(sh.text_frame.text.strip())
            if has_txt:
                raw_txt = sh.text_frame.text.strip()
                low_txt = raw_txt.lower()
                for bp in BOILERPLATE_CRIT:
                    if bp in low_txt:
                        issues.append(dict(slide=i, shape=name, kind="boilerplate-leak", severity="critical",
                                           detail=f"template/broken placeholder text detected: {bp!r}"))
                if BOILERPLATE_WARN.search(raw_txt):
                    issues.append(dict(slide=i, shape=name, kind="unfinished-marker", severity="warning",
                                       detail="TBD/TODO placeholder marker found in slide text"))

                need = est_tf_height_in(sh.text_frame, w)
                if need > h * 1.30 + 0.05:
                    issues.append(dict(slide=i, shape=name, kind="text-overflow", severity="critical",
                                       detail=f"need~{need:.2f}in > box{h:.2f}in"))
                elif need > h * 1.08 + 0.03:
                    issues.append(dict(slide=i, shape=name, kind="text-tight", severity="warning",
                                       detail=f"need~{need:.2f}in vs box{h:.2f}in"))
                # min font check on main body shapes (ignore captions <=12pt & footers)
                sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
                if pt in ("TITLE", "CENTER_TITLE") and sizes:
                    title_pt_max = max(title_pt_max, max(sizes))
                elif not is_ftr and sizes and (sh.is_placeholder or max(sizes) >= 12.5):
                    body_pt_max = max(body_pt_max, max(sizes))
                if sh.is_placeholder and not is_ftr and pt not in ("TITLE", "CENTER_TITLE"):
                    if sizes and min(sizes) < 12.0:
                        issues.append(dict(slide=i, shape=name, kind="tiny-font", severity="critical",
                                           detail=f"body font {min(sizes):.1f}pt < 12pt floor"))
                if not sh.is_placeholder and not is_ftr:
                    if not sizes or max(sizes) >= 12.5:
                        manual_textboxes += 1
                    else:
                        slide_caption_texts.append(raw_txt)
                if not is_ftr:
                    if pt not in ("TITLE", "CENTER_TITLE") and (sh.is_placeholder or not sizes or max(sizes) >= 12.5):
                        slide_body_texts.append(raw_txt)
                    # tighten effective vertical box to actual text height for realistic overlap check
                    eff_h = min(h, max(need + 0.08, 0.25))
                    content_boxes.append((name, "text", (l, t, w, eff_h)))

            elif sh.has_table:
                has_pic_or_tbl = True
                est_th = est_table_height_in(sh.table, w)
                eff_th = max(h, est_th)
                if t + h > SH + margin_in or t + eff_th > SH + 0.35:
                    issues.append(dict(slide=i, shape=name, kind="table-overflow", severity="critical",
                                       detail=f"wrapped table bottom~{t + eff_th:.2f}in > slideH={SH:.2f}in"))
                elif t + eff_th > SH + margin_in:
                    issues.append(dict(slide=i, shape=name, kind="table-tight", severity="warning",
                                       detail=f"wrapped table bottom~{t + eff_th:.2f}in near slideH={SH:.2f}in"))
                content_boxes.append((name, "table", (l, t, w, min(eff_th, h + 0.25))))
                for row in sh.table.rows:
                    row_txt = " | ".join(c.text_frame.text.strip() for c in row.cells)
                    slide_body_texts.append(row_txt)
            elif getattr(sh, "has_chart", False):
                has_pic_or_tbl = True
                content_boxes.append((name, "chart", (l, t, w, h)))
                try:
                    for plot in sh.chart.plots:
                        for ser in plot.series:
                            vals = [str(v) for v in ser.values if v is not None]
                            if vals:
                                slide_body_texts.append(" ".join(vals))
                except Exception:
                    pass
            elif is_pic:
                has_pic_or_tbl = True
                content_boxes.append((name, "picture", (l, t, w, h)))
                try:
                    native = sh.image.size
                    aspect_native = native[0] / native[1]
                    aspect_placed = w / h if h else 0
                    if abs(aspect_native - aspect_placed) / max(aspect_native, 1e-6) > 0.025:
                        issues.append(dict(slide=i, shape=name, kind="aspect-distortion", severity="critical",
                                           detail=f"native={aspect_native:.3f} placed={aspect_placed:.3f}"))
                    eff_dpi = native[0] / w if w else 0
                    if eff_dpi < 80 and min(native) > 200:
                        issues.append(dict(slide=i, shape=name, kind="low-readability", severity="warning",
                                           detail=f"image shown at ~{eff_dpi:.0f} DPI (<80) — enlarge box or re-crop tighter"))
                except Exception:
                    pass

        # Typography hierarchy check: body font must remain smaller than title font on content slides
        if i > 1 and title_pt_max > 0 and body_pt_max > title_pt_max + 1.0:
            issues.append(dict(
                slide=i, shape="typography", kind="hierarchy-inversion", severity="warning",
                detail=f"body font {body_pt_max:.1f}pt > title font {title_pt_max:.1f}pt"))

        # Slide-level figure reference consistency check
        if has_pic_or_tbl and slide_caption_texts and slide_body_texts:
            cap_figs = {
                (m.group(1) or m.group(2)).upper()
                for txt in slide_caption_texts for m in FIG_REF_RE.finditer(txt)
            }
            body_figs = {
                (m.group(1) or m.group(2)).upper()
                for txt in slide_body_texts for m in FIG_REF_RE.finditer(txt)
            }
            if cap_figs and body_figs and cap_figs.isdisjoint(body_figs):
                issues.append(dict(
                    slide=i, shape="fig-ref", kind="fig-ref-mismatch", severity="warning",
                    detail=f"caption shows Fig {sorted(cap_figs)} but body references Fig {sorted(body_figs)}"))

        # Optional quantitative grounding check against paper_text.txt
        if paper_nums is not None and i > 1:
            seen_on_slide = set()
            for txt in slide_body_texts:
                for m in NUM_TOKEN_RE.finditer(txt):
                    tok = m.group(1).strip()
                    # skip years (1900-2099) and duration/slide-count metadata phrases
                    core_plain = tok.replace(",", "").rstrip("%x× ")
                    if core_plain.isdigit() and 1900 <= int(core_plain) <= 2099:
                        continue
                    tail_ctx = txt[m.end():m.end() + 14].lower()
                    if re.match(r"^\s*(?:slides?|pages?|min(?:utes?)?|sec(?:onds?)?|分钟|页)\b", tail_ctx):
                        continue
                    if tok in seen_on_slide:
                        continue
                    seen_on_slide.add(tok)
                    num_checked += 1
                    variants = _norm_num_variants(tok)
                    if any(v in paper_nums for v in variants):
                        num_grounded += 1
                    else:
                        unverified_nums.append(dict(slide=i, token=tok, context=txt[:110]))

        if has_tph:
            slides_with_title_ph += 1
        if has_bph:
            slides_with_body_ph += 1

        # 2D pairwise overlap check among real content shapes
        for a_idx in range(len(content_boxes)):
            na, ka, ba = content_boxes[a_idx]
            area_a = max(ba[2] * ba[3], 1e-4)
            for b_idx in range(a_idx + 1, len(content_boxes)):
                nb, kb, bb = content_boxes[b_idx]
                area_b = max(bb[2] * bb[3], 1e-4)
                ov = _rect_overlap(ba, bb)
                if ov <= 0.12:
                    continue
                ratio = ov / min(area_a, area_b)
                if ratio > 0.22:
                    issues.append(dict(
                        slide=i, shape=f"{na} <-> {nb}", kind="shape-overlap", severity="critical",
                        detail=f"overlap={ov:.2f}sq.in ({ratio * 100:.0f}% of smaller box)"))
                elif ratio > 0.10:
                    issues.append(dict(
                        slide=i, shape=f"{na} <-> {nb}", kind="shape-overlap", severity="warning",
                        detail=f"overlap={ov:.2f}sq.in ({ratio * 100:.0f}% of smaller box)"))

        # notes check
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame.text.strip():
            issues.append(dict(slide=i, shape="-", kind="missing-notes", severity="warning",
                               detail="no speaker notes / evidence line"))

    # consistency of title geometry within the SAME layout name
    geom_by_layout = defaultdict(list)
    for i, slide in enumerate(prs.slides, 1):
        lo_name = slide.slide_layout.name if slide.slide_layout is not None else "?"
        tph = None
        for sh in slide.shapes:
            pt = str(sh.placeholder_format.type) if sh.is_placeholder else ""
            if pt.startswith("TITLE") and not pt.startswith("SUBTITLE") and not pt.startswith("CENTER_TITLE"):
                tph = sh
                break
        if tph is not None and i > 1:
            geom_by_layout[lo_name].append((i, round(Emu(tph.left).inches, 2), round(Emu(tph.top).inches, 2)))
    for lo_name, group in geom_by_layout.items():
        if len(group) < 2:
            continue
        ref = group[0]
        for g in group[1:]:
            if abs(g[1] - ref[1]) > 0.05 or abs(g[2] - ref[2]) > 0.05:
                issues.append(dict(
                    slide=g[0], shape="title-ph", kind="geometry-inconsistent", severity="warning",
                    detail=f"title box {g[1:]!r} differs from slide {ref[0]} {ref[1:]!r} (layout {lo_name!r})"))

    crit = [x for x in issues if x["severity"] == "critical"]
    warn = [x for x in issues if x["severity"] == "warning"]
    n_slides = len(prs.slides)
    report = dict(
        deck=str(deck_path),
        slides=n_slides,
        critical=len(crit),
        warnings=len(warn),
        inheritance_summary=dict(
            slides_with_title_ph=slides_with_title_ph,
            slides_with_body_ph=slides_with_body_ph,
            manual_textboxes=manual_textboxes,
            title_inheritance_rate=round(slides_with_title_ph / max(n_slides, 1), 3),
            layouts_used=sorted(layouts_used),
        ),
        issues=issues,
    )
    # Keep optional numeric check strictly advisory (never inflates warnings/result.json)
    if paper_nums is not None:
        report["numeric_advisory"] = dict(
            checked=num_checked,
            grounded=num_grounded,
            grounding_rate=round(num_grounded / max(num_checked, 1), 3),
            unverified=unverified_nums[:20],
        )
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--json", default="deck_check.json")
    ap.add_argument("--margin", type=float, default=0.06)
    ap.add_argument("--paper-text", default=None,
                    help="optional paper_text.txt for quantitative number grounding verification")
    ap.add_argument("--task-text", default=None,
                    help="optional task.md / case.json to whitelist constraint numbers")
    a = ap.parse_args()
    p = Path(a.deck)
    if not p.exists():
        print(f"FATAL: {p} not found", file=sys.stderr)
        return 2
    ptxt = Path(a.paper_text) if a.paper_text else None
    ttxt = Path(a.task_text) if a.task_text else None
    rep = check(p, a.margin, paper_text_path=ptxt, task_text_path=ttxt)
    Path(a.json).parent.mkdir(parents=True, exist_ok=True)
    Path(a.json).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    ng = rep.get("numeric_advisory")
    ng_str = f" num_advisory={ng['grounded']}/{ng['checked']}" if ng else ""
    print(f"CHECK slides={rep['slides']} critical={rep['critical']} warnings={rep['warnings']} "
          f"title_inherit={rep['inheritance_summary']['title_inheritance_rate']}{ng_str} -> {a.json}")
    for x in rep["issues"][:20]:
        print(f"  [{x['severity']}] s{x['slide']} {x['kind']}: {x['detail']}")
    return 1 if rep["critical"] else 0


if __name__ == "__main__":
    sys.exit(main())
