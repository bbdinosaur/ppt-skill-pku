#!/usr/bin/env python3
"""build_deck.py — Step 4 of the academic-slides skill.

Builds /app/output/presentation.pptx from work/slides_plan.json using the REAL
target template (layout/placeholder inheritance; masters never modified) or the
real starting deck in edit mode. Images are inserted aspect-preserved (never
cropped/stretched). Speaker notes and per-slide content_map sidecar are written.

Usage:
  python3 build_deck.py --plan work/slides_plan.json --work work \
      [--template inputs/target-template.pptx] [--start inputs/starting-deck.pptx] \
      [--out /app/output/presentation.pptx]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckutil as du  # noqa: E402

from pptx import Presentation  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402


def resolve_path(p, base):
    if not p:
        return None
    q = Path(p)
    return q if q.is_absolute() else base / q


def shrink_ph(prs, slide, ph):
    _, _, pw, phh = du.ph_geom(prs, slide, ph)
    w_pt = (pw.pt if pw else 0) or 0
    h_pt = (phh.pt if phh else 0) or 0
    du.shrink_body_to_fit(ph.text_frame, w_pt, h_pt,
                          du.ph_inherited_pt(ph, du.DEFAULT_BODY_PT))


def shrink_shape(sh, default_pt):
    w_pt = (sh.width.pt if sh.width else 0) or 0
    h_pt = (sh.height.pt if sh.height else 0) or 0
    du.shrink_body_to_fit(sh.text_frame, w_pt, h_pt, default_pt)


def build_create_slide(prs, spec, images_base: Path, accent=None):
    """Create one slide from a plan spec using template layout inheritance."""
    hint = spec.get("layout_hint") or "auto"
    layout, layout_name = du.pick_layout(prs, hint)
    slide = prs.slides.add_slide(layout)
    du.fix_degenerate_placeholders(prs, slide)

    role = spec.get("role", "custom")
    title = spec.get("title", "")
    tph = du.slide_title_ph(slide)
    if title:
        if tph is not None:
            tph.text_frame.text = str(title)
            shrink_title(prs, slide, tph)
        else:
            SW = int(prs.slide_width)
            du.fit_textbox(slide, (int(SW * 0.065), int(prs.slide_height * 0.05),
                                   int(SW * 0.87), int(prs.slide_height * 0.12)),
                           title, size_pt=du.DEFAULT_TITLE_PT, bold=True)
    du.fill_subtitle(slide, spec.get("subtitle", ""), prs)

    if role not in ("title", "qa") and accent:
        du.accent_bar(slide, prs, accent)

    bullets = spec.get("bullets") or []
    two_col = spec.get("two_col") or {}
    image = spec.get("image") or {}
    table = spec.get("table") or {}
    chart = spec.get("chart") or {}
    sources = spec.get("sources") or []

    img_path = None
    img_missing = False
    if image.get("path"):
        p = resolve_path(image["path"], images_base)
        if p and p.is_file():
            img_path = p
        else:
            img_missing = True

    body_phs = du.body_placeholders(slide)
    region = du.content_region(prs, slide)
    rl, rt, rw, rh = region
    src_strip = Inches(0.36) if (sources and role not in ("title", "qa")) else 0

    has_split = bool(img_path or img_missing or table or chart) and (bullets or two_col)

    # --- bullets (placeholder first for native list styles; textbox fallback)
    if two_col:
        left_b = two_col.get("left_bullets") or []
        right_b = two_col.get("right_bullets") or []
        if len(body_phs) >= 2:
            du.fill_bullets_tf(body_phs[0].text_frame, left_b)
            du.fill_bullets_tf(body_phs[1].text_frame, right_b)
            for ph in body_phs:
                shrink_ph(prs, slide, ph)
        else:
            ph = body_phs[0] if body_phs else None
            if ph is not None:
                try:
                    new_h = max(int((ph.height or rh) - src_strip), Inches(1.0))
                    du.set_ph_size(ph, int(rw * 0.46), new_h)
                    du.fill_bullets_tf(ph.text_frame, left_b)
                    shrink_ph(prs, slide, ph)
                except Exception:
                    pass
            # right column textbox (match left placeholder's inherited font size)
            colw = int(rw * 0.50)
            tx = du.fit_textbox(slide, (rl + rw - colw, rt, colw, rh - src_strip), "")
            du.fill_bullets_tf(tx.text_frame, right_b)
            pt_left = du.ph_inherited_pt(ph, du.DEFAULT_BODY_PT) if ph is not None else du.DEFAULT_BODY_PT
            for para in tx.text_frame.paragraphs:
                para.font.size = Pt(pt_left)
                for r in para.runs:
                    r.font.size = Pt(pt_left)
            shrink_shape(tx, du.DEFAULT_BODY_PT)
    elif bullets:
        if body_phs:
            ph = body_phs[0]
            if has_split:
                try:
                    _, _, pw0, ph0 = du.ph_geom(prs, slide, ph)
                    du.set_ph_size(ph, int(rw * 0.45), int(ph0 or (rh - src_strip)))
                except Exception:
                    try:
                        ph.width = Emu(int(rw * 0.45))
                    except Exception:
                        pass
            du.fill_bullets_tf(ph.text_frame, bullets)
            shrink_ph(prs, slide, ph)
        else:
            colw = int(rw * (0.45 if has_split else 1.0))
            tx = du.fit_textbox(slide, (rl, rt, colw, rh - src_strip), "")
            du.fill_bullets_tf(tx.text_frame, bullets)
            shrink_shape(tx, du.DEFAULT_BODY_PT)

    # --- image (aspect-preserved, centered in its slot; caption + ref)
    if img_path:
        if has_split:
            slot = (rl + int(rw * 0.47), rt, int(rw * 0.50), rh - src_strip)
        else:
            slot = (rl, rt, rw, rh - src_strip)
        cap = image.get("caption") or ""
        ref = image.get("ref") or ""
        cap_text = (f"{ref} {cap}".strip()) if (cap or ref) else None
        du.add_picture_fit(slide, img_path, slot, cap_text=cap_text)
    elif img_missing or image.get("missing"):
        if has_split:
            slot = (rl + int(rw * 0.47), rt, int(rw * 0.50), rh - src_strip)
        else:
            slot = (rl, rt, rw, rh - src_strip)
        cap = image.get("caption") or ""
        ref = image.get("ref") or ""
        du.fit_textbox(slide, slot, f"[图缺失/仅位图] {ref} {cap}".strip(),
                       size_pt=12, italic=True)

    # --- native editable chart (data-bound)
    if chart.get("categories") or chart.get("series"):
        if has_split:
            cslot = (rl + int(rw * 0.47), rt, int(rw * 0.50), rh - src_strip)
        else:
            cslot = (rl, rt, rw, rh - src_strip)
        try:
            du.add_native_chart(slide, cslot, chart)
            cap = chart.get("caption") or ""
            ref = chart.get("ref") or ""
            if cap or ref:
                du.fit_textbox(slide, (cslot[0], cslot[1] + cslot[3] - Inches(0.30),
                                       cslot[2], Inches(0.28)),
                               f"{ref} {cap}".strip(), size_pt=11)
        except Exception as e:
            du.fit_textbox(slide, cslot, f"[图表生成失败: {e}]", size_pt=12, italic=True)

    # --- table
    if table.get("rows"):
        rows = table["rows"]
        if has_split:
            trect = (rl + int(rw * 0.47), rt + Inches(0.1), int(rw * 0.50),
                     min(rh - Inches(0.2), Inches(0.35) * len(rows)))
        else:
            trect = (rl, rt, rw, min(rh - Inches(0.4), Inches(0.38) * len(rows)))
        du.add_table(slide, trect, rows, header=table.get("header", True),
                     font_pt=float(table.get("font_pt", 12)))

    # --- bottom source line (claim traceability on-slide)
    if sources and role not in ("title", "qa"):
        du.source_line(slide, (rl, rt + rh - Inches(0.30), rw, Inches(0.28)), sources,
                       prs.slide_width, lang=getattr(prs, "_deck_lang", "zh"))

    du.set_notes(slide, spec.get("notes", ""))
    return {"role": role, "layout": layout_name, "title": title, "image_missing": img_missing}


def shrink_title(prs, slide, tph, min_pt=20.0):
    """Step title font down (1pt steps) until estimated height fits the title box."""
    try:
        _, _, tw, th = du.ph_geom(prs, slide, tph)
        if not (tw and th):
            return
        size = du.ph_inherited_pt(tph, du.DEFAULT_TITLE_PT, getattr(prs, "_master_styles", None) or None)
        new = size
        while new > min_pt and du.tf_text_height_pt(tph.text_frame, tw.pt, new) > th.pt * 0.90:
            new -= 1.0
        if new < size - 0.5:
            for p in tph.text_frame.paragraphs:
                p.font.size = Pt(new)
                for r in p.runs:
                    if r.font.size is None or r.font.size.pt > new:
                        r.font.size = Pt(new)
    except Exception:
        pass


def patch_next_partname(prs):
    """python-pptx allocates slide partnames by len(sldIdLst)+1, which collides after
    slide deletion (numbering gap). Patch to lowest-unused allocation."""
    try:
        from pptx.parts.presentation import PresentationPart
        from pptx.opc.packuri import PackURI
        pkg = prs.part.package

        def _next(self):
            used = {str(p.partname) for p in pkg.iter_parts()}
            tmpl = "/ppt/slides/slide%d.xml"
            n = 1
            while tmpl % n in used:
                n += 1
            return PackURI(tmpl % n)

        PresentationPart._next_slide_partname = property(_next)
    except Exception:
        pass


def apply_edit_ops(prs, edits, images_base: Path, work: Path, cm_changes, accent=None):
    appended = {}
    for e in edits:
        op = e.get("op", "")
        sidx = int(e.get("slide", 0)) - 1
        if op == "set_title":
            slide = prs.slides[sidx]
            tph = du.slide_title_ph(slide)
            if tph is not None:
                tph.text_frame.text = str(e.get("text", ""))
                shrink_title(prs, slide, tph)
            cm_changes.append({"op": "set_title", "slide": sidx + 1, "to": e.get("text", ""),
                                "sources": e.get("sources") or []})
        elif op == "set_body":
            slide = prs.slides[sidx]
            bphs = du.body_placeholders(slide)
            if bphs:
                du.fill_bullets_tf(bphs[0].text_frame, e.get("bullets") or [])
                shrink_ph(prs, slide, bphs[0])
            cm_changes.append({"op": "set_body", "slide": sidx + 1,
                               "n_bullets": len(e.get("bullets") or []),
                               "sources": e.get("sources") or []})
        elif op == "set_notes":
            du.set_notes(prs.slides[sidx], e.get("text", ""))
            cm_changes.append({"op": "set_notes", "slide": sidx + 1})
        elif op == "append":
            spec = e.get("slide_spec") or {}
            info = build_create_slide(prs, spec, images_base, accent)
            appended[len(prs.slides)] = spec
            cm_changes.append({"op": "append", "slide": len(prs.slides),
                                "role": info["role"], "title": info["title"]})
        elif op == "delete":
            du.delete_slide(prs, sidx)
            cm_changes.append({"op": "delete", "slide": sidx + 1})
        elif op == "move":
            du.move_slide(prs, sidx, int(e.get("to", 1)) - 1)
            cm_changes.append({"op": "move", "slide": sidx + 1, "to": e.get("to")})
    # kept pages must not lose speaker notes (completeness): add a neutral marker note
    for i, slide in enumerate(prs.slides, 1):
        has_notes = slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip()
        if not has_notes:
            du.set_notes(slide, "（保留页：内容未改动；本页讲稿可沿用上一版本的说法，或用一句话过渡到新增结果。）")
    return appended


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", required=True)
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--out", default="/app/output/presentation.pptx")
    ap.add_argument("--template", default=None)
    ap.add_argument("--start", default=None)
    args = ap.parse_args()

    work = Path(args.work)
    plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
    inv = du.load_inventory(work)
    images_base = Path(inv.get("case_dir") or Path(args.plan).resolve().parent)

    template = args.template or (inv.get("template") or {}).get("path")
    start = args.start or (inv.get("start_deck") or {}).get("path")
    mode = plan.get("mode", "auto")
    if mode == "auto":
        wo_mode = (inv.get("workorder") or {}).get("mode", "create")
        mode = "edit" if (wo_mode == "edit" and start) else "create"

    if mode == "edit" and start and Path(start).is_file():
        prs = Presentation(str(start))
        base_layouts = len(prs.slide_layouts)
    elif template and Path(template).is_file():
        prs = Presentation(str(template))
        # start from a clean slide set if template carries sample pages
        while len(prs.slides) > 0:
            du.delete_slide(prs, 0)
    else:
        prs = Presentation()
        print("BUILD_WARN: no template found, using default blank template")

    cm_edits: list = []
    accent = du.theme_accent_color(prs, template or start)
    master_styles = du.load_master_styles(prs, template or start)
    patch_next_partname(prs)
    wo_lang = str((inv.get("workorder") or {}).get("language") or plan.get("language") or "zh")
    prs._deck_lang = "en" if wo_lang.lower().startswith("en") else "zh"
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    cm = []
    if mode == "edit":
        appended = apply_edit_ops(prs, plan.get("edits") or [], images_base, work, cm_edits, accent)
        # content_map covers ALL final slides (kept pages included) for evidence audit
        op_sources = {}
        for ch in cm_edits:
            if ch.get("sources"):
                op_sources[ch["slide"]] = (op_sources.get(ch["slide"]) or []) + ch["sources"]
        for i, slide in enumerate(prs.slides, 1):
            spec = appended.get(i) or {}
            tph = du.slide_title_ph(slide)
            title = tph.text_frame.text if (tph is not None and tph.has_text_frame) else ""
            notes = ""
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
            cm.append({"slide": i, "role": spec.get("role", "existing"), "layout": "",
                       "title": title,
                       "bullets": [b.get("t") if isinstance(b, dict) else str(b)
                                   for b in (spec.get("bullets") or [])],
                       "notes": notes,
                       "sources": (spec.get("sources") or []) + op_sources.get(i, []),
                       "image": spec.get("image") or {}, "table": spec.get("table") or {}, "chart": spec.get("chart") or {}})
    else:
        for spec in plan.get("slides") or []:
            info = build_create_slide(prs, spec, images_base, accent)
            img_spec = dict(spec.get("image") or {})
            if info.pop("image_missing", False):
                img_spec["missing"] = True
            cm.append({"slide": len(prs.slides), **info,
                       "bullets": [b.get("t") if isinstance(b, dict) else str(b)
                                   for b in (spec.get("bullets") or [])],
                       "notes": spec.get("notes", ""),
                       "sources": spec.get("sources") or [],
                       "image": img_spec, "table": spec.get("table") or {}, "chart": spec.get("chart") or {}})

    prs.save(str(out))
    try:
        cp = prs.core_properties
        cp.title = plan.get("deck_title") or "Academic Deck"
        cp.author = "academic-slides-pku skill"
        cp.subject = "paper-to-slides with evidence mapping"
    except Exception:
        pass
    (work / "content_map.json").write_text(
        json.dumps({"deck_title": plan.get("deck_title", ""), "mode": mode,
                    "out": str(out), "slides": cm,
                    "edits_applied": cm_edits},
                   ensure_ascii=False, indent=1), encoding="utf-8")
    du.log_change(work, "build", f"mode={mode} slides={len(prs.slides)} out={out}")
    print(f"BUILD_OK slides={len(prs.slides)} mode={mode} out={out}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        import traceback
        traceback.print_exc()
        du.log_change(Path("/tmp"), "build_failed", str(e))
        print(f"BUILD_FAILED: {e}")
        sys.exit(1)
