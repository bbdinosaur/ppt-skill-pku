#!/usr/bin/env python3
"""build_deck.py — Step 4 of the academic-slides skill.

Builds /app/output/presentation.pptx from work/slides_plan.json using the REAL
target template (layout/placeholder inheritance; masters never modified) or the
real starting deck in edit mode. Images are inserted aspect-preserved (never
cropped/stretched). Supports native data-bound charts, styled adaptive tables,
KPI highlight cards, takeaway callouts, rich edit ops and per-slide diffs.

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
from pptx.enum.text import PP_ALIGN  # noqa: E402
from pptx.util import Emu, Inches, Pt  # noqa: E402


def resolve_path(p, base: Path, work: Path | None = None):
    """Resolve image/asset path across case_dir, common subdirs, and work/figs."""
    if not p:
        return None
    q = Path(p)
    if q.is_absolute():
        return q if q.is_file() else None
    bases = [base, base / "inputs", base / "artifacts", base / "materials", base / "material"]
    if work is not None:
        bases.extend([work, work / "figs"])
    for b in bases:
        cand = b / q
        if cand.is_file():
            return cand
    # Last resort: match unique filename under base or work/figs
    name = q.name
    for search_root in ([work / "figs"] if work is not None else []) + [base]:
        if search_root.is_dir():
            hits = [h for h in search_root.rglob(name) if h.is_file()]
            if hits:
                return hits[0]
    return base / q


def shrink_ph(prs, slide, ph, min_pt=11.0, compact_spacing_pt=None):
    _, _, pw, phh = du.ph_geom(prs, slide, ph)
    w_pt = (pw.pt if pw else 0) or 0
    h_pt = (phh.pt if phh else 0) or 0
    du.shrink_body_to_fit(
        ph.text_frame, w_pt, h_pt,
        du.ph_inherited_pt(ph, du.DEFAULT_BODY_PT, getattr(prs, "_master_styles", None)),
        min_pt=min_pt, compact_spacing_pt=compact_spacing_pt,
    )


def shrink_shape(sh, default_pt, min_pt=11.0, compact_spacing_pt=None):
    w_pt = (sh.width.pt if sh.width else 0) or 0
    h_pt = (sh.height.pt if sh.height else 0) or 0
    du.shrink_body_to_fit(sh.text_frame, w_pt, h_pt, default_pt,
                          min_pt=min_pt, compact_spacing_pt=compact_spacing_pt)


def shrink_title(prs, slide, tph, min_pt=17.0, is_cover: bool = False):
    """Step title font down (1pt steps) until estimated height fits cleanly above the accent divider.
    On content slides, also avoid awkward 1-word orphan wraps by stepping down gently (to >=22pt)
    when a moderate single-paragraph title can sit on one crisp line."""
    try:
        _, _, tw, th = du.ph_geom(prs, slide, tph)
        if not (tw and th):
            return
        size = du.ph_inherited_pt(tph, du.DEFAULT_TITLE_PT, getattr(prs, "_master_styles", None))
        # Cap oversized 44pt blank-default titles on content slides so claim titles stay crisp and never touch the divider
        new = min(size, 32.0) if (not is_cover and size > 34.0) else size
        avail_w_pt = max(tw.pt - 18.0, 120.0)
        # Prefer a clean single-line claim title if a gentle shrink (>=22pt) eliminates orphan line wraps
        if not is_cover and len(tph.text_frame.paragraphs) == 1:
            t_txt = (tph.text_frame.text or "").strip()
            if 0 < du.em_width(t_txt) <= 46.0 and du.est_lines(t_txt, new, avail_w_pt) > 1:
                one_line_floor = max(min_pt, 22.0)
                cand = new
                while cand > one_line_floor and du.est_lines(t_txt, cand, avail_w_pt) > 1:
                    cand -= 1.0
                if du.est_lines(t_txt, cand, avail_w_pt) == 1:
                    new = cand
        while new > min_pt and du.tf_text_height_pt(tph.text_frame, avail_w_pt, new) > th.pt * 0.80:
            new -= 1.0
        if new < size - 0.5:
            for p in tph.text_frame.paragraphs:
                p.font.size = Pt(new)
                for r in p.runs:
                    if r.font.size is None or r.font.size.pt > new:
                        r.font.size = Pt(new)
    except Exception:
        pass


def _clamp_ph_to_region(prs, slide, ph, target_l, target_t, target_w, target_h):
    """Keep placeholder inside target region atomically (preserves inherited left/top!)."""
    try:
        l0, t0, w0, h0 = du.ph_geom(prs, slide, ph)
        cur_l = int(l0) if du.ph_ok(l0) else int(target_l)
        cur_t = int(t0) if du.ph_ok(t0) else int(target_t)
        new_t = max(cur_t, int(target_t))
        max_b = int(target_t) + int(target_h)
        cur_h = int(h0) if du.ph_ok(h0) else int(target_h)
        new_h = max(int(Inches(0.8)), min(cur_h, max_b - new_t))
        new_w = int(target_w) if target_w is not None else (int(w0) if du.ph_ok(w0) else int(target_h))
        if new_t != cur_t or target_w is not None or new_h < cur_h:
            du.set_ph_rect(prs, slide, ph, left=cur_l, top=new_t, width=new_w, height=new_h)
    except Exception:
        pass


def _cap_body_pt(prs, ph, bullets, max_pt: float):
    """Cap oversized default template body fonts (e.g. 28pt default) to crisp academic
    typography without touching font.name (template font family/color fully inherited)."""
    try:
        if any(isinstance(b, dict) and b.get("size") for b in (bullets or [])):
            return  # author explicitly chose per-bullet sizes
        inh = du.ph_inherited_pt(ph, du.DEFAULT_BODY_PT, getattr(prs, "_master_styles", None))
        if inh <= max_pt + 0.5:
            return
        for p in ph.text_frame.paragraphs:
            lvl = getattr(p, "level", 0) or 0
            cap = max(12.0, max_pt - lvl * 1.5)
            p.font.size = Pt(cap)
            for r in p.runs:
                if r.font.size is None or r.font.size.pt > cap:
                    r.font.size = Pt(cap)
    except Exception:
        pass


def build_create_slide(prs, spec, images_base: Path, work: Path | None = None,
                       accent=None, palette=None):
    """Create one slide from a plan spec using template layout inheritance."""
    role = spec.get("role", "custom")
    tighten = int(spec.get("_tighten", 0) or 0)
    b_min_pt = 10.0 if tighten >= 2 else (10.5 if tighten == 1 else 11.0)
    t_min_pt = 15.0 if tighten >= 2 else (16.0 if tighten == 1 else 17.0)
    c_space = 1.0 if tighten >= 2 else (2.0 if tighten == 1 else None)
    two_col = spec.get("two_col") or {}
    hint = (spec.get("layout_hint") or "auto").lower().strip()
    if hint == "auto":
        if role in ("title", "qa"):
            hint = "title"
        elif role == "section":
            hint = "section"
        elif two_col:
            hint = "two_content"
    layout, layout_name = du.pick_layout(prs, hint)
    slide = prs.slides.add_slide(layout)
    du.fix_degenerate_placeholders(prs, slide)

    title = du.strip_md(str(spec.get("title", "")))
    tph = du.slide_title_ph(slide)
    SW = int(prs.slide_width)
    needs_light_ink = du.dark_slide_needs_light_ink(prs, slide)
    if title:
        if tph is not None:
            # Widen 4:3-stuck-on-16:9 narrow title placeholders (e.g. 9.0" on 13.33" canvas)
            # so titles align cleanly across the 16:9 slide & avoid off-center / 1-word orphan wraps
            tl0, tt0, tw0, th0 = du.ph_geom(prs, slide, tph)
            if du.ph_ok(tw0) and du.ph_ok(th0) and int(tw0) < int(SW * 0.72) and (int(tl0 or 0) < int(SW * 0.15)):
                if role in ("title", "section", "qa"):
                    du.set_ph_rect(prs, slide, tph, left=int(SW * 0.08), top=tt0, width=int(SW * 0.84), height=th0)
                else:
                    du.set_ph_rect(prs, slide, tph, left=int(SW * 0.065), top=tt0, width=int(SW * 0.87), height=th0)
            tph.text_frame.text = du.glue_num_units(title)
            shrink_title(prs, slide, tph, min_pt=t_min_pt, is_cover=(role in ("title", "section", "qa")))
            if needs_light_ink:
                du.ensure_readable_tf_color(tph.text_frame)
        else:
            tb_t = du.fit_textbox(slide, (int(SW * 0.065), int(prs.slide_height * 0.05),
                                          int(SW * 0.87), int(prs.slide_height * 0.12)),
                                  title, size_pt=du.DEFAULT_TITLE_PT, bold=True,
                                  color="#F1F5F9" if needs_light_ink else None)
            shrink_shape(tb_t, du.DEFAULT_TITLE_PT, min_pt=t_min_pt)
    subtitle_txt = du.strip_md(str(spec.get("subtitle", "")))
    if subtitle_txt.strip() and role in ("title", "section", "qa"):
        sub_ph0 = du.get_ph(slide, du.SUBTITLE_TYPES)
        if sub_ph0 is not None:
            sl0, st0, sw0, sh0 = du.ph_geom(prs, slide, sub_ph0)
            if du.ph_ok(sw0) and du.ph_ok(sh0) and int(sw0) < int(SW * 0.72) and (int(sl0 or 0) < int(SW * 0.22)):
                du.set_ph_rect(prs, slide, sub_ph0, left=int(SW * 0.10), top=st0, width=int(SW * 0.80), height=sh0)
    du.fill_subtitle(slide, subtitle_txt, prs)
    if needs_light_ink and subtitle_txt.strip():
        sub_ph = du.get_ph(slide, du.SUBTITLE_TYPES)
        if sub_ph is not None and sub_ph.has_text_frame:
            du.ensure_readable_tf_color(sub_ph.text_frame, (0xCB, 0xD5, 0xE1))

    if role not in ("title", "qa") and accent:
        du.accent_bar(slide, prs, accent)

    bullets = spec.get("bullets") or []
    image = spec.get("image") or {}
    images_list = [im for im in (spec.get("images") or []) if isinstance(im, dict)][:2]
    table = spec.get("table") or {}
    chart = spec.get("chart") or {}
    highlights = spec.get("highlights") or []
    steps = spec.get("steps") or []
    takeaway = spec.get("takeaway") or ""
    sources = spec.get("sources") or []

    img_path = None
    img_missing = False
    if image.get("path"):
        p = resolve_path(image["path"], images_base, work)
        if p and p.is_file():
            img_path = p
        else:
            img_missing = True

    resolved_imgs = []
    for im in images_list:
        p = resolve_path(im.get("path"), images_base, work) if im.get("path") else None
        resolved_imgs.append((im, p if (p and p.is_file()) else None))

    body_phs = du.body_placeholders(slide)
    if subtitle_txt.strip():
        # Never overwrite a filled SUBTITLE placeholder with bullets on title/qa slides
        body_phs = [ph for ph in body_phs if ph.placeholder_format.type not in du.SUBTITLE_TYPES]
    rl, rt, rw, rh = du.content_region(prs, slide)
    if subtitle_txt.strip() and role in ("title", "qa"):
        # Push extra bullets below the filled title/subtitle area so they never overlap
        sub_bot = None
        for ph in slide.placeholders:
            if ph.placeholder_format.type in (du.SUBTITLE_TYPES | du.TITLE_TYPES):
                _, st, _, sh = du.ph_geom(prs, slide, ph)
                if st is not None and sh is not None:
                    cand_bot = int(st) + int(sh) + int(Inches(0.14))
                    sub_bot = max(sub_bot, cand_bot) if sub_bot is not None else cand_bot
        slide_bot = int(prs.slide_height) - int(Inches(0.32))
        if sub_bot is not None and sub_bot > rt:
            avail_below = slide_bot - sub_bot
            if avail_below >= int(Inches(0.48)):
                rt = sub_bot
                rh = avail_below
            else:
                rt = max(sub_bot, slide_bot - int(Inches(0.55)))
                rh = max(int(Inches(0.45)), slide_bot - rt)
    src_strip = int(Inches(0.34 if tighten else 0.36)) if (sources and role not in ("title", "qa")) else 0
    tk_strip = int(Inches(0.44 if tighten else 0.50)) if (takeaway and role not in ("title", "qa")) else 0
    bot_strip = src_strip + tk_strip

    top_off = 0
    if highlights and role not in ("title", "qa"):
        hi_h = int(Inches(0.96 if tighten >= 2 else (1.06 if tighten == 1 else 1.22)))
        du.add_highlights(slide, (rl, rt, rw, hi_h), highlights, accent=accent)
        top_off = hi_h + int(Inches(0.10 if tighten else 0.14))

    has_img = bool(img_path or img_missing or image.get("missing") or resolved_imgs)
    has_chart = bool(chart.get("categories") or chart.get("series"))
    has_table = bool(table.get("rows"))
    has_visual = has_img or has_chart or has_table

    if steps and role not in ("title", "qa"):
        has_below = bool(bullets or two_col or has_visual)
        avail_for_steps = max(int(Inches(1.1)), rh - top_off - bot_strip)
        if not has_below:
            st_h = min(int(Inches(2.05)), max(int(Inches(1.50)), int(avail_for_steps * 0.50)))
            st_t = rt + top_off + max(0, (avail_for_steps - st_h) // 3)
            du.add_step_flow(slide, (rl, st_t, rw, st_h), steps, accent=accent)
            top_off = (st_t - rt) + st_h + int(Inches(0.14))
        else:
            st_h = int(Inches(1.16 if tighten >= 2 else (1.28 if tighten == 1 else 1.44)))
            du.add_step_flow(slide, (rl, rt + top_off, rw, st_h), steps, accent=accent)
            top_off += st_h + int(Inches(0.12 if tighten else 0.16))

    ct = rt + top_off
    min_ch = int(Inches(0.45)) if role in ("title", "qa") else int(Inches(1.2))
    ch_avail = max(min_ch, rh - top_off - bot_strip)
    # If both two_col and a visual appear, fold two_col cleanly into left-column bullets to prevent overlap
    if two_col and has_visual:
        merged = []
        if two_col.get("left_title"):
            merged.append({"t": str(two_col["left_title"]), "lvl": 0, "bold": True})
        merged.extend(list(two_col.get("left_bullets") or []))
        if two_col.get("right_title"):
            merged.append({"t": str(two_col["right_title"]), "lvl": 0, "bold": True})
        merged.extend(list(two_col.get("right_bullets") or []))
        bullets = merged + list(bullets)
        two_col = {}
    has_split = has_visual and bool(bullets or two_col)

    # Smart left/right split ratio: short bullets need less width so figures/tables shine
    left_frac = 0.44
    if has_split and bullets and len(bullets) <= 3:
        max_len = max((du.em_width(b.get("t", "") if isinstance(b, dict) else str(b)) for b in bullets), default=20)
        tot_len = sum(du.em_width(b.get("t", "") if isinstance(b, dict) else str(b)) for b in bullets)
        if len(bullets) <= 2 and tot_len <= 42:
            left_frac = 0.33
        elif max_len < 22:
            left_frac = 0.38
    right_l = rl + int(rw * (left_frac + 0.025))
    right_w = rl + rw - right_l

    # --- bullets (placeholder first for native list styles; textbox fallback)
    from pptx.dml.color import RGBColor as _RGBColor
    col_hdr_rgb = _RGBColor(*(du._safe_accent_for_dark_bg(accent) if du.slide_is_dark(slide) else du._safe_accent_for_light_bg(accent)))
    if two_col:
        left_b = list(two_col.get("left_bullets") or [])
        right_b = list(two_col.get("right_bullets") or [])
        if two_col.get("left_title"):
            left_b = [{"t": str(two_col["left_title"]), "lvl": 0, "bold": True, "_col_header": True}] + left_b
        if two_col.get("right_title"):
            right_b = [{"t": str(two_col["right_title"]), "lvl": 0, "bold": True, "_col_header": True}] + right_b
        is_sparse_tc = (not has_visual and not highlights and not steps and not tighten
                        and max(len(left_b), len(right_b)) <= 5)
        tc_cap = 13.0 if tighten >= 2 else (14.5 if tighten == 1 else (20.0 if is_sparse_tc else 17.5))

        def _style_sparse_tc(tf_obj):
            if not is_sparse_tc:
                return
            for pi, p in enumerate(tf_obj.paragraphs):
                try:
                    lvl = getattr(p, "level", 0) or 0
                    p.space_after = Pt(10 if pi == 0 else (12 if lvl == 0 else 6))
                    p.line_spacing = 1.18
                except Exception:
                    pass

        if len(body_phs) >= 2 and not has_visual:
            col_gap = int(Inches(0.36))
            col_w = (rw - col_gap) // 2
            du.set_ph_rect(prs, slide, body_phs[0], left=rl, top=ct, width=col_w, height=ch_avail)
            du.set_ph_rect(prs, slide, body_phs[1], left=rl + col_w + col_gap, top=ct, width=col_w, height=ch_avail)
            du.fill_bullets_tf(body_phs[0].text_frame, left_b, accent_rgb=col_hdr_rgb)
            du.fill_bullets_tf(body_phs[1].text_frame, right_b, accent_rgb=col_hdr_rgb)
            _cap_body_pt(prs, body_phs[0], left_b, tc_cap)
            _cap_body_pt(prs, body_phs[1], right_b, tc_cap)
            for ph in body_phs[:2]:
                _style_sparse_tc(ph.text_frame)
                shrink_ph(prs, slide, ph, min_pt=b_min_pt, compact_spacing_pt=c_space)
                if needs_light_ink:
                    du.ensure_readable_tf_color(ph.text_frame)
            # Subtle vertical center divider hairline (<0.08" wide -> pure decorative shape)
            if not highlights and not steps:
                try:
                    from pptx.enum.shapes import MSO_SHAPE
                    div_x = rl + col_w + (col_gap // 2)
                    div_h = max(int(Inches(1.6)), ch_avail - int(Inches(0.35)))
                    vdiv = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, Emu(div_x), Emu(ct + int(Inches(0.06))),
                        Emu(int(Inches(0.014))), Emu(div_h),
                    )
                    vdiv.fill.solid()
                    vdiv.fill.fore_color.rgb = _RGBColor(0x33, 0x4E, 0x73) if du.slide_is_dark(slide) else _RGBColor(0xD8, 0xE1, 0xEC)
                    vdiv.line.fill.background()
                    du._flat_shape(vdiv)
                except Exception:
                    pass
        else:
            ph = body_phs[0] if body_phs else None
            if ph is not None:
                _clamp_ph_to_region(prs, slide, ph, rl, ct, int(rw * 0.46), ch_avail)
                du.fill_bullets_tf(ph.text_frame, left_b, accent_rgb=col_hdr_rgb)
                _cap_body_pt(prs, ph, left_b, min(tc_cap, 20.0 if is_sparse_tc else 17.0))
                _style_sparse_tc(ph.text_frame)
                shrink_ph(prs, slide, ph, min_pt=b_min_pt, compact_spacing_pt=c_space)
                if needs_light_ink:
                    du.ensure_readable_tf_color(ph.text_frame)
            colw = int(rw * 0.50)
            tx = du.fit_textbox(slide, (rl + rw - colw, ct, colw, ch_avail), "")
            du.fill_bullets_tf(tx.text_frame, right_b, accent_rgb=col_hdr_rgb)
            cap_r = 20.0 if is_sparse_tc else min(17.0, tc_cap)
            pt_left = min(cap_r, du.ph_inherited_pt(ph, du.DEFAULT_BODY_PT, getattr(prs, "_master_styles", None))) if ph is not None else cap_r
            for para in tx.text_frame.paragraphs:
                para.font.size = Pt(pt_left)
                for r in para.runs:
                    r.font.size = Pt(pt_left)
            _style_sparse_tc(tx.text_frame)
            shrink_shape(tx, pt_left, min_pt=b_min_pt, compact_spacing_pt=c_space)
            if needs_light_ink:
                du.ensure_readable_tf_color(tx.text_frame)
    elif bullets:
        rendered_cards = False
        # Preserve native template BODY placeholder slot filling on regular content slides;
        # only use agenda numbered strips on agenda/outline/toc, or keypoint cards when layout has no BODY placeholder.
        if not has_split and not highlights and not steps and not tighten:
            all_top = all((int(b.get("lvl", 0) or 0) == 0 if isinstance(b, dict) else True) for b in bullets)
            if role in ("agenda", "outline", "toc") and 3 <= len(bullets) <= 6 and all_top:
                du.add_agenda_rows(slide, (rl, ct, rw, ch_avail), bullets, accent=accent)
                rendered_cards = True
            elif (not body_phs
                  and role not in ("title", "section", "qa", "closing", "thanks", "references", "agenda", "outline", "toc")
                  and 2 <= len(bullets) <= 4 and all_top):
                parsed = []
                for b in bullets:
                    bt = str(b.get("t") if isinstance(b, dict) else b).strip()
                    matched = None
                    for sep in ("：", ": "):
                        if sep in bt:
                            hd, rest = bt.split(sep, 1)
                            if 0 < du.em_width(hd) <= 16 and rest.strip():
                                matched = (hd.strip(), rest.strip())
                                break
                    if matched is None:
                        parsed = []
                        break
                    parsed.append(matched)
                if len(parsed) == len(bullets):
                    du.add_keypoint_cards(slide, (rl, ct, rw, ch_avail), parsed, accent=accent)
                    rendered_cards = True
        if not rendered_cards:
            is_sparse_pure_text = (not has_split and not highlights and not steps and not tighten and len(bullets) <= 5)
            is_sparse_with_steps = (bool(steps) and not has_split and not highlights and not tighten and len(bullets) <= 4)
            b_cap = (13.0 if tighten >= 2 else (14.5 if tighten == 1 else (
                17.5 if has_split else (22.0 if is_sparse_pure_text else (19.5 if is_sparse_with_steps else 20.0))
            )))
            if body_phs:
                ph = body_phs[0]
                target_w = int(rw * left_frac) if has_split else None
                _clamp_ph_to_region(prs, slide, ph, rl, ct, target_w, ch_avail)
                du.fill_bullets_tf(ph.text_frame, bullets)
                _cap_body_pt(prs, ph, bullets, b_cap)
                if is_sparse_pure_text or is_sparse_with_steps:
                    for p in ph.text_frame.paragraphs:
                        try:
                            lvl = getattr(p, "level", 0) or 0
                            p.space_after = Pt((14 if is_sparse_pure_text else 11) if lvl == 0 else 7)
                            p.line_spacing = 1.18 if is_sparse_pure_text else 1.15
                        except Exception:
                            pass
                shrink_ph(prs, slide, ph, min_pt=b_min_pt, compact_spacing_pt=c_space)
                if needs_light_ink:
                    du.ensure_readable_tf_color(ph.text_frame)
            else:
                colw = int(rw * left_frac) if has_split else rw
                tx = du.fit_textbox(slide, (rl, ct, colw, ch_avail), "")
                du.fill_bullets_tf(tx.text_frame, bullets)
                if is_sparse_pure_text or is_sparse_with_steps:
                    for p in tx.text_frame.paragraphs:
                        try:
                            lvl = getattr(p, "level", 0) or 0
                            p.space_after = Pt((14 if is_sparse_pure_text else 11) if lvl == 0 else 7)
                            p.line_spacing = 1.18 if is_sparse_pure_text else 1.15
                        except Exception:
                            pass
                shrink_shape(tx, min(b_cap, du.DEFAULT_BODY_PT), min_pt=b_min_pt, compact_spacing_pt=c_space)
                if needs_light_ink:
                    du.ensure_readable_tf_color(tx.text_frame)

    # --- guaranteed non-overlapping visual slots (handles 1, 2, or 3 visuals cleanly!)
    active_vis = [k for k, on in (("table", has_table), ("chart", has_chart), ("image", has_img)) if on]
    dual_visual = (len(active_vis) >= 2) and not (bullets or two_col)
    visual_slots: dict[str, tuple[int, int, int, int]] = {}
    if len(active_vis) == 1:
        visual_slots[active_vis[0]] = (right_l, ct, right_w, ch_avail) if has_split else (rl, ct, rw, ch_avail)
    elif len(active_vis) >= 2:
        if has_split:
            # Stack active visuals vertically inside the right column with clean gap (zero overlap!)
            v_gap = int(Inches(0.14))
            each_h = max(int(Inches(1.1)), (ch_avail - v_gap * (len(active_vis) - 1)) // len(active_vis))
            for idx_v, vk in enumerate(active_vis):
                visual_slots[vk] = (right_l, ct + idx_v * (each_h + v_gap), right_w, each_h)
        else:
            # Side-by-side horizontally across full content width (table first on left)
            h_gap = int(Inches(0.22))
            each_w = max(int(Inches(1.8)), (rw - h_gap * (len(active_vis) - 1)) // len(active_vis))
            for idx_v, vk in enumerate(active_vis):
                visual_slots[vk] = (rl + idx_v * (each_w + h_gap), ct, each_w, ch_avail)

    dark_canvas = du.slide_is_dark(slide)
    cap_color = "#D8E2EC" if dark_canvas else "#444444"
    # --- image(s) (aspect-preserved, centered in slot; supports single image or 2 side-by-side images)
    if resolved_imgs and not img_path:
        base_slot = visual_slots.get("image", (rl, ct, rw, ch_avail))
        bl, bt, bw, bh = base_slot
        n_im = len(resolved_imgs)
        gap_im = int(Inches(0.22)) if n_im > 1 else 0
        slot_w = (bw - gap_im * (n_im - 1)) // n_im
        for idx_im, (im_spec, im_p) in enumerate(resolved_imgs):
            s_rect = (bl + idx_im * (slot_w + gap_im), bt, slot_w, bh)
            cap = im_spec.get("caption") or ""
            ref = im_spec.get("ref") or ""
            cap_text = du.format_caption(ref, cap) or None
            if im_p:
                du.add_picture_fit(slide, im_p, s_rect, cap_text=cap_text)
            else:
                img_missing = True
                du.fit_textbox(slide, s_rect, f"[图缺失/仅位图] {ref} {cap}".strip(),
                               size_pt=12, italic=True, color="#666666")
    elif img_path:
        slot = visual_slots.get("image", (rl, ct, rw, ch_avail))
        cap = image.get("caption") or ""
        ref = image.get("ref") or ""
        cap_text = du.format_caption(ref, cap) or None
        du.add_picture_fit(slide, img_path, slot, cap_text=cap_text)
    elif img_missing or image.get("missing"):
        slot = visual_slots.get("image", (rl, ct, rw, ch_avail))
        cap = image.get("caption") or ""
        ref = image.get("ref") or ""
        du.fit_textbox(slide, slot, f"[图缺失/仅位图] {ref} {cap}".strip(),
                       size_pt=12, italic=True, color="#666666")

    # --- native editable chart (data-bound; reserves bottom strip for auto-fitted caption!)
    if has_chart:
        cslot = visual_slots.get("chart", (rl, ct, rw, ch_avail))
        cap = chart.get("caption") or ""
        ref = chart.get("ref") or ""
        cap_text = du.format_caption(ref, cap)
        cap_box_h = du.caption_height_emu(cap_text, cslot[2], 11.0)
        cap_h = (cap_box_h + int(Inches(0.04))) if cap_text else 0
        chart_rect = (cslot[0], cslot[1], cslot[2], max(int(Inches(1.2)), cslot[3] - cap_h))
        try:
            du.add_native_chart(slide, chart_rect, chart, palette=palette)
            if cap_text:
                du.add_caption_box(
                    slide,
                    (cslot[0], cslot[1] + chart_rect[3] + int(Inches(0.02)), cslot[2], cap_box_h),
                    cap_text, cap_pt=11.0, color=cap_color,
                )
        except Exception as e:
            du.fit_textbox(slide, cslot, f"[图表生成失败: {e}]", size_pt=12, italic=True)

    # --- table (adaptive widths + theme accent header + smart wrapped-row height & caption)
    if has_table:
        rows = [["" if c is None else str(c) for c in r] for r in table["rows"]]
        tslot = visual_slots.get("table", (rl, ct, rw, ch_avail))
        t_l, t_t, t_w, t_max_h = tslot[0], tslot[1] + int(Inches(0.04)), tslot[2], max(int(Inches(1.1)), tslot[3] - int(Inches(0.08)))
        t_cap = table.get("caption") or ""
        t_ref = table.get("ref") or ""
        t_cap_text = du.format_caption(t_ref, t_cap)
        t_cap_box_h = du.caption_height_emu(t_cap_text, t_w, 11.0)
        t_cap_h = (t_cap_box_h + int(Inches(0.06))) if t_cap_text else 0
        t_box_h = max(int(Inches(1.0)), t_max_h - t_cap_h)
        nrows, ncols = len(rows), max((len(r) for r in rows), default=1)
        if "font_pt" in table and table.get("font_pt") is not None:
            req_pt = float(table["font_pt"])
        elif (has_split or dual_visual) and t_w < int(Inches(6.0)):
            req_pt = 12.0 if ncols >= 4 else 13.0
        else:
            req_pt = 14.5 if (nrows <= 5 and ncols <= 5) else (13.0 if nrows <= 7 else 12.0)
        if tighten >= 2:
            req_pt = min(req_pt, 10.5)
        elif tighten == 1:
            req_pt = min(req_pt, 11.5)
        # Estimate wrapped-row height using the EXACT same sqrt column-width weights as add_table
        col_widths_emu = du.compute_table_col_widths(rows, t_w)
        col_w_est = [max(cw / 12700.0 - 16.0, 22.0) for cw in col_widths_emu]

        def _est_table_h_pt(pt_sz: float) -> float:
            tot = 0.0
            for r in rows:
                ml = 1
                for ci, c in enumerate(r):
                    ml = max(ml, du.est_lines(c, pt_sz, col_w_est[min(ci, ncols - 1)]))
                tot += ml * pt_sz * 1.44 + 12.0
            return tot

        t_pt = req_pt
        while t_pt > 10.0 and _est_table_h_pt(t_pt) * 12700 * 1.12 > t_box_h:
            t_pt -= 0.5
        base_h = int(Inches(0.44 if t_pt >= 13.5 else 0.36) * nrows)
        est_h = int(_est_table_h_pt(t_pt) * 12700 * 1.16)
        t_h = min(t_box_h, max(base_h, est_h))
        # Gently vertically center compact tables so they don't float tiny at top of empty void
        if t_h < int(t_box_h * 0.65):
            t_t = t_t + min(int(Inches(0.42)), (t_box_h - t_h) // 4)
        gf_tbl = du.add_table(
            slide, (t_l, t_t, t_w, t_h), rows,
            header=table.get("header", True), font_pt=t_pt, accent=accent,
            highlight_row=table.get("highlight_row"), max_h=t_box_h,
        )
        actual_t_h = max(int(gf_tbl.height) if getattr(gf_tbl, "height", None) else t_h, est_h)
        actual_t_h = min(t_box_h, actual_t_h)
        if t_cap_text:
            cap_y = min(t_t + actual_t_h + int(Inches(0.10)), ct + ch_avail - t_cap_box_h)
            du.add_caption_box(
                slide,
                (t_l, cap_y, t_w, t_cap_box_h),
                t_cap_text, cap_pt=11.0, color=cap_color,
            )

    # --- bottom takeaway callout (above source line)
    if takeaway and role not in ("title", "qa"):
        tk_y = rt + rh - src_strip - int(Inches(0.46))
        du.add_takeaway_bar(slide, (rl, tk_y, rw, int(Inches(0.42))), takeaway, accent=accent)

    # --- bottom source line (claim traceability on-slide)
    if sources and role not in ("title", "qa"):
        du.source_line(slide, (rl, rt + rh - int(Inches(0.30)), rw, int(Inches(0.28))), sources,
                       prs.slide_width, lang=getattr(prs, "_deck_lang", "zh"))

    notes_text = str(spec.get("notes") or "").strip()
    if not notes_text:
        parts = [str(title or role)]
        if spec.get("subtitle"):
            parts.append(str(spec["subtitle"]))
        if takeaway:
            parts.append(str(takeaway))
        elif bullets:
            b_prev = [b.get("t", "") if isinstance(b, dict) else str(b) for b in bullets[:2]]
            parts.extend(x for x in b_prev if x)
        notes_text = " — ".join(p.strip() for p in parts if p and p.strip())[:240]
    du.set_notes(slide, notes_text)
    # Remove any untouched empty body/subtitle placeholders so PowerPoint never shows ghost "Click to add text"
    for ph in list(slide.placeholders):
        try:
            pt = ph.placeholder_format.type
            if pt in du.BODY_TYPES or (pt in du.SUBTITLE_TYPES and not str(spec.get("subtitle") or "").strip()):
                if not ph.has_text_frame or not (ph.text_frame.text or "").strip():
                    el = ph._element
                    if el.getparent() is not None:
                        el.getparent().remove(el)
        except Exception:
            pass
    if tighten:
        du.clamp_slide_shapes(prs, slide)
    return {"role": role, "layout": layout_name, "title": title, "image_missing": img_missing}


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


def apply_edit_ops(prs, edits, images_base: Path, work: Path, cm_changes,
                   accent=None, palette=None):
    """Apply in-place edit operations; returns (specs_by_eid, sources_by_eid).
    Tracks slide identity by id(slide._element) so delete+append never collides."""
    before_snaps = {}
    before_order = []
    for i, slide in enumerate(prs.slides, 1):
        eid = id(slide._element)
        before_order.append(eid)
        before_snaps[eid] = {"orig_slide": i, **du.extract_slide_snapshot(slide, prs)}

    specs_by_eid = {}
    sources_by_eid = {}
    replaced_meta = {}  # new_eid -> {orig_slide, title_before}
    zero_hit_edits = []

    def _rec_src(eid, srcs):
        if srcs:
            sources_by_eid[eid] = (sources_by_eid.get(eid) or []) + list(srcs)

    for e in edits:
        op = str(e.get("op", "")).strip()
        sraw = e.get("slide", 0)
        sidx = (int(sraw) - 1) if (isinstance(sraw, int) or str(sraw).isdigit()) else -1

        if op == "set_title" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            du.fix_degenerate_placeholders(prs, slide)
            new_title = du.glue_num_units(du.strip_md(str(e.get("text") if e.get("text") is not None else (e.get("title") or ""))))
            tsh = du.find_slide_title_shape(slide, prs)
            if tsh is not None and tsh.has_text_frame:
                du.set_text_preserving_style(tsh.text_frame, new_title)
                if tsh.is_placeholder:
                    shrink_title(prs, slide, tsh)
                else:
                    shrink_shape(tsh, du.effective_pt(tsh, du.DEFAULT_TITLE_PT), min_pt=16.0)
            _rec_src(id(slide._element), e.get("sources"))
            cm_changes.append({"op": "set_title", "slide": sidx + 1, "to": new_title,
                               "sources": e.get("sources") or []})

        elif op == "set_body" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            du.fix_degenerate_placeholders(prs, slide)
            bshs = du.find_slide_body_shapes(slide, prs)
            blist = e.get("bullets") or []
            if bshs:
                du.fill_bullets_tf(bshs[0].text_frame, blist)
                if bshs[0].is_placeholder:
                    _cap_body_pt(prs, bshs[0], blist, 20.0)
                    shrink_ph(prs, slide, bshs[0])
                else:
                    shrink_shape(bshs[0], du.DEFAULT_BODY_PT)
            elif blist:
                rl, rt, rw, rh = du.content_region(prs, slide)
                has_vis_on_slide = any(
                    getattr(sh, "has_table", False) or getattr(sh, "has_chart", False)
                    or (sh.shape_type is not None and "PICTURE" in str(sh.shape_type))
                    for sh in slide.shapes
                )
                bw = int(rw * 0.44) if has_vis_on_slide else rw
                tx = du.fit_textbox(slide, (rl, rt, bw, rh - int(Inches(0.36))), "")
                du.fill_bullets_tf(tx.text_frame, blist)
                shrink_shape(tx, min(18.0, du.DEFAULT_BODY_PT))
            _rec_src(id(slide._element), e.get("sources"))
            cm_changes.append({"op": "set_body", "slide": sidx + 1,
                               "n_bullets": len(blist),
                               "sources": e.get("sources") or []})

        elif op == "set_two_col" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            du.fix_degenerate_placeholders(prs, slide)
            bshs = du.find_slide_body_shapes(slide, prs)
            from pptx.dml.color import RGBColor as _RGBColor
            col_hdr_rgb = _RGBColor(*(du._safe_accent_for_dark_bg(accent) if du.slide_is_dark(slide) else du._safe_accent_for_light_bg(accent)))
            left_b = list(e.get("left_bullets") or [])
            right_b = list(e.get("right_bullets") or [])
            if e.get("left_title"):
                left_b = [{"t": str(e["left_title"]), "lvl": 0, "bold": True, "_col_header": True}] + left_b
            if e.get("right_title"):
                right_b = [{"t": str(e["right_title"]), "lvl": 0, "bold": True, "_col_header": True}] + right_b
            is_sparse_et = max(len(left_b), len(right_b)) <= 5
            et_cap = 20.0 if is_sparse_et else 17.5

            def _style_et(tf_obj):
                if not is_sparse_et:
                    return
                for pi, p in enumerate(tf_obj.paragraphs):
                    try:
                        lvl = getattr(p, "level", 0) or 0
                        p.space_after = Pt(10 if pi == 0 else (12 if lvl == 0 else 6))
                        p.line_spacing = 1.18
                    except Exception:
                        pass

            if len(bshs) >= 2:
                du.fill_bullets_tf(bshs[0].text_frame, left_b, accent_rgb=col_hdr_rgb)
                du.fill_bullets_tf(bshs[1].text_frame, right_b, accent_rgb=col_hdr_rgb)
                for sh, bl in ((bshs[0], left_b), (bshs[1], right_b)):
                    _style_et(sh.text_frame)
                    if sh.is_placeholder:
                        _cap_body_pt(prs, sh, bl, et_cap)
                        shrink_ph(prs, slide, sh)
                    else:
                        shrink_shape(sh, et_cap)
            elif len(bshs) == 1:
                du.fill_bullets_tf(bshs[0].text_frame, left_b + right_b, accent_rgb=col_hdr_rgb)
                _style_et(bshs[0].text_frame)
                if bshs[0].is_placeholder:
                    _cap_body_pt(prs, bshs[0], left_b + right_b, 20.0)
                    shrink_ph(prs, slide, bshs[0])
                else:
                    shrink_shape(bshs[0], du.DEFAULT_BODY_PT)
            elif left_b or right_b:
                rl, rt, rw, rh = du.content_region(prs, slide)
                gap = int(Inches(0.36))
                cw = (rw - gap) // 2
                ch_h = rh - int(Inches(0.36))
                tx1 = du.fit_textbox(slide, (rl, rt, cw, ch_h), "")
                tx2 = du.fit_textbox(slide, (rl + cw + gap, rt, cw, ch_h), "")
                du.fill_bullets_tf(tx1.text_frame, left_b, accent_rgb=col_hdr_rgb)
                du.fill_bullets_tf(tx2.text_frame, right_b, accent_rgb=col_hdr_rgb)
                for tx_e in (tx1, tx2):
                    for para in tx_e.text_frame.paragraphs:
                        para.font.size = Pt(et_cap)
                        for r in para.runs:
                            r.font.size = Pt(et_cap)
                    _style_et(tx_e.text_frame)
                shrink_shape(tx1, et_cap)
                shrink_shape(tx2, et_cap)
                try:
                    from pptx.enum.shapes import MSO_SHAPE
                    div_x = rl + cw + (gap // 2)
                    div_h = max(int(Inches(1.6)), ch_h - int(Inches(0.35)))
                    vdiv = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, Emu(div_x), Emu(rt + int(Inches(0.06))),
                        Emu(int(Inches(0.014))), Emu(div_h),
                    )
                    vdiv.fill.solid()
                    vdiv.fill.fore_color.rgb = _RGBColor(0x33, 0x4E, 0x73) if du.slide_is_dark(slide) else _RGBColor(0xD8, 0xE1, 0xEC)
                    vdiv.line.fill.background()
                    du._flat_shape(vdiv)
                except Exception:
                    pass
            _rec_src(id(slide._element), e.get("sources"))
            cm_changes.append({"op": "set_two_col", "slide": sidx + 1,
                               "sources": e.get("sources") or []})

        elif op == "replace_text":
            old_t, new_t = str(e.get("old", "")), str(e.get("new", ""))
            if isinstance(sraw, list):
                targets = [int(x) - 1 for x in sraw if (isinstance(x, int) or str(x).isdigit()) and 0 <= int(x) - 1 < len(prs.slides)]
            elif str(sraw).lower() == "all" or sidx < 0:
                targets = list(range(len(prs.slides)))
            else:
                targets = [sidx]
            total_hits = 0
            for ti in targets:
                if 0 <= ti < len(prs.slides):
                    sl_t = prs.slides[ti]
                    hits = du.replace_text_in_slide(sl_t, old_t, new_t)
                    if hits:
                        total_hits += hits
                        _rec_src(id(sl_t._element), e.get("sources"))
                        # If replacement lengthened text and caused overflow, gently fit down
                        if len(new_t) > len(old_t):
                            for sh_r in sl_t.shapes:
                                if not (sh_r.has_text_frame and sh_r.text_frame.text.strip()):
                                    continue
                                _, _, sw_r, sh_h = du.shape_geom(prs, sl_t, sh_r)
                                if sw_r and sh_h and sw_r.pt > 20 and sh_h.pt > 12:
                                    cur_pt = du.effective_pt(sh_r, du.DEFAULT_BODY_PT)
                                    if du.tf_text_height_pt(sh_r.text_frame, sw_r.pt, cur_pt) > sh_h.pt * 1.04:
                                        du.shrink_body_to_fit(sh_r.text_frame, sw_r.pt, sh_h.pt, cur_pt, min_pt=10.5)
            if total_hits == 0 and old_t:
                zero_hit_edits.append({"op": "replace_text", "slide": sraw, "old": old_t, "new": new_t})
            cm_changes.append({"op": "replace_text", "slide": sraw, "old": old_t, "new": new_t,
                               "hits": total_hits, "sources": e.get("sources") or []})

        elif op == "set_source" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            srcs = e.get("sources") or []
            if srcs:
                rl, rt, rw, rh = du.content_region(prs, slide)
                # Remove any existing bottom Source/来源 textbox to prevent overlap on re-sourcing
                for sh_old in list(slide.shapes):
                    try:
                        if sh_old.has_text_frame and not sh_old.is_placeholder:
                            t_old = (sh_old.text_frame.text or "").strip()
                            if (t_old.startswith(("来源：", "来源:", "Source:", "Source："))
                                    and sh_old.top is not None and int(sh_old.top) > int(prs.slide_height) * 0.78):
                                sh_old._element.getparent().remove(sh_old._element)
                    except Exception:
                        pass
                du.source_line(slide, (rl, rt + rh - int(Inches(0.30)), rw, int(Inches(0.28))),
                               srcs, prs.slide_width, lang=getattr(prs, "_deck_lang", "zh"))
                _rec_src(id(slide._element), srcs)
            cm_changes.append({"op": "set_source", "slide": sidx + 1, "sources": srcs})

        elif op == "set_notes" and 0 <= sidx < len(prs.slides):
            notes_val = e.get("text") if e.get("text") is not None else (e.get("notes") or "")
            du.set_notes(prs.slides[sidx], str(notes_val))
            cm_changes.append({"op": "set_notes", "slide": sidx + 1})

        elif op == "set_table" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            tbl_spec = e.get("table") if isinstance(e.get("table"), dict) else e
            rows = [["" if c is None else str(c) for c in r] for r in (tbl_spec.get("rows") or [])]
            if rows:
                existing_tbl = next((sh for sh in slide.shapes if getattr(sh, "has_table", False)), None)
                nrows, ncols = len(rows), max(len(r) for r in rows)
                if existing_tbl is not None and len(existing_tbl.table.rows) == nrows and len(existing_tbl.table.columns) == ncols:
                    tbl_e = existing_tbl.table
                    for ri, row in enumerate(rows):
                        for ci, val in enumerate(row):
                            clean_v = du.glue_num_units(du.strip_md(val))
                            du.set_text_preserving_style(tbl_e.cell(ri, ci).text_frame, clean_v)
                    # Guard against lengthened cell text overflowing the existing table frame
                    try:
                        col_ws_pt = [c.width.pt for c in tbl_e.columns]
                        h_pt_e = existing_tbl.height.pt if existing_tbl.height else 0
                        if h_pt_e > 20:
                            def _cur_tbl_est():
                                tot_p = 0.0
                                for r_o in tbl_e.rows:
                                    ml_o, rpt_o = 1, 0.0
                                    for ci_o, c_o in enumerate(r_o.cells):
                                        sz_o = 12.0
                                        for p_o in c_o.text_frame.paragraphs:
                                            for rn_o in p_o.runs:
                                                if rn_o.font.size is not None:
                                                    sz_o = rn_o.font.size.pt
                                                    break
                                        rpt_o = max(rpt_o, sz_o)
                                        ml_o = max(ml_o, du.est_lines(c_o.text, sz_o, max(col_ws_pt[ci_o] - 12, 10)))
                                    tot_p += ml_o * (rpt_o or 12.0) * 1.28 + 4
                                return tot_p
                            for step_pt in (11.5, 10.5, 10.0):
                                if _cur_tbl_est() <= h_pt_e * 1.15 + 4:
                                    break
                                for r_o in tbl_e.rows:
                                    for c_o in r_o.cells:
                                        for p_o in c_o.text_frame.paragraphs:
                                            for rn_o in p_o.runs:
                                                if rn_o.font.size is None or rn_o.font.size.pt > step_pt:
                                                    rn_o.font.size = Pt(step_pt)
                    except Exception:
                        pass
                else:
                    rl, rt, rw, rh = du.content_region(prs, slide)
                    if existing_tbl is not None and existing_tbl.width and existing_tbl.height:
                        tl, tt = int(existing_tbl.left), int(existing_tbl.top)
                        tw = min(int(existing_tbl.width), max(int(Inches(2.0)), rl + rw - tl))
                        max_th = max(int(Inches(1.1)), min(rh - int(Inches(0.32)), rt + rh - tt - int(Inches(0.28))))
                        try:
                            existing_tbl._element.getparent().remove(existing_tbl._element)
                        except Exception:
                            pass
                        slot = (tl, tt, tw, min(max_th, int(Inches(0.42) * nrows)))
                    else:
                        max_th = rh - int(Inches(0.36))
                        slot = (rl, rt, rw, min(max_th, int(Inches(0.42) * nrows)))
                    du.add_table(slide, slot, rows, header=tbl_spec.get("header", True),
                                 font_pt=float(tbl_spec.get("font_pt") or 12.5),
                                 accent=accent, highlight_row=tbl_spec.get("highlight_row"),
                                 max_h=max_th)
            eid = id(slide._element)
            sp_cur = dict(specs_by_eid.get(eid) or {})
            sp_cur["table"] = {
                "ref": tbl_spec.get("ref"), "caption": tbl_spec.get("caption"), "rows": rows,
            }
            specs_by_eid[eid] = sp_cur
            _rec_src(eid, e.get("sources"))
            cm_changes.append({"op": "set_table", "slide": sidx + 1,
                               "rows": len(rows), "sources": e.get("sources") or []})

        elif op == "set_chart" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            ch_spec = e.get("chart") if isinstance(e.get("chart"), dict) else e
            cats = [str(c) for c in (ch_spec.get("categories") or [])]
            sers = ch_spec.get("series") or []
            if cats and sers:
                existing_ch = next((sh for sh in slide.shapes if getattr(sh, "has_chart", False)), None)
                if existing_ch is not None:
                    req_type = str(ch_spec.get("type") or "").strip().lower()
                    type_mismatch = False
                    if req_type:
                        from pptx.enum.chart import XL_CHART_TYPE
                        wanted_ct = {
                            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                            "stacked_bar": XL_CHART_TYPE.COLUMN_STACKED, "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
                            "stacked_hbar": XL_CHART_TYPE.BAR_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS,
                            "pie": XL_CHART_TYPE.PIE,
                        }.get(req_type)
                        try:
                            if wanted_ct is not None and existing_ch.chart.chart_type != wanted_ct:
                                type_mismatch = True
                        except Exception:
                            pass
                    if type_mismatch:
                        slot = (int(existing_ch.left), int(existing_ch.top),
                                int(existing_ch.width), int(existing_ch.height))
                        merged_ch = dict(ch_spec)
                        try:
                            och = existing_ch.chart
                            if not merged_ch.get("title") and och.has_title:
                                merged_ch["title"] = och.chart_title.text_frame.text.strip()
                            for ax_attr, tkey in (("category_axis", "x_title"), ("value_axis", "y_title")):
                                if not merged_ch.get(tkey):
                                    ax = getattr(och, ax_attr, None)
                                    if ax is not None and getattr(ax, "has_title", False):
                                        merged_ch[tkey] = ax.axis_title.text_frame.text.strip()
                        except Exception:
                            pass
                        try:
                            existing_ch._element.getparent().remove(existing_ch._element)
                        except Exception:
                            pass
                        du.add_native_chart(slide, slot, merged_ch, palette=palette)
                        ch_spec = merged_ch
                    else:
                        try:
                            from pptx.chart.data import CategoryChartData
                            cd = CategoryChartData()
                            cd.categories = [du.strip_md(c) for c in cats]
                            any_pct = str(ch_spec.get("unit") or "").strip() == "%"
                            for sr in sers:
                                raw_v = list(sr.get("values") or [])[:len(cats)]
                                pvals = []
                                for v in raw_v:
                                    fv, isp = du.parse_chart_num(v)
                                    if isp:
                                        any_pct = True
                                    pvals.append(fv)
                                while len(pvals) < len(cats):
                                    pvals.append(0.0)
                                cd.add_series(du.strip_md(str(sr.get("name", "series"))), tuple(pvals))
                            existing_ch.chart.replace_data(cd)
                            if ch_spec.get("title"):
                                existing_ch.chart.has_title = True
                                existing_ch.chart.chart_title.text_frame.text = du.strip_md(str(ch_spec["title"]))
                            if any_pct and existing_ch.chart.plots:
                                try:
                                    dl = existing_ch.chart.plots[0].data_labels
                                    dl.number_format = '0.0"%"'
                                except Exception:
                                    pass
                        except Exception:
                            slot = (int(existing_ch.left), int(existing_ch.top),
                                    int(existing_ch.width), int(existing_ch.height))
                            try:
                                existing_ch._element.getparent().remove(existing_ch._element)
                            except Exception:
                                pass
                            du.add_native_chart(slide, slot, ch_spec, palette=palette)
                else:
                    rl, rt, rw, rh = du.content_region(prs, slide)
                    du.add_native_chart(slide, (rl, rt, rw, rh - int(Inches(0.36))), ch_spec, palette=palette)
            eid = id(slide._element)
            sp_cur = dict(specs_by_eid.get(eid) or {})
            sp_cur["chart"] = dict(ch_spec)
            specs_by_eid[eid] = sp_cur
            _rec_src(eid, e.get("sources"))
            cm_changes.append({"op": "set_chart", "slide": sidx + 1,
                               "series": len(sers), "sources": e.get("sources") or []})

        elif op == "set_image" and 0 <= sidx < len(prs.slides):
            slide = prs.slides[sidx]
            img_spec = e.get("image") if isinstance(e.get("image"), dict) else e
            raw_p = img_spec.get("path")
            im_p = resolve_path(raw_p, images_base, work) if raw_p else None
            cap = img_spec.get("caption") or ""
            ref = img_spec.get("ref") or ""
            cap_text = du.format_caption(ref, cap) or None
            existing_pic = next(
                (sh for sh in slide.shapes if sh.shape_type is not None and "PICTURE" in str(sh.shape_type)),
                None,
            )
            if existing_pic is not None and existing_pic.width and existing_pic.height:
                pl, pt_y, pw, ph = (int(existing_pic.left), int(existing_pic.top),
                                    int(existing_pic.width), int(existing_pic.height))
                slot = (pl, pt_y, pw, ph)
                # Remove old caption textbox immediately below replaced picture when adding a new caption
                if cap_text:
                    pic_bot = pt_y + ph
                    for sh_c in list(slide.shapes):
                        try:
                            if sh_c is existing_pic or sh_c.is_placeholder or not sh_c.has_text_frame:
                                continue
                            cl, ct_y, cw, ch = int(sh_c.left), int(sh_c.top), int(sh_c.width), int(sh_c.height)
                            ctext = (sh_c.text_frame.text or "").strip()
                            if (ctext and len(ctext) <= 180 and ch <= int(Inches(0.75))
                                    and pic_bot - int(Inches(0.15)) <= ct_y <= pic_bot + int(Inches(0.55))
                                    and cl < pl + pw and cl + cw > pl
                                    and not ctext.startswith(("来源", "Source"))):
                                sh_c._element.getparent().remove(sh_c._element)
                                slot = (pl, pt_y, pw, max(ph, (ct_y + ch) - pt_y))
                                break
                        except Exception:
                            pass
                try:
                    existing_pic._element.getparent().remove(existing_pic._element)
                except Exception:
                    pass
            else:
                rl, rt, rw, rh = du.content_region(prs, slide)
                slot = (rl + int(rw * 0.48), rt, int(rw * 0.52), rh - int(Inches(0.36)))
            if im_p and im_p.is_file():
                du.add_picture_fit(slide, im_p, slot, cap_text=cap_text)
            else:
                du.fit_textbox(slide, slot, f"[图缺失/仅位图] {ref} {cap}".strip(),
                               size_pt=12, italic=True, color="#666666")
            eid = id(slide._element)
            sp_cur = dict(specs_by_eid.get(eid) or {})
            sp_cur["image"] = {
                "path": str(im_p or raw_p or ""), "ref": ref, "caption": cap,
                **({"missing": True} if not (im_p and im_p.is_file()) else {}),
            }
            specs_by_eid[eid] = sp_cur
            _rec_src(eid, e.get("sources"))
            cm_changes.append({"op": "set_image", "slide": sidx + 1,
                               "path": str(im_p or raw_p or ""), "sources": e.get("sources") or []})

        elif op == "replace_slide" and 0 <= sidx < len(prs.slides):
            old_slide = prs.slides[sidx]
            old_eid = id(old_slide._element)
            old_meta = before_snaps.get(old_eid)
            spec = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else {
                k: v for k, v in e.items() if k not in ("op", "slide", "at")
            }
            info = build_create_slide(prs, spec, images_base, work, accent, palette)
            new_idx = len(prs.slides) - 1
            new_slide = prs.slides[new_idx]
            new_eid = id(new_slide._element)
            specs_by_eid[new_eid] = spec
            if old_meta:
                replaced_meta[new_eid] = {"orig_slide": old_meta["orig_slide"],
                                          "title_before": old_meta["title"],
                                          "old_eid": old_eid}
            du.move_slide(prs, new_idx, sidx)
            du.delete_slide(prs, sidx + 1)
            cm_changes.append({"op": "replace_slide", "slide": sidx + 1,
                               "role": info["role"], "title": info["title"]})

        elif op == "append":
            spec = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else {
                k: v for k, v in e.items() if k not in ("op", "slide", "at")
            }
            info = build_create_slide(prs, spec, images_base, work, accent, palette)
            new_slide = prs.slides[len(prs.slides) - 1]
            specs_by_eid[id(new_slide._element)] = spec
            cm_changes.append({"op": "append", "slide": len(prs.slides),
                               "role": info["role"], "title": info["title"]})

        elif op == "insert_at":
            spec = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else {
                k: v for k, v in e.items() if k not in ("op", "slide", "at")
            }
            at = max(1, min(int(e.get("at", len(prs.slides) + 1)), len(prs.slides) + 1))
            info = build_create_slide(prs, spec, images_base, work, accent, palette)
            new_idx = len(prs.slides) - 1
            new_slide = prs.slides[new_idx]
            specs_by_eid[id(new_slide._element)] = spec
            if at - 1 < new_idx:
                du.move_slide(prs, new_idx, at - 1)
            cm_changes.append({"op": "insert_at", "at": at,
                               "role": info["role"], "title": info["title"]})

        elif op == "delete" and 0 <= sidx < len(prs.slides):
            du.delete_slide(prs, sidx)
            cm_changes.append({"op": "delete", "slide": sidx + 1})

        elif op == "move" and 0 <= sidx < len(prs.slides):
            dst = max(0, min(int(e.get("to", 1)) - 1, len(prs.slides) - 1))
            du.move_slide(prs, sidx, dst)
            cm_changes.append({"op": "move", "slide": sidx + 1, "to": dst + 1})

    # Language-aware fallback note for kept slides lacking notes
    is_en = getattr(prs, "_deck_lang", "zh") == "en"
    default_kept_note = (
        "(Kept slide: content unchanged from starting deck; retain previous talking points or transition to updated results.)"
        if is_en else
        "（保留页：内容未改动；本页讲稿可沿用上一版本的说法，或用一句话过渡到新增结果。）"
    )
    for slide in prs.slides:
        has_notes = slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip()
        if not has_notes:
            du.set_notes(slide, default_kept_note)

    # Compute structured per-slide diff (before vs after)
    after_eids = {id(s._element) for s in prs.slides}
    replaced_old_eids = {m["old_eid"] for m in replaced_meta.values()}
    diff_rows = []
    for i, slide in enumerate(prs.slides, 1):
        eid = id(slide._element)
        after_snap = du.extract_slide_snapshot(slide, prs)
        if eid in replaced_meta:
            rm = replaced_meta[eid]
            diff_rows.append({"slide": i, "orig_slide": rm["orig_slide"], "status": "replaced",
                              "title_before": rm["title_before"], "title_after": after_snap["title"],
                              "changes": [f"replaced slide {rm['orig_slide']} ('{rm['title_before'][:30]}')"]})
        elif eid not in before_snaps:
            diff_rows.append({"slide": i, "orig_slide": None, "status": "added",
                              "title_after": after_snap["title"], "changes": ["new slide"]})
        else:
            b = before_snaps[eid]
            changes = []
            if b["orig_slide"] != i:
                changes.append(f"reordered {b['orig_slide']} -> {i}")
            if b["title"] != after_snap["title"]:
                changes.append(f"title: '{b['title'][:40]}' -> '{after_snap['title'][:40]}'")
            if b["bullets"] != after_snap["bullets"]:
                changes.append(f"body paragraphs: {len(b['bullets'])} -> {len(after_snap['bullets'])}")
            if b["tables"] != after_snap["tables"]:
                changes.append("table updated")
            if b["charts"] != after_snap["charts"]:
                changes.append("chart updated")
            if b.get("image_sigs", b["images"]) != after_snap.get("image_sigs", after_snap["images"]):
                changes.append("image updated")
            if b["notes"] != after_snap["notes"] and after_snap["notes"] != default_kept_note:
                changes.append("notes updated")
            status_lbl = "modified" if changes else "unchanged"
            diff_rows.append({"slide": i, "orig_slide": b["orig_slide"], "status": status_lbl,
                              "title_before": b["title"], "title_after": after_snap["title"],
                              "changes": changes})
    deleted = [{"orig_slide": before_snaps[eid]["orig_slide"], "status": "deleted",
                "title_before": before_snaps[eid]["title"]}
               for eid in before_order if eid not in after_eids and eid not in replaced_old_eids]
    slide_diff = {"before_count": len(before_order), "after_count": len(prs.slides),
                  "slides": diff_rows, "deleted": deleted, "ops_applied": cm_changes,
                  "zero_hit_edits": zero_hit_edits}
    (work / "slide_diff.json").write_text(
        json.dumps(slide_diff, ensure_ascii=False, indent=1), encoding="utf-8")

    return specs_by_eid, sources_by_eid


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", required=True)
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--out", default="/app/output/presentation.pptx")
    ap.add_argument("--template", default=None)
    ap.add_argument("--start", default=None)
    args = ap.parse_args()

    work = Path(args.work)
    plan = du.normalize_plan(json.loads(Path(args.plan).read_text(encoding="utf-8")))
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
    elif template and Path(template).is_file():
        prs = Presentation(str(template))
        while len(prs.slides) > 0:
            du.delete_slide(prs, 0)
    else:
        prs = Presentation()
        print("BUILD_WARN: no template found, using default blank template")

    cm_edits: list = []
    theme_src = template or start
    accent = du.theme_accent_color(prs, theme_src)
    palette = du.theme_palette(prs, theme_src)
    du.load_master_styles(prs, theme_src)
    patch_next_partname(prs)
    wo_lang = str((inv.get("workorder") or {}).get("language") or plan.get("language") or "zh")
    prs._deck_lang = "en" if wo_lang.lower().startswith("en") else "zh"
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    cm = []
    if mode == "edit":
        specs_by_eid, sources_by_eid = apply_edit_ops(
            prs, plan.get("edits") or [], images_base, work, cm_edits, accent, palette
        )
        for i, slide in enumerate(prs.slides, 1):
            eid = id(slide._element)
            spec = specs_by_eid.get(eid) or {}
            snap = du.extract_slide_snapshot(slide, prs)
            # Prefer live post-edit snapshot bullets so replace_text / set_two_col / set_body
            # changes are always reflected accurately in content_map & number_fidelity checks.
            spec_bullets = []
            tc_sp = spec.get("two_col") or {}
            if tc_sp:
                if tc_sp.get("left_title"):
                    spec_bullets.append(str(tc_sp["left_title"]))
                spec_bullets.extend(b.get("t") if isinstance(b, dict) else str(b) for b in (tc_sp.get("left_bullets") or []))
                if tc_sp.get("right_title"):
                    spec_bullets.append(str(tc_sp["right_title"]))
                spec_bullets.extend(b.get("t") if isinstance(b, dict) else str(b) for b in (tc_sp.get("right_bullets") or []))
            spec_bullets.extend(b.get("t") if isinstance(b, dict) else str(b)
                                for b in (spec.get("bullets") or []))
            snap_tables = [{"rows": t.get("all_rows") or t.get("preview") or []} for t in (snap.get("tables") or [])]
            tbl_cm = spec.get("table") or (snap_tables[0] if snap_tables else {})
            snap_charts = list(snap.get("charts") or [])
            ch_cm = spec.get("chart") or (snap_charts[0] if snap_charts else {})
            cm.append({
                "slide": i,
                "role": spec.get("role", "existing"),
                "layout": snap["layout"],
                "title": snap["title"],
                "subtitle": spec.get("subtitle", ""),
                "bullets": snap["bullets"] if snap["bullets"] else spec_bullets,
                "notes": snap["notes"],
                "sources": (spec.get("sources") or []) + sources_by_eid.get(eid, []),
                "image": spec.get("image") or {},
                "images": spec.get("images") or [],
                "table": tbl_cm,
                "tables": snap_tables[1:] if len(snap_tables) > 1 else [],
                "chart": ch_cm,
                "charts": snap_charts[1:] if len(snap_charts) > 1 else [],
                "highlights": spec.get("highlights") or [],
                "steps": spec.get("steps") or [],
                "takeaway": spec.get("takeaway") or "",
            })
    else:
        for spec in plan.get("slides") or []:
            info = build_create_slide(prs, spec, images_base, work, accent, palette)
            img_spec = dict(spec.get("image") or {})
            if info.pop("image_missing", False):
                img_spec["missing"] = True
            bullets_flat = []
            two_col = spec.get("two_col") or {}
            if two_col:
                if two_col.get("left_title"):
                    bullets_flat.append(str(two_col["left_title"]))
                bullets_flat.extend(
                    b.get("t") if isinstance(b, dict) else str(b)
                    for b in (two_col.get("left_bullets") or [])
                )
                if two_col.get("right_title"):
                    bullets_flat.append(str(two_col["right_title"]))
                bullets_flat.extend(
                    b.get("t") if isinstance(b, dict) else str(b)
                    for b in (two_col.get("right_bullets") or [])
                )
            bullets_flat.extend(
                b.get("t") if isinstance(b, dict) else str(b)
                for b in (spec.get("bullets") or [])
            )
            cm.append({
                "slide": len(prs.slides),
                **info,
                "subtitle": spec.get("subtitle", ""),
                "bullets": bullets_flat,
                "notes": spec.get("notes", ""),
                "sources": spec.get("sources") or [],
                "image": img_spec,
                "images": spec.get("images") or [],
                "table": spec.get("table") or {},
                "chart": spec.get("chart") or {},
                "highlights": spec.get("highlights") or [],
                "steps": spec.get("steps") or [],
                "takeaway": spec.get("takeaway") or "",
            })

    # Set core metadata BEFORE saving so it persists in the PPTX package
    try:
        cp = prs.core_properties
        cp.title = plan.get("deck_title") or "Academic Deck"
        cp.author = "academic-slides-pku skill"
        cp.subject = "paper-to-slides with evidence mapping"
    except Exception:
        pass
    prs.save(str(out))

    (work / "content_map.json").write_text(
        json.dumps({"deck_title": plan.get("deck_title", ""), "mode": mode,
                    "out": str(out), "slides": cm, "edits_applied": cm_edits},
                   ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
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
