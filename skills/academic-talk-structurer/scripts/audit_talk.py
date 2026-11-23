#!/usr/bin/env python3
"""Deterministic Rendered-Pixel + Semantic Editorial Quality Auditor (non-blocking).

Inspects rendered preview PNGs (via optional PIL) and structured talk artifacts
(slide_plan.json, deck_spec.json, figs/manifest.json, paper_tables.json, qa_plan.json)
to give BOTH vision-capable and text-only agents objective visual & editorial polish guidance.

Usage:
  python3 audit_talk.py --work /app/work --out-dir /app/output [--spec /app/work/deck_spec.json] \
      [--out /app/work/quality_audit.json]

Exit: 0 always (advisory diagnostics only; hard errors belong to preflight gates).
"""
import argparse
import json
import re
from pathlib import Path


def _speaking_units(s: str) -> float:
    if not s:
        return 0.0
    cjk = sum(1 for c in s if "\u2e80" <= c <= "\uffef")
    latin = len(re.findall(r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", s))
    return float(cjk) + latin * 1.5


def _sort_prev_key(p: Path):
    m = re.search(r"(\d+)$", p.stem)
    return (int(m.group(1)) if m else 9999, p.name)


def audit_pixels(out_dir: Path, plan_slides: list) -> dict:
    """Compute objective per-slide visual telemetry from rendered preview-*.png via optional PIL."""
    previews = sorted(out_dir.glob("preview*.png"), key=_sort_prev_key)
    if not previews:
        return {"ran": False, "reason": "no preview PNGs found", "slides": [], "flags": []}
    try:
        from PIL import Image
    except ImportError:
        return {"ran": False, "reason": "Pillow not installed", "slides": [], "flags": []}

    slide_stats = []
    flags = []
    role_by_idx = {int(s.get("pptx_slide") or (i + 1)): str(s.get("role") or "") for i, s in enumerate(plan_slides or [])}

    for idx, p in enumerate(previews, 1):
        role = role_by_idx.get(idx, "")
        try:
            with Image.open(str(p)) as im:
                rgb = im.convert("RGB")
                # Downsample to ~320px width for fast, deterministic pixel stats (~3ms/slide)
                w0, h0 = rgb.size
                sw = min(320, max(64, w0))
                sh = max(36, int(round(h0 * sw / max(1, w0))))
                resample = getattr(getattr(Image, "Resampling", Image), "BILINEAR", 2)
                small = rgb.resize((sw, sh), resample)
                px = list(small.get_flattened_data()) if hasattr(small, "get_flattened_data") else list(small.getdata())
        except Exception:
            continue

        # Estimate slide background color from median of 4 interior-margin probe patches (avoids outer chrome)
        samples = []
        for ry in (0.18, 0.50, 0.82):
            for rx in (0.04, 0.96):
                samples.append(px[int(ry * (sh - 1)) * sw + int(rx * (sw - 1))])
        bg_r = sorted(c[0] for c in samples)[len(samples) // 2]
        bg_g = sorted(c[1] for c in samples)[len(samples) // 2]
        bg_b = sorted(c[2] for c in samples)[len(samples) // 2]

        def is_ink(c):
            return (abs(c[0] - bg_r) + abs(c[1] - bg_g) + abs(c[2] - bg_b)) > 48

        # Analyze safe interior body band (exclude outer 3% bleed & top 10%/bottom 8% template banners)
        y_top_b = int(sh * 0.14)
        y_bot_b = int(sh * 0.90)
        y_mid_b = int(sh * 0.55)
        x_l = int(sw * 0.03)
        x_r = int(sw * 0.97)
        tot_body = 0
        ink_body = 0
        ink_lower = 0
        tot_lower = 0
        for yy in range(y_top_b, y_bot_b):
            row_off = yy * sw
            for xx in range(x_l, x_r):
                tot_body += 1
                ink = is_ink(px[row_off + xx])
                if ink:
                    ink_body += 1
                if yy >= y_mid_b:
                    tot_lower += 1
                    if ink:
                        ink_lower += 1
        ink_pct = round(100.0 * ink_body / max(1, tot_body), 1)
        lower_ink_pct = round(100.0 * ink_lower / max(1, tot_lower), 1)

        # Edge-clipping check on left/right 2px strip inside body band (excluding full-width horizontal banners)
        edge_hits = 0
        edge_tot = max(1, 2 * (y_bot_b - y_top_b))
        for yy in range(y_top_b, y_bot_b):
            row_off = yy * sw
            if is_ink(px[row_off + 1]):
                edge_hits += 1
            if is_ink(px[row_off + sw - 2]):
                edge_hits += 1
        edge_ratio = edge_hits / edge_tot

        stat = {
            "slide": idx,
            "role": role,
            "ink_pct": ink_pct,
            "lower_half_ink_pct": lower_ink_pct,
            "edge_touch_pct": round(100.0 * edge_ratio, 1),
        }
        slide_stats.append(stat)

        if role not in ("title", "thanks"):
            if ink_pct < 1.8:
                flags.append({"slide": idx, "type": "near_blank_slide", "detail": f"Slide {idx} body ink={ink_pct}% (<1.8%)"})
            elif ink_pct > 55.0:
                flags.append({"slide": idx, "type": "visually_dense_slide", "detail": f"Slide {idx} body ink={ink_pct}% (>55%)"})
            if 0.04 < edge_ratio < 0.35:
                flags.append({"slide": idx, "type": "edge_clip_risk", "detail": f"Slide {idx} content touches outer margin ({edge_ratio:.0%})"})

    return {
        "ran": True,
        "slides_checked": len(slide_stats),
        "mean_ink_pct": round(sum(s["ink_pct"] for s in slide_stats) / max(1, len(slide_stats)), 1),
        "slides": slide_stats,
        "flags": flags,
    }


def _norm_asset_key(label: str):
    s = str(label or "").strip()
    mf = re.search(r"(?:Fig(?:ure)?\.?\s*(S?\d+)|图\s*(\d+))", s, re.I)
    if mf:
        return ("fig", (mf.group(1) or mf.group(2)).upper(), f"Fig.{mf.group(1) or mf.group(2)}")
    mt = re.search(r"(?:Tab(?:le)?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", s, re.I)
    if mt:
        return ("tab", (mt.group(1) or mt.group(2)).upper(), f"Table {mt.group(1) or mt.group(2)}")
    return None


def audit_editorial(work: Path, spec: dict, plan: dict, manifest: list, tables: list, qa: list) -> dict:
    """Audit narrative arc, visual diversity, source asset usage, quantitative claims, oral notes, and Q&A."""
    slides = (plan or {}).get("slides") or []
    meta = (plan or {}).get("meta") or (spec or {}).get("meta") or {}
    content = [s for s in slides if s.get("role") not in ("title", "thanks", "qa_backup", "preserved")]
    n_content = len(content)

    # 1) Visual diversity & consecutive bullet-only streak
    max_bullet_streak = 0
    cur_streak = 0
    bullet_only_slides = []
    vis_slides = 0
    for s in content:
        has_vis = bool(s.get("visuals") or s.get("figure") or s.get("has_table"))
        if has_vis:
            vis_slides += 1
            cur_streak = 0
        else:
            cur_streak += 1
            max_bullet_streak = max(max_bullet_streak, cur_streak)
            bullet_only_slides.append(s.get("pptx_slide"))
    vis_pct = round(100.0 * vis_slides / max(1, n_content), 1)

    # 2) Clean source figure / table utilization
    used_figs = {Path(str(s.get("figure") or "")).name for s in slides if s.get("figure")}
    cited_keys = set()
    for s in slides:
        blobs = [str(e.get("locator") or "") for e in (s.get("evidence") or [])]
        blobs += [str(v) for v in (s.get("visuals") or [])]
        if s.get("figure_label"):
            blobs.append(str(s["figure_label"]))
        for b in blobs:
            nk = _norm_asset_key(b)
            if nk:
                cited_keys.add((nk[0], nk[1]))

    unused_clean_figs = []
    for m in (manifest or []):
        if m.get("method") not in ("raster", "raster_composite", "vector_crop", "starting_deck_pic"):
            continue
        px = m.get("px") or [0, 0]
        if m.get("low_res") or (px[0] and px[0] < 360):
            continue
        fname = Path(str(m.get("file") or "")).name
        nk = _norm_asset_key(m.get("label") or "")
        if fname not in used_figs and (not nk or (nk[0], nk[1]) not in cited_keys):
            unused_clean_figs.append(f"{m.get('label') or fname} (p.{m.get('page')}, {px[0]}x{px[1]}px)")

    has_any_table_or_chart = any(
        s.get("has_table") or any(str(v).startswith(("table", "chart")) for v in (s.get("visuals") or []))
        for s in content
    )
    unused_tables = []
    for tb in (tables or []):
        nk = _norm_asset_key(tb.get("label") or "")
        if nk and (nk[0], nk[1]) not in cited_keys:
            unused_tables.append(f"{tb.get('label')} (p.{tb.get('page')})")

    # 3) Result claim quantitative specificity (ignore pure hardware/model IDs like 8xA100, FP16, ResNet-50)
    HW_MODEL_RE = re.compile(r"\b(?:\d+\s*[x×]\s*)?(?:[AHV]100|H200|RTX\s*\d+|FP\d+|BF16|INT\d+|ResNet[- ]?\d+|GPT[- ]?\d+|Llama[- ]?\d+|Qwen[- ]?\d+|ViT[-A-Z/\d]+)\b", re.I)
    unquant_results = []
    for s in content:
        if s.get("role") == "result":
            blob = HW_MODEL_RE.sub("", f"{s.get('title') or ''} {s.get('claim') or ''}")
            if not re.search(r"\d", blob):
                unquant_results.append(s.get("pptx_slide"))

    # 4) Oral script richness (speaking units + bridge)
    thin_notes = []
    for s in content:
        nu = float(s.get("notes_units") or _speaking_units(str(s.get("notes") or "")))
        if nu < 75.0:
            thin_notes.append((s.get("pptx_slide"), round(nu)))

    # 5) 4-dim Q&A grounding
    qa_list = qa if isinstance(qa, list) else []
    qa_cats = {str(q.get("category") or "").strip().lower() for q in qa_list if isinstance(q, dict)}
    qa_missing_cats = [c for c in ("mechanism", "data", "baseline", "limitation") if c not in qa_cats]
    qa_ungrounded = sum(
        1 for q in qa_list
        if isinstance(q, dict) and not re.search(r"(?:Fig|Tab|§|Sec|p\.|Slide|图|表|\d)", str(q.get("answer_key") or ""), re.I)
    )

    # 6) Explicit tradeoffs & checkable goal
    tradeoffs = meta.get("tradeoffs") or []
    goal = str(meta.get("goal") or "").strip()

    # Build prioritized polish opportunities (top 3 highest-leverage recommendations)
    opps = []
    if max_bullet_streak >= 3 or (n_content >= 4 and vis_pct < 50.0):
        opps.append(
            f"Visual diversity is {vis_pct}% (bullet-only slides: {bullet_only_slides[:5]}); "
            "add figure/table/chart/steps/kpis/two_col so no 3 consecutive slides are text-only."
        )
    if unused_clean_figs and not any(s.get("figure") for s in content):
        opps.append(
            f"Deck uses 0 extracted source figures despite clean asset(s) ready: {', '.join(unused_clean_figs[:3])}."
        )
    elif unused_tables and not has_any_table_or_chart:
        opps.append(
            f"Deck has no native table/chart despite extracted source table(s): {', '.join(unused_tables[:3])}."
        )
    if unquant_results:
        opps.append(
            f"Result slide(s) {unquant_results} lack concrete numbers in title/claim; put exact source metric + baseline delta in claim."
        )
    if thin_notes:
        sample = ", ".join(f"s{p}({u}u)" for p, u in thin_notes[:4])
        opps.append(
            f"Short oral notes on {sample} (<75 units); enrich with motivation, figure/table walkthrough, and transition sentence."
        )
    if len(qa_list) < 4 or qa_missing_cats:
        opps.append(
            f"Q&A has {len(qa_list)} item(s)"
            + (f", missing categories: {', '.join(qa_missing_cats)}" if qa_missing_cats else "")
            + "; cover mechanism/data/baseline/limitation with locator-grounded answer_keys."
        )
    elif qa_ungrounded > 0:
        opps.append(
            f"{qa_ungrounded} Q&A answer_key(s) lack concrete paper locators or numbers; anchor each answer to exact Fig/Table/§/metric."
        )
    if len(tradeoffs) < 2 or not goal:
        opps.append(
            "Strengthen editorial auditability: fill verifiable meta.goal and >=2 concrete source-linked meta.tradeoffs."
        )

    return {
        "content_slides": n_content,
        "visual_diversity_pct": vis_pct,
        "max_bullet_only_streak": max_bullet_streak,
        "unused_clean_figures": unused_clean_figs[:6],
        "unused_tables": unused_tables[:6],
        "unquantified_result_slides": unquant_results,
        "thin_notes_slides": [p for p, _ in thin_notes],
        "qa_total": len(qa_list),
        "qa_missing_categories": qa_missing_cats,
        "qa_ungrounded_answers": qa_ungrounded,
        "has_goal": bool(goal),
        "tradeoffs_count": len(tradeoffs),
        "polish_opportunities": opps,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--out-dir", default="/app/output")
    ap.add_argument("--spec")
    ap.add_argument("--out")
    a = ap.parse_args()

    work = Path(a.work)
    out_dir = Path(a.out_dir)

    def load_j(p: Path, default):
        if p and p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return default
        return default

    spec_p = Path(a.spec) if a.spec else (work / "deck_spec.json")
    spec = load_j(spec_p, {})
    plan = load_j(work / "slide_plan.json", {})
    manifest = load_j(work / "figs" / "manifest.json", [])
    tables = load_j(work / "paper_tables.json", [])
    qa = load_j(work / "qa_plan.json", [])

    pix = audit_pixels(out_dir, plan.get("slides") or [])
    ed = audit_editorial(work, spec, plan, manifest, tables, qa)

    top_opps = list(ed["polish_opportunities"])
    for fl in (pix.get("flags") or [])[:2]:
        top_opps.insert(0, f"[visual-pixel] {fl['detail']}")

    rep = {
        "pixel_audit": pix,
        "editorial_audit": ed,
        "top_polish_opportunities": top_opps[:4],
    }
    out_p = Path(a.out) if a.out else (work / "quality_audit.json")
    out_p.write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")

    pix_summary = (
        f"pixels={pix.get('slides_checked', 0)} slides (mean_ink={pix.get('mean_ink_pct', '?')}%, flags={len(pix.get('flags') or [])})"
        if pix.get("ran") else f"pixels=skipped({pix.get('reason', 'n/a')})"
    )
    print(
        f"quality audit: {pix_summary} | vis_div={ed['visual_diversity_pct']}% "
        f"| bullet_streak={ed['max_bullet_only_streak']} | thin_notes={len(ed['thin_notes_slides'])} "
        f"| qa={ed['qa_total']}(miss={len(ed['qa_missing_categories'])})"
    )
    if rep["top_polish_opportunities"]:
        print("== POLISH OPPORTUNITIES (advisory) ==")
        for opp in rep["top_polish_opportunities"][:3]:
            print(f"  * {opp}")


if __name__ == "__main__":
    main()
