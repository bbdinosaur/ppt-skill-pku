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

        # Estimate slide background color from median of outer-margin probe patches (outside the 4% content frame)
        samples = []
        for ry in (0.18, 0.50, 0.82):
            for rx in (0.015, 0.985):
                samples.append(px[int(ry * (sh - 1)) * sw + int(rx * (sw - 1))])
        bg_r = sorted(c[0] for c in samples)[len(samples) // 2]
        bg_g = sorted(c[1] for c in samples)[len(samples) // 2]
        bg_b = sorted(c[2] for c in samples)[len(samples) // 2]

        def is_ink(c):
            return (abs(c[0] - bg_r) + abs(c[1] - bg_g) + abs(c[2] - bg_b)) > 64

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


def _norm_asset_keys(label: str):
    """Extract all referenced figure/table asset keys (supports compound locators like 'Figs. 1-2' or 'Tables 1 & 2')."""
    s = str(label or "").strip()
    out = []
    for mf in re.finditer(r"(?:Fig(?:ure)?s?\.?\s*(S?\d+)(?:\s*(?:[-–—~&]|and)\s*(S?\d+))?|图\s*(\d+)(?:\s*[-–—~、和与]\s*(\d+))?)", s, re.I):
        for g in mf.groups():
            if g:
                out.append(("fig", g.upper(), f"Fig.{g}"))
    for mt in re.finditer(r"(?:Tab(?:le)?s?\.?\s*(S?\d+|[IVX]+)\b(?:\s*(?:[-–—~&]|and)\s*(S?\d+|[IVX]+)\b)?|表\s*(\d+)(?:\s*[-–—~、和与]\s*(\d+))?)", s, re.I):
        for g in mt.groups():
            if g:
                out.append(("tab", g.upper(), f"Table {g}"))
    return out


def audit_editorial(work: Path, spec: dict, plan: dict, manifest: list, tables: list, qa: list,
                    lint: dict = None, numchk: dict = None) -> dict:
    """Audit narrative arc, visual diversity, source asset usage, quantitative claims, pacing, oral notes, and Q&A."""
    slides = (plan or {}).get("slides") or []
    meta = (plan or {}).get("meta") or (spec or {}).get("meta") or {}
    content = [s for s in slides if s.get("role") not in ("title", "thanks", "qa_backup", "preserved")]
    n_content = len(content)
    total_slides = len(slides)

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
    used_figs = set()
    for s in slides:
        if s.get("figure"):
            used_figs.add(Path(str(s["figure"])).name)
        for fp_i in (s.get("figures") or []):
            if fp_i:
                used_figs.add(Path(str(fp_i)).name)
    cited_keys = set()
    for s in slides:
        blobs = [str(e.get("locator") or "") for e in (s.get("evidence") or [])]
        blobs += [str(v) for v in (s.get("visuals") or [])]
        if s.get("figure_label"):
            blobs.append(str(s["figure_label"]))
        for b in blobs:
            for nk in _norm_asset_keys(b):
                cited_keys.add((nk[0], nk[1]))

    unused_clean_figs = []
    clean_fig_by_key = {}
    for m in (manifest or []):
        if m.get("method") not in ("raster", "raster_composite", "vector_crop", "starting_deck_pic", "input_image"):
            continue
        px = m.get("px") or [0, 0]
        if m.get("low_res") or (px[0] and px[0] < 360):
            continue
        fname = Path(str(m.get("file") or "")).name
        nks = _norm_asset_keys(m.get("label") or "")
        for nk in nks:
            clean_fig_by_key.setdefault((nk[0], nk[1]), m)
        if fname not in used_figs and (not nks or not any((nk[0], nk[1]) in cited_keys for nk in nks)):
            unused_clean_figs.append(f"{m.get('label') or fname} (p.{m.get('page')}, {px[0]}x{px[1]}px)")

    cited_fig_missing_vis = []
    for s in content:
        if s.get("figure") or s.get("has_table") or any(str(v).startswith(("table", "chart", "figure")) for v in (s.get("visuals") or [])):
            continue
        for e in (s.get("evidence") or []):
            for nk in _norm_asset_keys(str(e.get("locator") or "")):
                if nk[0] == "fig" and (nk[0], nk[1]) in clean_fig_by_key:
                    m_hit = clean_fig_by_key[(nk[0], nk[1])]
                    cited_fig_missing_vis.append(f"s{s.get('pptx_slide')} cites {nk[2]} (embed `{m_hit.get('file')}`)")
                    break

    has_any_table_or_chart = any(
        s.get("has_table") or any(str(v).startswith(("table", "chart")) for v in (s.get("visuals") or []))
        for s in content
    )
    unused_tables = []
    for tb in (tables or []):
        nks = _norm_asset_keys(tb.get("label") or "")
        if nks and not any((nk[0], nk[1]) in cited_keys for nk in nks):
            unused_tables.append(f"{tb.get('label')} (p.{tb.get('page')})")

    # 3) Result claim quantitative specificity + TOC/overclaim/missing-evidence checks
    HW_MODEL_RE = re.compile(r"\b(?:\d+\s*[x×]\s*)?(?:[AHV]100|H200|RTX\s*\d+|FP\d+|BF16|INT\d+|ResNet[- ]?\d+|GPT[- ]?\d+|Llama[- ]?\d+|Qwen[- ]?\d+|ViT[-A-Z/\d]+)\b", re.I)
    TOC_TITLE_RE = re.compile(
        r"^(?:\d+\.?\s*)?(?:introduction|background|related\s+work|method(?:ology)?|experiments?|results?|discussion|limitations?|conclusion(?:s)?|overview|outline|引言|背景|相关工作|方法|实验|结果|讨论|局限|结论|目录)$",
        re.I,
    )
    unquant_results = []
    toc_title_slides = []
    uncited_content = []
    for s in content:
        pno = s.get("pptx_slide")
        if s.get("role") == "result":
            blob = HW_MODEL_RE.sub("", f"{s.get('title') or ''} {s.get('claim') or ''}")
            if not re.search(r"\d", blob):
                unquant_results.append(pno)
        sttl = str(s.get("title") or "").strip()
        if sttl and TOC_TITLE_RE.match(sttl):
            toc_title_slides.append(pno)
        if s.get("role") not in ("roadmap", "conclusion") and not any(
            str(e.get("locator") or "").strip() for e in (s.get("evidence") or []) if isinstance(e, dict)
        ):
            uncited_content.append(pno)

    s3_overclaims = int((lint or {}).get("s3_overclaim_warnings") or 0)
    verbatim_notes_slides = []
    title_claim_dup_slides = []
    for w_ln in (lint or {}).get("warnings") or []:
        ws = str(w_ln)
        mv = re.match(r"^s(\d+): notes repeat slide bullets verbatim", ws)
        if mv:
            verbatim_notes_slides.append(int(mv.group(1)))
        md = re.match(r"^s(\d+): title and claim are nearly identical", ws)
        if md:
            title_claim_dup_slides.append(int(md.group(1)))

    # Narrative non-TOC structure check: ensure risk/limitation coverage on >=5-content-slide decks and no 3+ background opening
    roles_seq = [str(s.get("role") or "").lower() for s in content]
    has_risk_slide = any(r in ("limitation", "progress") for r in roles_seq)
    bg_lead_run = 0
    for r in roles_seq:
        if r == "roadmap":
            continue
        if r in ("motivation", "related"):
            bg_lead_run += 1
        else:
            break
    narrative_issues = []
    if n_content >= 5 and not has_risk_slide:
        narrative_issues.append("Deck has >=5 content slides but no limitation/progress slide; add honest boundary/risk coverage.")
    if bg_lead_run >= 3:
        narrative_issues.append(f"Deck opens with {bg_lead_run} consecutive background/motivation slides; compress background and move to core mechanism/evidence faster.")

    # 4) Oral script richness & talk pacing alignment
    thin_notes = []
    total_timed_s = 0
    for s in slides:
        if s.get("role") not in ("qa_backup", "preserved") and s.get("time_s"):
            try:
                total_timed_s += int(round(float(s["time_s"])))
            except (TypeError, ValueError):
                pass
    for s in content:
        nu = float(s.get("notes_units") or _speaking_units(str(s.get("notes") or "")))
        if nu < 75.0:
            thin_notes.append((s.get("pptx_slide"), round(nu)))

    dur_min = meta.get("duration_minutes")
    budget_s = int(round(float(dur_min) * 60)) if dur_min else None
    timing_drift_pct = (
        round(100.0 * abs(total_timed_s - budget_s) / budget_s, 1)
        if (budget_s and budget_s > 0 and total_timed_s > 0) else 0.0
    )

    # 5) 4-dim Q&A grounding + cite_slide bounds + numeric caliber
    qa_list = qa if isinstance(qa, list) else []
    qa_cats = {str(q.get("category") or "").strip().lower() for q in qa_list if isinstance(q, dict)}
    qa_missing_cats = [c for c in ("mechanism", "data", "baseline", "limitation") if c not in qa_cats]
    qa_ungrounded = sum(
        1 for q in qa_list
        if isinstance(q, dict) and not re.search(r"(?:Fig|Tab|§|Sec|p\.|Slide|图|表|\d)", str(q.get("answer_key") or ""), re.I)
    )
    qa_bad_cites = [
        qi for qi, q in enumerate(qa_list, 1)
        if isinstance(q, dict) and isinstance(q.get("cite_slide"), int) and total_slides > 0
        and not (1 <= q["cite_slide"] <= total_slides)
    ]
    qa_num_suspects = (numchk or {}).get("qa_suspects") or []

    # 6) Explicit tradeoffs, checkable goal, and visual-walkthrough cues in notes
    tradeoffs = [str(t).strip() for t in (meta.get("tradeoffs") or []) if len(str(t).strip()) >= 10]
    goal = str(meta.get("goal") or "").strip()
    VIS_CUE_RE = re.compile(
        r"(?:fig(?:ure)?\.?\s*\d|tab(?:le)?\.?\s*[\dIVX]+|chart|plot|curve|diagram|panel|row|column|bar|left|right|top|bottom|card|highlighted|see\s+the|shown\s+(?:on|in|above|below)|图\s*\d?|表\s*\d?|左侧|右侧|上方|下方|曲线|柱状|对比|卡片|高亮|流程图)",
        re.I,
    )
    vis_no_cue_slides = []
    for s in content:
        has_fig_tbl_ch = bool(
            s.get("figure") or s.get("figures") or s.get("has_table")
            or any(str(v).startswith(("figure", "table", "chart")) for v in (s.get("visuals") or []))
        )
        if has_fig_tbl_ch:
            nt_full = str(s.get("notes_full") or s.get("notes") or "")
            if nt_full and not VIS_CUE_RE.search(nt_full):
                vis_no_cue_slides.append(s.get("pptx_slide"))

    # Build prioritized polish opportunities (highest-leverage recommendations first)
    opps = []
    if s3_overclaims > 0:
        opps.append(
            f"{s3_overclaims} slide(s) use S1-strong claim verbs on S3/NA evidence; soften wording or upgrade to quantitative S1/S2 evidence."
        )
    if uncited_content:
        opps.append(
            f"Content slide(s) {uncited_content[:5]} have no bound source evidence locator; attach exact Fig/Table/§/p.N."
        )
    if max_bullet_streak >= 3 or (n_content >= 4 and vis_pct < 50.0):
        opps.append(
            f"Visual diversity is {vis_pct}% (bullet-only slides: {bullet_only_slides[:5]}); "
            "add figure/table/chart/steps/kpis/two_col so no 3 consecutive slides are text-only."
        )
    if unused_clean_figs and not any(s.get("figure") for s in content):
        opps.append(
            f"Deck uses 0 extracted source figures despite clean asset(s) ready: {', '.join(unused_clean_figs[:3])}."
        )
    elif cited_fig_missing_vis:
        opps.append(
            f"Visual-evidence gap: {'; '.join(cited_fig_missing_vis[:2])}."
        )
    elif unused_tables and not has_any_table_or_chart:
        opps.append(
            f"Deck has no native table/chart despite extracted source table(s): {', '.join(unused_tables[:3])}."
        )
    if toc_title_slides:
        opps.append(
            f"Slide(s) {toc_title_slides[:5]} use generic paper-TOC titles (e.g., Method/Results); rewrite as takeaway headlines."
        )
    if title_claim_dup_slides:
        opps.append(
            f"Slide(s) {title_claim_dup_slides[:4]} repeat title verbatim in claim; make title the punchy headline and claim the specific evidence takeaway."
        )
    if narrative_issues:
        opps.append(narrative_issues[0])
    if unquant_results:
        opps.append(
            f"Result slide(s) {unquant_results} lack concrete numbers in title/claim; put exact source metric + baseline delta in claim."
        )
    if qa_num_suspects:
        sample_qn = ", ".join(f"qa#{x.get('qa')}:{x.get('value')}" for x in qa_num_suspects[:3])
        opps.append(
            f"Q&A answer_key has unverified number(s) ({sample_qn}); align every Q&A figure with source text."
        )
    if verbatim_notes_slides:
        opps.append(
            f"Slide(s) {verbatim_notes_slides[:4]} speaker notes repeat slide bullets verbatim; rewrite as spoken explanation + transition."
        )
    elif thin_notes:
        sample = ", ".join(f"s{p}({u}u)" for p, u in thin_notes[:4])
        opps.append(
            f"Short oral notes on {sample} (<75 units); enrich with motivation, figure/table walkthrough, and transition sentence."
        )
    if timing_drift_pct > 22.0 and budget_s:
        opps.append(
            f"Talk pacing mismatch: timed slides sum to {total_timed_s}s vs {budget_s}s budget ({timing_drift_pct}% off); rebalance time_s."
        )
    if len(qa_list) < 4 or qa_missing_cats:
        opps.append(
            f"Q&A has {len(qa_list)} item(s)"
            + (f", missing categories: {', '.join(qa_missing_cats)}" if qa_missing_cats else "")
            + "; cover mechanism/data/baseline/limitation with locator-grounded answer_keys."
        )
    elif qa_ungrounded > 0 or qa_bad_cites:
        opps.append(
            f"Tighten Q&A citations ({qa_ungrounded} ungrounded, bad cite_slide: {qa_bad_cites}); anchor each answer to exact Fig/Table/§ and valid slide."
        )
    if len(tradeoffs) < 2 or not goal:
        opps.append(
            "Strengthen editorial auditability: fill verifiable meta.goal and >=2 concrete source-linked meta.tradeoffs."
        )
    if len(vis_no_cue_slides) >= 2:
        opps.append(
            f"Visual-walkthrough cue missing in notes on slide(s) {vis_no_cue_slides[:4]}; tell the speaker how to guide the audience through the on-page figure/table/chart."
        )

    readiness = {
        "visual_diversity_ok": (max_bullet_streak < 3 and (n_content < 4 or vis_pct >= 50.0)),
        "source_assets_used": not (unused_clean_figs and not any(s.get("figure") for s in content)) and not cited_fig_missing_vis,
        "claims_calibrated": (not unquant_results and s3_overclaims == 0 and not uncited_content),
        "titles_distinct_takeaways": (not toc_title_slides and not title_claim_dup_slides),
        "notes_oral_ready": (not thin_notes and not verbatim_notes_slides),
        "timing_aligned": (timing_drift_pct <= 22.0),
        "qa_4dim_grounded": (len(qa_list) >= 4 and not qa_missing_cats and qa_ungrounded == 0 and not qa_bad_cites and not qa_num_suspects),
        "tradeoffs_auditable": (bool(goal) and len(tradeoffs) >= 2),
        "narrative_non_toc": (not narrative_issues),
    }

    return {
        "content_slides": n_content,
        "visual_diversity_pct": vis_pct,
        "max_bullet_only_streak": max_bullet_streak,
        "unused_clean_figures": unused_clean_figs[:6],
        "cited_fig_missing_vis": cited_fig_missing_vis[:4],
        "unused_tables": unused_tables[:6],
        "toc_title_slides": toc_title_slides,
        "title_claim_dup_slides": title_claim_dup_slides,
        "uncited_content_slides": uncited_content,
        "s3_overclaims": s3_overclaims,
        "unquantified_result_slides": unquant_results,
        "narrative_issues": narrative_issues,
        "thin_notes_slides": [p for p, _ in thin_notes],
        "verbatim_notes_slides": verbatim_notes_slides,
        "timing_drift_pct": timing_drift_pct,
        "qa_total": len(qa_list),
        "qa_missing_categories": qa_missing_cats,
        "qa_ungrounded_answers": qa_ungrounded,
        "qa_bad_cite_slides": qa_bad_cites,
        "qa_numeric_suspects": len(qa_num_suspects),
        "has_goal": bool(goal),
        "tradeoffs_count": len(tradeoffs),
        "readiness_checks": readiness,
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
    lint = load_j(work / "lint_report.json", {})
    numchk = load_j(work / "number_check.json", {})

    pix = audit_pixels(out_dir, plan.get("slides") or [])
    ed = audit_editorial(work, spec, plan, manifest, tables, qa, lint=lint, numchk=numchk)
    ed["readiness_checks"]["pixels_clean"] = (not pix.get("flags"))

    top_opps = list(ed["polish_opportunities"])
    for fl in (pix.get("flags") or [])[:2]:
        top_opps.insert(0, f"[visual-pixel] {fl['detail']}")

    rc = ed["readiness_checks"]
    passed_n = sum(1 for v in rc.values() if v)
    total_rc = len(rc)
    rep = {
        "pixel_audit": pix,
        "editorial_audit": ed,
        "readiness_score": f"{passed_n}/{total_rc}",
        "readiness_checks": rc,
        "top_polish_opportunities": top_opps[:5],
    }
    out_p = Path(a.out) if a.out else (work / "quality_audit.json")
    out_p.write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")

    pix_summary = (
        f"pixels={pix.get('slides_checked', 0)} slides (mean_ink={pix.get('mean_ink_pct', '?')}%, flags={len(pix.get('flags') or [])})"
        if pix.get("ran") else f"pixels=skipped({pix.get('reason', 'n/a')})"
    )
    print(
        f"quality audit [{passed_n}/{total_rc}]: {pix_summary} | vis_div={ed['visual_diversity_pct']}% "
        f"| bullet_streak={ed['max_bullet_only_streak']} | thin_notes={len(ed['thin_notes_slides'])} "
        f"| qa={ed['qa_total']}(miss={len(ed['qa_missing_categories'])})"
    )
    if rep["top_polish_opportunities"]:
        print(f"== POLISH SCOREBOARD: {passed_n}/{total_rc} | ACTIONABLE POLISH PLAN (fix in 1 round -> re-run deliver.py) ==")
        for idx_o, opp in enumerate(rep["top_polish_opportunities"][:5], 1):
            print(f"  {idx_o}. {opp}")
    else:
        print(f"== POLISH SCOREBOARD: {passed_n}/{total_rc} READY ==")


if __name__ == "__main__":
    main()
