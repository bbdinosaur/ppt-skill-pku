#!/usr/bin/env python3
"""S6 result.json assembler: merges contract fields from case.json + work artifacts into result.json.

Usage:
  python3 build_result.py --case /input/case.json --work /app/work --out /app/output/result.json \
      [--fill /app/work/result_fill.json]

result_fill.json (agent-authored, honest values): any keys to override/extend, e.g.
  {"report_goal": "...", "audience": "...", "tradeoffs": [...], "uncertainty": [...], "limits": [...]}
Core fields auto-derived: case_id/status/sources/deliverables/slide_plan/evidence_map/qa_prep/rendering/changes.
Exit: 0 always (writes honest partial result), prints missing contract fields so agent can complete.
"""
import argparse
import json
import re
from pathlib import Path

META_WRAPPER_KEYS = {
    "required_fields", "optional_fields", "fields", "properties", "schema",
    "format", "type", "description", "notes", "version", "strict",
}


def find_contract(obj, path=""):
    hits = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}" if path else str(k)
            if any(nd in str(k).lower() for nd in ("output_contract", "output_field", "required_output", "result_schema", "required_fields")):
                hits.append((p, v))
            hits += find_contract(v, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            hits += find_contract(v, f"{path}[{i}]")
    return hits


def contract_fields(case):
    fields = []
    for p, v in find_contract(case):
        if isinstance(v, dict):
            for k, vv in v.items():
                kl = str(k).lower()
                if kl in ("required_fields", "fields", "optional_fields") and isinstance(vv, list):
                    fields += [str(x) for x in vv if isinstance(x, str)]
                elif kl == "properties" and isinstance(vv, dict):
                    fields += [str(x) for x in vv.keys()]
                elif kl not in META_WRAPPER_KEYS:
                    fields.append(str(k))
        elif isinstance(v, list):
            fields += [str(x) for x in v if isinstance(x, str)]
    seen, out = set(), []
    for f in fields:
        if f not in seen and f.lower() not in META_WRAPPER_KEYS:
            seen.add(f)
            out.append(f)
    return out


def collect_sources(case_path: Path, inspect: dict, plan_meta: dict, ingest: dict = None):
    srcs = []
    seen = set()

    def add(kind, path_str):
        if not path_str or path_str in seen:
            return
        seen.add(path_str)
        srcs.append({"type": kind, "path": str(path_str), "exists": Path(path_str).exists()})

    if case_path.exists():
        add("case", str(case_path))
    if isinstance(inspect, dict):
        sp = (inspect.get("source_pdf") or {}).get("path")
        if sp:
            add("source_pdf", sp)
        for pf in inspect.get("pptx_files") or []:
            add("pptx_input", pf.get("path"))
    if isinstance(ingest, dict):
        for ep in ingest.get("extra_pdfs") or []:
            add("extra_pdf", ep)
        for np_src in ingest.get("non_pdf_sources") or []:
            add("source_document", np_src)
    if plan_meta.get("template"):
        add("template", plan_meta["template"])
    if plan_meta.get("starting_deck"):
        add("starting_deck", plan_meta["starting_deck"])
    return srcs


def compute_source_coverage_audit(outline: dict, tables: list, plan_slides: list):
    """Cross-reference extracted source figures/tables against slide evidence/visuals for editorial audit."""
    def norm_item(label: str):
        s = str(label or "").strip()
        m_f = re.search(r"(?:Fig(?:ure)?\.?\s*(S?\d+)|图\s*(\d+))", s, re.I)
        if m_f:
            return ("figure", (m_f.group(1) or m_f.group(2)).upper(), f"Fig.{m_f.group(1) or m_f.group(2)}")
        m_t = re.search(r"(?:Tab(?:le)?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", s, re.I)
        if m_t:
            return ("table", (m_t.group(1) or m_t.group(2)).upper(), f"Table {m_t.group(1) or m_t.group(2)}")
        return None

    source_items = {}
    for c in (outline or {}).get("captions") or []:
        n = norm_item(c.get("label"))
        if n:
            source_items[(n[0], n[1])] = {"label": n[2], "page": c.get("page"), "caption": str(c.get("caption", ""))[:70]}
    for tb in (tables or []):
        n = norm_item(tb.get("label"))
        if n and (n[0], n[1]) not in source_items:
            source_items[(n[0], n[1])] = {"label": n[2], "page": tb.get("page"), "caption": str(tb.get("caption", ""))[:70]}

    cited_map = {}
    for s in plan_slides or []:
        pno = s.get("pptx_slide")
        role = s.get("role")
        blobs = [str(e.get("locator", "")) for e in (s.get("evidence") or [])]
        if s.get("figure_label"):
            blobs.append(str(s["figure_label"]))
        for v in (s.get("visuals") or []):
            blobs.append(str(v))
        joined = " ; ".join(blobs)
        for m in re.finditer(r"(?:Fig(?:ure)?\.?\s*(S?\d+)|图\s*(\d+))", joined, re.I):
            key = ("figure", (m.group(1) or m.group(2)).upper())
            cited_map.setdefault(key, []).append({"slide": pno, "role": role})
        for m in re.finditer(r"(?:Tab(?:le)?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", joined, re.I):
            key = ("table", (m.group(1) or m.group(2)).upper())
            cited_map.setdefault(key, []).append({"slide": pno, "role": role})

    included, omitted = [], []
    for key, info in source_items.items():
        hits = cited_map.get(key) or []
        if hits:
            uniq_slides = sorted({h["slide"] for h in hits if h.get("slide")})
            included.append({"item": info["label"], "page": info["page"], "slides": uniq_slides})
        else:
            omitted.append({"item": info["label"], "page": info["page"], "note": "omitted or compressed (check tradeoffs)"})

    # Section-level selection ledger (maps source paper sections -> final slides)
    sec_mapped, sec_omitted = [], []
    slide_sec_blobs = []
    for s in plan_slides or []:
        pno = s.get("pptx_slide")
        blob = " ; ".join(str(e.get("locator", "")) for e in (s.get("evidence") or []))
        slide_sec_blobs.append((pno, blob))
    seen_secs = set()
    for sec in (outline or {}).get("sections") or []:
        hdg = str(sec.get("heading") or "").strip()
        if not hdg or hdg.lower() in seen_secs:
            continue
        seen_secs.add(hdg.lower())
        m_num = re.match(r"^(?:#+\s*)?(\d+(?:\.\d+)?|[IVX]{1,4})\b", hdg)
        sec_num = m_num.group(1) if m_num else None
        pg = sec.get("page")
        hits = []
        for pno, blob in slide_sec_blobs:
            if not pno or not blob:
                continue
            if sec_num and re.search(rf"(?:§\s*|Sec(?:tion)?\.?\s*|第\s*){re.escape(sec_num)}(?:\.\d+)?\b", blob, re.I):
                hits.append(pno)
            elif pg and re.search(rf"(?:^|[\s,;(（/])p\.?\s*{int(pg)}\b", blob, re.I):
                hits.append(pno)
            elif re.search(r"(?:Intro|Method|Experiment|Result|Limitation|Conclusion|Abstract|引言|方法|实验|结果|局限|结论)", hdg, re.I):
                kw = re.search(r"(Intro\w*|Method\w*|Experiment\w*|Result\w*|Limitation\w*|Conclusion\w*|Abstract|引言|方法|实验|结果|局限|结论)", hdg, re.I).group(1)
                if re.search(re.escape(kw[:5]), blob, re.I):
                    hits.append(pno)
        uniq = sorted(set(hits))
        if uniq:
            sec_mapped.append({"section": hdg[:60], "page": pg, "slides": uniq})
        else:
            sec_omitted.append({"section": hdg[:60], "page": pg})

    return {
        "source_items_detected": len(source_items),
        "included": included,
        "omitted_or_deferred": omitted,
        "sections_mapped": sec_mapped[:18],
        "sections_omitted_or_compressed": sec_omitted[:12],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="/input/case.json")
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--out", default="/app/output/result.json")
    ap.add_argument("--fill")
    a = ap.parse_args()
    work = Path(a.work)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    case_path = Path(a.case)

    case = {}
    if case_path.exists():
        case = json.loads(case_path.read_text(encoding="utf-8"))
    case_id = case.get("case_id") or case.get("id") or "unknown-case"

    def load(name):
        p = work / name
        return json.loads(p.read_text(encoding="utf-8")) if p.exists() else None

    inspect = load("inspect.json") or {}
    ingest = load("ingest_summary.json") or {}
    outline = load("paper_outline.json") or {}
    paper_tables = load("paper_tables.json") or []
    render_meta = load("render_meta.json") or {}
    plan = load("slide_plan.json") or {"meta": {}, "slides": []}
    plan_meta = plan.get("meta") or {}
    qa = load("qa_plan.json") or []
    fill = {}
    if a.fill and Path(a.fill).exists():
        fill = json.loads(Path(a.fill).read_text(encoding="utf-8"))
    lint = load("lint_report.json")
    numchk = load("number_check.json")
    evchk = load("evidence_check.json")
    budchk = load("budget_check.json")
    qaud = load("quality_audit.json")
    changes = load("changes.json") or ["v1: initial build per SKILL.md pipeline"]

    outdir = out.parent
    deck = plan_meta.get("deck")
    deliver = {"deck": deck}
    pdf = previews = None
    if deck:
        dp = Path(deck)
        pdf = dp.with_suffix(".pdf")
        if pdf.exists():
            deliver["pdf"] = str(pdf)
        def _prev_key(p: Path):
            m = re.search(r"(\d+)$", p.stem)
            return (int(m.group(1)) if m else 9999, p.name)
        seen_p, prevs = set(), []
        for pv in sorted(outdir.glob("preview*.png"), key=_prev_key):
            if pv not in seen_p:
                seen_p.add(pv)
                prevs.append(str(pv))
        if prevs:
            previews = prevs
            deliver["previews"] = previews
    if (outdir / "speaker_notes.md").exists():
        deliver["speaker_notes"] = str(outdir / "speaker_notes.md")
    else:
        notes_md = sorted(outdir.glob("speaker_notes.*"))
        if notes_md:
            deliver["speaker_notes"] = str(notes_md[0])
    if (outdir / "speaker_notes.pdf").exists():
        deliver["speaker_notes_pdf"] = str(outdir / "speaker_notes.pdf")
    if (work / "slide_plan.json").exists():
        deliver["slide_plan"] = str(work / "slide_plan.json")
    if (work / "deck_spec.json").exists():
        deliver["deck_spec"] = str(work / "deck_spec.json")

    pages = None
    if pdf and Path(pdf).exists():
        try:
            try:
                import pymupdf as fitz
            except ImportError:
                import fitz
            pages = fitz.open(str(pdf)).page_count
        except Exception:
            pages = len(previews) if previews else None
    fonts = plan_meta.get("fonts") or {}
    plan_n = len(plan.get("slides", []))
    deck_n = None
    if deck and Path(deck).exists():
        try:
            from pptx import Presentation
            deck_n = len(Presentation(deck).slides)
        except Exception:
            pass
    if deck_n is not None and plan_n and deck_n != plan_n:
        consistency = f"deck has {deck_n} slides but slide_plan has {plan_n} entries; investigate"
    else:
        consistency = "ok"
    rendering = {
        "consistency": consistency,
        "fonts": fonts,
        "non_embedded_fonts": render_meta.get("non_embedded_fonts", []),
        "render_dpi": render_meta.get("dpi"),
        "pdf_pages": pages,
        "previews_count": len(previews) if previews else 0,
        "layout_lint": (lint or {}).get("n_issues", "not_run"),
        "lint_issues": (lint or {}).get("issues", [])[:20],
        "lint_warnings": (lint or {}).get("warnings", [])[:20],
        "low_res_figures": (lint or {}).get("low_res_figures", []),
        "symbol_warnings": (lint or {}).get("symbol_warnings", []),
        "number_check": ("not_run" if numchk is None else
                         {"checked": numchk.get("checked"), "suspects": numchk.get("n_suspects")}),
        "evidence_check": ("not_run" if evchk is None else
                           {"checked": evchk.get("checked"), "dangling": evchk.get("n_dangling")}),
        "budget_check": ("not_run" if budchk is None else
                         {"audience_class": budchk.get("audience_class"),
                          "counts": budchk.get("counts"),
                          "issues": len(budchk.get("issues") or []),
                          "warnings": len(budchk.get("warnings") or [])}),
    }
    if render_meta and render_meta.get("ok") is False:
        rendering["export_status"] = "failed"
        rendering["export_error"] = render_meta.get("error")

    traced_by_slide = (numchk or {}).get("traced_by_slide") or {}
    evidence_map = []
    strength_counts = {"S1": 0, "S2": 0, "S3": 0, "NA": 0}
    content_slides_n = 0
    content_with_ev_n = 0
    content_with_vis_n = 0
    for s in plan.get("slides", []):
        evs = s.get("evidence") or []
        role = s.get("role")
        if role not in ("title", "thanks", "qa_backup", "preserved"):
            content_slides_n += 1
            if any(e.get("locator") for e in evs):
                content_with_ev_n += 1
            if s.get("visuals") or s.get("figure") or s.get("has_table"):
                content_with_vis_n += 1
        for e in evs:
            st = str(e.get("strength", "")).upper()
            if st in strength_counts:
                strength_counts[st] += 1
        pno_key = str(s.get("pptx_slide") or "")
        ev_entry = {
            "pptx_slide": s.get("pptx_slide"),
            "role": role,
            "title": s.get("title"),
            "claim": s.get("claim"),
            "visuals": s.get("visuals") or ([s["figure_label"]] if s.get("figure_label") else []),
            "evidence": evs,
        }
        vnums = traced_by_slide.get(pno_key)
        if vnums:
            ev_entry["verified_numbers"] = vnums
        evidence_map.append(ev_entry)

    evidence_coverage = {
        "content_slides": content_slides_n,
        "slides_with_evidence": content_with_ev_n,
        "coverage_pct": round(100.0 * content_with_ev_n / max(1, content_slides_n), 1),
        "slides_with_visuals": content_with_vis_n,
        "visual_diversity_pct": round(100.0 * content_with_vis_n / max(1, content_slides_n), 1),
        "strength_counts": strength_counts,
        "s3_overclaim_warnings": (lint or {}).get("s3_overclaim_warnings", 0),
        "dangling_locators": (evchk or {}).get("n_dangling", "not_run"),
        "untraced_numbers": (numchk or {}).get("n_suspects", "not_run"),
    }

    slide_plan_out = [
        {"pptx_slide": s.get("pptx_slide"), "role": s.get("role"), "title": s.get("title"),
         "claim": s.get("claim"), "time_s": s.get("time_s"), "notes_chars": s.get("notes_chars"),
         "visuals": s.get("visuals") or ([s["figure_label"]] if s.get("figure_label") else []),
         "evidence_locators": [e.get("locator") for e in (s.get("evidence") or []) if e.get("locator")]}
        for s in plan.get("slides", [])
    ]

    total_timed_s = sum(int(s.get("time_s") or 0) for s in plan.get("slides", [])
                        if s.get("role") not in ("qa_backup", "preserved"))
    total_notes_units = sum(float(s.get("notes_units") or 0.0) for s in plan.get("slides", [])
                            if s.get("role") not in ("title", "thanks", "qa_backup", "preserved"))
    content_timed_s = sum(int(s.get("time_s") or 0) for s in plan.get("slides", [])
                          if s.get("role") not in ("title", "thanks", "qa_backup", "preserved"))
    budget_s = int(float(plan_meta.get("duration_minutes") or 0) * 60) if plan_meta.get("duration_minutes") else None
    pacing_status = "on_target"
    if budget_s and total_timed_s > budget_s * 1.12:
        pacing_status = "over_budget"
    elif budget_s and total_timed_s < budget_s * 0.65:
        pacing_status = "under_budget"
    timing_summary = {
        "total_timed_s": total_timed_s,
        "budget_s": budget_s,
        "pacing_status": pacing_status,
        "speaking_rate_units_per_min": round(total_notes_units / max(0.5, content_timed_s / 60.0), 1) if content_timed_s else 0.0,
        "total_slides": len(slide_plan_out),
        "backup_slides": sum(1 for s in slide_plan_out if s.get("role") == "qa_backup"),
    }

    page_budget_obj = load("page_budget.json") or {}
    audience_adaptation = {
        "audience": fill.get("audience") or plan_meta.get("audience") or case.get("audience"),
        "audience_class": (budchk or {}).get("audience_class") or page_budget_obj.get("audience_class"),
        "category_counts": (budchk or {}).get("counts"),
        "avg_seconds_per_content_slide": round(total_timed_s / max(1, content_slides_n), 1),
        "strategy": page_budget_obj.get("audience_strategy") or {},
        "budget_warnings": [w.get("msg") for w in (budchk or {}).get("warnings", [])][:6],
    }

    # Validate QA slide references and auto-bind backup_slide when a qa_backup slide exists
    total_slides_n = len(slide_plan_out)
    backup_nums = [s.get("pptx_slide") for s in slide_plan_out if s.get("role") == "qa_backup" and s.get("pptx_slide")]
    qa_out_of_range = []
    qa_clean = []
    for qi, qitem in enumerate(qa if isinstance(qa, list) else []):
        if not isinstance(qitem, dict):
            continue
        qcopy = dict(qitem)
        if qcopy.get("backup_slide") in (None, "") and backup_nums:
            qcopy["backup_slide"] = backup_nums[0]
        for ref_k in ("cite_slide", "backup_slide"):
            rv = qcopy.get(ref_k)
            if isinstance(rv, int) and total_slides_n and not (1 <= rv <= total_slides_n):
                qa_out_of_range.append(f"qa[{qi}].{ref_k}={rv} outside 1..{total_slides_n}")
        qa_clean.append(qcopy)
    qa = qa_clean

    complete_core = bool(deck) and Path(deck).exists() if deck else False
    scanned_no_ocr = bool(ingest.get("likely_scanned") and not ingest.get("ocr_used"))
    ptxt_size = (work / "paper.txt").stat().st_size if (work / "paper.txt").exists() else 0
    auto_limits = list(fill.get("limits", []))
    if scanned_no_ocr and ptxt_size < 200:
        complete_core = False
        if not any("scan" in str(x).lower() or "扫描" in str(x) for x in auto_limits):
            auto_limits.append("Source PDF appears scanned with <200 bytes extractable text and no OCR; verify figure/page content manually.")
    status = "complete" if complete_core else "partial"

    source_coverage_audit = compute_source_coverage_audit(outline, paper_tables, plan.get("slides", []))
    starting_deck_diff = None
    if plan_meta.get("mode") == "edit_deck" or plan_meta.get("orig_slides_count"):
        orig_n = int(plan_meta.get("orig_slides_count") or 0)
        pruned_list = list(plan_meta.get("pruned_orig_slides") or [])
        mapping, kept_n, edited_n, added_n, reordered_n = [], 0, 0, 0, 0
        for s in plan.get("slides", []):
            os_no = s.get("orig_slide")
            fs_no = s.get("pptx_slide")
            act = str(s.get("edit_action") or "new")
            if os_no is not None:
                if act == "keep":
                    kept_n += 1
                else:
                    edited_n += 1
                if fs_no is not None and os_no != fs_no:
                    reordered_n += 1
                mapping.append({"orig_slide": os_no, "final_slide": fs_no, "action": act, "title": str(s.get("title", ""))[:60]})
            elif s.get("role") != "preserved":
                added_n += 1
                mapping.append({"orig_slide": None, "final_slide": fs_no, "action": act, "title": str(s.get("title", ""))[:60]})
        starting_deck_diff = {
            "orig_slides": orig_n,
            "final_slides": len(slide_plan_out),
            "kept": kept_n,
            "edited": edited_n,
            "reordered": reordered_n,
            "added": added_n,
            "pruned_orig_slides": pruned_list,
            "mapping": mapping,
        }
        source_coverage_audit["starting_deck_diff"] = starting_deck_diff
        if pruned_list:
            source_coverage_audit["pruned_starting_deck_slides"] = pruned_list

    qa_cats = sorted({str(q.get("category")).strip() for q in qa if isinstance(q, dict) and q.get("category")})
    qa_cats_l = {c.lower() for c in qa_cats}
    qa_cats_missing = [c for c in ("mechanism", "data", "baseline", "limitation") if c not in qa_cats_l]
    qa_num_adv = (numchk or {}).get("qa_suspects", []) if isinstance(numchk, dict) else []

    auto_uncertainty = list(fill.get("uncertainty", [])) if isinstance(fill.get("uncertainty"), list) else (
        [str(fill["uncertainty"])] if fill.get("uncertainty") else []
    )
    lbl_warn = ingest.get("label_mismatch_warning") if isinstance(ingest, dict) else None
    if lbl_warn and not any(re.search(r"label|domain|topic|标签|领域", str(x), re.I) for x in auto_uncertainty):
        auto_uncertainty.append(f"Auto-detected label discrepancy: {lbl_warn}")
    if plan_meta.get("mode") == "no_template" and not any(re.search(r"template|default|模板", str(x), re.I) for x in auto_uncertainty):
        auto_uncertainty.append("No organizer template or starting deck provided; built with default 16:9 academic theme.")

    raw_tradeoffs = (
        fill.get("tradeoffs")
        or plan_meta.get("tradeoffs")
        or fill.get("editorial_decisions")
        or fill.get("design_tradeoffs")
        or []
    )
    if isinstance(raw_tradeoffs, str):
        raw_tradeoffs = [x.strip() for x in re.split(r"[;\n；]+", raw_tradeoffs) if x.strip()]
    else:
        raw_tradeoffs = [str(x).strip() for x in raw_tradeoffs if str(x).strip()]
    if not raw_tradeoffs:
        # Auto-synthesize auditable structural tradeoffs from source_coverage_audit / starting_deck_diff so tradeoffs is never empty
        inc_items = [x.get("item") for x in (source_coverage_audit.get("included") or []) if x.get("item")]
        omit_items = [x.get("item") for x in (source_coverage_audit.get("omitted_or_deferred") or []) if x.get("item")]
        omit_secs = [x.get("section") for x in (source_coverage_audit.get("sections_omitted_or_compressed") or []) if x.get("section")]
        if inc_items:
            raw_tradeoffs.append(f"Prioritized key visual/table evidence on main slides: {', '.join(inc_items[:6])}.")
        if omit_items or omit_secs:
            om = ", ".join((omit_items + omit_secs)[:5])
            raw_tradeoffs.append(f"Compressed or deferred secondary source items ({om}) to stay within {timing_summary.get('budget_s') or total_timed_s}s and {len(slide_plan_out)}-slide budget.")
        if starting_deck_diff:
            raw_tradeoffs.append(
                f"Adapted starting deck ({starting_deck_diff['orig_slides']} -> {starting_deck_diff['final_slides']} slides): "
                f"kept={starting_deck_diff['kept']}, edited={starting_deck_diff['edited']}, reordered={starting_deck_diff['reordered']}, "
                f"added={starting_deck_diff['added']}, pruned={starting_deck_diff['pruned_orig_slides']}."
            )

    auto_goal = fill.get("report_goal") or plan_meta.get("goal")
    if not auto_goal:
        # Synthesize honest default verifiable objective from audience + headline claim if omitted
        headline = next((s.get("claim") for s in (plan.get("slides") or []) if s.get("role") in ("result", "conclusion", "title") and s.get("claim")), "")
        if headline:
            auto_goal = f"Deliver an evidence-bound talk for {audience_adaptation['audience']} demonstrating: {str(headline)[:140]}"

    qaud_summary = None
    if isinstance(qaud, dict):
        pa = qaud.get("pixel_audit") or {}
        ea = qaud.get("editorial_audit") or {}
        qaud_summary = {
            "pixel_check_ran": bool(pa.get("ran")),
            "pixel_slides_checked": pa.get("slides_checked", 0),
            "mean_body_ink_pct": pa.get("mean_ink_pct"),
            "pixel_flags": pa.get("flags") or [],
            "max_bullet_only_streak": ea.get("max_bullet_only_streak", 0),
            "unused_clean_figures": ea.get("unused_clean_figures") or [],
            "polish_opportunities": qaud.get("top_polish_opportunities") or [],
        }

    result = {
        "case_id": case_id,
        "status": status,
        "sources": collect_sources(case_path, inspect, plan_meta, ingest),
        "deliverables": deliver,
        "report_goal": auto_goal,
        "audience": audience_adaptation["audience"],
        "audience_adaptation": audience_adaptation,
        "tradeoffs": raw_tradeoffs,
        "timing_summary": timing_summary,
        "evidence_coverage": evidence_coverage,
        "source_coverage_audit": source_coverage_audit,
        "slide_plan": slide_plan_out,
        "evidence_map": evidence_map,
        "qa_prep": qa,
        "qa_integrity": {
            "total": len(qa),
            "categories": qa_cats,
            "categories_missing": qa_cats_missing,
            "backup_slides": backup_nums,
            "out_of_range": qa_out_of_range,
            "qa_numeric_advisories": len(qa_num_adv),
        },
        "rendering": rendering,
        "quality_audit": qaud_summary,
        "uncertainty": auto_uncertainty,
        "changes": fill.get("changes", changes),
        "limits": auto_limits,
    }
    for k, v in fill.items():
        if k == "uncertainty":
            result[k] = auto_uncertainty
        elif k == "tradeoffs":
            result[k] = raw_tradeoffs
        elif k == "limits":
            result[k] = auto_limits
        else:
            result[k] = v

    # Semantic alias map so case.json output_contract fields never silently turn null when data exists
    alias_sources = {
        "deck_path": str(deck) if deck else None,
        "pptx_path": str(deck) if deck else None,
        "presentation_path": str(deck) if deck else None,
        "deck_file": str(deck) if deck else None,
        "result_path": str(out),
        "pdf_path": deliver.get("pdf"),
        "preview_paths": deliver.get("previews"),
        "preview_images": deliver.get("previews"),
        "previews": deliver.get("previews"),
        "speaker_notes_path": deliver.get("speaker_notes"),
        "notes_path": deliver.get("speaker_notes"),
        "slides": slide_plan_out,
        "slide_outline": slide_plan_out,
        "page_plan": slide_plan_out,
        "evidence_mapping": evidence_map,
        "evidence_trace": evidence_map,
        "source_mapping": evidence_map,
        "qa_preparation": qa,
        "qa_list": qa,
        "qa_items": qa,
        "render_check": rendering,
        "export_record": rendering,
        "rendering_check": rendering,
        "goal": result.get("report_goal"),
        "objective": result.get("report_goal"),
        "design_tradeoffs": result.get("tradeoffs"),
        "editorial_decisions": result.get("tradeoffs"),
        "omissions": result.get("tradeoffs"),
        "selection_rationale": result.get("tradeoffs"),
        "narrative_order": slide_plan_out,
        "timing_plan": timing_summary,
        "speaker_notes": deliver.get("speaker_notes"),
        "change_log": result.get("changes"),
        "version_history": result.get("changes"),
        "limitations": result.get("limits"),
        "starting_deck_diff": starting_deck_diff,
        "deck_diff": starting_deck_diff,
        "slide_diff": starting_deck_diff,
        "revision_summary": starting_deck_diff,
    }

    cfields = contract_fields(case)
    missing = []
    for f in cfields:
        if f not in result or result.get(f) is None:
            fl = f.lower()
            if fl in alias_sources and alias_sources[fl] is not None:
                result[f] = alias_sources[fl]
            elif f not in result:
                result[f] = None
                missing.append(f)
    missing = [f for f in missing if result.get(f) is None]

    out.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    json.loads(out.read_text(encoding="utf-8"))
    print(f"result.json -> {out} status={status}")
    print(f"contract fields checked: {cfields}")
    if missing:
        print("STILL NULL (fill these):", ", ".join(missing))
    if not complete_core:
        print("WARN: deck missing -> status=partial; inspect build_log.json warnings")
    for d, v in deliver.items():
        if isinstance(v, str) and not Path(v).exists():
            print(f"WARN: deliverable path does not exist: {d}={v}")


if __name__ == "__main__":
    main()
