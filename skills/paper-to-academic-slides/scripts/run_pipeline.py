#!/usr/bin/env python3
"""One-command pipeline wrapper for paper-to-academic-slides.

Two subcommands minimize agent tool-call overhead while preserving every
quality gate (layout geometry check + mandatory S4.5 evidence verification
+ PDF/PNG render check):

  prep  : S0 (resolve_case.py) + S1 (extract_paper.py)
  build : S3 (build_deck.py) + S4 (check_layout.py)
        + S4.5 (verify_evidence.py) + S5 (finalize_output.py)

All underlying scripts remain standalone CLIs in the same directory.
"""
import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent

HARD_EV_TYPES = {"image_ref_mismatch", "generic_section_title", "missing_notes"}
HARD_LAY_TYPES = {"template_placeholder_remnant"}
POLISH_EV_TYPES = {
    "number_absent_from_paper", "unit_mismatch_with_paper", "missing_limitations_slide",
    "notes_duplicate_slide_text", "schematic_unlabeled", "schematic_unmarked",
    "unattributed_comparison", "missing_quote_on_slide", "cover_title_mismatch_paper",
    "duplicate_figure_across_slides", "image_missing_caption", "quote_page_mismatch",
    "ref_page_mismatch", "pacing_budget_mismatch", "slide_line_ref_not_found",
    "embedded_figure_uncited", "dual_media_crowded",
}
POLISH_LAY_TYPES = {"text_overflow", "overlap", "text_overlap", "table_overflow", "too_dense", "master_footer_overlap"}


def _detect_label_discrepancy(cr, ctx):
    """Detect whether case.json domain/topic/title metadata contradicts the actual extracted paper.
    Only compares same-script labels to avoid false alarms on bilingual tasks."""
    labels = cr.get("case_labels") or {}
    if not labels or not isinstance(ctx, dict):
        return None
    pages = ctx.get("pages") or []
    p12 = " ".join((pg.get("text") or "") for pg in pages[:2])
    title_g = ctx.get("title_guess") or ""
    blob = (title_g + "\n" + p12).lower()
    if len(blob.strip()) < 120:
        return None
    blob_cjk = sum(1 for ch in blob if "\u4e00" <= ch <= "\u9fff") > 30
    stop_en = {
        "study", "paper", "research", "analysis", "report", "academic", "slides",
        "presentation", "science", "scientific", "review", "method", "model", "data",
        "system", "approach", "based", "using", "toward", "towards", "task", "case",
        "deck", "general", "other", "misc", "unknown", "demo", "test",
    }
    mismatches = []
    for k, raw_v in labels.items():
        v = str(raw_v or "").strip()
        if len(v) < 4:
            continue
        v_cjk = sum(1 for ch in v if "\u4e00" <= ch <= "\u9fff") >= 2
        if v_cjk != blob_cjk:
            continue
        if v_cjk:
            runs = re.findall(r"[\u4e00-\u9fff]{2,}", v)
            grams = list(dict.fromkeys(r[i:i + 2] for r in runs for i in range(max(0, len(r) - 1))))
            if len(grams) >= 2 and not any(g in blob for g in grams):
                mismatches.append(f"{k}='{v[:60]}'")
        else:
            toks = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z0-9\-]{3,}", v) if w.lower() not in stop_en]
            toks = list(dict.fromkeys(toks))
            if len(toks) >= 2 and not any(w in blob for w in toks):
                mismatches.append(f"{k}='{v[:60]}'")
    if not mismatches:
        return None
    return (
        f"case.json label(s) [{', '.join(mismatches)}] do not appear in actual paper pages 1-2 "
        f"(paper title_guess='{title_g[:80]}'); STRICTLY ground slides in the actual paper file, NOT the case label."
    )


def _run_py(script_name, args, timeout=240):
    cmd = [sys.executable, str(SCRIPTS_DIR / script_name)] + list(args)
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return proc.returncode, proc.stdout.strip(), proc.stderr.strip()
    except subprocess.TimeoutExpired:
        return 124, "", f"{script_name} timed out after {timeout}s"
    except Exception as e:
        return 1, "", str(e)


def _load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _pick_preview_samples(slides_spec, previews):
    """Select 2-3 representative preview PNGs: cover, densest/media slide, and closing/table slide."""
    if not previews:
        return []
    n = len(previews)
    if n <= 3:
        return list(previews)
    picks = [0]
    slides = (slides_spec or {}).get("slides", [])
    best_media, best_dense = None, None
    best_dense_score = -1
    for i, s in enumerate(slides):
        if i == 0 or i >= n:
            continue
        has_media = bool(s.get("table") or s.get("image") or s.get("images") or s.get("chart"))
        if has_media and best_media is None:
            best_media = i
        b_len = sum(len((b.get("text", "") if isinstance(b, dict) else str(b)))
                    for b in (s.get("bullets") or []))
        if s.get("table"):
            b_len += 250
        if b_len > best_dense_score:
            best_dense_score = b_len
            best_dense = i
    for idx in (best_media, best_dense, n - 1):
        if idx is not None and 0 <= idx < n and idx not in picks:
            picks.append(idx)
        if len(picks) >= 3:
            break
    while len(picks) < 3 and len(picks) < n:
        mid = n // 2
        if mid not in picks:
            picks.append(mid)
        else:
            break
    return [previews[i] for i in sorted(picks)]


def _write_initial_fallback_result(out_dir, cr, extract_ok):
    """Write an honest partial fallback /app/output/result.json right after prep
    (only if one does not already exist) so even an interrupted run leaves a valid deliverable."""
    os.makedirs(out_dir, exist_ok=True)
    rp = os.path.join(out_dir, "result.json")
    if os.path.isfile(rp) and os.path.getsize(rp) > 100:
        return rp
    sources = []
    for k in ("paper_source", "paper_pdf", "template_pptx", "starting_deck"):
        v = cr.get(k)
        if v and v not in sources:
            sources.append(v)
    gaps = ["Initial extraction completed; slide deck build and render pending."]
    if not extract_ok:
        gaps.append("Paper extraction reported issues; see case_resolved.json / prep output.")
    result = {
        "case_id": cr.get("case_id") or "unknown",
        "task": "paper_to_academic_slides",
        "status": "partial",
        "deliverables": {
            "pptx": None,
            "pdf": None,
            "previews": [],
            "evidence_map_json": None,
            "speaker_notes": None,
        },
        "deck": {"n_slides": 0, "notes_complete": False},
        "outline": [],
        "evidence_map": [],
        "sources": sources,
        "render_record": {"source_pptx": None, "exported_pdf": None, "preview_pngs": []},
        "checks": {
            "pptx_opens": False,
            "pdf_exported": False,
            "previews_rendered": 0,
            "geometry_fail": None,
            "evidence_errors": None,
        },
        "gaps_and_uncertainty": gaps,
        "uncertainty": gaps,
        "change_records": [{"step": "prep", "detail": "resolved case and wrote initial partial result.json"}],
    }
    oc = cr.get("output_contract")
    if isinstance(oc, dict):
        req_fields = oc.get("required_fields") or oc.get("fields") or list(oc.keys())
        if isinstance(req_fields, list):
            for rf in req_fields:
                if isinstance(rf, str) and rf not in result:
                    result[rf] = None if "path" in rf or "file" in rf else gaps
    with open(rp, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    return rp


def _build_reading_pack(work, cr, ctx, idx_path, extra_materials):
    """Assemble a self-contained <=45KB reading_pack.md bundling full task brief,
    starting deck dump, supplemental materials, figure/table-first paper_index.md,
    untruncated core pages, and an explicit remaining-pages directory."""
    paper_dir = os.path.join(work, "paper")
    os.makedirs(paper_dir, exist_ok=True)
    pack_path = os.path.join(paper_dir, "reading_pack.md")
    MAX_CHARS = 45000
    parts = ["# Unified Reading Pack (High-Completeness Briefing)\n"]
    used = len(parts[0])

    def _add(block, cap=None):
        nonlocal used
        if not block:
            return
        if cap and len(block) > cap:
            block = block[:cap].rstrip() + "\n... [truncated; read source file if needed]\n"
        rem = MAX_CHARS - used
        if rem <= 200:
            return
        if len(block) > rem:
            block = block[:rem - 60].rstrip() + "\n... [budget cap reached]\n"
        parts.append(block)
        used += len(block)

    # 1. Task & constraints (never chop real task.md instructions!)
    tmd_path = cr.get("task_md")
    tmd_raw = ""
    if tmd_path and os.path.isfile(tmd_path):
        try:
            tmd_raw = open(tmd_path, encoding="utf-8", errors="ignore").read().strip()
        except Exception:
            pass
    lim_pages = (ctx or {}).get("limitations_pages") or []
    ocr_pages = (ctx or {}).get("ocr_pages") or []
    unext_figs = (ctx or {}).get("unextracted_figures") or []
    img_form_pages = (ctx or {}).get("image_formula_pages") or []
    n_figs = len((ctx or {}).get("figures") or [])
    n_tbls = len((ctx or {}).get("tables") or [])
    cons = cr.get("constraints") or {}
    eff_mode = (cr.get("effective_template_mode")
                or (cr.get("template_info") or {}).get("recommended_mode")
                or ("layouts" if cr.get("template_pptx") else "blank"))
    hdr_lines = [
        "## 1. Case, Ground-Truth Integrity & Constraints",
        f"- case_id: `{cr.get('case_id')}` | template_mode: `{eff_mode}`",
        f"- constraints: `{json.dumps(cons, ensure_ascii=False)}`",
        f"- grounding_summary: pages={(ctx or {}).get('n_pages', 0)}, extracted_figures={n_figs}, extracted_tables={n_tbls}, limitations_pages={lim_pages}"
        + (f", ocr_pages={ocr_pages}" if ocr_pages else "")
        + (f", image_formula_pages={img_form_pages}" if img_form_pages else "")
        + (f", unextracted_figures={unext_figs[:6]}" if unext_figs else "")
        + (" [REQUIRED: include a Limitations slide]" if lim_pages else ""),
        "- ground_truth_rule: ALWAYS derive claims/numbers/authors strictly from actual paper/material files below; if any case label contradicts the paper, trust the paper and log discrepancy in gaps_and_uncertainty.",
    ]
    ms_c = cons.get("max_slides")
    dur_c = cons.get("duration_minutes") or cons.get("duration")
    if ms_c or dur_c:
        plan_bits = []
        if ms_c:
            plan_bits.append(f"<= {ms_c} slides total")
        if dur_c:
            plan_bits.append(f"~{dur_c} min total (allocate `[~X.X min]` per slide in `notes` summing to ~{dur_c} min)")
        hdr_lines.append(f"- deck_pacing_target: {' | '.join(plan_bits)}")
    if cr.get("label_discrepancy"):
        hdr_lines.append(f"- **CRITICAL LABEL MISMATCH WARNING**: {cr['label_discrepancy']}")
    if unext_figs:
        hdr_lines.append(f"- unextracted_figures_note: `{unext_figs[:6]}` have no cropped PNG — do NOT invent image paths; cite in text/evidence or plot their verbatim table/text numbers via native `table` or `chart`.")
    if cr.get("missing"):
        hdr_lines.append(f"- missing_inputs: `{json.dumps(cr.get('missing'), ensure_ascii=False)}`")
    if cr.get("output_contract"):
        hdr_lines.append(f"- output_contract: `{json.dumps(cr.get('output_contract'), ensure_ascii=False)[:900]}`")
    if cr.get("expected_outputs"):
        hdr_lines.append(f"- expected_outputs: `{json.dumps(cr.get('expected_outputs'), ensure_ascii=False)[:600]}`")
    if tmd_raw:
        hdr_lines.append("\n### Task Instructions (`task.md`)\n" + tmd_raw[:12000])
        if len(tmd_raw) > 12000:
            hdr_lines.append(f"... [task.md has {len(tmd_raw)} chars; read `{tmd_path}` for remainder]")
    _add("\n".join(hdr_lines) + "\n\n")

    # 2. Full Starting deck dump (if any)
    sd_md = cr.get("starting_deck_dump_md")
    if sd_md and os.path.isfile(sd_md):
        try:
            sd_txt = open(sd_md, encoding="utf-8", errors="ignore").read().strip()
            if sd_txt:
                tip = ("*(In `--reuse` mode, you may set `\"source_slide\": N` on any slide to select/reorder "
                       "starting-deck slide N, and `\"keep_existing_media\": true` to keep its native chart/diagram.)*\n\n")
                _add("## 2. Starting Deck Inspection\n" + tip + sd_txt + "\n\n", cap=9000)
        except Exception:
            pass

    # 3. Supplemental materials digest (images + csv/md/txt/json)
    if extra_materials:
        em_lines = ["## 3. Supplemental Input Materials"]
        text_budget = 4500
        for mp in extra_materials[:14]:
            if not os.path.isfile(mp):
                continue
            ext = os.path.splitext(mp)[1].lower()
            if ext in (".png", ".jpg", ".jpeg"):
                try:
                    from PIL import Image
                    w_i, h_i = Image.open(mp).size
                    ar = round(w_i / max(h_i, 1), 2)
                    tag = " [BANNER]" if ar >= 2.2 else ""
                    em_lines.append(f"- Image: `{mp}` ({w_i}x{h_i}, aspect={ar}{tag})")
                except Exception:
                    em_lines.append(f"- Image: `{mp}`")
            elif ext in (".csv", ".tsv", ".md", ".txt", ".json") and text_budget > 200:
                try:
                    raw = open(mp, encoding="utf-8", errors="ignore").read().strip()
                    take = min(1800, text_budget)
                    em_lines.append(f"- Text/Data `{mp}` ({len(raw)} chars):\n```\n{raw[:take]}\n```")
                    text_budget -= min(len(raw), take)
                except Exception:
                    em_lines.append(f"- Material: `{mp}`")
            else:
                em_lines.append(f"- File: `{mp}`")
        if len(em_lines) > 1:
            _add("\n".join(em_lines) + "\n\n", cap=5200)

    # 4. Paper index (Figures & Tables FIRST so they are never truncated!)
    if idx_path and os.path.isfile(idx_path):
        try:
            idx_txt = open(idx_path, encoding="utf-8", errors="ignore").read().strip()
            # Drop redundant Page 1 Excerpt since Section 5 includes full verbatim Page 1
            p1_marker = "\n## Page 1 Excerpt"
            if p1_marker in idx_txt:
                idx_txt = idx_txt.split(p1_marker, 1)[0].rstrip()
            _add("## 4. Paper Index (Figures, Tables, Sections, Captions, Quantitative Highlights)\n"
                 + idx_txt + "\n\n", cap=16000)
        except Exception:
            pass

    # 5. UNTRUNCATED verbatim text of core pages + Remaining Pages manifest
    included_pages = []
    pages_list = (ctx or {}).get("pages") or []
    pages_by_no = {p["page"]: p.get("text", "") for p in pages_list if isinstance(p, dict) and p.get("page")}
    secs_by_page = {}
    for s in ((ctx or {}).get("sections") or []):
        sp = s.get("page")
        if sp:
            secs_by_page.setdefault(sp, []).append(f"{s.get('num', '')} {s.get('title', '')}".strip())

    if pages_by_no:
        all_pnos = sorted(pages_by_no.keys())
        rec = [int(p) for p in ((ctx or {}).get("recommended_read_pages") or []) if int(p) in pages_by_no]
        # Priority order: p.1 -> recommended core pages -> remaining pages in reading order
        ordered = []
        for pno in [1] + rec + all_pnos:
            if pno in pages_by_no and pno not in ordered:
                ordered.append(pno)

        p_lines = ["## 5. Core Pages Verbatim Text (UNTRUNCATED — copy quotes/numbers verbatim)"]
        budget_for_pages = max(0, MAX_CHARS - used - 900)  # reserve 900 chars for Remaining Pages manifest
        p_used = len(p_lines[0])
        for pno in ordered:
            ptxt = (pages_by_no.get(pno) or "").strip()
            if not ptxt:
                continue
            block = f"\n### Page {pno}\n{ptxt}\n"
            if p_used + len(block) <= budget_for_pages:
                p_lines.append(block)
                p_used += len(block)
                included_pages.append(pno)
            elif not included_pages and budget_for_pages > 1200:
                # Guarantee at least Page 1 fits even on extreme mega-pages
                take = budget_for_pages - p_used - 80
                p_lines.append(f"\n### Page {pno}\n{ptxt[:take].rstrip()}\n... [see pages/p{pno:02d}.txt]\n")
                included_pages.append(pno)
                break
            elif pno in ([1] + rec) and (budget_for_pages - p_used) >= 2400:
                # Include substantial prefix of a recommended page rather than skipping it entirely
                take = budget_for_pages - p_used - 80
                p_lines.append(f"\n### Page {pno} (excerpt; full in pages/p{pno:02d}.txt)\n{ptxt[:take].rstrip()}\n")
                included_pages.append(pno)
                break

        omitted = [p for p in all_pnos if p not in included_pages]
        if omitted:
            p_lines.append("\n### Remaining Pages (read via `/app/work/paper/pages/pNN.txt` if needed)")
            for op in omitted[:20]:
                sec_lbl = "; ".join(secs_by_page.get(op, [])[:3]) or "(continuation / references / appendix)"
                p_lines.append(f"- `pages/p{op:02d}.txt` ({len(pages_by_no.get(op, ''))} chars): {sec_lbl}")
        _add("\n".join(p_lines) + "\n")

    with open(pack_path, "w", encoding="utf-8") as f:
        f.write("".join(parts))
    return pack_path, sorted(included_pages)


def cmd_prep(a):
    work = os.path.abspath(a.work)
    out_dir = os.path.abspath(a.out_dir)
    os.makedirs(work, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)
    att_file = os.path.join(work, ".build_attempts.json")
    if os.path.isfile(att_file):
        try:
            os.remove(att_file)
        except Exception:
            pass
    cr_path = os.path.join(work, "case_resolved.json")
    paper_dir = os.path.join(work, "paper")

    rc0, out0, err0 = _run_py("resolve_case.py", [
        "--case-json", a.case_json,
        "--input-dir", a.input_dir,
        "--out", cr_path,
    ])
    cr = _load_json(cr_path) or {}
    tmd_raw = cr.get("task_md_excerpt") or ""
    tmd_trunc = bool(cr.get("task_md_truncated")) or (len(tmd_raw) > 4500)
    tinfo = cr.get("template_info") or {}
    summary = {
        "stage": "prep",
        "ok": rc0 == 0 and bool(cr),
        "case_resolved": cr_path,
        "case_id": cr.get("case_id"),
        "task_md": cr.get("task_md"),
        "task_md_excerpt": tmd_raw[:4500],
        "task_md_truncated": tmd_trunc,
        "paper_source": cr.get("paper_source") or cr.get("paper_pdf"),
        "template_pptx": cr.get("template_pptx"),
        "starting_deck": cr.get("starting_deck"),
        "starting_deck_dump_md": cr.get("starting_deck_dump_md"),
        "n_existing_slides": max(
            tinfo.get("n_existing_slides", 0),
            (cr.get("starting_deck_info") or {}).get("n_existing_slides", 0),
        ),
        "template_mode": cr.get("effective_template_mode")
                         or tinfo.get("recommended_mode")
                         or ("layouts" if cr.get("template_pptx") else "blank"),
        "constraints": cr.get("constraints") or {},
        "output_contract": cr.get("output_contract"),
        "expected_outputs": cr.get("expected_outputs") or {},
        "extra_materials": [
            m for m in (cr.get("materials") or [])
            if m != (cr.get("paper_source") or cr.get("paper_pdf"))
        ],
        "missing": cr.get("missing") or [],
    }
    if rc0 != 0:
        summary["resolve_error"] = err0 or out0

    paper_src = summary["paper_source"]
    if paper_src and os.path.isfile(paper_src):
        rc1, out1, err1 = _run_py("extract_paper.py", [
            paper_src, "--out", paper_dir, "--dpi", str(a.extract_dpi),
        ])
        ctx_path = os.path.join(paper_dir, "paper_context.json")
        idx_path = os.path.join(paper_dir, "paper_index.md")
        ctx = _load_json(ctx_path) or {}
        summary["extract_ok"] = rc1 == 0 and os.path.isfile(ctx_path)
        summary["paper_index_md"] = idx_path if os.path.isfile(idx_path) else None
        summary["pages_dir"] = os.path.join(paper_dir, "pages")
        summary["paper_context_json"] = ctx_path if os.path.isfile(ctx_path) else None
        summary["n_pages"] = ctx.get("n_pages", 0)
        summary["n_captions"] = len(ctx.get("captions") or [])
        summary["n_figures"] = len(ctx.get("figures") or [])
        summary["n_tables"] = len(ctx.get("tables") or [])
        summary["key_highlights_count"] = len(ctx.get("key_highlights") or [])
        summary["limitations_pages"] = ctx.get("limitations_pages") or []
        summary["recommended_read_pages"] = ctx.get("recommended_read_pages") or []
        summary["title_guess"] = ctx.get("title_guess") or ""
        if ctx.get("limits"):
            summary["extract_limits"] = ctx["limits"]
        if ctx.get("unextracted_figures"):
            summary["unextracted_figures"] = ctx["unextracted_figures"]
        if ctx.get("image_formula_pages"):
            summary["image_formula_pages"] = ctx["image_formula_pages"]
        lbl_disc = _detect_label_discrepancy(cr, ctx)
        if lbl_disc:
            cr["label_discrepancy"] = lbl_disc
            summary["label_discrepancy"] = lbl_disc
            try:
                with open(cr_path, "w", encoding="utf-8") as cf:
                    json.dump(cr, cf, ensure_ascii=False, indent=1)
            except Exception:
                pass
        if rc1 != 0:
            summary["extract_error"] = err1 or out1
            summary["ok"] = False
        # Build one-shot reading_pack.md bundling task, starting deck, extra materials, index, and core pages
        try:
            pack_path, inc_pages = _build_reading_pack(
                work, cr, ctx, summary.get("paper_index_md"), summary.get("extra_materials") or []
            )
            summary["reading_pack_md"] = pack_path
            summary["reading_pack_included_pages"] = inc_pages
        except Exception as e:
            summary["reading_pack_error"] = str(e)
    else:
        summary["extract_ok"] = False
        summary["ok"] = False
        summary["extract_error"] = "No paper PDF or text material found."

    # Immediately save an honest partial fallback result.json so output is never missing
    summary["fallback_result_json"] = _write_initial_fallback_result(
        out_dir, cr, bool(summary.get("extract_ok"))
    )
    if summary.get("reading_pack_md"):
        extra_note = ""
        if tmd_trunc and summary.get("task_md"):
            extra_note = " (also read FULL task_md because excerpt was truncated)"
        summary["next_step"] = (
            f"FAST PATH: Read `{summary['reading_pack_md']}` ONCE{extra_note} "
            f"(already includes task, starting deck, extra materials, paper_index, and verbatim pages "
            f"{summary.get('reading_pack_included_pages', [])}); only read other pages/pNN.txt if a needed detail is absent; "
            "then write /app/work/slides.json and run: python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py build /app/work/slides.json --work /app/work --out-dir /app/output"
        )
    else:
        next_reads = ["paper_index_md"]
        if tmd_trunc and summary.get("task_md"):
            next_reads.append("FULL task_md (excerpt was truncated!)")
        if summary.get("starting_deck_dump_md"):
            next_reads.append("starting_deck_dump_md (inspect existing slides before adapting)")
        rec_p = summary.get("recommended_read_pages") or []
        if rec_p:
            next_reads.append(f"recommended pages {[f'p{int(p):02d}.txt' for p in rec_p]}")
        else:
            next_reads.append("key pages/pNN.txt")
        summary["next_step"] = (
            f"Read {', '.join(next_reads)}; "
            "then write /app/work/slides.json and run: python3 ${SKILL_DIR:-/skills/weekly-skill}/scripts/run_pipeline.py build /app/work/slides.json --work /app/work --out-dir /app/output"
        )
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    return 0 if summary["ok"] else 1


def cmd_build(a):
    work = os.path.abspath(a.work)
    out_dir = os.path.abspath(a.out_dir)
    os.makedirs(work, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    slides_json = os.path.abspath(a.slides_json)
    slides_spec = _load_json(slides_json)
    if not slides_spec or not isinstance(slides_spec.get("slides"), list):
        print(json.dumps({
            "stage": "build", "ok": False,
            "error": f"Invalid or missing slides.json: {slides_json}",
        }, ensure_ascii=False))
        return 1

    cr_path = a.case_resolved or os.path.join(work, "case_resolved.json")
    cr = _load_json(cr_path) or {}
    meta = slides_spec.get("meta") or {}

    # Resolve template and reuse mode cleanly (honoring dual-PPTX effective_template_mode)
    tinfo = cr.get("template_info") or {}
    rec_mode = cr.get("effective_template_mode") or tinfo.get("recommended_mode")
    if a.reuse:
        use_reuse = True
    elif a.layouts:
        use_reuse = False
    elif "reuse" in meta:
        use_reuse = bool(meta["reuse"])
    else:
        use_reuse = (rec_mode == "reuse")

    template_path = a.template
    if not template_path:
        if use_reuse and cr.get("starting_deck") and os.path.isfile(cr["starting_deck"]):
            template_path = cr["starting_deck"]
        elif cr.get("template_pptx") and os.path.isfile(cr["template_pptx"]):
            template_path = cr["template_pptx"]

    deck_path = os.path.join(work, "deck.pptx")
    build_args = [slides_json, "--out", deck_path]
    if template_path:
        build_args += ["--template", template_path]
    if use_reuse and template_path:
        build_args.append("--reuse")

    rc_b, out_b, err_b = _run_py("build_deck.py", build_args)
    if rc_b != 0 or not os.path.isfile(deck_path):
        print(json.dumps({
            "stage": "build", "ok": False, "step": "build_deck",
            "error": err_b or out_b,
        }, ensure_ascii=False))
        return 1

    # S4: Geometry & layout check
    lay_path = os.path.join(work, "layout_report.json")
    rc_l, _out_l, err_l = _run_py("check_layout.py", [deck_path, "--report", lay_path])
    lay = _load_json(lay_path) or {}
    if rc_l != 0 and not lay:
        lay = {"n_fail": 1, "n_warn": 0, "issues": [{"sev": "fail", "type": "check_layout_error", "detail": err_l[:200]}]}

    # S4.5: Mandatory evidence cross-check
    ctx_path = a.ctx or os.path.join(work, "paper", "paper_context.json")
    ev_path = os.path.join(work, "evidence_report.json")
    ev = None
    if os.path.isfile(ctx_path):
        ve_args = [slides_json, "--ctx", ctx_path, "--report", ev_path]
        if os.path.isfile(cr_path):
            ve_args += ["--case-resolved", cr_path]
        rc_e, _out_e, err_e = _run_py("verify_evidence.py", ve_args)
        ev = _load_json(ev_path) or {}
        if rc_e != 0 and not ev:
            ev = {"n_error": 1, "n_warn": 0, "issues": [{"sev": "error", "type": "verify_evidence_error", "detail": err_e[:200]}]}

    # S5: Export PDF/PNG previews + result.json
    fin_args = [
        deck_path,
        "--slides-json", slides_json,
        "--out-dir", out_dir,
        "--work", work,
        "--dpi", str(a.dpi),
    ]
    if os.path.isfile(cr_path):
        fin_args += ["--case-resolved", cr_path]
    if a.case_id:
        fin_args += ["--case-id", a.case_id]
    if template_path:
        fin_args += ["--template", template_path]
    rc_f, out_f, err_f = _run_py("finalize_output.py", fin_args)
    res = _load_json(os.path.join(out_dir, "result.json")) or {}

    lay_fail = int(lay.get("n_fail", 0))
    lay_warn = int(lay.get("n_warn", 0))
    ev_err = int(ev.get("n_error", 0)) if isinstance(ev, dict) else 0
    ev_warn = int(ev.get("n_warn", 0)) if isinstance(ev, dict) else 0

    fail_items = [i for i in (lay.get("issues") or []) if i.get("sev") == "fail"][:15]
    warn_items = [i for i in (lay.get("issues") or []) if i.get("sev") == "warn"][:10]
    ev_err_items = [i for i in ((ev or {}).get("issues") or []) if i.get("sev") == "error"][:15]
    ev_warn_items = [i for i in ((ev or {}).get("issues") or []) if i.get("sev") == "warn"][:10]

    previews = (res.get("deliverables") or {}).get("previews") or []
    checks = res.get("checks") or {}
    rrec = res.get("render_record") or {}
    blank_slides = checks.get("visual_blank_slides") or []
    tofu_slides = checks.get("visual_tofu_slides") or []

    # Track build attempts per work dir so the 1-pass rubric polish gate never infinite-loops
    att_file = os.path.join(work, ".build_attempts.json")
    att_state = _load_json(att_file) or {}
    attempt = int(att_state.get("count", 0)) + 1
    try:
        with open(att_file, "w", encoding="utf-8") as af:
            json.dump({"count": attempt}, af)
    except Exception:
        pass

    hard_ev_warns = [i for i in ((ev or {}).get("issues") or [])
                     if i.get("sev") == "warn" and i.get("type") in HARD_EV_TYPES]
    hard_lay_warns = [i for i in (lay.get("issues") or [])
                      if i.get("sev") == "warn" and i.get("type") in HARD_LAY_TYPES]
    polish_ev = [i for i in ((ev or {}).get("issues") or [])
                 if i.get("sev") == "warn" and i.get("type") in POLISH_EV_TYPES]
    polish_lay = [i for i in (lay.get("issues") or [])
                  if i.get("sev") == "warn" and i.get("type") in POLISH_LAY_TYPES]

    render_broken = (checks.get("pdf_pages_match") is False) or (checks.get("render_text_ok") is False)
    hard_action = (
        (lay_fail > 0) or (ev_err > 0) or bool(blank_slides) or bool(tofu_slides)
        or render_broken or bool(hard_ev_warns) or bool(hard_lay_warns)
    )
    polish_items = polish_ev + polish_lay
    polish_gate = (attempt == 1) and bool(polish_items)
    # Hard cap at attempt <= 2 so the agent never wastes budget looping on edge warnings
    action_required = (hard_action or polish_gate) if attempt <= 2 else False

    summary = {
        "stage": "build",
        "attempt": attempt,
        "ok": rc_f == 0 and not action_required,
        "status": res.get("status", "unknown"),
        "result_json": os.path.join(out_dir, "result.json"),
        "n_slides": (res.get("deck") or {}).get("n_slides"),
        "template_mode": (res.get("deck") or {}).get("template_mode"),
        "pdf_exported": checks.get("pdf_exported"),
        "pdf_pages_match": checks.get("pdf_pages_match"),
        "render_text_ok": checks.get("render_text_ok"),
        "visual_pixel_ok": checks.get("visual_pixel_ok"),
        "visual_blank_slides": blank_slides,
        "visual_tofu_slides": tofu_slides,
        "font_substitution_detected": rrec.get("font_substitution_detected"),
        "layout_fail": lay_fail,
        "layout_warn": lay_warn,
        "evidence_errors": ev_err,
        "evidence_warns": ev_warn,
        "action_required": action_required,
        "polish_recommended": bool(polish_items),
        "inspect_previews": _pick_preview_samples(slides_spec, previews),
    }
    if fail_items:
        summary["layout_fail_details"] = fail_items
    if warn_items:
        summary["layout_warn_details"] = warn_items
    if ev_err_items:
        summary["evidence_error_details"] = ev_err_items
    if ev_warn_items:
        summary["evidence_warn_details"] = ev_warn_items
    if action_required:
        checklist = []
        for bs in blank_slides:
            checklist.append(f"S{bs} [visual_blank_slide]: rendered preview appears nearly blank; verify content/background contrast")
        for ts in tofu_slides:
            checklist.append(f"S{ts} [visual_tofu_glyph]: missing glyph/tofu detected in rendered PDF text; check font/special symbols")
        if checks.get("pdf_pages_match") is False:
            checklist.append("S0 [pdf_pages_match]: exported PDF page count differs from PPTX slide count")
        if checks.get("render_text_ok") is False:
            checklist.append("S0 [render_text_ok]: rendered PDF text missing slide titles/content")
        for it in (fail_items + ev_err_items + hard_ev_warns + hard_lay_warns + polish_items)[:12]:
            checklist.append(f"S{it.get('slide', 0)} [{it.get('type')}]: {it.get('detail', '')}")
        checklist = checklist[:12]
        summary["polish_checklist"] = checklist
        summary["fix_instruction"] = (
            f"SURGICAL FIX PASS (attempt {attempt}/2): resolve the {len(checklist)} high-impact item(s) in "
            f"polish_checklist inside {slides_json} in ONE edit, then re-run: "
            f"python3 ${{SKILL_DIR:-/skills/weekly-skill}}/scripts/run_pipeline.py build {slides_json} --work {work} --out-dir {out_dir}"
        )
    else:
        summary["next_step"] = (
            "All automated geometry, evidence, PDF text, and pixel visual checks passed. "
            "If running on a multimodal vision model, you may optionally inspect inspect_previews; "
            "text-only models MUST NOT read binary PNG files and can finalize immediately."
        )
    if rc_f != 0:
        summary["finalize_error"] = err_f or out_f

    print(json.dumps(summary, ensure_ascii=False, indent=1))
    return 0


def main():
    ap = argparse.ArgumentParser(description="One-command runner for paper-to-academic-slides.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_prep = sub.add_parser("prep", help="S0 resolve_case + S1 extract_paper + fallback result.json")
    p_prep.add_argument("--case-json", default="/input/case.json")
    p_prep.add_argument("--input-dir", default="/input")
    p_prep.add_argument("--work", default="/app/work")
    p_prep.add_argument("--out-dir", default="/app/output")
    p_prep.add_argument("--extract-dpi", type=int, default=180)

    p_build = sub.add_parser("build", help="S3 build + S4 layout + S4.5 evidence + S5 render/export")
    p_build.add_argument("slides_json", help="Path to slides.json")
    p_build.add_argument("--work", default="/app/work")
    p_build.add_argument("--out-dir", default="/app/output")
    p_build.add_argument("--case-resolved", default=None)
    p_build.add_argument("--case-id", default=None)
    p_build.add_argument("--template", default=None)
    p_build.add_argument("--ctx", default=None)
    p_build.add_argument("--dpi", type=int, default=100)
    mode_g = p_build.add_mutually_exclusive_group()
    mode_g.add_argument("--reuse", action="store_true", help="Force in-place reuse of starting deck")
    mode_g.add_argument("--layouts", action="store_true", help="Force clean rebuild from template layouts")

    a = ap.parse_args()
    if a.cmd == "prep":
        sys.exit(cmd_prep(a))
    elif a.cmd == "build":
        sys.exit(cmd_build(a))


if __name__ == "__main__":
    main()
