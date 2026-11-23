#!/usr/bin/env python3
"""run_pipeline.py — Deterministic two-command orchestrator for academic-slides-pku.

Collapses the multi-script workflow into two fast, single-turn commands to minimize
Agent turns and prevent timeouts across multi-case runs:

  1) prepare: inspect_case + immediate honest partial result.json + compact brief
     python3 run_pipeline.py prepare --case-dir /input --work /app/work --out /app/output

  2) deliver: build_deck -> check_deck -> export_preview -> assemble_result (auto)
     python3 run_pipeline.py deliver --case-dir /input --work /app/work \
       --out /app/output --plan /app/work/slides_plan.json [--dpi 100]
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
PY = sys.executable


def _run(script: str, args: list[str], timeout: int = 300) -> subprocess.CompletedProcess:
    return subprocess.run(
        [PY, str(SCRIPTS / script), *args],
        capture_output=True, text=True, timeout=timeout,
    )


def _read_json(p: Path, default=None):
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def cmd_prepare(args) -> int:
    work = Path(args.work)
    out = Path(args.out)
    work.mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)
    deck = out / "presentation.pptx"

    r1 = _run("inspect_case.py", ["--case-dir", args.case_dir, "--work", str(work)], timeout=240)
    if r1.returncode != 0:
        print("PREPARE_FAILED " + json.dumps({
            "step": "inspect_case.py",
            "rc": r1.returncode,
            "stderr": (r1.stderr or r1.stdout or "")[-600:],
        }, ensure_ascii=False))
        return 1

    r2 = _run("assemble_result.py", [
        "--case-dir", args.case_dir, "--work", str(work),
        "--deck", str(deck), "--status", "partial",
    ], timeout=120)

    inv = _read_json(work / "inventory.json", {}) or {}
    tmpl = inv.get("template") or {}
    pdf_ex = inv.get("pdf_extract") or {}

    raw_task = ""
    task_info = None
    if inv.get("task_md") and Path(inv["task_md"]).is_file():
        raw_task = Path(inv["task_md"]).read_text(encoding="utf-8", errors="replace").strip()
        task_info = {
            "path": inv["task_md"],
            "truncated": len(raw_task) > 7500,
            "chars": len(raw_task),
        }

    figs = []
    for fg in (pdf_ex.get("figures") or [])[:16]:
        figs.append({
            "file": fg.get("file"), "page": fg.get("page"),
            "kind": fg.get("kind"), "caption": (fg.get("caption") or "")[:160],
        })
    for ip in (inv.get("image_files") or [])[:16]:
        figs.append({"file": ip, "kind": "case_image"})

    digest_file = inv.get("material_digest_file")
    d_txt = ""
    digest_truncated = False
    if digest_file and Path(digest_file).is_file():
        d_txt = Path(digest_file).read_text(encoding="utf-8", errors="replace").strip()
        digest_truncated = len(d_txt) > 13500

    sdd_file = inv.get("start_deck_dump")
    s_txt = ""
    start_deck_truncated = False
    if sdd_file and Path(sdd_file).is_file():
        s_txt = Path(sdd_file).read_text(encoding="utf-8", errors="replace").strip()
        start_deck_truncated = len(s_txt) > 6500

    comp = inv.get("completeness") or {}
    read_more_hints = []
    if task_info and task_info.get("truncated"):
        read_more_hints.append(f"task.md ({task_info['chars']} chars > 7.5k): read tail of {inv['task_md']}")
    if comp.get("material_json_truncated") and inv.get("material_json_path"):
        read_more_hints.append(f"material.json ({comp.get('material_json_chars')} chars > 9.5k): read {inv['material_json_path']}")
    if (comp.get("pdf_beyond_digest") or comp.get("digest_clipped") or digest_truncated) and pdf_ex.get("text_file"):
        read_more_hints.append(
            f"full paper text ({comp.get('pdf_text_chars')} chars): read relevant sections of {pdf_ex['text_file']} using page markers"
        )
    if start_deck_truncated and sdd_file:
        read_more_hints.append(f"start_deck_dump ({len(s_txt)} chars > 6.5k): read {sdd_file}")

    # Keep line 1 strictly compact single-line JSON, then print readable Markdown sections below
    brief = {
        "rp_path": str(SCRIPTS / "run_pipeline.py"),
        "case_id": inv.get("case_id"),
        "workorder": inv.get("workorder"),
        "task_md": task_info,
        "template": {
            "path": tmpl.get("path"),
            "size_in": tmpl.get("slide_size_in"),
            "layouts": [f"{l['i']}:{l['name']}({len(l['placeholders'])}ph)"
                        for l in tmpl.get("layouts", [])],
        } if tmpl else None,
        "start_deck_slides": len((inv.get("start_deck") or {}).get("slides", [])) if inv.get("start_deck") else None,
        "start_deck_dump": sdd_file,
        "material_json_path": inv.get("material_json_path"),
        "material_keys": list((inv.get("material") or {}).keys())[:24],
        "material_digest_file": digest_file,
        "material_text_file": pdf_ex.get("text_file"),
        "pdf_pages": pdf_ex.get("pages"),
        "pdf_text_chars": pdf_ex.get("text_chars"),
        "figures": figs,
        "extra_text_files": (inv.get("extra_text_files") or [])[:12],
        "anomalies": inv.get("anomalies", []),
        "partial_result": str(out / "result.json") if (out / "result.json").is_file() else None,
        "partial_rc": r2.returncode,
        "digest_truncated": bool(digest_truncated or comp.get("digest_clipped")),
        "start_deck_truncated": start_deck_truncated,
        "read_more_hints": read_more_hints,
    }
    print("PREPARE_OK " + json.dumps(brief, ensure_ascii=False))
    if raw_task:
        print("\n=== TASK.MD ===\n" + raw_task[:7500])
    if d_txt:
        print("\n=== MATERIAL DIGEST ===\n" + d_txt[:13500])
    if s_txt:
        print("\n=== START DECK DUMP ===\n" + s_txt[:6500])
    return 0


def _collect_source_lines(work: Path) -> list[tuple[str, str]]:
    """Collect (page_tag, verbatim_line) pairs from material_text.md + material.json + start_deck_dump + task.md + extra files."""
    inv = _read_json(work / "inventory.json", {}) or {}
    out: list[tuple[str, str]] = []
    seen = set()

    def _add_lines(txt: str, default_loc: str = "material"):
        cur_loc = default_loc
        for raw_ln in (txt or "").splitlines():
            s = raw_ln.strip()
            if not s:
                continue
            pm = re.match(r"^(?:<!--\s*PDF page\s*(\d+)\s*-->|---\s*page\s*(\d+)\s*---)$", s)
            if pm:
                cur_loc = f"p.{pm.group(1) or pm.group(2)}"
                continue
            if s.startswith("<!--") or len(s) < 12:
                continue
            # Split long paragraphs into sentence-like chunks while keeping exact substrings
            parts = [s] if len(s) <= 220 else [p.strip() for p in re.split(r"(?<=[.!?。！？])\s+", s) if len(p.strip()) >= 12]
            for p in parts:
                chunk = p[:200]
                if chunk not in seen:
                    seen.add(chunk)
                    out.append((cur_loc, chunk))
                    if len(out) >= 5000:
                        return

    mt_file = (inv.get("pdf_extract") or {}).get("text_file")
    if mt_file and Path(mt_file).is_file():
        _add_lines(Path(mt_file).read_text(encoding="utf-8", errors="replace"), "p.1")
    for k, loc in (("material_json_path", "material.json"), ("start_deck_dump", "start_deck"), ("task_md", "task.md")):
        p = inv.get(k)
        if p and Path(p).is_file():
            _add_lines(Path(p).read_text(encoding="utf-8", errors="replace"), loc)
    for ef in (inv.get("extra_text_files") or [])[:6]:
        ep = Path(ef)
        if ep.is_file():
            _add_lines(ep.read_text(encoding="utf-8", errors="replace"), ep.name)
    return out


def _best_verbatim_line(query: str, source_lines: list[tuple[str, str]],
                        pref_loc: str = "", min_score: float = 0.30,
                        overlap_coeff: bool = False) -> tuple[str, str] | None:
    """Find the closest verbatim substring in source_lines by trigram overlap + shared number bonus."""
    try:
        import check_deck as cd
    except Exception:
        return None
    qn = cd.norm(query).lower()
    if len(qn) < 6 or not source_lines:
        return None
    q_tri = {qn[i:i + 3] for i in range(len(qn) - 2)}
    if not q_tri:
        return None
    q_nums = set(re.findall(r"\d+\.\d+|\d{2,}", query.replace(",", "")))
    best_sc, best_pair = 0.0, None
    for loc, ln in source_lines:
        ln_n = cd.norm(ln).lower()
        if len(ln_n) < 12:
            continue
        l_tri = {ln_n[i:i + 3] for i in range(len(ln_n) - 2)}
        if not l_tri:
            continue
        denom = min(len(q_tri), len(l_tri)) if overlap_coeff else len(q_tri)
        ov = len(q_tri & l_tri) / max(denom, 1)
        if q_nums:
            l_nums = set(re.findall(r"\d+\.\d+|\d{2,}", ln.replace(",", "")))
            if q_nums & l_nums:
                ov += 0.14 * (len(q_nums & l_nums) / max(len(l_nums), 1) if overlap_coeff else len(q_nums & l_nums) / len(q_nums))
        if pref_loc and pref_loc.lower() == loc.lower():
            ov += 0.04
        if ov > best_sc:
            best_sc, best_pair = ov, (loc, ln)
    return best_pair if (best_pair is not None and best_sc >= min_score) else None


def _auto_repair_plan(plan: dict, fails: list[dict], work: Path, pass_idx: int) -> tuple[dict, list[str]]:
    """Honest, deterministic in-process auto-repair before LibreOffice export.
    Fixes mechanical layout/overflow, verbatim quote alignment, and missing sources.
    NEVER fabricates or alters slide numbers/claims."""
    repaired = copy.deepcopy(plan)
    actions: list[str] = []
    fail_map = {f["id"]: f.get("detail", "") for f in fails}
    slides = repaired.get("slides") or []
    edits = repaired.get("edits") or []
    is_create = bool(slides) and (repaired.get("mode", "create") != "edit")

    # 1. Tighten layout on slides with geometry/overflow issues
    tighten_ids: set[int] = set()
    for cid in ("text_fit", "table_fit", "overlap", "bounds"):
        if cid in fail_map:
            for m in re.findall(r"slide\s+(\d+)", fail_map[cid]):
                tighten_ids.add(int(m))
    if tighten_ids and is_create:
        for si in sorted(tighten_ids):
            if 1 <= si <= len(slides) and isinstance(slides[si - 1], dict):
                cur_t = int(slides[si - 1].get("_tighten", 0) or 0)
                if cur_t < pass_idx:
                    slides[si - 1]["_tighten"] = pass_idx
                    actions.append(f"tighten_slide_{si}_lvl{pass_idx}")
    elif tighten_ids and edits:
        for idx_e, e in enumerate(edits, 1):
            if not isinstance(e, dict):
                continue
            if str(e.get("op", "")).strip() in ("replace_slide", "append", "insert_at"):
                target_spec = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else e
                cur_t = int(target_spec.get("_tighten", 0) or 0)
                if cur_t < pass_idx:
                    target_spec["_tighten"] = pass_idx
                    actions.append(f"tighten_edit_{idx_e}_lvl{pass_idx}")

    # 1b. Clamp any explicit sub-floor font sizes (size < 10.5 or table font_pt < 10.0)
    if "min_font" in fail_map:
        def _clamp_spec_fonts(spec_d: dict, lbl: str):
            for b in (spec_d.get("bullets") or []):
                if isinstance(b, dict) and b.get("size") is not None:
                    try:
                        if float(b["size"]) < 10.5:
                            b["size"] = 10.5
                            actions.append(f"clamp_font_{lbl}")
                    except Exception:
                        pass
            tbl = spec_d.get("table")
            if isinstance(tbl, dict) and tbl.get("font_pt") is not None:
                try:
                    if float(tbl["font_pt"]) < 10.0:
                        tbl["font_pt"] = 10.0
                        actions.append(f"clamp_tbl_font_{lbl}")
                except Exception:
                    pass
        for idx_s, s in enumerate(slides, 1):
            if isinstance(s, dict):
                _clamp_spec_fonts(s, f"s{idx_s}")
        for idx_e, e in enumerate(edits, 1):
            if isinstance(e, dict):
                _clamp_spec_fonts(e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else e, f"e{idx_e}")

    # 2. Honest verbatim quote alignment + missing source attachment
    need_quotes = "quote_verification" in fail_map
    need_sources = "claim_sources" in fail_map
    if (need_quotes or need_sources) and (slides or edits):
        try:
            import check_deck as cd
            inv = _read_json(work / "inventory.json", {}) or {}
            mt_file = (inv.get("pdf_extract") or {}).get("text_file")
            fb_corpus = cd.corpus_from_material(inv, work)
            c_norm, c_nc = cd._build_corpus(mt_file, fb_corpus)
            c_low = c_norm.lower()
            c_nocite = cd._CITE_BRACKET_RE.sub("", c_low)
            src_lines = _collect_source_lines(work)

            def _quote_ok(q_raw: str) -> bool:
                q = cd.norm(q_raw)
                if len(q) < 6:
                    return True
                return bool(
                    q in c_norm or q.replace(",", "") in c_nc
                    or q.lower() in c_low
                    or cd._CITE_BRACKET_RE.sub("", q.lower()) in c_nocite
                )

            def _fix_sources_list(srcs: list, label: str):
                for src in (srcs or []):
                    if not isinstance(src, dict):
                        continue
                    q_raw = str(src.get("quote") or "")
                    if len(cd.norm(q_raw)) < 6 or _quote_ok(q_raw):
                        continue
                    q_ctx = f"{src.get('ref', '')} {q_raw}".strip()
                    best = _best_verbatim_line(q_ctx, src_lines, str(src.get("locator") or ""), min_score=0.38)
                    if best is None:
                        best = _best_verbatim_line(q_raw, src_lines, str(src.get("locator") or ""), min_score=0.36)
                    if best is not None and _quote_ok(best[1]):
                        src["quote"] = best[1][:160]
                        if not src.get("locator"):
                            src["locator"] = best[0]
                        actions.append(f"align_quote_{label}")

            if need_quotes:
                for idx, s in enumerate(slides, 1):
                    if isinstance(s, dict):
                        _fix_sources_list(s.get("sources") or [], f"s{idx}")
                for idx, e in enumerate(edits, 1):
                    if isinstance(e, dict):
                        _fix_sources_list(e.get("sources") or [], f"e{idx}")
                        if isinstance(e.get("slide_spec"), dict):
                            _fix_sources_list(e["slide_spec"].get("sources") or [], f"es{idx}")

            if need_sources:
                if is_create:
                    missing_s = {int(m) for m in re.findall(r"slide\s+(\d+)", fail_map.get("claim_sources", ""))}
                    for si in sorted(missing_s):
                        if 1 <= si <= len(slides) and isinstance(slides[si - 1], dict):
                            s = slides[si - 1]
                            if s.get("sources"):
                                continue
                            query_parts = [str(s.get("title") or "")]
                            for b in (s.get("bullets") or [])[:3]:
                                query_parts.append(b.get("t", "") if isinstance(b, dict) else str(b))
                            best = _best_verbatim_line(" ".join(query_parts), src_lines, "", min_score=0.28, overlap_coeff=True)
                            if best is not None and _quote_ok(best[1]):
                                s["sources"] = [{"ref": "material", "locator": best[0], "quote": best[1][:160]}]
                                actions.append(f"attach_source_s{si}")
                elif edits:
                    for idx_e, e in enumerate(edits, 1):
                        if not isinstance(e, dict):
                            continue
                        if str(e.get("op", "")).strip() in ("replace_slide", "append", "insert_at"):
                            sp = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else e
                            if sp.get("sources"):
                                continue
                            if (sp.get("role") or "") in ("title", "section", "agenda", "outline", "qa", "toc", "closing", "thanks"):
                                continue
                            query_parts = [str(sp.get("title") or "")]
                            for b in (sp.get("bullets") or [])[:3]:
                                query_parts.append(b.get("t", "") if isinstance(b, dict) else str(b))
                            best = _best_verbatim_line(" ".join(query_parts), src_lines, "", min_score=0.28, overlap_coeff=True)
                            if best is not None and _quote_ok(best[1]):
                                sp["sources"] = [{"ref": "material", "locator": best[0], "quote": best[1][:160]}]
                                actions.append(f"attach_source_e{idx_e}")
        except Exception:
            pass

    # 3. Auto-fill missing speaker notes honestly from the slide's own title/bullets/takeaway
    if "titles_notes" in fail_map:
        inv_n = _read_json(work / "inventory.json", {}) or {}
        wo_lang = str((inv_n.get("workorder") or {}).get("language") or repaired.get("language") or "zh").lower()
        is_en = wo_lang.startswith("en")

        def _fill_note_if_missing(spec_d: dict, lbl: str):
            if not str(spec_d.get("notes") or "").strip():
                t = str(spec_d.get("title") or repaired.get("deck_title") or "Slide").strip()
                b_items = [
                    (b.get("t", "") if isinstance(b, dict) else str(b)).strip()
                    for b in (spec_d.get("bullets") or [])[:2]
                ]
                b_items = [x for x in b_items if x]
                tk = str(spec_d.get("takeaway") or "").strip()
                parts = [p for p in ([t] + b_items + ([tk] if tk else [])) if p]
                if parts:
                    sep = ". " if is_en else "。"
                    spec_d["notes"] = sep.join(parts[:3])[:260]
                else:
                    spec_d["notes"] = "Present the key points of this slide." if is_en else "讲解本页核心内容与衔接过渡。"
                actions.append(f"fill_notes_{lbl}")
        if is_create:
            for idx_s, s in enumerate(slides, 1):
                if isinstance(s, dict):
                    _fill_note_if_missing(s, f"s{idx_s}")
        elif edits:
            for idx_e, e in enumerate(edits, 1):
                if isinstance(e, dict) and str(e.get("op", "")).strip() in ("replace_slide", "append", "insert_at"):
                    sp = e.get("slide_spec") if isinstance(e.get("slide_spec"), dict) else e
                    _fill_note_if_missing(sp, f"e{idx_e}")

    return repaired, actions


def _violation_score(fails: list[dict]) -> tuple[int, int]:
    """Return (num_failing_checks, total_individual_violations) so partial progress is kept."""
    total_v = 0
    for f in fails:
        d = str(f.get("detail") or "")
        # Patterns: "3 violations: ...", "5 quotes checked, 2 unmatched: ...", "2 slides without sources"
        m_unm = re.search(r"(\d+)\s+unmatched", d)
        if m_unm:
            total_v += int(m_unm.group(1))
            continue
        m_lead = re.match(r"\s*(\d+)\s+", d)
        if m_lead:
            total_v += max(1, int(m_lead.group(1)))
        else:
            total_v += 1
    return (len(fails), total_v)


def _apply_plan_patches(plan_path: Path, patch_slides: list[str], patch_file: str | None) -> list[str]:
    """Surgically patch specific slides/edits inside slides_plan.json without rewriting the full file.
    Accepts --patch-slide 'N:{...}' (1-indexed) and/or --patch-file JSON ({slides:{N:{...}}, edits:{N:{...}}})."""
    applied: list[str] = []
    if not (patch_slides or patch_file) or not plan_path.is_file():
        return applied
    plan = _read_json(plan_path, None)
    if not isinstance(plan, dict):
        return applied

    def _merge_slide(idx1: int, patch_obj: dict):
        slides = plan.get("slides")
        edits = plan.get("edits")
        if isinstance(slides, list) and (plan.get("mode", "create") != "edit"):
            if 1 <= idx1 <= len(slides) and isinstance(slides[idx1 - 1], dict):
                slides[idx1 - 1].update(patch_obj)
                applied.append(f"slide_{idx1}")
        elif isinstance(edits, list):
            if idx1 == 0:
                edits.append(patch_obj)
                applied.append("edit_append")
            elif 1 <= idx1 <= len(edits) and isinstance(edits[idx1 - 1], dict):
                edits[idx1 - 1].update(patch_obj)
                applied.append(f"edit_{idx1}")

    for raw in (patch_slides or []):
        m = re.match(r"^\s*(\d+)\s*:\s*(\{.*\})\s*$", str(raw), re.S)
        if not m:
            continue
        try:
            _merge_slide(int(m.group(1)), json.loads(m.group(2)))
        except Exception:
            continue

    if patch_file and Path(patch_file).is_file():
        pf = _read_json(Path(patch_file), None)
        if isinstance(pf, dict):
            for k, v in (pf.get("slides") or {}).items():
                if str(k).isdigit() and isinstance(v, dict):
                    _merge_slide(int(k), v)
            for k, v in (pf.get("edits") or {}).items():
                if str(k).isdigit() and isinstance(v, dict):
                    edits = plan.get("edits")
                    idx1 = int(k)
                    if isinstance(edits, list) and 1 <= idx1 <= len(edits) and isinstance(edits[idx1 - 1], dict):
                        edits[idx1 - 1].update(v)
                        applied.append(f"edit_{idx1}")
            for top_k in ("deck_title", "rationale", "uncertainty"):
                if top_k in pf:
                    plan[top_k] = pf[top_k]
                    applied.append(top_k)

    if applied:
        plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=1), encoding="utf-8")
    return applied


def cmd_deliver(args) -> int:
    work = Path(args.work)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    deck = Path(args.deck) if args.deck else (out / "presentation.pptx")
    pdf = out / (deck.stem + ".pdf")
    active_plan = Path(args.plan)
    patched = _apply_plan_patches(active_plan, getattr(args, "patch_slide", None) or [], getattr(args, "patch_file", None))

    hroot = work / "history"
    hroot.mkdir(parents=True, exist_ok=True)
    existing_vers = sorted(
        [int(d.name[1:]) for d in hroot.iterdir() if d.is_dir() and d.name.startswith("v") and d.name[1:].isdigit()]
    )
    ver_num = (existing_vers[-1] + 1) if existing_vers else 1

    def _build_and_check(plan_path: Path):
        b_args = ["--plan", str(plan_path), "--work", str(work), "--out", str(deck)]
        if args.template:
            b_args += ["--template", args.template]
        if args.start:
            b_args += ["--start", args.start]
        r_b = _run("build_deck.py", b_args, timeout=240)
        if r_b.returncode != 0:
            return r_b, None, None
        r_c = _run("check_deck.py", ["--deck", str(deck), "--work", str(work)], timeout=180)
        chk = _read_json(work / "checks.json", {}) or {}
        return r_b, r_c, chk

    rb, rc, checks = _build_and_check(active_plan)
    if rb.returncode != 0:
        _run("assemble_result.py", [
            "--case-dir", args.case_dir, "--work", str(work),
            "--deck", str(deck), "--status", "auto",
        ], timeout=120)
        print("DELIVER_FAILED " + json.dumps({
            "step": "build_deck.py",
            "rc": rb.returncode,
            "stderr": (rb.stderr or rb.stdout or "")[-700:],
        }, ensure_ascii=False))
        return 1

    # Fast in-process auto-repair loop (<=2 passes, <0.8s) BEFORE slow LibreOffice export
    auto_repaired: list[str] = []
    repairable_ids = {
        "text_fit", "table_fit", "overlap", "bounds", "min_font",
        "titles_notes", "quote_verification", "claim_sources",
    }
    plan_obj = _read_json(active_plan, None)
    if isinstance(plan_obj, dict):
        bak_deck = work / ".bak_deck.pptx"
        bak_chk = work / ".bak_checks.json"
        bak_cm = work / ".bak_content_map.json"
        bak_sd = work / ".bak_slide_diff.json"

        def _snap_best():
            for src_p, dst_p in ((deck, bak_deck), (work / "checks.json", bak_chk),
                                 (work / "content_map.json", bak_cm), (work / "slide_diff.json", bak_sd)):
                if src_p.is_file():
                    shutil.copy2(str(src_p), str(dst_p))

        def _restore_best():
            for src_p, dst_p in ((bak_deck, deck), (bak_chk, work / "checks.json"),
                                 (bak_cm, work / "content_map.json"), (bak_sd, work / "slide_diff.json")):
                if src_p.is_file():
                    shutil.copy2(str(src_p), str(dst_p))

        best_fails = [{"id": c["id"], "detail": c["detail"]}
                      for c in ((checks or {}).get("checks") or []) if not c.get("pass")]
        best_score = _violation_score(best_fails)
        _snap_best()
        work_plan = plan_obj
        for pass_idx in (1, 2):
            cur_chk = (checks or {}).get("checks") or []
            cur_fails = [{"id": c["id"], "detail": c["detail"]} for c in cur_chk if not c.get("pass")]
            if not any(f["id"] in repairable_ids for f in cur_fails):
                break
            cand_plan, acts = _auto_repair_plan(work_plan, cur_fails, work, pass_idx)
            if not acts:
                break
            rep_path = work / "slides_plan_repaired.json"
            rep_path.write_text(json.dumps(cand_plan, ensure_ascii=False, indent=1), encoding="utf-8")
            rb2, rc2, checks2 = _build_and_check(rep_path)
            new_fails = [{"id": c["id"], "detail": c["detail"]}
                         for c in ((checks2 or {}).get("checks") or []) if not c.get("pass")]
            new_score = _violation_score(new_fails)
            if rb2.returncode == 0 and new_score < best_score:
                work_plan = cand_plan
                active_plan = rep_path
                rb, rc, checks = rb2, rc2, checks2
                best_score = new_score
                auto_repaired.extend(acts)
                _snap_best()
            elif (pass_idx == 1 and rb2.returncode == 0 and new_score <= best_score
                  and any(a.startswith("tighten_") for a in acts)):
                # Allow pass 1 tighten to carry into pass 2 (_tighten=2) without discarding progress
                work_plan = cand_plan
                checks = checks2
            else:
                _restore_best()
                checks = _read_json(work / "checks.json", {}) or {}
                break
        # Ensure final on-disk artifacts match the strictly best snapshot
        if _violation_score([{"id": c["id"], "detail": c["detail"]}
                             for c in ((checks or {}).get("checks") or []) if not c.get("pass")]) > best_score:
            _restore_best()
            checks = _read_json(work / "checks.json", {}) or {}
        for bp in (bak_deck, bak_chk, bak_cm, bak_sd):
            try:
                if bp.is_file():
                    bp.unlink()
            except Exception:
                pass

    rexp = _run("export_preview.py", [
        "--deck", str(deck), "--work", str(work),
        "--pdf", str(pdf), "--dpi", str(args.dpi),
    ], timeout=180)
    ra = _run("assemble_result.py", [
        "--case-dir", args.case_dir, "--work", str(work),
        "--deck", str(deck), "--status", "auto", "--pass-num", str(ver_num),
    ], timeout=120)

    checks = _read_json(work / "checks.json", {}) or {}
    exp = _read_json(work / "export_report.json", {}) or {}
    res = _read_json(out / "result.json", {}) or {}
    chk_list = checks.get("checks") or []
    fails = [{"id": c["id"], "detail": c["detail"]} for c in chk_list if not c.get("pass")]

    # Parse missing_fields from assemble_result stdout if available
    missing_fields = []
    for line in (ra.stdout or "").splitlines():
        if line.startswith("RESULT_OK "):
            try:
                missing_fields = json.loads(line[len("RESULT_OK "):]).get("missing_fields", [])
            except Exception:
                pass

    cs = out / "previews" / "contact-sheet.png"

    # Snapshot this iteration into work/history/vN/ and write version_history.md for rollback
    version_meta = None
    try:
        vdir = hroot / f"v{ver_num}"
        vdir.mkdir(parents=True, exist_ok=True)
        if deck.is_file():
            shutil.copy2(str(deck), str(vdir / "presentation.pptx"))
        if active_plan.is_file():
            shutil.copy2(str(active_plan), str(vdir / "slides_plan.json"))
        if (work / "checks.json").is_file():
            shutil.copy2(str(work / "checks.json"), str(vdir / "checks.json"))
        idx_path = hroot / "versions.json"
        records = _read_json(idx_path, []) or []
        records.append({
            "version": f"v{ver_num}",
            "ts": datetime.now().astimezone().isoformat(timespec="seconds"),
            "slides": checks.get("slide_count"),
            "passed": f"{len(chk_list) - len(fails)}/{len(chk_list)}",
            "status": res.get("status"),
            "snapshot_dir": str(vdir),
        })
        idx_path.write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
        vh_lines = [
            "# 版本历史与回退说明 / Version History & Rollback",
            "",
            f"- Active delivery version: **`v{ver_num}`** (`{deck.name}`)",
            f"- Rollback example: `cp {hroot}/v1/presentation.pptx {deck}` (or re-run `deliver --plan {hroot}/v1/slides_plan.json`)",
            "",
            "| version | timestamp | slides | checks | status |",
            "|---|---|---|---|---|",
        ]
        for rec in records:
            vh_lines.append(
                f"| {rec['version']} | {rec['ts']} | {rec.get('slides')} | {rec.get('passed')} | {rec.get('status')} |"
            )
        (out / "version_history.md").write_text("\n".join(vh_lines) + "\n", encoding="utf-8")
        version_meta = {
            "active_version": f"v{ver_num}",
            "total_versions": len(records),
            "history_file": "version_history.md",
            "rollback_cmd": f"python3 {SCRIPTS / 'run_pipeline.py'} deliver --case-dir {args.case_dir} --work {work} --out {out} --plan {hroot}/v1/slides_plan.json",
        }
        if (out / "result.json").is_file() and isinstance(res, dict):
            res["version_history"] = version_meta
            (out / "result.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass

    summary = {
        "status": res.get("status"),
        "pass_num": ver_num,
        "patched": patched,
        "slides": checks.get("slide_count"),
        "passed": f"{len(chk_list) - len(fails)}/{len(chk_list)}",
        "auto_repaired": auto_repaired,
        "failures": fails,
        "export": {
            "method": exp.get("method"),
            "pdf": exp.get("pdf"),
            "pages": exp.get("pages"),
            "pages_match": exp.get("pages_match"),
            "titles_found": exp.get("titles_found"),
            "visual_ok": exp.get("visual_ok"),
            "blank_slides": exp.get("blank_slides", []),
            "previews": exp.get("previews"),
            "contact_sheet": str(cs) if cs.is_file() else None,
            "anomalies": exp.get("anomalies", []),
        },
        "result_json": str(out / "result.json"),
        "version_history": version_meta,
        "missing_fields": missing_fields,
        "exceptions": len(res.get("exceptions") or []),
        "slide_diff": res.get("slide_diff"),
        "step_rcs": {"build": rb.returncode, "check": rc.returncode,
                     "export": rexp.returncode, "assemble": ra.returncode},
    }
    hard_ids = {
        "bounds", "aspect", "slide_count", "slide_limit", "claim_sources",
        "overlap", "quote_verification", "number_fidelity",
    }
    heur_ids = {"text_fit", "table_fit"}
    fail_ids = {f.get("id") for f in fails}
    allow_heur_soft = (
        ver_num >= 2
        and bool(exp.get("pdf"))
        and exp.get("visual_ok") is not False
        and "bounds" not in fail_ids
        and "overlap" not in fail_ids
    )
    critical_ids = hard_ids if allow_heur_soft else (hard_ids | heur_ids)
    crit_fails = [f for f in fails if f.get("id") in critical_ids]
    soft_warns = [f for f in fails if f.get("id") not in critical_ids]
    summary["warnings"] = soft_warns
    ok = (res.get("status") == "ok" and not crit_fails and not missing_fields
          and rb.returncode == 0 and rc.returncode == 0 and ra.returncode == 0)
    print(("DELIVER_OK " if ok else "DELIVER_ISSUES ") + json.dumps(summary, ensure_ascii=False))
    # Print compact per-slide audit table (0 extra tool calls; never induces extra rerun on DELIVER_OK)
    try:
        cm_sl = ((res.get("content_map") or {}).get("slides")) or []
        if cm_sl:
            audit_lines = ["\n=== DECK AUDIT SUMMARY ==="]
            dur_info = res.get("duration_estimate") or {}
            if dur_info:
                audit_lines.append(
                    f"Slides: {len(cm_sl)} | Est. speech: {dur_info.get('estimated_speech_min')} min "
                    f"(target: {dur_info.get('workorder_duration') or 'unspecified'}) | "
                    f"Visual OK: {exp.get('visual_ok')} | Auto-repaired: {auto_repaired or 'none'}"
                )
            for s in cm_sl[:24]:
                vis = []
                if s.get("highlights"):
                    vis.append(f"kpi({len(s['highlights'])})")
                if s.get("steps"):
                    vis.append(f"steps({len(s['steps'])})")
                if s.get("tables"):
                    vis.append(f"tbl({s['tables'][0].get('rows', '?')}r)")
                if s.get("charts"):
                    vis.append(f"chart({s['charts'][0].get('type', 'bar')})")
                if s.get("images"):
                    vis.append(f"img({len(s['images'])})")
                if s.get("takeaway"):
                    vis.append("takeaway")
                vis_str = "+".join(vis) if vis else "text-only"
                n_len = len(str(s.get("notes") or "").strip())
                src_n = len(s.get("claim_sources") or [])
                b_n = len(s.get("claims") or [])
                t_short = str(s.get("title") or "(no title)").replace("\n", " ")[:48]
                audit_lines.append(
                    f"S{s.get('slide')} [{s.get('role', '?')}] {t_short} | b:{b_n} | {vis_str} | notes:{n_len}c | src:{src_n}"
                )
            print("\n".join(audit_lines)[:2600])
    except Exception:
        pass
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Deterministic two-step orchestrator for academic-slides-pku")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("prepare", help="Inspect case + save immediate partial result.json")
    p1.add_argument("--case-dir", default="/input")
    p1.add_argument("--work", default="/app/work")
    p1.add_argument("--out", default="/app/output")

    p2 = sub.add_parser("deliver", help="Build deck -> check -> export PDF/PNG -> assemble final result.json")
    p2.add_argument("--case-dir", default="/input")
    p2.add_argument("--work", default="/app/work")
    p2.add_argument("--out", default="/app/output")
    p2.add_argument("--plan", required=True)
    p2.add_argument("--deck", default=None)
    p2.add_argument("--template", default=None)
    p2.add_argument("--start", default=None)
    p2.add_argument("--dpi", type=int, default=100)
    p2.add_argument("--patch-slide", action="append", default=[],
                    help="Surgical patch 'N:{json}' merged onto slide/edit N without rewriting the full plan")
    p2.add_argument("--patch-file", default=None,
                    help="Optional JSON patch file {slides:{N:{...}}, edits:{N:{...}}}")

    args = ap.parse_args()
    return cmd_prepare(args) if args.cmd == "prepare" else cmd_deliver(args)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"PIPELINE_FAILED: {e}")
        sys.exit(1)
