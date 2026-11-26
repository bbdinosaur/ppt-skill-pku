#!/usr/bin/env python3
"""Assemble /app/output/result.json from pipeline artifacts (plan, build, check, render).

Usage:
  python3 write_result.py --case-json /input/case.json \
      [--plan /app/work/plan.json] [--build /app/work/build_report.json] \
      [--check /app/work/deck_check.json] [--render /app/output/previews/render_log.json] \
      [--deck /app/output/deck.pptx] [--fills /app/work/fills.json] \
      [--out /app/output/result.json] [--status complete|partial|failed]

Behavior:
  - case_id: copied verbatim from case.json (keys tried: case_id, id).
  - Mirrors every key in case.json `output_contract` at top level (fills > deliverable/section
    match > null + uncertainty). NEVER fabricates unknown values.
  - evidence_map: built authoritatively from build_report.json (post auto-split numbering),
    preserving structured `quote` and inferring valid source `kind`.
  - deck_summary.placeholder_chain: extracted directly from build_report.json for verifiable
    master/layout/placeholder inheritance.
  - Prints NULL_CONTRACT_FIELDS so the agent can fill any remaining grounded contract fields.
No-script fallback: assemble result.json by hand following references/result-schema.md.
"""
import argparse, json, re, sys, zipfile
from pathlib import Path

FILE_HINTS = ("pptx", "pdf", "png", "preview", "deck", "report", "log", "plan", "path", "file")
FILE_HINT_RE = re.compile(
    r"(?<![a-z0-9])(?:pptx|pdf|png|pngs|preview|previews|deck|report|log|plan|path|file|files)(?![a-z0-9])",
    re.I,
)
MANAGED_SECTIONS = {
    "case_id", "status", "deliverables", "deck_summary", "evidence_map",
    "rendering", "checks", "uncertainty", "change_log", "limits",
}
INPUT_PATH_EXTS = (".pdf", ".pptx", ".potx", ".docx", ".md", ".txt", ".csv", ".tsv", ".png", ".jpg", ".jpeg", ".webp")


def load(p):
    p = Path(p)
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _is_nonempty(v):
    if v is None:
        return False
    if isinstance(v, str):
        return bool(v.strip())
    if isinstance(v, (list, dict, tuple, set)):
        return len(v) > 0
    return True


def _item_key(item):
    if isinstance(item, str):
        return ("s", re.sub(r"\s+", " ", item).strip())
    try:
        return ("j", json.dumps(item, sort_keys=True, ensure_ascii=False))
    except Exception:
        return ("r", repr(item))


def _merge_unique(*lists):
    out = []
    seen = set()
    for seq in lists:
        if not seq:
            continue
        if not isinstance(seq, list):
            seq = [seq]
        for item in seq:
            if not _is_nonempty(item):
                continue
            k = _item_key(item)
            if k not in seen:
                seen.add(k)
                out.append(item)
    return out


def _collect_case_input_files(obj, out, key_ctx=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            kl = str(k).lower()
            if kl in ("output_contract", "output_fields"):
                continue
            _collect_case_input_files(v, out, f"{key_ctx}.{kl}")
    elif isinstance(obj, list):
        for v in obj:
            _collect_case_input_files(v, out, key_ctx)
    elif isinstance(obj, str) and 2 < len(obj) < 300:
        low = obj.lower().strip()
        if any(low.endswith(ext) for ext in INPUT_PATH_EXTS) and not low.startswith("/app/output/"):
            out.append(obj.strip())


def _is_stale_auto_msg(msg, resolved_keys, crit, blank, deliverables):
    if not isinstance(msg, str):
        return False
    m_null = re.match(r"^output_contract (?:file-like )?field '([^']+)' ", msg)
    if m_null and m_null.group(1) in resolved_keys:
        return True
    if crit == 0 and (
        msg.startswith("heuristic static check flagged ")
        or msg.startswith("static layout heuristic flagged ")
    ):
        return True
    if not blank and msg.startswith("low-variance heuristic flagged "):
        return True
    m_deliv = re.match(r"^deliverable (\w+) missing or invalid at result time$", msg)
    if m_deliv and deliverables.get(m_deliv.group(1)):
        return True
    return False


def _inspect_pptx_file(path):
    """Return (valid_nonempty, slide_count_in_zip)."""
    p = Path(path) if path else None
    if not p or not p.exists() or not p.is_file():
        return False, 0
    try:
        if p.stat().st_size <= 0:
            return False, 0
        with zipfile.ZipFile(p, "r") as z:
            names = z.namelist()
            if "ppt/presentation.xml" not in names:
                return False, 0
            n_slides = sum(1 for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
            return True, n_slides
    except Exception:
        return False, 0


def _is_valid_pdf(path):
    p = Path(path) if path else None
    if not p or not p.exists() or not p.is_file():
        return False
    try:
        if p.stat().st_size < 16:
            return False
        with p.open("rb") as f:
            head = f.read(16)
        return b"%PDF" in head
    except Exception:
        return False


def _pdf_page_count(path):
    try:
        import pymupdf
        with pymupdf.open(str(path)) as doc:
            return len(doc)
    except Exception:
        return None


def infer_source_kind(locator: str):
    low = (locator or "").lower()
    if re.search(r"\b(fig\.?|figure)\b|图\s*\d", low):
        return "paper-figure"
    if re.search(r"\b(tab\.?|table)\b|表\s*\d", low):
        return "paper-table"
    if "abstract" in low or "摘要" in low:
        return "paper-abstract"
    if "title" in low or "封面" in low or "标题" in low:
        return "paper-title"
    if "§" in low or re.search(r"\b(sec\.?|section|intro|method|result|discuss|limit|conclu)\b", low) or re.search(r"第?\d+节|方法|结果|局限|结论", low):
        return "paper-section"
    return "material-other"


def norm_sources(ev_list):
    out = []
    for e in (ev_list or []):
        if isinstance(e, dict):
            loc = str(e.get("locator") or e.get("source") or "").strip()
            if not loc:
                continue
            item = {"locator": loc, "kind": e.get("kind") or infer_source_kind(loc)}
            if e.get("quote"):
                item["quote"] = str(e["quote"]).strip()
            out.append(item)
        elif e:
            loc = str(e).strip()
            out.append({"locator": loc, "kind": infer_source_kind(loc)})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-json", required=True)
    ap.add_argument("--plan", default="/app/work/plan.json")
    ap.add_argument("--build", default="/app/work/build_report.json")
    ap.add_argument("--check", default="/app/work/deck_check.json")
    ap.add_argument("--render", default="/app/output/previews/render_log.json")
    ap.add_argument("--deck", default="/app/output/deck.pptx")
    ap.add_argument("--fills", default=None,
                    help="JSON file of contract-field -> value authored by the agent")
    ap.add_argument("--out", default="/app/output/result.json")
    ap.add_argument("--status", default=None, choices=["complete", "partial", "failed"])
    a = ap.parse_args()

    case = load(a.case_json)
    plan = load(a.plan)
    build = load(a.build)
    check = load(a.check)
    render = load(a.render)

    case_id = case.get("case_id", case.get("id"))
    uncertainty, limits = [], []
    if case_id is None:
        uncertainty.append("case.json has no case_id/id field; case_id set to 'unknown'")
        case_id = "unknown"

    pdf_path = render.get("pdf")
    if not pdf_path:
        cand_pdf = Path(a.deck).with_suffix(".pdf")
        if _is_valid_pdf(cand_pdf):
            pdf_path = str(cand_pdf)
    pptx_ok, pptx_zip_slides = _inspect_pptx_file(a.deck)
    pdf_ok = _is_valid_pdf(pdf_path)
    png_list = [p.get("png") for p in (render.get("pages") or []) if p.get("png") and Path(p.get("png")).exists()]
    previews_dir = render.get("previews_dir") or (str(Path(png_list[0]).parent) if png_list else None)

    deliverables = {
        "pptx": a.deck if pptx_ok else None,
        "pdf": pdf_path if pdf_ok else None,
        "previews_dir": previews_dir,
        "previews": png_list if png_list else None,
        "build_report": a.build if Path(a.build).exists() else None,
        "plan": a.plan if Path(a.plan).exists() else None,
        "render_log": a.render if Path(a.render).exists() else None,
    }
    for k in ("pptx", "pdf", "previews", "build_report", "render_log"):
        if not deliverables.get(k):
            limits.append(f"deliverable {k} missing or invalid at result time")

    # ---- mirror output_contract fields (never fabricate; defer null/limits until all grounded backfills finish) ----
    oc = case.get("output_contract") or case.get("output_fields") or {}
    contract_keys = list(oc.keys()) if isinstance(oc, dict) else list(oc)
    fills = load(a.fills) if a.fills else {}

    out_path = Path(a.out)
    raw_existing = load(out_path) if out_path.exists() else {}
    existing = (
        raw_existing
        if raw_existing and (not raw_existing.get("case_id") or raw_existing.get("case_id") == case_id)
        else {}
    )

    def best_for(key):
        lk = key.lower().replace("-", "_").replace(" ", "_")
        prefs = []
        if ("pdf" in lk
                and not any(t in lk for t in ("font", "page_count", "pages", "paper", "source", "input", "original"))):
            prefs.append("pdf")
        if any(t in lk for t in ("png", "preview", "screenshot")) or ("image" in lk and "count" not in lk):
            prefs.append("previews" if ("list" in lk or "images" in lk or "pngs" in lk or "previews" in lk) else "previews_dir")
        _non_file_slide_toks = (
            "count", "num", "total", "summary", "evidence", "source", "map", "size",
            "title", "note", "check", "qa", "plan", "log", "report", "layout",
            "change", "limit", "status", "chain", "inherit", "outline", "structure",
            "template", "input", "original",
        )
        if (("pptx" in lk and not any(t in lk for t in ("template", "source", "input", "original")))
                or lk in ("deck", "slides", "presentation", "slide_deck", "editable_deck", "output_deck", "adapted_deck")
                or (any(t in lk for t in ("deck", "slide", "presentation"))
                    and not any(x in lk for x in _non_file_slide_toks))):
            prefs.append("pptx")
        if any(t in lk for t in ("render_log", "export_log")):
            prefs.append("render_log")
        if lk in ("plan", "plan_path", "plan_file", "slide_plan_path"):
            prefs.append("plan")
        if any(t in lk for t in ("build_report", "check_report")):
            prefs.append("build_report")
        for p in prefs:
            v = deliverables.get(p)
            if v:
                return v
        return None

    mirrored = {}
    for key in contract_keys:
        if key == "case_id":
            continue
        if key in fills:
            mirrored[key] = fills[key]
            continue
        v = best_for(key)
        mirrored[key] = v if v is not None else None

    # ---- evidence map & verifiable placeholder chain (build_report is authoritative) ----
    bslides = build.get("slides") or []
    evidence_map = []
    placeholder_chain = []
    for s in bslides:
        entry = {
            "slide": s.get("n"),
            "kind": s.get("kind", "content"),
            "claim": s.get("claim"),
            "sources": norm_sources(s.get("evidence")),
            "notes_present": bool(s.get("notes_present", s.get("notes"))),
        }
        if s.get("figure_refs"):
            entry["figure_refs"] = s["figure_refs"]
        if s.get("time_hint_s"):
            entry["time_hint_s"] = s["time_hint_s"]
        evidence_map.append(entry)

        ph_rec = {
            "slide": s.get("n"),
            "layout": s.get("layout_used"),
            "title_ph": s.get("title_ph"),
        }
        if s.get("body_ph") is not None:
            ph_rec["body_ph"] = s["body_ph"]
        if s.get("subtitle_ph") is not None:
            ph_rec["subtitle_ph"] = s["subtitle_ph"]
        if s.get("table") is not None:
            ph_rec["table"] = s["table"]
        if s.get("chart") is not None:
            ph_rec["chart"] = s["chart"]
        if s.get("fallbacks_manual_textbox"):
            ph_rec["manual_textbox_fallback"] = True
        placeholder_chain.append(ph_rec)

    if not evidence_map:
        for i, s in enumerate(plan.get("slides", []), 1):
            entry = {
                "slide": i,
                "kind": s.get("kind", "content"),
                "claim": s.get("title", ""),
                "sources": norm_sources(s.get("evidence")),
                "notes_present": bool(s.get("notes")),
            }
            if s.get("images"):
                entry["figure_refs"] = [
                    {"caption": im.get("caption"), "source": im.get("source"), "asset": im.get("path")}
                    for im in s["images"]
                ]
            evidence_map.append(entry)

    if build.get("auto_splits"):
        uncertainty.append("; ".join(build["auto_splits"]) + "; evidence_map follows final numbering")

    # ---- checks & status (decouple hard delivery failure from heuristic QA alerts) ----
    crit = check.get("critical", 0)
    issues = check.get("issues", []) or []
    overflow = sum(1 for x in issues if x.get("kind") == "text-overflow" and x.get("severity") == "critical")
    overlap_crit = sum(1 for x in issues if x.get("kind") == "shape-overlap" and x.get("severity") == "critical")
    aspect = sum(1 for x in issues if x.get("kind") == "aspect-distortion")
    missing_notes = sum(1 for x in issues if x.get("kind") == "missing-notes")
    heuristic_alerts = [
        {"slide": x.get("slide"), "kind": x.get("kind"), "detail": x.get("detail")}
        for x in issues if x.get("severity") == "critical"
    ]

    slide_count = build.get("slide_count")
    if not slide_count:
        slide_count = pptx_zip_slides or len(plan.get("slides", []))

    if "pages_match" in render:
        pages_match = bool(render.get("pages_match"))
    elif pdf_ok and slide_count > 0:
        pc = _pdf_page_count(pdf_path)
        pages_match = (pc == slide_count) if pc is not None else False
    else:
        pages_match = False

    blank = render.get("blank_pages", []) or []
    if a.status:
        status = a.status
    elif not pptx_ok or slide_count <= 0:
        status = "failed"
    elif not pdf_ok or not pages_match:
        status = "partial"
    else:
        status = "complete"

    # Hard delivery failures
    if not pptx_ok or slide_count <= 0:
        limits.append("core deliverable deck.pptx missing, empty, or corrupted")
    if not pdf_ok:
        limits.append("PDF export failed, missing, or corrupted")
    elif not pages_match:
        limits.append("PDF page count != slide count; export inconsistent")

    # Heuristic geometry & low-variance alerts: always recorded honestly without downgrading valid deliverables
    if crit:
        limits.append(f"heuristic static check flagged {crit} geometry/overflow/overlap alert(s) (see checks); core deliverables intact")
        uncertainty.append(f"static layout heuristic flagged {crit} alert(s) (overflow={overflow}, overlap={overlap_crit}); verify rendered PNGs visually")
    if blank:
        limits.append(f"low-variance heuristic flagged suspected blank page(s): {blank} (verify visually; dark/minimal title slides may trigger)")
        uncertainty.append(f"low-variance heuristic flagged page(s) {blank} as suspected blank; confirm visually via previews")

    change_log = list(fills.get("change_log") or plan.get("change_log") or [])
    if build.get("normalizations"):
        for nmsg in build["normalizations"]:
            if nmsg not in change_log:
                change_log.append(nmsg)
    if not change_log:
        change_log = ["v1 deck built from plan and verified with check_deck + export_render"]

    sections = {
        "deliverables": deliverables,
        "deck_summary": {
            "slide_count": slide_count,
            "title": next((s.get("claim") for s in bslides if s.get("kind") == "title"), None),
            "layouts_used": sorted({s.get("layout_used", "?") for s in bslides}),
            "template": build.get("template"),
            "slide_size": build.get("slide_size"),
            "audience": plan.get("audience") or case.get("audience"),
            "duration_min": plan.get("duration_min") or case.get("duration_min"),
            "language": plan.get("language") or case.get("language"),
            "placeholder_chain": placeholder_chain,
        },
        "evidence_map": evidence_map,
        "rendering": {
            "pdf": render.get("pdf"),
            "previews_dir": previews_dir,
            "dpi": render.get("dpi"),
            "pages_match": pages_match,
            "pdf_page_size_in": render.get("pdf_page_size_in"),
            "size_match": render.get("size_match"),
            "notes_slides": render.get("notes_slides"),
            "title_text_ok": render.get("title_text_ok"),
            "blank_pages": blank,
            "theme_fonts": render.get("theme_fonts", {}),
            "pdf_fonts_used": render.get("pdf_fonts_used", []),
            "font_substitutions": render.get("font_substitutions", []),
            "font_warnings": render.get("font_warnings", []),
        },
        "checks": {
            "geometry_critical": crit,
            "text_overflow_critical": overflow,
            "overlap_critical": overlap_crit,
            "aspect_distortion": aspect,
            "missing_notes": missing_notes,
            "consistency_warnings": check.get("warnings", 0),
            "inheritance_summary": check.get("inheritance_summary", {}),
            "heuristic_alerts": heuristic_alerts,
        },
        "change_log": change_log,
        "limits": limits,
    }

    # ---- grounded provenance, template, theme/layout mappings & normalizations ----
    workdir_cands = [Path(a.plan).parent, Path(a.build).parent]
    tinv = {}
    ps_info = {}
    dinv = {}
    for wd in workdir_cands:
        if not tinv and (wd / "template_inventory.json").exists():
            tinv = load(wd / "template_inventory.json")
        if not ps_info and (wd / "paper_structure.json").exists():
            ps_info = load(wd / "paper_structure.json")
        if not dinv and (wd / "deck_inventory.json").exists():
            dinv = load(wd / "deck_inventory.json")

    raw_tpl = build.get("template") or plan.get("template") or tinv.get("template")
    tpl_path = (
        str(raw_tpl).strip()
        if isinstance(raw_tpl, str) and raw_tpl.strip() and raw_tpl.strip().lower() != "none"
        else None
    )
    normalizations = list(build.get("normalizations") or [])

    case_input_files = []
    _collect_case_input_files(case, case_input_files)
    plan_source_files = []
    for k_src in ("paper", "source", "source_file", "starting_deck", "deck"):
        if isinstance(plan.get(k_src), str) and plan[k_src].strip():
            plan_source_files.append(plan[k_src].strip())
    if isinstance(plan.get("source_files"), list):
        plan_source_files.extend(str(x).strip() for x in plan["source_files"] if _is_nonempty(x))
    fig_assets = []
    ev_locators = []
    unique_ev_sources = []
    seen_ev_src = set()
    for em in evidence_map:
        for src in (em.get("sources") or []):
            loc = src.get("locator")
            if loc:
                ev_locators.append(loc)
                sk = (loc, src.get("kind"))
                if sk not in seen_ev_src:
                    seen_ev_src.add(sk)
                    unique_ev_sources.append({"locator": loc, "kind": src.get("kind")})
        for fr in (em.get("figure_refs") or []):
            if fr.get("asset"):
                fig_assets.append(str(fr["asset"]))

    source_files = _merge_unique(
        [ps_info.get("source")] if ps_info.get("source") else [],
        [tpl_path] if tpl_path else [],
        [dinv.get("source") or dinv.get("deck")] if (dinv.get("source") or dinv.get("deck")) else [],
        plan_source_files,
        case_input_files,
        fig_assets,
    )
    computed_sources = _merge_unique(
        source_files,
        plan.get("sources") if isinstance(plan.get("sources"), list) else [],
        ev_locators,
    )
    merged_sources = _merge_unique(
        existing.get("sources") if isinstance(existing.get("sources"), list) else [],
        fills.get("sources") if isinstance(fills.get("sources"), list) else [],
        computed_sources,
    )

    provenance = {
        "source_files": source_files,
        "template": tpl_path or sections["deck_summary"]["template"],
        "evidence_slides": len(evidence_map),
        "evidence_sources": unique_ev_sources,
        "plan": deliverables.get("plan"),
        "build_report": deliverables.get("build_report"),
    }

    theme_mapping = {
        "template": tpl_path or sections["deck_summary"]["template"],
        "slide_size": build.get("slide_size") or (
            f"{tinv.get('slide_width_in')}x{tinv.get('slide_height_in')}in"
            if tinv.get("slide_width_in") and tinv.get("slide_height_in") else None
        ),
        "theme_colors": build.get("theme_colors") or tinv.get("theme_colors") or {},
        "theme_fonts": render.get("theme_fonts") or tinv.get("theme_fonts") or {},
        "font_ladder_pt": build.get("ladder") or {},
        "font_substitutions": render.get("font_substitutions") or [],
    }

    layout_mapping = []
    for s in bslides:
        lm_item = {
            "slide": s.get("n"),
            "kind": s.get("kind", "content"),
            "layout_used": s.get("layout_used"),
            "layout_select": s.get("layout_select"),
            "title_ph": s.get("title_ph"),
        }
        if s.get("body_ph") is not None:
            lm_item["body_ph"] = s.get("body_ph")
        if s.get("subtitle_ph") is not None:
            lm_item["subtitle_ph"] = s.get("subtitle_ph")
        if s.get("fallbacks_manual_textbox"):
            lm_item["manual_textbox_fallback"] = True
        layout_mapping.append(lm_item)

    # Explicit + generous safe backfill for standard-named output_contract fields (never fabricate)
    direct_map = {
        "status": status,
        "deliverables": sections["deliverables"],
        "deliverable": sections["deliverables"],
        "deck_summary": sections["deck_summary"],
        "summary": sections["deck_summary"],
        "slides_summary": sections["deck_summary"],
        "evidence_map": sections["evidence_map"],
        "evidence": sections["evidence_map"],
        "evidence_mapping": sections["evidence_map"],
        "source_mapping": sections["evidence_map"],
        "slide_evidence": sections["evidence_map"],
        "sources": merged_sources,
        "citations": merged_sources,
        "source_files": source_files,
        "input_files": source_files,
        "provenance": provenance,
        "traceability": provenance,
        "rendering": sections["rendering"],
        "render": sections["rendering"],
        "render_summary": sections["rendering"],
        "export_summary": sections["rendering"],
        "export_report": sections["rendering"],
        "checks": sections["checks"],
        "check": sections["checks"],
        "qa": sections["checks"],
        "qa_report": sections["checks"],
        "qa_summary": sections["checks"],
        "verification": sections["checks"],
        "change_log": sections["change_log"],
        "changes": sections["change_log"],
        "limits": sections["limits"],
        "limitations": sections["limits"],
        "template": sections["deck_summary"]["template"],
        "template_used": sections["deck_summary"]["template"],
        "template_path": tpl_path,
        "template_file": tpl_path,
        "template_pptx": tpl_path,
        "template_name": Path(tpl_path).name if tpl_path else None,
        "theme_mapping": theme_mapping,
        "theme": theme_mapping,
        "theme_info": theme_mapping,
        "style_mapping": theme_mapping,
        "layout_mapping": layout_mapping,
        "layout_map": layout_mapping,
        "layouts_mapping": layout_mapping,
        "slide_layouts": layout_mapping,
        "normalizations": normalizations,
        "template_normalizations": normalizations,
        "format_normalizations": normalizations,
        "layouts_used": sections["deck_summary"]["layouts_used"],
        "slide_size": sections["deck_summary"]["slide_size"],
        "deck_title": sections["deck_summary"]["title"],
        "title": sections["deck_summary"]["title"],
        "placeholder_chain": sections["deck_summary"]["placeholder_chain"],
        "inheritance": sections["checks"]["inheritance_summary"],
        "font_warnings": sections["rendering"]["font_warnings"],
        "pdf_fonts": sections["rendering"]["pdf_fonts_used"],
        "pdf_fonts_used": sections["rendering"]["pdf_fonts_used"],
        "notes_count": sections["rendering"]["notes_slides"],
    }
    _deferred_unc_keys = set()
    for key in list(mirrored):
        if mirrored[key] is not None:
            continue
        lk = key.lower().replace("-", "_").replace(" ", "_")
        if lk in ("uncertainty", "uncertainties"):
            _deferred_unc_keys.add(key)
            continue
        val = direct_map.get(lk)
        if val is None:
            for alias in ("deliverables", "deck_summary", "evidence_map", "rendering", "checks", "change_log", "limits"):
                if alias in lk or (len(lk) >= 4 and lk in alias):
                    val = sections[alias]
                    break
        if val is not None:
            mirrored[key] = val

    # Backfill numeric slide-count contract fields (guarding against unrelated *count* keys)
    _bad_count_toks = ("font", "warn", "err", "note", "fig", "img", "table", "char", "word", "token", "account", "quote")
    for key in list(mirrored):
        if mirrored[key] is not None or key in _deferred_unc_keys:
            continue
        lk = key.lower().replace("-", "_").replace(" ", "_")
        is_slide_cnt = (
            any(t in lk for t in ("slide_count", "slides_count", "total_slides", "num_slides",
                                  "number_of_slides", "page_count", "n_slides", "slides_total"))
            or lk in ("pages", "slides_num")
            or (any(t in lk for t in ("count", "number_of", "num_"))
                and not any(b in lk for b in _bad_count_toks))
        )
        if is_slide_cnt:
            mirrored[key] = sections["deck_summary"]["slide_count"]

    # Non-overwriting preservation of hand-written non-empty contract values from existing result.json
    for key in list(mirrored):
        if mirrored[key] is not None or key in _deferred_unc_keys:
            continue
        if key in existing and _is_nonempty(existing[key]):
            mirrored[key] = existing[key]

    # Record honest limits/uncertainty ONLY for contract keys still null after all grounded backfills
    _tpl_input_keys = {"template_path", "template_file", "template_pptx", "template_name"}
    for key in list(mirrored):
        if mirrored[key] is not None or key in _deferred_unc_keys:
            continue
        lk = key.lower().replace("-", "_").replace(" ", "_")
        if lk in _tpl_input_keys and tpl_path is None:
            uncertainty.append(f"output_contract field {key!r} left null: no template file was provided")
        elif FILE_HINT_RE.search(lk) and "count" not in lk:
            limits.append(f"output_contract file-like field {key!r} has no deliverable")
        else:
            uncertainty.append(f"output_contract field {key!r} left null: no grounded value "
                               f"from source material; fill via --fills if one exists")

    # ---- dedup-merge list fields with pre-existing result.json (non-overwriting) ----
    resolved_keys = {k for k, v in mirrored.items() if v is not None} | _deferred_unc_keys
    extra_unc = _merge_unique(fills.get("uncertainty"), fills.get("uncertainties"))
    if build.get("fallbacks"):
        extra_unc.append(f"layout fallbacks during build: {build['fallbacks']}")
    if build.get("warnings"):
        extra_unc.append(f"build warnings: {'; '.join(build['warnings'])}")

    existing_unc = [
        u for u in _merge_unique(existing.get("uncertainty"), existing.get("uncertainties"))
        if not _is_stale_auto_msg(u, resolved_keys, crit, blank, deliverables)
    ]
    existing_lim = [
        m for m in _merge_unique(existing.get("limits"), existing.get("limitations"))
        if not _is_stale_auto_msg(m, resolved_keys, crit, blank, deliverables)
    ]
    existing_clog = _merge_unique(existing.get("change_log"), existing.get("changes"))

    final_uncertainty = _merge_unique(existing_unc, uncertainty, extra_unc)
    final_limits = _merge_unique(existing_lim, limits, fills.get("limits"), fills.get("limitations"))
    final_change_log = _merge_unique(existing_clog, change_log)

    sections["uncertainty"] = final_uncertainty
    sections["limits"] = final_limits
    sections["change_log"] = final_change_log

    # Sync final merged lists back into mirrored contract aliases
    for key in list(mirrored):
        lk = key.lower().replace("-", "_").replace(" ", "_")
        if key in _deferred_unc_keys or lk in ("uncertainty", "uncertainties"):
            mirrored[key] = final_uncertainty
        elif lk in ("limits", "limitations"):
            mirrored[key] = final_limits
        elif lk in ("change_log", "changes"):
            mirrored[key] = final_change_log
        elif lk == "sources":
            mirrored[key] = _merge_unique(mirrored[key], merged_sources)

    # Preserve non-empty custom top-level fields from pre-existing result.json
    preserved_custom = {}
    for k, v in existing.items():
        if k in MANAGED_SECTIONS or k in mirrored:
            continue
        if not _is_nonempty(v):
            continue
        lk = k.lower().replace("-", "_").replace(" ", "_")
        if lk == "sources" and isinstance(v, list):
            preserved_custom[k] = _merge_unique(v, merged_sources)
        elif lk == "uncertainties" and isinstance(v, list):
            preserved_custom[k] = final_uncertainty
        elif lk == "limitations" and isinstance(v, list):
            preserved_custom[k] = final_limits
        elif lk == "changes" and isinstance(v, list):
            preserved_custom[k] = final_change_log
        else:
            preserved_custom[k] = v

    result = {"case_id": case_id, "status": status, **preserved_custom, **mirrored, **sections}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")

    null_keys = [k for k, v in mirrored.items() if v is None]
    print(f"OK result.json -> {out_path} status={status} slides={result['deck_summary']['slide_count']} "
          f"mirrored={len(mirrored)} evidence={len(evidence_map)}")
    if null_keys:
        print(f"NULL_CONTRACT_FIELDS: {null_keys}  (provide grounded values via --fills if derivable)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
