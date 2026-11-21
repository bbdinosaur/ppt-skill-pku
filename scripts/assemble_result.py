#!/usr/bin/env python3
"""assemble_result.py — Step 6 of the academic-slides skill.

Assembles /app/output/result.json with ALL output-contract root fields
(case_id, status, editable_source, export_artifact, content_map,
compliance_checks, exceptions) plus honest extras: plan, uncertainty,
time/change records, repro commands. Also writes plan.json + change_log.md.

Usage: python3 assemble_result.py --case-dir /input --work work \
          --deck out/presentation.pptx --status auto
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckutil as du  # noqa: E402

CONTRACT_FIELDS = ["case_id", "status", "editable_source", "export_artifact",
                   "content_map", "compliance_checks", "exceptions"]


def read_json(p: Path, default=None):
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def build_change_log_md(work: Path, out_md: Path):
    lines = ["# 修改记录 / Change log", ""]
    jl = work / "change_log.jsonl"
    if jl.is_file():
        for line in jl.read_text(encoding="utf-8").splitlines():
            try:
                e = json.loads(line)
                lines.append(f"- `{e.get('ts','')}` **{e.get('action','')}** — {e.get('detail','')}")
            except Exception:
                continue
    else:
        lines.append("- (no recorded changes)")
    out_md.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", default="/input")
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--deck", default="/app/output/presentation.pptx")
    ap.add_argument("--status", default="auto", choices=("auto", "ok", "partial", "failed"))
    args = ap.parse_args()

    work = Path(args.work)
    out_dir = Path(args.deck).parent
    inv = read_json(work / "inventory.json", {}) or {}
    content_map = read_json(work / "content_map.json", {}) or {}
    checks = read_json(work / "checks.json", {}) or {}
    export = read_json(work / "export_report.json", {}) or {}
    plan = read_json(work / "slides_plan.json", {}) or {}

    deck = Path(args.deck)
    pdf = Path(export.get("pdf") or (deck.parent / (deck.stem + ".pdf")))
    mode = content_map.get("mode") or plan.get("mode") or "create"

    exceptions = []
    # figure-missing / bitmap-only dispositions from the build
    for s in (content_map.get("slides") or []):
        img = s.get("image") or {}
        if img.get("missing") and img.get("path"):
            exceptions.append({"code": "FIGURE_MISSING",
                               "detail": f"slide {s.get('slide')}: image file not found: {img.get('path')}",
                               "disposition": "on-slide placeholder with ref+caption; no fabricated figure"})
    # input-level anomalies
    for a in inv.get("anomalies", []):
        if "not found" in a or "not identified" in a or "missing" in a or "unreadable" in a:
            code = "MISSING_INPUT" if ("not found" in a or "missing" in a) else "INPUT_ANOMALY"
            exceptions.append({"code": code, "detail": a, "disposition": "recorded; worked around or partial"})
    # check failures
    for c in checks.get("checks", []):
        if not c.get("pass"):
            code = {"slide_count": "PAGE_COUNT_MISMATCH", "slide_limit": "PAGE_COUNT_MISMATCH",
                    "bounds": "OUT_OF_BOUNDS", "text_fit": "OVERFLOW", "overlap": "OVERLAP",
                    "aspect": "ASPECT_CHANGED", "titles_notes": "NOTE_MISSING",
                    "min_font": "FONT_TOO_SMALL", "font_inheritance": "FONT_OVERRIDE",
                    "claim_sources": "EVIDENCE_MISSING", "quote_verification": "QUOTE_NOT_FOUND",
                    "number_fidelity": "NUMBER_NOT_FOUND", "table_fit": "TABLE_OVERFLOW",
                    }.get(c["id"], "CHECK_FAILED")
            exceptions.append({"code": code, "detail": c["detail"],
                               "disposition": "reported honestly; see content_map for affected slides"})
    # export anomalies
    for a in export.get("anomalies", []):
        code = "EXPORT_FALLBACK" if "LibreOffice" in a else "EXPORT_ANOMALY"
        exceptions.append({"code": code, "detail": a,
                           "disposition": "fallback preview retained for inspection"})

    # status logic
    deck_ok = deck.is_file() and deck.stat().st_size > 5000
    critical_bad = [c for c in checks.get("checks", [])
                    if not c.get("pass") and c["id"] in
                    ("bounds", "aspect", "slide_count", "slide_limit", "claim_sources",
                     "overlap", "text_fit", "table_fit", "quote_verification", "number_fidelity")]
    if args.status == "auto":
        if not deck_ok:
            status = "failed"
        elif critical_bad or not export.get("pdf"):
            status = "partial"
        else:
            status = "ok"
    else:
        status = args.status

    # content_map → contract shape (slide-level claims + evidence)
    cm_slides = []
    for s in content_map.get("slides", []):
        cm_slides.append({
            "slide": s.get("slide"),
            "role": s.get("role"),
            "layout": s.get("layout"),
            "title": s.get("title"),
            "claims": s.get("bullets", []),
            "images": ([{"path": (s.get("image") or {}).get("path"),
                         "ref": (s.get("image") or {}).get("ref"),
                         "caption": (s.get("image") or {}).get("caption")}]
                       if s.get("image") else []),
            "tables": ([{"rows": len((s.get("table") or {}).get("rows", []))}] if s.get("table") else []),
            "notes": s.get("notes", ""),
            "claim_sources": s.get("sources", []),
        })
    if mode == "edit" and content_map.get("edits_applied"):
        pass  # edits recorded in change_log + content_map file

    wo = inv.get("workorder") or {}
    result = {
        "case_id": inv.get("case_id") or (read_json(Path(args.case_dir) / "case.json", {}) or {}).get("case_id", "UNKNOWN"),
        "status": status,
        "editable_source": str(deck) if deck_ok else None,
        "export_artifact": str(pdf) if pdf.is_file() else None,
        "content_map": {
            "deck_title": content_map.get("deck_title") or plan.get("deck_title") or "",
            "mode": mode,
            "slide_count": len(cm_slides),
            "audience": wo.get("audience"),
            "language": wo.get("language"),
            "slides": cm_slides,
        },
        "compliance_checks": {
            "summary": f"{sum(1 for c in checks.get('checks', []) if c.get('pass'))}/"
                       f"{len(checks.get('checks', []))} passed",
            "export": {"method": export.get("method"), "pages": export.get("pages"),
                       "pages_match": export.get("pages_match"),
                       "titles_found": export.get("titles_found"),
                       "fonts": export.get("fonts")},
            "checks": checks.get("checks", []),
        },
        "exceptions": exceptions,
        # ---- honest extras beyond the contract ----
        "plan": {"file": "plan.json", "rationale": plan.get("rationale", ""),
                 "layout_roles": [s.get("role") for s in cm_slides]},
        "uncertainty": plan.get("uncertainty", []),
        "time_records": {"assembled_at": du.now_iso(),
                         "inspect_log": (work / "change_log.jsonl").is_file()},
        "change_log": "change_log.md",
        "repro": {"scripts": "skill scripts/ (inspect_case.py → build_deck.py → "
                             "check_deck.py → export_preview.py → assemble_result.py)",
                  "plan_file": "plan.json",
                  "notes": "rerun scripts in order with the same --work dir to reproduce"},
    }

    out = out_dir / "result.json"
    out_dir.mkdir(parents=True, exist_ok=True)

    # ---- duration budget (informational, non-blocking): notes speech-time vs workorder duration
    wo_dur = wo.get("duration")
    est_sec = 0.0
    for s in cm_slides:
        est_sec += max(15.0, len(s.get("notes") or "") / 4.0)  # ~4 chars/sec CJK speech
    est_min = round(est_sec / 60.0, 1)
    result["compliance_checks"]["duration_estimate"] = {
        "estimated_speech_min": est_min, "workorder_duration": wo_dur,
        "note": "informational; ~4 chars/sec, >=15s per slide"}

    # ---- QA audit report (human-reviewable summary next to the deck)
    try:
        q = ["# QA 审计报告 / QA audit report", ""]
        q.append(f"- case_id: `{result['case_id']}`  status: **{status}**  "
                 f"slides: {len(cm_slides)}  generated: {du.now_iso()}")
        q.append(f"- editable_source: `{result['editable_source']}`")
        q.append(f"- export_artifact: `{result['export_artifact']}`")
        q.append(f"- duration estimate: ~{est_min} min (workorder: {wo_dur})")
        q.append("")
        q.append("## 合规检查 / compliance checks")
        q.append("| check | pass | detail |")
        q.append("|---|---|---|")
        for c in checks.get("checks", []):
            mark = "✅" if c["pass"] else "❌"
            q.append(f"| {c['id']} | {mark} | {str(c['detail'])[:160]} |")
        q.append("")
        q.append("## 异常与处置 / exceptions")
        q.append("| code | detail | disposition |")
        q.append("|---|---|---|")
        for e in exceptions:
            q.append(f"| {e.get('code')} | {str(e.get('detail'))[:140]} | {str(e.get('disposition'))[:80]} |")
        if not exceptions:
            q.append("| (none) | | |")
        q.append("")
        q.append("## 页级证据映射 / page-level evidence")
        q.append("| slide | role | title | claim_sources |")
        q.append("|---|---|---|---|")
        for s in cm_slides:
            srcs = "; ".join(f"{x.get('ref','')}({x.get('locator','')})"
                              for x in (s.get("claim_sources") or [])) or "—"
            q.append(f"| {s.get('slide')} | {s.get('role')} | "
                     f"{str(s.get('title') or '')[:36]} | {srcs[:120]} |")
        q.append("")
        q.append("## 复现 / reproduce")
        q.append("```bash")
        q.append("python3 scripts/selftest.py   # 端到端自测（临时算例，无预置答案）")
        q.append("```")
        q.append(f"- plan rationale: {plan.get('rationale', '') or '(见 plan.json)'}")
        q.append("- change log: `change_log.md`; uncertainty: "
                 + json.dumps(result.get("uncertainty", []), ensure_ascii=False)[:200])
        (out_dir / "qa_report.md").write_text("\n".join(q) + "\n", encoding="utf-8")
        result["qa_report"] = "qa_report.md"
    except Exception as e:
        exceptions.append({"code": "QA_REPORT_FAILED", "detail": str(e)[:200],
                           "disposition": "non-blocking"})

    out.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")

    # plan + change records next to the deck
    if plan:
        (out_dir / "plan.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1),
                                           encoding="utf-8")
    build_change_log_md(work, out_dir / "change_log.md")
    try:  # narrative outline: per-slide title + talking points (讲解逻辑复核用)
        ol = [f"# 讲稿大纲 / Outline — {result['content_map'].get('deck_title', '')}", ""]
        for s in cm_slides:
            head = f"{s.get('slide')}. [{s.get('role')}] {s.get('title') or '(无标题页)'}"
            ol.append(head)
            for b in (s.get("claims") or [])[:6]:
                ol.append(f"    - {b if isinstance(b, str) else b.get('t', '')}")
            if s.get("notes"):
                ol.append(f"    讲稿: {s['notes']}")
            ol.append("")
        (out_dir / "outline.md").write_text("\n".join(ol) + "\n", encoding="utf-8")
        result["outline"] = "outline.md"
    except Exception:
        pass

    missing = [f for f in CONTRACT_FIELDS if f not in result]
    du.log_change(work, "assemble", f"status={status} exceptions={len(exceptions)}")
    print("RESULT_OK " + json.dumps({"result": str(out), "status": status,
                                     "missing_fields": missing,
                                     "exceptions": len(exceptions)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
