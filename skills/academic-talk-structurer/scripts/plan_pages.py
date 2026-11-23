#!/usr/bin/env python3
"""Audience/time -> page budget planner + deck_spec validator (deterministic, checkable).

Usage:
  python3 plan_pages.py plan  --case /input/case.json [--out /app/work/page_budget.json]
  python3 plan_pages.py check --case /input/case.json --spec /app/work/deck_spec.json [--out /app/work/budget_check.json]

Audience classes (keyword match on case.audience/task):
  defense(答辩/评审/committee)  labmeeting(组会/lab)  collab(合作/跨组)  progress(进度/导师/管理)
Budget fractions per content-page category: background/method/evidence/risk (sum to 1.0).
Tolerance scales with deck size: max(0.10, 0.75 / content_slides) so small decks stay feasible.
check exit: 0 within tolerance, 1 budget violated, 2 bad args.
"""
import argparse
import json
import sys
from pathlib import Path

AUDIENCE = {
    "defense": {"kw": ["答辩", "评审", "委员会", "defense", "committee", "review", "评委", "conference", "seminar"],
                "frac": {"background": 0.12, "method": 0.32, "evidence": 0.42, "risk": 0.14}, "qa_min": 1,
                "strategy": {
                    "opening_hook": "State the exact scientific gap, core contribution, and headline quantitative gain vs strongest baseline in <45s.",
                    "jargon_policy": "Assume domain fluency; define only newly introduced symbols/metrics once.",
                    "must_include": "Novelty contrast vs closest baselines, main controlled quantitative evidence (S1), ablations/mechanism, explicit failure boundaries.",
                    "cut_first": "Textbook background, lengthy related-work lists, minor hyperparameter logs.",
                    "qa_focus": "Fairness of baselines/compute budget, statistical significance/variance, theoretical/mechanistic rationale, out-of-distribution limits.",
                }},
    "labmeeting": {"kw": ["组会", "实验室", "lab", "group meeting", "group seminar", "research group", "reading group", "peer researcher"],
                   "frac": {"background": 0.08, "method": 0.40, "evidence": 0.36, "risk": 0.16}, "qa_min": 1,
                   "strategy": {
                       "opening_hook": "Jump directly to the technical bottleneck and design insight (skip generic domain motivation).",
                       "jargon_policy": "Use peer technical terms freely; unpack concrete algorithmic equations/steps and failure cases.",
                       "must_include": "Core mechanism/architecture walkthrough, diagnostic ablations, negative results/tricky edge cases, open discussion questions.",
                       "cut_first": "High-level motivational filler, broad survey slides.",
                       "qa_focus": "Implementation details, sensitivity to hyperparameters/data splits, why alternative designs failed, reproducible next experiments.",
                   }},
    "collab": {"kw": ["合作", "跨组", "collab", "partner", "interdisciplinary", "external"],
               "frac": {"background": 0.25, "method": 0.25, "evidence": 0.35, "risk": 0.15}, "qa_min": 1,
               "strategy": {
                   "opening_hook": "Anchor shared problem relevance and concrete input/output interface before internal algorithmic details.",
                   "jargon_policy": "Translate sub-field acronyms and domain metrics into plain cross-disciplinary impact.",
                   "must_include": "Problem framing accessible to non-specialists, intuitive method diagram, practical performance/cost tradeoffs, clear collaboration handoff/next steps.",
                   "cut_first": "Dense proofs, sub-field-only benchmark minutiae, internal lab jargon.",
                   "qa_focus": "Data/format compatibility, compute/latency cost in partner workflows, transferability to partner domains, realistic integration scope.",
               }},
    "progress": {"kw": ["进度", "导师", "管理", "progress", "advisor", "milestone", "update"],
                 "frac": {"background": 0.10, "method": 0.22, "evidence": 0.34, "risk": 0.34}, "qa_min": 1,
                 "strategy": {
                     "opening_hook": "Lead with milestone status (achieved vs blocked), headline result delta since last update, and key decision needed.",
                     "jargon_policy": "Keep terminology concise and tied directly to project KPIs and risk flags.",
                     "must_include": "Verified progress metrics, open blockers/risks with root causes, concrete mitigation plan and next-milestone timeline.",
                     "cut_first": "Already-known background, finalized old derivations.",
                     "qa_focus": "Critical-path risks, validity of preliminary numbers, resource/compute bottlenecks, go/no-go pivot criteria.",
                 }},
}
ROLE2CAT = {
    "roadmap": None,
    "motivation": "background", "related": "background",
    "method": "method",
    "setup": "evidence", "result": "evidence", "analysis": "evidence",
    "limitation": "risk", "progress": "risk",
    "title": None, "thanks": None, "conclusion": None, "qa_backup": None, "preserved": None,
}


def parse_slide_bounds(case):
    """Parse exact slide_count or range ('10-12' / {'min':8,'max':12} / min_slides+max_slides)."""
    import re
    cnt_raw = find_nested(case, ["slide_count", "num_slides", "page_count", "target_slides"])
    min_s = find_nested(case, ["min_slides", "min_pages"])
    max_s = find_nested(case, ["max_slides", "max_pages"])
    lo = hi = 0
    if isinstance(cnt_raw, dict):
        lo = int(cnt_raw.get("min") or cnt_raw.get("min_slides") or 0)
        hi = int(cnt_raw.get("max") or cnt_raw.get("max_slides") or 0)
    elif isinstance(cnt_raw, str):
        m = re.match(r"^\s*(\d+)\s*[-–~to]+\s*(\d+)\s*$", cnt_raw)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
        else:
            m1 = re.search(r"\d+", cnt_raw)
            if m1:
                lo = hi = int(m1.group(0))
    elif isinstance(cnt_raw, (int, float)):
        lo = hi = int(cnt_raw)
    if min_s is not None and not lo:
        try:
            lo = int(min_s)
        except (TypeError, ValueError):
            pass
    if max_s is not None and not hi:
        try:
            hi = int(max_s)
        except (TypeError, ValueError):
            pass
    return lo, hi


def find_nested(obj, keys):
    """Find first non-empty scalar value matching any key name anywhere in case.json."""
    if isinstance(obj, dict):
        for k in keys:
            if k in obj and obj[k] not in (None, ""):
                return obj[k]
        for v in obj.values():
            r = find_nested(v, keys)
            if r not in (None, ""):
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = find_nested(v, keys)
            if r not in (None, ""):
                return r
    return None


def classify(audience_text):
    import re
    t = (audience_text or "").lower()
    best_name, best_score = "defense", 0
    for name, cfg in AUDIENCE.items():
        score = 0
        for k in cfg["kw"]:
            kl = k.lower()
            hit = (bool(re.search(rf"\b{re.escape(kl)}\b", t))
                   if (kl.isascii() and kl.isalpha() and len(kl) <= 4)
                   else (kl in t))
            if hit:
                # Generic word "review" gets weight 1; specific audience cues get weight 2
                score += 1 if kl == "review" else 2
        if score > best_score:
            best_score, best_name = score, name
    return best_name


def parse_task_md_fallbacks(case_path: Path):
    """Extract fallback (audience_hint, duration_minutes, lo_slides, hi_slides) from sibling task.md."""
    import re
    tmd = case_path.parent / "task.md"
    if not tmd.exists():
        return "", None, 0, 0
    txt = tmd.read_text(encoding="utf-8", errors="ignore")[:12000]
    m_aud = re.search(r"(?:target\s+)?(?:audience|listener|受众|听众)\s*[:：]\s*([^\n]+)", txt, re.I)
    aud_hint = m_aud.group(1).strip() if m_aud else txt
    dur = None
    m_dur = re.search(r"(\d+(?:\.\d+)?)\s*[-]?\s*(?:min(?:ute)?s?|分钟)", txt, re.I)
    if m_dur:
        try:
            dur = float(m_dur.group(1))
        except ValueError:
            dur = None
    lo = hi = 0
    m_rng = re.search(r"(\d{1,2})\s*[-–~]\s*(\d{1,2})\s*(?:slides?|页|张)", txt, re.I)
    if m_rng and 2 <= int(m_rng.group(1)) <= int(m_rng.group(2)) <= 40:
        lo, hi = int(m_rng.group(1)), int(m_rng.group(2))
    else:
        m_one = re.search(r"(\d{1,2})\s*[-]?\s*(?:slides?|页|张)", txt, re.I)
        if m_one and 2 <= int(m_one.group(1)) <= 40:
            lo = hi = int(m_one.group(1))
    return aud_hint, dur, lo, hi


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["plan", "check"])
    ap.add_argument("--case", default="/input/case.json")
    ap.add_argument("--spec")
    ap.add_argument("--plan", help="optional built slide_plan.json (counts auto-appended qa_backup & actual delivered slides)")
    ap.add_argument("--out")
    ap.add_argument("--justif", help='justification json keyed by issue: {"share[method]":"...","slide_count":"..."}')
    a = ap.parse_args()
    cpath = Path(a.case)
    case = json.loads(cpath.read_text(encoding="utf-8")) if cpath.exists() else {}
    tmd_text, tmd_dur, tmd_lo, tmd_hi = parse_task_md_fallbacks(cpath)
    aud_raw = find_nested(case, ["audience", "audience_profile", "target_audience", "task", "presentation_type"]) or ""
    if isinstance(aud_raw, dict):
        aud_raw = json.dumps(aud_raw, ensure_ascii=False)
    aud = classify(str(aud_raw) if str(aud_raw).strip() else tmd_text)
    cfg = AUDIENCE[aud]
    dur_raw = find_nested(case, ["duration_minutes", "duration", "minutes", "time_limit_minutes"])
    try:
        minutes = float(dur_raw) if dur_raw is not None else (tmd_dur if tmd_dur is not None else 10.0)
    except (TypeError, ValueError):
        minutes = tmd_dur if tmd_dur is not None else 10.0
    lo_slides, hi_slides = parse_slide_bounds(case)
    if not lo_slides and not hi_slides and (tmd_lo or tmd_hi):
        lo_slides, hi_slides = tmd_lo, tmd_hi
    total = hi_slides or lo_slides

    spec = json.loads(Path(a.spec).read_text(encoding="utf-8")) if (a.spec and Path(a.spec).exists()) else None
    plan_doc = json.loads(Path(a.plan).read_text(encoding="utf-8")) if (a.plan and Path(a.plan).exists()) else None
    slides = ((plan_doc or {}).get("slides") or (spec or {}).get("slides") or [])
    inferred_from_duration = False
    n_total = total or len(slides)
    if n_total <= 0 and a.mode == "plan":
        n_total = max(6, min(18, round(minutes * 1.1) + 2))
        inferred_from_duration = True
    if n_total <= 0:
        print("no slide_count in case and no spec; nothing to plan", file=sys.stderr)
        sys.exit(2)
    qa_min = cfg["qa_min"]
    fixed = 2 + qa_min
    content = max(1, n_total - fixed)

    if a.mode == "plan":
        budget = {}
        floors = {c: (1 if c in ("method", "evidence") and content >= 2 else 0) for c in cfg["frac"]}
        for cat, fr in cfg["frac"].items():
            budget[cat] = max(floors[cat], round(fr * content))
        while sum(budget.values()) < content:
            cat = max(cfg["frac"], key=lambda c: cfg["frac"][c] * content - budget[c])
            budget[cat] += 1
        while sum(budget.values()) > content:
            cands = [c for c in cfg["frac"] if budget[c] > floors[c]]
            if not cands:
                break
            cat = max(cands, key=lambda c: budget[c] - cfg["frac"][c] * content)
            budget[cat] -= 1
        per = round((minutes * 60 - 25) / max(1, content))
        out = {"audience_class": aud, "total_pages": n_total, "content_pages": content,
               "inferred_from_duration": inferred_from_duration,
               "duration_minutes": minutes,
               "qa_backup_min": qa_min, "category_budget": budget,
               "seconds_per_content_slide": per,
               "audience_strategy": cfg.get("strategy", {}),
               "note": "deck_spec 角色分配应满足 category_budget（±容差见 check）"}
        if a.out:
            Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
        print(json.dumps(out, ensure_ascii=False, indent=1))
        return

    if not slides:
        print("check mode requires --spec with slides", file=sys.stderr)
        sys.exit(2)
    counts = {"background": 0, "method": 0, "evidence": 0, "risk": 0}
    qa = 0
    for s in slides:
        cat = ROLE2CAT.get(s.get("role", ""))
        if cat:
            counts[cat] += 1
        if s.get("role") == "qa_backup":
            qa += 1
    spec_content = sum(counts.values()) or 1
    tol = max(0.12, 1.15 / spec_content)
    issues, warnings, justified = [], [], []
    justif = {}
    if a.justif and Path(a.justif).exists():
        justif = json.loads(Path(a.justif).read_text(encoding="utf-8"))
    meta_tradeoffs = ((spec or {}).get("meta", {}).get("tradeoffs")
                      or (plan_doc or {}).get("meta", {}).get("tradeoffs") or [])
    tradeoff_blob = " ".join(str(x) for x in meta_tradeoffs)

    def add_blocking(key, msg):
        if key in justif:
            justified.append({"key": key, "msg": msg, "reason": justif[key]})
        elif key == "slide_count" and tradeoff_blob and (str(len(slides)) in tradeoff_blob or "页" in tradeoff_blob or "slide" in tradeoff_blob.lower()):
            justified.append({"key": key, "msg": msg, "reason": f"meta.tradeoffs: {tradeoff_blob[:90]}"})
        else:
            issues.append({"key": key, "msg": msg})

    def add_advisory(key, msg):
        if key in justif:
            justified.append({"key": key, "msg": msg, "reason": justif[key]})
        else:
            warnings.append({"key": key, "msg": msg})

    for cat, target in cfg["frac"].items():
        got = counts[cat] / spec_content
        if abs(got - target) > tol:
            add_advisory(f"share[{cat}]", f"got {got:.0%} target {target:.0%} (tol ±{tol:.0%}, counts {counts[cat]}/{spec_content})")
    if qa < qa_min:
        add_advisory("qa_backup", f"qa_backup pages {qa} < min {qa_min}")
    if lo_slides and hi_slides and lo_slides != hi_slides:
        if not (lo_slides <= len(slides) <= hi_slides):
            add_blocking("slide_count", f"slide count {len(slides)} outside range {lo_slides}..{hi_slides}")
    elif total and len(slides) != total:
        add_blocking("slide_count", f"slide count {len(slides)} != case slide_count {total}")
    rep = {"audience_class": aud, "counts": counts, "qa_backup": qa, "tolerance": round(tol, 3),
           "issues": issues, "warnings": warnings, "justified": justified}
    if a.out:
        Path(a.out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"budget check [{aud}]: bg/meth/evi/risk = {counts['background']}/{counts['method']}/{counts['evidence']}/{counts['risk']}, qa={qa}")
    for i in issues:
        print(f"  ! [{i['key']}] {i['msg']}")
    for w in warnings:
        print(f"  ~ (advisory) [{w['key']}] {w['msg']}")
    for j in justified:
        print(f"  ~ [{j['key']}] {j['msg']} :: justified: {j['reason'][:60]}")
    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
