#!/usr/bin/env python3
"""Evidence locator validation: every cited Fig/Table/section/page must exist in the source.

Usage:
  python3 check_evidence.py --spec /app/work/deck_spec.json --source /app/work/paper.txt \
      [--manifest /app/work/figs/manifest.json] [--pages 3] [--out /app/work/evidence_check.json]

Checks per evidence entry: figure/table labels appear in source text (or figs manifest),
section numbers appear in source, page numbers within PDF page count. Exit: 0 ok, 1 dangling, 3 no source.
No-script fallback: grep each cited label (e.g. "Figure 2") in paper.txt manually.
"""
import argparse
import json
import re
import sys
from pathlib import Path

FIG_RE = re.compile(r"(?:Fig(?:ure)?s?\.?\s*(S?\d+)[a-z]?|图\s*(\d+))", re.IGNORECASE)
TAB_RE = re.compile(r"(?:Tab(?:le)?s?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", re.IGNORECASE)
SEC_RE = re.compile(r"(?:§\s*|Sec(?:tion)?s?\.?\s*|第\s*)(\d+(?:\.\d+){0,2}|[IVX]{1,4}(?:-[A-Z])?\b)", re.IGNORECASE)
APP_RE = re.compile(r"(?:App(?:endix)?\.?\s*|附录\s*)([A-Z](?:\.\d+)?)\b", re.IGNORECASE)
EQ_RE = re.compile(r"(?:Eq(?:uation)?\.?\s*\(?(\d+)\)?|公式\s*\(?(\d+)\)?)", re.IGNORECASE)
ALG_RE = re.compile(r"(?:Alg(?:orithm)?\.?\s*(\d+)\b|算法\s*(\d+))", re.IGNORECASE)
SLIDE_RE = re.compile(r"(?:(?:Orig(?:inal)?\s+)?Slide\s*(\d+)\b|原\s*[Pp]\s*(\d+))", re.IGNORECASE)
PAGE_RE = re.compile(r"(?:^|[\s,;(（/])p\.?\s*(\d+)\b", re.IGNORECASE)
CONTENT_ROLES = {"motivation", "related", "method", "setup", "result", "analysis", "limitation", "progress"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--source", required=True)
    ap.add_argument("--manifest")
    ap.add_argument("--tables")
    ap.add_argument("--pages", type=int)
    ap.add_argument("--out")
    a = ap.parse_args()
    src = Path(a.source)
    if not src.exists():
        print(f"FATAL: source text missing: {src}", file=sys.stderr)
        sys.exit(3)
    text = src.read_text(encoding="utf-8", errors="ignore")
    manifest_labels, manifest_pages = set(), set()
    if a.manifest and Path(a.manifest).exists():
        mf = json.loads(Path(a.manifest).read_text(encoding="utf-8"))
        for m in mf:
            lbl = (m.get("label") or "")
            mf_fig = FIG_RE.search(lbl)
            mf_tab = TAB_RE.search(lbl)
            if mf_fig:
                manifest_labels.add(("figure", (mf_fig.group(1) or mf_fig.group(2)).upper()))
            elif mf_tab:
                manifest_labels.add(("table", (mf_tab.group(1) or mf_tab.group(2)).upper()))
            if m.get("page"):
                manifest_pages.add(int(m["page"]))
    if a.tables and Path(a.tables).exists():
        try:
            tbls = json.loads(Path(a.tables).read_text(encoding="utf-8"))
            for tb in (tbls if isinstance(tbls, list) else []):
                lbl = str(tb.get("label") or "")
                tb_m = TAB_RE.search(lbl)
                if tb_m:
                    manifest_labels.add(("table", (tb_m.group(1) or tb_m.group(2)).upper()))
                if tb.get("page"):
                    manifest_pages.add(int(tb["page"]))
        except Exception:
            pass
    page_markers = [int(x) for x in re.findall(r"--- PAGE (\d+)(?: \([^)]+\))? ---", text)]
    deck_markers = [int(x) for x in re.findall(r"--- DECK SLIDE (\d+) ---", text)]
    pages = max(a.pages or 0, max(page_markers) if page_markers else 0, max(manifest_pages) if manifest_pages else 0) or 99
    deck_slides_max = max(deck_markers) if deck_markers else 0

    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    dangling, warnings, checked = [], [], 0
    for i, s in enumerate(spec.get("slides", []), 1):
        evs = s.get("evidence") or []
        if isinstance(evs, str):
            evs = [{"locator": evs}]
        role = str(s.get("role", "")).lower()
        if role in CONTENT_ROLES and not any(str(e.get("locator", "") if isinstance(e, dict) else e).strip() for e in evs):
            warnings.append(f"s{i} ({role}): no evidence locator bound")
        for ev in evs:
            loc = str(ev.get("locator", "") if isinstance(ev, dict) else ev)
            etype = str(ev.get("type", "") if isinstance(ev, dict) else "")
            ev_hits = 0
            for rex, kind in ((FIG_RE, "figure"), (TAB_RE, "table")):
                for mm in rex.finditer(loc):
                    num = mm.group(1) or mm.group(2)
                    checked += 1
                    ev_hits += 1
                    pat = {
                        "figure": r"(?:Fig(?:ure)?\.?\s*NUM(?![0-9])|图\s*NUM(?![0-9]))",
                        "table": r"(?:Tab(?:le)?\.?\s*NUM(?![0-9])|表\s*NUM(?![0-9]))",
                    }[kind].replace("NUM", re.escape(num))
                    if not (re.search(pat, text, re.IGNORECASE) or ((kind, num.upper()) in manifest_labels)):
                        dangling.append({"slide": i, "locator": loc, "issue": f"{kind} {num} not found in source"})
            for sm in SEC_RE.finditer(loc):
                num = sm.group(1)
                checked += 1
                ev_hits += 1
                base_num = num.split("-", 1)[0]
                if (
                    not re.search(rf"(?m)^\s*{re.escape(num)}(?=[\s\.:A-Za-z\u4e00-\u9fff])", text)
                    and not re.search(rf"\b{re.escape(num)}\b", text)
                    and not (base_num != num and re.search(rf"\b{re.escape(base_num)}\b", text))
                ):
                    dangling.append({"slide": i, "locator": loc, "issue": f"section {num} not found in source"})
            for am in APP_RE.finditer(loc):
                ap_id = am.group(1)
                checked += 1
                ev_hits += 1
                if not re.search(rf"(?:App(?:endix)?\.?\s*|Supp(?:lementary)?\.?\s*|附录\s*|补充材料\s*){re.escape(ap_id)}\b", text, re.IGNORECASE):
                    warnings.append(f"s{i}: appendix {ap_id} not matched verbatim in source ({loc!r}) — verify label")
            for em in EQ_RE.finditer(loc):
                eq_n = em.group(1) or em.group(2)
                checked += 1
                ev_hits += 1
                if not re.search(rf"(?:Eq(?:uation)?\.?\s*\(?{re.escape(eq_n)}\)?|公式\s*\(?{re.escape(eq_n)}\)?|\(\s*{re.escape(eq_n)}\s*\))", text, re.IGNORECASE):
                    warnings.append(f"s{i}: equation ({eq_n}) not matched verbatim in extracted text ({loc!r}) — verify math numbering")
            for alm in ALG_RE.finditer(loc):
                al_n = alm.group(1) or alm.group(2)
                checked += 1
                ev_hits += 1
                if not re.search(rf"(?:Alg(?:orithm)?\.?\s*{re.escape(al_n)}\b|算法\s*{re.escape(al_n)})", text, re.IGNORECASE):
                    warnings.append(f"s{i}: algorithm {al_n} not matched verbatim in extracted text ({loc!r}) — verify float numbering")
            for slm in SLIDE_RE.finditer(loc):
                sn = int(slm.group(1) or slm.group(2))
                checked += 1
                ev_hits += 1
                if deck_slides_max and (sn < 1 or sn > deck_slides_max):
                    dangling.append({"slide": i, "locator": loc, "issue": f"starting-deck slide {sn} outside 1..{deck_slides_max}"})
            for pm in PAGE_RE.finditer(loc):
                num = int(pm.group(1))
                checked += 1
                ev_hits += 1
                if num < 1 or num > pages:
                    dangling.append({"slide": i, "locator": loc, "issue": f"page {num} outside 1..{pages}"})
            if loc.strip() and ev_hits == 0 and role in CONTENT_ROLES:
                if not re.search(r"(?:Abstract|Introduction|Method|Result|Discussion|Limitation|Conclusion|摘要|引言|方法|实验|讨论|局限|结论|\.(?:csv|tsv|md|txt|json|docx|xlsx|tex)\b)", loc, re.I):
                    warnings.append(f"s{i}: vague evidence locator {loc!r} (add specific Fig/Table/§/Eq/p.N/Slide N)")
            if etype in ("equation",) or not loc.strip():
                continue
    rep = {"checked": checked, "n_dangling": len(dangling), "dangling": dangling[:50], "warnings": warnings[:20]}
    if a.out:
        Path(a.out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"evidence check: {checked} locators checked, {len(dangling)} dangling" + (f", {len(warnings)} advisory" if warnings else ""))
    for d in dangling:
        print(f"  [s{d['slide']}] {d['locator']} :: {d['issue']}")
    for w in warnings[:6]:
        print(f"  ~ (advisory) {w}")
    sys.exit(1 if dangling else 0)


if __name__ == "__main__":
    main()
