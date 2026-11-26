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

FIG_RE = re.compile(r"(?:Fig(?:ure)?s?\.?\s*(S?\d+)[a-z]?(?:\s*(?:[-–—~&]|and)\s*(S?\d+)[a-z]?)?|图\s*(\d+)(?:\s*[-–—~、和与]\s*(\d+))?)", re.IGNORECASE)
TAB_RE = re.compile(r"(?:Tab(?:le)?s?\.?\s*(S?\d+|[IVX]+)\b(?:\s*(?:[-–—~&]|and)\s*(S?\d+|[IVX]+)\b)?|表\s*(\d+)(?:\s*[-–—~、和与]\s*(\d+))?)", re.IGNORECASE)
SEC_RE = re.compile(r"(?:§\s*|Sec(?:tion)?s?\.?\s*|第\s*)(\d+(?:\.\d+){0,2}|[IVX]{1,4}(?:-[A-Z])?\b)", re.IGNORECASE)
APP_RE = re.compile(r"(?:App(?:endix)?\.?\s*|附录\s*)([A-Z](?:\.\d+)?)\b", re.IGNORECASE)
EQ_RE = re.compile(r"(?:Eq(?:uation)?\.?\s*\(?(\d+)\)?|公式\s*\(?(\d+)\)?)", re.IGNORECASE)
ALG_RE = re.compile(r"(?:Alg(?:orithm)?\.?\s*(\d+)\b|算法\s*(\d+))", re.IGNORECASE)
SLIDE_RE = re.compile(r"(?:(?:Orig(?:inal)?\s+)?Slide\s*(\d+)\b|原\s*[Pp]\s*(\d+))", re.IGNORECASE)
PAGE_RE = re.compile(r"(?:^|[\s,;(（/])pp?\.?\s*(\d+)(?:\s*[-–—~]\s*(\d+))?\b", re.IGNORECASE)
CONTENT_ROLES = {"motivation", "related", "method", "setup", "result", "analysis", "limitation", "progress"}

_ARABIC_TO_ROMAN = {
    "1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI",
    "7": "VII", "8": "VIII", "9": "IX", "10": "X", "11": "XI", "12": "XII",
}
_ROMAN_TO_ARABIC = {v: k for k, v in _ARABIC_TO_ROMAN.items()}


def _roman_arabic_alts(num: str):
    """Return equivalent Roman/Arabic variants for Table/Section matching (e.g., 2 <-> II)."""
    u = str(num or "").strip().upper()
    alts = [num]
    if u in _ARABIC_TO_ROMAN:
        alts.append(_ARABIC_TO_ROMAN[u])
    elif u in _ROMAN_TO_ARABIC:
        alts.append(_ROMAN_TO_ARABIC[u])
    return alts


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
                for g_v in (mf_fig.group(1), mf_fig.group(3)):
                    if g_v:
                        manifest_labels.add(("figure", g_v.upper()))
            elif mf_tab:
                for g_v in (mf_tab.group(1), mf_tab.group(3)):
                    if g_v:
                        for alt_v in _roman_arabic_alts(g_v):
                            manifest_labels.add(("table", alt_v.upper()))
            if m.get("page"):
                manifest_pages.add(int(m["page"]))
    if a.tables and Path(a.tables).exists():
        try:
            tbls = json.loads(Path(a.tables).read_text(encoding="utf-8"))
            for tb in (tbls if isinstance(tbls, list) else []):
                lbl = str(tb.get("label") or "")
                tb_m = TAB_RE.search(lbl)
                if tb_m:
                    for g_v in (tb_m.group(1), tb_m.group(3)):
                        if g_v:
                            for alt_v in _roman_arabic_alts(g_v):
                                manifest_labels.add(("table", alt_v.upper()))
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
                    nums = [g for g in mm.groups() if g]
                    for num in nums:
                        checked += 1
                        ev_hits += 1
                        alts = _roman_arabic_alts(num) if kind == "table" else [num]
                        matched = False
                        for cand_n in alts:
                            pat = {
                                "figure": r"(?:Fig(?:ure)?\.?\s*NUM(?![0-9])|图\s*NUM(?![0-9]))",
                                "table": r"(?:Tab(?:le)?\.?\s*NUM(?![0-9A-Za-z])|表\s*NUM(?![0-9]))",
                            }[kind].replace("NUM", re.escape(cand_n))
                            if re.search(pat, text, re.IGNORECASE) or ((kind, cand_n.upper()) in manifest_labels):
                                matched = True
                                break
                        if not matched:
                            dangling.append({"slide": i, "locator": loc, "issue": f"{kind} {num} not found in source"})
            for sm in SEC_RE.finditer(loc):
                num = sm.group(1)
                checked += 1
                ev_hits += 1
                base_num = num.split("-", 1)[0].split(".", 1)[0]
                sec_alts = [num]
                if base_num != num:
                    sec_alts.append(base_num)
                for ra in _roman_arabic_alts(base_num):
                    if ra not in sec_alts:
                        sec_alts.append(ra)
                sec_ok = False
                for sn_cand in sec_alts:
                    if re.search(rf"(?m)^\s*{re.escape(sn_cand)}(?=[\s\.:A-Za-z\u4e00-\u9fff])", text):
                        sec_ok = True
                        break
                    # For pure Roman numerals like "I", avoid matching English pronoun "I" mid-sentence; require section-like boundary
                    if re.fullmatch(r"[IVX]{1,4}", sn_cand, re.I):
                        if re.search(rf"(?m)(?:^|\n)\s*{re.escape(sn_cand)}[\.\s]+[A-Z\u4e00-\u9fff]|(?:Sec(?:tion)?\.?|§)\s*{re.escape(sn_cand)}\b", text):
                            sec_ok = True
                            break
                    elif re.search(rf"\b{re.escape(sn_cand)}\b", text):
                        sec_ok = True
                        break
                if not sec_ok:
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
                for pg_s in (pm.group(1), pm.group(2)):
                    if not pg_s:
                        continue
                    num = int(pg_s)
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
