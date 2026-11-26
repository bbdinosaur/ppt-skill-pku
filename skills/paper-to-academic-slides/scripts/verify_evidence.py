#!/usr/bin/env python3
"""Cross-verify slides.json evidence against paper_context.json.

Usage:
  python3 verify_evidence.py slides.json --ctx /app/work/paper/paper_context.json \
      [--report /app/work/evidence_report.json]

Checks (error = must fix; warn = quality visibility):
  E1 slide_count_exceeds_limit      slides > meta.max_slides
  E2 ref_not_found                  figure/table/algorithm id absent from paper captions
  E3 quote_not_found                quote text not found (NFKC/ligature-normalized) in paper
  E4 page_out_of_range              cited page > paper n_pages or < 1
  W1 quote_page_mismatch            quote found on a different page than stated
  W2 number_not_on_cited_page       key numeric token in bullets/table absent from cited pages
  W3 notes_duplicate_slide_text     notes mostly repeat the slide bullets
  W4 no_evidence                    slide (except cover) lacks paper_refs
  W5 missing_notes                  slide notes shorter than 20 chars
  W6 generic_section_title          title is a bare section label instead of a claim
  W7 missing_limitations_slide      paper has limitations section but deck omits it
Exit code 0 always; severity inside report.
"""
import argparse
import json
import os
import re
import sys
import unicodedata

LIGATURES = {
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl",
    "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st",
    "\u2212": "-", "\u2013": "-", "\u2014": "-", "\u00a0": " ",
}
KEY_NUM_RE = re.compile(r"(?<![A-Za-z])(?:\d+\.\d+|\d{2,})")
REF_PARSE_RE = re.compile(
    r"(?:extended\s+data\s+|supplementary\s+)?"
    r"(fig(?:ure)?s?|tab(?:le)?s?|图|表|alg(?:orithm)?s?|listings?|schemes?|eq(?:uation)?s?|式)"
    r"\s*\.?\s*\(?((?:[sa]?\d+[a-z]?)|(?:[IVX]{1,4}\b))\)?",
    re.IGNORECASE,
)
GENERIC_TITLES = {
    "introduction", "background", "motivation", "method", "methods", "methodology",
    "method overview", "methodology overview", "proposed method", "our method",
    "approach", "model", "framework", "architecture", "experiments", "experimental setup",
    "experimental results", "results", "main results", "evaluation", "ablation",
    "ablation study", "analysis", "discussion", "limitations", "limitations and future work",
    "conclusion", "conclusions", "summary", "references", "related work", "overview",
    "引言", "背景", "研究背景", "动机", "问题", "问题定义", "方法", "方法概述", "方法介绍",
    "核心方法", "模型", "实验", "实验设置", "实验设计", "实验结果", "实验结果与分析",
    "结果", "主结果", "主要结果", "消融", "消融实验", "讨论", "局限", "局限性",
    "结论", "总结", "总结与展望",
}
ID_NUM_STRIP_RE = re.compile(
    r"(?:§|sec(?:tion)?\.?|fig(?:ure)?\.?|tab(?:le)?\.?|eq(?:uation)?\.?|alg(?:orithm)?\.?|pp?\.|图|表|式|第)\s*\d+(?:[.\-–]\d+)?\s*页?",
    re.I,
)
YEAR_RE = re.compile(r"^(?:19\d\d|20[0-3]\d)$")
LIMIT_WORDS = re.compile(r"(?i)limitation|drawback|shortcoming|future\s+work|caveat|局限|不足|限制|边界")


COMPARE_WORDS_RE = re.compile(
    r"(?i)\b(?:outperforms?|beats?|exceeds?|surpasses?|vs\.?|versus|compared\s+to|over\s+baseline)\b|优于|超过|高于|对比基线"
)
ATTRIB_HINT_RE = re.compile(
    r"(?:\([^)]{2,30}\)|et\s+al\.|(?i:\b(?:full[\s\-]\w+|standard|vanilla|w/o|without|ablation|unpruned|dense|bert|gpt|llama|resnet|vit|clip|transformer|diffusion|h2o|snapkv|equiformer|nequip)\b)|[A-Z][A-Za-z]*-[A-Za-z0-9]+|[A-Z]{2,}[0-9]+|[A-Z][a-z]+[A-Z][A-Za-z0-9\-]*)",
)


def _parse_page_int(val):
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return int(val)
    m = re.search(r"\d+", str(val))
    return int(m.group(0)) if m else None


def _strip_num_commas(s):
    return re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", s or "")


def clean_str(s):
    if not s:
        return ""
    for k, v in LIGATURES.items():
        if k in s:
            s = s.replace(k, v)
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"([A-Za-z])-\s*\n\s*([a-z])", r"\1\2", s)
    return s


def norm(s):
    return re.sub(r"\s+", "", clean_str(s))


def alnum(s):
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", clean_str(s).lower())


def canonical_ref_key(rtype, rid):
    """Map 'Fig. 3', 'Figure 3', '图3', 'Fig. 3a' -> ('fig:3a', 'fig:3')."""
    s = clean_str(str(rid or "")).strip()
    m = REF_PARSE_RE.search(s)
    if m:
        k = m.group(1).lower().rstrip(".")
        cid = m.group(2).lower()
    else:
        k = (rtype or "").lower()
        dm = re.search(r"([sa]?\d+[a-z]?)", s, re.IGNORECASE)
        if not dm:
            return alnum(s), None
        cid = dm.group(1).lower()

    if k in ("fig", "figs", "figure", "figures", "图", "scheme", "schemes"):
        grp = "fig" if not k.startswith("scheme") else "scheme"
    elif k in ("tab", "tabs", "table", "tables", "表"):
        grp = "tab"
    elif k in ("alg", "algs", "algorithm", "algorithms"):
        grp = "alg"
    elif k in ("listing", "listings"):
        grp = "listing"
    elif k in ("eq", "eqs", "equation", "equations", "式"):
        grp = "eq"
    else:
        grp = k
    full = f"{grp}:{cid}"
    base_cid = re.sub(r"[a-z]+$", "", cid)
    base = f"{grp}:{base_cid}" if base_cid and base_cid != cid else None
    return full, base


def char_bigrams(s):
    s = norm(s).lower()
    return {s[i:i + 2] for i in range(len(s) - 1) if len(s[i:i + 2]) == 2}


def _fuzzy_window_quote_match(quote_raw, text_raw, text_alnum):
    """PDF-artifact-tolerant fallback when exact substring fails (e.g., superscript footnote
    digit or broken math glyph): requires 100% of numeric tokens AND >=85% of >=3-char
    tokens (at least 3 tokens) to co-occur inside a compact text window."""
    q_clean = clean_str(quote_raw)
    t_clean = clean_str(text_raw)
    if not q_clean or not t_clean:
        return False
    q_nums = re.findall(r"\d+(?:\.\d+)?", _strip_num_commas(q_clean))
    t_num_src = _strip_num_commas(t_clean)
    for num in q_nums:
        if num not in t_num_src:
            return False
    # Tokenize English words (>=3 chars) and CJK 2-char chunks
    en_toks = [w.lower() for w in re.findall(r"[A-Za-z0-9]{3,}", q_clean)]
    cjk_runs = re.findall(r"[\u4e00-\u9fff]{2,}", q_clean)
    cjk_toks = []
    for run in cjk_runs:
        for i in range(0, len(run) - 1, 2):
            cjk_toks.append(run[i:i + 2])
    num_toks = [re.sub(r"[^0-9]", "", n) for n in q_nums if len(re.sub(r"[^0-9]", "", n)) >= 2]
    toks = list(dict.fromkeys(en_toks + cjk_toks + num_toks))
    if len(toks) < 3:
        return False
    # Locate token positions inside text_alnum
    hits = []
    for tk in toks:
        pos = text_alnum.find(tk)
        if pos >= 0:
            hits.append(pos)
    if len(hits) / float(len(toks)) < 0.85:
        return False
    hits.sort()
    win_lim = max(160, int(len(alnum(q_clean)) * 2.4) + 120)
    # Check if >=85% of tokens fall inside one compact window
    need = int((len(toks) * 0.85) + 0.999)
    for i in range(len(hits)):
        j = i + need - 1
        if j < len(hits) and (hits[j] - hits[i]) <= win_lim:
            return True
    return False


SCHEMATIC_CAP_RE = re.compile(
    r"\b(?:overview|architecture|framework|pipeline|workflow|schematic|flowchart|topology)\b|(?:框架|流程|架构|原理图|总览图)",
    re.IGNORECASE,
)
SCHEMATIC_LABEL_RE = re.compile(
    r"\b(?:schematic|illustrative|conceptual|diagram|non[- ]?experimental|not\s+experimental|architecture|pipeline|workflow|framework)\b"
    r"|(?:示意|非实验|概念图|架构图|流程图|原理图|框架图)",
    re.IGNORECASE,
)
UNIT_PAIR_RE = re.compile(
    r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)\s*"
    r"(%|ms|msec|μs|us|ns|secs?|s(?![A-Za-z0-9])|mins?|hrs?|h(?![A-Za-z0-9])|"
    r"gb|mb|kb|tb|gib|mib|kib|ev|mev|kev|kcal/mol|kj/mol|"
    r"db|khz|mhz|ghz|hz|nm|μm|um|mm|cm|km|mw|kw|w(?![A-Za-z0-9])|"
    r"gflops|tflops|mflops|flops|tok/s|fps|mev/a|mev/å)",
    re.IGNORECASE,
)
UNIT_ALIAS_MAP = {
    "msec": "ms", "sec": "s", "secs": "s", "mins": "min", "hrs": "hr", "h": "hr",
    "μs": "us", "μm": "um", "gib": "gb", "mib": "mb", "kib": "kb", "mev/å": "mev/a",
}
STRICT_SCI_UNITS = {
    "%", "ms", "us", "ns", "gb", "mb", "kb", "tb", "ev", "mev", "kev",
    "kcal/mol", "kj/mol", "db", "hz", "khz", "mhz", "ghz", "nm", "um",
    "mm", "cm", "km", "gflops", "tflops", "mflops", "flops", "tok/s", "fps", "mev/a",
}
UNIT_DIM_MAP = {
    "ns": "time", "us": "time", "ms": "time", "s": "time", "min": "time", "hr": "time",
    "kb": "mem", "mb": "mem", "gb": "mem", "tb": "mem",
    "mev": "energy", "ev": "energy", "kev": "energy", "kcal/mol": "energy", "kj/mol": "energy",
    "nm": "len", "um": "len", "mm": "len", "cm": "len", "km": "len",
    "hz": "freq", "khz": "freq", "mhz": "freq", "ghz": "freq",
    "mw": "power", "w": "power", "kw": "power",
    "flops": "compute", "mflops": "compute", "gflops": "compute", "tflops": "compute",
}
COVER_META_NUM_RE = re.compile(
    r"(?:\b20\d{2}[-/.]\d{1,2}(?:[-/.]\d{1,2})?\b|\d{4}年\d{1,2}月(?:\d{1,2}日)?|"
    r"\b\d{1,2}\s*[-–]?\s*(?:min(?:ute)?s?|slides?|pages?)\b|\d{1,2}\s*(?:分钟|页|张))",
    re.IGNORECASE,
)
PACE_TAG_RE = re.compile(
    r"(?:\[\s*~?(\d+(?:\.\d+)?)\s*(?:min(?:ute)?s?|m|分钟)\s*\]|~(\d+(?:\.\d+)?)\s*(?:min|分钟))",
    re.IGNORECASE,
)


HDR_UNIT_RE = re.compile(
    r"[\(\[（【/]\s*"
    r"(%|ms|msec|μs|us|ns|secs?|s|mins?|hrs?|h|"
    r"gb|mb|kb|tb|gib|mib|kib|ev|mev|kev|kcal/mol|kj/mol|"
    r"db|khz|mhz|ghz|hz|nm|μm|um|mm|cm|km|mw|kw|w|"
    r"gflops|tflops|mflops|flops|tok/s|fps|mev/a|mev/å)"
    r"\s*[\)\]）】]?",
    re.IGNORECASE,
)


def _canon_unit(raw_u):
    u = (raw_u or "").strip().lower()
    return UNIT_ALIAS_MAP.get(u, u)


def _extract_header_unit(hdr):
    """Extract canonical unit declared in a table column header like 'Latency (ms)' or 'Acc (%)'."""
    if not hdr:
        return None
    m = HDR_UNIT_RE.search(clean_str(str(hdr)))
    if not m:
        return None
    return _canon_unit(m.group(1)), m.group(1)


def _extract_num_units(text):
    out = []
    if not text:
        return out
    cleaned = _strip_num_commas(ID_NUM_STRIP_RE.sub(" ", clean_str(text)))
    for m in UNIT_PAIR_RE.finditer(cleaned):
        num = m.group(1)
        raw_u = m.group(2)
        cu = _canon_unit(raw_u)
        out.append((num, cu, raw_u))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("slides_json")
    ap.add_argument("--ctx", default="/app/work/paper/paper_context.json")
    ap.add_argument("--case-resolved", default=None)
    ap.add_argument("--report", default="/app/work/evidence_report.json")
    a = ap.parse_args()

    with open(a.slides_json, encoding="utf-8") as f:
        spec = json.load(f)
    try:
        with open(a.ctx, encoding="utf-8") as f:
            ctx = json.load(f)
    except Exception as e:
        report = {"ok": False, "n_error": 1, "n_warn": 0,
                  "issues": [{"sev": "error", "type": "ctx_missing", "detail": str(e)}]}
        os.makedirs(os.path.dirname(os.path.abspath(a.report)) or ".", exist_ok=True)
        with open(a.report, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=1)
        print(json.dumps(report, ensure_ascii=False))
        sys.exit(0)

    slides = spec.get("slides", [])
    meta = spec.get("meta", {})
    page_text = {p["page"]: clean_str(p.get("text", "")) for p in ctx.get("pages", [])}
    page_alnum = {p: alnum(t) for p, t in page_text.items()}
    page_norm = {p: norm(t) for p, t in page_text.items()}
    # Include extracted table rows in page_alnum/page_norm so verbatim table quotes match cleanly
    for tbl in ctx.get("tables", []):
        tp = tbl.get("page")
        if tp:
            tblob = " ".join(tbl.get("headers") or []) + " " + " ".join(
                " ".join(str(c) for c in r) for r in (tbl.get("rows") or []))
            page_alnum[tp] = page_alnum.get(tp, "") + alnum(tblob)
            page_norm[tp] = page_norm.get(tp, "") + norm(tblob)
    n_pages = ctx.get("n_pages", 0)

    # Load supplemental text/csv/json materials if provided via --case-resolved
    extra_text_blob = ""
    cr_data = {}
    if a.case_resolved and os.path.isfile(a.case_resolved):
        try:
            with open(a.case_resolved, encoding="utf-8") as cf:
                cr = json.load(cf)
            cr_data = cr if isinstance(cr, dict) else {}
            main_src = os.path.abspath(cr.get("paper_source") or cr.get("paper_pdf") or "")
            chunks = []
            for mp in (cr.get("materials") or [])[:15]:
                if not isinstance(mp, str) or not os.path.isfile(mp):
                    continue
                if os.path.abspath(mp) == main_src:
                    continue
                if os.path.splitext(mp)[1].lower() in (".md", ".txt", ".csv", ".tsv", ".json"):
                    try:
                        chunks.append(clean_str(open(mp, encoding="utf-8", errors="ignore").read()[:50000]))
                    except Exception:
                        pass
            extra_text_blob = "\n".join(chunks)
        except Exception:
            pass
    extra_alnum = alnum(extra_text_blob)
    extra_norm = _strip_num_commas(norm(extra_text_blob))

    # Build (number -> set of canonical units) map from full paper + tables + extra materials
    paper_num_units = {}
    unit_corpus_chunks = list(page_text.values())
    for tbl in ctx.get("tables", []):
        hdrs_t = tbl.get("headers") or []
        unit_corpus_chunks.append(" ".join(hdrs_t))
        col_units = [_extract_header_unit(h) for h in hdrs_t]
        for r in (tbl.get("rows") or []):
            unit_corpus_chunks.append(" ".join(str(c) for c in r))
            for ci, cell_v in enumerate(r):
                if ci < len(col_units) and col_units[ci]:
                    cu_h, _ru_h = col_units[ci]
                    for nm in re.findall(r"(?<![A-Za-z0-9.])\d+(?:\.\d+)?", _strip_num_commas(clean_str(str(cell_v)))):
                        paper_num_units.setdefault(nm, set()).add(cu_h)
    if extra_text_blob:
        unit_corpus_chunks.append(extra_text_blob)
    for chunk in unit_corpus_chunks:
        for num, cu, _ru in _extract_num_units(chunk):
            paper_num_units.setdefault(num, set()).add(cu)

    # Index captions by norm_id, canonical key, and raw alnum (list per key to handle multi-page mentions)
    cap_index = {}
    for c in ctx.get("captions", []):
        cid = c.get("id", "")
        ctype = c.get("kind", "")
        full_k, base_k = canonical_ref_key(ctype, cid)
        for key in filter(None, [c.get("norm_id"), full_k, base_k, alnum(cid)]):
            cap_index.setdefault(key, []).append(c)
        # Also allow scheme captions to match figure refs and vice versa
        if full_k.startswith("scheme:"):
            cap_index.setdefault("fig:" + full_k.split(":", 1)[1], []).append(c)

    sec_titles = [(s.get("num", ""), s.get("title", "")) for s in ctx.get("sections", [])]
    # Map extracted figure file paths (abs & basename) to their caption norm_id for cross-check
    fig_file_meta = {}
    for fg in ctx.get("figures", []):
        fp = fg.get("file")
        if fp:
            fig_file_meta[os.path.abspath(fp)] = fg
            fig_file_meta[os.path.basename(fp)] = fg
    issues = []

    max_slides = meta.get("max_slides")
    if max_slides and len(slides) > int(max_slides):
        issues.append({"slide": 0, "sev": "error", "type": "slide_count_exceeds_limit",
                       "detail": f"{len(slides)} > max_slides={max_slides}"})

    has_limitation_slide = False
    seen_slide_imgs = {}
    paced_minutes = []
    for si, sld in enumerate(slides, 1):
        ev = sld.get("evidence") or {}
        refs = ev.get("paper_refs", []) or []
        slide_line = (ev.get("slide_line") or "").strip()
        title = (sld.get("title") or "").strip()
        bullets = [b.get("text", "") if isinstance(b, dict) else str(b)
                   for b in sld.get("bullets", [])]
        right_bullets = [b.get("text", "") if isinstance(b, dict) else str(b)
                         for b in sld.get("right_bullets", [])]
        all_bullets = bullets + right_bullets
        col_headings = [h for h in ((sld.get("left_heading") or "").strip(),
                                    (sld.get("right_heading") or "").strip()) if h]
        kpi_texts = [f"{k.get('value', '')} {k.get('label', '')}".strip()
                     for k in (sld.get("kpis") or sld.get("metrics") or []) if isinstance(k, dict)]
        takeaway_txt = (sld.get("takeaway") or "").strip()
        notes = sld.get("notes", "") or ""
        kind = sld.get("kind", "content")
        table_spec = sld.get("table") or {}

        if LIMIT_WORDS.search(title + " " + " ".join(col_headings + all_bullets)):
            has_limitation_slide = True

        # W6 generic bare-label title (strip leading section numbers like "3. " or "§4 ")
        title_stripped = re.sub(r"^(?:§\s*)?\d+(?:\.\d+)*[\s.、:：-]+", "", title).lower().strip(" :：.-")
        if kind != "cover" and title_stripped in GENERIC_TITLES:
            issues.append({"slide": si, "sev": "warn", "type": "generic_section_title",
                           "detail": f"title '{title}' is a section label, not a claim sentence"})

        if kind not in ("cover",) and not refs:
            issues.append({"slide": si, "sev": "error", "type": "no_evidence",
                           "detail": "non-cover slide has no paper_refs (every content slide must cite paper)"})
        if kind not in ("cover",) and not slide_line:
            issues.append({"slide": si, "sev": "warn", "type": "missing_slide_line",
                           "detail": "evidence.slide_line empty (auto-synthesized fallback will be used)"})
        elif kind not in ("cover",) and slide_line and cap_index:
            for sl_m in REF_PARSE_RE.finditer(clean_str(slide_line)):
                sl_k = sl_m.group(1).lower().rstrip(".")
                if sl_k in ("eq", "eqs", "equation", "equations", "式"):
                    continue
                sl_full, sl_base = canonical_ref_key(sl_k, sl_m.group(0))
                if not (cap_index.get(sl_full) or (sl_base and cap_index.get(sl_base))):
                    sl_al = alnum(sl_m.group(0))
                    if sl_al and not any(sl_al in t for t in page_alnum.values()) and sl_al not in extra_alnum:
                        issues.append({"slide": si, "sev": "warn", "type": "slide_line_ref_not_found",
                                       "detail": f"slide_line cites '{sl_m.group(0)}' absent from paper"})
        if len(notes) < 20 and kind != "cover":
            issues.append({"slide": si, "sev": "warn", "type": "missing_notes",
                           "detail": f"notes len={len(notes)}"})

        # Verify embedded image files exist and match cited figure IDs
        slide_imgs = []
        if sld.get("image"):
            slide_imgs.append(sld["image"])
        for p in (sld.get("images") or []):
            if p and p not in slide_imgs:
                slide_imgs.append(p)
        for im_p in slide_imgs:
            if not isinstance(im_p, str) or not os.path.isfile(im_p):
                issues.append({"slide": si, "sev": "error", "type": "image_file_missing",
                               "detail": f"image file not found: {im_p}"})
            elif kind != "cover":
                abs_im = os.path.abspath(im_p)
                if abs_im in seen_slide_imgs:
                    issues.append({"slide": si, "sev": "warn", "type": "duplicate_figure_across_slides",
                                   "detail": f"{os.path.basename(im_p)} already embedded on slide {seen_slide_imgs[abs_im]}"})
                else:
                    seen_slide_imgs[abs_im] = si
        pm_note = PACE_TAG_RE.search(notes)
        if pm_note:
            try:
                paced_minutes.append(float(pm_note.group(1) or pm_note.group(2)))
            except Exception:
                pass

        cited_pages = set()
        cited_fig_norms = set()
        has_any_quote = False
        for r in refs:
            rtype = (r.get("type") or "page").lower()
            rid_raw = str(r.get("id", ""))
            page = _parse_page_int(r.get("page"))
            quote = r.get("quote", "")
            if quote and len(alnum(quote)) >= 8:
                has_any_quote = True

            if rtype in ("figure", "table", "algorithm", "listing", "equation", "scheme"):
                full_k, base_k = canonical_ref_key(rtype, rid_raw)
                if full_k.startswith("fig:") or full_k.startswith("scheme:"):
                    cited_fig_norms.add(full_k.replace("scheme:", "fig:"))
                    if base_k:
                        cited_fig_norms.add(base_k.replace("scheme:", "fig:"))
                caps = (cap_index.get(full_k) or
                        (cap_index.get(base_k) if base_k else None) or
                        cap_index.get(alnum(rid_raw)))
                if not caps:
                    # Fallback: check if normalized reference label appears verbatim in paper text
                    grp, _, cid_part = full_k.partition(":")
                    variants = {alnum(rid_raw)}
                    if cid_part:
                        pref_map = {
                            "fig": ("fig", "figure", "图"),
                            "tab": ("tab", "table", "表"),
                            "alg": ("alg", "algorithm"),
                            "scheme": ("scheme",),
                            "eq": ("eq", "equation", "式"),
                            "listing": ("listing",),
                        }
                        for pfx in pref_map.get(grp, (grp,)):
                            variants.add(pfx + cid_part)
                    variants = {v for v in variants if len(v) >= 2}
                    found_pages = [p for p, t in sorted(page_alnum.items())
                                   if any(v in t for v in variants)]
                    if rtype == "equation" and found_pages:
                        cited_pages.update(found_pages[:2])
                    elif found_pages and page and int(page) in found_pages:
                        cited_pages.add(int(page))
                        issues.append({"slide": si, "sev": "warn", "type": "ref_caption_unparsed",
                                       "detail": f"{rid_raw} found in text on p.{page} (caption line not parsed)"})
                    elif found_pages:
                        cited_pages.update(found_pages[:2])
                        issues.append({"slide": si, "sev": "warn", "type": "ref_page_mismatch",
                                       "detail": f"{rid_raw} found in text on p.{found_pages[:3]} (caption unparsed)"})
                    else:
                        issues.append({"slide": si, "sev": "error", "type": "ref_not_found",
                                       "detail": f"{rid_raw} not in paper captions"})
                else:
                    cap_pages = sorted(set(int(c["page"]) for c in caps if c.get("page")))
                    cited_pages.update(cap_pages)
                    if page and cap_pages and int(page) not in cap_pages:
                        issues.append({"slide": si, "sev": "warn", "type": "ref_page_mismatch",
                                       "detail": f"{rid_raw} stated p.{page} but caption on p.{cap_pages}"})
                        page = cap_pages[0]
                    elif not page and cap_pages:
                        page = cap_pages[0]
            elif rtype == "section":
                num = rid_raw.strip()
                if num.lower() in ("abstract", "摘要") and ctx.get("abstract_page"):
                    hit = True
                else:
                    hit = False
                    num_clean = re.sub(r"^[§\s]+", "", num).rstrip(".")
                    num_head = num_clean.split()[0].rstrip(".") if num_clean else ""
                    for snum, stitle in sec_titles:
                        sn_c = (snum or "").strip().rstrip(".")
                        if (num_head and sn_c and (num_head == sn_c or num_head.startswith(sn_c + "."))) or \
                                (len(num) > 3 and alnum(num) in alnum(f"{snum} {stitle}")):
                            hit = True
                            break
                if not hit and len(num) > 3:
                    issues.append({"slide": si, "sev": "warn", "type": "section_not_found",
                                   "detail": f"section '{rid_raw}' not in detected headers"})

            if page is not None:
                try:
                    pi = int(page)
                    if pi < 1 or pi > max(n_pages, 1):
                        issues.append({"slide": si, "sev": "error", "type": "page_out_of_range",
                                       "detail": f"p.{pi} but paper has {n_pages} pages"})
                except Exception:
                    pass

            if quote and len(alnum(quote)) >= 8:
                frags = [alnum(x) for x in re.split(r"[;；|]|(?:\.{2,}|…+)", quote) if len(alnum(x)) >= 6]
                if not frags:
                    frags = [alnum(quote)]
                pages_searched = ([page] if page in page_alnum else []) + \
                    [p for p in sorted(page_alnum) if p != page]
                found_on = []
                for p in pages_searched:
                    t = page_alnum.get(p, "")
                    if all(fr in t for fr in frags):
                        found_on = [p]
                        break
                if not found_on:
                    # Also check across adjacent page boundary (sentence wrapped across pages)
                    for p in sorted(page_alnum):
                        if p + 1 in page_alnum:
                            joined = page_alnum[p] + page_alnum[p + 1]
                            if all(fr in joined for fr in frags):
                                found_on = [p, p + 1]
                                break
                if not found_on:
                    # PDF-artifact-tolerant window check (100% numbers + >=85% content tokens)
                    for p in pages_searched:
                        if _fuzzy_window_quote_match(quote, page_text.get(p, ""), page_alnum.get(p, "")):
                            found_on = [p]
                            break
                if not found_on and extra_alnum:
                    if all(fr in extra_alnum for fr in frags) or \
                            _fuzzy_window_quote_match(quote, extra_text_blob, extra_alnum):
                        found_on = [page] if page else []
                        if not found_on:
                            cited_pages.add(1)
                        # Grounded in supplemental case material
                        found_on = found_on or ["material"]
                if not found_on:
                    issues.append({"slide": si, "sev": "error", "type": "quote_not_found",
                                   "detail": f"quote not in paper: '{quote[:65]}'"})
                elif page and found_on and found_on != ["material"] and page not in found_on:
                    issues.append({"slide": si, "sev": "warn", "type": "quote_page_mismatch",
                                   "detail": f"quote found on p.{found_on} not p.{page}"})
                cited_pages.update(p for p in found_on if isinstance(p, int))
            elif page:
                cited_pages.add(page)

        if kind not in ("cover",) and refs and not has_any_quote:
            issues.append({"slide": si, "sev": "warn", "type": "missing_quote_on_slide",
                           "detail": "paper_refs present but none includes a verbatim quote (>=8 alnum chars)"})

        # Cross-check embedded extracted figure file against slide's cited figure IDs & schematic labeling
        slide_img_cap = (sld.get("image_caption") or sld.get("caption") or "").strip()
        if not slide_img_cap and isinstance(sld.get("image_captions"), list):
            slide_img_cap = " ".join(str(x) for x in sld["image_captions"])
        for im_p in slide_imgs:
            if isinstance(im_p, str):
                fg = fig_file_meta.get(os.path.abspath(im_p)) or fig_file_meta.get(os.path.basename(im_p))
                fg_norm = (fg or {}).get("norm_id")
                if fg_norm and fg_norm.startswith("fig:"):
                    if cited_fig_norms:
                        fg_base = re.sub(r"[a-z]+$", "", fg_norm)
                        if fg_norm not in cited_fig_norms and fg_base not in cited_fig_norms:
                            issues.append({
                                "slide": si, "sev": "warn", "type": "image_ref_mismatch",
                                "detail": f"embedded {os.path.basename(im_p)} belongs to {fg.get('caption')} ({fg_norm}), "
                                          f"but slide cites {sorted(cited_fig_norms)}",
                            })
                    elif kind != "cover":
                        issues.append({
                            "slide": si, "sev": "warn", "type": "embedded_figure_uncited",
                            "detail": f"embedded {os.path.basename(im_p)} ({fg.get('caption') or fg_norm}) is not cited in evidence.paper_refs",
                        })
                if fg_norm:
                    src_caps = cap_index.get(fg_norm) or []
                    src_cap_txt = " ".join(c.get("text", "") for c in src_caps)
                    if src_cap_txt and SCHEMATIC_CAP_RE.search(src_cap_txt):
                        lbl_blob = f"{slide_img_cap} {notes}"
                        if not SCHEMATIC_LABEL_RE.search(lbl_blob):
                            issues.append({
                                "slide": si, "sev": "warn", "type": "schematic_unlabeled",
                                "detail": f"{fg.get('caption') or fg_norm} appears to be a schematic/architecture figure; "
                                          "explicitly label as schematic/non-experimental in image_caption or notes",
                            })

        # Key number consistency across title, bullets, right_bullets, kpis, takeaway, and native table cells
        corpus = _strip_num_commas("".join(page_norm.get(p, "") for p in cited_pages if p))
        full_paper_norm = {p: _strip_num_commas(pt) for p, pt in page_norm.items()}
        check_texts = ([title] if kind != "cover" else []) + col_headings + list(all_bullets) + kpi_texts
        if takeaway_txt:
            check_texts.append(takeaway_txt)
        for row in (table_spec.get("rows") or []):
            check_texts.extend(str(c) for c in row)
        chart_spec = sld.get("chart")
        if isinstance(chart_spec, dict):
            for ser in (chart_spec.get("series") or []):
                if isinstance(ser, dict):
                    for v in (ser.get("values") or []):
                        check_texts.append(str(v))
            if chart_spec.get("title"):
                check_texts.append(str(chart_spec["title"]))
            if chart_spec.get("note"):
                check_texts.append(str(chart_spec["note"]))

        # Guard against blindly copying a misleading case.json domain label onto the Cover title
        if kind == "cover" and si == 1 and title and full_paper_norm:
            p12_raw = clean_str(" ".join((pg.get("text") or "") for pg in (ctx.get("pages") or [])[:2])).lower()
            p12_cjk = sum(1 for ch in p12_raw if "\u4e00" <= ch <= "\u9fff") > 30
            t_cjk = sum(1 for ch in title if "\u4e00" <= ch <= "\u9fff") >= 4
            if t_cjk == p12_cjk and len(p12_raw) >= 120:
                stop_en = {"study", "paper", "research", "analysis", "report", "academic", "slides", "presentation",
                           "review", "using", "based", "toward", "towards", "method", "model", "system", "approach"}
                if t_cjk:
                    toks_c = [w for w in re.findall(r"[\u4e00-\u9fff]{2,}", title) if len(w) >= 2]
                    c_grams = []
                    for w in toks_c:
                        for gi in range(0, len(w) - 1):
                            c_grams.append(w[gi:gi + 2])
                    c_grams = list(dict.fromkeys(c_grams))
                    if len(c_grams) >= 3 and not any(g in p12_raw for g in c_grams):
                        issues.append({
                            "slide": si, "sev": "warn", "type": "cover_title_mismatch_paper",
                            "detail": "cover title shares no key terms with actual paper pages 1-2; verify ground-truth paper title",
                        })
                else:
                    toks_e = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z0-9\-]{3,}", title)
                              if w.lower() not in stop_en]
                    toks_e = list(dict.fromkeys(toks_e))
                    if len(toks_e) >= 3 and not any(w in p12_raw for w in toks_e):
                        issues.append({
                            "slide": si, "sev": "warn", "type": "cover_title_mismatch_paper",
                            "detail": "cover title shares no key terms with actual paper pages 1-2; verify ground-truth paper title",
                        })
        seen_miss = set()
        for txt in check_texts:
            base_t = COVER_META_NUM_RE.sub(" ", clean_str(txt)) if kind == "cover" else clean_str(txt)
            cleaned_t = _strip_num_commas(ID_NUM_STRIP_RE.sub(" ", base_t))
            for m_obj in KEY_NUM_RE.finditer(cleaned_t):
                m = m_obj.group(0)
                if m in seen_miss or YEAR_RE.match(m):
                    continue
                if corpus and m in corpus:
                    continue
                seen_miss.add(m)
                if extra_norm and m in extra_norm:
                    continue
                other_p = [p for p, pt in sorted(full_paper_norm.items()) if m in pt][:3]
                if other_p:
                    if corpus:
                        issues.append({"slide": si, "sev": "warn", "type": "number_not_on_cited_page",
                                       "detail": f"'{m}' not on cited pages (found on p.{other_p})"})
                elif full_paper_norm:
                    prefix_ch = cleaned_t[m_obj.start() - 1] if m_obj.start() > 0 else ""
                    delta_hint = " (if a delta, also retain verbatim raw baseline & proposed values)" if prefix_ch in "+-" else ""
                    issues.append({"slide": si, "sev": "warn", "type": "number_absent_from_paper",
                                   "detail": f"'{m}' ABSENT FROM FULL PAPER — verify no rounding/fabrication{delta_hint}"})

        # Scientific unit-mutation consistency check (catches e.g. '38 ms' mutated to '38 s', incl. table header units)
        if kind != "cover" and paper_num_units:
            unit_check_texts = list(check_texts)
            tbl_hdrs = table_spec.get("headers") or []
            if tbl_hdrs and (table_spec.get("rows") or []):
                slide_col_units = [_extract_header_unit(h) for h in tbl_hdrs]
                for row in (table_spec.get("rows") or []):
                    for ci, cell_v in enumerate(row):
                        if ci < len(slide_col_units) and slide_col_units[ci]:
                            _cu_s, ru_s = slide_col_units[ci]
                            for nm in re.findall(r"(?<![A-Za-z0-9.])\d+(?:\.\d+)?", _strip_num_commas(clean_str(str(cell_v)))):
                                unit_check_texts.append(f"{nm} {ru_s}")
            if isinstance(chart_spec, dict) and (chart_spec.get("unit") or "").strip():
                cu_chart = str(chart_spec["unit"]).strip()
                for ser in (chart_spec.get("series") or []):
                    if isinstance(ser, dict):
                        for v in (ser.get("values") or []):
                            unit_check_texts.append(f"{v} {cu_chart}")
            seen_unit_err = set()
            all_paper_norm_blob = "".join(full_paper_norm.values()) + extra_norm
            for utxt in unit_check_texts:
                for num, cu, raw_u in _extract_num_units(utxt):
                    if (num, cu) in seen_unit_err or YEAR_RE.match(num):
                        continue
                    known_units = paper_num_units.get(num)
                    if not known_units or cu in known_units:
                        continue
                    # For 1-digit integers, require either a strict scientific unit or same physical dimension
                    if "." not in num and len(num) < 2:
                        cu_dim = UNIT_DIM_MAP.get(cu)
                        same_dim = bool(cu_dim and any(UNIT_DIM_MAP.get(ku) == cu_dim for ku in known_units))
                        if not same_dim and cu not in STRICT_SCI_UNITS:
                            continue
                    # Double-check normalized adjacency before alerting
                    if f"{num}{cu}" in all_paper_norm_blob.lower():
                        continue
                    seen_unit_err.add((num, cu))
                    issues.append({
                        "slide": si, "sev": "warn", "type": "unit_mismatch_with_paper",
                        "detail": f"'{num} {raw_u}' conflicts with paper unit(s) {sorted(known_units)} for '{num}'",
                    })
        if kind != "cover" and slide_imgs and (table_spec or chart_spec) and len(all_bullets) >= 4:
            issues.append({
                "slide": si, "sev": "warn", "type": "dual_media_crowded",
                "detail": f"slide combines image + {'table' if table_spec else 'chart'} with {len(all_bullets)} bullets; trim bullets or split across slides",
            })

        # Check explicit baseline attribution on comparative claims without a table
        if kind != "cover" and not table_spec:
            slide_blob = f"{title} {' '.join(col_headings + all_bullets + kpi_texts)}"
            if COMPARE_WORDS_RE.search(slide_blob) and len(KEY_NUM_RE.findall(slide_blob)) >= 2:
                if not ATTRIB_HINT_RE.search(slide_blob):
                    issues.append({"slide": si, "sev": "warn", "type": "unattributed_comparison",
                                   "detail": "comparative claim with multiple numbers lacks explicit baseline/author name"})

        # Notes vs slide text duplication
        if notes and all_bullets:
            nb, sb = char_bigrams(notes), char_bigrams(" ".join(all_bullets) + title)
            if nb and sb:
                j = len(nb & sb) / max(1, len(nb | sb))
                if j > 0.55:
                    issues.append({"slide": si, "sev": "warn", "type": "notes_duplicate_slide_text",
                                   "detail": f"overlap {j:.0%}"})

        # Bullet density reminder (keep slides crisp, not pasted paragraphs)
        max_b_lim = 8 if right_bullets else 6
        if kind != "cover" and all_bullets:
            if len(all_bullets) > max_b_lim:
                issues.append({"slide": si, "sev": "warn", "type": "bullet_too_dense",
                               "detail": f"{len(all_bullets)} bullets (>{max_b_lim} recommended)"})
            else:
                long_b = [b for b in all_bullets if len(b.split()) > 45 or len(b) > 160]
                if long_b:
                    issues.append({"slide": si, "sev": "warn", "type": "bullet_too_dense",
                                   "detail": f"long bullet ({len(long_b[0])} chars); tighten to key phrase"})

        # Schematic/architecture figure distinction & caption presence reminder
        has_img = bool(sld.get("image") or sld.get("images"))
        if kind != "cover" and has_img:
            cap_txt = (sld.get("image_caption") or sld.get("caption") or "").strip()
            if not cap_txt and isinstance(sld.get("image_captions"), list):
                cap_txt = " ".join(str(x).strip() for x in sld["image_captions"] if str(x).strip())
            if not cap_txt:
                issues.append({"slide": si, "sev": "warn", "type": "image_missing_caption",
                               "detail": "embedded figure has no image_caption (add Fig. ID, conditions, or schematic note)"})
            if re.search(r"(?:framework|architecture|pipeline|overview|workflow|框架|架构|流程|总览)", title, re.I):
                combined_note = f"{notes} {cap_txt} {' '.join(bullets)}"
                if not re.search(r"(?:示意|架构图|流程图|框架图|schematic|diagram|conceptual|illustration|non-experimental)", combined_note, re.I):
                    issues.append({"slide": si, "sev": "warn", "type": "schematic_unmarked",
                                   "detail": "method/framework figure should explicitly note '示意图/schematic, non-experimental'"})

    # W7 missing limitations slide when paper explicitly has limitations
    lim_pages = ctx.get("limitations_pages") or []
    if lim_pages and (not max_slides or int(max_slides) >= 7) and not has_limitation_slide:
        issues.append({"slide": 0, "sev": "warn", "type": "missing_limitations_slide",
                       "detail": f"paper discusses limitations on p.{lim_pages[:4]} but no slide covers limitations"})

    # Advisory pacing budget check when target duration_minutes is specified
    target_dur = (cr_data.get("constraints") or {}).get("duration_minutes")
    if target_dur and paced_minutes and len(paced_minutes) >= max(2, len(slides) // 2):
        try:
            td = float(target_dur)
            tot_p = sum(paced_minutes)
            if td > 0 and abs(tot_p - td) / td > 0.40:
                issues.append({"slide": 0, "sev": "warn", "type": "pacing_budget_mismatch",
                               "detail": f"speaker notes pacing sums to {tot_p:.1f} min vs target {td:.1f} min"})
        except Exception:
            pass

    n_err = sum(1 for i in issues if i["sev"] == "error")
    n_warn = sum(1 for i in issues if i["sev"] == "warn")
    report = {"ok": n_err == 0, "n_error": n_err, "n_warn": n_warn,
              "n_slides": len(slides), "issues": issues}
    os.makedirs(os.path.dirname(os.path.abspath(a.report)) or ".", exist_ok=True)
    with open(a.report, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
