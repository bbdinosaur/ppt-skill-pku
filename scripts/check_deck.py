#!/usr/bin/env python3
"""check_deck.py — Step 5a of the academic-slides skill.

Static geometry + content compliance checks on the produced deck:
  C1 slide_count vs workorder     C2 shapes within slide bounds
  C3 text-overflow heuristic      C4 content-shape overlap
  C5 image aspect preserved       C6 title + speaker notes present
  C7 min font size                C8 template font-name overrides
  C9 claim-source quotes exist in real material text (anti-hallucination)

Usage: python3 check_deck.py --deck out/presentation.pptx --work work
Writes work/checks.json and prints CHECK_OK/CHECK_ISSUES summary.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckutil as du  # noqa: E402

from pptx import Presentation  # noqa: E402

EXEMPT_ROLES = {"title", "section", "agenda", "outline", "qa", "toc", "closing", "thanks", "references", "existing"}


def corpus_from_material(inv, work: Path | None = None):
    """Build combined reference corpus from material.json + workorder + task.md + case.json + start_deck_dump."""
    import json as _json
    buf = []

    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if isinstance(v, (str, int, float)):
                    buf.append(f"{k}: {v}")
                else:
                    buf.append(str(k))
                    walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
        elif isinstance(o, (str, int, float)):
            buf.append(str(o))

    p = inv.get("material_json_path")
    if p and Path(p).is_file():
        try:
            walk(_json.loads(Path(p).read_text(encoding="utf-8")))
        except Exception:
            try:
                buf.append(Path(p).read_text(encoding="utf-8", errors="replace"))
            except Exception:
                pass
    if inv.get("workorder_raw"):
        walk(inv["workorder_raw"])
    # Include task.md and case.json so edit_instructions / venue dates / prompt numbers verify cleanly
    tmd = inv.get("task_md")
    if tmd and Path(tmd).is_file():
        try:
            buf.append(Path(tmd).read_text(encoding="utf-8", errors="replace"))
        except Exception:
            pass
    cdir = inv.get("case_dir")
    if cdir:
        cj = Path(cdir) / "case.json"
        if cj.is_file():
            try:
                walk(_json.loads(cj.read_text(encoding="utf-8")))
            except Exception:
                pass
    sdd = inv.get("start_deck_dump")
    if sdd and Path(sdd).is_file():
        try:
            buf.append(Path(sdd).read_text(encoding="utf-8", errors="replace"))
        except Exception:
            pass
    if work is not None:
        sdd2 = work / "start_deck_dump.md"
        if sdd2.is_file() and str(sdd2) != str(sdd):
            try:
                buf.append(sdd2.read_text(encoding="utf-8", errors="replace"))
            except Exception:
                pass
    for ef in (inv.get("extra_text_files") or []):
        ep = Path(ef)
        if ep.is_file():
            try:
                buf.append(ep.read_text(encoding="utf-8", errors="replace"))
            except Exception:
                pass
    return "\n".join(buf)


_NORM_TRANS = str.maketrans({
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl",
    "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
    "\u2013": "-", "\u2014": "-", "\u2212": "-", "\u00ad": "", "\u00a0": " ",
    "\uff0c": ",", "\uff1a": ":", "\uff1b": ";", "\uff08": "(", "\uff09": ")",
})


def norm(s: str) -> str:
    # Remove PDF hyphenated line-breaks, ligatures, smart punctuation, and whitespace
    s = re.sub(r"-\s*\n\s*", "", s or "")
    s = s.translate(_NORM_TRANS)
    return re.sub(r"\s+", "", s)


def check_text_fit(prs, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            _, _, w, h = du.shape_geom(prs, slide, sh)
            if not w or not h or not h.pt:
                continue
            w_pt, h_pt = w.pt, h.pt
            fb = du.DEFAULT_TITLE_PT if (sh.is_placeholder and sh.placeholder_format.type in du.TITLE_TYPES) else du.DEFAULT_BODY_PT
            size = du.effective_pt(sh, fb)
            if size <= 0:
                size = fb
            est = du.tf_text_height_pt(sh.text_frame, w_pt, size)
            if est > h_pt * 1.06 + 2:
                issues.append(f"slide {si} '{sh.name}': est text height {est:.0f}pt > box {h_pt:.0f}pt")
    return issues


def check_bounds(prs, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    tol = int(min(SW, SH) * 0.01)
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            l, t, w, h = du.shape_geom(prs, slide, sh)
            if l is None or t is None or w is None or h is None:
                continue
            l, t, w_i, h_i = int(l), int(t), int(w), int(h)
            has_content = (
                (sh.has_text_frame and bool(sh.text_frame.text.strip()))
                or bool(getattr(sh, "has_table", False))
                or bool(getattr(sh, "has_chart", False))
                or (sh.shape_type is not None and "PICTURE" in str(sh.shape_type))
            )
            # Ignore pure decorative full-bleed template banners/backgrounds without content
            if not has_content and (w_i > 0.85 * SW or h_i > 0.85 * SH or w_i * h_i > 0.70 * SW * SH):
                continue
            r, b = l + w_i, t + h_i
            if l < -tol or t < -tol or r > SW + tol or b > SH + tol:
                issues.append(f"slide {si} '{sh.name}' out of bounds l={l} t={t} r={r} b={b} (slide {SW}x{SH})")
    return issues


def check_overlap(prs, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        boxes = []
        for sh in slide.shapes:
            l, t, w, h = du.shape_geom(prs, slide, sh)
            if l is None or t is None or w is None or h is None:
                continue
            l, t, w, h = int(l), int(t), int(w), int(h)
            if w * h > 0.80 * SW * SH:
                continue  # full-bleed background
            if sh.is_placeholder and sh.placeholder_format.type in du.SKIP_TYPES:
                continue
            if h <= 0 or w <= 0:
                continue
            if h < int(0.08 * 914400) or w < int(0.08 * 914400):
                continue  # thin decorative accent line
            has_text = sh.has_text_frame and sh.text_frame.text.strip()
            is_pic = sh.shape_type is not None and "PICTURE" in str(sh.shape_type)
            is_tbl = bool(getattr(sh, "has_table", False))
            is_chart = bool(getattr(sh, "has_chart", False))
            if sh.is_placeholder:
                ph_t = sh.placeholder_format.type
                if ph_t in du.TITLE_TYPES and not has_text:
                    continue
                if ph_t in du.BODY_TYPES and not has_text:
                    continue  # empty placeholder renders nothing
            if not (has_text or is_pic or is_tbl or is_chart):
                continue
            boxes.append((sh.name, l, t, l + w, t + h))
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ix = max(0, min(a[3], b[3]) - max(a[1], b[1]))
                iy = max(0, min(a[4], b[4]) - max(a[2], b[2]))
                inter = ix * iy
                if inter <= 0:
                    continue
                area_a = (a[3] - a[1]) * (a[4] - a[2])
                area_b = (b[3] - b[1]) * (b[4] - b[2])
                smaller = max(min(area_a, area_b), 1)
                # caption/source under a picture intentionally sits below it; ignore tiny strips
                if inter / smaller > 0.20:
                    issues.append(f"slide {si}: overlap {a[0]} x {b[0]} ({100 * inter // smaller}%)")
    return issues


def check_aspect(prs, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            try:
                if "PICTURE" not in str(sh.shape_type):
                    continue
                nat_w, nat_h = sh.image.size
                pw, ph = int(sh.width), int(sh.height)
                if nat_w <= 0 or nat_h <= 0 or pw <= 0 or ph <= 0:
                    continue
                ar_nat, ar_place = nat_h / nat_w, ph / pw
                if abs(ar_nat - ar_place) / ar_nat > 0.02:
                    issues.append(f"slide {si} '{sh.name}': aspect changed {ar_nat:.3f} -> {ar_place:.3f}")
            except Exception:
                pass
    return issues


def check_notes_titles(prs, content_map, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    cm = {s.get("slide"): s for s in (content_map or {}).get("slides", [])}
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        tsh = du.find_slide_title_shape(slide, prs)
        has_title = tsh is not None and tsh.has_text_frame and tsh.text_frame.text.strip()
        if not has_title:
            spec = cm.get(si) or {}
            if spec.get("role") not in ("qa", "closing", "thanks"):
                issues.append(f"slide {si}: no title text")
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        if not notes:
            issues.append(f"slide {si}: missing speaker notes")
    return issues


def check_table_fit(prs, skip_slides: set[int] | None = None):
    """Tables auto-grow downward in PowerPoint: estimate wrapped-cell height vs frame."""
    issues = []
    skip = skip_slides or set()
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            try:
                if not sh.has_table:
                    continue
            except Exception:
                continue
            tbl = sh.table
            ncols = len(tbl.columns)
            col_ws = [c.width.pt for c in tbl.columns]
            total = 0.0
            for row in tbl.rows:
                max_lines = 1
                row_pt = 0.0
                for ci, cell in enumerate(row.cells):
                    t = cell.text
                    size = 12.0
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            if r.font.size is not None:
                                size = r.font.size.pt
                                break
                    row_pt = max(row_pt, size)
                    max_lines = max(max_lines, du.est_lines(t, size, max(col_ws[ci] - 12, 10)) if ci < ncols else 1)
                total += max_lines * (row_pt or 12.0) * 1.28 + 4
            h_pt = sh.height.pt if sh.height else 0
            if total > h_pt * 1.18 + 6:
                issues.append(f"slide {si} '{sh.name}': est table height {total:.0f}pt > frame {h_pt:.0f}pt")
    return issues


def check_min_font(prs, min_pt, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is not None and r.font.size.pt < min_pt and r.text.strip():
                        issues.append(f"slide {si} '{sh.name}': font {r.font.size.pt}pt < {min_pt}pt")
    return issues


def check_font_overrides(prs, skip_slides: set[int] | None = None):
    issues = []
    skip = skip_slides or set()
    for si, slide in enumerate(prs.slides, 1):
        if si in skip:
            continue
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name is not None:
                        issues.append(f"slide {si} '{sh.name}': explicit font '{r.font.name}' overrides template")
    return issues


NUM_RE = re.compile(r"\d+\.\d+|\d{3,}")


def _build_corpus_raw(material_text_file, fallback_corpus=""):
    parts = []
    if material_text_file and Path(material_text_file).is_file():
        parts.append(Path(material_text_file).read_text(encoding="utf-8", errors="replace"))
    if fallback_corpus:
        parts.append(fallback_corpus)
    return "\n".join(parts)


def _build_corpus(material_text_file, fallback_corpus=""):
    raw = _build_corpus_raw(material_text_file, fallback_corpus)
    c_norm = norm(raw)
    c_nocomma = c_norm.replace(",", "")
    return c_norm, c_nocomma


def _nearest_numbers(target: str, corpus_nums: list[tuple[float, str]], k: int = 2) -> str:
    try:
        tv = float(target)
    except Exception:
        return ""
    if not corpus_nums:
        return ""
    tol = max(abs(tv) * 0.15, 2.0)
    cands = sorted({(abs(v - tv), s) for v, s in corpus_nums if abs(v - tv) <= tol})
    if not cands:
        return ""
    return " (closest in material: " + ", ".join(s for _, s in cands[:k]) + ")"


def _nearest_quote_line(quote: str, raw_lines: list[str]) -> str:
    """Return the closest real source line by character-trigram overlap for 1-turn fixes."""
    qn = norm(quote).lower()
    if len(qn) < 6 or not raw_lines:
        return ""
    q_tri = {qn[i:i + 3] for i in range(len(qn) - 2)}
    if not q_tri:
        return ""
    best_sc, best_ln = 0.0, ""
    for ln in raw_lines:
        ln_n = norm(ln).lower()
        if len(ln_n) < 6:
            continue
        l_tri = {ln_n[i:i + 3] for i in range(len(ln_n) - 2)}
        if not l_tri:
            continue
        sc = len(q_tri & l_tri) / len(q_tri)
        if sc > best_sc:
            best_sc, best_ln = sc, ln
    return f" -> closest real line: '{best_ln[:90]}'" if best_sc >= 0.30 else ""


def _number_in_corpus(m: str, corpus: str, corpus_nc: str, corpus_vals: set[float]) -> bool:
    """Check whether number token `m` is faithfully present in corpus, allowing only
    mathematically exact representation equivalences (trailing zeros, leading-zero .85,
    exact decimal<->percent 0.854<->85.4, exact k/M scale 1200<->1.2k)."""
    if m in corpus or m in corpus_nc:
        return True
    if "." in m:
        stripped = m.rstrip("0").rstrip(".")
        if stripped and (stripped in corpus_nc) and re.search(rf"(?<![\d.]){re.escape(stripped)}(?![\d.])", corpus_nc):
            return True
        # Leading-zero stat convention: paper '.85' <-> slide '0.85'
        if m.startswith("0.") and m[1:] in corpus_nc:
            return True
    try:
        v = float(m)
    except Exception:
        return False
    # Exact decimal <-> percent (e.g. 0.854 <-> 85.4%) and exact k/M/B scale (1200 <-> 1.2k)
    for scale in (1.0, 0.01, 100.0, 1000.0, 1e6, 1e9):
        sv = v * scale
        if any(abs(sv - cv) <= max(1e-12, abs(sv) * 1e-6) for cv in corpus_vals):
            return True
    return False


_SUP_MAP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺", "0123456789-+")
_SCI_E_RE = re.compile(r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)\s*[eE]\s*([+-]?\d+)(?!\d)")
_SCI_X10_RE = re.compile(r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)\s*[×xX\*]\s*10\^?\s*\(?([+-]?\d+)\)?")
_CORPUS_SUFFIX_RE = re.compile(r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)\s*([kKmMbB])(?![A-Za-z0-9])")
_SLIDE_SUFFIX_RE = re.compile(r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)([kKMB])(?![A-Za-z0-9])")
_LEAD_DOT_RE = re.compile(r"(?<![A-Za-z0-9.])(\.\d+)(?!\d)")


def _val_in_corpus_vals(val: float, corpus_vals: set[float]) -> bool:
    return any(abs(val - cv) <= max(1e-12, abs(val) * 1e-6) for cv in corpus_vals)


_UNIT_EXP_RE = re.compile(r"(?:(?<![A-Za-z0-9.])10\^?\s*\(?([+-]?\d+)\)?|(?<![A-Za-z0-9.])\d+(?:\.\d+)?\s*[eE]\s*([+-]?\d+)(?!\d))")


def _extract_special_corpus_vals(raw_corpus: str) -> tuple[set[float], set[int]]:
    """Parse explicit scientific notation (1.2e-4, 3.5×10^-5), unit exponents (1e-6, 10^-6),
    and suffixed magnitudes (1.2k, 3.4M, 2.5B) from raw whitespace-preserving corpus."""
    vals: set[float] = set()
    exps: set[int] = set()
    c_sup = re.sub(r"(?<=\d),(?=\d)", "", raw_corpus.translate(_SUP_MAP))
    for rx in (_SCI_E_RE, _SCI_X10_RE):
        for m in rx.finditer(c_sup):
            try:
                ev = int(m.group(2))
                vals.add(float(m.group(1)) * (10.0 ** ev))
                exps.add(ev)
            except Exception:
                pass
    for m in _UNIT_EXP_RE.finditer(c_sup):
        try:
            exps.add(int(m.group(1) or m.group(2)))
        except Exception:
            pass
    mults = {"k": 1e3, "m": 1e6, "b": 1e9}
    for m in _CORPUS_SUFFIX_RE.finditer(c_sup):
        try:
            vals.add(float(m.group(1)) * mults[m.group(2).lower()])
        except Exception:
            pass
    return vals, exps


def check_numbers(content_map, material_text_file, fallback_corpus="", skip_slides: set[int] | None = None):
    """Number fidelity: substantive numbers on slides (decimals or 3+ digits) must be
    findable in the combined source corpus (PDF + material.json + start deck)."""
    issues = []
    checked = 0
    if not content_map:
        return issues, checked
    skip = skip_slides or set()
    raw_corpus = _build_corpus_raw(material_text_file, fallback_corpus)
    corpus, corpus_nc = _build_corpus(material_text_file, fallback_corpus)
    if not corpus:
        return issues, checked
    corpus_nums = []
    raw_nc = re.sub(r"(?<=\d),(?=\d)", "", raw_corpus)
    corpus_vals, corpus_exps = _extract_special_corpus_vals(raw_corpus)
    for num_s in set(NUM_RE.findall(raw_nc)) | set(NUM_RE.findall(corpus_nc)):
        try:
            fv = float(num_s)
            corpus_nums.append((fv, num_s))
            corpus_vals.add(fv)
        except Exception:
            pass
    # Also collect leading-dot decimals like .854 in APA/stats tables
    for dm in set(re.findall(r"(?<!\d)\.\d+", raw_nc)):
        try:
            corpus_vals.add(float("0" + dm))
        except Exception:
            pass
    for s in (content_map.get("slides") or []):
        if s.get("slide") in skip:
            continue
        role = (s.get("role") or "").lower()
        texts = [s.get("title") or ""] + [b if isinstance(b, str) else b.get("t", "")
                                          for b in (s.get("bullets") or [])]
        if s.get("subtitle") and role not in ("title", "qa", "closing", "thanks"):
            texts.append(str(s["subtitle"]))
        if s.get("takeaway"):
            texts.append(str(s["takeaway"]))
        for hi in (s.get("highlights") or []):
            if isinstance(hi, dict):
                texts.extend(str(hi.get(k, "")) for k in ("value", "label", "sub"))
        for st in (s.get("steps") or []):
            if isinstance(st, dict):
                texts.extend(str(st.get(k, "")) for k in ("title", "desc", "detail"))
            elif isinstance(st, str):
                texts.append(st)
        tbl = s.get("table") or {}
        if tbl.get("caption"):
            texts.append(str(tbl["caption"]))
        for row in (tbl.get("rows") or []):
            texts.extend(str(c) for c in row)
        img = s.get("image") or {}
        if img.get("caption"):
            texts.append(img["caption"])
        for im in (s.get("images") or []):
            if isinstance(im, dict) and im.get("caption"):
                texts.append(str(im["caption"]))
        ch = s.get("chart") or {}
        for ck in ("title", "caption", "x_title", "y_title"):
            if ch.get(ck):
                texts.append(str(ch[ck]))
        texts.extend(str(c) for c in (ch.get("categories") or []))
        for sr in (ch.get("series") or []):
            texts.append(str(sr.get("name", "")))
            texts.extend(str(v) for v in (sr.get("values") or []))
        seen = set()
        sec_tag_re = re.compile(
            r"(?:Sec(?:tion)?\.?|§|Fig(?:ure)?\.?|Table|Tab\.?|Eq(?:uation)?\.?|Alg\.?|Appendix|v|"
            r"图|表|公式|式|第)\s*\(?\d+\.\d+\)?",
            re.I,
        )
        mults_s = {"k": 1e3, "K": 1e3, "M": 1e6, "B": 1e9}
        for t in texts:
            t_clean = sec_tag_re.sub(" ", str(t).translate(_SUP_MAP).replace(",", ""))
            # 1) Check full scientific-notation tokens first (catches 100x exponent hallucinations
            #    and allows exact bidirectional match: paper 0.00035 <-> slide 3.5×10^-4)
            for rx in (_SCI_E_RE, _SCI_X10_RE):
                for sm in rx.finditer(t_clean):
                    tok = sm.group(0).strip()
                    if tok in seen:
                        continue
                    seen.add(tok)
                    checked += 1
                    try:
                        mv = float(sm.group(1))
                        ev = int(sm.group(2))
                        sv = mv * (10.0 ** ev)
                    except Exception:
                        continue
                    ok_sci = _val_in_corpus_vals(sv, corpus_vals) or (
                        ev in corpus_exps and _val_in_corpus_vals(mv, corpus_vals)
                    )
                    if not ok_sci:
                        hint = _nearest_numbers(f"{sv:g}", corpus_nums)
                        issues.append(f"slide {s.get('slide')}: number '{tok}' not found in material{hint}")
                t_clean = rx.sub(" ", t_clean)
            # 2) Check suffixed magnitudes (e.g. 2.4k, 1.8M, 2.5B)
            for sm in _SLIDE_SUFFIX_RE.finditer(t_clean):
                tok = sm.group(0).strip()
                if tok in seen:
                    continue
                seen.add(tok)
                checked += 1
                try:
                    base_v = float(sm.group(1))
                    suf = sm.group(2)
                    sv = base_v * mults_s[suf]
                except Exception:
                    continue
                ok_s = _val_in_corpus_vals(sv, corpus_vals) or (
                    suf == "K" and base_v >= 100 and _val_in_corpus_vals(base_v, corpus_vals)
                )
                if not ok_s:
                    hint = _nearest_numbers(f"{sv:g}", corpus_nums)
                    issues.append(f"slide {s.get('slide')}: number '{tok}' not found in material{hint}")
            t_clean = _SLIDE_SUFFIX_RE.sub(" ", t_clean)
            # 3) Check APA leading-dot decimals (e.g. .854, .001)
            for dm in _LEAD_DOT_RE.findall(t_clean):
                if dm in seen:
                    continue
                seen.add(dm)
                checked += 1
                if not _number_in_corpus("0" + dm, corpus, corpus_nc, corpus_vals):
                    hint = _nearest_numbers("0" + dm, corpus_nums)
                    issues.append(f"slide {s.get('slide')}: number '{dm}' not found in material{hint}")
            t_clean = _LEAD_DOT_RE.sub(" ", t_clean)
            # 4) Standard decimals and 3+ digit integers
            for m in NUM_RE.findall(t_clean):
                if m in seen:
                    continue
                if re.match(r"^(19|20)\d{2}$", m):
                    continue  # publication/citation/venue year
                if m == "100":
                    continue  # generic 100% scale ceiling
                seen.add(m)
                checked += 1
                if _number_in_corpus(m, corpus, corpus_nc, corpus_vals):
                    continue
                hint = _nearest_numbers(m, corpus_nums)
                issues.append(f"slide {s.get('slide')}: number '{m}' not found in material{hint}")
    return issues, checked


_CITE_BRACKET_RE = re.compile(r"\[\d+(?:[,;\-–]\s*\d+)*\]")


def check_quotes(content_map, material_text_file, fallback_corpus="", skip_slides: set[int] | None = None):
    """Verify on-slide claim quotes really exist in the source material (anti-hallucination)."""
    issues = []
    checked = 0
    if not content_map:
        return issues, checked
    skip = skip_slides or set()
    raw_corpus = _build_corpus_raw(material_text_file, fallback_corpus)
    text, text_nc = _build_corpus(material_text_file, fallback_corpus)
    text_low = text.lower()
    text_nocite = _CITE_BRACKET_RE.sub("", text_low)
    raw_lines = [ln.strip() for ln in raw_corpus.splitlines() if 12 <= len(ln.strip()) <= 260][:2400]
    for s in (content_map.get("slides") or []):
        if s.get("slide") in skip:
            continue
        for src in (s.get("sources") or []):
            q = norm(src.get("quote", ""))
            if len(q) < 6:
                continue
            checked += 1
            if text and (
                q in text
                or q.replace(",", "") in text_nc
                or q.lower() in text_low
                or _CITE_BRACKET_RE.sub("", q.lower()) in text_nocite
            ):
                continue
            hint = _nearest_quote_line(src.get("quote", ""), raw_lines)
            issues.append(f"slide {s.get('slide')}: quote not found in material: '{src.get('quote', '')[:45]}...'{hint}")
    return issues, checked


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--min-font", type=float, default=10.0)
    ap.add_argument("--expected-count", type=int, default=None)
    ap.add_argument("--max-count", type=int, default=None)
    args = ap.parse_args()

    work = Path(args.work)
    inv = du.load_inventory(work)
    wo = inv.get("workorder") or {}
    def parse_count(v):
        if v is None:
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            if re.search(r"\d+\s*[-–—~至到]\s*\d+", str(v)):
                return None  # range handled by min/max below
            m = re.search(r"\d+", str(v))
            return int(m.group()) if m else None

    expected = args.expected_count or parse_count(wo.get("slide_count"))
    min_count = parse_count(wo.get("min_slides"))
    max_count = args.max_count or parse_count(wo.get("max_slides"))
    if isinstance(wo.get("slide_count"), str):
        rm = re.search(r"(\d+)\s*[-–—~至到]\s*(\d+)", wo["slide_count"])
        if rm:
            min_count = min_count or int(rm.group(1))
            max_count = max_count or int(rm.group(2))

    content_map = {}
    cmf = work / "content_map.json"
    if cmf.is_file():
        try:
            content_map = json.loads(cmf.read_text(encoding="utf-8"))
        except Exception:
            pass

    prs = Presentation(args.deck)
    n = len(prs.slides)
    checks = []

    def add(cid, ok, detail):
        checks.append({"id": cid, "pass": bool(ok), "detail": detail})

    if expected:
        add("slide_count", n == int(expected), f"{n} slides == {expected}")
    elif min_count and max_count:
        add("slide_count", int(min_count) <= n <= int(max_count),
            f"{n} slides in [{min_count}, {max_count}]")
    elif min_count:
        add("slide_count", n >= int(min_count), f"{n} slides >= {min_count}")
    else:
        add("slide_count", True, f"{n} slides (no exact workorder count)")
    if max_count and not (min_count and not expected):
        add("slide_limit", n <= int(max_count), f"{n} <= {max_count}")

    # In edit mode, exempt untouched starting-deck slides from pre-existing geometry/font
    # warnings (and modified in-place slides from pre-existing starting-deck font overrides)
    unchanged_slides: set[int] = set()
    existing_font_slides: set[int] = set()
    sdf = work / "slide_diff.json"
    if content_map.get("mode") == "edit" and sdf.is_file():
        try:
            sd = json.loads(sdf.read_text(encoding="utf-8"))
            for row in (sd.get("slides") or []):
                si = row.get("slide")
                if not isinstance(si, int):
                    continue
                if row.get("status") == "unchanged":
                    unchanged_slides.add(si)
                    existing_font_slides.add(si)
                elif row.get("status") == "modified":
                    existing_font_slides.add(si)
        except Exception:
            pass

    b = check_bounds(prs, unchanged_slides)
    add("bounds", not b, f"{len(b)} violations" + (": " + "; ".join(b[:6]) if b else ""))
    ov = check_text_fit(prs, unchanged_slides)
    add("text_fit", not ov, f"{len(ov)} potential overflows" + (": " + "; ".join(ov[:6]) if ov else ""))
    ol = check_overlap(prs, unchanged_slides)
    add("overlap", not ol, f"{len(ol)} overlaps" + (": " + "; ".join(ol[:6]) if ol else ""))
    asp = check_aspect(prs, unchanged_slides)
    add("aspect", not asp, f"{len(asp)} aspect changes" + (": " + "; ".join(asp[:4]) if asp else ""))
    nt = check_notes_titles(prs, content_map, unchanged_slides)
    add("titles_notes", not nt, f"{len(nt)} missing" + (": " + "; ".join(nt[:6]) if nt else ""))
    mf = check_min_font(prs, args.min_font, existing_font_slides)
    add("min_font", not mf, f"{len(mf)} violations" + (": " + "; ".join(mf[:4]) if mf else ""))
    tbl = check_table_fit(prs, unchanged_slides)
    add("table_fit", not tbl, f"{len(tbl)} potential table overflows" + (": " + "; ".join(tbl[:4]) if tbl else ""))
    fo = check_font_overrides(prs, existing_font_slides)
    add("font_inheritance", not fo, f"{len(fo)} overrides" + (": " + "; ".join(fo[:4]) if fo else ""))

    # evidence: every content slide cites >=1 source; quotes verified against material text
    mt_file = (inv.get("pdf_extract") or {}).get("text_file")
    ev_issues = []
    cm_slides = content_map.get("slides") or []
    for s in cm_slides:
        if (s.get("role") or "") in EXEMPT_ROLES or s.get("slide") in unchanged_slides:
            continue
        if not (s.get("sources") or []):
            ev_issues.append(f"slide {s.get('slide')} ({s.get('role')}): no claim_source")
    fb_corpus = corpus_from_material(inv, work)
    qi, qchecked = check_quotes(content_map, mt_file, fb_corpus, unchanged_slides)
    add("claim_sources", not ev_issues, f"{len(ev_issues)} slides without sources" + (": " + "; ".join(ev_issues[:6]) if ev_issues else ""))
    add("quote_verification", not qi, f"{qchecked} quotes checked, {len(qi)} unmatched" + (": " + "; ".join(qi[:6]) if qi else ""))
    ni, nchecked = check_numbers(content_map, mt_file, fb_corpus, unchanged_slides)
    add("number_fidelity", not ni, f"{nchecked} numbers checked, {len(ni)} unmatched" + (": " + "; ".join(ni[:8]) if ni else ""))

    all_pass = all(c["pass"] for c in checks)
    (work / "checks.json").write_text(
        json.dumps({"all_pass": all_pass, "slide_count": n, "checks": checks},
                   ensure_ascii=False, indent=1), encoding="utf-8")
    du.log_change(work, "check", f"all_pass={all_pass} issues={sum(1 for c in checks if not c['pass'])}")
    fails = [c for c in checks if not c["pass"]]
    print(("CHECK_OK " if all_pass else "CHECK_ISSUES ") +
          json.dumps({"slides": n, "passed": f"{len(checks) - len(fails)}/{len(checks)}",
                      "failures": [{"id": c["id"], "detail": c["detail"]} for c in fails]},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"CHECK_FAILED: {e}")
        sys.exit(1)
