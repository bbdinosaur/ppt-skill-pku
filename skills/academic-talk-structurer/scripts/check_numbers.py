#!/usr/bin/env python3
"""Numeric fidelity check: every number on slides (including tables) must be traceable to source text.

Usage:
  python3 check_numbers.py --spec /app/work/deck_spec.json --source /app/work/paper.txt \
      [--out /app/work/number_check.json] [--allow /app/work/number_allow.json]

Checks claims/bullets/titles/captions/tables of all slides (notes excluded: speech register).
Match rule: exact substring in source (comma-normalized) OR float within 0.5% of any source number.
Years (1900-2100) and month/day tokens inside dates (YYYY-MM) are skipped.
Optional allowlist (--allow): [{"value":"5.5","reason":"derived 75.2-69.7 (Table 2)"}] — justified numbers are
reported separately and do not affect exit code. Exit: 0 clean/justified, 1 suspects, 3 source missing.
No-script fallback: grep each slide number in paper.txt manually.
"""
import argparse
import json
import re
import sys
from pathlib import Path

NUM_RE = re.compile(r"\d+(?:\.\d+)?[eE][+-]?\d+|\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?")
YEAR_RE = re.compile(r"^(19|20)\d{2}$")
DATECTX_RE = re.compile(r"\d{4}[-/.]\d{1,2}([-/.]\d{1,2})?|[01]?\d月\d{1,2}[日号]?")
SUP_MAP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻", "0123456789+-")
SCI_PAT_RE = re.compile(
    r"(?:(\d+(?:\.\d+)?)\s*(?:\\times|\\cdot|[×xX*])\s*)?10(?:\s*\^\s*\{?\s*([+−-]?\d+)\s*\}?|\s*([⁺⁻]?[⁰¹²³⁴⁵⁶⁷⁸⁹]+)|\s*([−-]\s*\d+))"
)


def extract_sci_values(s: str):
    """Parse ASCII/LaTeX/PDF-Unicode scientific notation (e.g. 2.5×10^-4, 2.5×10⁻⁴, 10−3)."""
    out = []
    for sm in SCI_PAT_RE.finditer(s or ""):
        try:
            mant = float(sm.group(1)) if sm.group(1) else 1.0
            raw_exp = sm.group(2) or sm.group(3) or sm.group(4) or ""
            exp_s = raw_exp.translate(SUP_MAP).replace("−", "-").replace(" ", "")
            exp = int(exp_s)
            if -18 <= exp <= 18:
                out.append(mant * (10.0 ** exp))
        except ValueError:
            pass
    return out


def normalize_sci_text(s: str) -> str:
    """Replace 10^k / a×10⁻⁴ expressions with float tokens so NUM_RE checks the true scientific value."""
    def _rep(sm):
        try:
            mant = float(sm.group(1)) if sm.group(1) else 1.0
            raw_exp = sm.group(2) or sm.group(3) or sm.group(4) or ""
            exp = int(raw_exp.translate(SUP_MAP).replace("−", "-").replace(" ", ""))
            if -18 <= exp <= 18:
                return f"{mant * (10.0 ** exp):g}"
        except ValueError:
            pass
        return sm.group(0)
    return SCI_PAT_RE.sub(_rep, s or "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--source", required=True)
    ap.add_argument("--tables", help="optional paper_tables.json path to include clean structured table numbers")
    ap.add_argument("--qa", help="optional qa_plan.json path to audit answer_key numbers (advisory)")
    ap.add_argument("--out")
    ap.add_argument("--allow", dest="allow", help="allowlist json: [{value,reason}]")
    a = ap.parse_args()
    src = Path(a.source)
    if not src.exists():
        print(f"FATAL: source text missing: {src} (run S1 pdf->text first)", file=sys.stderr)
        sys.exit(3)
    text = src.read_text(encoding="utf-8", errors="ignore")
    if a.tables and Path(a.tables).exists():
        try:
            tbls = json.loads(Path(a.tables).read_text(encoding="utf-8"))
            tb_bits = []
            for tb in (tbls if isinstance(tbls, list) else []):
                tb_bits.append(" ".join(str(x) for x in (tb.get("headers") or [])))
                for row in (tb.get("rows") or []):
                    tb_bits.append(" ".join(str(x) for x in (row or [])))
            if tb_bits:
                text = text + "\n" + "\n".join(tb_bits)
        except Exception:
            pass
    src_nums = set()
    src_tokens = set()
    UNIT_MULT = {"k": 1e3, "K": 1e3, "m": 1e6, "M": 1e6, "b": 1e9, "B": 1e9, "万": 1e4, "亿": 1e8}
    for m in NUM_RE.finditer(text):
        tok = m.group(0).replace(",", "")
        # Do not let letter-glued hardware/model identifiers (e.g. A100, H100, FP16, L2) pollute src_tokens or src_nums
        if m.start() > 0 and text[m.start() - 1].isascii() and text[m.start() - 1].isalpha():
            continue
        src_tokens.add(tok)
        try:
            val = float(tok)
            src_nums.add(val)
            # Also register scaled forms (e.g. 2M -> 2e6, 200万 -> 2e6; 0.763 <-> 76.3%; 450ms <-> 0.45s; 2.5 min <-> 150s)
            nxt = text[m.end(): m.end() + 12].strip()
            if re.match(r"^thousand\b", nxt, re.I):
                src_nums.add(val * 1e3)
            elif re.match(r"^million\b", nxt, re.I):
                src_nums.add(val * 1e6)
            elif re.match(r"^billion\b", nxt, re.I):
                src_nums.add(val * 1e9)
            elif nxt and nxt[0] in UNIT_MULT and not re.match(r"^(ms|mb|min)\b", nxt, re.I):
                src_nums.add(val * UNIT_MULT[nxt[0]])
                if nxt[0] in ("k", "K"):
                    src_nums.add(val * 1024.0)
                elif nxt[0] in ("m", "M"):
                    src_nums.add(val * 1048576.0)
            if re.match(r"^ms\b", nxt, re.I):
                src_nums.add(val / 1000.0)
            elif re.match(r"^(?:min|mins|minute|minutes|分钟)\b", nxt, re.I) and 0 < val < 180:
                src_nums.add(val * 60.0)
                src_nums.add(val / 60.0)
            elif re.match(r"^s\b", nxt) and 0 < val < 7200:
                if val < 60:
                    src_nums.add(val * 1000.0)
                if val >= 15:
                    src_nums.add(val / 60.0)
            elif re.match(r"^GB\b", nxt, re.I):
                src_nums.add(val * 1000.0)
                src_nums.add(val * 1024.0)
            elif re.match(r"^MB\b", nxt, re.I):
                src_nums.add(val / 1000.0)
                src_nums.add(val / 1024.0)
                src_nums.add(val * 1000.0)
                src_nums.add(val * 1024.0)
            elif re.match(r"^KB\b", nxt, re.I):
                src_nums.add(val / 1000.0)
                src_nums.add(val / 1024.0)
            elif re.match(r"^TB\b", nxt, re.I):
                src_nums.add(val * 1000.0)
                src_nums.add(val * 1024.0)
            elif re.match(r"^(?:us|µs|μs)\b", nxt, re.I):
                src_nums.add(val / 1000.0)
            elif re.match(r"^(?:hr|hrs|hour|hours)\b", nxt, re.I) and 0 < val < 24:
                src_nums.add(val * 60.0)
            if 0 < val < 1:
                src_nums.add(val * 100.0)
            elif 1 < val <= 100 and (nxt.startswith("%")):
                src_nums.add(val / 100.0)
        except ValueError:
            pass
    # Also register simple fractional shorthands in source text (e.g. 1/4 -> 25%, 1/3 -> 33.3%)
    for fm in re.finditer(r"(?<![\d/])([1-9])\s*/\s*([2-9]|10)(?![\d/])", text):
        try:
            a_f, b_f = float(fm.group(1)), float(fm.group(2))
            if 0 < a_f < b_f:
                src_nums.add(a_f / b_f)
                src_nums.add((a_f / b_f) * 100.0)
        except ValueError:
            pass
    # Also register 10^k / a×10^k / Unicode superscript & minus scientific notation from source text
    for val_sci in extract_sci_values(text):
        src_nums.add(val_sci)
        src_tokens.add(f"{val_sci:g}")

    src_list = list(src_nums)

    def close(v, is_approx=False):
        for s in src_list:
            if is_approx and abs(v - s) <= max(0.5, abs(s) * 0.04):
                return True
            tol = max(1e-12, abs(s) * 0.015) if (0 < abs(s) < 0.01 or 0 < abs(v) < 0.01) else max(0.01, abs(s) * 0.005)
            if abs(v - s) <= tol:
                return True
        return False

    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    allow = {}
    if a.allow and Path(a.allow).exists():
        for e in json.loads(Path(a.allow).read_text(encoding="utf-8")):
            allow[str(e.get("value")).replace(",", "")] = str(e.get("reason", ""))

    def expand_units(val: float, after_str: str):
        cands = [val]
        ast_s = after_str.strip()
        ach = ast_s[:1]
        if re.match(r"^thousand\b", ast_s, re.I):
            cands.append(val * 1e3)
        elif re.match(r"^million\b", ast_s, re.I):
            cands.append(val * 1e6)
        elif re.match(r"^billion\b", ast_s, re.I):
            cands.append(val * 1e9)
        elif ach in UNIT_MULT and not re.match(r"^(ms|mb|min)\b", ast_s, re.I):
            cands.append(val * UNIT_MULT[ach])
            if ach in ("k", "K"):
                cands.append(val * 1024.0)
            elif ach in ("m", "M"):
                cands.append(val * 1048576.0)
        if re.match(r"^ms\b", ast_s, re.I):
            cands.append(val / 1000.0)
        elif re.match(r"^(?:min|mins|minute|minutes|分钟)\b", ast_s, re.I) and 0 < val < 180:
            cands.extend([val * 60.0, val / 60.0])
        elif re.match(r"^s\b", ast_s) and 0 < val < 7200:
            if val < 60:
                cands.append(val * 1000.0)
            if val >= 15:
                cands.append(val / 60.0)
        elif re.match(r"^GB\b", ast_s, re.I):
            cands.extend([val * 1000.0, val * 1024.0])
        elif re.match(r"^MB\b", ast_s, re.I):
            cands.extend([val / 1000.0, val / 1024.0, val * 1000.0, val * 1024.0])
        elif re.match(r"^KB\b", ast_s, re.I):
            cands.extend([val / 1000.0, val / 1024.0])
        elif re.match(r"^TB\b", ast_s, re.I):
            cands.extend([val * 1000.0, val * 1024.0])
        elif re.match(r"^(?:us|µs|μs)\b", ast_s, re.I):
            cands.append(val / 1000.0)
        elif re.match(r"^(?:hr|hrs|hour|hours)\b", ast_s, re.I) and 0 < val < 24:
            cands.append(val * 60.0)
        return cands

    def check_derived_math(val: float, traced: list, kind_hint: str = "delta"):
        if val <= 0 or not traced:
            return None
        # Each item is (float_val, tok_id, unit_kind) so unit-scaled variants of the SAME token are never paired against each other
        norm = [(x[0], x[1], (x[2] if len(x) >= 3 else "plain")) if isinstance(x, tuple) else (x, idx, "plain") for idx, x in enumerate(traced)]
        seen_pairs = set()
        pairs = []
        for ia in range(len(norm)):
            av, ida, _ = norm[ia]
            for ib in range(ia + 1, len(norm)):
                bv, idb, _ = norm[ib]
                if ida == idb or av == bv:
                    continue
                lo, hi = (av, bv) if av < bv else (bv, av)
                if (lo, hi) not in seen_pairs:
                    seen_pairs.add((lo, hi))
                    pairs.append((lo, hi))
        tol = max(0.06, val * 0.018)
        tol_pct = max(0.55, val * 0.025) if "." not in f"{val:g}" else max(0.15, val * 0.018)

        def _try_diff():
            for av, bv in pairs:
                if av > 0 and bv / av >= 50.0:
                    continue
                diff = abs(bv - av)
                if abs(diff - val) <= tol:
                    return f"auto-verified delta |{bv:g} - {av:g}| = {diff:.3g}"
            return None

        def _try_ratio():
            for av, bv in pairs:
                if av <= 0:
                    continue
                ratio = bv / av
                if abs(ratio - val) <= tol:
                    return f"auto-verified ratio {bv:g} / {av:g} = {ratio:.3g}x"
            return None

        def _try_rel():
            for av, bv in pairs:
                if av > 0 and bv / av >= 50.0:
                    continue
                diff = abs(bv - av)
                if bv > 0:
                    rp2 = (diff / bv) * 100.0
                    if abs(rp2 - val) <= tol_pct:
                        return f"auto-verified rel-reduction ({bv:g}-{av:g})/{bv:g} = {rp2:.1f}%"
                if av > 0:
                    rp1 = (diff / av) * 100.0
                    if abs(rp1 - val) <= tol_pct:
                        return f"auto-verified rel-change ({bv:g}-{av:g})/{av:g} = {rp1:.1f}%"
            return None

        def _try_recip():
            # Equivalent speedup/compression <-> reduction % (ONLY when source token is actually a ratio/pct!)
            seen_v = set()
            r_tol_pct = 0.55 if "." not in f"{val:g}" else 0.25
            for av, _, ukind in norm:
                if (av, ukind) in seen_v:
                    continue
                seen_v.add((av, ukind))
                if kind_hint == "pct" and ukind == "ratio" and 1.05 < av <= 50.0:
                    eq_pct = (1.0 - 1.0 / av) * 100.0
                    if abs(eq_pct - val) <= r_tol_pct:
                        return f"auto-verified equiv 1 - 1/{av:g}x = {eq_pct:.1f}%"
                elif kind_hint == "ratio" and ukind == "pct" and 5.0 <= av <= 99.0:
                    eq_ratio = 1.0 / (1.0 - av / 100.0)
                    if abs(eq_ratio - val) <= min(tol, 0.15):
                        return f"auto-verified equiv 1/(1-{av:g}%) = {eq_ratio:.3g}x"
            return None

        order = {
            "pct": (_try_rel, _try_recip, _try_diff),
            "ratio": (_try_ratio, _try_recip, _try_diff),
        }.get(kind_hint, (_try_diff, _try_ratio, _try_rel))
        for fn in order:
            hit = fn()
            if hit:
                return hit
        return None

    suspects, justified, checked = [], [], 0
    deck_verified_derived = {}  # tok -> reason (verified on an evidence/content slide, reusable on title/conclusion)
    slide_traced_by_slide = {}
    slide_verified_tokens = {}  # slide_num -> list of distinct verified raw tokens
    pending_cross_slide = []
    CITE_SLIDE_BEFORE_RE = re.compile(r"(?:Fig(?:ure)?\.?|Tab(?:le)?\.?|Sec(?:tion)?\.?|Eq(?:uation)?\.?|Alg(?:orithm)?\.?|Slide|Slides|p\.|pp\.|§|P|图|表|第|公式|算法|原P)\s*\(?$", re.I)
    MODEL_ID_BEFORE_RE = re.compile(r"(?:ResNet|VGG|ViT[-A-Za-z/]*|GPT|Llama|Qwen|BERT|RoBERTa|T5|CIFAR|ImageNet|FP|BF|INT|NF|MXFP|W\d+A|RTX\s*|L|Top|Stage|Phase|Step|Layer|Block|Rank|Head|Seed|Part|Task|Case|Exp|Run|Fold|Bit|v|V)[-_/: ]?$|(?:(?<![A-Za-z])|\d\s*[x×])[AHV]$", re.I)
    meta_dur = None
    try:
        if (spec.get("meta") or {}).get("duration_minutes") is not None:
            meta_dur = float(spec["meta"]["duration_minutes"])
    except (TypeError, ValueError):
        meta_dur = None
    for i, s in enumerate(spec.get("slides", []), 1):
        role_s = str(s.get("role") or "").lower()
        texts = [s.get("title", ""), s.get("claim", "")]
        for b in s.get("bullets") or []:
            texts.append(b.get("text", "") if isinstance(b, dict) else str(b))
        fig = s.get("figure") or {}
        if isinstance(fig, dict) and fig.get("caption"):
            texts.append(fig["caption"])
        for f_it in (s.get("figures") or [])[:2]:
            if isinstance(f_it, dict) and f_it.get("caption"):
                texts.append(str(f_it["caption"]))
        tbl = s.get("table") or {}
        if isinstance(tbl, dict) and tbl:
            if tbl.get("caption"):
                texts.append(str(tbl["caption"]))
            for h in tbl.get("headers") or []:
                texts.append(str(h))
            for r in tbl.get("rows") or []:
                for cell in (r or []):
                    texts.append(str(cell))
        ch_spec = s.get("chart") or {}
        if isinstance(ch_spec, dict) and ch_spec:
            for fld in ("title", "caption", "y_label"):
                if ch_spec.get(fld):
                    texts.append(str(ch_spec[fld]))
            for cat in ch_spec.get("categories") or []:
                texts.append(str(cat))
            for ser in ch_spec.get("series") or []:
                if isinstance(ser, dict):
                    if ser.get("name"):
                        texts.append(str(ser["name"]))
                    for val in ser.get("values") or []:
                        texts.append(str(val))
        for k in s.get("kpis") or []:
            if isinstance(k, dict):
                for fld in ("value", "label", "delta"):
                    if k.get(fld):
                        texts.append(str(k[fld]))
        for st in s.get("steps") or []:
            if isinstance(st, dict):
                for fld in ("title", "detail"):
                    if st.get(fld):
                        texts.append(str(st[fld]))
        tc = s.get("two_col") or {}
        if isinstance(tc, dict):
            for fld in ("left_title", "right_title"):
                if tc.get(fld):
                    texts.append(str(tc[fld]))
            for col_k in ("left_bullets", "right_bullets"):
                for b in tc.get(col_k) or []:
                    texts.append(b.get("text", "") if isinstance(b, dict) else str(b))
        co = s.get("callout")
        if isinstance(co, str) and co.strip():
            texts.append(co)
        elif isinstance(co, dict):
            for fld in ("title", "label", "text", "body", "detail"):
                if co.get(fld):
                    texts.append(str(co[fld]))

        # Two-pass per slide: collect traceable source values on this slide first, then verify derived deltas/ratios
        slide_traced = []
        slide_pending = []
        for field_txt in texts:
            if not isinstance(field_txt, str):
                field_txt = str(field_txt)
            field_txt = normalize_sci_text(field_txt)
            for m in NUM_RE.finditer(field_txt):
                raw_tok = m.group(0)
                tok = raw_tok.replace(",", "")
                if YEAR_RE.match(tok):
                    continue
                if DATECTX_RE.search(field_txt[max(0, m.start() - 6): m.end() + 4]) and len(tok) <= 2:
                    continue
                after = field_txt[m.end(): m.end() + 14]
                before = field_txt[max(0, m.start() - 3): m.start()]
                before_w = field_txt[max(0, m.start() - 16): m.start()]
                has_metric_unit = bool(re.match(r"^\s*(%|pt\b|pts\b|pp\b|percentage\s+points?|百分点|ms\b|s\b|min\b|mins\b|minutes?\b|h\b|hr\b|hrs\b|×|x\b|GB\b|MB\b|TB\b|KB\b|k\b|K\b|M\b|B\b|thousand\b|million\b|billion\b|万|亿|倍|次)", after, re.I))
                has_signed = bool(re.search(r"[+±−-]\s*$", before))
                # Skip citation locators/ranges (Fig.2, Table 1, §3.2, p.3-5, Slide 4) and standard model/format/bitwidth identifiers (ResNet-50, Qwen2.5, Llama-3.1, FP16, INT8, 8-bit, 3D)
                if not has_metric_unit and (CITE_SLIDE_BEFORE_RE.search(before_w) or MODEL_ID_BEFORE_RE.search(before_w)):
                    continue
                if "." not in tok and not has_metric_unit and re.match(r"^\s*[- ]?(?:bit|bits|D\b|dim\b|way\b|shot\b|hop\b|fold\b|stage\b|step\b|layer\b|head\b|node\b|GPU\b|GPUs\b|slides?\b|页)", after, re.I):
                    continue
                # Skip talk duration metadata (e.g. "6 min" / "10 分钟") on cover/roadmap/thanks slides
                if role_s in ("title", "thanks", "roadmap") and meta_dur is not None and re.match(r"^\s*(?:min|mins|minute|minutes|分钟)\b", after, re.I):
                    try:
                        if abs(float(tok) - meta_dur) < 0.05:
                            continue
                    except ValueError:
                        pass
                # Skip small structural/ordinal integers (0..20) with no metric unit or decimal (e.g. "3 steps", "Top-1", "Phase 2")
                if "." not in tok and not has_metric_unit and not has_signed:
                    try:
                        if 0 <= int(tok) <= 20:
                            continue
                    except ValueError:
                        pass
                checked += 1
                try:
                    v = float(tok)
                except ValueError:
                    continue
                cand_vals = expand_units(v, after)
                is_approx_n = bool(re.search(r"[~≈∼约]\s*$", before))
                if tok in src_tokens or tok.lower() in src_tokens or any(close(cv, is_approx_n) for cv in cand_vals):
                    tid = ("s", i, len(slide_traced))
                    ukind = "pct" if re.match(r"^\s*%", after) else ("ratio" if re.match(r"^\s*(?:×|x\b|倍)", after) else "plain")
                    for cv in cand_vals:
                        slide_traced.append((cv, tid, ukind))
                    unit_m = re.match(r"^(\s*(?:%|pt|pts|pp|ms|s|min|GB|MB|TB|KB|×|x|k|K|M|B|倍))(?![A-Za-z])", after)
                    disp_tok = raw_tok + (unit_m.group(1).strip() if unit_m else "")
                    vt_list = slide_verified_tokens.setdefault(str(i), [])
                    if disp_tok not in vt_list and len(vt_list) < 12:
                        vt_list.append(disp_tok)
                    continue
                ctx = field_txt[max(0, m.start() - 18): m.end() + 12]
                wide_ctx = field_txt[max(0, m.start() - 40): m.end() + 18]
                is_pct = bool(re.match(r"^\s*%", after))
                is_ratio = bool(re.match(r"^\s*(?:×|x\b|倍)", after))
                kind_hint = "pct" if is_pct else ("ratio" if is_ratio else "delta")
                is_derived_ctx = bool(
                    has_signed
                    or is_pct or is_ratio
                    or re.match(r"^\s*(?:pt|pts|pp|percentage\s+points?|百分点)\b", after, re.I)
                    or re.search(r"\b(?:by|saves?|cuts?|trims?|reduces?|drops?|gains?|loses?|faster|slower|lower|higher)\b|减少|降低|节省|提升|缩小|百分点", wide_ctx, re.I)
                )
                slide_pending.append((raw_tok, tok, v, ctx, is_derived_ctx, kind_hint))

        slide_traced_by_slide[i] = list(slide_traced)

        for raw_tok, tok, v, ctx, is_derived_ctx, kind_hint in slide_pending:
            entry = {"slide": i, "value": raw_tok, "context": ctx}
            if tok in allow:
                entry["reason"] = allow[tok]
                justified.append(entry)
                continue
            if is_derived_ctx:
                derived_why = check_derived_math(v, slide_traced, kind_hint)
                if derived_why:
                    entry["reason"] = derived_why
                    justified.append(entry)
                    deck_verified_derived.setdefault(tok, f"{derived_why} (on s{i})")
                    continue
                pending_cross_slide.append((tok, entry))
                continue
            suspects.append(entry)

    # Carry over derived metrics that were mathematically verified on a content slide (e.g. reused on Cover/Conclusion)
    for tok, entry in pending_cross_slide:
        if tok in deck_verified_derived:
            entry["reason"] = f"consistent with {deck_verified_derived[tok]}"
            justified.append(entry)
        else:
            suspects.append(entry)

    # Advisory Q&A numeric caliber check: catch hallucinated numbers inside qa_prep[].answer_key
    qa_items = []
    if a.qa and Path(a.qa).exists():
        try:
            qa_items = json.loads(Path(a.qa).read_text(encoding="utf-8"))
        except Exception:
            qa_items = []
    elif isinstance(spec.get("qa_prep"), list):
        qa_items = spec.get("qa_prep")
    qa_suspects = []
    CITE_BEFORE_RE = re.compile(r"(?:Fig(?:ure)?\.?|Tab(?:le)?\.?|Sec(?:tion)?\.?|Eq(?:uation)?\.?|Alg(?:orithm)?\.?|Slide|p\.|§|P|图|表|第|公式|算法|原P)\s*\(?$", re.I)
    for qi, qobj in enumerate(qa_items if isinstance(qa_items, list) else [], 1):
        if not isinstance(qobj, dict):
            continue
        ans = normalize_sci_text(str(qobj.get("answer_key") or ""))
        cite_sl = qobj.get("cite_slide")
        qa_traced = list(slide_traced_by_slide.get(cite_sl, [])) if isinstance(cite_sl, int) else []
        qa_pending = []
        for m in NUM_RE.finditer(ans):
            raw_tok = m.group(0)
            tok = raw_tok.replace(",", "")
            if YEAR_RE.match(tok):
                continue
            after = ans[m.end(): m.end() + 14]
            before_w = ans[max(0, m.start() - 14): m.start()]
            before = ans[max(0, m.start() - 3): m.start()]
            if CITE_BEFORE_RE.search(before_w) or MODEL_ID_BEFORE_RE.search(before_w):
                continue
            has_unit = bool(re.match(r"^\s*(%|pt\b|pts\b|pp\b|百分点|ms\b|s\b|min\b|mins\b|×|x\b|GB\b|MB\b|KB\b|TB\b|k\b|K\b|M\b|B\b|thousand\b|million\b|billion\b|万|亿|倍)", after, re.I))
            has_sign = bool(re.search(r"[+±−-]\s*$", before))
            is_q_approx = bool(re.search(r"(?:[~≈]|约|about|approx\.?|roughly|nearly)\s*$", before_w, re.I))
            if "." not in tok and "e" not in tok.lower() and not has_unit and not has_sign:
                try:
                    if 0 <= int(tok) <= 20:
                        continue
                except ValueError:
                    pass
            try:
                qv = float(tok)
            except ValueError:
                continue
            qcands = expand_units(qv, after)
            if tok in src_tokens or tok.lower() in src_tokens or tok in allow or tok in deck_verified_derived or any(close(cv, is_approx=is_q_approx) for cv in qcands):
                qtid = ("q", qi, len(qa_traced))
                qukind = "pct" if re.match(r"^\s*%", after) else ("ratio" if re.match(r"^\s*(?:×|x\b|倍)", after) else "plain")
                for cv in qcands:
                    qa_traced.append((cv, qtid, qukind))
                continue
            qctx = ans[max(0, m.start() - 16): m.end() + 14]
            q_hint = "pct" if re.match(r"^\s*%", after) else ("ratio" if re.match(r"^\s*(?:×|x\b|倍)", after) else "delta")
            qa_pending.append((raw_tok, tok, qv, qctx, q_hint))
        for raw_tok, tok, qv, qctx, q_hint in qa_pending:
            if check_derived_math(qv, qa_traced, q_hint):
                continue
            qa_suspects.append({"qa": qi, "value": raw_tok, "context": qctx})

    rep = {"checked": checked, "n_suspects": len(suspects), "suspects": suspects[:50],
           "justified": justified[:50], "qa_suspects": qa_suspects[:20],
           "traced_by_slide": slide_verified_tokens}
    if a.out:
        Path(a.out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"number check: {checked} numbers checked, {len(suspects)} suspect(s), {len(justified)} justified"
          + (f", {len(qa_suspects)} qa-advisory" if qa_suspects else ""))
    for s in suspects:
        print(f"  [s{s['slide']}] {s['value']!r} not traceable :: ...{s['context']}...")
    if suspects:
        print("  FIX: replace with exact value from paper.txt, OR if derived/read from chart add {\"value\":\"...\",\"reason\":\"...\"} to deck_spec.json top-level \"number_allow\"")
    for s in justified:
        print(f"  [s{s['slide']}] {s['value']!r} justified: {s.get('reason', '')[:60]}")
    for qs in qa_suspects[:6]:
        print(f"  ~ (qa-advisory) [qa#{qs['qa']}] {qs['value']!r} not in source :: ...{qs['context']}...")
    sys.exit(1 if suspects else 0)


if __name__ == "__main__":
    main()
