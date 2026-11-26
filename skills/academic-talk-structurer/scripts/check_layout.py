#!/usr/bin/env python3
"""S4 static lint for a PPTX: out-of-bounds, text overflow estimate, shape overlap, picture distortion.

Usage:
  python3 check_layout.py talk.pptx [--json /app/work/lint_report.json] [--plan /app/work/slide_plan.json]

Geometry lint + optional plan-level notes/timing lint (--plan):
  Bilingual speaking-rate units = CJK chars + 1.5 * Latin words.
  notes_too_short: units < 40; notes_overrun: units > 4.8 * time_s; time_missing on content slides.
Exit: 0 clean, 1 issues found (list printed), 3 fatal (python-pptx missing / file unreadable).
No-script fallback: open the rendered PNG previews and inspect visually (SKILL.md S5).
"""
import argparse
import json
import math
import re
import sys
from pathlib import Path

EMU_IN = 914400.0


def char_units(s):
    u = 0.0
    for c in s:
        u += 1.0 if "\u2e80" <= c <= "\uffef" or "\u3000" <= c <= "\u303f" else 0.55
    return u


def speaking_units(s):
    """Language-aware speaking length: 1 CJK char ~= 1 unit; 1 English word ~= 1.5 units."""
    if not s:
        return 0.0
    cjk = sum(1 for c in s if "\u2e80" <= c <= "\uffef")
    latin_words = len(re.findall(r"[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*", s))
    return float(cjk) + latin_words * 1.5


def master_font_sizes(prs):
    """(title_pt, body_pt) parsed from master txStyles; fallback (36, 20)."""
    try:
        from pptx.oxml.ns import qn
        m = prs.slide_masters[0].element
        ts = m.find(qn("p:txStyles"))

        def sz_of(tag):
            if ts is None:
                return None
            el = ts.find(qn(tag))
            if el is None:
                return None
            lvl1 = el.find(qn("a:lvl1pPr"))
            if lvl1 is None:
                return None
            dr = lvl1.find(qn("a:defRPr"))
            if dr is None:
                return None
            sz = dr.get("sz")
            return float(sz) / 100.0 if sz else None

        return (sz_of("p:titleStyle") or 36.0, sz_of("p:bodyStyle") or 20.0)
    except Exception:
        return 36.0, 20.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--plan", dest="plan")
    a = ap.parse_args()
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    except ImportError:
        print("FATAL: python-pptx missing", file=sys.stderr)
        sys.exit(3)
    try:
        prs = Presentation(a.pptx)
    except Exception as e:
        print(f"FATAL: cannot open {a.pptx}: {e!r}", file=sys.stderr)
        sys.exit(3)
    SW, SH = prs.slide_width, prs.slide_height
    title_sz, body_sz = master_font_sizes(prs)
    TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
    RAW_LATEX_RE = re.compile(r"\\(?:frac|sum|prod|int|sqrt|mathbb|mathcal|begin|end|approx|leq|geq|times|cdot|alpha|beta|gamma|lambda|sigma)\b")
    GENERIC_TITLE_RE = re.compile(
        r"^(?:\d+\.?\s*)?(?:introduction|background|related\s+work|method(?:ology)?|experiments?|results?|discussion|limitations?|conclusion(?:s)?|overview|outline|引言|背景|相关工作|方法|实验|结果|讨论|局限|结论|目录)$",
        re.IGNORECASE,
    )
    issues, warnings, low_res_figs, symbol_warns = [], [], [], []

    for si, slide in enumerate(prs.slides, 1):
        boxes = []
        slide_shapes = list(slide.shapes)
        slide_area = float(SW * SH) if (SW and SH) else 0.0
        for sh_idx, sh in enumerate(slide_shapes):
            l, t, w, h = sh.left, sh.top, sh.width, sh.height
            if None in (l, t, w, h):
                continue
            st = str(sh.shape_type)
            txt = sh.text_frame.text.strip() if sh.has_text_frame else ""
            is_table = bool(getattr(sh, "has_table", False))
            is_chart = bool(getattr(sh, "has_chart", False))
            is_pic = (sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
            is_grp = (sh.shape_type == MSO_SHAPE_TYPE.GROUP)
            area_ratio = (float(w) * float(h)) / slide_area if slide_area > 0 else 0.0
            # Full-bleed underlying background picture (>=65% slide area sitting beneath foreground content) is exempt from overlap/OOB
            is_underlying_bg_pic = False
            if is_pic and area_ratio >= 0.65:
                has_content_above = any(
                    (
                        (o.has_text_frame and o.text_frame.text.strip())
                        or getattr(o, "has_table", False)
                        or getattr(o, "has_chart", False)
                        or (o.shape_type == MSO_SHAPE_TYPE.PICTURE and None not in (o.width, o.height) and (float(o.width) * float(o.height)) < 0.65 * slide_area)
                    )
                    for o in slide_shapes[sh_idx + 1:]
                )
                has_text_below = any((o.has_text_frame and o.text_frame.text.strip()) for o in slide_shapes[:sh_idx])
                is_underlying_bg_pic = has_content_above or (not has_text_below)
            is_bg_like = (
                (w > SW * 0.88 and h > SH * 0.80)
                or is_underlying_bg_pic
                or (t <= int(0.35 * EMU_IN) and w >= SW * 0.80 and h <= int(1.35 * EMU_IN))
                or (t >= SH - int(0.85 * EMU_IN) and w >= SW * 0.80 and h <= int(0.90 * EMU_IN))
                or ((l <= int(0.35 * EMU_IN) or l + w >= SW - int(0.35 * EMU_IN)) and w <= int(0.70 * EMU_IN) and h >= SH * 0.65)
            )
            is_tiny_chrome = (w < int(0.25 * EMU_IN) or h < int(0.20 * EMU_IN))
            is_content = bool(txt or is_table or is_chart or ((is_pic or is_grp) and not is_bg_like and not is_tiny_chrome))
            # Only flag genuine out-of-bounds (>0.10in) on content shapes, ignoring full-bleed template chrome
            edge_tol = int(0.10 * EMU_IN)
            if is_content and (l < -edge_tol or t < -edge_tol or l + w > SW + edge_tol or t + h > SH + edge_tol):
                issues.append({"slide": si, "type": "out_of_bounds", "shape": st,
                               "detail": f"L{l/EMU_IN:.2f} T{t/EMU_IN:.2f} W{w/EMU_IN:.2f} H{h/EMU_IN:.2f} slide {SW/EMU_IN:.2f}x{SH/EMU_IN:.2f}"})
            if is_content:
                label = txt if txt else ("table" if is_table else ("chart" if is_chart else ("group" if is_grp else "pic")))
                boxes.append((l, t, w, h, label, st))
            if txt:
                if any(ch == "\ufffd" or "\ue000" <= ch <= "\uf8ff" for ch in txt):
                    msg = f"s{si}: garbled/PUA replacement glyph in slide text ({txt[:36]!r})"
                    symbol_warns.append(msg)
                    warnings.append(msg)
                elif RAW_LATEX_RE.search(txt):
                    msg = f"s{si}: raw LaTeX command visible on slide ({RAW_LATEX_RE.search(txt).group(0)}) — verbalize or use clean Unicode"
                    symbol_warns.append(msg)
                    warnings.append(msg)
            if txt and sh.has_text_frame:
                tf = sh.text_frame
                ml = tf.margin_left if tf.margin_left is not None else 91440
                mr = tf.margin_right if tf.margin_right is not None else 91440
                mt = tf.margin_top if tf.margin_top is not None else 45720
                mb = tf.margin_bottom if tf.margin_bottom is not None else 45720
                usable_w_pt = max(10.0, (w - ml - mr) / EMU_IN * 72.0)
                usable_h_pt = max(10.0, (h - mt - mb) / EMU_IN * 72.0)
                need_pt, assumed = 0.0, False
                is_title = bool(sh.is_placeholder and sh.placeholder_format.type in TITLE_TYPES)
                def_sz = title_sz if is_title else body_sz
                for p in tf.paragraphs:
                    sz = None
                    for r in p.runs:
                        if r.font.size:
                            sz = r.font.size.pt
                            break
                    if sz is None and p.font.size:
                        sz = p.font.size.pt
                    if sz is None:
                        sz = def_sz
                        assumed = True
                    units = char_units(p.text)
                    if units <= 0:
                        continue
                    lines = max(1, math.ceil(units * sz / usable_w_pt))
                    need_pt += lines * sz * 1.22
                if need_pt > usable_h_pt * 1.16 + 4:
                    if si == 1 and assumed:
                        warnings.append(f"s1: cover text height estimate ({need_pt:.0f}pt vs {usable_h_pt:.0f}pt, master size assumed) — verify visually")
                    else:
                        issues.append({"slide": si, "type": "text_overflow", "shape": st,
                                       "detail": f"need~{need_pt:.0f}pt vs box {usable_h_pt:.0f}pt :: {txt[:40]!r}" +
                                                 (" (sizes assumed: set explicit font sizes)" if assumed else "")})
            if is_pic and not is_bg_like:
                try:
                    nat_w, nat_h = sh.image.size
                    disp = (w / h) if h else 0
                    nat = nat_w / nat_h if nat_h else 0
                    if nat and disp and abs(disp - nat) / nat > 0.12:
                        issues.append({"slide": si, "type": "picture_distortion", "shape": st,
                                       "detail": f"aspect {disp:.2f} vs native {nat:.2f}"})
                    disp_in = (w / EMU_IN)
                    if disp_in > 1.0:
                        eff_dpi = nat_w / disp_in
                        if eff_dpi < 90:
                            low_res_figs.append({"slide": si, "px": [nat_w, nat_h], "dpi": round(eff_dpi, 1)})
                            sev = "very low-res" if eff_dpi < 45 else "low-res"
                            warnings.append(f"s{si}: {sev} picture ({nat_w}px over {disp_in:.1f}in = {eff_dpi:.0f} DPI) — prefer higher-res source or native table/chart")
                except Exception:
                    pass
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                l1, t1, w1, h1, x1, s1 = boxes[i]
                l2, t2, w2, h2, x2, s2 = boxes[j]
                ix = max(0, min(l1 + w1, l2 + w2) - max(l1, l2))
                iy = max(0, min(t1 + h1, t2 + h2) - max(t1, t2))
                inter = ix * iy
                denom = min(w1 * h1, w2 * h2) or 1
                if inter / denom > 0.22 and inter > 60000:
                    issues.append({"slide": si, "type": "overlap", "shape": f"{s1} x {s2}",
                                   "detail": f"{x1[:24]!r} vs {x2[:24]!r} inter={inter/denom:.0%}"})

    rep = {
        "file": a.pptx, "n_issues": len(issues), "issues": issues, "warnings": warnings,
        "low_res_figures": low_res_figs, "symbol_warnings": symbol_warns,
    }
    if a.plan:
        try:
            plan = json.loads(Path(a.plan).read_text(encoding="utf-8"))
            p_meta = plan.get("meta") or {}
            if not str(p_meta.get("goal") or "").strip():
                warnings.append("meta.goal is empty; state a verifiable presentation objective for editorial audit")
            tr_list = p_meta.get("tradeoffs") or []
            if len(tr_list) < 2:
                warnings.append(f"meta.tradeoffs has {len(tr_list)} item(s) (<2); document explicit keep/cut/reorder decisions for editorial audit")
            total_timed_s = 0.0
            seen_titles = {}
            s3_overclaims = 0
            visual_slide_count = 0
            content_slide_count = 0
            is_scanned_pdf = False
            ing_p = Path(a.plan).parent / "ingest_summary.json"
            if ing_p.exists():
                try:
                    is_scanned_pdf = bool(json.loads(ing_p.read_text(encoding="utf-8")).get("likely_scanned"))
                except Exception:
                    pass
            for s in plan.get("slides", []):
                role = s.get("role", "")
                pno_s = s.get("pptx_slide")
                fig_p_str = str(s.get("figure") or "")
                if fig_p_str and not is_scanned_pdf and re.search(r"(?:^|[/\\])page_p\d+\.png$", fig_p_str, re.I):
                    issues.append({"slide": pno_s, "type": "uncropped_page_render", "shape": "figure",
                                   "detail": f"uncropped full-page PDF screenshot ({Path(fig_p_str).name}) embedded; use a cropped figure or native table/chart instead"})
                stitle = str(s.get("title") or "").strip()
                if stitle and role not in ("preserved",):
                    tk = stitle.lower()
                    if tk in seen_titles:
                        warnings.append(f"s{pno_s}: duplicate slide title with s{seen_titles[tk]} ({stitle[:36]!r})")
                    else:
                        seen_titles[tk] = pno_s
                if role == "qa_backup" and s.get("time_s"):
                    warnings.append(f"s{pno_s}: qa_backup slide should have time_s=null (excluded from main talk budget)")
                if role in ("title", "thanks", "qa_backup", "preserved"):
                    if role == "title" and s.get("time_s"):
                        total_timed_s += float(s["time_s"])
                    continue
                content_slide_count += 1
                if s.get("visuals") or s.get("figure") or s.get("has_table"):
                    visual_slide_count += 1
                if s.get("time_s"):
                    total_timed_s += float(s["time_s"])
                if role not in ("roadmap", "conclusion") and stitle and GENERIC_TITLE_RE.match(stitle):
                    warnings.append(f"s{pno_s}: generic section title {stitle!r}; rewrite as a speakable takeaway")
                full_notes = s.get("notes_full") or ""
                if s.get("notes_units") is not None:
                    nu = float(s["notes_units"])
                elif full_notes:
                    nu = speaking_units(full_notes)
                else:
                    nu = float(s.get("notes_chars") or 0)
                ts = s.get("time_s")
                if nu < 35:
                    issues.append({"slide": s.get("pptx_slide"), "type": "notes_too_short",
                                   "shape": "notes", "detail": f"{nu:.0f} speaking units < 35; write a speakable script"})
                if ts is None:
                    issues.append({"slide": s.get("pptx_slide"), "type": "time_missing",
                                   "shape": "plan", "detail": "content slide without time_s"})
                elif nu > 5.2 * float(ts):
                    issues.append({"slide": s.get("pptx_slide"), "type": "notes_overrun",
                                   "shape": "notes", "detail": f"{nu:.0f} speaking units > 5.2x{ts}s budget; trim script"})
                # Advisory: detect verbatim bullet copying into notes + wall-of-text density + claim bloat
                b_texts = [b.get("text", "") if isinstance(b, dict) else str(b) for b in (s.get("bullets") or [])]
                if len(b_texts) > 5:
                    warnings.append(f"s{s.get('pptx_slide')}: {len(b_texts)} bullets (>5); condense to <=4 crisp points")
                body_units = sum(speaking_units(bt) for bt in b_texts)
                if body_units > 180:
                    warnings.append(f"s{s.get('pptx_slide')}: dense slide bullets ({body_units:.0f} units > 180); move prose detail to speaker notes")
                b_chunks = [c.strip() for c in b_texts if len(c.strip()) >= 12]
                if b_chunks and full_notes:
                    copied = sum(1 for c in b_chunks if c[:18].lower() in full_notes.lower())
                    if copied / len(b_chunks) >= 0.70 and len(b_chunks) >= 2:
                        warnings.append(f"s{s.get('pptx_slide')}: notes repeat slide bullets verbatim ({copied}/{len(b_chunks)}); rewrite as oral delivery")
                # Advisory: detect title/claim duplication & evidence-strength calibration
                evs = s.get("evidence") or []
                strengths = {str(e.get("strength", "")).upper() for e in evs if e.get("strength")}
                claim_txt = str(s.get("claim") or "").strip()
                if stitle and claim_txt:
                    nt = re.sub(r"[\W_]+", "", stitle.lower())
                    nc = re.sub(r"[\W_]+", "", claim_txt.lower())
                    if len(nt) >= 12 and len(nc) >= 12 and (nt == nc or (min(len(nt), len(nc)) / max(len(nt), len(nc)) >= 0.85 and (nt in nc or nc in nt))):
                        warnings.append(f"s{pno_s}: title and claim are nearly identical; make title the headline and claim the specific evidence-backed takeaway")
                if claim_txt and speaking_units(claim_txt) > 45:
                    warnings.append(f"s{s.get('pptx_slide')}: claim is lengthy ({speaking_units(claim_txt):.0f} units); tighten to 1 crisp sentence (<=30 CJK / <=18 EN words)")
                if strengths and strengths <= {"S3", "NA"} and re.search(
                    r"(显著优于|严格证明|首次突破|大幅领先|proves|conclusively|significantly outperforms|definitively)", claim_txt, re.I
                ):
                    s3_overclaims += 1
                    warnings.append(f"s{s.get('pptx_slide')}: claim wording sounds S1-strong while evidence is {sorted(strengths)}; calibrate verb strength")
                if strengths == {"S1"} and not re.search(r"\d", claim_txt + " " + " ".join(b_texts)) and not (s.get("figure") or s.get("has_table")):
                    warnings.append(f"s{pno_s}: evidence tagged S1 (quantitative+baseline) but slide shows no numeric metric, figure, or table — verify strength tag")
                tail = (s.get("notes_tail") or "")
                if tail and not any(k in tail.lower() for k in (
                    "下面", "接下来", "然后", "最后", "综上", "这引出", "下一", "因此", "总结", "谢谢",
                    "由此", "进一步", "回到", "展望", "小结", "问答", "欢迎", "请看",
                    "next", "this leads", "leads us", "brings us", "finally", "turn to", "let's look",
                    "let us", "now let", "moving on", "in summary", "to wrap up", "takeaway", "question", "thank"
                )):
                    warnings.append(f"s{s.get('pptx_slide')}: notes may lack a bridge sentence at the end (advisory)")
                if role in ("method", "result", "analysis", "setup"):
                    if full_notes and not any(k in full_notes.lower() for k in (
                        "图", "表", "右", "左", "如下", "曲线", "数据", "流程", "步骤", "对比", "卡片",
                        "fig", "table", "chart", "plot", "right", "left", "shown", "panel", "bar", "diagram", "flow", "pipeline", "stage", "step", "card"
                    )):
                        warnings.append(f"s{s.get('pptx_slide')}: evidence-page notes may not point at the on-page figure/table (advisory)")
            if content_slide_count >= 4 and visual_slide_count == 0:
                warnings.append("deck has 0 visual components across content slides; use figure/table/steps/kpis/two_col to break bullet walls")
            rep["s3_overclaim_warnings"] = s3_overclaims
            dur_min = p_meta.get("duration_minutes")
            if dur_min and total_timed_s > 0:
                budget_s = float(dur_min) * 60.0
                if budget_s > 0 and abs(total_timed_s - budget_s) / budget_s > 0.22:
                    warnings.append(f"timing: total timed slides = {total_timed_s:.0f}s vs budget {budget_s:.0f}s ({dur_min}min); align slide time_s")
        except Exception as e:
            warnings.append(f"plan lint skipped: {e!r}")
        rep["n_issues"] = len(issues)
        rep["n_warnings"] = len(warnings)
    if a.json_out:
        Path(a.json_out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"lint {a.pptx}: {len(issues)} issue(s), {len(warnings)} warning(s)")
    for it in issues:
        print(f"  [s{it['slide']}] {it['type']}: {it['detail'][:140]}")
    for w in warnings[:8]:
        print(f"  ~ {w[:140]}")
    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
