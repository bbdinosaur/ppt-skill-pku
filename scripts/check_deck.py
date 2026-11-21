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

EXEMPT_ROLES = {"title", "section", "agenda", "qa", "toc", "existing"}


def corpus_from_material(inv):
    """Fallback corpus: all string values from material.json (used when no PDF text)."""
    import json as _json
    p = inv.get("material_json_path")
    if not p or not Path(p).is_file():
        return ""
    try:
        obj = _json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return ""
    buf = []
    def walk(o):
        if isinstance(o, dict):
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
        elif isinstance(o, str):
            buf.append(o)
    walk(obj)
    return "\n".join(buf)


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def check_text_fit(prs):
    issues = []
    for si, slide in enumerate(prs.slides, 1):
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


def check_bounds(prs):
    issues = []
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    tol = int(min(SW, SH) * 0.01)
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            l, t, w, h = du.shape_geom(prs, slide, sh)
            if l is None or t is None or w is None or h is None:
                continue
            l, t, r, b = int(l), int(t), int(l) + int(w), int(t) + int(h)
            if l < -tol or t < -tol or r > SW + tol or b > SH + tol:
                issues.append(f"slide {si} '{sh.name}' out of bounds l={l} t={t} r={r} b={b} (slide {SW}x{SH})")
    return issues


def check_overlap(prs):
    issues = []
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    for si, slide in enumerate(prs.slides, 1):
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
            has_text = sh.has_text_frame and sh.text_frame.text.strip()
            is_pic = sh.shape_type is not None and "PICTURE" in str(sh.shape_type)
            if sh.is_placeholder:
                ph_t = sh.placeholder_format.type
                if ph_t in du.TITLE_TYPES and not has_text:
                    continue
                if ph_t in du.BODY_TYPES and not has_text:
                    continue  # empty placeholder renders nothing
            if not (has_text or is_pic):
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


def check_aspect(prs):
    issues = []
    for si, slide in enumerate(prs.slides, 1):
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


def check_notes_titles(prs, content_map):
    issues = []
    cm = {s.get("slide"): s for s in (content_map or {}).get("slides", [])}
    for si, slide in enumerate(prs.slides, 1):
        tph = du.slide_title_ph(slide)
        has_title = tph is not None and tph.text_frame.text.strip()
        if not has_title:
            spec = cm.get(si) or {}
            if spec.get("role") not in ("qa",):
                issues.append(f"slide {si}: no title text")
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        if not notes:
            issues.append(f"slide {si}: missing speaker notes")
    return issues


def check_table_fit(prs):
    """Tables auto-grow downward in PowerPoint: estimate wrapped-cell height vs frame."""
    issues = []
    for si, slide in enumerate(prs.slides, 1):
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
                for ci, cell in enumerate(row.cells):
                    t = cell.text
                    size = 12.0
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            if r.font.size is not None:
                                size = r.font.size.pt
                                break
                    max_lines = max(max_lines, du.est_lines(t, size, max(col_ws[ci] - 12, 10)) if ci < ncols else 1)
                total += max_lines * 12.0 * 1.3 + 4
            h_pt = sh.height.pt if sh.height else 0
            if total > h_pt * 1.18 + 6:
                issues.append(f"slide {si} '{sh.name}': est table height {total:.0f}pt > frame {h_pt:.0f}pt")
    return issues


def check_min_font(prs, min_pt):
    issues = []
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not (sh.has_text_frame and sh.text_frame.text.strip()):
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is not None and r.font.size.pt < min_pt and r.text.strip():
                        issues.append(f"slide {si} '{sh.name}': font {r.font.size.pt}pt < {min_pt}pt")
    return issues


def check_font_overrides(prs):
    issues = []
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name is not None:
                        issues.append(f"slide {si} '{sh.name}': explicit font '{r.font.name}' overrides template")
    return issues


NUM_RE = re.compile(r"\d+\.\d+|\d{3,}")


def check_numbers(content_map, material_text_file, fallback_corpus=""):
    """Number fidelity: substantive numbers on slides (decimals or 3+ digits) must be
    findable in the material corpus — catches copy errors and rounding-by-memory."""
    issues = []
    checked = 0
    if not content_map:
        return issues, checked
    corpus = ""
    if material_text_file and Path(material_text_file).is_file():
        corpus = norm(Path(material_text_file).read_text(encoding="utf-8", errors="replace"))
    if not corpus:
        corpus = norm(fallback_corpus)
    if not corpus:
        return issues, checked
    for s in (content_map.get("slides") or []):
        texts = [s.get("title") or ""] + [b if isinstance(b, str) else b.get("t", "")
                                            for b in (s.get("bullets") or [])]
        tbl = s.get("table") or {}
        for row in (tbl.get("rows") or []):
            texts.extend(str(c) for c in row)
        img = s.get("image") or {}
        if img.get("caption"):
            texts.append(img["caption"])
        ch = s.get("chart") or {}
        texts.extend(str(c) for c in (ch.get("categories") or []))
        for sr in (ch.get("series") or []):
            texts.append(str(sr.get("name", "")))
            texts.extend(str(v) for v in (sr.get("values") or []))
        seen = set()
        for t in texts:
            for m in NUM_RE.findall(str(t)):
                if m in seen:
                    continue
                seen.add(m)
                checked += 1
                if m not in corpus:
                    issues.append(f"slide {s.get('slide')}: number '{m}' not found in material")
    return issues, checked


def check_quotes(content_map, material_text_file, fallback_corpus=""):
    """Verify on-slide claim quotes really exist in the source material (anti-hallucination)."""
    issues = []
    checked = 0
    if not content_map:
        return issues, checked
    text = ""
    if material_text_file and Path(material_text_file).is_file():
        text = norm(Path(material_text_file).read_text(encoding="utf-8", errors="replace"))
    if not text:
        text = norm(fallback_corpus)
    for s in (content_map.get("slides") or []):
        for src in (s.get("sources") or []):
            q = norm(src.get("quote", ""))
            if len(q) < 6:
                continue
            checked += 1
            if text and q in text:
                continue
            issues.append(f"slide {s.get('slide')}: quote not found in material: '{src.get('quote', '')[:50]}...'")
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
            import re as _re
            m = _re.search(r"\d+", str(v))
            return int(m.group()) if m else None
    expected = args.expected_count or parse_count(wo.get("slide_count"))
    max_count = args.max_count or parse_count(wo.get("max_slides"))

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

    add("slide_count", True if not expected else n == int(expected),
        f"{n} slides" + (f" == {expected}" if expected else " (no workorder count)"))
    if max_count:
        add("slide_limit", n <= int(max_count), f"{n} <= {max_count}")

    b = check_bounds(prs)
    add("bounds", not b, f"{len(b)} violations" + (": " + "; ".join(b[:6]) if b else ""))
    ov = check_text_fit(prs)
    add("text_fit", not ov, f"{len(ov)} potential overflows" + (": " + "; ".join(ov[:6]) if ov else ""))
    ol = check_overlap(prs)
    add("overlap", not ol, f"{len(ol)} overlaps" + (": " + "; ".join(ol[:6]) if ol else ""))
    asp = check_aspect(prs)
    add("aspect", not asp, f"{len(asp)} aspect changes" + (": " + "; ".join(asp[:4]) if asp else ""))
    nt = check_notes_titles(prs, content_map)
    add("titles_notes", not nt, f"{len(nt)} missing" + (": " + "; ".join(nt[:6]) if nt else ""))
    mf = check_min_font(prs, args.min_font)
    add("min_font", not mf, f"{len(mf)} violations" + (": " + "; ".join(mf[:4]) if mf else ""))
    tbl = check_table_fit(prs)
    add("table_fit", not tbl, f"{len(tbl)} potential table overflows" + (": " + "; ".join(tbl[:4]) if tbl else ""))
    fo = check_font_overrides(prs)
    add("font_inheritance", not fo, f"{len(fo)} overrides" + (": " + "; ".join(fo[:4]) if fo else ""))

    # evidence: every content slide cites >=1 source; quotes verified against material text
    mt_file = (inv.get("pdf_extract") or {}).get("text_file")
    ev_issues = []
    cm_slides = content_map.get("slides") or []
    for s in cm_slides:
        if (s.get("role") or "") in EXEMPT_ROLES:
            continue
        if not (s.get("sources") or []):
            ev_issues.append(f"slide {s.get('slide')} ({s.get('role')}): no claim_source")
    qi, qchecked = check_quotes(content_map, mt_file, corpus_from_material(inv))
    add("claim_sources", not ev_issues, f"{len(ev_issues)} slides without sources" + (": " + "; ".join(ev_issues[:6]) if ev_issues else ""))
    add("quote_verification", not qi, f"{qchecked} quotes checked, {len(qi)} unmatched" + (": " + "; ".join(qi[:6]) if qi else ""))
    ni, nchecked = check_numbers(content_map, mt_file, corpus_from_material(inv))
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
