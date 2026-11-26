#!/usr/bin/env python3
"""Static geometry/typography checker for a generated PPTX.

Usage:
  python3 check_layout.py deck.pptx [--report /app/work/layout_report.json]

Checks per slide:
  1. out-of-bounds shapes (beyond slide rect, tolerance 0.03in)
  2. estimated text overflow (paragraph text vs frame height, CJK-aware widths)
  3. picture/table vs non-empty text overlap (>8% of the smaller area)
  4. tiny fonts (<9pt anywhere)
  5. dense slides (>10 paragraphs on one slide)
  6. table overflow & tiny table cell fonts
Exit code 0 always; severity is in the report: "ok" | "warn" | "fail".
"""
import argparse
import json
import os
import re
import sys

from pptx import Presentation

EMU_IN = 914400.0
TOL = 0.03
REMNANT_RE = re.compile(
    r"(?i)(?:click\s+to\s+add|lorem\s+ipsum|insert\s+title|subtitle\s+here|单击此处添加|点击添加)"
)


def est_text_width(text, pt):
    w = 0.0
    for ch in text:
        o = ord(ch)
        if o > 0x2E7F:
            w += 1.02
        elif ch in " ilj.,:;'|!":
            w += 0.30
        elif ch.isupper() or ch.isdigit():
            w += 0.62
        else:
            w += 0.52
    return w * pt / 72.0


def est_lines(text, pt, width_in):
    if width_in <= 0.2:
        return 99
    total = 0
    for para in text.split("\n"):
        if not para.strip():
            total += 1
            continue
        tokens = re.findall(r"\S+\s*", para)
        cur, lines = 0.0, 1
        for tok in tokens:
            wl = est_text_width(tok, pt)
            if wl > width_in:
                if cur > 0:
                    lines += 1
                    cur = 0.0
                extra = max(1, int(wl / width_in))
                lines += extra
                cur = wl - extra * width_in
            elif cur + wl > width_in and cur > 0:
                lines += 1
                cur = wl
            else:
                cur += wl
        total += lines
    return total


def frame_text_info(shape):
    if not shape.has_text_frame:
        return None
    tf = shape.text_frame
    w = shape.width / EMU_IN
    h = shape.height / EMU_IN
    inner_w = max(w - (tf.margin_left + tf.margin_right) / EMU_IN, 0.2)
    inner_h = max(h - (tf.margin_top + tf.margin_bottom) / EMU_IN, 0.15)
    total_h = 0.0
    max_pt, min_pt, n_para = 0.0, 99.0, 0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        if not txt.strip():
            continue
        n_para += 1
        pt = None
        for r in p.runs:
            if r.font.size:
                pt = r.font.size.pt
                break
        if pt is None:
            try:
                pt = p.font.size.pt
            except Exception:
                pt = 18.0
        pt = pt or 18.0
        max_pt = max(max_pt, pt)
        min_pt = min(min_pt, pt)
        lvl_indent = 0.35 * (p.level or 0)
        line_mult = 1.20 if pt < 12.0 else 1.28
        total_h += est_lines(txt, pt, max(inner_w - lvl_indent, 0.5)) * line_mult * pt / 72.0
    return {"w": w, "h": h, "inner_w": inner_w, "inner_h": inner_h,
            "est_h": total_h, "max_pt": max_pt,
            "min_pt": (min_pt if n_para else 0.0), "n_para": n_para}


def bbox_of(shape):
    try:
        return (shape.left / EMU_IN, shape.top / EMU_IN,
                (shape.left + shape.width) / EMU_IN,
                (shape.top + shape.height) / EMU_IN)
    except Exception:
        return None


def inter_area(a, b):
    x0, y0 = max(a[0], b[0]), max(a[1], b[1])
    x1, y1 = min(a[2], b[2]), min(a[3], b[3])
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--report", default="/app/work/layout_report.json")
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    SW, SH = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
    issues = []
    stats = {"slides": len(prs.slides._sldIdLst),
             "slide_w": round(SW, 2), "slide_h": round(SH, 2)}

    for si, slide in enumerate(prs.slides, 1):
        boxes = []
        for shape in slide.shapes:
            bb = bbox_of(shape)
            if bb is None:
                continue
            name = getattr(shape, "name", "?")
            has_tf = getattr(shape, "has_text_frame", False)
            has_tbl = getattr(shape, "has_table", False)
            has_cht = getattr(shape, "has_chart", False)
            stype = "pic" if shape.shape_type == 13 else (
                "table" if has_tbl else ("chart" if has_cht else ("text" if has_tf else str(shape.shape_type))))

            # 1 bounds
            if bb[0] < -TOL or bb[1] < -TOL or bb[2] > SW + TOL or bb[3] > SH + TOL:
                issues.append({"slide": si, "sev": "fail", "type": "out_of_bounds",
                               "shape": name, "detail": f"{stype} bbox={[round(v,2) for v in bb]}"})

            # 2 text fit
            info = frame_text_info(shape)
            has_text = bool(info and info["n_para"] > 0)
            if info and info["n_para"]:
                if info["est_h"] > info["inner_h"] * 1.45 + 0.15:
                    issues.append({"slide": si, "sev": "fail", "type": "text_overflow",
                                   "shape": name,
                                   "detail": (f"severe overflow: est {info['est_h']:.2f}in > frame {info['inner_h']:.2f}in "
                                              f"({info['n_para']} paras, {info['max_pt']}pt)")})
                elif info["est_h"] > info["inner_h"] * 1.15 + 0.08:
                    issues.append({"slide": si, "sev": "warn", "type": "text_overflow",
                                   "shape": name,
                                   "detail": (f"est {info['est_h']:.2f}in > frame {info['inner_h']:.2f}in "
                                              f"({info['n_para']} paras, {info['max_pt']}pt)")})
                if info["min_pt"] and info["min_pt"] < 9:
                    issues.append({"slide": si, "sev": "fail", "type": "font_too_small",
                                   "shape": name, "detail": f"{info['min_pt']}pt"})
                if info["n_para"] > 10:
                    issues.append({"slide": si, "sev": "warn", "type": "too_dense",
                                   "shape": name, "detail": f"{info['n_para']} paragraphs"})
                raw_t = shape.text_frame.text or ""
                if REMNANT_RE.search(raw_t):
                    issues.append({"slide": si, "sev": "warn", "type": "template_placeholder_remnant",
                                   "shape": name, "detail": f"remnant text: '{raw_t.strip()[:45]}'"})
            elif has_tbl:
                tbl = shape.table
                est_h, min_pt, n_cells = 0.0, 99.0, 0
                col_ws = [c.width / EMU_IN for c in tbl.columns]
                for r in tbl.rows:
                    row_h = 0.0
                    for ci, c in enumerate(r.cells):
                        t = c.text.strip()
                        if not t:
                            continue
                        n_cells += 1
                        pt = None
                        for p in c.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size:
                                    pt = run.font.size.pt
                                    break
                            if pt:
                                break
                        pt = pt or 12.0
                        min_pt = min(min_pt, pt)
                        cw = col_ws[ci] if ci < len(col_ws) else 2.0
                        row_h = max(row_h, est_lines(t, pt, max(cw - 0.1, 0.5)) * 1.25 * pt / 72.0 + 0.08)
                    est_h += max(row_h, 0.30)
                if est_h > (shape.height / EMU_IN) * 1.18 + 0.08:
                    issues.append({"slide": si, "sev": "warn", "type": "table_overflow",
                                   "shape": name,
                                   "detail": f"est {est_h:.2f}in > {shape.height/EMU_IN:.2f}in"})
                if n_cells and min_pt < 9:
                    issues.append({"slide": si, "sev": "fail", "type": "font_too_small",
                                   "shape": name, "detail": f"table cell {min_pt}pt"})

            if stype in ("pic", "table", "chart") or (stype == "text" and has_text):
                # Exclude full-bleed background images (>=60% slide area) from overlap checks
                is_bg_pic = (stype == "pic" and (bb[2] - bb[0]) * (bb[3] - bb[1]) >= 0.60 * SW * SH)
                if not is_bg_pic:
                    boxes.append((bb, stype, name))

        # 3 picture/table/chart vs non-empty text overlap (severe >25% -> fail)
        media = [b for b in boxes if b[1] in ("pic", "table", "chart")]
        texts = [b for b in boxes if b[1] == "text"]
        for mb in media:
            for tb in texts:
                ia = inter_area(mb[0], tb[0])
                smaller = min((mb[0][2] - mb[0][0]) * (mb[0][3] - mb[0][1]),
                              (tb[0][2] - tb[0][0]) * (tb[0][3] - tb[0][1]))
                if smaller > 0 and ia / smaller > 0.08:
                    ratio = ia / smaller
                    issues.append({"slide": si, "sev": ("fail" if ratio > 0.25 else "warn"),
                                   "type": "overlap",
                                   "detail": f"{mb[2]} x {tb[2]} ({ratio:.0%})"})

        # 3b media vs media overlap (e.g., figure colliding with table/chart or another figure)
        for mi in range(len(media)):
            for mj in range(mi + 1, len(media)):
                ma, mb = media[mi], media[mj]
                ia = inter_area(ma[0], mb[0])
                smaller = min((ma[0][2] - ma[0][0]) * (ma[0][3] - ma[0][1]),
                              (mb[0][2] - mb[0][0]) * (mb[0][3] - mb[0][1]))
                if smaller > 0 and ia / smaller > 0.08:
                    ratio = ia / smaller
                    issues.append({"slide": si, "sev": ("fail" if ratio > 0.20 else "warn"),
                                   "type": "overlap",
                                   "detail": f"{ma[2]} x {mb[2]} ({ratio:.0%})"})

        # 4 non-empty text vs non-empty text overlap (severe >35% -> fail)
        for ti in range(len(texts)):
            for tj in range(ti + 1, len(texts)):
                a_b, b_b = texts[ti], texts[tj]
                ia = inter_area(a_b[0], b_b[0])
                smaller = min((a_b[0][2] - a_b[0][0]) * (a_b[0][3] - a_b[0][1]),
                              (b_b[0][2] - b_b[0][0]) * (b_b[0][3] - b_b[0][1]))
                if smaller > 0 and ia / smaller > 0.15:
                    ratio = ia / smaller
                    issues.append({"slide": si, "sev": ("fail" if ratio > 0.35 else "warn"),
                                   "type": "text_overlap",
                                   "detail": f"{a_b[2]} x {b_b[2]} ({ratio:.0%})"})

        # 5 template layout/master non-placeholder bottom footer collision check
        slide_area = SW * SH
        master_footers = []
        for container in (getattr(slide, "slide_layout", None),
                          getattr(getattr(slide, "slide_layout", None), "slide_master", None)):
            if container is None:
                continue
            for shp in container.shapes:
                if getattr(shp, "is_placeholder", False):
                    continue
                mbb = bbox_of(shp)
                if mbb is None:
                    continue
                mw, mh = mbb[2] - mbb[0], mbb[3] - mbb[1]
                if mw * mh >= 0.55 * slide_area:
                    continue
                if mbb[1] >= SH - 0.85 and mh <= 0.65:
                    has_vis = bool(getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip())
                    if has_vis or (mw >= SW * 0.25 and mh >= 0.04):
                        master_footers.append((mbb, getattr(shp, "name", "MasterFooter")))
        for mbb, mname in master_footers:
            for sbb, _st, sname in boxes:
                ia = inter_area(mbb, sbb)
                smaller = min((mbb[2] - mbb[0]) * (mbb[3] - mbb[1]),
                              (sbb[2] - sbb[0]) * (sbb[3] - sbb[1]))
                if smaller > 0 and ia / smaller > 0.15:
                    ratio = ia / smaller
                    issues.append({"slide": si, "sev": ("fail" if ratio > 0.35 else "warn"),
                                   "type": "master_footer_overlap",
                                   "detail": f"{sname} x {mname} ({ratio:.0%})"})

    n_fail = sum(1 for i in issues if i["sev"] == "fail")
    n_warn = sum(1 for i in issues if i["sev"] == "warn")
    report = {"ok": n_fail == 0, "n_fail": n_fail, "n_warn": n_warn,
              "issues": issues, "stats": stats}
    os.makedirs(os.path.dirname(os.path.abspath(a.report)) or ".", exist_ok=True)
    with open(a.report, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
