#!/usr/bin/env python3
"""Extract paper structure (title, abstract, sections + excerpts, full figure/table captions,
structured table rows, key quantitative stats) from a PDF, dump page-tagged text, and crop
raster/vector figure & table regions to PNG for slide insertion.

Usage:
  python3 extract_paper.py PAPER.pdf --out paper_structure.json --figdir figs [--text-out paper_text.txt] [--dpi 200]

Detection is LINE-level & 2-column aware. Every locator carries a PDF page number.
No-script fallback: `pdftotext -layout` + manual captions & pymupdf clip.
"""
import argparse, json, re, sys
from pathlib import Path

try:
    import pymupdf
except ImportError:
    pymupdf = None

ABSTRACT_MAX_CHARS = 2600
SEC_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)\s*[.:\s]\s*(.{2,90})$")
FIG_RE = re.compile(
    r"^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Figure|Fig\.?|图|Graph)\s*"
    r"([SA]?\d+[a-z]?|[IVX]{1,4})\s*[:.\-–—]\s*(.{4,})$"
    r"|^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Figure|Fig\.?|图|Graph)\s*"
    r"([SA]?\d+[a-z]?)\s+([A-Z\u4e00-\u9fff].{4,})$"
)
TAB_RE = re.compile(
    r"^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Table|表|Tab\.?)\s*"
    r"([SA]?\d+[a-z]?|[IVX]{1,4})\s*[:.\-–—]\s*(.{4,})$"
    r"|^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Table|表|Tab\.?)\s*"
    r"([SA]?\d+[a-z]?|[IVX]{1,4})\s+([A-Z\u4e00-\u9fff].{4,})$"
)
FIG_HEAD_ONLY = re.compile(
    r"^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Figure|Fig\.?|图|Graph)\s*"
    r"([SA]?\d+[a-z]?|[IVX]{1,4})\s*[:.\-–—]?\s*$"
)
TAB_HEAD_ONLY = re.compile(
    r"^\s*(?:(?i:Extended\s+Data|Supplementary)\s+)?(?i:Table|表|Tab\.?)\s*"
    r"([SA]?\d+[a-z]?|[IVX]{1,4})\s*[:.\-–—]?\s*$"
)


def _cap_num(m):
    raw = (m.group(1) or m.group(3) or "").strip() if m.lastindex and m.lastindex >= 3 else (m.group(1) or "").strip()
    return int(raw) if raw.isdigit() else raw


def _slug(v):
    return re.sub(r"[^A-Za-z0-9_-]", "", str(v)) or "x"
UNNUM_HEADS = {
    "introduction", "background", "related work", "methods", "method", "materials and methods",
    "methodology", "approach", "results", "experiments", "experimental setup", "evaluation",
    "discussion", "limitations", "limitation", "conclusion", "conclusions", "summary",
    "future work", "acknowledgments", "references",
    "结论", "引言", "绪论", "方法", "实验", "结果", "讨论", "局限", "局限性", "背景", "总结",
}
ABSTRACT_HEAD = re.compile(r"^\s*(abstract|摘\s*要|ABSTRACT)\s*[.:\-—–]?\s*", re.I)


def _line_meta(line):
    spans = line.get("spans", [])
    sizes = [sp["size"] for sp in spans] or [0]
    text = "".join(sp["text"] for sp in spans)
    text = re.sub(r"\s+", " ", text).strip()
    bold_n = sum(1 for sp in spans if "Bold" in sp.get("font", "") or "bold" in sp.get("font", ""))
    bold = bool(spans) and bold_n / len(spans) >= 0.5
    return dict(text=text, fmax=max(sizes), fmin=min(sizes), bbox=line["bbox"], bold=bold)


def _page_lines(page):
    d = page.get_text("dict")
    out = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            m = _line_meta(ln)
            if m["text"]:
                m["x0"], m["y0"], m["x1"], m["y1"] = m["bbox"]
                out.append(m)
    out.sort(key=lambda r: (round(r["y0"], 1), r["x0"]))
    return out


def _order_lines(lines, page_width):
    """Reading order: 2-column aware (left column before right), full-width bands stay in place."""
    if len(lines) < 12 or not page_width:
        return sorted(lines, key=lambda r: (round(r["y0"], 1), r["x0"]))
    left = [r for r in lines if r["x1"] < page_width * 0.52]
    right = [r for r in lines if r["x0"] > page_width * 0.48]
    rest = [r for r in lines if r not in left and r not in right]
    if len(left) >= 6 and len(right) >= 6 and len(left) + len(right) >= 0.78 * len(lines):
        col_top = min(r["y0"] for r in left + right)
        top_fw = sorted([r for r in rest if r["y1"] <= col_top + 2], key=lambda r: r["y0"])
        bottom_fw = sorted([r for r in rest if r["y1"] > col_top + 2], key=lambda r: r["y0"])
        return (top_fw
                + sorted(left, key=lambda r: (round(r["y0"], 1), r["x0"]))
                + sorted(right, key=lambda r: (round(r["y0"], 1), r["x0"]))
                + bottom_fw)
    return sorted(lines, key=lambda r: (round(r["y0"], 1), r["x0"]))


def _col_band(bbox, page_w):
    """Return horizontal (x_min, x_max) column band for a caption bbox."""
    x0, _, x1, _ = bbox
    if page_w <= 0:
        return 0.0, 1e6
    if x1 < page_w * 0.54:
        return 0.0, page_w * 0.54
    if x0 > page_w * 0.46:
        return page_w * 0.46, page_w
    return 0.0, page_w


def _overlap_x(b1, b2, pad=8.0):
    return b1[2] + pad >= b2[0] and b2[2] + pad >= b1[0]


def _trim_white(out_path: Path):
    try:
        from PIL import Image
        with Image.open(out_path) as im:
            g = im.convert("L")
            bbox = g.point(lambda v: 0 if v > 245 else 255).getbbox()
            if bbox and bbox[2] - bbox[0] > 40 and bbox[3] - bbox[1] > 30:
                pad = 6
                box = (max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                       min(im.width, bbox[2] + pad), min(im.height, bbox[3] + pad))
                im.crop(box).save(out_path)
                return box[2] - box[0], box[3] - box[1]
            return im.width, im.height
    except Exception:
        return None


def _merge_caption(page_lines, start_idx, page_w):
    """Merge continuation lines directly below a figure/table caption in the same column."""
    base = page_lines[start_idx]
    cx0, cy0, cx1, cy1 = base["bbox"]
    col_l, col_r = _col_band(base["bbox"], page_w)
    parts = [base["text"]]
    last_y1 = cy1
    fsize = max(base["fmax"], 8.0)
    for nb in page_lines[start_idx + 1:]:
        nx0, ny0, nx1, ny1 = nb["bbox"]
        if ny0 < last_y1 - 2 or ny0 - last_y1 > fsize * 1.45:
            break
        if nx1 < col_l or nx0 > col_r:
            break
        if FIG_RE.match(nb["text"]) or TAB_RE.match(nb["text"]) or SEC_RE.match(nb["text"]):
            break
        if abs(nb["fmax"] - base["fmax"]) > 1.8:
            break
        parts.append(nb["text"])
        last_y1 = max(last_y1, ny1)
        cx0, cx1, cy1 = min(cx0, nx0), max(cx1, nx1), max(cy1, ny1)
        if len(parts) >= 5:
            break
    return " ".join(parts)[:500], [cx0, cy0, cx1, cy1]


def extract(pdf_path: Path, figdir: Path, dpi: int, text_out: Path = None):
    doc = pymupdf.open(str(pdf_path))
    all_lines = []
    lines_by_page = {}
    for pno in range(len(doc)):
        page = doc[pno]
        ordered = _order_lines(_page_lines(page), page.rect.width)
        for m in ordered:
            m["page"] = pno + 1
            all_lines.append(m)
        lines_by_page[pno + 1] = ordered

    if text_out is not None:
        text_out.parent.mkdir(parents=True, exist_ok=True)
        chunks = []
        for pno in range(1, len(doc) + 1):
            chunks.append(f"=== PAGE {pno} ===")
            chunks.extend(r["text"] for r in lines_by_page[pno])
        text_out.write_text("\n".join(chunks), encoding="utf-8")

    notes = []
    sizes_sorted = sorted(r["fmin"] for r in all_lines)
    body_font = sizes_sorted[len(sizes_sorted) // 2] if sizes_sorted else 10.0

    # ---- title: largest-font lines on first 2 pages, top area ----
    cands = [r for r in all_lines if r["page"] <= 2 and r["y0"] < 320 and len(r["text"]) >= 8]
    title_lines = []
    if cands:
        best = max(cands, key=lambda r: r["fmax"])
        if best["fmax"] >= body_font + 1.5:
            title_lines = [r for r in all_lines
                           if r["page"] == best["page"] and abs(r["fmax"] - best["fmax"]) < 0.7
                           and r["y0"] < best["y0"] + 3.5 * best["fmax"]]
    title = None
    if title_lines:
        title_lines.sort(key=lambda r: r["y0"])
        title = dict(text=" ".join(r["text"] for r in title_lines)[:300],
                     page=title_lines[0]["page"], font=round(title_lines[0]["fmax"], 1))
    elif all_lines:
        title = dict(text=all_lines[0]["text"][:300], page=all_lines[0]["page"],
                     font=round(all_lines[0]["fmax"], 1))
        notes.append("title fallback: first line used")

    # ---- abstract: absorb lines until a heading-like line ----
    abstract = None
    for i, r in enumerate(all_lines[:140]):
        if ABSTRACT_HEAD.match(r["text"]):
            parts = [ABSTRACT_HEAD.sub("", r["text"], count=1)]
            for nb in all_lines[i + 1:]:
                if nb["page"] > r["page"] + 1:
                    break
                if (nb["fmin"] >= body_font + 1.2 or nb["bold"]) and (
                        SEC_RE.match(nb["text"]) or re.sub(r"[^\w\s]", "", nb["text"].lower()).strip() in UNNUM_HEADS):
                    break
                if FIG_RE.match(nb["text"]) or TAB_RE.match(nb["text"]):
                    break
                parts.append(nb["text"])
                if sum(len(p) for p in parts) > ABSTRACT_MAX_CHARS:
                    break
            abstract = dict(text=" ".join(p for p in parts if p).strip()[:ABSTRACT_MAX_CHARS], page=r["page"])
            break

    # ---- sections + concise excerpts ----
    sections, seen, sec_indices = [], set(), []
    for idx, r in enumerate(all_lines):
        m = SEC_RE.match(r["text"])
        heading_like = (r["fmax"] >= body_font + 1.0 or r["bold"]) and len(r["text"]) <= 110
        if m and heading_like:
            num, name = m.group(1), m.group(2).strip()
            if re.match(r"^\d", name) or num in seen:
                continue
            seen.add(num)
            depth = num.count(".") + 1
            sections.append(dict(number=num, title=name, page=r["page"], level=depth,
                                 detected="bold" if r["bold"] else "font"))
            sec_indices.append(idx)
        elif heading_like and not m:
            key_core = re.sub(r"[^\w\s]", "", r["text"].strip().lower()).strip()
            if key_core in UNNUM_HEADS and key_core not in seen:
                seen.add(key_core)
                sections.append(dict(number="", title=r["text"].strip(), page=r["page"], level=1,
                                     detected="bold" if r["bold"] else "font"))
                sec_indices.append(idx)

    for k, s in enumerate(sections):
        start = sec_indices[k] + 1
        end = sec_indices[k + 1] if k + 1 < len(sec_indices) else len(all_lines)
        buf = []
        total_len = 0
        for r in all_lines[start:end]:
            if FIG_RE.match(r["text"]) or TAB_RE.match(r["text"]):
                continue
            buf.append(r["text"])
            total_len += len(r["text"]) + 1
            if total_len >= 520:
                break
        s["excerpt"] = " ".join(buf)[:520]

    # ---- figure / table captions (multi-line merged; supports Arabic, Roman & S/A labels) ----
    figdir.mkdir(parents=True, exist_ok=True)
    figures, tables = [], []
    seen_fig, seen_tab = set(), set()
    for pno in range(1, len(doc) + 1):
        pw = doc[pno - 1].rect.width
        plines = lines_by_page[pno]
        for idx, r in enumerate(plines):
            fm, tm = FIG_RE.match(r["text"]), TAB_RE.match(r["text"])
            if not fm and not tm and idx + 1 < len(plines):
                nxt = plines[idx + 1]
                if 0 <= nxt["y0"] - r["y1"] <= 18 and len(nxt["text"]) >= 8:
                    fm = FIG_HEAD_ONLY.match(r["text"])
                    if not fm:
                        tm = TAB_HEAD_ONLY.match(r["text"])
            if fm:
                num = _cap_num(fm)
                if (str(num), pno) in seen_fig:
                    continue
                seen_fig.add((str(num), pno))
                cap_txt, cap_box = _merge_caption(plines, idx, pw)
                figures.append(dict(number=num, caption=cap_txt, page=pno, bbox=cap_box))
            elif tm:
                num = _cap_num(tm)
                if (str(num), pno) in seen_tab:
                    continue
                seen_tab.add((str(num), pno))
                cap_txt, cap_box = _merge_caption(plines, idx, pw)
                tables.append(dict(number=num, caption=cap_txt, page=pno, bbox=cap_box))

    # ---- table crops + structured cell rows (with LaTeX booktabs borderless fallback) ----
    for t in tables:
        page = doc[t["page"] - 1]
        pw, ph = page.rect.width, page.rect.height
        col_l, col_r = _col_band(t["bbox"], pw)
        tb_objs = []
        for kw in ({}, {"vertical_strategy": "text", "horizontal_strategy": "lines"}):
            try:
                found = page.find_tables(**kw)
                cands = [tb for tb in found.tables if _overlap_x(tb.bbox, (col_l, 0, col_r, 0), pad=18)]
                if not cands and (col_r - col_l) < pw * 0.8:
                    # spanning full-width table under a short caption
                    cands = [tb for tb in found.tables if (tb.bbox[2] - tb.bbox[0]) > pw * 0.55]
                if cands:
                    tb_objs = cands
                    break
            except Exception:
                continue
        cy0, cy1 = t["bbox"][1], t["bbox"][3]
        cand = None
        if tb_objs:
            below = [tb for tb in tb_objs if tb.bbox[1] >= cy0 - 6]
            above = [tb for tb in tb_objs if tb.bbox[3] <= cy1 + 6]
            cand = min(below, key=lambda tb: abs(tb.bbox[1] - cy1)) if below else (
                   min(above, key=lambda tb: abs(cy0 - tb.bbox[3])) if above else None)

        if cand is not None:
            cb = cand.bbox
            clip = pymupdf.Rect(max(0, cb[0] - 4), max(0, cb[1] - 4),
                                min(pw, cb[2] + 4), min(ph, cb[3] + 4))
            pix = page.get_pixmap(clip=clip, dpi=dpi)
            out = figdir / f"table{_slug(t['number'])}_p{t['page']}.png"
            pix.save(str(out))
            t["image"] = str(out)
            t["image_w"], t["image_h"] = pix.width, pix.height
            try:
                raw_rows = cand.extract() or []
                t["rows"] = [[re.sub(r"\s+", " ", (c or "")).strip() for c in row] for row in raw_rows[:20]]
            except Exception:
                t["rows"] = []
        else:
            # Booktabs horizontal-rule band fallback for LaTeX tables without vertical borders
            h_rules = []
            try:
                for d in page.get_drawings():
                    r = d.get("rect")
                    if r is None:
                        continue
                    bw, bh = r.x1 - r.x0, r.y1 - r.y0
                    if bw >= 60 and bh <= 3.0 and r.y0 >= cy1 - 4 and r.y0 <= min(ph, cy1 + 360):
                        h_rules.append((r.x0, r.y0, r.x1, r.y1))
            except Exception:
                h_rules = []
            if len(h_rules) >= 2:
                rx0 = min(b[0] for b in h_rules)
                ry0 = min(b[1] for b in h_rules)
                rx1 = max(b[2] for b in h_rules)
                ry1 = max(b[3] for b in h_rules)
                if (ry1 - ry0) >= 24 and (rx1 - rx0) >= 60:
                    clip = pymupdf.Rect(max(0, rx0 - 6), max(0, ry0 - 4),
                                        min(pw, rx1 + 6), min(ph, ry1 + 6))
                    pix = page.get_pixmap(clip=clip, dpi=dpi)
                    out = figdir / f"table{_slug(t['number'])}_p{t['page']}.png"
                    pix.save(str(out))
                    t["image"] = str(out)
                    t["image_w"], t["image_h"] = pix.width, pix.height
                    t["rows"] = []
                else:
                    t["image"], t["rows"] = None, []
            else:
                t["image"], t["rows"] = None, []

    # ---- figure crops: column-aware + full-width rescue + raster/vector + top-caption fallback ----
    all_caps_by_page = {}
    for item in figures + tables:
        all_caps_by_page.setdefault(item["page"], []).append(item["bbox"])

    for f in figures:
        page = doc[f["page"] - 1]
        pw, ph = page.rect.width, page.rect.height
        cap_x0, cap_y0, cap_x1, cap_y1 = f["bbox"]
        col_l, col_r = _col_band(f["bbox"], pw)

        def _prose_lines_between(y_top, y_bot, x_l, x_r):
            """Count genuine full-width prose paragraph lines (ignoring subplot captions & legends)."""
            cnt = 0
            band_w = max(x_r - x_l, 80.0)
            for ln in lines_by_page[f["page"]]:
                rb = ln["bbox"]
                if rb[1] >= y_top + 4 and rb[3] <= y_bot - 4 and _overlap_x(rb, (x_l, 0, x_r, 0), pad=4):
                    txt = ln["text"].strip()
                    if re.match(r"^\(?[a-fA-F]\)|^(?:Fig|Table|Panel)\b", txt):
                        continue
                    if len(txt) >= 50 and (rb[2] - rb[0]) >= band_w * 0.65 and ln["fmax"] >= body_font - 0.8:
                        cnt += 1
            return cnt

        # Full-width spanning graphic rescue on 2-column pages when caption itself is short
        if (col_r - col_l) < pw * 0.8:
            for info in page.get_image_info():
                bb = info["bbox"]
                if (bb[2] - bb[0]) > pw * 0.58 and bb[0] < pw * 0.42 and bb[2] > pw * 0.58:
                    if bb[3] <= cap_y0 + 14 and cap_y0 - bb[3] < 70 and _prose_lines_between(bb[3], cap_y0, col_l, col_r) < 4:
                        col_l, col_r = 0.0, pw
                        break

        def _y_ceil_for(box_x0, box_x1):
            """Vertical ceiling = bottom of nearest prior caption that actually overlaps [box_x0, box_x1] in X."""
            yc = 0.0
            for ob in all_caps_by_page.get(f["page"], []):
                if ob is f["bbox"]:
                    continue
                if ob[3] <= cap_y0 - 10 and _overlap_x(ob, (box_x0, 0, box_x1, 0), pad=4):
                    yc = max(yc, ob[3] + 4)
            return yc

        y_ceil = _y_ceil_for(col_l, col_r)

        # 1) raster images in column & vertical band (ignore tiny logos <30pt & distant images blocked by dense prose)
        img_boxes = []
        for info in page.get_image_info():
            bb = info["bbox"]
            bw, bh = bb[2] - bb[0], bb[3] - bb[1]
            if bw < 30 or bh < 30:
                continue
            bb_ceil = _y_ceil_for(bb[0], bb[2])
            if bb[3] <= cap_y0 + 10 and bb[1] >= bb_ceil - 12 and _overlap_x(bb, (col_l, 0, col_r, 0), pad=10):
                if _prose_lines_between(bb[3], cap_y0, bb[0], bb[2]) >= 4:
                    continue
                img_boxes.append(bb)

        clip = None
        if img_boxes:
            ix0 = min(bb[0] for bb in img_boxes)
            ix1 = max(bb[2] for bb in img_boxes)
            iy0 = min(bb[1] for bb in img_boxes)
            box_ceil = _y_ceil_for(ix0, ix1)
            clip = pymupdf.Rect(
                max(col_l, ix0 - 6),
                max(box_ceil, iy0 - 6),
                min(col_r, ix1 + 6),
                cap_y0 - 2,
            )
        else:
            # 2) vector drawings (matplotlib / TikZ / LaTeX plots) in column & vertical band
            draw_boxes = []
            try:
                for d in page.get_drawings():
                    r = d.get("rect")
                    if r is None:
                        continue
                    bb = (r.x0, r.y0, r.x1, r.y1)
                    bw, bh = bb[2] - bb[0], bb[3] - bb[1]
                    if bw < 2 and bh < 2:
                        continue
                    if bh < 2.5 and bw > pw * 0.65:  # skip horizontal header/footer rules
                        continue
                    if bb[3] <= cap_y0 + 4 and bb[1] >= _y_ceil_for(bb[0], bb[2]) - 8 and _overlap_x(bb, (col_l, 0, col_r, 0), pad=8):
                        draw_boxes.append(bb)
            except Exception:
                draw_boxes = []

            if draw_boxes:
                # pick cluster closest above caption (stop if vertical gap > 48pt)
                draw_boxes.sort(key=lambda b: b[3], reverse=True)
                cluster = [draw_boxes[0]]
                cur_top = draw_boxes[0][1]
                for bb in draw_boxes[1:]:
                    if cur_top - bb[3] <= 48:
                        cluster.append(bb)
                        cur_top = min(cur_top, bb[1])
                    else:
                        break
                left = min(bb[0] for bb in cluster)
                top = min(bb[1] for bb in cluster)
                right = max(bb[2] for bb in cluster)
                bottom = max(bb[3] for bb in cluster)
                # if vector plot spans center gutter despite short caption, widen band
                if left < pw * 0.42 and right > pw * 0.58:
                    col_l, col_r = 0.0, pw
                # include nearby axis labels / legends around the vector plot
                for r in lines_by_page[f["page"]]:
                    rb = r["bbox"]
                    if (rb[3] <= cap_y0 - 1 and rb[1] >= top - 28 and rb[3] >= top - 22
                            and rb[0] >= max(col_l, left - 45) and rb[2] <= min(col_r, right + 45)
                            and len(r["text"]) <= 60):
                        left = min(left, rb[0])
                        top = min(top, rb[1])
                        right = max(right, rb[2])
                        bottom = max(bottom, rb[3])
                if (right - left) >= 40 and (bottom - top) >= 35 and _prose_lines_between(bottom, cap_y0, left, right) < 4:
                    box_ceil = _y_ceil_for(left, right)
                    clip = pymupdf.Rect(max(col_l, left - 8), max(box_ceil, top - 8),
                                        min(col_r, right + 8), min(cap_y0 - 2, bottom + 8))

            # 3) top-captioned figure fallback: search raster images BELOW caption
            if clip is None:
                y_floor = ph
                for ob in all_caps_by_page.get(f["page"], []):
                    if ob is f["bbox"]:
                        continue
                    if ob[1] >= cap_y1 + 10 and _overlap_x(ob, (col_l, 0, col_r, 0), pad=10):
                        y_floor = min(y_floor, ob[1] - 4)
                below_imgs = []
                for info in page.get_image_info():
                    bb = info["bbox"]
                    bw, bh = bb[2] - bb[0], bb[3] - bb[1]
                    if bw >= 30 and bh >= 30 and bb[1] >= cap_y1 - 6 and bb[3] <= y_floor + 6:
                        if _overlap_x(bb, (col_l, 0, col_r, 0), pad=10):
                            below_imgs.append(bb)
                if below_imgs:
                    clip = pymupdf.Rect(
                        max(col_l, min(bb[0] for bb in below_imgs) - 6),
                        max(cap_y1 + 2, min(bb[1] for bb in below_imgs) - 6),
                        min(col_r, max(bb[2] for bb in below_imgs) + 6),
                        min(y_floor, max(bb[3] for bb in below_imgs) + 6),
                    )

            # 4) fallback: whitespace gap above caption in same column
            if clip is None:
                prev_bottom = y_ceil
                for r in lines_by_page[f["page"]]:
                    rb = r["bbox"]
                    if rb[3] <= cap_y0 + 2 and _overlap_x(rb, (col_l, 0, col_r, 0), pad=4) and not FIG_RE.match(r["text"]):
                        prev_bottom = max(prev_bottom, rb[3])
                if cap_y0 - prev_bottom > 40:
                    clip = pymupdf.Rect(max(col_l, cap_x0 - 160), max(y_ceil, prev_bottom + 2),
                                        min(col_r, cap_x1 + 160), cap_y0 - 2)

        if clip is None or clip.width < 30 or clip.height < 30:
            f["image"] = None
            notes.append(f"figure {f['number']} p{f['page']}: no croppable region")
            continue
        pix = page.get_pixmap(clip=clip, dpi=dpi)
        out = figdir / f"fig{_slug(f['number'])}_p{f['page']}.png"
        pix.save(str(out))
        f["image"] = str(out)
        f["image_w"], f["image_h"] = pix.width, pix.height
        trimmed = _trim_white(out)
        if trimmed:
            f["image_w"], f["image_h"] = trimmed

    result = dict(
        source=str(pdf_path),
        pages=len(doc),
        text_file=str(text_out) if text_out else None,
        title=title,
        abstract=abstract,
        sections=sections,
        figures=figures,
        tables=tables,
        body_font=round(body_font, 1),
        key_stats=key_stats(all_lines),
        notes=notes,
    )
    return result


def key_stats(lines, cap=20):
    """Quantitative claims (percentages / multipliers / metrics with units) + page locators."""
    stats, seen = [], set()
    pat = re.compile(
        r"(\d+(?:\.\d+)?\s*%|\d+(?:\.\d+)?\s*[x×]\b|[+-]\d+(?:\.\d+)?\s*(?:%|pts|pp|dB|ms|s|min)\b|"
        r"\b\d+(?:\.\d+)?\s*(?:±|\+/-)\s*\d+(?:\.\d+)?|\bp\s*[<≤]\s*0\.\d+)", re.I)
    verb = re.compile(
        r"(improve|reduce|cut|gain|increase|decrease|outperform|achieve|reach|save|boost|speedup|"
        r"accuracy|recall|coverage|precision|f1|auc|rmse|mae|loss|latency|cost|time|error|drop|higher|lower|"
        r"提升|降低|减少|提高|达到|优于|节省|误差|准确率|覆盖率|耗时)", re.I)
    for r in lines:
        if len(stats) >= cap:
            break
        if len(r["text"]) > 240:
            continue
        if pat.search(r["text"]) and verb.search(r["text"]):
            key = r["text"][:100]
            if key in seen:
                continue
            seen.add(key)
            stats.append(dict(text=r["text"][:220], page=r["page"]))
    return stats


def extract_non_pdf(src: Path, text_out: Path = None):
    """Extract structure from .md / .txt / .docx / .csv research materials using the same schema."""
    ext = src.suffix.lower()
    raw_lines = []
    tables = []
    if ext == ".docx":
        from docx import Document
        doc = Document(str(src))
        for p in doc.paragraphs:
            t = re.sub(r"\s+", " ", p.text).strip()
            if t:
                raw_lines.append((t, (p.style.name or "").lower()))
        for idx, tb in enumerate(doc.tables, 1):
            rows = [[re.sub(r"\s+", " ", c.text).strip() for c in r.cells] for r in tb.rows]
            tables.append(dict(number=idx, caption=f"Table {idx} (docx)", page=1, image=None, rows=rows[:20]))
    elif ext == ".csv":
        import csv
        with src.open(encoding="utf-8", errors="ignore") as f:
            rows = list(csv.reader(f))
        tables.append(dict(number=1, caption=f"Table 1 ({src.name})", page=1, image=None, rows=rows[:30]))
        raw_lines = [(", ".join(r), "") for r in rows[:60]]
    else:
        txt = src.read_text(encoding="utf-8", errors="ignore")
        md_tbl_buf = []
        for ln in txt.splitlines():
            s = ln.strip()
            if s.startswith("|") and s.endswith("|") and s.count("|") >= 2:
                cells = [c.strip() for c in s.strip("|").split("|")]
                if not all(re.match(r"^[:\-]+$", c) for c in cells if c):
                    md_tbl_buf.append(cells)
                continue
            elif md_tbl_buf:
                if len(md_tbl_buf) >= 2:
                    tables.append(dict(number=len(tables) + 1, caption=f"Table {len(tables) + 1} (markdown)",
                                       page=1, image=None, rows=md_tbl_buf[:20]))
                md_tbl_buf = []
            if s:
                raw_lines.append((s, "heading" if s.startswith("#") else ""))
        if len(md_tbl_buf) >= 2:
            tables.append(dict(number=len(tables) + 1, caption=f"Table {len(tables) + 1} (markdown)",
                               page=1, image=None, rows=md_tbl_buf[:20]))

    all_lines = []
    sections = []
    title = None
    for idx, (text, st) in enumerate(raw_lines, 1):
        clean = re.sub(r"^#+\s*", "", text).strip()
        all_lines.append(dict(text=clean, page=1))
        if title is None and len(clean) >= 4:
            title = dict(text=clean[:300], page=1, font=None)
        if "heading" in st or text.startswith("#") or SEC_RE.match(clean):
            lvl = len(text) - len(text.lstrip("#")) if text.startswith("#") else 1
            sections.append(dict(number="", title=clean[:110], page=1, level=max(1, lvl),
                                 detected="markup", excerpt=""))
    if text_out is not None:
        text_out.parent.mkdir(parents=True, exist_ok=True)
        text_out.write_text("\n".join([f"=== SOURCE {src.name} ==="] + [r["text"] for r in all_lines]),
                            encoding="utf-8")
    return dict(
        source=str(src), pages=1, text_file=str(text_out) if text_out else None,
        title=title, abstract=None, sections=sections, figures=[], tables=tables,
        body_font=None, key_stats=key_stats(all_lines), notes=[f"extracted non-PDF source ({ext})"],
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="path to source paper (.pdf, or .md/.txt/.docx/.csv)")
    ap.add_argument("--out", default="paper_structure.json")
    ap.add_argument("--figdir", default="figs")
    ap.add_argument("--text-out", default=None, help="optional page-tagged full text dump (default: <out_dir>/paper_text.txt)")
    ap.add_argument("--dpi", type=int, default=200)
    a = ap.parse_args()
    src = Path(a.pdf).resolve()
    if not src.exists():
        print(f"FATAL: {src} not found", file=sys.stderr)
        return 2
    out_path = Path(a.out)
    text_out = Path(a.text_out) if a.text_out else out_path.parent / "paper_text.txt"
    if src.suffix.lower() != ".pdf":
        data = extract_non_pdf(src, text_out=text_out)
    else:
        if pymupdf is None:
            print("FATAL: pymupdf missing; fallback: pdftotext + manual captions", file=sys.stderr)
            return 2
        data = extract(src, Path(a.figdir), a.dpi, text_out=text_out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    n_img = sum(1 for f in data["figures"] if f.get("image"))
    n_tab_rows = sum(1 for t in data["tables"] if t.get("rows"))
    print(f"OK pages={data['pages']} sections={len(data['sections'])} "
          f"figures={len(data['figures'])}(cropped {n_img}) "
          f"tables={len(data['tables'])}(rows {n_tab_rows}) key_stats={len(data['key_stats'])} "
          f"-> {out_path} (text -> {text_out})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
