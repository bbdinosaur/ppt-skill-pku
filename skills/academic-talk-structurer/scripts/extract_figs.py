#!/usr/bin/env python3
"""S1 figure/table & text extractor from source PDF:
- Extracts raster images paired with Fig/Table captions
- Auto-crops vector figures (LaTeX/TikZ/matplotlib PDF drawings) above captions at 180 DPI
- Falls back to page render only if vector crop cannot be bounded
- Optionally exports full page-marked text (--text /app/work/paper.txt) in one pass

Usage:
  python3 extract_figs.py --pdf /input/assets/paper.pdf --out /app/work/figs \
      [--text /app/work/paper.txt] [--max-pages 8] [--min-px 140]

Writes PNGs to --out and manifest <out>/manifest.json:
  [{"file":..., "label":"Fig.2", "page":3, "method":"raster|vector_crop|page_render_needs_crop", "bbox":[...]}]
Fallback (no pymupdf): exports text via pdfplumber/pdftotext and prints pdftoppm commands.
"""
import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

CAPTION_RE = re.compile(
    r"^\s*((?:Extended\s+Data\s+)?Fig(?:ure)?\.?\s*S?\d+[a-z]?|图\s*\d+|Tab(?:le)?\.?\s*(?:S?\d+|[IVX]+)|表\s*\d+)",
    re.IGNORECASE,
)

SECTION_RE = re.compile(
    r"^\s*((?:\d+(?:\.\d+)?|[IVX]+)\.?\s+[A-Z\u4e00-\u9fff][^\n]{2,65}|(?:Abstract|Introduction|Related\s+Work|Method(?:ology)?|Experiments?|Results?|Discussion|Limitations?|Conclusion(?:s)?|摘要|引言|相关工作|方法|实验|结果|讨论|局限|结论)\s*$)",
    re.IGNORECASE,
)


def parse_captions(page):
    """Return list of dicts {label, kind, x0, y0, x1, y1, text} for captions on page.
    Inspects per-line bboxes so captions vertically close to adjacent prose/ticks are never missed.
    """
    caps = []
    seen = set()
    try:
        blocks = (page.get_text("dict") or {}).get("blocks") or []
    except Exception:
        blocks = []
    for blk in blocks:
        if blk.get("type", 0) != 0:
            continue
        lines = blk.get("lines") or []
        for li, ln in enumerate(lines):
            txt = "".join(sp.get("text", "") for sp in (ln.get("spans") or [])).strip()
            if not txt:
                continue
            m = CAPTION_RE.match(txt)
            if not m:
                continue
            rest = txt[m.end():].strip()
            # Guard against wrapped prose lines ("Figure 2 shows...") inside a paragraph:
            # non-first lines must look like a real caption (colon/period/dash separator or standalone label)
            if li > 0 and rest and not re.match(r"^[:：.\-–—]|^$", rest):
                continue
            if li == 0 and re.match(r"^(?:shows?|demonstrates?|illustrates?|presents?|compares?|reports?|summarizes?|depicts?|lists?|gives?|indicates?|reveals?)\b", rest, re.I):
                continue
            lbl = m.group(1).rstrip(". ")
            kind = "table" if re.match(r"^(?:tab|表)", lbl, re.I) else "figure"
            bb = ln.get("bbox") or blk.get("bbox") or (0, 0, 0, 0)
            key = (lbl.lower(), round(float(bb[1]) / 8.0))
            if key in seen:
                continue
            seen.add(key)
            caps.append({
                "label": lbl, "kind": kind,
                "x0": float(bb[0]), "y0": float(bb[1]),
                "x1": float(bb[2]), "y1": float(bb[3]),
                "text": txt[:90],
            })
    return caps


def dump_text_fitz(doc, text_path: Path):
    chunks = []
    for pno in range(doc.page_count):
        chunks.append(f"--- PAGE {pno + 1} ---\n{doc[pno].get_text('text')}")
    text_path.parent.mkdir(parents=True, exist_ok=True)
    text_path.write_text("\n\n".join(chunks), encoding="utf-8")
    return doc.page_count


def dump_text_fallback(pdf_path: str, text_path: Path):
    text_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            chunks = [f"--- PAGE {i + 1} ---\n{pg.extract_text() or ''}" for i, pg in enumerate(pdf.pages)]
        text_path.write_text("\n\n".join(chunks), encoding="utf-8")
        return len(chunks)
    except Exception:
        pass
    try:
        r = subprocess.run(["pdftotext", pdf_path, str(text_path)], capture_output=True, text=True, timeout=60)
        if r.returncode == 0 and text_path.exists():
            return -1
    except Exception:
        pass
    return 0


def find_vector_crop(page, cap):
    """Find tight clip rect for a vector figure (above Fig caption) or table (below Table caption)."""
    pr = page.rect
    pw, ph = pr.width, pr.height
    drawings = []
    try:
        for d in page.get_drawings():
            r = d.get("rect")
            if not r:
                continue
            # skip page-wide border frames or hairline page rules
            if r.width > pw * 0.94 and r.height > ph * 0.90:
                continue
            if r.width < 2 and r.height < 2:
                continue
            drawings.append(r)
    except Exception:
        return None
    if not drawings:
        return None

    # search window: above figure caption, or below table caption (within ~55% of page height)
    if cap["kind"] == "figure":
        win_y0 = max(pr.y0 + 18, cap["y0"] - ph * 0.58)
        win_y1 = cap["y0"] - 1
    else:
        win_y0 = cap["y1"] + 1
        win_y1 = min(pr.y1 - 18, cap["y1"] + ph * 0.48)
    if win_y1 - win_y0 < 35:
        return None

    # determine column span: full-width if wide caption, centered caption, single-column page, or drawings cross midline
    cap_w = cap["x1"] - cap["x0"]
    mid = pw / 2.0
    cap_cx = 0.5 * (cap["x0"] + cap["x1"])
    is_single_col_page = False
    try:
        wide_blocks = sum(
            1 for b in page.get_text("blocks")
            if len(b) >= 5 and (len(b) < 7 or b[6] == 0) and (float(b[2]) - float(b[0])) > pw * 0.58 and len((b[4] or "").strip()) > 60
        )
        is_single_col_page = wide_blocks >= 2
    except Exception:
        pass
    band_drawings = [r for r in drawings if r.y1 >= win_y0 and r.y0 <= win_y1]
    band_crosses_mid = any(r.x0 < pw * 0.44 and r.x1 > pw * 0.56 for r in band_drawings) or (
        any(r.x1 < mid for r in band_drawings) and any(r.x0 > mid for r in band_drawings) and abs(cap_cx - mid) < pw * 0.16
    )
    if cap_w > pw * 0.50 or abs(cap_cx - mid) < pw * 0.14 or is_single_col_page or band_crosses_mid:
        col_x0, col_x1 = pr.x0 + 8, pr.x1 - 8
    elif cap_cx < mid:
        col_x0, col_x1 = pr.x0 + 8, mid + 18
    else:
        col_x0, col_x1 = mid - 18, pr.x1 - 8

    cand = []
    for r in drawings:
        if r.x1 < col_x0 or r.x0 > col_x1:
            continue
        if r.y1 < win_y0 or r.y0 > win_y1:
            continue
        cand.append(r)
    if not cand:
        return None

    # cluster drawings contiguous with the caption (stop at large vertical gaps > 36pt)
    if cap["kind"] == "figure":
        cand.sort(key=lambda r: -r.y1)
        cur_top = cap["y0"]
        kept = []
        for r in cand:
            if cur_top - r.y1 > 48 and kept:
                break
            kept.append(r)
            cur_top = min(cur_top, r.y0)
    else:
        cand.sort(key=lambda r: r.y0)
        cur_bot = cap["y1"]
        kept = []
        for r in cand:
            if r.y0 - cur_bot > 48 and kept:
                break
            kept.append(r)
            cur_bot = max(cur_bot, r.y1)

    if not kept:
        return None
    x0 = min(r.x0 for r in kept)
    y0 = min(r.y0 for r in kept)
    x1 = max(r.x1 for r in kept)
    y1 = max(r.y1 for r in kept)

    # include axis labels / legend text blocks overlapping or tightly adjacent to drawing cluster
    for b in page.get_text("blocks"):
        if len(b) < 5 or (len(b) >= 7 and b[6] != 0):
            continue
        bx0, by0, bx1, by1 = float(b[0]), float(b[1]), float(b[2]), float(b[3])
        # skip the caption itself and distant body paragraphs
        if abs(by0 - cap["y0"]) < 3 and abs(bx0 - cap["x0"]) < 10:
            continue
        if bx1 >= x0 - 40 and bx0 <= x1 + 32 and by1 >= y0 - 18 and by0 <= min(cap["y0"] - 1, y1 + 22):
            # only absorb short chart labels/ticks/legends (avoid grabbing a body paragraph)
            txt = " ".join((b[4] or "").split())
            if txt and len(txt) <= 120 and (bx1 - bx0) <= (x1 - x0) + 70 and not (txt.endswith((".", "。")) and len(txt) > 45):
                x0, y0 = min(x0, bx0), min(y0, by0)
                x1, y1 = max(x1, bx1), max(y1, by1)

    if (x1 - x0) < 70 or (y1 - y0) < 45:
        return None
    pad = 8.0
    return [
        max(pr.x0, x0 - pad),
        max(pr.y0, y0 - pad),
        min(pr.x1, x1 + pad),
        min(pr.y1, y1 + pad),
    ]


def build_paper_outline(doc):
    """Extract section headings, captions, and limitation/conclusion pages for fast S1 anchor reading."""
    sections, captions, lim_pages = [], [], []
    for pno in range(doc.page_count):
        pg = doc[pno]
        txt = pg.get_text("text") or ""
        if re.search(r"\b(limitation|threats?\s+to\s+validity|future\s+work|conclusion)\b|局限|不足|结论", txt, re.I):
            lim_pages.append(pno + 1)
        for line in txt.splitlines():
            s = line.strip()
            if not s or len(s) > 75:
                continue
            if SECTION_RE.match(s) and not CAPTION_RE.match(s):
                sections.append({"page": pno + 1, "heading": s})
        for c in parse_captions(pg):
            captions.append({"page": pno + 1, "label": c["label"], "kind": c["kind"], "caption": c["text"]})
    return {"pages": doc.page_count, "sections": sections[:40], "captions": captions[:40], "limitation_or_conclusion_pages": lim_pages}


def extract_tables_from_pdf(doc, pdf_path: str):
    """Extract structured tables from PDF pages (pdfplumber first for exact decimals/spacing, PyMuPDF fallback)."""
    tables = []
    # Pre-parse Table captions per page using fitz for accurate bbox-to-caption pairing
    tcaps_by_page = {}
    for pno in range(doc.page_count):
        tcaps_by_page[pno] = [c for c in parse_captions(doc[pno]) if c["kind"] == "table"]

    def _clean_rows(raw):
        cleaned = []
        for row in (raw or []):
            if not row:
                continue
            vals = [" ".join(str(c or "").split()) for c in row]
            if any(vals):
                cleaned.append(vals)
        if len(cleaned) < 2:
            return None
        ncols = max(len(r) for r in cleaned)
        keep_cols = [ci for ci in range(ncols) if any(ci < len(r) and r[ci] for r in cleaned)]
        if len(keep_cols) < 2:
            return None
        cleaned = [[r[ci] if ci < len(r) else "" for ci in keep_cols] for r in cleaned]
        # Filter out false-positive prose splits (real tables have compact cells and/or numeric data)
        non_empty = [c for r in cleaned for c in r if c]
        if not non_empty:
            return None
        avg_len = sum(len(c) for c in non_empty) / len(non_empty)
        if avg_len > 48:
            return None
        return cleaned

    matched_caps = set()

    def _record_table(pno, ti, raw, bbox, n_on_page, forced_cap=None):
        cleaned = _clean_rows(raw)
        if not cleaned:
            return False
        tcaps = tcaps_by_page.get(pno, [])
        label, cap_txt = f"Table_p{pno+1}_{ti+1}", ""
        if forced_cap is not None:
            label, cap_txt = forced_cap["label"], forced_cap["text"]
            matched_caps.add((pno, label))
        elif bbox and tcaps:
            best = min(
                ((min(abs(c["y1"] - bbox[1]), abs(c["y0"] - bbox[3])), c) for c in tcaps),
                key=lambda x: x[0],
            )
            if best[0] < 160:
                label, cap_txt = best[1]["label"], best[1]["text"]
                matched_caps.add((pno, label))
        elif len(tcaps) == 1 and n_on_page == 1:
            label, cap_txt = tcaps[0]["label"], tcaps[0]["text"]
            matched_caps.add((pno, label))
        tables.append({
            "page": pno + 1, "label": label, "caption": cap_txt,
            "headers": cleaned[0], "rows": cleaned[1:14],
        })
        return True

    plumber_ok = False
    try:
        import time as _time
        import pdfplumber
        t_start = _time.monotonic()
        with pdfplumber.open(pdf_path) as pdf:
            plumber_ok = True
            n_pages = len(pdf.pages)
            for pno, pg in enumerate(pdf.pages):
                if _time.monotonic() - t_start > 10.0:
                    break
                pbx0, pby0, pbx1, pby1 = (float(x) for x in getattr(pg, "bbox", (0.0, 0.0, float(pg.width), float(pg.height))))
                has_tcap = bool(tcaps_by_page.get(pno))
                # Only run pdfplumber table extraction on pages with a Table caption (or tiny <=4p docs) — 10x faster & zero false positives
                if not has_tcap and n_pages > 4:
                    continue
                found = []
                try:
                    found = pg.find_tables() or []
                except Exception:
                    found = []
                for ti, tb in enumerate(found):
                    try:
                        raw = tb.extract()
                        bbox = list(tb.bbox) if getattr(tb, "bbox", None) else None
                        _record_table(pno, ti, raw, bbox, len(found))
                    except Exception:
                        continue
                # Booktabs fallback: LaTeX papers use horizontal rules (\toprule/\midrule/\bottomrule) without vertical lines
                unmatched = [c for c in tcaps_by_page.get(pno, []) if (pno, c["label"]) not in matched_caps]
                if unmatched:
                    pw, ph = float(pg.width), float(pg.height)
                    # Collect horizontal rules (both PDF lines and thin horizontal rects)
                    hrules = []
                    for ln in (pg.lines or []):
                        if abs(float(ln["y0"]) - float(ln["y1"])) <= 3.0 and (float(ln["x1"]) - float(ln["x0"])) >= 80.0:
                            hrules.append((float(ln["x0"]), float(ln["top"]), float(ln["x1"]), float(ln["bottom"])))
                    for rc in (pg.rects or []):
                        if abs(float(rc["bottom"]) - float(rc["top"])) <= 3.0 and (float(rc["x1"]) - float(rc["x0"])) >= 80.0:
                            hrules.append((float(rc["x0"]), float(rc["top"]), float(rc["x1"]), float(rc["bottom"])))

                    def parse_booktabs_band(crop_pg, bx0, by0, bx1, by1):
                        # Filter horizontal rules inside this column/vertical band
                        in_rules = sorted(
                            [r for r in hrules if r[2] >= bx0 + 20 and r[0] <= bx1 - 20 and by0 - 4 <= r[1] <= by1 + 4],
                            key=lambda r: r[1],
                        )
                        if len(in_rules) < 2:
                            return None, None
                        # Cluster rules into one contiguous booktabs table (stop if gap > 220pt)
                        cluster = [in_rules[0]]
                        for r in in_rules[1:]:
                            if r[1] - cluster[-1][1] <= 220.0:
                                cluster.append(r)
                            else:
                                break
                        if len(cluster) < 2:
                            return None, None
                        rx0 = max(pbx0, min(r[0] for r in cluster) - 4.0)
                        rx1 = min(pbx1, max(r[2] for r in cluster) + 4.0)
                        ry0 = max(pby0, cluster[0][1] - 2.0)
                        ry1 = min(pby1, cluster[-1][3] + 2.0)
                        if ry1 - ry0 < 24.0 or rx1 - rx0 < 80.0:
                            return None, None
                        try:
                            tbox = pg.crop((rx0, ry0, rx1, ry1))
                            words = tbox.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False) or []
                        except Exception:
                            return None, None
                        if len(words) < 4:
                            return None, None
                        # 1) Cluster words into visual rows by vertical top coordinate
                        words.sort(key=lambda w: (float(w["top"]), float(w["x0"])))
                        wrows = []
                        for w in words:
                            wt = float(w["top"])
                            if not wrows or abs(wt - wrows[-1]["y"]) > 5.0:
                                wrows.append({"y": wt, "words": [w]})
                            else:
                                wrows[-1]["words"].append(w)
                        if len(wrows) < 2:
                            return None, None
                        # 2) Merge adjacent words within same row into column phrases (gap <= 7pt)
                        row_cells = []
                        spans = []
                        for wr in wrows:
                            ws = sorted(wr["words"], key=lambda w: float(w["x0"]))
                            cells = []
                            for w in ws:
                                wx0, wx1 = float(w["x0"]), float(w["x1"])
                                if cells and (wx0 - cells[-1]["x1"]) <= 7.0:
                                    cells[-1]["x1"] = max(cells[-1]["x1"], wx1)
                                    cells[-1]["text"] += " " + w["text"]
                                else:
                                    cells.append({"x0": wx0, "x1": wx1, "text": w["text"]})
                            row_cells.append(cells)
                            for c_it in cells:
                                spans.append((c_it["x0"], c_it["x1"]))
                        # 3) Project horizontal spans to find disjoint column intervals (whitespace gutters)
                        spans.sort()
                        col_intervals = []
                        for sx0, sx1 in spans:
                            if not col_intervals or sx0 > col_intervals[-1][1] + 2.0:
                                col_intervals.append([sx0, sx1])
                            else:
                                col_intervals[-1][1] = max(col_intervals[-1][1], sx1)
                        if len(col_intervals) < 2:
                            return None, None
                        # 4) Assign each cell to its column index
                        out_rows = []
                        for cells in row_cells:
                            row_arr = [""] * len(col_intervals)
                            for c_it in cells:
                                mid_c = 0.5 * (c_it["x0"] + c_it["x1"])
                                ci = min(range(len(col_intervals)),
                                         key=lambda k: 0.0 if col_intervals[k][0] <= mid_c <= col_intervals[k][1]
                                         else min(abs(mid_c - col_intervals[k][0]), abs(mid_c - col_intervals[k][1])))
                                row_arr[ci] = (row_arr[ci] + " " + c_it["text"]).strip()
                            out_rows.append(row_arr)
                        return out_rows, [rx0, ry0, rx1, ry1]

                    for ti_u, cap in enumerate(unmatched):
                        cap_w = cap["x1"] - cap["x0"]
                        if cap_w > pw * 0.45:
                            cx0, cx1 = 0.0, pw
                        else:
                            mid = pw / 2.0
                            cx0, cx1 = (0.0, mid + 14.0) if (cap["x0"] + cap["x1"]) / 2.0 < mid else (mid - 14.0, pw)
                        bands = [
                            (max(0.0, cx0), min(ph, cap["y1"] + 1.0), min(pw, cx1), min(ph, cap["y1"] + ph * 0.48)),
                            (max(0.0, cx0), max(0.0, cap["y0"] - ph * 0.42), min(pw, cx1), max(0.0, cap["y0"] - 1.0)),
                        ]
                        got_bt = False
                        for bx0, by0, bx1, by1 in bands:
                            if by1 - by0 < 36 or bx1 - bx0 < 80:
                                continue
                            bt_rows, bt_bbox = parse_booktabs_band(pg, bx0, by0, bx1, by1)
                            if bt_rows and _record_table(pno, 100 + ti_u, bt_rows, bt_bbox, 1, forced_cap=cap):
                                got_bt = True
                                break
                        if got_bt:
                            continue
    except Exception:
        plumber_ok = False

    if not plumber_ok:
        for pno in range(doc.page_count):
            page = doc[pno]
            try:
                ft = page.find_tables()
                tbl_list = list(ft.tables) if ft else []
            except Exception:
                tbl_list = []
            for ti, tb in enumerate(tbl_list):
                try:
                    raw = tb.extract()
                    bbox = list(tb.bbox) if getattr(tb, "bbox", None) else None
                    _record_table(pno, ti, raw, bbox, len(tbl_list))
                except Exception:
                    continue
    # Fallback 3: borderless whitespace-gutter table parser for unmatched Table captions (e.g. plain-text PDFs)
    for pno in range(doc.page_count):
        unm = [c for c in tcaps_by_page.get(pno, []) if (pno, c["label"]) not in matched_caps]
        if not unm:
            continue
        page = doc[pno]
        pw, ph = float(page.rect.width), float(page.rect.height)
        try:
            words = page.get_text("words") or []
        except Exception:
            words = []
        if not words:
            continue
        for ti_b, cap in enumerate(unm):
            by0, by1 = cap["y1"] + 1.0, min(ph - 12.0, cap["y1"] + ph * 0.30)
            cand_w = [w for w in words if by0 <= w[1] <= by1 and w[3] <= by1 + 6.0]
            if len(cand_w) < 4:
                continue
            cand_w.sort(key=lambda w: (round(w[1] / 3.0), w[0]))
            wrows = []
            for w in cand_w:
                if not wrows or abs(w[1] - wrows[-1][0]) > 4.0:
                    wrows.append([w[1], [w]])
                else:
                    wrows[-1][1].append(w)
            # Stop at large vertical gap (>22pt) indicating body prose below table
            kept_wr = []
            prev_y = cap["y1"]
            for ry, ws in wrows:
                if kept_wr and (ry - prev_y) > 22.0:
                    break
                kept_wr.append((ry, sorted(ws, key=lambda x: x[0])))
                prev_y = ry
            if len(kept_wr) < 2:
                continue
            row_cells, spans = [], []
            for _, ws in kept_wr:
                cells = []
                for w in ws:
                    if cells and (w[0] - cells[-1][1]) <= 9.0:
                        cells[-1][1] = max(cells[-1][1], w[2])
                        cells[-1][2] += " " + w[4]
                    else:
                        cells.append([w[0], w[2], w[4]])
                # Stop if line looks like a prose sentence (1 wide cell > 65 chars)
                if len(cells) == 1 and len(cells[0][2]) > 65:
                    break
                row_cells.append(cells)
                for c_it in cells:
                    spans.append((c_it[0], c_it[1]))
            if len(row_cells) < 2:
                continue
            spans.sort()
            cols = []
            for sx0, sx1 in spans:
                if not cols or sx0 > cols[-1][1] + 4.0:
                    cols.append([sx0, sx1])
                else:
                    cols[-1][1] = max(cols[-1][1], sx1)
            if len(cols) < 2 or len(cols) > 10:
                continue
            grid = []
            for cells in row_cells:
                arr = [""] * len(cols)
                for cx0, cx1, ctxt in cells:
                    mc = 0.5 * (cx0 + cx1)
                    ci = min(range(len(cols)), key=lambda k: 0.0 if cols[k][0] <= mc <= cols[k][1] else min(abs(mc - cols[k][0]), abs(mc - cols[k][1])))
                    arr[ci] = (arr[ci] + " " + ctxt).strip()
                grid.append(arr)
            # Require at least 2 numeric cells so prose columns aren't mistaken for a table
            num_cells = sum(1 for r in grid for c in r if re.search(r"\d", c))
            if num_cells >= 2:
                _record_table(pno, 200 + ti_b, grid, [cols[0][0], kept_wr[0][0], cols[-1][1], kept_wr[-1][0] + 12.0], 1, forced_cap=cap)
    return tables


def absorb_nearby_labels(page, rect_xy, cap, pr, is_single_raster=False):
    """Expand a figure crop rect to include tightly adjacent subfigure/axis labels without bleeding into body prose."""
    x0, y0, x1, y1 = rect_xy
    fig_w = max(1.0, x1 - x0)
    for b in page.get_text("blocks"):
        if len(b) < 5 or (len(b) >= 7 and b[6] != 0):
            continue
        bx0, by0, bx1, by1 = float(b[0]), float(b[1]), float(b[2]), float(b[3])
        if cap and abs(by0 - cap["y0"]) < 4 and abs(bx0 - cap["x0"]) < 12:
            continue
        if cap and cap["y0"] >= y1 - 2 and by1 > cap["y0"] - 1:
            continue
        if cap and cap["y1"] <= y0 + 2 and by0 < cap["y1"] + 1:
            continue
        txt = " ".join((b[4] or "").split())
        if not txt or len(txt) > 58 or txt.count("\n") > 1:
            continue
        # Reject prose sentences that belong to surrounding two-column body text
        if len(txt.split()) > 9 or (txt.endswith((".", "。")) and len(txt) > 28):
            continue
        bw = bx1 - bx0
        bxc = 0.5 * (bx0 + bx1)
        # Case 1: top/bottom subfigure label (e.g. "(a) Throughput") horizontally aligned within figure width
        horiz_aligned = (bx0 >= x0 - 8 and bx1 <= x1 + 8) or (x0 - 4 <= bxc <= x1 + 4 and bw <= fig_w * 0.92)
        vert_adj = (y0 - 14 <= by1 <= y0 + 4) or (y1 - 4 <= by0 <= y1 + 14)
        # Case 2: narrow vertical axis/tick label directly beside composite figure (skip for self-contained single raster)
        narrow_side = (not is_single_raster) and (bw <= 36) and (by0 >= y0 - 4 and by1 <= y1 + 4) and (
            (x0 - 14 <= bx1 <= x0 + 2) or (x1 - 2 <= bx0 <= x1 + 14)
        )
        if (horiz_aligned and vert_adj) or narrow_side:
            x0, y0 = min(x0, bx0), min(y0, by0)
            x1, y1 = max(x1, bx1), max(y1, by1)
    pad = 0.0 if is_single_raster else 4.0
    return [
        max(pr.x0, x0 - pad), max(pr.y0, y0 - pad),
        min(pr.x1, x1 + pad), min(pr.y1, y1 + pad),
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", default="figs")
    ap.add_argument("--text", help="optional path to dump full page-marked text (e.g. /app/work/paper.txt)")
    ap.add_argument("--outline", help="optional path to save paper_outline.json (default: sibling of --text)")
    ap.add_argument("--tables", help="optional path to save paper_tables.json (default: sibling of --text)")
    ap.add_argument("--max-pages", type=int, default=8)
    ap.add_argument("--min-px", type=int, default=140)
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
    except ImportError:
        if a.text:
            n = dump_text_fallback(a.pdf, Path(a.text))
            print(f"text fallback -> {a.text} (pages={n})")
        print("pymupdf missing. Fallback commands:\n"
              f"  pdftoppm -png -r 150 {a.pdf} {out}/page\n"
              "then crop figure regions manually (PIL) and name them fig_p<N>.png", file=sys.stderr)
        sys.exit(3)

    doc = fitz.open(a.pdf)
    if a.text:
        npages = dump_text_fitz(doc, Path(a.text))
        print(f"text -> {a.text} ({npages} pages)")
    outline_path = Path(a.outline) if a.outline else (Path(a.text).parent / "paper_outline.json" if a.text else None)
    if outline_path:
        ol = build_paper_outline(doc)
        outline_path.parent.mkdir(parents=True, exist_ok=True)
        outline_path.write_text(json.dumps(ol, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"outline -> {outline_path} (sections={len(ol['sections'])}, captions={len(ol['captions'])})")

    tables_path = Path(a.tables) if a.tables else (Path(a.text).parent / "paper_tables.json" if a.text else None)
    if tables_path:
        tbls = extract_tables_from_pdf(doc, a.pdf)
        tables_path.parent.mkdir(parents=True, exist_ok=True)
        tables_path.write_text(json.dumps(tbls, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"tables -> {tables_path} (tables={len(tbls)})")

    manifest = []
    seen_xref = set()
    matched_caps = set()  # (pno, label)

    # pass 1: embedded raster images with caption pairing + WYSIWYG composite cropping (preserves rotation/SMask/overlays)
    for pno in range(doc.page_count):
        page = doc[pno]
        pr = page.rect
        caps = parse_captions(page)
        grouped = {}   # cap_label -> {"cap": cap, "rects": [...]}
        singles = []   # (xref, smask, rect_or_None)
        for img in page.get_images(full=True):
            xref = img[0]
            smask = img[1] if len(img) > 1 else 0
            if xref in seen_xref:
                continue
            seen_xref.add(xref)
            rects = []
            try:
                rects = [r for r in page.get_image_rects(xref) if r.width >= 36 and r.height >= 36]
            except Exception:
                rects = []
            # Check native pixel size before keeping tiny logos/icons
            try:
                base_info = doc.extract_image(xref)
                nat_w = int(base_info.get("width", 0))
                nat_h = int(base_info.get("height", 0))
                if nat_w < a.min_px or nat_h < a.min_px:
                    continue
            except Exception:
                pass
            if rects and caps:
                r0 = rects[0]
                best = None
                for c in caps:
                    d = abs(c["y0"] - r0.y1) if c["y0"] >= r0.y1 else abs(c["y1"] - r0.y0)
                    if best is None or d < best[0]:
                        best = (d, c)
                if best and best[0] < 140:
                    c_hit = best[1]
                    matched_caps.add((pno, c_hit["label"]))
                    grp_e = grouped.setdefault(c_hit["label"], {"cap": c_hit, "rects": [], "items": []})
                    grp_e["rects"].append(r0)
                    grp_e["items"].append((xref, smask, r0))
                    continue
            singles.append((xref, smask, rects[0] if rects else None))

        def _extract_native_pixmap(xref_id, smask_id):
            base_pix = fitz.Pixmap(doc, xref_id)
            if smask_id > 0:
                mask_pix = fitz.Pixmap(doc, smask_id)
                px = fitz.Pixmap(base_pix, mask_pix)
                base_pix = mask_pix = None
            else:
                px = base_pix
            if px.n - px.alpha > 3:
                px = fitz.Pixmap(fitz.csRGB, px)
            if px.alpha:
                flat = fitz.Pixmap(fitz.csRGB, px.irect, False)
                flat.set_rect(flat.irect, (255, 255, 255))
                flat.copy(px, px.irect)
                px = flat
            return px

        # Render grouped caption-paired figures: native full-res for clean single rasters (zero text bleed), WYSIWYG clip for multi-panel/overlaid
        page_drawings = None
        for lbl, grp in grouped.items():
            rs = grp["rects"]
            items = grp["items"]
            cap = grp["cap"]
            pix = None
            crop = None
            method_tag = "raster_composite" if len(rs) > 1 else "raster"
            if len(items) == 1 and page.rotation == 0:
                xref_i, smask_i, r_only = items[0]
                if page_drawings is None:
                    try:
                        page_drawings = [d["rect"] for d in page.get_drawings() if d.get("rect")]
                    except Exception:
                        page_drawings = []
                ov_draw = sum(1 for dr in page_drawings if dr.intersects(r_only) and (dr.width > 8 or dr.height > 8))
                if ov_draw < 6:
                    try:
                        cand_px = _extract_native_pixmap(xref_i, smask_i)
                        if cand_px.width >= a.min_px and cand_px.height >= a.min_px // 2:
                            pix = cand_px
                            crop = [r_only.x0, r_only.y0, r_only.x1, r_only.y1]
                    except Exception:
                        pix = None
            if pix is None:
                ux0 = min(r.x0 for r in rs)
                uy0 = min(r.y0 for r in rs)
                ux1 = max(r.x1 for r in rs)
                uy1 = max(r.y1 for r in rs)
                crop = absorb_nearby_labels(page, [ux0, uy0, ux1, uy1], cap, pr, is_single_raster=(len(rs) == 1))
                clip = fitz.Rect(*crop)
                dpi = 220 if clip.width < 180 else 180
                pix = page.get_pixmap(clip=clip, dpi=dpi)
            if pix.width >= a.min_px and pix.height >= a.min_px // 2:
                idx = len([m for m in manifest if m["page"] == pno + 1])
                fname = f"fig_p{pno+1}_{idx}.png"
                pix.save(str(out / fname))
                manifest.append({
                    "file": str(out / fname),
                    "label": lbl, "caption": cap["text"],
                    "page": pno + 1,
                    "method": method_tag,
                    "bbox": crop, "px": [pix.width, pix.height],
                    "low_res": bool(pix.width < 420 or pix.height < 220),
                })
            pix = None

        # Render uncaptioned raster images
        for xref, smask, r0 in singles:
            if r0 is not None:
                dpi = 220 if r0.width < 180 else 180
                pix = page.get_pixmap(clip=r0, dpi=dpi)
                bbox = [r0.x0, r0.y0, r0.x1, r0.y1]
            else:
                try:
                    base_pix = fitz.Pixmap(doc, xref)
                    if smask > 0:
                        mask_pix = fitz.Pixmap(doc, smask)
                        pix = fitz.Pixmap(base_pix, mask_pix)
                        base_pix = mask_pix = None
                    else:
                        pix = base_pix
                except Exception:
                    continue
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                if pix.alpha:
                    flat = fitz.Pixmap(fitz.csRGB, pix.irect, False)
                    flat.set_rect(flat.irect, (255, 255, 255))
                    flat.copy(pix, pix.irect)
                    pix = flat
                bbox = None
            if pix.width < a.min_px or pix.height < a.min_px:
                pix = None
                continue
            idx = len([m for m in manifest if m["page"] == pno + 1])
            fname = f"fig_p{pno+1}_{idx}.png"
            pix.save(str(out / fname))
            manifest.append({
                "file": str(out / fname),
                "label": fname.replace(".png", ""),
                "page": pno + 1, "method": "raster",
                "bbox": bbox, "px": [pix.width, pix.height],
                "low_res": bool(pix.width < 420 or pix.height < 220),
            })
            pix = None

    # pass 2: vector-figure auto-crop for unmatched captions, with page-render fallback
    rendered_pages = 0
    for pno in range(doc.page_count):
        page = doc[pno]
        caps = [c for c in parse_captions(page) if (pno, c["label"]) not in matched_caps]
        if not caps:
            continue
        unbounded = []
        for c in caps:
            # Skip vector-cropping tables AND do not pollute figure manifest with uncropped full-page renders for tables
            # (tables belong in paper_tables.json / native editable PPTX tables, never pasted as raw A4 page screenshots)
            if c["kind"] == "table":
                continue
            crop = find_vector_crop(page, c)
            if crop:
                clip = fitz.Rect(*crop)
                pix = page.get_pixmap(clip=clip, dpi=180)
                if pix.width >= a.min_px and pix.height >= a.min_px // 2:
                    idx = len([m for m in manifest if m["page"] == pno + 1])
                    fname = f"vec_p{pno+1}_{idx}.png"
                    pix.save(str(out / fname))
                    manifest.append({
                        "file": str(out / fname),
                        "label": c["label"], "caption": c["text"],
                        "page": pno + 1, "method": "vector_crop",
                        "bbox": crop, "px": [pix.width, pix.height],
                        "low_res": bool(pix.width < 420 or pix.height < 220),
                    })
                    pix = None
                    continue
                pix = None
            unbounded.append(c)

        if unbounded and rendered_pages < a.max_pages:
            pix = page.get_pixmap(dpi=150)
            fname = f"page_p{pno+1}.png"
            pix.save(str(out / fname))
            for c in unbounded:
                manifest.append({
                    "file": str(out / fname),
                    "label": c["label"], "caption": c["text"],
                    "page": pno + 1, "method": "page_render_needs_crop",
                    "bbox": [c["x0"], c["y0"], c["x1"], c["y1"]],
                    "px": [pix.width, pix.height],
                })
            pix = None
            rendered_pages += 1

    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"extracted {len(manifest)} entries -> {out}/manifest.json")
    for m in manifest:
        lr = " [LOW-RES]" if m.get("low_res") else ""
        print(f"   [{m['method']:>22}] {m['file']} label={m['label']} page={m['page']} px={m.get('px')}{lr}")


if __name__ == "__main__":
    main()
