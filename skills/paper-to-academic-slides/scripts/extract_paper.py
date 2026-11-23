#!/usr/bin/env python3
"""Extract paper structure from PDF (or Markdown/text material):
per-page text, captions, sections, bitmap + vector figures, structured tables,
and limitations page hints.

Usage:
  python3 extract_paper.py PAPER.pdf --out /app/work/paper [--dpi 180]

Outputs (under --out):
  paper_context.json      pages text, captions (with norm_id), sections,
                          figures, tables, limitations_pages, title_guess
  figs/fig_pNN_i.png      extracted bitmap figures (page-clipped for fidelity)
  figs/vec_pNN_i.png      caption-anchored vector figure crops (matplotlib/LaTeX/TikZ)

Dependencies: pymupdf (fitz). Falls back to pdftotext or plain text parsing.
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

CAPTION_RE = re.compile(
    r"^\s*(?P<prefix>(?:Extended\s+Data\s+|Supplementary\s+)?)?"
    r"(?P<kind>Fig(?:ure)?|Tab(?:le)?|图|表|Algorithm|Listing|Scheme|Eq(?:uation)?|式)"
    r"\s*\.?\s*(?P<id>(?:[SA]?\d+[a-z]?)|(?:[IVX]{1,4}\b))"
    r"(?P<sep>[\s:：.．)|—–-]*)(?P<rest>.{0,240})",
    re.IGNORECASE,
)
INLINE_PROSE_START_RE = re.compile(
    r"^(?:shows?|illustrates?|presents?|compares?|lists?|reports?|demonstrates?|"
    r"summarizes?|depicts?|displays?|gives?|provides?|plots?|reveals?|indicates?|"
    r"confirms?|suggests?|highlights?|details?|is|are|was|were|can|may|will|"
    r"has|have|and|or|in|of|to|for|with|from|we|our|as|at|by|on)\b",
    re.IGNORECASE,
)
REF_SECTION_RE = re.compile(
    r"(?im)^\s*(?:\d+\.?\s+|[IVX]+\.?\s+)?(?:references|bibliography|acknowledg(?:e)?ments?)\s*$"
)
SECTION_RE = re.compile(
    r"^\s*(?P<num>(?:\d+(?:\.\d+){0,2}\.?)|(?:[IVX]{1,4}\.?))\s+(?P<title>[A-Z\u4e00-\u9fff][\w\-\u4e00-\u9fff ,/&:]{1,75})\s*$"
)
NAMED_SECTION_RE = re.compile(
    r"^\s*(Abstract|Introduction|Background|Related\s+Work|Methods?|Methodology|"
    r"Approach|Model|Framework|Experiments?|Experimental\s+Setup|Results?(?:\s+and\s+Discussion)?|"
    r"Evaluation|Ablation(?:\s+Study)?|Discussion|Limitations?(?:\s+and\s+Future\s+Work)?|"
    r"Conclusion(?:s)?|Summary|摘\s*要|引\s*言|方\s*法|实\s*验|结\s*果|讨\s*论|局\s*限|结\s*论)\s*$",
    re.IGNORECASE,
)
TITLE_BLACKLIST = re.compile(
    r"(?i)online|downloaded|see discussions|received|accepted|published|citation:|access|"
    r"this article|http|doi\.org|figure can be|preprint|arxiv:|copyright|open access|"
    r"proceedings of|journal of|transactions on|vol\.\s*\d"
)
STOP_SECTIONS = {"references", "acknowledg", "bibliography", "funding", "conflict of interest"}
LIMIT_RE = re.compile(r"(?i)\blimitations?\b|\bdrawbacks?\b|\bshortcomings?\b|\bfuture\s+work\b|局限|不足|限制")


def clean_text(s):
    if not s:
        return ""
    for k, v in LIGATURES.items():
        if k in s:
            s = s.replace(k, v)
    s = unicodedata.normalize("NFKC", s)
    # Rejoin hyphenated line breaks like "perfor-\nmance"
    s = re.sub(r"([A-Za-z])-\n([a-z])", r"\1\2", s)
    return s


def norm_caption_id(kind, cid):
    k = (kind or "").lower().rstrip(".")
    if k in ("fig", "figure", "图"):
        grp = "fig"
    elif k in ("tab", "table", "表"):
        grp = "tab"
    elif k in ("alg", "algorithm"):
        grp = "alg"
    elif k in ("listing",):
        grp = "listing"
    elif k in ("scheme",):
        grp = "scheme"
    elif k in ("eq", "equation", "式"):
        grp = "eq"
    else:
        grp = k
    return f"{grp}:{(cid or '').lower().strip()}"


def span_lines(page):
    """Yield (text, y0, y1, x0, x1, max_fontsize) in natural reading order
    (2-column-aware: top spanning -> left column -> right column -> bottom spanning)."""
    out = []
    try:
        pw = float(page.rect.width) or 612.0
        ph = float(page.rect.height) or 792.0
        d = page.get_text("dict")
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                txt = "".join(s.get("text", "") for s in ln.get("spans", []))
                txt = clean_text(txt)
                if not txt.strip():
                    continue
                sizes = [s.get("size", 10) for s in ln.get("spans", [])]
                bb = ln.get("bbox", [0, 0, 0, 0])
                out.append((txt, bb[1], bb[3], bb[0], bb[2], max(sizes) if sizes else 10.0))

        # Detect 2-column layout and reorder cleanly so quotes never split across columns
        mid = pw * 0.5
        left_narrow = [r for r in out if (r[4] - r[3]) < pw * 0.50 and r[4] <= mid + 18]
        right_narrow = [r for r in out if (r[4] - r[3]) < pw * 0.50 and r[3] >= mid - 18]
        if len(left_narrow) >= 4 and len(right_narrow) >= 4:
            l_span = max(r[2] for r in left_narrow) - min(r[1] for r in left_narrow)
            r_span = max(r[2] for r in right_narrow) - min(r[1] for r in right_narrow)
            if l_span >= ph * 0.08 and r_span >= ph * 0.08:
                col_top = min(min(r[1] for r in left_narrow), min(r[1] for r in right_narrow))
                col_bot = max(max(r[2] for r in left_narrow), max(r[2] for r in right_narrow))
                top_s, left_s, right_s, bot_s = [], [], [], []
                for r in out:
                    w_r = r[4] - r[3]
                    if r[2] <= col_top + 4 and w_r >= pw * 0.42:
                        top_s.append(r)
                    elif r[1] >= col_bot - 4 and w_r >= pw * 0.42:
                        bot_s.append(r)
                    elif r[4] <= mid + 18:
                        left_s.append(r)
                    elif r[3] >= mid - 18:
                        right_s.append(r)
                    elif r[1] < (col_top + col_bot) * 0.5:
                        top_s.append(r)
                    else:
                        bot_s.append(r)
                key_y = lambda r: (round(r[1], 1), round(r[3], 1))
                out = (sorted(top_s, key=key_y) + sorted(left_s, key=key_y) +
                       sorted(right_s, key=key_y) + sorted(bot_s, key=key_y))
    except Exception:
        txt = clean_text(page.get_text("text") or "")
        for i, t in enumerate(txt.splitlines()):
            if t.strip():
                out.append((t, float(i), float(i + 12), 0.0, 500.0, 10.0))
    return out


def guess_title(doc):
    try:
        page = doc[0]
        ph = page.rect.height
        best, best_size = "", 0.0
        for txt, y0, _y1, _x0, _x1, size in span_lines(page)[:60]:
            t = txt.strip()
            if not t or y0 > ph * 0.55 or TITLE_BLACKLIST.search(t):
                continue
            if 8 <= len(t) <= 180 and size > best_size and not CAPTION_RE.match(t):
                best, best_size = t, size
        return best
    except Exception:
        return ""


def render_clip(page, bbox, fn, dpi=180):
    """Render a clipped region of a PDF page to PNG (preserves white bg, rotation, vector)."""
    import pymupdf
    x0, y0, x1, y1 = bbox
    rect = pymupdf.Rect(x0, y0, x1, y1) & page.rect
    if rect.width < 40 or rect.height < 30:
        return None
    mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, clip=rect, alpha=False)
    pix.save(fn)
    if os.path.isfile(fn) and os.path.getsize(fn) > 1200:
        return {"w": int(rect.width), "h": int(rect.height),
                "bbox": [round(rect.x0, 1), round(rect.y0, 1),
                         round(rect.x1, 1), round(rect.y1, 1)]}
    try:
        os.remove(fn)
    except Exception:
        pass
    return None


def extract_bitmap_figure(page, info, outdir, page_no, idx, dpi=180, lines=None, max_y=None):
    """Extract a bitmap figure via page clip (avoids inverted SMasks & CMYK artifacts),
    gently capturing adjacent short axis/tick labels without touching prose or captions."""
    bbox = info.get("bbox")
    if not bbox or len(bbox) != 4:
        return None
    x0, y0, x1, y1 = bbox
    w, h = x1 - x0, y1 - y0
    if w < 45 or h < 35:
        return None
    ar = w / max(h, 1.0)
    if ar > 7.5 or ar < 0.14:
        return None  # banner / decorative bar
    pw, ph = page.rect.width, page.rect.height
    # skip tiny publisher logo at very top of page 1
    if page_no == 1 and y1 < ph * 0.11 and w < pw * 0.35:
        return None
    # Expand gently for short adjacent axis/tick/sub-panel labels (hybrid bitmap+PDF-text plots)
    ex0, ey0, ex1, ey1 = x0, y0, x1, y1
    if lines:
        y_cap_lim = (max_y - 3.0) if max_y is not None else ph
        for txt, ly0, ly1, lx0, lx1, _sz in lines:
            st = txt.strip()
            if not st or len(st) > 36 or CAPTION_RE.match(st):
                continue
            if ly1 >= y_cap_lim:
                continue
            if ly0 >= y0 - 20 and ly1 <= y1 + 20 and lx0 >= x0 - 28 and lx1 <= x1 + 22:
                ex0 = min(ex0, lx0)
                ey0 = min(ey0, ly0)
                ex1 = max(ex1, lx1)
                ey1 = max(ey1, ly1)
    fn = os.path.join(outdir, f"fig_p{page_no:02d}_{idx}.png")
    meta = render_clip(page, (max(0, ex0 - 3), max(0, ey0 - 3), min(pw, ex1 + 3), min(ph, ey1 + 3)), fn, dpi)
    if meta:
        return {"file": fn, "page": page_no, "source": "bitmap_clip", **meta}
    return None


def crop_vector_figure_for_caption(page, cap, all_caps_on_page, lines, outdir, page_no, idx, dpi=180):
    """Locate vector drawings / chart labels adjacent to a figure caption and render clip
    (with header/footer rule filtering and a bottom-up prose barrier)."""
    pr = page.rect
    pw, ph = pr.width, pr.height
    cy = cap["y"]
    if cy is None:
        return None

    # Gather vector drawing boxes + embedded image boxes (captures hybrid photo+vector plots)
    dboxes = []
    try:
        drawings = page.get_drawings()
        step = max(1, len(drawings) // 2500) if len(drawings) > 2500 else 1
        for dr in drawings[::step]:
            r = dr.get("rect")
            if not r:
                continue
            rw, rh = r.width, r.height
            if rw > pw * 0.94 and rh > ph * 0.90:
                continue
            if rw < 2 and rh < 2:
                continue
            # Exclude thin page header/footer separator rules
            if rh < 6.0 and (r.y1 < ph * 0.095 or r.y0 > ph * 0.91):
                continue
            dboxes.append((r.x0, r.y0, r.x1, r.y1))
    except Exception:
        pass
    try:
        for im in page.get_image_info(xrefs=False):
            b = im.get("bbox")
            if b and (b[2] - b[0]) > 15 and (b[3] - b[1]) > 15:
                dboxes.append((b[0], b[1], b[2], b[3]))
    except Exception:
        pass

    if not dboxes:
        return None

    # Find paragraph body lines to avoid eating prose
    cap_ys_above = [c["y"] for c in all_caps_on_page if c["y"] is not None and c["y"] < cy - 20]
    upper_bound = max(cap_ys_above) + 28 if cap_ys_above else ph * 0.06

    # Search above caption first (standard for Figure captions), then below if empty
    def cluster_in_band(y_lo, y_hi, above_caption=True):
        cands = [b for b in dboxes if b[1] >= y_lo - 4 and b[3] <= y_hi + 4]
        if len(cands) < 2:
            return None
        # Keep boxes in the same column region as the caption if single-column caption
        cx0 = cap.get("x0", 0)
        cx1 = cap.get("x1", pw)
        col_mode = "full"
        if (cx1 - cx0) < pw * 0.55:
            mid = pw * 0.5
            if cx1 < mid + 25:
                cands = [b for b in cands if b[0] < mid + 20]
                col_mode = "left"
            elif cx0 > mid - 25:
                cands = [b for b in cands if b[2] > mid - 20]
                col_mode = "right"
        if len(cands) < 2:
            return None

        # Bottom-up prose barrier: stop before dense body paragraphs above the figure
        if above_caption:
            mid = pw * 0.5
            prose_ys = []
            for txt, ly0, ly1, lx0, lx1, _sz in lines:
                st = txt.strip()
                if len(st) < 52 or CAPTION_RE.match(st):
                    continue
                if ly1 >= y_hi - 8 or ly0 <= y_lo:
                    continue
                if col_mode == "left" and lx0 > mid + 15:
                    continue
                if col_mode == "right" and lx1 < mid - 15:
                    continue
                prose_ys.append((ly0, ly1))
            if prose_ys:
                # Find lowest prose line that sits strictly above at least 2 figure primitives
                for _py0, py1 in sorted(prose_ys, key=lambda t: t[1], reverse=True):
                    below_prose = [b for b in cands if b[1] >= py1 - 2]
                    if len(below_prose) >= 2:
                        span_h = max(b[3] for b in below_prose) - min(b[1] for b in below_prose)
                        if span_h >= ph * 0.07:
                            cands = below_prose
                            break

        x0 = min(b[0] for b in cands)
        y0 = min(b[1] for b in cands)
        x1 = max(b[2] for b in cands)
        y1 = max(b[3] for b in cands)
        # Expand slightly to include axis tick/legend short text lines inside/near drawings
        for txt, ly0, ly1, lx0, lx1, _sz in lines:
            st = txt.strip()
            if len(st) <= 45 and not CAPTION_RE.match(st) and ly0 >= y0 - 22 and ly1 <= y1 + 14 and ly1 < cy - 2:
                if lx0 >= x0 - 35 and lx1 <= x1 + 35:
                    x0 = min(x0, lx0)
                    y0 = min(y0, ly0)
                    x1 = max(x1, lx1)
                    y1 = max(y1, ly1)
        if (x1 - x0) < pw * 0.16 or (y1 - y0) < ph * 0.08:
            return None
        if (x1 - x0) * (y1 - y0) < 0.022 * pw * ph:
            return None
        return (max(0, x0 - 8), max(0, y0 - 8), min(pw, x1 + 8), min(ph, y1 + 6))

    bbox = cluster_in_band(upper_bound, cy - 3, above_caption=True)
    if bbox is None:
        cap_ys_below = [c["y"] for c in all_caps_on_page if c["y"] is not None and c["y"] > cy + 20]
        lower_bound = min(cap_ys_below) - 10 if cap_ys_below else ph * 0.94
        bbox = cluster_in_band(cy + 16, lower_bound, above_caption=False)
    if bbox is None:
        return None

    fn = os.path.join(outdir, f"vec_p{page_no:02d}_{idx}.png")
    meta = render_clip(page, bbox, fn, dpi)
    if meta:
        return {"file": fn, "page": page_no, "source": "vector_crop",
                "caption": cap["id"], "norm_id": cap["norm_id"], **meta}
    return None


def extract_page_tables(page, page_no, page_caps, page_text=""):
    """Extract structured tables via PyMuPDF find_tables (with booktabs fallback)."""
    out = []
    if not hasattr(page, "find_tables"):
        return out
    has_tcap = any(c["kind"] == "table" for c in page_caps)
    # Fast-skip pure prose / reference pages without any table caption or table keyword
    if not has_tcap and not re.search(r"(?:\btable\b|\btab\.|表\s*\d)", page_text or "", re.I):
        return out
    tbls = []
    try:
        tf = page.find_tables()
        tbls = getattr(tf, "tables", []) or []
    except Exception:
        tbls = []
    # Booktabs LaTeX fallback (horizontal rules only, no vertical lines)
    if not tbls and has_tcap:
        try:
            tf2 = page.find_tables(vertical_strategy="text", horizontal_strategy="lines")
            tbls = getattr(tf2, "tables", []) or []
        except Exception:
            tbls = []
    # Borderless scientific table fallback (whitespace-aligned columns near a Table caption)
    if not tbls and has_tcap:
        try:
            import pymupdf
            tcaps_p = [c for c in page_caps if c["kind"] == "table" and c["y"] is not None]
            pw_t, ph_t = page.rect.width, page.rect.height
            for tc in tcaps_p[:2]:
                band_y0 = max(0.0, tc["y"] + 8.0)
                band_y1 = min(ph_t, tc["y"] + ph_t * 0.48)
                if band_y1 - band_y0 < 50:
                    continue
                tf3 = page.find_tables(
                    clip=pymupdf.Rect(0, band_y0, pw_t, band_y1),
                    vertical_strategy="text",
                    horizontal_strategy="text",
                )
                for cand in (getattr(tf3, "tables", []) or [])[:1]:
                    raw_c = cand.extract() or []
                    clean_c = [[clean_text(str(x or "")).strip() for x in rr]
                               for rr in raw_c if any((x or "").strip() for x in rr)]
                    if len(clean_c) >= 3 and len(clean_c[0]) >= 3:
                        n_num = sum(1 for rr in clean_c[1:] for x in rr if re.search(r"\d", x))
                        if n_num >= 3:
                            tbls.append(cand)
        except Exception:
            pass
    for ti, tbl in enumerate(tbls[:4]):
        try:
            # Read each cell via page.get_text(clip=...) to avoid find_tables char-order quirks (e.g. '712 .')
            import pymupdf
            rows = []
            if getattr(tbl, "rows", None):
                for r in tbl.rows:
                    rvals = []
                    r_lines_per_col = []
                    for cell in r.cells:
                        if cell is None:
                            rvals.append("")
                            r_lines_per_col.append([])
                        else:
                            raw_ct = clean_text(page.get_text("text", clip=pymupdf.Rect(cell)) or "")
                            clines = [re.sub(r"\s+", " ", ln).strip() for ln in raw_ct.splitlines() if ln.strip()]
                            r_lines_per_col.append(clines)
                            rvals.append(" ".join(clines).strip())
                    # Detect LaTeX booktabs merged mega-row (no horizontal lines between data rows)
                    non_empty_counts = [len(cl) for cl in r_lines_per_col if cl]
                    if (len(r_lines_per_col) >= 2 and len(non_empty_counts) >= 2 and
                            max(non_empty_counts) >= 3 and
                            min(non_empty_counts) == max(non_empty_counts)):
                        n_sub = max(non_empty_counts)
                        for si_r in range(n_sub):
                            sub_row = [cl[si_r] if si_r < len(cl) else "" for cl in r_lines_per_col]
                            if any(sub_row):
                                rows.append(sub_row)
                    elif any(rvals):
                        rows.append(rvals)
            # If booktabs still yielded only 2 giant rows with multi-line data, re-split via text-strategy inside bbox
            if len(rows) <= 2 and getattr(tbl, "bbox", None):
                try:
                    tf_sub = page.find_tables(
                        clip=pymupdf.Rect(tbl.bbox),
                        vertical_strategy="text",
                        horizontal_strategy="text",
                    )
                    sub_tbls = getattr(tf_sub, "tables", []) or []
                    if sub_tbls:
                        sub_raw = sub_tbls[0].extract() or []
                        sub_rows = [[clean_text(str(c or "")).replace("\n", " ").strip() for c in rr]
                                    for rr in sub_raw if any((c or "").strip() for c in rr)]
                        if len(sub_rows) > len(rows) and len(sub_rows[0]) >= 2:
                            rows = sub_rows
                except Exception:
                    pass
            if not rows:
                raw = tbl.extract() or []
                rows = [[clean_text(str(c or "")).replace("\n", " ").strip() for c in r]
                        for r in raw if any((c or "").strip() for c in r)]
            if len(rows) < 2 or len(rows[0]) < 2:
                continue
            # Strip completely empty columns (common with merged LaTeX booktabs cells)
            max_c = max(len(r) for r in rows)
            non_empty_cols = [ci for ci in range(max_c)
                              if any(ci < len(r) and r[ci].strip() for r in rows)]
            if len(non_empty_cols) >= 2 and len(non_empty_cols) < max_c:
                rows = [[r[ci] if ci < len(r) else "" for ci in non_empty_cols] for r in rows]
            bb = list(tbl.bbox) if getattr(tbl, "bbox", None) else [0, 0, 0, 0]
            # Match nearest Table caption on same page
            tcaps = [c for c in page_caps if c["kind"] == "table" and c["y"] is not None]
            cap_id = None
            if tcaps:
                cap_id = min(tcaps, key=lambda c: abs(c["y"] - bb[1]))["id"]
            out.append({
                "page": page_no,
                "table_idx": ti,
                "caption": cap_id,
                "n_rows_total": len(rows) - 1,
                "headers": rows[0][:9],
                "rows": [r[:9] for r in rows[1:11]],
            })
        except Exception:
            continue
    return out


QUANT_HIGHLIGHT_RE = re.compile(
    r"(?:\d+(?:\.\d+)?\s*(?:%|×|x\b|GB|MB|KB|ms|s\b|tok/s|fps|dB|Hz|MHz|GHz|nm|mm|km|eV|kcal|mol|W\b|mW|FLOPs|GFLOPs|TFLOPs|params|M\b|B\b))"
    r"|(?:\b(?:F1|AUC|ROC|RMSE|MAE|BLEU|ROUGE|PSNR|SSIM|IoU|mAP|accuracy|acc|top-1|perplexity|PPL|speedup|throughput|latency|error|loss|outperform|improv|reduc|提升|降低|加速|准确率|精度)\b.{0,28}\d+(?:\.\d+)?)",
    re.IGNORECASE,
)


def extract_key_highlights(pages, max_total=16):
    """Extract high-signal verbatim quantitative sentences (skipping references) so
    the agent grasps key metrics, units, and page locations in a single read."""
    out = []
    in_refs = False
    for pg in pages or []:
        pno = pg.get("page", 0)
        txt = pg.get("text") or ""
        if REF_SECTION_RE.search(txt):
            in_refs = True
        if in_refs:
            continue
        # Split gently on sentence boundaries (English & CJK)
        flat = re.sub(r"\s+", " ", txt).strip()
        sents = re.split(r"(?<=[.!?。！？])\s+", flat)
        picked = 0
        for s in sents:
            st = s.strip()
            if len(st) < 28 or len(st) > 220:
                continue
            if CAPTION_RE.match(st) or TITLE_BLACKLIST.search(st):
                continue
            if QUANT_HIGHLIGHT_RE.search(st):
                out.append({"page": pno, "text": st})
                picked += 1
                if picked >= 2 or len(out) >= max_total:
                    break
        if len(out) >= max_total:
            break
    return out


def recommend_key_pages(sections, limit_pages, n_pages, tables=None, key_highlights=None):
    """Recommend 3-5 highest-value pages (Method, Main Results/Tables, Limitations, Conclusion)
    for fast targeted reading under tight step budgets."""
    role_pats = [
        ("method", re.compile(r"(?i)method|approach|model|framework|architecture|algorithm|方法|模型|架构")),
        ("results", re.compile(r"(?i)experiment|result|evaluation|benchmark|ablation|实验|结果|消融")),
        ("conclusion", re.compile(r"(?i)conclusion|summary|discussion|结论|总结|讨论")),
    ]
    rec = []
    for _role, pat in role_pats:
        for s in sections or []:
            if pat.search(s.get("title") or ""):
                p = s.get("page")
                if p and p not in rec:
                    rec.append(p)
                    break
    # Include main table pages (often separate from the Experimental Setup start page!)
    for t in (tables or [])[:2]:
        tp = t.get("page")
        if tp and tp not in rec:
            rec.append(tp)
    # Include page with highest quantitative highlight density if still room
    if key_highlights:
        hl_cnt = {}
        for h in key_highlights:
            hp = h.get("page")
            if hp and hp > 1:
                hl_cnt[hp] = hl_cnt.get(hp, 0) + 1
        for hp, _c in sorted(hl_cnt.items(), key=lambda kv: (-kv[1], kv[0])):
            if hp not in rec and len(rec) < 5:
                rec.append(hp)
                break
    for lp in (limit_pages or [])[:1]:
        if lp not in rec:
            rec.append(lp)
    if not rec and n_pages:
        rec = [p for p in (1, max(2, n_pages // 2), n_pages) if 1 <= p <= n_pages]
    return sorted(set(rec))[:5]


SCHEMATIC_HINT_RE = re.compile(
    r"\b(?:overview|architecture|framework|pipeline|workflow|schematic|diagram|illustration|flowchart|topology|structure)\b"
    r"|(?:框架|流程|架构|示意|原理图|总览图|结构图)",
    re.IGNORECASE,
)


def detect_tesseract_lang():
    """Return best available tesseract language string ('eng+chi_sim' or 'eng'), or None."""
    import shutil
    import subprocess
    if not shutil.which("tesseract"):
        return None
    try:
        r = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True, timeout=6)
        langs = set(ln.strip().lower() for ln in (r.stdout + "\n" + r.stderr).splitlines())
        if "chi_sim" in langs and "eng" in langs:
            return "eng+chi_sim"
        if "eng" in langs:
            return "eng"
    except Exception:
        return "eng"
    return "eng"


def ocr_pdf_page(page, lang="eng", dpi=200):
    """Run tesseract OCR on a low-text/scanned PDF page and return cleaned text."""
    import subprocess
    import tempfile
    tmp_png = None
    try:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        fd, tmp_png = tempfile.mkstemp(prefix="ocr_p_", suffix=".png")
        os.close(fd)
        pix.save(tmp_png)
        r = subprocess.run(
            ["tesseract", tmp_png, "stdout", "-l", lang, "--psm", "6"],
            capture_output=True, text=True, timeout=25,
        )
        if r.returncode == 0 and r.stdout:
            return clean_text(r.stdout)
    except Exception:
        pass
    finally:
        if tmp_png and os.path.exists(tmp_png):
            try:
                os.remove(tmp_png)
            except Exception:
                pass
    return ""


def write_paper_index(ctx, outdir):
    """Write a high-priority-first paper_index.md so figures, tables, sections,
    and captions appear at the top and are never lost if downstream excerpts are capped."""
    rec_pages = ctx.get("recommended_read_pages") or []
    cap_by_norm = {c.get("norm_id"): c.get("text", "") for c in (ctx.get("captions") or []) if c.get("norm_id")}
    lines = [
        f"# Paper Index: {ctx.get('title_guess') or '(untitled)'}",
        f"- Source: `{ctx.get('source_pdf', '')}` ({ctx.get('source_type', 'pdf')}, {ctx.get('n_pages', 0)} pages)",
        f"- Abstract page: {ctx.get('abstract_page')} | Limitations pages: {ctx.get('limitations_pages') or []} | Recommended read pages: {rec_pages}",
    ]
    if ctx.get("ocr_pages"):
        lines.append(f"- OCR-recovered scanned pages: {ctx['ocr_pages']}")
    if ctx.get("limits"):
        lines.append(f"- Extraction limits: {'; '.join(ctx['limits'])}")

    figs = ctx.get("figures") or []
    if figs:
        lines.append("\n## Extracted Figures")
        for f in figs[:30]:
            cap = f.get("caption") or "(unmatched)"
            fw, fh = f.get("w") or 0, f.get("h") or 0
            ar = round(fw / max(fh, 1), 2) if fw and fh else None
            ar_tag = f", ar={ar}{' banner' if ar and ar >= 2.2 else ''}" if ar else ""
            ctext = cap_by_norm.get(f.get("norm_id"), "")
            sch_tag = " [SCHEMATIC — label non-experimental]" if (ctext and SCHEMATIC_HINT_RE.search(ctext)) else ""
            lines.append(f"- p.{f['page']} | {cap} | `{f['file']}` ({f.get('source', 'clip')}, {fw}x{fh}{ar_tag}){sch_tag}")

    tbls = ctx.get("tables") or []
    if tbls:
        lines.append("\n## Extracted Tables (verbatim)")
        for t in tbls[:10]:
            lines.append(f"\n### {t.get('caption') or 'Table'} (p.{t['page']}, {t.get('n_rows_total')} rows)")
            hdrs = t.get("headers") or []
            if hdrs:
                lines.append("| " + " | ".join(hdrs) + " |")
                lines.append("| " + " | ".join(["---"] * len(hdrs)) + " |")
            for r in (t.get("rows") or [])[:6]:
                pad = list(r) + [""] * max(0, len(hdrs) - len(r))
                lines.append("| " + " | ".join(pad[:max(len(hdrs), len(r))]) + " |")

    secs = ctx.get("sections") or []
    if secs:
        lines.append("\n## Sections")
        for s in secs[:35]:
            num = (s.get("num") or "").strip()
            lines.append(f"- p.{s['page']}: {num} {s['title']}".replace("  ", " "))

    caps = ctx.get("captions") or []
    if caps:
        lines.append("\n## Captions")
        for c in caps[:40]:
            lines.append(f"- [{c.get('norm_id', '')}] **{c['id']}** (p.{c['page']}): {c.get('text', '')[:160]}")

    hls = ctx.get("key_highlights") or []
    if hls:
        lines.append("\n## Key Quantitative Highlights (verbatim)")
        for h in hls[:16]:
            lines.append(f"- p.{h['page']}: {h['text']}")

    pages = ctx.get("pages") or []
    if pages and (pages[0].get("text") or "").strip():
        p1_raw = pages[0]["text"].strip()
        p1_snip = p1_raw[:1400]
        if len(p1_raw) > 1400 and "\n" in p1_snip:
            p1_snip = p1_snip.rsplit("\n", 1)[0]
        lines.append("\n## Page 1 Excerpt (Title / Authors / Abstract)")
        lines.append(p1_snip)

    idx_path = os.path.join(outdir, "paper_index.md")
    with open(idx_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    # Also write individual per-page text files for zero-truncation targeted reading
    pdir = os.path.join(outdir, "pages")
    os.makedirs(pdir, exist_ok=True)
    for pg in ctx.get("pages") or []:
        pno = pg.get("page", 0)
        pfn = os.path.join(pdir, f"p{int(pno):02d}.txt")
        with open(pfn, "w", encoding="utf-8") as pf:
            pf.write(pg.get("text", "") or "")
    return idx_path


def extract_docx_text(path):
    """Extract paragraphs and Markdown-formatted tables from .docx using stdlib zipfile+xml."""
    import zipfile
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_p = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
    w_tbl = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"
    body = root.find("w:body", ns)
    if body is None:
        return "\n".join(
            "".join(t.text or "" for t in p.findall(".//w:t", ns)).strip()
            for p in root.findall(".//w:p", ns)
            if "".join(t.text or "" for t in p.findall(".//w:t", ns)).strip()
        )
    out = []
    for child in body:
        if child.tag == w_p:
            txt = "".join(t.text or "" for t in child.findall(".//w:t", ns)).strip()
            if txt:
                out.append(txt)
        elif child.tag == w_tbl:
            rows = []
            for tr in child.findall("w:tr", ns):
                cells = []
                for tc in tr.findall("w:tc", ns):
                    ct = " ".join(
                        "".join(t.text or "" for t in p.findall(".//w:t", ns)).strip()
                        for p in tc.findall(".//w:p", ns)
                    ).strip().replace("|", "/")
                    cells.append(ct)
                if any(cells):
                    rows.append(cells)
            if len(rows) >= 2 and len(rows[0]) >= 2:
                out.append("| " + " | ".join(rows[0]) + " |")
                out.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
                for r in rows[1:]:
                    pad = r + [""] * max(0, len(rows[0]) - len(r))
                    out.append("| " + " | ".join(pad[:len(rows[0])]) + " |")
    return "\n".join(out)


def extract_text_material(path, outdir):
    """Support Markdown/plain-text/docx research materials with identical paper_context schema."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        raw = clean_text(extract_docx_text(path))
    else:
        raw = clean_text(open(path, "r", encoding="utf-8", errors="ignore").read())
    chunks = [c.strip() for c in re.split(r"\n(?=#+\s)", raw) if c.strip()]
    if not chunks:
        chunks = [raw]
    pages, sections, captions, tables, limit_pages = [], [], [], [], []
    title = ""
    for i, ch in enumerate(chunks, 1):
        pages.append({"page": i, "nchars": len(ch), "text": ch})
        if LIMIT_RE.search(ch):
            limit_pages.append(i)
        for ln in ch.splitlines():
            hm = re.match(r"^(#{1,3})\s+(.+)", ln.strip())
            if hm:
                if not title and len(hm.group(1)) == 1:
                    title = hm.group(2).strip()
                sections.append({"page": i, "num": hm.group(1), "title": hm.group(2).strip()[:80]})
            cm = CAPTION_RE.match(ln.strip())
            if cm:
                cid = f"{cm.group('kind')} {cm.group('id')}"
                captions.append({
                    "kind": norm_caption_id(cm.group("kind"), cm.group("id")).split(":")[0],
                    "id": cid,
                    "norm_id": norm_caption_id(cm.group("kind"), cm.group("id")),
                    "page": i, "y": None, "text": (cm.group("rest") or "").strip()[:220],
                })
        # Markdown pipe tables
        tlines = [ln.strip() for ln in ch.splitlines() if "|" in ln and ln.strip().startswith("|")]
        if len(tlines) >= 3:
            parsed = [[c.strip() for c in ln.strip("|").split("|")]
                      for ln in tlines if not re.match(r"^[\|\s:\-]+$", ln)]
            if len(parsed) >= 2:
                tables.append({"page": i, "table_idx": len(tables), "caption": None,
                               "n_rows_total": len(parsed) - 1,
                               "headers": parsed[0][:8], "rows": [r[:8] for r in parsed[1:9]]})
    key_hls = extract_key_highlights(pages)
    rec_pages = recommend_key_pages(sections[:80], limit_pages, len(pages), tables=tables, key_highlights=key_hls)
    ctx = {
        "source_pdf": os.path.abspath(path),
        "source_type": "text",
        "title_guess": title or os.path.basename(path),
        "n_pages": len(pages),
        "abstract_page": 1,
        "limitations_pages": limit_pages,
        "recommended_read_pages": rec_pages,
        "key_highlights": key_hls,
        "pages": pages,
        "captions": captions,
        "sections": sections[:80],
        "figures": [],
        "tables": tables,
        "limits": [],
    }
    outjson = os.path.join(outdir, "paper_context.json")
    os.makedirs(outdir, exist_ok=True)
    with open(outjson, "w", encoding="utf-8") as f:
        json.dump(ctx, f, ensure_ascii=False, indent=1)
    idx_path = write_paper_index(ctx, outdir)
    print(json.dumps({"ok": True, "out": outjson, "index": idx_path, "source_type": "text",
                      "n_pages": len(pages), "captions": len(captions),
                      "figures": 0, "tables": len(tables),
                      "key_highlights": len(key_hls),
                      "recommended_read_pages": rec_pages,
                      "limitations_pages": limit_pages,
                      "title_guess": ctx["title_guess"], "limits": []}, ensure_ascii=False))
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", default="/app/work/paper")
    ap.add_argument("--dpi", type=int, default=180)
    ap.add_argument("--max-pages", type=int, default=60)
    a = ap.parse_args()

    if os.path.splitext(a.pdf)[1].lower() in (".md", ".txt", ".docx"):
        return extract_text_material(a.pdf, a.out)

    limits = []
    figdir = os.path.join(a.out, "figs")
    os.makedirs(figdir, exist_ok=True)

    doc = None
    try:
        import pymupdf
        doc = pymupdf.open(a.pdf)
    except Exception as e:
        limits.append(f"pymupdf open failed: {e}; trying pdftotext fallback")

    pages, captions, sections, figures, tables, limit_pages, ocr_pages = [], [], [], [], [], [], []
    title = ""
    ocr_lang = None
    ocr_checked = False

    if doc is not None:
        if doc.needs_pass:
            try:
                doc.authenticate("")
            except Exception:
                limits.append("PDF appears encrypted; text/images may be empty")
        n = min(doc.page_count, a.max_pages)
        if doc.page_count > a.max_pages:
            limits.append(f"page cap {a.max_pages} < {doc.page_count}; truncated")

        in_refs = False
        for i in range(n):
            page = doc[i]
            lines = span_lines(page)
            text = clean_text("\n".join(t.strip() for t, *_ in lines))
            ocr_used = False
            # Automatic Tesseract OCR fallback for scanned / image-only pages (< 120 native chars)
            if len(text.strip()) < 120 and len(ocr_pages) < 15:
                if not ocr_checked:
                    ocr_lang = detect_tesseract_lang()
                    ocr_checked = True
                if ocr_lang:
                    ocr_txt = ocr_pdf_page(page, lang=ocr_lang, dpi=200)
                    if len(ocr_txt.strip()) > len(text.strip()) + 40:
                        text = ocr_txt
                        ocr_used = True
                        ocr_pages.append(i + 1)
                        # Populate synthetic lines for caption/section parsing on OCR pages
                        for oln in text.splitlines():
                            ost = oln.strip()
                            if not ost:
                                continue
                            cm_o = CAPTION_RE.match(ost)
                            if cm_o:
                                k_raw = cm_o.group("kind").lower().rstrip(".")
                                kind_o = {"fig": "figure", "figure": "figure", "图": "figure",
                                          "tab": "table", "table": "table", "表": "table",
                                          "scheme": "figure", "eq": "equation", "equation": "equation"}.get(k_raw, k_raw)
                                captions.append({
                                    "kind": kind_o,
                                    "id": f"{cm_o.group('kind')} {cm_o.group('id')}",
                                    "norm_id": norm_caption_id(cm_o.group("kind"), cm_o.group("id")),
                                    "page": i + 1, "y": None,
                                    "text": (cm_o.group("rest") or "").strip()[:300],
                                })
                            sm_o = SECTION_RE.match(ost)
                            if sm_o and len(ost) < 85 and not any(sw in sm_o.group("title").lower() for sw in STOP_SECTIONS):
                                sections.append({"page": i + 1, "num": sm_o.group("num"), "title": sm_o.group("title").strip()})
                            elif NAMED_SECTION_RE.match(ost) and len(ost) < 50:
                                sections.append({"page": i + 1, "num": "", "title": ost})
                        if not title and i == 0:
                            for oln in text.splitlines()[:8]:
                                oln_s = oln.strip()
                                if 12 <= len(oln_s) <= 180 and not TITLE_BLACKLIST.search(oln_s):
                                    title = oln_s
                                    break
            pages.append({"page": i + 1, "nchars": len(text), "text": text, "ocr_used": ocr_used})
            if REF_SECTION_RE.search(text):
                # Check if limitations text appears BEFORE References on this same page
                pre_ref = REF_SECTION_RE.split(text, maxsplit=1)[0]
                if not in_refs and LIMIT_RE.search(pre_ref):
                    limit_pages.append(i + 1)
                in_refs = True
            elif not in_refs and LIMIT_RE.search(text):
                limit_pages.append(i + 1)

            max_sz = max((s for *_, s in lines), default=10.0)
            page_caps = []
            for li_idx, (txt, y0, _y1, x0, x1, size) in enumerate(lines):
                st = txt.strip()
                m = CAPTION_RE.match(st)
                if m and size < 1.6 * max_sz:
                    raw_id = m.group("id")
                    prefix = (m.group("prefix") or "").strip()
                    sep = m.group("sep") or ""
                    rest = (m.group("rest") or "").strip()
                    # Roman IDs must be strictly uppercase (avoid matching lowercase words like 'i'/'in')
                    if re.fullmatch(r"[ivxIVX]{1,4}", raw_id) and raw_id != raw_id.upper():
                        m = None
                    # Skip inline body prose mentions like "Fig. 3 shows that..."
                    has_punct_sep = bool(re.search(r"[:：.．|—–-]", sep))
                    if m and not has_punct_sep and rest and INLINE_PROSE_START_RE.match(rest):
                        m = None
                    # Support standalone header line like "TABLE I" + up to 3 continuation caption lines
                    if m:
                        prev_y1 = _y1
                        cont_parts = [rest] if rest else []
                        for nxt_i in range(li_idx + 1, min(li_idx + 4, len(lines))):
                            ntxt, ny0, ny1, nx0, nx1, nsz = lines[nxt_i]
                            nst = ntxt.strip()
                            if not nst:
                                break
                            dy = ny0 - prev_y1
                            if dy < -2 or dy > 16:
                                break
                            if CAPTION_RE.match(nst) or SECTION_RE.match(nst) or NAMED_SECTION_RE.match(nst):
                                break
                            if nsz > size + 1.2:
                                break
                            # Same column check (avoid merging across two-column gutter)
                            if (nx1 < x0 - 30) or (nx0 > x1 + 30):
                                break
                            cont_parts.append(nst)
                            prev_y1 = ny1
                            if len(" ".join(cont_parts)) >= 320:
                                break
                        rest = " ".join(p for p in cont_parts if p).strip()
                    if m and (len(rest) >= 3 or re.match(r"^\s*[\w(\u4e00-\u9fff]", rest) or has_punct_sep):
                        k_raw = m.group("kind").lower().rstrip(".")
                        kind = {"fig": "figure", "figure": "figure", "图": "figure",
                                "tab": "table", "table": "table", "表": "table",
                                "algorithm": "algorithm", "listing": "listing",
                                "scheme": "figure", "eq": "equation", "equation": "equation",
                                "式": "equation"}.get(k_raw, k_raw)
                        cid = f"{prefix} {m.group('kind')} {raw_id}".strip() if prefix else f"{m.group('kind')} {raw_id}"
                        c_obj = {
                            "kind": kind,
                            "id": cid,
                            "norm_id": norm_caption_id(m.group("kind"), raw_id),
                            "page": i + 1, "y": round(y0, 1),
                            "x0": round(x0, 1), "x1": round(x1, 1),
                            "text": rest[:340],
                        }
                        page_caps.append(c_obj)
                        captions.append(c_obj)

                sm = SECTION_RE.match(st)
                if sm and size >= 10.0 and len(st) < 85 and \
                        not any(st_w in sm.group("title").lower() for st_w in STOP_SECTIONS):
                    sections.append({"page": i + 1, "num": sm.group("num"),
                                     "title": sm.group("title").strip()})
                elif NAMED_SECTION_RE.match(st) and size >= 10.0 and len(st) < 50:
                    sections.append({"page": i + 1, "num": "", "title": st})

            # 1) Bitmap images on page (2-column-aware caption matching + multi-panel composite merge)
            try:
                infos = page.get_image_info(xrefs=True)
            except Exception:
                infos = []
            parea = abs(page.rect)
            pw_p, ph_p = page.rect.width, page.rect.height
            page_figs = []

            def _match_caption_for_bbox(bx0, by0, bx1, by1):
                """Match figure caption preferring same-column X overlap, then min vertical distance."""
                def _x_aligned(c):
                    cx0 = c.get("x0", 0.0)
                    cx1 = c.get("x1", pw_p)
                    if (cx1 - cx0) >= pw_p * 0.55 or (bx1 - bx0) >= pw_p * 0.55:
                        return True
                    ov = min(bx1, cx1) - max(bx0, cx0)
                    return ov >= -25.0

                below = [c for c in page_caps if c["kind"] == "figure" and c["y"] is not None
                         and c["y"] >= by1 - 20]
                if below:
                    aligned = [c for c in below if _x_aligned(c)]
                    pool = aligned or below
                    return min(pool, key=lambda c: c["y"] - by1)
                above = [c for c in page_caps if c["kind"] == "figure" and c["y"] is not None
                         and 0 <= by0 - c["y"] <= 85]
                if above:
                    aligned = [c for c in above if _x_aligned(c)]
                    pool = aligned or above
                    return min(pool, key=lambda c: by0 - c["y"])
                return None

            for j, info in enumerate(infos):
                b = info.get("bbox")
                if not b or (b[2] - b[0]) * (b[3] - b[1]) < 0.015 * parea:
                    continue
                best_c = _match_caption_for_bbox(b[0], b[1], b[2], b[3])
                cap_y_lim = best_c["y"] if (best_c and best_c["y"] is not None and best_c["y"] >= b[3] - 20) else None
                f = extract_bitmap_figure(page, info, figdir, i + 1, j, a.dpi, lines=lines, max_y=cap_y_lim)
                if f:
                    if best_c:
                        f["caption"] = best_c["id"]
                        f["norm_id"] = best_c["norm_id"]
                    page_figs.append(f)

            # Merge multi-panel bitmap sub-figures sharing the same caption into a composite figure
            by_norm = {}
            for pf in page_figs:
                nid = pf.get("norm_id")
                if nid:
                    by_norm.setdefault(nid, []).append(pf)
            comp_figs = []
            for comp_i, (nid, grp) in enumerate(by_norm.items()):
                if len(grp) >= 2:
                    ux0 = min(g["bbox"][0] for g in grp)
                    uy0 = min(g["bbox"][1] for g in grp)
                    ux1 = max(g["bbox"][2] for g in grp)
                    uy1 = max(g["bbox"][3] for g in grp)
                    # Only merge if union bbox is reasonably compact (<= 75% of page height)
                    if (uy1 - uy0) <= ph_p * 0.75:
                        cap_obj = next((c for c in page_caps if c.get("norm_id") == nid), None)
                        cap_lim = cap_obj["y"] - 3.0 if (cap_obj and cap_obj["y"] is not None and cap_obj["y"] >= uy1 - 15) else ph_p
                        for txt, ly0, ly1, lx0, lx1, _sz in lines:
                            st = txt.strip()
                            if len(st) <= 40 and not CAPTION_RE.match(st) and ly1 < cap_lim:
                                if ly0 >= uy0 - 18 and ly1 <= uy1 + 18 and lx0 >= ux0 - 20 and lx1 <= ux1 + 20:
                                    ux0, uy0 = min(ux0, lx0), min(uy0, ly0)
                                    ux1, uy1 = max(ux1, lx1), max(uy1, ly1)
                        cfn = os.path.join(figdir, f"fig_p{i+1:02d}_comp_{comp_i}.png")
                        cmeta = render_clip(page, (max(0, ux0 - 4), max(0, uy0 - 4), min(pw_p, ux1 + 4), min(ph_p, uy1 + 4)), cfn, a.dpi)
                        if cmeta:
                            comp_figs.append({
                                "file": cfn, "page": i + 1, "source": "composite_panels",
                                "caption": grp[0].get("caption"), "norm_id": nid,
                                "n_panels": len(grp), **cmeta,
                            })
            # List composite merged figures first so agent naturally selects the complete multi-panel figure
            figures.extend(comp_figs)
            figures.extend(page_figs)

            # 2) Vector figure fallback for any figure caption on this page without a bitmap
            matched_norms = {f.get("norm_id") for f in page_figs if f.get("norm_id")}
            v_idx = 0
            for c in page_caps:
                if c["kind"] != "figure" or c["norm_id"] in matched_norms:
                    continue
                vf = crop_vector_figure_for_caption(page, c, page_caps, lines, figdir, i + 1, v_idx, a.dpi)
                if vf:
                    figures.append(vf)
                    matched_norms.add(c["norm_id"])
                    v_idx += 1

            # 3) Structured tables (skip pure references pages without captions)
            if (not in_refs) or page_caps:
                tables.extend(extract_page_tables(page, i + 1, page_caps, page_text=text))

        title = guess_title(doc)
        try:
            meta = doc.metadata or {}
            if not title:
                title = clean_text((meta.get("title") or "").strip())
        except Exception:
            pass
        doc.close()
    else:
        import subprocess
        try:
            r = subprocess.run(["pdftotext", a.pdf, "-"], capture_output=True, timeout=120)
            txt = clean_text(r.stdout.decode("utf-8", "ignore"))
            for i, chunk in enumerate(txt.split("\f")):
                if chunk.strip():
                    pages.append({"page": i + 1, "nchars": len(chunk), "text": chunk})
                    if LIMIT_RE.search(chunk):
                        limit_pages.append(i + 1)
            for m in re.finditer(
                r"(?m)^\s*(Fig(?:ure)?\.?|Tab(?:le)?\.?|Scheme|图|表)\s*([SA]?\d+[a-z]?)[\s:：.．](.{0,200})",
                txt, re.IGNORECASE,
            ):
                k = m.group(1).lower().rstrip(".")
                captions.append({
                    "kind": "figure" if k in ("fig", "figure", "scheme", "图") else "table",
                    "id": f"{m.group(1)} {m.group(2)}",
                    "norm_id": norm_caption_id(m.group(1), m.group(2)),
                    "page": txt[:m.start()].count("\f") + 1,
                    "y": None, "text": m.group(3).strip()[:200],
                })
        except Exception as e:
            limits.append(f"pdftotext fallback failed: {e}")

    abstract_page = None
    for p in pages[:3]:
        if re.search(r"(?im)^\s*abstract\b|摘\s*要", p["text"]):
            abstract_page = p["page"]
            break

    # Deduplicate captions by (norm_id, page)
    seen = set()
    deduped = []
    for c in captions:
        k = (c.get("norm_id") or c["id"].lower(), c["page"])
        if k not in seen:
            seen.add(k)
            c_clean = {kk: vv for kk, vv in c.items() if kk not in ("x0", "x1")}
            deduped.append(c_clean)
    captions = deduped

    # Deduplicate sections
    seen_s, sec_dedup = set(), []
    for s in sections:
        sk = (s["page"], s["title"].lower())
        if sk not in seen_s:
            seen_s.add(sk)
            sec_dedup.append(s)

    total_chars = sum(p["nchars"] for p in pages)
    if ocr_pages:
        limits.append(f"tesseract OCR applied to {len(ocr_pages)} scanned/low-text page(s): {ocr_pages}")
    if total_chars < 500:
        limits.append(f"low extracted text ({total_chars} chars); PDF may be scanned/image-only")

    for fg in figures:
        fw, fh = fg.get("w") or 0, fg.get("h") or 0
        if fw and fh:
            fg["aspect"] = round(fw / float(fh), 2)

    key_hls = extract_key_highlights(pages)
    rec_pages = recommend_key_pages(sec_dedup[:80], limit_pages, len(pages), tables=tables, key_highlights=key_hls)

    ctx = {
        "source_pdf": os.path.abspath(a.pdf),
        "source_type": "pdf",
        "title_guess": title,
        "n_pages": len(pages),
        "abstract_page": abstract_page,
        "limitations_pages": limit_pages,
        "ocr_pages": ocr_pages,
        "recommended_read_pages": rec_pages,
        "key_highlights": key_hls,
        "pages": pages,
        "captions": captions,
        "sections": sec_dedup[:80],
        "figures": figures,
        "tables": tables[:20],
        "limits": limits,
    }
    outjson = os.path.join(a.out, "paper_context.json")
    os.makedirs(a.out, exist_ok=True)
    with open(outjson, "w", encoding="utf-8") as f:
        json.dump(ctx, f, ensure_ascii=False, indent=1)
    idx_path = write_paper_index(ctx, a.out)
    print(json.dumps({"ok": True, "out": outjson, "index": idx_path, "n_pages": len(pages),
                      "captions": len(captions), "figures": len(figures),
                      "tables": len(tables), "limitations_pages": limit_pages,
                      "key_highlights": len(key_hls),
                      "recommended_read_pages": rec_pages,
                      "title_guess": title, "limits": limits}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
