#!/usr/bin/env python3
"""inspect_case.py — Step 1 of the academic-slides skill.

Inventories the runtime case: case.json, material.json, artifacts/slide-workorder.json,
PPTX templates / starting decks, PDFs and figure images. Extracts PDF text + embedded
figures so later steps can work from real content instead of labels.

Usage:
  python3 inspect_case.py --case-dir /input --work /app/work
Outputs: work/inventory.json (+ work/material_text.md, work/figs/*) and a compact summary.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckutil  # noqa: E402

try:
    from pptx import Presentation  # noqa: E402
    from pptx.util import Emu  # noqa: E402
except ImportError:
    Presentation = None  # type: ignore[assignment]
    Emu = None  # type: ignore[assignment]


def find_first(root: Path, names, work: Path | None = None):
    for n in names:
        for cand in (root / n, root / "artifacts" / n, root / "inputs" / n,
                     root / "materials" / n, root / "data" / n):
            if cand.is_file():
                return cand
    # Safe recursive fallback for nested case layouts
    for n in names:
        for cand in sorted(root.rglob(n)):
            if cand.is_file() and (work is None or not _skip_path(cand, root, work)):
                return cand
    return None


def theme_fonts(pptx_path: Path):
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml", n)]
            if not names:
                return {}
            xml = z.read(sorted(names)[0]).decode("utf-8", "replace")
            out = {}
            for tag in ("majorFont", "minorFont"):
                m = re.search(rf'<a:{tag}>.*?<a:latin typeface="([^"]*)"', xml, re.S)
                out[tag] = m.group(1) if m else None
            return out
    except Exception:
        return {}


def describe_pptx(path: Path):
    d = {"path": str(path), "slide_size_in": None, "layouts": [], "slides": []}
    if Presentation is None:
        d["error"] = "python-pptx not installed"
        return d
    try:
        prs = Presentation(str(path))
        d["slide_size_in"] = [round(prs.slide_width / 914400, 2), round(prs.slide_height / 914400, 2)]
        tf = theme_fonts(path)
        if tf:
            d["theme_fonts"] = tf
        for i, layout in enumerate(prs.slide_layouts):
            phs = []
            for ph in layout.placeholders:
                phs.append({
                    "idx": ph.placeholder_format.idx,
                    "type": str(ph.placeholder_format.type).split(" ")[0],
                    "name": ph.name,
                })
            d["layouts"].append({"i": i, "name": layout.name, "placeholders": phs})
        for i, slide in enumerate(prs.slides):
            shapes = []
            for sh in slide.shapes:
                info = {"name": sh.name, "kind": str(sh.shape_type).split(" ")[0]}
                if sh.is_placeholder:
                    info["ph_type"] = str(sh.placeholder_format.type).split(" ")[0]
                if sh.has_text_frame and sh.text_frame.text.strip():
                    info["text"] = sh.text_frame.text.strip().replace("\n", " ⏎ ")[:80]
                shapes.append(info)
            d["slides"].append({"i": i + 1, "shapes": shapes[:14]})
    except Exception as e:
        d["error"] = str(e)
    return d


FIG_CAP_RE = re.compile(
    r"^(?:(?:Extended\s+Data|Supplementary|Supp\.?)\s+)?(?:Figure|Fig\.?|图)\s*"
    r"([SA]?\d+(?:[.\-]\d+)?[a-zA-Z]?|[IVX]{1,5})\b[\s:：.\-—]*(.*)$",
    re.I,
)
TBL_CAP_RE = re.compile(
    r"^(?:(?:Extended\s+Data|Supplementary|Supp\.?)\s+)?(?:Table|Tab\.?|表)\s*"
    r"([SA]?\d+(?:[.\-]\d+)?[a-zA-Z]?|[IVX]{1,5})\b[\s:：.\-—]*(.*)$",
    re.I,
)


def _crop_vector_figure(page, cap_bbox, fitz, fig_path: Path):
    """Render/crop the vector figure region immediately above a Figure caption.
    Column-aware for 2-column IEEE/NeurIPS/CVPR papers so left-column captions never grab right-column drawings."""
    pr = page.rect
    x0, y0, x1, y1 = cap_bbox
    cap_w = max(0.0, x1 - x0)
    is_single_col = cap_w < pr.width * 0.58
    if is_single_col:
        mid_x = pr.x0 + pr.width * 0.5
        if (x0 + x1) * 0.5 <= mid_x:
            col_x0, col_x1 = pr.x0 + 6.0, max(x1 + 24.0, mid_x + 8.0)
        else:
            col_x0, col_x1 = min(x0 - 24.0, mid_x - 8.0), pr.x1 - 6.0
    else:
        col_x0, col_x1 = pr.x0 + 8.0, pr.x1 - 8.0
    # Search vector drawings above the caption within the same column band
    top_limit = max(pr.y0 + 24.0, y0 - pr.height * 0.52)
    rects = []
    try:
        for d in page.get_drawings():
            r = d.get("rect")
            if r is None:
                continue
            if r.width < 2 and r.height < 2:
                continue
            if r.width > pr.width * 0.94 and r.height > pr.height * 0.90:
                continue  # page frame
            if r.y1 <= y0 + 4.0 and r.y0 >= top_limit - 10.0 and r.y1 > pr.y0 + 20.0:
                if is_single_col and (r.x1 < col_x0 or r.x0 > col_x1):
                    continue  # drawing belongs to the other column
                rects.append(r)
    except Exception:
        rects = []
    if rects:
        ux0 = max(col_x0, min(r.x0 for r in rects) - 14.0)
        uy0 = max(pr.y0 + 16.0, min(r.y0 for r in rects) - 14.0)
        ux1 = min(col_x1, max(r.x1 for r in rects) + 14.0)
        uy1 = min(y0 - 1.0, max(r.y1 for r in rects) + 10.0)
        if ux1 - ux0 >= 80.0 and uy1 - uy0 >= 55.0:
            clip = fitz.Rect(ux0, uy0, ux1, uy1)
        else:
            clip = None
    else:
        clip = None
    if clip is None:
        # Fallback: conservative band above the caption within its column if enough vertical space exists
        band_top = max(pr.y0 + 36.0, y0 - pr.height * 0.34)
        if y0 - band_top < 70.0:
            return None
        clip = fitz.Rect(max(pr.x0 + 16.0, col_x0), band_top, min(pr.x1 - 16.0, col_x1), y0 - 2.0)
    pix = page.get_pixmap(clip=clip, dpi=160)
    if pix.width < 60 or pix.height < 50:
        return None
    pix.save(str(fig_path))
    return pix.width, pix.height


def extract_pdf(pdf_path: Path, work: Path):
    """Extract per-page text, embedded raster images, vector figure crops (by caption),
    structured PDF tables (via PyMuPDF find_tables), and a structured Paper Map."""
    out = {"path": str(pdf_path), "pages": 0, "text_file": None, "figures": [],
           "tables_index": [], "tables_extracted": [], "anomalies": []}
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz as fitz
        doc = fitz.open(str(pdf_path))
        out["pages"] = doc.page_count
        chunks = []
        fig_index_lines = []
        tbl_index_lines = []
        figdir = work / "figs"
        for pno in range(doc.page_count):
            page = doc[pno]
            chunks.append(f"\n\n<!-- PDF page {pno + 1} -->\n" + page.get_text("text"))
            # Detect figure & table caption blocks with coordinates
            fig_caps = []
            page_tbl_caps = []
            page_has_tbl = False
            try:
                for b in page.get_text("dict").get("blocks", []):
                    if b.get("type", 0) != 0:
                        continue
                    blines = [
                        (re.sub(r"\s+", " ", "".join(s.get("text", "") for s in ln.get("spans", []))).strip(),
                         ln.get("bbox"))
                        for ln in b.get("lines", [])
                    ]
                    blines = [(t, bb) for t, bb in blines if t]
                    if not blines:
                        continue
                    # Split horizontally disjoint 2-column caption lines (e.g. left Fig. S1 + right Fig. 2 at same Y)
                    cap_groups = []
                    for lt, lbb in blines:
                        is_new_cap = bool(FIG_CAP_RE.match(lt) or TBL_CAP_RE.match(lt))
                        if cap_groups and is_new_cap:
                            prev_bb = cap_groups[-1][1]
                            if lbb and prev_bb and (lbb[0] > prev_bb[2] - 4.0 or prev_bb[0] > lbb[2] - 4.0):
                                cap_groups.append([lt, lbb])
                                continue
                        if cap_groups:
                            cap_groups[-1][0] += " " + lt
                            if lbb and cap_groups[-1][1]:
                                pb = cap_groups[-1][1]
                                cap_groups[-1][1] = (min(pb[0], lbb[0]), min(pb[1], lbb[1]),
                                                     max(pb[2], lbb[2]), max(pb[3], lbb[3]))
                        else:
                            cap_groups.append([lt, lbb or b.get("bbox")])
                    for btxt, bbox_g in cap_groups:
                        btxt = re.sub(r"\s+", " ", btxt).strip()
                        if not btxt:
                            continue
                        mf = FIG_CAP_RE.match(btxt)
                        if mf:
                            fig_caps.append((mf.group(1), btxt[:180], bbox_g))
                            fig_index_lines.append(f"- p{pno + 1}: {btxt[:140]}")
                        else:
                            mt = TBL_CAP_RE.match(btxt)
                            if mt:
                                page_has_tbl = True
                                page_tbl_caps.append(btxt[:180])
                                tbl_index_lines.append(f"- p{pno + 1}: {btxt[:140]}")
                                out["tables_index"].append({"page": pno + 1, "id": mt.group(1), "caption": btxt[:180]})
            except Exception:
                pass

            # Extract structured PDF tables via PyMuPDF find_tables() using cell-clip text
            # Strictly gated on actual Table caption blocks on this page (<0.25s on 25p PDFs!)
            # Falls back to vertical_strategy="text" for borderless LaTeX booktabs academic tables
            if len(out["tables_extracted"]) < 8 and page_has_tbl and hasattr(page, "find_tables"):
                def _extract_tbl_list(tf_obj):
                    found = []
                    for tb in (getattr(tf_obj, "tables", None) or []):
                        clean_rows = []
                        tb_rows = getattr(tb, "rows", None) or []
                        for row in tb_rows[:12]:
                            cr = []
                            for cell in (getattr(row, "cells", None) or [])[:10]:
                                if cell is None:
                                    cr.append("")
                                else:
                                    txt_c = page.get_text("text", clip=fitz.Rect(*cell))
                                    cr.append(re.sub(r"\s+", " ", txt_c).strip()[:60])
                            if any(cr):
                                clean_rows.append(cr)
                        if len(clean_rows) >= 2 and len(clean_rows[0]) >= 2:
                            # Reject prose false positives: require >=2 non-empty cells in >=2 rows
                            dense_rows = sum(1 for r in clean_rows if sum(1 for c in r if c) >= 2)
                            if dense_rows >= 2:
                                found.append(clean_rows)
                    return found

                try:
                    found_tbls = _extract_tbl_list(page.find_tables())
                    if not found_tbls:
                        try:
                            found_tbls = _extract_tbl_list(
                                page.find_tables(vertical_strategy="text", horizontal_strategy="lines")
                            )
                        except Exception:
                            found_tbls = []
                    for ti_p, clean_rows in enumerate(found_tbls):
                        t_cap = page_tbl_caps[ti_p] if ti_p < len(page_tbl_caps) else (page_tbl_caps[0] if page_tbl_caps else "")
                        out["tables_extracted"].append({"page": pno + 1, "caption": t_cap, "rows": clean_rows})
                        if len(out["tables_extracted"]) >= 8:
                            break
                except Exception:
                    pass

            # Extract embedded raster images
            page_rasters = []
            for i, info in enumerate(page.get_images(full=True)):
                xref = info[0]
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    if pix.width < 40 or pix.height < 40:
                        continue
                    figdir.mkdir(parents=True, exist_ok=True)
                    fp = figdir / f"pdf_p{pno + 1}_img{i}.png"
                    pix.save(str(fp))
                    cap_txt = fig_caps[len(page_rasters)][1] if len(page_rasters) < len(fig_caps) else ""
                    rec = {"file": str(fp), "page": pno + 1, "w": pix.width, "h": pix.height,
                           "kind": "raster", "caption": cap_txt}
                    out["figures"].append(rec)
                    page_rasters.append(rec)
                except Exception:
                    out["anomalies"].append(f"p{pno + 1}: image xref {xref} extract failed")

            # Vector-figure fallback crop: if page has Figure captions without matching raster images
            if fig_caps and len(page_rasters) < len(fig_caps):
                for ci, (fid, cap_txt, bbox) in enumerate(fig_caps[len(page_rasters):], 1):
                    if not bbox:
                        continue
                    try:
                        figdir.mkdir(parents=True, exist_ok=True)
                        fp = figdir / f"pdf_p{pno + 1}_fig{fid}_vec{ci}.png"
                        res = _crop_vector_figure(page, bbox, fitz, fp)
                        if res:
                            out["figures"].append({"file": str(fp), "page": pno + 1,
                                                   "w": res[0], "h": res[1],
                                                   "kind": "vector_crop", "figure_id": fid,
                                                   "caption": cap_txt})
                    except Exception:
                        pass

        text = "".join(chunks)
        # paper map: per-page heading candidates + figures + tables for cheap navigation
        try:
            heads = []
            for pno in range(doc.page_count):
                page = doc[pno]
                sizes = []
                lines = {}
                for b in page.get_text("dict")["blocks"]:
                    for ln in b.get("lines", []):
                        txt = "".join(s["text"] for s in ln.get("spans", [])).strip()
                        if not txt:
                            continue
                        mx = max((s.get("size", 0) for s in ln.get("spans", [])), default=0)
                        sizes.append(mx)
                        lines[(round(ln["bbox"][1], 1), txt)] = mx
                if not sizes:
                    continue
                sizes.sort()
                med = sizes[len(sizes) // 2] or 10.0
                cand = [t for (_, t), s in sorted(lines.items())
                        if s >= med * 1.18 and 2 < len(t) < 60]
                for t in cand[:5]:
                    heads.append(f"- p{pno + 1}: {t}")
            header_parts = []
            if heads:
                header_parts.append("<!-- 论文导航地图 / paper map (heading candidates by page) -->\n"
                                    + "\n".join(heads))
            if fig_index_lines:
                header_parts.append("<!-- 图表索引 / figures index -->\n" + "\n".join(fig_index_lines[:24]))
            if tbl_index_lines:
                header_parts.append("<!-- 数据表索引 / tables index -->\n" + "\n".join(tbl_index_lines[:24]))
            if header_parts:
                text = "\n" + "\n\n".join(header_parts) + "\n" + text
        except Exception:
            pass
        tfp = work / "material_text.md"
        tfp.write_text(text, encoding="utf-8")
        out["text_file"] = str(tfp)
        out["text_chars"] = len(text)
        doc.close()
    except Exception as e:
        # Fallback chain when PyMuPDF is missing or fails: pdfplumber -> pdftotext
        fb_chunks = []
        try:
            import pdfplumber
            with pdfplumber.open(str(pdf_path)) as pdoc:
                out["pages"] = len(pdoc.pages)
                for pno, pg in enumerate(pdoc.pages, 1):
                    fb_chunks.append(f"\n\n<!-- PDF page {pno} -->\n" + (pg.extract_text() or ""))
        except Exception:
            try:
                import shutil as _sh, subprocess as _sp
                pt = _sh.which("pdftotext")
                if pt:
                    r = _sp.run([pt, str(pdf_path), "-"], capture_output=True, text=True, timeout=60)
                    if r.returncode == 0 and r.stdout.strip():
                        fb_chunks.append("\n\n<!-- PDF page 1 -->\n" + r.stdout)
            except Exception:
                pass
        if fb_chunks:
            text = "".join(fb_chunks)
            tfp = work / "material_text.md"
            tfp.write_text(text, encoding="utf-8")
            out["text_file"] = str(tfp)
            out["text_chars"] = len(text)
            out["anomalies"].append(f"PyMuPDF unavailable ({e}); used text fallback")
        else:
            out["anomalies"].append(f"PDF extract failed: {e}")
    return out


def normalize_workorder(wo: dict):
    """Best-effort normalization of common workorder fields (schema-adaptive, recursive)."""
    out = {}
    flat = {}
    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k not in flat or flat.get(k) in (None, ""):
                    flat.setdefault(k, v)
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(wo)
    def g(*keys, default=None):
        for k in keys:
            v = flat.get(k)
            if v not in (None, "") and isinstance(v, (str, int, float, bool)):
                return v
        return default
    mode = g("mode", "task_type", "operation", "deck_mode", default="")
    # Normalize separators so word boundaries work on snake_case / kebab-case / paths
    m = re.sub(r"[-_/]+", " ", str(mode).lower()).strip()
    # Strip benign non-edit phrases before keyword matching (e.g. "generate_editable_slides",
    # "editable_pptx", "template_adaptation", "adapt_paper_to_slides", "生成可编辑PPT")
    m = re.sub(r"\beditable\b", " ", m)
    m = re.sub(
        r"\b(?:(?:template|paper|manuscript|article)\s+adapt\w*"
        r"|adapt\w*(?:\s+(?:to|for|from|with|a|the))*\s+(?:template|paper|manuscript|article))\b",
        " ", m,
    )
    m = re.sub(r"(?:可编辑|模板适配|适配模板)", " ", m)
    is_create = (
        bool(re.search(r"\b(?:create|creating|created|creation|generate|generating|generated|generation|build|building|built|new)\b", m))
        or any(k in m for k in ("生成", "创建", "制作", "新建"))
    )
    is_edit = (
        bool(re.search(r"\b(?:edit|edits|editing|edited|revise|revising|revised|revision|modify|modifying|modified|modification|adapt|adapting|adapted|adaptation|update|updating|updated)\b", m))
        or any(k in m for k in ("更新", "修改", "编辑", "修订"))
    )
    if is_create and not is_edit:
        out["mode"] = "create"
    elif is_edit and not is_create:
        out["mode"] = "edit"
    else:
        out["mode"] = "auto"
    for keys, dst in (
        (("slide_count", "total_slides", "slides", "num_slides", "page_count", "页数"), "slide_count"),
        (("max_slides", "slide_limit", "max_pages", "页数上限"), "max_slides"),
        (("min_slides", "min_pages", "页数下限"), "min_slides"),
        (("audience", "受众", "target_audience"), "audience"),
        (("language", "语言"), "language"),
        (("duration", "time_limit", "时长"), "duration"),
        (("title", "deck_title", "题目"), "deck_title"),
        (("source_preservation", "fidelity_rule", "preservation"), "source_preservation"),
        (("template", "target_template", "template_path", "模板"), "template"),
        (("starting_deck", "start_deck", "input_deck", "base_deck", "existing_deck"), "starting_deck"),
    ):
        v = g(*keys)
        if v is not None:
            out[dst] = v
    # Parse range-style slide_count (e.g. "10-12", "8–12", "10~12") into min/max
    sc_val = out.get("slide_count")
    if isinstance(sc_val, str):
        rm = re.search(r"(\d+)\s*[-–—~至到]\s*(\d+)", sc_val)
        if rm:
            lo, hi = int(rm.group(1)), int(rm.group(2))
            if lo <= hi:
                out.setdefault("min_slides", lo)
                out.setdefault("max_slides", hi)
                out.pop("slide_count", None)
    # preserve list/dict edit instructions if present
    for k in ("edit_instructions", "edits", "instructions", "requirements", "changes", "tasks"):
        v = flat.get(k)
        if isinstance(v, (list, dict)) and v:
            out["edit_instructions"] = v
            break
    return out


def dump_start_deck(start_path: Path, out_md: Path) -> str | None:
    """Write a complete human/LLM-readable slide-by-slide dump of starting-deck.pptx."""
    if Presentation is None:
        return None
    try:
        prs = Presentation(str(start_path))
        n_slides = len(prs.slides)
        note_cap = 180 if n_slides > 10 else 420
        lines = [f"# Starting Deck Dump: `{start_path.name}` ({n_slides} slides)", ""]
        for i, slide in enumerate(prs.slides, 1):
            snap = deckutil.extract_slide_snapshot(slide, prs)
            lines.append(f"## Slide {i} [{snap['layout']}]")
            lines.append(f"- **Title**: {snap['title'] or '(none)'}")
            if snap["bullets"]:
                lines.append("- **Body**:")
                for b in snap["bullets"]:
                    lines.append(f"  - {b}")
            if snap["tables"]:
                lines.append(f"- **Tables**: {len(snap['tables'])}")
                for ti_s, tbl_s in enumerate(snap["tables"][:4], 1):
                    lines.append(f"  - Table {ti_s}: {tbl_s.get('preview', [])[:14]}")
            if snap["charts"]:
                lines.append(f"- **Charts**: {snap['charts']}")
            if snap["images"]:
                lines.append(f"- **Images**: {snap['images']}")
            if snap["notes"]:
                nt = snap["notes"]
                lines.append(f"- **Notes**: {nt[:note_cap]}" + (" ..." if len(nt) > note_cap else ""))
            lines.append("")
        out_md.write_text("\n".join(lines), encoding="utf-8")
        return str(out_md)
    except Exception:
        return None


def extract_pdfs(pdf_paths, work: Path):
    """Extract text + embedded figures + tables from one or more PDFs into work/material_text.md."""
    if len(pdf_paths) == 1:
        return extract_pdf(pdf_paths[0], work)
    merged = {"paths": [str(p) for p in pdf_paths], "pages": 0, "text_file": None,
              "figures": [], "tables_index": [], "tables_extracted": [], "anomalies": []}
    all_chunks = []
    for idx, p in enumerate(pdf_paths, 1):
        sub_work = work / f"_pdf_{idx}"
        sub_work.mkdir(parents=True, exist_ok=True)
        r = extract_pdf(p, sub_work)
        merged["pages"] += r.get("pages") or 0
        merged["anomalies"].extend(r.get("anomalies") or [])
        merged["tables_index"].extend(r.get("tables_index") or [])
        merged["tables_extracted"].extend(r.get("tables_extracted") or [])
        # move extracted figs into unified work/figs with prefix
        figdir = work / "figs"
        figdir.mkdir(parents=True, exist_ok=True)
        for fg in r.get("figures") or []:
            src = Path(fg["file"])
            dst = figdir / f"pdf{idx}_{src.name}"
            try:
                src.replace(dst)
                fg["file"] = str(dst)
                fg["pdf"] = p.name
                merged["figures"].append(fg)
            except Exception:
                pass
        if r.get("text_file") and Path(r["text_file"]).is_file():
            all_chunks.append(f"\n\n<!-- SOURCE PDF {idx}: {p.name} -->\n"
                              + Path(r["text_file"]).read_text(encoding="utf-8", errors="replace"))
        try:
            import shutil
            shutil.rmtree(sub_work, ignore_errors=True)
        except Exception:
            pass
    text = "".join(all_chunks)
    tfp = work / "material_text.md"
    tfp.write_text(text, encoding="utf-8")
    merged["text_file"] = str(tfp)
    merged["text_chars"] = len(text)
    return merged


def _extract_xlsx_md(xlsx_path: Path, max_sheets: int = 6, max_rows: int = 30, max_cols: int = 12) -> str:
    """Zero-dependency stdlib XLSX sheet-to-Markdown extractor for benchmark/data tables."""
    import xml.etree.ElementTree as ET
    out_lines: list[str] = []
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            shared: list[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                root_s = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root_s.iter():
                    if si.tag.endswith("}si") or si.tag == "si":
                        txt = "".join(t.text or "" for t in si.iter() if t.tag.endswith("}t") or t.tag == "t")
                        shared.append(txt.strip())
            sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))[:max_sheets]
            for sname in sheets:
                root = ET.fromstring(z.read(sname))
                rows_out: list[list[str]] = []
                for row in root.iter():
                    if not (row.tag.endswith("}row") or row.tag == "row"):
                        continue
                    vals: list[str] = []
                    for c in list(row)[:max_cols]:
                        if not (c.tag.endswith("}c") or c.tag == "c"):
                            continue
                        ct = c.attrib.get("t", "")
                        val = ""
                        if ct == "inlineStr":
                            val = "".join(t.text or "" for t in c.iter() if t.tag.endswith("}t") or t.tag == "t").strip()
                        else:
                            v_el = next((ch for ch in c if ch.tag.endswith("}v") or ch.tag == "v"), None)
                            if v_el is not None and v_el.text is not None:
                                raw_v = v_el.text.strip()
                                if ct == "s" and raw_v.isdigit() and int(raw_v) < len(shared):
                                    val = shared[int(raw_v)]
                                else:
                                    val = raw_v
                        vals.append(re.sub(r"\s+", " ", val).replace("|", "/"))
                    if any(vals):
                        rows_out.append(vals)
                    if len(rows_out) >= max_rows:
                        break
                if rows_out:
                    out_lines.append(f"### {xlsx_path.name} ({Path(sname).stem})")
                    ncols = max(len(r) for r in rows_out)
                    for ri, r in enumerate(rows_out):
                        r_pad = r + [""] * (ncols - len(r))
                        out_lines.append("| " + " | ".join(r_pad) + " |")
                        if ri == 0:
                            out_lines.append("| " + " | ".join("---" for _ in range(ncols)) + " |")
    except Exception:
        return ""
    return "\n".join(out_lines)


def _extract_docx_md(docx_path: Path, max_paras: int = 400, max_tbl_rows: int = 30) -> str:
    """Zero-dependency stdlib DOCX text + Markdown table extractor (works without python-docx)."""
    import xml.etree.ElementTree as ET
    out_lines: list[str] = []
    try:
        with zipfile.ZipFile(docx_path) as z:
            if "word/document.xml" not in z.namelist():
                return ""
            root = ET.fromstring(z.read("word/document.xml"))
            body = next((el for el in root.iter() if el.tag.endswith("}body") or el.tag == "body"), root)
            n_p = 0
            for child in list(body):
                tag = child.tag.rsplit("}", 1)[-1]
                if tag == "p" and n_p < max_paras:
                    txt = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t") or t.tag == "t").strip()
                    if txt:
                        out_lines.append(txt)
                        n_p += 1
                elif tag == "tbl":
                    rows_out: list[list[str]] = []
                    for tr in child:
                        if tr.tag.rsplit("}", 1)[-1] != "tr":
                            continue
                        vals: list[str] = []
                        for tc in tr:
                            if tc.tag.rsplit("}", 1)[-1] != "tc":
                                continue
                            ct = " ".join(
                                "".join(t.text or "" for t in p.iter() if t.tag.endswith("}t") or t.tag == "t").strip()
                                for p in tc if p.tag.rsplit("}", 1)[-1] == "p"
                            ).strip()
                            vals.append(re.sub(r"\s+", " ", ct).replace("|", "/"))
                        if any(vals):
                            rows_out.append(vals)
                        if len(rows_out) >= max_tbl_rows:
                            break
                    if rows_out:
                        ncols = max(len(r) for r in rows_out)
                        for ri, r in enumerate(rows_out):
                            r_pad = r + [""] * (ncols - len(r))
                            out_lines.append("| " + " | ".join(r_pad) + " |")
                            if ri == 0:
                                out_lines.append("| " + " | ".join("---" for _ in range(ncols)) + " |")
                        out_lines.append("")
    except Exception:
        return ""
    return "\n".join(out_lines)


def _skip_path(p: Path, case_dir: Path, work: Path) -> bool:
    """Exclude work/ and output/ subdirectories when scanning case_dir."""
    try:
        pr = p.resolve()
        wr = work.resolve()
        if pr == wr or wr in pr.parents:
            return True
    except Exception:
        pass
    try:
        rel_parts = [x.lower() for x in p.relative_to(case_dir).parts[:-1]]
    except Exception:
        rel_parts = [x.lower() for x in p.parts[:-1]]
    return any(part in ("output", "outputs", "previews", "__pycache__") for part in rel_parts)


def write_material_digest(inv: dict, work: Path) -> str | None:
    """Write a high-fidelity deterministic digest (<=22KB) of the case materials:
    paper map, figure/table index + verbatim captions, extracted PDF & CSV tables,
    full material.json (up to 14KB verbatim), opening/abstract verbatim, and
    sentence-split page-balanced narrative/quantitative lines across ALL pages."""
    lines = [f"# Material Digest: `{inv.get('case_id', 'UNKNOWN')}`", ""]
    pdf_ex = inv.get("pdf_extract") or {}
    # 1. Paper map & figure/table index from material_text.md header comments
    mt_path = pdf_ex.get("text_file")
    mt_raw = ""
    if mt_path and Path(mt_path).is_file():
        mt_raw = Path(mt_path).read_text(encoding="utf-8", errors="replace")
        m_p1 = re.search(r"<!--\s*PDF page 1\s*-->|---\s*page 1\s*---", mt_raw)
        hdr_end = m_p1.start() if m_p1 else -1
        if hdr_end > 0:
            hdr = mt_raw[:hdr_end].strip()
            if hdr:
                lines.append("## Paper Map & Indices")
                lines.append(hdr[:2800])
                lines.append("")

    # 2. Available real figures (PDF extracted + shipped)
    figs = []
    for fg in (pdf_ex.get("figures") or [])[:16]:
        figs.append(f"- `{fg.get('file')}` (p.{fg.get('page')}, {fg.get('kind')}): {(fg.get('caption') or '')[:160]}")
    for ip in (inv.get("image_files") or [])[:16]:
        figs.append(f"- `{ip}` (case image)")
    if figs:
        lines.append("## Available Real Figures")
        lines.extend(figs)
        lines.append("")

    # 2b. Extracted structured PDF tables (exact cell values + matched caption for zero-typo slides)
    ext_tables = pdf_ex.get("tables_extracted") or []
    if ext_tables:
        lines.append("## Extracted PDF Tables")
        for ti, tb in enumerate(ext_tables[:8], 1):
            rows = tb.get("rows") or []
            if not rows:
                continue
            t_cap = str(tb.get("caption") or "").strip()
            hdr_lbl = f"### PDF Table {ti} (p.{tb.get('page')})" + (f" — {t_cap[:120]}" if t_cap else "")
            lines.append(hdr_lbl)
            lines.append("| " + " | ".join(rows[0]) + " |")
            lines.append("|" + "|".join(["---"] * len(rows[0])) + "|")
            for r in rows[1:13]:
                padded = r + [""] * max(0, len(rows[0]) - len(r))
                lines.append("| " + " | ".join(padded[:len(rows[0])]) + " |")
            lines.append("")

    # 2c. Supplementary CSV/TSV data tables (exact numeric tables visible in Step 1)
    csv_files = [Path(p) for p in (inv.get("extra_text_files") or [])
                 if Path(p).suffix.lower() in (".csv", ".tsv") and Path(p).is_file()][:4]
    if csv_files:
        import csv as _csv
        lines.append("## Supplementary CSV/TSV Tables")
        for cf in csv_files:
            try:
                delim = "\t" if cf.suffix.lower() == ".tsv" else ","
                with cf.open("r", encoding="utf-8", errors="replace", newline="") as fh:
                    rows = []
                    for r in _csv.reader(fh, delimiter=delim):
                        cr = [c.strip()[:56] for c in r[:10]]
                        if any(cr):
                            rows.append(cr)
                        if len(rows) >= 13:
                            break
                if len(rows) >= 2 and len(rows[0]) >= 2:
                    lines.append(f"### `{cf.name}`")
                    lines.append("| " + " | ".join(rows[0]) + " |")
                    lines.append("|" + "|".join(["---"] * len(rows[0])) + "|")
                    for r in rows[1:]:
                        padded = r + [""] * max(0, len(rows[0]) - len(r))
                        lines.append("| " + " | ".join(padded[:len(rows[0])]) + " |")
                    lines.append("")
            except Exception:
                continue

    # 2d. Supplementary Markdown/Text/JSON files (non-CSV/TSV) so non-PDF cases are complete in Step 1
    other_txts = [Path(p) for p in (inv.get("extra_text_files") or [])
                  if Path(p).suffix.lower() not in (".csv", ".tsv") and Path(p).is_file()][:4]
    if other_txts:
        lines.append("## Supplementary Text/Data Files")
        budget = 3600
        for tf_p in other_txts:
            if budget <= 120:
                break
            try:
                raw_t = tf_p.read_text(encoding="utf-8", errors="replace").strip()
                if raw_t:
                    chunk = raw_t[:min(budget, 1800)]
                    lines.append(f"### `{tf_p.name}`")
                    lines.append(chunk + (" ..." if len(raw_t) > len(chunk) else ""))
                    lines.append("")
                    budget -= len(chunk)
            except Exception:
                continue

    # 3. Full structured material.json (dynamic anti-starvation budget: <=4.8KB verbatim when PDF
    #    is also present so PDF Abstract/Captions/Key Sentences NEVER get clipped out of Step 1 stdout;
    #    <=9.5KB verbatim when no PDF is present)
    mat = inv.get("material") or {}
    mat_chars = 0
    mat_truncated = False
    mat_verbatim_cap = 4800 if len(mt_raw) > 2000 else 9500
    if mat:
        lines.append("## material.json Full Content")
        mat_str = json.dumps(mat, ensure_ascii=False, indent=1)
        mat_chars = len(mat_str)
        if mat_chars <= mat_verbatim_cap:
            lines.append("```json\n" + mat_str + "\n```")
        else:
            mat_truncated = True
            per_hi = 950 if len(mt_raw) > 2000 else 1600
            per_lo = 360 if len(mt_raw) > 2000 else 600
            for k, v in list(mat.items())[:24]:
                vs = json.dumps(v, ensure_ascii=False)
                cap_k = per_hi if k.lower() in (
                    "tables", "figures", "key_numbers", "results", "claims",
                    "metrics", "limitations", "sections", "abstract", "method", "methods"
                ) else per_lo
                lines.append(f"- **{k}**: {vs[:cap_k]}" + (" ..." if len(vs) > cap_k else ""))
        lines.append("")

    # 4. Opening/Abstract + Verbatim Captions + Sentence-Split Page-Balanced Key Lines
    pdf_text_chars = len(mt_raw)
    if mt_raw:
        m_p1 = re.search(r"<!--\s*PDF page 1\s*-->|---\s*page 1\s*---", mt_raw)
        body = mt_raw[m_p1.start():] if m_p1 else mt_raw
        # 4a. Opening / Abstract verbatim (de-hyphenated across LaTeX line wraps, never drops framing)
        m_p2 = re.search(r"<!--\s*PDF page 2\s*-->|---\s*page 2\s*---", body)
        p1_txt = (body[:m_p2.start()] if m_p2 else body[:1800]).strip()
        p1_clean = re.sub(r"<!--.*?-->", "", p1_txt)
        p1_clean = re.sub(r"([A-Za-z]{2,})-\s*\n\s*([a-z]{2,})", r"\1\2", p1_clean).strip()
        if p1_clean:
            lines.append("## Opening & Abstract (p.1 verbatim)")
            lines.append(p1_clean[:1400] + (" ..." if len(p1_clean) > 1400 else ""))
            lines.append("")

        q_re = re.compile(
            r"(\d+\.\d+|\d{2,}\s*(%|ms|GB|MB|x|×|k\b|M\b|B\b|万|亿)|Table\s*[IVX\d]|Figure\s*[SA\d]|Fig\.\s*[SA\d]|"
            r"表\s*\d|图\s*\d|\b(outperform|achiev|reduc|improv|limitation|conclud|propos|ablation|"
            r"method|architect|dataset|benchmark|baseline|future|trade-?off|bottleneck|evaluat|significan)\w*|"
            r"(提升|降低|优于|局限|不足|未来|结论|消融|提出|架构|方法|数据集|基线|权衡|瓶颈|实验))",
            re.I,
        )
        # De-hyphenate LaTeX line-end wraps and merge wrapped 2-column prose lines per page
        body_clean = re.sub(r"([A-Za-z]{2,})-\s*\n\s*([a-z]{2,})", r"\1\2", body)
        cur_page = "p.1"
        by_page: dict[str, list[str]] = {}
        page_order: list[str] = []
        captions_found: list[str] = []
        seen = set()
        seen_cap = set()
        raw_page_lines: list[tuple[str, str]] = []
        for ln in body_clean.splitlines():
            s = ln.strip()
            if not s:
                raw_page_lines.append((cur_page, ""))
                continue
            pm = re.match(r"^(?:<!--\s*PDF page\s*(\d+)\s*-->|---\s*page\s*(\d+)\s*---)$", s)
            if pm:
                cur_page = f"p.{pm.group(1) or pm.group(2)}"
                if cur_page not in by_page:
                    by_page[cur_page] = []
                    page_order.append(cur_page)
                raw_page_lines.append((cur_page, ""))
                continue
            if s.startswith("<!--"):
                continue
            raw_page_lines.append((cur_page, s))
        # Re-join wrapped prose lines within the same page/paragraph (preserving captions & headings)
        merged_blocks: list[tuple[str, str]] = []
        buf_pg, buf_txt = "p.1", ""
        for pg, s in raw_page_lines:
            if not s or pg != buf_pg or FIG_CAP_RE.match(s) or TBL_CAP_RE.match(s):
                if buf_txt:
                    merged_blocks.append((buf_pg, buf_txt))
                    buf_txt = ""
                if s:
                    buf_pg, buf_txt = pg, s
                else:
                    buf_pg = pg
            elif (buf_txt and not re.search(r"[.!?。！？:：]$", buf_txt)
                  and len(buf_txt) < 520 and not re.match(r"^(?:\d+\.?\s+[A-Z]|[IVX]+\.\s+)", s)):
                joiner = "" if (re.search(r"[\u4e00-\u9fff]$", buf_txt) and re.match(r"^[\u4e00-\u9fff]", s)) else " "
                buf_txt = buf_txt + joiner + s
            else:
                if buf_txt:
                    merged_blocks.append((buf_pg, buf_txt))
                buf_pg, buf_txt = pg, s
        if buf_txt:
            merged_blocks.append((buf_pg, buf_txt))

        for cur_page, s in merged_blocks:
            # Capture verbatim figure/table captions across all pages
            if (FIG_CAP_RE.match(s) or TBL_CAP_RE.match(s)) and s not in seen_cap and len(captions_found) < 24:
                seen_cap.add(s)
                captions_found.append(f"- [{cur_page}] {s[:240]}")
            # Split long paragraphs into verbatim sentences so prose paragraphs (>220 chars) are NEVER dropped
            segs = [s] if len(s) <= 240 else [
                seg.strip() for seg in re.split(r"(?<=[.!?。！？])\s+", s) if len(seg.strip()) >= 18
            ]
            for seg in segs:
                if len(seg) < 18:
                    continue
                seg_clip = seg[:240]
                if q_re.search(seg_clip) and seg_clip not in seen:
                    seen.add(seg_clip)
                    if cur_page not in by_page:
                        by_page[cur_page] = []
                        page_order.append(cur_page)
                    if len(by_page[cur_page]) < 16:
                        by_page[cur_page].append(seg_clip)
        if captions_found:
            lines.append("## Verbatim Figure & Table Captions")
            lines.extend(captions_found)
            lines.append("")
        # Round-robin across pages so late-paper Results/Ablations/Limitations/Conclusions are never dropped
        q_lines = []
        for idx_slot in range(16):
            for pg in page_order:
                bucket = by_page.get(pg) or []
                if idx_slot < len(bucket):
                    q_lines.append(f"- [{pg}] {bucket[idx_slot]}")
                    if len(q_lines) >= 90:
                        break
            if len(q_lines) >= 90:
                break
        if q_lines:
            lines.append("## Verbatim Key Sentences Across All Pages (exact quotes)")
            lines.extend(q_lines)
            lines.append("")

    out_md = work / "material_digest.md"
    full_txt = "\n".join(lines)
    txt = full_txt[:16000]
    out_md.write_text(txt, encoding="utf-8")
    inv["completeness"] = {
        "material_json_chars": mat_chars,
        "material_json_truncated": mat_truncated,
        "pdf_text_chars": pdf_text_chars,
        "pdf_beyond_digest": bool(pdf_text_chars > 14000),
        "digest_chars": len(txt),
        "digest_clipped": len(full_txt) > 16000,
    }
    return str(out_md)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", default="/input")
    ap.add_argument("--work", default="/app/work")
    args = ap.parse_args()
    case_dir = Path(args.case_dir)
    work = Path(args.work)
    work.mkdir(parents=True, exist_ok=True)

    inv = {"case_dir": str(case_dir), "work": str(work), "anomalies": []}

    case_json = find_first(case_dir, ("case.json",), work)
    if case_json:
        try:
            inv["case"] = json.loads(case_json.read_text(encoding="utf-8"))
        except Exception as e:
            inv["anomalies"].append(f"case.json unreadable: {e}")
            inv["case"] = {"raw_path": str(case_json)}
    else:
        inv["anomalies"].append("case.json not found")
        inv["case"] = {}

    inv["case_id"] = inv["case"].get("case_id") or inv["case"].get("id") or "UNKNOWN"
    oc_raw = inv["case"].get("output_contract")
    if isinstance(oc_raw, list):
        inv["output_contract"] = [str(x) for x in oc_raw if isinstance(x, str)]
    elif isinstance(oc_raw, dict):
        req = oc_raw.get("required_fields") or oc_raw.get("fields")
        if isinstance(req, list):
            inv["output_contract"] = [str(x) for x in req if isinstance(x, str)]
        else:
            inv["output_contract"] = [str(k) for k in oc_raw.keys()]

    wo_path = find_first(case_dir, ("slide-workorder.json", "workorder.json"), work)
    if wo_path:
        try:
            inv["workorder_raw"] = json.loads(wo_path.read_text(encoding="utf-8"))
            wo_raw = inv["workorder_raw"]
            if isinstance(wo_raw, list) and wo_raw and isinstance(wo_raw[0], dict):
                wo_raw = wo_raw[0]  # array-form workorder: use first object
            inv["workorder"] = normalize_workorder(wo_raw if isinstance(wo_raw, dict) else {})
        except Exception as e:
            inv["anomalies"].append(f"workorder unreadable: {e}")
    else:
        # Fallback: extract workorder-like constraints embedded directly in case.json
        case_wo = normalize_workorder(inv["case"]) if isinstance(inv.get("case"), dict) else {}
        if any(k in case_wo for k in ("slide_count", "max_slides", "min_slides", "audience",
                                      "language", "edit_instructions", "template", "starting_deck")):
            inv["workorder"] = case_wo
        else:
            inv["anomalies"].append("slide-workorder.json not found — infer constraints from case.json/task.md")

    task_md = find_first(case_dir, ("task.md",), work)
    if task_md:
        inv["task_md"] = str(task_md)

    # ---- PPTX inputs (explicit workorder paths first, then filename heuristics)
    pptx_files = sorted(p for p in case_dir.rglob("*.pptx")
                        if p.is_file() and not _skip_path(p, case_dir, work))
    wo_dict = inv.get("workorder") or {}

    def _resolve_pptx_hint(hint_val):
        if not hint_val:
            return None
        hp = Path(str(hint_val))
        for cand in (hp, case_dir / hp, case_dir / "inputs" / hp, case_dir / "artifacts" / hp):
            if cand.is_file() and cand.suffix.lower() == ".pptx":
                return cand
        for p in pptx_files:
            if p.name.lower() == hp.name.lower():
                return p
        return None

    template = _resolve_pptx_hint(wo_dict.get("template"))
    start_deck = _resolve_pptx_hint(wo_dict.get("starting_deck"))
    for p in pptx_files:
        n = p.name.lower()
        if template is None and ("template" in n or "模板" in n):
            template = p
        elif start_deck is None and any(k in n for k in (
            "start", "deck", "draft", "old", "base", "existing", "original", "initial", "v1"
        )) and (template is None or p.resolve() != template.resolve()):
            start_deck = p
    wo_mode = wo_dict.get("mode", "auto")
    if start_deck is None and template is not None and len(pptx_files) >= 2:
        others = [p for p in pptx_files if p.resolve() != template.resolve()]
        if others:
            start_deck = others[0]
    if start_deck is None and wo_mode == "edit" and pptx_files:
        non_tmpl = [p for p in pptx_files if template is None or p.resolve() != template.resolve()]
        # Never coerce a lone identified template.pptx into start_deck when no independent deck exists
        if non_tmpl:
            start_deck = non_tmpl[0]

    inv["template"] = describe_pptx(template) if template else None
    inv["start_deck"] = describe_pptx(start_deck) if start_deck else None
    inv["pptx_files"] = [str(p) for p in pptx_files]
    if template is None and pptx_files:
        if start_deck is not None and wo_mode == "edit":
            # In edit mode without a separate template, starting deck IS the template (no anomaly)
            inv["template"] = inv["start_deck"]
        else:
            by_size = sorted(pptx_files, key=lambda p: p.stat().st_size, reverse=True)
            inv["template"] = describe_pptx(by_size[0])
            if len(by_size) > 1 and inv.get("start_deck") is None:
                inv["start_deck"] = describe_pptx(by_size[1])
            inv["anomalies"].append("template not identified by name — used size heuristic")
    if start_deck is not None:
        inv["start_deck_dump"] = dump_start_deck(start_deck, work / "start_deck_dump.md")

    # ---- material
    mat_json_path = find_first(case_dir, ("material.json", "materials.json", "content.json"), work)
    mat_text_path = find_first(case_dir, ("paper.md", "material.md", "source.md"), work)
    inv["material_json_path"] = str(mat_json_path) if mat_json_path else None
    if mat_json_path:
        try:
            inv["material"] = json.loads(mat_json_path.read_text(encoding="utf-8"))
        except Exception as e:
            inv["anomalies"].append(f"material.json unreadable: {e}")
            inv["material"] = {}
    else:
        inv["material"] = {}

    pdfs = sorted(p for p in case_dir.rglob("*.pdf")
                  if p.is_file() and not _skip_path(p, case_dir, work))
    inv["pdf_inputs"] = []
    inv["pdf_extract"] = None
    for p in pdfs:
        inv["pdf_inputs"].append(str(p))
    if pdfs:
        inv["pdf_extract"] = extract_pdfs(pdfs, work)

    text_exts = (".txt", ".md")
    primary_docx: Path | None = None
    if not inv.get("pdf_extract"):
        docx_p = next((p for p in sorted(case_dir.rglob("*.docx"))
                       if p.is_file() and not _skip_path(p, case_dir, work)), None)
        txt_p = mat_text_path or next((p for p in sorted(case_dir.rglob("*"))
                      if p.is_file() and p.suffix.lower() in text_exts
                      and p.name not in ("task.md",) and not _skip_path(p, case_dir, work)), None)
        src = docx_p or txt_p
        if src is not None:
            try:
                if docx_p is not None:
                    primary_docx = docx_p
                    text = _extract_docx_md(docx_p)
                else:
                    text = src.read_text(encoding="utf-8", errors="replace")
                tfp = work / "material_text.md"
                tfp.write_text(text, encoding="utf-8")
                inv["pdf_extract"] = {"text_file": str(tfp), "pages": None,
                                       "figures": [], "anomalies": [],
                                       "source": str(src), "text_chars": len(text)}
            except Exception as e:
                inv["anomalies"].append(f"material text extract failed: {e}")

    # ---- supplementary text/data files (CSV/TSV/MD/TXT/TEX/BIB/extra JSON/XLSX/DOCX alongside PDFs)
    supp_exts = (".csv", ".tsv", ".md", ".txt", ".tex", ".bib", ".json")
    skip_names = {"case.json", "task.md", "slide-workorder.json", "workorder.json",
                  "material.json", "materials.json", "content.json"}
    inv["extra_text_files"] = [
        str(p) for p in sorted(case_dir.rglob("*"))
        if p.is_file() and p.suffix.lower() in supp_exts
        and p.name not in skip_names and not _skip_path(p, case_dir, work)
        and p.stat().st_size <= 2_000_000
    ][:20]
    xlsx_files = [
        p for p in sorted(case_dir.rglob("*.xlsx"))
        if p.is_file() and not _skip_path(p, case_dir, work) and p.stat().st_size <= 5_000_000
    ][:4]
    if xlsx_files:
        x_chunks = [md for xp in xlsx_files if (md := _extract_xlsx_md(xp))]
        if x_chunks:
            x_md_path = work / "xlsx_tables.md"
            x_md_path.write_text("\n\n".join(x_chunks), encoding="utf-8")
            inv["extra_text_files"].append(str(x_md_path))
    supp_docx = [
        p for p in sorted(case_dir.rglob("*.docx"))
        if p.is_file() and not _skip_path(p, case_dir, work) and p.stat().st_size <= 5_000_000
        and (primary_docx is None or p.resolve() != primary_docx.resolve())
    ][:3]
    if supp_docx:
        d_chunks = [f"### {dp.name}\n" + md for dp in supp_docx if (md := _extract_docx_md(dp))]
        if d_chunks:
            d_md_path = work / "docx_text.md"
            d_md_path.write_text("\n\n".join(d_chunks), encoding="utf-8")
            inv["extra_text_files"].append(str(d_md_path))

    # ---- figure images shipped with the case
    img_exts = (".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff", ".bmp", ".svg", ".emf")
    inv["image_files"] = [str(p) for p in sorted(case_dir.rglob("*"))
                          if p.is_file() and p.suffix.lower() in img_exts
                          and not _skip_path(p, case_dir, work)]

    # material figure paths: check existence
    fig_status = []
    mat = inv.get("material") or {}
    figs = mat.get("figures") or mat.get("figs") or []
    if isinstance(figs, list):
        for f in figs:
            if isinstance(f, dict):
                fp = f.get("file") or f.get("path") or f.get("image") or ""
                if fp:
                    cand = Path(fp)
                    if not cand.is_absolute():
                        for base in (case_dir, case_dir / "inputs", case_dir / "artifacts",
                                     case_dir / "materials", case_dir / "material"):
                            c2 = base / fp
                            if c2.is_file():
                                cand = c2
                                break
                        else:
                            cand = case_dir / fp
                    fig_status.append({"id": f.get("id", "?"), "path": str(cand),
                                       "exists": cand.is_file()})
    if fig_status:
        inv["material_figures_status"] = fig_status
        missing = [f for f in fig_status if not f["exists"]]
        if missing:
            inv["anomalies"].append(f"{len(missing)} referenced figure file(s) missing")

    try:
        inv["material_digest_file"] = write_material_digest(inv, work)
    except Exception:
        inv["material_digest_file"] = None

    (work / "inventory.json").write_text(
        json.dumps(inv, ensure_ascii=False, indent=1), encoding="utf-8")
    deckutil.log_change(work, "inspect", f"case_id={inv['case_id']} pdfs={len(pdfs)} pptx={len(pptx_files)} anomalies={len(inv['anomalies'])}")

    # compact stdout summary for the agent
    slim = {
        "case_id": inv["case_id"],
        "has_task_md": bool(inv.get("task_md")),
        "workorder": inv.get("workorder"),
        "template": None if not inv.get("template") else {
            "path": inv["template"]["path"], "size_in": inv["template"].get("slide_size_in"),
            "layouts": [f"{l['i']}:{l['name']}({len(l['placeholders'])}ph)"
                        for l in inv["template"].get("layouts", [])][:14]},
        "start_deck_slides": None if not inv.get("start_deck") else len(inv["start_deck"].get("slides", [])),
        "start_deck_dump": inv.get("start_deck_dump"),
        "material_keys": list((inv.get("material") or {}).keys()),
        "material_digest_file": inv.get("material_digest_file"),
        "pdf_pages": (inv.get("pdf_extract") or {}).get("pages"),
        "pdf_text_chars": (inv.get("pdf_extract") or {}).get("text_chars"),
        "pdf_figures": len(((inv.get("pdf_extract") or {}).get("figures")) or []),
        "pdf_figure_files": [{"file": f.get("file"), "page": f.get("page"),
                              "kind": f.get("kind"), "caption": (f.get("caption") or "")[:90]}
                             for f in (((inv.get("pdf_extract") or {}).get("figures")) or [])[:8]],
        "image_files": len(inv.get("image_files", [])),
        "image_paths": inv.get("image_files", [])[:10],
        "extra_text_files": inv.get("extra_text_files", [])[:8],
        "material_text_file": (inv.get("pdf_extract") or {}).get("text_file"),
        "anomalies": inv.get("anomalies", []),
    }
    print("INSPECT_OK " + json.dumps(slim, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
