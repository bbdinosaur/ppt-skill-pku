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

from pptx import Presentation  # noqa: E402
from pptx.util import Emu  # noqa: E402


def find_first(root: Path, names):
    for n in names:
        for cand in (root / n, root / "artifacts" / n, root / "inputs" / n):
            if cand.is_file():
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


def extract_pdf(pdf_path: Path, work: Path):
    """Extract per-page text and embedded images from a PDF. Returns dict summary."""
    out = {"path": str(pdf_path), "pages": 0, "text_file": None, "figures": [], "anomalies": []}
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz as fitz
        doc = fitz.open(str(pdf_path))
        out["pages"] = doc.page_count
        chunks = []
        for pno in range(doc.page_count):
            page = doc[pno]
            chunks.append(f"\n\n<!-- PDF page {pno + 1} -->\n" + page.get_text("text"))
            for i, info in enumerate(page.get_images(full=True)):
                xref = info[0]
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    if pix.width < 40 or pix.height < 40:
                        continue
                    figdir = work / "figs"
                    figdir.mkdir(parents=True, exist_ok=True)
                    fp = figdir / f"pdf_p{pno + 1}_img{i}.png"
                    pix.save(str(fp))
                    out["figures"].append({"file": str(fp), "page": pno + 1,
                                           "w": pix.width, "h": pix.height})
                except Exception:
                    out["anomalies"].append(f"p{pno + 1}: image xref {xref} extract failed")
        text = "".join(chunks)
        # paper map: per-page heading candidates (font-size heuristic) for cheap navigation
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
            if heads:
                text = ("\n<!-- 论文导航地图 / paper map (heading candidates by page) -->\n"
                        + "\n".join(heads) + "\n" + text)
        except Exception:
            pass
        tfp = work / "material_text.md"
        tfp.write_text(text, encoding="utf-8")
        out["text_file"] = str(tfp)
        out["text_chars"] = len(text)
        doc.close()
    except Exception as e:
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
    m = str(mode).lower()
    if any(k in m for k in ("edit", "adapt", "revise", "modify", "更新", "修改", "编辑")):
        out["mode"] = "edit"
    elif any(k in m for k in ("create", "generate", "build", "new", "生成", "创建", "制作")):
        out["mode"] = "create"
    else:
        out["mode"] = "auto"
    for keys, dst in (
        (("slide_count", "total_slides", "slides", "num_slides", "page_count", "页数"), "slide_count"),
        (("max_slides", "slide_limit", "max_pages", "页数上限"), "max_slides"),
        (("min_slides",), "min_slides"),
        (("audience", "受众", "target_audience"), "audience"),
        (("language", "语言"), "language"),
        (("duration", "time_limit", "时长"), "duration"),
        (("title", "deck_title", "题目"), "deck_title"),
    ):
        v = g(*keys)
        if v is not None:
            out[dst] = v
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", default="/input")
    ap.add_argument("--work", default="/app/work")
    args = ap.parse_args()
    case_dir = Path(args.case_dir)
    work = Path(args.work)
    work.mkdir(parents=True, exist_ok=True)

    inv = {"case_dir": str(case_dir), "work": str(work), "anomalies": []}

    case_json = find_first(case_dir, ("case.json",))
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

    wo_path = find_first(case_dir, ("slide-workorder.json", "workorder.json"))
    if wo_path:
        try:
            inv["workorder_raw"] = json.loads(wo_path.read_text(encoding="utf-8"))
            inv["workorder"] = normalize_workorder(inv["workorder_raw"])
        except Exception as e:
            inv["anomalies"].append(f"workorder unreadable: {e}")
    else:
        inv["anomalies"].append("slide-workorder.json not found — infer constraints from case.json/task.md")

    task_md = find_first(case_dir, ("task.md",))
    if task_md:
        inv["task_md"] = str(task_md)

    # ---- PPTX inputs
    pptx_files = sorted(p for p in case_dir.rglob("*.pptx") if p.is_file())
    template = None
    start_deck = None
    for p in pptx_files:
        n = p.name.lower()
        if "template" in n or "模板" in n:
            template = p
        elif any(k in n for k in ("start", "deck", "draft", "old", "base")) and n != (template.name.lower() if template else ""):
            if start_deck is None:
                start_deck = p
    inv["template"] = describe_pptx(template) if template else None
    inv["start_deck"] = describe_pptx(start_deck) if start_deck else None
    inv["pptx_files"] = [str(p) for p in pptx_files]
    if template is None and pptx_files:
        # fall back: biggest pptx treated as template, next as start deck
        by_size = sorted(pptx_files, key=lambda p: p.stat().st_size, reverse=True)
        inv["template"] = describe_pptx(by_size[0])
        if len(by_size) > 1 and inv.get("start_deck") is None:
            inv["start_deck"] = describe_pptx(by_size[1])
        inv["anomalies"].append("template not identified by name — used size heuristic")

    # ---- material
    mat_path = find_first(case_dir, ("material.json", "materials.json", "paper.md",
                                     "material.md", "source.md", "content.json"))
    inv["material_json_path"] = str(mat_path) if mat_path else None
    if mat_path:
        try:
            inv["material"] = json.loads(mat_path.read_text(encoding="utf-8"))
        except Exception as e:
            inv["anomalies"].append(f"material.json unreadable: {e}")
            inv["material"] = {}

    pdfs = sorted(p for p in case_dir.rglob("*.pdf") if p.is_file() and "output" not in p.parent.name.lower())
    inv["pdf_inputs"] = []
    inv["pdf_extract"] = None
    for p in pdfs:
        inv["pdf_inputs"].append(str(p))
    if pdfs:
        inv["pdf_extract"] = extract_pdf(pdfs[0], work)

    text_exts = (".txt", ".md")
    if not inv.get("pdf_extract"):
        docx_p = next((p for p in sorted(case_dir.rglob("*.docx")) if p.is_file() and "output" not in p.parent.name.lower()), None)
        txt_p = next((p for p in sorted(case_dir.rglob("*")) if p.is_file() and p.suffix.lower() in text_exts
                      and p.name not in ("task.md",) and "output" not in p.parent.name.lower()), None)
        src = docx_p or txt_p
        if src is not None:
            try:
                chunks = []
                if docx_p is not None:
                    import docx
                    d = docx.Document(str(docx_p))
                    for para in d.paragraphs:
                        if para.text.strip():
                            chunks.append(para.text)
                    for i, tbl in enumerate(d.tables):
                        for row in tbl.rows:
                            chunks.append(" | ".join(c.text.strip() for c in row.cells))
                else:
                    chunks.append(src.read_text(encoding="utf-8", errors="replace"))
                text = "\n".join(chunks)
                tfp = work / "material_text.md"
                tfp.write_text(text, encoding="utf-8")
                inv["pdf_extract"] = {"text_file": str(tfp), "pages": None,
                                       "figures": [], "anomalies": [],
                                       "source": str(src), "text_chars": len(text)}
            except Exception as e:
                inv["anomalies"].append(f"material text extract failed: {e}")

    # ---- figure images shipped with the case
    img_exts = (".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff", ".bmp", ".svg", ".emf")
    inv["image_files"] = [str(p) for p in sorted(case_dir.rglob("*"))
                          if p.is_file() and p.suffix.lower() in img_exts]

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

    (work / "inventory.json").write_text(
        json.dumps(inv, ensure_ascii=False, indent=1), encoding="utf-8")
    deckutil.log_change(work, "inspect", f"case_id={inv['case_id']} pdfs={len(pdfs)} pptx={len(pptx_files)} anomalies={len(inv['anomalies'])}")

    # compact stdout summary for the agent
    slim = {
        "case_id": inv["case_id"],
        "workorder": inv.get("workorder"),
        "template": None if not inv.get("template") else {
            "path": inv["template"]["path"], "size_in": inv["template"].get("slide_size_in"),
            "layouts": [f"{l['i']}:{l['name']}({len(l['placeholders'])}ph)"
                        for l in inv["template"].get("layouts", [])][:14]},
        "start_deck_slides": None if not inv.get("start_deck") else len(inv["start_deck"].get("slides", [])),
        "material_keys": list((inv.get("material") or {}).keys()),
        "pdf_pages": (inv.get("pdf_extract") or {}).get("pages"),
        "pdf_text_chars": (inv.get("pdf_extract") or {}).get("text_chars"),
        "pdf_figures": len(((inv.get("pdf_extract") or {}).get("figures")) or []),
        "image_files": len(inv.get("image_files", [])),
        "material_text_file": (inv.get("pdf_extract") or {}).get("text_file"),
        "anomalies": inv.get("anomalies", []),
    }
    print("INSPECT_OK " + json.dumps(slim, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
