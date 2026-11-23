#!/usr/bin/env python3
"""S0 case inspector: parse case.json, discover input assets, inventory template/deck/PDF.

Usage:
  python3 inspect_case.py [--case /input/case.json] [--work /app/work]

Outputs: human-readable summary on stdout + <work>/inspect.json for later steps.
Exit codes: 0 ok, 2 case.json missing/unreadable.
"""
import argparse
import json
import sys
from pathlib import Path

PPTX_EXT = {".pptx", ".ppt", ".potx"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg"}

TEMPLATE_KW = ["template", "模板", "master", "母版", "potx"]
DECK_KW = ["deck", "start", "initial", "初始", "已有", "base", "draft", "existing"]
SOURCE_KW = ["paper", "source", "论文", "材料", "report", "manuscript", "article"]


def classify(name: str, ext: str):
    n = name.lower()
    if ext in PPTX_EXT:
        if any(k in n for k in TEMPLATE_KW):
            return "template"
        if any(k in n for k in DECK_KW):
            return "deck"
        return "pptx_ambiguous"
    if ext == ".pdf":
        return "pdf"
    if ext in IMG_EXT:
        return "image"
    if ext == ".json":
        return "json"
    return "other"


def find_keys(obj, needles, path=""):
    hits = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}" if path else str(k)
            if any(nd in str(k).lower() for nd in needles):
                hits.append((p, v))
            hits += find_keys(v, needles, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            hits += find_keys(v, needles, f"{path}[{i}]")
    return hits


def collect_strings(obj):
    out = []
    if isinstance(obj, dict):
        for v in obj.values():
            out += collect_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            out += collect_strings(v)
    elif isinstance(obj, str):
        out.append(obj)
    return out


def inventory_pptx(path: Path):
    info = {"path": str(path), "exists": path.exists()}
    if not path.exists():
        return info
    try:
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(path)) as z:
            _xml = z.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
        _faces = _re.findall(r'<a:latin typeface="([^"]+)"', _xml)
        info["fonts"] = {"major": _faces[0] if _faces else "", "minor": _faces[1] if len(_faces) > 1 else ""}
    except Exception:
        pass
    try:
        from pptx import Presentation
        from pptx.enum.shapes import PP_PLACEHOLDER
        prs = Presentation(str(path))
        info["slide_size"] = [prs.slide_width, prs.slide_height]
        info["n_slides"] = len(prs.slides)
        layouts = []
        for lo in prs.slide_layouts:
            phs = []
            for ph in lo.placeholders:
                t = ph.placeholder_format.type
                phs.append({"idx": ph.placeholder_format.idx, "type": str(t)})
            layouts.append({"name": lo.name, "placeholders": phs})
        info["layouts"] = layouts
        slide_inventory = []
        for i, sl in enumerate(prs.slides):
            t, body_bits = "", []
            n_pics = n_tbls = 0
            for sh in sl.shapes:
                if sh.is_placeholder and sh.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                    if not t and sh.has_text_frame:
                        t = sh.text_frame.text.strip()[:90]
                elif sh.has_text_frame:
                    tx = " ".join(sh.text_frame.text.split())
                    if tx:
                        body_bits.append(tx)
                if getattr(sh, "has_table", False):
                    n_tbls += 1
                elif sh.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                    n_pics += 1
            has_notes = False
            try:
                has_notes = bool(sl.notes_slide.notes_text_frame.text.strip())
            except Exception:
                pass
            slide_inventory.append({
                "index": i, "slide_num": i + 1, "layout": getattr(sl.slide_layout, "name", ""),
                "title": t, "body_preview": " | ".join(body_bits)[:140],
                "n_pictures": n_pics, "n_tables": n_tbls, "has_notes": has_notes,
            })
        info["slides"] = slide_inventory
        info["slide_titles"] = [{"i": s["slide_num"], "title": s["title"]} for s in slide_inventory]
        info["ok"] = True
    except Exception as e:  # noqa: BLE001
        info["ok"] = False
        info["error"] = repr(e)
    return info


def inventory_pdf(path: Path):
    info = {"path": str(path), "exists": path.exists()}
    if not path.exists():
        return info
    for mod in ("fitz", "pdfplumber"):
        try:
            if mod == "fitz":
                try:
                    import pymupdf as fitz
                except ImportError:
                    import fitz
                doc = fitz.open(str(path))
                info["pages"] = doc.page_count
                txt = doc[0].get_text()[:300] if doc.page_count else ""
                low = [i + 1 for i in range(doc.page_count) if len(doc[i].get_text().strip()) < 30]
                if low:
                    info["low_text_pages"] = low
                    if len(low) > doc.page_count / 2:
                        info["likely_scanned"] = True
            else:
                import pdfplumber
                with pdfplumber.open(str(path)) as pdf:
                    info["pages"] = len(pdf.pages)
                    txt = (pdf.pages[0].extract_text() or "")[:300] if pdf.pages else ""
            info["first_page_snippet"] = txt.replace("\n", " | ")
            info["tool"] = mod
            break
        except Exception as e:  # noqa: BLE001
            info.setdefault("errors", []).append(f"{mod}: {e!r}")
    return info


def extract_constraints_digest(case):
    """Pull key presentation constraints regardless of nesting depth."""
    digest = {}
    for key, needles in (
        ("audience", ["audience", "listener", "受众", "听众"]),
        ("duration_minutes", ["duration", "minutes", "time_limit", "时长"]),
        ("slide_count", ["slide_count", "num_slides", "page_count", "max_slides", "页数"]),
        ("language", ["language", "lang", "语言"]),
        ("report_goal", ["goal", "objective", "purpose", "目标"]),
        ("output_file", ["output_file", "deck_file", "pptx_name", "output_pptx", "deliverable"]),
    ):
        hits = find_keys(case, needles)
        if hits:
            digest[key] = [{"path": p, "value": v} for p, v in hits[:3]]
    return digest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="/input/case.json")
    ap.add_argument("--work", default="/app/work")
    a = ap.parse_args()
    cpath = Path(a.case)
    if not cpath.exists():
        print(f"FATAL: case.json not found at {cpath}", file=sys.stderr)
        sys.exit(2)
    case = json.loads(cpath.read_text(encoding="utf-8"))
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)

    rep = {"case_path": str(cpath), "case_keys": list(case.keys()) if isinstance(case, dict) else []}
    rep["constraints_digest"] = extract_constraints_digest(case)

    # check sibling task.md so S0 sees task instructions in one call
    task_md = cpath.parent / "task.md"
    if task_md.exists():
        t_txt = task_md.read_text(encoding="utf-8", errors="ignore")
        rep["task_md_path"] = str(task_md)
        rep["task_md_preview"] = t_txt[:2000]

    # output contract (verbatim so agent sees exact required fields)
    rep["contract_hits"] = [
        {"path": p, "value": v} for p, v in find_keys(case, ["output_contract", "output_field", "required_output", "result_schema", "contract"])
    ][:10]

    # scan sibling files under the case dir + paths mentioned inside case.json (deduped)
    assets, seen_paths = [], set()
    root = cpath.parent
    for p in sorted(root.rglob("*")):
        if p.is_file():
            rp = str(p.resolve())
            seen_paths.add(rp)
            assets.append({"path": str(p), "size": p.stat().st_size,
                           "class": classify(p.name, p.suffix.lower())})
    for s in collect_strings(case):
        sp = Path(s)
        if sp.suffix.lower() in PPTX_EXT | IMG_EXT | {".pdf"} and len(s) < 400:
            cand = sp if sp.is_absolute() else (root / sp)
            rp = str(cand.resolve()) if cand.exists() else s
            if rp in seen_paths:
                continue
            seen_paths.add(rp)
            assets.append({"path": str(cand) if cand.exists() else s, "size": cand.stat().st_size if cand.exists() else None,
                           "class": classify(sp.name, sp.suffix.lower()),
                           "exists_hint": cand.exists()})
    rep["assets"] = assets

    pdfs = [Path(x["path"]) for x in assets if x["class"] == "pdf" and Path(x["path"]).exists()]
    if pdfs:
        pdfs.sort(key=lambda p: (not any(k in p.name.lower() for k in SOURCE_KW), -p.stat().st_size))
        rep["source_pdf"] = inventory_pdf(pdfs[0])
    ppts = [Path(x["path"]) for x in assets if x["class"] in ("template", "deck", "pptx_ambiguous") and Path(x["path"]).exists()]
    if ppts:
        ppts.sort(key=lambda p: 0 if any(k in p.name.lower() for k in TEMPLATE_KW) else 1)
        rep["pptx_files"] = [inventory_pptx(p) for p in ppts[:4]]

    (work / "inspect.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")

    # human summary
    print("== case keys:", rep["case_keys"])
    if rep["constraints_digest"]:
        print("== constraints:", json.dumps(rep["constraints_digest"], ensure_ascii=False)[:600])
    for h in rep["contract_hits"]:
        print("== contract:", h["path"], "=", json.dumps(h["value"], ensure_ascii=False)[:800])
    for ass in assets:
        print(f"   [{ass['class']:>14}] {ass['path']}" + ("" if ass.get("exists_hint") is None else f" exists={ass['exists_hint']}"))
    if "source_pdf" in rep:
        sp = rep["source_pdf"]
        print(f"== source pdf: {sp['path']} pages={sp.get('pages')} tool={sp.get('tool')}" +
              (" LIKELY_SCANNED" if sp.get("likely_scanned") else ""))
    for pf in rep.get("pptx_files", []):
        names = [l["name"] for l in pf.get("layouts", [])]
        print(f"== pptx: {pf['path']} slides={pf.get('n_slides')} size={pf.get('slide_size')} ok={pf.get('ok')}")
        print("   layouts:", "; ".join(names))
        if pf.get("slides"):
            for sl in pf["slides"][:25]:
                print(f"   [idx={sl['index']}] {sl['title'] or '(no title)'} | pics={sl['n_pictures']} tbls={sl['n_tables']} notes={sl['has_notes']} :: {sl['body_preview'][:75]}")
    if rep.get("task_md_path"):
        print(f"== task.md found: {rep['task_md_path']} ({len(rep.get('task_md_preview',''))} chars previewed in inspect.json)")
    print(f"inspect.json -> {work/'inspect.json'}")


if __name__ == "__main__":
    main()
