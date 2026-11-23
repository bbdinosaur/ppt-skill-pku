#!/usr/bin/env python3
"""One-click S0+S1 ingest: inspect case, extract source text/outline/figures, scan QA hooks.

Usage:
  python3 ingest.py [--case /input/case.json] [--work /app/work] [--dpi 180]

Outputs under <work>:
  inspect.json, paper.txt, paper_outline.json, figs/manifest.json, qa_hooks.json, ingest_summary.json
Exit: 0 ok, 2 case.json missing.
No-script fallback: read /input/case.json + task.md, run pdftotext + pdftoppm manually.
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PY = sys.executable

TEMPLATE_KW = ("template", "模板", "master", "母版", "potx")
DECK_KW = ("starting_deck", "start_deck", "initial_deck", "existing_deck", "base_deck", "draft", "start", "initial", "existing", "初始", "已有")


def run_py(script: str, args: list, timeout: int = 240):
    try:
        return subprocess.run([PY, str(HERE / script)] + args, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        return subprocess.CompletedProcess(
            args=[PY, str(HERE / script)] + args,
            returncode=124,
            stdout=(e.stdout.decode("utf-8", errors="ignore") if isinstance(e.stdout, bytes) else (e.stdout or "")),
            stderr=f"TIMEOUT after {timeout}s running {script}",
        )


def walk_kv(obj, path=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}" if path else str(k)
            yield p, k, v
            yield from walk_kv(v, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from walk_kv(v, f"{path}[{i}]")


def resolve_pptx_roles(case: dict, inspect: dict, case_dir: Path):
    """Distinguish starting_deck vs template robustly using case keys, filenames, and slide counts."""
    pptx_info = {str(Path(p["path"]).resolve()): p for p in inspect.get("pptx_files", []) if p.get("exists")}
    starting_deck = None
    template = None

    # 1) Explicit case.json key hints
    for path_k, k, v in walk_kv(case):
        if not isinstance(v, str) or not v.lower().endswith((".pptx", ".potx")):
            continue
        if "/output/" in v.replace("\\", "/") or path_k.lower().startswith("output"):
            continue
        cand = Path(v) if Path(v).is_absolute() else (case_dir / v)
        if not cand.exists():
            continue
        rp = str(cand.resolve())
        kl = (path_k + " " + k).lower()
        if any(w in kl for w in ("start", "initial", "existing", "base_deck", "draft", "deck")) and not any(w in kl for w in ("template", "output", "out_")):
            starting_deck = starting_deck or str(cand)
        elif any(w in kl for w in ("template", "master", "theme", "potx")):
            template = template or str(cand)

    # 2) Fallback to inspect.json pptx_files classification + slide count heuristic
    for p in inspect.get("pptx_files", []):
        if not p.get("exists") or not p.get("ok"):
            continue
        pstr = p["path"]
        name_l = Path(pstr).name.lower()
        n_slides = int(p.get("n_slides") or 0)
        if pstr in (starting_deck, template):
            continue
        if any(w in name_l for w in DECK_KW) and not any(w in name_l for w in TEMPLATE_KW):
            starting_deck = starting_deck or pstr
        elif any(w in name_l for w in TEMPLATE_KW):
            template = template or pstr
        elif n_slides >= 2:
            starting_deck = starting_deck or pstr
        else:
            template = template or pstr

    mode = "edit_deck" if starting_deck else ("from_template" if template else "default_theme")
    return mode, starting_deck, template


def resolve_output_names(case: dict):
    """Infer target deck filename and result filename from case.json / output_contract."""
    deck_name = "talk.pptx"
    result_name = "result.json"
    for path_k, k, v in walk_kv(case):
        if not isinstance(v, str):
            continue
        pl = path_k.lower()
        vl = v.lower()
        is_out_ctx = any(w in pl for w in ("output", "contract", "deliver", "target", "result", "deck_path", "pptx_path")) or "/output/" in vl
        if is_out_ctx and vl.endswith(".pptx"):
            deck_name = Path(v).name
        elif is_out_ctx and vl.endswith(".json") and "case" not in Path(v).name.lower():
            result_name = Path(v).name
    return deck_name, result_name


def extract_extra_pdf_text(pdf_path: Path) -> str:
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        doc = fitz.open(str(pdf_path))
        chunks = [f"--- EXTRA PDF: {pdf_path.name} ({len(doc)} pages) ---"]
        for i, pg in enumerate(doc, 1):
            chunks.append(f"--- PAGE {i} ({pdf_path.name}) ---\n{pg.get_text('text')}")
        doc.close()
        return "\n".join(chunks)
    except Exception:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            chunks = [f"--- EXTRA PDF: {pdf_path.name} ({len(pdf.pages)} pages) ---"]
            for i, pg in enumerate(pdf.pages, 1):
                chunks.append(f"--- PAGE {i} ({pdf_path.name}) ---\n{pg.extract_text() or ''}")
            return "\n".join(chunks)
    except Exception:
        return ""


def _read_xlsx_rows(p: Path):
    """Read first 2 sheets from .xlsx (openpyxl if installed, else stdlib zipfile+XML fallback)."""
    sheets = []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(p), data_only=True, read_only=True)
        for sname in wb.sheetnames[:2]:
            ws = wb[sname]
            rows = []
            for r in ws.iter_rows(values_only=True):
                vals = [" ".join(str(c).split()) if c is not None else "" for c in r[:12]]
                if any(vals):
                    rows.append(vals)
                if len(rows) >= 40:
                    break
            if rows:
                sheets.append((sname, rows))
        wb.close()
        return sheets
    except Exception:
        pass
    try:
        import zipfile
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(str(p)) as z:
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                root_s = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root_s.iter():
                    if si.tag.endswith("}si") or si.tag == "si":
                        shared.append("".join(t.text or "" for t in si.iter() if t.tag.endswith("}t") or t.tag == "t"))
            for name in sorted(n for n in z.namelist() if re.match(r"^xl/worksheets/sheet\d+\.xml$", n))[:2]:
                root_w = ET.fromstring(z.read(name))
                rows = []
                for row in root_w.iter():
                    if not (row.tag.endswith("}row") or row.tag == "row"):
                        continue
                    vals = []
                    for c in row:
                        if not (c.tag.endswith("}c") or c.tag == "c"):
                            continue
                        ct = c.attrib.get("t", "")
                        v_el = next((ch for ch in c if ch.tag.endswith("}v") or ch.tag == "v"), None)
                        raw_v = v_el.text if v_el is not None and v_el.text is not None else ""
                        if ct == "s" and raw_v.isdigit() and int(raw_v) < len(shared):
                            vals.append(" ".join(shared[int(raw_v)].split()))
                        else:
                            vals.append(" ".join(raw_v.split()))
                    if any(vals):
                        rows.append(vals[:12])
                    if len(rows) >= 40:
                        break
                if rows:
                    sheets.append((Path(name).stem, rows))
    except Exception:
        pass
    return sheets


def extract_non_pdf_sources(case_dir: Path, work: Path, append: bool = False, case_obj=None):
    """Extract text & tables from non-PDF input files (.md/.txt/.docx/.csv/.tsv/.tex/.json/.xlsx).

    When append=True (a primary PDF already exists), skips task.md and unreferenced .json files,
    and appends concise supplementary material to paper.txt & paper_tables.json so multi-file facts stay traceable.
    """
    import csv
    cands = []
    case_blob = json.dumps(case_obj, ensure_ascii=False).lower() if case_obj else ""
    TEXT_EXT = {".md", ".txt", ".docx", ".csv", ".tsv", ".tex", ".json", ".xlsx"}
    for p in sorted(case_dir.rglob("*")):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in TEXT_EXT:
            continue
        if p.name.lower() in ("case.json", "manifest.json", "result.json"):
            continue
        if append and p.name.lower() == "task.md":
            continue
        if append and ext == ".json" and p.name.lower() not in case_blob:
            continue
        cands.append(p)
    primary = [p for p in cands if p.name.lower() != "task.md"] or ([] if append else cands)
    if not primary:
        return []
    chunks, sections, tables, used_paths = [], [], [], []
    cap_chars = 15000 if append else 40000
    for p in primary[:8]:
        ext = p.suffix.lower()
        txt = ""
        if ext == ".docx":
            try:
                import docx
                d = docx.Document(str(p))
                lines = [para.text for para in d.paragraphs if para.text.strip()]
                for ti, tb in enumerate(d.tables, 1):
                    trows = []
                    for row in tb.rows:
                        vals = [" ".join(c.text.split()) for c in row.cells]
                        if any(vals):
                            trows.append(vals)
                            lines.append(" | ".join(vals))
                    if len(trows) >= 2 and len(trows[0]) >= 2:
                        tables.append({
                            "page": 1, "label": f"Table {ti} ({p.name})", "caption": f"Table {ti} ({p.name})",
                            "headers": trows[0], "rows": trows[1:14],
                        })
                txt = "\n".join(lines)[:cap_chars]
            except Exception:
                # Stdlib zipfile + XML fallback when python-docx is unavailable
                try:
                    import zipfile
                    import xml.etree.ElementTree as ET
                    with zipfile.ZipFile(str(p)) as z:
                        if "word/document.xml" in z.namelist():
                            root_d = ET.fromstring(z.read("word/document.xml"))
                            W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                            d_lines = []
                            body = root_d.find(f"{W_NS}body")
                            body_children = list(body) if body is not None else list(root_d)
                            t_cnt = 0
                            for child in body_children:
                                if child.tag == f"{W_NS}p":
                                    pt = "".join(t.text or "" for t in child.iter(f"{W_NS}t")).strip()
                                    if pt:
                                        d_lines.append(pt)
                                elif child.tag == f"{W_NS}tbl":
                                    t_cnt += 1
                                    trows = []
                                    for tr in child.findall(f"{W_NS}tr"):
                                        vals = []
                                        for tc in tr.findall(f"{W_NS}tc"):
                                            ct = " ".join("".join(t.text or "" for t in tc.iter(f"{W_NS}t")).split())
                                            vals.append(ct)
                                        if any(vals):
                                            trows.append(vals)
                                            d_lines.append(" | ".join(vals))
                                    if len(trows) >= 2 and len(trows[0]) >= 2:
                                        tables.append({
                                            "page": 1, "label": f"Table {t_cnt} ({p.name})", "caption": f"Table {t_cnt} ({p.name})",
                                            "headers": trows[0], "rows": trows[1:14],
                                        })
                            txt = "\n".join(d_lines)[:cap_chars]
                except Exception:
                    txt = ""
        elif ext in (".csv", ".tsv"):
            try:
                delim = "\t" if ext == ".tsv" else ","
                with p.open("r", encoding="utf-8", errors="ignore", newline="") as fh:
                    rows = [[str(c).strip() for c in r[:12]] for r in csv.reader(fh, delimiter=delim)]
                rows = [r for r in rows if any(r)]
                if len(rows) >= 2 and len(rows[0]) >= 2:
                    tables.append({
                        "page": 1, "label": f"Table ({p.name})", "caption": f"Structured data from {p.name}",
                        "headers": rows[0], "rows": rows[1:20],
                    })
                txt = "\n".join(" | ".join(r) for r in rows[:60])[:cap_chars]
            except Exception:
                txt = p.read_text(encoding="utf-8", errors="ignore")[:cap_chars]
        elif ext == ".xlsx":
            sheets = _read_xlsx_rows(p)
            s_lines = []
            for sname, srows in sheets:
                s_lines.append(f"[Sheet: {sname}]")
                for r in srows:
                    s_lines.append(" | ".join(r))
                if len(srows) >= 2 and len(srows[0]) >= 2:
                    tables.append({
                        "page": 1, "label": f"Table ({p.name}:{sname})", "caption": f"{p.name} sheet {sname}",
                        "headers": srows[0], "rows": srows[1:18],
                    })
            txt = "\n".join(s_lines)[:cap_chars]
        elif ext == ".json":
            try:
                raw_j = p.read_text(encoding="utf-8", errors="ignore")
                if len(raw_j) <= 40000:
                    obj_j = json.loads(raw_j)
                    txt = json.dumps(obj_j, ensure_ascii=False, indent=1)[:cap_chars]
            except Exception:
                txt = ""
        else:
            txt = p.read_text(encoding="utf-8", errors="ignore")[:cap_chars]
        if txt.strip():
            used_paths.append(str(p))
            tag = f"--- SUPPLEMENT ({p.name}) ---" if append else f"--- PAGE 1 ({p.name}) ---"
            chunks.append(f"{tag}\n{txt}")
            if not append:
                for line in txt.splitlines():
                    m = re.match(r"^(#{1,3}\s+.+|\d+(?:\.\d+)*\.?\s+\S.{2,70})$", line.strip())
                    if m:
                        sections.append({"page": 1, "heading": m.group(1).strip()[:90]})
    if not chunks:
        return []
    ptxt = work / "paper.txt"
    if append and ptxt.exists() and ptxt.read_text(encoding="utf-8", errors="ignore").strip():
        cur = ptxt.read_text(encoding="utf-8", errors="ignore")
        ptxt.write_text(cur + "\n\n" + "\n\n".join(chunks), encoding="utf-8")
    else:
        ptxt.write_text("\n\n".join(chunks), encoding="utf-8")
        if not (work / "paper_outline.json").exists():
            (work / "paper_outline.json").write_text(
                json.dumps({"pages": 1, "sections": sections[:30], "captions": []}, ensure_ascii=False, indent=1),
                encoding="utf-8",
            )
    if tables:
        pt_p = work / "paper_tables.json"
        existing = []
        if pt_p.exists():
            try:
                existing = json.loads(pt_p.read_text(encoding="utf-8"))
                if not isinstance(existing, list):
                    existing = []
            except Exception:
                existing = []
        pt_p.write_text(json.dumps(existing + tables, ensure_ascii=False, indent=1), encoding="utf-8")
    (work / "figs").mkdir(parents=True, exist_ok=True)
    if not (work / "figs" / "manifest.json").exists():
        (work / "figs" / "manifest.json").write_text("[]", encoding="utf-8")
    return used_paths


LABEL_STOP = {
    "paper", "study", "research", "academic", "talk", "presentation", "slide", "slides", "deck",
    "method", "methods", "analysis", "results", "system", "systems", "model", "models", "data",
    "general", "science", "scientific", "engineering", "computer", "review", "defense", "seminar",
    "meeting", "brief", "report", "draft", "case", "task", "test", "demo",
    "论文", "研究", "学术", "汇报", "报告", "方法", "实验", "结果", "分析", "系统", "模型", "数据", "科学",
}


def detect_label_mismatch(case_labels: dict, raw_text: str):
    """Detect zero-overlap misleading domain/topic labels in case.json vs actual input file text."""
    if not case_labels or not raw_text or len(raw_text.strip()) < 80:
        return None
    low = raw_text.lower()
    tokens = []
    for k in ("domain", "field", "topic", "subject", "category"):
        val = case_labels.get(k)
        if not val or not isinstance(val, str):
            continue
        for w in re.findall(r"[a-zA-Z]{4,}", val.lower()):
            for sub in re.split(r"[_-]+", w):
                if len(sub) >= 4 and sub not in LABEL_STOP:
                    tokens.append(sub)
        for seg in re.findall(r"[\u4e00-\u9fff]{2,}", val):
            if seg not in LABEL_STOP:
                tokens.append(seg)
                if len(seg) >= 4:
                    for j in range(0, len(seg) - 1, 2):
                        bg = seg[j:j + 2]
                        if len(bg) == 2 and bg not in LABEL_STOP:
                            tokens.append(bg)
    uniq = sorted(set(tokens))
    # Allow single-word misleading domain labels when specific (>=5 ASCII chars or >=2 CJK chars)
    if not uniq or (len(uniq) == 1 and uniq[0].isascii() and len(uniq[0]) < 5):
        return None
    hits = [t for t in uniq if (t in low or (len(t) >= 5 and t[:5] in low))]
    if not hits:
        return (
            f"case.json labels {json.dumps(case_labels, ensure_ascii=False)[:180]} have 0 term overlap "
            f"with actual source text (checked terms: {', '.join(uniq[:8])}). "
            f"Follow actual input files ONLY and record this discrepancy in result.uncertainty."
        )
    return None


def write_reading_map(work: Path, ptxt: Path, outline: dict, paper_tables: list, fig_manifest: list, qa_hooks: list, page_budget: dict = None, starting_deck=None, inspect=None, case_meta=None):
    """Build a compact, truthful navigation index (reading_map.md) pointing to exact paper.txt pages, line ranges, tables, figures, and starting-deck slides."""
    if not ptxt.exists():
        return None
    raw = ptxt.read_text(encoding="utf-8", errors="ignore")
    if not raw.strip():
        return None
    raw_lines = raw.splitlines()
    total_lines = len(raw_lines)
    page_line_starts = {}
    for ln_no, ln_txt in enumerate(raw_lines, 1):
        mp = re.match(r"^--- PAGE (\d+)(?: \([^)]+\))? ---$", ln_txt.strip())
        if mp:
            pno_m = int(mp.group(1))
            if pno_m not in page_line_starts:
                page_line_starts[pno_m] = ln_no
    sorted_pnos = sorted(page_line_starts.keys())
    page_line_ranges = {}
    for idx_p, pno_m in enumerate(sorted_pnos):
        l_start = page_line_starts[pno_m]
        l_end = (page_line_starts[sorted_pnos[idx_p + 1]] - 1) if (idx_p + 1 < len(sorted_pnos)) else total_lines
        page_line_ranges[pno_m] = (l_start, max(l_start, l_end))
    pages = re.split(r"--- PAGE (\d+)(?: \([^)]+\))? ---", raw)
    page_map = {}
    if len(pages) >= 3:
        for i in range(1, len(pages) - 1, 2):
            pno = int(pages[i])
            page_map.setdefault(pno, "")
            page_map[pno] += "\n" + pages[i + 1]
    else:
        page_map[1] = raw
    total_pages = max(page_map.keys()) if page_map else int((outline or {}).get("pages") or 1)
    total_chars = len(raw)

    lines = ["# Source Reading Map & Asset Inventory", ""]
    if case_meta:
        cm = case_meta
        lines.append(
            f"- **Case**: `{cm.get('case_id')}` | **Mode**: `{cm.get('mode')}` | "
            f"**Deck**: `{cm.get('deck_name')}` | **Result**: `{cm.get('result_name')}`"
        )
        if cm.get("constraints"):
            lines.append(f"- **Constraints**: `{json.dumps(cm['constraints'], ensure_ascii=False)[:320]}`")
        if cm.get("contract"):
            lines.append(f"- **Output Contract**: `{json.dumps(cm['contract'], ensure_ascii=False)[:320]}`")
        if cm.get("template"):
            lines.append(f"- **Template**: `{cm['template']}`")
        if cm.get("non_pdf_sources"):
            lines.append(f"- **Supplementary Files (merged into `paper.txt`)**: `{', '.join(Path(x).name for x in cm['non_pdf_sources'])}`")
        if cm.get("case_labels"):
            lines.append(f"- **Case Labels (verify against actual source!)**: `{json.dumps(cm['case_labels'], ensure_ascii=False)[:220]}`")
        # Extract opening actual lines from page 1 so agent immediately verifies actual paper topic vs labels
        p1_lines = [ln.strip() for ln in (page_map.get(1) or raw[:600]).splitlines() if ln.strip() and not ln.strip().startswith("---")]
        if p1_lines:
            lines.append(f"- **Actual Source Opening**: {' / '.join(p1_lines[:2])[:180]}")
        if cm.get("label_mismatch_warning"):
            lines.append(f"> **LABEL MISMATCH ALERT**: {cm['label_mismatch_warning']}")
        lines.append("> **Source-Truth Rule**: Base ALL claims & numbers strictly on actual input files (`paper.txt` & extracted assets). If any `case.json` domain/topic label conflicts with actual files, follow the actual files and log the discrepancy in `uncertainty`.")
        lines.append("")
    if inspect and inspect.get("task_md_path"):
        tmd_p = Path(inspect["task_md_path"])
        tmd_txt = tmd_p.read_text(encoding="utf-8", errors="ignore").strip() if tmd_p.exists() else str(inspect.get("task_md_preview") or "").strip()
        if tmd_txt:
            lines.append("## Task Instructions (`task.md`)")
            lines.append(tmd_txt[:3200])
            if len(tmd_txt) > 3200:
                lines.append(f"> *[task.md truncated at 3200/{len(tmd_txt)} chars — read `{tmd_p}` for remaining lines]*")
            lines.append("")
    if total_chars <= 22000:
        lines.append(f"> `paper.txt` is compact ({total_chars} chars, {total_pages} pages) — read it directly in full before designing `deck_spec.json`.")
    else:
        lines.append(f"> `paper.txt` is {total_chars} chars ({total_pages} pages) — use the Section Page Ranges below to read the actual Introduction, Method, Main Results, and Limitations pages directly (never author slides from snippets alone).")
    lines.append("")

    # 0) Starting Deck Full Content Inventory (when editing/reorganizing an existing deck)
    sd_dump_path = work / "starting_deck_dump.json"
    if starting_deck and sd_dump_path.exists():
        try:
            sd_dump = json.loads(sd_dump_path.read_text(encoding="utf-8"))
        except Exception:
            sd_dump = []
        if sd_dump:
            lines.append(f"## 0. Starting Deck Inventory ({Path(starting_deck).name}, {len(sd_dump)} slides — use 0-based `reuse.slide_index`)")
            for sl_i in sd_dump[:24]:
                lines.append(
                    f"- **`[idx={sl_i.get('index')}]` (Slide {sl_i.get('slide_num')}) {str(sl_i.get('title') or '(untitled)')[:75]}** "
                    f"| layout={sl_i.get('layout')} | pics={sl_i.get('n_pictures', 0)} tbls={len(sl_i.get('tables') or [])} notes={bool(sl_i.get('notes'))}"
                )
                for b_txt in (sl_i.get("bullets") or [])[:5]:
                    lines.append(f"  - {str(b_txt)[:120]}")
                for tb_i in (sl_i.get("tables") or [])[:1]:
                    hdr = tb_i.get("headers") or []
                    if hdr:
                        lines.append("  - [Table] " + " | ".join(str(x)[:18] for x in hdr[:6]) + f" ({len(tb_i.get('rows') or [])} rows)")
                if sl_i.get("notes"):
                    lines.append(f"  - [Notes] {str(sl_i['notes'])[:120]}")
            lines.append("")

    # 1) Section Page-Range Map with exact paper.txt line ranges (for 1-call offset/limit reads on long papers)
    secs = (outline or {}).get("sections") or []
    if secs:
        lines.append("## 1. Paper Section Page Map (`paper.txt`, with line ranges)")
        for si, sec in enumerate(secs[:18]):
            p_start = int(sec.get("page") or 1)
            p_end = int(secs[si + 1].get("page") or p_start) if si + 1 < len(secs) else total_pages
            prng = f"p.{p_start}" if p_end <= p_start else f"p.{p_start}–{p_end}"
            ls = (page_line_ranges.get(p_start) or (None, None))[0]
            le = (page_line_ranges.get(p_end) or page_line_ranges.get(p_start) or (None, None))[1]
            lrng = f" L{ls}–{le}" if (ls and le) else ""
            lines.append(f"- `{prng}{lrng}`: {sec.get('heading')}")
        lim_pages = (outline or {}).get("limitation_or_conclusion_pages") or []
        if lim_pages:
            lp_tags = []
            for lp in lim_pages[:4]:
                rng = page_line_ranges.get(int(lp))
                lp_tags.append(f"p.{lp} (L{rng[0]}–{rng[1]})" if rng else f"p.{lp}")
            lines.append(f"- Limitations / Conclusion pages: {', '.join(lp_tags)}")
        lines.append("")

    # 2) Extracted Figures & High-Confidence Tables
    tbl_caps = [c for c in ((outline or {}).get("captions") or []) if c.get("kind") == "table"]
    if fig_manifest or paper_tables or tbl_caps:
        lines.append("## 2. Extracted Figures & Structured Tables")
        if fig_manifest:
            for m in fig_manifest[:14]:
                lr = " [LOW-RES]" if m.get("low_res") else ""
                lines.append(f"- Figure `{m.get('label')}` (p.{m.get('page')}, {m.get('method')}, {m.get('px')}){lr}: `{m.get('file')}` — {str(m.get('caption',''))[:85]}")
        if paper_tables:
            lines.append("- Structured Tables (`paper_tables.json` — render as native `table`/`chart`, verify numbers in `paper.txt`):")
            for tb in paper_tables[:6]:
                lines.append(f"  - **{tb.get('label')} (p.{tb.get('page')})**: {str(tb.get('caption',''))[:85]}")
                hdr = tb.get("headers") or []
                if hdr:
                    lines.append("    | " + " | ".join(str(x)[:22] for x in hdr[:7]) + " |")
                    lines.append("    | " + " | ".join("---" for _ in hdr[:7]) + " |")
                    for row in (tb.get("rows") or [])[:5]:
                        lines.append("    | " + " | ".join(str(x)[:22] for x in row[:7]) + " |")
        parsed_tbl_lbls = {str(tb.get("label", "")).lower() for tb in (paper_tables or [])}
        unparsed_tbls = [c for c in tbl_caps if str(c.get("label", "")).lower() not in parsed_tbl_lbls]
        if unparsed_tbls:
            lines.append("- Additional Table Captions in `paper.txt` (build native `table`/`chart` directly from `paper.txt`, do NOT embed full-page PNG):")
            for tc in unparsed_tbls[:6]:
                lines.append(f"  - `{tc.get('label')}` (p.{tc.get('page')}): {str(tc.get('caption',''))[:90]}")
        lines.append("")

    # 3) Audience Budget & Q&A Hooks
    if page_budget or qa_hooks:
        lines.append("## 3. Audience Budget & Q&A Hooks (`qa_hooks.json`)")
        if page_budget:
            cb = page_budget.get("category_budget") or {}
            lines.append(
                f"- Budget [`{page_budget.get('audience_class')}`]: total={page_budget.get('total_pages')} "
                f"(content={page_budget.get('content_pages')}, qa_min={page_budget.get('qa_backup_min')}, "
                f"~{page_budget.get('seconds_per_content_slide')}s/slide) | "
                f"bg={cb.get('background',0)}, method={cb.get('method',0)}, evidence={cb.get('evidence',0)}, risk={cb.get('risk',0)}"
            )
            strat = page_budget.get("audience_strategy") or {}
            if strat:
                lines.append(f"  - **Hook**: {strat.get('opening_hook', '')}")
                lines.append(f"  - **Must Include**: {strat.get('must_include', '')}")
                lines.append(f"  - **Cut First**: {strat.get('cut_first', '')}")
                lines.append(f"  - **Q&A Focus**: {strat.get('qa_focus', '')}")
        for h in (qa_hooks or [])[:8]:
            lines.append(f"- [QA/{h.get('type')} @ p.{h.get('page')}] {h.get('snippet')}")
        lines.append("")

    out_p = work / "reading_map.md"
    out_p.write_text("\n".join(lines), encoding="utf-8")
    return str(out_p)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="/input/case.json")
    ap.add_argument("--work", default="/app/work")
    a = ap.parse_args()

    cpath = Path(a.case)
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)

    # 1) S0 inspect_case
    r0 = run_py("inspect_case.py", ["--case", str(cpath), "--work", str(work)])
    if r0.returncode != 0:
        sys.stderr.write(r0.stderr or r0.stdout)
        sys.exit(r0.returncode)

    case = json.loads(cpath.read_text(encoding="utf-8"))
    inspect = json.loads((work / "inspect.json").read_text(encoding="utf-8"))
    mode, starting_deck, template = resolve_pptx_roles(case, inspect, cpath.parent)
    deck_name, result_name = resolve_output_names(case)

    # 2) S1 extract_figs on primary PDF (+ append secondary PDFs if any) or non-PDF fallback
    src_pdf = (inspect.get("source_pdf") or {}).get("path")
    fig_manifest = []
    outline = {}
    paper_tables = []
    extra_pdf_paths = []
    non_pdf_sources = []
    if src_pdf and Path(src_pdf).exists():
        r1 = run_py("extract_figs.py", [
            "--pdf", str(src_pdf),
            "--out", str(work / "figs"),
            "--text", str(work / "paper.txt"),
        ])
        if r1.returncode != 0:
            print("WARN: extract_figs returned non-zero:", (r1.stderr or r1.stdout).strip()[:200])
        # Append any additional source PDFs in input dir so supplementary numbers/locators stay traceable
        extra_pdfs = [
            Path(x["path"]) for x in inspect.get("assets", [])
            if x.get("class") == "pdf" and Path(x["path"]).exists()
            and Path(x["path"]).resolve() != Path(src_pdf).resolve()
        ]
        extra_pdf_paths = [str(ep) for ep in extra_pdfs[:3]]
        if extra_pdfs and (work / "paper.txt").exists():
            extras = [extract_extra_pdf_text(ep) for ep in extra_pdfs[:3]]
            extras = [e for e in extras if e.strip()]
            if extras:
                cur = (work / "paper.txt").read_text(encoding="utf-8", errors="ignore")
                (work / "paper.txt").write_text(cur + "\n\n" + "\n\n".join(extras), encoding="utf-8")
        # ALSO append any supplementary non-PDF research/data files (.csv/.tsv/.md/.txt/.docx/.xlsx/.tex + referenced .json)
        non_pdf_sources = extract_non_pdf_sources(cpath.parent, work, append=True, case_obj=case)
    else:
        non_pdf_sources = extract_non_pdf_sources(cpath.parent, work, append=False, case_obj=case)

    # Extract structured starting_deck_dump.json & append text to paper.txt so inherited facts/numbers stay traceable
    if starting_deck and Path(starting_deck).exists():
        try:
            from pptx import Presentation
            from pptx.enum.shapes import PP_PLACEHOLDER
            prs = Presentation(starting_deck)
            deck_chunks = [f"--- STARTING DECK: {Path(starting_deck).name} ({len(prs.slides)} slides) ---"]
            sd_dump = []
            sd_tables = []
            sd_sections = []
            sd_figs = []
            figs_dir = work / "figs"
            figs_dir.mkdir(parents=True, exist_ok=True)
            def _iter_shapes(shapes_col):
                for _sh in shapes_col:
                    yield _sh
                    if getattr(_sh, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
                        try:
                            yield from _iter_shapes(_sh.shapes)
                        except Exception:
                            pass

            for idx, sl in enumerate(prs.slides, 1):
                title_t = ""
                bullets_t = []
                tbls_s = []
                charts_s = []
                bits = []
                n_pics = 0
                for sh in _iter_shapes(sl.shapes):
                    if sh.shape_type == 13:
                        n_pics += 1
                        # Export genuine content figures from starting deck (skip small icons, logos, header/footer strips)
                        try:
                            sh_top = sh.top if sh.top is not None else 0
                            sh_h = sh.height if sh.height is not None else 0
                            in_chrome = (sh_top < int(0.90 * 914400) and sh_h < int(0.95 * 914400)) or (sh_top > int(prs.slide_height - 0.85 * 914400))
                            im_blob = sh.image.blob
                            im_ext = (sh.image.ext or "png").lower()
                            if (not in_chrome) and im_ext in ("png", "jpg", "jpeg", "bmp", "gif", "webp"):
                                px_info = None
                                try:
                                    from PIL import Image as _PILIm
                                    import io as _io
                                    with _PILIm.open(_io.BytesIO(im_blob)) as _im:
                                        w_px, h_px = _im.size
                                        ar = max(w_px, h_px) / max(1, min(w_px, h_px))
                                        if min(w_px, h_px) >= 120 and max(w_px, h_px) >= 220 and ar <= 4.5:
                                            px_info = [w_px, h_px]
                                except Exception:
                                    if len(im_blob) >= 8192:
                                        px_info = [0, 0]
                                if px_info is not None:
                                    fp_out = figs_dir / f"deck_s{idx}_p{n_pics}.{im_ext}"
                                    fp_out.write_bytes(im_blob)
                                    sd_figs.append({
                                        "file": str(fp_out),
                                        "label": f"Slide {idx} Fig {n_pics}",
                                        "page": idx,
                                        "method": "starting_deck_pic",
                                        "px": px_info,
                                        "caption": f"Embedded image from starting deck Slide {idx}",
                                    })
                        except Exception:
                            pass
                    is_t = bool(sh.is_placeholder and sh.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE))
                    if sh.has_text_frame and sh.text_frame.text.strip():
                        raw_t = sh.text_frame.text.strip()
                        bits.append(raw_t)
                        if is_t and not title_t:
                            title_t = " ".join(raw_t.split())[:100]
                        else:
                            for ln in raw_t.splitlines():
                                lns = ln.strip(" •·-\t")
                                if lns and lns != title_t:
                                    bullets_t.append(lns[:180])
                    if getattr(sh, "has_table", False):
                        rows_clean = []
                        for row in sh.table.rows:
                            rc = [" ".join(c.text.split()) for c in row.cells]
                            rows_clean.append(rc)
                            bits.append(" | ".join(rc))
                        if len(rows_clean) >= 2:
                            tb_obj = {
                                "label": f"Slide {idx} Table",
                                "page": idx,
                                "caption": title_t or f"Slide {idx} table",
                                "headers": rows_clean[0],
                                "rows": rows_clean[1:12],
                            }
                            tbls_s.append(tb_obj)
                            sd_tables.append(tb_obj)
                    if getattr(sh, "has_chart", False):
                        try:
                            ch = sh.chart
                            ch_title = ""
                            if getattr(ch, "has_title", False):
                                try:
                                    ch_title = " ".join(ch.chart_title.text_frame.text.split())[:80]
                                except Exception:
                                    pass
                            pl0 = ch.plots[0] if ch.plots else None
                            cats = [str(c.label) for c in (pl0.categories if pl0 else [])]
                            ser_rows = []
                            for ser in (ch.series or []):
                                vals = [("" if v is None else f"{v:g}" if isinstance(v, float) else str(v)) for v in (ser.values or [])]
                                sname = str(ser.name or "Series")
                                ser_rows.append([sname] + vals[:len(cats) or len(vals)])
                            if ser_rows:
                                hdr = ["Series"] + (cats if cats else [f"C{k+1}" for k in range(len(ser_rows[0]) - 1)])
                                ch_lbl = f"Slide {idx} Chart" + (f" ({ch_title[:32]})" if ch_title else "")
                                bits.append(f"[chart] {ch_lbl}: " + " | ".join(hdr))
                                for sr in ser_rows:
                                    bits.append(" | ".join(sr))
                                ch_tb = {
                                    "label": ch_lbl,
                                    "page": idx,
                                    "caption": ch_title or title_t or f"Slide {idx} chart data",
                                    "headers": hdr[:10],
                                    "rows": [sr[:10] for sr in ser_rows[:10]],
                                }
                                charts_s.append(ch_tb)
                                sd_tables.append(ch_tb)
                        except Exception:
                            pass
                nt = ""
                try:
                    nt = sl.notes_slide.notes_text_frame.text.strip()
                    if nt:
                        bits.append(f"[notes] {nt}")
                except Exception:
                    pass
                sd_dump.append({
                    "index": idx - 1,
                    "slide_num": idx,
                    "layout": getattr(sl.slide_layout, "name", ""),
                    "title": title_t,
                    "bullets": bullets_t[:10],
                    "tables": tbls_s,
                    "charts": charts_s,
                    "n_pictures": n_pics,
                    "notes": nt[:400],
                })
                if title_t:
                    sd_sections.append({"page": idx, "heading": f"Slide {idx}: {title_t}"})
                if bits:
                    deck_chunks.append(f"--- DECK SLIDE {idx} ---\n" + "\n".join(bits))
            (work / "starting_deck_dump.json").write_text(
                json.dumps(sd_dump, ensure_ascii=False, indent=1), encoding="utf-8"
            )
            deck_txt = "\n\n".join(deck_chunks)
            ptxt = work / "paper.txt"
            if ptxt.exists() and ptxt.read_text(encoding="utf-8", errors="ignore").strip():
                ptxt.write_text(ptxt.read_text(encoding="utf-8", errors="ignore") + "\n\n" + deck_txt, encoding="utf-8")
            else:
                # Deck-only case (no PDF/DOCX): format with PAGE markers so downstream anchors/evidence/QA work seamlessly
                paged_chunks = [f"--- PAGE {d['slide_num']} (starting_deck) ---\n{d['title']}\n" + "\n".join(d["bullets"]) + (f"\n[notes] {d['notes']}" if d["notes"] else "") for d in sd_dump]
                for tb in sd_tables:
                    paged_chunks.append(f"--- PAGE {tb['page']} (starting_deck_table) ---\n" + " | ".join(tb["headers"]) + "\n" + "\n".join(" | ".join(r) for r in tb["rows"]))
                ptxt.write_text("\n\n".join(paged_chunks) + "\n\n" + deck_txt, encoding="utf-8")
                if not (work / "paper_outline.json").exists():
                    (work / "paper_outline.json").write_text(
                        json.dumps({"pages": len(sd_dump), "sections": sd_sections, "captions": []}, ensure_ascii=False, indent=1),
                        encoding="utf-8",
                    )
                (work / "figs").mkdir(parents=True, exist_ok=True)
                if not (work / "figs" / "manifest.json").exists():
                    (work / "figs" / "manifest.json").write_text("[]", encoding="utf-8")
            if sd_tables:
                pt_p = work / "paper_tables.json"
                cur_t = []
                if pt_p.exists():
                    try:
                        cur_t = json.loads(pt_p.read_text(encoding="utf-8"))
                        if not isinstance(cur_t, list):
                            cur_t = []
                    except Exception:
                        cur_t = []
                seen_tl = {t.get("label") for t in cur_t if isinstance(t, dict)}
                cur_t.extend(t for t in sd_tables if t["label"] not in seen_tl)
                pt_p.write_text(json.dumps(cur_t, ensure_ascii=False, indent=1), encoding="utf-8")
            if sd_figs:
                mf_p = work / "figs" / "manifest.json"
                cur_mf = []
                if mf_p.exists():
                    try:
                        cur_mf = json.loads(mf_p.read_text(encoding="utf-8"))
                        if not isinstance(cur_mf, list):
                            cur_mf = []
                    except Exception:
                        cur_mf = []
                seen_f = {m.get("file") for m in cur_mf if isinstance(m, dict)}
                cur_mf.extend(m for m in sd_figs if m["file"] not in seen_f)
                mf_p.write_text(json.dumps(cur_mf, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:
            pass

    if (work / "figs" / "manifest.json").exists():
        try:
            fig_manifest = json.loads((work / "figs" / "manifest.json").read_text(encoding="utf-8"))
        except Exception:
            fig_manifest = []
    if (work / "paper_outline.json").exists():
        try:
            outline = json.loads((work / "paper_outline.json").read_text(encoding="utf-8"))
        except Exception:
            outline = {}
    if (work / "paper_tables.json").exists():
        try:
            paper_tables = json.loads((work / "paper_tables.json").read_text(encoding="utf-8"))
        except Exception:
            paper_tables = []

    # 2b) Scanned-PDF handling: render page PNGs + optional tesseract OCR fallback
    likely_scanned = bool((inspect.get("source_pdf") or {}).get("likely_scanned"))
    ocr_used = False
    scanned_page_pngs = []
    if likely_scanned and src_pdf and Path(src_pdf).exists():
        figs_dir = work / "figs"
        figs_dir.mkdir(parents=True, exist_ok=True)
        try:
            try:
                import pymupdf as fitz
            except ImportError:
                import fitz
            doc_s = fitz.open(str(src_pdf))
            for pi in range(min(8, doc_s.page_count)):
                pp = figs_dir / f"scanned_p{pi + 1}.png"
                if not pp.exists():
                    doc_s[pi].get_pixmap(dpi=150).save(str(pp))
                scanned_page_pngs.append(str(pp))
            doc_s.close()
        except Exception:
            pass
        if scanned_page_pngs and shutil.which("tesseract"):
            ocr_chunks = []
            for idx_p, pp in enumerate(scanned_page_pngs, 1):
                try:
                    ro = subprocess.run(["tesseract", pp, "stdout", "--psm", "6"], capture_output=True, text=True, timeout=30)
                    if ro.returncode == 0 and ro.stdout.strip():
                        ocr_chunks.append(f"--- PAGE {idx_p} (OCR) ---\n{ro.stdout.strip()}")
                except Exception:
                    continue
            if ocr_chunks:
                ocr_used = True
                (work / "paper.txt").write_text("\n\n".join(ocr_chunks), encoding="utf-8")

    # 3) QA hook scan + deterministic audience/time page budget + garbled-symbol check
    qa_hooks = []
    garbled_glyphs = 0
    if (work / "paper.txt").exists() and (work / "paper.txt").stat().st_size > 20:
        ptxt_content = (work / "paper.txt").read_text(encoding="utf-8", errors="ignore")
        garbled_glyphs = sum(1 for ch in ptxt_content if ch == "\ufffd" or "\ue000" <= ch <= "\uf8ff")
        rq = run_py("qa_draft.py", ["--source", str(work / "paper.txt"), "--out", str(work / "qa_hooks.json"), "--suggest", "--max", "10"])
        if rq.returncode == 0 and (work / "qa_hooks.json").exists():
            try:
                qa_hooks = json.loads((work / "qa_hooks.json").read_text(encoding="utf-8")).get("hooks", [])
            except Exception:
                qa_hooks = []

    page_budget = {}
    rp_plan = run_py("plan_pages.py", ["plan", "--case", str(cpath), "--out", str(work / "page_budget.json")])
    if rp_plan.returncode == 0 and (work / "page_budget.json").exists():
        try:
            page_budget = json.loads((work / "page_budget.json").read_text(encoding="utf-8"))
        except Exception:
            page_budget = {}

    loose_imgs = [
        x["path"] for x in inspect.get("assets", [])
        if x.get("class") == "image" and Path(x["path"]).exists()
    ]

    case_labels = {}
    for _pk, _k, _v in walk_kv(case):
        if _k in ("domain", "field", "topic", "subject", "category", "paper_title", "title") and isinstance(_v, str) and _v.strip():
            if _k not in case_labels:
                case_labels[_k] = _v.strip()
    ptxt_raw_for_lbl = (work / "paper.txt").read_text(encoding="utf-8", errors="ignore") if (work / "paper.txt").exists() else ""
    label_mismatch_warning = detect_label_mismatch(case_labels, ptxt_raw_for_lbl)
    reading_map_path = write_reading_map(
        work, work / "paper.txt", outline, paper_tables, fig_manifest, qa_hooks,
        page_budget=page_budget, starting_deck=starting_deck, inspect=inspect,
        case_meta={
            "case_id": case.get("case_id") or case.get("id"),
            "mode": mode,
            "deck_name": deck_name,
            "result_name": result_name,
            "template": template,
            "non_pdf_sources": non_pdf_sources,
            "case_labels": case_labels,
            "label_mismatch_warning": label_mismatch_warning,
            "constraints": inspect.get("constraints_digest"),
            "contract": (inspect.get("contract_hits") or [{}])[0].get("value") if inspect.get("contract_hits") else None,
        },
    )

    summary = {
        "case_path": str(cpath),
        "work_dir": str(work),
        "recommended_mode": mode,
        "starting_deck": starting_deck,
        "template": template,
        "recommended_deck_name": deck_name,
        "recommended_result_name": result_name,
        "source_pdf": src_pdf,
        "extra_pdfs": extra_pdf_paths,
        "non_pdf_sources": non_pdf_sources,
        "likely_scanned": likely_scanned,
        "ocr_used": ocr_used,
        "scanned_page_pngs": scanned_page_pngs,
        "garbled_glyph_count": garbled_glyphs,
        "label_mismatch_warning": label_mismatch_warning,
        "page_budget": page_budget,
        "paper_txt": str(work / "paper.txt") if (work / "paper.txt").exists() else None,
        "reading_map": reading_map_path,
        "paper_tables": str(work / "paper_tables.json") if (work / "paper_tables.json").exists() else None,
        "loose_images": loose_imgs,
        "n_extracted_figs": len(fig_manifest),
        "n_extracted_tables": len(paper_tables),
        "n_qa_hooks": len(qa_hooks),
    }
    (work / "ingest_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")

    # Compact executive digest for agent
    print("== INGEST DIGEST ==")
    print(f"case_id={case.get('case_id') or case.get('id')} | mode={mode} | target_deck={deck_name} | target_result={result_name}")
    if label_mismatch_warning:
        print(f"  ALERT [LABEL_MISMATCH]: {label_mismatch_warning}")
    if inspect.get("constraints_digest"):
        print("constraints:", json.dumps(inspect["constraints_digest"], ensure_ascii=False)[:600])
    if page_budget:
        cb = page_budget.get("category_budget") or {}
        inf_tag = " (inferred from duration)" if page_budget.get("inferred_from_duration") else ""
        print(f"page_budget [{page_budget.get('audience_class')}]: total={page_budget.get('total_pages')}{inf_tag} "
              f"(content={page_budget.get('content_pages')}, qa_min={page_budget.get('qa_backup_min')}, "
              f"~{page_budget.get('seconds_per_content_slide')}s/slide) | "
              f"bg={cb.get('background',0)} method={cb.get('method',0)} evidence={cb.get('evidence',0)} risk={cb.get('risk',0)}")
    for h in inspect.get("contract_hits", [])[:3]:
        print("contract:", h["path"], "=", json.dumps(h["value"], ensure_ascii=False)[:480])
    if template:
        print(f"template: {template}")
        tpl_res = str(Path(template).resolve())
        for pf in inspect.get("pptx_files", []):
            if pf.get("path") and str(Path(pf["path"]).resolve()) == tpl_res:
                sz = pf.get("slide_size")
                sz_s = f"{sz[0]/914400:.2f}x{sz[1]/914400:.2f}in" if sz and len(sz) == 2 else "?"
                los = [l["name"] for l in pf.get("layouts", [])][:10]
                print(f"  canvas={sz_s} | layouts: " + "; ".join(los))
    if starting_deck:
        print(f"starting_deck: {starting_deck}")
        sd_res = str(Path(starting_deck).resolve())
        for pf in inspect.get("pptx_files", []):
            if pf.get("path") and str(Path(pf["path"]).resolve()) == sd_res:
                sz = pf.get("slide_size")
                sz_s = f"{sz[0]/914400:.2f}x{sz[1]/914400:.2f}in" if sz and len(sz) == 2 else "?"
                print(f"  canvas={sz_s} | slides={pf.get('n_slides')}")
                for sl in pf.get("slides", [])[:20]:
                    print(f"  [idx={sl['index']}] {sl['title'] or '(no title)'} | pics={sl['n_pictures']} tbls={sl['n_tables']} notes={sl['has_notes']} :: {sl['body_preview'][:65]}")
    if inspect.get("task_md_path"):
        tmd_p = Path(inspect["task_md_path"])
        tmd_txt = tmd_p.read_text(encoding="utf-8", errors="ignore") if tmd_p.exists() else inspect.get("task_md_preview", "")
        print(f"task.md ({len(tmd_txt)} chars): {tmd_p}")
        print("  content:", tmd_txt.strip()[:1800])
        if len(tmd_txt) > 1800:
            print(f"  [TRUNCATED: read full {tmd_p} before drafting deck_spec.json]")
    if outline:
        secs = [f"p{s['page']}:{s['heading']}" for s in outline.get("sections", [])[:12]]
        print(f"paper: pages={outline.get('pages')} scanned={summary['likely_scanned']} ocr={ocr_used} garbled_glyphs={garbled_glyphs} | sections: " + " ; ".join(secs))
    if likely_scanned and not ocr_used:
        print(f"  WARN: PDF appears scanned and tesseract OCR unavailable -> rendered {len(scanned_page_pngs)} page PNGs in {work/'figs'}; inspect visually and record partial/limits")
    if garbled_glyphs >= 10:
        print(f"  WARN: {garbled_glyphs} garbled/PUA glyphs detected in extracted text (likely math symbols); verbalize equations on slides and verify numbers carefully")
    if fig_manifest:
        print(f"extracted figs/tables ({len(fig_manifest)}):")
        for m in fig_manifest[:14]:
            lr = " [LOW-RES]" if m.get("low_res") else ""
            print(f"  [{m.get('method')}] {m.get('file')} | label={m.get('label')} p{m.get('page')} px={m.get('px')}{lr}")
    if paper_tables:
        print(f"structured tables ({len(paper_tables)} in {work/'paper_tables.json'}):")
        for tb in paper_tables[:8]:
            hdr = " | ".join(str(x)[:16] for x in (tb.get("headers") or [])[:6])
            print(f"  [{tb.get('label')}@p{tb.get('page')}] {len(tb.get('rows') or [])}x{len(tb.get('headers') or [])}: {hdr[:85]}")
    if loose_imgs:
        print(f"loose input images ({len(loose_imgs)}): " + ", ".join(loose_imgs[:8]))
    if qa_hooks:
        print(f"qa hooks ({len(qa_hooks)} in {work/'qa_hooks.json'}): top 3 ->")
        for h in qa_hooks[:3]:
            print(f"  [{h['type']}@p{h['page']}] {h['snippet'][:80]}")
    print(f"read next: {work/'reading_map.md'} + {work/'paper.txt'} (full if compact, else key Method/Results/Limitations pages)")
    print(f"artifacts ready: {work/'paper.txt'}, {work/'paper_outline.json'}, {work/'inspect.json'}, {work/'ingest_summary.json'}")


if __name__ == "__main__":
    main()
