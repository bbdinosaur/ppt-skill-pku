#!/usr/bin/env python3
"""Step-1 one-stop briefing: inventory /input, template layout slots, paper figures/tables/key sections,
and optional starting deck in a single bounded-stdout command (default <= 32 KB).

Usage:
  python3 prepare_brief.py [--input /input] [--workdir /app/work] [--max-bytes 32000]
      [--case PATH] [--task PATH] [--template PATH] [--paper PATH] [--deck PATH]

Outputs written to --workdir:
  - template_inventory.json (when template exists)
  - paper_structure.json, paper_text.txt, figs/*.png (when paper/material exists)
  - deck_inventory.json, plan_draft.json, deck_imgs/* (when starting deck exists)
  - brief.txt (exact copy of compact stdout briefing)
No-script fallback: run inspect_template.py + extract_paper.py (+ inspect_deck.py) individually.
"""
import argparse, json, os, re, sys
from pathlib import Path

sys.dont_write_bytecode = True
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import inspect_template
import extract_paper
import inspect_deck

try:
    from PIL import Image
except ImportError:
    Image = None

TPL_NAME_RE = re.compile(r"(template|tpl|theme|master|blank|brand|style|model)", re.I)
DECK_NAME_RE = re.compile(r"(start|draft|old|legacy|existing|current|original|source_ppt|input_deck|base_deck|v0)", re.I)
TIER3_SEC_RE = re.compile(
    r"(experiment|result|evaluat|ablat|bench|compar|perform|finding|discuss|limit|threat|conclu|future|takeaway|"
    r"实验|结果|评估|消融|对比|讨论|局限|结论)", re.I)
TIER1_SEC_RE = re.compile(
    r"(related\s+work|prior\s+work|acknowledg|ethic|broader\s+impact|reference|bibliograph|appendix|supplement|"
    r"相关工作|致谢|参考文献|附录)", re.I)
REF_STOP_RE = re.compile(r"^\s*(\d+\.?\s+)?(references|bibliography|acknowledg(e)?ments?|参考文献|致谢)\s*$", re.I)
SIGNAL_SENT_RE = re.compile(
    r"(\d+(?:\.\d+)?\s*(?:%|x|×|ms|s|gb|mb|kb|hz|khz|mhz|ghz|mm|cm|nm|ev|k\b|m\b|b\b|db|pt|fps)|"
    r"±|\bp\s*[<≤=]\s*0\.\d|f1|auc|bleu|rouge|rmse|mae|top-\d|table\s+\w+|fig(?:ure)?\.?\s+\w+|"
    r"outperform|improv|reduc|faster|lower|higher|gain|achieve|limitation|bottleneck|however|fail|trade-?off|"
    r"提升|降低|优于|局限|不足|误差)", re.I)


def _utf8_trunc(text: str, max_bytes: int, suffix: str = " ...[truncated]") -> str:
    raw = text.encode("utf-8")
    if len(raw) <= max_bytes:
        return text
    sb = suffix.encode("utf-8")
    keep = max(0, max_bytes - len(sb))
    cut = raw[:keep].decode("utf-8", errors="ignore").rstrip()
    return cut + suffix


def _clean_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _collect_case_hints(obj, tpl_hints, deck_hints, paper_hints, key_ctx=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            _collect_case_hints(v, tpl_hints, deck_hints, paper_hints, f"{key_ctx}.{k}".lower())
    elif isinstance(obj, list):
        for v in obj:
            _collect_case_hints(v, tpl_hints, deck_hints, paper_hints, key_ctx)
    elif isinstance(obj, str) and len(obj) < 300:
        low = obj.lower()
        base = Path(obj).name.lower()
        if any(ext in low for ext in (".pptx", ".potx")):
            if any(t in key_ctx for t in ("template", "tpl", "theme", "master")):
                tpl_hints.add(base)
            elif any(t in key_ctx for t in ("start", "deck", "draft", "source", "old", "existing", "input")):
                deck_hints.add(base)
        elif any(ext in low for ext in (".pdf", ".docx", ".md", ".txt", ".csv")):
            if any(t in key_ctx for t in ("paper", "pdf", "source", "manuscript", "article", "material")):
                paper_hints.add(base)


def _scan_input(input_dir: Path):
    files = []
    if not input_dir.exists():
        return files
    for root, dirs, fnames in os.walk(input_dir):
        rel_depth = len(Path(root).relative_to(input_dir).parts)
        dirs[:] = sorted(d for d in dirs if not d.startswith(".") and d != "__pycache__" and rel_depth < 4)
        for fn in sorted(fnames):
            if fn.startswith("."):
                continue
            p = Path(root) / fn
            if p.is_file():
                files.append(p)
    return files


def _classify_inputs(files, case_data, a):
    tpl_hints, deck_hints, paper_hints = set(), set(), set()
    _collect_case_hints(case_data, tpl_hints, deck_hints, paper_hints)

    case_p = Path(a.case) if a.case else next((f for f in files if f.name.lower() == "case.json"), None)
    task_p = Path(a.task) if a.task else next(
        (f for f in files if f.name.lower() in ("task.md", "prompt.md", "instructions.md")), None)

    ppt_files = [f for f in files if f.suffix.lower() in (".pptx", ".potx")]
    pdf_files = [f for f in files if f.suffix.lower() == ".pdf"]
    doc_files = [
        f for f in files
        if f.suffix.lower() in (".md", ".txt", ".docx", ".csv", ".tsv", ".json")
        and f not in (case_p, task_p)
        and f.name.lower() not in ("case.json", "task.md")
    ]
    img_files = [f for f in files if f.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp")]

    tpl_p = Path(a.template) if a.template else None
    deck_p = Path(a.deck) if a.deck else None

    if tpl_p is None and deck_p is None and ppt_files:
        if len(ppt_files) == 1:
            p0 = ppt_files[0]
            tpl_p = p0
            # Check if the single .pptx also contains multiple slides (could be a starting deck + template)
            if p0.suffix.lower() == ".pptx":
                try:
                    inv0 = inspect_template.inspect(p0)
                    if inv0.get("existing_slides", 0) >= 2 or p0.name.lower() in deck_hints or DECK_NAME_RE.search(p0.name):
                        deck_p = p0
                except Exception:
                    pass
        else:
            scored = []
            for p in ppt_files:
                nm = p.name.lower()
                is_potx = p.suffix.lower() == ".potx"
                t_score = (10 if is_potx else 0) + (6 if nm in tpl_hints else 0) + (4 if TPL_NAME_RE.search(nm) else 0)
                d_score = (6 if nm in deck_hints else 0) + (4 if DECK_NAME_RE.search(nm) else 0)
                n_slides = 0
                try:
                    n_slides = inspect_template.inspect(p).get("existing_slides", 0)
                except Exception:
                    pass
                scored.append((p, t_score, d_score, n_slides))
            # Pick best template (prefer high t_score, fewer existing slides)
            scored_tpl = sorted(scored, key=lambda x: (-x[1], x[2], x[3], x[0].name))
            tpl_p = scored_tpl[0][0]
            rem = [x for x in scored if x[0] != tpl_p]
            if rem:
                scored_deck = sorted(rem, key=lambda x: (-x[2], -x[3], x[0].name))
                if scored_deck[0][2] > 0 or scored_deck[0][3] >= 1:
                    deck_p = scored_deck[0][0]
    elif tpl_p is None and ppt_files:
        rem = [p for p in ppt_files if p != deck_p]
        if rem:
            tpl_p = rem[0]
    elif deck_p is None and ppt_files:
        rem = [p for p in ppt_files if p != tpl_p]
        if rem:
            deck_p = rem[0]

    paper_p = Path(a.paper) if a.paper else None
    extra_pdfs = []
    if paper_p is None:
        if pdf_files:
            ranked_pdf = sorted(
                pdf_files,
                key=lambda p: (0 if p.name.lower() in paper_hints else 1, -p.stat().st_size, p.name),
            )
            paper_p = ranked_pdf[0]
            extra_pdfs = ranked_pdf[1:]
        elif doc_files:
            ranked_doc = sorted(
                doc_files,
                key=lambda p: (0 if p.name.lower() in paper_hints else 1,
                               0 if p.suffix.lower() in (".docx", ".md", ".txt") else 1,
                               -p.stat().st_size),
            )
            paper_p = ranked_doc[0]
            doc_files = [d for d in doc_files if d != paper_p]
    else:
        extra_pdfs = [p for p in pdf_files if p.resolve() != paper_p.resolve()]
        doc_files = [d for d in doc_files if d.resolve() != paper_p.resolve()]

    return dict(
        case=case_p, task=task_p, template=tpl_p, deck=deck_p,
        paper=paper_p, extra_pdfs=extra_pdfs, extra_docs=doc_files, images=img_files,
    )


def _fmt_template(inv: dict, max_bytes: int = 4400) -> str:
    lines = [
        "## 2. TEMPLATE LAYOUT SLOTS & THEME",
        f"- file: `{inv.get('template')}` | size: {inv.get('slide_width_in')}x{inv.get('slide_height_in')} in "
        f"| is_potx={inv.get('is_potx')} | existing_slides={inv.get('existing_slides')} | layouts={inv.get('n_layouts_total')}",
        f"- theme_fonts: {json.dumps(inv.get('theme_fonts') or {}, ensure_ascii=False)}",
        f"- theme_colors: {json.dumps(inv.get('theme_colors') or {}, ensure_ascii=False)}",
    ]
    all_lo = []
    for m_i, m in enumerate(inv.get("masters") or []):
        m_dm = m.get("decor_margins") or {}
        if any(v > 0 for v in m_dm.values()) or m.get("bg_color"):
            lines.append(f"- master[{m_i}]: bg={m.get('bg_color')} decor_margins={m_dm}")
        for lo in (m.get("layouts") or []):
            all_lo.append(lo)

    lines.append("- layouts (use exact `name` for `layout_hint` in plan.json):")
    for idx, lo in enumerate(all_lo):
        bg_tag = f"{lo.get('bg_color') or 'inherit'}({'DARK' if lo.get('dark_bg') else 'light'})"
        ph_bits = []
        for ph in (lo.get("placeholders") or []):
            pt = str(ph.get("type", "")).replace("PLACEHOLDER", "").strip()
            if pt in ("FOOTER", "SLIDE_NUMBER", "DATE", "HEADER"):
                continue
            ph_bits.append(f"{ph['idx']}:{pt}@{ph['left']},{ph['top']},{ph['width']}x{ph['height']}")
        dm = lo.get("decor_margins") or {}
        dm_tag = f" decor={dm}" if any(v > 0 for v in dm.values()) else ""
        row = (f"  [{idx}] \"{lo['name']}\" | bg={bg_tag} | "
               f"title={int(bool(lo.get('has_title')))} sub={int(bool(lo.get('has_subtitle')))} "
               f"body={lo.get('text_body_count', 0)} pic={lo.get('picture_count', 0)} "
               f"foot={int(bool(lo.get('has_footer')))} snum={int(bool(lo.get('has_slide_number')))}{dm_tag}"
               + (f" | phs: {', '.join(ph_bits[:6])}" if ph_bits else ""))
        cand = "\n".join(lines + [row])
        if len(cand.encode("utf-8")) > max_bytes - 120:
            lines.append(f"  ... ({len(all_lo) - idx} more layouts in template_inventory.json)")
            break
        lines.append(row)
    return "\n".join(lines)


def _fmt_table_rows(rows, max_rows=9, max_cols=8) -> list:
    out = []
    for r_i, row in enumerate((rows or [])[:max_rows]):
        cells = [_clean_ws(str(c))[:36] for c in row[:max_cols]]
        out.append("    | " + " | ".join(cells) + " |")
    if rows and len(rows) > max_rows:
        out.append(f"    ... ({len(rows) - max_rows} more rows in paper_structure.json)")
    return out


def _fmt_paper_core(ps: dict, max_bytes: int = 9800) -> str:
    lines = ["## 3. PAPER METADATA, FIGURES, TABLES & KEY STATS"]
    title = (ps.get("title") or {}).get("text") or "(untitled)"
    lines.append(f"- source: `{ps.get('source')}` | pages={ps.get('pages')} | body_font={ps.get('body_font')}pt")
    lines.append(f"- title: {_clean_ws(title)}")
    abs_txt = _clean_ws((ps.get("abstract") or {}).get("text") or "")
    if abs_txt:
        lines.append(f"- abstract (p.{(ps.get('abstract') or {}).get('page', 1)}): {_utf8_trunc(abs_txt, 1400)}")

    figs = ps.get("figures") or []
    lines.append(f"- figures ({len(figs)}):")
    for f in figs:
        img_info = (f"`{f['image']}` ({f.get('image_w')}x{f.get('image_h')})"
                    if f.get("image") else f"NO_CROP (use pymupdf clip on p.{f.get('page')})")
        cap = _clean_ws(f.get("caption") or "")[:240]
        lines.append(f"  * Fig. {f.get('number')} (p.{f.get('page')}): {cap} | img={img_info}")

    tabs = ps.get("tables") or []
    lines.append(f"- tables ({len(tabs)}):")
    for t in tabs:
        rows = t.get("rows") or []
        img_tag = f" | img=`{t['image']}`" if t.get("image") else ""
        cap = _clean_ws(t.get("caption") or "")[:220]
        lines.append(f"  * Table {t.get('number')} (p.{t.get('page')}): {cap} | rows={len(rows)}{img_tag}")
        if rows:
            for rline in _fmt_table_rows(rows, max_rows=8, max_cols=8):
                if len(("\n".join(lines + [rline])).encode("utf-8")) > max_bytes - 1600:
                    lines.append("    ... (remaining table rows in paper_structure.json)")
                    break
                lines.append(rline)

    stats = ps.get("key_stats") or []
    if stats:
        lines.append(f"- key_stats ({len(stats)}):")
        for st in stats[:18]:
            row = f"  * [p.{st.get('page', '?')}] {_clean_ws(st.get('text', ''))[:200]}"
            if len(("\n".join(lines + [row])).encode("utf-8")) > max_bytes - 120:
                break
            lines.append(row)

    return _utf8_trunc("\n".join(lines), max_bytes)


def _split_pages(text_file: Path):
    if not text_file or not text_file.exists():
        return []
    raw = text_file.read_text(encoding="utf-8", errors="ignore")
    pages = []
    cur_p, cur_lines = 1, []
    for ln in raw.splitlines():
        m = re.match(r"^===\s*PAGE\s+(\d+)\s*===$", ln.strip())
        if m:
            if cur_lines:
                pages.append((cur_p, "\n".join(cur_lines)))
            cur_p = int(m.group(1))
            cur_lines = []
        else:
            cur_lines.append(ln)
    if cur_lines:
        pages.append((cur_p, "\n".join(cur_lines)))
    return pages


def _extract_signal_sentences(text: str, already: str, budget_bytes: int) -> str:
    if budget_bytes < 120 or not text:
        return ""
    sents = re.split(r"(?<=[.!?。！？])\s+", _clean_ws(text))
    picked, used = [], 0
    already_low = already.lower()
    for s in sents:
        sc = s.strip()
        if len(sc) < 25 or len(sc) > 320:
            continue
        if sc[:50].lower() in already_low:
            continue
        if SIGNAL_SENT_RE.search(sc):
            b = len(sc.encode("utf-8")) + 3
            if used + b > budget_bytes:
                break
            picked.append(sc)
            used += b
    return " | ".join(picked)


def _fmt_paper_sections(ps: dict, text_file: Path, budget_bytes: int) -> str:
    if budget_bytes < 500:
        return ""
    pages = _split_pages(text_file)
    # Build ordered full lines with page tags, stopping at References
    tagged_lines = []
    ref_reached = False
    for pno, ptxt in pages:
        for ln in ptxt.splitlines():
            s = ln.strip()
            if not s:
                continue
            if pno > 1 and REF_STOP_RE.match(s):
                ref_reached = True
                break
            tagged_lines.append((pno, s))
        if ref_reached:
            break

    secs = ps.get("sections") or []
    lines = ["## 4. PAPER SECTIONS & PRIORITY EXCERPTS (Results/Experiments Protected)"]

    # Try locating each section in tagged_lines to get full section body (not just 520-char excerpt)
    usable_secs = []
    if len(secs) >= 2 and tagged_lines:
        anchors = []
        cursor = 0
        for s in secs:
            stitle = _clean_ws(s.get("title") or "")
            snum = _clean_ws(str(s.get("number") or ""))
            spage = s.get("page") or 1
            hit = None
            needle = stitle[:28].lower()
            for idx in range(cursor, len(tagged_lines)):
                pno, txt = tagged_lines[idx]
                if pno < spage:
                    continue
                if pno > spage + 1 and hit is None:
                    break
                if needle and needle in txt.lower():
                    hit = idx
                    cursor = idx + 1
                    break
            anchors.append(hit)

        for i, s in enumerate(secs):
            stitle = _clean_ws(s.get("title") or "")
            snum = _clean_ws(str(s.get("number") or ""))
            label = f"{snum} {stitle}".strip()
            if REF_STOP_RE.match(stitle) or REF_STOP_RE.match(label):
                usable_secs.append(dict(label=label, page=s.get("page", 1), tier=0, body=""))
                continue
            if TIER1_SEC_RE.search(label):
                tier = 1
            elif TIER3_SEC_RE.search(label):
                tier = 3
            else:
                tier = 2

            full_body = ""
            if anchors[i] is not None:
                start = anchors[i] + 1
                nxt = next((a for a in anchors[i + 1:] if a is not None), len(tagged_lines))
                chunk_parts = []
                last_p = s.get("page") or 1
                for pno, txt in tagged_lines[start:nxt]:
                    if pno != last_p:
                        chunk_parts.append(f"[p.{pno}]")
                        last_p = pno
                    chunk_parts.append(txt)
                full_body = " ".join(chunk_parts)
            if not full_body:
                full_body = _clean_ws(s.get("excerpt") or "")
            usable_secs.append(dict(label=label, page=s.get("page", 1), tier=tier, body=full_body))

    if usable_secs:
        avail = max(400, budget_bytes - 160)
        weights = {3: 3.0, 2: 2.0, 1: 0.35, 0: 0.0}
        w_sum = sum(weights[s["tier"]] for s in usable_secs) or 1.0
        for s in usable_secs:
            if s["tier"] == 0:
                lines.append(f"- [p.{s['page']}] {s['label']} (omitted — references/ack)")
                continue
            if s["tier"] == 1:
                stub = _utf8_trunc(_clean_ws(s["body"]), 180)
                lines.append(f"- [p.{s['page']}] {s['label']} (compact): {stub}")
                continue
            if not s["body"]:
                lines.append(f"- [p.{s['page']}] {s['label']}")
                continue
            alloc = int(avail * (weights[s["tier"]] / w_sum))
            alloc = max(360, min(alloc, 2600 if s["tier"] == 3 else 1800))
            head_b = int(alloc * 0.62)
            tail_b = alloc - head_b
            head = _utf8_trunc(s["body"], head_b, suffix=" ...")
            sig = _extract_signal_sentences(s["body"], head, tail_b)
            entry = f"- [p.{s['page']}] {s['label']}: {head}"
            if sig:
                entry += f" || KEY_DATA: {sig}"
            lines.append(entry)
    elif pages:
        # Balanced page-by-page fallback when section headings weren't cleanly detected
        valid_pages = []
        for pno, ptxt in pages:
            cleaned_lines = []
            stop_here = False
            for ln in ptxt.splitlines():
                if pno > 1 and REF_STOP_RE.match(ln.strip()):
                    stop_here = True
                    break
                if ln.strip():
                    cleaned_lines.append(ln.strip())
            if cleaned_lines:
                valid_pages.append((pno, " ".join(cleaned_lines)))
            if stop_here:
                break
        if valid_pages:
            per_page = max(380, min(2200, (budget_bytes - 200) // len(valid_pages)))
            for pno, ptxt in valid_pages:
                head_b = int(per_page * 0.65)
                head = _utf8_trunc(ptxt, head_b, suffix=" ...")
                sig = _extract_signal_sentences(ptxt, head, per_page - head_b)
                row = f"- [p.{pno}] {head}" + (f" || KEY_DATA: {sig}" if sig else "")
                lines.append(row)

    return _utf8_trunc("\n".join(lines), budget_bytes)


def _fmt_starting_deck(dinv: dict, draft_path: Path, max_bytes: int = 4400) -> str:
    lines = [
        "## 5. STARTING DECK SUMMARY (Editable Migration Draft Ready)",
        f"- deck: `{dinv.get('deck')}` | slides={dinv.get('slide_count')} | size={dinv.get('slide_size')} "
        f"| master_fp={dinv.get('master_fingerprint')}",
        f"- draft_plan: `{draft_path}`",
    ]
    for s in (dinv.get("slides") or []):
        title = _clean_ws(s.get("title") or "")[:90]
        n_b = sum(len(ph.get("paragraphs") or []) for ph in (s.get("placeholders") or [])
                  if "TITLE" not in str(ph.get("type", ""))) + len(s.get("floating_text") or [])
        parts = [f"s{s['n']} [{s.get('layout')}] \"{title}\"",
                 f"bullets~{n_b} tbl={len(s.get('tables') or [])} pic={len(s.get('pictures') or [])} "
                 f"chart={len(s.get('charts') or [])} notes={'Y' if s.get('notes') else 'N'}"]
        if s.get("tables"):
            tb0 = s["tables"][0].get("rows") or []
            if tb0:
                parts.append(f"tbl_hdr={tb0[0][:5]}")
        if s.get("charts"):
            ch0 = s["charts"][0]
            parts.append(f"chart={ch0.get('type')}({len(ch0.get('series') or [])}s)")
        row = "  * " + " | ".join(parts)
        if len(("\n".join(lines + [row])).encode("utf-8")) > max_bytes - 100:
            lines.append("  ... (remaining slides in deck_inventory.json / plan_draft.json)")
            break
        lines.append(row)
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="One-stop Step-1 briefing for academic slide adaptation.")
    ap.add_argument("--input", default="/input", help="Input directory containing case.json, task.md, template, paper")
    ap.add_argument("--workdir", default="/app/work", help="Working directory to write structured inventories & figs")
    ap.add_argument("--max-bytes", type=int, default=32000, help="Hard UTF-8 stdout cap (default 32000 B)")
    ap.add_argument("--case", default=None)
    ap.add_argument("--task", default=None)
    ap.add_argument("--template", default=None)
    ap.add_argument("--paper", default=None)
    ap.add_argument("--deck", default=None)
    ap.add_argument("--dpi", type=int, default=200, help="Figure crop DPI for extract_paper (default 200)")
    a = ap.parse_args()

    input_dir = Path(a.input)
    workdir = Path(a.workdir)
    workdir.mkdir(parents=True, exist_ok=True)
    max_bytes = max(8000, min(a.max_bytes, 44000))

    files = _scan_input(input_dir)
    case_path = Path(a.case) if a.case else next((f for f in files if f.name.lower() == "case.json"), None)
    case_data = {}
    if case_path and case_path.exists():
        try:
            case_data = json.loads(case_path.read_text(encoding="utf-8"))
        except Exception as e:
            case_data = {"_parse_error": str(e)}

    roles = _classify_inputs(files, case_data, a)
    saved = []

    # ---- Block 1: Input Inventory & Case Contract ----
    b1 = ["## 1. INPUT INVENTORY, CASE CONTRACT & TASK"]
    b1.append(f"- input_dir: `{input_dir}` ({len(files)} file(s))")
    for f in files[:30]:
        tag = []
        if roles["case"] and f == roles["case"]:
            tag.append("CASE")
        if roles["task"] and f == roles["task"]:
            tag.append("TASK")
        if roles["template"] and f == roles["template"]:
            tag.append("TEMPLATE")
        if roles["deck"] and f == roles["deck"]:
            tag.append("STARTING_DECK")
        if roles["paper"] and f == roles["paper"]:
            tag.append("PAPER")
        dim = ""
        if f in roles["images"] and Image is not None:
            try:
                with Image.open(f) as im:
                    dim = f" {im.width}x{im.height}px"
            except Exception:
                pass
        b1.append(f"  * `{f}` ({f.stat().st_size} B{dim})" + (f" [{','.join(tag)}]" if tag else ""))

    if case_data:
        oc = case_data.get("output_contract") or case_data.get("output_fields") or {}
        oc_keys = list(oc.keys()) if isinstance(oc, dict) else list(oc)
        b1.append(f"- case_id: `{case_data.get('case_id', case_data.get('id', 'MISSING'))}`")
        b1.append(f"- output_contract ({len(oc_keys)} keys): {json.dumps(oc_keys, ensure_ascii=False)}")
        meta_extras = {
            k: v for k, v in case_data.items()
            if k not in ("case_id", "id", "output_contract", "output_fields") and len(str(v)) < 300
        }
        if meta_extras:
            b1.append(f"- case_meta: {json.dumps(meta_extras, ensure_ascii=False)}")

    if roles["task"] and roles["task"].exists():
        t_raw = roles["task"].read_text(encoding="utf-8", errors="ignore").strip()
        if len(t_raw.encode("utf-8")) <= 4000:
            b1.append("### task.md\n" + t_raw)
        else:
            b1.append("### task.md (head + tail)\n" + _utf8_trunc(t_raw, 2800, "\n...[middle omitted]...\n")
                      + "\n" + t_raw[-900:])
    block1 = _utf8_trunc("\n".join(b1), 6200)

    # ---- Block 2: Template Inspection ----
    block2 = "## 2. TEMPLATE LAYOUT SLOTS & THEME\n- TEMPLATE: none detected (use `\"template\": \"none\"` in plan.json and record in uncertainty)"
    if roles["template"] and roles["template"].exists():
        try:
            tinv = inspect_template.inspect(roles["template"])
            t_out = workdir / "template_inventory.json"
            t_out.write_text(json.dumps(tinv, ensure_ascii=False, indent=1), encoding="utf-8")
            saved.append(str(t_out))
            block2 = _fmt_template(tinv, max_bytes=4400)
        except Exception as e:
            block2 = f"## 2. TEMPLATE LAYOUT SLOTS & THEME\n- ERROR inspecting `{roles['template']}`: {e}"

    # ---- Block 3 & 4: Paper Extraction ----
    ps_data = None
    text_out = workdir / "paper_text.txt"
    block3 = "## 3. PAPER METADATA, FIGURES, TABLES & KEY STATS\n- PAPER: none detected"
    if roles["paper"] and roles["paper"].exists():
        try:
            p_src = roles["paper"].resolve()
            p_out = workdir / "paper_structure.json"
            figdir = workdir / "figs"
            if p_src.suffix.lower() == ".pdf":
                ps_data = extract_paper.extract(p_src, figdir, a.dpi, text_out=text_out)
            else:
                ps_data = extract_paper.extract_non_pdf(p_src, text_out=text_out)
            p_out.write_text(json.dumps(ps_data, ensure_ascii=False, indent=1), encoding="utf-8")
            saved.extend([str(p_out), str(text_out), str(figdir)])
            block3 = _fmt_paper_core(ps_data, max_bytes=9800)
        except Exception as e:
            block3 = f"## 3. PAPER METADATA, FIGURES, TABLES & KEY STATS\n- ERROR extracting `{roles['paper']}`: {e}"

    # ---- Block 5: Starting Deck (if any) ----
    block5 = ""
    if roles["deck"] and roles["deck"].exists():
        try:
            d_imgdir = workdir / "deck_imgs"
            dinv = inspect_deck.inspect(roles["deck"], d_imgdir)
            d_out = workdir / "deck_inventory.json"
            dr_out = workdir / "plan_draft.json"
            d_out.write_text(json.dumps(dinv, ensure_ascii=False, indent=1), encoding="utf-8")
            dr_out.write_text(json.dumps(inspect_deck.draft(dinv), ensure_ascii=False, indent=1), encoding="utf-8")
            saved.extend([str(d_out), str(dr_out), str(d_imgdir)])
            block5 = _fmt_starting_deck(dinv, dr_out, max_bytes=4400)
        except Exception as e:
            block5 = f"## 5. STARTING DECK SUMMARY\n- ERROR inspecting `{roles['deck']}`: {e}"

    # ---- Block 6: Supplementary Materials (extra PDFs / docs) ----
    extra_lines = []
    for ep in roles["extra_pdfs"][:3]:
        try:
            ep_out = workdir / f"extra_{ep.stem}_text.txt"
            ep_ps = extract_paper.extract(ep.resolve(), workdir / f"figs_{ep.stem}", a.dpi, text_out=ep_out)
            saved.append(str(ep_out))
            extra_lines.append(
                f"- extra PDF `{ep}`: pages={ep_ps.get('pages')} figs={len(ep_ps.get('figures') or [])} "
                f"tabs={len(ep_ps.get('tables') or [])} title={_clean_ws((ep_ps.get('title') or {}).get('text') or '')[:120]}"
            )
        except Exception as e:
            extra_lines.append(f"- extra PDF `{ep}`: error {e}")
    for ed in roles["extra_docs"][:4]:
        try:
            if ed.suffix.lower() == ".docx":
                ep_ps = extract_paper.extract_non_pdf(ed.resolve(), text_out=workdir / f"extra_{ed.stem}_text.txt")
                snip = " ".join(s.get("title", "") for s in (ep_ps.get("sections") or [])[:6])
                extra_lines.append(f"- extra DOCX `{ed}`: sections={snip[:220]}")
            else:
                raw = ed.read_text(encoding="utf-8", errors="ignore")
                extra_lines.append(f"- extra `{ed}` ({len(raw)} chars): {_utf8_trunc(_clean_ws(raw), 420)}")
        except Exception as e:
            extra_lines.append(f"- extra `{ed}`: error {e}")
    block6 = ("## 6. SUPPLEMENTARY MATERIALS\n" + "\n".join(extra_lines)) if extra_lines else ""
    if block6:
        block6 = _utf8_trunc(block6, 2200)

    # ---- Allocate all remaining byte budget to Block 4 (Priority Paper Sections) ----
    fixed_parts = [block1, block2, block3] + ([block5] if block5 else []) + ([block6] if block6 else [])
    fixed_bytes = sum(len(p.encode("utf-8")) + 2 for p in fixed_parts)
    footer_reserve = 550
    sec_budget = max(1200, max_bytes - fixed_bytes - footer_reserve)
    block4 = _fmt_paper_sections(ps_data, text_out, sec_budget) if ps_data else ""

    ordered = [block1, block2, block3]
    if block4:
        ordered.append(block4)
    if block5:
        ordered.append(block5)
    if block6:
        ordered.append(block6)

    body = "\n\n".join(ordered)
    brief_path = workdir / "brief.txt"
    saved.append(str(brief_path))
    footer = (
        f"\n\n## SAVED ARTIFACTS ({len(saved)})\n"
        + "\n".join(f"- `{p}`" for p in saved)
        + f"\nBRIEF_READY bytes={len(body.encode('utf-8'))}/{max_bytes} "
          f"(next: write /app/work/plan.json and run Step 3 build+check+export+result)"
    )
    final_out = _utf8_trunc(body, max_bytes - len(footer.encode("utf-8")) - 16) + footer
    brief_path.write_text(final_out, encoding="utf-8")
    print(final_out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
