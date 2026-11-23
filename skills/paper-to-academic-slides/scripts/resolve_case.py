#!/usr/bin/env python3
"""Resolve runtime case inputs: read case.json & task.md, locate materials,
inspect template/starting deck, and extract output_contract requirements.

Usage:
  python3 resolve_case.py [--case-json /input/case.json] [--input-dir /input]
                          [--out /app/work/case_resolved.json]

Outputs a JSON with resolved paths, template inspection, output_contract, and
constraints. Never fails hard: missing pieces are recorded for honest degradation.
"""
import argparse
import json
import os
import re
import sys

PDF_KEYS = ["paper", "paper_pdf", "pdf", "source", "paper_path", "source_pdf",
            "article", "material", "material_pdf", "paper_file", "input_pdf"]
TEXT_MAT_KEYS = ["paper_md", "material_md", "source_md", "manuscript", "paper_txt",
                 "text_file", "report", "notes_file"]
PPTX_KEYS = ["template", "template_pptx", "pptx_template", "deck", "starting_deck",
             "pptx", "presentation", "template_path", "deck_path", "slide_template",
             "base_deck", "input_pptx"]
SLIDES_KEYS = ["max_slides", "slide_limit", "slide_count", "n_slides", "num_slides",
               "page_limit", "max_pages", "target_slides"]
AUD_KEYS = ["audience", "target_audience", "duration_minutes", "time_minutes",
            "minutes", "duration", "language", "style", "title", "task",
            "instructions", "notes", "requirements", "constraints", "topic"]
MATERIAL_EXTS = {".pdf", ".md", ".txt", ".csv", ".tsv", ".png", ".jpg", ".jpeg", ".svg", ".docx"}


def find_files(root, exts):
    hits = []
    if not root or not os.path.isdir(root):
        return hits
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for fn in sorted(filenames):
            if os.path.splitext(fn)[1].lower() in exts:
                hits.append(os.path.join(dirpath, fn))
    return hits


def resolve_path(val, base_dir, input_dir):
    if not isinstance(val, str) or not val.strip():
        return None
    if os.path.isabs(val) and os.path.exists(val):
        return val
    for root in (base_dir, input_dir, "/input"):
        if root:
            cand = os.path.join(root, val)
            if os.path.exists(cand):
                return os.path.abspath(cand)
    return val if os.path.isabs(val) else os.path.join(base_dir or input_dir or "/input", val)


def extract_theme_fonts(prs):
    """Read major/minor theme fonts from slide master XML if available."""
    fonts = {}
    try:
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if "theme" in str(rel.reltype).lower():
                xml = rel.target_part.blob.decode("utf-8", "ignore")
                maj = re.search(r"<a:majorFont>.*?<a:latin[^>]*typeface=\"([^\"]+)\"", xml, re.S)
                minf = re.search(r"<a:minorFont>.*?<a:latin[^>]*typeface=\"([^\"]+)\"", xml, re.S)
                ea_maj = re.search(r"<a:majorFont>.*?<a:ea[^>]*typeface=\"([^\"]+)\"", xml, re.S)
                ea_min = re.search(r"<a:minorFont>.*?<a:ea[^>]*typeface=\"([^\"]+)\"", xml, re.S)
                if maj and maj.group(1):
                    fonts["major_latin"] = maj.group(1)
                if minf and minf.group(1):
                    fonts["minor_latin"] = minf.group(1)
                if ea_maj and ea_maj.group(1):
                    fonts["major_ea"] = ea_maj.group(1)
                if ea_min and ea_min.group(1):
                    fonts["minor_ea"] = ea_min.group(1)
                break
    except Exception:
        pass
    return fonts


def inspect_pptx(path, case_raw_text=""):
    """Inspect a template or starting deck so the agent knows layouts and reuse mode."""
    info = {
        "path": os.path.abspath(path),
        "ok": False,
        "n_existing_slides": 0,
        "existing_slide_titles": [],
        "n_nonempty_slides": 0,
        "slide_size_in": None,
        "layouts": [],
        "has_footer_placeholder": False,
        "theme_fonts": {},
        "recommended_mode": "layouts",
    }
    try:
        from pptx import Presentation
        prs = Presentation(path)
        info["ok"] = True
        info["slide_size_in"] = [round(prs.slide_width / 914400.0, 3),
                                 round(prs.slide_height / 914400.0, 3)]
        info["theme_fonts"] = extract_theme_fonts(prs)
        n_slides = len(prs.slides)
        info["n_existing_slides"] = n_slides
        nonempty = 0
        titles = []
        for idx, sld in enumerate(prs.slides, 1):
            if idx > 15:
                break
            stext = []
            for shp in sld.shapes:
                if shp.has_text_frame and shp.text_frame.text.strip():
                    stext.append(shp.text_frame.text.strip().splitlines()[0][:80])
            if stext:
                nonempty += 1
            titles.append({"slide": idx, "title": stext[0] if stext else ""})
        info["existing_slide_titles"] = titles
        info["n_nonempty_slides"] = nonempty

        for li, lay in enumerate(prs.slide_layouts):
            phs = []
            for ph in lay.placeholders:
                pt = int(ph.placeholder_format.type)
                phs.append(pt)
                if pt in (13, 15, 16):
                    info["has_footer_placeholder"] = True
            info["layouts"].append({"idx": li, "name": lay.name, "ph_types": phs})

        low = (case_raw_text + " " + os.path.basename(path)).lower()
        reuse_hint = any(k in low for k in (
            "starting_deck", "starting deck", "adapt", "revise", "existing deck",
            "existing slides", "edit deck", "起始", "现有演示", "修改现有"
        ))
        if reuse_hint and n_slides >= 1:
            info["recommended_mode"] = "reuse"
        elif n_slides >= 3 and nonempty >= 2 and "template" not in os.path.basename(path).lower():
            info["recommended_mode"] = "reuse"
        else:
            info["recommended_mode"] = "layouts"
    except Exception as e:
        info["error"] = str(e)
    return info


def dump_starting_deck(path, out_md_path, figs_dir=None):
    """Dump existing slides of a starting deck / non-empty template into Markdown
    and export content pictures to figs_dir so the agent sees exact existing titles,
    bullets, tables, pictures (with reusable paths), charts, groups, and notes."""
    try:
        from pptx import Presentation
        prs = Presentation(path)
        if len(prs.slides) == 0:
            return None
        EMU_IN = 914400.0
        SW = prs.slide_width / EMU_IN
        SH = prs.slide_height / EMU_IN
        slide_area = SW * SH
        if figs_dir is None:
            figs_dir = os.path.join(os.path.dirname(os.path.abspath(out_md_path)), "starting_deck_figs")
        lines = [
            f"# Starting Deck Inspection: `{os.path.basename(path)}`",
            f"- Slides: {len(prs.slides)} | Size: {SW:.2f} x {SH:.2f} in",
        ]
        for idx, sld in enumerate(prs.slides, 1):
            lname = getattr(getattr(sld, "slide_layout", None), "name", "?")
            lines.append(f"\n## Slide {idx} (layout: {lname})")
            texts = []
            tables = []
            media_items = []
            pic_i = 0
            for shp in sld.shapes:
                try:
                    w_in = shp.width / EMU_IN
                    h_in = shp.height / EMU_IN
                    t_in = shp.top / EMU_IN
                    l_in = shp.left / EMU_IN
                except Exception:
                    w_in = h_in = t_in = l_in = 0.0
                if shp.shape_type == 13:
                    is_bg = (w_in * h_in >= 0.65 * slide_area)
                    is_corner_logo = (w_in <= 2.5 and h_in <= 1.3 and
                                      (t_in <= 0.85 or t_in + h_in >= SH - 0.85))
                    if is_bg:
                        media_items.append(f"[BackgroundImage {w_in:.1f}x{h_in:.1f}in, protected]")
                    elif is_corner_logo:
                        media_items.append(f"[CornerLogo {w_in:.1f}x{h_in:.1f}in at ({l_in:.1f},{t_in:.1f}), protected]")
                    else:
                        pic_i += 1
                        saved_path = None
                        try:
                            img = shp.image
                            ext = (img.ext or "png").lower()
                            if ext not in ("png", "jpg", "jpeg"):
                                ext = "png"
                            os.makedirs(figs_dir, exist_ok=True)
                            saved_path = os.path.abspath(
                                os.path.join(figs_dir, f"s{idx:02d}_pic{pic_i}.{ext}")
                            )
                            with open(saved_path, "wb") as imgf:
                                imgf.write(img.blob)
                        except Exception:
                            saved_path = None
                        ar = round(w_in / max(h_in, 0.05), 2)
                        if saved_path:
                            media_items.append(
                                f"[Picture {w_in:.1f}x{h_in:.1f}in ar={ar}] `{saved_path}`"
                            )
                        else:
                            media_items.append(f"[Picture {w_in:.1f}x{h_in:.1f}in ar={ar}]")
                elif getattr(shp, "has_chart", False):
                    media_items.append(
                        f"[NativeChart {w_in:.1f}x{h_in:.1f}in — set keep_existing_media:true to preserve]"
                    )
                elif shp.shape_type == 6:
                    if w_in >= 1.8 and h_in >= 1.2 and (w_in * h_in < 0.65 * slide_area):
                        media_items.append(
                            f"[GroupedDiagram {w_in:.1f}x{h_in:.1f}in — set keep_existing_media:true to preserve]"
                        )
                elif getattr(shp, "has_table", False):
                    trows = []
                    for r in shp.table.rows:
                        trows.append(" | ".join(c.text.strip().replace("\n", " ")[:40] for c in r.cells))
                    tables.append(trows[:6])
                elif shp.has_text_frame and shp.text_frame.text.strip():
                    role = ""
                    if shp.is_placeholder:
                        pt = int(shp.placeholder_format.type)
                        if shp.placeholder_format.idx == 0 or pt in (1, 3):
                            role = "[Title] "
                        elif pt in (13, 15, 16):
                            role = "[Footer] "
                    body_t = " / ".join(ln.strip() for ln in shp.text_frame.text.splitlines() if ln.strip())
                    texts.append(f"{role}{body_t[:320]}")
            for t in texts[:8]:
                lines.append(f"- {t}")
            for mi in media_items[:6]:
                lines.append(f"- Media: {mi}")
            for ti, tr in enumerate(tables, 1):
                lines.append(f"- Table {ti}: " + " ; ".join(tr[:4]))
            notes = ""
            try:
                if sld.has_notes_slide:
                    notes = sld.notes_slide.notes_text_frame.text.strip()
            except Exception:
                pass
            if notes:
                lines.append(f"- Speaker Notes: {notes[:280]}")
            else:
                lines.append("- Speaker Notes: (empty)")
        os.makedirs(os.path.dirname(os.path.abspath(out_md_path)) or ".", exist_ok=True)
        with open(out_md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        return os.path.abspath(out_md_path)
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-json", default="/input/case.json")
    ap.add_argument("--input-dir", default="/input")
    ap.add_argument("--task-md", default=None)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()

    resolved = {
        "case_json": None,
        "task_md": None,
        "task_md_excerpt": None,
        "task_md_truncated": False,
        "case_id": None,
        "paper_pdf": None,
        "paper_source": None,
        "template_pptx": None,
        "starting_deck": None,
        "starting_deck_dump_md": None,
        "template_info": None,
        "constraints": {},
        "output_contract": None,
        "expected_outputs": {},
        "materials": [],
        "extra_files": [],
        "missing": [],
        "raw_keys": [],
        "notes": [],
    }

    case_raw_text = ""
    case = {}
    cj = a.case_json
    cj_dir = os.path.dirname(cj) if cj else a.input_dir
    if cj and os.path.isfile(cj):
        resolved["case_json"] = os.path.abspath(cj)
        try:
            with open(cj, "r", encoding="utf-8") as f:
                case_raw_text = f.read()
                case = json.loads(case_raw_text)
        except Exception as e:
            resolved["missing"].append(f"case.json parse error: {e}")
            case = {}
    else:
        resolved["missing"].append(f"case.json not found at {cj}")

    # Read task.md if present alongside case.json or in input_dir
    tmd_cands = [a.task_md] if a.task_md else [
        os.path.join(cj_dir or "", "task.md"),
        os.path.join(a.input_dir, "task.md"),
        "/input/task.md",
    ]
    for tc in tmd_cands:
        if tc and os.path.isfile(tc):
            resolved["task_md"] = os.path.abspath(tc)
            try:
                tcontent = open(tc, "r", encoding="utf-8", errors="ignore").read()
                resolved["task_md_excerpt"] = tcontent[:10000]
                resolved["task_md_truncated"] = len(tcontent) > 10000
                case_raw_text += "\n" + tcontent[:10000]
            except Exception:
                pass
            break

    if isinstance(case, dict) and case:
        # Preserve top-level output_contract / deliverables schema verbatim
        for ok_key in ("output_contract", "output_schema", "required_outputs",
                       "deliverables_contract", "outputs"):
            if ok_key in case:
                resolved["output_contract"] = case[ok_key]
                break

        flat = {}

        def walk(obj, prefix=""):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    walk(v, f"{prefix}.{k}" if prefix else k)
            elif isinstance(obj, list):
                for i, v in enumerate(obj):
                    walk(v, f"{prefix}[{i}]")
            else:
                flat[prefix] = obj

        walk(case)
        resolved["raw_keys"] = sorted(flat.keys())[:200]
        low = {k.lower(): (k, v) for k, v in flat.items()}

        def grab(keys):
            for key in keys:
                if key in low:
                    return low[key][1]
            for lk, (_orig, v) in sorted(low.items()):
                for key in keys:
                    if lk == key or lk.endswith("." + key) or lk.endswith("_" + key):
                        return v
            return None

        resolved["case_id"] = grab(["case_id", "id", "caseid"])
        for key in PDF_KEYS:
            v = grab([key])
            if isinstance(v, str) and v.lower().endswith(".pdf"):
                resolved["paper_pdf"] = resolve_path(v, cj_dir, a.input_dir)
                break
        if not resolved["paper_pdf"]:
            for key in TEXT_MAT_KEYS:
                v = grab([key])
                if isinstance(v, str) and v.lower().endswith((".md", ".txt", ".docx")):
                    resolved["paper_source"] = resolve_path(v, cj_dir, a.input_dir)
                    break

        # Distinguish starting_deck and template_pptx if both exist
        for k_name in ("starting_deck", "base_deck", "deck_path", "input_pptx"):
            v = grab([k_name])
            if isinstance(v, str) and v.lower().endswith((".pptx", ".ppt")):
                resolved["starting_deck"] = resolve_path(v, cj_dir, a.input_dir)
                break
        for key in PPTX_KEYS:
            v = grab([key])
            if isinstance(v, str) and v.lower().endswith((".pptx", ".ppt")):
                resolved["template_pptx"] = resolve_path(v, cj_dir, a.input_dir)
                break
        if resolved["starting_deck"] and not resolved["template_pptx"]:
            resolved["template_pptx"] = resolved["starting_deck"]

        for key in SLIDES_KEYS:
            v = grab([key])
            if isinstance(v, (int, float)) and v:
                resolved["constraints"]["max_slides"] = int(v)
                break
            if isinstance(v, str):
                m = re.search(r"\d+", v)
                if m:
                    resolved["constraints"]["max_slides"] = int(m.group(0))
                    break
        for key in AUD_KEYS:
            v = grab([key])
            if isinstance(v, (str, int, float)) and len(str(v)) < 500:
                resolved["constraints"][key] = v

        # Capture any explicit expected output filenames (broad contract coverage, incl. nested lists)
        for k, v in flat.items():
            kl = k.lower()
            if isinstance(v, str) and not v.startswith("/input"):
                if any(x in kl for x in ("output", "deliverable", "artifact", "expected",
                                          "pptx_path", "pdf_path", "deck_file", "slides_file",
                                          "notes_file", "evidence_file", "required_files", "files")):
                    if any(v.lower().endswith(ext) for ext in (".pptx", ".pdf", ".json", ".md")):
                        resolved["expected_outputs"][k] = v
        # Also scan output_contract values directly for filenames
        oc = resolved.get("output_contract")
        if oc is not None:
            def _scan_oc(obj, pref="output_contract"):
                if isinstance(obj, dict):
                    for kk, vv in obj.items():
                        _scan_oc(vv, f"{pref}.{kk}")
                elif isinstance(obj, list):
                    for ii, vv in enumerate(obj):
                        _scan_oc(vv, f"{pref}[{ii}]")
                elif isinstance(obj, str) and not obj.startswith("/input"):
                    if any(obj.lower().endswith(ext) for ext in (".pptx", ".pdf", ".json", ".md")):
                        resolved["expected_outputs"].setdefault(pref, obj)
            _scan_oc(oc)

    # Extract fallback constraints from task.md when absent in case.json
    if resolved["task_md_excerpt"]:
        tmd_txt = resolved["task_md_excerpt"]
        if "max_slides" not in resolved["constraints"]:
            ms = re.search(
                r"(?:max(?:imum)?\s*(?:of\s*)?(\d{1,2})\s*slides?|(\d{1,2})\s*slides?\s*(?:max|limit|cap)|"
                r"(?:不超过|最多|上限|控制在)\s*(\d{1,2})\s*[页张])",
                tmd_txt, re.I,
            )
            if ms:
                val = next((int(g) for g in ms.groups() if g), None)
                if val and 3 <= val <= 40:
                    resolved["constraints"]["max_slides"] = val
        if "duration_minutes" not in resolved["constraints"] and "duration" not in resolved["constraints"]:
            mdur = re.search(r"(\d{1,2})\s*(?:min(?:ute)?s?|分钟)", tmd_txt, re.I)
            if mdur:
                resolved["constraints"]["duration_minutes"] = int(mdur.group(1))

    # Fallback discovery in input dir
    pdfs = find_files(a.input_dir, {".pdf"})
    pptxs = find_files(a.input_dir, {".pptx", ".ppt"})
    mats = [p for p in find_files(a.input_dir, MATERIAL_EXTS)
            if os.path.basename(p).lower() not in ("case.json", "task.md")]
    resolved["materials"] = mats

    if not resolved["paper_pdf"] and pdfs:
        resolved["paper_pdf"] = pdfs[0]
    if pptxs:
        resolved["extra_files"] = pptxs
        if len(pptxs) >= 2 and (not resolved["template_pptx"] or not resolved["starting_deck"]):
            # Smartly classify multiple PPTX files into template vs starting_deck
            tpl_cands, start_cands = [], []
            for pp in pptxs:
                bn = os.path.basename(pp).lower()
                is_tpl_name = any(w in bn for w in ("template", "master", "theme", "模板", "母版"))
                is_start_name = any(w in bn for w in ("start", "draft", "existing", "base", "initial", "deck", "起始", "初稿"))
                nsld = 0
                try:
                    from pptx import Presentation
                    nsld = len(Presentation(pp).slides)
                except Exception:
                    pass
                if is_tpl_name and not is_start_name:
                    tpl_cands.append((nsld, pp))
                elif is_start_name and not is_tpl_name:
                    start_cands.append((-nsld, pp))
                elif nsld <= 2:
                    tpl_cands.append((nsld, pp))
                else:
                    start_cands.append((-nsld, pp))
            tpl_cands.sort()
            start_cands.sort()
            if not resolved["template_pptx"] and tpl_cands:
                resolved["template_pptx"] = tpl_cands[0][1]
            if not resolved["starting_deck"] and start_cands:
                resolved["starting_deck"] = start_cands[0][1]
        if not resolved["template_pptx"]:
            resolved["template_pptx"] = pptxs[0]

    if resolved["paper_pdf"] and not os.path.isfile(resolved["paper_pdf"]):
        resolved["missing"].append(f"paper path does not exist: {resolved['paper_pdf']}")
        resolved["paper_pdf"] = None
    if not resolved["paper_source"]:
        if resolved["paper_pdf"]:
            resolved["paper_source"] = resolved["paper_pdf"]
        else:
            md_txt = [p for p in mats if os.path.splitext(p)[1].lower() in (".md", ".txt", ".docx")]
            if md_txt:
                resolved["paper_source"] = md_txt[0]
    if not resolved["paper_source"]:
        resolved["missing"].append("paper PDF or text material not located")

    if resolved["template_pptx"] and os.path.isfile(resolved["template_pptx"]):
        tinfo = inspect_pptx(resolved["template_pptx"], case_raw_text)
        resolved["template_info"] = tinfo
        if resolved["starting_deck"] and os.path.abspath(resolved["starting_deck"]) == os.path.abspath(resolved["template_pptx"]):
            tinfo["recommended_mode"] = "reuse"
    elif resolved["template_pptx"]:
        resolved["missing"].append(f"template path does not exist: {resolved['template_pptx']}")
        resolved["template_pptx"] = None
    if resolved["starting_deck"] and os.path.isfile(resolved["starting_deck"]):
        if not resolved["template_pptx"] or os.path.abspath(resolved["starting_deck"]) != os.path.abspath(resolved["template_pptx"]):
            resolved["starting_deck_info"] = inspect_pptx(resolved["starting_deck"], case_raw_text)
            resolved["notes"].append("both template_pptx and starting_deck present; inspect starting_deck_info if task requires editing existing slides")

    out = a.out or os.path.join(os.environ.get("PPT_WORK", "/app/work"), "case_resolved.json")
    work_dir = os.path.dirname(os.path.abspath(out)) or "."
    # Dump existing slide contents whenever starting_deck or non-empty pptx is present
    dump_target = None
    if resolved["starting_deck"] and os.path.isfile(resolved["starting_deck"]):
        dump_target = resolved["starting_deck"]
    elif (resolved["template_pptx"] and os.path.isfile(resolved["template_pptx"]) and
          (resolved.get("template_info") or {}).get("n_existing_slides", 0) > 0):
        dump_target = resolved["template_pptx"]
    if dump_target:
        resolved["starting_deck_dump_md"] = dump_starting_deck(
            dump_target, os.path.join(work_dir, "starting_deck_dump.md")
        )

    os.makedirs(work_dir, exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(resolved, f, ensure_ascii=False, indent=1)
    print(json.dumps(resolved, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
