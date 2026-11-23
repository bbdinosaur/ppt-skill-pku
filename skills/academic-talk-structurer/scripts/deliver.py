#!/usr/bin/env python3
"""One-click S3+S4+S5+S6 pipeline: build PPTX, run 4-gate preflight, render PDF/PNG + contact sheets,
export speaker notes, and assemble contract-aligned result.json.

Usage:
  python3 deliver.py --case /input/case.json --work /app/work --out-dir /app/output \
      [--spec /app/work/deck_spec.json] [--out <deck.pptx>] [--template <tpl.pptx>] [--deck <start.pptx>] \
      [--dpi 100] [--skip-render] [--notes-pdf]

Exit: 0 build + all preflight gates clean; 1 build failed or preflight issue found (deliverables + result.json still written).
No-script fallback: run build_deck.py, preflight.sh, render_preview.sh, export_notes.py, build_result.py manually.
"""
import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PY = sys.executable


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


def _coerce_int(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    m = re.search(r"-?\d+", str(v))
    return int(m.group(0)) if m else None


def normalize_spec(spec: dict):
    """Lossless schema resilience for minor LLM formatting variance (e.g. string evidence/figure/time_s)."""
    if not isinstance(spec, dict):
        return spec, []
    fixes = []
    meta = spec.get("meta")
    if not isinstance(meta, dict):
        meta = {}
        spec["meta"] = meta
    if isinstance(meta.get("tradeoffs"), str):
        meta["tradeoffs"] = [x.strip() for x in re.split(r"[;\n；]+", meta["tradeoffs"]) if x.strip()]
        fixes.append("meta.tradeoffs coerced str->list")
    if not meta.get("tradeoffs"):
        for tk in ("editorial_decisions", "editorial_tradeoffs", "decisions", "cuts_and_tradeoffs", "design_tradeoffs"):
            cand_t = meta.get(tk) or (spec.get("result_fill") or {}).get(tk)
            if cand_t:
                meta["tradeoffs"] = cand_t if isinstance(cand_t, list) else [x.strip() for x in re.split(r"[;\n；]+", str(cand_t)) if x.strip()]
                fixes.append(f"meta.tradeoffs<-{tk}")
                break

    for idx, sl in enumerate(spec.get("slides") or [], 1):
        if not isinstance(sl, dict):
            continue
        # bullets: str -> list
        if isinstance(sl.get("bullets"), str):
            sl["bullets"] = [x.strip(" -•·") for x in sl["bullets"].splitlines() if x.strip()]
            fixes.append(f"s{idx}.bullets str->list")
        # time_s: "60s" -> 60; qa_backup always excluded from main talk timing (time_s=None)
        if str(sl.get("role") or "").lower() == "qa_backup" and sl.get("time_s") is not None:
            sl["time_s"] = None
            fixes.append(f"s{idx}.qa_backup.time_s->null")
        elif isinstance(sl.get("time_s"), str):
            iv = _coerce_int(sl["time_s"])
            sl["time_s"] = iv
            fixes.append(f"s{idx}.time_s str->{iv}")
        # figure: "path.png" -> {"path": "path.png"}; coerce file/src/image/img -> path
        if isinstance(sl.get("figure"), str):
            sl["figure"] = {"path": sl["figure"]}
            fixes.append(f"s{idx}.figure str->dict")
        elif isinstance(sl.get("figure"), dict) and not sl["figure"].get("path"):
            for fk in ("file", "src", "image", "img", "fig_path"):
                if sl["figure"].get(fk):
                    sl["figure"]["path"] = sl["figure"][fk]
                    fixes.append(f"s{idx}.figure.{fk}->path")
                    break
        # reuse: int -> {"slide_index": N, "action": "edit"}; coerce slide/index -> slide_index
        ru = sl.get("reuse")
        if isinstance(ru, int) and not isinstance(ru, bool):
            sl["reuse"] = {"slide_index": ru, "action": "edit"}
            fixes.append(f"s{idx}.reuse int->dict")
        elif isinstance(ru, dict) and "slide_index" not in ru:
            for rk in ("index", "slide", "orig_index", "orig_slide_index"):
                if rk in ru and ru[rk] is not None:
                    ru["slide_index"] = _coerce_int(ru[rk])
                    ru.setdefault("action", "edit")
                    fixes.append(f"s{idx}.reuse.{rk}->slide_index")
                    break
        # table: 2D list or list[dict] -> {"headers": ..., "rows": ...}; string highlight_row -> int index
        if isinstance(sl.get("table"), list) and sl["table"]:
            if isinstance(sl["table"][0], list):
                sl["table"] = {"headers": sl["table"][0], "rows": sl["table"][1:]}
                fixes.append(f"s{idx}.table 2d-list->dict")
            elif isinstance(sl["table"][0], dict):
                hdrs = list(sl["table"][0].keys())
                sl["table"] = {"headers": hdrs, "rows": [[str(r.get(h, "")) for h in hdrs] for r in sl["table"] if isinstance(r, dict)]}
                fixes.append(f"s{idx}.table dict-list->dict")
        if isinstance(sl.get("table"), dict):
            tb_d = sl["table"]
            if isinstance(tb_d.get("rows"), list) and tb_d["rows"] and isinstance(tb_d["rows"][0], dict):
                hdrs = tb_d.get("headers") or list(tb_d["rows"][0].keys())
                tb_d["headers"] = hdrs
                tb_d["rows"] = [[str(r.get(h, "")) for h in hdrs] for r in tb_d["rows"] if isinstance(r, dict)]
                fixes.append(f"s{idx}.table.rows dict->list")
            if isinstance(tb_d.get("highlight_row"), str):
                hl_q = tb_d["highlight_row"].strip().lower()
                found_hl = next((ri for ri, r in enumerate(tb_d.get("rows") or []) if r and hl_q in str(r[0]).lower()), None)
                tb_d["highlight_row"] = found_hl
                fixes.append(f"s{idx}.table.highlight_row str->{found_hl}")
        # kpis: coerce alt keys (val/number -> value; name/metric/title -> label; change/diff/gain -> delta)
        if isinstance(sl.get("kpis"), list):
            k_chg = False
            for kp in sl["kpis"]:
                if not isinstance(kp, dict):
                    continue
                if "value" not in kp:
                    for ak in ("val", "number", "metric_value", "score"):
                        if ak in kp and kp[ak] not in (None, ""):
                            kp["value"] = str(kp[ak])
                            k_chg = True
                            break
                if "label" not in kp:
                    for lk in ("name", "metric", "title", "desc"):
                        if lk in kp and kp[lk] not in (None, ""):
                            kp["label"] = str(kp[lk])
                            k_chg = True
                            break
                if "delta" not in kp:
                    for dk in ("change", "diff", "gain", "improvement", "sub"):
                        if dk in kp and kp[dk] not in (None, ""):
                            kp["delta"] = str(kp[dk])
                            k_chg = True
                            break
            if k_chg:
                fixes.append(f"s{idx}.kpis alt-keys->value/label/delta")
        # two_col: coerce nested {"left":{"title","bullets"},"right":{...}} -> flat left_title/left_bullets/right_title/right_bullets
        tc = sl.get("two_col")
        if isinstance(tc, dict) and ("left" in tc or "right" in tc) and "left_bullets" not in tc:
            new_tc = dict(tc)
            for side in ("left", "right"):
                sub = tc.get(side)
                if isinstance(sub, dict):
                    if sub.get("title") and not new_tc.get(f"{side}_title"):
                        new_tc[f"{side}_title"] = sub["title"]
                    bs = sub.get("bullets") or sub.get("items") or sub.get("points")
                    if isinstance(bs, list) and not new_tc.get(f"{side}_bullets"):
                        new_tc[f"{side}_bullets"] = bs
                elif isinstance(sub, list) and not new_tc.get(f"{side}_bullets"):
                    new_tc[f"{side}_bullets"] = sub
            sl["two_col"] = new_tc
            fixes.append(f"s{idx}.two_col nested->flat")
        # chart: coerce Chart.js (labels/datasets) or simple (x/y) schemas -> categories/series
        ch = sl.get("chart")
        if isinstance(ch, dict):
            ch_chg = False
            if not ch.get("categories"):
                for ck in ("labels", "x", "x_labels"):
                    if isinstance(ch.get(ck), list) and ch[ck]:
                        ch["categories"] = ch[ck]
                        ch_chg = True
                        break
            if not ch.get("series"):
                if isinstance(ch.get("datasets"), list) and ch["datasets"]:
                    ds_out = []
                    for ds in ch["datasets"]:
                        if isinstance(ds, dict):
                            ds_out.append({
                                "name": ds.get("name") or ds.get("label") or "Series",
                                "values": ds.get("values") or ds.get("data") or [],
                            })
                    if ds_out:
                        ch["series"] = ds_out
                        ch_chg = True
                elif isinstance(ch.get("y"), list) and ch["y"]:
                    ch["series"] = [{"name": ch.get("y_label") or ch.get("title") or "Value", "values": ch["y"]}]
                    ch_chg = True
            if ch_chg:
                fixes.append(f"s{idx}.chart alt-schema->categories/series")
        # steps: list of str -> list of {title, detail}; coerce alt keys (name/heading -> title; desc/description/text -> detail)
        if isinstance(sl.get("steps"), list):
            new_steps = []
            changed_st = False
            for st in sl["steps"]:
                if isinstance(st, str):
                    changed_st = True
                    m = re.match(r"^([^:：]{2,28})[:：]\s*(.+)$", st.strip())
                    new_steps.append({"title": m.group(1).strip(), "detail": m.group(2).strip()} if m else {"title": st.strip(), "detail": ""})
                elif isinstance(st, dict):
                    st_c = dict(st)
                    if not st_c.get("title"):
                        for tk in ("name", "heading", "step"):
                            if st_c.get(tk):
                                st_c["title"] = str(st_c[tk])
                                changed_st = True
                                break
                    if not st_c.get("detail"):
                        for dk in ("desc", "description", "text", "body"):
                            if st_c.get(dk):
                                st_c["detail"] = str(st_c[dk])
                                changed_st = True
                                break
                    new_steps.append(st_c)
            if changed_st:
                sl["steps"] = new_steps
                fixes.append(f"s{idx}.steps normalized")
        # evidence: str or list[str] -> list[{type, locator, strength}] + default strength
        role_l = str(sl.get("role") or "").lower()
        def_str = "NA" if role_l in ("title", "thanks", "roadmap", "qa_backup") else "S2"
        evs = sl.get("evidence")
        if isinstance(evs, str):
            evs = [evs]
        if isinstance(evs, list):
            new_evs = []
            changed_ev = False
            for ev in evs:
                if isinstance(ev, str) and ev.strip():
                    changed_ev = True
                    el = ev.strip()
                    et = "figure" if re.search(r"fig|图", el, re.I) else ("table" if re.search(r"tab|表", el, re.I) else "section")
                    new_evs.append({"type": et, "locator": el, "strength": def_str})
                elif isinstance(ev, dict):
                    if ev.get("locator") and not ev.get("strength"):
                        ev["strength"] = def_str
                        changed_ev = True
                    new_evs.append(ev)
            if changed_ev or isinstance(sl.get("evidence"), str):
                sl["evidence"] = new_evs
                fixes.append(f"s{idx}.evidence normalized")

    qa = None
    for qk in ("qa_prep", "qa_plan", "qa_list", "qa", "qa_preparation"):
        cand_q = spec.get(qk)
        if cand_q is None and isinstance(spec.get("result_fill"), dict):
            cand_q = spec["result_fill"].get(qk)
        if isinstance(cand_q, dict):
            cand_q = cand_q.get("items") or cand_q.get("questions") or cand_q.get("qa_prep")
        if isinstance(cand_q, list) and cand_q:
            qa = cand_q
            if qk != "qa_prep" or not isinstance(spec.get("qa_prep"), list):
                spec["qa_prep"] = qa
                fixes.append(f"qa_prep<-{qk}")
            break
    if isinstance(qa, list):
        for q in qa:
            if not isinstance(q, dict):
                continue
            if not q.get("category"):
                for ck in ("cat", "type", "dimension"):
                    if q.get(ck):
                        q["category"] = str(q[ck]).strip().lower()
                        break
            if not q.get("question") and q.get("q"):
                q["question"] = q["q"]
            if not q.get("answer_key"):
                for alt_k in ("answer", "key", "response", "a"):
                    if q.get(alt_k):
                        q["answer_key"] = q[alt_k]
                        break
            for rk in ("cite_slide", "backup_slide"):
                if isinstance(q.get(rk), str):
                    q[rk] = _coerce_int(q[rk])
    return spec, fixes


def sync_inline_files(spec_path: Path, work: Path):
    """If deck_spec.json embeds qa_prep/result_fill/number_allow/budget_justify, sync to <work>/*.json
    when the standalone file is missing or older than deck_spec.json."""
    try:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
    except Exception:
        return {}, spec_path
    spec, fixes = normalize_spec(spec)
    eff_spec_path = spec_path
    if fixes:
        eff_spec_path = work / "deck_spec.norm.json"
        eff_spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"  [normalize] applied {len(fixes)} schema fixup(s): {', '.join(fixes[:6])}")
    spec_mtime = spec_path.stat().st_mtime

    def maybe_write(fname: str, payload):
        if payload is None:
            return
        dst = work / fname
        if not dst.exists() or spec_mtime >= dst.stat().st_mtime - 0.5:
            dst.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    qa = spec.get("qa_prep") if isinstance(spec.get("qa_prep"), list) else spec.get("qa_plan")
    if qa:
        maybe_write("qa_plan.json", qa)

    fill = dict(spec.get("result_fill")) if isinstance(spec.get("result_fill"), dict) else {}
    for k in ("uncertainty", "limits", "changes", "report_goal", "audience", "tradeoffs"):
        if k in spec and k not in fill:
            fill[k] = spec[k]
    if fill:
        maybe_write("result_fill.json", fill)

    if isinstance(spec.get("number_allow"), list) and spec["number_allow"]:
        maybe_write("number_allow.json", spec["number_allow"])
    if isinstance(spec.get("budget_justify"), dict) and spec["budget_justify"]:
        maybe_write("budget_justify.json", spec["budget_justify"])
    return spec, eff_spec_path


def sort_preview_key(p: Path):
    m = re.search(r"(\d+)$", p.stem)
    return (int(m.group(1)) if m else 9999, p.name)


def build_contact_sheets(previews: list, work: Path, cols: int = 2, rows_per_sheet: int = 3, tile_w: int = 620):
    """Combine individual slide preview PNGs into compact 2x3 contact sheets under <work>/ for fast visual audit."""
    for old in work.glob("contact_sheet_*.png"):
        try:
            old.unlink()
        except Exception:
            pass
    if not previews:
        return []
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return []
    per_sheet = cols * rows_per_sheet
    sheets = []
    for idx in range(0, len(previews), per_sheet):
        batch = previews[idx : idx + per_sheet]
        imgs = [Image.open(p).convert("RGB") for p in batch]
        w0, h0 = imgs[0].size
        th = max(120, int(tile_w * h0 / max(1, w0)))
        n_rows = (len(imgs) + cols - 1) // cols
        pad = 18
        header = 20
        sheet = Image.new("RGB", (cols * tile_w + (cols + 1) * pad, n_rows * (th + header) + (n_rows + 1) * pad), (226, 232, 240))
        draw = ImageDraw.Draw(sheet)
        for j, im in enumerate(imgs):
            r, c = divmod(j, cols)
            x = pad + c * (tile_w + pad)
            y = pad + r * (th + header + pad)
            slide_no = idx + j + 1
            draw.rectangle([x, y, x + 76, y + header - 2], fill=(30, 41, 59))
            draw.text((x + 8, y + 3), f"Slide {slide_no:02d}", fill=(248, 250, 252))
            resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1)
            sheet.paste(im.resize((tile_w, th), resample), (x, y + header))
        out_p = work / f"contact_sheet_{idx // per_sheet + 1}.png"
        sheet.save(out_p)
        sheets.append(str(out_p))
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="/input/case.json")
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--out-dir", default="/app/output")
    ap.add_argument("--spec")
    ap.add_argument("--out", help="explicit output pptx path (overrides auto-detection)")
    ap.add_argument("--template", help="explicit template pptx")
    ap.add_argument("--deck", help="explicit starting deck pptx")
    ap.add_argument("--result", help="explicit output result.json path")
    ap.add_argument("--dpi", type=int, default=100)
    ap.add_argument("--skip-render", action="store_true")
    ap.add_argument("--notes-pdf", action="store_true")
    a = ap.parse_args()

    work = Path(a.work)
    out_dir = Path(a.out_dir)
    work.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    cpath = Path(a.case)
    spec_path = Path(a.spec) if a.spec else (work / "deck_spec.json")
    if not spec_path.exists():
        print(f"FATAL: deck_spec.json missing at {spec_path}", file=sys.stderr)
        sys.exit(2)

    ingest = {}
    if (work / "ingest_summary.json").exists():
        try:
            ingest = json.loads((work / "ingest_summary.json").read_text(encoding="utf-8"))
        except Exception:
            ingest = {}
    inspect = {}
    if (work / "inspect.json").exists():
        try:
            inspect = json.loads((work / "inspect.json").read_text(encoding="utf-8"))
        except Exception:
            inspect = {}

    spec, spec_path = sync_inline_files(spec_path, work)
    meta = spec.get("meta") if isinstance(spec.get("meta"), dict) else {}

    # Resolve output deck path & result path
    deck_name = ingest.get("recommended_deck_name") or "talk.pptx"
    result_name = ingest.get("recommended_result_name") or "result.json"
    deck_out = Path(a.out) if a.out else (out_dir / deck_name)
    result_out = Path(a.result) if a.result else (out_dir / result_name)

    # Resolve starting_deck vs template
    deck_arg = a.deck or meta.get("starting_deck")
    tpl_arg = a.template or meta.get("template")
    if not deck_arg and not tpl_arg:
        mode_hint = str(meta.get("build_mode") or "").lower()
        if mode_hint == "from_template" and ingest.get("template"):
            tpl_arg = ingest["template"]
        elif ingest.get("starting_deck") and mode_hint != "from_template":
            deck_arg = ingest["starting_deck"]
        elif ingest.get("template"):
            tpl_arg = ingest["template"]

    # 1) Build PPTX
    b_args = ["--spec", str(spec_path), "--out", str(deck_out), "--work", str(work)]
    if deck_arg and Path(deck_arg).exists():
        b_args += ["--deck", str(deck_arg)]
    elif tpl_arg and Path(tpl_arg).exists():
        b_args += ["--template", str(tpl_arg)]

    rb = run_py("build_deck.py", b_args)
    print("== [1/5] build_deck ==")
    print((rb.stdout or "").strip())
    if rb.returncode != 0:
        print("BUILD FAILED:", (rb.stderr or "").strip(), file=sys.stderr)
        # Still emit honest partial result.json
        r_args = ["--case", str(cpath), "--work", str(work), "--out", str(result_out)]
        if (work / "result_fill.json").exists():
            r_args += ["--fill", str(work / "result_fill.json")]
        run_py("build_result.py", r_args)
        sys.exit(1)

    # 2) Four-gate preflight
    print("== [2/5] preflight gates ==")
    gate_fail = 0
    fix_hints = []
    # 2a layout + notes/timing
    rl = run_py("check_layout.py", [str(deck_out), "--json", str(work / "lint_report.json"), "--plan", str(work / "slide_plan.json")])
    print("  [layout]", (rl.stdout or rl.stderr).strip().replace("\n", "\n    "))
    if rl.returncode != 0:
        gate_fail += 1
        fix_hints.append("layout: trim overflowing bullets/notes or adjust figure/table region on flagged slide(s)")

    # 2b number traceability (+ advisory Q&A numeric caliber check)
    src_txt = work / "paper.txt"
    if src_txt.exists():
        n_args = ["--spec", str(spec_path), "--source", str(src_txt), "--out", str(work / "number_check.json")]
        if (work / "paper_tables.json").exists():
            n_args += ["--tables", str(work / "paper_tables.json")]
        if (work / "qa_plan.json").exists():
            n_args += ["--qa", str(work / "qa_plan.json")]
        if (work / "number_allow.json").exists():
            n_args += ["--allow", str(work / "number_allow.json")]
        rn = run_py("check_numbers.py", n_args)
        print("  [numbers]", (rn.stdout or rn.stderr).strip().replace("\n", "\n    "))
        if rn.returncode != 0:
            gate_fail += 1
            fix_hints.append("numbers: align flagged values with paper.txt or document derived math in top-level number_allow")

        # 2c evidence locators
        e_args = ["--spec", str(spec_path), "--source", str(src_txt), "--out", str(work / "evidence_check.json")]
        if (work / "figs" / "manifest.json").exists():
            e_args += ["--manifest", str(work / "figs" / "manifest.json")]
        if (work / "paper_tables.json").exists():
            e_args += ["--tables", str(work / "paper_tables.json")]
        sp_pages = (inspect.get("source_pdf") or {}).get("pages")
        if sp_pages:
            e_args += ["--pages", str(sp_pages)]
        re_ = run_py("check_evidence.py", e_args)
        print("  [evidence]", (re_.stdout or re_.stderr).strip().replace("\n", "\n    "))
        if re_.returncode != 0:
            gate_fail += 1
            fix_hints.append("evidence: fix dangling Fig/Table/Section/page/Slide locator against reading_map.md")
    else:
        print("  [numbers/evidence] skipped (paper.txt absent)")

    # 2d figure paths exist
    bad_figs = []
    for idx, sl in enumerate(spec.get("slides", []), 1):
        fp = (sl.get("figure") or {}).get("path")
        if fp and not Path(fp).exists():
            alts = [
                work / fp, work / "figs" / Path(fp).name, work / Path(fp).name,
                cpath.parent / fp, work.parent / fp, work.parent.parent / fp,
            ]
            if not any(alt.exists() for alt in alts):
                bad_figs.append((idx, fp))
    if bad_figs:
        gate_fail += 1
        fix_hints.append("figs: replace missing figure path(s) with exact path from reading_map.md or use table/chart/steps")
        for idx, fp in bad_figs:
            print(f"  [figs] ! [s{idx}] missing figure: {fp}")
    else:
        print("  [figs] ok")

    # 2e page budget (check actual built slide_plan.json when available so auto-appended qa_backup is counted accurately)
    if cpath.exists():
        p_args = ["check", "--case", str(cpath), "--spec", str(spec_path), "--out", str(work / "budget_check.json")]
        if (work / "slide_plan.json").exists():
            p_args += ["--plan", str(work / "slide_plan.json")]
        if (work / "budget_justify.json").exists():
            p_args += ["--justif", str(work / "budget_justify.json")]
        rp = run_py("plan_pages.py", p_args)
        print("  [budget]", (rp.stdout or rp.stderr).strip().replace("\n", "\n    "))
        if rp.returncode != 0:
            gate_fail += 1
            fix_hints.append("budget: match requested slide_count/qa_backup or justify constraint conflict in meta.tradeoffs / budget_justify")

    # 3) Render PDF + PNG previews + contact sheets
    sheets = []
    previews = []
    print("== [3/5] render preview ==")
    if a.skip_render:
        print("  skipped (--skip-render)")
    else:
        for old in out_dir.glob("preview*.png"):
            try:
                old.unlink()
            except Exception:
                pass
        try:
            rr = subprocess.run(
                ["sh", str(HERE / "render_preview.sh"), str(deck_out), str(out_dir), str(a.dpi)],
                capture_output=True, text=True, timeout=300,
            )
        except subprocess.TimeoutExpired:
            rr = subprocess.CompletedProcess(
                args=["render_preview.sh"], returncode=124, stdout="", stderr="render_preview.sh timed out after 300s"
            )
        non_emb = []
        for ln in (rr.stdout + "\n" + rr.stderr).splitlines():
            if "fonts not embedded" in ln:
                raw_f = ln.split("fonts not embedded in PDF:", 1)[-1].split("(", 1)[0]
                non_emb = [f.strip() for f in raw_f.split(",") if f.strip()]
        if rr.returncode != 0:
            print("  WARN render_preview failed:", (rr.stderr or rr.stdout).strip()[-240:])
            (work / "render_meta.json").write_text(json.dumps({
                "ok": False, "dpi": a.dpi, "error": (rr.stderr or rr.stdout).strip()[-240:],
                "non_embedded_fonts": non_emb, "contact_sheets": [],
            }, ensure_ascii=False, indent=1), encoding="utf-8")
        else:
            previews = sorted(out_dir.glob("preview*.png"), key=sort_preview_key)
            sheets = build_contact_sheets(previews, work)
            (work / "render_meta.json").write_text(json.dumps({
                "ok": True, "dpi": a.dpi, "non_embedded_fonts": non_emb, "contact_sheets": sheets,
            }, ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"  PDF: {deck_out.with_suffix('.pdf')} | previews: {len(previews)} PNGs in {out_dir}")
            if non_emb:
                print(f"  WARN non-embedded fonts in PDF: {', '.join(non_emb[:8])}")
            if sheets:
                print("  contact sheets (inspect these first): " + ", ".join(sheets))

    # 4) Export speaker notes handout
    print("== [4/5] speaker notes ==")
    en_args = ["--deck", str(deck_out), "--plan", str(work / "slide_plan.json"), "--out", str(out_dir / "speaker_notes.md")]
    if (work / "qa_plan.json").exists():
        en_args += ["--qa", str(work / "qa_plan.json")]
    if (work / "paper_outline.json").exists():
        en_args += ["--outline", str(work / "paper_outline.json")]
    if a.notes_pdf:
        en_args += ["--pdf"]
    ren = run_py("export_notes.py", en_args)
    print("  " + (ren.stdout or ren.stderr).strip().replace("\n", "\n  "))

    # 4b) Deterministic rendered-pixel + semantic quality audit & real revision provenance (changes.json)
    raud = run_py("audit_talk.py", ["--work", str(work), "--out-dir", str(out_dir), "--spec", str(spec_path), "--out", str(work / "quality_audit.json")])
    audit_out = (raud.stdout or "").strip()
    try:
        cur_plan = json.loads((work / "slide_plan.json").read_text(encoding="utf-8")) if (work / "slide_plan.json").exists() else {}
        cur_slides = cur_plan.get("slides") or []
        cur_sig = [
            (s.get("pptx_slide"), s.get("role"), s.get("title"), s.get("claim"), s.get("visuals"), s.get("notes_chars"))
            for s in cur_slides
        ]
        prev_p = work / "_prev_plan_sig.json"
        chg_p = work / "changes.json"
        chg_list = []
        if chg_p.exists():
            try:
                loaded_c = json.loads(chg_p.read_text(encoding="utf-8"))
                if isinstance(loaded_c, list):
                    chg_list = loaded_c
            except Exception:
                chg_list = []
        prev_sig = None
        if prev_p.exists():
            try:
                prev_sig = json.loads(prev_p.read_text(encoding="utf-8"))
            except Exception:
                prev_sig = None
        qa_aud = json.loads((work / "quality_audit.json").read_text(encoding="utf-8")) if (work / "quality_audit.json").exists() else {}
        vis_d = (qa_aud.get("editorial_audit") or {}).get("visual_diversity_pct", "?")
        mode_tag = (cur_plan.get("meta") or {}).get("mode") or "build"
        if prev_sig is None or not chg_list:
            entry = f"v1: initial {mode_tag} build ({len(cur_slides)} slides, gates_failed={gate_fail}, visual_div={vis_d}%)"
            chg_list = [entry]
        elif prev_sig == cur_sig:
            # Re-run on identical spec: refresh status on latest entry without duplicating history
            v_idx = len(chg_list)
            base_lbl = "initial" if v_idx == 1 else "refined"
            chg_list[-1] = f"v{v_idx}: {base_lbl} {mode_tag} build ({len(cur_slides)} slides, gates_failed={gate_fail}, visual_div={vis_d}%)"
        else:
            changed_s = []
            max_n = max(len(prev_sig), len(cur_sig))
            for si_c in range(max_n):
                if si_c >= len(prev_sig) or si_c >= len(cur_sig) or prev_sig[si_c] != cur_sig[si_c]:
                    changed_s.append(si_c + 1)
            v_idx = len(chg_list) + 1
            chg_list.append(
                f"v{v_idx}: refined {mode_tag} deck ({len(cur_slides)} slides, modified slides {changed_s[:10]}, gates_failed={gate_fail}, visual_div={vis_d}%)"
            )
        chg_p.write_text(json.dumps(chg_list[-8:], ensure_ascii=False, indent=1), encoding="utf-8")
        prev_p.write_text(json.dumps(cur_sig, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

    # 5) Assemble result.json
    print("== [5/5] result.json ==")
    res_args = ["--case", str(cpath), "--work", str(work), "--out", str(result_out)]
    if (work / "result_fill.json").exists():
        res_args += ["--fill", str(work / "result_fill.json")]
    rres = run_py("build_result.py", res_args)
    print("  " + (rres.stdout or rres.stderr).strip().replace("\n", "\n  "))

    try:
        rj = json.loads(result_out.read_text(encoding="utf-8"))
        ec = rj.get("evidence_coverage") or {}
        ts = rj.get("timing_summary") or {}
        qi = rj.get("qa_integrity") or {}
        sc = ec.get("strength_counts") or {}
        print(
            f"  [audit] coverage={ec.get('coverage_pct', '?')}% visual_div={ec.get('visual_diversity_pct', '?')}% "
            f"strengths=S1:{sc.get('S1',0)}/S2:{sc.get('S2',0)}/S3:{sc.get('S3',0)} "
            f"qa={qi.get('total',0)}({len(qi.get('categories') or [])}cats) "
            f"timing={ts.get('total_timed_s',0)}s/{ts.get('budget_s') or '?'}s"
        )
        miss_q = qi.get("categories_missing") or []
        if int(qi.get("total") or 0) < 4 or miss_q:
            print(f"  [qa-advisory] qa_prep has {qi.get('total',0)} item(s)" + (f", missing categories: {', '.join(miss_q)}" if miss_q else "") + "; aim for >=4 covering mechanism/data/baseline/limitation")
    except Exception:
        pass
    if audit_out:
        print("  " + audit_out.replace("\n", "\n  "))

    if fix_hints:
        print("== NEXT FIX ==")
        for fh in fix_hints:
            print(f"  - {fh}")
    print(f"== DELIVER SUMMARY: gates_failed={gate_fail} | deck={deck_out} | result={result_out} ==")
    sys.exit(0 if gate_fail == 0 else 1)


if __name__ == "__main__":
    main()
