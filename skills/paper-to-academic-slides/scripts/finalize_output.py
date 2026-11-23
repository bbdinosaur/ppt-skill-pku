#!/usr/bin/env python3
"""Finalize outputs: export PDF + PNG previews (100 DPI), verify rendering,
export standalone evidence_map.json & speaker_notes.md, align with
case.json output_contract, and assemble result.json.

Usage:
  python3 finalize_output.py deck.pptx --slides-json slides.json \
      --case-id CASE --out-dir /app/output --work /app/work \
      [--template T.pptx] [--case-resolved /app/work/case_resolved.json] [--dpi 100]

Never invents data: missing pieces are recorded as honest gaps.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata

LIGATURES = {
    "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl",
    "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st",
}


def norm_alnum(t):
    s = t or ""
    for k, v in LIGATURES.items():
        if k in s:
            s = s.replace(k, v)
    s = unicodedata.normalize("NFKC", s).lower()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", s)


def sh(cmd, timeout=180):
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=timeout)
        return r.returncode, r.stdout.decode("utf-8", "ignore")[-800:], \
            r.stderr.decode("utf-8", "ignore")[-800:]
    except Exception as e:
        return -1, "", str(e)


def count_pptx_slides(p):
    from pptx import Presentation
    return len(Presentation(p).slides._sldIdLst)


def count_pdf_pages(p):
    try:
        import pymupdf
        with pymupdf.open(p) as d:
            return d.page_count
    except Exception:
        return -1


def render_previews(pdf, outdir, dpi=100):
    import pymupdf
    os.makedirs(outdir, exist_ok=True)
    outs = []
    with pymupdf.open(pdf) as d:
        for i, page in enumerate(d):
            fn = os.path.join(outdir, f"slide-{i+1:02d}.png")
            page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72.0, dpi / 72.0)).save(fn)
            outs.append(fn)
    return outs


def tool_versions():
    vers = {"python": sys.version.split()[0]}
    for mod in ("pptx", "pymupdf", "PIL"):
        try:
            m = __import__(mod)
            vers[mod] = getattr(m, "__version__", "installed")
        except Exception:
            vers[mod] = "missing"
    rc, so, _ = sh(["soffice", "--version"], timeout=15)
    if rc == 0 and so.strip():
        vers["soffice"] = so.strip().splitlines()[0]
    return vers


def load_json_opt(path):
    if path and os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--slides-json", default=None)
    ap.add_argument("--case-id", default=None)
    ap.add_argument("--out-dir", default="/app/output")
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--template", default=None)
    ap.add_argument("--case-resolved", default=None)
    ap.add_argument("--dpi", type=int, default=100)
    a = ap.parse_args()

    out = os.path.abspath(a.out_dir)
    os.makedirs(out, exist_ok=True)
    os.makedirs(os.path.join(out, "previews"), exist_ok=True)
    log, critical_gaps, advisory_gaps = [], [], []

    cr_path = a.case_resolved or os.path.join(a.work, "case_resolved.json")
    case_res = load_json_opt(cr_path) or {}
    if case_res.get("label_discrepancy"):
        advisory_gaps.append(f"label_discrepancy: {case_res['label_discrepancy']}")
    paper_ctx = load_json_opt(os.path.join(a.work, "paper", "paper_context.json")) or {}
    for lim_msg in (paper_ctx.get("limits") or [])[:6]:
        if isinstance(lim_msg, str) and lim_msg.strip():
            advisory_gaps.append(f"extract_limit: {lim_msg.strip()}")
    case_id = a.case_id or case_res.get("case_id")
    template_path = a.template or case_res.get("template_pptx")

    deck_src = os.path.abspath(a.deck)
    deck_out = os.path.join(out, "slides.pptx")
    shutil.copyfile(deck_src, deck_out)
    log.append(f"deck copied -> {deck_out} ({os.path.getsize(deck_out)} B)")

    # Mirror PPTX if case.json explicitly requested a custom pptx output filename
    extra_copies = []
    for _k, exp_name in (case_res.get("expected_outputs") or {}).items():
        if isinstance(exp_name, str) and exp_name.lower().endswith(".pptx"):
            bn = os.path.basename(exp_name)
            if bn and bn != "slides.pptx":
                dst = os.path.join(out, bn)
                shutil.copyfile(deck_out, dst)
                extra_copies.append(dst)
                log.append(f"mirrored requested output -> {dst}")

    pdf_out = os.path.join(out, "slides.pdf")
    lo_prof = f"/tmp/lo_profile_{os.getpid()}"
    try:
        rc, _so, se = sh([
            "soffice",
            f"-env:UserInstallation=file://{lo_prof}",
            "--headless", "--norestore", "--nolockcheck",
            "--convert-to", "pdf", "--outdir", out, deck_out,
        ], timeout=180)
    finally:
        shutil.rmtree(lo_prof, ignore_errors=True)
    if not os.path.isfile(pdf_out) or os.path.getsize(pdf_out) < 2000:
        critical_gaps.append(f"PDF export failed (rc={rc}); soffice stderr: {se.strip()[:200]}")
        pdf_out = None
    else:
        log.append(f"PDF exported -> {pdf_out} ({os.path.getsize(pdf_out)} B)")

    render_dpi = min(max(int(a.dpi), 72), 108)
    previews = []
    if pdf_out:
        try:
            previews = render_previews(pdf_out, os.path.join(out, "previews"), render_dpi)
            log.append(f"{len(previews)} PNG previews rendered at {render_dpi} DPI")
        except Exception as e:
            critical_gaps.append(f"preview render failed: {e}")

    try:
        n_slides = count_pptx_slides(deck_out)
    except Exception as e:
        n_slides = 0
        critical_gaps.append(f"pptx open/count failed: {e}")
    n_pdf = count_pdf_pages(pdf_out) if pdf_out else -1
    if pdf_out and n_pdf != n_slides:
        critical_gaps.append(f"pdf pages ({n_pdf}) != pptx slides ({n_slides})")
    if pdf_out and n_slides > 0 and len(previews) < n_slides:
        critical_gaps.append(f"incomplete previews ({len(previews)}/{n_slides})")

    # Notes presence check + speaker notes markdown export
    notes_ok, notes_missing, notes_text = True, [], []
    try:
        from pptx import Presentation
        prs = Presentation(deck_out)
        for i, s in enumerate(prs.slides, 1):
            t = ""
            if s.has_notes_slide:
                t = s.notes_slide.notes_text_frame.text.strip()
            if len(t) < 10:
                notes_missing.append(i)
                notes_ok = False
            title = ""
            # Prefer the actual Title placeholder (idx==0 or TITLE/CTITLE) over decorative shapes
            for ph in s.placeholders:
                try:
                    if ph.placeholder_format.idx == 0 or int(ph.placeholder_format.type) in (1, 3):
                        if ph.has_text_frame and ph.text_frame.text.strip():
                            title = ph.text_frame.text.strip().replace("\n", " ")
                            break
                except Exception:
                    pass
            if not title:
                for shp in s.shapes:
                    if shp.has_text_frame and shp.text_frame.text.strip():
                        title = shp.text_frame.text.strip().splitlines()[0]
                        break
            notes_text.append((i, title, t))
    except Exception as e:
        critical_gaps.append(f"notes check failed: {e}")

    notes_md_path = os.path.join(out, "speaker_notes.md")
    try:
        with open(notes_md_path, "w", encoding="utf-8") as f:
            f.write("# 讲稿（逐页，与页面文字分离）\n\n")
            for i, title, t in notes_text:
                f.write(f"## Slide {i} — {title[:80]}\n\n{t}\n\n" if t else
                        f"## Slide {i} — {title[:80]}\n\n（无备注）\n\n")
    except Exception:
        pass

    # Render text integrity + real PDF font audit + automated pixel visual check
    render_text_ok, render_bad, tofu_slides = None, [], []
    pdf_fonts_used = []
    if pdf_out and notes_text:
        try:
            import pymupdf
            pdf_pages = {}
            font_set = set()
            with pymupdf.open(pdf_out) as _d:
                for idx_p, pg in enumerate(_d, 1):
                    raw_t = pg.get_text("text") or ""
                    pdf_pages[idx_p] = norm_alnum(raw_t)
                    if "\ufffd" in raw_t:
                        tofu_slides.append(idx_p)
                    try:
                        for f_tup in pg.get_fonts():
                            fname = f_tup[3] if len(f_tup) > 3 else ""
                            if fname:
                                # Strip subset prefix like "ABCDEE+LiberationSans"
                                clean_fn = fname.split("+", 1)[-1]
                                font_set.add(clean_fn)
                    except Exception:
                        pass
            pdf_fonts_used = sorted(font_set)[:20]
            for i, title, _t in notes_text:
                nt = norm_alnum(title)
                pg_txt = pdf_pages.get(i, "")
                if len(nt) >= 4:
                    head = nt[:10]
                    mid = nt[len(nt) // 3: len(nt) // 3 + 8] if len(nt) >= 14 else head
                    if head not in pg_txt and mid not in pg_txt:
                        render_bad.append(i)
            render_text_ok = (not render_bad) and (not tofu_slides)
            if render_bad:
                advisory_gaps.append(f"slide titles missing from rendered PDF pages: {render_bad} "
                                     "(possible font substitution/glyph loss)")
            if tofu_slides:
                advisory_gaps.append(f"replacement glyph U+FFFD found in rendered PDF pages: {tofu_slides}")
            log.append(f"render text check: {'ok' if render_text_ok else {'missing': render_bad, 'tofu': tofu_slides}} | pdf_fonts={pdf_fonts_used[:6]}")
        except Exception as e:
            advisory_gaps.append(f"render text check failed: {e}")

    visual_blank_slides = []
    visual_ink_ratios = []
    if previews:
        try:
            from PIL import Image
            for idx_im, ppath in enumerate(previews, 1):
                with Image.open(ppath) as im:
                    rgb = im.convert("RGB")
                    rgb.thumbnail((320, 180))
                    w_i, h_i = rgb.size
                    px = list(rgb.getdata())
                if not px:
                    continue
                corners = [px[0], px[w_i - 1], px[(h_i - 1) * w_i], px[h_i * w_i - 1]]
                bg = tuple(sorted(c[k] for c in corners)[1] for k in range(3))
                non_bg = sum(1 for r, g, b in px
                             if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) > 35)
                ratio = round(non_bg / max(len(px), 1), 4)
                visual_ink_ratios.append({"slide": idx_im, "ink_ratio": ratio})
                if ratio < 0.004:
                    visual_blank_slides.append(idx_im)
            if visual_blank_slides:
                advisory_gaps.append(f"nearly blank rendered slides detected: {visual_blank_slides}")
            log.append(f"visual pixel check: {len(previews)} pngs, blank={visual_blank_slides}")
        except Exception as e:
            log.append(f"visual pixel check skipped: {e}")

    build_report_path = os.path.join(a.work, "deck_build_report.json")
    build_rep = load_json_opt(build_report_path)
    layout_report_path = os.path.join(a.work, "layout_report.json")
    layout = load_json_opt(layout_report_path)
    ev_report_path = os.path.join(a.work, "evidence_report.json")
    ev_report = load_json_opt(ev_report_path)

    # Evidence map from slides.json
    evidence_map, outline, deck_meta = [], [], {}
    if a.slides_json and os.path.isfile(a.slides_json):
        try:
            with open(a.slides_json, encoding="utf-8") as f:
                spec = json.load(f)
            deck_meta = spec.get("meta", {})
            for i, sld in enumerate(spec.get("slides", []), 1):
                ev = sld.get("evidence") or {}
                evidence_map.append({
                    "slide": i,
                    "kind": sld.get("kind", "content"),
                    "claim": sld.get("title", ""),
                    "paper_refs": ev.get("paper_refs", []),
                    "slide_line": ev.get("slide_line", ""),
                    "figure_file": sld.get("image") or (sld.get("images") or [None])[0],
                    "has_table": bool(sld.get("table")),
                    "has_chart": bool(sld.get("chart")),
                    "has_kpis": bool(sld.get("kpis") or sld.get("metrics")),
                    "two_column": bool(sld.get("right_bullets")),
                    "notes_present": bool(sld.get("notes")),
                    "notes_chars": len((sld.get("notes") or "").strip()),
                })
                outline.append({
                    "slide": i,
                    "title": sld.get("title", ""),
                    "kind": sld.get("kind", "content"),
                    "n_bullets": len(sld.get("bullets", [])) + len(sld.get("right_bullets", [])),
                })
        except Exception as e:
            critical_gaps.append(f"slides.json parse failed: {e}")

    # Annotate per-slide verification status
    if evidence_map and ev_report:
        by_slide = {}
        for iss in ev_report.get("issues", []):
            by_slide.setdefault(iss.get("slide"), []).append(
                {"sev": iss.get("sev"), "type": iss.get("type")})
        for item in evidence_map:
            sv = by_slide.get(item["slide"], [])
            has_err = any(x.get("sev") == "error" for x in sv)
            item["verify"] = "error" if has_err else ("warn" if sv else "ok")
            if sv:
                item["verify_issues"] = sv[:6]

    # Export standalone evidence_map.json, human-auditable evidence_map.md, and reproducible slides_outline.json
    ev_map_out = os.path.join(out, "evidence_map.json")
    try:
        with open(ev_map_out, "w", encoding="utf-8") as f:
            json.dump({"case_id": case_id, "n_slides": n_slides,
                       "evidence_map": evidence_map}, f, ensure_ascii=False, indent=1)
    except Exception:
        ev_map_out = None

    ev_md_out = os.path.join(out, "evidence_map.md")
    try:
        md_lines = [
            f"# Slide-Level Evidence Map (`{case_id or 'case'}`)",
            "",
            "| Slide | Claim Title | Visual / Table | Paper References & Verbatim Quotes | Verify |",
            "|---|---|---|---|---|",
        ]
        for em in evidence_map:
            vis = []
            if em.get("has_kpis"):
                vis.append("KPIs")
            if em.get("figure_file"):
                vis.append(f"`{os.path.basename(em['figure_file'])}`")
            if em.get("has_table"):
                vis.append("Native Table")
            if em.get("two_column") and not em.get("figure_file") and not em.get("has_table"):
                vis.append("2-Col")
            vis_s = " + ".join(vis) if vis else "Text"
            ref_parts = []
            for r in (em.get("paper_refs") or [])[:3]:
                rid = str(r.get("id") or r.get("type") or "").strip()
                pg = r.get("page")
                qt = (r.get("quote") or "").strip().replace("|", "/").replace("\n", " ")[:70]
                tag = f"{rid} (p.{pg})" if pg else rid
                ref_parts.append(f"{tag}: \"{qt}\"" if qt else tag)
            refs_s = "; ".join(ref_parts) or "(none)"
            claim_s = (em.get("claim") or "").replace("|", "/").replace("\n", " ")[:90]
            md_lines.append(f"| {em['slide']} | {claim_s} | {vis_s} | {refs_s} | {em.get('verify', 'ok')} |")
        with open(ev_md_out, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines) + "\n")
    except Exception:
        ev_md_out = None

    outline_json_out = None
    if a.slides_json and os.path.isfile(a.slides_json):
        try:
            outline_json_out = os.path.join(out, "slides_outline.json")
            shutil.copyfile(os.path.abspath(a.slides_json), outline_json_out)
        except Exception:
            outline_json_out = None

    # Mirror any requested custom PDF / evidence_map / speaker_notes / outline filenames from output_contract
    for _k, exp_name in (case_res.get("expected_outputs") or {}).items():
        if not isinstance(exp_name, str):
            continue
        bn = os.path.basename(exp_name)
        bl = bn.lower()
        if not bn:
            continue
        src_map = None
        if bl.endswith(".pdf") and bl != "slides.pdf" and pdf_out and os.path.isfile(pdf_out):
            src_map = pdf_out
        elif bl.endswith(".md") and "evidence" in bl and bl != "evidence_map.md" and ev_md_out and os.path.isfile(ev_md_out):
            src_map = ev_md_out
        elif bl.endswith(".md") and bl != "speaker_notes.md" and os.path.isfile(notes_md_path):
            src_map = notes_md_path
        elif bl.endswith(".json") and bl not in ("result.json", "evidence_map.json") and "evidence" in bl and ev_map_out and os.path.isfile(ev_map_out):
            src_map = ev_map_out
        elif bl.endswith(".json") and bl not in ("result.json", "slides_outline.json") and ("outline" in bl or "slides" in bl) and outline_json_out and os.path.isfile(outline_json_out):
            src_map = outline_json_out
        if src_map:
            dst = os.path.join(out, bn)
            if dst not in extra_copies:
                shutil.copyfile(src_map, dst)
                extra_copies.append(dst)
                log.append(f"mirrored requested output -> {dst}")

    ev_err = ev_report.get("n_error") if ev_report else None
    all_gaps = (critical_gaps + advisory_gaps +
                (layout.get("issues", []) if layout else [])[:20] +
                ([{"sev": i.get("sev"), "type": i.get("type"), "slide": i.get("slide"),
                   "detail": i.get("detail")} for i in ev_report.get("issues", [])]
                 if ev_report else []))

    # Core deliverable completeness determines top-level status:
    # - 'failed'  : core editable PPTX missing/corrupt (<5KB or 0 slides)
    # - 'partial' : core deliverable missing or critical export/parse step broke
    # - 'success' : valid PPTX + PDF + complete previews + output_contract fields present;
    #               heuristic layout/evidence advisories remain honestly recorded in gaps/issues/uncertainty
    pptx_valid = os.path.isfile(deck_out) and os.path.getsize(deck_out) >= 5000 and n_slides > 0
    if not pptx_valid:
        status = "failed"
    elif critical_gaps or not pdf_out or not previews or len(previews) < n_slides:
        status = "partial"
    else:
        status = "success"

    sources_list = []
    if case_res.get("paper_source"):
        sources_list.append({"role": "paper", "path": case_res["paper_source"]})
    if template_path:
        sources_list.append({"role": "template", "path": os.path.abspath(template_path)})
    for m in (case_res.get("materials") or [])[:15]:
        if m not in {s["path"] for s in sources_list}:
            sources_list.append({"role": "material", "path": m})

    template_info = case_res.get("template_info") or {}
    theme_fonts = template_info.get("theme_fonts", {})
    # Font substitution audit: check whether template Latin theme fonts appear in exported PDF fonts
    font_sub_detected = False
    font_audit_note = "ok"
    if theme_fonts and pdf_fonts_used:
        pdf_f_low = " ".join(pdf_fonts_used).lower()
        missing_tf = []
        for tf_k in ("major_latin", "minor_latin"):
            tf_v = (theme_fonts.get(tf_k) or "").strip()
            if tf_v and tf_v.lower().split()[0] not in pdf_f_low:
                missing_tf.append(tf_v)
        if missing_tf:
            font_sub_detected = True
            font_audit_note = (f"template theme font(s) {sorted(set(missing_tf))} substituted during PDF export "
                               f"(rendered with {pdf_fonts_used[:5]}); layout verified post-render")

    tpl_mode = (build_rep.get("template_mode") if build_rep else
                ("reuse" if deck_meta.get("reuse") else
                 ("layouts" if template_path else "blank")))
    render_record = {
        "source_pptx": deck_out,
        "exported_pdf": pdf_out,
        "preview_pngs": previews,
        "preview_dpi": render_dpi,
        "template": os.path.abspath(template_path) if template_path else None,
        "template_mode": tpl_mode,
        "theme_accent_hex": build_rep.get("theme_accent_hex") if build_rep else None,
        "cleared_template_slides": build_rep.get("cleared_template_slides", 0) if build_rep else 0,
        "layouts_used": build_rep.get("layouts_used", []) if build_rep else [],
        "theme_fonts": theme_fonts,
        "font_audit": build_rep.get("font_audit") if build_rep else None,
        "pdf_fonts_used": pdf_fonts_used,
        "font_substitution_detected": font_sub_detected,
        "font_audit_note": font_audit_note,
        "tool_versions": tool_versions(),
        "pdf_pages_match": (n_pdf == n_slides) if pdf_out else None,
        "render_text_ok": render_text_ok,
        "visual_ink_ratios": visual_ink_ratios,
    }

    change_log = [
        {"step": "extract", "detail": f"parsed paper/material into {a.work}/paper/paper_context.json"},
        {"step": "build", "detail": f"built {n_slides}-slide editable PPTX (mode={tpl_mode})"},
        {"step": "verify_and_render", "detail": f"layout fail={layout.get('n_fail') if layout else None}, "
                                                f"evidence err={ev_err}, pdf={bool(pdf_out)}, previews={len(previews)}, "
                                                f"font_sub={font_sub_detected}"},
    ]
    # In --reuse mode, record authentic per-slide diff against original starting deck titles (honoring source_slide)
    if tpl_mode == "reuse":
        src_info = case_res.get("starting_deck_info") or template_info
        orig_titles = {t.get("slide"): t.get("title", "") for t in (src_info.get("existing_slide_titles") or [])}
        rmap_by_sn = {m.get("slide"): m for m in ((build_rep or {}).get("reuse_mapping") or [])}
        slide_diffs = []
        for item in outline:
            sn = item["slide"]
            rm = rmap_by_sn.get(sn)
            src_no = rm.get("source_slide") if rm is not None else sn
            old_t = orig_titles.get(src_no) if src_no is not None else None
            if src_no is None or old_t is None:
                slide_diffs.append({"slide": sn, "source_slide": None, "action": "added", "new_title": item["title"][:80]})
            elif old_t.strip() != item["title"].strip():
                slide_diffs.append({"slide": sn, "source_slide": src_no, "action": "revised",
                                    "old_title": old_t[:60], "new_title": item["title"][:80]})
            else:
                slide_diffs.append({"slide": sn, "source_slide": src_no, "action": "updated_body_notes_evidence",
                                    "title": item["title"][:80]})
        if slide_diffs:
            change_log.append({"step": "reuse_slide_diff", "slides": slide_diffs[:20]})

    # Parse optional speaker-notes pacing tags (e.g. "[~1.0 min]", "[~1.5m]", "1.5分钟")
    pace_re = re.compile(r"(?:\[\s*~?(\d+(?:\.\d+)?)\s*(?:min(?:ute)?s?|m|分钟)\s*\]|~(\d+(?:\.\d+)?)\s*(?:min|分钟))", re.I)
    paced_vals = []
    for _ni, _nt, n_txt in notes_text:
        pm = pace_re.search(n_txt or "")
        if pm:
            try:
                paced_vals.append(float(pm.group(1) or pm.group(2)))
            except Exception:
                pass
    target_dur = (case_res.get("constraints") or {}).get("duration_minutes") or (case_res.get("constraints") or {}).get("duration")

    # Strictly data-derived 6-element academic presentation audit
    ev_issues_all = (ev_report.get("issues") or []) if isinstance(ev_report, dict) else []
    content_em = [e for e in evidence_map if e.get("kind") != "cover" and e.get("slide", 0) > 1]
    content_out = [o for o in outline if o.get("kind") != "cover" and o.get("slide", 0) > 1]
    all_titles_blob = " ".join((o.get("title") or "") for o in outline).lower()
    notes_lens = [e.get("notes_chars", 0) for e in evidence_map if e.get("notes_present")]
    six_element_audit = {
        "storyline_coverage": {
            "has_cover": bool(outline),
            "has_limitations": bool(re.search(r"limitation|drawback|caveat|future|局限|不足|限制|边界", all_titles_blob)),
            "has_conclusion": bool(re.search(r"conclu|takeaway|summary|结论|总结|展望", all_titles_blob)) or (
                any(o.get("kind") == "closing" for o in outline)
            ),
            "n_slides": n_slides,
        },
        "slide_structure": {
            "avg_bullets_per_content_slide": round(
                sum(o.get("n_bullets", 0) for o in content_out) / max(len(content_out), 1), 2
            ) if content_out else 0.0,
            "generic_title_warns": sum(1 for i in ev_issues_all if i.get("type") == "generic_section_title"),
        },
        "figure_table_grounding": {
            "n_figures_embedded": sum(1 for e in evidence_map if e.get("figure_file")),
            "n_tables_embedded": sum(1 for e in evidence_map if e.get("has_table")),
            "n_charts_embedded": sum(1 for e in evidence_map if e.get("has_chart")),
            "content_slides_with_refs": sum(1 for e in content_em if e.get("paper_refs")),
            "content_slides_with_quotes": sum(
                1 for e in content_em
                if any(len((r.get("quote") or "").strip()) >= 8 for r in (e.get("paper_refs") or []))
            ),
            "total_content_slides": len(content_em),
        },
        "speaker_notes_separation": {
            "notes_complete": notes_ok,
            "avg_notes_chars": round(sum(notes_lens) / max(len(notes_lens), 1), 1) if notes_lens else 0.0,
            "notes_dup_warns": sum(1 for i in ev_issues_all if i.get("type") == "notes_duplicate_slide_text"),
            "paced_minutes_total": round(sum(paced_vals), 2) if paced_vals else None,
            "target_duration_minutes": target_dur,
        },
        "template_compliance": {
            "template_mode": render_record["template_mode"],
            "theme_accent_hex": render_record["theme_accent_hex"],
            "layouts_used_count": len(set(render_record["layouts_used"])),
            "geometry_fail": layout.get("n_fail", None) if layout else None,
        },
        "export_render_verification": {
            "pdf_exported": bool(pdf_out),
            "pdf_pages_match": (n_pdf == n_slides) if pdf_out else None,
            "previews_rendered": len(previews),
            "render_text_ok": render_text_ok,
            "visual_pixel_ok": (not visual_blank_slides) and (not tofu_slides) if previews else None,
            "font_substitution_detected": font_sub_detected,
        },
    }

    result = {
        "case_id": case_id,
        "task": "paper_to_academic_slides",
        "status": status,
        "deliverables": {
            "pptx": deck_out,
            "pdf": pdf_out,
            "previews": previews,
            "slides_outline_json": outline_json_out or (os.path.abspath(a.slides_json) if a.slides_json else None),
            "evidence_map_json": ev_map_out,
            "evidence_map_md": ev_md_out,
            "speaker_notes": notes_md_path,
            "build_report": build_report_path if os.path.isfile(build_report_path) else None,
            "layout_report": layout_report_path if os.path.isfile(layout_report_path) else None,
            "evidence_report": ev_report_path if os.path.isfile(ev_report_path) else None,
            "extra_copies": extra_copies,
        },
        "deck": {
            "n_slides": n_slides,
            "slide_size_in": layout.get("stats", {}) if layout else {},
            "template": render_record["template"],
            "template_mode": render_record["template_mode"],
            "n_figures_embedded": sum(1 for e in evidence_map if e.get("figure_file")),
            "n_tables_embedded": sum(1 for e in evidence_map if e.get("has_table")),
            "n_charts_embedded": sum(1 for e in evidence_map if e.get("has_chart")),
            "notes_complete": notes_ok,
            "notes_missing_slides": notes_missing,
        },
        "outline": outline,
        "evidence_map": evidence_map,
        "six_element_audit": six_element_audit,
        "sources": sources_list,
        "render_record": render_record,
        "checks": {
            "pptx_opens": True,
            "pdf_exported": bool(pdf_out),
            "pdf_pages_match": (n_pdf == n_slides) if pdf_out else None,
            "previews_rendered": len(previews),
            "render_text_ok": render_text_ok,
            "visual_blank_slides": visual_blank_slides,
            "visual_tofu_slides": tofu_slides,
            "visual_pixel_ok": (not visual_blank_slides) and (not tofu_slides) if previews else None,
            "geometry_fail": layout.get("n_fail", None) if layout else None,
            "geometry_warn": layout.get("n_warn", None) if layout else None,
            "evidence_errors": ev_err,
            "evidence_warns": ev_report.get("n_warn") if ev_report else None,
        },
        "gaps_and_uncertainty": all_gaps,
        "uncertainty": all_gaps,
        "gaps": all_gaps,
        "issues": all_gaps,
        "critical_gaps": critical_gaps,
        "pipeline_log": log,
        "change_log": change_log,
        "change_records": change_log,
    }

    # Align any extra required top-level fields declared in case.json output_contract
    oc = case_res.get("output_contract")
    if isinstance(oc, dict):
        explicit_req = oc.get("required_fields") or oc.get("fields")
        req_fields = explicit_req if isinstance(explicit_req, list) else list(oc.keys())
        if isinstance(req_fields, list):
            alias_map = {
                "summary": f"{n_slides}-slide editable academic deck ({status})",
                "editable_artifact": deck_out,
                "editable_pptx": deck_out,
                "pptx_path": deck_out,
                "deck_path": deck_out,
                "pdf_path": pdf_out,
                "artifacts": result["deliverables"],
                "files": result["deliverables"],
                "slide_evidence_map": evidence_map,
                "evidence": evidence_map,
                "citations": evidence_map,
                "notes": notes_md_path,
                "speaker_notes": notes_md_path,
                "rendered_pdf": pdf_out,
                "rendered_previews": previews,
                "previews": previews,
                "verification": result["checks"],
                "render_verification": render_record,
                "fonts": {"theme_fonts": theme_fonts, "pdf_fonts_used": pdf_fonts_used,
                          "font_substitution_detected": font_sub_detected},
                "limitations": all_gaps,
                "gaps": all_gaps,
                "issues": all_gaps,
            }
            for rf in req_fields:
                if isinstance(rf, str) and rf not in result and rf in alias_map:
                    result[rf] = alias_map[rf]
            # Only downgrade to partial if explicitly required output_contract fields remain missing
            if isinstance(explicit_req, list):
                missing_req = [rf for rf in explicit_req
                               if isinstance(rf, str) and (rf not in result or result[rf] is None)]
                if missing_req and status == "success":
                    msg = f"missing required output_contract field(s): {missing_req}"
                    critical_gaps.append(msg)
                    all_gaps.insert(0, msg)
                    status = "partial"
                    result["status"] = status
                    if "summary" in result and isinstance(result["summary"], str):
                        result["summary"] = f"{n_slides}-slide editable academic deck ({status})"

    rp = os.path.join(out, "result.json")
    with open(rp, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    print(json.dumps({"ok": status != "failed", "status": status, "result_json": rp,
                      "n_slides": n_slides, "pdf": bool(pdf_out),
                      "previews": len(previews), "gaps": critical_gaps + advisory_gaps}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
