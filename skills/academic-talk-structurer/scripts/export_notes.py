#!/usr/bin/env python3
"""Export a printable speaker handout (markdown) from the built deck: claim, time, script, evidence, QA keys.

Usage:
  python3 export_notes.py --deck /app/output/talk.pptx --plan /app/work/slide_plan.json \
      [--qa /app/work/qa_plan.json] [--out /app/output/speaker_notes.md]

No-script fallback: copy notes from PowerPoint notes pane manually per slide_plan.json order.
Exit 0 always (honest artifact); warns on missing inputs.
"""
import argparse
import json
import re
from pathlib import Path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--qa")
    ap.add_argument("--outline")
    ap.add_argument("--out", default="speaker_notes.md")
    ap.add_argument("--evidence-map", dest="evidence_map", help="also export standalone slide-level evidence_map.md for independent review")
    ap.add_argument("--pdf", dest="pdf", action="store_true", help="also render PDF via reportlab (skip if unavailable)")
    a = ap.parse_args()
    plan = json.loads(Path(a.plan).read_text(encoding="utf-8")) if Path(a.plan).exists() else {"slides": []}
    slides = plan.get("slides") or []

    notes_by_pos = {}
    try:
        from pptx import Presentation
        prs = Presentation(a.deck)
        for i, sl in enumerate(prs.slides, 1):
            try:
                notes_by_pos[i] = sl.notes_slide.notes_text_frame.text.strip()
            except Exception:
                notes_by_pos[i] = ""
    except Exception as e:
        print(f"warn: cannot read deck notes ({e!r}); falling back to plan-only handout")

    meta = plan.get("meta") or {}
    lines = ["# 讲稿讲义与取舍复核 / Speaker Notes & Editorial Summary", "",
             f"- deck: `{a.deck}`",
             f"- 总时长预算: {meta.get('duration_minutes', '?')} min"
             f"（每内容页约 {meta.get('seconds_per_content_slide', meta.get('seconds_per_timed_slide', '?'))}s）"]
    if meta.get("goal"):
        lines.append(f"- 报告目标 / Goal: {meta['goal']}")
    if meta.get("audience"):
        lines.append(f"- 受众定位 / Audience: {meta['audience']}")
    if meta.get("tradeoffs"):
        lines.append("- 关键内容取舍 / Editorial Tradeoffs:")
        for tr in meta["tradeoffs"]:
            lines.append(f"  - {tr}")
    # Compact verification summary for fast independent review
    c_slides = [s for s in slides if s.get("role") not in ("title", "thanks", "qa_backup", "preserved")]
    if c_slides:
        ev_bound = sum(1 for s in c_slides if any(e.get("locator") for e in (s.get("evidence") or []) if isinstance(e, dict)))
        vis_bound = sum(1 for s in c_slides if (s.get("visuals") or s.get("figure") or s.get("has_table")))
        s_cnts = {"S1": 0, "S2": 0, "S3": 0}
        for s in c_slides:
            for e in (s.get("evidence") or []):
                st_u = str(e.get("strength") or "").upper() if isinstance(e, dict) else ""
                if st_u in s_cnts:
                    s_cnts[st_u] += 1
        t_sum = sum(int(s.get("time_s") or 0) for s in slides if s.get("role") not in ("qa_backup", "preserved"))
        bud_str = f"/{int(round(float(meta['duration_minutes']) * 60))}s" if meta.get("duration_minutes") else ""
        qa_sum_s = ""
        if a.qa and Path(a.qa).exists():
            try:
                _qa_l = json.loads(Path(a.qa).read_text(encoding="utf-8"))
                if isinstance(_qa_l, list):
                    _cats = sorted({str(q.get("category") or "").lower() for q in _qa_l if isinstance(q, dict) and q.get("category")})
                    qa_sum_s = f" | Q&A={len(_qa_l)}({len(_cats)}维)"
            except Exception:
                pass
        rd_sum_s = ""
        _qa_aud_p = Path(a.plan).parent / "quality_audit.json"
        if _qa_aud_p.exists():
            try:
                _qa_aud = json.loads(_qa_aud_p.read_text(encoding="utf-8"))
                if _qa_aud.get("readiness_score"):
                    rd_sum_s = f" | 就绪核验={_qa_aud['readiness_score']}"
            except Exception:
                pass
        lines.append(
            f"- 复核摘要 / Audit Summary: 证据覆盖 {ev_bound}/{len(c_slides)} ({round(100*ev_bound/len(c_slides))}%) "
            f"| 视觉页 {vis_bound}/{len(c_slides)} ({round(100*vis_bound/len(c_slides))}%) "
            f"| 强度 S1:{s_cnts['S1']}/S2:{s_cnts['S2']}/S3:{s_cnts['S3']} "
            f"| 节奏 {t_sum}s{bud_str}{qa_sum_s}{rd_sum_s}"
        )
    lines.append("")

    # Compact Slide-by-Slide Evidence & Timing Matrix for independent audit
    if slides:
        lines.append("## 页级主张、证据与节奏总览 / Slide-by-Slide Evidence & Timing Matrix")
        lines.append("")
        lines.append("| Slide | Role | Title | One-Sentence Claim | Evidence Locator | Strength | Visual | Time |")
        lines.append("|---|---|---|---|---|---|---|---|")
        for s in slides:
            pos = s.get("pptx_slide")
            role = s.get("role", "")
            ttl = str(s.get("title") or "").replace("|", "/")[:48]
            clm = str(s.get("claim") or "").replace("|", "/")[:64]
            evs = s.get("evidence") or []
            locs = ", ".join(str(e.get("locator") or "") for e in evs if e.get("locator")).replace("|", "/")[:40] or "—"
            strs = "/".join(sorted({str(e.get("strength") or "") for e in evs if e.get("strength")})) or "—"
            vis = ", ".join(str(v) for v in (s.get("visuals") or [])) or (str(s.get("figure_label") or "") or ("table" if s.get("has_table") else "—"))
            vis = vis.replace("|", "/")[:28]
            ts = s.get("time_s")
            tstr = f"{ts}s" if ts else "backup"
            lines.append(f"| P{pos} | {role} | {ttl} | {clm} | {locs} | {strs} | {vis} | {tstr} |")
        lines.append("")

    total_s = 0
    for s in slides:
        pos = s.get("pptx_slide")
        ts = s.get("time_s")
        total_s += ts or 0
        lines.append(f"## P{pos} · {s.get('title', '')}")
        lines.append(f"- 角色: {s.get('role')} | 时长: {f'{ts}s' if ts else '备用/不计'}")
        if s.get("claim"):
            lines.append(f"- 主张: {s['claim']}")
        ev = s.get("evidence") or []
        if ev:
            lines.append("- 证据: " + "; ".join(f"{e.get('locator')} [{e.get('strength')}]" for e in ev))
        nt = notes_by_pos.get(pos, "")
        if nt:
            lines.append(f"- 讲稿: {nt}")
        lines.append("")
    lines.append(f"> 口播合计约 {total_s}s（不含备用页与开场缓冲）。")
    if a.qa and Path(a.qa).exists():
        qa = json.loads(Path(a.qa).read_text(encoding="utf-8"))
        backup_nums = [s.get("pptx_slide") for s in slides if s.get("role") == "qa_backup" and s.get("pptx_slide")]
        default_bk = backup_nums[0] if backup_nums else None
        lines.append("")
        lines.append("## 追问应答速查 / Q&A keys")
        lines.append("")
        for q in qa:
            cs = q.get("cite_slide")
            bs = q.get("backup_slide") or default_bk
            cat = str(q.get("category") or "").strip()
            cat_tag = f"`[{cat}]` " if cat else ""
            ref_parts = []
            if cs:
                ref_parts.append(f"引用 P{cs}")
            if bs:
                ref_parts.append(f"备用页 P{bs}")
            ref_str = f"（{'，'.join(ref_parts)}）" if ref_parts else ""
            lines.append(f"- {cat_tag}**{q.get('question')}** → {q.get('answer_key')}{ref_str}")
        lines.append("")
    if a.outline and Path(a.outline).exists():
        try:
            ol = json.loads(Path(a.outline).read_text(encoding="utf-8"))
            caps = ol.get("captions") or []
            if caps:
                cited = {}
                for s in slides:
                    pos = s.get("pptx_slide")
                    blob = " ".join(
                        [str(e.get("locator", "")) for e in (s.get("evidence") or [])]
                        + [str(s.get("figure_label") or "")]
                        + [str(v) for v in (s.get("visuals") or [])]
                    )
                    for m in re.finditer(r"(?:Fig(?:ure)?\.?\s*(S?\d+)|图\s*(\d+))", blob, re.I):
                        cited.setdefault(("fig", (m.group(1) or m.group(2)).upper()), set()).add(pos)
                    for m in re.finditer(r"(?:Tab(?:le)?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", blob, re.I):
                        cited.setdefault(("tab", (m.group(1) or m.group(2)).upper()), set()).add(pos)
                lines.append("## 源材料图表覆盖复核 / Source-to-Deck Coverage Audit")
                lines.append("")
                seen_k = set()
                for c in caps:
                    lbl = str(c.get("label", ""))
                    mf = re.search(r"(?:Fig(?:ure)?\.?\s*(S?\d+)|图\s*(\d+))", lbl, re.I)
                    mt = re.search(r"(?:Tab(?:le)?\.?\s*(S?\d+|[IVX]+)\b|表\s*(\d+))", lbl, re.I)
                    key = ("fig", (mf.group(1) or mf.group(2)).upper()) if mf else (("tab", (mt.group(1) or mt.group(2)).upper()) if mt else None)
                    if not key or key in seen_k:
                        continue
                    seen_k.add(key)
                    hits = sorted(x for x in cited.get(key, set()) if x)
                    status = ("已纳入 " + ", ".join(f"P{x}" for x in hits)) if hits else "未上正文（已压缩/略去，见取舍说明）"
                    lines.append(f"- `{lbl}` (p.{c.get('page', '?')}): {status}")
                lines.append("")
            secs = ol.get("sections") or []
            if secs:
                lines.append("### 源论文章节映射 / Source Section Coverage")
                lines.append("")
                seen_sc = set()
                for sec in secs[:14]:
                    hdg = str(sec.get("heading") or "").strip()
                    if not hdg or hdg.lower() in seen_sc:
                        continue
                    seen_sc.add(hdg.lower())
                    m_num = re.match(r"^(?:#+\s*)?(\d+(?:\.\d+)?|[IVX]{1,4})\b", hdg)
                    s_num = m_num.group(1) if m_num else None
                    pg = sec.get("page")
                    s_hits = []
                    for s in slides:
                        pno = s.get("pptx_slide")
                        ev_b = " ; ".join(str(e.get("locator", "")) for e in (s.get("evidence") or []))
                        if not pno or not ev_b:
                            continue
                        if s_num and re.search(rf"(?:§\s*|Sec(?:tion)?\.?\s*|第\s*){re.escape(s_num)}(?:\.\d+)?\b", ev_b, re.I):
                            s_hits.append(pno)
                        elif pg and re.search(rf"(?:^|[\s,;(（/])p\.?\s*{int(pg)}\b", ev_b, re.I):
                            s_hits.append(pno)
                    uniq_s = sorted(set(s_hits))
                    st_s = ("→ " + ", ".join(f"P{x}" for x in uniq_s)) if uniq_s else "→ 压缩/略去（见取舍）"
                    lines.append(f"- `{hdg[:55]}` (p.{pg}): {st_s}")
                lines.append("")
        except Exception:
            pass
    if meta.get("mode") == "edit_deck" or meta.get("orig_slides_count"):
        lines.append("## 起始演示文稿改动复核 / Starting Deck Editorial Diff")
        lines.append("")
        pruned = meta.get("pruned_orig_slides") or []
        lines.append(f"- 原始页数: {meta.get('orig_slides_count', '?')} → 成稿页数: {len(slides)}"
                     + (f" | 裁剪原页: {pruned}" if pruned else ""))
        for s in slides:
            os_no = s.get("orig_slide")
            act = s.get("edit_action") or "new"
            src_tag = f"原P{os_no} ({act})" if os_no is not None else f"新增 ({act})"
            lines.append(f"- **P{s.get('pptx_slide')}** ← {src_tag}: {s.get('title', '')}")
        lines.append("")
    Path(a.out).write_text("\n".join(lines), encoding="utf-8")
    print(f"speaker handout -> {a.out} ({len(slides)} slides)")
    if a.evidence_map:
        work_dir = Path(a.plan).parent
        def _load_w(fn):
            p_w = work_dir / fn
            if p_w.exists():
                try:
                    return json.loads(p_w.read_text(encoding="utf-8"))
                except Exception:
                    return None
            return None
        numchk = _load_w("number_check.json") or {}
        evchk = _load_w("evidence_check.json") or {}
        qaud = _load_w("quality_audit.json") or {}
        rmeta = _load_w("render_meta.json") or {}
        chgs = _load_w("changes.json") or []
        traced_map = numchk.get("traced_by_slide") or {}
        qa_items = []
        if a.qa and Path(a.qa).exists():
            try:
                _qraw = json.loads(Path(a.qa).read_text(encoding="utf-8"))
                if isinstance(_qraw, list):
                    qa_items = [q for q in _qraw if isinstance(q, dict)]
            except Exception:
                qa_items = []

        ev_lines = [
            "# 页级证据映射与取舍复核报告 / Slide-Level Evidence Map & Audit Ledger",
            "",
            f"- **演示文稿 / Deck**: `{a.deck}` (`mode={meta.get('mode', '?')}`, {len(slides)} slides, {meta.get('duration_minutes', '?')} min)",
        ]
        if meta.get("goal"):
            ev_lines.append(f"- **报告目标 / Goal**: {meta['goal']}")
        if meta.get("audience"):
            ev_lines.append(f"- **听众定位 / Audience**: {meta['audience']}")
        n_checked = numchk.get("checked", "?")
        n_susp = numchk.get("n_suspects", "?")
        e_checked = evchk.get("checked", "?")
        e_dang = evchk.get("n_dangling", "?")
        rd_sc = qaud.get("readiness_score", "?")
        ev_lines.append(
            f"- **核验状态 / Verification**: `readiness={rd_sc}` | "
            f"`numbers={n_checked} checked ({n_susp} suspects)` | "
            f"`locators={e_checked} checked ({e_dang} dangling)` | "
            f"`render_dpi={rmeta.get('dpi', '?')}`"
        )
        ev_lines.append("")
        ev_lines.append("## 1. 逐页主张—证据—强度—溯源数字总表 / Slide-by-Slide Traceability Matrix")
        ev_lines.append("")
        ev_lines.append("| Slide | Role | Takeaway Claim | Evidence Locators | Strength | Verified Numbers | Visual Asset | Time |")
        ev_lines.append("|---|---|---|---|---|---|---|---|")
        for s in slides:
            pos = s.get("pptx_slide")
            role = str(s.get("role") or "")
            clm = str(s.get("claim") or s.get("title") or "").replace("|", "/")[:62]
            evs = [e for e in (s.get("evidence") or []) if isinstance(e, dict) and e.get("locator")]
            locs = "; ".join(str(e.get("locator")).replace("|", "/") for e in evs) or "—"
            strs = "/".join(sorted({str(e.get("strength") or "NA").upper() for e in evs})) or "NA"
            vnums = ", ".join(str(x) for x in (traced_map.get(str(pos)) or [])[:6]) or "—"
            vis_list = s.get("visuals") or ([s["figure_label"]] if s.get("figure_label") else [])
            vis = ", ".join(str(v).replace("|", "/")[:24] for v in vis_list[:2]) or "—"
            ts = s.get("time_s")
            tstr = f"{ts}s" if ts else "backup"
            ev_lines.append(f"| P{pos} | {role} | {clm} | {locs} | {strs} | {vnums} | {vis} | {tstr} |")
        ev_lines.append("")

        if meta.get("tradeoffs"):
            ev_lines.append("## 2. 关键内容取舍 / Editorial Tradeoffs")
            ev_lines.append("")
            for tr in meta["tradeoffs"]:
                ev_lines.append(f"- {tr}")
            ev_lines.append("")

        # Reuse computed source coverage & starting deck diff sections if present in speaker_notes
        for hdr_marker in ("## 源材料图表覆盖复核 / Source-to-Deck Coverage Audit", "## 起始演示文稿改动复核 / Starting Deck Editorial Diff"):
            if hdr_marker in lines:
                st_i = lines.index(hdr_marker)
                en_i = next((j for j in range(st_i + 1, len(lines)) if lines[j].startswith("## ")), len(lines))
                ev_lines.extend(lines[st_i:en_i])

        if qa_items:
            backup_nums = [s.get("pptx_slide") for s in slides if s.get("role") == "qa_backup" and s.get("pptx_slide")]
            def_bk = backup_nums[0] if backup_nums else None
            ev_lines.append("## 3. 四维问答溯源矩阵 / 4-Dimension Q&A Grounding Matrix")
            ev_lines.append("")
            ev_lines.append("| # | Category | Anticipated Question | Source-Grounded Answer Key | Cite | Backup |")
            ev_lines.append("|---|---|---|---|---|---|")
            for qi, q in enumerate(qa_items, 1):
                cat = str(q.get("category") or "—")
                qq = str(q.get("question") or "").replace("|", "/")[:72]
                ak = str(q.get("answer_key") or "").replace("|", "/")[:120]
                cs = f"P{q['cite_slide']}" if q.get("cite_slide") else "—"
                bs = f"P{q.get('backup_slide') or def_bk}" if (q.get("backup_slide") or def_bk) else "—"
                ev_lines.append(f"| Q{qi} | `{cat}` | {qq} | {ak} | {cs} | {bs} |")
            ev_lines.append("")

        ev_lines.append("## 4. 模板、渲染与版本修改记录 / Template, Rendering & Revision Log")
        ev_lines.append("")
        fonts_d = meta.get("fonts") or {}
        if fonts_d:
            ev_lines.append(f"- **字体 / Fonts**: major=`{fonts_d.get('major') or 'default'}`, minor=`{fonts_d.get('minor') or 'default'}`")
        if rmeta.get("non_embedded_fonts"):
            ev_lines.append(f"- **未嵌入字体告警**: {', '.join(rmeta['non_embedded_fonts'][:8])}")
        if chgs:
            for c_item in chgs:
                ev_lines.append(f"- {c_item}")
        ev_lines.append("")
        Path(a.evidence_map).write_text("\n".join(ev_lines), encoding="utf-8")
        print(f"evidence map -> {a.evidence_map} ({len(slides)} slides)")
    if a.pdf:
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas as _cv
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            import textwrap
            all_text = "\n".join(lines)
            has_cjk = any("\u2e80" <= ch <= "\uffef" for ch in all_text)
            font = "Helvetica"
            if has_cjk:
                try:
                    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
                    font = "STSong-Light"
                except Exception:
                    pass
            if font != "Helvetica":
                # emoji/symbols missing from CID fonts -> plain text markers
                lines = [ln.replace("⏱", "[时间]") for ln in lines]
            pdf_path = Path(a.out).with_suffix(".pdf")
            c = _cv.Canvas(str(pdf_path), pagesize=A4)
            W, H = A4
            y = H - 50
            # CJK glyphs are ~2x wide: halve wrap width when CJK font is active
            wrap_w = 46 if font != "Helvetica" else 88
            for ln in lines:
                for seg in textwrap.wrap(ln, wrap_w) or [""]:
                    if y < 50:
                        c.showPage(); y = H - 50
                    c.setFont(font, 9.5)
                    c.drawString(45, y, seg[:max(60, int(wrap_w * 1.4))]); y -= 13
            c.save()
            print(f"handout pdf -> {pdf_path} (font={font})")
        except Exception as e:
            print(f"warn: pdf export skipped ({e!r}); keep markdown only")


if __name__ == "__main__":
    main()
