#!/usr/bin/env python3
"""Toolchain self-test: verifies this Skill's scripts work in the current environment (~5s).

Usage: python3 selftest.py [--keep]
Builds a throwaway 3-slide deck (title + figure + table + English notes) + vector-PDF crop test
in a temp dir (clearly synthetic, NOT a benchmark fixture), runs lint + extract + result assembly.
Exit 0 = all pass, 1 = any fail.
"""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PY = sys.executable


def run(args):
    return subprocess.run([PY] + args, capture_output=True, text=True, timeout=120)


def main():
    keep = "--keep" in sys.argv
    tmp = Path(tempfile.mkdtemp(prefix="pkselftest_"))
    results = []

    # 1) python-pptx import
    try:
        import pptx  # noqa: F401
        results.append(("python-pptx import", True, ""))
    except ImportError as e:
        results.append(("python-pptx import", False, str(e)))

    # 2) tiny figure via PIL + vector PDF via reportlab (if available)
    fig = tmp / "fig.png"
    try:
        from PIL import Image
        Image.new("RGB", (480, 320), (70, 114, 176)).save(fig)
        results.append(("PIL figure gen", True, ""))
    except ImportError as e:
        results.append(("PIL figure gen", False, f"{e}; deck falls back to text-only"))

    # 3) mini spec with steps + kpis + figure + native table + two_col + English notes
    spec = {
        "meta": {"title": "selftest", "duration_minutes": 4, "claim_on_slide": True,
                 "tradeoffs": ["synthetic test slide selection"]},
        "slides": [
            {"role": "title", "title": "Synthetic Self-Test Deck", "claim": "Toolchain verification",
             "bullets": ["Synthetic Author | Test Venue"],
             "notes": "Welcome everyone. This synthetic title slide verifies cover layout without claim bar clutter.", "time_s": 20},
            {"role": "roadmap", "title": "Talk Roadmap", "claim": "Three steps: gap -> mechanism -> evidence",
             "steps": [
                 {"label": "01", "title": "Gap", "detail": "Accuracy drops at large batch"},
                 {"label": "02", "title": "Method", "detail": "Adaptive preconditioning"},
                 {"label": "03", "title": "Evidence", "detail": "Controlled benchmarks"},
             ],
             "evidence": [{"type": "section", "locator": "§1–§3", "strength": "NA"}],
             "notes": "Our talk follows three clear steps from the large-batch accuracy gap, through our adaptive preconditioning mechanism, to controlled empirical evidence across multiple batch scales. Next, let's examine the main result.", "time_s": 45},
            {"role": "result", "title": "Main Result: Accuracy Gain", "claim": "Method raises top-1 to 76.3% (+3.2pt vs baseline)",
             "kpis": [
                 {"value": "76.3%", "label": "Top-1 @ 8k batch", "delta": "+3.2pt"},
                 {"value": "1.6x", "label": "Wall-clock speedup", "delta": "90 epochs"},
             ],
             "bullets": ["Consistent gain under matched budget", "Zero layout overflow"],
             "figure": {"path": str(fig), "region": "right", "caption": "Synthetic plot"} if fig.exists() else None,
             "evidence": [{"type": "figure", "locator": "Fig.1, p.1", "strength": "S1"}],
             "notes": "Looking at the figure on the right, our method raises top-1 accuracy to 76.3 percent under the exact same compute budget, outperforming the linear scaling baseline by 3.2 points. Next, let's check the ablation table.", "time_s": 60},
            {"role": "analysis", "title": "Ablation Comparison", "claim": "Benefit scales monotonically across batch sizes",
             "bullets": ["Gain widens at larger batch", "No degradation at small batch"],
             "table": {"headers": ["Batch", "Base", "Ours", "Delta"],
                       "rows": [["2048", "75.8", "76.8", "+1.0"], ["8192", "73.1", "76.3", "+3.2"]],
                       "region": "right", "caption": "Table 1 summary", "highlight_row": 1},
             "evidence": [{"type": "table", "locator": "Table 1, p.1", "strength": "S1"}],
             "notes": "In the table on the right, the improvement grows smoothly from 1.0 point at 2048 batch to 3.2 points at 8192 batch, confirming our scaling hypothesis. Next, let's review limitations.", "time_s": 60},
            {"role": "limitation", "title": "Scope and Mitigation", "claim": "Validated on CNNs; Transformer scaling remains open",
             "two_col": {
                 "left_title": "Known Limitations",
                 "left_bullets": ["Validated only on ResNet-50", "Adds 7% step overhead"],
                 "right_title": "Mitigation & Next Steps",
                 "right_bullets": ["Layer-wise schedule for ViT", "Fuse kernel to cut overhead"],
             },
             "evidence": [{"type": "section", "locator": "§4, p.1", "strength": "S3"}],
             "notes": "We keep our scope honest: current validation covers ResNet-50 with a modest 7 percent overhead, and we outline two concrete next steps on the right. Finally, thank you for listening.", "time_s": 50},
        ],
    }
    (tmp / "spec.json").write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    deck = tmp / "selftest.pptx"
    r = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "spec.json"), "--out", str(deck), "--work", str(tmp)])
    results.append(("build_deck (cards+fig+table)", r.returncode == 0 and deck.exists(), r.stderr.strip()[-140:]))

    # 3b) edit_deck mode with prune + reorder + duplicate slide_index safety + preserved notes
    edit_spec = {
        "meta": {"title": "edited", "duration_minutes": 3, "keep_only_spec_slides": True},
        "slides": [
            {"role": "title", "reuse": {"slide_index": 0, "action": "keep"}, "title": "Synthetic Self-Test Deck", "claim": "Kept cover",
             "time_s": 20},
            {"role": "analysis", "reuse": {"slide_index": 3, "action": "keep"}, "title": "Ablation Comparison", "claim": "Reordered ahead",
             "notes": "Reordered ablation slide ahead of limitation slide to test edit_deck pruning and reordering cleanly. Finally, thank you.", "time_s": 55},
            {"role": "analysis", "reuse": {"slide_index": 3, "action": "edit", "clear_visuals": True}, "title": "Split Follow-up Slide",
             "claim": "Duplicate reuse index auto-clones layout safely", "bullets": ["Safety: zero dropped slides on split"],
             "notes": "This second slide intentionally reuses slide_index 3 to verify automatic layout cloning without losing content. Finally, thank you.", "time_s": 45},
        ],
    }
    (tmp / "edit_spec.json").write_text(json.dumps(edit_spec, ensure_ascii=False), encoding="utf-8")
    edited = tmp / "edited.pptx"
    r_ed = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "edit_spec.json"), "--deck", str(deck), "--out", str(edited), "--work", str(tmp / "ed")])
    ed_ok = r_ed.returncode == 0 and edited.exists()
    if ed_ok:
        try:
            from pptx import Presentation as _Prs
            ed_prs = _Prs(str(edited))
            ed_plan = json.loads((tmp / "ed" / "slide_plan.json").read_text(encoding="utf-8"))
            ed_ok = len(ed_prs.slides) == 3 and ed_plan["slides"][0]["notes_chars"] > 10
        except Exception:
            ed_ok = False
    results.append(("edit_deck (prune+split+keep)", ed_ok, (r_ed.stdout + r_ed.stderr).strip()[-140:]))

    # 4) layout + bilingual notes lint (must pass with 0 issues)
    r = run([str(HERE / "check_layout.py"), str(deck), "--json", str(tmp / "lint.json"), "--plan", str(tmp / "slide_plan.json")])
    results.append(("check_layout (bilingual)", r.returncode == 0, (r.stdout + r.stderr).strip()[-160:]))

    # 5) vector PDF figure extraction + text dump (if fitz + reportlab available)
    vec_ok, vec_note = True, ""
    try:
        try:
            import pymupdf as fitz  # noqa: F401
        except ImportError:
            import fitz  # noqa: F401
        from reportlab.pdfgen import canvas as _cv
        from reportlab.lib.pagesizes import letter
        vpdf = tmp / "vec.pdf"
        c = _cv.Canvas(str(vpdf), pagesize=letter)
        c.setStrokeColorRGB(0.1, 0.3, 0.6)
        c.rect(72, 460, 260, 150, stroke=1, fill=0)
        c.line(80, 480, 320, 590)
        c.drawString(72, 440, "Fig. 1 Synthetic vector plot showing scaling trend.")
        c.save()
        r2 = run([str(HERE / "extract_figs.py"), "--pdf", str(vpdf), "--out", str(tmp / "figs"), "--text", str(tmp / "paper.txt")])
        mf = json.loads((tmp / "figs" / "manifest.json").read_text(encoding="utf-8")) if (tmp / "figs" / "manifest.json").exists() else []
        vec_ok = (r2.returncode == 0 and any(m.get("method") == "vector_crop" for m in mf) and (tmp / "paper.txt").exists())
        vec_note = (r2.stdout + r2.stderr).strip()[-140:]
    except ImportError:
        vec_note = "fitz/reportlab absent (advisory skip)"
    results.append(("extract_figs (vector)", vec_ok, vec_note))

    # 6) one-click ingest + deliver pipeline check (skip-render for speed)
    pipe_dir = tmp / "pipe"
    (pipe_dir / "input").mkdir(parents=True, exist_ok=True)
    (pipe_dir / "input" / "source.md").write_text(
        "# 1 Introduction\nBaseline reaches 73.1% at 8192 batch.\n# 2 Results\nMethod reaches 76.3% (+3.2pt) with 1.6x speedup.\n# 3 Limitations\nValidated on ResNet-50 with 7% overhead.\n",
        encoding="utf-8",
    )
    (pipe_dir / "input" / "case.json").write_text(json.dumps({
        "case_id": "selftest-pipe",
        "audience": "defense committee",
        "duration_minutes": 4,
        "language": "en",
        "output_contract": {"deck_path": "/app/output/pipe_talk.pptx", "required_fields": ["case_id", "status", "slide_plan", "evidence_map", "qa_prep", "rendering"]},
    }, ensure_ascii=False), encoding="utf-8")
    r_ing = run([str(HERE / "ingest.py"), "--case", str(pipe_dir / "input" / "case.json"), "--work", str(pipe_dir / "work")])
    pipe_spec = dict(spec)
    pipe_spec["qa_prep"] = [{"question": "What is the overhead?", "answer_key": "7% step overhead (Section 3)", "cite_slide": 3, "backup_slide": None}]
    pipe_spec["slides"] = [
        {"role": "title", "title": "Synthetic Pipeline Test", "claim": "Pipeline check", "bullets": ["Author"], "notes": "Welcome everyone to this quick pipeline verification slide.", "time_s": 20},
        {"role": "result", "title": "Main Result: 76.3% Top-1 (+3.2pt) at 8192 Batch", "claim": "Method reaches 76.3% (+3.2pt) with 1.6x speedup", "bullets": ["Comparison: 76.3% vs 73.1% at 8192 batch"],
         "chart": {"type": "bar", "categories": ["8192"], "series": [{"name": "Baseline", "values": [73.1]}, {"name": "Ours", "values": [76.3]}], "region": "right", "caption": "8192 batch top-1 (%)"},
         "evidence": [{"type": "section", "locator": "§2", "strength": "S1"}],
         "notes": "Our method reaches 76.3 percent at 8192 batch as shown in the right chart, outperforming the 73.1 percent baseline by 3.2 points with a 1.6x speedup. Next, let us review limitations.", "time_s": 60},
        {"role": "limitation", "title": "Scope Boundaries: ResNet-50 and 7% Overhead", "claim": "Validated on ResNet-50 with 7% overhead", "bullets": ["Scope: ResNet-50 vision backbone only", "Cost: 7% step overhead"],
         "evidence": [{"type": "section", "locator": "§3", "strength": "S3"}],
         "notes": "Current validation is explicitly scoped to ResNet-50 vision backbones and introduces roughly 7 percent step overhead during training, which we highlight transparently before moving to our summary and taking audience questions.", "time_s": 50},
        {"role": "qa_backup", "title": "Backup Q&A", "claim": "Overhead reference", "bullets": ["Overhead: 7% step overhead (§3)"],
         "evidence": [{"type": "section", "locator": "§3", "strength": "S3"}],
         "notes": "Backup slide for overhead follow-up questions.", "time_s": None},
    ]
    (pipe_dir / "work" / "deck_spec.json").write_text(json.dumps(pipe_spec, ensure_ascii=False), encoding="utf-8")
    r_del = run([str(HERE / "deliver.py"), "--case", str(pipe_dir / "input" / "case.json"), "--work", str(pipe_dir / "work"), "--out-dir", str(pipe_dir / "out"), "--skip-render"])
    pipe_ok = (r_ing.returncode == 0 and r_del.returncode == 0 and (pipe_dir / "out" / "pipe_talk.pptx").exists() and (pipe_dir / "out" / "result.json").exists())
    if pipe_ok:
        pr = json.loads((pipe_dir / "out" / "result.json").read_text(encoding="utf-8"))
        notes_md = (pipe_dir / "out" / "speaker_notes.md").read_text(encoding="utf-8") if (pipe_dir / "out" / "speaker_notes.md").exists() else ""
        ev_md = (pipe_dir / "out" / "evidence_map.md").read_text(encoding="utf-8") if (pipe_dir / "out" / "evidence_map.md").exists() else ""
        pipe_ok = (
            pr.get("qa_prep", [{}])[0].get("backup_slide") == 4
            and any(s.get("type") == "source_document" for s in pr.get("sources", []))
            and "source_coverage_audit" in pr
            and "visual_diversity_pct" in (pr.get("evidence_coverage") or {})
            and "| Slide | Role |" in notes_md
            and "| Slide | Role | Takeaway Claim |" in ev_md
            and bool((pr.get("deliverables") or {}).get("evidence_map"))
            and (pipe_dir / "work" / "page_budget.json").exists()
            and (pipe_dir / "work" / "reading_map.md").exists()
        )
    results.append(("ingest+deliver+chart+evmap (e2e)", pipe_ok, (r_ing.stderr + r_del.stdout + r_del.stderr).strip()[-180:]))

    # 6c) dual side-by-side figures + enhanced math/ASCII polish (x_t, ℝ^d, >=)
    dual_spec = {
        "meta": {"title": "Dual Fig Math Test", "duration_minutes": 5, "language": "en"},
        "slides": [
            {"role": "result", "title": "Dual Panel Comparison on ℝ^d", "claim": "Score g_t >= 0.18 cuts cost 2.4x",
             "figures": [{"path": str(fig), "caption": "Left panel: Fig.1"}, {"path": str(fig), "caption": "Right panel: Fig.2"}],
             "evidence": [{"type": "figure", "locator": "Figs. 1–2", "strength": "S1"}],
             "notes": "Compare left and right panels showing g_t >= 0.18 achieves a 2.4x reduction.", "time_s": 60}
        ],
    }
    (tmp / "dual_spec.json").write_text(json.dumps(dual_spec, ensure_ascii=False), encoding="utf-8")
    r_dual = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "dual_spec.json"), "--out", str(tmp / "dual.pptx"), "--work", str(tmp)])
    dual_ok = r_dual.returncode == 0 and (tmp / "dual.pptx").exists()
    if dual_ok:
        try:
            from pptx import Presentation as _PrsDual
            from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
            _p_d = _PrsDual(str(tmp / "dual.pptx"))
            _s0 = _p_d.slides[0]
            _n_pics = sum(1 for sh in _s0.shapes if sh.shape_type == _MST.PICTURE)
            _all_t = " ".join(sh.text_frame.text for sh in _s0.shapes if sh.has_text_frame)
            dual_ok = (_n_pics == 2 and "ℝᵈ" in _all_t and "gₜ" in _all_t and "≥" in _all_t)
        except Exception:
            dual_ok = False
    results.append(("build_deck (dual_figs+math)", dual_ok, (r_dual.stdout + r_dual.stderr).strip()[-140:]))

    # 6d) multi-media zero-overlap (fig+tbl without bullets, fig+tbl with bullets, and edit_deck reuse fig+tbl)
    mm_spec = {
        "meta": {"title": "Multi-Media Zero Overlap", "duration_minutes": 6, "language": "en"},
        "slides": [
            {"role": "result", "title": "Side-by-Side Figure and Table",
             "claim": "Zero overlap when figure and table coexist without bullets",
             "figure": {"path": str(fig), "caption": "Fig. 1: Synthetic curve"},
             "table": {"headers": ["Method", "Acc", "Lat"], "rows": [["Base", "80.1%", "40ms"], ["Ours", "84.5%", "22ms"]], "caption": "Table 1: Main results"},
             "evidence": [{"type": "figure", "locator": "Fig. 1, Table 1", "strength": "S1"}],
             "notes": "Left panel shows Fig. 1 and right panel shows Table 1 with zero overlap.", "time_s": 60},
            {"role": "analysis", "title": "Bullets Above Side-by-Side Media",
             "claim": "Top bullet band sits cleanly above paired figure and table",
             "bullets": ["Top band summarizes key latency vs accuracy tradeoff cleanly.", "Bottom-left figure and bottom-right table stay completely disjoint."],
             "figure": {"path": str(fig), "caption": "Fig. 2: Latency profile"},
             "table": {"headers": ["Setting", "Gain"], "rows": [["Static", "+1.2%"], ["Ours", "+4.4%"]], "caption": "Table 2: Ablation"},
             "evidence": [{"type": "table", "locator": "Fig. 2, Table 2", "strength": "S1"}],
             "notes": "Read the top bullets first, then inspect Fig. 2 on the bottom left and Table 2 on the bottom right.", "time_s": 60},
        ],
    }
    (tmp / "mm_spec.json").write_text(json.dumps(mm_spec, ensure_ascii=False), encoding="utf-8")
    r_mm = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "mm_spec.json"), "--out", str(tmp / "mm.pptx"), "--work", str(tmp)])
    r_mm_lint = run([str(HERE / "check_layout.py"), str(tmp / "mm.pptx"), "--json", str(tmp / "mm_lint.json")])
    # Also verify edit_deck (is_reuse) with fig+tbl does NOT delete the newly inserted figure when adding the table
    mm_reuse_spec = {
        "meta": {"title": "Reuse Multi-Media", "duration_minutes": 5, "language": "en", "keep_only_spec_slides": True},
        "slides": [
            {"reuse": {"slide_index": 0, "action": "edit", "clear_visuals": False},
             "role": "result", "title": "Reused Slide With Both Fig and Table",
             "claim": "Single-pass clear_old preserves both newly inserted media items",
             "figure": {"path": str(fig), "caption": "New Fig. 1"},
             "table": {"headers": ["Model", "Score"], "rows": [["Old", "70.0"], ["New", "78.5"]], "caption": "New Table 1"},
             "notes": "Both new figure and new table survive reuse.", "time_s": 60},
        ],
    }
    (tmp / "mm_reuse_spec.json").write_text(json.dumps(mm_reuse_spec, ensure_ascii=False), encoding="utf-8")
    r_mm_reuse = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "mm_reuse_spec.json"), "--deck", str(tmp / "mm.pptx"), "--out", str(tmp / "mm_reuse.pptx"), "--work", str(tmp)])
    r_mm_reuse_lint = run([str(HERE / "check_layout.py"), str(tmp / "mm_reuse.pptx"), "--json", str(tmp / "mm_reuse_lint.json")])
    mm_ok = (r_mm.returncode == 0 and r_mm_lint.returncode == 0 and r_mm_reuse.returncode == 0 and r_mm_reuse_lint.returncode == 0)
    if mm_ok:
        try:
            from pptx import Presentation as _PrsMM
            from pptx.enum.shapes import MSO_SHAPE_TYPE as _MSTMM
            _pm = _PrsMM(str(tmp / "mm.pptx"))
            _s0_pic = [sh for sh in _pm.slides[0].shapes if sh.shape_type == _MSTMM.PICTURE]
            _s0_tbl = [sh for sh in _pm.slides[0].shapes if getattr(sh, "has_table", False)]
            _s1_pic = [sh for sh in _pm.slides[1].shapes if sh.shape_type == _MSTMM.PICTURE]
            _s1_tbl = [sh for sh in _pm.slides[1].shapes if getattr(sh, "has_table", False)]
            _pr = _PrsMM(str(tmp / "mm_reuse.pptx"))
            _sr_pic = [sh for sh in _pr.slides[0].shapes if sh.shape_type == _MSTMM.PICTURE]
            _sr_tbl = [sh for sh in _pr.slides[0].shapes if getattr(sh, "has_table", False)]
            mm_ok = (
                len(_s0_pic) == 1 and len(_s0_tbl) == 1 and (_s0_pic[0].left + _s0_pic[0].width) <= _s0_tbl[0].left
                and len(_s1_pic) == 1 and len(_s1_tbl) == 1 and (_s1_pic[0].left + _s1_pic[0].width) <= _s1_tbl[0].left
                and len(_sr_pic) == 1 and len(_sr_tbl) == 1
            )
        except Exception:
            mm_ok = False
    results.append(("build_deck (multi-media+reuse)", mm_ok, (r_mm.stdout + r_mm_lint.stdout + r_mm_reuse_lint.stdout).strip()[-160:]))

    # 6e) dark bg detection (full-bleed PICTURE, blipFill, gradFill, schemeClr+lumMod/shade, Z-order) + >=65% bg lint exemption
    dark_ok = False
    dark_note = ""
    try:
        import importlib.util as _ilu
        from PIL import Image as _PImg
        from pptx import Presentation as _PrsDk
        from pptx.util import Inches as _In
        from pptx.oxml import parse_xml as _pxml
        from pptx.oxml.ns import nsdecls as _nsd, qn as _qn
        _spec_bd = _ilu.spec_from_file_location("bd_mod", str(HERE / "build_deck.py"))
        _bd = _ilu.module_from_spec(_spec_bd)
        _spec_bd.loader.exec_module(_bd)

        dk_img = tmp / "dark_bg.png"
        _PImg.new("RGB", (640, 360), (16, 28, 48)).save(str(dk_img))
        lt_img = tmp / "light_bg.png"
        _PImg.new("RGB", (640, 360), (242, 246, 250)).save(str(lt_img))

        _pdk = _PrsDk()
        _pdk.slide_width, _pdk.slide_height = _In(13.333), _In(7.5)
        # Slide 1: dark full-bleed PICTURE shape (covering 76% of slide) -> must detect dark & exempt overlap
        _sl1 = _pdk.slides.add_slide(_pdk.slide_layouts[0])
        _pic_bg = _sl1.shapes.add_picture(str(dk_img), _In(0.4), _In(0.4), width=_In(12.2), height=_In(6.2))
        _sl1.shapes._spTree.remove(_pic_bg._element)
        _sl1.shapes._spTree.insert(2, _pic_bg._element)
        _d1 = _bd.detect_slide_dark_bg(_sl1, {}, _pdk.slide_width, _pdk.slide_height)

        # Slide 2: dark <p:bg><p:bgPr><a:gradFill> multi-stop gradient
        _sl2 = _pdk.slides.add_slide(_pdk.slide_layouts[1])
        _sl2._element.cSld.insert(0, _pxml(
            f'<p:bg {_nsd("p", "a")}><p:bgPr><a:gradFill><a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="0F172A"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="1E3A5F"/></a:gs>'
            f'</a:gsLst></a:gradFill></p:bgPr></p:bg>'
        ))
        _d2 = _bd.detect_slide_dark_bg(_sl2, {}, _pdk.slide_width, _pdk.slide_height)

        # Slide 3: dark <a:schemeClr val="bg1"><a:lumMod val="20000"/> modifier
        _sl3 = _pdk.slides.add_slide(_pdk.slide_layouts[1])
        _sl3._element.cSld.insert(0, _pxml(
            f'<p:bg {_nsd("p", "a")}><p:bgPr><a:solidFill>'
            f'<a:schemeClr val="bg1"><a:lumMod val="20000"/></a:schemeClr>'
            f'</a:solidFill></p:bgPr></p:bg>'
        ))
        _d3 = _bd.detect_slide_dark_bg(_sl3, {"lt1": (255, 255, 255)}, _pdk.slide_width, _pdk.slide_height)

        # Slide 4: dark <p:bg><p:bgPr><a:blipFill> via related image part
        _sl4 = _pdk.slides.add_slide(_pdk.slide_layouts[1])
        _, _rid4 = _sl4.part.get_or_add_image_part(str(dk_img))
        _sl4._element.cSld.insert(0, _pxml(
            f'<p:bg {_nsd("p", "a", "r")}><p:bgPr><a:blipFill>'
            f'<a:blip r:embed="{_rid4}"/><a:stretch><a:fillRect/></a:stretch>'
            f'</a:blipFill></p:bgPr></p:bg>'
        ))
        _d4 = _bd.detect_slide_dark_bg(_sl4, {}, _pdk.slide_width, _pdk.slide_height)

        # Slide 5: light full-bleed PICTURE overriding dark gradient master (Z-order: slide > master)
        _sl5 = _pdk.slides.add_slide(_pdk.slide_layouts[1])
        _sl5._element.cSld.insert(0, _pxml(
            f'<p:bg {_nsd("p", "a")}><p:bgPr><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></p:bgPr></p:bg>'
        ))
        _sl5.shapes.add_picture(str(lt_img), _In(0.0), _In(0.0), width=_In(13.333), height=_In(7.5))
        _d5 = _bd.detect_slide_dark_bg(_sl5, {}, _pdk.slide_width, _pdk.slide_height)

        _pdk.save(str(tmp / "dark_tpl.pptx"))
        # Build on dark_tpl via edit_deck (reuse slide 0 with dark full-bleed photo) and verify light text + 0 overlap lint
        dk_spec = {
            "meta": {"title": "Dark Photo Cover", "duration_minutes": 5, "language": "en", "keep_only_spec_slides": True},
            "slides": [
                {"reuse": {"slide_index": 0, "action": "edit", "clear_visuals": False},
                 "role": "title", "title": "Readable Light Title Over Dark Cover Photo",
                 "bullets": ["Author · Institute · 2026"], "notes": "Cover over dark full-bleed photo.", "time_s": 25},
                {"reuse": {"slide_index": 1, "action": "edit", "clear_visuals": False},
                 "role": "result", "title": "Dark Gradient Slide Contrast Check",
                 "claim": "Gradient background triggers light typography automatically",
                 "bullets": ["Light body text over dark gradient background."], "notes": "Gradient dark bg check.", "time_s": 60},
            ],
        }
        (tmp / "dk_spec.json").write_text(json.dumps(dk_spec, ensure_ascii=False), encoding="utf-8")
        r_dk_b = run([str(HERE / "build_deck.py"), "--spec", str(tmp / "dk_spec.json"), "--deck", str(tmp / "dark_tpl.pptx"), "--out", str(tmp / "dk_out.pptx"), "--work", str(tmp)])
        r_dk_l = run([str(HERE / "check_layout.py"), str(tmp / "dk_out.pptx"), "--json", str(tmp / "dk_lint.json")])
        _pdk_out = _PrsDk(str(tmp / "dk_out.pptx"))
        _t0_rgb = _pdk_out.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb
        _t1_rgb = _pdk_out.slides[1].shapes.title.text_frame.paragraphs[0].runs[0].font.color.rgb
        _t0_lum = _bd.luminance((_t0_rgb[0], _t0_rgb[1], _t0_rgb[2])) if _t0_rgb else 0
        _t1_lum = _bd.luminance((_t1_rgb[0], _t1_rgb[1], _t1_rgb[2])) if _t1_rgb else 0
        dark_ok = (_d1 is True and _d2 is True and _d3 is True and _d4 is True and _d5 is False
                   and r_dk_b.returncode == 0 and r_dk_l.returncode == 0 and _t0_lum >= 160 and _t1_lum >= 160)
        dark_note = f"flags={(_d1,_d2,_d3,_d4,_d5)} lum=({_t0_lum:.0f},{_t1_lum:.0f}) lint={r_dk_l.returncode}"
    except Exception as _e_dk:
        dark_note = repr(_e_dk)
    results.append(("dark_bg (pic+blip+grad+mod+lint)", dark_ok, dark_note))

    # 6b) exact numeric token + derived delta + scientific notation + unit conversion + Roman/range evidence + Q&A audit
    (tmp / "num_src.txt").write_text(
        "--- PAGE 1 ---\nIII. METHOD\nFigure 1: Pipeline. Figure 2: Scaling. TABLE II: Results.\n"
        "Total samples 1420 with accuracy 76.3% vs baseline 73.1%, lr=3e-4, latency 450ms, cuts 1/4 memory (Eq. (1), Algorithm 1).\n",
        encoding="utf-8",
    )
    (tmp / "num_spec.json").write_text(json.dumps({"slides": [{"title": "Test", "claim": "Accuracy reaches 42% on subset"}]}), encoding="utf-8")
    r_num = run([str(HERE / "check_numbers.py"), "--spec", str(tmp / "num_spec.json"), "--source", str(tmp / "num_src.txt")])
    (tmp / "num_spec_ok.json").write_text(json.dumps({
        "slides": [{"role": "result", "title": "Gain", "claim": "76.3% vs 73.1% (+3.2pt), lr=3e-4, 0.45s, 25% memory cut",
                    "evidence": [{"type": "table", "locator": "Table 2, Figs. 1–2, §3, Eq. (1), Alg. 1", "strength": "S1"}]}],
        "qa_prep": [{"question": "Latency & gain?", "answer_key": "0.45s (450ms) on 8xA100 FP16 at lr=3e-4, +3.2pt vs Eq. (1), bogus 99.9%", "cite_slide": 1}],
    }), encoding="utf-8")
    (tmp / "ev_spec_bad.json").write_text(json.dumps({
        "slides": [{"role": "result", "title": "Bad", "evidence": [{"type": "table", "locator": "Table 5, Figs. 1–9", "strength": "S1"}]}]
    }), encoding="utf-8")
    r_num_ok = run([str(HERE / "check_numbers.py"), "--spec", str(tmp / "num_spec_ok.json"), "--source", str(tmp / "num_src.txt"), "--out", str(tmp / "num_ok.json")])
    r_ev_ok = run([str(HERE / "check_evidence.py"), "--spec", str(tmp / "num_spec_ok.json"), "--source", str(tmp / "num_src.txt")])
    r_ev_bad = run([str(HERE / "check_evidence.py"), "--spec", str(tmp / "ev_spec_bad.json"), "--source", str(tmp / "num_src.txt")])
    qa_caught = False
    if (tmp / "num_ok.json").exists():
        try:
            q_sus = json.loads((tmp / "num_ok.json").read_text(encoding="utf-8")).get("qa_suspects") or []
            qa_caught = (len(q_sus) == 1 and q_sus[0].get("value") == "99.9")
        except Exception:
            qa_caught = False
    results.append(("check_numbers+ev (sci+unit+roman+qa)", r_num.returncode == 1 and r_num_ok.returncode == 0 and r_ev_ok.returncode == 0 and r_ev_bad.returncode == 1 and qa_caught,
                    (r_num.stdout + r_num_ok.stdout + r_ev_ok.stdout + r_ev_bad.stdout).strip()[-140:]))

    # 7) audit_talk + result assembly + soffice availability
    r_aud = run([str(HERE / "audit_talk.py"), "--work", str(pipe_dir / "work"), "--out-dir", str(pipe_dir / "out"),
                 "--spec", str(pipe_dir / "work" / "deck_spec.json"), "--out", str(pipe_dir / "work" / "quality_audit.json")])
    aud_ok = (r_aud.returncode == 0 and (pipe_dir / "work" / "quality_audit.json").exists())
    if aud_ok:
        qa_j = json.loads((pipe_dir / "work" / "quality_audit.json").read_text(encoding="utf-8"))
        aud_ok = "editorial_audit" in qa_j and "pixel_audit" in qa_j and "readiness_checks" in qa_j
    results.append(("audit_talk (pixel+editorial)", aud_ok, (r_aud.stdout + r_aud.stderr).strip()[-120:]))

    (tmp / "qa_plan.json").write_text("[]", encoding="utf-8")
    r = run([str(HERE / "build_result.py"), "--work", str(tmp), "--out", str(tmp / "result.json")])
    ok = r.returncode == 0 and (tmp / "result.json").exists()
    results.append(("build_result", ok, r.stderr.strip()[-120:]))

    has_soffice = bool(shutil.which("soffice") or shutil.which("libreoffice"))
    results.append(("soffice (render)", has_soffice, "advisory: export via GUI if missing"))

    print("== selftest ==")
    fail = 0
    for name, ok_, note in results:
        mark = "PASS" if ok_ else "FAIL"
        if not ok_:
            fail += 1
        print(f"  [{mark}] {name}" + (f" :: {note}" if note and not ok_ else ""))
    print(f"== {len(results) - fail}/{len(results)} components pass ==" + (f" (temp: {tmp}, kept)" if keep else ""))
    if not keep:
        shutil.rmtree(tmp, ignore_errors=True)
    sys.exit(1 if fail else 0)


if __name__ == "__main__":
    main()
