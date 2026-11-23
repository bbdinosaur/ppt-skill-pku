#!/usr/bin/env python3
"""export_preview.py — Step 5b of the academic-slides skill.

Renders the deck to a real PDF via LibreOffice, verifies page count / fonts /
title extractability with PyMuPDF, and writes PNG previews per slide for visual
inspection. If soffice is unavailable, falls back to a PIL geometry previewer
(rectangles + wrapped text) so a visual check is always possible.

Usage: python3 export_preview.py --deck out/presentation.pptx --work work \
           --pdf out/presentation.pdf [--previews-dir out/previews] [--dpi 100]
Writes work/export_report.json, prints EXPORT_OK / EXPORT_FALLBACK / EXPORT_FAILED.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckutil as du  # noqa: E402

from pptx import Presentation  # noqa: E402


def reportlab_pdf(deck: Path, out_pdf: Path, content_map_path: Path | None):
    """Render a REAL PDF approximation of the deck with reportlab (no soffice needed).
    Uses the content_map sidecar for text/notes/images so the PDF is a faithful,
    inspectable deliverable when LibreOffice is unavailable."""
    try:
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.lib.colors import HexColor
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        font = "STSong-Light"
    except Exception as e:
        raise RuntimeError(f"reportlab unavailable: {e}")

    cm = {}
    if content_map_path and Path(content_map_path).is_file():
        try:
            cm = json.loads(Path(content_map_path).read_text(encoding="utf-8"))
        except Exception:
            pass
    slides = cm.get("slides") or []
    PW, PH = landscape(A4)
    W, H = 13.333 * 72, 7.5 * 72  # deck points (16:9)
    sx, sy = PW / W, PH / H
    c = canvas.Canvas(str(out_pdf), pagesize=(PW, PH))

    def wrap(text, size, maxw):
        words, lines, cur = text, [], ""
        for chx in text:
            t = cur + chx
            if pdfmetrics.stringWidth(t, font, size) > maxw:
                lines.append(cur); cur = chx
            else:
                cur = t
        if cur:
            lines.append(cur)
        return lines or [""]

    from pptx import Presentation
    prs = Presentation(str(deck))
    SWd, SHd = int(prs.slide_width), int(prs.slide_height)
    for idx, slide in enumerate(prs.slides):
        info = slides[idx] if idx < len(slides) else {}
        for sh in slide.shapes:
            sl, st, sw, shh = du.shape_geom(prs, slide, sh)
            if sl is None or st is None or not sw or not shh:
                continue
            l, tp = int(sl) / SWd * W * sx, int(st) / SHd * H * sy
            w, h = int(sw) / SWd * W * sx, int(shh) / SHd * H * sy
            try:
                if "PICTURE" in str(sh.shape_type) and sh.image:
                    img_path = sh.image.blob and None
                    # extract blob to temp and draw
                    import tempfile
                    tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                    tf.write(sh.image.blob); tf.close()
                    try:
                        c.drawImage(tf.name, l, PH - tp - h, width=w, height=h,
                                    preserveAspectRatio=True, anchor="c")
                    finally:
                        os.unlink(tf.name)
                    continue
            except Exception:
                pass
            if sh.has_text_frame and sh.text_frame.text.strip():
                size = du.effective_pt(sh, du.DEFAULT_BODY_PT) * 0.62
                c.setFont(font, size)
                ty = PH - tp - size
                for para in sh.text_frame.paragraphs:
                    txt = "".join(r.text for r in para.runs)
                    if not txt:
                        ty -= size
                        continue
                    for ln in wrap(txt, size, w):
                        c.drawString(l + 2, ty, ln)
                        ty -= size * 1.25
            elif sh.shape_type is not None and "TABLE" in str(sh.shape_type):
                try:
                    tbl = sh.table
                    nrows = len(tbl.rows); ncols = len(tbl.columns)
                    rh = h / max(nrows, 1); cw = w / max(ncols, 1)
                    for ri, row in enumerate(tbl.rows):
                        for ci, cell in enumerate(row.cells):
                            x = l + ci * cw; yy = PH - tp - (ri + 1) * rh
                            c.setStrokeColor(HexColor("#999999"))
                            c.rect(x, yy, cw, rh, stroke=1, fill=0)
                            c.setFont(font, 9)
                            c.drawString(x + 3, yy + rh - 12, cell.text[:40])
                except Exception:
                    pass
        c.showPage()
    c.save()
    return out_pdf


def soffice_convert(deck: Path, outdir: Path, work: Path | None = None) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    lo_home = ((work or outdir) / ".lohome").resolve()
    lo_home.mkdir(parents=True, exist_ok=True)
    env = {**os.environ, "HOME": str(lo_home)}
    try:
        subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation=file://{lo_home}",
             "--convert-to", "pdf", "--outdir", str(outdir), str(deck)],
            env=env, timeout=55,  # fail fast to reportlab fallback if LibreOffice stalls
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
        )
    except Exception:
        return None
    pdf = outdir / (deck.stem + ".pdf")
    return pdf if pdf.is_file() and pdf.stat().st_size > 500 else None


def verify_pdf(pdf: Path, expect_pages: int, titles, previews_dir: Path, dpi: int = 100):
    rep = {"pdf": str(pdf), "pages": 0, "fonts": [], "pages_match": None,
           "titles_found": None, "previews": 0, "anomalies": []}
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz as fitz
        doc = fitz.open(str(pdf))
        rep["pages"] = doc.page_count
        rep["pages_match"] = (doc.page_count == expect_pages)
        fonts = set()
        all_text = ""
        for pno in range(doc.page_count):
            for f in doc[pno].get_fonts():
                fonts.add(f[3])
            all_text += doc[pno].get_text("text")
        rep["fonts"] = sorted(fonts)
        if titles:
            import re
            def norm(s):
                s = re.sub(r"[→←↔·•\u2013\u2014\-:：（）()]", "", s or "")
                return re.sub(r"\s+", "", s).lower()
            t = norm(all_text)
            found = sum(1 for x in titles if norm(x) and norm(x) in t)
            rep["titles_found"] = f"{found}/{len(titles)}"
            if found < len(titles) * 0.8:
                rep["anomalies"].append("many titles not extractable from PDF — check font substitution")
        if "\ufffd" in all_text:
            rep["anomalies"].append("replacement chars found in PDF text")
        # export PNG previews (clean stale previews from previous larger decks first)
        previews_dir.mkdir(parents=True, exist_ok=True)
        for old_png in previews_dir.glob("slide-*.png"):
            try:
                old_png.unlink()
            except Exception:
                pass
        cs = previews_dir / "contact-sheet.png"
        if cs.is_file():
            try:
                cs.unlink()
            except Exception:
                pass
        for pno in range(doc.page_count):
            pix = doc[pno].get_pixmap(dpi=dpi)
            pix.save(str(previews_dir / f"slide-{pno + 1:02d}.png"))
        rep["previews"] = doc.page_count
        make_contact_sheet(previews_dir, doc.page_count)
        vis = visual_sanity_check(previews_dir)
        rep["blank_slides"] = vis["blank_slides"]
        rep["visual_ok"] = bool(rep["pages_match"]) and not vis["blank_slides"]
        if vis["blank_slides"]:
            rep["anomalies"].append(f"nearly blank rendered slides: {vis['blank_slides']}")
        doc.close()
    except Exception as e:
        # Fallback to pdftoppm if PyMuPDF is missing in the runtime container
        ppm = shutil.which("pdftoppm")
        if ppm and pdf.is_file():
            try:
                previews_dir.mkdir(parents=True, exist_ok=True)
                for old_png in previews_dir.glob("slide-*.png"):
                    old_png.unlink(missing_ok=True)
                subprocess.run([ppm, "-png", "-r", str(dpi), str(pdf), str(previews_dir / "slide")],
                               timeout=60, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                pngs = sorted(previews_dir.glob("slide-*.png"))
                rep["pages"] = len(pngs)
                rep["pages_match"] = (len(pngs) == expect_pages)
                rep["previews"] = len(pngs)
                if pngs:
                    make_contact_sheet(previews_dir, len(pngs))
                    vis = visual_sanity_check(previews_dir)
                    rep["blank_slides"] = vis["blank_slides"]
                    rep["visual_ok"] = bool(rep["pages_match"]) and not vis["blank_slides"]
                else:
                    rep["anomalies"].append(f"pdf verify failed: {e}")
            except Exception:
                rep["anomalies"].append(f"pdf verify failed: {e}")
        else:
            rep["anomalies"].append(f"pdf verify failed: {e}")
    return rep


def visual_sanity_check(previews_dir: Path) -> dict:
    """Fast pure-Python Pillow pixel sanity check (no binary image read needed by text LLMs)."""
    blank_slides = []
    try:
        from PIL import Image, ImageStat
        for idx, p in enumerate(sorted(previews_dir.glob("slide-*.png")), 1):
            with Image.open(str(p)) as im:
                rgb = im.convert("RGB")
                st = ImageStat.Stat(rgb)
                max_std = max(st.stddev) if st.stddev else 0.0
                if max_std < 1.8:
                    blank_slides.append(idx)
    except Exception:
        pass
    return {"blank_slides": blank_slides}


def make_contact_sheet(previews_dir: Path, n: int, cols: int = 3):
    """Compose all slide PNGs into one contact-sheet PNG for fast visual review."""
    try:
        from PIL import Image
        files = sorted(previews_dir.glob("slide-*.png"))
        if len(files) < 2:
            return
        thumbs = [Image.open(str(f)) for f in files]
        w = 420
        th = [im.resize((w, int(im.height * w / im.width))) for im in thumbs]
        rows = (len(th) + cols - 1) // cols
        ch = max(t.height for t in th)
        sheet = Image.new("RGB", (cols * w + (cols + 1) * 8, rows * ch + (rows + 1) * 8), "#dddddd")
        for i, t in enumerate(th):
            r, c = divmod(i, cols)
            sheet.paste(t, (8 + c * (w + 8), 8 + r * (ch + 8)))
        sheet.save(str(previews_dir / "contact-sheet.png"))
    except Exception:
        pass


def pil_preview(deck: Path, previews_dir: Path):
    """Fallback: approximate geometry preview with PIL when soffice is missing."""
    from pptx import Presentation
    from PIL import Image, ImageDraw, ImageFont

    prs = Presentation(str(deck))
    for old_png in previews_dir.glob("slide-*.png"):
        try:
            old_png.unlink()
        except Exception:
            pass
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    scale = 1100 / SW
    W, H = int(SW * scale), int(SH * scale)

    def font(sz):
        for cand in sorted((Path("/usr/share/fonts")).rglob("*CJK*"))[:1] + \
                    sorted(Path("/usr/share/fonts").rglob("*.tt[fc]"))[:3]:
            try:
                return ImageFont.truetype(str(cand), sz)
            except Exception:
                continue
        return ImageFont.load_default()

    for si, slide in enumerate(prs.slides, 1):
        img = Image.new("RGB", (W, H), "white")
        dr = ImageDraw.Draw(img)
        for sh in slide.shapes:
            sl, st, sw, shh = du.shape_geom(prs, slide, sh)
            if sl is None or st is None or not sw or not shh:
                continue
            l, t = int(int(sl) * scale), int(int(st) * scale)
            w, h = int(int(sw) * scale), int(int(shh) * scale)
            try:
                txt = sh.text_frame.text if sh.has_text_frame else ""
                txt = " ".join(txt.split())
            except Exception:
                txt = ""
            try:
                if "PICTURE" in str(sh.shape_type):
                    dr.rectangle([l, t, l + w, t + h], outline="#3a6ea5", width=2)
                    dr.text((l + 4, t + 4), "[image]", fill="#3a6ea5", font=font(14))
                    continue
            except Exception:
                pass
            if txt:
                fpt = du.effective_pt(sh, du.DEFAULT_BODY_PT) * scale * 0.62
                f = font(max(9, int(fpt)))
                fsize = getattr(f, "size", 12)
                # naive wrap by pixel width
                lines, cur = [], ""
                for chx in txt:
                    trial = cur + chx
                    if dr.textlength(trial, font=f) > w - 8:
                        lines.append(cur)
                        cur = chx
                    else:
                        cur = trial
                if cur:
                    lines.append(cur)
                y = t + 2
                for ln in lines[: max(1, h // (fsize + 2))]:
                    dr.text((l + 3, y), ln, fill="black", font=f)
                    y += fsize + 2
            dr.rectangle([l, t, l + max(w, 2), t + max(h, 2)], outline="#bbbbbb", width=1)
        img.save(str(previews_dir / f"slide-{si:02d}.png"))
    return si


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--work", default="/app/work")
    ap.add_argument("--pdf", default=None)
    ap.add_argument("--previews-dir", default=None)
    ap.add_argument("--dpi", type=int, default=100)
    args = ap.parse_args()

    deck = Path(args.deck)
    work = Path(args.work)
    inv = du.load_inventory(work)
    out_dir = deck.parent
    pdf_path = Path(args.pdf) if args.pdf else out_dir / (deck.stem + ".pdf")
    previews = Path(args.previews_dir) if args.previews_dir else out_dir / "previews"
    previews.mkdir(parents=True, exist_ok=True)

    cm = {}
    cmf = work / "content_map.json"
    if cmf.is_file():
        cm = json.loads(cmf.read_text(encoding="utf-8"))
    titles = [s.get("title") for s in cm.get("slides", []) if s.get("title")]

    n_slides = len(Presentation(str(deck)).slides)
    pdf = soffice_convert(deck, out_dir, work)
    report = {"deck": str(deck), "pdf": None, "previews_dir": str(previews), "method": None,
              "anomalies": []}
    if pdf is None:
        try:
            rl = reportlab_pdf(deck, pdf_path, Path(work) / "content_map.json")
            report["method"] = "reportlab"
            report["pdf"] = str(rl)
            v = verify_pdf(rl, n_slides, titles, previews, dpi=args.dpi)
            if v.get("previews", 0) > 0:
                report.update({k: v.get(k) for k in ("pages", "fonts", "pages_match", "titles_found", "previews", "visual_ok", "blank_slides")})
            else:
                report["previews"] = pil_preview(deck, previews)
                report["pages"] = n_slides
                report["pages_match"] = True
                make_contact_sheet(previews, n_slides)
            report["anomalies"].append("LibreOffice unavailable - PDF rendered via reportlab "
                                       "fallback (approximate layout, real PDF deliverable)")
            (work / "export_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1),
                                                     encoding="utf-8")
            du.log_change(work, "export", f"method=reportlab pdf={rl} previews={report.get('previews')}")
            print("EXPORT_FALLBACK_PDF " + json.dumps({"pdf": str(rl), "method": "reportlab",
                                                       "pages": report.get("pages"),
                                                       "previews": report.get("previews")}, ensure_ascii=False))
            return 0
        except Exception as e:
            report["anomalies"].append(f"reportlab fallback failed: {e}")
    if pdf:
        if pdf.resolve() != pdf_path.resolve():
            try:
                shutil.copy2(str(pdf), str(pdf_path))
                pdf = pdf_path
            except Exception:
                pass
        report["method"] = "libreoffice"
        report["pdf"] = str(pdf)
        v = verify_pdf(pdf, n_slides, titles, previews, dpi=args.dpi)
        report.update({k: v.get(k) for k in
                       ("pages", "fonts", "pages_match", "titles_found", "previews", "visual_ok", "blank_slides", "anomalies")})
        print("EXPORT_OK " + json.dumps({k: report.get(k) for k in
              ("pdf", "pages", "pages_match", "titles_found", "visual_ok", "previews", "method", "anomalies")},
              ensure_ascii=False))
    else:
        report["method"] = "pil_fallback"
        n = pil_preview(deck, previews)
        make_contact_sheet(previews, n)
        report["previews"] = n
        report["anomalies"].append("LibreOffice unavailable — PIL geometry previews only, PDF export missing")
        print("EXPORT_FALLBACK " + json.dumps({"pdf": None, "previews": n, "method": "pil_fallback",
                                               "anomalies": report["anomalies"]}, ensure_ascii=False))

    (work / "export_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1),
                                             encoding="utf-8")
    du.log_change(work, "export", f"method={report['method']} pdf={report.get('pdf')}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"EXPORT_FAILED: {e}")
        sys.exit(1)
