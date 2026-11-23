#!/usr/bin/env python3
"""Export deck -> PDF via LibreOffice (isolated profile + auto-retry), render pages to PNG
(default 108 DPI), and verify:
  (1) PDF page count == slide count & PDF page dimensions match PPTX slide size,
  (2) per-page title text survives export (ASCII/CJK normalized; largest-font fallback),
  (3) dark-background-safe blank-page / missing-figure detection,
  (4) real font-substitution & missing-glyph audit (classifying metric-equivalent vs fallback).

Usage:
  python3 export_render.py DECK.pptx --outdir /app/output/previews [--pdf-out /app/output/deck.pdf] [--dpi 108]
Exit codes: 0 = verified, 1 = mismatch detected, 2 = export failed.

No-script fallback:
  soffice --headless --convert-to pdf --outdir /app/output DECK.pptx
  pdftoppm -png -r 108 /app/output/DECK.pdf /app/output/previews/page
"""
import argparse, json, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

# Common Linux/LibreOffice metric-compatible or CJK family equivalents
FONT_EQUIV = {
    "calibri": ("carlito", "liberationsans", "dejavusans"),
    "cambria": ("caladea", "liberationserif"),
    "arial": ("liberationsans", "arimo", "notosans"),
    "helvetica": ("liberationsans", "arimo", "nimbus"),
    "timesnewroman": ("liberationserif", "tinos", "notoserif"),
    "georgia": ("liberationserif", "gelasio", "dejavuserif"),
    "consolas": ("liberationmono", "dejavusansmono", "inconsolata"),
    "couriernew": ("liberationmono", "cousine"),
    "simsun": ("notoserifcjk", "notosanscjk", "arpl"),
    "simhei": ("notosanscjk", "wenquanyi", "sourcehansans"),
    "microsoftyahei": ("notosanscjk", "wenquanyi", "sourcehansans"),
    "dengxian": ("notosanscjk", "wenquanyi"),
    "pingfangsc": ("notosanscjk", "wenquanyi"),
    "kaiti": ("notoserifcjk", "arpl", "notosanscjk"),
}


def run_soffice(deck: Path, outdir: Path, timeout=180, retries=1):
    """Run LibreOffice headless export with an isolated UserInstallation profile + retry."""
    outdir.mkdir(parents=True, exist_ok=True)
    last_err = ""
    for attempt in range(retries + 1):
        prof_dir = Path(tempfile.mkdtemp(prefix="lo_prof_"))
        try:
            cmd = [
                "soffice",
                f"-env:UserInstallation=file://{prof_dir}",
                "--headless", "--norestore", "--nofirststartwizard",
                "--convert-to", "pdf",
                "--outdir", str(outdir),
                str(deck),
            ]
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            pdf = outdir / (deck.stem + ".pdf")
            if pdf.exists() and pdf.stat().st_size > 200:
                return pdf, (r.stderr[-200:] if r.returncode else ""), attempt
            last_err = f"attempt {attempt + 1}: soffice rc={r.returncode} stderr={r.stderr[-300:]}"
        except subprocess.TimeoutExpired:
            last_err = f"attempt {attempt + 1}: soffice timeout after {timeout}s"
        except FileNotFoundError:
            return None, "soffice binary not found in PATH", attempt
        finally:
            shutil.rmtree(prof_dir, ignore_errors=True)
    return None, last_err, retries


def norm(s):
    """Normalize text for title survival matching (keeps CJK + alphanumeric, drops punctuation)."""
    return "".join(c for c in s.lower() if c.isalnum())


def read_deck_theme_fonts(deck: Path):
    fonts = {}
    try:
        with zipfile.ZipFile(deck) as z:
            names = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml", n))
            if not names:
                return fonts
            xml = z.read(names[0]).decode("utf-8", "ignore")
        for tag in ("majorFont", "minorFont"):
            block = re.search(rf"<a:{tag}>(.*?)</a:{tag}>", xml, re.S)
            if not block:
                continue
            inner = block.group(1)
            m_lat = re.search(r'<a:latin\s+typeface="([^"]*)"', inner)
            if m_lat and m_lat.group(1):
                fonts[tag] = m_lat.group(1)
            m_ea = re.search(r'<a:ea\s+typeface="([^"]*)"', inner)
            if m_ea and m_ea.group(1):
                fonts[f"{tag}_ea"] = m_ea.group(1)
    except Exception:
        pass
    return fonts


def _clean_pdf_font_name(raw_name):
    return raw_name.split("+", 1)[-1] if "+" in raw_name else raw_name


def _slide_title_probe(slide):
    """Return title placeholder text, or largest-font non-footer text shape as fallback."""
    for sh in slide.shapes:
        if sh.is_placeholder:
            pt = str(sh.placeholder_format.type)
            if (pt.startswith("TITLE") or pt.startswith("CENTER_TITLE")) and sh.has_text_frame:
                txt = sh.text_frame.text.strip()
                if txt:
                    return txt
    best_pt, best_txt = -1.0, ""
    for sh in slide.shapes:
        if not sh.has_text_frame or sh.name in ("__adaptor_footer__", "__adaptor_snum__"):
            continue
        txt = sh.text_frame.text.strip()
        if len(txt) < 3:
            continue
        max_pt = 0.0
        for p in sh.text_frame.paragraphs:
            if p.font.size:
                max_pt = max(max_pt, p.font.size.pt)
            for r in p.runs:
                if r.font.size:
                    max_pt = max(max_pt, r.font.size.pt)
        if max_pt > best_pt:
            best_pt, best_txt = max_pt, txt
    return best_txt


def verify(pdf: Path, deck: Path, outdir: Path, dpi: int):
    import pymupdf
    from pptx import Presentation
    from pptx.util import Emu
    from PIL import Image, ImageStat

    prs = Presentation(str(deck))
    n_slides = len(prs.slides)
    deck_w_in = round(Emu(prs.slide_width).inches, 2)
    deck_h_in = round(Emu(prs.slide_height).inches, 2)
    titles = []
    pic_area = []
    notes_count = 0
    for slide in prs.slides:
        titles.append(_slide_title_probe(slide))
        area = 0.0
        for sh in slide.shapes:
            if (sh.shape_type is not None and "PICTURE" in str(sh.shape_type)) or getattr(sh, "has_chart", False):
                try:
                    area += Emu(sh.width).inches * Emu(sh.height).inches
                except (TypeError, AttributeError):
                    pass
        pic_area.append(area)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            notes_count += 1

    theme_fonts = read_deck_theme_fonts(deck)
    doc = pymupdf.open(str(pdf))
    pages = []
    all_ok = True
    size_match = True
    pdf_size_in = None
    font_warnings = []
    pdf_fonts = set()

    for pno in range(len(doc)):
        page = doc[pno]
        pw_in = round(page.rect.width / 72.0, 2)
        ph_in = round(page.rect.height / 72.0, 2)
        if pdf_size_in is None:
            pdf_size_in = f"{pw_in}x{ph_in}in"
        if abs(pw_in - deck_w_in) > 0.12 or abs(ph_in - deck_h_in) > 0.12:
            size_match = False

        text = page.get_text()
        for fmeta in page.get_fonts(full=True):
            if len(fmeta) >= 4 and fmeta[3]:
                pdf_fonts.add(_clean_pdf_font_name(fmeta[3]))

        pix = page.get_pixmap(dpi=dpi)
        img_path = outdir / f"page-{pno + 1:02d}.png"
        pix.save(str(img_path))
        with Image.open(img_path) as im:
            g = im.convert("L")
            hist = g.histogram()
            total_px = max(im.width * im.height, 1)
            dom_val = max(range(256), key=lambda k: hist[k])
            n_non_bg = sum(hist[v] for v in range(256) if abs(v - dom_val) > 18)
            non_bg_ratio = n_non_bg / total_px
            stddev = ImageStat.Stat(g).stddev[0]

        t_expect = titles[pno] if pno < len(titles) else ""
        t_ok = True
        if t_expect:
            probe = norm(t_expect[:40])
            page_norm = norm(text[:1000])
            if probe and probe not in page_norm:
                if len(probe) >= 6 and probe[:10] not in norm(text):
                    t_ok = False
                    all_ok = False
                    font_warnings.append(f"p{pno + 1}: title text not matched verbatim in exported PDF")
        if "\ufffd" in text or "\u25a1" in text:
            font_warnings.append(f"p{pno + 1}: replacement/tofu glyph detected in exported PDF")

        pages.append(dict(
            page=pno + 1,
            png=str(img_path),
            non_bg_ratio=round(non_bg_ratio, 4),
            pixel_stddev=round(stddev, 2),
            chars=len(text),
            title_found=t_ok,
        ))

    # Font substitution audit: classify exact match vs metric-compatible substitution vs unmapped fallback
    pdf_fonts_sorted = sorted(pdf_fonts)
    pdf_norm_map = [(f, f.lower().replace(" ", "").replace("-", "").replace("_", "")) for f in pdf_fonts_sorted]
    font_substitutions = []
    for role, fam in theme_fonts.items():
        if not fam or fam.startswith("+"):
            continue
        fam_key = fam.lower().replace(" ", "").replace("-", "").replace("_", "")
        if not fam_key:
            continue
        if any(fam_key in fn for _, fn in pdf_norm_map):
            continue
        equivs = FONT_EQUIV.get(fam_key, ())
        matched_eq = [orig for orig, fn in pdf_norm_map if any(eq in fn for eq in equivs)]
        if matched_eq:
            font_substitutions.append(dict(
                role=role, requested=fam, substituted_with=matched_eq[:3], metric_compatible=True))
            font_warnings.append(
                f"theme {role}={fam!r} rendered via metric-compatible fallback {matched_eq[:2]}")
        else:
            font_substitutions.append(dict(
                role=role, requested=fam, substituted_with=pdf_fonts_sorted[:4], metric_compatible=False))
            font_warnings.append(
                f"theme {role}={fam!r} not found verbatim in PDF fonts {pdf_fonts_sorted[:6]}")

    pages_match = len(doc) == n_slides
    blank = [
        p["page"] for k, p in enumerate(pages)
        if (p["chars"] < 5 and p["non_bg_ratio"] < 0.01)
        or (k < len(pic_area) and pic_area[k] > 1.0 and p["non_bg_ratio"] < 0.012 and p["pixel_stddev"] < 4.0)
    ]
    log = dict(
        pdf=str(pdf),
        previews_dir=str(outdir),
        dpi=dpi,
        slides=n_slides,
        pdf_pages=len(doc),
        pages_match=pages_match,
        deck_size_in=f"{deck_w_in}x{deck_h_in}in",
        pdf_page_size_in=pdf_size_in,
        size_match=size_match,
        notes_slides=notes_count,
        title_text_ok=all_ok,
        blank_pages=blank,
        theme_fonts=theme_fonts,
        pdf_fonts_used=pdf_fonts_sorted,
        font_substitutions=font_substitutions,
        font_warnings=font_warnings,
        pages=pages,
    )
    return log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--outdir", default="render", help="directory for preview PNGs and render_log.json")
    ap.add_argument("--pdf-out", default=None, help="optional final path for exported PDF (e.g. /app/output/deck.pdf)")
    ap.add_argument("--dpi", type=int, default=108, help="PNG preview DPI (recommended 96-120)")
    a = ap.parse_args()
    deck = Path(a.deck).resolve()
    if not deck.exists():
        print(f"FATAL: {deck} not found", file=sys.stderr)
        return 2
    outdir = Path(a.outdir)
    pdf, msg, retries_used = run_soffice(deck, outdir)
    if pdf is None:
        print(f"FATAL export: {msg}", file=sys.stderr)
        outdir.mkdir(parents=True, exist_ok=True)
        (outdir / "render_log.json").write_text(
            json.dumps(dict(status="export-failed", detail=msg, retries_used=retries_used),
                       ensure_ascii=False, indent=1), encoding="utf-8")
        return 2
    if a.pdf_out:
        pdf_dst = Path(a.pdf_out).resolve()
        pdf_dst.parent.mkdir(parents=True, exist_ok=True)
        if pdf_dst != pdf.resolve():
            shutil.move(str(pdf), str(pdf_dst))
            pdf = pdf_dst
    log = verify(pdf, deck, outdir, a.dpi)
    log["soffice"] = msg or "ok"
    log["retries_used"] = retries_used
    rc = 0 if (log["pages_match"] and log["title_text_ok"] and not log["blank_pages"]) else 1
    log["status"] = "verified" if rc == 0 else "mismatch"
    (outdir / "render_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"EXPORT slides={log['slides']} pdf_pages={log['pdf_pages']} match={log['pages_match']} "
          f"size_ok={log['size_match']} title_ok={log['title_text_ok']} blank={log['blank_pages']} "
          f"fonts={len(log['pdf_fonts_used'])} -> {outdir}/render_log.json")
    return rc


if __name__ == "__main__":
    sys.exit(main())
