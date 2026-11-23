#!/usr/bin/env python3
"""Inventory an existing deck (starting-deck scenarios) and optionally draft a plan.

Usage:
  python3 inspect_deck.py DECK.pptx [--json deck_inventory.json] [--draft plan_draft.json] [--imgdir imgs]

Captures placeholders (including pictures, tables, and charts hosted inside placeholders),
free-floating textboxes, native PPTX tables, native PPTX charts, grouped shapes,
pictures (saved as PNG/JPG), speaker notes, paragraph indent levels, and master/theme fingerprint.
No-script fallback: python-pptx loop over prs.slides / slide.shapes (see SKILL.md §3).
"""
import argparse, json, os, re, sys, tempfile, zipfile
from pathlib import Path


def inches(v):
    try:
        return round(v / 914400, 2)
    except TypeError:
        return None


def ph_kind(t):
    if t.startswith("CENTER_TITLE") or t.startswith("TITLE"):
        return "title"
    if t.startswith("SUBTITLE"):
        return "subtitle"
    if t.startswith("PICTURE") or t.startswith("CLIP_ART"):
        return "picture"
    if any(t.startswith(k) for k in ("BODY", "OBJECT", "TEXT", "CONTENT", "VERTICAL_BODY")):
        return "body"
    if any(t.startswith(k) for k in ("FOOTER", "SLIDE_NUMBER", "DATE", "HEADER")):
        return "chrome"
    return None


def _theme_fonts(pptx_path: Path):
    fonts = {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
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


def _iter_shapes(shapes):
    """Yield leaf shapes, recursively unpacking grouped shapes."""
    for sh in shapes:
        if sh.shape_type is not None and "GROUP" in str(sh.shape_type):
            try:
                yield from _iter_shapes(sh.shapes)
                continue
            except Exception:
                pass
        yield sh


def _extract_paragraphs(sh):
    """Return list of {'text': ..., 'level': int} preserving bullet hierarchy."""
    if not sh.has_text_frame:
        return []
    out = []
    for p in sh.text_frame.paragraphs:
        txt = p.text.strip()
        if txt:
            out.append(dict(text=txt[:400], level=int(p.level or 0)))
    return out


def _try_get_image(sh):
    """Safely extract embedded image from either a Picture shape or a populated Picture/Content placeholder."""
    st = str(sh.shape_type) if sh.shape_type is not None else ""
    if "PICTURE" in st:
        try:
            return sh.image
        except Exception:
            return None
    if sh.is_placeholder:
        # Populated PicturePlaceholder / SlidePlaceholder holding an image has a blip child
        try:
            return sh.image
        except Exception:
            return None
    return None


def _extract_chart(sh):
    """Extract editable series/categories metadata from a native PPTX chart shape."""
    try:
        ch = sh.chart
        ctype = str(ch.chart_type).split(" ")[0]
        title = ""
        try:
            if ch.has_title and ch.chart_title.has_text_frame:
                title = ch.chart_title.text_frame.text.strip()
        except Exception:
            pass
        cats = []
        series_list = []
        if ch.plots:
            plot = ch.plots[0]
            try:
                cats = [str(c) for c in plot.categories][:30]
            except Exception:
                pass
            for ser in plot.series:
                vals = []
                try:
                    vals = [round(v, 4) if isinstance(v, float) else v for v in ser.values][:30]
                except Exception:
                    pass
                series_list.append(dict(name=str(ser.name or ""), values=vals))
        return dict(
            chart_type=ctype,
            title=title[:160],
            categories=cats,
            series=series_list[:10],
            left=inches(sh.left), top=inches(sh.top),
            width=inches(sh.width), height=inches(sh.height),
        )
    except Exception as e:
        return dict(chart_type="unknown", error=str(e)[:120],
                    left=inches(sh.left), top=inches(sh.top),
                    width=inches(sh.width), height=inches(sh.height))


def inspect(deck: Path, imgdir: Path = None):
    from pptx import Presentation
    tmp_name = None
    try:
        prs = Presentation(str(deck))
    except ValueError as e:
        if "template.main" not in str(e):
            raise
        tmp_fd, tmp_name = tempfile.mkstemp(suffix=".pptx")
        os.close(tmp_fd)
        with zipfile.ZipFile(deck) as zin, zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "[Content_Types].xml":
                    data = data.replace(
                        b"presentationml.template.main+xml",
                        b"presentationml.presentation.main+xml")
                zout.writestr(item, data)
        prs = Presentation(tmp_name)

    try:
        slides = []
        for i, slide in enumerate(prs.slides, 1):
            rec = dict(
                n=i,
                layout=slide.slide_layout.name if slide.slide_layout is not None else "?",
                placeholders=[],
                free_textboxes=[],
                tables=[],
                charts=[],
                pictures=[],
                notes=None,
            )
            for sh in _iter_shapes(slide.shapes):
                # 1) Record placeholder metadata (never block table/chart/picture extraction!)
                is_ph = False
                if sh.is_placeholder:
                    is_ph = True
                    pf = sh.placeholder_format
                    t = str(pf.type).split(" ")[0]
                    paras = _extract_paragraphs(sh)
                    txt = "\n".join(p["text"] for p in paras)
                    rec["placeholders"].append(dict(
                        idx=pf.idx, type=t, role=ph_kind(t), text=txt[:800],
                        paragraphs=paras[:20],
                        has_table=bool(sh.has_table),
                        has_chart=bool(sh.has_chart),
                        left=inches(sh.left), top=inches(sh.top),
                        width=inches(sh.width), height=inches(sh.height)))

                # 2) Extract native tables (both standalone and placeholder-hosted)
                if sh.has_table:
                    rows = [[re.sub(r"\s+", " ", cell.text).strip() for cell in row.cells]
                            for row in sh.table.rows]
                    rec["tables"].append(dict(
                        from_placeholder=is_ph,
                        rows=rows[:20],
                        left=inches(sh.left), top=inches(sh.top),
                        width=inches(sh.width), height=inches(sh.height)))
                    continue

                # 3) Extract native charts (both standalone and placeholder-hosted)
                if sh.has_chart:
                    ch_rec = _extract_chart(sh)
                    ch_rec["from_placeholder"] = is_ph
                    rec["charts"].append(ch_rec)
                    continue

                # 4) Extract pictures (both standalone and placeholder-hosted)
                img = _try_get_image(sh)
                if img is not None:
                    p = dict(from_placeholder=is_ph,
                             left=inches(sh.left), top=inches(sh.top),
                             width=inches(sh.width), height=inches(sh.height))
                    if imgdir is not None:
                        try:
                            imgdir.mkdir(parents=True, exist_ok=True)
                            ext = img.ext.lower() if img.ext and img.ext.lower() in ("png", "jpg", "jpeg") else "png"
                            fn = imgdir / f"slide{i:02d}_img{len(rec['pictures']) + 1}.{ext}"
                            fn.write_bytes(img.blob)
                            p["saved"] = str(fn)
                            p["native_px"] = list(img.size)
                        except Exception as e:
                            p["save_error"] = str(e)
                    rec["pictures"].append(p)
                    continue

                # 5) Free-floating textboxes (non-placeholder shapes with text)
                if not is_ph and sh.has_text_frame:
                    paras = _extract_paragraphs(sh)
                    txt = "\n".join(p["text"] for p in paras)
                    if txt:
                        rec["free_textboxes"].append(dict(
                            name=sh.name, text=txt[:800], paragraphs=paras[:20],
                            left=inches(sh.left), top=inches(sh.top),
                            width=inches(sh.width), height=inches(sh.height)))

            if slide.has_notes_slide:
                rec["notes"] = slide.notes_slide.notes_text_frame.text.strip()[:1200]
            slides.append(rec)

        layout_names = [lo.name for m in prs.slide_masters for lo in m.slide_layouts]
        return dict(
            deck=str(deck),
            slide_count=len(slides),
            master_fingerprint=dict(
                slide_width_in=inches(prs.slide_width),
                slide_height_in=inches(prs.slide_height),
                theme_fonts=_theme_fonts(deck),
                n_layouts=len(layout_names),
                layouts=layout_names[:25],
            ),
            slides=slides,
        )
    finally:
        if tmp_name and os.path.exists(tmp_name):
            try:
                os.remove(tmp_name)
            except OSError:
                pass


def _chart_to_plan(ch):
    ct = (ch.get("chart_type") or "").upper()
    kind = "line" if "LINE" in ct else ("bar" if "BAR" in ct else "column")
    return dict(
        type=kind,
        title=ch.get("title") or "",
        categories=ch.get("categories") or [],
        series=ch.get("series") or [],
        show_values=True,
    )


def draft(inv):
    """Build a plan skeleton preserving placeholders, free textboxes, two-column splits, tables, charts, and pictures."""
    out_slides = []
    sw = (inv.get("master_fingerprint") or {}).get("slide_width_in") or 13.33
    for s in inv["slides"]:
        lo = (s["layout"] or "").lower()
        phs_title = [p for p in s["placeholders"] if p["role"] == "title" and p["text"]]
        phs_body = [p for p in s["placeholders"]
                    if p["role"] == "body" and p["text"] and not p.get("has_table") and not p.get("has_chart")]
        phs_sub = [p for p in s["placeholders"] if p["role"] == "subtitle" and p["text"]]
        title = phs_title[0]["text"].splitlines()[0] if phs_title else ""
        sub = phs_sub[0]["text"] if phs_sub else ""

        free_sorted = sorted(s["free_textboxes"], key=lambda x: (x.get("top") or 99, x.get("left") or 99))
        if not title and free_sorted:
            title = free_sorted[0]["text"].splitlines()[0]
            free_sorted = free_sorted[1:]

        body_items = []
        for p in phs_body:
            body_items.extend(p.get("paragraphs") or [{"text": ln.strip(), "level": 0}
                                                       for ln in p["text"].splitlines() if ln.strip()])
        for ft in free_sorted:
            body_items.extend(ft.get("paragraphs") or [{"text": ln.strip(), "level": 0}
                                                        for ln in ft["text"].splitlines() if ln.strip()])

        # Detect genuine two-column body layout (2 side-by-side text placeholders)
        phs_body_sorted = sorted(phs_body, key=lambda p: (p.get("left") or 0))
        is_two_col = (
            len(phs_body_sorted) >= 2
            and abs((phs_body_sorted[1].get("left") or 0) - (phs_body_sorted[0].get("left") or 0)) >= sw * 0.25
            and not s["pictures"] and not s["tables"] and not s.get("charts")
        )

        if "section" in lo:
            kind = "section"
        elif s["n"] == 1 and (sub or "title" in lo):
            kind = "title"
        elif is_two_col:
            kind = "two_col"
        elif s["pictures"]:
            kind = "figure"
        else:
            kind = "content"

        entry = dict(kind=kind, layout_hint=s["layout"], title=title[:120], notes=s.get("notes") or "")
        if kind == "title":
            entry["subtitle"] = sub or ("\n".join(b["text"] for b in body_items[:3]) if body_items else "")
        elif kind == "two_col":
            l_paras = phs_body_sorted[0].get("paragraphs") or []
            r_paras = phs_body_sorted[1].get("paragraphs") or []
            entry["left"] = {"bullets": l_paras[:6]}
            entry["right"] = {"bullets": r_paras[:6]}
        elif body_items:
            entry["bullets"] = body_items[:8]

        if s["pictures"]:
            entry["images"] = [
                {"path": p.get("saved"), "caption": "", "source": f"from starting deck, slide {s['n']}"}
                for p in s["pictures"] if p.get("saved")
            ]
        if s["tables"]:
            tb0 = s["tables"][0]["rows"]
            if len(tb0) >= 2:
                entry["table"] = {"headers": tb0[0], "rows": tb0[1:],
                                  "caption": f"Table from starting deck slide {s['n']}"}
        if s.get("charts"):
            ch0 = s["charts"][0]
            if ch0.get("categories") and ch0.get("series"):
                entry["chart"] = _chart_to_plan(ch0)
                entry["chart"]["source"] = f"from starting deck, slide {s['n']}"
        out_slides.append(entry)
    return dict(
        template=None,
        output=None,
        slides=out_slides,
        _meta="DRAFT extracted from starting deck — review titles/bullets/tables/charts/evidence, fill template/output, add paper locators before building.",
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--json", default=None)
    ap.add_argument("--draft", default=None)
    ap.add_argument("--imgdir", default="deck_imgs")
    a = ap.parse_args()
    deck = Path(a.deck)
    if not deck.exists():
        print(f"FATAL: {deck} not found", file=sys.stderr)
        return 2
    inv = inspect(deck, Path(a.imgdir) if (a.draft or a.imgdir != "deck_imgs") else None)
    js = json.dumps(inv, ensure_ascii=False, indent=1)
    if a.json:
        Path(a.json).parent.mkdir(parents=True, exist_ok=True)
        Path(a.json).write_text(js, encoding="utf-8")
        n_tbl = sum(len(s["tables"]) for s in inv["slides"])
        n_pic = sum(len(s["pictures"]) for s in inv["slides"])
        n_ch = sum(len(s["charts"]) for s in inv["slides"])
        print(f"OK {inv['slide_count']} slides (tables={n_tbl} pics={n_pic} charts={n_ch}) -> {a.json}")
    else:
        print(js)
    if a.draft:
        Path(a.draft).parent.mkdir(parents=True, exist_ok=True)
        Path(a.draft).write_text(json.dumps(draft(inv), ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"DRAFT plan skeleton -> {a.draft} (review before use)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
