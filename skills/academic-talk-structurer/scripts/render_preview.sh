#!/bin/sh
# S5 render preview: PPTX -> PDF -> PNGs (for visual self-check with image reading).
# Usage: render_preview.sh <deck.pptx> <outdir> [dpi]
# Env: SOFFICE_BIN to override soffice/libreoffice autodetect. pdffonts (if present) flags non-embedded fonts.
# No-script fallback: run the two commands manually (see SKILL.md S5), or export PDF via GUI.
set -e
SRC="$1"; OUT="$2"; DPI="${3:-100}"
[ -f "$SRC" ] || { echo "missing input: $SRC"; exit 2; }
mkdir -p "$OUT"
SOFFICE="${SOFFICE_BIN:-}"
if [ -z "$SOFFICE" ]; then command -v soffice >/dev/null 2>&1 && SOFFICE=soffice; fi
if [ -z "$SOFFICE" ]; then command -v libreoffice >/dev/null 2>&1 && SOFFICE=libreoffice; fi
[ -n "$SOFFICE" ] || { echo "no soffice/libreoffice found; export PDF manually then run pdftoppm"; exit 4; }
BASE=$(basename "$SRC"); BASE="${BASE%.*}"
PDF="$OUT/$BASE.pdf"
ok=0
for try in 1 2; do
  PROF="/tmp/lo_pk_skill_$try"
  if timeout 240 "$SOFFICE" --headless -env:UserInstallation=file://$PROF --convert-to pdf --outdir "$OUT" "$SRC" && [ -f "$PDF" ]; then ok=1; break; fi
  echo "convert attempt $try failed; retrying"
done
[ "$ok" = 1 ] || { echo "pdf convert failed twice (check fonts/embedded media); deliver pptx + lint, mark rendering failed"; exit 5; }
if command -v pdftoppm >/dev/null 2>&1; then
  timeout 180 pdftoppm -png -r "$DPI" "$PDF" "$OUT/preview" || echo "pdftoppm failed; previews skipped"
else
  echo "pdftoppm missing; keep $PDF for manual inspection"
fi
if command -v pdffonts >/dev/null 2>&1; then
  NONEMB=$(pdffonts "$PDF" 2>/dev/null | awk 'NR>2 && $5=="no" {print $1}' | tr '\n' ',' )
  [ -n "$NONEMB" ] && echo "warn: fonts not embedded in PDF: $NONEMB (check for tofu/glyph loss; see SKILL.md 异常处置)"
fi
echo "PDF: $PDF"
ls "$OUT"/preview*.png 2>/dev/null | head -40 || true
