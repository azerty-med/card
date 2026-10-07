#!/bin/sh
# usage: preview.sh file.docx out.png -> converts with LibreOffice and rasterises page 1
S=/root/.claude/skills/synced/ee36b141-f323-4afc-bc1d-3a5c37e81c42_bbf89af0-7744-41fb-8705-c383c374dda0/docx
D=$(mktemp -d)
python3 $S/scripts/office/soffice.py --headless --convert-to pdf --outdir $D "$1" >/dev/null 2>&1
pdfinfo $D/*.pdf | grep Pages
pdftoppm -png -r 130 -singlefile $D/*.pdf "${2%.png}"
rm -rf $D
