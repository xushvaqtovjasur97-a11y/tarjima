#!/bin/bash
# render (pdf) + UNO bloklar → fix2 (gaplarni ko‘chirish) → takrorlash.  Foydalanish: loop2.sh <docx> <ishchi papka> [N]
F=$1; W=$2; N=${3:-8}
mkdir -p $W
for it in $(seq 1 $N); do
  cp $F $W/cur.docx; rm -f $W/cur.pdf
  timeout 900 python3 /home/user/tarjima/tools/sahifalash/render_uno.py $W/cur.docx $W/cur.pdf --fnline 4.1 --fntext 4.1 >/dev/null
  timeout 900 python3 /home/user/tarjima/tools/sahifalash/bloklar_uno.py $W/cur.docx $W/bl.json >/dev/null
  echo "=== takror $it ($(pdfinfo $W/cur.pdf | grep Pages))"
  out=$(FIXSTATE=$W/fixed.json python3 /home/user/tarjima/tools/sahifalash/fix2.py $F $W/cur.pdf $W/bl.json)
  echo "$out"
  echo "$out" | grep -q "oldinga" || break
done
