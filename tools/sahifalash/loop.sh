#!/bin/bash
# render → fix (har bir bo‘limda birinchi muammo) → takrorlash
S=/tmp/claude-0/-home-user-tarjima/315c132a-88b7-5e49-bb67-7f6a05e43e7d/scratchpad
F=/home/user/tarjima/Dissertatsiya_Maxmudov_O.T_asosiy.docx
N=${1:-8}
for it in $(seq 1 $N); do
  rm -f $S/lay/cur.pdf
  cp $F $S/lay/cur.docx
  (cd $S/lay && timeout 600 soffice --headless --convert-to pdf cur.docx >/dev/null 2>&1)
  echo "=== takror $it"
  out=$(python3 $S/lay/fix.py $S/lay/cur.pdf $F)
  echo "$out"
  python3 -c "import sys; sys.path.insert(0,'/home/user/tarjima/tools'); from pardoz import pardoz; pardoz('$F')" >/dev/null
  echo "$out" | grep -q "ko‘chirildi\|kichraytirildi" || break
done
