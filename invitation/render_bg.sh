#!/bin/sh
# usage: render_bg.sh in.html out.jpg  -> 1748x2480 (A5 at 300 dpi)
B=$(ls -d /opt/pw-browsers/chromium-*/chrome-linux/chrome | head -1)
T=$(mktemp --suffix=.png)
$B --headless=new --no-sandbox --disable-gpu --virtual-time-budget=10000 --window-size=1748,2700 --hide-scrollbars --screenshot=$T "file://$(realpath $1)" 2>/dev/null
python3 -c "from PIL import Image; Image.open('$T').convert('RGB').crop((0,0,1748,2480)).save('$2',quality=92)"
rm -f $T
