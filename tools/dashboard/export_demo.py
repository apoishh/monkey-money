"""Build the online demo in ../../docs from the SAMPLE workbook (for GitHub Pages).

    python tools\\dashboard\\export_demo.py

Only ever run this against the fictional sample 'Monkey Money.xlsx' - whatever it reads is published.
"""
import json
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
DOCS = os.path.join(ROOT, "docs")
SAMPLE = os.path.join(ROOT, "Monkey Money.xlsx")
sys.path.insert(0, HERE)
import data  # noqa: E402

if os.path.basename(data.WORKBOOK) != "Monkey Money.xlsx" or "--workbook" in sys.argv:
    sys.exit("Refusing: the demo must be built from the sample 'Monkey Money.xlsx' only.")

os.makedirs(DOCS, exist_ok=True)
d = data.load(SAMPLE)
d["source"] = {"file": "Monkey Money.xlsx (sample)", "saved": d["source"]["saved"]}
with open(os.path.join(DOCS, "data.json"), "w", encoding="utf-8") as f:
    json.dump(d, f, separators=(",", ":"), default=str)

static = os.path.join(HERE, "static")
for name in ("app.js", "style.css"):
    shutil.copyfile(os.path.join(static, name), os.path.join(DOCS, name))
html = open(os.path.join(static, "index.html"), encoding="utf-8").read()
html = html.replace('href="/static/style.css"', 'href="style.css"').replace('src="/static/app.js"', 'src="app.js"')
html = html.replace("<title>Monkey Money</title>", "<title>Monkey Money – demo</title>")
banner = ('<div class="banner demo">Demo with fictional sample data. '
          '<a id="repo-link" href="https://github.com/">Get Monkey Money</a> to run it on your own PC with your own workbook.</div>')
html = re.sub(r'(<main>\s*)', r"\1" + banner + "\n  ", html, count=1)
open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8").write(html)
open(os.path.join(DOCS, ".nojekyll"), "w").close()
print(f"Demo written to {DOCS}: {len(d['transactions'])} sample transactions")
