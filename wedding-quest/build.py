"""Build the game from src/template.html + quests.json.

Outputs:
  wedding-quest.html   page fragment for the Claude Artifact version
  web/Index.html       full HTML document for the Google Apps Script web app
  web/Code.gs          server script for the web app (copied from src/Code.gs)
"""
import json, pathlib, shutil

here = pathlib.Path(__file__).parent
data = json.loads((here / "quests.json").read_text(encoding="utf-8"))
tpl = (here / "src" / "template.html").read_text(encoding="utf-8")
assert "/*__QUESTS__*/null" in tpl
page = tpl.replace("/*__QUESTS__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
(here / "wedding-quest.html").write_text(page, encoding="utf-8")

head, body = page.split("</style>\n", 1)
web = ("<!DOCTYPE html>\n<html lang=\"id\">\n<head>\n<meta charset=\"utf-8\">\n"
       "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
       "<base target=\"_top\">\n" + head + "</style>\n</head>\n<body>\n" + body + "</body>\n</html>\n")
(here / "web").mkdir(exist_ok=True)
(here / "web" / "Index.html").write_text(web, encoding="utf-8")
shutil.copy(here / "src" / "Code.gs", here / "web" / "Code.gs")

# GitHub Pages copy (repo-root docs/): same page, talking to the Apps Script deployment as an API.
API_URL = "https://script.google.com/macros/s/AKfycbxUxUF28OJ87brLOvb_Q9PD9kd4EVQIjJFNrKfCrLhH80ZcBNfo0qk6D1amBUmESTPn/exec"
docs = here.parent / "docs"
docs.mkdir(exist_ok=True)
pages = web.replace('<base target="_top">\n', "").replace("<script>\n\"use strict\";", f'<script>window.WQ_API = "{API_URL}";</script>\n<script>\n"use strict";', 1)
assert "window.WQ_API" in pages
(docs / "index.html").write_text(pages, encoding="utf-8")
print("built", len(page), "bytes,", len(data["quests"]), "quests")
