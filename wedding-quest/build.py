"""Build wedding-quest.html from src/template.html + quests.json."""
import json, pathlib
here = pathlib.Path(__file__).parent
data = json.loads((here / "quests.json").read_text(encoding="utf-8"))
tpl = (here / "src" / "template.html").read_text(encoding="utf-8")
assert "/*__QUESTS__*/null" in tpl
out = tpl.replace("/*__QUESTS__*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
(here / "wedding-quest.html").write_text(out, encoding="utf-8")
print("built", len(out), "bytes,", len(data["quests"]), "quests")
