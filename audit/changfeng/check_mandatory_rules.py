import json, sys
from pathlib import Path
rules = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
mandatory = ['\u65bd\u5de5\u529f\u80fd\u5206\u533a\u793a\u610f\u56fe', '\u65bd\u5de5\u8fdb\u5ea6\u6a2a\u9053\u56fe', '\u65bd\u5de5\u8fdb\u5ea6\u7f51\u7edc\u56fe']
for name in mandatory:
    if not any(x.get('name') == name and x.get('hard') is True for x in rules.get('required_figures', [])):
        raise SystemExit('MISSING_MANDATORY_FIGURE_RULE=' + name)
print('MANDATORY_FIGURE_RULES=PASS')
