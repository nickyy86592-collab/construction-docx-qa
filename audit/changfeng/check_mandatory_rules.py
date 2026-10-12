import json, sys
from pathlib import Path
rules = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
mandatory = ['施工功能分区示意图', '施工进度横道图', '施工进度网络图']
for name in mandatory:
    if not any(x.get('name') == name and x.get('hard') is True and x.get('color_required') is True for x in rules.get('required_figures', [])):
        raise SystemExit('MISSING_MANDATORY_COLOR_FIGURE_RULE=' + name)
slots = rules.get('required_real_photo_slots')
if not isinstance(slots,list) or not slots or any(not x.get('name') for x in slots):
    raise SystemExit('MISSING_MANDATORY_REAL_PHOTO_SLOTS')
if rules.get('photo_delivery_stage') not in ('draft','final'):
    raise SystemExit('MISSING_PHOTO_DELIVERY_STAGE')
print('MANDATORY_COLOR_FIGURE_AND_PHOTO_RULES=PASS')
