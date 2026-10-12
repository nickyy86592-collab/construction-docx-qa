#!/usr/bin/env python3
"""Compatibility entry point for the canonical Word audit. No independent audit rules."""
from pathlib import Path
import importlib.util
import json
import sys
_spec = importlib.util.spec_from_file_location('_word_delivery_gate', Path(__file__).with_name('Word正式交付硬审（不含页数）.py'))
_module = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _module
_spec.loader.exec_module(_module)

def main():
    argv = sys.argv[1:]
    if '--photo-stage' in argv:
        i = argv.index('--photo-stage')
        if i + 1 >= len(argv) or argv[i+1] not in ('draft', 'final'):
            raise SystemExit('photo-stage must be draft or final')
        stage = argv[i+1]
        del argv[i:i+2]
        if '--rules' not in argv:
            raise SystemExit('--photo-stage requires explicit --rules')
        index = argv.index('--rules')+1
        source = Path(argv[index])
        rules = json.loads(source.read_text(encoding='utf-8-sig'))
        rules['photo_delivery_stage'] = stage
        report_index = argv.index('--report')+1 if '--report' in argv else None
        snapshot = Path(argv[report_index]).with_suffix('.effective-rules.json') if report_index is not None else source.with_name(source.stem+'.effective-rules.json')
        snapshot.parent.mkdir(parents=True, exist_ok=True)
        snapshot.write_text(json.dumps(rules,ensure_ascii=False,indent=2),encoding='utf-8')
        argv[index] = str(snapshot)
    sys.argv[1:] = argv
    return _module.main()

if __name__ == '__main__':
    raise SystemExit(main())
