#!/usr/bin/env python3
"""Compatibility entry point; constants and execution come from the synced project script."""
from pathlib import Path
import importlib.util
_spec = importlib.util.spec_from_file_location('_project_page_gate', Path(__file__).with_name('长丰义井镇项目页数硬审.py'))
_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_module)
EXPECTED_BODY_PAGES = _module.EXPECTED_BODY_PAGES
PAGE_TOLERANCE = _module.PAGE_TOLERANCE
FRONT_MATTER_PAGES = _module.FRONT_MATTER_PAGES
if __name__ == '__main__':
    raise SystemExit(_module.main())
