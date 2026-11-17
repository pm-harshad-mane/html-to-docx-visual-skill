#!/usr/bin/env python3
"""Build dist/html-to-docx-visual.zip: the ready-to-upload skill package.

The archive contains a single top-level folder named after the skill
(`name:` in SKILL.md) holding only runtime files. Output is deterministic
(fixed timestamps and order), so rebuilding unchanged sources yields an
identical zip and `git diff` only shows real changes.

Usage: python scripts/build_skill_zip.py
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
INCLUDE = ['SKILL.md', 'README.md', 'requirements.txt', 'scripts/convert.py', 'examples/sample.html']
FIXED_TIME = (2020, 1, 1, 0, 0, 0)


def skill_name() -> str:
    match = re.search(r'^name:\s*(\S+)\s*$', (REPO / 'SKILL.md').read_text(encoding='utf8'), re.M)
    if not match:
        sys.exit('SKILL.md has no `name:` in its frontmatter')
    return match.group(1)


def main() -> int:
    name = skill_name()
    missing = [p for p in INCLUDE if not (REPO / p).is_file()]
    if missing:
        sys.exit(f'Missing files: {", ".join(missing)}')
    out = REPO / 'dist' / f'{name}.zip'
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for rel in sorted(INCLUDE):
            info = zipfile.ZipInfo(f'{name}/{rel}', date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o755 if rel.endswith('.py') else 0o644) << 16
            zf.writestr(info, (REPO / rel).read_bytes())
    print(f'Built {out.relative_to(REPO)} ({out.stat().st_size:,} bytes, {len(INCLUDE)} files)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
