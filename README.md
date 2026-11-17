# HTML to DOCX Visual — portable Agent Skill

This is an Agent Skills–format skill for ChatGPT-compatible Agent Skills environments and Claude. The main instructions live in `SKILL.md`. The included Python converter is optional but useful when a shell, Chromium and Pandoc are available.

## Install

**Ready-made package:** download [`dist/html-to-docx-visual.zip`](dist/html-to-docx-visual.zip). It contains the correctly named `html-to-docx-visual/` folder with only the runtime files, so there's nothing to build. Maintainers: after changing any packaged file, run `python scripts/build_skill_zip.py` and commit the updated zip.

- **Claude (web/desktop chat):** Upload the zip via **Customize → Skills → + → Create skill → Upload a skill**. Enable code execution if prompted.
- **Claude Code:** Unzip and copy the `html-to-docx-visual/` folder to `~/.claude/skills/html-to-docx-visual/` (or `.claude/skills/` in a project).
- **Codex:** Unzip and copy the `html-to-docx-visual/` folder to `~/.agents/skills/html-to-docx-visual/` (or `.agents/skills/` in a project).
- **ChatGPT:** On eligible workspace plans, open **Plugins → Skills → Create → Upload from your computer**. Skill upload availability depends on plan and workspace permissions. On plans without skill upload, give ChatGPT this `SKILL.md` as instructions and the HTML as an attachment, or use Codex.

## Run standalone from terminal

Requires Python 3.10+, Chromium, Pandoc (3.x recommended for figure captions), and the listed Python libraries. Optional: LibreOffice (`soffice`) and `pdftoppm` (Poppler) for previews.

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
python scripts/convert.py your-document.html your-document.docx --review
```

A detection manifest (`conversion_report.json`) and extracted figure PNGs are created in a sibling `<name>_export_assets/` directory, so an operator can verify uncertain automatic detections.

| Option | Purpose |
|---|---|
| `--review` | Render the DOCX to PDF + page PNGs for visual QA (needs LibreOffice and `pdftoppm`) |
| `--asset-root DIR` | Folder local images/CSS/JS may come from (default: the HTML file's folder) |
| `--allow-network` | Let the page and Pandoc fetch remote resources. Off by default; use only for trusted HTML |
| `--viewport-width`, `--scale`, `--wait-ms` | Layout width (auto-widened for wide content), screenshot sharpness, extra wait for JS rendering |

Diagram markup hints: `data-export-as-image="true"` forces capture (works on any element, including `<svg>`, `<canvas>`, and elements inside lists/tables); `aria-label` or `data-caption` provides alt text, and `data-caption` also adds a Word caption.

## Security model

The converter renders HTML with JavaScript enabled, so treat input as untrusted:

- **Network is blocked by default.** HTTP(S), DNS, WebSocket, service workers and loopback access are cut off. Blocked requests are listed in the report. Remote `<img>` elements are replaced by their alt text rather than fetched by Pandoc.
- **Local files are confined to the asset root.** `..` paths, absolute paths, `file:` URLs and symlinks that resolve outside it are never inlined or embedded in the DOCX, and are listed under `security.blocked_local_references`.
- `<iframe>`, `<object>`, `<embed>`, media, and hidden (`display:none`) content are stripped before Pandoc runs.
- Residual risk: a page may still *display* (not read or transmit) an image from elsewhere on disk inside a diagram screenshot. Run unknown HTML in a container or VM without sensitive files, and follow your organisation's InfoSec policies for untrusted content.

## Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest tests
```

The end-to-end tests skip automatically when Pandoc or Playwright's Chromium is unavailable.

## Caveats

Automatic detection of arbitrary HTML/CSS visual layouts is inherently heuristic. The script detects common cases but cannot guarantee perfect identification of every visual diagram. Add `data-export-as-image="true"` to outer diagram containers for high certainty, without editing the visual design. Scripts running in some hosted chat sandboxes may be restricted.

Main rules: preserve editable body content, capture the browser-rendered visual composition as one image, insert it in reading order, and check the finished Word pages.
