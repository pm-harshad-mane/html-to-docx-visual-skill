# HTML to DOCX Visual — portable Agent Skill

Converts AI-generated HTML documents into editable Word (`.docx`) files while preserving CSS, div/span, SVG, Canvas, and JavaScript-created diagrams as sharp embedded images. Use it for HTML-to-Word export, especially technical reports and architecture documents with browser-rendered diagrams.

## Why this skill exists

AI assistants often produce rich HTML reports where diagrams aren't images at all. They're built from styled `<div>`s and `<span>`s, Flexbox/Grid layouts, CSS arrows, inline SVG, Canvas, or JavaScript-generated DOM. Ordinary HTML-to-Word converters handle those badly: a flowchart collapses into a pile of loose paragraphs, connectors and colours disappear, and Canvas or script-drawn charts vanish entirely. Screenshotting whole pages has the opposite problem: everything looks right, but nothing can be edited.

This skill does both jobs, each part the right way:

| Part of the HTML | In the Word document |
|---|---|
| Headings, paragraphs, lists, links | Editable text using Word heading styles |
| Data tables (`<table>`) | Real, editable Word tables |
| Diagrams (CSS/div/span, SVG, Canvas, JS-rendered) | One sharp, high-resolution image each, rendered by a real browser exactly as displayed and placed where the diagram was |

The original HTML stays the editable source for the diagrams.

**Typical uses:** architecture and design documents, technical reports, request-flow and process write-ups, and any AI-generated HTML that has to be shared, reviewed or archived as a Word file.

## How it works

1. Renders the HTML in headless Chromium, waiting for fonts and scripts.
2. Finds diagrams, using explicit `data-export-as-image="true"` markers first, then meaningful class names (`flowchart`, `architecture-diagram`, …), then layout heuristics and standalone SVG/Canvas.
3. Screenshots each diagram at high resolution and swaps it for an image in an export-only copy of the HTML.
4. Converts that copy to `.docx` with Pandoc and tidies the layout (margins, table borders, image sizing).
5. Writes a `conversion_report.json` listing every detected diagram, plus optional page previews for checking.

The main instructions live in `SKILL.md`, in Agent Skills format for Claude and other Agent Skills–compatible environments (Codex, ChatGPT). The included Python converter (`scripts/convert.py`) is optional but recommended when a shell, Chromium and Pandoc are available.

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
