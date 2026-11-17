# HTML to DOCX Visual — portable Agent Skill

This is an Agent Skills–format skill for ChatGPT-compatible Agent Skills environments and Claude. The main instructions live in `SKILL.md`. The included Python converter is optional but useful when a shell, Chromium and Pandoc are available.

## Install

- **Claude (web/desktop chat):** ZIP the enclosing `html-to-docx-visual/` folder, then **Customize → Skills → + → Create skill → Upload a skill**. Enable code execution if prompted.
- **Claude Code:** Copy the folder to `~/.claude/skills/html-to-docx-visual/` (or `.claude/skills/` in a project).
- **Codex:** Copy the folder to `~/.agents/skills/html-to-docx-visual/` (or `.agents/skills/` in a project).
- **ChatGPT:** On eligible workspace plans, open **Plugins → Skills → Create → Upload from your computer**. Skill upload availability depends on plan and workspace permissions. On plans without skill upload, give ChatGPT this `SKILL.md` as instructions and the HTML as an attachment, or use Codex.

## Run standalone from terminal

Requires Python 3.10+, Chromium, Pandoc, and the listed Python libraries. Optional: LibreOffice and `pdftoppm` for previews.

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
python scripts/convert.py your-document.html your-document.docx --review
```

A detection manifest and extracted figure PNGs are created in a sibling work directory, so an operator can verify uncertain automatic detections.

## Caveats

Automatic detection of arbitrary HTML/CSS visual layouts is inherently heuristic. The script detects common cases but cannot guarantee perfect identification of every visual diagram. Add `data-export-as-image="true"` to outer diagram containers for high certainty, without editing the visual design. Scripts running in some hosted chat sandboxes may be restricted. Never execute unknown HTML/JavaScript with sensitive local files accessible.

Main rules: preserve editable body content, capture the browser-rendered visual composition as one image, insert it in reading order, and check the finished Word pages.
