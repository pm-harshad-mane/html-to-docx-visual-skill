---
name: html-to-docx-visual
description: Convert AI-generated HTML documents into editable Word (.docx) files while preserving CSS, div/span, SVG, Canvas, and JavaScript-created diagrams as sharp embedded images. Use for HTML-to-Word export, especially technical reports and architecture documents with browser-rendered diagrams.
---

# HTML to DOCX: preserve browser-generated diagrams

## Goal

Convert the user's HTML into a professional `.docx`. **Keep prose, headings, hyperlinks, lists, and real data tables editable in Word.** **Render visual diagrams exactly as the browser displays them and embed them as images in their original reading-order positions.** The HTML file remains the editable source of truth for diagrams.

Do not attempt to reconstruct CSS diagrams using Word shapes. Do not screenshot entire document pages. Do not treat an HTML `<table>` used for tabular data as a diagram.

## What counts as a diagram

A diagram can be built entirely from `<div>`, `<span>`, Flexbox, CSS Grid, backgrounds, arrows, absolute positioning, and pseudo-elements. It might instead be SVG, Mermaid output, Canvas, or dynamically generated JavaScript DOM. Judge **visual function**, not merely the HTML tag name. Flowcharts, architecture drawings, request flows, timelines, decision trees, annotated process maps, swimlanes, and composed schematics are diagram candidates.

## Procedure

1. **Inspect the HTML and assets.** Identify CSS, fonts, scripts, images, and any dependencies. Do not execute untrusted HTML with filesystem/network access beyond what is required for rendering; work in a restricted or isolated browser context whenever possible. If it has external dependencies, explain any that could not load.
2. **Render in Chromium or another modern browser**, including local CSS/JS. Wait for page load, fonts, and diagram rendering. If animations exist, freeze their state before capture.
3. **Locate diagram containers** in this priority order:
   - Elements marked `data-export-as-image="true"`, `data-docx-render="image"`, or `data-skill-diagram`.
   - Containers with semantically meaningful class/id names (e.g. `architecture-diagram`, `flowchart`, `process-flow`, `topology`, `sequence-diagram`).
   - Unmarked, visually composed DOM groups: repeated bordered/colored nodes with connecting arrows or lines, CSS Grid/Flex arrangement, positioned connectors, SVG/Canvas, or pseudo-element connectors.
   - Avoid marking normal card grids, tables, lists, navigation, or entire document sections as diagrams. If uncertain, inspect screenshots and revise the selection; **automated inference is not infallible**.
4. **Capture each diagram as a separate high-resolution image** at its rendered bounding box. Prefer PNG for mixed CSS/div/span content; optionally retain standalone SVG when fully supported. Include labels, shadows, connectors, and padding without clipping. Preserve aspect ratio. For wide figures, fit within Word's printable width without cropping. If the diagram spans beyond the viewport, enlarge viewport before screenshot.
5. **Replace the entire selected diagram container in an export-only copy of the HTML** with a single `<img>` at the same DOM position, with useful alt text; never delete adjacent paragraphs or duplicate diagram text as paragraphs.
6. **Convert the modified HTML to DOCX**, using Pandoc, a DOCX library, or other available converter. Keep headings as named Word styles, tables as real Word tables, lists and hyperlinks intact, and captions adjacent to images. Use a consistent document design, sensible margins, and page breaks. Do not rasterize normal content.
7. **Verify output.** Render every DOCX page to PDF or PNG if tooling permits. Inspect for missing content, broken images, clipped text/connectors, tiny diagram labels, split images, overflow, broken tables, and misplaced figures. Correct and rerender. Report diagram count and any uncertain or missed figures.

## Included converter

When a local shell is available and `python`, `playwright`, `beautifulsoup4`, `Pillow`, `python-docx`, `pandoc`, and Chromium exist, use `scripts/convert.py` as a starting point:

```bash
python scripts/convert.py input.html output.docx --review
```

The converter provides first-pass automatic diagram detection (explicit markers, named containers, CSS layout heuristics, and standalone `<svg>`/`<canvas>`), browser element screenshots, a detection manifest, and DOCX export. It un-clips diagrams inside scroll containers, widens the viewport for wide content, sizes images at their true CSS size (capped to the printable area), and drops `display:none` content. **Inspect the manifest and the final Word document**; redo any imperfect selections. Use `--review` for an optional PDF/page-image preview if LibreOffice and `pdftoppm` are available. If automatic detection missed a diagram, identify the correct outer container, add `data-export-as-image="true"` (and optionally `data-caption="..."`) to that element in an export copy, and rerun.

**Security defaults:** network access is blocked and only local files inside the HTML's folder (`--asset-root`) can be embedded. Pass `--allow-network` only for trusted HTML that needs remote fonts or images, and say so to the user. Report any entries under `security.blocked_local_references` or `removed_remote_images` in the manifest.

If required tools are unavailable, follow the procedure with whatever browser/Word-generation tools exist, and state the limitation rather than claiming successful diagram preservation.

## Output and quality bar

- Deliver the `.docx`, plus the original HTML unchanged (when needed) and optionally a short conversion report.
- Keep text content and actual tables editable; images may not be editable as Word shapes.
- Diagram screenshots should preserve CSS composition accurately, not merely approximate it.
- Never claim the export is visually verified unless its pages were actually rendered and inspected.
- If the user requests only the final Word file, do not burden them with intermediate page captures.

## Example

This AI-generated HTML is a **diagram**, not a Word-ready set of separate paragraphs:

```html
<div class="request-flow" data-export-as-image="true">
  <span class="node">Publisher</span>
  <span class="arrow">→</span>
  <span class="node">Decision Fabric</span>
  <span class="arrow">→</span>
  <span class="node">Vendor</span>
</div>
```

Capture the outer `<div>` as **one image** and place that image between the adjacent Word paragraphs. Keep any separate headings, prose, and real tables as editable document content.
