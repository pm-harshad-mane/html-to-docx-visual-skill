#!/usr/bin/env python3
"""Browser-rendered CSS/HTML diagrams -> PNG, surrounding HTML -> editable DOCX.

Requires: playwright, beautifulsoup4, Pillow, pandoc, Chromium. Recommended: python-docx.
Optional: LibreOffice and pdftoppm for preview images. Review the manifest for misses.

Security defaults: the page renders with network access blocked, and only local
files inside the asset root (default: the HTML file's folder) can be inlined or
embedded in the DOCX. Use --allow-network / --asset-root to relax deliberately.
"""
from __future__ import annotations

import argparse
import base64
import json
import math
import mimetypes
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

# DOCX page geometry in inches. Pandoc's default reference.docx is US Letter;
# postprocess_docx() re-fits images against the real page size afterwards.
MARGIN_TOP, MARGIN_BOTTOM, MARGIN_SIDE = 0.82, 0.8, 0.88
DEFAULT_PAGE_W, DEFAULT_PAGE_H = 8.5, 11.0
CAPTION_ALLOWANCE_IN = 0.6  # vertical room kept free for a caption line
CSS_PX_PER_INCH = 96
MAX_VIEWPORT_WIDTH = 4096
VIEWPORT_HEIGHT = 900

# Elements (other than captured diagrams) that are never passed to Pandoc:
# they either don't render in Word or could make Pandoc fetch extra resources.
STRIP_TAGS = 'script, noscript, style, template, iframe, object, embed, video, audio, source, track, link, base'
# Parents under which a diagram image may be wrapped in its own block element.
BLOCK_PARENTS = {'body', 'div', 'section', 'article', 'main', 'header', 'footer', 'aside',
                 'li', 'td', 'th', 'blockquote', 'details', 'dd'}


DETECT = r"""() => {
  const namePattern = /(diagram|flowchart|architecture|pipeline|process[-_ ]?flow|sequence[-_ ]?diagram|swimlane|topology|decision[-_ ]?tree|visuali[sz]ation|request[-_ ]?flow|timeline)/i;
  const arrowPattern = /[→←↑↓➜⇢⟶⟵⤷⇨⇒]/;
  const connectorClass = /connector|arrow|edge|line-between/i;
  const SVG_NS = 'http://www.w3.org/2000/svg';
  const elements = [...document.querySelectorAll('body *')];
  const order = new Map(elements.map((e, i) => [e, i]));
  const styleCache = new Map();
  const cs = e => { let s = styleCache.get(e); if (!s) { s = getComputedStyle(e); styleCache.set(e, s); } return s; };
  const className = e => typeof e.className === 'string' ? e.className : (e.getAttribute('class') || '');

  // Ancestors of connector-classed elements, computed in one pass instead of
  // running querySelectorAll on every element (O(n) instead of O(n^2)).
  const hasConnector = new Set();
  for (const c of elements) {
    if (!connectorClass.test(className(c))) continue;
    for (let n = c.parentElement; n && !hasConnector.has(n); n = n.parentElement) hasConnector.add(n);
  }

  const shown = e => { const s = cs(e); return s.display !== 'none' && s.visibility !== 'hidden' && Number(s.opacity) > 0.01; };
  const hasSize = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const bigEnough = e => { const r = e.getBoundingClientRect(); return r.width >= 160 && r.height >= 55; };
  const withinExcluded = e => Boolean(e.closest('nav, aside, header, footer, table, ul, ol, form'));
  const kidsOf = e => [...e.children].filter(c => {
    const r = c.getBoundingClientRect();
    return cs(c).display !== 'none' && r.width >= 12 && r.height >= 10;
  });
  const painted = s => s.borderStyle !== 'none' || s.boxShadow !== 'none' ||
    (s.backgroundColor !== 'rgba(0, 0, 0, 0)' && s.backgroundColor !== 'transparent');
  const depthOf = e => { let d = 0; for (let n = e; n.parentElement; n = n.parentElement) d++; return d; };
  const isExplicit = e => {
    const attr = e.getAttribute('data-export-as-image');
    return (attr !== null && attr.toLowerCase() !== 'false') || e.getAttribute('data-docx-render') === 'image' || e.hasAttribute('data-skill-diagram');
  };

  function feature(e, text) {
    const kids = kidsOf(e), r = e.getBoundingClientRect();
    if (kids.length < 2 || text.length >= 2600 || r.height > 1800) return false;
    const layout = /grid|flex/.test(cs(e).display);
    const arrow = arrowPattern.test(text) || hasConnector.has(e);
    const boxes = kids.filter(c => painted(cs(c))).length;
    const positions = kids.filter(c => /absolute|fixed/.test(cs(c).position)).length;
    const graphics = e.querySelector('svg, canvas') !== null;
    return (layout && arrow && (boxes >= 1 || kids.length >= 3)) || (graphics && kids.length > 2) || (layout && positions >= 2 && boxes >= 2);
  }

  function describe(e, i) {
    const text = x => (x || '').replace(/\s+/g, ' ').trim();
    const by = e.getAttribute('aria-labelledby');
    const labelled = by ? by.split(/\s+/).map(id => text(document.getElementById(id)?.textContent)).filter(Boolean).join(' ') : '';
    const fig = e.closest('figure');
    const figcaption = fig ? text(fig.querySelector('figcaption')?.textContent) : '';
    const svgTitle = e.namespaceURI === SVG_NS ? text(e.querySelector('title')?.textContent) : '';
    const heading = text(e.querySelector('h1, h2, h3, h4, h5, h6, [role=heading]')?.textContent);
    const found = [e.getAttribute('aria-label'), labelled, e.getAttribute('data-caption'), e.getAttribute('title'),
                   e.getAttribute('alt'), svgTitle, figcaption, heading].map(text).find(Boolean);
    return (found || 'Diagram ' + (i + 1)).slice(0, 160);
  }

  const candidates = [], skipped = [];
  for (const e of elements) {
    if (e.namespaceURI === SVG_NS && e.ownerSVGElement) continue;  // shapes inside an SVG; only the outer <svg> counts
    const name = (e.id || className(e).trim() || e.tagName.toLowerCase()).slice(0, 120);
    // Explicit markers always win: they bypass exclusion zones, tag filters and size floors.
    if (isExplicit(e)) {
      if (shown(e) && hasSize(e)) candidates.push({e, rank: 4});
      else skipped.push({name, reason: 'explicitly marked but hidden or zero-size'});
      continue;
    }
    if (/^(script|style|noscript|template|iframe|img|p|h[1-6])$/i.test(e.tagName)) continue;
    if (withinExcluded(e) || !shown(e) || !bigEnough(e)) continue;
    if (/^(svg|canvas)$/i.test(e.tagName)) { candidates.push({e, rank: 1}); continue; }
    const text = e.innerText || '';
    const r = e.getBoundingClientRect();
    if (namePattern.test(className(e) + ' ' + (e.id || '')) && kidsOf(e).length >= 1 && text.length < 2600 && r.height < 1600) candidates.push({e, rank: 3});
    else if (feature(e, text)) candidates.push({e, rank: 2});
  }
  for (const c of candidates) c.depth = depthOf(c.e);
  // Heuristic containers: deepest first (tightest fit). Everything else: outermost first.
  candidates.sort((a, b) => b.rank - a.rank || (a.rank === 2 ? b.depth - a.depth : a.depth - b.depth));

  const chosen = [];
  const overlaps = e => chosen.some(d => d.e.contains(e) || e.contains(d.e));
  for (const c of candidates) {
    if (overlaps(c.e)) continue;
    // Include the immediately surrounding styled wrapper when the detected
    // element is the wrapper's only meaningful child (preserves borders/padding).
    if (c.rank <= 2) {
      const parent = c.e.parentElement;
      if (parent && !/^(BODY|MAIN|ARTICLE)$/.test(parent.tagName) && kidsOf(parent).length === 1 &&
          shown(parent) && bigEnough(parent) && painted(cs(parent)) &&
          parent.getBoundingClientRect().height < 1800 && !overlaps(parent)) c.e = parent;
    }
    chosen.push(c);
  }
  chosen.sort((a, b) => order.get(a.e) - order.get(b.e));
  const labels = {4: 'explicit', 3: 'named', 2: 'heuristic', 1: 'graphic'};
  const diagrams = chosen.map((c, i) => {
    const id = 'diagram-' + String(i + 1).padStart(3, '0');
    c.e.setAttribute('data-html-docx-capture-id', id);
    const r = c.e.getBoundingClientRect();
    return {
      id, detection: labels[c.rank], element: c.e.tagName.toLowerCase(),
      name: (c.e.id || className(c.e).trim() || c.e.tagName.toLowerCase()).slice(0, 120),
      width: Math.round(r.width), height: Math.round(r.height), alt: describe(c.e, i),
      caption: (c.e.getAttribute('data-caption') || '').replace(/\s+/g, ' ').trim().slice(0, 300),
    };
  });
  return {diagrams, skipped};
}"""

# Un-clip a captured element whose content overflows its box (scroll containers,
# fixed-size boxes with spilling content) so the screenshot shows all of it.
EXPAND = r"""(id) => {
  const e = document.querySelector(`[data-html-docx-capture-id="${id}"]`);
  if (!e) return {expanded: false, right: 0, width: 0, height: 0};
  let expanded = false;
  if (e instanceof HTMLElement) {
    const s = getComputedStyle(e), px = v => parseFloat(v) || 0;
    const clipX = /hidden|clip/.test(s.overflowX), clipY = /hidden|clip/.test(s.overflowY);
    if (!clipX && e.clientWidth > 0 && e.scrollWidth > e.clientWidth + 1) {
      const w = s.boxSizing === 'border-box'
        ? e.scrollWidth + px(s.borderLeftWidth) + px(s.borderRightWidth)
        : e.scrollWidth - px(s.paddingLeft) - px(s.paddingRight);
      e.style.setProperty('overflow-x', 'visible', 'important');
      e.style.setProperty('max-width', 'none', 'important');
      e.style.setProperty('width', w + 'px', 'important');
      expanded = true;
    }
    if (!clipY && e.clientHeight > 0 && e.scrollHeight > e.clientHeight + 1) {
      e.style.setProperty('overflow-y', 'visible', 'important');
      e.style.setProperty('max-height', 'none', 'important');
      e.style.setProperty('height', 'auto', 'important');
      expanded = true;
    }
  }
  const r = e.getBoundingClientRect();
  return {expanded, right: Math.ceil(r.right + window.scrollX), width: Math.round(r.width), height: Math.round(r.height)};
}"""

# Mark HTML elements that are not rendered (display:none), so hidden tabs,
# modals and templates don't leak into the Word document as visible text.
MARK_HIDDEN = r"""() => {
  let count = 0;
  for (const e of document.querySelectorAll('body *')) {
    if (e.namespaceURI !== 'http://www.w3.org/1999/xhtml') continue;
    if (/^(script|style|noscript|template)$/i.test(e.tagName)) continue;
    if (e.closest('[data-html-docx-capture-id]')) continue;
    if (getComputedStyle(e).display !== 'none') continue;
    if (!e.parentElement || !e.parentElement.closest('[data-html-docx-hidden]')) count++;
    e.setAttribute('data-html-docx-hidden', '');
  }
  return count;
}"""


def is_remote(ref: str) -> bool:
    ref = ref.strip()
    return ref.startswith('//') or urlparse(ref).scheme.lower() in ('http', 'https')


def is_local_ref(ref: str) -> bool:
    ref = ref.strip()
    if not ref or ref.startswith(('data:', '#', '//')):
        return False
    return urlparse(ref).scheme.lower() in ('', 'file')


def confined_path(ref: str, base_dir: Path, root: Path) -> Path | None:
    """Resolve a local URL reference to an existing file inside ``root``.

    Returns None for remote/data/fragment references, missing files, and any
    path that escapes ``root`` (via ``..``, absolute paths, file: URLs or symlinks).
    """
    ref = (ref or '').strip()
    if not is_local_ref(ref):
        return None
    parsed = urlparse(ref)
    if parsed.scheme.lower() == 'file':
        candidate = Path(url2pathname(parsed.path))
    else:
        candidate = base_dir / url2pathname(unquote(parsed.path))
    try:
        candidate = candidate.resolve(strict=True)
        root = root.resolve(strict=True)
    except (OSError, RuntimeError):
        return None
    if not candidate.is_file() or not candidate.is_relative_to(root):
        return None
    return candidate


def fit_width_inches(css_width: float, css_height: float,
                     page_w: float = DEFAULT_PAGE_W, page_h: float = DEFAULT_PAGE_H) -> float:
    """Word width for a figure: its true CSS size (96 px/in), shrunk to fit the printable area."""
    printable_w = page_w - 2 * MARGIN_SIDE
    printable_h = page_h - MARGIN_TOP - MARGIN_BOTTOM - CAPTION_ALLOWANCE_IN
    w, h = css_width / CSS_PX_PER_INCH, css_height / CSS_PX_PER_INCH
    if w <= 0 or h <= 0:
        return math.floor(printable_w * 100) / 100
    # Round down so the rounded size can never exceed the printable area.
    return math.floor(w * min(1.0, printable_w / w, printable_h / h) * 100) / 100


def run_checked(args: list[str], **kwargs):
    proc = subprocess.run(args, text=True, capture_output=True, **kwargs)
    if proc.returncode:
        raise RuntimeError(f"Command failed: {' '.join(args)}\n{proc.stderr[-2200:]}")
    return proc


def find_soffice() -> str | None:
    for name in ('soffice', 'libreoffice'):
        found = shutil.which(name)
        if found:
            return found
    for candidate in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
                      r'C:\Program Files\LibreOffice\program\soffice.exe'):
        if Path(candidate).is_file():
            return candidate
    return None


def preview(docx: Path, target: Path) -> dict:
    lo, ppt = find_soffice(), shutil.which('pdftoppm')
    if not lo or not ppt:
        return {'created': False, 'reason': 'LibreOffice (soffice) and/or pdftoppm not installed'}
    target.mkdir(parents=True, exist_ok=True)
    for old in target.glob('page-*.png'):
        old.unlink()
    # Isolate LO configuration, avoiding locks with the user's application.
    with tempfile.TemporaryDirectory(prefix='html-docx-lo-') as temp:
        run_checked([lo, '-env:UserInstallation=' + Path(temp).as_uri(), '--headless', '--convert-to', 'pdf',
                     '--outdir', str(target), str(docx)], timeout=180)
    pdf = target / f'{docx.stem}.pdf'
    if not pdf.exists():
        return {'created': False, 'reason': 'LibreOffice did not produce PDF'}
    run_checked([ppt, '-png', '-r', '110', str(pdf), str(target / 'page')], timeout=180)
    return {'created': True, 'pdf': str(pdf), 'pages': len(list(target.glob('page-*.png')))}


def _data_uri(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or 'application/octet-stream'
    return f'data:{mime};base64,' + base64.b64encode(path.read_bytes()).decode('ascii')


def inline_local_dependencies(source: Path, root: Path, blocked: list[str]) -> str:
    """Fallback for browsers/policies that disallow file navigation.

    Inlines relative CSS, JS and IMG assets that live inside ``root``; anything
    outside it is left unresolved and recorded in ``blocked``. External network
    resources and JS fetch/XHR are not embedded.
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(source.read_text(encoding='utf8'), 'html.parser')

    def local(ref: str, base: Path) -> Path | None:
        path = confined_path(ref, base, root)
        if path is None and is_local_ref(ref):
            blocked.append(ref.strip())
        return path

    def embed_css_urls(css: str, base: Path) -> str:
        def replace(match):
            path = local(match.group(1).strip().strip('"\''), base)
            return f'url({_data_uri(path)})' if path else match.group(0)
        return re.sub(r'url\(\s*([^)]*?)\s*\)', replace, css)

    for item in list(soup.select('link[href]')):
        if 'stylesheet' not in (item.get('rel') or []):
            continue
        path = local(item['href'], source.parent)
        if path:
            new = soup.new_tag('style')
            new.string = embed_css_urls(path.read_text(encoding='utf8'), path.parent)
            item.replace_with(new)
    for item in list(soup.select('style')):
        if item.string:
            item.string.replace_with(embed_css_urls(str(item.string), source.parent))
    for item in list(soup.select('script[src]')):
        path = local(item['src'], source.parent)
        if path:
            new = soup.new_tag('script')
            new.string = path.read_text(encoding='utf8')
            item.replace_with(new)
    for item in soup.select('img[src]'):
        path = local(item['src'], source.parent)
        if path:
            item['src'] = _data_uri(path)
    return str(soup)


def _widen_viewport(page, needed: int, current: int) -> int:
    needed = min(int(needed), MAX_VIEWPORT_WIDTH)
    if needed <= current:
        return current
    page.set_viewport_size({'width': needed, 'height': VIEWPORT_HEIGHT})
    page.wait_for_timeout(250)
    return needed


def render_and_capture(source: Path, root: Path, args, assets: Path, report: dict) -> tuple[str, list[dict]]:
    """Render the page, detect diagrams, screenshot them; return (rendered HTML, diagram infos)."""
    from PIL import Image
    from playwright.sync_api import sync_playwright

    security = report['security']
    with sync_playwright() as play:
        launch: dict = {'headless': True}
        browser_path = shutil.which('chromium') or shutil.which('chromium-browser') or shutil.which('google-chrome')
        if browser_path:
            launch['executable_path'] = browser_path
        if not args.allow_network:
            # Defence in depth: a dead proxy (loopback included), failing DNS, and no
            # non-proxied WebRTC, on top of the request router and offline mode below.
            launch['proxy'] = {'server': 'http://127.0.0.1:9', 'bypass': '<-loopback>'}
            launch['args'] = ['--host-resolver-rules=MAP * ~NOTFOUND', '--dns-prefetch-disable',
                              '--force-webrtc-ip-handling-policy=disable_non_proxied_udp']
        try:
            browser = play.chromium.launch(**launch)
        except Exception as exc:
            raise RuntimeError('Chromium browser unavailable. Run `python -m playwright install chromium` or install Chromium.') from exc
        try:
            context = browser.new_context(
                viewport={'width': args.viewport_width, 'height': VIEWPORT_HEIGHT}, device_scale_factor=args.scale,
                reduced_motion='reduce', color_scheme='light', service_workers='block', accept_downloads=False,
                offline=not args.allow_network)
            if not args.allow_network:
                def guard(route):
                    url = route.request.url
                    if url.startswith(('file:', 'data:', 'blob:', 'about:')):
                        route.continue_()
                    else:
                        security['blocked_requests'].append(url[:300])
                        route.abort('blockedbyclient')
                context.route('**/*', guard)
            page = context.new_page()
            try:
                page.goto(source.as_uri(), wait_until='load', timeout=30000)
            except Exception as exc:
                if 'ERR_BLOCKED_BY_ADMINISTRATOR' not in str(exc) and 'ERR_ACCESS_DENIED' not in str(exc):
                    raise
                print('Browser policy blocked file navigation; using local-content fallback. External resources may not render.', file=sys.stderr)
                page.close()
                page = context.new_page()
                page.set_content(inline_local_dependencies(source, root, security['blocked_local_references']),
                                 wait_until='load', timeout=30000)
            try:
                page.evaluate('async () => await document.fonts.ready')
            except Exception:
                pass
            page.wait_for_timeout(args.wait_ms)

            # Lay out wide pages at their natural width instead of an artificial edge.
            width = _widen_viewport(page, page.evaluate('document.documentElement.scrollWidth'), args.viewport_width)
            detected = page.evaluate(DETECT)
            diagrams, report['skipped'] = detected['diagrams'], detected['skipped']
            if not diagrams:
                print('WARNING: No diagrams detected. Check output; consider explicit data-export-as-image markers.', file=sys.stderr)

            rightmost = 0
            for info in diagrams:
                result = page.evaluate(EXPAND, info['id'])
                info['expanded_overflow'] = result['expanded']
                rightmost = max(rightmost, result['right'])
            width = _widen_viewport(page, rightmost + 32, width)
            if rightmost + 32 > MAX_VIEWPORT_WIDTH:
                report['warnings'].append(f'A diagram is wider than {MAX_VIEWPORT_WIDTH}px and may be clipped.')
            report['viewport_width'] = width
            report['hidden_elements_removed'] = page.evaluate(MARK_HIDDEN)

            for info in diagrams:
                png = assets / f"{info['id']}.png"
                page.locator(f'[data-html-docx-capture-id="{info["id"]}"]').screenshot(
                    path=str(png), animations='disabled', timeout=30000)
                with Image.open(png) as im:
                    info['image_pixels'] = [im.width, im.height]
                info['image'] = str(png)
            html = page.content()
        finally:
            browser.close()
    return html, diagrams


def build_export_html(html: str, diagrams: list[dict], source: Path, root: Path,
                      allow_network: bool, report: dict) -> str:
    from bs4 import BeautifulSoup

    security = report['security']
    soup = BeautifulSoup(html, 'html.parser')

    for info in diagrams:
        node = soup.find(attrs={'data-html-docx-capture-id': info['id']})
        if node is None:
            raise RuntimeError(f'Could not locate detected element {info["id"]} in export DOM')
        image = soup.new_tag('img', attrs={
            'src': f'images/{info["id"]}.png', 'alt': info['alt'],
            'width': f'{info["word_width_inches"]:.2f}in', 'data-diagram-snapshot': 'true'})
        in_figure = node.find_parent('figure') is not None
        block = node.parent is not None and node.parent.name in BLOCK_PARENTS
        if block and not in_figure and info.get('caption'):
            wrapper = soup.new_tag('figure')
            wrapper.append(image)
            caption = soup.new_tag('figcaption')
            caption.string = info['caption']
            wrapper.append(caption)
        elif block and not in_figure:
            wrapper = soup.new_tag('p')
            wrapper.append(image)
        else:  # inline context or existing <figure>: keep the surrounding structure valid
            wrapper = image
        node.replace_with(wrapper)

    for element in soup.select('[data-html-docx-hidden]'):
        if not element.decomposed:
            element.decompose()
    for element in soup.select(STRIP_TAGS):
        if not element.decomposed:
            element.decompose()

    # Only snapshots and confined local images may reach Pandoc; it would
    # otherwise read any path or fetch any URL it is given.
    for img in soup.find_all('img'):
        if img.get('data-diagram-snapshot') == 'true':
            del img['data-diagram-snapshot']
            continue
        if img.has_attr('srcset'):
            del img['srcset']
        src = (img.get('src') or '').strip()
        if not src or src.startswith('data:'):
            continue
        if is_remote(src):
            if allow_network:
                continue
            security['removed_remote_images'].append(src[:300])
        else:
            path = confined_path(src, source.parent, root)
            if path is not None:
                img['src'] = str(path)
                continue
            security['blocked_local_references'].append(src[:300])
        alt = (img.get('alt') or '').strip()
        img.replace_with(f'[image: {alt}]' if alt else '')

    for element in soup.select('[data-html-docx-capture-id]'):
        del element['data-html-docx-capture-id']
    return str(soup)


def _set_table_borders(table, size: int = 4, color: str = 'A0A7B4') -> None:
    """Give a table visible borders directly in XML (Pandoc's reference.docx has no 'Table Grid' style)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn('w:tblBorders'))
    if existing is not None:
        tbl_pr.remove(existing)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        borders.append(el)
    # Keep OOXML schema order: tblBorders precedes these siblings.
    successors = {qn(f'w:{n}') for n in ('shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription')}
    for index, child in enumerate(tbl_pr):
        if child.tag in successors:
            tbl_pr.insert(index, borders)
            break
    else:
        tbl_pr.append(borders)


def postprocess_docx(target: Path) -> None:
    """Apply restrained Word defaults to the editable content Pandoc generated."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt

    doc = Document(target)
    for sec in doc.sections:
        sec.top_margin, sec.bottom_margin = Inches(MARGIN_TOP), Inches(MARGIN_BOTTOM)
        sec.left_margin = sec.right_margin = Inches(MARGIN_SIDE)
    first = doc.sections[0]
    page_w = first.page_width or Inches(DEFAULT_PAGE_W)
    page_h = first.page_height or Inches(DEFAULT_PAGE_H)
    printable_w = page_w - first.left_margin - first.right_margin
    printable_h = page_h - first.top_margin - first.bottom_margin - Inches(CAPTION_ALLOWANCE_IN)
    for shape in doc.inline_shapes:  # safety net for non-Letter reference documents
        w, h = shape.width, shape.height
        if w and h:
            scale = min(1.0, printable_w / w, printable_h / h)
            if scale < 1.0:
                shape.width, shape.height = int(w * scale), int(h * scale)
    for tbl in doc.tables:
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(tbl)
        for cell in tbl._cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9.5)
    for para in doc.paragraphs:
        if para._p.xpath('.//w:drawing'):
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(target)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description='Export HTML to DOCX with CSS diagrams as PNG images.')
    p.add_argument('html', type=Path)
    p.add_argument('docx', type=Path)
    p.add_argument('--review', action='store_true', help='Create optional PDF and page PNGs for visual QA')
    p.add_argument('--viewport-width', type=int, default=1400)
    p.add_argument('--scale', type=float, default=2.5, help='Device scale factor for diagram screenshots')
    p.add_argument('--wait-ms', type=int, default=600, help='Extra wait after load for JS-rendered diagrams')
    p.add_argument('--asset-root', type=Path, default=None,
                   help='Only local files inside this folder may be loaded into the DOCX (default: the HTML file\'s folder)')
    p.add_argument('--allow-network', action='store_true',
                   help='Let the page and Pandoc fetch remote resources (off by default; only for trusted HTML)')
    args = p.parse_args(argv)

    source, target = args.html.expanduser().resolve(), args.docx.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() not in ('.html', '.htm'):
        p.error(f'HTML file not found: {source}')
    root = args.asset_root.expanduser().resolve() if args.asset_root else source.parent
    if not root.is_dir() or not source.is_relative_to(root):
        p.error(f'--asset-root must be an existing folder that contains the HTML file: {root}')
    if args.scale <= 0 or args.viewport_width < 320 or args.wait_ms < 0:
        p.error('--scale must be > 0, --viewport-width >= 320 and --wait-ms >= 0')
    if not shutil.which('pandoc'):
        p.error('pandoc is required, but was not found in PATH')
    try:
        import bs4, PIL, playwright  # noqa: F401,E401
    except ImportError as exc:
        p.error(f'missing Python dependency ({exc.name}); run: python -m pip install -r requirements.txt')

    target.parent.mkdir(parents=True, exist_ok=True)
    work = target.parent / f'{target.stem}_export_assets'
    assets = work / 'images'
    assets.mkdir(parents=True, exist_ok=True)
    for stale in assets.glob('diagram-*.png'):
        stale.unlink()

    report: dict = {
        'input': str(source), 'output': str(target), 'warnings': [],
        'security': {'network': 'allowed' if args.allow_network else 'blocked', 'asset_root': str(root),
                     'blocked_requests': [], 'blocked_local_references': [], 'removed_remote_images': []},
    }
    html, diagrams = render_and_capture(source, root, args, assets, report)
    for info in diagrams:
        css_w, css_h = (px / args.scale for px in info['image_pixels'])
        info['word_width_inches'] = fit_width_inches(css_w, css_h)

    staged_html = work / 'export.html'
    staged_html.write_text(build_export_html(html, diagrams, source, root, args.allow_network, report), encoding='utf8')
    # Resource path is only the work dir: every non-snapshot image is already an absolute, confined path.
    run_checked(['pandoc', str(staged_html), '--from=html', '--to=docx', '--standalone',
                 f'--resource-path={work}', '-o', str(target)], timeout=180)
    try:
        postprocess_docx(target)
    except ImportError:
        report['warnings'].append('python-docx not installed; skipped Word layout cleanup (margins, table borders, image fit).')
        print(report['warnings'][-1], file=sys.stderr)
    if not target.exists() or target.stat().st_size < 1000:
        raise RuntimeError('Expected DOCX output was not created')

    security = report['security']
    for key in ('blocked_local_references', 'removed_remote_images'):
        security[key] = sorted(set(security[key]))
    if security['blocked_local_references']:
        print(f'Blocked {len(security["blocked_local_references"])} local reference(s) outside the asset root or missing; see report.', file=sys.stderr)
    if security['removed_remote_images']:
        print(f'Removed {len(security["removed_remote_images"])} remote image(s); rerun with --allow-network only for trusted HTML.', file=sys.stderr)
    report.update({
        'diagram_count': len(diagrams), 'diagrams': diagrams,
        'note': 'Automatic detection is heuristic. Inspect the diagram images and all DOCX pages; rerun after marking any missed diagram.',
    })
    if args.review:
        try:
            report['preview'] = preview(target, work / 'preview')
        except Exception as exc:
            report['preview'] = {'created': False, 'reason': str(exc)}
    manifest = work / 'conversion_report.json'
    manifest.write_text(json.dumps(report, indent=2), encoding='utf8')
    print(f'Created: {target}\nDetected diagrams: {len(diagrams)}\nReport: {manifest}')
    if args.review:
        print('Preview:', report.get('preview'))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
