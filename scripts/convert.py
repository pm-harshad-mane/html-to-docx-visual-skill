#!/usr/bin/env python3
"""Browser-rendered CSS/HTML diagrams -> PNG, surrounding HTML -> editable DOCX.

Requires: playwright, beautifulsoup4, Pillow, pandoc, Chromium.
Optional: LibreOffice, pdftoppm for preview images. Review manifest for misses.
"""
from __future__ import annotations

import argparse
import base64
import mimetypes
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import unquote, urlparse

from bs4 import BeautifulSoup
from PIL import Image
from playwright.sync_api import sync_playwright


DETECT = r"""() => {
  const namePattern = /(diagram|flowchart|architecture|pipeline|process[-_ ]?flow|sequence[-_ ]?diagram|swimlane|topology|decision[-_ ]?tree|visuali[sz]ation|request[-_ ]?flow|timeline)/i;
  const arrowPattern = /[→←↑↓➜⇢⟶⟵⤷⇨⇒]/;
  const candidates = [];
  const elements = [...document.querySelectorAll('body *')];
  function visible(e) {
    const s = getComputedStyle(e), r = e.getBoundingClientRect();
    return s.display !== 'none' && s.visibility !== 'hidden' && Number(s.opacity) > 0.01 && r.width >= 160 && r.height >= 55;
  }
  function withinExcluded(e) {
    return Boolean(e.closest('nav, aside, header, footer, table, ul, ol, form'));
  }
  function nearestMeaningfulChildren(e) {
    return [...e.children].filter(c => {
      const r = c.getBoundingClientRect(), s = getComputedStyle(c);
      return s.display !== 'none' && r.width >= 12 && r.height >= 10;
    });
  }
  function feature(e) {
    const s = getComputedStyle(e), kids = nearestMeaningfulChildren(e);
    const area = e.getBoundingClientRect();
    const short = (e.innerText || '').length < 2600;
    if (kids.length < 2 || !short || area.height > 1800) return false;
    const layout = /grid|flex/.test(s.display);
    const arrow = arrowPattern.test(e.innerText || '') || [...e.querySelectorAll('[class]')].some(c => /connector|arrow|edge|line-between/i.test(c.className || ''));
    const boxes = kids.filter(c => {
      const x = getComputedStyle(c);
      return x.borderStyle !== 'none' || x.backgroundColor !== 'rgba(0, 0, 0, 0)' || x.boxShadow !== 'none';
    }).length;
    const positions = kids.filter(c => /absolute|fixed/.test(getComputedStyle(c).position)).length;
    const graphics = e.querySelector('svg, canvas') !== null;
    return ((layout && arrow && boxes >= 1) || (layout && arrow && kids.length >= 3) || (graphics && kids.length > 2) || (layout && positions >= 2 && boxes >= 2));
  }
  for (const e of elements) {
    if (withinExcluded(e) || !visible(e)) continue;
    if (/^(script|style|noscript|iframe|img|svg|canvas|p|h[1-6])$/i.test(e.tagName)) continue;
    const attr = e.getAttribute('data-export-as-image');
    const explicit = (attr !== null && attr.toLowerCase() !== 'false') || e.getAttribute('data-docx-render') === 'image' || e.hasAttribute('data-skill-diagram');
    const named = namePattern.test((typeof e.className === 'string' ? e.className : '') + ' ' + (e.id || ''));
    let rank = 0;
    if (explicit) rank = 3;
    else if (named && nearestMeaningfulChildren(e).length >= 1 && (e.innerText || '').length < 2600 && e.getBoundingClientRect().height < 1600) rank = 2;
    else if (feature(e)) rank = 1;
    if (!rank) continue;
    const depth = (()=>{let n=e,d=0;while(n.parentElement){d++;n=n.parentElement;}return d;})();
    candidates.push({e, rank, depth, name: (e.id || (typeof e.className === 'string' ? e.className.trim() : '')) || e.tagName.toLowerCase()});
  }
  candidates.sort((a,b) => b.rank - a.rank || (a.rank===1 ? b.depth-a.depth : a.depth-b.depth));
  const chosen=[];
  for (const c of candidates) {
    if (chosen.some(d => d.e.contains(c.e) || c.e.contains(d.e))) continue;
    // Include the immediately surrounding styled wrapper when the detected
    // flow is the wrapper's only meaningful child (preserves borders/padding).
    if (c.rank===1) {
      const parent=c.e.parentElement;
      if (parent && parent.tagName!=='BODY' && parent.tagName!=='MAIN' && parent.tagName!=='ARTICLE'
          && nearestMeaningfulChildren(parent).length===1 && visible(parent)) {
        const ps=getComputedStyle(parent);
        const styled=ps.borderStyle!=='none' || ps.boxShadow!=='none' ||
          (ps.backgroundColor!=='rgba(0, 0, 0, 0)' && ps.backgroundColor!=='transparent');
        if (styled && parent.getBoundingClientRect().height < 1800) c.e=parent;
      }
    }
    chosen.push(c);
  }
  chosen.sort((a,b) => elements.indexOf(a.e) - elements.indexOf(b.e));
  return chosen.map((c,i)=>{
    const id='diagram-'+String(i+1).padStart(3,'0');
    c.e.setAttribute('data-html-docx-capture-id', id);
    const r=c.e.getBoundingClientRect();
    return {id, detection: c.rank===3?'explicit':c.rank===2?'named':'heuristic', element:c.e.tagName.toLowerCase(), name:c.name.slice(0,120), width:Math.round(r.width), height:Math.round(r.height), alt:(c.e.getAttribute('aria-label') || c.e.getAttribute('data-caption') || c.name || 'Visual diagram').slice(0,160)};
  });
}"""


def which_or_none(name: str) -> str | None:
    return shutil.which(name)


def run_checked(args: list[str], **kwargs):
    proc = subprocess.run(args, text=True, capture_output=True, **kwargs)
    if proc.returncode:
        raise RuntimeError(f"Command failed: {' '.join(args)}\n{proc.stderr[-2200:]}")
    return proc


def preview(docx: Path, target: Path) -> dict:
    lo, ppt = which_or_none('libreoffice'), which_or_none('pdftoppm')
    if not lo or not ppt:
        return {'created': False, 'reason': 'LibreOffice and/or pdftoppm not installed'}
    target.mkdir(parents=True, exist_ok=True)
    # Isolate LO configuration, avoiding locks with the user's application.
    with tempfile.TemporaryDirectory(prefix='html-docx-lo-') as temp:
        run_checked([lo, '-env:UserInstallation=file://' + temp, '--headless', '--convert-to', 'pdf', '--outdir', str(target), str(docx)], timeout=90)
    pdf = target / f'{docx.stem}.pdf'
    if not pdf.exists():
        return {'created': False, 'reason': 'LibreOffice did not produce PDF'}
    run_checked([ppt, '-f', '1', '-l', '99', '-png', '-r', '110', str(pdf), str(target/'page')], timeout=90)
    return {'created': True, 'pdf': str(pdf), 'pages': len(list(target.glob('page-*.png')))}


def inline_local_dependencies(source: Path) -> str:
    """Fallback for browsers/policies that disallow file navigation.

    This can handle relative local CSS, JS, and IMG assets; external network
    resources and JS fetch/XHR are not automatically embeddable.
    """
    soup = BeautifulSoup(source.read_text(encoding='utf8'), 'html.parser')
    def local(path: str) -> Path | None:
        if not path or path.startswith(('data:', 'http:', 'https:', '#')): return None
        parsed = urlparse(path)
        if parsed.scheme: return None
        item = (source.parent / unquote(parsed.path)).resolve()
        return item if item.is_file() else None
    def embedded_css_urls(css: str) -> str:
        def replace(match):
            raw = match.group(1).strip().strip('\"\'')
            item = local(raw)
            if item is None: return match.group(0)
            mime = mimetypes.guess_type(item.name)[0] or 'application/octet-stream'
            blob = base64.b64encode(item.read_bytes()).decode('ascii')
            return f'url(data:{mime};base64,{blob})'
        return re.sub(r'url\(\s*([^)]*?)\s*\)', replace, css)
    for item in list(soup.select('link[href]')):
        if 'stylesheet' not in item.get('rel', []): continue
        path = local(item['href'])
        if path:
            new = soup.new_tag('style')
            new.string = embedded_css_urls(path.read_text(encoding='utf8'))
            item.replace_with(new)
    for item in list(soup.select('style')):
        if item.string:
            item.string.replace_with(embedded_css_urls(str(item.string)))
    for item in list(soup.select('script[src]')):
        path = local(item['src'])
        if path:
            new = soup.new_tag('script')
            new.string = path.read_text(encoding='utf8')
            item.replace_with(new)
    for item in soup.select('img[src]'):
        path = local(item['src'])
        if path:
            mime = mimetypes.guess_type(path.name)[0] or 'application/octet-stream'
            item['src'] = 'data:' + mime + ';base64,' + base64.b64encode(path.read_bytes()).decode('ascii')
    return str(soup)


def main() -> int:
    p = argparse.ArgumentParser(description='Export HTML to DOCX with CSS diagrams as PNG images.')
    p.add_argument('html', type=Path)
    p.add_argument('docx', type=Path)
    p.add_argument('--review', action='store_true', help='Create optional PDF and page PNGs for visual QA')
    p.add_argument('--viewport-width', type=int, default=1400)
    p.add_argument('--scale', type=float, default=2.5)
    p.add_argument('--wait-ms', type=int, default=600)
    args = p.parse_args()
    source, target = args.html.expanduser().resolve(), args.docx.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() not in ('.html','.htm'):
        p.error(f'HTML file not found: {source}')
    if not which_or_none('pandoc'):
        p.error('pandoc is required, but was not found in PATH')
    target.parent.mkdir(parents=True,exist_ok=True)
    work = target.parent / f'{target.stem}_export_assets'
    assets = work / 'images'
    assets.mkdir(parents=True,exist_ok=True)

    with sync_playwright() as play:
        browser_path = which_or_none('chromium') or which_or_none('chromium-browser') or which_or_none('google-chrome')
        launch_args = {'headless':True}
        if browser_path: launch_args['executable_path'] = browser_path
        try:
            browser = play.chromium.launch(**launch_args)
        except Exception as exc:
            raise RuntimeError('Chromium browser unavailable. Run `python -m playwright install chromium` or install Chromium.') from exc
        context = browser.new_context(viewport={'width':args.viewport_width,'height':900}, device_scale_factor=args.scale, reduced_motion='reduce', color_scheme='light')
        page = context.new_page()
        try:
            page.goto(source.as_uri(), wait_until='load',timeout=30000)
        except Exception as exc:
            if 'ERR_BLOCKED_BY_ADMINISTRATOR' not in str(exc) and 'ERR_ACCESS_DENIED' not in str(exc):
                raise
            print('Browser policy blocked file navigation; using local-content fallback. External resources may not render.',file=sys.stderr)
            page.close()
            page = context.new_page()
            page.set_content(inline_local_dependencies(source),wait_until='load',timeout=30000)
        try: page.evaluate('async ()=>await document.fonts.ready')
        except Exception: pass
        page.wait_for_timeout(args.wait_ms)
        result = page.evaluate(DETECT)
        if not result:
            print('WARNING: No diagrams detected. Check output; consider explicit data-export-as-image markers.',file=sys.stderr)
        # Add room for content that extends wider than the original viewport.
        export_images = {}
        for idx, info in enumerate(result, 1):
            target_png = assets / (info['id']+'.png')
            loc = page.locator(f'[data-html-docx-capture-id="{info["id"]}"]')
            loc.screenshot(path=str(target_png), animations='disabled', timeout=30000)
            iw, ih = Image.open(target_png).size
            info['image_pixels'] = [iw,ih]
            info['image'] = str(target_png)
            export_images[info['id']] = target_png
        html = page.content()
        browser.close()

    soup = BeautifulSoup(html,'html.parser')
    for script in soup.select('script,noscript'):
        script.decompose()
    for style in soup.select('style'):
        style.decompose()
    for info in result:
        node = soup.find(attrs={'data-html-docx-capture-id':info['id']})
        if node is None:
            raise RuntimeError(f'Could not locate detected element {info["id"]} in export DOM')
        filename = Path('images') / f'{info["id"]}.png'
        with Image.open(export_images[info['id']]) as im:
            width_inches = min(6.2, max(2.5, (im.width/args.scale)/150))
        image = soup.new_tag('img', src=str(filename), alt='Diagram: '+info['alt'], width=f'{width_inches:.2f}in', **{'data-diagram-snapshot':'true'})
        para = soup.new_tag('p')
        para.append(image)
        node.replace_with(para)
        info['word_width_inches'] = round(width_inches,2)
    # Recover image references originally relative to the source HTML.
    for img in soup.find_all('img'):
        src = img.get('src','')
        if not src or img.get('data-diagram-snapshot')=='true' or src.startswith(('http:','https:','data:','file:','#')):
            continue
        parsed = urlparse(src)
        if parsed.scheme: continue
        source_img = (source.parent / unquote(parsed.path)).resolve()
        if source_img.is_file():
            img['src'] = str(source_img)
    for img in soup.select('img[data-diagram-snapshot]'):
        del img['data-diagram-snapshot']
    for e in soup.select('[data-html-docx-capture-id]'):
        del e['data-html-docx-capture-id']
    staged_html = work/'export.html'
    staged_html.write_text(str(soup), encoding='utf8')
    cmd = ['pandoc', str(staged_html), '--from=html', '--to=docx', '--standalone', f'--resource-path={work}:{source.parent}', '-o', str(target)]
    run_checked(cmd, timeout=120)
    # Apply restrained Word defaults to elements Pandoc generated as editable content.
    try:
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        doc = Document(target)
        sec=doc.sections[0]
        sec.top_margin=Inches(0.82); sec.bottom_margin=Inches(0.8)
        sec.left_margin=Inches(0.88); sec.right_margin=Inches(0.88)
        for tbl in doc.tables:
            tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
            try: tbl.style='Table Grid'
            except KeyError: pass
            for cell in tbl._cells:
                for para in cell.paragraphs:
                    for run in para.runs: run.font.size=Pt(9.5)
        for para in doc.paragraphs:
            if para._p.xpath('.//w:drawing'):
                para.alignment=WD_ALIGN_PARAGRAPH.CENTER
        doc.save(target)
    except ImportError:
        print('python-docx not installed; skipped optional Word-style cleanup.',file=sys.stderr)
    if not target.exists() or target.stat().st_size < 1000:
        raise RuntimeError('Expected DOCX output was not created')
    report = {
        'input':str(source),'output':str(target),'diagram_count':len(result),'diagrams':result,
        'warning':'Automatic detection is heuristic. Inspect the diagram images and all DOCX pages; rerun after marking any missed diagram.',
    }
    if args.review:
        try: report['preview'] = preview(target, work/'preview')
        except Exception as exc: report['preview']={'created':False,'reason':str(exc)}
    manifest = work / 'conversion_report.json'
    manifest.write_text(json.dumps(report,indent=2),encoding='utf8')
    print(f'Created: {target}\nDetected diagrams: {len(result)}\nReport: {manifest}')
    if args.review:
        print('Preview:', report.get('preview'))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
