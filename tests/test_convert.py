"""Tests for scripts/convert.py.

Unit tests need only the standard library (plus bs4 where noted). The end-to-end
tests are skipped unless Pandoc, Playwright and a Chromium build are available.
"""
from __future__ import annotations

import json
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / 'scripts'))
import convert  # noqa: E402  (module has no third-party imports at load time)

PNG_1PX = b'\x89PNG\r\n\x1a\n' + bytes(16)  # signature stub; content is never decoded


# --- path confinement -------------------------------------------------------

def test_confined_path_allows_file_inside_root(tmp_path):
    (tmp_path / 'img').mkdir()
    (tmp_path / 'img' / 'a b.png').write_bytes(PNG_1PX)
    found = convert.confined_path('img/a%20b.png', tmp_path, tmp_path)
    assert found == (tmp_path / 'img' / 'a b.png').resolve()


def test_confined_path_rejects_traversal(tmp_path):
    root = tmp_path / 'site'
    root.mkdir()
    (tmp_path / 'secret.txt').write_text('token')
    assert convert.confined_path('../secret.txt', root, root) is None
    assert convert.confined_path(str(tmp_path / 'secret.txt'), root, root) is None


def test_confined_path_rejects_file_url_outside_root(tmp_path):
    root = tmp_path / 'site'
    root.mkdir()
    secret = tmp_path / 'secret.txt'
    secret.write_text('token')
    assert convert.confined_path(secret.as_uri(), root, root) is None


def test_confined_path_rejects_symlink_escape(tmp_path):
    root = tmp_path / 'site'
    root.mkdir()
    secret = tmp_path / 'secret.txt'
    secret.write_text('token')
    try:
        (root / 'link.txt').symlink_to(secret)
    except OSError:
        pytest.skip('symlinks not supported here')
    assert convert.confined_path('link.txt', root, root) is None


@pytest.mark.parametrize('ref', ['https://example.com/a.png', '//example.com/a.png',
                                 'data:image/png;base64,AAAA', '#frag', '', 'javascript:alert(1)'])
def test_confined_path_ignores_non_local(tmp_path, ref):
    assert convert.confined_path(ref, tmp_path, tmp_path) is None


def test_ref_classifiers():
    assert convert.is_remote('https://x/y.png') and convert.is_remote('//x/y.png')
    assert not convert.is_remote('img/y.png')
    assert convert.is_local_ref('../y.png') and convert.is_local_ref('file:///etc/hosts')
    assert not convert.is_local_ref('data:,x') and not convert.is_local_ref('https://x')


# --- figure sizing -----------------------------------------------------------

def test_fit_width_uses_css_inches():
    assert convert.fit_width_inches(480, 200) == 5.0  # 480 css px == 5 in


def test_fit_width_caps_to_printable_width():
    printable = convert.DEFAULT_PAGE_W - 2 * convert.MARGIN_SIDE
    assert printable - 0.01 <= convert.fit_width_inches(3000, 300) <= printable


def test_fit_width_caps_tall_figures_to_page_height():
    width = convert.fit_width_inches(600, 3000)
    printable_h = convert.DEFAULT_PAGE_H - convert.MARGIN_TOP - convert.MARGIN_BOTTOM - convert.CAPTION_ALLOWANCE_IN
    assert width * 3000 / 600 <= printable_h + 0.01


# --- inline fallback ---------------------------------------------------------

def test_inline_fallback_does_not_embed_files_outside_root(tmp_path):
    pytest.importorskip('bs4')
    root = tmp_path / 'site'
    (root / 'css').mkdir(parents=True)
    (tmp_path / 'secret.png').write_bytes(PNG_1PX)
    (root / 'ok.png').write_bytes(PNG_1PX)
    (root / 'bg.png').write_bytes(PNG_1PX)
    # url() inside an external stylesheet resolves relative to the stylesheet.
    (root / 'css' / 'site.css').write_text('.a{background:url(../bg.png)} .b{background:url(../../secret.png)}')
    html = root / 'index.html'
    html.write_text('<link rel="stylesheet" href="css/site.css"><img src="ok.png"><img src="../secret.png">')
    blocked: list[str] = []
    out = convert.inline_local_dependencies(html, root, blocked)
    assert out.count('data:image/png;base64,') == 2  # ok.png and bg.png only
    assert '../secret.png' in blocked and '../../secret.png' in blocked


# --- end to end --------------------------------------------------------------

def _require_toolchain():
    if not shutil.which('pandoc'):
        pytest.skip('pandoc not installed')
    pytest.importorskip('playwright')
    pytest.importorskip('bs4')
    pytest.importorskip('PIL')


def _run(html: Path, out: Path, *extra: str) -> dict:
    try:
        convert.main([str(html), str(out), *extra])
    except RuntimeError as exc:
        if 'Chromium browser unavailable' in str(exc):
            pytest.skip('Chromium not installed for Playwright')
        raise
    return json.loads((out.parent / f'{out.stem}_export_assets' / 'conversion_report.json').read_text())


def _docx_parts(path: Path) -> tuple[str, list[str]]:
    with zipfile.ZipFile(path) as z:
        return z.read('word/document.xml').decode('utf8'), [n for n in z.namelist() if n.startswith('word/media/')]


def test_sample_end_to_end(tmp_path):
    _require_toolchain()
    out = tmp_path / 'sample.docx'
    report = _run(REPO / 'examples' / 'sample.html', out)
    xml, media = _docx_parts(out)
    assert report['diagram_count'] == 1 and report['diagrams'][0]['detection'] == 'explicit'
    assert len(media) == 1
    assert '<w:tbl>' in xml and 'Responsibilities' in xml and 'w:tblBorders' in xml
    assert 'Decision Fabric' not in xml  # diagram text lives only inside the image


def test_security_hidden_content_and_bare_svg(tmp_path):
    _require_toolchain()
    site = tmp_path / 'site'
    site.mkdir()
    (tmp_path / 'secret.png').write_bytes(PNG_1PX)
    html = site / 'doc.html'
    html.write_text("""<!doctype html><html><body>
      <h1>Doc</h1><p>Intro text.</p>
      <img src="../secret.png" alt="leak">
      <img src="https://example.invalid/remote.png" alt="remote">
      <div style="display:none">HIDDEN_TEXT</div>
      <ul><li><div data-export-as-image="true" style="width:300px;height:80px;background:#eef">Marked in list</div></li></ul>
      <svg width="400" height="120"><title>Bare chart</title><rect width="400" height="120" fill="#36c"/></svg>
      <script>fetch('https://example.invalid/beacon').catch(()=>{});</script>
    </body></html>""")
    out = tmp_path / 'doc.docx'
    report = _run(html, out)
    xml, media = _docx_parts(out)
    kinds = sorted(d['detection'] for d in report['diagrams'])
    assert kinds == ['explicit', 'graphic']
    assert len(media) == 2  # the two captures; neither the secret nor the remote image
    assert 'HIDDEN_TEXT' not in xml and 'Intro text.' in xml
    sec = report['security']
    assert '../secret.png' in sec['blocked_local_references']
    assert sec['removed_remote_images'] == ['https://example.invalid/remote.png']
    assert report['hidden_elements_removed'] >= 1
    assert next(d for d in report['diagrams'] if d['detection'] == 'graphic')['alt'] == 'Bare chart'
