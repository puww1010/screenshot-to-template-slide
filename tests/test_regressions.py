#!/usr/bin/env python3
"""Regression tests for problems found in real use. Run:  python3 tests/test_regressions.py
Needs python-pptx (only to generate throw-away templates) and the same tools build_slide.py needs."""
import os, re, shutil, subprocess, sys, tempfile, zipfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, 'scripts'); sys.path.insert(0, SCRIPTS)
import slidekit as sk  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402

TMP = tempfile.mkdtemp(prefix='sts_test_')
FAILS = []


def check(name, cond, detail=''):
    print(('PASS ' if cond else 'FAIL ') + name + ('' if cond else f'  -> {detail}'))
    if not cond: FAILS.append(name)


def template(w, h, name):
    p = Presentation(); p.slide_width = Inches(w); p.slide_height = Inches(h)
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = 'x'
    path = os.path.join(TMP, name); p.save(path); return path


def page(body, name):
    path = os.path.join(TMP, name)
    open(path, 'w', encoding='utf-8').write(f'import sys; sys.path.insert(0, {SCRIPTS!r})\nfrom slidekit import *\n' + body)
    return path


def build(tpl, pg, *extra):
    out = os.path.join(TMP, os.path.basename(pg) + '.pptx')
    r = subprocess.run([sys.executable, os.path.join(SCRIPTS, 'build_slide.py'), '--template', tpl, '--page', pg,
                        '--out', out, '--pptx-skill', 'none', *extra], capture_output=True, text=True)
    return r.returncode, r.stdout + r.stderr, out


def slide_xml(out):
    return zipfile.ZipFile(out).read('ppt/slides/slide1.xml').decode('utf-8')


T10 = template(10, 5.625, 't10.pptx')
T20 = template(20, 11.25, 't20.pptx')

# 1. para(lnpct): percent in, OOXML units (1/1000 %) out; OOXML-unit input still accepted; nonsense rejected.
#    Bug: lnpct=120 was written as spcPct val="120" (= 0.12 %), PowerPoint collapsed every line onto one.
check('para lnpct=120 -> 120000', 'spcPct val="120000"' in sk.para([sk.run('a', 12)], lnpct=120))
check('para lnpct=110000 kept', 'spcPct val="110000"' in sk.para([sk.run('a', 12)], lnpct=110000))
for bad in (5, 400, 900000):
    try:
        sk.para([sk.run('a', 12)], lnpct=bad); check(f'para lnpct={bad} rejected', False, 'no error raised')
    except ValueError:
        check(f'para lnpct={bad} rejected', True)

# 2. XML lint: spacing/size values PowerPoint takes literally (LibreOffice silently fixes them) must fail the build.
raw = page("def build():\n    return [text('T',.5,1,9,1,['<a:p><a:pPr><a:lnSpc><a:spcPct val=\"120\"/></a:lnSpc></a:pPr>'+run('one two',12)+'</a:p>'])]\n", 'bad_lnspc.py')
code, log, _ = build(T10, raw)
check('lint fails on spcPct val=120', code != 0 and '[FAIL] lint' in log, log[-300:])
raw = page("def build():\n    return [text('T',.5,1,9,1,[para([run('big',12).replace('sz=\"1200\"','sz=\"180000\"')])])]\n", 'bad_sz.py')
code, log, _ = build(T10, raw)
check('lint fails on sz=180000 (1800 pt)', code != 0 and '[FAIL] lint' in log, log[-300:])
ok = page("def build():\n    return [title_ph('Hello'), text('T',.5,1,9,1,[para([run('fine',12)],lnpct=115)])]\n", 'ok.py')
code, log, _ = build(T10, ok)
check('lint passes a normal page', code == 0 and '[lint] OK' in log, log[-300:])

# 3. --lang zh: CJK runs (including the title placeholder) are tagged zh-CN, Latin-only runs keep en-US.
zh = page("def build():\n    return [title_ph('标题 Title'), text('T',.5,1,9,1,[para([run('中文内容',12), run('English',12)])])]\n", 'zh.py')
code, log, out = build(T10, zh, '--lang', 'zh')
x = slide_xml(out) if code == 0 else ''
cjk_runs = re.findall(r'<a:r><a:rPr\b([^>]*)>(?:(?!</a:r>).)*?<a:t[^>]*>([^<]*[一-鿿][^<]*)</a:t>', x, re.S)
check('--lang zh tags every CJK run zh-CN', code == 0 and cjk_runs and all('lang="zh-CN"' in a for a, _ in cjk_runs), log[-300:])
check('--lang zh leaves Latin-only runs en-US', 'lang="en-US"' in x)
code, log, _ = build(T10, zh)
check('CJK text without --lang zh -> WARN', code == 0 and '[WARN]' in log and 'CJK' in log, log[-300:])

# 4. Canvas other than 10" x 5.625": slidekit coordinates/font sizes need scaling -> WARN with the factor.
code, log, _ = build(T20, ok)
check('20" canvas -> WARN with scale x2.0', code == 0 and re.search(r'\[WARN\] canvas 20(\.0+)? x 11\.25', log) and 'x2.0' in log, log[-400:])
# attribute order cy before cx (Google Slides exports write it this way) must be detected too
import shutil
T20b = os.path.join(TMP, 't20_cy_first.pptx'); d = os.path.join(TMP, 'unz'); shutil.rmtree(d, True)
zipfile.ZipFile(T20).extractall(d); pp = os.path.join(d, 'ppt/presentation.xml'); px = open(pp, encoding='utf-8').read()
px = re.sub(r'<p:sldSz cx="(\d+)" cy="(\d+)"', r'<p:sldSz cy="\2" cx="\1"', px); open(pp, 'w', encoding='utf-8').write(px)
subprocess.run(['zip', '-qXDr', T20b, '.'], cwd=d, check=True)
code, log, _ = build(T20b, ok)
check('20" canvas with cy before cx -> WARN', code == 0 and '[WARN] canvas 20' in log, log[-400:])
code, log, _ = build(T10, ok)
check('10" canvas -> no canvas WARN', '[WARN] canvas' not in log, log[-300:])

# 5. Fonts that are not standard on Windows/macOS -> NOTE to install/embed on the presenting machine.
code, log, _ = build(T10, ok, '--font', 'Figtree')
check('non-system font -> NOTE', code == 0 and '[NOTE]' in log and 'Figtree' in log, log[-300:])
code, log, _ = build(T10, ok, '--font', 'Arial')
check('Arial -> no font NOTE', '[NOTE] font' not in log, log[-300:])

# 6. Section dividers are text only: the shipped example has no illustration shapes, and a page that puts an
#    illustration card on a section/divider layout gets a WARN (other layouts don't).
DEMO = os.path.join(ROOT, 'examples', 'demo-template.pptx')
code, log, out = build(DEMO, os.path.join(ROOT, 'examples', 'section_divider.py'), '--layout', 'ACME Section Divider')
names = re.findall(r'<p:cNvPr [^>]*name="([^"]*)"', slide_xml(out)) if code == 0 else []
check('section_divider example is text only', code == 0 and names and
      all(n in ('Title', 'Subtitle', 'Progress label') or n.startswith('Progress') for n in names[1:]), str(names))
check('text-only section divider -> no WARN', code == 0 and '[WARN] section divider' not in log, log[-300:])
card = page("def build():\n    return [title_ph('Chapter'), stage_card(), icon('bug',5,1.2,.4)]\n", 'card_divider.py')
code, log, _ = build(T10, card, '--layout', 'Section Header')
check('illustration on section layout -> WARN', code == 0 and '[WARN] section divider' in log, log[-300:])
code, log, _ = build(T10, card)
check('illustration on content layout -> no section WARN', code == 0 and '[WARN] section divider' not in log, log[-300:])

# 7. Reliability: the build must not depend on the `zip` / `fc-list` binaries, and user mistakes (a page.py without
#    build(), a template that is not a .pptx) must end with a readable [FAIL] line instead of a Python traceback.
NOBIN = os.path.join(TMP, 'nobin'); os.makedirs(NOBIN, exist_ok=True)
for tool in ('python3', 'soffice', 'libreoffice', 'pdftoppm'):
    w = shutil.which(tool)
    if w and not os.path.exists(os.path.join(NOBIN, tool)): os.symlink(w, os.path.join(NOBIN, tool))
env = dict(os.environ, PATH=NOBIN)
r = subprocess.run([sys.executable, os.path.join(SCRIPTS, 'build_slide.py'), '--template', T10, '--page', ok, '--out',
                    os.path.join(TMP, 'nozip.pptx'), '--pptx-skill', 'none', '--qa-dir', os.path.join(TMP, 'qa_nozip')],
                   capture_output=True, text=True, env=env)
check('build works without zip / fc-list on PATH', r.returncode == 0 and '[lint] OK' in r.stdout, (r.stdout + r.stderr)[-400:])
if r.returncode == 0:
    zn = zipfile.ZipFile(os.path.join(TMP, 'nozip.pptx')).namelist()
    check('pure-Python zip: [Content_Types].xml first, no directory entries', zn[0] == '[Content_Types].xml' and not any(n.endswith('/') for n in zn), str(zn[:3]))
nobuild = page("x = 1\n", 'nobuild.py')
code, log, _ = build(T10, nobuild)
check('page.py without build() -> [FAIL], no traceback', code != 0 and '[FAIL]' in log and 'Traceback' not in log, log[-300:])
bad_tpl = os.path.join(TMP, 'not_a_template.pptx'); open(bad_tpl, 'w').write('hello')
code, log, _ = build(bad_tpl, ok)
check('invalid template -> [FAIL], no traceback', code != 0 and '[FAIL]' in log and 'Traceback' not in log, log[-300:])
code, log, _ = build(T10, ok, '--layout', 'No Such Layout')
check('unknown --layout -> [FAIL] listing layouts', code != 0 and '[FAIL]' in log and 'Title Only' in log, log[-300:])

# 8. doctor.py: reports a missing package file, and --selftest builds a demo slide end to end.
DOC = os.path.join(SCRIPTS, 'doctor.py')
r = subprocess.run([sys.executable, DOC, '--selftest'], capture_output=True, text=True)
check('doctor --selftest passes on a complete package', r.returncode == 0 and '[selftest] OK' in r.stdout, (r.stdout + r.stderr)[-400:])
cp = os.path.join(TMP, 'pkg'); shutil.rmtree(cp, True); shutil.copytree(ROOT, cp, ignore=shutil.ignore_patterns('.git', '__pycache__'))
os.remove(os.path.join(cp, 'references', 'lessons.md'))
r = subprocess.run([sys.executable, os.path.join(cp, 'scripts', 'doctor.py')], capture_output=True, text=True)
check('doctor flags a missing package file', r.returncode != 0 and 'references/lessons.md' in r.stdout, (r.stdout + r.stderr)[-400:])

print(f'\n{len(FAILS)} failed' if FAILS else '\nall passed')
sys.exit(1 if FAILS else 0)
