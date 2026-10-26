#!/usr/bin/env python3
"""build_slide.py — build ONE editable slide on top of a .pptx template, then QA it.

Pipeline: unpack template -> keep only the first slide -> (optional) switch its layout ->
          inject shapes returned by page.build() -> unify fonts -> zip -> validate ->
          stress render (+6% font size) -> diff against template.

Usage:
  python3 build_slide.py --template T.pptx --page my_page.py --out OUT.pptx \
      [--layout "Section Divider" | slideLayout19.xml] [--font auto|"Helvetica Neue"|none] \
      [--qa-dir DIR] [--pptx-skill /path/to/skills/pptx] [--notes script.txt] [--lang en|zh] [--ea-font "PingFang SC"]

page.py must define  build()  returning a list of shape-XML strings (use slidekit helpers).
"""
import argparse, glob, os, re, shutil, subprocess, sys, tempfile, zipfile, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import slidekit as sk  # noqa: E402


def find_pptx_skill(explicit):
    """Optional helper: Anthropic's official `pptx` skill (validate.py / soffice.py / clean.py).
    Pass --pptx-skill none to force the built-in fallbacks."""
    if explicit: return None if explicit.lower() == 'none' else explicit
    if os.environ.get('PPTX_SKILL'): return os.environ['PPTX_SKILL']
    for pat in ['/sessions/*/mnt/.claude/skills/pptx', os.path.expanduser('~/.claude/skills/pptx'),
                os.path.join(HERE, '..', '..', 'pptx')]:
        hits = glob.glob(pat)
        if hits and os.path.exists(os.path.join(hits[0], 'scripts/office/validate.py')): return hits[0]
    return None


def read(p): return open(p, encoding='utf-8').read()
def write(p, s): open(p, 'w', encoding='utf-8').write(s)


def keep_first_slide(u):
    pres = os.path.join(u, 'ppt/presentation.xml'); x = read(pres)
    ids = re.findall(r'<p:sldId [^>]*/>', x)
    if len(ids) > 1:
        x = x.replace(''.join(ids), ids[0]) if ''.join(ids) in x else re.sub(r'(<p:sldIdLst>).*?(</p:sldIdLst>)', lambda m: m.group(1) + ids[0] + m.group(2), x, flags=re.S)
        write(pres, x); return True
    return False


def prune_unlisted_slides(u):
    """Built-in clean: delete slide parts (and their notes) that are no longer in <p:sldIdLst>."""
    pres = read(os.path.join(u, 'ppt/presentation.xml')); prp = os.path.join(u, 'ppt/_rels/presentation.xml.rels'); prels = read(prp)
    keep = set(re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres)); ctp = os.path.join(u, '[Content_Types].xml'); ct = read(ctp)
    for rel in re.findall(r'<Relationship [^>]*/>', prels):
        if '/relationships/slide"' not in rel: continue
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        if rid in keep: continue
        tgt = re.search(r'Target="([^"]+)"', rel).group(1); sp = os.path.normpath(os.path.join(u, 'ppt', tgt))
        srel = os.path.join(os.path.dirname(sp), '_rels', os.path.basename(sp) + '.rels')
        parts = [sp]
        if os.path.exists(srel):
            for n in re.findall(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"', read(srel)):
                np_ = os.path.join(u, 'ppt/notesSlides', n); parts += [np_, os.path.join(u, 'ppt/notesSlides/_rels', n + '.rels')]
            parts.append(srel)
        for f in parts:
            if os.path.exists(f):
                os.remove(f); pn = '/' + os.path.relpath(f, u).replace(os.sep, '/')
                ct = re.sub(r'<Override PartName="%s"[^>]*/>' % re.escape(pn), '', ct)
        prels = prels.replace(rel, '')
    write(prp, prels); write(ctp, ct)


def basic_validate(path):
    """Built-in package check used when the pptx skill is absent: XML well-formed, rels targets exist,
    every part has a content type, shape ids unique per slide."""
    import xml.dom.minidom as md
    z = zipfile.ZipFile(path); names = set(z.namelist()); issues = []
    ct = z.read('[Content_Types].xml').decode('utf8')
    defaults = set(re.findall(r'<Default Extension="([^"]+)"', ct)); overrides = set(re.findall(r'<Override PartName="([^"]+)"', ct))
    for n in names:
        if n.endswith('/'): continue
        if n.endswith('.xml') or n.endswith('.rels'):
            try: md.parseString(z.read(n))
            except Exception as e: issues.append(f'XML parse error in {n}: {e}')
        if n != '[Content_Types].xml' and '/' + n not in overrides and n.rsplit('.', 1)[-1] not in defaults:
            issues.append(f'no content type for {n}')
        if n.endswith('.rels'):
            base = os.path.dirname(os.path.dirname(n))
            for rel in re.findall(r'<Relationship [^>]*/>', z.read(n).decode('utf8')):
                if 'TargetMode="External"' in rel: continue
                t = re.search(r'Target="([^"]+)"', rel).group(1)
                full = t.lstrip('/') if t.startswith('/') else os.path.normpath(os.path.join(base, t)).replace(os.sep, '/')
                if full not in names: issues.append(f'{n}: missing target {t}')
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            ids = re.findall(r'<p:cNvPr id="(\d+)"', z.read(n).decode('utf8'))
            d = {i for i in ids if ids.count(i) > 1}
            if d: issues.append(f'{n}: duplicate shape ids {sorted(d)}')
    return issues


def render_pdf(pptx, outdir, skill):
    if skill:
        subprocess.run([sys.executable, os.path.join(skill, 'scripts/office/soffice.py'), '--headless', '--convert-to', 'pdf', '--outdir', outdir, pptx], capture_output=True)
    else:
        exe = shutil.which('soffice') or shutil.which('libreoffice')
        if not exe: print('[qa render] LibreOffice not found — skipped'); return None
        prof = tempfile.mkdtemp(prefix='lo_')
        try:
            subprocess.run([exe, f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf', '--outdir', outdir, pptx],
                           capture_output=True, timeout=240)
        except subprocess.TimeoutExpired:
            print('[qa render] LibreOffice timed out — skipped'); return None
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + '.pdf')
    return pdf if os.path.exists(pdf) else None


def slide_text(pptx):
    r = subprocess.run([sys.executable, '-I', '-m', 'markitdown', pptx], capture_output=True, text=True)
    if r.returncode == 0 and r.stdout.strip(): return [l for l in r.stdout.splitlines() if l.strip()]
    x = zipfile.ZipFile(pptx).read(first_slide_path_zip(pptx)).decode('utf8')
    return [''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', p)) for p in re.findall(r'<a:p>.*?</a:p>', x, re.S) if '<a:t' in p]


def first_slide_path_zip(pptx):
    z = zipfile.ZipFile(pptx); pres = z.read('ppt/presentation.xml').decode('utf8')
    rid = re.search(r'<p:sldId [^>]*r:id="([^"]+)"', pres).group(1)
    rel = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, z.read('ppt/_rels/presentation.xml.rels').decode('utf8')).group(0)
    return 'ppt/' + re.search(r'Target="([^"]+)"', rel).group(1)


def first_slide_path(u):
    pres = read(os.path.join(u, 'ppt/presentation.xml'))
    rid = re.search(r'<p:sldId [^>]*r:id="([^"]+)"', pres).group(1)
    rels = read(os.path.join(u, 'ppt/_rels/presentation.xml.rels'))
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return os.path.join(u, 'ppt', tgt.group(1))


def resolve_layout(u, want):
    if want.endswith('.xml'): return want
    for f in sorted(glob.glob(os.path.join(u, 'ppt/slideLayouts/slideLayout*.xml'))):
        x = read(f)
        mn = re.search(r'matchingName="([^"]*)"', x); cn = re.search(r'<p:cSld name="([^"]*)"', x)
        if (mn and mn.group(1) == want) or (cn and cn.group(1) == want): return os.path.basename(f)
    sys.exit(f'layout "{want}" not found; run list_layouts.py on the template')


def detect_title_font(u, slide_path):
    rels = read(slide_path.replace('slides/', 'slides/_rels/') + '.rels')
    lay = re.search(r'slideLayouts/(slideLayout\d+\.xml)', rels).group(1)
    lx = read(os.path.join(u, 'ppt/slideLayouts', lay))
    lrels = read(os.path.join(u, 'ppt/slideLayouts/_rels', lay + '.rels'))
    master = re.search(r'slideMasters/(slideMaster\d+\.xml)', lrels).group(1)
    mx = read(os.path.join(u, 'ppt/slideMasters', master))
    for src in (lx, mx):
        m = re.search(r'<p:ph type="(?:title|ctrTitle)".*?</p:sp>', src, re.S)
        if m:
            f = re.search(r'<a:latin typeface="([^"+][^"]*)"', m.group(0))
            if f: return f.group(1)
    m = re.search(r'<p:titleStyle>.*?<a:latin typeface="([^"]+)"', mx, re.S)
    if m and not m.group(1).startswith('+'): return m.group(1)
    theme = sorted(glob.glob(os.path.join(u, 'ppt/theme/theme*.xml')))[0]
    m = re.search(r'<a:majorFont><a:latin typeface="([^"]+)"', read(theme))
    return m.group(1) if m else None


def unify_fonts(slide_path, font, scale=1.0, ea_font=None):
    """Add latin/ea/cs/sym typeface to every run we wrote (runs with content-bearing rPr).
    Runs inheriting from placeholders (self-closing rPr) are left alone. scale>1 = stress copy.
    ea_font: separate East-Asian (CJK) font for Chinese pages."""
    x = read(slide_path)
    tag = (f'<a:latin typeface="{font}"/><a:ea typeface="{ea_font or font}"/><a:cs typeface="{font}"/><a:sym typeface="{font}"/>') if font else ''
    def fix(m):
        s = m.group(0)
        return s if (not tag or 'typeface=' in s) else s.replace('</a:rPr>', tag + '</a:rPr>')
    x = re.sub(r'<a:rPr\b[^>]*[^/]>.*?</a:rPr>', fix, x, flags=re.S)
    if ea_font:  # placeholder runs (self-closing rPr) keep template latin font but get the CJK font
        x = re.sub(r'<a:rPr\b([^>]*)/>', lambda m: f'<a:rPr{m.group(1)}><a:ea typeface="{ea_font}"/></a:rPr>', x)
    if scale != 1.0:
        x = re.sub(r'(<a:rPr\b[^>]*\bsz=")(\d+)"', lambda m: f'{m.group(1)}{int(int(m.group(2)) * scale)}"', x)
    write(slide_path, x)


CJK = re.compile('[%s]' % ''.join('%s-%s' % (chr(a), chr(b)) for a, b in ((0x3000, 0x303f), (0x3400, 0x4dbf), (0x4e00, 0x9fff), (0xff00, 0xffef))))
RUN_RE = re.compile(r'(<a:r>\s*<a:rPr\b)([^>]*?)(/?>)((?:(?!</a:r>).)*?<a:t[^>]*>)([^<]*)(</a:t>)', re.S)


def tag_cjk_lang(slide_path, lang='zh-CN'):
    """--lang zh: tag every run whose text contains CJK with lang=zh-CN (slidekit writes en-US).
    PowerPoint picks fallback fonts and line-breaking rules from this attribute."""
    x = read(slide_path)
    def fix(m):
        if not CJK.search(m.group(5)): return m.group(0)
        attrs = re.sub(r'\slang="[^"]*"', '', m.group(2))
        return m.group(1) + f' lang="{lang}"' + attrs + ''.join(m.group(k) for k in range(3, 7))
    write(slide_path, RUN_RE.sub(fix, x))


def lint_slide(x, lang):
    """Static checks for values PowerPoint takes literally but LibreOffice silently corrects,
    so the QA render cannot catch them. Returns (fails, warns)."""
    fails, warns = [], []
    for v in re.findall(r'<a:lnSpc><a:spcPct val="(-?\d+)"', x):
        if not 50000 <= int(v) <= 300000:
            fails.append(f'line spacing spcPct val="{v}" = {int(v)/1000:g} % (expected 50-300 %, i.e. 50000-300000; '
                         f'pass a percent such as lnpct=120 to slidekit.para)')
    for v in re.findall(r'<a:lnSpc><a:spcPts val="(-?\d+)"', x):
        if not 100 <= int(v) <= 100000: fails.append(f'line spacing spcPts val="{v}" = {int(v)/100:g} pt')
    for tag, v in re.findall(r'<a:(spcBef|spcAft)><a:spcPts val="(-?\d+)"', x):
        if not 0 <= int(v) <= 20000: fails.append(f'{tag} spcPts val="{v}" = {int(v)/100:g} pt (expected 0-200 pt)')
    for v in re.findall(r'<a:(?:rPr|defRPr|endParaRPr)\b[^>]*\bsz="(-?\d+)"', x):
        if not 100 <= int(v) <= 40000: fails.append(f'font size sz="{v}" = {int(v)/100:g} pt (expected 1-400 pt)')
    if lang != 'zh':
        n = sum(1 for m in RUN_RE.finditer(x) if CJK.search(m.group(5)) and 'lang="zh' not in m.group(2))
        if n: warns.append(f'{n} run(s) contain CJK text but are not tagged zh-CN: build Chinese pages with --lang zh')
    return sorted(set(fails)), warns


def canvas_note(u):
    tag = re.search(r'<p:sldSz\b[^>]*>', read(os.path.join(u, 'ppt/presentation.xml')))
    cx = tag and re.search(r'\bcx="(\d+)"', tag.group(0)); cy = tag and re.search(r'\bcy="(\d+)"', tag.group(0))
    if not (cx and cy): return
    w, h = int(cx.group(1)) / 914400, int(cy.group(1)) / 914400
    if abs(w - 10) > .01 or abs(h - 5.625) > .01:
        print(f'[WARN] canvas {w:g} x {h:g} in: slidekit defaults and references/page-patterns.md assume 10 x 5.625 in; '
              f'scale coordinates and font sizes by x{w/10:.1f}')


CROSS_FONTS = {'Arial', 'Arial Black', 'Aptos', 'Calibri', 'Cambria', 'Candara', 'Consolas', 'Constantia', 'Corbel',
               'Courier New', 'Georgia', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'}
MAC_FONTS = {'Helvetica', 'Helvetica Neue', 'PingFang SC', 'PingFang TC', 'Hiragino Sans GB', 'Songti SC', 'Heiti SC'}
WIN_FONTS = {'Segoe UI', 'Microsoft YaHei', 'Microsoft JhengHei', 'SimSun', 'SimHei', 'DengXian', 'Yu Gothic'}


def font_notes(*fonts):
    for f in dict.fromkeys(x for x in fonts if x):
        if f in CROSS_FONTS: continue
        if f in MAC_FONTS: print(f'[NOTE] font "{f}" ships with macOS only; on Windows PowerPoint substitutes it')
        elif f in WIN_FONTS: print(f'[NOTE] font "{f}" ships with Windows only; on macOS PowerPoint substitutes it')
        else: print(f'[NOTE] font "{f}" is not a standard Windows/macOS font: install it on the presenting machine, '
                    f'or embed fonts when saving in PowerPoint')


def zipdir(d, out):
    if os.path.exists(out): os.remove(out)
    subprocess.check_call(['zip', '-qXDr', os.path.abspath(out), '.'], cwd=d)


def ensure_font_alias(font, ea=None):
    """For QA rendering only: map missing fonts to metric-similar ones so overflow checks are meaningful."""
    have = subprocess.run(['fc-list', ':family'], capture_output=True, text=True).stdout
    al = ''
    if font and font not in have: al += f'<alias binding="same"><family>{font}</family><prefer><family>Liberation Sans</family></prefer></alias>'
    if ea and ea not in have: al += f'<alias binding="same"><family>{ea}</family><prefer><family>Noto Sans CJK SC</family></prefer></alias>'
    if not al: return
    cfg = os.path.expanduser('~/.config/fontconfig'); os.makedirs(cfg, exist_ok=True)
    write(os.path.join(cfg, 'fonts.conf'), f'<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>{al}</fontconfig>')


def set_notes(u, slide_path, notes_text):
    """Put the script into the slide's speaker notes (body placeholder of its notesSlide)."""
    rp = slide_path.replace('slides/', 'slides/_rels/') + '.rels'; rels = read(rp)
    m = re.search(r'Target="\.\./notesSlides/(notesSlide\d+\.xml)"', rels)
    if not m:
        print('[notes] slide has no notesSlide part — skipped (add notes in PowerPoint, or use a template slide that has notes)'); return
    npth = os.path.join(u, 'ppt/notesSlides', m.group(1)); x = read(npth)
    paras = ''.join(f'<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t xml:space="preserve">{sk.esc(t)}</a:t></a:r></a:p>' if t.strip() else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
                    for t in notes_text.strip().splitlines())
    def fill(mm):
        sp = mm.group(0)
        return re.sub(r'(<p:txBody>.*?<a:lstStyle/>|<p:txBody>.*?<a:lstStyle>.*?</a:lstStyle>).*?(</p:txBody>)', lambda t: t.group(1) + paras + t.group(2), sp, count=1, flags=re.S)
    x2 = re.sub(r'<p:sp>(?:(?!</p:sp>).)*?<p:ph type="body"[^>]*/>.*?</p:sp>', fill, x, count=1, flags=re.S)
    write(npth, x2); print('[notes] written to', m.group(1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--template', required=True); ap.add_argument('--page', required=True); ap.add_argument('--out', required=True)
    ap.add_argument('--layout'); ap.add_argument('--font', default='auto'); ap.add_argument('--qa-dir')
    ap.add_argument('--pptx-skill'); ap.add_argument('--stress', type=float, default=1.06)
    ap.add_argument('--notes', help='text file whose content goes into the speaker notes (script stays OFF the slide)')
    ap.add_argument('--lang', default='en'); ap.add_argument('--ea-font', default=None)
    a = ap.parse_args()
    work = tempfile.mkdtemp(prefix='sts_'); u = os.path.join(work, 'u')
    zipfile.ZipFile(a.template).extractall(u)
    skill = find_pptx_skill(a.pptx_skill)
    if keep_first_slide(u):
        prune_unlisted_slides(u)
        if skill: subprocess.run([sys.executable, os.path.join(skill, 'scripts/clean.py'), u], check=False, capture_output=True)
    slide = first_slide_path(u)
    if a.layout:
        lay = resolve_layout(u, a.layout); rp = slide.replace('slides/', 'slides/_rels/') + '.rels'
        write(rp, re.sub(r'slideLayouts/slideLayout\d+\.xml', 'slideLayouts/' + lay, read(rp)))
    spec = importlib.util.spec_from_file_location('page', a.page); page = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(page)
    sk.C.i = 10
    sk.inject(slide, page.build())
    canvas_note(u)
    if a.lang == 'zh': tag_cjk_lang(slide)
    font = detect_title_font(u, slide) if a.font == 'auto' else (None if a.font == 'none' else a.font)
    ea = a.ea_font or ('PingFang SC' if a.lang == 'zh' else None)
    if a.notes: set_notes(u, slide, read(a.notes))
    q = os.path.join(work, 'q'); shutil.copytree(u, q)
    unify_fonts(slide, font, ea_font=ea)
    zipdir(u, a.out)
    unify_fonts(slide.replace(u, q, 1), font, a.stress, ea_font=ea); stress = os.path.join(work, 'stress.pptx'); zipdir(q, stress)
    print(f'[build] {a.out}  font={font}  pptx-skill={"yes" if skill else "no (built-in fallbacks)"}')
    font_notes(font, ea)
    if skill:
        r = subprocess.run([sys.executable, os.path.join(skill, 'scripts/office/validate.py'), a.out, '--original', a.template], capture_output=True, text=True)
        print('[validate]', (r.stdout + r.stderr).strip().splitlines()[-1] if (r.stdout + r.stderr).strip() else r.returncode)
    else:
        iss = basic_validate(a.out)
        print('[validate:basic]', 'PASSED' if not iss else 'ISSUES:\n  ' + '\n  '.join(iss))
    # template diff
    t, o = zipfile.ZipFile(a.template), zipfile.ZipFile(a.out)
    tn = set(t.namelist())
    changed = [n for n in o.namelist() if n not in tn or t.read(n) != o.read(n)]
    print('[diff vs template] changed parts:', changed)
    slide_part = os.path.relpath(slide, u).replace(os.sep, '/')
    if slide_part not in changed:
        sys.exit(f'[FAIL] {slide_part} is identical to the template — page shapes were not injected')
    fails, warns = lint_slide(read(slide), a.lang)
    for w in warns: print('[WARN]', w)
    if fails:
        sys.exit('[FAIL] lint: PowerPoint would render these literally (LibreOffice hides them):\n  ' + '\n  '.join(fails))
    print('[lint] OK')
    if a.qa_dir:
        os.makedirs(a.qa_dir, exist_ok=True); ensure_font_alias(font, ea)
        pdf = render_pdf(stress, work, skill)
        if pdf and shutil.which('pdftoppm'):
            base = os.path.join(a.qa_dir, os.path.splitext(os.path.basename(a.out))[0] + '_stress')
            for f in glob.glob(base + '*.jpg'): os.remove(f)
            subprocess.run(['pdftoppm', '-jpeg', '-r', '150', pdf, base])
            print('[qa render]', sorted(glob.glob(base + '*.jpg')))
        elif pdf: print('[qa render] PDF at', pdf, '(pdftoppm not found, open the PDF to inspect)')
    print('[text on slide]\n' + '\n'.join(slide_text(a.out)))


if __name__ == '__main__':
    main()
