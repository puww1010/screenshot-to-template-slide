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
    if explicit: return explicit
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
    if keep_first_slide(u) and skill:
        subprocess.run([sys.executable, os.path.join(skill, 'scripts/clean.py'), u], check=False)
    slide = first_slide_path(u)
    if a.layout:
        lay = resolve_layout(u, a.layout); rp = slide.replace('slides/', 'slides/_rels/') + '.rels'
        write(rp, re.sub(r'slideLayouts/slideLayout\d+\.xml', 'slideLayouts/' + lay, read(rp)))
    spec = importlib.util.spec_from_file_location('page', a.page); page = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(page)
    sk.C.i = 10
    sk.inject(slide, page.build())
    font = detect_title_font(u, slide) if a.font == 'auto' else (None if a.font == 'none' else a.font)
    ea = a.ea_font or ('PingFang SC' if a.lang == 'zh' else None)
    if a.notes: set_notes(u, slide, read(a.notes))
    q = os.path.join(work, 'q'); shutil.copytree(u, q)
    unify_fonts(slide, font, ea_font=ea)
    zipdir(u, a.out)
    unify_fonts(slide.replace(u, q, 1), font, a.stress, ea_font=ea); stress = os.path.join(work, 'stress.pptx'); zipdir(q, stress)
    print(f'[build] {a.out}  font={font}')
    if skill:
        r = subprocess.run([sys.executable, os.path.join(skill, 'scripts/office/validate.py'), a.out, '--original', a.template], capture_output=True, text=True)
        print('[validate]', (r.stdout + r.stderr).strip().splitlines()[-1] if (r.stdout + r.stderr).strip() else r.returncode)
    # template diff
    t, o = zipfile.ZipFile(a.template), zipfile.ZipFile(a.out)
    tn = set(t.namelist())
    changed = [n for n in o.namelist() if n not in tn or t.read(n) != o.read(n)]
    print('[diff vs template] changed parts:', changed)
    if a.qa_dir and skill:
        os.makedirs(a.qa_dir, exist_ok=True); ensure_font_alias(font, ea)
        subprocess.run([sys.executable, os.path.join(skill, 'scripts/office/soffice.py'), '--headless', '--convert-to', 'pdf', '--outdir', work, stress],
                       capture_output=True)
        pdf = os.path.join(work, 'stress.pdf'); base = os.path.join(a.qa_dir, os.path.splitext(os.path.basename(a.out))[0] + '_stress')
        for f in glob.glob(base + '*.jpg'): os.remove(f)
        subprocess.run(['pdftoppm', '-jpeg', '-r', '150', pdf, base])
        print('[qa render]', sorted(glob.glob(base + '*.jpg')))
    r = subprocess.run([sys.executable, '-I', '-m', 'markitdown', a.out], capture_output=True, text=True)
    print('[text on slide]\n' + '\n'.join(l for l in r.stdout.splitlines() if l.strip()))


if __name__ == '__main__':
    main()
