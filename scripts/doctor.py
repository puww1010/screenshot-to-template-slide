#!/usr/bin/env python3
"""doctor.py -- check that this skill can actually run here, before building any slide.

Usage:
  python3 scripts/doctor.py              # check package files, Python and optional tools
  python3 scripts/doctor.py --selftest   # ...and build one demo slide end to end (about 10-60 s)

Exit code 0 = ready (optional tools may still be missing, see the table); 1 = cannot build, fix the [FAIL] lines first.
"""
import os, platform, shutil, subprocess, sys, tempfile, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# Every file SKILL.md tells the agent to run or read. If one is missing, the package is incomplete
# (for example only SKILL.md was copied) and the workflow will break halfway.
MANIFEST = [
    'SKILL.md',
    'scripts/slidekit.py', 'scripts/build_slide.py', 'scripts/list_layouts.py', 'scripts/merge_deck.py', 'scripts/doctor.py',
    'examples/page_template.py', 'examples/stat_timeline.py', 'examples/how_it_works_tabs.py',
    'examples/illustrated_flow.py', 'examples/section_divider.py', 'examples/illustrated_flow_io.md',
    'examples/demo-template.pptx',
    'references/page-patterns.md', 'references/lessons.md', 'references/prompt-patterns.md',
    'references/template-preset-example.md', 'references/testing.md', 'references/troubleshooting.md',
]

OS = platform.system()
HINTS = {
    'soffice': {'Darwin': 'brew install --cask libreoffice', 'Linux': 'sudo apt install libreoffice-impress',
                'Windows': 'winget install TheDocumentFoundation.LibreOffice'},
    'pdftoppm': {'Darwin': 'brew install poppler', 'Linux': 'sudo apt install poppler-utils',
                 'Windows': 'winget install oschwartz10612.Poppler'},
    'markitdown': {'*': 'pip install markitdown'},
    'pptx': {'*': 'pip install python-pptx'},
}


def hint(key):
    h = HINTS.get(key, {})
    return h.get(OS) or h.get('*') or ''


def has_module(name):
    return importlib.util.find_spec(name) is not None


def find_pptx_skill():
    sys.path.insert(0, HERE)
    try:
        import build_slide
        return build_slide.find_pptx_skill(None)
    except Exception:
        return None


def main():
    selftest = '--selftest' in sys.argv[1:]
    fails, rows = [], []

    # 1. Python
    ok = sys.version_info >= (3, 8)
    rows.append(('Python >= 3.8', 'OK' if ok else 'MISSING', platform.python_version(), 'required'))
    if not ok: fails.append(f'Python {platform.python_version()} is too old; install Python 3.8 or newer')

    # 2. Package completeness
    missing = [f for f in MANIFEST if not os.path.isfile(os.path.join(ROOT, f))]
    rows.append(('package files', 'OK' if not missing else 'MISSING', f'{len(MANIFEST) - len(missing)}/{len(MANIFEST)}', 'required'))
    for f in missing: print(f'[FAIL] package file missing: {f}')
    if missing:
        fails.append('the skill package is incomplete. Install the full .skill package (Releases page of '
                     'github.com/puww1010/screenshot-to-template-slide), not SKILL.md alone')

    # 3. Optional tools: what is lost without each one
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    rows.append(('LibreOffice', 'OK' if soffice else 'missing', soffice or hint('soffice'),
                 'optional: without it no QA render; lint + text check still run, inspect the page in PowerPoint'))
    pdftoppm = shutil.which('pdftoppm')
    rows.append(('pdftoppm', 'OK' if pdftoppm else 'missing', pdftoppm or hint('pdftoppm'),
                 'optional: without it the QA render stays a PDF instead of JPG'))
    md = has_module('markitdown')
    rows.append(('markitdown', 'OK' if md else 'missing', '' if md else hint('markitdown'),
                 'optional: built-in text extraction is used instead'))
    skill = find_pptx_skill()
    rows.append(('Anthropic pptx skill', 'OK' if skill else 'missing', skill or '',
                 'optional: built-in validate / render fallbacks are used instead'))
    pp = has_module('pptx')
    rows.append(('python-pptx', 'OK' if pp else 'missing', '' if pp else hint('pptx'),
                 'optional: only needed for tests/test_regressions.py'))

    w = max(len(r[0]) for r in rows)
    print('\n[doctor] ' + f'{OS} / Python {platform.python_version()} / skill at {ROOT}')
    for name, state, detail, note in rows:
        print(f'  {name:<{w}}  {state:<8} {detail}')
        if state != 'OK': print(f'  {"":<{w}}           -> {note}')

    # 4. Self-test: one real build on the bundled demo template
    if selftest and not fails:
        out = tempfile.mkdtemp(prefix='doctor_')
        cmd = [sys.executable, os.path.join(HERE, 'build_slide.py'), '--template', os.path.join(ROOT, 'examples/demo-template.pptx'),
               '--page', os.path.join(ROOT, 'examples/stat_timeline.py'), '--out', os.path.join(out, 'selftest.pptx')]
        if soffice: cmd += ['--qa-dir', os.path.join(out, 'qa')]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=400)
        log = r.stdout + r.stderr
        if r.returncode == 0 and '[lint] OK' in log:
            print(f'[selftest] OK  demo slide built: {os.path.join(out, "selftest.pptx")}')
        else:
            fails.append('selftest build failed:\n    ' + '\n    '.join(log.strip().splitlines()[-8:]))

    if fails:
        for f in fails: print(f'[FAIL] {f}')
        print('[doctor] NOT READY. Fix the [FAIL] lines above (see references/troubleshooting.md).')
        return 1
    print('[doctor] READY' + ('' if selftest else '  (run with --selftest to build one demo slide)'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
