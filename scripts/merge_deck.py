#!/usr/bin/env python3
"""merge_deck.py — merge single-slide PPTX files (all built on the SAME template) into one deck, in order.

Usage: python3 merge_deck.py --out DECK.pptx page1.pptx page2.pptx ...
- The first file is the base (its masters/layouts/theme are kept).
- Each later slide is re-pointed to the base layout with the same matchingName/cSld name
  (works even if a page was re-saved in PowerPoint).
- Speaker notes and any extra media/rels of each slide are carried over.
"""
import argparse, os, re, shutil, subprocess, sys, tempfile, zipfile

def rd(p): return open(p, encoding='utf-8').read()
def wr(p, s): open(p, 'w', encoding='utf-8').write(s)

def layout_key(x):
    m = re.search(r'matchingName="([^"]*)"', x) or re.search(r'<p:cSld name="([^"]*)"', x)
    return m.group(1) if m else None

def first_slide(u):
    pres = rd(f'{u}/ppt/presentation.xml'); rid = re.search(r'<p:sldId [^>]*r:id="([^"]+)"', pres).group(1)
    rels = rd(f'{u}/ppt/_rels/presentation.xml.rels')
    m = re.search(r'<Relationship [^>]*Id="%s"[^>]*/>' % rid, rels)
    return 'ppt/' + re.search(r'Target="([^"]+)"', m.group(0)).group(1)

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--out', required=True); ap.add_argument('files', nargs='+'); a = ap.parse_args()
    work = tempfile.mkdtemp(prefix='merge_'); base = f'{work}/base'
    zipfile.ZipFile(a.files[0]).extractall(base)
    # base layouts by key
    lay_by_key = {}
    for f in sorted(os.listdir(f'{base}/ppt/slideLayouts')):
        if f.endswith('.xml'): lay_by_key.setdefault(layout_key(rd(f'{base}/ppt/slideLayouts/{f}')), f)
    ct = rd(f'{base}/[Content_Types].xml'); pres = rd(f'{base}/ppt/presentation.xml'); prels = rd(f'{base}/ppt/_rels/presentation.xml.rels')
    def next_num(folder, prefix):
        ns = [int(m) for m in re.findall(prefix + r'(\d+)\.xml', ' '.join(os.listdir(folder)))] if os.path.isdir(folder) else []
        return max(ns + [0]) + 1
    max_id = max([int(i) for i in re.findall(r'<p:sldId id="(\d+)"', pres)] + [255])
    max_rid = max([int(i) for i in re.findall(r'Id="rId(\d+)"', prels)] + [0])
    for k, f in enumerate(a.files[1:], 2):
        src = f'{work}/s{k}'; zipfile.ZipFile(f).extractall(src)
        sp = first_slide(src); sx = rd(f'{src}/{sp}'); srels_p = f'{src}/{os.path.dirname(sp)}/_rels/{os.path.basename(sp)}.rels'
        srels = rd(srels_p) if os.path.exists(srels_p) else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        n = next_num(f'{base}/ppt/slides', 'slide'); new_slide = f'slide{n}.xml'
        def fix_rel(m):
            rel = m.group(0); tgt = re.search(r'Target="([^"]+)"', rel).group(1); typ = re.search(r'Type="([^"]+)"', rel).group(1)
            if typ.endswith('/slideLayout'):
                key = layout_key(rd(os.path.normpath(f'{src}/ppt/slides/{tgt}')))
                if key not in lay_by_key: sys.exit(f'{f}: layout "{key}" not in base template')
                return rel.replace(tgt, '../slideLayouts/' + lay_by_key[key])
            if typ.endswith('/notesSlide'):
                nn = next_num(f'{base}/ppt/notesSlides', 'notesSlide'); os.makedirs(f'{base}/ppt/notesSlides/_rels', exist_ok=True)
                srcn = os.path.normpath(f'{src}/ppt/slides/{tgt}'); shutil.copy(srcn, f'{base}/ppt/notesSlides/notesSlide{nn}.xml')
                nr = rd(os.path.join(os.path.dirname(srcn), '_rels', os.path.basename(srcn) + '.rels'))
                nr = re.sub(r'\.\./slides/slide\d+\.xml', f'../slides/{new_slide}', nr)
                wr(f'{base}/ppt/notesSlides/_rels/notesSlide{nn}.xml.rels', nr)
                nonlocal_ct.append(f'<Override PartName="/ppt/notesSlides/notesSlide{nn}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>')
                return rel.replace(tgt, f'../notesSlides/notesSlide{nn}.xml')
            if 'TargetMode="External"' in rel: return rel
            srcf = os.path.normpath(f'{src}/ppt/slides/{tgt}'); ext = os.path.splitext(srcf)[1]
            dstn = f'{os.path.splitext(os.path.basename(srcf))[0]}_m{k}{ext}'; dst_dir = os.path.normpath(f'{base}/ppt/slides/' + os.path.dirname(tgt))
            os.makedirs(dst_dir, exist_ok=True); shutil.copy(srcf, f'{dst_dir}/{dstn}')
            if f'Extension="{ext[1:]}"' not in ct and ext:
                nonlocal_ct.append(f'<Default Extension="{ext[1:]}" ContentType="image/{ext[1:]}"/>')
            return rel.replace(tgt, os.path.dirname(tgt) + '/' + dstn)
        nonlocal_ct = []
        srels = re.sub(r'<Relationship [^>]*/>', fix_rel, srels)
        os.makedirs(f'{base}/ppt/slides/_rels', exist_ok=True)
        wr(f'{base}/ppt/slides/{new_slide}', sx); wr(f'{base}/ppt/slides/_rels/{new_slide}.rels', srels)
        ct = ct.replace('</Types>', ''.join(nonlocal_ct) + f'<Override PartName="/ppt/slides/{new_slide}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
        max_id += 1; max_rid += 1
        prels = prels.replace('</Relationships>', f'<Relationship Id="rId{max_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{new_slide}"/></Relationships>')
        pres = pres.replace('</p:sldIdLst>', f'<p:sldId id="{max_id}" r:id="rId{max_rid}"/></p:sldIdLst>')
        print(f'+ {os.path.basename(f)} -> {new_slide}')
    wr(f'{base}/[Content_Types].xml', ct); wr(f'{base}/ppt/presentation.xml', pres); wr(f'{base}/ppt/_rels/presentation.xml.rels', prels)
    if os.path.exists(a.out): os.remove(a.out)
    subprocess.check_call(['zip', '-qXDr', os.path.abspath(a.out), '.'], cwd=base)
    print('[merged]', a.out, len(a.files), 'slides')

if __name__ == '__main__':
    main()
