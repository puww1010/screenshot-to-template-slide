#!/usr/bin/env python3
"""List every slide layout in a template: file, matchingName, cSld name and placeholders (with position in inches).
Usage: python3 list_layouts.py template.pptx [--grep Section]"""
import re, sys, zipfile
z=zipfile.ZipFile(sys.argv[1]); g=sys.argv[3] if len(sys.argv)>3 and sys.argv[2]=='--grep' else None
for n in sorted([n for n in z.namelist() if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$',n)],key=lambda s:int(re.findall(r'\d+',s)[-1])):
    x=z.read(n).decode('utf8'); mn=(re.search(r'matchingName="([^"]*)"',x) or [None,''])[1]; cn=(re.search(r'<p:cSld name="([^"]*)"',x) or [None,''])[1]
    if g and g.lower() not in (mn+cn).lower(): continue
    phs=[]
    for sp in re.findall(r'<p:sp>.*?</p:sp>',x,re.S):
        ph=re.search(r'<p:ph([^>]*)/?>',sp)
        if not ph: continue
        t=(re.search(r'type="(\w+)"',ph.group(1)) or [None,'body'])[1]
        o=re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"',sp)
        phs.append(t+('@'+','.join(f'{int(v)/914400:.2f}' for v in o.groups()) if o else ''))
    print(f'{n.split("/")[-1]:22s} {mn[:40]:40s} {cn[:28]:28s} {" ".join(phs)}')
