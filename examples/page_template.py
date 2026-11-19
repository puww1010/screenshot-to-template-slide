"""page_template.py -- copy this file to start a new page, then fill in build().

Build it with:
  python3 scripts/build_slide.py --template T.pptx --page my_page.py --out OUT.pptx --qa-dir <abs path>/QA [--lang zh]

Rules that matter most (SKILL.md section 2): every word on the slide comes from the source screenshot, never from the
speaker script; keep the source layout; draw icons with slidekit shapes, no pictures; use theme colours only.
"""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
from slidekit import *

# Canvas scale: slidekit and references/page-patterns.md use a 10 x 5.625 in canvas. If build_slide.py prints
# "[WARN] canvas 20 x 11.25 in ... x2.0", set K = 2 and every coordinate / font size below scales with it.
K = 1


def build():
    S = [title_ph('Title copied verbatim from the screenshot')]          # inherits the layout's title style

    # Subtitle (optional): 11 pt, dk2. lnpct is a PERCENT (115 = 115 %).
    S.append(text('Subtitle', .5*K, .8*K, 9*K, .3*K,
                  [para([run('Subtitle copied verbatim from the screenshot', 11*K, 'dk2')], lnpct=115)]))

    # One card with an icon badge, a bold label and a short description.
    x, y, w, h = .5*K, 1.4*K, 2.8*K, 1.8*K
    S.append(shadow(shape('Card 1', x, y, w, h, 'roundRect', fill=sf('lt1'), av={'adj': 8000})))
    S.append(badge('shield_ok', x + w/2, y + .45*K, .5*K, 'solid'))
    S.append(text('Card 1 title', x + .1*K, y + .8*K, w - .2*K, .3*K, [para([run('Label', 11*K, 'dk1', b=True)], algn='ctr')]))
    S.append(text('Card 1 text', x + .1*K, y + 1.1*K, w - .2*K, .6*K, [para([run('One short line of text', 9*K, 'dk2')], algn='ctr')]))

    # Arrows, chevrons and badges that sit ON TOP of cards must be appended after the cards.
    return S
