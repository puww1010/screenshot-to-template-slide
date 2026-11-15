"""Example: chapter/section divider (use with --layout "<your template's Section Divider layout>").
Pattern: TEXT ONLY on the template's own background -- title (inherits the layout style), subtitle, progress meter.
No illustration card: section pages read better as a clean typographic break."""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from slidekit import *

def build():
    S=[title_ph('Proactive Discovery'),
       text('Subtitle',.8,3.2,8.4,.4,[para([run('Find issues before they find you',18,'lt1')],algn='ctr')])]
    S+=progress_meter(3,5)
    return S
