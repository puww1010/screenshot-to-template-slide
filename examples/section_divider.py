"""Example: chapter/section divider (use with --layout "<your template's Section Divider layout>").
Pattern: white stage card with an illustration on top, template title (inherits white 36pt), subtitle, progress meter."""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from slidekit import *

def build():
    S=[title_ph('Proactive Discovery'), stage_card()]
    # illustration: scout rover -> magnifier finds a bug in a network -> protecting shield
    S.append(shape('Rover tread',1.85,1.66,.86,.2,'roundRect',fill=sf('dk1'),av={'adj':50000}))
    for wx in (1.97,2.28,2.59): S.append(shape('Wheel',wx-.06,1.7,.12,.12,'ellipse',fill=sf('lt1')))
    S.append(shape('Rover body',1.92,1.42,.72,.26,'roundRect',fill=sf('lt2',SOFT),line=K(),av={'adj':20000}))
    S.append(vline('Mast',2.3,1.12,1.42,K(1.5)))
    S.append(shape('Camera',2.14,.94,.34,.22,'roundRect',fill=sf('lt1'),line=K(),av={'adj':20000}))
    S.append(shape('Lens',2.36,.98,.14,.14,'ellipse',fill=sf('accent1'),line=K(1)))
    S.append(path('Scan beam',2.5,.72,1.0,.82,[('M',0,.4),('L',.9,0),('C',1,.33,1,.67,.9,1)],closed=True,fill=sf('accent1',TINT)))
    N=[(3.7,.95,'hexagon'),(4.25,1.62,'ellipse'),(4.95,1.02,'hexagon'),(5.55,1.68,'ellipse'),(6.15,1.0,'hexagon')]
    for a,b in [(0,1),(0,2),(1,2),(1,3),(2,4),(3,4)]:
        S.append(pathabs('Link',[('M',N[a][0],N[a][1]),('L',N[b][0],N[b][1])],line=ln('dk2',1,mods=SOFT)))
    for (x,y,g) in N:
        d=.32 if g=='hexagon' else .26
        S.append(shape('Node',x-d/2,y-d/2,d,d,g,fill=sf('lt1'),line=K(1.1)))
    S.append(magnifier(4.95,1.05)); S.append(icon('bug',4.89,.99,.2,'accent1',1.0))
    S.append(hline('To shield',6.45,7.05,1.3,ln('accent1',1.5,'dash',head=True)))
    S.append(green_shield(7.15,.85,.6,.75))
    S.append(text('Subtitle',.8,3.2,8.4,.4,[para([run('Find issues before they find you',18,'lt1')],algn='ctr')]))
    S+=progress_meter(3,5)
    return S
