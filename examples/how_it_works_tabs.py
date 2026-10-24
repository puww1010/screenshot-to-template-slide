"""Example: 'How it works' process page — title, subtitle, tab strip (active step), step cards with chevrons,
highlighted gate card, banner and closing note. Text comes verbatim from the source screenshot."""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from slidekit import *
GT='<a:lumMod val="10000"/><a:lumOff val="90000"/>'

def build():
    S=[title_ph('How it works: automated build pipeline'),
       text('Subtitle',.5,.8,9,.24,[para([run('Every release is built, tested and verified before it reaches production',11,'dk2')])],anchor='ctr')]
    S+=tabs(['Collect','Build','Deliver'],active=1)
    steps=[('chip','Input received','From partners, customers and 20+ sources'),
           ('test','Prototype','AI builds a working prototype to prove the idea'),
           ('shield_ok','Package built','Automated build covers the change and its variants')]
    cw,g,y,h=2.0,.3,1.6,1.7
    for k,(kd,t,d) in enumerate(steps):
        x=.5+k*(cw+g); cx=x+cw/2
        S.append(shape(f'Step {k+1}',x,y,cw,h,'roundRect',fill=sf('lt2',SOFT),av={'adj':8000}))
        S.append(badge(kd,cx,y+.38,.5,'solid'))
        S.append(text(f'Step {k+1} title',x+.08,y+.72,cw-.16,.24,[para([run(t,10.5,'dk1',b=True)],algn='ctr')],anchor='ctr'))
        S.append(text(f'Step {k+1} text',x+.1,y+1.0,cw-.2,.62,[para([run(d,8.5,'dk2')],algn='ctr')]))
        S.append(chevron(x+cw+g/2,y+h/2))
    qx=.5+3*(cw+g)
    S.append(shape('Gate',qx,y,9.5-qx,h,'roundRect',fill=sf('accent4',GT),line=ln('accent4',1.25),av={'adj':8000}))
    S.append(text('Gate text',qx+.14,y+.1,9.5-qx-.24,h-.2,[para([run('QA gate',10,'accent4',b=True)]),
        para([run('Acceptance testing',9.5,'dk1',b=True)],before=4),para([run('Must pass every scenario and edge case',8.5,'dk2')]),
        para([run('Regression testing',9.5,'dk1',b=True)],before=5),para([run('Must not break existing workflows or integrations',8.5,'dk2')])]))
    S.append(shape('Banner',.5,3.6,9,.46,'roundRect',fill=sf('accent1',GT),line=ln('accent1',1),av={'adj':20000}))
    S.append(text('Banner text',.75,3.6,8.6,.46,[para([run('Released continuously in hours',10.5,'accent1',b=True),
        run('   No release window, no restarts, no manual steps',9.5,'dk1')])],anchor='ctr'))
    S.append(shape('Impact box',.5,4.24,9,.78,'roundRect',fill=sf('lt2',SOFT),av={'adj':12000}))
    S.append(shape('Check ring',.66,4.43,.4,.4,'ellipse',fill=sf('lt1'),line=ln('accent4',1.75)))
    S.append(path('Check',.77,4.57,.18,.12,[('M',0,.5),('L',.38,1),('L',1,0)],line=ln('accent4',2,cap=True)))
    S.append(text('Impact text',1.2,4.32,8.15,.62,[para([run('Operational impact: ',9.5,'accent4',b=True),
        run('every change is tested automatically before it reaches customers, so speed never comes at the cost of stability, even at high volume.',9.5,'dk1')],lnpct=105000)],anchor='ctr'))
    return S
