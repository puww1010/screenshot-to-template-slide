"""Example: big-stat + attack-chain + THEN/NOW timeline + verdict banner (the 'window has collapsed' pattern).
All numbers here are ILLUSTRATIVE for a fictional company (ACME demo)."""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from slidekit import *

def build():
    S=[title_ph('The response window has collapsed')]
    # hero stat ring + caption
    S+=ring(.55,1.08,1.55,41)
    S.append(text('Stat caption',2.35,1.12,2.55,1.5,[
        para([run('of critical alerts escalate into incidents within ',13,'dk2'),run('24 hours',13,'accent1',b=True),run(' of first detection',13,'dk2')],lnpct=110000),
        para([run('Illustrative data for a fictional company.',10,'dk2',i=True)],before=8)]))
    # chain
    xs=[5.6,7.35,9.1]; cy=1.55; d=.64
    for k,(x,(kd,st,lb)) in enumerate(zip(xs,[('search','light','Detected'),('bolt','tint','Escalated'),('warn','solid','Impacted')])):
        S.append(badge(kd,x,cy,d,st))
        S.append(text(f'Chain label {k+1}',x-.7,cy+d/2+.08,1.4,.25,[para([run(lb,11,'dk1',b=True)],algn='ctr')]))
        if k<2: S.append(hline(f'Chain arrow {k+1}',x+d/2+.12,xs[k+1]-d/2-.12,cy,ln('dk2',1,'dash',head=True)))
    S.append(path('24h bracket',xs[0]-.3,2.3,xs[2]-xs[0]+.6,.08,[('M',0,0),('L',0,1),('L',1,1),('L',1,0)],line=ln('accent1',1)))
    S.append(text('24h label',xs[0]-.3,2.44,xs[2]-xs[0]+.6,.25,[para([run('All inside 24 hours',10,'accent1',b=True)],algn='ctr')]))
    # THEN / NOW lanes
    TX=1.3; TW=9.5-TX; h=.52; ty=2.98
    S.append(text('Then label',.5,ty,.75,h,[para([run('THEN',11,'dk1',b=True)])],anchor='ctr'))
    S.append(shape('Then track',TX,ty,TW,h,'roundRect',fill=sf('lt2'),av={'adj':50000}))
    sx=[TX+.12,TX+2.2,TX+4.3]
    for k,((kd,lb),x) in enumerate(zip([('test','Triage the alert'),('calendar','Schedule the fix'),('deploy','Roll out the change')],sx)):
        S.append(badge(kd,x+.2,ty+h/2,.4,'light'))
        S.append(text(f'Step {k+1}',x+.5,ty,1.5,h,[para([run(lb,11,'dk1')])],anchor='ctr'))
        if k<2: S.append(hline(f'Step arrow {k+1}',x+.5+len(lb)*.075+.08,sx[k+1]-.06,ty+h/2,ln('dk2',.75,head=True)))
    S.append(text('Then span',TX+5.9,ty,TW-6.1,h,[para([run('Days, even weeks',11,'dk2',b=True)],algn='r')],anchor='ctr'))
    ny=ty+h+.18
    S.append(text('Now label',.5,ny,.75,h,[para([run('NOW',11,'accent1',b=True)])],anchor='ctr'))
    S.append(shape('Now track',TX,ny,TW,h,'roundRect',line=ln('dk2',.75,'dash',mods=SOFT),av={'adj':50000}))
    S.append(shape('Now pill',TX,ny,.75,h,'roundRect',fill=sf('accent1'),av={'adj':50000}))
    S.append(icon('clock',TX+.375,ny+h/2,.28,'lt1',1.25))
    S.append(text('Now text',TX+.92,ny,3,h,[para([run('Less than a day',13,'accent1',b=True)])],anchor='ctr'))
    # verdict banner
    y=4.4; vh=.7
    S.append(shape('Verdict panel',.5,y,9,vh,'roundRect',fill=sf('lt2',SOFT),av={'adj':14000}))
    S.append(icon('shield_broken',.95,y+vh/2,.4,'accent1',1.5))
    S.append(text('Verdict text',1.4,y,7.9,vh,[para([run('Waiting for the next maintenance window isn’t just slow anymore — it’s ',14,'dk1'),
                                                  run('a real risk',14,'accent1',b=True)])],anchor='ctr'))
    return S
