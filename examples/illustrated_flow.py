"""Example: illustration-heavy flow (source slide is mostly drawings). Tinted stage, white shadowed cards with arrows,
drawings above/below each card. Source had no title -> minimal title taken from the script (tell the user)."""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from slidekit import *

def build():
    S=[title_ph('Inside the delivery pipeline'),
       shape('Stage',.5,1.0,9.0,3.9,'roundRect',fill=sf('accent1','<a:lumMod val="10000"/><a:lumOff val="90000"/>'),av={'adj':5000})]
    cw,cg,y,h,x0=1.8,.53,2.48,.92,.6
    cx=[x0+k*(cw+cg)+cw/2 for k in range(4)]
    S+=[magnifier(cx[0]-.38,1.75),compass_rose(cx[0]+.38,1.75),server(cx[0],4.0),
        binary(cx[1]-.55,1.42),gears(cx[1],4.02),
        sandbox(cx[2],1.78),shield_half(cx[2],4.02),
        gate(cx[3],1.78),check(cx[3]-.35,4.0),rocket(cx[3]+.3,4.0)]
    for k,(t,d) in enumerate([('Plan','Pick scope & strategy'),('Build','Parallel work & candidates'),
                              ('Verify','Tested in a sandbox'),('Release','Score, approve, publish')]):
        x=x0+k*(cw+cg)
        S.append(shadow(shape(f'Card {k+1}',x,y,cw,h,'roundRect',fill=W,av={'adj':10000}),alpha=18000,blur=63500))
        S.append(text(f'Card {k+1} title',x+.16,y+.1,cw-.25,.34,[para([run(t,15,'dk1',b=True)])],anchor='ctr'))
        S.append(text(f'Card {k+1} text',x+.16,y+.46,cw-.25,.4,[para([run(d,10,'dk2')])]))
        if k<3: S.append(hline(f'Flow {k+1}',x+cw+.04,x+cw+cg-.04,y+h/2,ln('accent1',1.5,head=True)))
    return S
