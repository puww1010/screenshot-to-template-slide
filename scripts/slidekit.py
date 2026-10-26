"""slidekit — native, editable OOXML shapes for template-based slides.

All coordinates are in inches on a 10" x 5.625" (16:9) canvas. Colours are theme scheme colours
(dk1, lt1, dk2, lt2, accent1..accent6) so the template palette is always respected.
Usage:  import sys; sys.path.insert(0, SKILL_SCRIPTS); from slidekit import *
        shapes=[title_ph('...'), text(...), badge('shield_ok',...)]; inject('unpacked/ppt/slides/slide1.xml', shapes)
Icon kinds (icon()/badge()): globe bolt bug clock test calendar deploy shield_broken chat terminal chip target
  book warn shield_ok tl_broken burst search shield_spark vault hourglass people user building cloud factory
Illustration helpers: magnifier compass_rose server compass_dial gears binary sandbox test_tube shield_half gate key
  check rocket code_stack vault green_shield ic_rack ic_vm ic_containers ic_sase mini_badge label
"""
import re, os, math
E=914400
def emu(v): return int(round(v*E))
class Ctx:
    def __init__(s): s.i=10
    def nid(s): s.i+=1; return s.i
C=Ctx()
def esc(t): return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def sf(c,mods=''): return f'<a:solidFill><a:schemeClr val="{c}">{mods}</a:schemeClr></a:solidFill>'
TINT='<a:lumMod val="20000"/><a:lumOff val="80000"/>'      # light tint of accent
SOFT='<a:lumMod val="50000"/><a:lumOff val="50000"/>'      # lighter lt2
def ln(c,w=1.0,dash=None,head=False,mods='',cap=None):
    d=f'<a:prstDash val="{dash}"/>' if dash else ''
    h='<a:tailEnd type="triangle" w="med" len="med"/>' if head else ''
    cp=' cap="rnd"' if cap else ''
    return f'<a:ln w="{int(w*12700)}"{cp}>{sf(c,mods)}{d}<a:round/>{h}</a:ln>'
NOLN='<a:ln><a:noFill/></a:ln>'
def run(t,sz,color='dk1',b=False,i=False,spc=None):
    s=f' spc="{spc}"' if spc else ''
    return f'<a:r><a:rPr lang="en-US" sz="{int(sz*100)}" b="{int(b)}" i="{int(i)}"{s} dirty="0">{sf(color)}</a:rPr><a:t xml:space="preserve">{esc(t)}</a:t></a:r>'
def para(runs,algn='l',lnpct=None,before=0):
    """lnpct: line spacing in PERCENT (115 = 115 %). OOXML units (115000) are still accepted.
    OOXML spcPct is in 1/1000 %, so a raw 120 would mean 0.12 % and PowerPoint stacks every line on top of each other."""
    lnx=''
    if lnpct:
        v=int(round(lnpct*1000)) if lnpct<1000 else int(lnpct)
        if not 50000<=v<=300000: raise ValueError(f'para: line spacing {lnpct!r} = {v/1000:g} % is outside 50-300 %')
        lnx=f'<a:lnSpc><a:spcPct val="{v}"/></a:lnSpc>'
    bf=f'<a:spcBef><a:spcPts val="{int(before*100)}"/></a:spcBef>' if before else ''
    return f'<a:p><a:pPr algn="{algn}">{lnx}{bf}<a:buNone/></a:pPr>{"".join(runs)}</a:p>'
def body(paras,anchor='t'):
    p=paras or ['<a:p><a:endParaRPr lang="en-US"/></a:p>']
    return f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(p)}</p:txBody>'
def xfrm(x,y,w,h,extra=''):
    return f'<a:xfrm{extra}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{max(emu(w),1)}" cy="{max(emu(h),1)}"/></a:xfrm>'
def shape(name,x,y,w,h,geom='rect',fill=None,line=None,paras=None,anchor='t',txbox=False,av=None,rot=0):
    a='<a:avLst>'+''.join(f'<a:gd name="{k}" fmla="val {v}"/>' for k,v in (av or {}).items())+'</a:avLst>'
    tb=' txBox="1"' if txbox else ''
    r=f' rot="{int(rot*60000)}"' if rot else ''
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{C.nid()}" name="{esc(name)}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{xfrm(x,y,w,h,r)}<a:prstGeom prst="{geom}">{a}</a:prstGeom>{fill or "<a:noFill/>"}{line or NOLN}</p:spPr>{body(paras,anchor)}</p:sp>')
def text(name,x,y,w,h,paras,anchor='t'):
    return shape(name,x,y,w,h,paras=paras,anchor=anchor,txbox=True)
def path(name,x,y,w,h,segs,closed=False,fill=None,line=None):
    """segs: list of ('M',u,v) ('L',u,v) ('C',u1,v1,u2,v2,u3,v3) in 0..1 of box"""
    P=lambda u,v:f'<a:pt x="{int(u*1000)}" y="{int(v*1000)}"/>'
    out=''
    for s in segs:
        if s[0]=='M': out+=f'<a:moveTo>{P(s[1],s[2])}</a:moveTo>'
        elif s[0]=='L': out+=f'<a:lnTo>{P(s[1],s[2])}</a:lnTo>'
        elif s[0]=='C': out+=f'<a:cubicBezTo>{P(s[1],s[2])}{P(s[3],s[4])}{P(s[5],s[6])}</a:cubicBezTo>'
    if closed: out+='<a:close/>'
    fl='' if (closed and fill) else ' fill="none"'
    geom=f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="1000" h="1000"{fl}>{out}</a:path></a:pathLst></a:custGeom>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{C.nid()}" name="{esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{xfrm(x,y,w,h)}{geom}{fill if (closed and fill) else "<a:noFill/>"}{line or NOLN}</p:spPr>{body(None)}</p:sp>')
def hline(name,x1,x2,y,line): return path(name,x1,y,x2-x1,0,[('M',0,0),('L',1,0)],line=line)
def vline(name,x,y1,y2,line): return path(name,x,y1,0,y2-y1,[('M',0,0),('L',0,1)],line=line)
def group(name,x,y,w,h,children):
    X=xfrm(x,y,w,h).replace('</a:xfrm>',f'<a:chOff x="{emu(x)}" y="{emu(y)}"/><a:chExt cx="{max(emu(w),1)}" cy="{max(emu(h),1)}"/></a:xfrm>')
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{C.nid()}" name="{esc(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr>{X}</p:grpSpPr>{"".join(children)}</p:grpSp>')

# ---------- icon library (all native, editable) ----------
def _icon_base(kind,cx,cy,s,col='dk1',w=1.25):
    x0,y0=cx-s/2,cy-s/2
    B=lambda u1,v1,u2,v2:(x0+u1*s,y0+v1*s,(u2-u1)*s,(v2-v1)*s)
    L=ln(col,w,cap=True)
    def seg(u1,v1,u2,v2,c=None):
        bx=min(u1,u2);by=min(v1,v2);bw=abs(u2-u1);bh=abs(v2-v1)
        nu=lambda u:(u-bx)/bw if bw else 0
        nv=lambda v:(v-by)/bh if bh else 0
        return path('line',x0+bx*s,y0+by*s,bw*s,bh*s,[('M',nu(u1),nv(v1)),('L',nu(u2),nv(v2))],line=c or L)
    def poly(pts,closed=False,fill=None,line=L):
        us=[p[0] for p in pts];vs=[p[1] for p in pts]
        bx,by=min(us),min(vs);bw=max(us)-bx or 1e-3;bh=max(vs)-by or 1e-3
        segs=[('M' if k==0 else 'L',(u-bx)/bw,(v-by)/bh) for k,(u,v) in enumerate(pts)]
        return path('stroke',x0+bx*s,y0+by*s,bw*s,bh*s,segs,closed=closed,fill=fill,line=line)
    k=[]
    if kind=='globe':
        k.append(shape('globe',*B(0,0,1,1),'ellipse',line=L))
        k.append(shape('meridian',*B(.3,0,.7,1),'ellipse',line=L))
        k+= [seg(0,.5,1,.5), seg(.07,.26,.93,.26), seg(.07,.74,.93,.74)]
    elif kind=='bolt':
        k.append(shape('bolt',*B(.15,0,.85,1),'lightningBolt',fill=sf(col)))
    elif kind=='bug':
        k.append(shape('body',*B(.3,.36,.7,.96),'ellipse',fill=sf(col)))
        k.append(shape('head',*B(.37,.14,.63,.38),'ellipse',fill=sf(col)))
        for (a,b,c,d) in [(.3,.52,.08,.42),(.3,.68,.06,.68),(.32,.84,.1,.94),(.45,.18,.34,.03)]:
            k.append(seg(a,b,c,d)); k.append(seg(1-a,b,1-c,d))
        k.append(seg(.5,.44,.5,.92,ln('lt1',w*.8)))
    elif kind=='clock':
        k.append(shape('face',*B(0,0,1,1),'ellipse',line=L))
        k+= [seg(.5,.5,.5,.2), seg(.5,.5,.72,.62)]
    elif kind=='test':
        k.append(shape('board',*B(.16,.1,.84,1),'roundRect',line=L,av={'adj':12000}))
        k.append(shape('clip',*B(.34,.02,.66,.18),'roundRect',fill=sf(col),av={'adj':30000}))
        k.append(poly([(.32,.58),(.46,.72),(.68,.44)]))
    elif kind=='calendar':
        k.append(shape('cal',*B(.06,.14,.94,.96),'roundRect',line=L,av={'adj':12000}))
        k+= [seg(.06,.38,.94,.38), seg(.3,.04,.3,.24), seg(.7,.04,.7,.24)]
        for v in (.56,.78):
            for u in (.3,.5,.7):
                k.append(shape('day',*B(u-.05,v-.05,u+.05,v+.05),'ellipse',fill=sf(col)))
    elif kind=='deploy':
        k.append(shape('rack1',*B(.02,.12,.7,.44),'roundRect',line=L,av={'adj':20000}))
        k.append(shape('rack2',*B(.02,.56,.7,.88),'roundRect',line=L,av={'adj':20000}))
        for v in (.28,.72): k.append(shape('led',*B(.13,v-.05,.23,v+.05),'ellipse',fill=sf(col)))
        k.append(seg(.88,.96,.88,.36)); k.append(poly([(.76,.5),(.88,.36),(1,.5)]))
    elif kind=='shield_broken':
        k.append(path('shield',*B(.06,0,.94,1),[('M',.5,0),('L',1,.16),('L',1,.5),('C',1,.78,.74,.92,.5,1),('C',.26,.92,0,.78,0,.5),('L',0,.16)],closed=True,line=L))
        k.append(poly([(.54,.14),(.42,.44),(.6,.56),(.46,.86)]))
    return group(f'Icon {kind}',x0,y0,s,s,k)

def badge(kind,cx,cy,d,style='light'):
    """circle badge with icon. style: light(white/dk2 stroke), tint(accent tint), solid(accent1/white icon)"""
    if style=='solid': f,l,ic=sf('accent1'),None,'lt1'
    elif style=='tint': f,l,ic=sf('accent1',TINT),None,'accent1'
    else: f,l,ic=sf('lt1'),ln('dk2',.75),'dk1'
    circ=shape(f'Badge {kind}',cx-d/2,cy-d/2,d,d,'ellipse',fill=f,line=l)
    return group(f'Badge {kind}',cx-d/2,cy-d/2,d,d,[circ,icon(kind,cx,cy,d*.52,ic,1.25 if d>.5 else 1.0)])

def title(t):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{C.nid()}" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{esc(t)}</a:t></a:r></a:p></p:txBody></p:sp>')


def ring(x,y,d,pct):
    th=11000
    end=(270+360*pct/100)%360
    return [shape('Ring track',x,y,d,d,'donut',fill=sf('lt2'),av={'adj':th}),
            shape('Ring 29%',x,y,d,d,'blockArc',fill=sf('accent1'),av={'adj1':16200000,'adj2':int(end*60000),'adj3':th}),
            text('Ring value',x,y,d,d,[para([run(f'{pct}%',30,'accent1',b=True)],algn='ctr')],anchor='ctr')]


# ---------- icon set 2 ----------
def icon2(kind,cx,cy,s,col='dk1',w=1.25):
    x0,y0=cx-s/2,cy-s/2
    B=lambda u1,v1,u2,v2:(x0+u1*s,y0+v1*s,(u2-u1)*s,(v2-v1)*s)
    L=ln(col,w,cap=True)
    def poly(pts,closed=False,fill=None,line=L):
        us=[p[0] for p in pts];vs=[p[1] for p in pts]
        bx,by=min(us),min(vs);bw=max(us)-bx or 1e-3;bh=max(vs)-by or 1e-3
        segs=[('M' if k==0 else 'L',(u-bx)/bw,(v-by)/bh) for k,(u,v) in enumerate(pts)]
        return path('stroke',x0+bx*s,y0+by*s,bw*s,bh*s,segs,closed=closed,fill=fill,line=line)
    seg=lambda a,b,c,d,l=None: poly([(a,b),(c,d)],line=l or L)
    dot=lambda u,v,r=.05: shape('dot',*B(u-r,v-r,u+r,v+r),'ellipse',fill=sf(col))
    k=[]
    if kind=='chat':
        k.append(shape('bubble',*B(0,.06,1,.76),'roundRect',line=L,av={'adj':30000}))
        k.append(poly([(.24,.76),(.2,.98),(.46,.76)]))
        k+=[dot(.3,.41),dot(.5,.41),dot(.7,.41)]
    elif kind=='terminal':
        k.append(shape('window',*B(0,.1,1,.9),'roundRect',line=L,av={'adj':12000}))
        k.append(seg(0,.3,1,.3))
        k.append(poly([(.18,.46),(.34,.58),(.18,.7)]))
        k.append(seg(.42,.72,.68,.72))
    elif kind=='chip':
        k.append(shape('core',*B(.18,.18,.82,.82),'roundRect',line=L,av={'adj':14000},
                       paras=[para([run('AI',max(6,round(s*18)),col,b=True)],algn='ctr')],anchor='ctr'))
        for t in (.36,.5,.64):
            k+=[seg(t,0,t,.18),seg(t,.82,t,1),seg(0,t,.18,t),seg(.82,t,1,t)]
    elif kind=='target':
        k.append(shape('ring',*B(.1,.1,.9,.9),'ellipse',line=L))
        k+=[seg(.5,0,.5,.3),seg(.5,.7,.5,1),seg(0,.5,.3,.5),seg(.7,.5,1,.5),dot(.5,.5,.07)]
    elif kind=='book':
        k.append(poly([(.5,.22),(.03,.08),(.03,.84),(.5,.98)],closed=True))
        k.append(poly([(.5,.22),(.97,.08),(.97,.84),(.5,.98)],closed=True))
    elif kind=='warn':
        k.append(shape('tri',*B(0,.04,1,.96),'triangle',line=L))
        k+=[seg(.5,.4,.5,.66),dot(.5,.78,.045)]
    elif kind=='shield_ok':
        k.append(path('shield',*B(.06,0,.94,1),[('M',.5,0),('L',1,.16),('L',1,.5),('C',1,.78,.74,.92,.5,1),('C',.26,.92,0,.78,0,.5),('L',0,.16)],closed=True,line=L))
        k.append(poly([(.32,.5),(.46,.64),(.7,.38)]))
    elif kind=='tl_broken':
        k+=[seg(0,.5,.4,.5),dot(.12,.5,.06),dot(.3,.5,.06)]
        k.append(poly([(.44,.3),(.5,.42),(.44,.56),(.5,.7)]))
        k.append(poly([(.58,.34),(.64,.46),(.58,.6),(.64,.74)]))
        k.append(seg(.68,.6,1,.78))
    else:
        return _icon_base(kind,cx,cy,s,col,w)
    return group(f'Icon {kind}',x0,y0,s,s,k)

def nottag(cx,cy,d=.2):
    return group('Not tag',cx-d/2,cy-d/2,d,d,[shape('Not circle',cx-d/2,cy-d/2,d,d,'ellipse',fill=sf('accent1'),line=ln('lt1',1)),
        path('x1',cx-d*.2,cy-d*.2,d*.4,d*.4,[('M',0,0),('L',1,1)],line=ln('lt1',1.25,cap=True)),
        path('x2',cx-d*.2,cy-d*.2,d*.4,d*.4,[('M',1,0),('L',0,1)],line=ln('lt1',1.25,cap=True))])


# ---------- icon set 3 ----------
def icon3(kind,cx,cy,s,col='dk1',w=1.25):
    x0,y0=cx-s/2,cy-s/2
    B=lambda u1,v1,u2,v2:(x0+u1*s,y0+v1*s,(u2-u1)*s,(v2-v1)*s)
    L=ln(col,w,cap=True)
    def seg(a,b,c,d,l=None):
        bx,by=min(a,c),min(b,d); bw=abs(c-a) or 1e-3; bh=abs(d-b) or 1e-3
        return path('line',x0+bx*s,y0+by*s,bw*s,bh*s,[('M',(a-bx)/bw,(b-by)/bh),('L',(c-bx)/bw,(d-by)/bh)],line=l or L)
    shield=lambda u1,v1,u2,v2: path('shield',*B(u1,v1,u2,v2),[('M',.5,0),('L',1,.16),('L',1,.5),('C',1,.78,.74,.92,.5,1),('C',.26,.92,0,.78,0,.5),('L',0,.16)],closed=True,line=L)
    k=[]
    if kind=='burst':
        k.append(icon2('bug',cx,cy,s*.5,col,w))
        for a in range(0,360,45):
            r1,r2=.33,.48; t=math.radians(a)
            k.append(seg(.5+r1*math.cos(t),.5+r1*math.sin(t),.5+r2*math.cos(t),.5+r2*math.sin(t)))
    elif kind=='search':
        k.append(shape('lens',*B(.04,.04,.7,.7),'ellipse',line=ln(col,w*1.2)))
        k.append(seg(.62,.62,.96,.96,ln(col,w*1.8,cap=True)))
        k.append(path('glint',*B(.2,.18,.36,.34),[('M',0,1),('C',0,.4,.4,0,1,0)],line=ln(col,w*.8,cap=True)))
    elif kind=='shield_spark':
        k.append(shield(.02,.1,.78,1))
        k.append(shape('spark',*B(.6,0,1,.4),'star4',fill=sf(col),av={'adj':16000}))
        k.append(shape('spark2',*B(.26,.42,.52,.68),'star4',fill=sf(col),av={'adj':16000}))
    elif kind=='vault':
        k.append(shape('door',*B(0,.04,1,.88),'roundRect',line=L,av={'adj':12000}))
        k.append(shape('dial',*B(.3,.2,.7,.6),'ellipse',line=L))
        for a in (0,120,240):
            t=math.radians(a-90)
            k.append(seg(.5,.4,.5+.2*math.cos(t),.4+.2*math.sin(t)))
        k.append(seg(.84,.3,.84,.54))
        k+=[seg(.14,.88,.14,.98),seg(.86,.88,.86,.98)]
    else:
        return icon2(kind,cx,cy,s,col,w)
    return group(f'Icon {kind}',x0,y0,s,s,k)

# ---------- icon set 4 ----------
def icon4(kind,cx,cy,s,col='dk1',w=1.25):
    x0,y0=cx-s/2,cy-s/2
    B=lambda u1,v1,u2,v2:(x0+u1*s,y0+v1*s,(u2-u1)*s,(v2-v1)*s)
    L=ln(col,w,cap=True)
    def poly(pts,closed=False,fill=None,line=L):
        us=[p[0] for p in pts];vs=[p[1] for p in pts]
        bx,by=min(us),min(vs);bw=max(us)-bx or 1e-3;bh=max(vs)-by or 1e-3
        segs=[('M' if k==0 else 'L',(u-bx)/bw,(v-by)/bh) for k,(u,v) in enumerate(pts)]
        return path('stroke',x0+bx*s,y0+by*s,bw*s,bh*s,segs,closed=closed,fill=fill,line=line)
    def person(u,v,r):  # head + shoulders centered at u, head top v
        return [shape('head',*B(u-r*.45,v,u+r*.45,v+r*.9),'ellipse',line=L),
                path('shoulders',*B(u-r,v+r*1.05,u+r,v+r*2.0),[('M',0,1),('C',0,.1,1,.1,1,1)],line=L)]
    k=[]
    if kind=='hourglass':
        k+=[poly([(.12,.02),(.88,.02)]),poly([(.12,.98),(.88,.98)]),
            poly([(.22,.04),(.78,.04),(.5,.5),(.78,.96),(.22,.96),(.5,.5)],closed=True)]
        k.append(poly([(.36,.86),(.64,.86),(.5,.7)],closed=True,fill=sf(col)))
    elif kind=='people':
        k+=person(.32,.12,.26)+person(.7,.2,.22)
    elif kind=='user':
        k+=person(.5,.05,.42)
    elif kind=='building':
        k.append(shape('tower',*B(.22,.04,.78,.98),'rect',line=L))
        for v in (.25,.47,.69):
            for u in (.38,.62): k.append(shape('win',*B(u-.05,v-.05,u+.05,v+.05),'rect',fill=sf(col)))
    elif kind=='cloud':
        k.append(shape('cloud',*B(0,.15,1,.85),'cloud',line=L))
    elif kind=='factory':
        k.append(poly([(0,.98),(0,.5),(.3,.3),(.3,.5),(.6,.3),(.6,.5),(.82,.5),(.82,.06),(.98,.06),(.98,.98)],closed=True))
    else:
        return icon3(kind,cx,cy,s,col,w)
    return group(f'Icon {kind}',x0,y0,s,s,k)

icon=icon4

# ---------- drawing helpers ----------
K=lambda w=1.25: ln('dk1',w,cap=True)
def seg(name,x1,y1,x2,y2,line,head=False):
    bx,by=min(x1,x2),min(y1,y2); bw=abs(x2-x1) or .001; bh=abs(y2-y1) or .001
    return path(name,bx,by,bw,bh,[('M',(x1-bx)/bw,(y1-by)/bh),('L',(x2-bx)/bw,(y2-by)/bh)],line=line)
def pathabs(name,pts,closed=False,fill=None,line=None):
    """pts: list of ('M'|'L',x,y) or ('C',x1,y1,x2,y2,x3,y3) absolute inches"""
    xs=[];ys=[]
    for p in pts:
        xs+=p[1::2]; ys+=p[2::2]
    bx,by=min(xs),min(ys); bw=(max(xs)-bx) or .001; bh=(max(ys)-by) or .001
    segs=[]
    for p in pts:
        n=[p[0]]
        for i in range(1,len(p),2): n+=[(p[i]-bx)/bw,(p[i+1]-by)/bh]
        segs.append(tuple(n))
    return path(name,bx,by,bw,bh,segs,closed=closed,fill=fill,line=line)
WHITE=sf('lt1')


W=sf('lt1'); TINT2=sf('accent1',TINT)

def magnifier(cx,cy):
    return group('Magnifier',cx-.25,cy-.25,.5,.5,[seg('Handle',cx+.09,cy+.09,cx+.24,cy+.24,ln('dk1',3,cap=True)),
        shape('Lens',cx-.24,cy-.24,.36,.36,'ellipse',fill=W,line=K(1.5))])
def compass_rose(cx,cy):
    return group('Compass rose',cx-.24,cy-.24,.48,.48,[shape('Rose ring',cx-.24,cy-.24,.48,.48,'ellipse',fill=TINT2,line=K(1.25)),
        shape('Rose star',cx-.18,cy-.18,.36,.36,'star4',fill=sf('dk1'),av={'adj':12500}),
        shape('Rose star 2',cx-.13,cy-.13,.26,.26,'star4',fill=sf('accent1'),av={'adj':12500},rot=45)])
def server(cx,cy):
    S=[]
    for k in range(3):
        y=cy-.26+k*.17
        S.append(shape('Server unit',cx-.27,y,.54,.14,'roundRect',fill=W,line=K(1),av={'adj':25000}))
        S.append(shape('LED',cx-.2,y+.045,.05,.05,'ellipse',fill=sf('accent1')))
        S.append(hline('Slot',cx-.08,cx+.18,y+.07,K(.75)))
    S+= [vline('Stem',cx,cy+.25,cy+.33,K(1)),hline('Base',cx-.18,cx+.18,cy+.33,K(1.25))]
    return group('Server',cx-.27,cy-.26,.54,.6,S)
def compass_dial(cx,cy):
    return group('Compass',cx-.22,cy-.22,.44,.44,[shape('Dial',cx-.22,cy-.22,.44,.44,'ellipse',fill=W,line=K(1.5)),
        shape('Dial inner',cx-.16,cy-.16,.32,.32,'ellipse',fill=TINT2),
        pathabs('Needle',[('M',cx+.1,cy-.1),('L',cx+.03,cy+.03),('L',cx-.1,cy+.1),('L',cx-.03,cy-.03)],closed=True,fill=sf('dk1'))])
def gears(cx,cy):
    return group('Gears',cx-.25,cy-.22,.55,.5,[shape('Gear big',cx-.25,cy-.22,.36,.36,'gear9',fill=W,line=K(1.1),av={'adj1':10000,'adj2':1763}),
        shape('Hub',cx-.12,cy-.09,.1,.1,'ellipse',fill=sf('dk1')),
        shape('Gear small',cx+.06,cy+.04,.24,.24,'gear6',fill=TINT2,line=K(1.1),av={'adj1':15000,'adj2':3526})])
def binary(x,y):
    S=[]
    for k,b in enumerate(['0110101011','1101001101','1010110010','0101101011']):
        S.append(text('Binary',x+(k%2)*.04,y+k*.15,.78,.14,[para([run(b,7,'dk2',b=True)])]).replace('<a:rPr lang="en-US" sz="700"','<a:rPr lang="en-US" sz="700"',1))
        S.append(hline('Binary arrow',x+.84,x+1.02,y+k*.15+.07,ln('dk1',1,head=True)))
    return group('Binary stream',x,y,1.06,.6,S)
def sandbox(cx,cy):
    S=[pathabs('Castle',[('M',cx-.32,cy+.22),('L',cx-.32,cy-.16),('L',cx-.24,cy-.16),('L',cx-.24,cy-.24),('L',cx-.14,cy-.24),('L',cx-.14,cy-.16),
            ('L',cx-.05,cy-.16),('L',cx-.05,cy-.24),('L',cx+.05,cy-.24),('L',cx+.05,cy-.16),('L',cx+.14,cy-.16),('L',cx+.14,cy-.24),
            ('L',cx+.24,cy-.24),('L',cx+.24,cy-.16),('L',cx+.32,cy-.16),('L',cx+.32,cy+.22)],closed=True,fill=sf('lt2'),line=K(1.25)),
       pathabs('Door',[('M',cx-.08,cy+.22),('L',cx-.08,cy+.04),('C',cx-.08,cy-.06,cx+.08,cy-.06,cx+.08,cy+.04),('L',cx+.08,cy+.22)],closed=True,fill=sf('dk1')),
       shape('Sandbox base',cx-.45,cy+.17,.9,.22,'ellipse',fill=TINT2,line=K(1),paras=[para([run('sand-box',7,'dk1',b=True)],algn='ctr')],anchor='ctr')]
    return group('Sandbox',cx-.45,cy-.24,.9,.63,S)
def test_tube(cx,cy):
    T=[shape('Tube',cx-.06,cy-.24,.12,.48,'roundRect',fill=W,line=K(1.25),av={'adj':50000}),
       shape('Liquid',cx-.04,cy+.02,.08,.19,'roundRect',fill=sf('accent1'),av={'adj':50000})]
    S=[group('Tube body',cx-.06,cy-.24,.12,.48,T).replace('<a:xfrm>','<a:xfrm rot="1500000">',1)]
    for (x1,y1,x2,y2) in [(cx+.12,cy-.3,cx+.16,cy-.38),(cx+.2,cy-.24,cx+.28,cy-.28),(cx+.04,cy-.34,cx+.04,cy-.42)]:
        S.append(seg('Spark',x1,y1,x2,y2,ln('accent1',1.25,cap=True)))
    return group('Test tube',cx-.16,cy-.42,.46,.7,S)
def shield_half(cx,cy):
    sh=[('M',.5,0),('L',1,.16),('L',1,.5),('C',1,.78,.74,.92,.5,1),('C',.26,.92,0,.78,0,.5),('L',0,.16)]
    return group('Shield',cx-.2,cy-.25,.4,.5,[path('Shield',cx-.2,cy-.25,.4,.5,sh,closed=True,fill=W,line=K(1.5)),
        path('Shield half',cx,cy-.21,.16,.42,[('M',0,0),('L',1,.12),('L',1,.48),('C',1,.72,.5,.88,0,1)],closed=True,fill=sf('dk2',SOFT))])
def gate(cx,cy):
    S=[shape('Post L',cx-.42,cy-.3,.1,.62,'rect',fill=sf('dk2'),line=K(1)),shape('Post R',cx+.32,cy-.3,.1,.62,'rect',fill=sf('dk2'),line=K(1)),
       shape('Cap L',cx-.44,cy-.36,.14,.07,'rect',fill=sf('dk1')),shape('Cap R',cx+.3,cy-.36,.14,.07,'rect',fill=sf('dk1')),
       pathabs('Arch',[('M',cx-.32,cy-.12),('C',cx-.16,cy-.32,cx+.16,cy-.32,cx+.32,cy-.12)],line=K(1.5)),
       hline('Rail',cx-.32,cx+.32,cy+.18,K(1.25))]
    for k in range(7):
        x=cx-.27+k*.09
        S.append(vline('Bar',x,cy-.2+abs(k-3)*.025,cy+.3,K(1)))
    S.append(shape('Lock',cx-.04,cy+.02,.08,.1,'rect',fill=sf('accent1')))
    return group('Gate',cx-.44,cy-.36,.88,.68,S)
def key(cx,cy):
    return group('Key',cx-.22,cy-.08,.46,.16,[shape('Bow',cx-.22,cy-.08,.16,.16,'ellipse',fill=TINT2,line=K(1.5)),
        hline('Shaft',cx-.06,cx+.24,cy,K(1.75)),vline('Tooth',cx+.14,cy,cy+.07,K(1.5)),vline('Tooth',cx+.21,cy,cy+.07,K(1.5))])
def check(cx,cy):
    return pathabs('Check',[('M',cx-.2,cy),('L',cx-.06,cy+.15),('L',cx+.22,cy-.18)],line=ln('accent4',4,cap=True))
def rocket(cx,cy):
    body=[('M',.5,0),('C',.85,.25,.8,.75,.68,.85),('L',.32,.85),('C',.2,.75,.15,.25,.5,0)]
    S=[path('Flame',cx-.08,cy+.18,.16,.2,[('M',0,0),('L',1,0),('L',.5,1)],closed=True,fill=sf('accent1')),
       path('Fin L',cx-.2,cy+.02,.12,.2,[('M',1,0),('L',0,1),('L',1,.8)],closed=True,fill=sf('dk2'),line=K(1)),
       path('Fin R',cx+.08,cy+.02,.12,.2,[('M',0,0),('L',1,1),('L',0,.8)],closed=True,fill=sf('dk2'),line=K(1)),
       path('Rocket body',cx-.14,cy-.32,.28,.56,body,closed=True,fill=W,line=K(1.25)),
       shape('Window',cx-.05,cy-.12,.1,.1,'ellipse',fill=TINT2,line=K(1))]
    g=group('Rocket',cx-.2,cy-.32,.4,.7,S)
    return g.replace('<a:xfrm>','<a:xfrm rot="2700000">',1)


SH=[('M',.5,0),('L',1,.16),('L',1,.5),('C',1,.78,.74,.92,.5,1),('C',.26,.92,0,.78,0,.5),('L',0,.16)]
def label(t,cx,y,w=1.0):
    return text('Label '+t,cx-w/2,y,w,.2,[para([run(t,9,'dk1',b=True)],algn='ctr')])
def code_stack(x,y,n):
    S=[]
    for k in range(n):
        xx,yy=x+k*.06,y+k*.09
        S.append(shape('Code doc',xx,yy,.36,.22,'roundRect',fill=W,line=K(1),av={'adj':15000}))
        S.append(path('Code chevron',xx+.05,yy+.06,.05,.1,[('M',1,0),('L',0,.5),('L',1,1)],line=K(1)))
        S.append(hline('Code line',xx+.13,xx+.3,yy+.11,ln('dk2',1,cap=True)))
    return group('Variant stack',x,y,.36+(n-1)*.06,.22+(n-1)*.09,S)
def vault(x,y,w,h):
    return group('Vault',x,y,w,h,[shape('Vault body',x,y,w,h,'roundRect',fill=sf('lt2'),line=K(1.25),av={'adj':12000}),
        shape('Vault door',x+w*.1,y+h*.12,w*.62,h*.76,'roundRect',fill=W,line=K(1),av={'adj':10000}),
        shape('Dial',x+w*.24,y+h*.32,w*.34,w*.34,'ellipse',fill=TINT2,line=K(1)),
        vline('Hinge',x+w*.85,y+h*.3,y+h*.7,K(1.5))])
def green_shield(x,y,w,h):
    return group('Signature shield',x,y,w,h,[path('Shield',x,y,w,h,SH,closed=True,fill=sf('accent4'),line=K(1.25)),
        shape('Plus',x+w*.27,y+h*.25,w*.46,w*.46,'mathPlus',fill=W,av={'adj1':23520})])


def ic_rack(cx,cy):
    S=[]
    for k in range(2):
        y=cy-.13+k*.14
        S+= [shape('Unit',cx-.17,y,.34,.12,'roundRect',fill=W,line=K(1),av={'adj':20000}),
             shape('LED',cx-.13,y+.04,.04,.04,'ellipse',fill=sf('accent1')),hline('Port',cx-.04,cx+.12,y+.06,K(.75))]
    return group('Icon rack',cx-.17,cy-.13,.34,.26,S)
def ic_vm(cx,cy):
    return group('Icon VM',cx-.17,cy-.16,.36,.32,[
        pathabs('Cube front',[('M',cx-.15,cy-.06),('L',cx+.03,cy-.06),('L',cx+.03,cy+.14),('L',cx-.15,cy+.14)],closed=True,fill=W,line=K(1)),
        pathabs('Cube top',[('M',cx-.15,cy-.06),('L',cx-.07,cy-.14),('L',cx+.11,cy-.14),('L',cx+.03,cy-.06)],closed=True,fill=sf('lt2'),line=K(1)),
        pathabs('Cube side',[('M',cx+.03,cy-.06),('L',cx+.11,cy-.14),('L',cx+.11,cy+.06),('L',cx+.03,cy+.14)],closed=True,fill=sf('lt2'),line=K(1)),
        shape('Cloud',cx-.01,cy+.0,.2,.14,'cloud',fill=TINT2,line=K(1))])
def ic_containers(cx,cy):
    S=[]
    for (dx,dy) in [(-.12,.02),(.0,.02),(.12,.02),(-.06,-.12),(.06,-.12)]:
        S.append(shape('Box',cx+dx-.055,cy+dy,.11,.11,'rect',fill=TINT2 if dy<0 else W,line=K(1)))
    return group('Icon containers',cx-.18,cy-.12,.36,.25,S)
def ic_sase(cx,cy):
    return group('Icon SASE',cx-.18,cy-.16,.36,.32,[
        shape('Globe',cx-.02,cy-.12,.2,.2,'ellipse',fill=W,line=K(1)),shape('Meridian',cx+.04,cy-.12,.08,.2,'ellipse',line=K(.75)),
        hline('Equator',cx-.02,cx+.18,cy-.02,K(.75)),
        path('Shield',cx-.18,cy-.14,.2,.25,SH,closed=True,fill=TINT2,line=K(1))])
def mini_badge(cx,cy,ok=True):
    if ok:
        return group('Shielded',cx-.13,cy-.13,.26,.26,[shape('Ok',cx-.13,cy-.13,.26,.26,'ellipse',fill=sf('accent4'),line=ln('lt1',1.25)),
            path('Tick',cx-.06,cy-.03,.12,.07,[('M',0,.5),('L',.38,1),('L',1,0)],line=ln('lt1',1.5,cap=True))])
    return group('Blocked',cx-.13,cy-.13,.26,.26,[shape('Stop',cx-.13,cy-.13,.26,.26,'ellipse',fill=sf('accent1'),line=ln('lt1',1.25)),
        path('x',cx-.05,cy-.05,.1,.1,[('M',0,0),('L',1,1)],line=ln('lt1',1.5,cap=True)),
        path('x',cx-.05,cy-.05,.1,.1,[('M',1,0),('L',0,1)],line=ln('lt1',1.5,cap=True))])


# ======================= page-level helpers =======================
def shadow(xml,alpha=20000,blur=76200,dist=19050):
    """Add a soft drop shadow to a shape XML that has <a:ln><a:noFill/></a:ln> or any </a:ln>."""
    eff=(f'<a:effectLst><a:outerShdw blurRad="{blur}" dist="{dist}" dir="5400000" rotWithShape="0">'
         f'<a:schemeClr val="dk1"><a:alpha val="{alpha}"/></a:schemeClr></a:outerShdw></a:effectLst>')
    return xml.replace('</a:ln></p:spPr>','</a:ln>'+eff+'</p:spPr>',1)

def hl_run(t,sz,color='dk1',b=True):
    """Text run with highlighter (theme accent2) behind it. Needs Office 2019+/365."""
    return run(t,sz,color,b).replace('</a:solidFill></a:rPr>','</a:solidFill><a:highlight><a:schemeClr val="accent2"/></a:highlight></a:rPr>',1)

def title_ph(t, x=None,y=None,w=None,h=None, sz=None, algn=None):
    """Title placeholder. With no args it fully inherits the layout (preferred).
    Pass x/y/w/h/sz/algn only when the source slide demands it (slide-level override, template untouched)."""
    sp='<p:spPr/>' if x is None else f'<p:spPr>{xfrm(x,y,w,h)}</p:spPr>'
    ppr=f'<a:pPr algn="{algn}"/>' if algn else ''
    rpr=f' sz="{int(sz*100)}"' if sz else ''
    bp='<a:bodyPr anchor="ctr"/>' if x is not None else '<a:bodyPr/>'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{C.nid()}" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>{sp}'
            f'<p:txBody>{bp}<a:lstStyle/><a:p>{ppr}<a:r><a:rPr lang="en-US"{rpr} dirty="0"/><a:t>{esc(t)}</a:t></a:r></a:p></p:txBody></p:sp>')

def tabs(labels, active, y=1.12, x0=.5, total_w=9.0, h=.3, gap=.12, sz=10):
    """Process tab strip ('1 Collect | 2 Build | 3 Deliver'); `active` is 0-based."""
    n=len(labels); w=(total_w-gap*(n-1))/n; S=[]
    for k,t in enumerate(labels):
        on=k==active; x=x0+k*(w+gap)
        S.append(shape(f'Tab {k+1}',x,y,w,h,'roundRect',fill=sf('accent1') if on else sf('lt2',SOFT),av={'adj':50000},
            paras=[para([run(f'{k+1}  ',sz,'lt1' if on else 'dk2',b=True),run(t,sz,'lt1' if on else 'dk2',b=on)],algn='ctr')],anchor='ctr'))
    return S

def chevron(cx,cy,w=.12,h=.16,c='accent1'):
    return shape('Flow chevron',cx-w/2,cy-h/2,w,h,'chevron',fill=sf(c),av={'adj':50000})

def progress_meter(n_done,total,cx=5.0,y=4.15,on_color='lt1',label=True):
    """Agenda/section progress meter (pills + '0N / 0M')."""
    pw,g=.5,.08; tw=total*pw+(total-1)*g; x0=cx-tw/2; S=[]
    for j in range(total):
        a=100000 if j<n_done else 40000
        S.append(shape(f'Progress {j+1}',x0+j*(pw+g),y,pw,.08,'roundRect',fill=sf(on_color,f'<a:alpha val="{a}"/>'),av={'adj':50000}))
    if label: S.append(text('Progress label',cx-1,y+.16,2,.22,[para([run(f'{n_done:02d} / {total:02d}',10,on_color,b=True,spc=100)],algn='ctr')]))
    return S

def stage_card(x=1.5,y=.5,w=7.0,h=1.55,fill='lt1'):
    """White rounded card with soft shadow, used to hold illustrations on busy section-divider backgrounds."""
    return shadow(shape('Illustration card',x,y,w,h,'roundRect',fill=sf(fill),av={'adj':14000}),alpha=20000,blur=101600,dist=25400)

def inject(slide_xml_path, shapes):
    """Replace everything inside spTree (after grpSpPr) of a slide with `shapes`."""
    x=open(slide_xml_path,encoding='utf-8').read()
    x,n=re.subn(r'(</p:grpSpPr>|<p:grpSpPr/>).*(</p:spTree>)',lambda m:m.group(1)+''.join(shapes)+m.group(2),x,count=1,flags=re.S)
    if n==0: raise RuntimeError(f'inject: no <p:grpSpPr>…</p:spTree> found in {slide_xml_path}; shapes NOT injected')
    open(slide_xml_path,'w',encoding='utf-8').write(x)
