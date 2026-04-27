/* OSAMA DESIGN TOOLS */

Office.onReady(() => {
  renderSVGLibrary();
});

function getShapes(context) {
  return context.presentation.getSelectedShapes();
}

function showStatus(msg, type) {
  const el = document.getElementById('status');
  if (!el) return;
  el.textContent = msg;
  el.className = type;
  setTimeout(() => { el.className = ''; }, 3000);
}

// --- CORNER RADIUS ---
async function applyCornerRadius(radiusPt) {
  const radius = parseFloat(radiusPt);
  if (isNaN(radius)) return;
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/width,items/height,items/adjustments,items/type");
    await context.sync();
    shapes.items.forEach(shape => {
      try {
        const minSide = Math.min(shape.width, shape.height);
        if (minSide > 0) shape.adjustments.set(0, Math.min(2 * radius / minSide, 0.5));
      } catch (e) {}
    });
    await context.sync();
  });
}

async function applyCornerRadiusToAll(radiusPt) {
  const radius = parseFloat(radiusPt);
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/width,items/height,items/adjustments");
    await context.sync();
    shapes.items.forEach(shape => {
      try {
        const minSide = Math.min(shape.width, shape.height);
        if (minSide > 0) shape.adjustments.set(0, Math.min(2 * radius / minSide, 0.5));
      } catch (e) {}
    });
    await context.sync();
  });
}

// --- FILL & OPACITY ---
async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch (e) {} });
    await context.sync();
  });
}

function onFillColorInput(hex) {
  const field = document.getElementById('fillHex');
  if (field) field.value = hex.toUpperCase();
  applyFillColor(hex);
}

function syncFillHexInput() {
  let hex = document.getElementById('fillHex').value.trim();
  if (!hex.startsWith('#')) hex = '#' + hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    const picker = document.getElementById('fillColor');
    if (picker) picker.value = hex;
    applyFillColor(hex);
  }
}

async function applyNoFill() {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.transparency = 1; } catch (e) {} });
    await context.sync();
  });
}

async function applyOpacity(val) {
  const trans = 1 - (parseFloat(val) / 100);
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.transparency = trans; } catch (e) {} });
    await context.sync();
  });
}

// --- BORDER ---
async function applyBorderColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/lineFormat");
    await context.sync();
    shapes.items.forEach(s => {
      try { s.lineFormat.color = hex; s.lineFormat.visible = true; } catch (e) {}
    });
    await context.sync();
  });
}

function syncBorderHexInput() {
  let hex = document.getElementById('borderHex').value.trim();
  if (!hex.startsWith('#')) hex = '#' + hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    const picker = document.getElementById('borderColor');
    if (picker) picker.value = hex;
    applyBorderColor(hex);
  }
}

async function applyBorderWidth(val) {
  const pt = parseFloat(val);
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/lineFormat");
    await context.sync();
    shapes.items.forEach(s => {
      try {
        s.lineFormat.visible = pt > 0;
        if (pt > 0) s.lineFormat.weight = pt;
      } catch (e) {}
    });
    await context.sync();
  });
}

// --- CONVERSION ---
async function convertToRoundRect() {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
    await context.sync();
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    for (const s of shapes.items) {
      const L = s.left, T = s.top, W = s.width, H = s.height;
      let fill = "#4472C4";
      try { fill = s.fill.foregroundColor || fill; } catch(e){}
      s.delete();
      await context.sync();
      const ns = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.roundRectangle,
        { left: L, top: T, width: W, height: H }
      );
      ns.fill.setSolidColor(fill);
      await context.sync();
    }
    showStatus('✓ Converted to Round Rectangle', 'ok');
  });
}

// --- SVG INSERT ---
// All _svg* helpers below are required by insertSVGCode — do not remove them.

function _svgDims(el) {
  let w = parseFloat(el.getAttribute('width')), h = parseFloat(el.getAttribute('height'));
  const vb = el.getAttribute('viewBox');
  if (vb) { const p = vb.trim().split(/[\s,]+/); w = w||parseFloat(p[2]); h = h||parseFloat(p[3]); }
  return { w: w||200, h: h||200 };
}

function _svgColor(c) {
  if (!c || c==='none'||c==='transparent') return null;
  c = c.trim();
  if (c==='currentColor'||c==='currentcolor') return '000000';
  if (c[0]==='#') { let h=c.slice(1); if(h.length===3) h=h[0]+h[0]+h[1]+h[1]+h[2]+h[2]; return h.toLowerCase().padStart(6,'0'); }
  const m=c.match(/rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m) return [m[1],m[2],m[3]].map(n=>parseInt(n).toString(16).padStart(2,'0')).join('');
  return {black:'000000',white:'ffffff',red:'ff0000',green:'008000',blue:'0000ff',yellow:'ffff00',
    orange:'ffa500',purple:'800080',gray:'808080',grey:'808080',pink:'ffc0cb',cyan:'00ffff',
    lime:'00ff00',navy:'000080',teal:'008080',silver:'c0c0c0',brown:'a52a2a',maroon:'800000'}[c.toLowerCase()]||'000000';
}

function _svgProp(el, prop, inh) {
  const sm = new RegExp('(?:^|;)\\s*'+prop+'\\s*:\\s*([^;]+)').exec(el.getAttribute('style')||'');
  const v = (sm?sm[1].trim():null)||el.getAttribute(prop);
  return (v && v!=='inherit') ? v : (inh&&inh[prop])||null;
}

function _svgParsePath(d) {
  const toks=[]; let i=0;
  d.replace(/([MmZzLlHhVvCcSsQqTtAa])|([+-]?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?)/g,
    (_,c,n)=>{ if(c) toks.push({t:'c',v:c}); else if(n!=null) toks.push({t:'n',v:parseFloat(n)}); });
  const out=[];
  const n=()=>(i<toks.length&&toks[i].t==='n')?toks[i++].v:0;
  const hn=()=>i<toks.length&&toks[i].t==='n';
  while(i<toks.length){
    if(toks[i].t!=='c'){i++;continue;}
    const c=toks[i++].v;
    if(c==='Z'||c==='z'){out.push({c:'Z'});continue;}
    do{switch(c){
      case 'M':case 'm':out.push({c,a:[n(),n()]});break;
      case 'L':case 'l':out.push({c,a:[n(),n()]});break;
      case 'H':case 'h':out.push({c,a:[n()]});break;
      case 'V':case 'v':out.push({c,a:[n()]});break;
      case 'C':case 'c':out.push({c,a:[n(),n(),n(),n(),n(),n()]});break;
      case 'S':case 's':out.push({c,a:[n(),n(),n(),n()]});break;
      case 'Q':case 'q':out.push({c,a:[n(),n(),n(),n()]});break;
      case 'T':case 't':out.push({c,a:[n(),n()]});break;
      case 'A':case 'a':out.push({c,a:[n(),n(),n(),n(),n(),n(),n()]});break;
      default:if(hn())n();
    }}while(hn());
  }
  return out;
}

function _svgArc2Bez(x1,y1,rx,ry,rot,laf,sf,x2,y2) {
  if(rx<=0||ry<=0) return [[x1,y1,x2,y2,x2,y2]];
  const φ=rot*Math.PI/180, cφ=Math.cos(φ), sφ=Math.sin(φ);
  const dx=(x1-x2)/2, dy=(y1-y2)/2, xp=cφ*dx+sφ*dy, yp=-sφ*dx+cφ*dy;
  let rxs=rx*rx, rys=ry*ry; const xps=xp*xp, yps=yp*yp;
  const λ=xps/rxs+yps/rys; if(λ>1){const s=Math.sqrt(λ);rx*=s;ry*=s;rxs=rx*rx;rys=ry*ry;}
  let sq=Math.sqrt(Math.max(0,(rxs*rys-rxs*yps-rys*xps)/Math.max(rxs*yps+rys*xps,1e-10)));
  if(laf===sf) sq=-sq;
  const cxp=sq*rx*yp/ry, cyp=-sq*ry*xp/rx;
  const cx=cφ*cxp-sφ*cyp+(x1+x2)/2, cy=sφ*cxp+cφ*cyp+(y1+y2)/2;
  const ang=(ux,uy,vx,vy)=>{let a=Math.acos(Math.max(-1,Math.min(1,(ux*vx+uy*vy)/Math.sqrt((ux*ux+uy*uy)*(vx*vx+vy*vy)))));if(ux*vy-uy*vx<0)a=-a;return a;};
  let t1=ang(1,0,(xp-cxp)/rx,(yp-cyp)/ry), dt=ang((xp-cxp)/rx,(yp-cyp)/ry,(-xp-cxp)/rx,(-yp-cyp)/ry);
  if(!sf&&dt>0)dt-=2*Math.PI; if(sf&&dt<0)dt+=2*Math.PI;
  const ns=Math.ceil(Math.abs(dt)/(Math.PI/2)), dts=dt/ns, out=[];
  for(let k=0;k<ns;k++){
    const a1=t1+k*dts, a2=t1+(k+1)*dts;
    const α=Math.sin(dts)*(Math.sqrt(4+3*Math.pow(Math.tan(dts/2),2))-1)/3;
    const [c1,s1,c2,s2]=[Math.cos(a1),Math.sin(a1),Math.cos(a2),Math.sin(a2)];
    const ex=cx+rx*cφ*c2-ry*sφ*s2, ey=cy+rx*sφ*c2+ry*cφ*s2;
    out.push([cx+rx*cφ*c1-ry*sφ*s1+α*(-(rx*cφ*s1+ry*sφ*c1)),
              cy+rx*sφ*c1+ry*cφ*s1+α*(-(rx*sφ*s1-ry*cφ*c1)),
              ex-α*(-(rx*cφ*s2+ry*sφ*c2)), ey-α*(-(rx*sφ*s2-ry*cφ*c2)), ex, ey]);
  }
  return out;
}

function _svgNorm(raw) {
  const R=v=>Math.round(v*100)/100, out=[];
  let px=0,py=0,sx=0,sy=0,cpx=null,cpy=null;
  for(const {c,a=[]} of raw){
    switch(c){
      case 'M':px=a[0];py=a[1];sx=px;sy=py;out.push({c:'M',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'm':px+=a[0];py+=a[1];sx=px;sy=py;out.push({c:'M',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'L':px=a[0];py=a[1];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'l':px+=a[0];py+=a[1];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'H':px=a[0];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'h':px+=a[0];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'V':py=a[0];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'v':py+=a[0];out.push({c:'L',x:R(px),y:R(py)});cpx=cpy=null;break;
      case 'C':{const[x1,y1,x2,y2,x,y]=a;out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});cpx=x2;cpy=y2;px=x;py=y;break;}
      case 'c':{const[a1,b1,a2,b2,dx,dy]=a;const[x1,y1,x2,y2,x,y]=[px+a1,py+b1,px+a2,py+b2,px+dx,py+dy];out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});cpx=x2;cpy=y2;px=x;py=y;break;}
      case 'S':{const[x2,y2,x,y]=a;const x1=cpx!=null?2*px-cpx:px,y1=cpy!=null?2*py-cpy:py;out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});cpx=x2;cpy=y2;px=x;py=y;break;}
      case 's':{const[dx2,dy2,dx,dy]=a;const x2=px+dx2,y2=py+dy2,x=px+dx,y=py+dy;const x1=cpx!=null?2*px-cpx:px,y1=cpy!=null?2*py-cpy:py;out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});cpx=x2;cpy=y2;px=x;py=y;break;}
      case 'Q':{const[qx,qy,x,y]=a;out.push({c:'C',x1:R(px+2/3*(qx-px)),y1:R(py+2/3*(qy-py)),x2:R(x+2/3*(qx-x)),y2:R(y+2/3*(qy-y)),x:R(x),y:R(y)});cpx=qx;cpy=qy;px=x;py=y;break;}
      case 'q':{const[dqx,dqy,dx,dy]=a;const qx=px+dqx,qy=py+dqy,x=px+dx,y=py+dy;out.push({c:'C',x1:R(px+2/3*(qx-px)),y1:R(py+2/3*(qy-py)),x2:R(x+2/3*(qx-x)),y2:R(y+2/3*(qy-y)),x:R(x),y:R(y)});cpx=qx;cpy=qy;px=x;py=y;break;}
      case 'T':{const[x,y]=a;const qx=cpx!=null?2*px-cpx:px,qy=cpy!=null?2*py-cpy:py;out.push({c:'C',x1:R(px+2/3*(qx-px)),y1:R(py+2/3*(qy-py)),x2:R(x+2/3*(qx-x)),y2:R(y+2/3*(qy-y)),x:R(x),y:R(y)});cpx=qx;cpy=qy;px=x;py=y;break;}
      case 't':{const[dx,dy]=a;const x=px+dx,y=py+dy,qx=cpx!=null?2*px-cpx:px,qy=cpy!=null?2*py-cpy:py;out.push({c:'C',x1:R(px+2/3*(qx-px)),y1:R(py+2/3*(qy-py)),x2:R(x+2/3*(qx-x)),y2:R(y+2/3*(qy-y)),x:R(x),y:R(y)});cpx=qx;cpy=qy;px=x;py=y;break;}
      case 'A':{const[rx,ry,r,laf,sf,x,y]=a;_svgArc2Bez(px,py,rx,ry,r,laf,sf,x,y).forEach(([x1,y1,x2,y2,ex,ey])=>out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(ex),y:R(ey)}));cpx=cpy=null;px=x;py=y;break;}
      case 'a':{const[rx,ry,r,laf,sf,dx,dy]=a;const x=px+dx,y=py+dy;_svgArc2Bez(px,py,rx,ry,r,laf,sf,x,y).forEach(([x1,y1,x2,y2,ex,ey])=>out.push({c:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(ex),y:R(ey)}));cpx=cpy=null;px=x;py=y;break;}
      case 'Z':out.push({c:'Z'});px=sx;py=sy;cpx=cpy=null;break;
    }
  }
  return out;
}

function _svgCmdsToXML(cmds) {
  return cmds.map(o=>{
    const r=v=>Math.round(v);
    if(o.c==='M') return `<a:moveTo><a:pt x="${r(o.x)}" y="${r(o.y)}"/></a:moveTo>`;
    if(o.c==='L') return `<a:lnTo><a:pt x="${r(o.x)}" y="${r(o.y)}"/></a:lnTo>`;
    if(o.c==='C') return `<a:cubicBezTo><a:pt x="${r(o.x1)}" y="${r(o.y1)}"/><a:pt x="${r(o.x2)}" y="${r(o.y2)}"/><a:pt x="${r(o.x)}" y="${r(o.y)}"/></a:cubicBezTo>`;
    if(o.c==='Z') return `<a:close/>`;
    return '';
  }).join('');
}

function _svgTfParse(t) {
  if(!t) return {a:1,b:0,c:0,d:1,e:0,f:0};
  const mx=t.match(/matrix\(\s*([^,\s)]+)[,\s]+([^,\s)]+)[,\s]+([^,\s)]+)[,\s]+([^,\s)]+)[,\s]+([^,\s)]+)[,\s]+([^,\s)]+)\s*\)/);
  if(mx) return {a:+mx[1],b:+mx[2],c:+mx[3],d:+mx[4],e:+mx[5],f:+mx[6]};
  let a=1,b=0,cc=0,d=1,e=0,f=0;
  const tr=t.match(/translate\(\s*([^,)\s]+)(?:[,\s]+([^)\s]+))?\)/);
  if(tr){e=+tr[1]||0;f=+tr[2]||0;}
  const sc=t.match(/scale\(\s*([^,)\s]+)(?:[,\s]+([^)\s]+))?\)/);
  if(sc){a=+sc[1]||1;d=+(sc[2]||sc[1])||1;}
  const ro=t.match(/rotate\(\s*([^,)\s]+)(?:[,\s]+([^,)\s]+)[,\s]+([^)\s]+))?\)/);
  if(ro){const θ=(+ro[1])*Math.PI/180,cx2=+ro[2]||0,cy2=+ro[3]||0;a=Math.cos(θ);b=Math.sin(θ);cc=-Math.sin(θ);d=Math.cos(θ);e=cx2-cx2*a+cy2*Math.sin(θ);f=cy2-cx2*Math.sin(θ)-cy2*Math.cos(θ);}
  return {a,b,c:cc,d,e,f};
}

function _svgTfPt(m,x,y){return{x:m.a*x+m.c*y+m.e, y:m.b*x+m.d*y+m.f};}

function _svgTfMul(p,q){return{a:p.a*q.a+p.c*q.b, b:p.b*q.a+p.d*q.b, c:p.a*q.c+p.c*q.d, d:p.b*q.c+p.d*q.d, e:p.a*q.e+p.c*q.f+p.e, f:p.b*q.e+p.d*q.f+p.f};}

function _svgCollect(el, inh, tf) {
  const shapes=[];
  const tag=(el.tagName||'').toLowerCase().replace(/^.*:/,'');
  const fill=_svgProp(el,'fill',inh), stroke=_svgProp(el,'stroke',inh), sw=_svgProp(el,'stroke-width',inh)||'1';
  const ci={fill,stroke,'stroke-width':sw};
  const myTf=_svgTfParse(el.getAttribute&&el.getAttribute('transform'));
  const m=_svgTfMul(tf,myTf);

  if(tag==='g'||tag==='svg'){for(const ch of (el.children||[]))shapes.push(..._svgCollect(ch,ci,m));return shapes;}

  const fc=_svgColor(fill), sc=_svgColor(stroke), swn=Math.max(1,Math.round(parseFloat(sw)*9525));

  if(tag==='path'){
    const d=el.getAttribute('d');if(!d)return shapes;
    const cmds=_svgNorm(_svgParsePath(d)).map(o=>{
      if(o.c==='M'||o.c==='L'){const p=_svgTfPt(m,o.x,o.y);return{...o,x:p.x,y:p.y};}
      if(o.c==='C'){const p1=_svgTfPt(m,o.x1,o.y1),p2=_svgTfPt(m,o.x2,o.y2),p=_svgTfPt(m,o.x,o.y);return{...o,x1:p1.x,y1:p1.y,x2:p2.x,y2:p2.y,x:p.x,y:p.y};}
      return o;
    });
    if(cmds.length)shapes.push({t:'path',cmds,fill:fc,stroke:sc,sw:swn});
    return shapes;
  }
  if(tag==='rect'){
    const x=+el.getAttribute('x')||0,y=+el.getAttribute('y')||0,w=+el.getAttribute('width')||0,h=+el.getAttribute('height')||0,rx=+el.getAttribute('rx')||0;
    const p1=_svgTfPt(m,x,y),p2=_svgTfPt(m,x+w,y+h);
    shapes.push({t:'rect',x:p1.x,y:p1.y,w:p2.x-p1.x,h:p2.y-p1.y,rx,fill:fc,stroke:sc,sw:swn});
    return shapes;
  }
  if(tag==='circle'){
    const cx=+el.getAttribute('cx')||0,cy=+el.getAttribute('cy')||0,r=+el.getAttribute('r')||0;
    const p=_svgTfPt(m,cx-r,cy-r);shapes.push({t:'ellipse',x:p.x,y:p.y,w:r*2,h:r*2,fill:fc,stroke:sc,sw:swn});return shapes;
  }
  if(tag==='ellipse'){
    const cx=+el.getAttribute('cx')||0,cy=+el.getAttribute('cy')||0,rx2=+el.getAttribute('rx')||0,ry2=+el.getAttribute('ry')||0;
    const p=_svgTfPt(m,cx-rx2,cy-ry2);shapes.push({t:'ellipse',x:p.x,y:p.y,w:rx2*2,h:ry2*2,fill:fc,stroke:sc,sw:swn});return shapes;
  }
  if(tag==='polygon'||tag==='polyline'){
    const pts=(el.getAttribute('points')||'').trim().split(/[\s,]+/).map(Number).filter(n=>!isNaN(n));
    if(pts.length<2)return shapes;
    const cmds=[];
    for(let k=0;k<pts.length-1;k+=2){const p=_svgTfPt(m,pts[k],pts[k+1]);cmds.push({c:k===0?'M':'L',x:p.x,y:p.y});}
    if(tag==='polygon')cmds.push({c:'Z'});
    shapes.push({t:'path',cmds,fill:fc,stroke:sc,sw:swn});return shapes;
  }
  if(tag==='line'){
    const p1=_svgTfPt(m,+el.getAttribute('x1')||0,+el.getAttribute('y1')||0);
    const p2=_svgTfPt(m,+el.getAttribute('x2')||0,+el.getAttribute('y2')||0);
    shapes.push({t:'path',cmds:[{c:'M',x:p1.x,y:p1.y},{c:'L',x:p2.x,y:p2.y}],fill:null,stroke:sc||'000000',sw:swn});return shapes;
  }
  return shapes;
}

function _svgToSpXML(shapes,vbW,vbH,cx,cy) {
  let xml='',id=10;
  const eu=v=>Math.round(v/vbW*cx), ev=v=>Math.round(v/vbH*cy);
  for(const s of shapes){
    const fxml=s.fill?`<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>`:`<a:noFill/>`;
    const lxml=s.stroke?`<a:ln w="${s.sw}"><a:solidFill><a:srgbClr val="${s.stroke}"/></a:solidFill></a:ln>`:`<a:ln><a:noFill/></a:ln>`;
    if(s.t==='path'){
      const pxml=_svgCmdsToXML(s.cmds);if(!pxml)continue;
      xml+=`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id++}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
        <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>
        <a:pathLst><a:path w="${Math.round(vbW)}" h="${Math.round(vbH)}">${pxml}</a:path></a:pathLst></a:custGeom>
        ${fxml}${lxml}</p:spPr></p:sp>`;
    } else if(s.t==='rect'){
      const hasR=s.rx>0,adj=hasR?Math.round(Math.min(s.rx/Math.max(s.w,s.h,1),0.5)*50000):0;
      const geom=hasR?`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>`:`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
      xml+=`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="R${id++}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="${eu(s.x)}" y="${ev(s.y)}"/><a:ext cx="${eu(s.w)}" cy="${ev(s.h)}"/></a:xfrm>
        ${geom}${fxml}${lxml}</p:spPr></p:sp>`;
    } else if(s.t==='ellipse'){
      xml+=`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="E${id++}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="${eu(s.x)}" y="${ev(s.y)}"/><a:ext cx="${eu(s.w)}" cy="${ev(s.h)}"/></a:xfrm>
        <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${fxml}${lxml}</p:spPr></p:sp>`;
    }
  }
  return xml;
}

function _buildOoxml(spTree) {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/ppt/presentation.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"><pkg:xmlData>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldMasterIdLst/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/ppt/_rels/presentation.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/ppt/slides/slide1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"><pkg:xmlData>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${spTree}
</p:spTree></p:cSld></p:sld></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/ppt/slides/_rels/slide1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/></pkg:xmlData></pkg:part>
</pkg:package>`;
}

// --- MAIN INSERT FUNCTION ---
// Strategy:
//   1. Try PowerPoint.run API to insert native shapes (works on Windows)
//   2. If that fails, try Office.CoercionType.XmlSvg as fallback (image, works on Mac)
async function insertSVGCode(svgCode) {
  if (!svgCode || !svgCode.includes('<svg')) {
    showStatus('Paste valid SVG first', 'err');
    return;
  }

  // Ensure SVG has width/height attributes
  let svg = svgCode.trim();
  if (!svg.match(/\swidth\s*=/i)) {
    const vbM = svg.match(/viewBox=["']\s*[\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)/);
    const w = vbM ? vbM[1] : '200';
    const h = vbM ? vbM[2] : '200';
    svg = svg.replace(/<svg/, `<svg width="${w}" height="${h}"`);
  }

  // --- Method 1: Native shapes via PowerPoint JS API ---
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svg, 'image/svg+xml');
    const svgEl = doc.querySelector('svg');

    if (svgEl) {
      const { w: vbW, h: vbH } = _svgDims(svgEl);
      // EMU: 1pt = 12700, slide default 6in wide = 5486400 EMU
      const scale = 5486400 / Math.max(vbW, vbH, 1);
      const cxEmu = Math.round(vbW * scale);
      const cyEmu = Math.round(vbH * scale);
      const ident = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
      const shapes = _svgCollect(svgEl, {}, ident);

      if (shapes.length > 0) {
        const spTree = _svgToSpXML(shapes, vbW, vbH, cxEmu, cyEmu);
        const ooxml = _buildOoxml(spTree);

        // Try to insert via OOXML coercion
        try {
          await new Promise((resolve, reject) => {
            Office.context.document.setSelectedDataAsync(
              ooxml,
              { coercionType: Office.CoercionType.Ooxml },
              (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  resolve();
                } else {
                  reject(new Error(result.error ? result.error.message : 'OOXML failed'));
                }
              }
            );
          });

          showStatus('✓ SVG inserted as native shapes!', 'ok');
          return;
        } catch (ooError) {
          console.warn('OOXML insert failed, trying XmlSvg fallback:', ooError.message);
          // Fall through to XmlSvg method
        }
      }
    }
  } catch (parseError) {
    console.warn('SVG parsing failed:', parseError.message);
  }

  // --- Method 2: XmlSvg fallback (Mac-compatible, inserts as image) ---
  try {
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        svg,
        { coercionType: Office.CoercionType.XmlSvg },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error ? result.error.message : 'XmlSvg failed'));
          }
        }
      );
    });
    showStatus('✓ SVG inserted as image!', 'ok');
  } catch (fallbackError) {
    showStatus('Error inserting SVG: ' + fallbackError.message, 'err');
  }
}

// --- IMPROVED INSERT WITH VALIDATION ---
async function insertSVGToSlide() {
  const code = document.getElementById('svg-input').value.trim();
  
  // Check for empty textarea
  if (!code) {
    showStatus('❌ Paste SVG code first', 'err');
    return;
  }
  
  // Check for required SVG tags
  if (!code.includes('<svg')) {
    showStatus('❌ Error: Missing <svg> tag', 'err');
    return;
  }
  
  if (!code.includes('</svg>')) {
    showStatus('❌ Error: Missing closing </svg> tag', 'err');
    return;
  }
  
  // Validate SVG syntax with DOMParser
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(code, 'image/svg+xml');
    
    // Check for parse errors
    if (doc.getElementsByTagName('parsererror').length > 0) {
      showStatus('❌ Error: SVG has syntax errors - check your code', 'err');
      console.error('SVG Parse Error:', doc.getElementsByTagName('parsererror')[0].textContent);
      return;
    }
  } catch (e) {
    showStatus('❌ Error: Invalid SVG format', 'err');
    console.error('SVG Validation Error:', e.message);
    return;
  }
  
  // If validation passes, insert
  showStatus('⏳ Inserting SVG...', 'ok');
  await insertSVGCode(code);
}

// --- SVG LIBRARY ---
function getSVGLibrary() {
  try { return JSON.parse(localStorage.getItem('svgLibrary') || '[]'); }
  catch (e) { return []; }
}

function saveSVGToLibrary() {
  const code = document.getElementById('svg-input').value.trim();
  if (!code) { showStatus('Paste SVG code first', 'err'); return; }
  const nameEl = document.getElementById('svgName');
  const name = (nameEl && nameEl.value.trim()) || ('Shape ' + Date.now());
  const lib = getSVGLibrary();
  lib.push({ id: Date.now(), name: name, code: code });
  localStorage.setItem('svgLibrary', JSON.stringify(lib));
  if (nameEl) nameEl.value = '';
  renderSVGLibrary();
  showStatus('✓ Saved to library!', 'ok');
}

function deleteSVGFromLibrary(id) {
  const lib = getSVGLibrary().filter(s => s.id !== id);
  localStorage.setItem('svgLibrary', JSON.stringify(lib));
  renderSVGLibrary();
}

async function insertSVGFromLibrary(id) {
  const item = getSVGLibrary().find(s => s.id === id);
  if (item) await insertSVGCode(item.code);
}

function renderSVGLibrary() {
  const grid = document.getElementById('svgLibraryGrid');
  if (!grid) return;
  const lib = getSVGLibrary();
  if (!lib.length) {
    grid.innerHTML = '<div style="font-size:10px;color:#aaa;padding:10px;text-align:center;grid-column:1/-1">Library is empty</div>';
    return;
  }
  grid.innerHTML = lib.map(s => `
    <div class="shape-item" onclick="insertSVGFromLibrary(${s.id})">
      <div class="preview">${s.code}</div>
      <span class="lbl">${s.name}</span>
      <button class="del-btn" onclick="event.stopPropagation();deleteSVGFromLibrary(${s.id})">×</button>
    </div>`).join('');
}

// --- BULK COLOR REPLACE ---
function normalizeHex(val) {
  if (!val) return '';
  return val.replace('#', '').toUpperCase().trim();
}

async function readColorFromSelection() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/foregroundColor,items/fill/type");
      await context.sync();
      if (!shapes.items.length) { showStatus('Select a shape first', 'err'); return; }
      const fc = shapes.items[0].fill.foregroundColor;
      if (!fc) { showStatus('Shape has no solid fill', 'err'); return; }
      const hex = '#' + normalizeHex(fc);
      document.getElementById('findColor').value = hex;
      showStatus('✓ Color read: ' + hex, 'ok');
    });
  } catch (e) { showStatus('Could not read color', 'err'); }
}

async function bulkColorReplace() {
  const find = normalizeHex(document.getElementById('findColor').value);
  const replace = document.getElementById('replaceColor').value;
  let count = 0;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/fill/foregroundColor,items/fill/type,items/lineFormat/color");
    await context.sync();
    shapes.items.forEach(s => {
      try {
        if (normalizeHex(s.fill.foregroundColor) === find) {
          s.fill.setSolidColor(replace);
          count++;
        }
      } catch (e) {}
      try {
        if (normalizeHex(s.lineFormat.color) === find) {
          s.lineFormat.color = replace;
          count++;
        }
      } catch (e) {}
    });
    await context.sync();
  });
  showStatus('✓ Replaced ' + count + ' item(s)', count > 0 ? 'ok' : 'err');
}

// --- EXPOSE FUNCTIONS ---
window.applyCornerRadius = applyCornerRadius;
window.applyCornerRadiusToAll = applyCornerRadiusToAll;
window.applyFillColor = applyFillColor;
window.onFillColorInput = onFillColorInput;
window.syncFillHexInput = syncFillHexInput;
window.applyNoFill = applyNoFill;
window.applyOpacity = applyOpacity;
window.applyBorderColor = applyBorderColor;
window.syncBorderHexInput = syncBorderHexInput;
window.applyBorderWidth = applyBorderWidth;
window.convertToRoundRect = convertToRoundRect;
window.insertSVGToSlide = insertSVGToSlide;
window.saveSVGToLibrary = saveSVGToLibrary;
window.deleteSVGFromLibrary = deleteSVGFromLibrary;
window.insertSVGFromLibrary = insertSVGFromLibrary;
window.readColorFromSelection = readColorFromSelection;
window.bulkColorReplace = bulkColorReplace;