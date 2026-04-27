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
      const L = s.left, T = s.top, W = s.width, H = s.height, F = s.fill.foregroundColor;
      s.delete();
      const ns = slide.shapes.addGeometricShape("RoundRectangle", { left: L, top: T, width: W, height: H });
      ns.fill.setSolidColor(F);
    }
    await context.sync();
  });
}

// ─── SVG → DrawingML shape converter ────────────────────────────────────────

function svgGetDimensions(svgEl) {
  let w = parseFloat(svgEl && svgEl.getAttribute('width'));
  let h = parseFloat(svgEl && svgEl.getAttribute('height'));
  if (!w || !h) {
    const vb = svgEl && svgEl.getAttribute('viewBox');
    if (vb) {
      const p = vb.trim().split(/[\s,]+/);
      if (p.length >= 4) { w = w || parseFloat(p[2]); h = h || parseFloat(p[3]); }
    }
  }
  return { w: w || 200, h: h || 200 };
}

function svgParseColor(c) {
  if (!c || c === 'none' || c === 'transparent') return null;
  c = c.trim();
  if (c.startsWith('#')) {
    let h = c.slice(1);
    if (h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
    return h.toLowerCase().padStart(6,'0');
  }
  if (c.startsWith('rgb')) {
    const m = c.match(/rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (m) return [m[1],m[2],m[3]].map(n=>parseInt(n).toString(16).padStart(2,'0')).join('');
  }
  const named={black:'000000',white:'ffffff',red:'ff0000',green:'008000',blue:'0000ff',
    yellow:'ffff00',orange:'ffa500',purple:'800080',gray:'808080',grey:'808080',
    pink:'ffc0cb',cyan:'00ffff',magenta:'ff00ff',lime:'00ff00',navy:'000080',
    teal:'008080',silver:'c0c0c0',brown:'a52a2a',maroon:'800000'};
  return named[c.toLowerCase()] || '000000';
}

function svgGetProp(el, prop, inherited) {
  const style = el.getAttribute('style') || '';
  const m = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([^;]+)').exec(style);
  const val = (m ? m[1].trim() : null) || el.getAttribute(prop);
  if (val && val !== 'inherit') return val;
  return (inherited && inherited[prop]) || null;
}

function svgParsePath(d) {
  const tokens = [];
  d.replace(/([MmZzLlHhVvCcSsQqTtAa])|([+-]?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?)/g,
    (_, cmd, num) => { if (cmd) tokens.push({t:'c',v:cmd}); else if (num!=null) tokens.push({t:'n',v:parseFloat(num)}); });
  const out = [];
  let i = 0;
  function n() { return (i<tokens.length && tokens[i].t==='n') ? tokens[i++].v : 0; }
  function hasN() { return i<tokens.length && tokens[i].t==='n'; }
  while (i < tokens.length) {
    if (tokens[i].t !== 'c') { i++; continue; }
    const cmd = tokens[i++].v;
    if (cmd==='Z'||cmd==='z') { out.push({cmd:'Z'}); continue; }
    do {
      switch(cmd) {
        case 'M': case 'm': out.push({cmd, a:[n(),n()]}); break;
        case 'L': case 'l': out.push({cmd, a:[n(),n()]}); break;
        case 'H': case 'h': out.push({cmd, a:[n()]}); break;
        case 'V': case 'v': out.push({cmd, a:[n()]}); break;
        case 'C': case 'c': out.push({cmd, a:[n(),n(),n(),n(),n(),n()]}); break;
        case 'S': case 's': out.push({cmd, a:[n(),n(),n(),n()]}); break;
        case 'Q': case 'q': out.push({cmd, a:[n(),n(),n(),n()]}); break;
        case 'T': case 't': out.push({cmd, a:[n(),n()]}); break;
        case 'A': case 'a': out.push({cmd, a:[n(),n(),n(),n(),n(),n(),n()]}); break;
        default: if(hasN()) n();
      }
    } while (hasN());
  }
  return out;
}

function svgArcToBeziers(x1,y1,rx,ry,xRot,laf,sf,x2,y2) {
  if (rx<=0||ry<=0) return [[x1,y1,x2,y2,x2,y2]];
  const phi=xRot*Math.PI/180, cp=Math.cos(phi), sp=Math.sin(phi);
  const dx=(x1-x2)/2, dy=(y1-y2)/2;
  const xp=cp*dx+sp*dy, yp=-sp*dx+cp*dy;
  let rxs=rx*rx, rys=ry*ry;
  const xps=xp*xp, yps=yp*yp;
  const lam=xps/rxs+yps/rys;
  if(lam>1){const s=Math.sqrt(lam);rx*=s;ry*=s;rxs=rx*rx;rys=ry*ry;}
  const num=Math.max(0,rxs*rys-rxs*yps-rys*xps);
  const den=rxs*yps+rys*xps;
  let sq=Math.sqrt(num/Math.max(den,1e-10));
  if(laf===sf) sq=-sq;
  const cxp=sq*rx*yp/ry, cyp=-sq*ry*xp/rx;
  const cx=cp*cxp-sp*cyp+(x1+x2)/2, cy=sp*cxp+cp*cyp+(y1+y2)/2;
  function ang(ux,uy,vx,vy){const n=Math.sqrt(ux*ux+uy*uy)*Math.sqrt(vx*vx+vy*vy);let a=Math.acos(Math.max(-1,Math.min(1,(ux*vx+uy*vy)/n)));if(ux*vy-uy*vx<0)a=-a;return a;}
  let t1=ang(1,0,(xp-cxp)/rx,(yp-cyp)/ry);
  let dt=ang((xp-cxp)/rx,(yp-cyp)/ry,(-xp-cxp)/rx,(-yp-cyp)/ry);
  if(!sf&&dt>0) dt-=2*Math.PI; if(sf&&dt<0) dt+=2*Math.PI;
  const ns=Math.ceil(Math.abs(dt)/(Math.PI/2)), dts=dt/ns, bezs=[];
  for(let k=0;k<ns;k++){
    const a1=t1+k*dts, a2=t1+(k+1)*dts;
    const al=Math.sin(dts)*(Math.sqrt(4+3*Math.pow(Math.tan(dts/2),2))-1)/3;
    const c1=Math.cos(a1),s1=Math.sin(a1),c2=Math.cos(a2),s2=Math.sin(a2);
    const d1x=-(rx*cp*s1+ry*sp*c1), d1y=-(rx*sp*s1-ry*cp*c1);
    const d2x=-(rx*cp*s2+ry*sp*c2), d2y=-(rx*sp*s2-ry*cp*c2);
    const ex=cx+rx*cp*c2-ry*sp*s2, ey=cy+rx*sp*c2+ry*cp*s2;
    bezs.push([cx+rx*cp*c1-ry*sp*s1+al*d1x, cy+rx*sp*c1+ry*cp*s1+al*d1y,
               ex-al*d2x, ey-al*d2y, ex, ey]);
  }
  return bezs;
}

function svgNormalizePath(raw) {
  const out=[], R=n=>Math.round(n*100)/100;
  let cx=0,cy=0,sx=0,sy=0,pcpx=null,pcpy=null;
  for(const {cmd,a=[]} of raw){
    switch(cmd){
      case 'M': cx=a[0];cy=a[1];sx=cx;sy=cy;out.push({cmd:'M',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'm': cx+=a[0];cy+=a[1];sx=cx;sy=cy;out.push({cmd:'M',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'L': cx=a[0];cy=a[1];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'l': cx+=a[0];cy+=a[1];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'H': cx=a[0];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'h': cx+=a[0];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'V': cy=a[0];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'v': cy+=a[0];out.push({cmd:'L',x:R(cx),y:R(cy)});pcpx=pcpy=null;break;
      case 'C':{const[x1,y1,x2,y2,x,y]=a;out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=x2;pcpy=y2;cx=x;cy=y;break;}
      case 'c':{const[a1,b1,a2,b2,dx,dy]=a;const x1=cx+a1,y1=cy+b1,x2=cx+a2,y2=cy+b2,x=cx+dx,y=cy+dy;out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=x2;pcpy=y2;cx=x;cy=y;break;}
      case 'S':{const[x2,y2,x,y]=a;const x1=pcpx!=null?2*cx-pcpx:cx,y1=pcpy!=null?2*cy-pcpy:cy;out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=x2;pcpy=y2;cx=x;cy=y;break;}
      case 's':{const[dx2,dy2,dx,dy]=a;const x2=cx+dx2,y2=cy+dy2,x=cx+dx,y=cy+dy;const x1=pcpx!=null?2*cx-pcpx:cx,y1=pcpy!=null?2*cy-pcpy:cy;out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=x2;pcpy=y2;cx=x;cy=y;break;}
      case 'Q':{const[qx1,qy1,x,y]=a;const x1=cx+2/3*(qx1-cx),y1=cy+2/3*(qy1-cy),x2=x+2/3*(qx1-x),y2=y+2/3*(qy1-y);out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=qx1;pcpy=qy1;cx=x;cy=y;break;}
      case 'q':{const[dqx,dqy,dx,dy]=a;const qx1=cx+dqx,qy1=cy+dqy,x=cx+dx,y=cy+dy;const x1=cx+2/3*(qx1-cx),y1=cy+2/3*(qy1-cy),x2=x+2/3*(qx1-x),y2=y+2/3*(qy1-y);out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=qx1;pcpy=qy1;cx=x;cy=y;break;}
      case 'T':{const[x,y]=a;const qx1=pcpx!=null?2*cx-pcpx:cx,qy1=pcpy!=null?2*cy-pcpy:cy;const x1=cx+2/3*(qx1-cx),y1=cy+2/3*(qy1-cy),x2=x+2/3*(qx1-x),y2=y+2/3*(qy1-y);out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=qx1;pcpy=qy1;cx=x;cy=y;break;}
      case 't':{const[dx,dy]=a;const x=cx+dx,y=cy+dy;const qx1=pcpx!=null?2*cx-pcpx:cx,qy1=pcpy!=null?2*cy-pcpy:cy;const x1=cx+2/3*(qx1-cx),y1=cy+2/3*(qy1-cy),x2=x+2/3*(qx1-x),y2=y+2/3*(qy1-y);out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(x),y:R(y)});pcpx=qx1;pcpy=qy1;cx=x;cy=y;break;}
      case 'A':{const[rx,ry,rot,laf,sf,x,y]=a;svgArcToBeziers(cx,cy,rx,ry,rot,laf,sf,x,y).forEach(([x1,y1,x2,y2,ex,ey])=>out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(ex),y:R(ey)}));pcpx=pcpy=null;cx=x;cy=y;break;}
      case 'a':{const[rx,ry,rot,laf,sf,dx,dy]=a;const x=cx+dx,y=cy+dy;svgArcToBeziers(cx,cy,rx,ry,rot,laf,sf,x,y).forEach(([x1,y1,x2,y2,ex,ey])=>out.push({cmd:'C',x1:R(x1),y1:R(y1),x2:R(x2),y2:R(y2),x:R(ex),y:R(ey)}));pcpx=pcpy=null;cx=x;cy=y;break;}
      case 'Z': out.push({cmd:'Z'});cx=sx;cy=sy;pcpx=pcpy=null;break;
    }
  }
  return out;
}

function svgNormCmdsToXML(cmds) {
  return cmds.map(c=>{
    if(c.cmd==='M') return `<a:moveTo><a:pt x="${Math.round(c.x)}" y="${Math.round(c.y)}"/></a:moveTo>`;
    if(c.cmd==='L') return `<a:lnTo><a:pt x="${Math.round(c.x)}" y="${Math.round(c.y)}"/></a:lnTo>`;
    if(c.cmd==='C') return `<a:cubicBezTo><a:pt x="${Math.round(c.x1)}" y="${Math.round(c.y1)}"/><a:pt x="${Math.round(c.x2)}" y="${Math.round(c.y2)}"/><a:pt x="${Math.round(c.x)}" y="${Math.round(c.y)}"/></a:cubicBezTo>`;
    if(c.cmd==='Z') return `<a:close/>`;
    return '';
  }).join('');
}

function svgParseTransform(t) {
  if (!t) return {tx:0,ty:0,sx:1,sy:1};
  let tx=0,ty=0,sx=1,sy=1;
  const tr=t.match(/translate\(\s*([^,)\s]+)(?:[,\s]+([^)\s]+))?\)/);
  if(tr){tx=parseFloat(tr[1])||0;ty=parseFloat(tr[2])||0;}
  const sc=t.match(/scale\(\s*([^,)\s]+)(?:[,\s]+([^)\s]+))?\)/);
  if(sc){sx=parseFloat(sc[1])||1;sy=parseFloat(sc[2]||sc[1])||1;}
  const mx=t.match(/matrix\(\s*([^,\s]+)[,\s]+([^,\s]+)[,\s]+([^,\s]+)[,\s]+([^,\s]+)[,\s]+([^,\s]+)[,\s]+([^)\s]+)\)/);
  if(mx){sx=parseFloat(mx[1]);sy=parseFloat(mx[4]);tx=parseFloat(mx[5]);ty=parseFloat(mx[6]);}
  return {tx,ty,sx,sy};
}

function svgApplyTf(tf, x, y) { return {x: tf.tx + tf.sx*x, y: tf.ty + tf.sy*y}; }
function svgComposeTf(parent, child) {
  return {tx:parent.tx+parent.sx*child.tx, ty:parent.ty+parent.sy*child.ty, sx:parent.sx*child.sx, sy:parent.sy*child.sy};
}

function svgCollectShapes(el, inh, tf) {
  const shapes = [];
  const tag = el.tagName && el.tagName.toLowerCase().replace(/^.*:/, '');
  const fill = svgGetProp(el, 'fill', inh);
  const stroke = svgGetProp(el, 'stroke', inh);
  const sw = svgGetProp(el, 'stroke-width', inh) || '1';
  const childInh = {fill, stroke, 'stroke-width': sw};
  const myTf = svgParseTransform(el.getAttribute && el.getAttribute('transform'));
  const accTf = svgComposeTf(tf, myTf);

  if (tag==='g' || tag==='svg') {
    for (const ch of (el.children||[])) shapes.push(...svgCollectShapes(ch, childInh, accTf));
    return shapes;
  }

  const fillColor = svgParseColor(fill);
  const strokeColor = svgParseColor(stroke);
  const strokeW = Math.max(1, Math.round(parseFloat(sw) * Math.max(accTf.sx, accTf.sy) * 9525));

  function tf2(x,y){return svgApplyTf(accTf,x,y);}

  if (tag==='path') {
    const d = el.getAttribute('d'); if(!d) return shapes;
    const raw = svgParsePath(d);
    const norm = svgNormalizePath(raw);
    const tfd = norm.map(c=>{
      if(c.cmd==='M'||c.cmd==='L'){const p=tf2(c.x,c.y);return{...c,x:p.x,y:p.y};}
      if(c.cmd==='C'){const p1=tf2(c.x1,c.y1),p2=tf2(c.x2,c.y2),p=tf2(c.x,c.y);return{...c,x1:p1.x,y1:p1.y,x2:p2.x,y2:p2.y,x:p.x,y:p.y};}
      return c;
    });
    if(tfd.length) shapes.push({type:'path',cmds:tfd,fill:fillColor,stroke:strokeColor,sw:strokeW});
    return shapes;
  }

  if (tag==='rect') {
    let x=parseFloat(el.getAttribute('x'))||0, y=parseFloat(el.getAttribute('y'))||0;
    let w=parseFloat(el.getAttribute('width'))||0, h=parseFloat(el.getAttribute('height'))||0;
    const rx=parseFloat(el.getAttribute('rx'))||0;
    const p1=tf2(x,y), p2=tf2(x+w,y+h);
    shapes.push({type:'rect',x:p1.x,y:p1.y,w:p2.x-p1.x,h:p2.y-p1.y,rx:rx*accTf.sx,fill:fillColor,stroke:strokeColor,sw:strokeW});
    return shapes;
  }

  if (tag==='circle') {
    const cx=parseFloat(el.getAttribute('cx'))||0, cy=parseFloat(el.getAttribute('cy'))||0;
    const r=parseFloat(el.getAttribute('r'))||0;
    const p=tf2(cx-r,cy-r); const d=r*2*accTf.sx;
    shapes.push({type:'ellipse',x:p.x,y:p.y,w:d,h:d,fill:fillColor,stroke:strokeColor,sw:strokeW});
    return shapes;
  }

  if (tag==='ellipse') {
    const cx=parseFloat(el.getAttribute('cx'))||0, cy=parseFloat(el.getAttribute('cy'))||0;
    const rx=parseFloat(el.getAttribute('rx'))||0, ry=parseFloat(el.getAttribute('ry'))||0;
    const p=tf2(cx-rx,cy-ry);
    shapes.push({type:'ellipse',x:p.x,y:p.y,w:rx*2*accTf.sx,h:ry*2*accTf.sy,fill:fillColor,stroke:strokeColor,sw:strokeW});
    return shapes;
  }

  if (tag==='polygon'||tag==='polyline') {
    const pts=(el.getAttribute('points')||'').trim().split(/[\s,]+/).map(parseFloat).filter(n=>!isNaN(n));
    if(pts.length<2) return shapes;
    const cmds=[];
    for(let k=0;k<pts.length-1;k+=2){const p=tf2(pts[k],pts[k+1]);cmds.push({cmd:k===0?'M':'L',x:p.x,y:p.y});}
    if(tag==='polygon') cmds.push({cmd:'Z'});
    shapes.push({type:'path',cmds,fill:fillColor,stroke:strokeColor,sw:strokeW});
    return shapes;
  }

  if (tag==='line') {
    const p1=tf2(parseFloat(el.getAttribute('x1'))||0, parseFloat(el.getAttribute('y1'))||0);
    const p2=tf2(parseFloat(el.getAttribute('x2'))||0, parseFloat(el.getAttribute('y2'))||0);
    shapes.push({type:'path',cmds:[{cmd:'M',x:p1.x,y:p1.y},{cmd:'L',x:p2.x,y:p2.y}],fill:null,stroke:strokeColor||'000000',sw:strokeW});
    return shapes;
  }

  return shapes;
}

function svgShapesToSpXML(shapes, vbW, vbH, cx, cy, gx, gy) {
  let xml='', id=10;
  const emu = v => Math.round(v / vbW * cx);
  const emv = v => Math.round(v / vbH * cy);

  for (const s of shapes) {
    const fillXML = s.fill ? `<a:solidFill><a:srgbClr val="${s.fill}"/></a:solidFill>` : `<a:noFill/>`;
    const lnXML   = s.stroke
      ? `<a:ln w="${s.sw}"><a:solidFill><a:srgbClr val="${s.stroke}"/></a:solidFill></a:ln>`
      : `<a:ln><a:noFill/></a:ln>`;

    if (s.type === 'path') {
      const pathXML = svgNormCmdsToXML(s.cmds);
      if (!pathXML) continue;
      xml += `<p:sp>
        <p:nvSpPr><p:cNvPr id="${id++}" name="Shape${id}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${gx}" y="${gy}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
            <a:rect l="l" t="t" r="r" b="b"/>
            <a:pathLst><a:path w="${Math.round(vbW)}" h="${Math.round(vbH)}">${pathXML}</a:path></a:pathLst>
          </a:custGeom>
          ${fillXML}${lnXML}
        </p:spPr>
      </p:sp>`;
    } else if (s.type === 'rect') {
      const hasR = s.rx > 0;
      const geom = hasR
        ? `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${Math.round(Math.min(s.rx/Math.max(s.w,1), s.rx/Math.max(s.h,1), 0.5)*50000)}"/></a:avLst></a:prstGeom>`
        : `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
      xml += `<p:sp>
        <p:nvSpPr><p:cNvPr id="${id++}" name="Rect${id}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${gx+emu(s.x)}" y="${gy+emv(s.y)}"/><a:ext cx="${emu(s.w)}" cy="${emv(s.h)}"/></a:xfrm>
          ${geom}${fillXML}${lnXML}
        </p:spPr>
      </p:sp>`;
    } else if (s.type === 'ellipse') {
      xml += `<p:sp>
        <p:nvSpPr><p:cNvPr id="${id++}" name="Ellipse${id}"/><p:cNvSpPr><a:spLocks noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${gx+emu(s.x)}" y="${gy+emv(s.y)}"/><a:ext cx="${emu(s.w)}" cy="${emv(s.h)}"/></a:xfrm>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
          ${fillXML}${lnXML}
        </p:spPr>
      </p:sp>`;
    }
  }
  return xml;
}

function buildShapesOoxml(spTreeContent) {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    </Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/ppt/presentation.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml">
    <pkg:xmlData>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:sldMasterIdLst/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
      </p:presentation>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/ppt/_rels/presentation.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    </Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/ppt/slides/slide1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml">
    <pkg:xmlData>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld><p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
          ${spTreeContent}
        </p:spTree></p:cSld>
      </p:sld>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/ppt/slides/_rels/slide1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/></pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

aasync function insertSVGCode(svgCode) {
  if (!svgCode || !svgCode.includes('<svg')) {
    showStatus('Paste valid SVG first', 'err');
    return;
  }
  try {
    let svg = svgCode;
    if (!svg.includes('viewBox')) {
      const wm = svg.match(/width=["']([^"']*)["']/);
      const hm = svg.match(/height=["']([^"']*)["']/);
      if (wm && hm) {
        svg = svg.replace(/<svg/, `<svg viewBox="0 0 ${parseFloat(wm[1])||100} ${parseFloat(hm[1])||100}"`);
      }
    }
    svg = svg.replace(/width=["'][^"']*["']/i, 'width="200"');
    svg = svg.replace(/height=["'][^"']*["']/i, 'height="200"');
    if (!svg.match(/width=/i)) {
      svg = svg.replace(/<svg/, '<svg width="200" height="200"');
    }
    const base64 = btoa(unescape(encodeURIComponent(svg)));
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image, imageLeft:100, imageTop:100, imageWidth:200, imageHeight:200 },
        r => {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(r.error.message));
        }
      );
    });
    showStatus('✓ SVG inserted. Right-click → Convert to Shape', 'ok');
  } catch(e) {
    showStatus('Error: ' + e.message, 'err');
  }
}

async function insertSVGToSlide() {
  const code = document.getElementById('svg-input').value.trim();
  if (!code) { showStatus('Paste SVG code first', 'err'); return; }
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
  showStatus('Saved to library!', 'ok');
}

function deleteSVGFromLibrary(id) {
  const lib = getSVGLibrary().filter(function(s) { return s.id !== id; });
  localStorage.setItem('svgLibrary', JSON.stringify(lib));
  renderSVGLibrary();
}

async function insertSVGFromLibrary(id) {
  const item = getSVGLibrary().find(function(s) { return s.id === id; });
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
  grid.innerHTML = lib.map(function(s) {
    return '<div class="shape-item" onclick="insertSVGFromLibrary(' + s.id + ')">' +
      '<div class="preview">' + s.code + '</div>' +
      '<span class="lbl">' + s.name + '</span>' +
      '<button class="del-btn" onclick="event.stopPropagation();deleteSVGFromLibrary(' + s.id + ')">×</button>' +
      '</div>';
  }).join('');
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
      showStatus('Color read: ' + hex, 'ok');
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

    shapes.items.forEach(function(s) {
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

  showStatus('Replaced ' + count + ' item(s)', count > 0 ? 'ok' : 'err');
}

// Expose all functions globally for HTML inline event handlers
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
