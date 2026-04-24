Office.onReady(() => {
  loadPaletteUI();
  loadTableLibraryUI();
  loadParaStyleLibraryUI();
  loadSystemFonts();
  renderStandardShapes();
  loadSVGLibraryUI();
});

// ── TABS ──────────────────────────────────────────────
function switchTab(name) {
  const names = ['shapes','library','table','text','colors'];
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', names[i]===name));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
}

// ── STATUS ────────────────────────────────────────────
function showStatus(msg, type='ok') {
  const el = document.getElementById('status');
  el.textContent = msg; el.className = type;
  setTimeout(() => el.className='', 3000);
}

function toggleBtn(el) { el.classList.toggle('on'); }

// ── ALIGN MODE ────────────────────────────────────────
let alignMode = 'slide';
function setAlignMode(mode) {
  alignMode = mode;
  document.getElementById('mode-slide').classList.toggle('active', mode==='slide');
  document.getElementById('mode-sel').classList.toggle('active', mode==='selection');
}

// ── GET SELECTED SHAPES (الطريقة الصحيحة) ────────────
async function getSelectedShapesWithProps(context, props) {
  const shapes = context.presentation.getSelectedShapes();
  // Step 1: load items first
  shapes.load("items");
  await context.sync();

  if (!shapes.items.length) return [];

  // Step 2: load properties for each item
  shapes.items.forEach(s => s.load(props));
  await context.sync();

  return shapes.items;
}

// ── ALIGN ─────────────────────────────────────────────
async function alignShapes(direction) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");

      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first","err");

      shapes.items.forEach(s => s.load("left,top,width,height"));
      await context.sync();

      const items = shapes.items;
      let refL, refR, refT, refB;

      if (alignMode === 'slide') {
        refL=0; refR=slide.width; refT=0; refB=slide.height;
      } else {
        refL = Math.min(...items.map(s=>s.left));
        refR = Math.max(...items.map(s=>s.left+s.width));
        refT = Math.min(...items.map(s=>s.top));
        refB = Math.max(...items.map(s=>s.top+s.height));
      }

      items.forEach(s => {
        switch(direction) {
          case 'left':   s.left = refL; break;
          case 'right':  s.left = refR - s.width; break;
          case 'center': s.left = refL + (refR-refL-s.width)/2; break;
          case 'top':    s.top  = refT; break;
          case 'bottom': s.top  = refB - s.height; break;
          case 'middle': s.top  = refT + (refB-refT-s.height)/2; break;
        }
      });

      await context.sync();
      showStatus(`✓ Aligned ${direction}`);
    });
  } catch(e) { showStatus(e.message,"err"); }
}

// ── DISTRIBUTE ────────────────────────────────────────
async function distributeShapes(axis) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length < 3) return showStatus("Select at least 3 shapes","err");

      shapes.items.forEach(s => s.load("left,top,width,height"));
      await context.sync();

      const items = [...shapes.items];

      if (axis === 'h') {
        items.sort((a,b) => a.left - b.left);
        const first = items[0].left;
        const last  = items[items.length-1].left + items[items.length-1].width;
        const totalW = items.reduce((s,sh)=>s+sh.width,0);
        const gap = (last - first - totalW) / (items.length-1);
        let x = first;
        items.forEach(s => { s.left=x; x+=s.width+gap; });
      } else {
        items.sort((a,b) => a.top - b.top);
        const first = items[0].top;
        const last  = items[items.length-1].top + items[items.length-1].height;
        const totalH = items.reduce((s,sh)=>s+sh.height,0);
        const gap = (last - first - totalH) / (items.length-1);
        let y = first;
        items.forEach(s => { s.top=y; y+=s.height+gap; });
      }

      await context.sync();
      showStatus("✓ Distributed");
    });
  } catch(e) { showStatus(e.message,"err"); }
}

// ── MATCH SIZE ────────────────────────────────────────
async function matchSize(type) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length < 2) return showStatus("Select at least 2 shapes","err");

      shapes.items.forEach(s => s.load("width,height"));
      await context.sync();

      const refW = shapes.items[0].width;
      const refH = shapes.items[0].height;

      shapes.items.forEach(s => {
        if (type==='width'||type==='both')  s.width  = refW;
        if (type==='height'||type==='both') s.height = refH;
      });

      await context.sync();
      showStatus("✓ Size matched");
    });
  } catch(e) { showStatus(e.message,"err"); }
}

// ── CONVERT TO ROUND RECT ─────────────────────────────
async function convertToRoundRect() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first","err");

      shapes.items.forEach(s => s.load("left,top,width,height,fill/foregroundColor,lineFormat/color,lineFormat/weight,lineFormat/visible"));
      await context.sync();

      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      let count=0;

      for (const shape of shapes.items) {
        const L=shape.left, T=shape.top, W=shape.width, H=shape.height;
        let fill="#4472C4", lc="#000000", lw=1, lv=false;
        try { fill = shape.fill.foregroundColor || fill; }    catch(e){}
        try { lc   = shape.lineFormat.color || lc; }          catch(e){}
        try { lw   = shape.lineFormat.weight || lw; }         catch(e){}
        try { lv   = shape.lineFormat.visible; }              catch(e){}

        shape.delete();
        await context.sync();

        const ns = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.roundRectangle,
          { left:L, top:T, width:W, height:H }
        );
        ns.fill.setSolidColor(fill);
        ns.lineFormat.color   = lc;
        ns.lineFormat.weight  = lw;
        ns.lineFormat.visible = lv;
        await context.sync();
        count++;
      }
      showStatus(`✓ Converted ${count} shape(s)`);
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── CORNER RADIUS ─────────────────────────────────────
// Formula: adjValue = radiusPt / shortSide (max 0.5)
// shortSide = min(width, height)
async function applyCornerRadius() {
  const radiusPt = parseFloat(document.getElementById('radiusInput').value);
  if (isNaN(radiusPt) || radiusPt < 0) return showStatus("Enter valid radius","err");

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first","err");

      shapes.items.forEach(s => s.load("width,height"));
      await context.sync();

      let applied=0, skipped=0;

      for (const shape of shapes.items) {
        try {
          const shortSide = Math.min(shape.width, shape.height);
          const adjValue  = Math.min(radiusPt / shortSide, 0.5);

          shape.adjustments.load("items");
          await context.sync();

          if (shape.adjustments.items && shape.adjustments.items.length > 0) {
            shape.adjustments.items[0].value = adjValue;
            await context.sync();
            applied++;
          } else {
            skipped++;
          }
        } catch(e) { skipped++; }
      }

      if (applied > 0) showStatus(`✓ Radius ${radiusPt}pt → ${applied} shape(s)`);
      else showStatus("Select Round Rectangle shapes first","err");
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── FILL ──────────────────────────────────────────────
async function applyFillColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── OPACITY ───────────────────────────────────────────
async function applyOpacity(value) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      shapes.items.forEach(s => { try { s.fill.transparency=1-parseFloat(value)/100; } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── BORDER COLOR ──────────────────────────────────────
async function applyBorderColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      shapes.items.forEach(s => { try { s.lineFormat.color=hex; s.lineFormat.visible=true; } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── BORDER WIDTH ──────────────────────────────────────
async function applyBorderWidth(value) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      const pt = parseFloat(value);
      shapes.items.forEach(s => {
        try {
          if (pt===0) s.lineFormat.visible=false;
          else { s.lineFormat.visible=true; s.lineFormat.weight=pt; }
        } catch(e){}
      });
      await context.sync();
    });
  } catch(e){}
}

// ── STANDARD SHAPES ───────────────────────────────────
const STANDARD_SHAPES = [
  { name:"Rect",       type:"rectangle",            svg:'<rect x="2" y="4" width="16" height="12" fill="#555"/>' },
  { name:"RoundRect",  type:"roundRectangle",       svg:'<rect x="2" y="4" width="16" height="12" rx="3" fill="#555"/>' },
  { name:"Ellipse",    type:"ellipse",              svg:'<ellipse cx="10" cy="10" rx="8" ry="6" fill="#555"/>' },
  { name:"Triangle",   type:"isoscelesTriangle",    svg:'<polygon points="10,2 18,18 2,18" fill="#555"/>' },
  { name:"RtTri",      type:"rightTriangle",        svg:'<polygon points="2,18 18,18 2,2" fill="#555"/>' },
  { name:"Diamond",    type:"diamond",              svg:'<polygon points="10,2 18,10 10,18 2,10" fill="#555"/>' },
  { name:"Pentagon",   type:"regularPentagon",      svg:'<polygon points="10,2 18,7 15,17 5,17 2,7" fill="#555"/>' },
  { name:"Hexagon",    type:"hexagon",              svg:'<polygon points="5,2 15,2 20,10 15,18 5,18 0,10" fill="#555"/>' },
  { name:"Octagon",    type:"octagon",              svg:'<polygon points="6,2 14,2 18,6 18,14 14,18 6,18 2,14 2,6" fill="#555"/>' },
  { name:"Parallelgm", type:"parallelogram",        svg:'<polygon points="5,17 2,3 15,3 18,17" fill="#555"/>' },
  { name:"Trapezoid",  type:"trapezoid",            svg:'<polygon points="3,17 17,17 14,3 6,3" fill="#555"/>' },
  { name:"Star4",      type:"star4",                svg:'<polygon points="10,2 12,8 18,10 12,12 10,18 8,12 2,10 8,8" fill="#555"/>' },
  { name:"Star5",      type:"star5",                svg:'<polygon points="10,2 12,7 18,7 13,11 15,17 10,13 5,17 7,11 2,7 8,7" fill="#555"/>' },
  { name:"Star6",      type:"star6",                svg:'<polygon points="10,2 13,7 18,6 15,11 17,16 12,14 10,18 8,14 3,16 5,11 2,6 7,7" fill="#555"/>' },
  { name:"Arrow R",    type:"rightArrow",           svg:'<polygon points="2,7 13,7 13,4 18,10 13,16 13,13 2,13" fill="#555"/>' },
  { name:"Arrow L",    type:"leftArrow",            svg:'<polygon points="18,7 7,7 7,4 2,10 7,16 7,13 18,13" fill="#555"/>' },
  { name:"Arrow U",    type:"upArrow",              svg:'<polygon points="10,2 16,8 13,8 13,18 7,18 7,8 4,8" fill="#555"/>' },
  { name:"Arrow D",    type:"downArrow",            svg:'<polygon points="10,18 4,12 7,12 7,2 13,2 13,12 16,12" fill="#555"/>' },
  { name:"Chevron",    type:"chevron",              svg:'<polygon points="2,5 12,5 17,10 12,15 2,15 7,10" fill="#555"/>' },
  { name:"Callout",    type:"wedgeRectCallout",     svg:'<rect x="2" y="2" width="13" height="10" rx="1" fill="#555"/><polygon points="5,12 3,18 9,12" fill="#555"/>' },
  { name:"Heart",      type:"heart",                svg:'<path d="M10,16 C6,12 2,9 2,6 A4,4,0,0,1,10,5 A4,4,0,0,1,18,6 C18,9 14,12 10,16Z" fill="#555"/>' },
  { name:"Cross",      type:"plus",                 svg:'<rect x="7" y="2" width="6" height="16" fill="#555"/><rect x="2" y="7" width="16" height="6" fill="#555"/>' },
  { name:"Moon",       type:"moon",                 svg:'<path d="M13,4 A8,8,0,1,0,13,16 A5,5,0,1,1,13,4Z" fill="#555"/>' },
  { name:"Donut",      type:"donut",                svg:'<circle cx="10" cy="10" r="8" fill="#555"/><circle cx="10" cy="10" r="4" fill="white"/>' },
  { name:"Frame",      type:"frame",                svg:'<rect x="2" y="2" width="16" height="16" fill="none" stroke="#555" stroke-width="3"/>' },
];

function renderStandardShapes() {
  const grid = document.getElementById('standardShapeGrid');
  if (!grid) return;
  grid.innerHTML = STANDARD_SHAPES.map((s,i) => `
    <div class="shape-item" onclick="insertStandardShape(${i})" title="${s.name}">
      <svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">${s.svg}</svg>
      <span>${s.name}</span>
    </div>
  `).join('');
}

async function insertStandardShape(index) {
  const s = STANDARD_SHAPES[index];
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");
      await context.sync();

      const shapeType = PowerPoint.GeometricShapeType[s.type];
      if (!shapeType) return showStatus("Shape not supported","err");

      const shape = slide.shapes.addGeometricShape(shapeType, {
        left:  (slide.width  - 150) / 2,
        top:   (slide.height - 150) / 2,
        width:  150,
        height: 150
      });
      shape.fill.setSolidColor("#4472C4");
      shape.lineFormat.visible = false;
      await context.sync();
      showStatus(`✓ ${s.name} inserted`);
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── SVG LIBRARY ───────────────────────────────────────
function getSVGLib()     { try { return JSON.parse(localStorage.getItem("svgLibrary")||"[]"); } catch { return []; } }
function saveSVGLib(lib) { localStorage.setItem("svgLibrary", JSON.stringify(lib)); }

function addSVGToLibrary() {
  const code = document.getElementById('svg-input').value.trim();
  if (!code || !code.includes('<svg')) return showStatus("Paste valid SVG first","err");
  const name = document.getElementById('svgName').value.trim() || `SVG ${getSVGLib().length+1}`;
  const lib  = getSVGLib();
  lib.push({ id:Date.now(), name, code });
  saveSVGLib(lib);
  loadSVGLibraryUI();
  document.getElementById('svg-input').value = '';
  document.getElementById('svgName').value   = '';
  showStatus(`✓ "${name}" added`);
}

function loadSVGLibraryUI() {
  const lib  = getSVGLib();
  const grid = document.getElementById('svgLibraryGrid');
  if (!grid) return;
  if (!lib.length) {
    grid.innerHTML = '<div style="font-size:11px;color:#aaa;grid-column:span 5;text-align:center;padding:8px">No SVGs yet</div>';
    return;
  }
  grid.innerHTML = lib.map(item => {
    const preview = item.code
      .replace(/width="[^"]*"/, 'width="28"')
      .replace(/height="[^"]*"/, 'height="28"');
    return `
      <div class="shape-item" onclick="insertSVGFromLibrary(${item.id})" title="${item.name}">
        <div style="width:28px;height:28px;overflow:hidden;display:flex;align-items:center;justify-content:center">${preview}</div>
        <span>${item.name}</span>
        <span class="del-shape" onclick="event.stopPropagation();deleteSVGFromLib(${item.id})">×</span>
      </div>`;
  }).join('');
}

async function insertSVGFromLibrary(id) {
  const item = getSVGLib().find(i=>i.id===id);
  if (!item) return;
  try {
    const base64 = btoa(unescape(encodeURIComponent(item.code)));
    Office.context.document.setSelectedDataAsync(
      base64,
      { coercionType:Office.CoercionType.Image, imageLeft:100, imageTop:100, imageWidth:200, imageHeight:200 },
      r => {
        if (r.status===Office.AsyncResultStatus.Failed) showStatus("Error: "+r.error.message,"err");
        else showStatus("✓ Inserted! Right-click → Convert to Shape");
      }
    );
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

function deleteSVGFromLib(id) { saveSVGLib(getSVGLib().filter(i=>i.id!==id)); loadSVGLibraryUI(); }

// ── TABLE ─────────────────────────────────────────────
async function createTable() {
  const rows=parseInt(document.getElementById('tblRows').value);
  const cols=parseInt(document.getElementById('tblCols').value);
  const rowH=parseFloat(document.getElementById('tblRowH').value);
  const colWI=parseFloat(document.getElementById('tblColW').value);
  const hBg=document.getElementById('tblHeaderBg').value;
  const hFg=document.getElementById('tblHeaderFg').value;
  const hSz=parseFloat(document.getElementById('tblHeaderSize').value);
  const hFn=document.getElementById('tblHeaderFont').value||'Calibri';
  const hBo=document.getElementById('tblHeaderBold').classList.contains('on');
  const hIt=document.getElementById('tblHeaderItalic').classList.contains('on');
  const hCp=document.getElementById('tblHeaderCaps').classList.contains('on');
  const hAl=document.getElementById('tblHeaderAlign').value;
  const r1=document.getElementById('tblRow1').value;
  const r2=document.getElementById('tblRow2').value;
  const bFg=document.getElementById('tblBodyFg').value;
  const bSz=parseFloat(document.getElementById('tblBodySize').value);
  const bFn=document.getElementById('tblBodyFont').value||'Calibri';
  const bAl=document.getElementById('tblBodyAlign').value;
  const bc=document.getElementById('tblBorder').value;
  const bw=parseFloat(document.getElementById('tblBorderW').value);
  const pd=parseFloat(document.getElementById('tblPadding').value);
  const aMap={left:'Left',center:'Center',right:'Right'};

  try {
    await PowerPoint.run(async (context) => {
      const slide=context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");
      await context.sync();

      const tw=slide.width*0.8;
      const cw=colWI>0?colWI:tw/cols;

      const scp=Array(rows).fill("").map((_,r)=>Array(cols).fill("").map(()=>({
        fill:{color:r===0?hBg:(r%2===0?r2:r1)},
        font:{color:r===0?hFg:bFg,size:r===0?hSz:bSz,name:r===0?hFn:bFn,bold:r===0?hBo:false,italic:r===0?hIt:false,allCaps:r===0?hCp:false},
        margins:{top:pd,bottom:pd,left:pd,right:pd},
        horizontalAlignment:aMap[r===0?hAl:bAl]||'Left',
        borders:{bottom:{color:bc,weight:bw},top:{color:bc,weight:bw},left:{color:bc,weight:bw},right:{color:bc,weight:bw}}
      })));

      const values=Array(rows).fill("").map((_,r)=>Array(cols).fill("").map((_,c)=>r===0?`Header ${c+1}`:""));
      const columns=Array(cols).fill("").map(()=>({columnWidth:cw}));
      const rowsOpt=Array(rows).fill("").map(()=>({rowHeight:rowH}));

      const shape=slide.shapes.addTable(rows,cols,{values,specificCellProperties:scp,columns,rows:rowsOpt});
      shape.left=slide.width*0.1;
      shape.top=slide.height*0.15;
      await context.sync();
      showStatus("✓ Table inserted!");
    });
  } catch(e) { showStatus(e.message,"err"); }
}

function getTableLib() { try{return JSON.parse(localStorage.getItem("tableLibrary")||"[]");}catch{return[];} }
function saveTableLib(lib) { localStorage.setItem("tableLibrary",JSON.stringify(lib)); }

function saveTableStyle() {
  const lib=getTableLib(), name=prompt("Name:",`Table ${lib.length+1}`);
  if(!name) return;
  lib.push({id:Date.now(),name,
    rows:parseInt(document.getElementById('tblRows').value),
    cols:parseInt(document.getElementById('tblCols').value),
    rowH:parseFloat(document.getElementById('tblRowH').value),
    colW:parseFloat(document.getElementById('tblColW').value),
    headerBg:document.getElementById('tblHeaderBg').value,
    headerFg:document.getElementById('tblHeaderFg').value,
    headerSize:parseFloat(document.getElementById('tblHeaderSize').value),
    headerFont:document.getElementById('tblHeaderFont').value,
    headerBold:document.getElementById('tblHeaderBold').classList.contains('on'),
    headerItalic:document.getElementById('tblHeaderItalic').classList.contains('on'),
    headerCaps:document.getElementById('tblHeaderCaps').classList.contains('on'),
    headerAlign:document.getElementById('tblHeaderAlign').value,
    row1:document.getElementById('tblRow1').value,
    row2:document.getElementById('tblRow2').value,
    bodyFg:document.getElementById('tblBodyFg').value,
    bodySize:parseFloat(document.getElementById('tblBodySize').value),
    bodyFont:document.getElementById('tblBodyFont').value,
    bodyAlign:document.getElementById('tblBodyAlign').value,
    borderColor:document.getElementById('tblBorder').value,
    borderW:parseFloat(document.getElementById('tblBorderW').value),
    padding:parseFloat(document.getElementById('tblPadding').value),
  });
  saveTableLib(lib); loadTableLibraryUI(); showStatus(`✓ Saved "${name}"`);
}

function loadTableStyle(item) {
  document.getElementById('tblRows').value=item.rows;
  document.getElementById('tblCols').value=item.cols;
  document.getElementById('tblRowH').value=item.rowH;
  document.getElementById('tblColW').value=item.colW||0;
  document.getElementById('tblHeaderBg').value=item.headerBg;
  document.getElementById('tblHeaderFg').value=item.headerFg;
  document.getElementById('tblHeaderSize').value=item.headerSize;
  document.getElementById('tblHeaderFont').value=item.headerFont;
  document.getElementById('tblHeaderAlign').value=item.headerAlign;
  document.getElementById('tblRow1').value=item.row1;
  document.getElementById('tblRow2').value=item.row2;
  document.getElementById('tblBodyFg').value=item.bodyFg;
  document.getElementById('tblBodySize').value=item.bodySize;
  document.getElementById('tblBodyFont').value=item.bodyFont;
  document.getElementById('tblBodyAlign').value=item.bodyAlign;
  document.getElementById('tblBorder').value=item.borderColor;
  document.getElementById('tblBorderW').value=item.borderW;
  document.getElementById('tblBorderWVal').textContent=item.borderW;
  document.getElementById('tblPadding').value=item.padding;
  document.getElementById('tblHeaderBold').classList.toggle('on',!!item.headerBold);
  document.getElementById('tblHeaderItalic').classList.toggle('on',!!item.headerItalic);
  document.getElementById('tblHeaderCaps').classList.toggle('on',!!item.headerCaps);
  showStatus(`✓ Loaded "${item.name}"`);
}

function loadTableLibraryUI() {
  const lib=getTableLib(), el=document.getElementById('tableLibrary');
  if(!lib.length){el.innerHTML='<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved templates</div>';return;}
  el.innerHTML=lib.map(item=>`
    <div class="table-item" onclick="loadTableStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div class="table-info"><b>${item.name}</b>${item.rows}×${item.cols} · ${item.rowH}pt</div>
      <div style="display:flex;gap:4px;align-items:center">
        <div style="width:12px;height:12px;background:${item.headerBg};border-radius:2px"></div>
        <div style="width:12px;height:12px;background:${item.row1};border-radius:2px"></div>
        <span onclick="event.stopPropagation();deleteTableStyle(${item.id})" style="color:#ccc;font-size:16px;cursor:pointer">×</span>
      </div>
    </div>`).join('');
}
function deleteTableStyle(id){saveTableLib(getTableLib().filter(i=>i.id!==id));loadTableLibraryUI();}
function clearTableLib(){if(confirm("Clear all?")){saveTableLib([]);loadTableLibraryUI();}}

// ── FONT PICKER ───────────────────────────────────────
const FONTS=["Arial","Arial Black","Calibri","Calibri Light","Cambria","Century Gothic",
  "Comic Sans MS","Courier New","Garamond","Georgia","Helvetica","Impact","Palatino",
  "Rockwell","Tahoma","Times New Roman","Trebuchet MS","Verdana","Montserrat","Oswald",
  "Raleway","Roboto","Lato","Poppins","Inter","Futura","Gill Sans","Franklin Gothic Medium"];
let allFonts=[...FONTS];

function loadSystemFonts() {
  try {
    if('fonts' in document) {
      document.fonts.ready.then(()=>{
        const det=FONTS.filter(f=>document.fonts.check(`12px "${f}"`));
        if(det.length>5) allFonts=[...new Set([...det,...FONTS])];
        renderFontDropdown(allFonts);
      });
    }
  } catch(e){}
  renderFontDropdown(allFonts);
}
function renderFontDropdown(fonts) {
  const dd=document.getElementById('fontDropdown');
  if(dd) dd.innerHTML=fonts.map(f=>`<div class="font-option" style="font-family:'${f}'" onclick="selectFont('${f}')">${f}</div>`).join('');
}
function filterFonts(v) {
  renderFontDropdown(allFonts.filter(f=>f.toLowerCase().includes(v.toLowerCase())));
  document.getElementById('fontDropdown').classList.add('open');
}
function openFontDropdown() {
  renderFontDropdown(allFonts);
  document.getElementById('fontDropdown').classList.add('open');
}
function selectFont(n) {
  document.getElementById('txtFont').value=n;
  document.getElementById('fontDropdown').classList.remove('open');
  liveChar();
}
document.addEventListener('click', e=>{
  if(!e.target.closest('.font-wrap'))
    document.getElementById('fontDropdown').classList.remove('open');
});

// ── TEXT ──────────────────────────────────────────────
let charTimer=null;
function liveChar(){clearTimeout(charTimer);charTimer=setTimeout(applyChar,400);}

async function applyChar() {
  const fontName  = document.getElementById('txtFont').value;
  const style     = document.getElementById('txtStyle').value;
  const fontSize  = parseFloat(document.getElementById('txtSize').value);
  const fontColor = document.getElementById('txtColor').value;
  const bold    = style==='bold'||style==='bolditalic'||document.getElementById('txtBold').classList.contains('on');
  const italic  = style==='italic'||style==='bolditalic'||document.getElementById('txtItalic').classList.contains('on');
  const underline = document.getElementById('txtUnderline').classList.contains('on');
  const strike    = document.getElementById('txtStrike').classList.contains('on');
  const allCaps   = document.getElementById('txtCaps').classList.contains('on');
  const sup       = document.getElementById('txtSuper').classList.contains('on');
  const sub       = document.getElementById('txtSub').classList.contains('on');

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      for(const s of shapes.items){
        try{
          const font=s.textFrame.textRange.font;
          if(fontName) font.name=fontName;
          if(!isNaN(fontSize)) font.size=fontSize;
          font.color=fontColor; font.bold=bold; font.italic=italic;
          font.allCaps=allCaps; font.strikethrough=strike;
          font.superscript=sup; font.subscript=sub;
          font.underline=underline
            ?PowerPoint.ShapeFontUnderlineStyle.single
            :PowerPoint.ShapeFontUnderlineStyle.none;
        }catch(e){}
      }
      await context.sync();
      showStatus("✓ Character applied");
    });
  } catch(e){ showStatus(e.message,"err"); }
}

async function toggleLive(el,style){el.classList.toggle('on');await applyChar();}

async function applyPara() {
  const align=document.getElementById('txtAlign').value;
  const aMap={
    left:   PowerPoint.ParagraphHorizontalAlignment.left,
    center: PowerPoint.ParagraphHorizontalAlignment.center,
    right:  PowerPoint.ParagraphHorizontalAlignment.right,
    justify:PowerPoint.ParagraphHorizontalAlignment.justify
  };
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      for(const s of shapes.items){
        try{ s.textFrame.textRange.paragraphFormat.horizontalAlignment=aMap[align]; }catch(e){}
      }
      await context.sync();
      showStatus("✓ Paragraph applied");
    });
  } catch(e){ showStatus(e.message,"err"); }
}

// ── PARA STYLES ───────────────────────────────────────
function getParaStyles(){try{return JSON.parse(localStorage.getItem("paraStyles")||"[]");}catch{return[];}}
function saveParaStylesLib(s){localStorage.setItem("paraStyles",JSON.stringify(s));}

function saveParaStyle(){
  const styles=getParaStyles(), name=prompt("Style name:",`Style ${styles.length+1}`);
  if(!name) return;
  styles.push({
    id:Date.now(), name,
    fontName:document.getElementById('txtFont').value,
    style:document.getElementById('txtStyle').value,
    fontSize:parseFloat(document.getElementById('txtSize').value),
    fontColor:document.getElementById('txtColor').value,
    bold:document.getElementById('txtBold').classList.contains('on'),
    italic:document.getElementById('txtItalic').classList.contains('on'),
    underline:document.getElementById('txtUnderline').classList.contains('on'),
    strike:document.getElementById('txtStrike').classList.contains('on'),
    allCaps:document.getElementById('txtCaps').classList.contains('on'),
    superscript:document.getElementById('txtSuper').classList.contains('on'),
    subscript:document.getElementById('txtSub').classList.contains('on'),
    align:document.getElementById('txtAlign').value,
  });
  saveParaStylesLib(styles); loadParaStyleLibraryUI(); showStatus(`✓ "${name}" saved`);
}

function loadParaStyle(item){
  document.getElementById('txtFont').value=item.fontName;
  document.getElementById('txtStyle').value=item.style||'normal';
  document.getElementById('txtSize').value=item.fontSize;
  document.getElementById('txtColor').value=item.fontColor;
  document.getElementById('txtAlign').value=item.align;
  document.getElementById('txtBold').classList.toggle('on',!!item.bold);
  document.getElementById('txtItalic').classList.toggle('on',!!item.italic);
  document.getElementById('txtUnderline').classList.toggle('on',!!item.underline);
  document.getElementById('txtStrike').classList.toggle('on',!!item.strike);
  document.getElementById('txtCaps').classList.toggle('on',!!item.allCaps);
  document.getElementById('txtSuper').classList.toggle('on',!!item.superscript);
  document.getElementById('txtSub').classList.toggle('on',!!item.subscript);
  applyChar(); applyPara();
  showStatus(`✓ Loaded "${item.name}"`);
}

function loadParaStyleLibraryUI(){
  const styles=getParaStyles(), el=document.getElementById('paraStyleLibrary');
  if(!styles.length){el.innerHTML='<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved styles</div>';return;}
  el.innerHTML=styles.map(item=>`
    <div class="style-item" onclick="loadParaStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div>
        <div style="font-family:'${item.fontName}';color:${item.fontColor};font-weight:${item.bold?'bold':'normal'};font-style:${item.italic?'italic':'normal'};font-size:13px">${item.name}</div>
        <div class="style-meta">${item.fontName} · ${item.fontSize}pt</div>
      </div>
      <span onclick="event.stopPropagation();deleteParaStyle(${item.id})" style="color:#ccc;font-size:16px;cursor:pointer">×</span>
    </div>`).join('');
}
function deleteParaStyle(id){saveParaStylesLib(getParaStyles().filter(s=>s.id!==id));loadParaStyleLibraryUI();}
function clearParaStyles(){if(confirm("Clear?")){saveParaStylesLib([]);loadParaStyleLibraryUI();}}

// ── COLOR PALETTE ─────────────────────────────────────
let selectedColor=null;
function getPalette(){try{return JSON.parse(localStorage.getItem("colorPalette")||"[]");}catch{return[];}}
function savePalette(p){localStorage.setItem("colorPalette",JSON.stringify(p));}

function addColor(){
  const hex=document.getElementById('newColor').value;
  const name=document.getElementById('newColorName').value||hex;
  const p=getPalette(); p.push({id:Date.now(),hex,name}); savePalette(p); loadPaletteUI();
  document.getElementById('newColorName').value='';
}
function deleteColor(id){savePalette(getPalette().filter(c=>c.id!==id));loadPaletteUI();}

function loadPaletteUI(){
  const p=getPalette();
  const mk=(fn,del)=>p.map(c=>`
    <div class="swatch" style="background:${c.hex}" title="${c.name}" onclick="${fn}('${c.hex}','${c.name}')">
      ${del?`<span class="del" onclick="event.stopPropagation();deleteColor(${c.id})">×</span>`:''}
    </div>`).join('')||'<span style="font-size:11px;color:#aaa">No colors</span>';
  document.getElementById('paletteGrid').innerHTML=mk('selectColor',true);
  document.getElementById('applyPaletteGrid').innerHTML=mk('selectColor',false);
}

function selectColor(hex,name){
  selectedColor=hex;
  document.getElementById('selectedColorPreview').innerHTML=
    `<span style="display:inline-block;width:12px;height:12px;background:${hex};border-radius:2px;margin-right:4px;vertical-align:middle"></span>${name}`;
}

async function applyColorToSelected(type){
  if(!selectedColor) return showStatus("Click a color first","err");
  if(type==='fill') await applyFillColor(selectedColor);
  else              await applyBorderColor(selectedColor);
}

async function bulkColorReplace(){
  const fh=document.getElementById('findColor').value.replace('#','').toUpperCase();
  const rh=document.getElementById('replaceColor').value;
  try {
    await PowerPoint.run(async(context)=>{
      const slide=context.presentation.getSelectedSlides().getItemAt(0);
      const shapes=slide.shapes;
      shapes.load("items");
      await context.sync();
      shapes.items.forEach(s=>s.load("fill/foregroundColor,type"));
      await context.sync();
      let count=0;
      for(const s of shapes.items){
        try{
          if((s.fill.foregroundColor||'').replace('#','').toUpperCase()===fh){
            s.fill.setSolidColor(rh); count++;
          }
        }catch(e){}
      }
      await context.sync();
      showStatus(`✓ Replaced ${count} shape(s)`);
    });
  } catch(e){ showStatus(e.message,"err"); }
}