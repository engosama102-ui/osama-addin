Office.onReady(() => {
  loadPaletteUI();
  renderStandardShapes();
  loadSVGLibraryUI();
  loadGradientLibraryUI();
  updateGradientPreview();
});

// ── STATUS ────────────────────────────────────────────
function showStatus(msg, type='ok') {
  const el = document.getElementById('status');
  el.textContent = msg; el.className = type;
  setTimeout(() => el.className='', 3000);
}

// ── FILL MODE ─────────────────────────────────────────
function setFillMode(mode) {
  document.querySelectorAll('.fill-tab').forEach((t,i) =>
    t.classList.toggle('active', ['solid','gradient'][i]===mode));
  document.querySelectorAll('.fill-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('fill-'+mode).classList.add('active');
}

// ── FILL COLOR (live) ─────────────────────────────────
async function applyFillColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/type");
      await context.sync();
      shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── EYEDROPPER ────────────────────────────────────────
async function eyedropperFill() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/foregroundColor");
      await context.sync();
      if (!shapes.items.length) return showStatus("Select a shape first","err");
      const color = shapes.items[0].fill.foregroundColor;
      if (color) {
        document.getElementById('fillColor').value = color.startsWith('#') ? color : '#'+color;
        document.getElementById('newColor').value  = document.getElementById('fillColor').value;
        showStatus(`✓ Color picked: ${color}`);
      }
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── GRADIENT ──────────────────────────────────────────
function updateGradientPreview() {
  const c1    = document.getElementById('grad1').value;
  const c2    = document.getElementById('grad2').value;
  const angle = document.getElementById('gradAngle').value;
  document.getElementById('gradPreview').style.background =
    `linear-gradient(${angle}deg, ${c1}, ${c2})`;
}

async function applyGradient() {
  const c1    = document.getElementById('grad1').value;
  const c2    = document.getElementById('grad2').value;
  const angle = parseFloat(document.getElementById('gradAngle').value);

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/type");
      await context.sync();
      if (!shapes.items.length) return showStatus("Select shapes first","err");

      shapes.items.forEach(s => {
        try {
          s.fill.setLinearGradient(angle, [
            { color: c1.replace('#',''), position: 0 },
            { color: c2.replace('#',''), position: 1 }
          ]);
        } catch(e) {
          // fallback: solid color if gradient not supported
          try { s.fill.setSolidColor(c1); } catch(e2){}
        }
      });
      await context.sync();
      showStatus("✓ Gradient applied");
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── GRADIENT LIBRARY ──────────────────────────────────
function getGradLib()     { try { return JSON.parse(localStorage.getItem("gradLibrary")||"[]"); } catch { return []; } }
function saveGradLib(lib) { localStorage.setItem("gradLibrary", JSON.stringify(lib)); }

function saveGradient() {
  const lib   = getGradLib();
  const name  = prompt("Gradient name:", `Gradient ${lib.length+1}`);
  if (!name) return;
  lib.push({
    id:    Date.now(), name,
    c1:    document.getElementById('grad1').value,
    c2:    document.getElementById('grad2').value,
    angle: document.getElementById('gradAngle').value
  });
  saveGradLib(lib);
  loadGradientLibraryUI();
  showStatus(`✓ "${name}" saved`);
}

function loadGradientLibraryUI() {
  const lib = getGradLib();
  const el  = document.getElementById('gradLibrary');
  if (!lib.length) { el.innerHTML=''; return; }
  el.innerHTML = lib.map(item => `
    <div style="display:flex;align-items:center;gap:6px;margin-top:5px">
      <div style="flex:1;height:20px;border-radius:4px;cursor:pointer;
        background:linear-gradient(${item.angle}deg,${item.c1},${item.c2});
        border:1px solid #ddd"
        onclick="loadGradient(${JSON.stringify(item).replace(/"/g,'&quot;')})"
        title="${item.name}"></div>
      <span style="font-size:11px;color:#666;min-width:60px">${item.name}</span>
      <span onclick="deleteGradient(${item.id})" style="color:#ccc;font-size:14px;cursor:pointer">×</span>
    </div>`).join('');
}

function loadGradient(item) {
  document.getElementById('grad1').value = item.c1;
  document.getElementById('grad2').value = item.c2;
  document.getElementById('gradAngle').value = item.angle;
  document.getElementById('gradAngleVal').textContent = item.angle+'°';
  updateGradientPreview();
  applyGradient();
}

function deleteGradient(id) {
  saveGradLib(getGradLib().filter(i=>i.id!==id));
  loadGradientLibraryUI();
}

// ── OPACITY (live) ────────────────────────────────────
async function applyOpacity(value) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/type");
      await context.sync();
      shapes.items.forEach(s => { try { s.fill.transparency=1-parseFloat(value)/100; } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── BORDER COLOR (live) ───────────────────────────────
async function applyBorderColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/lineFormat/color");
      await context.sync();
      shapes.items.forEach(s => { try { s.lineFormat.color=hex; s.lineFormat.visible=true; } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

// ── BORDER WIDTH (live) ───────────────────────────────
async function applyBorderWidth(value) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/lineFormat/weight");
      await context.sync();
      const pt=parseFloat(value);
      shapes.items.forEach(s => {
        try {
          if(pt===0) s.lineFormat.visible=false;
          else { s.lineFormat.visible=true; s.lineFormat.weight=pt; }
        } catch(e){}
      });
      await context.sync();
    });
  } catch(e){}
}

// ── CONVERT TO ROUND RECT ─────────────────────────────
async function convertToRoundRect() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/visible");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first","err");

      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      let count=0;

      for (const shape of shapes.items) {
        const L=shape.left, T=shape.top, W=shape.width, H=shape.height;
        let fill="#4472C4", lc="#000000", lw=1, lv=false;
        try { fill=shape.fill.foregroundColor||fill; }  catch(e){}
        try { lc=shape.lineFormat.color||lc; }          catch(e){}
        try { lw=shape.lineFormat.weight||lw; }         catch(e){}
        try { lv=shape.lineFormat.visible; }            catch(e){}

        shape.delete();
        await context.sync();

        const ns = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.roundRectangle,
          { left:L, top:T, width:W, height:H }
        );
        ns.fill.setSolidColor(fill);
        ns.lineFormat.color=lc;
        ns.lineFormat.weight=lw;
        ns.lineFormat.visible=lv;
        await context.sync();
        count++;
      }
      showStatus(`✓ Converted ${count} shape(s)`);
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── CORNER RADIUS ─────────────────────────────────────
async function applyCornerRadius() {
  const radiusPt = parseFloat(document.getElementById('radiusInput').value);
  if (isNaN(radiusPt) || radiusPt < 0) return showStatus("Enter valid radius","err");

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/width,items/height");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first","err");

      let applied=0;
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
          }
        } catch(e){}
      }

      if (applied > 0) showStatus(`✓ Radius ${radiusPt}pt applied`);
      else showStatus("Select Round Rectangle first","err");
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── STANDARD SHAPES ───────────────────────────────────
const STANDARD_SHAPES = [
  { name:"Rect",      type:"rectangle",          svg:'<rect x="2" y="4" width="16" height="12" fill="#555"/>' },
  { name:"RoundRect", type:"roundRectangle",     svg:'<rect x="2" y="4" width="16" height="12" rx="3" fill="#555"/>' },
  { name:"Ellipse",   type:"ellipse",            svg:'<ellipse cx="10" cy="10" rx="8" ry="6" fill="#555"/>' },
  { name:"Triangle",  type:"isoscelesTriangle",  svg:'<polygon points="10,2 18,18 2,18" fill="#555"/>' },
  { name:"RtTri",     type:"rightTriangle",      svg:'<polygon points="2,18 18,18 2,2" fill="#555"/>' },
  { name:"Diamond",   type:"diamond",            svg:'<polygon points="10,2 18,10 10,18 2,10" fill="#555"/>' },
  { name:"Pentagon",  type:"regularPentagon",    svg:'<polygon points="10,2 18,7 15,17 5,17 2,7" fill="#555"/>' },
  { name:"Hexagon",   type:"hexagon",            svg:'<polygon points="5,2 15,2 20,10 15,18 5,18 0,10" fill="#555"/>' },
  { name:"Octagon",   type:"octagon",            svg:'<polygon points="6,2 14,2 18,6 18,14 14,18 6,18 2,14 2,6" fill="#555"/>' },
  { name:"Star4",     type:"star4",              svg:'<polygon points="10,2 12,8 18,10 12,12 10,18 8,12 2,10 8,8" fill="#555"/>' },
  { name:"Star5",     type:"star5",              svg:'<polygon points="10,2 12,7 18,7 13,11 15,17 10,13 5,17 7,11 2,7 8,7" fill="#555"/>' },
  { name:"Star6",     type:"star6",              svg:'<polygon points="10,2 13,7 18,6 15,11 17,16 12,14 10,18 8,14 3,16 5,11 2,6 7,7" fill="#555"/>' },
  { name:"Arrow R",   type:"rightArrow",         svg:'<polygon points="2,7 13,7 13,4 18,10 13,16 13,13 2,13" fill="#555"/>' },
  { name:"Arrow L",   type:"leftArrow",          svg:'<polygon points="18,7 7,7 7,4 2,10 7,16 7,13 18,13" fill="#555"/>' },
  { name:"Arrow U",   type:"upArrow",            svg:'<polygon points="10,2 16,8 13,8 13,18 7,18 7,8 4,8" fill="#555"/>' },
  { name:"Arrow D",   type:"downArrow",          svg:'<polygon points="10,18 4,12 7,12 7,2 13,2 13,12 16,12" fill="#555"/>' },
  { name:"Chevron",   type:"chevron",            svg:'<polygon points="2,5 12,5 17,10 12,15 2,15 7,10" fill="#555"/>' },
  { name:"Heart",     type:"heart",              svg:'<path d="M10,16 C6,12 2,9 2,6 A4,4,0,0,1,10,5 A4,4,0,0,1,18,6 C18,9 14,12 10,16Z" fill="#555"/>' },
  { name:"Cross",     type:"plus",               svg:'<rect x="7" y="2" width="6" height="16" fill="#555"/><rect x="2" y="7" width="16" height="6" fill="#555"/>' },
  { name:"Moon",      type:"moon",               svg:'<path d="M13,4 A8,8,0,1,0,13,16 A5,5,0,1,1,13,4Z" fill="#555"/>' },
  { name:"Callout",   type:"wedgeRectCallout",   svg:'<rect x="2" y="2" width="13" height="10" rx="1" fill="#555"/><polygon points="5,12 3,18 9,12" fill="#555"/>' },
  { name:"Donut",     type:"donut",              svg:'<circle cx="10" cy="10" r="8" fill="#555"/><circle cx="10" cy="10" r="4" fill="white"/>' },
  { name:"Parallelgm",type:"parallelogram",      svg:'<polygon points="5,17 2,3 15,3 18,17" fill="#555"/>' },
  { name:"Trapezoid", type:"trapezoid",          svg:'<polygon points="3,17 17,17 14,3 6,3" fill="#555"/>' },
  { name:"Frame",     type:"frame",              svg:'<rect x="2" y="2" width="16" height="16" fill="none" stroke="#555" stroke-width="3"/>' },
];

function renderStandardShapes() {
  const grid = document.getElementById('standardShapeGrid');
  if (!grid) return;
  grid.innerHTML = STANDARD_SHAPES.map((s,i) => `
    <div class="shape-item" onclick="insertStandardShape(${i})" title="${s.name}">
      <svg viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">${s.svg}</svg>
      <span>${s.name}</span>
    </div>`).join('');
}

async function insertStandardShape(index) {
  const s = STANDARD_SHAPES[index];
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");
      await context.sync();

      const shapeType = PowerPoint.GeometricShapeType[s.type];
      if (shapeType === undefined) return showStatus(`${s.name} not supported`,"err");

      const shape = slide.shapes.addGeometricShape(shapeType, {
        left:  (slide.width  - 150) / 2,
        top:   (slide.height - 150) / 2,
        width: 150, height: 150
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
  saveSVGLib(lib); loadSVGLibraryUI();
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
      .replace(/width="[^"]*"/, 'width="26"')
      .replace(/height="[^"]*"/, 'height="26"');
    return `
      <div class="shape-item" onclick="insertSVGFromLibrary(${item.id})" title="${item.name}">
        <div style="width:26px;height:26px;overflow:hidden;display:flex;align-items:center;justify-content:center">${preview}</div>
        <span>${item.name}</span>
        <span class="del-s" onclick="event.stopPropagation();deleteSVGFromLib(${item.id})">×</span>
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

// ── COLOR PALETTE ─────────────────────────────────────
let selectedColor=null;
function getPalette()   { try { return JSON.parse(localStorage.getItem("colorPalette")||"[]"); } catch { return []; } }
function savePalette(p) { localStorage.setItem("colorPalette", JSON.stringify(p)); }

function addColor() {
  const hex  = document.getElementById('newColor').value;
  const name = document.getElementById('newColorName').value||hex;
  const p    = getPalette(); p.push({id:Date.now(),hex,name}); savePalette(p); loadPaletteUI();
  document.getElementById('newColorName').value='';
}
function deleteColor(id) { savePalette(getPalette().filter(c=>c.id!==id)); loadPaletteUI(); }

function loadPaletteUI() {
  const p  = getPalette();
  const mk = (fn,del) => p.map(c=>`
    <div class="swatch" style="background:${c.hex}" title="${c.name}" onclick="${fn}('${c.hex}','${c.name}')">
      ${del?`<span class="del" onclick="event.stopPropagation();deleteColor(${c.id})">×</span>`:''}
    </div>`).join('')||'<span style="font-size:11px;color:#aaa">No colors</span>';
  document.getElementById('paletteGrid').innerHTML      = mk('selectColor',true);
  document.getElementById('applyPaletteGrid').innerHTML = mk('selectColor',false);
}

function selectColor(hex,name) {
  selectedColor=hex;
  document.getElementById('selectedColorPreview').innerHTML=
    `<span style="display:inline-block;width:12px;height:12px;background:${hex};border-radius:2px;margin-right:4px;vertical-align:middle"></span>${name}`;
}

async function applyColorToSelected(type) {
  if(!selectedColor) return showStatus("Click a color first","err");
  if(type==='fill') await applyFillColor(selectedColor);
  else              await applyBorderColor(selectedColor);
}

async function bulkColorReplace() {
  const fh = document.getElementById('findColor').value.replace('#','').toUpperCase();
  const rh = document.getElementById('replaceColor').value;
  try {
    await PowerPoint.run(async(context)=>{
      const slide  = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/fill/foregroundColor,items/type");
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