Office.onReady(() => {
  loadSVGLibraryUI();
});

function showStatus(msg, type='ok') {
  const el = document.getElementById('status');
  el.textContent = msg; el.className = type;
  setTimeout(() => el.className='', 3000);
}

function onFillColorInput(hex) {
  const field = document.getElementById('fillHex');
  if (field) field.value = hex;
  applyFillColor(hex);
}

function syncFillHexInput() {
  const input = document.getElementById('fillHex');
  const colorInput = document.getElementById('fillColor');
  if (!input || !colorInput) return;
  let hex = input.value.trim();
  if (!hex.startsWith('#')) hex = '#'+hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    colorInput.value = hex;
    applyFillColor(hex);
  }
}

function syncBorderHexInput() {
  const input = document.getElementById('borderHex');
  const colorInput = document.getElementById('borderColor');
  if (!input || !colorInput) return;
  let hex = input.value.trim();
  if (!hex.startsWith('#')) hex = '#'+hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    colorInput.value = hex;
    applyBorderColor(hex);
  }
}

async function applyFillColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/type");
      await context.sync();
      shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
      await context.sync();
    });
  } catch(e) {}
}

async function applyNoFill() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill/type");
      await context.sync();
      shapes.items.forEach(s => {
        try { s.fill.transparency = 1; } catch(e){}
        try { s.fill.setSolidColor("#FFFFFF"); } catch(e){}
      });
      await context.sync();
      showStatus("✓ No fill applied");
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

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

async function applyBorderColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/lineFormat/visible");
      await context.sync();
      shapes.items.forEach(s => { try { s.lineFormat.color=hex; s.lineFormat.visible=true; } catch(e){} });
      await context.sync();
    });
  } catch(e){}
}

async function applyBorderWidth(value) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/lineFormat/visible");
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
async function applyCornerRadius(radiusPt) {
  radiusPt = parseFloat(radiusPt);
  if (isNaN(radiusPt) || radiusPt < 0) return;

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height,items/type,items/adjustments");
      await context.sync();

      if (!shapes.items.length) return;

      let applied = 0;
      for (const shape of shapes.items) {
        try {
          if (shape.type !== PowerPoint.ShapeType.geometricShape) continue;
          const shortSide = Math.min(shape.width, shape.height);
          if (shortSide <= 0) continue;

          const adjValue = Math.min(2 * radiusPt / shortSide, 0.5);
          shape.adjustments.set(0, adjValue);
          applied++;
        } catch (e) {
          continue;
        }
      }

      if (applied === 0) return showStatus("No round rectangle selected","err");
      await context.sync();
    });
  } catch(e) {}
}

// ── APPLY RADIUS TO ALL SHAPES ────────────────────────
async function applyCornerRadiusToAll(radiusPt) {
  radiusPt = parseFloat(radiusPt);
  if (isNaN(radiusPt) || radiusPt < 0) return;

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/type,items/width,items/height,items/adjustments");
      await context.sync();

      if (!shapes.items.length) return showStatus("No shapes on slide","err");

      let applied = 0;
      for (const shape of shapes.items) {
        try {
          if (shape.type !== PowerPoint.ShapeType.geometricShape) continue;
          const shortSide = Math.min(shape.width, shape.height);
          if (shortSide <= 0) continue;

          const adjValue = Math.min(2 * radiusPt / shortSide, 0.5);
          shape.adjustments.set(0, adjValue);
          applied++;
        } catch (e) {
          continue;
        }
      }

      if (applied === 0) return showStatus("No round shapes found","err");
      await context.sync();
      showStatus(`✓ Applied radius to ${applied} shape(s)`);
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

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

function getSVGLib()     { try { return JSON.parse(localStorage.getItem("svgLibrary")||"[]"); } catch { return []; } }
function saveSVGLib(lib) { localStorage.setItem("svgLibrary", JSON.stringify(lib)); }

function addSVGToLibrary() {
  const code = document.getElementById('svg-input').value.trim();
  if (!code || !code.includes('<svg')) return showStatus("Paste valid SVG first","err");
  let normalizedCode = code;
  const svgMatch = code.match(/<svg[^>]*>/i);
  if (svgMatch) {
    let svgTag = svgMatch[0];
    if (!svgTag.includes('viewBox')) {
      const wm = svgTag.match(/width=[\"']([^\"']*)[\"']/);
      const hm = svgTag.match(/height=[\"']([^\"']*)[\"']/);
      if (wm && hm) {
        svgTag = svgTag.replace('>', ` viewBox="0 0 ${parseFloat(wm[1])} ${parseFloat(hm[1])}">`);
        normalizedCode = normalizedCode.replace(svgMatch[0], svgTag);
      }
    }
    normalizedCode = normalizedCode.replace(/\s*width=[\"'][^\"']*[\"']/i, '');
    normalizedCode = normalizedCode.replace(/\s*height=[\"'][^\"']*[\"']/i, '');
  }
  const name = document.getElementById('svgName').value.trim() || `SVG ${getSVGLib().length+1}`;
  const lib  = getSVGLib();
  lib.push({ id:Date.now(), name, code:normalizedCode });
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

async function insertSVGCode(svgCode) {
  if (!svgCode || !svgCode.includes('<svg')) return showStatus("Paste valid SVG first","err");
  let base64;
  try {
    if (!svgCode.includes('viewBox')) {
      const wm = svgCode.match(/width=["']([^"']*)["']/);
      const hm = svgCode.match(/height=["']([^"']*)["']/);
      if (wm && hm) {
        svgCode = svgCode.replace(/<svg/, `<svg viewBox="0 0 ${parseFloat(wm[1])||100} ${parseFloat(hm[1])||100}"`);
      }
    }
    if (!svgCode.includes('preserveAspectRatio')) {
      svgCode = svgCode.replace(/<svg/, `<svg preserveAspectRatio="xMidYMid meet"`);
    }
    svgCode = svgCode.replace(/width=["'][^"']*["']/i, 'width="200"');
    svgCode = svgCode.replace(/height=["'][^"']*["']/i, 'height="200"');
    if (!svgCode.match(/width=/i)) {
      svgCode = svgCode.replace(/<svg/, '<svg width="200" height="200"');
    }
    base64 = btoa(unescape(encodeURIComponent(svgCode)));
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");
      await context.sync();
      const image = slide.shapes.addImage(base64);
      image.left = Math.max(0, (slide.width - 200) / 2);
      image.top  = Math.max(0, (slide.height - 200) / 2);
      image.width  = 200;
      image.height = 200;
      await context.sync();
      showStatus("✓ SVG inserted");
    });
  } catch(e) {
    if (base64) {
      try {
        Office.context.document.setSelectedDataAsync(
          base64,
          { coercionType:Office.CoercionType.Image, imageLeft:100, imageTop:100, imageWidth:200, imageHeight:200 },
          r => {
            if (r.status===Office.AsyncResultStatus.Failed) showStatus("Error: "+r.error.message,"err");
            else showStatus("✓ SVG inserted");
          }
        );
        return;
      } catch(inner) {}
    }
    showStatus("Error: "+e.message,"err");
  }
}

async function insertSVGToSlide() {
  const code = document.getElementById('svg-input').value.trim();
  if (!code) return showStatus("Paste SVG code first","err");
  await insertSVGCode(code);
}

async function insertSVGFromLibrary(id) {
  const item = getSVGLib().find(i=>i.id===id);
  if (!item) return;
  await insertSVGCode(item.code);
}

function deleteSVGFromLib(id) { saveSVGLib(getSVGLib().filter(i=>i.id!==id)); loadSVGLibraryUI(); }

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