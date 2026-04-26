Office.onReady(() => {
  loadSVGLibraryUI();
  renderFillColorWindow();
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
  addRecentFillColor(hex);
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
          try { s.fill.setSolidColor(c1); } catch(e2){}
        }
      });
      await context.sync();
      showStatus("✓ Gradient applied");
    });
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

function getGradLib()     { try { return JSON.parse(localStorage.getItem("gradLibrary")||"[]"); } catch { return []; } }
function saveGradLib(lib) { localStorage.setItem("gradLibrary", JSON.stringify(lib)); }

function saveGradient() {
  const lib  = getGradLib();
  const name = `Gradient ${lib.length+1}`;
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
  if (!el) return;
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

async function applyCornerRadius(radiusPt) {
  radiusPt = parseFloat(radiusPt);
  if (isNaN(radiusPt) || radiusPt < 0) return;
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/width,items/height,items/type");
      await context.sync();
      if (!shapes.items.length) return;
      let applied = 0;
      for (const shape of shapes.items) {
        try {
          if (shape.type !== PowerPoint.ShapeType.geometricShape) continue;
          const shortSide = Math.min(shape.width, shape.height);
          if (shortSide <= 0) continue;
          const adjValue = Math.min(2 * radiusPt / shortSide, 0.5);
          shape.adjustments.load("items");
          await context.sync();
          if (shape.adjustments.items && shape.adjustments.items.length > 0) {
            shape.adjustments.items[0].value = adjValue;
            await context.sync();
            applied++;
          }
        } catch (e) { continue; }
      }
      if (applied === 0) showStatus("Select Round Rectangle first","err");
      else showStatus(`✓ Radius applied to ${applied} shape(s)`);
    });
  } catch(e) {}
}

async function applyCornerRadiusToAll(radiusPt) {
  radiusPt = parseFloat(radiusPt);
  if (isNaN(radiusPt) || radiusPt < 0) return;
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/type,items/width,items/height");
      await context.sync();
      if (!shapes.items.length) return showStatus("No shapes on slide","err");
      let applied = 0;
      for (const shape of shapes.items) {
        try {
          if (shape.type !== PowerPoint.ShapeType.geometricShape) continue;
          const shortSide = Math.min(shape.width, shape.height);
          if (shortSide <= 0) continue;
          const adjValue = Math.min(2 * radiusPt / shortSide, 0.5);
          shape.adjustments.load("items");
          await context.sync();
          if (shape.adjustments.items && shape.adjustments.items.length > 0) {
            shape.adjustments.items[0].value = adjValue;
            await context.sync();
            applied++;
          }
        } catch (e) { continue; }
      }
      if (applied === 0) showStatus("No round shapes found","err");
      else showStatus(`✓ Applied to ${applied} shape(s)`);
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

// ── COLOR ─────────────────────────────────────────────
function syncFillHexInput() {
  const input = document.getElementById('fillHex');
  const colorInput = document.getElementById('fillColor');
  if (!input || !colorInput) return;
  let hex = input.value.trim();
  if (!hex.startsWith('#')) hex = '#'+hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    colorInput.value = hex;
    applyFillColor(hex);
    addRecentFillColor(hex);
  }
}

function openMoreFillColors() {
  const colorInput = document.getElementById('fillColor');
  if (colorInput) colorInput.click();
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

const DEFAULT_THEME_COLORS = ["#FFFFFF","#000000","#4472C4","#ED7D31","#A5A5A5","#FFC000","#5B9BD5","#70AD47"];
const STANDARD_COLORS = ["#C00000","#FF0000","#FFC000","#FFFF00","#92D050","#00B050","#00B0F0","#0070C0","#002060","#7030A0"];

function getThemeColors() {
  try { return JSON.parse(localStorage.getItem("themeColors")) || DEFAULT_THEME_COLORS; }
  catch { return DEFAULT_THEME_COLORS; }
}
function saveThemeColors(colors) { localStorage.setItem("themeColors", JSON.stringify(colors)); }

let themeColorEditIndex = null;

function openThemeColorPicker(index) {
  themeColorEditIndex = index;
  const picker = document.getElementById('themeColorPicker');
  const colors = getThemeColors();
  if (!picker) return;
  picker.value = colors[index] || "#000000";
  picker.click();
}

function saveThemeColorFromPicker(value) {
  if (themeColorEditIndex === null) return;
  const colors = getThemeColors();
  colors[themeColorEditIndex] = value.toUpperCase();
  saveThemeColors(colors);
  renderFillColorWindow();
  themeColorEditIndex = null;
}

function deleteThemeColor(index) {
  const colors = getThemeColors();
  colors.splice(index, 1);
  saveThemeColors(colors);
  renderFillColorWindow();
}

// ── FIX: addThemeColor بدون prompt ────────────────────
function addThemeColor() {
  const colors = getThemeColors();
  colors.push("#4472C4");
  saveThemeColors(colors);
  renderFillColorWindow();
  // فتح الـ picker للون الجديد مباشرة
  themeColorEditIndex = colors.length - 1;
  const picker = document.getElementById('themeColorPicker');
  if (picker) {
    picker.value = "#4472C4";
    picker.click();
  }
}

function renderFillColorWindow() {
  const themeGrid    = document.getElementById('themeColorGrid');
  const standardGrid = document.getElementById('standardColorGrid');
  const recentGrid   = document.getElementById('recentColorGrid');

  const themeColors = getThemeColors();
  if (themeGrid) themeGrid.innerHTML = themeColors.map((hex, i) => `
    <div class="color-swatch" style="background:${hex}" title="Click to apply | Long press to edit"
      onclick="setFillColor('${hex}')"
      oncontextmenu="event.preventDefault();openThemeColorPicker(${i})">
      <div class="delete-btn" onclick="event.stopPropagation();deleteThemeColor(${i})">×</div>
    </div>`).join('');

  if (standardGrid) standardGrid.innerHTML = STANDARD_COLORS.map(hex => `
    <div class="color-swatch" style="background:${hex}" title="${hex}" onclick="setFillColor('${hex}')"></div>`).join('');

  const recent = JSON.parse(localStorage.getItem("recentFillColors")||"[]");
  if (recentGrid) {
    recentGrid.innerHTML = recent.length
      ? recent.map(hex => `<div class="color-swatch" style="background:${hex}" title="${hex}" onclick="setFillColor('${hex}')"></div>`).join('')
      : `<div class="color-swatch add-border" onclick="openMoreFillColors()">More</div>`;
  }
}

function setFillColor(hex) {
  const input    = document.getElementById('fillColor');
  const hexInput = document.getElementById('fillHex');
  if (input)    input.value    = hex;
  if (hexInput) hexInput.value = hex;
  applyFillColor(hex);
  addRecentFillColor(hex);
}

function addRecentFillColor(hex) {
  if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return;
  const recent  = JSON.parse(localStorage.getItem("recentFillColors")||"[]");
  const updated = [hex, ...recent.filter(c=>c!==hex)].slice(0,8);
  localStorage.setItem("recentFillColors", JSON.stringify(updated));
  renderFillColorWindow();
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