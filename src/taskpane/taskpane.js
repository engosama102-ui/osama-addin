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
async function insertSVGCode(svgCode) {
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
    svg = svg.replace(/(<svg[^>]*)\swidth=["'][^"']*["']/i, '$1');
    svg = svg.replace(/(<svg[^>]*)\sheight=["'][^"']*["']/i, '$1');
    svg = svg.replace(/<svg/, '<svg width="200" height="200"');

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
    showStatus('✓ SVG inserted! Right-click → Convert to Shape', 'ok');
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