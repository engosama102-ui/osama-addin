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

// --- SVG: convert to PNG via canvas, then insert ---
function svgToPngBase64(svgCode) {
  return new Promise((resolve, reject) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgCode, 'image/svg+xml');
    const svgEl = doc.querySelector('svg');
    const w = parseInt(svgEl && svgEl.getAttribute('width')) || 200;
    const h = parseInt(svgEl && svgEl.getAttribute('height')) || 200;

    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');

    const blob = new Blob([svgCode], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = function () {
      ctx.drawImage(img, 0, 0, w, h);
      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL('image/png').split(',')[1]);
    };
    img.onerror = function () {
      URL.revokeObjectURL(url);
      reject(new Error('SVG load failed'));
    };
    img.src = url;
  });
}

async function insertSVGCode(svgCode) {
  try {
    const base64 = await svgToPngBase64(svgCode);
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error.message));
        }
      );
    });
    showStatus('SVG inserted!', 'ok');
  } catch (e) {
    showStatus('Insert failed: ' + e.message, 'err');
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
