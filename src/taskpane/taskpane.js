/* OSAMA DESIGN TOOLS */

Office.onReady(() => {});

function getShapes(context) {
  return context.presentation.getSelectedShapes();
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

// --- SVG ---
async function insertSVGToSlide() {
  const code = document.getElementById('svg-input').value;
  const base64 = btoa(unescape(encodeURIComponent(code)));
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.addImage(base64);
    await context.sync();
  });
}

// Expose all functions globally so HTML event handlers can call them
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
