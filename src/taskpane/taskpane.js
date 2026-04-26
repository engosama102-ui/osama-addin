/* FIXED FOR WEBPACK - OSAMA ADDIN */
Office.onReady(() => {
  if (typeof loadSVGLibraryUI === "function") {
    loadSVGLibraryUI();
  }
});

// --- HELPER ---
function getShapes(context) {
  return context.presentation.getSelectedShapes();
}

// --- RADIUS (THE FIX) ---
export async function applyCornerRadius(radiusPt) {
  const radius = parseFloat(radiusPt);
  if (isNaN(radius)) return;
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/width,items/height,items/adjustments,items/type");
    await context.sync();
    shapes.items.forEach(shape => {
      try {
        const minSide = Math.min(shape.width, shape.height);
        const adjValue = Math.min(2 * radius / minSide, 0.5);
        shape.adjustments.set(0, adjValue);
      } catch (e) {}
    });
    await context.sync();
  });
}

export async function applyCornerRadiusToAll(radiusPt) {
  const radius = parseFloat(radiusPt);
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/width,items/height,items/adjustments");
    await context.sync();
    shapes.items.forEach(shape => {
      try {
        const minSide = Math.min(shape.width, shape.height);
        shape.adjustments.set(0, Math.min(2 * radius / minSide, 0.5));
      } catch (e) {}
    });
    await context.sync();
  });
}

// --- FILL & OPACITY ---
export async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
    await context.sync();
  });
}

export function onFillColorInput(hex) {
  const field = document.getElementById('fillHex');
  if (field) field.value = hex.toUpperCase();
  applyFillColor(hex);
}

export function syncFillHexInput() {
  let hex = document.getElementById('fillHex').value.trim();
  if (!hex.startsWith('#')) hex = '#' + hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    applyFillColor(hex);
  }
}

export async function applyNoFill() {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { s.fill.transparency = 1; });
    await context.sync();
  });
}

export async function applyOpacity(val) {
  const trans = 1 - (parseFloat(val) / 100);
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { s.fill.transparency = trans; });
    await context.sync();
  });
}

// --- BORDER ---
export async function applyBorderColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/lineFormat");
    await context.sync();
    shapes.items.forEach(s => { s.lineFormat.color = hex; s.lineFormat.visible = true; });
    await context.sync();
  });
}

export async function applyBorderWidth(val) {
  const pt = parseFloat(val);
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/lineFormat");
    await context.sync();
    shapes.items.forEach(s => {
      s.lineFormat.visible = pt > 0;
      if (pt > 0) s.lineFormat.weight = pt;
    });
    await context.sync();
  });
}

// --- CONVERSION & SVG ---
export async function convertToRoundRect() {
  await PowerPoint.run(async (context) => {
    const shapes = getShapes(context);
    shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
    await context.sync();
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    shapes.items.forEach(s => {
      const L=s.left, T=s.top, W=s.width, H=s.height, F=s.fill.foregroundColor;
      s.delete();
      const ns = slide.shapes.addGeometricShape("RoundRectangle", { left:L, top:T, width:W, height:H });
      ns.fill.setSolidColor(F);
    });
    await context.sync();
  });
}

export async function insertSVGToSlide() {
  const code = document.getElementById('svg-input').value;
  const base64 = btoa(unescape(encodeURIComponent(code)));
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.addImage(base64);
    await context.sync();
  });
}

// Placeholder for missing UI functions to stop Webpack errors
export function addSVGToLibrary() {}
export function bulkColorReplace() {}
export function loadSVGLibraryUI() {}