/* OSAMA ADDIN - RECOVERY MODE */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // السطر ده هيخليك تتأكد إن الكود اتحدث، لو ظهرت الرسالة يعني الكود وصل
    console.log("Add-in version 2.0 loaded");
    if (typeof loadSVGLibraryUI === "function") loadSVGLibraryUI();
  }
});

// 1. CORNER RADIUS (النسخة الأصلية المضمونة)
async function applyCornerRadius(val) {
  const radius = parseFloat(val);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/width,items/height,items/adjustments,items/type");
    await context.sync();

    shapes.items.forEach(shape => {
      try {
        // فحص بسيط: إذا كان شكل هندسي وله adjustments
        if (shape.adjustments) {
          const minSide = Math.min(shape.width, shape.height);
          const adjValue = Math.min(2 * radius / minSide, 0.5);
          shape.adjustments.set(0, adjValue);
        }
      } catch (e) {}
    });
    return context.sync();
  });
}

// 2. OPACITY
async function applyOpacity(val) {
  const trans = 1 - (parseFloat(val) / 100);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.transparency = trans; } catch(e){} });
    return context.sync();
  });
}

// 3. BORDER WIDTH
async function applyBorderWidth(val) {
  const pt = parseFloat(val);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/lineFormat");
    await context.sync();
    shapes.items.forEach(s => {
      try {
        if (pt === 0) s.lineFormat.visible = false;
        else { s.lineFormat.visible = true; s.lineFormat.weight = pt; }
      } catch(e){}
    });
    return context.sync();
  });
}

// 4. FILL COLOR
async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
    return context.sync();
  });
}

function onFillColorInput(hex) {
  const field = document.getElementById('fillHex');
  if (field) field.value = hex.toUpperCase();
  applyFillColor(hex);
}

// 5. CONVERT TO ROUND RECT
async function convertToRoundRect() {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
    await context.sync();
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    shapes.items.forEach(s => {
      const L=s.left, T=s.top, W=s.width, H=s.height, F=s.fill.foregroundColor;
      s.delete();
      const ns = slide.shapes.addGeometricShape("RoundRectangle", { left:L, top:T, width:W, height:H });
      ns.fill.setSolidColor(F);
    });
    return context.sync();
  });
}