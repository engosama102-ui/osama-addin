/* OSAMA ADDIN - Professional Recovery Version */

Office.onReady(() => {
  console.log("Add-in Ready");
  // استدعاء الدوال الأساسية فقط لو موجودة
  if (typeof loadSVGLibraryUI === "function") loadSVGLibraryUI();
});

// ── UNIVERSAL STATUS ──────────────────────────────────
function showStatus(msg, type='ok') {
  const el = document.getElementById('status');
  if (el) {
    el.textContent = msg; el.className = type;
    setTimeout(() => el.className='', 3000);
  }
}

// ── CORNER RADIUS (FIXED) ─────────────────────────────
async function applyCornerRadius(radiusPt) {
  const radius = parseFloat(radiusPt);
  if (isNaN(radius)) return;

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/width,items/height,items/type,items/adjustments");
    await context.sync();

    shapes.items.forEach((shape) => {
      try {
        // فحص بسيط عشان نضمن إننا بنعدل على شكل هندسي فقط
        if (shape.type === "GeometricShape" || shape.adjustments) {
          const shortSide = Math.min(shape.width, shape.height);
          if (shortSide > 0) {
            // المعادلة الأصلية اللي كانت شغالة معاك
            const adjValue = Math.min(2 * radius / shortSide, 0.5);
            shape.adjustments.set(0, adjValue);
          }
        }
      } catch (e) { /* تجاهل الأخطاء الفردية لكل شكل */ }
    });
    return context.sync();
  }).catch(err => console.log("Radius Error: " + err.message));
}

// ── OPACITY (FIXED) ───────────────────────────────────
async function applyOpacity(value) {
  const trans = 1 - (parseFloat(value) / 100);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/fill");
    await context.sync();

    shapes.items.forEach((shape) => {
      try {
        shape.fill.transparency = trans;
      } catch (e) { }
    });
    return context.sync();
  }).catch(err => console.log("Opacity Error"));
}

// ── BORDER WIDTH (FIXED) ──────────────────────────────
async function applyBorderWidth(value) {
  const pt = parseFloat(value);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/lineFormat");
    await context.sync();

    shapes.items.forEach((shape) => {
      try {
        if (pt === 0) {
          shape.lineFormat.visible = false;
        } else {
          shape.lineFormat.visible = true;
          shape.lineFormat.weight = pt;
        }
      } catch (e) { }
    });
    return context.sync();
  }).catch(err => console.log("Border Error"));
}

// ── BORDER COLOR (FIXED) ──────────────────────────────
async function applyBorderColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/lineFormat");
    await context.sync();

    shapes.items.forEach((shape) => {
      try {
        shape.lineFormat.color = hex;
        shape.lineFormat.visible = true;
      } catch (e) { }
    });
    return context.sync();
  });
}

// ── FILL COLOR ────────────────────────────────────────
async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/fill");
    await context.sync();
    shapes.items.forEach(s => {
      try { s.fill.setSolidColor(hex); } catch(e){}
    });
    await context.sync();
  });
}

function onFillColorInput(hex) {
  const field = document.getElementById('fillHex');
  if (field) field.value = hex.toUpperCase();
  applyFillColor(hex);
}

// ── CONVERSION TOOL (STABLE) ──────────────────────────
async function convertToRoundRect() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
      await context.sync();

      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      for (const shape of shapes.items) {
        const L=shape.left, T=shape.top, W=shape.width, H=shape.height, F=shape.fill.foregroundColor;
        shape.delete();
        const ns = slide.shapes.addGeometricShape("RoundRectangle", { left:L, top:T, width:W, height:H });
        ns.fill.setSolidColor(F);
      }
      await context.sync();
      showStatus("✓ Converted");
    });
  } catch(e) { showStatus("Select shape first","err"); }
}