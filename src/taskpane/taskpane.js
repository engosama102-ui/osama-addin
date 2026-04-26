/* PowerPoint Add-in Script
   Optimized for macOS & Professional Performance
*/

Office.onReady(() => {
  renderStandardShapes();
  loadSVGLibraryUI();
});

function showStatus(msg, type = 'ok') {
  const el = document.getElementById('status');
  if (!el) return;
  el.textContent = msg;
  el.className = type;
  setTimeout(() => el.className = '', 3000);
}

// ── COLOR HELPERS ─────────────────────────────────────
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
  if (!hex.startsWith('#')) hex = '#' + hex;
  if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
    colorInput.value = hex;
    applyFillColor(hex);
  }
}

async function applyFillColor(hex) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/fill");
      await context.sync();
      shapes.items.forEach(s => {
        try { s.fill.setSolidColor(hex); } catch (e) { }
      });
      await context.sync();
    });
  } catch (e) { }
}

// ── SHAPE CONVERSION (CRITICAL FIX) ───────────────────
async function convertToRoundRect() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/visible");
      await context.sync();

      if (!shapes.items.length) return showStatus("Select shapes first", "err");

      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      
      // Store data first to avoid context loss during deletion
      const pendingShapes = shapes.items.map(s => ({
        l: s.left, t: s.top, w: s.width, h: s.height,
        f: s.fill.foregroundColor,
        lc: s.lineFormat.color,
        lw: s.lineFormat.weight,
        lv: s.lineFormat.visible,
        ref: s
      }));

      for (const data of pendingShapes) {
        data.ref.delete();
        const ns = slide.shapes.addGeometricShape(
          PowerPoint.GeometricShapeType.roundRectangle,
          { left: data.l, top: data.t, width: data.w, height: data.h }
        );
        ns.fill.setSolidColor(data.f || "#4472C4");
        if (data.lv) {
          ns.lineFormat.color = data.lc;
          ns.lineFormat.weight = data.lw;
        } else {
          ns.lineFormat.visible = false;
        }
      }
      await context.sync();
      showStatus(`✓ Converted ${pendingShapes.length} shape(s)`);
    });
  } catch (e) { showStatus("Error: " + e.message, "err"); }
}

// ── CORNER RADIUS (FIXED FOR SLIDER) ──────────────────
async function applyCornerRadius(radiusPt) {
  const radius = parseFloat(radiusPt);
  if (isNaN(radius) || radius < 0) return;

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      // Must load geometricShapeType to verify it is a RoundRectangle
      shapes.load("items/width,items/height,items/adjustments,items/geometricShapeType");
      await context.sync();

      let count = 0;
      shapes.items.forEach(shape => {
        if (shape.geometricShapeType.toLowerCase() === "roundrectangle") {
          const shortSide = Math.min(shape.width, shape.height);
          // Adjustment logic: 0.0 (sharp) to 0.5 (max round)
          let adjValue = radius / shortSide;
          if (adjValue > 0.5) adjValue = 0.5;
          
          shape.adjustments.set(0, adjValue);
          count++;
        }
      });

      await context.sync();
    });
  } catch (e) { console.error(e); }
}

async function applyCornerRadiusToAll(radiusPt) {
  const radius = parseFloat(radiusPt);
  if (isNaN(radius) || radius < 0) return;

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/width,items/height,items/adjustments,items/geometricShapeType");
      await context.sync();

      let count = 0;
      shapes.items.forEach(shape => {
        if (shape.geometricShapeType.toLowerCase() === "roundrectangle") {
          const shortSide = Math.min(shape.width, shape.height);
          let adjValue = Math.min(radius / shortSide, 0.5);
          shape.adjustments.set(0, adjValue);
          count++;
        }
      });
      await context.sync();
      showStatus(`✓ Updated ${count} shapes`);
    });
  } catch (e) { showStatus(e.message, "err"); }
}

// ── SVG & LIBRARY LOGIC ───────────────────────────────
async function insertSVGCode(svgCode) {
  if (!svgCode || !svgCode.includes('<svg')) return showStatus("Invalid SVG", "err");
  
  try {
    const base64 = btoa(unescape(encodeURIComponent(svgCode)));
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.load("width,height");
      await context.sync();
      
      const image = slide.shapes.addImage(base64);
      image.width = 150;
      image.height = 150;
      image.left = (slide.width - 150) / 2;
      image.top = (slide.height - 150) / 2;
      
      await context.sync();
      showStatus("✓ SVG Inserted");
    });
  } catch (e) {
    showStatus("Insertion failed", "err");
  }
}

function getSVGLib() { 
  try { return JSON.parse(localStorage.getItem("svgLibrary") || "[]"); } 
  catch { return []; } 
}

function saveSVGLib(lib) { localStorage.setItem("svgLibrary", JSON.stringify(lib)); }

function loadSVGLibraryUI() {
  const lib = getSVGLib();
  const grid = document.getElementById('svgLibraryGrid');
  if (!grid) return;
  
  if (!lib.length) {
    grid.innerHTML = '<div class="empty-msg">Library Empty</div>';
    return;
  }

  grid.innerHTML = lib.map(item => `
    <div class="shape-item" onclick="insertSVGFromLibrary(${item.id})">
      <div class="preview">${item.code}</div>
      <span>${item.name}</span>
      <span class="del-btn" onclick="event.stopPropagation();deleteSVG(${item.id})">×</span>
    </div>
  `).join('');
}

async function insertSVGFromLibrary(id) {
  const item = getSVGLib().find(i => i.id === id);
  if (item) await insertSVGCode(item.code);
}

function deleteSVG(id) {
  const filtered = getSVGLib().filter(i => i.id !== id);
  saveSVGLib(filtered);
  loadSVGLibraryUI();
}

// ── BULK TOOLS ────────────────────────────────────────
async function bulkColorReplace() {
  const findHex = document.getElementById('findColor').value.toUpperCase().replace('#','');
  const replaceHex = document.getElementById('replaceColor').value;

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load("items/fill/foregroundColor");
      await context.sync();

      let count = 0;
      shapes.items.forEach(s => {
        const currentHex = (s.fill.foregroundColor || "").toUpperCase().replace('#','');
        if (currentHex === findHex) {
          s.fill.setSolidColor(replaceHex);
          count++;
        }
      });
      await context.sync();
      showStatus(`✓ Replaced ${count} items`);
    });
  } catch (e) { showStatus(e.message, "err"); }
}