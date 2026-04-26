/* OSAMA ADDIN - Final Stable Version */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("OSAMA ADDIN Ready");
    }
});

function showStatus(msg, type = 'ok') {
    const el = document.getElementById('status');
    if (!el) return;
    el.textContent = msg;
    el.className = type;
    setTimeout(() => el.className = '', 3000);
}

// ── CORNER RADIUS (THE FIX) ───────────────────────────
async function applyCornerRadius(radiusPt) {
    const radius = parseFloat(radiusPt);
    if (isNaN(radius)) return;

    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/width,items/height,items/adjustments,items/geometricShapeType");
        await context.sync();

        if (shapes.items.length === 0) return;

        shapes.items.forEach((shape) => {
            const type = shape.geometricShapeType.toLowerCase();
            // Check for Round Rectangle
            if (type === "roundrectangle" || type === "roundrect") {
                const shortSide = Math.min(shape.width, shape.height);
                // Equation: radius points / short side = ratio (0 to 0.5)
                let ratio = radius / shortSide;
                
                if (ratio > 0.5) ratio = 0.5;
                if (ratio < 0) ratio = 0;

                shape.adjustments.set(0, ratio);
            }
        });
        return context.sync(); // Force UI Update
    }).catch(err => console.error(err));
}

async function applyCornerRadiusToAll(radiusPt) {
    const radius = parseFloat(radiusPt);
    await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items/width,items/height,items/adjustments,items/geometricShapeType");
        await context.sync();

        shapes.items.forEach((shape) => {
            if (shape.geometricShapeType.toLowerCase() === "roundrectangle") {
                const shortSide = Math.min(shape.width, shape.height);
                let ratio = Math.min(radius / shortSide, 0.5);
                shape.adjustments.set(0, ratio);
            }
        });
        await context.sync();
        showStatus("Applied to all shapes");
    });
}

// ── COLOR & FILL ──────────────────────────────────────
function onFillColorInput(hex) {
    document.getElementById('fillHex').value = hex.toUpperCase();
    applyFillColor(hex);
}

function syncFillHexInput() {
    let hex = document.getElementById('fillHex').value.trim();
    if (!hex.startsWith('#')) hex = '#' + hex;
    if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
        document.getElementById('fillColor').value = hex;
        applyFillColor(hex);
    }
}

async function applyFillColor(hex) {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/fill");
        await context.sync();
        shapes.items.forEach(s => s.fill.setSolidColor(hex));
        await context.sync();
    });
}

async function applyNoFill() {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/fill");
        await context.sync();
        shapes.items.forEach(s => s.fill.transparency = 1);
        await context.sync();
        showStatus("No fill applied");
    });
}

async function applyOpacity(val) {
    const trans = 1 - (parseFloat(val) / 100);
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/fill");
        await context.sync();
        shapes.items.forEach(s => s.fill.transparency = trans);
        await context.sync();
    });
}

// ── BORDER ────────────────────────────────────────────
async function applyBorderColor(hex) {
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/lineFormat");
        await context.sync();
        shapes.items.forEach(s => {
            s.lineFormat.color = hex;
            s.lineFormat.visible = true;
        });
        await context.sync();
    });
}

async function applyBorderWidth(val) {
    const weight = parseFloat(val);
    await PowerPoint.run(async (context) => {
        const shapes = context.presentation.getSelectedShapes();
        shapes.load("items/lineFormat");
        await context.sync();
        shapes.items.forEach(s => {
            if (weight === 0) s.lineFormat.visible = false;
            else {
                s.lineFormat.visible = true;
                s.lineFormat.weight = weight;
            }
        });
        await context.sync();
    });
}

// ── CONVERSION ────────────────────────────────────────
async function convertToRoundRect() {
    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/visible");
            await context.sync();

            if (shapes.items.length === 0) return showStatus("Select shapes first", "err");

            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const data = shapes.items.map(s => ({
                l: s.left, t: s.top, w: s.width, h: s.height,
                f: s.fill.foregroundColor, lc: s.lineFormat.color,
                lw: s.lineFormat.weight, lv: s.lineFormat.visible,
                ref: s
            }));

            for (const item of data) {
                item.ref.delete();
                const ns = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
                    left: item.l, top: item.t, width: item.w, height: item.h
                });
                ns.fill.setSolidColor(item.f || "#4472C4");
                ns.lineFormat.visible = item.lv;
                if (item.lv) {
                    ns.lineFormat.color = item.lc;
                    ns.lineFormat.weight = item.lw;
                }
            }
            await context.sync();
            showStatus("Converted to Round Rect");
        });
    } catch (e) { showStatus(e.message, "err"); }
}