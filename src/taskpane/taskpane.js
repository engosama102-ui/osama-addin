/* OSAMA DESIGN TOOLS - CLEAN RECOVERY */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Add-in Ready");
    }
});

// ── CORNER RADIUS (Restored logic) ──
async function applyCornerRadius(radiusPt) {
    const radius = parseFloat(radiusPt);
    if (isNaN(radius)) return;

    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/width,items/height,items/type,items/adjustments");
            await context.sync();

            shapes.items.forEach(shape => {
                try {
                    // نستخدم نفس فحص النوع الذي كان يعمل عندك سابقاً
                    if (shape.type === "GeometricShape" || shape.adjustments) {
                        const minSide = Math.min(shape.width, shape.height);
                        if (minSide > 0) {
                            // المعادلة الأصلية 2 * r / side
                            const adjValue = Math.min(2 * radius / minSide, 0.5);
                            shape.adjustments.set(0, adjValue);
                        }
                    }
                } catch (e) { console.log("Shape error skipped"); }
            });
            return context.sync();
        });
    } catch (err) { console.log("Radius Error: " + err.message); }
}

// ── OPACITY ──
async function applyOpacity(val) {
    const trans = 1 - (parseFloat(val) / 100);
    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/fill");
            await context.sync();
            shapes.items.forEach(s => { try { s.fill.transparency = trans; } catch(e){} });
            return context.sync();
        });
    } catch(e) {}
}

// ── BORDER WIDTH ──
async function applyBorderWidth(val) {
    const pt = parseFloat(val);
    try {
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
    } catch(e) {}
}

// ── FILL COLOR ──
async function applyFillColor(hex) {
    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/fill");
            await context.sync();
            shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
            return context.sync();
        });
    } catch(e) {}
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
        document.getElementById('fillColor').value = hex;
        applyFillColor(hex);
    }
}

// ── CONVERT TO ROUND RECT ──
async function convertToRoundRect() {
    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
            await context.sync();

            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            for (const s of shapes.items) {
                const L=s.left, T=s.top, W=s.width, H=s.height, F=s.fill.foregroundColor;
                s.delete();
                const ns = slide.shapes.addGeometricShape("RoundRectangle", { left:L, top:T, width:W, height:H });
                ns.fill.setSolidColor(F);
            }
            return context.sync();
        });
    } catch(e) {}
}

// ── BORDER COLOR ──
async function applyBorderColor(hex) {
    try {
        await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load("items/lineFormat");
            await context.sync();
            shapes.items.forEach(s => { try { s.lineFormat.color = hex; s.lineFormat.visible = true; } catch(e){} });
            return context.sync();
        });
    } catch(e) {}
}