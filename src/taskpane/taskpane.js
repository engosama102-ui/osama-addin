/* OSAMA DESIGN TOOLS - DIRECT CODE */

Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("Office Ready");
    }
});

// 1. CORNER RADIUS
async function applyCornerRadius(val) {
    var radius = parseFloat(val);
    await PowerPoint.run(async function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/width,items/height,items/adjustments");
        await context.sync();

        for (var i = 0; i < shapes.items.length; i++) {
            var shape = shapes.items[i];
            try {
                var minSide = Math.min(shape.width, shape.height);
                // المعادلة الأصلية
                var adjValue = Math.min(2 * radius / minSide, 0.5);
                shape.adjustments.set(0, adjValue);
            } catch (e) { }
        }
        await context.sync();
    });
}

// 2. OPACITY
async function applyOpacity(val) {
    var trans = 1 - (parseFloat(val) / 100);
    await PowerPoint.run(async function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/fill");
        await context.sync();
        for (var i = 0; i < shapes.items.length; i++) {
            try { shapes.items[i].fill.transparency = trans; } catch (e) { }
        }
        await context.sync();
    });
}

// 3. BORDER WIDTH
async function applyBorderWidth(val) {
    var pt = parseFloat(val);
    await PowerPoint.run(async function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/lineFormat");
        await context.sync();
        for (var i = 0; i < shapes.items.length; i++) {
            try {
                if (pt === 0) shapes.items[i].lineFormat.visible = false;
                else {
                    shapes.items[i].lineFormat.visible = true;
                    shapes.items[i].lineFormat.weight = pt;
                }
            } catch (e) { }
        }
        await context.sync();
    });
}

// 4. FILL COLOR
async function applyFillColor(hex) {
    await PowerPoint.run(async function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/fill");
        await context.sync();
        for (var i = 0; i < shapes.items.length; i++) {
            try { shapes.items[i].fill.setSolidColor(hex); } catch (e) { }
        }
        await context.sync();
    });
}

function onFillColorInput(hex) {
    document.getElementById('fillHex').value = hex.toUpperCase();
    applyFillColor(hex);
}

// 5. CONVERT TO ROUND RECT
async function convertToRoundRect() {
    await PowerPoint.run(async function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor");
        await context.sync();
        var slide = context.presentation.getSelectedSlides().getItemAt(0);
        for (var i = 0; i < shapes.items.length; i++) {
            var s = shapes.items[i];
            var L=s.left, T=s.top, W=s.width, H=s.height, F=s.fill.foregroundColor;
            s.delete();
            var ns = slide.shapes.addGeometricShape("RoundRectangle", { left:L, top:T, width:W, height:H });
            ns.fill.setSolidColor(F);
        }
        await context.sync();
    });
}