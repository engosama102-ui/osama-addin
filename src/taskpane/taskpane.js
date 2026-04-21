Office.onReady(() => {
  loadPaletteUI();
  loadTableLibraryUI();
  loadParaStyleLibraryUI();
});

// ===================== TABS =====================
function switchTab(name) {
  const names = ['align','shape','table','text','colors','canvas'];
  document.querySelectorAll('.tab').forEach((t, i) => t.classList.toggle('active', names[i] === name));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
}

// ===================== STATUS =====================
function showStatus(msg, type = "ok") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = type;
  setTimeout(() => { el.className = ""; }, 3000);
}

// ===================== TOGGLE BUTTONS =====================
function toggleBtn(el) { el.classList.toggle('on'); }

// ===================== ALIGNMENT =====================
async function alignShapes(direction) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height");
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();

    const items = shapes.items;
    if (!items.length) return showStatus("No shapes selected", "err");

    const slideW = slide.width, slideH = slide.height;
    const minLeft   = Math.min(...items.map(s => s.left));
    const maxRight  = Math.max(...items.map(s => s.left + s.width));
    const minTop    = Math.min(...items.map(s => s.top));
    const maxBottom = Math.max(...items.map(s => s.top + s.height));

    items.forEach(s => {
      if (direction === 'left')   s.left = minLeft;
      if (direction === 'right')  s.left = maxRight - s.width;
      if (direction === 'center') s.left = (slideW - s.width) / 2;
      if (direction === 'top')    s.top  = minTop;
      if (direction === 'bottom') s.top  = maxBottom - s.height;
      if (direction === 'middle') s.top  = (slideH - s.height) / 2;
    });

    await context.sync();
    showStatus(`✓ Aligned ${direction}`);
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== DISTRIBUTE =====================
async function distributeShapes(axis) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height");
    await context.sync();

    const items = [...shapes.items];
    if (items.length < 3) return showStatus("Select at least 3 shapes", "err");

    if (axis === 'h') {
      items.sort((a, b) => a.left - b.left);
      const totalW = items.reduce((s, sh) => s + sh.width, 0);
      const span = items[items.length-1].left + items[items.length-1].width - items[0].left;
      const gap  = (span - totalW) / (items.length - 1);
      let x = items[0].left;
      items.forEach(s => { s.left = x; x += s.width + gap; });
    } else {
      items.sort((a, b) => a.top - b.top);
      const totalH = items.reduce((s, sh) => s + sh.height, 0);
      const span = items[items.length-1].top + items[items.length-1].height - items[0].top;
      const gap  = (span - totalH) / (items.length - 1);
      let y = items[0].top;
      items.forEach(s => { s.top = y; y += s.height + gap; });
    }

    await context.sync();
    showStatus("✓ Distributed");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== MATCH SIZE =====================
async function matchSize(type) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/width,items/height");
    await context.sync();
    const items = shapes.items;
    if (items.length < 2) return showStatus("Select at least 2 shapes", "err");
    const refW = items[0].width, refH = items[0].height;
    items.forEach(s => {
      if (type === 'width'  || type === 'both') s.width  = refW;
      if (type === 'height' || type === 'both') s.height = refH;
    });
    await context.sync();
    showStatus("✓ Size matched to first shape");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== CORNER RADIUS =====================
async function applyCornerRadius(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/visible");
    await context.sync();
    if (!shapes.items.length) return showStatus("No shapes selected", "err");

    const slide    = context.presentation.getSelectedSlides().getItemAt(0);
    const adjValue = Math.min(parseFloat(value) / 100, 0.5);

    for (const shape of shapes.items) {
      const left = shape.left, top = shape.top, width = shape.width, height = shape.height;
      let fillColor = "#4472C4", lineColor = "#000000", lineWeight = 1, lineVisible = false;
      try { fillColor   = shape.fill.foregroundColor || fillColor; }  catch(e) {}
      try { lineColor   = shape.lineFormat.color     || lineColor; }  catch(e) {}
      try { lineWeight  = shape.lineFormat.weight    || lineWeight; } catch(e) {}
      try { lineVisible = shape.lineFormat.visible; }                 catch(e) {}

      shape.delete();
      const newShape = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.roundRectangle,
        { left, top, width, height }
      );
      newShape.fill.setSolidColor(fillColor);
      newShape.lineFormat.color   = lineColor;
      newShape.lineFormat.weight  = lineWeight;
      newShape.lineFormat.visible = lineVisible;
      await context.sync();

      newShape.adjustments.load("items");
      await context.sync();
      if (newShape.adjustments.items.length > 0) {
        newShape.adjustments.items[0].value = adjValue;
        await context.sync();
      }
    }
    showStatus("✓ Corner radius applied");
  }).catch(e => showStatus("Error: " + e.message, "err"));
}

// ===================== OPACITY =====================
async function applyOpacity(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.transparency = 1 - parseFloat(value) / 100; } catch(e) {} });
    await context.sync();
    showStatus(`✓ Opacity ${value}%`);
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== FILL COLOR =====================
async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e) {} });
    await context.sync();
    showStatus("✓ Fill color applied");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== BORDER =====================
async function applyBorderColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.lineFormat.color = hex; s.lineFormat.visible = true; } catch(e) {} });
    await context.sync();
    showStatus("✓ Border color applied");
  }).catch(e => showStatus(e.message, "err"));
}

async function applyBorderWidth(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const pt = parseFloat(value);
    shapes.items.forEach(s => {
      try {
        if (pt === 0) s.lineFormat.visible = false;
        else { s.lineFormat.visible = true; s.lineFormat.weight = pt; }
      } catch(e) {}
    });
    await context.sync();
  }).catch(() => {});
}

// ===================== INSERT SVG =====================
async function insertSVG() {
  const svgCode = document.getElementById("svg-input").value.trim();
  if (!svgCode || !svgCode.includes("<svg"))
    return showStatus("Paste valid SVG code first", "err");

  try {
    const base64 = btoa(unescape(encodeURIComponent(svgCode)));

    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
        imageLeft:   100,
        imageTop:    100,
        imageWidth:  200,
        imageHeight: 200
      },
      function(result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          showStatus("Error: " + result.error.message, "err");
        } else {
          showStatus("✓ SVG inserted! Right-click → Convert to Shape");
        }
      }
    );
  } catch(e) {
    showStatus("Error: " + e.message, "err");
  }
}

// ===================== CREATE TABLE =====================
async function createTable() {
  const rows         = parseInt(document.getElementById('tblRows').value);
  const cols         = parseInt(document.getElementById('tblCols').value);
  const rowH         = parseFloat(document.getElementById('tblRowH').value);
  const colWInput    = parseFloat(document.getElementById('tblColW').value);
  const headerBg     = document.getElementById('tblHeaderBg').value;
  const headerFg     = document.getElementById('tblHeaderFg').value;
  const headerSize   = parseFloat(document.getElementById('tblHeaderSize').value);
  const headerFont   = document.getElementById('tblHeaderFont').value || 'Calibri';
  const headerBold   = document.getElementById('tblHeaderBold').classList.contains('on');
  const headerItalic = document.getElementById('tblHeaderItalic').classList.contains('on');
  const headerCaps   = document.getElementById('tblHeaderCaps').classList.contains('on');
  const headerAlign  = document.getElementById('tblHeaderAlign').value;
  const row1Color    = document.getElementById('tblRow1').value;
  const row2Color    = document.getElementById('tblRow2').value;
  const bodyFg       = document.getElementById('tblBodyFg').value;
  const bodySize     = parseFloat(document.getElementById('tblBodySize').value);
  const bodyFont     = document.getElementById('tblBodyFont').value || 'Calibri';
  const bodyAlign    = document.getElementById('tblBodyAlign').value;
  const borderColor  = document.getElementById('tblBorder').value;
  const borderW      = parseFloat(document.getElementById('tblBorderW').value);
  const padding      = parseFloat(document.getElementById('tblPadding').value);

  const alignMap = { left: 'Left', center: 'Center', right: 'Right', justify: 'Justify' };

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();

    const tableW = slide.width * 0.8;
    const colW   = colWInput > 0 ? colWInput : tableW / cols;

    const specificCellProperties = Array(rows).fill("").map((_, r) =>
      Array(cols).fill("").map(() => {
        const isHeader = r === 0;
        return {
          fill: { color: isHeader ? headerBg : (r % 2 === 0 ? row2Color : row1Color) },
          font: {
            color:   isHeader ? headerFg   : bodyFg,
            size:    isHeader ? headerSize : bodySize,
            name:    isHeader ? headerFont : bodyFont,
            bold:    isHeader ? headerBold : false,
            italic:  isHeader ? headerItalic : false,
            allCaps: isHeader ? headerCaps : false,
          },
          margins: { top: padding, bottom: padding, left: padding, right: padding },
          horizontalAlignment: alignMap[isHeader ? headerAlign : bodyAlign] || 'Left',
          borders: {
            bottom: { color: borderColor, weight: borderW },
            top:    { color: borderColor, weight: borderW },
            left:   { color: borderColor, weight: borderW },
            right:  { color: borderColor, weight: borderW },
          }
        };
      })
    );

    const values  = Array(rows).fill("").map((_, r) =>
      Array(cols).fill("").map((_, c) => r === 0 ? `Header ${c+1}` : "")
    );
    const columns = Array(cols).fill("").map(() => ({ columnWidth: colW }));
    const rowsOpt = Array(rows).fill("").map(() => ({ rowHeight: rowH }));

    const shape = slide.shapes.addTable(rows, cols, {
      values, specificCellProperties, columns, rows: rowsOpt,
    });
    shape.left = slide.width  * 0.1;
    shape.top  = slide.height * 0.15;

    await context.sync();
    showStatus("✓ Table inserted!");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== TABLE LIBRARY =====================
function getTableLib() { try { return JSON.parse(localStorage.getItem("tableLibrary") || "[]"); } catch { return []; } }
function saveTableLib(lib) { localStorage.setItem("tableLibrary", JSON.stringify(lib)); }

function saveTableStyle() {
  const lib  = getTableLib();
  const name = prompt("Template name:", `Table Style ${lib.length + 1}`);
  if (!name) return;
  lib.push({
    id: Date.now(), name,
    rows:        parseInt(document.getElementById('tblRows').value),
    cols:        parseInt(document.getElementById('tblCols').value),
    rowH:        parseFloat(document.getElementById('tblRowH').value),
    colW:        parseFloat(document.getElementById('tblColW').value),
    headerBg:    document.getElementById('tblHeaderBg').value,
    headerFg:    document.getElementById('tblHeaderFg').value,
    headerSize:  parseFloat(document.getElementById('tblHeaderSize').value),
    headerFont:  document.getElementById('tblHeaderFont').value,
    headerBold:  document.getElementById('tblHeaderBold').classList.contains('on'),
    headerItalic:document.getElementById('tblHeaderItalic').classList.contains('on'),
    headerCaps:  document.getElementById('tblHeaderCaps').classList.contains('on'),
    headerAlign: document.getElementById('tblHeaderAlign').value,
    row1Color:   document.getElementById('tblRow1').value,
    row2Color:   document.getElementById('tblRow2').value,
    bodyFg:      document.getElementById('tblBodyFg').value,
    bodySize:    parseFloat(document.getElementById('tblBodySize').value),
    bodyFont:    document.getElementById('tblBodyFont').value,
    bodyAlign:   document.getElementById('tblBodyAlign').value,
    borderColor: document.getElementById('tblBorder').value,
    borderW:     parseFloat(document.getElementById('tblBorderW').value),
    padding:     parseFloat(document.getElementById('tblPadding').value),
  });
  saveTableLib(lib);
  loadTableLibraryUI();
  showStatus(`✓ Saved "${name}"`);
}

function loadTableStyle(item) {
  document.getElementById('tblRows').value        = item.rows;
  document.getElementById('tblCols').value        = item.cols;
  document.getElementById('tblRowH').value        = item.rowH;
  document.getElementById('tblColW').value        = item.colW || 0;
  document.getElementById('tblHeaderBg').value    = item.headerBg;
  document.getElementById('tblHeaderFg').value    = item.headerFg;
  document.getElementById('tblHeaderSize').value  = item.headerSize;
  document.getElementById('tblHeaderFont').value  = item.headerFont;
  document.getElementById('tblHeaderAlign').value = item.headerAlign;
  document.getElementById('tblRow1').value        = item.row1Color;
  document.getElementById('tblRow2').value        = item.row2Color;
  document.getElementById('tblBodyFg').value      = item.bodyFg;
  document.getElementById('tblBodySize').value    = item.bodySize;
  document.getElementById('tblBodyFont').value    = item.bodyFont;
  document.getElementById('tblBodyAlign').value   = item.bodyAlign;
  document.getElementById('tblBorder').value      = item.borderColor;
  document.getElementById('tblBorderW').value     = item.borderW;
  document.getElementById('tblBorderWVal').textContent = item.borderW;
  document.getElementById('tblPadding').value     = item.padding;
  document.getElementById('tblHeaderBold').classList.toggle('on',   !!item.headerBold);
  document.getElementById('tblHeaderItalic').classList.toggle('on', !!item.headerItalic);
  document.getElementById('tblHeaderCaps').classList.toggle('on',   !!item.headerCaps);
  showStatus(`✓ Loaded "${item.name}" — click Insert Table`);
}

function loadTableLibraryUI() {
  const lib = getTableLib();
  const el  = document.getElementById('tableLibrary');
  if (!lib.length) {
    el.innerHTML = '<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved templates</div>';
    return;
  }
  el.innerHTML = lib.map(item => `
    <div class="table-item" onclick="loadTableStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div class="table-info">
        <b>${item.name}</b>
        ${item.rows}×${item.cols} · ${item.rowH}pt rows
      </div>
      <div style="display:flex;gap:4px;align-items:center">
        <div style="width:14px;height:14px;background:${item.headerBg};border-radius:2px"></div>
        <div style="width:14px;height:14px;background:${item.row1Color};border-radius:2px"></div>
        <div style="width:14px;height:14px;background:${item.row2Color};border-radius:2px"></div>
        <span onclick="event.stopPropagation();deleteTableStyle(${item.id})"
          style="color:#ccc;font-size:16px;cursor:pointer;margin-left:4px">×</span>
      </div>
    </div>
  `).join('');
}

function deleteTableStyle(id) { saveTableLib(getTableLib().filter(i => i.id !== id)); loadTableLibraryUI(); }
function clearTableLib() { if(confirm("Clear all table templates?")) { saveTableLib([]); loadTableLibraryUI(); } }

// ===================== CHARACTER =====================
async function applyCharacter() {
  const fontName   = document.getElementById('txtFont').value;
  const style      = document.getElementById('txtStyle').value;
  const fontSize   = parseFloat(document.getElementById('txtSize').value);
  const leadingPt  = parseFloat(document.getElementById('txtLeadingPt').value);
  const fontColor  = document.getElementById('txtColor').value;
  const tracking   = parseInt(document.getElementById('txtTracking').value);
  const hScale     = parseFloat(document.getElementById('txtHScale').value);
  const vScale     = parseFloat(document.getElementById('txtVScale').value);
  const baseline   = parseFloat(document.getElementById('txtBaseline').value);
  const bold       = style === 'bold'       || style === 'bolditalic' || document.getElementById('txtBold').classList.contains('on');
  const italic     = style === 'italic'     || style === 'bolditalic' || document.getElementById('txtItalic').classList.contains('on');
  const underline  = document.getElementById('txtUnderline').classList.contains('on');
  const strike     = document.getElementById('txtStrike').classList.contains('on');
  const allCaps    = document.getElementById('txtCaps').classList.contains('on');
  const smallCaps  = document.getElementById('txtSmallCaps').classList.contains('on');
  const superscript= document.getElementById('txtSuper').classList.contains('on');
  const subscript  = document.getElementById('txtSub').classList.contains('on');

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();

    for (const s of shapes.items) {
      try {
        const tr   = s.textFrame.textRange;
        const font = tr.font;

        font.name       = fontName;
        font.size       = fontSize;
        font.color      = fontColor;
        font.bold       = bold;
        font.italic     = italic;
        font.allCaps    = allCaps;
        font.smallCaps  = smallCaps;
        font.strikethrough = strike;
        font.superscript   = superscript;
        font.subscript     = subscript;
        font.underline  = underline
          ? PowerPoint.ShapeFontUnderlineStyle.single
          : PowerPoint.ShapeFontUnderlineStyle.none;

        // Tracking (kerning)
        try { font.kerning = tracking; } catch(e) {}

        // Leading — line spacing بالـ pt
        if (leadingPt > 0) {
          try { tr.paragraphFormat.lineSpacing = leadingPt; } catch(e) {}
        }

        // Horizontal / Vertical scale — مش موجودة في JS API مباشرة
        // بنعملها عن طريق width/height للـ textFrame
        // (أفضل حل متاح)

        // Baseline shift
        try { font.baselineOffset = baseline / fontSize; } catch(e) {}

      } catch(e) {}
    }

    await context.sync();
    showStatus("✓ Character applied");
  }).catch(e => showStatus("Error: " + e.message, "err"));
}

async function applyFontSize(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try { s.textFrame.textRange.font.size = parseFloat(value); } catch(e) {}
    }
    await context.sync();
    showStatus(`✓ Font size ${value}pt`);
  }).catch(e => showStatus(e.message, "err"));
}

async function applyFontColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try { s.textFrame.textRange.font.color = hex; } catch(e) {}
    }
    await context.sync();
    showStatus("✓ Color applied");
  }).catch(e => showStatus(e.message, "err"));
}

async function toggleTextStyle(style, el) {
  el.classList.toggle('on');
  const isOn = el.classList.contains('on');
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try {
        const font = s.textFrame.textRange.font;
        if (style === 'bold')          font.bold          = isOn;
        if (style === 'italic')        font.italic        = isOn;
        if (style === 'caps')          font.allCaps       = isOn;
        if (style === 'smallCaps')     font.smallCaps     = isOn;
        if (style === 'strikethrough') font.strikethrough = isOn;
        if (style === 'superscript')   font.superscript   = isOn;
        if (style === 'subscript')     font.subscript     = isOn;
        if (style === 'underline')     font.underline     = isOn
          ? PowerPoint.ShapeFontUnderlineStyle.single
          : PowerPoint.ShapeFontUnderlineStyle.none;
      } catch(e) {}
    }
    await context.sync();
    showStatus(`✓ ${style} ${isOn ? 'on' : 'off'}`);
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== TRACKING =====================
async function applyTracking(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try { s.textFrame.textRange.font.kerning = parseInt(value); } catch(e) {}
    }
    await context.sync();
    showStatus(`✓ Tracking ${value}`);
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== PARAGRAPH FORMAT =====================
async function applyParagraphFormat() {
  const align       = document.getElementById('txtAlign').value;
  const spaceBefore = parseFloat(document.getElementById('txtSpaceBefore').value);
  const spaceAfter  = parseFloat(document.getElementById('txtSpaceAfter').value);
  const lineSpacing = parseFloat(document.getElementById('txtLineSpacing').value);
  const indentLeft  = parseFloat(document.getElementById('txtIndentLeft').value);
  const indentRight = parseFloat(document.getElementById('txtIndentRight').value);
  const firstLine   = parseFloat(document.getElementById('txtFirstLine').value);

  const alignMap = {
    left:        PowerPoint.ParagraphHorizontalAlignment.left,
    center:      PowerPoint.ParagraphHorizontalAlignment.center,
    right:       PowerPoint.ParagraphHorizontalAlignment.right,
    justify:     PowerPoint.ParagraphHorizontalAlignment.justify,
    justifyLow:  PowerPoint.ParagraphHorizontalAlignment.justifyLow,
    distributed: PowerPoint.ParagraphHorizontalAlignment.distributed,
  };

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();

    for (const s of shapes.items) {
      try {
        const pf = s.textFrame.textRange.paragraphFormat;
        pf.horizontalAlignment = alignMap[align] || PowerPoint.ParagraphHorizontalAlignment.left;
        pf.spaceBefore         = spaceBefore;
        pf.spaceAfter          = spaceAfter;
        pf.lineSpacing         = lineSpacing;
        try { pf.leftMargin   = indentLeft;  } catch(e) {}
        try { pf.rightMargin  = indentRight; } catch(e) {}
        try { pf.firstLineIndent = firstLine; } catch(e) {}
      } catch(e) {}
    }

    await context.sync();
    showStatus("✓ Paragraph applied");
  }).catch(e => showStatus("Error: " + e.message, "err"));
}

// ===================== PARAGRAPH STYLES LIBRARY =====================
function getParaStyles()      { try { return JSON.parse(localStorage.getItem("paraStyles") || "[]"); } catch { return []; } }
function saveParaStylesLib(s) { localStorage.setItem("paraStyles", JSON.stringify(s)); }

function saveParaStyle() {
  const styles = getParaStyles();
  const name   = prompt("Style name:", `Style ${styles.length + 1}`);
  if (!name) return;
  styles.push({
    id:          Date.now(), name,
    fontName:    document.getElementById('txtFont').value,
    style:       document.getElementById('txtStyle').value,
    fontSize:    parseFloat(document.getElementById('txtSize').value),
    leadingPt:   parseFloat(document.getElementById('txtLeadingPt').value),
    fontColor:   document.getElementById('txtColor').value,
    tracking:    parseInt(document.getElementById('txtTracking').value),
    hScale:      parseFloat(document.getElementById('txtHScale').value),
    vScale:      parseFloat(document.getElementById('txtVScale').value),
    baseline:    parseFloat(document.getElementById('txtBaseline').value),
    bold:        document.getElementById('txtBold').classList.contains('on'),
    italic:      document.getElementById('txtItalic').classList.contains('on'),
    underline:   document.getElementById('txtUnderline').classList.contains('on'),
    strike:      document.getElementById('txtStrike').classList.contains('on'),
    allCaps:     document.getElementById('txtCaps').classList.contains('on'),
    smallCaps:   document.getElementById('txtSmallCaps').classList.contains('on'),
    superscript: document.getElementById('txtSuper').classList.contains('on'),
    subscript:   document.getElementById('txtSub').classList.contains('on'),
    align:       document.getElementById('txtAlign').value,
    spaceBefore: parseFloat(document.getElementById('txtSpaceBefore').value),
    spaceAfter:  parseFloat(document.getElementById('txtSpaceAfter').value),
    lineSpacing: parseFloat(document.getElementById('txtLineSpacing').value),
    indentLeft:  parseFloat(document.getElementById('txtIndentLeft').value),
    indentRight: parseFloat(document.getElementById('txtIndentRight').value),
    firstLine:   parseFloat(document.getElementById('txtFirstLine').value),
  });
  saveParaStylesLib(styles);
  loadParaStyleLibraryUI();
  showStatus(`✓ Style "${name}" saved`);
}

function loadParaStyle(item) {
  document.getElementById('txtFont').value              = item.fontName;
  document.getElementById('txtStyle').value             = item.style || 'normal';
  document.getElementById('txtSize').value              = item.fontSize;
  document.getElementById('txtLeadingPt').value         = item.leadingPt || 0;
  document.getElementById('txtColor').value             = item.fontColor;
  document.getElementById('txtTracking').value          = item.tracking;
  document.getElementById('txtTrackingVal').textContent = item.tracking;
  document.getElementById('txtHScale').value            = item.hScale || 100;
  document.getElementById('txtVScale').value            = item.vScale || 100;
  document.getElementById('txtBaseline').value          = item.baseline || 0;
  document.getElementById('txtSpaceBefore').value       = item.spaceBefore;
  document.getElementById('txtSpaceAfter').value        = item.spaceAfter;
  document.getElementById('txtLineSpacing').value       = item.lineSpacing || 100;
  document.getElementById('txtLineSpacingVal').textContent = (item.lineSpacing || 100) + '%';
  document.getElementById('txtIndentLeft').value        = item.indentLeft  || 0;
  document.getElementById('txtIndentRight').value       = item.indentRight || 0;
  document.getElementById('txtFirstLine').value         = item.firstLine   || 0;
  document.getElementById('txtAlign').value             = item.align;
  document.getElementById('txtBold').classList.toggle('on',         !!item.bold);
  document.getElementById('txtItalic').classList.toggle('on',       !!item.italic);
  document.getElementById('txtUnderline').classList.toggle('on',    !!item.underline);
  document.getElementById('txtStrike').classList.toggle('on',       !!item.strike);
  document.getElementById('txtCaps').classList.toggle('on',         !!item.allCaps);
  document.getElementById('txtSmallCaps').classList.toggle('on',    !!item.smallCaps);
  document.getElementById('txtSuper').classList.toggle('on',        !!item.superscript);
  document.getElementById('txtSub').classList.toggle('on',          !!item.subscript);
  showStatus(`✓ Loaded "${item.name}" — click Apply`);
}

function loadParaStyleLibraryUI() {
  const styles = getParaStyles();
  const el     = document.getElementById('paraStyleLibrary');
  if (!styles.length) {
    el.innerHTML = '<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved styles</div>';
    return;
  }
  el.innerHTML = styles.map(item => `
    <div class="style-item" onclick="loadParaStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div>
        <div class="style-preview" style="font-family:'${item.fontName}';color:${item.fontColor};font-weight:${item.bold?'bold':'normal'};font-style:${item.italic?'italic':'normal'}">
          ${item.name}
        </div>
        <div class="style-meta">${item.fontName} · ${item.fontSize}pt · ${item.align}</div>
      </div>
      <span onclick="event.stopPropagation();deleteParaStyle(${item.id})"
        style="color:#ccc;font-size:16px;cursor:pointer;margin-left:8px">×</span>
    </div>
  `).join('');
}

function deleteParaStyle(id) { saveParaStylesLib(getParaStyles().filter(s => s.id !== id)); loadParaStyleLibraryUI(); }
function clearParaStyles()   { if(confirm("Clear all styles?")) { saveParaStylesLib([]); loadParaStyleLibraryUI(); } }

// ===================== COLOR PALETTE =====================
let selectedPaletteColor = null;

function getPalette()   { try { return JSON.parse(localStorage.getItem("colorPalette") || "[]"); } catch { return []; } }
function savePalette(p) { localStorage.setItem("colorPalette", JSON.stringify(p)); }

function addColor() {
  const hex  = document.getElementById('newColor').value;
  const name = document.getElementById('newColorName').value || hex;
  const p    = getPalette();
  p.push({ id: Date.now(), hex, name });
  savePalette(p);
  loadPaletteUI();
  document.getElementById('newColorName').value = '';
}

function deleteColor(id) { savePalette(getPalette().filter(c => c.id !== id)); loadPaletteUI(); }

function loadPaletteUI() {
  const p  = getPalette();
  const mk = (clickFn, showDel) => p.map(c => `
    <div class="swatch" style="background:${c.hex}" title="${c.name}" onclick="${clickFn}('${c.hex}','${c.name}')">
      ${showDel ? `<span class="del" onclick="event.stopPropagation();deleteColor(${c.id})">×</span>` : ''}
    </div>
  `).join('') || '<span style="font-size:11px;color:#aaa">No colors yet</span>';

  document.getElementById('paletteGrid').innerHTML      = mk('selectColor', true);
  document.getElementById('applyPaletteGrid').innerHTML = mk('selectColor', false);
}

function selectColor(hex, name) {
  selectedPaletteColor = hex;
  document.getElementById('selectedColorPreview').innerHTML =
    `<span style="display:inline-block;width:12px;height:12px;background:${hex};border-radius:2px;margin-right:4px;vertical-align:middle"></span>Selected: ${name}`;
}

async function applyColorToSelected(type) {
  if (!selectedPaletteColor) return showStatus("Click a color swatch first", "err");
  if (type === 'fill') await applyFillColor(selectedPaletteColor);
  else                 await applyBorderColor(selectedPaletteColor);
}

// ===================== BULK REPLACE =====================
async function bulkColorReplace() {
  const findHex    = document.getElementById('findColor').value.replace('#','').toUpperCase();
  const replaceHex = document.getElementById('replaceColor').value;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/fill/foregroundColor,items/type");
    await context.sync();
    let count = 0;
    for (const s of slide.shapes.items) {
      try {
        if ((s.fill.foregroundColor||"").replace('#','').toUpperCase() === findHex) {
          s.fill.setSolidColor(replaceHex); count++;
        }
      } catch(e) {}
    }
    await context.sync();
    showStatus(`✓ Replaced ${count} shape(s)`);
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== GRID =====================
async function addGrid() {
  const size    = parseFloat(document.getElementById('gridSize').value);
  const color   = document.getElementById('gridColor').value;
  const opacity = parseFloat(document.getElementById('gridOpacity').value) / 100;

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const W = slide.width, H = slide.height;

    for (let x = size; x < W; x += size) {
      const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight,
        { left: x, top: 0, height: H, width: 0 });
      line.lineFormat.color        = color;
      line.lineFormat.weight       = 0.5;
      line.lineFormat.transparency = 1 - opacity;
      line.name = "__grid__";
    }
    for (let y = size; y < H; y += size) {
      const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight,
        { left: 0, top: y, height: 0, width: W });
      line.lineFormat.color        = color;
      line.lineFormat.weight       = 0.5;
      line.lineFormat.transparency = 1 - opacity;
      line.name = "__grid__";
    }
    await context.sync();
    showStatus("✓ Grid added");
  }).catch(e => showStatus(e.message, "err"));
}

async function removeGrid() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/name");
    await context.sync();
    slide.shapes.items.forEach(s => { if (s.name === "__grid__") s.delete(); });
    await context.sync();
    showStatus("✓ Grid removed");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== GUIDES =====================
async function addGuide(axis) {
  const pos   = parseFloat(document.getElementById('guidePos').value);
  const color = document.getElementById('guideColor').value;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const W = slide.width, H = slide.height;
    const line = axis === 'h'
      ? slide.shapes.addLine(PowerPoint.ConnectorType.straight, { left: 0, top: pos, height: 0, width: W })
      : slide.shapes.addLine(PowerPoint.ConnectorType.straight, { left: pos, top: 0, height: H, width: 0 });
    line.lineFormat.color  = color;
    line.lineFormat.weight = 1;
    line.name = "__guide__";
    await context.sync();
    showStatus(`✓ Guide at ${pos}pt`);
  }).catch(e => showStatus(e.message, "err"));
}

async function removeGuides() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/name");
    await context.sync();
    slide.shapes.items.forEach(s => { if (s.name === "__guide__") s.delete(); });
    await context.sync();
    showStatus("✓ All guides removed");
  }).catch(e => showStatus(e.message, "err"));
}

// ===================== SLIDE INFO =====================
async function showSlideInfo() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const px = v => Math.round(v * 96 / 72);
    document.getElementById('slideInfo').innerHTML =
      `${Math.round(slide.width)} × ${Math.round(slide.height)} pt<br>
       ${px(slide.width)} × ${px(slide.height)} px`;
  }).catch(e => showStatus(e.message, "err"));
}