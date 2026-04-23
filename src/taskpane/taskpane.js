Office.onReady(() => {
  loadPaletteUI();
  loadTableLibraryUI();
  loadParaStyleLibraryUI();
  loadSystemFonts();
});

// ── TABS ──────────────────────────────────────────────
function switchTab(name) {
  const names = ['shapes','table','text','colors','canvas'];
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', names[i]===name));
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
}

// ── STATUS ────────────────────────────────────────────
function showStatus(msg, type='ok') {
  const el = document.getElementById('status');
  el.textContent = msg; el.className = type;
  setTimeout(() => { el.className=''; }, 3000);
}

// ── TOGGLE ────────────────────────────────────────────
function toggleBtn(el) { el.classList.toggle('on'); }

// ── ALIGN MODE ────────────────────────────────────────
let alignMode = 'slide';
function setAlignMode(mode) {
  alignMode = mode;
  document.getElementById('mode-slide').classList.toggle('active', mode==='slide');
  document.getElementById('mode-selection').classList.toggle('active', mode==='selection');
}

// ── ALIGN ─────────────────────────────────────────────
async function alignShapes(direction) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height");
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();

    const items = shapes.items;
    if (!items.length) return showStatus("No shapes selected","err");

    let refLeft, refRight, refTop, refBottom;

    if (alignMode === 'slide') {
      refLeft = 0; refRight = slide.width;
      refTop  = 0; refBottom = slide.height;
    } else {
      refLeft   = Math.min(...items.map(s => s.left));
      refRight  = Math.max(...items.map(s => s.left + s.width));
      refTop    = Math.min(...items.map(s => s.top));
      refBottom = Math.max(...items.map(s => s.top + s.height));
    }

    const totalW = refRight - refLeft;
    const totalH = refBottom - refTop;

    items.forEach(s => {
      switch(direction) {
        case 'left':   s.left = refLeft; break;
        case 'right':  s.left = refRight - s.width; break;
        case 'center': s.left = refLeft + (totalW - s.width) / 2; break;
        case 'top':    s.top  = refTop; break;
        case 'bottom': s.top  = refBottom - s.height; break;
        case 'middle': s.top  = refTop + (totalH - s.height) / 2; break;
      }
    });

    await context.sync();
    showStatus(`✓ Aligned ${direction}`);
  }).catch(e => showStatus(e.message,"err"));
}

// ── DISTRIBUTE ────────────────────────────────────────
async function distributeShapes(axis) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height");
    await context.sync();

    const items = [...shapes.items];
    if (items.length < 3) return showStatus("Select at least 3 shapes","err");

    if (axis === 'h') {
      items.sort((a,b) => a.left - b.left);
      const first = items[0].left;
      const last  = items[items.length-1].left + items[items.length-1].width;
      const totalW = items.reduce((s,sh) => s + sh.width, 0);
      const gap = (last - first - totalW) / (items.length - 1);
      let x = first;
      items.forEach(s => { s.left = x; x += s.width + gap; });
    } else {
      items.sort((a,b) => a.top - b.top);
      const first = items[0].top;
      const last  = items[items.length-1].top + items[items.length-1].height;
      const totalH = items.reduce((s,sh) => s + sh.height, 0);
      const gap = (last - first - totalH) / (items.length - 1);
      let y = first;
      items.forEach(s => { s.top = y; y += s.height + gap; });
    }

    await context.sync();
    showStatus("✓ Distributed");
  }).catch(e => showStatus(e.message,"err"));
}

// ── MATCH SIZE ────────────────────────────────────────
async function matchSize(type) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/width,items/height");
    await context.sync();
    const items = shapes.items;
    if (items.length < 2) return showStatus("Select at least 2 shapes","err");
    const refW = items[0].width, refH = items[0].height;
    items.forEach(s => {
      if (type==='width'  || type==='both') s.width  = refW;
      if (type==='height' || type==='both') s.height = refH;
    });
    await context.sync();
    showStatus("✓ Size matched");
  }).catch(e => showStatus(e.message,"err"));
}

// ── CONVERT TO ROUND RECT ─────────────────────────────
async function convertToRoundRect() {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/left,items/top,items/width,items/height,items/fill/foregroundColor,items/lineFormat/color,items/lineFormat/weight,items/lineFormat/visible");
    await context.sync();

    if (!shapes.items.length) return showStatus("No shapes selected","err");

    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    let count = 0;

    for (const shape of shapes.items) {
      const left=shape.left, top=shape.top, width=shape.width, height=shape.height;
      let fill="#4472C4", lColor="#000", lWeight=1, lVisible=false;
      try { fill    = shape.fill.foregroundColor || fill; }    catch(e){}
      try { lColor  = shape.lineFormat.color     || lColor; }  catch(e){}
      try { lWeight = shape.lineFormat.weight    || lWeight; } catch(e){}
      try { lVisible= shape.lineFormat.visible; }              catch(e){}

      shape.delete();
      await context.sync();

      const ns = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.roundRectangle,
        { left, top, width, height }
      );
      ns.fill.setSolidColor(fill);
      ns.lineFormat.color   = lColor;
      ns.lineFormat.weight  = lWeight;
      ns.lineFormat.visible = lVisible;
      await context.sync();
      count++;
    }

    showStatus(`✓ Converted ${count} shape(s) to Round Rectangle`);
  }).catch(e => showStatus("Error: "+e.message,"err"));
}

// ── CORNER RADIUS ─────────────────────────────────────
async function applyCornerRadius() {
  const radiusPt = parseFloat(document.getElementById('radiusInput').value);
  if (isNaN(radiusPt) || radiusPt < 0) return showStatus("Enter a valid radius","err");

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/width,items/height");
    await context.sync();

    if (!shapes.items.length) return showStatus("No shapes selected","err");

    let applied = 0;
    for (const shape of shapes.items) {
      try {
        const minDim = Math.min(shape.width, shape.height);
        // adjValue = radius_pt / (minDim/2), capped at 0.5
        const adjValue = Math.min(radiusPt / (minDim / 2), 0.5);
        shape.adjustments.load("items");
        await context.sync();
        if (shape.adjustments.items && shape.adjustments.items.length > 0) {
          shape.adjustments.items[0].value = adjValue;
          await context.sync();
          applied++;
        }
      } catch(e) { /* shape doesn't support adjustments */ }
    }

    if (applied > 0) showStatus(`✓ Radius applied to ${applied} shape(s)`);
    else showStatus("Select a Round Rectangle shape first","err");
  }).catch(e => showStatus("Error: "+e.message,"err"));
}

// ── OPACITY ───────────────────────────────────────────
async function applyOpacity(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.transparency = 1 - parseFloat(value)/100; } catch(e){} });
    await context.sync();
  }).catch(()=>{});
}

// ── FILL COLOR ────────────────────────────────────────
async function applyFillColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.fill.setSolidColor(hex); } catch(e){} });
    await context.sync();
  }).catch(()=>{});
}

// ── BORDER COLOR ──────────────────────────────────────
async function applyBorderColor(hex) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    shapes.items.forEach(s => { try { s.lineFormat.color=hex; s.lineFormat.visible=true; } catch(e){} });
    await context.sync();
  }).catch(()=>{});
}

// ── BORDER WIDTH ──────────────────────────────────────
async function applyBorderWidth(value) {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const pt = parseFloat(value);
    shapes.items.forEach(s => {
      try {
        if (pt===0) s.lineFormat.visible = false;
        else { s.lineFormat.visible=true; s.lineFormat.weight=pt; }
      } catch(e){}
    });
    await context.sync();
  }).catch(()=>{});
}

// ── INSERT SVG ────────────────────────────────────────
async function insertSVG() {
  const svgCode = document.getElementById("svg-input").value.trim();
  if (!svgCode || !svgCode.includes("<svg"))
    return showStatus("Paste valid SVG code first","err");
  try {
    const base64 = btoa(unescape(encodeURIComponent(svgCode)));
    Office.context.document.setSelectedDataAsync(
      base64,
      { coercionType: Office.CoercionType.Image, imageLeft:100, imageTop:100, imageWidth:200, imageHeight:200 },
      r => {
        if (r.status === Office.AsyncResultStatus.Failed) showStatus("Error: "+r.error.message,"err");
        else showStatus("✓ Inserted! Right-click → Convert to Shape");
      }
    );
  } catch(e) { showStatus("Error: "+e.message,"err"); }
}

// ── TABLE ─────────────────────────────────────────────
async function createTable() {
  const rows        = parseInt(document.getElementById('tblRows').value);
  const cols        = parseInt(document.getElementById('tblCols').value);
  const rowH        = parseFloat(document.getElementById('tblRowH').value);
  const colWInput   = parseFloat(document.getElementById('tblColW').value);
  const headerBg    = document.getElementById('tblHeaderBg').value;
  const headerFg    = document.getElementById('tblHeaderFg').value;
  const headerSize  = parseFloat(document.getElementById('tblHeaderSize').value);
  const headerFont  = document.getElementById('tblHeaderFont').value||'Calibri';
  const headerBold  = document.getElementById('tblHeaderBold').classList.contains('on');
  const headerItalic= document.getElementById('tblHeaderItalic').classList.contains('on');
  const headerCaps  = document.getElementById('tblHeaderCaps').classList.contains('on');
  const headerAlign = document.getElementById('tblHeaderAlign').value;
  const row1        = document.getElementById('tblRow1').value;
  const row2        = document.getElementById('tblRow2').value;
  const bodyFg      = document.getElementById('tblBodyFg').value;
  const bodySize    = parseFloat(document.getElementById('tblBodySize').value);
  const bodyFont    = document.getElementById('tblBodyFont').value||'Calibri';
  const bodyAlign   = document.getElementById('tblBodyAlign').value;
  const borderColor = document.getElementById('tblBorder').value;
  const borderW     = parseFloat(document.getElementById('tblBorderW').value);
  const padding     = parseFloat(document.getElementById('tblPadding').value);
  const aMap = { left:'Left', center:'Center', right:'Right' };

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();

    const tableW = slide.width * 0.8;
    const colW   = colWInput > 0 ? colWInput : tableW / cols;

    const scp = Array(rows).fill("").map((_,r) =>
      Array(cols).fill("").map(() => ({
        fill: { color: r===0 ? headerBg : (r%2===0 ? row2 : row1) },
        font: {
          color:   r===0 ? headerFg   : bodyFg,
          size:    r===0 ? headerSize : bodySize,
          name:    r===0 ? headerFont : bodyFont,
          bold:    r===0 ? headerBold : false,
          italic:  r===0 ? headerItalic : false,
          allCaps: r===0 ? headerCaps : false,
        },
        margins: { top:padding, bottom:padding, left:padding, right:padding },
        horizontalAlignment: aMap[r===0 ? headerAlign : bodyAlign]||'Left',
        borders: {
          bottom:{color:borderColor,weight:borderW},
          top:{color:borderColor,weight:borderW},
          left:{color:borderColor,weight:borderW},
          right:{color:borderColor,weight:borderW},
        }
      }))
    );

    const values  = Array(rows).fill("").map((_,r) => Array(cols).fill("").map((_,c) => r===0?`Header ${c+1}`:""));
    const columns = Array(cols).fill("").map(() => ({ columnWidth:colW }));
    const rowsOpt = Array(rows).fill("").map(() => ({ rowHeight:rowH }));

    const shape = slide.shapes.addTable(rows, cols, { values, specificCellProperties:scp, columns, rows:rowsOpt });
    shape.left = slide.width*0.1;
    shape.top  = slide.height*0.15;
    await context.sync();
    showStatus("✓ Table inserted!");
  }).catch(e => showStatus(e.message,"err"));
}

// ── TABLE LIBRARY ─────────────────────────────────────
function getTableLib() { try { return JSON.parse(localStorage.getItem("tableLibrary")||"[]"); } catch { return []; } }
function saveTableLib(lib) { localStorage.setItem("tableLibrary", JSON.stringify(lib)); }

function saveTableStyle() {
  const lib  = getTableLib();
  const name = prompt("Template name:", `Table ${lib.length+1}`);
  if (!name) return;
  lib.push({
    id:Date.now(), name,
    rows:parseInt(document.getElementById('tblRows').value),
    cols:parseInt(document.getElementById('tblCols').value),
    rowH:parseFloat(document.getElementById('tblRowH').value),
    colW:parseFloat(document.getElementById('tblColW').value),
    headerBg:document.getElementById('tblHeaderBg').value,
    headerFg:document.getElementById('tblHeaderFg').value,
    headerSize:parseFloat(document.getElementById('tblHeaderSize').value),
    headerFont:document.getElementById('tblHeaderFont').value,
    headerBold:document.getElementById('tblHeaderBold').classList.contains('on'),
    headerItalic:document.getElementById('tblHeaderItalic').classList.contains('on'),
    headerCaps:document.getElementById('tblHeaderCaps').classList.contains('on'),
    headerAlign:document.getElementById('tblHeaderAlign').value,
    row1:document.getElementById('tblRow1').value,
    row2:document.getElementById('tblRow2').value,
    bodyFg:document.getElementById('tblBodyFg').value,
    bodySize:parseFloat(document.getElementById('tblBodySize').value),
    bodyFont:document.getElementById('tblBodyFont').value,
    bodyAlign:document.getElementById('tblBodyAlign').value,
    borderColor:document.getElementById('tblBorder').value,
    borderW:parseFloat(document.getElementById('tblBorderW').value),
    padding:parseFloat(document.getElementById('tblPadding').value),
  });
  saveTableLib(lib); loadTableLibraryUI(); showStatus(`✓ Saved "${name}"`);
}

function loadTableStyle(item) {
  document.getElementById('tblRows').value        = item.rows;
  document.getElementById('tblCols').value        = item.cols;
  document.getElementById('tblRowH').value        = item.rowH;
  document.getElementById('tblColW').value        = item.colW||0;
  document.getElementById('tblHeaderBg').value    = item.headerBg;
  document.getElementById('tblHeaderFg').value    = item.headerFg;
  document.getElementById('tblHeaderSize').value  = item.headerSize;
  document.getElementById('tblHeaderFont').value  = item.headerFont;
  document.getElementById('tblHeaderAlign').value = item.headerAlign;
  document.getElementById('tblRow1').value        = item.row1;
  document.getElementById('tblRow2').value        = item.row2;
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
  showStatus(`✓ Loaded "${item.name}"`);
}

function loadTableLibraryUI() {
  const lib = getTableLib();
  const el  = document.getElementById('tableLibrary');
  if (!lib.length) { el.innerHTML='<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved templates</div>'; return; }
  el.innerHTML = lib.map(item => `
    <div class="table-item" onclick="loadTableStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div class="table-info">
        <b>${item.name}</b>${item.rows}×${item.cols} · ${item.rowH}pt
      </div>
      <div style="display:flex;gap:4px;align-items:center">
        <div style="width:12px;height:12px;background:${item.headerBg};border-radius:2px"></div>
        <div style="width:12px;height:12px;background:${item.row1};border-radius:2px"></div>
        <span onclick="event.stopPropagation();deleteTableStyle(${item.id})" style="color:#ccc;font-size:16px;cursor:pointer">×</span>
      </div>
    </div>`).join('');
}
function deleteTableStyle(id) { saveTableLib(getTableLib().filter(i=>i.id!==id)); loadTableLibraryUI(); }
function clearTableLib() { if(confirm("Clear all?")) { saveTableLib([]); loadTableLibraryUI(); } }

// ── FONT PICKER ───────────────────────────────────────
const FONTS = [
  "Arial","Arial Black","Arial Narrow","Calibri","Calibri Light","Cambria",
  "Century Gothic","Comic Sans MS","Courier New","Franklin Gothic Medium",
  "Garamond","Georgia","Gill Sans","Helvetica","Impact","Lucida Console",
  "Myriad Pro","Open Sans","Optima","Palatino","Rockwell","Tahoma",
  "Times New Roman","Trebuchet MS","Verdana","Baskerville","Didot",
  "Bodoni MT","Bebas Neue","Montserrat","Oswald","Raleway","Roboto",
  "Lato","Source Sans Pro","Nunito","Poppins","Inter","DM Sans","Futura"
];
let allFonts = [...FONTS];

function loadSystemFonts() {
  try {
    if ('fonts' in document) {
      document.fonts.ready.then(() => {
        const detected = FONTS.filter(f => document.fonts.check(`12px "${f}"`));
        if (detected.length > 5) allFonts = [...new Set([...detected,...FONTS])];
        renderFontDropdown(allFonts);
      });
    }
  } catch(e) {}
  renderFontDropdown(allFonts);
}

function renderFontDropdown(fonts) {
  const dd = document.getElementById('fontDropdown');
  dd.innerHTML = fonts.map(f =>
    `<div class="font-option" style="font-family:'${f}'" onclick="selectFont('${f}')">${f}</div>`
  ).join('');
}
function filterFonts(val) {
  renderFontDropdown(allFonts.filter(f => f.toLowerCase().includes(val.toLowerCase())));
  document.getElementById('fontDropdown').classList.add('open');
}
function openFontDropdown() {
  renderFontDropdown(allFonts);
  document.getElementById('fontDropdown').classList.add('open');
}
function selectFont(name) {
  document.getElementById('txtFont').value = name;
  document.getElementById('fontDropdown').classList.remove('open');
  liveChar();
}
document.addEventListener('click', e => {
  if (!e.target.closest('.font-wrap'))
    document.getElementById('fontDropdown').classList.remove('open');
});

// ── TEXT ──────────────────────────────────────────────
let charTimer = null;
function scheduleChar() {
  clearTimeout(charTimer);
  charTimer = setTimeout(applyChar, 400);
}

function liveChar() { scheduleChar(); }

async function applyChar() {
  const fontName  = document.getElementById('txtFont').value;
  const style     = document.getElementById('txtStyle').value;
  const fontSize  = parseFloat(document.getElementById('txtSize').value);
  const fontColor = document.getElementById('txtColor').value;
  const bold      = style==='bold'||style==='bolditalic'||document.getElementById('txtBold').classList.contains('on');
  const italic    = style==='italic'||style==='bolditalic'||document.getElementById('txtItalic').classList.contains('on');
  const underline = document.getElementById('txtUnderline').classList.contains('on');
  const strike    = document.getElementById('txtStrike').classList.contains('on');
  const allCaps   = document.getElementById('txtCaps').classList.contains('on');
  const superscript = document.getElementById('txtSuper').classList.contains('on');
  const subscript   = document.getElementById('txtSub').classList.contains('on');

  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try {
        const font = s.textFrame.textRange.font;
        if (fontName) font.name = fontName;
        if (!isNaN(fontSize)) font.size = fontSize;
        font.color        = fontColor;
        font.bold         = bold;
        font.italic       = italic;
        font.allCaps      = allCaps;
        font.strikethrough= strike;
        font.superscript  = superscript;
        font.subscript    = subscript;
        font.underline    = underline
          ? PowerPoint.ShapeFontUnderlineStyle.single
          : PowerPoint.ShapeFontUnderlineStyle.none;
      } catch(e){}
    }
    await context.sync();
    showStatus("✓ Character applied");
  }).catch(e => showStatus(e.message,"err"));
}

async function toggleLive(el, style) {
  el.classList.toggle('on');
  await applyChar();
}

async function applyPara() {
  const align = document.getElementById('txtAlign').value;
  const aMap = {
    left:    PowerPoint.ParagraphHorizontalAlignment.left,
    center:  PowerPoint.ParagraphHorizontalAlignment.center,
    right:   PowerPoint.ParagraphHorizontalAlignment.right,
    justify: PowerPoint.ParagraphHorizontalAlignment.justify,
  };
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    for (const s of shapes.items) {
      try {
        s.textFrame.textRange.paragraphFormat.horizontalAlignment = aMap[align];
      } catch(e){}
    }
    await context.sync();
    showStatus("✓ Paragraph applied");
  }).catch(e => showStatus(e.message,"err"));
}

async function liveParaAlign() { await applyPara(); }

// ── PARA STYLES ───────────────────────────────────────
function getParaStyles()      { try { return JSON.parse(localStorage.getItem("paraStyles")||"[]"); } catch { return []; } }
function saveParaStylesLib(s) { localStorage.setItem("paraStyles", JSON.stringify(s)); }

function saveParaStyle() {
  const styles = getParaStyles();
  const name   = prompt("Style name:", `Style ${styles.length+1}`);
  if (!name) return;
  styles.push({
    id:Date.now(), name,
    fontName:  document.getElementById('txtFont').value,
    style:     document.getElementById('txtStyle').value,
    fontSize:  parseFloat(document.getElementById('txtSize').value),
    fontColor: document.getElementById('txtColor').value,
    bold:      document.getElementById('txtBold').classList.contains('on'),
    italic:    document.getElementById('txtItalic').classList.contains('on'),
    underline: document.getElementById('txtUnderline').classList.contains('on'),
    strike:    document.getElementById('txtStrike').classList.contains('on'),
    allCaps:   document.getElementById('txtCaps').classList.contains('on'),
    superscript:document.getElementById('txtSuper').classList.contains('on'),
    subscript: document.getElementById('txtSub').classList.contains('on'),
    align:     document.getElementById('txtAlign').value,
  });
  saveParaStylesLib(styles); loadParaStyleLibraryUI(); showStatus(`✓ Style "${name}" saved`);
}

function loadParaStyle(item) {
  document.getElementById('txtFont').value   = item.fontName;
  document.getElementById('txtStyle').value  = item.style||'normal';
  document.getElementById('txtSize').value   = item.fontSize;
  document.getElementById('txtColor').value  = item.fontColor;
  document.getElementById('txtAlign').value  = item.align;
  document.getElementById('txtBold').classList.toggle('on',      !!item.bold);
  document.getElementById('txtItalic').classList.toggle('on',    !!item.italic);
  document.getElementById('txtUnderline').classList.toggle('on', !!item.underline);
  document.getElementById('txtStrike').classList.toggle('on',    !!item.strike);
  document.getElementById('txtCaps').classList.toggle('on',      !!item.allCaps);
  document.getElementById('txtSuper').classList.toggle('on',     !!item.superscript);
  document.getElementById('txtSub').classList.toggle('on',       !!item.subscript);
  applyChar(); applyPara();
  showStatus(`✓ Loaded "${item.name}"`);
}

function loadParaStyleLibraryUI() {
  const styles = getParaStyles();
  const el     = document.getElementById('paraStyleLibrary');
  if (!styles.length) { el.innerHTML='<div style="font-size:12px;color:#aaa;text-align:center;padding:10px">No saved styles</div>'; return; }
  el.innerHTML = styles.map(item => `
    <div class="style-item" onclick="loadParaStyle(${JSON.stringify(item).replace(/"/g,'&quot;')})">
      <div>
        <div style="font-family:'${item.fontName}';color:${item.fontColor};font-weight:${item.bold?'bold':'normal'};font-style:${item.italic?'italic':'normal'};font-size:13px">${item.name}</div>
        <div class="style-meta">${item.fontName} · ${item.fontSize}pt · ${item.align}</div>
      </div>
      <span onclick="event.stopPropagation();deleteParaStyle(${item.id})" style="color:#ccc;font-size:16px;cursor:pointer">×</span>
    </div>`).join('');
}
function deleteParaStyle(id) { saveParaStylesLib(getParaStyles().filter(s=>s.id!==id)); loadParaStyleLibraryUI(); }
function clearParaStyles()   { if(confirm("Clear?")) { saveParaStylesLib([]); loadParaStyleLibraryUI(); } }

// ── COLOR PALETTE ─────────────────────────────────────
let selectedColor = null;
function getPalette()   { try { return JSON.parse(localStorage.getItem("colorPalette")||"[]"); } catch { return []; } }
function savePalette(p) { localStorage.setItem("colorPalette", JSON.stringify(p)); }

function addColor() {
  const hex  = document.getElementById('newColor').value;
  const name = document.getElementById('newColorName').value||hex;
  const p    = getPalette();
  p.push({id:Date.now(), hex, name});
  savePalette(p); loadPaletteUI();
  document.getElementById('newColorName').value='';
}
function deleteColor(id) { savePalette(getPalette().filter(c=>c.id!==id)); loadPaletteUI(); }

function loadPaletteUI() {
  const p  = getPalette();
  const mk = (fn, del) => p.map(c => `
    <div class="swatch" style="background:${c.hex}" title="${c.name}" onclick="${fn}('${c.hex}','${c.name}')">
      ${del?`<span class="del" onclick="event.stopPropagation();deleteColor(${c.id})">×</span>`:''}
    </div>`).join('')||'<span style="font-size:11px;color:#aaa">No colors</span>';

  document.getElementById('paletteGrid').innerHTML      = mk('selectColor', true);
  document.getElementById('applyPaletteGrid').innerHTML = mk('selectColor', false);
}

function selectColor(hex, name) {
  selectedColor = hex;
  document.getElementById('selectedColorPreview').innerHTML =
    `<span style="display:inline-block;width:12px;height:12px;background:${hex};border-radius:2px;margin-right:4px;vertical-align:middle"></span>${name}`;
}
async function applyColorToSelected(type) {
  if (!selectedColor) return showStatus("Click a color first","err");
  if (type==='fill') await applyFillColor(selectedColor);
  else               await applyBorderColor(selectedColor);
}

async function bulkColorReplace() {
  const findHex    = document.getElementById('findColor').value.replace('#','').toUpperCase();
  const replaceHex = document.getElementById('replaceColor').value;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/fill/foregroundColor,items/type");
    await context.sync();
    let count=0;
    for (const s of slide.shapes.items) {
      try {
        if ((s.fill.foregroundColor||'').replace('#','').toUpperCase()===findHex) {
          s.fill.setSolidColor(replaceHex); count++;
        }
      } catch(e){}
    }
    await context.sync();
    showStatus(`✓ Replaced ${count} shape(s)`);
  }).catch(e => showStatus(e.message,"err"));
}

// ── GRID ──────────────────────────────────────────────
async function addGrid() {
  const size  = parseFloat(document.getElementById('gridSize').value);
  const color = document.getElementById('gridColor').value;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const W=slide.width, H=slide.height;
    for (let x=size; x<W; x+=size) {
      const l = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {left:x,top:0,height:H,width:0});
      l.lineFormat.color=color; l.lineFormat.weight=0.5; l.name="__grid__";
    }
    for (let y=size; y<H; y+=size) {
      const l = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {left:0,top:y,height:0,width:W});
      l.lineFormat.color=color; l.lineFormat.weight=0.5; l.name="__grid__";
    }
    await context.sync();
    showStatus("✓ Grid added");
  }).catch(e => showStatus(e.message,"err"));
}

async function removeGrid() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/name");
    await context.sync();
    slide.shapes.items.forEach(s => { if(s.name==="__grid__") s.delete(); });
    await context.sync();
    showStatus("✓ Grid removed");
  }).catch(e => showStatus(e.message,"err"));
}

// ── GUIDES ────────────────────────────────────────────
async function addGuide(axis) {
  const pos   = parseFloat(document.getElementById('guidePos').value);
  const color = document.getElementById('guideColor').value;
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const W=slide.width, H=slide.height;
    const line = axis==='h'
      ? slide.shapes.addLine(PowerPoint.ConnectorType.straight, {left:0,top:pos,height:0,width:W})
      : slide.shapes.addLine(PowerPoint.ConnectorType.straight, {left:pos,top:0,height:H,width:0});
    line.lineFormat.color=color; line.lineFormat.weight=1; line.name="__guide__";
    await context.sync();
    showStatus(`✓ Guide at ${pos}pt`);
  }).catch(e => showStatus(e.message,"err"));
}

async function removeGuides() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load("items/name");
    await context.sync();
    slide.shapes.items.forEach(s => { if(s.name==="__guide__") s.delete(); });
    await context.sync();
    showStatus("✓ Guides removed");
  }).catch(e => showStatus(e.message,"err"));
}

// ── SLIDE INFO ────────────────────────────────────────
async function showSlideInfo() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("width,height");
    await context.sync();
    const px = v => Math.round(v*96/72);
    document.getElementById('slideInfo').innerHTML =
      `${Math.round(slide.width)} × ${Math.round(slide.height)} pt<br>${px(slide.width)} × ${px(slide.height)} px`;
  }).catch(e => showStatus(e.message,"err"));
}