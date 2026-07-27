/**
 * BPlan â€“ "Xuất báo cáo" feature
 * Layout (exported_layout.txt):
 *   HEADER      -> sheet 1 of template
 *   LINE 1      -> "Số học sinh: X"
 *   LINE 2      -> "Chi phí/người: Y VND / 1 ngày"
 *   TABLE       -> styled result table (7 cols)
 *   LINE1 | BOTTOM  -> "Ngân sách" | sheet 2 col C+
 *   LINE2 |         -> "Chi phí"
 *   LINE3 |         -> "Chênh lệch"
 *
 * After building, opens Univer Sheets modal for editing HEADER+FOOTER only.
 * On save: caches changed template to localStorage + offers download.
 * Requires: xlsx-js-style, SweetAlert2, _calcBorderFull/_calcBorderMed/_calcBorderThin,
 *           _calcAutoWidths, _calcAlign
 */

//  Template cache 
const _BP_TMPL_KEY = 'bplan_report_template_v1';

function _bpSaveTmplCache(uint8, cacheKey) {
  cacheKey = cacheKey || _BP_TMPL_KEY;
  try {
    let s = '';
    for (let i = 0; i < uint8.length; i++) s += String.fromCharCode(uint8[i]);
    localStorage.setItem(cacheKey, btoa(s));
  } catch (_) {}
}

function _bpLoadTmplCache(cacheKey) {
  cacheKey = cacheKey || _BP_TMPL_KEY;
  try {
    const raw = localStorage.getItem(cacheKey);
    if (!raw) return null;
    const bin = atob(raw);
    const u8  = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  } catch (_) { return null; }
}

//  Download helper 
function _bpDL(arrBuf, name) {
  const blob = new Blob([arrBuf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 150);
}

//  XLSX write helper 
function _bpWriteXlsx(wb) {
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: false });
}

// ── Debug helper: log cell styles for inspection ────────────────────────────────
function _bpDebugLogStyles(sheet, label, maxCells) {
  maxCells = maxCells || 5;
  const keys = Object.keys(sheet).filter(k => !k.startsWith('!')).slice(0, maxCells);
  console.log(`[BPlan] ${label} - Sample cells:`, keys.map(k => {
    const cell = sheet[k];
    // Log full style object for detailed inspection
    if (cell.s && typeof cell.s === 'object') {
      console.log(`  [BPlan] Cell ${k} style:`, JSON.stringify(cell.s, null, 2));
    }
    return {
      addr: k,
      v: cell.v,
      styleType: typeof cell.s,
      hasFont: cell.s && typeof cell.s === 'object' ? !!cell.s.font : false,
      hasFill: cell.s && typeof cell.s === 'object' ? !!cell.s.fill : false,
      hasBorder: cell.s && typeof cell.s === 'object' ? !!cell.s.border : false,
      hasAlign: cell.s && typeof cell.s === 'object' ? !!cell.s.alignment : false
    };
  }));
  
  // Also log merges for this sheet
  const merges = sheet['!merges'] || [];
  console.log(`[BPlan] ${label} - Merges (${merges.length}):`, merges.map(m => 
    `${XLSX.utils.encode_cell(m.s)}:${XLSX.utils.encode_cell(m.e)}`
  ).join(', ') || 'none');
}

// ═══════════════════════════════════════════════════════════════════════════════
// ROBUST STYLE EXTRACTION FROM XLSX TEMPLATE
// ═══════════════════════════════════════════════════════════════════════════════
// 
// xlsx-js-style 1.2.0-beta has issues with style preservation:
// - When reading Excel files, cell.s may be an XF index (integer) instead of object
// - The library doesn't always resolve these indices properly
// 
// SOLUTION: Use JSZip to manually extract styles from the XLSX (which is a ZIP file)
// and build a style lookup table that we can use to resolve XF indices.

// ── Parse XLSX XML to extract the complete style table ──────────────────────────
async function _bpExtractStylesFromXlsx(uint8Array) {
  const styleTable = {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [],      // Cell format records (maps xfId -> style)
    cellStyleXfs: [], // Cell style format records
    resolved: {}      // Cache: xfId -> resolved style object
  };
  
  try {
    // Use JSZip to extract the XLSX (it's a ZIP file)
    const zip = await JSZip.loadAsync(uint8Array);
    
    // Parse xl/styles.xml
    const stylesFile = zip.file('xl/styles.xml');
    if (!stylesFile) {
      console.warn('[BPlan] No styles.xml found in XLSX');
      return styleTable;
    }
    
    const stylesXml = await stylesFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(stylesXml, 'application/xml');
    
    // Helper: Parse color from XML
    const parseColor = (colorEl) => {
      if (!colorEl) return null;
      const rgb = colorEl.getAttribute('rgb');
      const indexed = colorEl.getAttribute('indexed');
      const theme = colorEl.getAttribute('theme');
      const tint = colorEl.getAttribute('tint');
      if (rgb) return { rgb: rgb };
      if (indexed) return { indexed: parseInt(indexed) };
      if (theme) return { theme: parseInt(theme), tint: tint ? parseFloat(tint) : undefined };
      return null;
    };
    
    // Helper: Parse font from XML
    const parseFont = (fontEl) => {
      const font = {};
      const nameEl = fontEl.querySelector('name');
      if (nameEl) font.name = nameEl.getAttribute('val');
      
      const szEl = fontEl.querySelector('sz');
      if (szEl) font.sz = parseFloat(szEl.getAttribute('val'));
      
      const colorEl = fontEl.querySelector('color');
      if (colorEl) font.color = parseColor(colorEl);
      
      const boldEl = fontEl.querySelector('b');
      if (boldEl) font.bold = true;
      
      const italicEl = fontEl.querySelector('i');
      if (italicEl) font.italic = true;
      
      const underlineEl = fontEl.querySelector('u');
      if (underlineEl) font.underline = underlineEl.getAttribute('val') || true;
      
      const strikeEl = fontEl.querySelector('strike');
      if (strikeEl) font.strike = true;
      
      return font;
    };
    
    // Helper: Parse fill from XML
    const parseFill = (fillEl) => {
      const patternFill = fillEl.querySelector('patternFill');
      if (!patternFill) return null;
      
      const patternType = patternFill.getAttribute('patternType');
      if (!patternType || patternType === 'none') return null;
      
      const fgColorEl = patternFill.querySelector('fgColor');
      const bgColorEl = patternFill.querySelector('bgColor');
      
      const fill = { patternType };
      if (fgColorEl) fill.fgColor = parseColor(fgColorEl);
      if (bgColorEl) fill.bgColor = parseColor(bgColorEl);
      
      return fill;
    };
    
    // Helper: Parse border from XML
    const parseBorder = (borderEl) => {
      const border = {};
      const sides = ['left', 'right', 'top', 'bottom', 'diagonal'];
      
      sides.forEach(side => {
        const sideEl = borderEl.querySelector(side);
        if (!sideEl) return;
        
        const style = sideEl.getAttribute('style');
        if (!style) return;
        
        const colorEl = sideEl.querySelector('color');
        border[side] = {
          style: style,
          color: colorEl ? parseColor(colorEl) : { rgb: '000000' }
        };
      });
      
      return Object.keys(border).length ? border : null;
    };
    
    // Helper: Parse alignment from XML
    const parseAlignment = (alignEl) => {
      if (!alignEl) return null;
      const alignment = {};
      
      const horizontal = alignEl.getAttribute('horizontal');
      const vertical = alignEl.getAttribute('vertical');
      const wrapText = alignEl.getAttribute('wrapText');
      const textRotation = alignEl.getAttribute('textRotation');
      const shrinkToFit = alignEl.getAttribute('shrinkToFit');
      const indent = alignEl.getAttribute('indent');
      
      if (horizontal) alignment.horizontal = horizontal;
      if (vertical) alignment.vertical = vertical;
      if (wrapText === '1' || wrapText === 'true') alignment.wrapText = true;
      if (textRotation) alignment.textRotation = parseInt(textRotation);
      if (shrinkToFit === '1' || shrinkToFit === 'true') alignment.shrinkToFit = true;
      if (indent) alignment.indent = parseInt(indent);
      
      return Object.keys(alignment).length ? alignment : null;
    };
    
    // Parse fonts
    const fontsEl = doc.querySelector('fonts');
    if (fontsEl) {
      const fontEls = fontsEl.querySelectorAll(':scope > font');
      fontEls.forEach(fontEl => {
        styleTable.fonts.push(parseFont(fontEl));
      });
    }
    
    // Parse fills
    const fillsEl = doc.querySelector('fills');
    if (fillsEl) {
      const fillEls = fillsEl.querySelectorAll(':scope > fill');
      fillEls.forEach(fillEl => {
        styleTable.fills.push(parseFill(fillEl));
      });
    }
    
    // Parse borders
    const bordersEl = doc.querySelector('borders');
    if (bordersEl) {
      const borderEls = bordersEl.querySelectorAll(':scope > border');
      borderEls.forEach(borderEl => {
        styleTable.borders.push(parseBorder(borderEl));
      });
    }
    
    // Parse cellXfs (cell format records)
    const cellXfsEl = doc.querySelector('cellXfs');
    if (cellXfsEl) {
      const xfEls = cellXfsEl.querySelectorAll(':scope > xf');
      xfEls.forEach((xfEl, idx) => {
        const xf = {
          numFmtId: parseInt(xfEl.getAttribute('numFmtId')) || 0,
          fontId: parseInt(xfEl.getAttribute('fontId')) || 0,
          fillId: parseInt(xfEl.getAttribute('fillId')) || 0,
          borderId: parseInt(xfEl.getAttribute('borderId')) || 0,
          xfId: parseInt(xfEl.getAttribute('xfId')) || 0,
          applyFont: xfEl.getAttribute('applyFont') === '1',
          applyFill: xfEl.getAttribute('applyFill') === '1',
          applyBorder: xfEl.getAttribute('applyBorder') === '1',
          applyAlignment: xfEl.getAttribute('applyAlignment') === '1',
          applyProtection: xfEl.getAttribute('applyProtection') === '1'
        };
        
        // Parse alignment if present
        const alignEl = xfEl.querySelector('alignment');
        if (alignEl) {
          xf.alignment = parseAlignment(alignEl);
        }
        
        styleTable.cellXfs.push(xf);
      });
    }
    
    console.log(`[BPlan] Extracted ${styleTable.fonts.length} fonts, ${styleTable.fills.length} fills, ${styleTable.borders.length} borders, ${styleTable.cellXfs.length} cellXfs`);
    
    return styleTable;
    
  } catch (e) {
    console.error('[BPlan] Failed to extract styles from XLSX:', e);
    return styleTable;
  }
}

// ── Resolve an XF index to a complete style object ───────────────────────────────
function _bpResolveXfToStyle(xfId, styleTable) {
  if (xfId == null || !styleTable || !styleTable.cellXfs[xfId]) {
    return null;
  }
  
  // Check cache first
  if (styleTable.resolved[xfId]) {
    return JSON.parse(JSON.stringify(styleTable.resolved[xfId]));
  }
  
  const xf = styleTable.cellXfs[xfId];
  const style = {};
  
  // Resolve font
  if (xf.fontId != null && styleTable.fonts[xf.fontId]) {
    style.font = JSON.parse(JSON.stringify(styleTable.fonts[xf.fontId]));
  }
  
  // Resolve fill
  if (xf.fillId != null && styleTable.fills[xf.fillId]) {
    const fill = styleTable.fills[xf.fillId];
    if (fill && fill.patternType !== 'none') {
      style.fill = JSON.parse(JSON.stringify(fill));
    }
  }
  
  // Resolve border
  if (xf.borderId != null && styleTable.borders[xf.borderId]) {
    const border = styleTable.borders[xf.borderId];
    if (border) {
      style.border = JSON.parse(JSON.stringify(border));
    }
  }
  
  // Copy alignment
  if (xf.alignment) {
    style.alignment = JSON.parse(JSON.stringify(xf.alignment));
  }
  
  // Cache the result
  styleTable.resolved[xfId] = style;
  
  return JSON.parse(JSON.stringify(style));
}

// ── Read XLSX with full style resolution using JSZip ─────────────────────────────
async function _bpReadXlsxWithStyles(uint8Array) {
  // First, extract styles using JSZip
  const styleTable = await _bpExtractStylesFromXlsx(uint8Array);
  
  // Also extract cell xfIds from sheet XML
  const cellXfIds = await _bpExtractCellXfIds(uint8Array);
  
  // Then read with xlsx-js-style to get the data
  const wb = XLSX.read(uint8Array, { type: 'array', cellStyles: true });
  
  // Debug: Check what JSZip extracted
  console.log('[BPlan] Style table from JSZip:', {
    fonts: styleTable.fonts.length,
    fills: styleTable.fills.length,
    borders: styleTable.borders.length,
    cellXfs: styleTable.cellXfs.length
  });
  
  // Log sample of JSZip-extracted styles
  if (styleTable.fonts.length > 0) {
    console.log('[BPlan] Sample fonts from JSZip:', styleTable.fonts.slice(0, 2));
  }
  if (styleTable.borders.length > 0) {
    console.log('[BPlan] Sample borders from JSZip:', styleTable.borders);
  }
  
  // Now resolve styles for all cells using JSZip style table
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const sheetIndex = wb.SheetNames.indexOf(sheetName);
    const cellKeys = Object.keys(sheet).filter(k => !k.startsWith('!'));
    
    let resolvedCount = 0;
    let fromXlsxLib = 0;
    let noStyleCount = 0;
    
    for (const addr of cellKeys) {
      const cell = sheet[addr];
      const xfIdKey = `${sheetIndex}:${addr}`;
      const xfId = cellXfIds[xfIdKey];
      
      if (xfId !== undefined && styleTable.cellXfs[xfId]) {
        // We have the xfId from JSZip - resolve the full style
        const resolvedStyle = _bpResolveXfToStyle(xfId, styleTable);
        if (resolvedStyle) {
          cell.s = resolvedStyle;
          resolvedCount++;
          continue;
        }
      }
      
      // Fallback: use what xlsx-js-style gave us (cleaned up)
      if (cell.s && typeof cell.s === 'object') {
        if (cell.s.patternType === 'none') {
          delete cell.s.patternType;
        }
        if (cell.s.fill && cell.s.fill.patternType === 'none') {
          delete cell.s.fill;
        }
        fromXlsxLib++;
      } else if (cell.s != null && typeof cell.s === 'number') {
        // XF index from xlsx-js-style - try to resolve
        const resolvedStyle = _bpResolveXfToStyle(cell.s, styleTable);
        if (resolvedStyle) {
          cell.s = resolvedStyle;
          resolvedCount++;
        }
      } else {
        noStyleCount++;
      }
    }
    
    console.log(`[BPlan] Sheet "${sheetName}": resolved=${resolvedCount}, fromXlsxLib=${fromXlsxLib}, noStyle=${noStyleCount}`);
  }
  
  // Store the style table on the workbook for later use
  wb._styleTable = styleTable;
  
  return { wb, styleTable };
}

// ── Extract cell xfIds from sheet XML ───────────────────────────────────────────
async function _bpExtractCellXfIds(uint8Array) {
  const xfIds = {}; // key: "sheetIndex:cellAddr" -> xfId
  
  try {
    const zip = await JSZip.loadAsync(uint8Array);
    
    // Read each sheet
    const sheetFiles = [];
    zip.folder('xl/worksheets').forEach((relativePath, file) => {
      if (relativePath.endsWith('.xml')) {
        sheetFiles.push(relativePath);
      }
    });
    
    // Sort sheets by number
    sheetFiles.sort((a, b) => {
      const numA = parseInt(a.match(/sheet(\d+)/)?.[1] || '0');
      const numB = parseInt(b.match(/sheet(\d+)/)?.[1] || '0');
      return numA - numB;
    });
    
    for (let i = 0; i < sheetFiles.length; i++) {
      const sheetPath = `xl/worksheets/${sheetFiles[i]}`;
      const sheetFile = zip.file(sheetPath);
      if (!sheetFile) continue;
      
      const sheetXml = await sheetFile.async('text');
      const parser = new DOMParser();
      const doc = parser.parseFromString(sheetXml, 'application/xml');
      
      // Find all cells with 's' attribute (style index)
      const cells = doc.querySelectorAll('c[s]');
      cells.forEach(cellEl => {
        const r = cellEl.getAttribute('r'); // Cell reference like "A1"
        const s = cellEl.getAttribute('s'); // Style index (xfId)
        if (r && s !== null) {
          xfIds[`${i}:${r}`] = parseInt(s);
        }
      });
    }
    
    console.log(`[BPlan] Extracted ${Object.keys(xfIds).length} cell xfIds from sheet XML`);
    
  } catch (e) {
    console.warn('[BPlan] Failed to extract cell xfIds:', e);
  }
  
  return xfIds;
}

//  Build report sheet (returns {ws, zones}) 
function _bpBuildSheet(result, totalPeople, budget, headerAoa, bottomAoa) {
  // 6 columns A:F only
  const HDRS  = ['Buổi','Tên món','Số lượng','Đơn vị','Giá tiền/đv','Chi phí'];
  const NC    = HDRS.length;  // 6
  const total = result.total_cost || 0;
  const diff  = result.difference != null ? result.difference : total - budget;
  const cpd   = Math.round(total / Math.max(totalPeople, 1) * 10) / 10;
  const rows  = [];

  // HEADER zone — push NC-wide EMPTY rows as placeholders.
  // _bpApplyTmplStyles will overwrite these with the real template cells (values + styles).
  // The AOA values are not used for HEADER cells; they are here only so that
  // aoa_to_sheet builds the correct !ref range and row count.
  const hCount = headerAoa.length;
  for (let i = 0; i < hCount; i++) rows.push(new Array(NC).fill(''));

  // Info lines
  const il1 = rows.length;
  rows.push([`Số học sinh: ${totalPeople}`, ...Array(NC-1).fill('')]);
  const il2 = rows.length;
  rows.push([`Chi phí /người: ${cpd.toLocaleString('vi')} VND / 1 ngày`, ...Array(NC-1).fill('')]);

  // Table
  const tS = rows.length;
  rows.push([...HDRS]);
  result.details.forEach(item => {
    const ppu = item.quantity ? item.cost / item.quantity : 0;
    rows.push([item.part||'', item.name||'', item.quantity, item.unit||'', ppu, item.cost]);
  });
  rows.push(['','','','','Tổng chi phí     ', total]);
  const tE = rows.length - 1;

  // Summary + FOOTER zone
  // Cols A:B (0–1) = summary text; cols C:F (2–5) = FOOTER placeholder (4 cols).
  // _bpApplyTmplStyles overwrites cols C:F with real template footer cells.
  const sumLines = [
    [`Ngân sách: ${budget.toLocaleString('vi')} VND`],
    [`Chi phí: ${total.toLocaleString('vi')} VND`],
    [`Chênh lệch: ${diff.toLocaleString('vi')} VND`],
  ];
  const bS = rows.length;
  const footerRowCount = Math.max(sumLines.length, bottomAoa.length);
  for (let si = 0; si < footerRowCount; si++) {
    const sumText = sumLines[si] ? sumLines[si][0] : '';
    // 6 cols total: A=sumText, B='', C:F='' (placeholders for template content)
    rows.push([sumText, '', '', '', '', '']);
  }
  const bE = rows.length - 1;

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = _calcAutoWidths(rows, 1, 8, 55);
  ws['!rows'] = rows.map((_, i) => ({
    hpx: i === tS ? 24 : (i === il1 || i === il2) ? 20 : 18
  }));

  // Info line styles
  const IS = {
    font:      { bold: true, sz: 11, italic: true },
    alignment: { horizontal: 'left', vertical: 'center' },
    fill:      { fgColor: { rgb: 'FFF2CC' } },
    border:    _calcBorderFull('thin', 'CCCCCC')
  };
  [il1, il2].forEach(ri => {
    for (let ci = 0; ci < NC; ci++) {
      const a = XLSX.utils.encode_cell({ r: ri, c: ci });
      if (!ws[a]) ws[a] = { v: '', t: 's' };
      ws[a].s = JSON.parse(JSON.stringify(IS));
    }
  });

  // Table styles
  for (let R = tS; R <= tE; R++) {
    const isH = R === tS, isF = R === tE, isE = R % 2 === 0;
    for (let C = 0; C < NC; C++) {
      const a = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[a]) ws[a] = { v: '', t: 's' };
      const cur = C === 4 || C === 5;
      if (isH) {
        ws[a].s = { font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          fill: { fgColor: { rgb: '2E5597' } }, border: _calcBorderFull('medium', '1D3A6E') };
      } else if (isF) {
        ws[a].s = { font: { bold: cur, sz: 11 },
          alignment: { horizontal: _calcAlign(rows[R][C]), vertical: 'center' },
          fill: { fgColor: { rgb: 'D9E1F2' } },
          border: { top: _calcBorderMed('2E5597'), bottom: _calcBorderMed('1D3A6E'),
            left:  C===0 ? _calcBorderMed('2E5597') : _calcBorderThin('9DB2CE'),
            right: C===NC-1 ? _calcBorderMed('2E5597') : _calcBorderThin('9DB2CE') } };
        if (cur && typeof ws[a].v === 'number') ws[a].z = '#,##0 "VND"';
      } else {
        ws[a].s = { font: { bold: C===0, sz: 10 },
          alignment: { horizontal: _calcAlign(rows[R][C]), vertical: 'center' },
          fill: { fgColor: { rgb: isE ? 'EEF2FA' : 'FFFFFF' } },
          border: { top: _calcBorderThin('9DB2CE'), bottom: _calcBorderThin('9DB2CE'),
            left:  C===0 ? _calcBorderMed('2E5597') : _calcBorderThin('9DB2CE'),
            right: C===NC-1 ? _calcBorderMed('2E5597') : _calcBorderThin('9DB2CE') } };
        if (cur && typeof ws[a].v === 'number') ws[a].z = '#,##0 "VND"';
      }
    }
  }

  // Summary col A:B styles (cols C:F get styles from template via _bpApplyTmplStyles)
  const SF = [{ rgb: 'E2EFDA' }, { rgb: 'FCE4D6' },
    Math.abs(diff) < 1 ? { rgb: 'E2EFDA' } : { rgb: 'FCE4D6' }];
  for (let si = 0; si < sumLines.length; si++) {
    const ri = bS + si;
    for (let ci = 0; ci < 2; ci++) {
      const a = XLSX.utils.encode_cell({ r: ri, c: ci });
      if (!ws[a]) ws[a] = { v: '', t: 's' };
      ws[a].s = { font: { bold: true, sz: 11 },
        alignment: { horizontal: 'left', vertical: 'center' },
        fill: { fgColor: SF[si] || SF[2] }, border: _calcBorderFull('thin', 'AAAAAA') };
    }
  }

  return { ws, rows, zones: { hCount, il1, il2, tS, tE, bS, bE, NC } };
}

// ── Extract all style properties from a cell, handling XF indices ───────────────
// This is a comprehensive style extractor that works with both inline styles
// and XF-indexed styles.
function _bpExtractCellStyle(cell, wb, sheet, styleTable) {
  if (!cell) return null;
  
  // Case 1: Already inline object - deep clone and clean up
  if (cell.s && typeof cell.s === 'object') {
    const style = JSON.parse(JSON.stringify(cell.s));
    
    // Clean up invalid/incomplete style objects
    // xlsx-js-style sometimes returns {"patternType": "none"} which is not useful
    if (style.patternType === 'none') {
      delete style.patternType;
    }
    if (style.fill && style.fill.patternType === 'none') {
      delete style.fill;
    }
    
    // Check if there's anything meaningful left
    const hasContent = style.font || style.fill || style.border || style.alignment || 
                       style.numFmt || Object.keys(style).length > 0;
    
    if (!hasContent && Object.keys(style).length === 0) {
      return null;
    }
    
    return style;
  }
  
  // Case 2: XF index (integer) - resolve from style table
  if (cell.s && typeof cell.s === 'number' && styleTable) {
    const resolvedStyle = _bpResolveXfToStyle(cell.s, styleTable);
    if (resolvedStyle) {
      return resolvedStyle;
    }
  }
  
  // Case 3: No style property
  return null;
}

// ── Copy cells (value + inline style) from a template sheet region into the
// report worksheet, with optional row/col offsets. Only cols 0–5 (A:F) are
// written. Merges from the source are translated and also clamped to A:F.
// FIXED: Now uses JSZip-extracted style table for proper XF resolution.
// FIXED: Properly handles merged cells to avoid duplicates.
function _bpCopySheetRegion(srcSheet, dstWs, rowCount, colCount,
                             srcRowOff, srcColOff, dstRowOff, dstColOff, maxCol, srcWb, styleTable) {
  maxCol = maxCol != null ? maxCol : 5;
  
  // Build a map of merge ranges for this source region
  // For each merge, ONLY the origin cell should be copied
  // Non-origin cells should be skipped entirely (they will appear empty due to merge)
  const mergeMap = {};
  const mergeOriginCells = new Set(); // Track which cells are merge origins
  
  (srcSheet['!merges'] || []).forEach(m => {
    const originKey = `${m.s.r},${m.s.c}`;
    mergeOriginCells.add(originKey);
    
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        const key = `${r},${c}`;
        if (!mergeMap[key]) {
          mergeMap[key] = { 
            origin: { r: m.s.r, c: m.s.c }, 
            isOrigin: r === m.s.r && c === m.s.c,
            mergeRange: m
          };
        }
      }
    }
  });
  
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const dstC = dstColOff + c;
      if (dstC > maxCol) continue;
      
      const srcR = srcRowOff + r;
      const srcC = srcColOff + c;
      const srcAddr = XLSX.utils.encode_cell({ r: srcR, c: srcC });
      const dstAddr = XLSX.utils.encode_cell({ r: dstRowOff + r, c: dstC });
      
      // Check if this source cell is part of a merge
      const mergeKey = `${srcR},${srcC}`;
      const mergeInfo = mergeMap[mergeKey];
      
      // Skip non-origin cells in a merge - they should appear empty due to the merge
      if (mergeInfo && !mergeInfo.isOrigin) {
        continue; // Don't create any cell - the merge will make it appear empty
      }
      
      const srcCell = srcSheet[srcAddr];
      
      if (!srcCell) {
        // No source cell exists - only create if not part of a merge
        if (!mergeInfo) {
          if (!dstWs[dstAddr]) dstWs[dstAddr] = { v: '', t: 's' };
        }
        continue;
      }
      
      // Build destination cell with value
      const dstCell = { 
        v: srcCell.v != null ? srcCell.v : '', 
        t: srcCell.t || 's' 
      };
      
      // Extract and copy style - use JSZip-extracted style table
      const style = _bpExtractCellStyle(srcCell, srcWb, srcSheet, styleTable);
      if (style) {
        dstCell.s = style;
      } else {
        // If no style found, log for debugging
        console.log(`[BPlan] No style found for cell ${srcAddr}`, {
          hasS: !!srcCell.s,
          sType: typeof srcCell.s,
          sValue: srcCell.s
        });
      }
      
      // Copy number format if present
      if (srcCell.z) dstCell.z = srcCell.z;
      
      dstWs[dstAddr] = dstCell;
    }
  }
}

// ── Apply template HEADER and FOOTER into the report worksheet.
// HEADER: template sheet1 rows 1-4 (0-indexed) → report rows 0-3 (cols 0–5)
// FOOTER: template sheet2 rows 0-4 (0-indexed) → report rows bS–bE, cols 2-5
// Both merges and inline styles are copied. Merges are clamped to cols A:F.
// FIXED: Now uses JSZip-extracted style table for proper XF resolution.
// FIXED: Handle template row offset (template starts at row 1 in 0-indexed, report starts at row 0)
function _bpApplyTmplStyles(ws, headerSheet, bottomSheet, headerAoa, bottomAoa, zones, tmplWb, styleTable) {
  const _HEADER_ROWS = zones.hCount;        // 4
  const _FOOTER_ROWS = zones.bE - zones.bS + 1; // 5
  const _FOOTER_COLS = 4;                   // template sheet2 cols A:D (4) → report C:F

  // -- HEADER: copy from template rows 0-3 (0-indexed, A1:F4) to report rows 0-3 --
  // Template HEADER starts at Excel row 1 (row 0 in 0-indexed)
  _bpCopySheetRegion(headerSheet, ws,
    _HEADER_ROWS, zones.NC,   // 4 rows × 6 cols
    0, 0,                     // src starts at row 0, col 0
    0, 0,                     // dst A1 (report row 0)
    5, tmplWb, styleTable);   // maxCol, workbook, style table

  // -- FOOTER: copy from template rows 0-4 (A1:D5) to report rows bS–bE, cols C:F --
  // Template FOOTER starts at Excel row 1 (row 0 in 0-indexed, A1:D5)
  _bpCopySheetRegion(bottomSheet, ws,
    _FOOTER_ROWS, _FOOTER_COLS, // 5 rows × 4 cols
    0, 0,                       // src starts at row 0 (A1:D5), col 0
    zones.bS, 2,                // dst row bS (C25), col C (index 2)
    5, tmplWb, styleTable);     // maxCol, workbook, style table

  // -- Merges: always reset to avoid accumulation across repeated calls --
  ws['!merges'] = [];

  // Helper: encode a merge as a canonical string key for deduplication
  const mergeKey = (sr, sc, er, ec) => `${sr},${sc},${er},${ec}`;
  const seen = new Set();
  const addMerge = (sr, sc, er, ec) => {
    // Clamp to valid column range 0–5 (A:F)
    sc = Math.max(0, Math.min(sc, 5));
    ec = Math.max(0, Math.min(ec, 5));
    if (sc > ec) return; // invalid after clamp
    const k = mergeKey(sr, sc, er, ec);
    if (seen.has(k)) return;
    seen.add(k);
    ws['!merges'].push({ s: { r: sr, c: sc }, e: { r: er, c: ec } });
  };

  // HEADER merges — template rows 0-3 (A1:F4) map 1-to-1 to report rows 0-3
  (headerSheet['!merges'] || []).forEach(m => {
    const reportStartRow = m.s.r;
    const reportEndRow = m.e.r;
    
    // Only include merges that fall within the report HEADER zone (rows 0-3)
    if (reportStartRow < 0 || reportStartRow >= _HEADER_ROWS) return;
    
    const clampedEndRow = Math.min(reportEndRow, _HEADER_ROWS - 1);
    if (clampedEndRow < reportStartRow) return;
    
    // Clamp columns to A:F (0-5)
    const startCol = Math.max(0, Math.min(m.s.c, 5));
    const endCol = Math.max(0, Math.min(m.e.c, 5));
    if (startCol > endCol) return;
    
    addMerge(reportStartRow, startCol, clampedEndRow, endCol);
  });

  // FOOTER merges — shift rows by +bS (template row 0 → report row bS), cols by +2
  (bottomSheet['!merges'] || []).forEach(m => {
    // Template FOOTER starts at row 0 (0-indexed, A1:D5), report FOOTER starts at row bS
    // Transform: template row → report row (template row + bS)
    const reportStartRow = m.s.r + zones.bS;
    const reportEndRow = m.e.r + zones.bS;
    
    // Only include merges that fall within the report FOOTER zone
    if (reportStartRow < zones.bS || reportStartRow > zones.bE) return;
    
    const clampedEndRow = Math.min(reportEndRow, zones.bE);
    if (clampedEndRow < reportStartRow) return;
    
    // Template footer is A:D (0-3), report footer is C:F (2-5)
    // Shift columns by +2
    const startCol = Math.max(0, Math.min(m.s.c + 2, 5));
    const endCol = Math.max(0, Math.min(m.e.c + 2, 5));
    if (startCol > endCol) return;
    
    addMerge(reportStartRow, startCol, clampedEndRow, endCol);
  });
  
  // -- Note on merge cell handling --
  // xlsx-js-style handles merges by only storing the value in the origin cell.
  // Non-origin cells should NOT be created as separate cell objects.
  // The merge range in !merges handles the visual span.
  // We do NOT need to copy styles to non-origin cells - the merge handles display.
  
  console.log(`[BPlan] Applied ${ws['!merges'].length} merges to report sheet`);
}

// ── Convert Univer IStyleData to xlsx-js-style inline cell.s object ──────────
// Pass skipBg=true when the cell was tinted by zone-indicator (yellow/grey) so
// we don't bake the tint colour into saved files.
function _univerStyleToXlsx(uStyle, skipBg) {
  if (!uStyle) return null;
  const s = {};
  if (uStyle.bl !== undefined || uStyle.it !== undefined || uStyle.fs !== undefined ||
      uStyle.ff !== undefined || uStyle.cl !== undefined || uStyle.ul !== undefined || uStyle.st !== undefined) {
    s.font = {};
    if (uStyle.bl !== undefined) s.font.bold    = !!uStyle.bl;
    if (uStyle.it !== undefined) s.font.italic  = !!uStyle.it;
    if (uStyle.fs !== undefined) s.font.sz      = uStyle.fs;
    if (uStyle.ff !== undefined) s.font.name    = uStyle.ff;
    if (uStyle.cl && uStyle.cl.rgb) s.font.color = { rgb: uStyle.cl.rgb.replace('#','') };
    if (uStyle.ul && uStyle.ul.s)   s.font.underline = true;
    if (uStyle.st && uStyle.st.s)   s.font.strike    = true;
    if (!Object.keys(s.font).length) delete s.font;
  }
  if (!skipBg && uStyle.bg) {
    // bg may be a plain RGB string ('#RRGGBB') or an object with .rgb
    const rawBg = typeof uStyle.bg === 'string' ? uStyle.bg : (uStyle.bg && uStyle.bg.rgb ? uStyle.bg.rgb : null);
    if (rawBg) {
      const rgb = rawBg.toUpperCase().replace('#','');
      // Skip zone-tint colours we injected (yellow #FFFFF0, grey #EEEEEE)
      if (rgb !== 'FFFFF0' && rgb !== 'EEEEEE')
        s.fill = { patternType: 'solid', fgColor: { rgb } };
    }
  }
  if (uStyle.ht !== undefined || uStyle.vt !== undefined || uStyle.tb !== undefined) {
    s.alignment = {};
    const hMap = { 1:'left', 2:'center', 3:'right' };
    const vMap = { 1:'bottom', 2:'middle', 3:'top' };
    if (uStyle.ht) s.alignment.horizontal = hMap[uStyle.ht] || 'left';
    if (uStyle.vt) s.alignment.vertical   = vMap[uStyle.vt] || 'bottom';
    if (uStyle.tb === 3) s.alignment.wrapText = true;
    if (!Object.keys(s.alignment).length) delete s.alignment;
  }
  if (uStyle.bd) {
    const styleMap = { 1:'thin', 2:'medium', 3:'thick', 4:'dashed', 5:'dotted', 6:'double',
                       7:'mediumDashed', 8:'dashDot', 9:'mediumDashDot', 10:'dashDotDot',
                       11:'mediumDashDotDot', 12:'slantDashDot' };
    s.border = {};
    ['t','b','l','r'].forEach(side => {
      const sideKey = { t:'top', b:'bottom', l:'left', r:'right' }[side];
      const bd = uStyle.bd[side];
      if (bd && bd.s)
        s.border[sideKey] = { style: styleMap[bd.s] || 'thin',
          color: { rgb: bd.cl && bd.cl.rgb ? bd.cl.rgb.replace('#','') : '000000' } };
    });
    if (!Object.keys(s.border).length) delete s.border;
  }
  return Object.keys(s).length ? s : null;
}

//  Load Univer CDN scripts once 
let _univerLoading = null;
function _bpLoadUniver() {
  if (_univerLoading) return _univerLoading;
  _univerLoading = (async () => {
    const add = (tag, attrs) => new Promise((res, rej) => {
      const el = document.createElement(tag);
      Object.entries(attrs).forEach(([k, v]) => el[k] = v);
      if (tag === 'script') { el.onload = res; el.onerror = rej; }
      else { el.onload = res; el.onerror = rej; }
      document.head.appendChild(el);
      if (tag === 'link') res(); // link tags don't block
    });
    // Peer deps
    await add('script', { src: 'https://unpkg.com/react@18.3.1/umd/react.production.min.js' });
    await add('script', { src: 'https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js' });
    await add('script', { src: 'https://unpkg.com/rxjs@7.8.1/dist/bundles/rxjs.umd.min.js' });
    // Univer preset CSS + JS
    await add('link',   { rel: 'stylesheet', href: 'https://unpkg.com/@univerjs/preset-sheets-core/lib/index.css' });
    await add('script', { src: 'https://unpkg.com/@univerjs/presets/lib/umd/index.js' });
    await add('script', { src: 'https://unpkg.com/@univerjs/preset-sheets-core/lib/umd/index.js' });
    await add('script', { src: 'https://unpkg.com/@univerjs/preset-sheets-core/lib/umd/locales/en-US.js' });
  })();
  return _univerLoading;
}

//  Convert XLSX workbook to Univer snapshot 
function _bpXlsxToUniverSnap(reportWb, zones, activeConfig) {
  // Build a minimal IWorkbookData snapshot from our worksheet
  const ws    = reportWb.Sheets[reportWb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const cellData = {};

  // Helper: convert xlsx-js-style border side to Univer border side
  function _xlsBorderSide(side) {
    if (!side || !side.style) return undefined;
    const styleMap = { thin: 1, medium: 2, thick: 3, dashed: 4, dotted: 5, double: 6 };
    return { s: styleMap[side.style] || 1, cl: { rgb: '#' + (side.color && side.color.rgb ? side.color.rgb : '000000') } };
  }

  for (let R = range.s.r; R <= range.e.r; R++) {
    cellData[R] = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (!cell) continue;
      const uCell = { v: cell.v != null ? cell.v : '', t: 1 };
      if (cell.s) {
        const s = cell.s;
        uCell.s = {};
        if (s.font) {
          uCell.s.bl = s.font.bold ? 1 : 0;
          uCell.s.it = s.font.italic ? 1 : 0;
          if (s.font.sz)    uCell.s.fs  = s.font.sz;
          if (s.font.color) uCell.s.cl  = { rgb: '#' + (s.font.color.rgb || '000000') };
        }
        if (s.fill && s.fill.fgColor && s.fill.fgColor.rgb)
          uCell.s.bg = '#' + s.fill.fgColor.rgb;
        if (s.alignment) {
          const hMap = { center: 2, right: 3, left: 1 };
          const vMap = { bottom: 1, middle: 2, center: 2, top: 3 };
          if (s.alignment.horizontal) uCell.s.ht = hMap[s.alignment.horizontal] || 1;
          if (s.alignment.vertical)   uCell.s.vt = vMap[s.alignment.vertical]   || 2;
          if (s.alignment.wrapText)   uCell.s.tb = 3;
        }
        // Map border
        if (s.border) {
          const bd = {};
          const t = _xlsBorderSide(s.border.top);
          const b = _xlsBorderSide(s.border.bottom);
          const l = _xlsBorderSide(s.border.left);
          const r = _xlsBorderSide(s.border.right);
          if (t) bd.t = t;
          if (b) bd.b = b;
          if (l) bd.l = l;
          if (r) bd.r = r;
          if (Object.keys(bd).length) uCell.s.bd = bd;
        }
      }
      cellData[R][C] = uCell;
    }
  }

  // Compute footer editable column offset: ONLY /calc menu has fColOff = 2 (protecting cols 0 & 1).
  // All other table types have fColOff = 0 (only protecting table data).
  const fColOff = (activeConfig && activeConfig.fColStart != null)
    ? activeConfig.fColStart
    : (activeConfig && activeConfig.tableType && activeConfig.tableType !== 'calc_menu' ? 0 : 2);

  // Locked cells: all rows except HEADER (0..hCount-1) and FOOTER cols fColOff+ (bS..bE)
  // We store a per-cell lock map: key = "r,c", value = original value
  const lockedCells = {};   // "r,c" -> original value
  const TOTAL_ROWS = range.e.r + 1;
  const TOTAL_COLS = range.e.c + 1;
  for (let r = 0; r < TOTAL_ROWS; r++) {
    const isHeaderRow = r < zones.hCount;
    const isFooterRow = r >= zones.bS && r <= zones.bE;
    for (let c = 0; c < TOTAL_COLS; c++) {
      const isFooterCol = isFooterRow && c >= fColOff;  // cols fColOff+ on footer rows = FOOTER (editable)
      if (isHeaderRow || isFooterCol) continue;   // editable — skip
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      lockedCells[`${r},${c}`] = cell ? cell.v : '';
    }
  }

  const snapshot = {
    id: 'bplan-report',
    sheetOrder: ['sheet1'],
    sheets: {
      sheet1: {
        id: 'sheet1', name: 'Báo cáo',
        cellData,
        rowCount: TOTAL_ROWS,
        columnCount: TOTAL_COLS,
        rowData: {},
        columnData: {},
        defaultRowHeight: 19,
        defaultColumnWidth: 100,
        mergeData: ws['!merges'] ? ws['!merges'].map(m => ({
          startRow: m.s.r, endRow: m.e.r, startColumn: m.s.c, endColumn: m.e.c
        })) : [],
        freeze: { startRow: 0, startColumn: 0, ySplit: 0, xSplit: 0 },
        hidden: 0,
        tabColor: '',
        _lockedCells: lockedCells   // per-cell lock map
      }
    },
    locale: 'vi-VN',
    name: 'Báo cáo BPlan',
    appVersion: '1.0.0',
    styles: {}
  };

  return snapshot;
}

//  Show Univer preview modal 
async function _bpShowPreview(reportWb, zones, tmplWb, tmplBytes, activeConfig) {
  const fColOff = (activeConfig && activeConfig.fColStart != null)
    ? activeConfig.fColStart
    : (activeConfig && activeConfig.tableType && activeConfig.tableType !== 'calc_menu' ? 0 : 2);

  let _pendingMerges = [];
  let _mergeChanges = [];
  await _bpLoadUniver();

  // Dispose previous instance
  if (window._bpUniverAPI) {
    try { window._bpUniverAPI.dispose && window._bpUniverAPI.dispose(); } catch (_) {}
    window._bpUniverAPI = null;
  }
  ['_bpOverlay','_bpToolbar','_bpUniverMount','_bpPopupZFix'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.remove();
  });

  // ── Inject CSS: bump Univer popup z-index above our toolbar (z-index 10010) ──
  // Univer renders .univer-popup at z-index:1070 and .univer-popup-mask at 1060
  // which are both below our toolbar. Override them so dropdowns show on top.
  const zFix = document.createElement('style');
  zFix.id = '_bpPopupZFix';
  zFix.textContent = `.univer-popup { z-index: 10020 !important; }
.univer-popup-mask { z-index: 10015 !important; }
[class*="univer-z-\\[1080\\]"] { z-index: 10020 !important; }
[class*="univer-z-\\[1081\\]"] { z-index: 10021 !important; }`;
  document.head.appendChild(zFix);

  // ── Semi-transparent backdrop ────────────────────────────────────────────────
  const overlay = document.createElement('div');
  overlay.id = '_bpOverlay';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9990;pointer-events:none;';
  document.body.appendChild(overlay);

  // ── Toolbar: fixed bar at top, z-index 10010 ────────────────────────────────
  const toolbar = document.createElement('div');
  toolbar.id = '_bpToolbar';
  toolbar.style.cssText = [
    'position:fixed','top:0','left:0','right:0','height:44px',
    'background:#2E5597','color:#fff','display:flex','align-items:center',
    'justify-content:space-between','padding:0 14px','gap:8px',
    'z-index:10010','font-size:.93rem','box-shadow:0 2px 8px #0005'
  ].join(';');
  toolbar.innerHTML = `
    <span style="font-weight:700;white-space:nowrap">
      <i class="fa-solid fa-file-excel" style="margin-right:6px"></i>
      Xem trước báo cáo
      <span style="font-size:.76rem;opacity:.8;font-weight:400;margin-left:8px">
        Vùng <strong style="color:#ffe082">vàng</strong> = HEADER/FOOTER (có thể sửa) &nbsp;|&nbsp;
        Vùng <strong style="color:#ccc">xám</strong> = dữ liệu (khoá)
      </span>
    </span>
    <div style="display:flex;gap:6px;flex-shrink:0">
      <button id="_bpBtnSaveTmpl" class="btn btn-sm btn-warning">
        <i class="fa-solid fa-floppy-disk"></i> Lưu template
      </button>
      <button id="_bpBtnExport" class="btn btn-sm btn-success">
        <i class="fa-solid fa-download"></i> Tải báo cáo (.xlsx)
      </button>
      <button id="_bpBtnClose" class="btn btn-sm btn-secondary">
        <i class="fa-solid fa-xmark"></i> Đóng
      </button>
    </div>`;
  document.body.appendChild(toolbar);

  // ── Univer host: fills viewport below toolbar ────────────────────────────────
  const univerMount = document.createElement('div');
  univerMount.id = '_bpUniverMount';
  univerMount.style.cssText = [
    'position:fixed','top:44px','left:0','right:0','bottom:0',
    'z-index:10000','background:#fff'
  ].join(';');
  document.body.appendChild(univerMount);

  //  Init Univer
  const { createUniver }             = UniverPresets;
  const { LocaleType, mergeLocales } = UniverCore;
  const { UniverSheetsCorePreset }   = UniverPresetSheetsCore;

  const { univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS) },
    presets: [UniverSheetsCorePreset({ container: univerMount })]
  });
  window._bpUniverAPI = univerAPI;

  //  Build snapshot
  const snap        = _bpXlsxToUniverSnap(reportWb, zones, activeConfig);
  const lockedCells = snap.sheets.sheet1._lockedCells;
  const TOTAL_ROWS  = snap.sheets.sheet1.rowCount;
  const TOTAL_COLS  = snap.sheets.sheet1.columnCount;

  // Initialize _pendingMerges from snapshot mergeData so initial template merges are retained
  _pendingMerges = (snap.sheets.sheet1.mergeData || [])
    .filter(m => {
      const isHeader = m.startRow < zones.hCount && m.endRow < zones.hCount;
      const isFooter = m.startRow >= zones.bS && m.startRow <= zones.bE && m.endRow >= zones.bS && m.endRow <= zones.bE;
      return isHeader || isFooter;
    })
    .map(m => ({
      startRow: m.startRow,
      endRow: m.endRow,
      startColumn: m.startColumn,
      endColumn: m.endColumn,
      isHeader: m.startRow < zones.hCount,
      isFooter: m.startRow >= zones.bS
    }));

  // Remove our private _lockedCells field before handing the snapshot to Univer —
  // unknown fields can cause createWorkbook() to throw or produce a corrupt model.
  delete snap.sheets.sheet1._lockedCells;

  let _populationDone = false;

  //  Create workbook — pass only the clean snapshot
  let univerSheet = null;
  try {
    const wb2 = univerAPI.createWorkbook(snap);
    univerSheet = wb2 ? wb2.getActiveSheet() : null;
  } catch (e) { console.warn('[BPlan] createWorkbook failed:', e); }
  if (!univerSheet) {
    // Fallback: create a blank workbook and populate cells manually below
    try {
      const wb2 = univerAPI.createWorkbook({ id: 'bp', name: 'Báo cáo',
        sheetOrder: ['s1'], sheets: { s1: { id: 's1', name: 'Báo cáo',
          rowCount: TOTAL_ROWS, columnCount: TOTAL_COLS } },
        locale: 'vi-VN', name: 'Báo cáo BPlan', appVersion: '1.0.0', styles: {} });
      univerSheet = wb2 ? wb2.getActiveSheet() : null;
      if (!univerSheet) { const wb3 = univerAPI.getActiveWorkbook(); univerSheet = wb3 ? wb3.getActiveSheet() : null; }
    } catch (e) { console.warn('[BPlan] fallback createWorkbook failed:', e); }
  }

  //  Populate cells + zone tint
  if (univerSheet) {
    const sheetData = snap.sheets.sheet1.cellData;
    for (let r = 0; r < TOTAL_ROWS; r++) {
      const rowData = sheetData[r] || {};
      for (let c = 0; c < TOTAL_COLS; c++) {
        const isLocked = (`${r},${c}`) in lockedCells;
        const cell = rowData[c];
        try {
          const rng = univerSheet.getRange(r, c);
          const v = cell ? cell.v : '';
          rng.setValue(v !== undefined ? v : '');
          if (cell && cell.s) {
            const s = cell.s;
            if (s.bl !== undefined) rng.setFontWeight(s.bl ? 'bold' : 'normal');
            if (s.it !== undefined) rng.setFontStyle(s.it ? 'italic' : 'normal');
            if (s.fs) rng.setFontSize(s.fs);
            if (s.cl && s.cl.rgb) rng.setFontColor(s.cl.rgb);
            if (s.bg) rng.setBackground(s.bg);
            if (s.ht) {
              const hMap = { 1: 'left', 2: 'center', 3: 'right' };
              if (hMap[s.ht]) rng.setHorizontalAlignment(hMap[s.ht]);
            }
          }
          const hasBg = cell && cell.s && cell.s.bg;
          if (!hasBg) rng.setBackground(isLocked ? '#EEEEEE' : '#FFFFF0');
        } catch (_) {}
      }
    }
    (snap.sheets.sheet1.mergeData || []).forEach(m => {
      try {
        univerSheet.getRange(
          m.startRow, m.startColumn,
          m.endRow - m.startRow + 1,
          m.endColumn - m.startColumn + 1
        ).merge();
      } catch (_) {}
    });
  }
  // Mark population complete — commands fired AFTER this point are user edits
  _populationDone = true;

  // ── Change tracking ──────────────────────────────────────────────────────────
  //
  // We use TWO complementary paths so edits are never missed:
  //
  // Path A — mutation mirror (real-time):
  //   Univer fires onCommandExecuted for MUTATIONS (not commands). The mutation
  //   that carries cell writes is 'sheet.mutation.set-range-values'. Its params
  //   carry the exact {row:{col:ICellData}} map written — read directly, no Facade.
  //   Also listen to style/format mutations so toolbar changes are captured.
  //
  // Path B — full scan on button click (safety net):
  //   Before export/save we call _scanEditableCells() which reads every editable
  //   cell with getValue() + getCellStyleData(). This is the ground truth.
  //   getValue() on FRange works for any arbitrary cell (confirmed from source).

  const _changes = {};  // "r,c" -> { v, uStyle }

  function _isEditable(r, c) {
    if (r < zones.hCount) return true;
    if (r >= zones.bS && r <= zones.bE && c >= fColOff) return true;
    return false;
  }

  // Helper to get effective merges (from live Univer sheet or _pendingMerges fallback)
  function _getEffectiveMerges() {
    const scanned = _scanMergeData();
    if (scanned && scanned.length > 0) {
      return scanned;
    }
    return _pendingMerges || [];
  }

  // Path B — full scan via Facade getValue() + getCellStyleData()
  // FIXED: Skip cells that are non-origin cells in merged ranges to avoid piling up content
  function _scanEditableCells() {
    if (!univerSheet) return;
    
    _mergeChanges = _getEffectiveMerges();

    // Build a map of merged cells from _mergeChanges
    const mergedCellMap = {}; // key: "r,c" -> { isOrigin: boolean, originCell: "r,c" }
    for (const m of _mergeChanges) {
      for (let r = m.startRow; r <= m.endRow; r++) {
        for (let c = m.startColumn; c <= m.endColumn; c++) {
          const key = `${r},${c}`;
          mergedCellMap[key] = {
            isOrigin: r === m.startRow && c === m.startColumn,
            originCell: `${m.startRow},${m.startColumn}`,
            isHeader: m.isHeader,
            isFooter: m.isFooter
          };
        }
      }
    }
    
    let scannedCount = 0;
    let skippedCount = 0;
    
    for (let r = 0; r < TOTAL_ROWS; r++) {
      const isHeader = r < zones.hCount;
      const isFooter = r >= zones.bS && r <= zones.bE;
      if (!isHeader && !isFooter) continue;
      const colStart = (isFooter && !isHeader) ? fColOff : 0;
      for (let c = colStart; c < TOTAL_COLS; c++) {
        if (!_isEditable(r, c)) continue;
        
        // FIXED: Skip non-origin cells in merged ranges - they should not have separate values
        const cellKey = `${r},${c}`;
        const mergeInfo = mergedCellMap[cellKey];
        if (mergeInfo && !mergeInfo.isOrigin) {
          // This is a non-origin cell in a merge - skip it
          // The value will be read from the origin cell only
          skippedCount++;
          continue;
        }
        
        let v = '', uStyle = null;
        try {
          const rng = univerSheet.getRange(r, c);
          try { const raw = rng.getValue(); v = raw != null ? raw : ''; } catch (_e) {}
          try { if (typeof rng.getCellStyleData === 'function') uStyle = rng.getCellStyleData('cell'); } catch (_e) {}
        } catch (_e) {}
        _changes[`${r},${c}`] = { v, uStyle };
        scannedCount++;
      }
    }
    
    console.log('[BPlan] _scanEditableCells: scanned', scannedCount, 'cells, skipped', skippedCount, 'merged non-origins');
    console.log('[BPlan] Sample scanned values (first 5):');
    let sampleCount = 0;
    for (const key in _changes) {
      if (sampleCount >= 5) break;
      const [r, c] = key.split(',').map(Number);
      const isHeader = r < zones.hCount;
      console.log(`  [${isHeader ? 'HEADER' : 'FOOTER'}] Cell ${key}: v="${_changes[key].v}"`);
      sampleCount++;
    }
  }
  
  // Helper to extract numeric bounds from any Univer merge object representation
  function _extractMergeBounds(m) {
    if (!m || typeof m !== 'object') return null;
    
    let sR = m.startRow != null ? m.startRow : (typeof m.getStartRow === 'function' ? m.getStartRow() : m._startRow);
    let eR = m.endRow != null ? m.endRow : (typeof m.getEndRow === 'function' ? m.getEndRow() : m._endRow);
    let sC = m.startColumn != null ? m.startColumn : (typeof m.getStartColumn === 'function' ? m.getStartColumn() : m._startColumn);
    let eC = m.endColumn != null ? m.endColumn : (typeof m.getEndColumn === 'function' ? m.getEndColumn() : m._endColumn);

    if ((sR == null || isNaN(sR)) && (m.range || m._range)) {
      const r = m.range || m._range;
      sR = r.startRow != null ? r.startRow : r._startRow;
      eR = r.endRow != null ? r.endRow : r._endRow;
      sC = r.startColumn != null ? r.startColumn : r._startColumn;
      eC = r.endColumn != null ? r.endColumn : r._endColumn;
    }

    if ((sR == null || isNaN(sR)) && typeof m.getRange === 'function') {
      try {
        const r = m.getRange();
        if (r) {
          sR = r.startRow != null ? r.startRow : r.row;
          eR = r.endRow != null ? r.endRow : (r.row != null ? r.row + (r.rowCount || 1) - 1 : undefined);
          sC = r.startColumn != null ? r.startColumn : r.col;
          eC = r.endColumn != null ? r.endColumn : (r.col != null ? r.col + (r.colCount || 1) - 1 : undefined);
        }
      } catch (_) {}
    }

    if (sR == null || isNaN(sR)) {
      if (m.row != null) {
        sR = m.row;
        eR = m.row + (m.rowCount || 1) - 1;
        sC = m.col;
        eC = m.col + (m.colCount || 1) - 1;
      } else if (m.startY != null) {
        sR = m.startY; eR = m.endY;
        sC = m.startX; eC = m.endX;
      }
    }

    if (sR != null && eR != null && sC != null && eC != null && !isNaN(sR) && !isNaN(eR) && !isNaN(sC) && !isNaN(eC)) {
      return { startRow: Number(sR), endRow: Number(eR), startColumn: Number(sC), endColumn: Number(eC) };
    }
    return null;
  }

  // ── Scan merge data from Univer sheet ─────────────────────────────────────────
  // Returns array of merge ranges in the editable regions (HEADER and FOOTER)
  function _scanMergeData() {
    if (!univerSheet) {
      console.log('[BPlan] _scanMergeData: No univerSheet available');
      return [];
    }
    
    const merges = [];
    try {
      let mergeData = null;
      if (typeof univerSheet.getMergeData === 'function') {
        mergeData = univerSheet.getMergeData();
      } else if (typeof univerSheet.getSheetData === 'function') {
        const sheetData = univerSheet.getSheetData();
        if (sheetData) mergeData = sheetData.mergeData;
      }
      
      let rawItems = [];
      if (Array.isArray(mergeData)) {
        rawItems = mergeData;
      } else if (mergeData && typeof mergeData === 'object') {
        if (Array.isArray(mergeData._collection)) {
          rawItems = mergeData._collection;
        } else if (mergeData._collection && typeof mergeData._collection.values === 'function') {
          rawItems = Array.from(mergeData._collection.values());
        } else if (mergeData._collection && typeof mergeData._collection === 'object') {
          rawItems = Object.values(mergeData._collection);
        } else if (typeof mergeData.getMerges === 'function') {
          rawItems = mergeData.getMerges();
        } else {
          rawItems = Object.values(mergeData).filter(x => x && typeof x === 'object');
        }
      }

      console.log('[BPlan] Found', rawItems.length, 'raw merge items from Univer');

      for (const raw of rawItems) {
        const parsed = _extractMergeBounds(raw);
        if (!parsed) continue;
        const { startRow, endRow, startColumn, endColumn } = parsed;
        if (startRow === endRow && startColumn === endColumn) continue; // skip 1x1

        const isHeaderMerge = startRow < zones.hCount;
        const isFooterMerge = startRow >= zones.bS;

        if (isHeaderMerge || isFooterMerge) {
          merges.push({
            startRow, endRow, startColumn, endColumn,
            isHeader: isHeaderMerge, isFooter: isFooterMerge
          });
        }
      }
    } catch (e) {
      console.warn('[BPlan] Could not scan merge data from Univer:', e);
    }
    
    console.log('[BPlan] _scanMergeData returning', merges.length, 'merges');
    return merges;
  }
  
  // Store for merge data captured from Univer
  _mergeChanges = [];
  
  // Store merge data directly from command mutations
  _pendingMerges = [];  // Merges from mutation commands

  // Path A — mutation-level intercept with correct mutation IDs
  // Confirmed IDs from Univer source: mutations fire through onCommandExecuted.
  const WRITE_MUTATIONS = [
    'sheet.mutation.set-range-values',          // cell value + style writes
    'sheet.mutation.set-range-styles',           // toolbar style changes (bold, colour…)
    'sheet.mutation.set-range-formatted-values', // paste with format
    'sheet.mutation.clear-selection-content',
    'sheet.mutation.remove-worksheet-merge',
    'sheet.mutation.add-worksheet-merge',
  ];
  
  const MERGE_MUTATIONS = [
    'sheet.mutation.add-worksheet-merge',
    'sheet.mutation.remove-worksheet-merge',
  ];

  let _interceptor = null;
  try {
    _interceptor = univerAPI.onCommandExecuted && univerAPI.onCommandExecuted((cmd) => {
      if (!_populationDone) return;
      if (!WRITE_MUTATIONS.includes(cmd.id)) return;

      // Handle merge mutations - extract merge data directly from command params
      if (MERGE_MUTATIONS.includes(cmd.id)) {
        console.log('[BPlan] Merge mutation:', cmd.id, cmd.params);
        
        const p = cmd.params || {};
        
        if (cmd.id === 'sheet.mutation.add-worksheet-merge') {
          // Extract merge range from params
          // Univer uses: { unitId, subUnitId, ranges: [{startRow, endRow, startColumn, endColumn}] }
          const ranges = p.ranges || [];
          console.log('[BPlan] Merge ranges from params:', ranges);
          
          for (const range of ranges) {
            const { startRow, endRow, startColumn, endColumn } = range;
            
            // Check if in editable region
            const isHeaderMerge = startRow < zones.hCount && endRow < zones.hCount;
            const isFooterMerge = startRow >= zones.bS && startRow <= zones.bE && 
                                   endRow >= zones.bS && endRow <= zones.bE;
            
            if (isHeaderMerge || isFooterMerge) {
              _pendingMerges.push({
                startRow,
                endRow,
                startColumn,
                endColumn,
                isHeader: isHeaderMerge,
                isFooter: isFooterMerge
              });
            }
          }
        } else if (cmd.id === 'sheet.mutation.remove-worksheet-merge') {
          // Remove the matching merge from _pendingMerges
          const ranges = p.ranges || [];
          for (const range of ranges) {
            _pendingMerges = _pendingMerges.filter(m => 
              !(m.startRow === range.startRow && m.endRow === range.endRow &&
                m.startColumn === range.startColumn && m.endColumn === range.endColumn)
            );
          }
        }
        
        _mergeChanges = [..._pendingMerges];
        console.log('[BPlan] Updated merge changes:', _mergeChanges.length, 'merges');
      }

      // Extract cell value map directly from mutation params (no Facade read needed)
      try {
        const p = cmd.params || {};
        // set-range-values: params.cellValue = {[row]:{[col]:ICellData}}
        const cv = p.cellValue || p.value || null;
        if (cv && typeof cv === 'object') {
          console.log('[BPlan] Mutation cellValue:', cmd.id, 'cells:', Object.keys(cv).length);
          for (const rStr of Object.keys(cv)) {
            const r = Number(rStr);
            const cols = cv[rStr];
            if (!cols || typeof cols !== 'object') continue;
            for (const cStr of Object.keys(cols)) {
              const c = Number(cStr);
              if (!_isEditable(r, c)) continue;
              const cell = cols[cStr];
              if (cell === null) { _changes[`${r},${c}`] = { v: '', uStyle: null }; continue; }
              const v = cell && cell.v != null ? cell.v : '';
              const uStyle = cell && cell.s && typeof cell.s === 'object' ? cell.s : null;
              _changes[`${r},${c}`] = { v, uStyle };
              console.log(`[BPlan] Captured mutation change: [${r},${c}] = "${v}"`);
            }
          }
        }
      } catch (_e) {}

      // Revert locked cells after a short settle
      setTimeout(() => {
        if (!univerSheet) return;
        for (const key in lockedCells) {
          const [ri, ci] = key.split(',').map(Number);
          const val = lockedCells[key];
          try {
            const cur = univerSheet.getRange(ri, ci).getValue();
            if (String(cur) !== String(val)) univerSheet.getRange(ri, ci).setValue(val != null ? val : '');
          } catch (_e) {}
        }
      }, 150);
    });
  } catch (_) {}

  function _cleanup() {
    if (_interceptor && typeof _interceptor.dispose === 'function') _interceptor.dispose();
    try { univerAPI.dispose && univerAPI.dispose(); } catch (_) {}
    window._bpUniverAPI = null;
    ['_bpOverlay','_bpToolbar','_bpUniverMount','_bpPopupZFix'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.remove();
    });
  }


  // ── Helper: export using _changes map (captured synchronously after edits) ──
  // Avoids unreliable Facade API bulk reads. Patches only changed cells onto reportWb.
  // Preserves original cell styles; only overlays style properties that Univer actually returned.
  // Also applies merge changes from Univer.
  // FIXED: Skip non-origin cells in merged ranges to avoid content piling up
  function _exportFromUniver() {
    try {
      const origWs = reportWb.Sheets[reportWb.SheetNames[0]];
      const patchedWs = JSON.parse(JSON.stringify(origWs));

      // Build a map of merged cells from _mergeChanges
      const mergedCellMap = {};
      for (const m of _mergeChanges) {
        for (let r = m.startRow; r <= m.endRow; r++) {
          for (let c = m.startColumn; c <= m.endColumn; c++) {
            const key = `${r},${c}`;
            mergedCellMap[key] = {
              isOrigin: r === m.startRow && c === m.startColumn,
              originCell: `${m.startRow},${m.startColumn}`
            };
          }
        }
      }

      for (const key in _changes) {
        const [r, c] = key.split(',').map(Number);
        if (!_isEditable(r, c)) continue;
        
        // FIXED: Skip non-origin cells in merged ranges
        const cellKey = `${r},${c}`;
        const mergeInfo = mergedCellMap[cellKey];
        if (mergeInfo && !mergeInfo.isOrigin) {
          // Non-origin cells in merges should not have separate values
          // The merge handles the visual display
          continue;
        }
        
        const { v, uStyle } = _changes[key];
        const addr = XLSX.utils.encode_cell({ r, c });

        // Ensure cell exists in patched sheet
        if (!patchedWs[addr]) {
          patchedWs[addr] = { v: '', t: 's' };
          const rng = XLSX.utils.decode_range(patchedWs['!ref'] || 'A1');
          if (r > rng.e.r) rng.e.r = r;
          if (c > rng.e.c) rng.e.c = c;
          patchedWs['!ref'] = XLSX.utils.encode_range(rng);
        }

        // Update value
        const isNum = typeof v === 'number';
        patchedWs[addr].v = isNum ? v : String(v ?? '');
        patchedWs[addr].t = isNum ? 'n' : 's';

        // Merge style: keep original styles and only overlay what Univer actually returned.
        // This preserves merges, borders, and alignment that Univer may not report back.
        if (uStyle) {
          const xlsxStyle = _univerStyleToXlsx(uStyle, true);
          if (xlsxStyle && Object.keys(xlsxStyle).length) {
            const origStyle = patchedWs[addr].s || {};
            // Deep-merge: xlsxStyle keys take precedence over origStyle, but
            // missing keys in xlsxStyle fall back to origStyle to preserve borders etc.
            const merged = JSON.parse(JSON.stringify(origStyle));
            for (const prop of ['font', 'fill', 'alignment']) {
              if (xlsxStyle[prop]) {
                merged[prop] = Object.assign({}, origStyle[prop] || {}, xlsxStyle[prop]);
              }
            }
            // Border: only replace if Univer returned a complete, non-empty border object
            if (xlsxStyle.border && Object.keys(xlsxStyle.border).length >= 2) {
              merged.border = Object.assign({}, origStyle.border || {}, xlsxStyle.border);
            }
            patchedWs[addr].s = merged;
          }
          // If xlsxStyle is null/empty, keep the original cell style untouched
        }
      }
      
      // Apply merge changes from Univer
      if (_mergeChanges.length > 0) {
        // Remove existing merges in editable regions and replace with Univer merges
        const existingMerges = patchedWs['!merges'] || [];
        const newMerges = existingMerges.filter(m => {
          // Keep merges that are NOT in editable regions (i.e., in the data table)
          const isInHeader = m.s.r < zones.hCount;
          const isInFooter = m.s.r >= zones.bS && m.s.r <= zones.bE;
          return !isInHeader && !isInFooter;
        });
        
        // Add merges from Univer
        for (const m of _mergeChanges) {
          newMerges.push({
            s: { r: m.startRow, c: m.startColumn },
            e: { r: m.endRow, c: m.endColumn }
          });
        }
        
        patchedWs['!merges'] = newMerges;
        console.log('[BPlan] Applied', _mergeChanges.length, 'merge changes to report');
      }

      // Auto-adjust column widths for full report (delete stale !cols first so fresh calc always wins)
      console.log('[BPlan][AutoFit] patchedWs !cols BEFORE delete:', JSON.stringify(patchedWs['!cols']));
      delete patchedWs['!cols'];
      console.log('[BPlan][AutoFit] patchedWs !cols AFTER delete (should be undefined):', patchedWs['!cols']);
      console.log('[BPlan][AutoFit] zones passed to _bpAutoFitCols:', JSON.stringify(zones));
      _bpAutoFitCols(patchedWs, 2, 80, zones);
      console.log('[BPlan][AutoFit] patchedWs !cols AFTER _bpAutoFitCols:', JSON.stringify(patchedWs['!cols']));

      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, patchedWs, 'Báo cáo');
      return _bpWriteXlsx(newWb);
    } catch (err) {
      console.error('[BPlan][AutoFit] _exportFromUniver threw exception — falling back to reportWb:', err);
      return null;
    }
  }

  // Helper to auto-adjust column widths matching Excel's native AutoFit behavior
  function _bpAutoFitCols(ws, minW = 2, maxW = 80, zones = null) {
    if (!ws) return;
    console.log('[BPlan][AutoFit] _bpAutoFitCols called. minW:', minW, 'maxW:', maxW, 'zones:', JSON.stringify(zones));
    console.log('[BPlan][AutoFit] Total ws keys:', Object.keys(ws).filter(k => !k.startsWith('!')).length);
    const colWidths = {};
    
    // Ignore multi-column merged cells for column width calculations
    const multiColMerges = new Set();
    if (ws['!merges'] && Array.isArray(ws['!merges'])) {
      ws['!merges'].forEach(m => {
        if (m.s.c !== m.e.c) {
          for (let r = m.s.r; r <= m.e.r; r++) {
            for (let c = m.s.c; c <= m.e.c; c++) {
              multiColMerges.add(`${r},${c}`);
            }
          }
        }
      });
    }
    console.log('[BPlan][AutoFit] multiColMerges count:', multiColMerges.size);

    let maxCol = 0;
    let skippedZone = 0, skippedMerge = 0, skippedEmpty = 0, measured = 0;
    const keys = Object.keys(ws).filter(k => !k.startsWith('!'));
    keys.forEach(k => {
      const cell = XLSX.utils.decode_cell(k);
      const c = cell.c;
      const r = cell.r;
      if (c > maxCol) maxCol = c;
      
      // If zones provided, restrict auto-fit measuring to table rows only
      if (zones && zones.tS != null && zones.tE != null) {
        if (r < zones.tS || r > zones.tE) { skippedZone++; return; }
      }

      if (multiColMerges.has(`${r},${c}`)) { skippedMerge++; return; }

      const rawVal = ws[k] ? ws[k].v : null;
      if (rawVal == null || rawVal === '') { skippedEmpty++; return; }

      let str = String(rawVal);
      if (typeof rawVal === 'number') {
        str = rawVal.toLocaleString('vi');
        if (ws[k].z && /VND/i.test(ws[k].z)) str += ' VND';
        else if (ws[k].z && /%/i.test(ws[k].z)) str += '%';
      }
      const lines = str.split('\n');
      let maxLen = 0;
      lines.forEach(line => {
        const w = typeof _calcTextWidth === 'function' ? _calcTextWidth(line) : _bpTextWidth(line);
        if (w > maxLen) maxLen = w;
      });
      const autoWidth = Math.ceil(maxLen + 1);
      console.log(`[BPlan][AutoFit]   cell ${k} (r=${r},c=${c}) val=${JSON.stringify(str)} → w=${maxLen.toFixed(1)} → autoWidth=${autoWidth}`);
      measured++;
      if (!colWidths[c] || autoWidth > colWidths[c]) {
        colWidths[c] = autoWidth;
      }
    });
    console.log(`[BPlan][AutoFit] Scan done: measured=${measured}, skippedZone=${skippedZone}, skippedMerge=${skippedMerge}, skippedEmpty=${skippedEmpty}`);
    console.log('[BPlan][AutoFit] colWidths per col:', JSON.stringify(colWidths));

    // Evaluate multi-column merges to ensure spanned text is not truncated
    if (ws['!merges'] && Array.isArray(ws['!merges'])) {
      ws['!merges'].forEach(m => {
        const colSpan = m.e.c - m.s.c + 1;
        if (colSpan > 1) {
          const originAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
          const cellObj = ws[originAddr];
          const rawVal = cellObj ? cellObj.v : null;
          if (rawVal != null && rawVal !== '') {
            let str = String(rawVal);
            if (typeof rawVal === 'number') {
              str = rawVal.toLocaleString('vi');
            }
            let maxLen = 0;
            str.split('\n').forEach(line => {
              let w = typeof _calcTextWidth === 'function' ? _calcTextWidth(line) : _bpTextWidth(line);
              if (cellObj && cellObj.s && cellObj.s.font && cellObj.s.font.bold) {
                w *= 1.1; // 10% wider for bold header text
              }
              if (w > maxLen) maxLen = w;
            });
            // Multi-column merges need extra padding (+3.5) due to cell margins across gridlines
            const neededTotalW = Math.ceil(maxLen + 3.5);
            
            // Calculate current total width of all spanned columns
            let currentTotalW = 0;
            for (let c = m.s.c; c <= m.e.c; c++) {
              currentTotalW += (colWidths[c] || minW);
            }
            
            // If total width is insufficient, distribute missing deficit evenly
            if (neededTotalW > currentTotalW) {
              const deficit = neededTotalW - currentTotalW;
              const addPerCol = Math.ceil(deficit / colSpan);
              for (let c = m.s.c; c <= m.e.c; c++) {
                colWidths[c] = (colWidths[c] || minW) + addPerCol;
              }
            }
          }
        }
      });
    }

    const cols = [];
    for (let c = 0; c <= Math.max(maxCol, 0); c++) {
      const fitW = colWidths[c] != null ? Math.max(minW, Math.min(colWidths[c], maxW)) : minW;
      cols[c] = { wch: fitW };
    }
    console.log('[BPlan][AutoFit] Final !cols:', JSON.stringify(cols));
    ws['!cols'] = cols;
  }



  //  Buttons
  document.getElementById('_bpBtnClose').onclick = _cleanup;

  document.getElementById('_bpBtnExport').onclick = () => {
    const fileName = `BPlan_BaoCao_${new Date().toISOString().slice(0,10)}.xlsx`;
    // Path B safety-net: scan all editable cells right before export so we
    // capture the latest state even if the mutation interceptor missed anything.
    _scanEditableCells();
    console.log('[BPlan] Export - using effective merges:', _mergeChanges.length);
    const bytes = _exportFromUniver();
    if (!bytes) {
      console.warn('[BPlan] Export - _exportFromUniver returned null! Using fallback reportWb (NO AUTOFIT).');
      console.log('[BPlan] Export fallback reportWb !cols:', JSON.stringify(reportWb.Sheets[reportWb.SheetNames[0]]['!cols']));
    } else {
      console.log('[BPlan] Export - _exportFromUniver succeeded, bytes length:', bytes.length);
    }
    _bpDL(bytes || _bpWriteXlsx(reportWb), fileName);
  };

  document.getElementById('_bpBtnSaveTmpl').onclick = async () => {
    _scanEditableCells();   // fresh scan before saving
    
    console.log('[BPlan] === SAVE TEMPLATE DEBUG ===');
    console.log('[BPlan] Scanned changes:', Object.keys(_changes).length);
    console.log('[BPlan] Pending merges:', _pendingMerges.length);
    console.log('[BPlan] Zones:', zones);
    
    try {
      // Create a DEEP COPY of the template workbook to avoid modifying the original
      // This ensures each save operation works with a fresh copy
      const hSheet = JSON.parse(JSON.stringify(tmplWb.Sheets[tmplWb.SheetNames[0]]));
      const fSheet = tmplWb.SheetNames[1] ? JSON.parse(JSON.stringify(tmplWb.Sheets[tmplWb.SheetNames[1]])) : null;
      
      // Create a new workbook object for the saved template
      const savedTmplWb = {
        SheetNames: [...tmplWb.SheetNames],
        Sheets: {}
      };
      savedTmplWb.Sheets[savedTmplWb.SheetNames[0]] = hSheet;
      if (fSheet) {
        savedTmplWb.Sheets[savedTmplWb.SheetNames[1]] = fSheet;
      }

      // Debug: Show what's in the template sheets BEFORE saving
      console.log('[BPlan] Template HEADER sheet cells BEFORE save:', 
        Object.keys(hSheet).filter(k => !k.startsWith('!')));
      if (fSheet) {
        console.log('[BPlan] Template FOOTER sheet cells BEFORE save:', 
          Object.keys(fSheet).filter(k => !k.startsWith('!')));
      }
      
      // Debug: Show what changes will be applied
      const headerChanges = Object.entries(_changes).filter(([k, v]) => {
        const [r, c] = k.split(',').map(Number);
        return r < zones.hCount;
      });
      const footerChanges = Object.entries(_changes).filter(([k, v]) => {
        const [r, c] = k.split(',').map(Number);
        return r >= zones.bS && r <= zones.bE && c >= fColOff;
      });
      console.log('[BPlan] Header changes to apply:', headerChanges.length);
      console.log('[BPlan] Footer changes to apply:', footerChanges.length);
      
      // Debug: Show sample of actual values being applied
      console.log('[BPlan] Sample header changes:');
      headerChanges.slice(0, 5).forEach(([k, val]) => {
        console.log(`  Cell ${k}: v="${val.v}"`);
      });
      console.log('[BPlan] Sample footer changes:');
      footerChanges.slice(0, 5).forEach(([k, val]) => {
        console.log(`  Cell ${k}: v="${val.v}"`);
      });

      // Build a map of merged cells to avoid writing to non-origin cells
      const mergedCellMap = {};
      for (const m of _mergeChanges) {
        for (let r = m.startRow; r <= m.endRow; r++) {
          for (let c = m.startColumn; c <= m.endColumn; c++) {
            const key = `${r},${c}`;
            mergedCellMap[key] = {
              isOrigin: r === m.startRow && c === m.startColumn,
              originCell: `${m.startRow},${m.startColumn}`,
              isHeader: m.isHeader,
              isFooter: m.isFooter
            };
          }
        }
      }

      for (const key in _changes) {
        const [r, c] = key.split(',').map(Number);
        const { v, uStyle } = _changes[key];
        const isHeader = r < zones.hCount;
        const isFooter = r >= zones.bS && r <= zones.bE && c >= fColOff;
        
        // Check if this cell is a non-origin in a merged range
        const cellKey = `${r},${c}`;
        const mergeInfo = mergedCellMap[cellKey];

        let targetSheet, targetAddr;
        if (isHeader) {
          targetSheet = hSheet;
          // Template HEADER starts at Excel row 1 (row 0 in 0-indexed)
          targetAddr = XLSX.utils.encode_cell({ r: r, c });
        } else if (isFooter && fSheet) {
          targetSheet = fSheet;
          // Transform: report row r → template row (r - bS), col c → (c - fColOff)
          targetAddr = XLSX.utils.encode_cell({ r: r - zones.bS, c: c - fColOff });
        } else continue;

        if (mergeInfo && !mergeInfo.isOrigin) {
          // Non-origin cells in merged ranges should be empty in template
          if (targetSheet[targetAddr]) {
            targetSheet[targetAddr].v = '';
            targetSheet[targetAddr].t = 's';
          }
          continue;
        }

        if (!targetSheet[targetAddr]) {
          targetSheet[targetAddr] = { v: '', t: 's' };
        }

        const isNum = typeof v === 'number';
        const strVal = v != null ? String(v) : '';
        const oldVal = targetSheet[targetAddr].v;

        targetSheet[targetAddr].v = isNum ? v : strVal;
        targetSheet[targetAddr].t = isNum ? 'n' : 's';
        const sheetName = isHeader ? 'HEADER' : 'FOOTER';
        console.log(`[BPlan] Wrote to ${sheetName} sheet ${targetAddr}: "${oldVal}" -> "${targetSheet[targetAddr].v}"`);

        // Merge style from Univer onto original template cell style
        if (uStyle && targetSheet[targetAddr]) {
          const xlsxStyle = _univerStyleToXlsx(uStyle, true);
          if (xlsxStyle && Object.keys(xlsxStyle).length) {
            const orig = targetSheet[targetAddr].s;
            const origStyle = (orig && typeof orig === 'object') ? orig : {};
            const merged = JSON.parse(JSON.stringify(origStyle));
            for (const prop of ['font', 'fill', 'alignment']) {
              if (xlsxStyle[prop]) merged[prop] = Object.assign({}, origStyle[prop] || {}, xlsxStyle[prop]);
            }
            if (xlsxStyle.border && Object.keys(xlsxStyle.border).length >= 2) {
              merged.border = Object.assign({}, origStyle.border || {}, xlsxStyle.border);
            }
            targetSheet[targetAddr].s = merged;
          }
        }
      }
      
      // Apply merge changes to template sheets
      if (_mergeChanges.length > 0) {
        // HEADER merges - apply to header sheet (rows 0..hCount-1)
        const headerMerges = _mergeChanges.filter(m => m.isHeader);
        if (headerMerges.length > 0 && hSheet) {
          hSheet['!merges'] = headerMerges.map(m => ({
            s: { r: m.startRow, c: m.startColumn },
            e: { r: m.endRow, c: m.endColumn }
          }));
          console.log('[BPlan] Applied', headerMerges.length, 'header merges to template');
        }
        
        // FOOTER merges - apply to footer sheet with coordinate transformation
        const footerMerges = _mergeChanges.filter(m => m.isFooter);
        if (footerMerges.length > 0 && fSheet) {
          fSheet['!merges'] = footerMerges.map(m => ({
            s: { r: m.startRow - zones.bS, c: m.startColumn - fColOff },
            e: { r: m.endRow - zones.bS, c: m.endColumn - fColOff }
          }));
          console.log('[BPlan] Applied', footerMerges.length, 'footer merges to template');
        }
      }

      // Helper to update !ref bounds on sheet
      const _bpUpdateSheetRef = (ws) => {
        if (!ws) return;
        const keys = Object.keys(ws).filter(k => !k.startsWith('!'));
        if (!keys.length) return;
        let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
        keys.forEach(k => {
          const cell = XLSX.utils.decode_cell(k);
          if (cell.r < minR) minR = cell.r;
          if (cell.c < minC) minC = cell.c;
          if (cell.r > maxR) maxR = cell.r;
          if (cell.c > maxC) maxC = cell.c;
        });
        ws['!ref'] = XLSX.utils.encode_range({
          s: { r: minR, c: minC },
          e: { r: maxR, c: maxC }
        });
      };

      _bpUpdateSheetRef(hSheet);
      if (fSheet) _bpUpdateSheetRef(fSheet);

      // Auto-adjust column widths for template sheets
      _bpAutoFitCols(hSheet, 12);
      if (fSheet) _bpAutoFitCols(fSheet, 14);

      console.log('[BPlan] HEADER sheet !ref AFTER save:', hSheet['!ref']);
      if (fSheet) console.log('[BPlan] FOOTER sheet !ref AFTER save:', fSheet['!ref']);

      const arr = _bpWriteXlsx(savedTmplWb);
      const cacheKey = (activeConfig && activeConfig.cacheKey) ? activeConfig.cacheKey : _BP_TMPL_KEY;
      const fileName = (activeConfig && activeConfig.tableType) ? `BPlan_Template_${activeConfig.tableType}.xlsx` : 'BPlan_Template.xlsx';
      _bpSaveTmplCache(new Uint8Array(arr), cacheKey);
      _bpDL(arr, fileName);

      Swal.fire({ icon: 'success', title: 'Đã lưu template',
        text: 'Template đã được cache và tải xuống.', timer: 2200, showConfirmButton: false });
    } catch (e) {
      Swal.fire({ icon: 'warning', title: 'Lưu thất bại', text: e.message });
    }
  };
}

async function bplanExportReport(getCurrentResult, getTotalPeople, getBudgetVal) {
  const currentResult = getCurrentResult();
  if (!currentResult || !currentResult.details) {
    Swal.fire({ icon: 'warning', title: 'Chưa có kết quả',
      text: 'Hãy tính toán trước khi xuất báo cáo.' });
    return;
  }

  const totalPeople = getTotalPeople();
  const budgetVal   = getBudgetVal();
  const cached      = _bpLoadTmplCache();

  //  Step 1: Template picker
  const { value: choice } = await Swal.fire({
    title: 'Chọn template báo cáo',
    html: `
      <div style="display:flex;flex-direction:column;gap:10px;text-align:left;font-size:.93rem">
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="cached" checked>
          ${cached ? 'Dùng template đã lưu (cache)' : 'Dùng template mặc định'}
        </label>
        ${cached ? `
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="default">
          Dùng template mặc định
        </label>` : ''}
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="upload">
          Tải lên file template (.xlsx)
        </label>
        <div id="_bpUpArea" style="display:none;margin-top:2px">
          <input type="file" id="_bpTmplFile" accept=".xlsx,.xls"
            style="font-size:.88em;width:100%">
        </div>
      </div>`,
    didOpen: () => {
      document.querySelectorAll('input[name="_bpTmpl"]').forEach(r => {
        r.addEventListener('change', () => {
          document.getElementById('_bpUpArea').style.display =
            r.value === 'upload' && r.checked ? 'block' : 'none';
        });
      });
    },
    confirmButtonText: 'Tiếp tục',
    showCancelButton: true,
    preConfirm: () => {
      const sel = document.querySelector('input[name="_bpTmpl"]:checked').value;
      if (sel === 'upload') {
        const f = document.getElementById('_bpTmplFile').files[0];
        if (!f) { Swal.showValidationMessage('Vui lòng chọn file template'); return false; }
        return { type: 'upload', file: f };
      }
      return { type: sel };
    }
  });
  if (!choice) return;

  //  Step 2: Load bytes
  let tmplBytes;
  try {
    if (choice.type === 'upload') {
      tmplBytes = new Uint8Array(await choice.file.arrayBuffer());
    } else if (choice.type === 'cached' && cached) {
      tmplBytes = cached;
    } else {
      const r = await fetch('/static/assets/template.xlsx');
      if (!r.ok) throw new Error('Không tải được /static/assets/template.xlsx');
      tmplBytes = new Uint8Array(await r.arrayBuffer());
    }
  } catch (e) {
    Swal.fire({ icon: 'error', title: 'Lỗi tải template', text: e.message }); return;
  }

  //  Step 3: Parse + validate with JSZip-based style extraction
  let tmplWb, styleTable;
  try {
    // Use JSZip to extract styles manually (more reliable than xlsx-js-style)
    const result = await _bpReadXlsxWithStyles(tmplBytes);
    tmplWb = result.wb;
    styleTable = result.styleTable;
    
    if (!tmplWb) {
      throw new Error('Không đọc được template');
    }
  }
  catch (e) { Swal.fire({ icon: 'error', title: 'Không đọc được template', text: e.message }); return; }

  if (!tmplWb || tmplWb.SheetNames.length < 2) {
    const found = tmplWb ? tmplWb.SheetNames.join(', ') : 'không xác định';
    Swal.fire({ icon: 'error', title: 'Template không hợp lệ',
      text: `File cần có ít nhất 2 sheet (Sheet 1 = HEADER, Sheet 2 = FOOTER). Tìm thấy: [${found}]` });
    return;
  }

  // Debug: Check style resolution
  try {
    console.log('[BPlan] === TEMPLATE STYLE DEBUG (JSZip-based) ===');
    const _checkSheet = tmplWb.Sheets[tmplWb.SheetNames[0]];
    const _checkKeys = Object.keys(_checkSheet).filter(k => !k.startsWith('!'));
    let _objectStyles = 0, _numberStyles = 0, _noStyles = 0;
    _checkKeys.forEach(k => {
      const cell = _checkSheet[k];
      if (!cell.s) _noStyles++;
      else if (typeof cell.s === 'object') _objectStyles++;
      else if (typeof cell.s === 'number') _numberStyles++;
    });
    console.log(`[BPlan] After JSZip extraction: ${_objectStyles} object, ${_numberStyles} number, ${_noStyles} none`);
    console.log(`[BPlan] Style table has ${styleTable.cellXfs.length} cell formats`);
    
    // Log sample cells
    _bpDebugLogStyles(tmplWb.Sheets[tmplWb.SheetNames[0]], 'HEADER sheet', 10);
    _bpDebugLogStyles(tmplWb.Sheets[tmplWb.SheetNames[1]], 'FOOTER sheet', 10);
    console.log('[BPlan] === END STYLE DEBUG ===');
  } catch (e) {
    console.warn('[BPlan] Style debug failed:', e);
  }

  const headerSheet = tmplWb.Sheets[tmplWb.SheetNames[0]];
  const bottomSheet = tmplWb.Sheets[tmplWb.SheetNames[1]];
  // Enforce fixed zone sizes: 4 HEADER rows, 5 FOOTER rows
  const _HEADER_ROWS = 4;
  const _FOOTER_ROWS = 5;
  let headerAoa = XLSX.utils.sheet_to_json(headerSheet, { header: 1, defval: '' });
  let bottomAoa = XLSX.utils.sheet_to_json(bottomSheet, { header: 1, defval: '' });
  while (headerAoa.length < _HEADER_ROWS) headerAoa.push([]);
  headerAoa = headerAoa.slice(0, _HEADER_ROWS);
  while (bottomAoa.length < _FOOTER_ROWS) bottomAoa.push([]);
  bottomAoa = bottomAoa.slice(0, _FOOTER_ROWS);

  //  Step 4: Build report
  const { ws, rows, zones } = _bpBuildSheet(
    currentResult, totalPeople, budgetVal, headerAoa, bottomAoa
  );
  _bpApplyTmplStyles(ws, headerSheet, bottomSheet, headerAoa, bottomAoa, zones, tmplWb, styleTable);
  
  // Debug: Log styles in the report sheet after applying template
  console.log('[BPlan] === REPORT SHEET STYLE DEBUG ===');
  _bpDebugLogStyles(ws, 'Report HEADER (rows 0-3)', 10);
  console.log(`[BPlan] Report FOOTER starts at row ${zones.bS}`);
  for (let i = 0; i < 5; i++) {
    const row = zones.bS + i;
    _bpDebugLogStyles(ws, `Report FOOTER row ${row}`, 6);
  }
  console.log('[BPlan] === END REPORT STYLE DEBUG ===');

  const reportWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(reportWb, ws, 'Báo cáo');

  //  Step 5: Open Univer preview
  const calcMenuConfig = {
    tableType: 'calc_menu',
    cacheKey: _BP_TMPL_KEY,
    hCount: 4, hCols: 6, fCount: 5, fCols: 4, fColStart: 2
  };
  await _bpShowPreview(reportWb, zones, tmplWb, tmplBytes, calcMenuConfig);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MULTI-TABLE TYPE EXPORT SUITE
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Auto-fit Column Width helper matching Excel's native AutoFit behavior.
 * Evaluates cell content lengths using Unicode display weights, applies a comfort padding (+3 chars),
 * and handles multi-column merges gracefully.
 */
function _bpAutoFitColsExact(ws, minW = 2, maxW = 80, zones = null) {
  if (!ws) return;
  console.log('[BPlan][AutoFitExact] _bpAutoFitColsExact called. minW:', minW, 'maxW:', maxW, 'zones:', JSON.stringify(zones));
  console.log('[BPlan][AutoFitExact] Total ws keys:', Object.keys(ws).filter(k => !k.startsWith('!')).length);
  const colWidths = {};
  
  const multiColMerges = new Set();
  if (ws['!merges'] && Array.isArray(ws['!merges'])) {
    ws['!merges'].forEach(m => {
      if (m.s.c !== m.e.c) {
        for (let r = m.s.r; r <= m.e.r; r++) {
          for (let c = m.s.c; c <= m.e.c; c++) {
            multiColMerges.add(`${r},${c}`);
          }
        }
      }
    });
  }
  console.log('[BPlan][AutoFitExact] multiColMerges count:', multiColMerges.size);

  let maxCol = 0;
  let skippedZone = 0, skippedMerge = 0, skippedEmpty = 0, measured = 0;
  const keys = Object.keys(ws).filter(k => !k.startsWith('!'));
  keys.forEach(k => {
    const cell = XLSX.utils.decode_cell(k);
    const c = cell.c;
    const r = cell.r;
    if (c > maxCol) maxCol = c;
    
    if (zones && zones.tS != null && zones.tE != null) {
      if (r < zones.tS || r > zones.tE) { skippedZone++; return; }
    }

    if (multiColMerges.has(`${r},${c}`)) { skippedMerge++; return; }

    const rawVal = ws[k] ? ws[k].v : null;
    if (rawVal == null || rawVal === '') { skippedEmpty++; return; }

    let str = String(rawVal);
    if (typeof rawVal === 'number') {
      str = rawVal.toLocaleString('vi');
      if (ws[k].z && ws[k].z.includes('VND')) str += ' VND';
    }
    const lines = str.split('\n');
    let maxLen = 0;
    lines.forEach(line => {
      const w = typeof _calcTextWidth === 'function' ? _calcTextWidth(line) : _bpTextWidth(line);
      if (w > maxLen) maxLen = w;
    });
    const autoWidth = Math.ceil(maxLen + 1);
    console.log(`[BPlan][AutoFitExact]   cell ${k} (r=${r},c=${c}) val=${JSON.stringify(str)} → w=${maxLen.toFixed(1)} → autoWidth=${autoWidth}`);
    measured++;
    if (!colWidths[c] || autoWidth > colWidths[c]) {
      colWidths[c] = autoWidth;
    }
  });
  console.log(`[BPlan][AutoFitExact] Scan done: measured=${measured}, skippedZone=${skippedZone}, skippedMerge=${skippedMerge}, skippedEmpty=${skippedEmpty}`);
  console.log('[BPlan][AutoFitExact] colWidths per col:', JSON.stringify(colWidths));

    // Evaluate multi-column merges to ensure spanned text is not truncated
    if (ws['!merges'] && Array.isArray(ws['!merges'])) {
      ws['!merges'].forEach(m => {
        const colSpan = m.e.c - m.s.c + 1;
        if (colSpan > 1) {
          const originAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
          const cellObj = ws[originAddr];
          const rawVal = cellObj ? cellObj.v : null;
          if (rawVal != null && rawVal !== '') {
            let str = String(rawVal);
            if (typeof rawVal === 'number') {
              str = rawVal.toLocaleString('vi');
            }
            let maxLen = 0;
            str.split('\n').forEach(line => {
              let w = typeof _calcTextWidth === 'function' ? _calcTextWidth(line) : _bpTextWidth(line);
              if (cellObj && cellObj.s && cellObj.s.font && cellObj.s.font.bold) {
                w *= 1.1; // 10% wider for bold header text
              }
              if (w > maxLen) maxLen = w;
            });
            const neededTotalW = Math.ceil(maxLen + 3.5);
            
            // Calculate current total width of all spanned columns
            let currentTotalW = 0;
            for (let c = m.s.c; c <= m.e.c; c++) {
              currentTotalW += (colWidths[c] || minW);
            }
            
            // If total width is insufficient, distribute missing deficit evenly
            if (neededTotalW > currentTotalW) {
              const deficit = neededTotalW - currentTotalW;
              const addPerCol = Math.ceil(deficit / colSpan);
              for (let c = m.s.c; c <= m.e.c; c++) {
                colWidths[c] = (colWidths[c] || minW) + addPerCol;
              }
            }
          }
        }
      });
    }

  const cols = [];
  for (let c = 0; c <= Math.max(maxCol, 0); c++) {
    const fitW = colWidths[c] != null ? Math.max(minW, Math.min(colWidths[c], maxW)) : minW;
    cols[c] = { wch: fitW };
  }
  console.log('[BPlan][AutoFitExact] Final !cols:', JSON.stringify(cols));
  ws['!cols'] = cols;
}

// Generate default template bytes matching dynamic region (hCols x hCount, fCols x fCount)
function _bpCreateDefaultTemplateBytes(hCount, hCols, fCount, fCols) {
  const wb = XLSX.utils.book_new();
  
  const hRows = [];
  for (let r = 0; r < hCount; r++) {
    const row = new Array(hCols).fill('');
    if (r === 0) row[0] = 'BÁO CÁO THỐNG KÊ BPLAN';
    hRows.push(row);
  }
  const hSheet = XLSX.utils.aoa_to_sheet(hRows);
  if (hCols >= 2) {
    hSheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: hCols - 1 } }];
  }
  XLSX.utils.book_append_sheet(wb, hSheet, 'HEADER');

  const fRows = [];
  for (let r = 0; r < fCount; r++) {
    const row = new Array(fCols).fill('');
    if (r === 0) row[0] = 'Xác nhận của nhà trường';
    fRows.push(row);
  }
  const fSheet = XLSX.utils.aoa_to_sheet(fRows);
  if (fCols >= 2) {
    fSheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: Math.min(1, fCols - 1) } }];
  }
  XLSX.utils.book_append_sheet(wb, fSheet, 'FOOTER');

  return new Uint8Array(_bpWriteXlsx(wb));
}

// Build a report worksheet from any HTMLTableElement with custom HEADER and FOOTER zones
function _bpBuildSheetFromTable(tableEl, hCount, hCols, fCount, fCols, fColStart, headerAoa, bottomAoa) {
  let clean = tableEl;
  if (tableEl && typeof _bpCloneForExport === 'function') {
    clean = _bpCloneForExport(tableEl);
  }
  
  let tableSheet;
  if (clean && clean.nodeType) {
    tableSheet = bplanTableToSheet(clean, { colPadding: 0 });
  } else {
    tableSheet = clean || {};
  }
  
  const ref = tableSheet['!ref'] || 'A1:A1';
  const range = XLSX.utils.decode_range(ref);
  const tableRowCount = range.e.r - range.s.r + 1;
  const tableColCount = range.e.c - range.s.c + 1;
  const NC = Math.max(hCols, tableColCount, fColStart + fCols);
  
  const rows = [];
  for (let i = 0; i < hCount; i++) {
    rows.push(new Array(NC).fill(''));
  }
  
  const tS = rows.length;
  for (let r = 0; r < tableRowCount; r++) {
    const row = new Array(NC).fill('');
    for (let c = 0; c < tableColCount; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
      const cell = tableSheet[cellAddr];
      row[c] = cell ? (cell.v != null ? cell.v : '') : '';
    }
    rows.push(row);
  }
  const tE = rows.length - 1;
  
  const bS = rows.length;
  for (let i = 0; i < fCount; i++) {
    rows.push(new Array(NC).fill(''));
  }
  const bE = rows.length - 1;
  
  const ws = XLSX.utils.aoa_to_sheet(rows);
  
  for (let r = 0; r < tableRowCount; r++) {
    const dstR = tS + r;
    for (let c = 0; c < tableColCount; c++) {
      const srcAddr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
      const dstAddr = XLSX.utils.encode_cell({ r: dstR, c: c });
      const srcCell = tableSheet[srcAddr];
      if (srcCell && ws[dstAddr]) {
        if (srcCell.s) ws[dstAddr].s = JSON.parse(JSON.stringify(srcCell.s));
        if (srcCell.t) ws[dstAddr].t = srcCell.t;
        if (srcCell.z) ws[dstAddr].z = srcCell.z;
      }
    }
  }
  
  const wsMerges = [];
  if (tableSheet['!merges'] && Array.isArray(tableSheet['!merges'])) {
    tableSheet['!merges'].forEach(m => {
      wsMerges.push({
        s: { r: m.s.r - range.s.r + tS, c: m.s.c - range.s.c },
        e: { r: m.e.r - range.s.r + tS, c: m.e.c - range.s.c }
      });
    });
  }
  ws['!merges'] = wsMerges;
  
  return { ws, rows, zones: { hCount, hCols, tS, tE, bS, bE, fCols, fColStart, NC } };
}

// Apply template styles for any custom table type and region
function _bpApplyTmplStylesCustom(ws, headerSheet, bottomSheet, headerAoa, bottomAoa, zones, tmplWb, styleTable, hCount, hCols, fCount, fCols, fColStart) {
  _bpCopySheetRegion(headerSheet, ws,
    hCount, hCols,
    0, 0,
    0, 0,
    zones.NC - 1, tmplWb, styleTable);

  _bpCopySheetRegion(bottomSheet, ws,
    fCount, fCols,
    0, 0,
    zones.bS, fColStart,
    zones.NC - 1, tmplWb, styleTable);

  const existingTableMerges = ws['!merges'] || [];
  ws['!merges'] = [...existingTableMerges];

  const mergeKey = (sr, sc, er, ec) => `${sr},${sc},${er},${ec}`;
  const seen = new Set();
  existingTableMerges.forEach(m => seen.add(mergeKey(m.s.r, m.s.c, m.e.r, m.e.c)));

  const addMerge = (sr, sc, er, ec) => {
    sc = Math.max(0, Math.min(sc, zones.NC - 1));
    ec = Math.max(0, Math.min(ec, zones.NC - 1));
    if (sc > ec) return;
    const k = mergeKey(sr, sc, er, ec);
    if (seen.has(k)) return;
    seen.add(k);
    ws['!merges'].push({ s: { r: sr, c: sc }, e: { r: er, c: ec } });
  };

  (headerSheet['!merges'] || []).forEach(m => {
    if (m.s.r < 0 || m.s.r >= hCount) return;
    const clampedEndRow = Math.min(m.e.r, hCount - 1);
    addMerge(m.s.r, m.s.c, clampedEndRow, m.e.c);
  });

  (bottomSheet['!merges'] || []).forEach(m => {
    const reportStartRow = m.s.r + zones.bS;
    const reportEndRow = m.e.r + zones.bS;
    if (reportStartRow < zones.bS || reportStartRow > zones.bE) return;
    const clampedEndRow = Math.min(reportEndRow, zones.bE);
    addMerge(reportStartRow, m.s.c + fColStart, clampedEndRow, m.e.c + fColStart);
  });
}

/**
 * Universal Export Function for custom table types & regions.
 */
async function bplanExportReportCustom(config) {
  const cacheKey = config.cacheKey || `bplan_report_template_${config.tableType}`;
  const hCount = config.hCount || 4;
  const hCols = config.hCols || 6;
  const fCount = config.fCount || 5;
  const fCols = config.fCols || hCols;
  const fColStart = config.fColStart != null ? config.fColStart : 0;
  const cached = _bpLoadTmplCache(cacheKey);

  const { value: choice } = await Swal.fire({
    title: `Chọn template - ${config.title || 'Báo cáo'}`,
    html: `
      <div style="display:flex;flex-direction:column;gap:10px;text-align:left;font-size:.93rem">
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="cached" checked>
          ${cached ? 'Dùng template đã lưu (cache)' : 'Dùng template mặc định'}
        </label>
        ${cached ? `
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="default">
          Dùng template mặc định
        </label>` : ''}
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="radio" name="_bpTmpl" value="upload">
          Tải lên file template (.xlsx)
        </label>
        <div id="_bpUpArea" style="display:none;margin-top:2px">
          <input type="file" id="_bpTmplFile" accept=".xlsx,.xls" style="font-size:.88em;width:100%">
        </div>
      </div>`,
    didOpen: () => {
      document.querySelectorAll('input[name="_bpTmpl"]').forEach(r => {
        r.addEventListener('change', () => {
          document.getElementById('_bpUpArea').style.display =
            r.value === 'upload' && r.checked ? 'block' : 'none';
        });
      });
    },
    confirmButtonText: 'Tiếp tục',
    showCancelButton: true,
    preConfirm: () => {
      const sel = document.querySelector('input[name="_bpTmpl"]:checked').value;
      if (sel === 'upload') {
        const f = document.getElementById('_bpTmplFile').files[0];
        if (!f) { Swal.showValidationMessage('Vui lòng chọn file template'); return false; }
        return { type: 'upload', file: f };
      }
      return { type: sel };
    }
  });
  if (!choice) return;

  let tmplBytes;
  try {
    if (choice.type === 'upload') {
      tmplBytes = new Uint8Array(await choice.file.arrayBuffer());
    } else if (choice.type === 'cached' && cached) {
      tmplBytes = cached;
    } else {
      tmplBytes = _bpCreateDefaultTemplateBytes(hCount, hCols, fCount, fCols);
    }
  } catch (e) {
    Swal.fire({ icon: 'error', title: 'Lỗi tải template', text: e.message }); return;
  }

  let tmplWb, styleTable;
  try {
    const res = await _bpReadXlsxWithStyles(tmplBytes);
    tmplWb = res.wb; styleTable = res.styleTable;
  } catch (e) {
    Swal.fire({ icon: 'error', title: 'Không đọc được template', text: e.message }); return;
  }

  if (!tmplWb || tmplWb.SheetNames.length < 2) {
    Swal.fire({ icon: 'error', title: 'Template không hợp lệ', text: 'File cần có ít nhất 2 sheet (Sheet 1 = HEADER, Sheet 2 = FOOTER).' });
    return;
  }

  const headerSheet = tmplWb.Sheets[tmplWb.SheetNames[0]];
  const bottomSheet = tmplWb.Sheets[tmplWb.SheetNames[1]];

  let headerAoa = XLSX.utils.sheet_to_json(headerSheet, { header: 1, defval: '' });
  let bottomAoa = XLSX.utils.sheet_to_json(bottomSheet, { header: 1, defval: '' });
  while (headerAoa.length < hCount) headerAoa.push([]);
  headerAoa = headerAoa.slice(0, hCount);
  while (bottomAoa.length < fCount) bottomAoa.push([]);
  bottomAoa = bottomAoa.slice(0, fCount);

  const { ws, zones } = _bpBuildSheetFromTable(config.tableEl, hCount, hCols, fCount, fCols, fColStart, headerAoa, bottomAoa);
  _bpApplyTmplStylesCustom(ws, headerSheet, bottomSheet, headerAoa, bottomAoa, zones, tmplWb, styleTable, hCount, hCols, fCount, fCols, fColStart);

  console.log('[BPlan][AutoFitExact] ws !cols BEFORE delete (bplanExportReportCustom):', JSON.stringify(ws['!cols']));
  delete ws['!cols'];
  console.log('[BPlan][AutoFitExact] zones passed to _bpAutoFitColsExact (no zone row-restriction for custom types):', JSON.stringify(zones));
  _bpAutoFitColsExact(ws, 2, 80, null); // null = scan ALL rows; multi-col merges still skipped
  console.log('[BPlan][AutoFitExact] ws !cols AFTER _bpAutoFitColsExact:', JSON.stringify(ws['!cols']));

  const reportWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(reportWb, ws, 'Báo cáo');

  await _bpShowPreview(reportWb, zones, tmplWb, tmplBytes, config);
}
