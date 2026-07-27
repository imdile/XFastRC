/**
 * BPlan Enhanced Excel Export Utility
 * Replaces TableToExcel with SheetJS (xlsx.full.min.js) for rich formatting.
 * Requires xlsx.full.min.js to be loaded before this script.
 */

// ─── Shared border / width / alignment helpers (used by bplan_report.js too) ─

/** Returns a border side object. */
function _calcBorderMed(rgb)  { return { style: 'medium', color: { rgb } }; }
function _calcBorderThin(rgb) { return { style: 'thin',   color: { rgb } }; }
/** Returns all-four-sides border with the given style and colour. */
function _calcBorderFull(style, rgb) {
  const s = { style, color: { rgb } };
  return { top: s, bottom: s, left: s, right: s };
}

/**
 * Estimate Excel column width (in wch units, Calibri 11pt) for a given string.
 *
 * Excel's wch unit ≈ width of one digit in the default font.
 * Rules used (empirically calibrated against Excel AutoFit):
 *  - Combining diacritics (Vietnamese tones like ̀ ́ ̃ ̉ ̣ and decomposed accents) → 0
 *  - Narrow ASCII chars (i j l f r t ! | 1 ) → 0.5
 *  - Normal ASCII letters/digits/spaces       → 1.0
 *  - Wide Latin (uppercase M/W)               → 1.2
 *  - CJK / Hangul / full-width                → 2.0
 */
function _calcTextWidth(str) {
  let w = 0;
  for (const ch of String(str || '')) {
    const cp = ch.codePointAt(0);
    // Combining diacritics (Vietnamese tone marks, decomposed accents) — zero width
    if (cp >= 0x0300 && cp <= 0x036F) { continue; }
    if (cp >= 0x1DC0 && cp <= 0x1DFF) { continue; }
    if (cp >= 0xFE20 && cp <= 0xFE2F) { continue; }
    // CJK / Hangul / full-width forms
    if (
      (cp >= 0x1100  && cp <= 0x115F)  ||
      (cp >= 0x2E80  && cp <= 0x303E)  ||
      (cp >= 0x3040  && cp <= 0x33FF)  ||
      (cp >= 0x3400  && cp <= 0x4DBF)  ||
      (cp >= 0x4E00  && cp <= 0x9FFF)  ||
      (cp >= 0xAC00  && cp <= 0xD7AF)  ||
      (cp >= 0xF900  && cp <= 0xFAFF)  ||
      (cp >= 0xFF00  && cp <= 0xFF60)  ||
      (cp >= 0xFFE0  && cp <= 0xFFE6)  ||
      (cp >= 0x20000 && cp <= 0x2A6DF)
    ) { w += 2.0; continue; }
    // Narrow ASCII
    if ('ijlft!|1:.,;()[]'.includes(ch)) { w += 0.6; continue; }
    // Wide uppercase
    if ('MW'.includes(ch)) { w += 1.2; continue; }
    // Default: normal character
    w += 1.0;
  }
  return w;
}

/**
 * Compute auto-fit column widths from a 2-D data array.
 * @param {Array[]} sheetData
 * @param {number} padding  extra chars added to each column
 * @param {number} min      minimum width
 * @param {number} max      maximum width
 * @returns {Array<{wch:number}>}
 */
function _calcAutoWidths(sheetData, padding, min, max) {
  padding = padding != null ? padding : 1;
  min = min != null ? min : 10;
  max = max != null ? max : 80;
  const ncols = Math.max(...sheetData.map(r => r.length));
  const widths = new Array(ncols).fill(min);
  sheetData.forEach(row => {
    row.forEach((cell, ci) => {
      let str = String(cell ?? '');
      if (typeof cell === 'number') {
        str = cell.toLocaleString('vi');
      }
      const w = _calcTextWidth(str) + padding;
      if (w > widths[ci]) widths[ci] = Math.min(w, max);
    });
  });
  return widths.map(w => ({ wch: Math.ceil(w) }));
}

/**
 * Infer horizontal alignment from a cell value.
 * Numbers/VND/% → right; short single-word codes → center; else left.
 */
function _calcAlign(val) {
  const t = String(val ?? '').trim();
  if (t === '') return 'left';
  if (/^-?[\d][\d\s.,]*(\s*(VND|%|đ))?$/.test(t)) return 'right';
  if (t.length <= 6 && !/\s/.test(t)) return 'center';
  return 'left';
}

// ─── Style constants ──────────────────────────────────────────────────────────

const _BP_HEADER = {
  font:      { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  fill:      { fgColor: { rgb: "2E5597" } },
  border: {
    top:    { style: "medium", color: { rgb: "1D3A6E" } },
    bottom: { style: "medium", color: { rgb: "1D3A6E" } },
    left:   { style: "medium", color: { rgb: "1D3A6E" } },
    right:  { style: "medium", color: { rgb: "1D3A6E" } }
  }
};

const _BP_FOOTER = {
  font:      { bold: true, sz: 11, color: { rgb: "1D3A6E" } },
  alignment: { horizontal: "center", vertical: "center" },
  fill:      { fgColor: { rgb: "D9E1F2" } },
  border: {
    top:    { style: "medium", color: { rgb: "2E5597" } },
    bottom: { style: "medium", color: { rgb: "2E5597" } },
    left:   { style: "medium", color: { rgb: "2E5597" } },
    right:  { style: "medium", color: { rgb: "2E5597" } }
  }
};

const _BP_ODD = {
  font:      { sz: 10 },
  alignment: { vertical: "center" },
  fill:      { fgColor: { rgb: "FFFFFF" } },
  border: {
    top:    { style: "thin", color: { rgb: "9DB2CE" } },
    bottom: { style: "thin", color: { rgb: "9DB2CE" } },
    left:   { style: "thin", color: { rgb: "9DB2CE" } },
    right:  { style: "thin", color: { rgb: "9DB2CE" } }
  }
};

const _BP_EVEN = {
  font:      { sz: 10 },
  alignment: { vertical: "center" },
  fill:      { fgColor: { rgb: "EEF2FA" } },
  border: {
    top:    { style: "thin", color: { rgb: "9DB2CE" } },
    bottom: { style: "thin", color: { rgb: "9DB2CE" } },
    left:   { style: "thin", color: { rgb: "9DB2CE" } },
    right:  { style: "thin", color: { rgb: "9DB2CE" } }
  }
};

// First-column data cells get bold + slightly stronger left border
const _BP_FIRST_COL_EXTRA = {
  font:   { bold: true, sz: 10 },
  border: { left: { style: "medium", color: { rgb: "2E5597" } } }
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

function _bpClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/** Alias for _calcTextWidth — same calibrated logic used by all auto-fit paths. */
function _bpTextWidth(str) {
  return _calcTextWidth(str);
}

/**
 * Infer horizontal alignment from cell text content.
 * Right-aligns numbers (with any thousand-separator style) and VND/% values.
 * Centers short codes (≤ 6 chars, no spaces, not a sentence).
 */
function _bpAlign(rawText) {
  const t = (rawText || "").trim();
  if (t === "") return "left";
  // Right-align: pure integer/decimal, optionally with VND or %
  // Covers: "1234", "1.234.567", "1,234,567", "1.234 VND", "-5", "98%"
  if (/^-?[\d][\d\s.,]*(\s*(VND|%|đ))?$/.test(t)) return "right";
  // Center-align: short single-word codes/labels (STT, OK, Có, Không, etc.)
  if (t.length <= 6 && !/\s/.test(t)) return "center";
  return "left";
}

// ─── Core conversion ──────────────────────────────────────────────────────────

/**
 * Convert a DOM <table> to a SheetJS worksheet with rich formatting.
 * Handles colspan / rowspan merges and thead / tbody / tfoot sections.
 *
 * @param {HTMLTableElement} tableEl
 * @param {Object}  [opts]
 * @param {number[]} [opts.colWidths]  – Override auto column widths (char units).
 * @param {number}   [opts.minColWidth=6]   – Minimum column width.
 * @param {number}   [opts.maxColWidth=60]  – Maximum column width.
 * @param {number}   [opts.colPadding=3]    – Extra padding added to auto width.
 * @returns {Object} SheetJS worksheet
 */
function bplanTableToSheet(tableEl, opts) {
  opts = Object.assign({ minColWidth: 6, maxColWidth: 80, colPadding: 0 }, opts || {});

  const rows    = Array.from(tableEl.rows);
  const numCols = Math.max(1, ...rows.map(r =>
    Array.from(r.cells).reduce((s, c) => s + (parseInt(c.colSpan) || 1), 0)
  ));

  // ── Build a 2-D grid respecting rowspan/colspan ──────────────────────────
  const grid   = [];   // grid[r][c] = cell descriptor | null
  const filled = [];   // filled[r][c] = true when occupied by a spanning ancestor

  rows.forEach((row, ri) => {
    if (!grid[ri])   grid[ri]   = new Array(numCols).fill(null);
    if (!filled[ri]) filled[ri] = new Array(numCols).fill(false);

    const section = (row.parentElement && row.parentElement.tagName.toLowerCase()) || "tbody";
    let ci = 0;

    Array.from(row.cells).forEach(cell => {
      while (ci < numCols && filled[ri][ci]) ci++;
      if (ci >= numCols) return;

      const colspan = Math.max(1, parseInt(cell.colSpan) || 1);
      const rowspan = Math.max(1, parseInt(cell.rowSpan) || 1);
      const rawText = (cell.innerText || cell.textContent || "").trim();

      // Numeric detection: strip thousand separators + VND
      let value = rawText;
      const stripped = rawText.replace(/[\s,.]/g, "").replace(/VND$/i, "").trim();
      if (stripped !== "" && !isNaN(stripped) && !/[a-df-wyzA-DF-WYZ]/.test(rawText)) {
        value = parseFloat(stripped);
      }

      grid[ri][ci] = { value, rawText, colspan, rowspan, cellEl: cell, section, colIndex: ci };

      // Mark spanned positions
      for (let dr = 0; dr < rowspan; dr++) {
        for (let dc = 0; dc < colspan; dc++) {
          if (!filled[ri + dr]) filled[ri + dr] = new Array(numCols).fill(false);
          filled[ri + dr][ci + dc] = true;
        }
      }
      ci += colspan;
    });
  });

  // ── Build worksheet ───────────────────────────────────────────────────────
  const ws     = {};
  const merges = [];

  // Track per-column max content width for auto-fit
  const colWidths = new Array(numCols).fill(opts.minColWidth);

  grid.forEach((row, ri) => {
    row.forEach((cell, ci) => {
      const addr = XLSX.utils.encode_cell({ r: ri, c: ci });

      // Spanned cells (null) must still get an empty entry so SheetJS doesn't
      // lose the column reference and drop data from column A.
      if (!cell) {
        ws[addr] = { v: '', t: 's' };
        return;
      }
      const isHeader = cell.section === "thead";
      const isFooter = cell.section === "tfoot";
      const isFirst  = ci === 0;
      const isEven   = ri % 2 === 0;

      // ── Value ──────────────────────────────────────────────────────────
      ws[addr] = {
        v: cell.value,
        t: typeof cell.value === "number" ? "n" : "s"
      };

      // ── Style ──────────────────────────────────────────────────────────
      let style;
      if (isHeader) {
        style = _bpClone(_BP_HEADER);
        if (cell.colspan > 1) style.alignment.wrapText = true;
      } else if (isFooter) {
        style = _bpClone(_BP_FOOTER);
        style.alignment.horizontal = _bpAlign(cell.rawText);
      } else {
        style = _bpClone(isEven ? _BP_EVEN : _BP_ODD);
        style.alignment.horizontal = _bpAlign(cell.rawText);
        // Bold + stronger left border for the first data column
        if (isFirst) {
          style.font.bold = true;
          style.border.left = _bpClone(_BP_FIRST_COL_EXTRA.border.left);
        }
      }

      // Right-column medium border on last column
      if (ci === numCols - 1 && !isHeader) {
        style.border.right = { style: "medium", color: { rgb: "2E5597" } };
      }
      // Bottom medium border on last data row before tfoot, and on tfoot rows
      if (isFooter) {
        style.border.bottom = { style: "medium", color: { rgb: "1D3A6E" } };
      }

      ws[addr].s = style;

      // ── Merge registration ─────────────────────────────────────────────
      if (cell.rowspan > 1 || cell.colspan > 1) {
        merges.push({
          s: { r: ri, c: ci },
          e: { r: ri + cell.rowspan - 1, c: ci + cell.colspan - 1 }
        });
      }

      // ── Auto-fit tracking ──────────────────────────────────────────────
      // Only count single-column cells (spanning cells skew the width calc)
      if (cell.colspan === 1) {
        const w = _bpTextWidth(cell.rawText) + opts.colPadding;
        if (w > colWidths[ci]) colWidths[ci] = w;
      }
    });
  });

  // ── Sheet metadata ────────────────────────────────────────────────────────
  const totalRows = grid.length;
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: totalRows - 1, c: numCols - 1 }
  });

  if (merges.length) ws["!merges"] = merges;

  // Row heights
  ws["!rows"] = grid.map(row => ({
    hpx: row.some(c => c && c.section === "thead") ? 24 : 19
  }));

  // Column widths – use override if supplied, else auto-fit clamped to max
  ws["!cols"] = (opts.colWidths && opts.colWidths.length)
    ? opts.colWidths.map(w => ({ wch: w }))
    : colWidths.map(w => ({ wch: Math.min(w, opts.maxColWidth) }));

  return ws;
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Prepare a clean clone of a table for export:
 * - Removes tablesorter filter rows (rows that contain only inputs/selects)
 * - Strips hidden rows (display:none)
 * - Removes the rwd-table responsive class so row sections are stable
 */
function _bpCloneForExport(tableEl) {
  const clone = tableEl.cloneNode(true);
  // Remove rows that are purely filter inputs (tablesorter injects these into thead)
  Array.from(clone.rows).forEach(row => {
    const cells = Array.from(row.cells);
    const allInputs = cells.length > 0 && cells.every(td =>
      td.querySelector('input,select') && (td.innerText || td.textContent || '').trim() === ''
    );
    if (allInputs) row.parentNode && row.parentNode.removeChild(row);
  });
  // Remove rows hidden by display:none (tablesorter filter hides non-matching rows)
  Array.from(clone.rows).forEach(row => {
    if (row.style.display === 'none') row.parentNode && row.parentNode.removeChild(row);
  });
  return clone;
}

/**
 * Export a single table to an Excel file.
 *
 * @param {HTMLTableElement} tableEl
 * @param {string}  fileName   e.g. "report.xlsx"
 * @param {string}  [sheetName="Sheet 1"]
 * @param {Object}  [opts]     passed to bplanTableToSheet
 */
function bplanExportTable(tableEl, fileName, sheetName, opts) {
  if (!tableEl) { console.warn("bplanExportTable: tableEl is null"); return; }
  const clean = _bpCloneForExport(tableEl);
  const wb = XLSX.utils.book_new();
  const ws = bplanTableToSheet(clean, opts || {});
  XLSX.utils.book_append_sheet(wb, ws, (sheetName || "Sheet 1").substring(0, 31));
  // Use write() + Blob for maximum compatibility across environments
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: false });
  const blob = new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}

/**
 * Export multiple tables into separate sheets of one workbook.
 *
 * @param {Array<{tableEl, sheetName, opts}>} sheets
 * @param {string} fileName
 */
function bplanExportMultiSheet(sheets, fileName) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ tableEl, sheetName, opts }) => {
    if (!tableEl) return;
    const clean = _bpCloneForExport(tableEl);
    const ws = bplanTableToSheet(clean, opts || {});
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || "Sheet").substring(0, 31));
  });
  const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: false });
  const blob = new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}
