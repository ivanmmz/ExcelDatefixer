import ExcelJS from 'exceljs';

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
const MS_PER_DAY = 86400000;

// ── CSV Parsing ───────────────────────────────────────────────────────────────

/**
 * Standard RFC 4180 CSV parser that handles double quotes and commas.
 */
export function parseCSV(text) {
  const lines = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    const next = text[i + 1];
    
    if (inQuotes) {
      if (c === '"') {
        if (next === '"') {
          cell += '"';
          i++; // Skip next quote
        } else {
          inQuotes = false;
        }
      } else {
        cell += c;
      }
    } else {
      if (c === '"') {
        inQuotes = true;
      } else if (c === ',') {
        row.push(cell);
        cell = '';
      } else if (c === '\n' || c === '\r') {
        row.push(cell);
        cell = '';
        if (row.length > 0 || c === '\n') {
          lines.push(row);
          row = [];
        }
        if (c === '\r' && next === '\n') {
          i++; // Skip \n
        }
      } else {
        cell += c;
      }
    }
  }
  if (cell || row.length > 0) {
    row.push(cell);
    lines.push(row);
  }
  return lines;
}

// ── Date Parsing ──────────────────────────────────────────────────────────────

/**
 * Parse any date-like value: Date object, Excel serial number, formula result, or date string.
 * Returns a Date or null.
 */
export function tryParseAnyDate(val) {
  if (val === null || val === undefined || val === '') return null;

  // Handle exceljs formula cell value format
  if (val && typeof val === 'object' && val.result !== undefined) {
    return tryParseAnyDate(val.result);
  }

  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  // If it's a number (Excel serial number)
  if (typeof val === 'number' && isFinite(val)) {
    const d = new Date(EXCEL_EPOCH_UTC + val * MS_PER_DAY);
    return isNaN(d.getTime()) ? null : d;
  }

  if (typeof val === 'string') {
    let s = val.trim();
    if (!s) return null;

    // Check if it's a time-only string, e.g. "14:02:44" or "14:02" or "02:00 PM"
    const timeRegex = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(\s*(am|pm))?$/i;
    const match = s.match(timeRegex);
    if (match) {
      let hours = parseInt(match[1], 10);
      const minutes = parseInt(match[2], 10);
      const seconds = match[3] ? parseInt(match[3], 10) : 0;
      const ampm = match[5];
      if (ampm) {
        if (ampm.toLowerCase() === 'pm' && hours < 12) hours += 12;
        if (ampm.toLowerCase() === 'am' && hours === 12) hours = 0;
      }
      const now = new Date();
      const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), hours, minutes, seconds, 0));
      if (!isNaN(d.getTime())) return d;
    }

    // Replace dots with slashes in the date part only (leaving decimals in milliseconds intact)
    const spaceIdx = s.indexOf(' ');
    let datePart = spaceIdx === -1 ? s : s.substring(0, spaceIdx);
    let timePart = spaceIdx === -1 ? '' : s.substring(spaceIdx);
    if (datePart.includes('.')) {
      const parts = datePart.split('.');
      if (parts.length >= 3) {
        datePart = parts.join('/');
      }
    }
    s = datePart + timePart;

    // Try parsing split parts
    const sep = s.includes('/') ? '/' : (s.includes('-') ? '-' : null);
    if (sep) {
      const parts = s.split(sep);
      if (parts.length >= 3) {
        const p0 = parts[0].trim();
        const p1 = parts[1].trim();
        const p2 = parts[2].trim().split(/\s/)[0]; // strip time if any

        const n0 = parseInt(p0, 10);
        const n1 = parseInt(p1, 10);
        const n2 = parseInt(p2, 10);

        if (!isNaN(n0) && !isNaN(n1) && !isNaN(n2)) {
          let hours = 0, minutes = 0, seconds = 0;
          const timePartStr = s.split(/\s+/)[1];
          if (timePartStr) {
            const tParts = timePartStr.split(':');
            if (tParts.length >= 2) {
              hours = parseInt(tParts[0], 10) || 0;
              minutes = parseInt(tParts[1], 10) || 0;
              seconds = parseInt(tParts[2], 10) || 0;
            }
          }

          if (p0.length === 4) {
            // YYYY-MM-DD
            const d = new Date(Date.UTC(n0, n1 - 1, n2, hours, minutes, seconds));
            if (!isNaN(d.getTime())) return d;
          } else {
            // Determine day and month based on value bounds (smart dayfirst=True)
            let day, month;
            if (n0 > 12 && n1 <= 12) {
              day = n0;
              month = n1;
            } else if (n1 > 12 && n0 <= 12) {
              day = n1;
              month = n0;
            } else {
              // Ambiguous, default to dayfirst (n0 is day, n1 is month)
              day = n0;
              month = n1;
            }
            const d = new Date(Date.UTC(n2, month - 1, day, hours, minutes, seconds));
            if (!isNaN(d.getTime())) return d;
          }
        }
      }
    }

    // Fallback to standard JS Date parsing
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      if (/Z|[+-]\d{2}:?\d{2}$/i.test(s)) {
        return d;
      }
      return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()));
    }
  }

  return null;
}

/**
 * Returns 'datetime', 'date', or null based on the Date object's content.
 */
function detectContentType(d) {
  if (!d) return null;
  const hasTime = d.getUTCHours() !== 0 || d.getUTCMinutes() !== 0 || d.getUTCSeconds() !== 0;
  return hasTime ? 'datetime' : 'date';
}

// ── Task 1: Swap Date ─────────────────────────────────────────────────────────

/**
 * Swap DD and MM in string date values for a given column.
 */
export function taskSwapDate(rows, colIdx) {
  for (let r = 0; r < rows.length; r++) {
    const val = rows[r][colIdx];
    let s = '';
    let isDateObj = false;
    let originalDateObj = null;

    if (val instanceof Date) {
      isDateObj = true;
      originalDateObj = val;
      const y = val.getUTCFullYear();
      const m = val.getUTCMonth() + 1;
      const d = val.getUTCDate();
      s = `${d}/${m}/${y}`;
    } else if (val && typeof val === 'object' && val.result !== undefined) {
      s = String(val.result);
    } else if (val !== null && val !== undefined) {
      s = String(val).trim();
    }

    if (!s) continue;

    const sep = s.includes('/') ? '/' : (s.includes('.') ? '.' : (s.includes('-') ? '-' : null));
    if (!sep) continue;

    const parts = s.split(sep);
    if (parts.length < 2) continue;

    let p0 = parts[0].trim().split(/\s/)[0];
    let p1 = parts[1].trim().split(/\s/)[0];

    // Handle YYYY-MM-DD where year is first (p0 has length 4)
    if (p0.length === 4 && parts.length >= 3) {
      const year = p0;
      const month = parts[1].trim();
      const day = parts[2].trim().split(/\s/)[0];
      if (/^\d+$/.test(month) && /^\d+$/.test(day)) {
        const nMonth = parseInt(month, 10);
        const nDay = parseInt(day, 10);
        if (nMonth <= 12 && nDay <= 12) {
          let swapped = `${day}/${month}/${year}`;
          const timePartStr = s.split(/\s+/)[1];
          if (timePartStr) swapped += ' ' + timePartStr;

          if (isDateObj && originalDateObj) {
            rows[r][colIdx] = new Date(Date.UTC(parseInt(year, 10), nDay - 1, nMonth, originalDateObj.getUTCHours(), originalDateObj.getUTCMinutes(), originalDateObj.getUTCSeconds(), originalDateObj.getUTCMilliseconds()));
          } else {
            rows[r][colIdx] = swapped;
          }
        }
      }
    } else {
      if (/^\d+$/.test(p0) && /^\d+$/.test(p1)) {
        const n0 = parseInt(p0, 10);
        const n1 = parseInt(p1, 10);
        if (n0 <= 12 && n1 <= 12) {
          let swapped = `${p1}/${p0}`;
          if (parts.length > 2) swapped += '/' + parts.slice(2).join('/');

          if (isDateObj && originalDateObj) {
            const yearStr = parts[2] ? parts[2].trim().split(/\s/)[0] : String(originalDateObj.getUTCFullYear());
            const year = parseInt(yearStr, 10);
            rows[r][colIdx] = new Date(Date.UTC(year, n0 - 1, n1, originalDateObj.getUTCHours(), originalDateObj.getUTCMinutes(), originalDateObj.getUTCSeconds(), originalDateObj.getUTCMilliseconds()));
          } else {
            rows[r][colIdx] = swapped;
          }
        }
      }
    }
  }
}

// ── Task 2: Fix Missing ───────────────────────────────────────────────────────

/**
 * Forward-fill gaps of null/empty values up to `threshold` rows long.
 * Skips columns listed in `protectedCols`.
 */
export function taskFixMissing(rows, threshold, protectedCols) {
  if (!rows.length) return;
  const numCols = rows[0].length;

  for (let c = 0; c < numCols; c++) {
    if (protectedCols.has(c)) continue;

    let i = 0;
    while (i < rows.length) {
      const isEmpty = v => v === null || v === undefined || v === '';
      if (isEmpty(rows[i][c])) {
        const start = i;
        while (i < rows.length && isEmpty(rows[i][c])) i++;
        const gapLen = i - start;
        if (gapLen <= threshold && start > 0) {
          const fillVal = rows[start - 1][c];
          for (let j = start; j < start + gapLen; j++) rows[j][c] = fillVal;
        }
      } else {
        i++;
      }
    }
  }
}

// ── Task 3: Standardize 1440 ──────────────────────────────────────────────────

/**
 * Normalize each calendar day to exactly 1440 rows (one per minute).
 * Groups by date from `dateCol`, deduplicates by minute, pads if short.
 */
export function taskStandardize1440(rows, dateCol, timeCol) {
  const groups = new Map(); // "YYYY-MM-DD" -> [{ row, timeStr }]

  for (const row of rows) {
    const dDt = tryParseAnyDate(row[dateCol]);
    const tDt = tryParseAnyDate(row[timeCol]);
    if (!dDt || !tDt) continue;

    const dateStr = `${dDt.getUTCFullYear()}-${String(dDt.getUTCMonth() + 1).padStart(2, '0')}-${String(dDt.getUTCDate()).padStart(2, '0')}`;
    const timeStr = `${String(tDt.getUTCHours()).padStart(2, '0')}:${String(tDt.getUTCMinutes()).padStart(2, '0')}`;

    if (!groups.has(dateStr)) groups.set(dateStr, []);
    groups.get(dateStr).push({ row, timeStr });
  }

  const result = [];
  for (const [, group] of groups) {
    // Deduplicate by minute, keep first occurrence
    const seen = new Set();
    const clean = group.filter(item => {
      if (seen.has(item.timeStr)) return false;
      seen.add(item.timeStr);
      return true;
    });

    if (clean.length === 0) continue;

    if (clean.length >= 1440) {
      result.push(...clean.slice(0, 1440));
    } else {
      result.push(...clean);
      const last = clean[clean.length - 1];
      const pad = 1440 - clean.length;
      for (let i = 0; i < pad; i++) result.push({ ...last, row: [...last.row] });
    }
  }

  return result.map(item => item.row);
}

// ── Task 4: Date to Value ─────────────────────────────────────────────────────

/**
 * Convert Date/time values to Excel serial numbers.
 * role 'Date' → days since 1899-12-30; role 'Time' → fraction of day.
 */
export function taskDateToValue(rows, targets) {
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      let val = rows[r][c];
      if (val === null || val === undefined) continue;

      // 1. If it's a formula, extract the result
      if (typeof val === 'object' && val.formula !== undefined) {
        val = val.result;
        if (val === null || val === undefined) {
          rows[r][c] = val;
          continue;
        }
      }

      // 2. If it's a Date object, convert to Excel serial number (decimal date + time)
      if (val instanceof Date) {
        if (!isNaN(val.getTime())) {
          rows[r][c] = (val.getTime() - EXCEL_EPOCH_UTC) / MS_PER_DAY;
        }
        continue;
      }

      // 3. If it's a string, try converting to number or date
      if (typeof val === 'string') {
        const trimmed = val.trim();
        if (trimmed === '') {
          rows[r][c] = null;
          continue;
        }

        // Try parsing as number first
        const num = Number(trimmed);
        if (!isNaN(num)) {
          rows[r][c] = num;
          continue;
        }

        // Try parsing as date
        const dt = tryParseAnyDate(trimmed);
        if (dt) {
          rows[r][c] = (dt.getTime() - EXCEL_EPOCH_UTC) / MS_PER_DAY;
          continue;
        }

        // Keep as trimmed string
        rows[r][c] = trimmed;
        continue;
      }

      // Keep other types as is
      rows[r][c] = val;
    }
  }
}


// ── Task 5: Smart Format ──────────────────────────────────────────────────────

/**
 * Apply Excel number formatting to date/time cells in the workbook.
 * Returns count of styled cells.
 */
export function taskSmartFormat(workbook, dateFormat, timeFormat) {
  const dtFmt = `${dateFormat} ${timeFormat}`;
  let count = 0;

  const sheets = workbook.worksheets;
  for (const sheet of sheets) {
    if (sheet.rowCount < 2) continue;

    // Dynamically detect Date and Time columns by reading row 1 headers
    let dateColIdx = 0;
    let timeColIdx = 1;
    const headerRow = sheet.getRow(1);
    if (headerRow) {
      headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const val = String(cell.value || '').toLowerCase().trim();
        if (val.includes('date') || val.includes('日期')) {
          dateColIdx = colNumber - 1;
        } else if (val.includes('time') || val.includes('时间')) {
          timeColIdx = colNumber - 1;
        }
      });
    }

    const sheetTargets = [[dateColIdx, 'Date'], [timeColIdx, 'Time']];

    for (let row = 2; row <= sheet.rowCount; row++) {
      for (const [colIdx, role] of sheetTargets) {
        if (colIdx === null || colIdx === undefined) continue;
        const cell = sheet.getRow(row).getCell(colIdx + 1);
        const dt = tryParseAnyDate(cell.value);
        if (!dt) continue;
        if (role === 'Date') {
          cell.value = dt;
          cell.numFmt = detectContentType(dt) === 'datetime' ? dtFmt : dateFormat;
        } else {
          cell.value = { formula: `TIME(${dt.getUTCHours()},${dt.getUTCMinutes()},${dt.getUTCSeconds()})` };
          cell.numFmt = timeFormat;
        }
        count++;
      }
    }
  }

  return count;
}

// ── Pipeline ──────────────────────────────────────────────────────────────────

/**
 * Process a single Excel/CSV file through the enabled tasks.
 * @param {File} file
 * @param {{ dateFormat, timeFormat, threshold, activeTasks: Set<string> }} params
 * @param {(msg: string) => void} logFunc
 * @param {(progress: number) => void} progressFunc  0–1
 * @returns {Promise<Blob>} Processed .xlsx blob
 */
export async function processFile(file, params, logFunc, progressFunc) {
  const { dateFormat, timeFormat, threshold, activeTasks } = params;

  logFunc(`[*] Processing: ${file.name}`);

  const workbook = new ExcelJS.Workbook();
  let sheet;

  if (file.name.toLowerCase().endsWith('.csv')) {
    const text = await file.text();
    const csvRows = parseCSV(text);
    sheet = workbook.addWorksheet('Sheet1');
    for (const r of csvRows) {
      sheet.addRow(r);
    }
  } else {
    await workbook.xlsx.load(await file.arrayBuffer());
  }

  const sheets = workbook.worksheets;
  const sheetCount = sheets.length;
  let currentSheet = 0;

  for (const sheet of sheets) {
    currentSheet++;
    const maxRow = sheet.rowCount;
    const maxCol = sheet.columnCount;
    logFunc(`  - Sheet: '${sheet.name}' (${maxRow - 1} data rows)`);
    progressFunc((currentSheet / sheetCount) * 0.3);

    // Extract data rows (skip header row 1)
    const rows = [];
    for (let r = 2; r <= maxRow; r++) {
      const row = [];
      for (let c = 1; c <= maxCol; c++) {
        row.push(sheet.getRow(r).getCell(c).value);
      }
      rows.push(row);
    }
    if (!rows.length) continue;

    // Dynamically detect Date and Time columns by reading row 1 headers
    let dateColIdx = 0;
    let timeColIdx = 1;
    const headerRow = sheet.getRow(1);
    if (headerRow) {
      headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const val = String(cell.value || '').toLowerCase().trim();
        if (val.includes('date') || val.includes('日期')) {
          dateColIdx = colNumber - 1;
        } else if (val.includes('time') || val.includes('时间')) {
          timeColIdx = colNumber - 1;
        }
      });
    }

    const sheetTargets = [[dateColIdx, 'Date'], [timeColIdx, 'Time']];
    const sheetProtectedCols = new Set([dateColIdx, timeColIdx]);

    if (activeTasks.has('date_format')) {
      taskSwapDate(rows, dateColIdx);
      logFunc(`    + Swapped Day/Month in Col ${String.fromCharCode(65 + dateColIdx)}.`);
    }

    if (activeTasks.has('fix_missing')) {
      taskFixMissing(rows, threshold, sheetProtectedCols);
      logFunc(`    + Filled numeric gaps (threshold: ${threshold}).`);
    }

    if (activeTasks.has('standard_1440')) {
      const normalized = taskStandardize1440(rows, dateColIdx, timeColIdx);
      rows.length = 0;
      rows.push(...normalized);
      logFunc(`    + Standardized to 1440 rows (result: ${rows.length}).`);
    }

    if (activeTasks.has('date_to_value')) {
      taskDateToValue(rows, sheetTargets);
      logFunc('    + Converted Date/Time to Excel serial values.');
    }

    // Write processed rows back to sheet
    const rowCount = sheet.rowCount;
    for (let i = rowCount; i > 1; i--) {
      sheet.spliceRows(i, 1);
    }
    for (const row of rows) {
      const r = sheet.addRow(row);
      // Convert Date objects in the time column to TIME() formula if not converting to float value
      if (!activeTasks.has('date_to_value') && r.getCell(timeColIdx + 1).value instanceof Date) {
        const d = r.getCell(timeColIdx + 1).value;
        r.getCell(timeColIdx + 1).value = { formula: `TIME(${d.getHours()},${d.getMinutes()},${d.getSeconds()})` };
      }
    }
  }

  if (activeTasks.has('date_format')) {
    logFunc('  - Applying smart visual formatting...');
    const styled = taskSmartFormat(workbook, dateFormat, timeFormat);
    logFunc(`    + Styled ${styled} cells.`);
  }

  progressFunc(0.9);
  const outBuffer = await workbook.xlsx.writeBuffer();
  progressFunc(1);

  logFunc(`[✓] Done: '${file.name}'`);
  return new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
