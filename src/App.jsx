import { useState, useRef } from 'react';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';

// ── Column widths (character width units) ─────────────────────────────────────
const COL_WIDTHS = [12, 11.5, 19, 26, 23.5, 22.5, 27, 9.5, 5.5, 14, 5.5, 13.5];

// ── Expected CSV columns ───────────────────────────────────────────────────────
const EXPECTED_COLUMNS = [
  'Asset Name', 'Asset Type', 'Make Name', 'Model Number', 'Serial Number',
  'Asset Location', 'Property Zone Served', 'Subject', 'Note', 'Property Name',
  'Billing Customer Name', 'Property Address Line 1', 'Property Address Line 2',
  'Property City', 'Property State', 'Property Zipcode',
];

// ── CSV parsing ────────────────────────────────────────────────────────────────
function parseCSVLine(line) {
  const fields = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  while (i < line.length) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === ',') {
        fields.push(field);
        field = '';
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }
  fields.push(field);
  return fields;
}

function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
  if (lines.length === 0) return { headers: [], rows: [] };

  const headers = parseCSVLine(lines[0]);
  const rows = lines.slice(1).map(line => {
    const values = parseCSVLine(line);
    const row = {};
    for (let i = 0; i < headers.length; i++) {
      const val = i < values.length ? values[i] : '';
      row[headers[i]] = val === '' ? null : val;
    }
    return row;
  });

  return { headers, rows };
}

// ── XLSX parsing ───────────────────────────────────────────────────────────────
function parseXLSX(arrayBuffer) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (raw.length === 0) return { headers: [], rows: [] };

  const headers = raw[0].map(h => (h == null ? '' : String(h).trim()));
  const rows = raw.slice(1).map(rawRow => {
    const row = {};
    for (let i = 0; i < headers.length; i++) {
      const cell = i < rawRow.length ? rawRow[i] : '';
      if (cell === '' || cell == null) {
        row[headers[i]] = null;
      } else {
        // Convert to string; strip trailing .0 from numbers stored as floats
        const str = String(cell).replace(/\.0$/, '');
        row[headers[i]] = str === '' ? null : str;
      }
    }
    return row;
  });

  return { headers, rows };
}

// ── Data transformation helpers ───────────────────────────────────────────────
function classifySubject(subject) {
  if (!subject) return 'other';
  const s = subject.toLowerCase().trim();
  const hasFilter = s.includes('filter') || s.includes('filters');
  const hasBelt = s.includes('belt') || s.includes('belts');
  if (hasFilter && hasBelt) return 'both';
  if (hasFilter) return 'filter';
  if (hasBelt) return 'belt';
  return 'other';
}

function parseQtyAndValue(note) {
  if (!note) return { qty: null, value: null };
  const s = note.trim();
  let m = s.match(/^[([]?(\d{1,2})[)\]]\s+(.+)/);
  if (m) return { qty: m[1], value: m[2].trim() };
  m = s.match(/^(\d{1,2})\s+-\s+(.+)/);
  if (m) return { qty: m[1], value: m[2].trim() };
  return { qty: null, value: s };
}

function buildSiteAddress(row) {
  const zipRaw = row['Property Zipcode'];
  const zip = zipRaw != null ? String(zipRaw).replace(/\.0$/, '') : null;
  return [
    row['Property Address Line 1'],
    row['Property Address Line 2'],
    row['Property City'],
    row['Property State'],
    zip,
  ].filter(v => v != null && v !== '').join(', ');
}

function transformRows(rows) {
  // Step A: group by Property Name
  const grouped = new Map();
  for (const row of rows) {
    const key = row['Property Name'] ?? null;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  }

  // Sort non-null A-Z case-insensitive; null goes last
  const sorted = [...grouped.entries()].sort(([a], [b]) => {
    if (a === null && b === null) return 0;
    if (a === null) return 1;
    if (b === null) return -1;
    return a.localeCompare(b, undefined, { sensitivity: 'base' });
  });

  const properties = [];
  let totalAssets = 0;

  for (const [propName, propRows] of sorted) {
    const firstRow = propRows[0];

    // Step F: property header metadata
    const meta = {
      agreementNumber: '',
      customerName: firstRow['Billing Customer Name'] ?? '',
      propertyName: propName ?? '(No Property)',
      siteAddress: buildSiteAddress(firstRow),
    };

    // Steps B-E: deduplicate assets and map columns
    const assetMap = new Map();
    for (const row of propRows) {
      const assetName = row['Asset Name'];
      if (!assetMap.has(assetName)) {
        assetMap.set(assetName, {
          assetId:      assetName ?? '',
          assetType:    row['Asset Type'] ?? '',
          manufacturer: row['Make Name'] ?? '',
          modelNo:      row['Model Number'] ?? '',
          serialNo:     row['Serial Number'] ?? '',
          location:     row['Asset Location'] ?? '',
          areaServed:   row['Property Zone Served'] ?? '',
          tonnage:      '',
          filterQty:    '',
          filters:      '',
          beltQty:      '',
          belts:        '',
        });
      }

      const asset = assetMap.get(assetName);
      const type = classifySubject(row['Subject']);

      if (type === 'filter') {
        const { qty, value } = parseQtyAndValue(row['Note']);
        asset.filterQty = qty ?? '';
        asset.filters   = value ?? '';
      } else if (type === 'belt') {
        const { qty, value } = parseQtyAndValue(row['Note']);
        asset.beltQty = qty ?? '';
        asset.belts   = value ?? '';
      } else if (type === 'both') {
        asset.filters = row['Note'] ?? '';
        asset.belts   = 'See Filters';
      }
      // 'other': keep asset record, ignore note
    }

    const assets = [...assetMap.values()];
    totalAssets += assets.length;
    properties.push({ meta, assets });
  }

  return { properties, totalAssets };
}

// ── Excel export ──────────────────────────────────────────────────────────────
function exportToExcel(properties) {
  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];
  const rowHeights = [];
  let r = 0;

  // Style building blocks
  const THIN     = { style: 'thin', color: { rgb: '000000' } };
  const border   = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  const fLabel   = { name: 'Arial', sz: 10, bold: true };
  const fValue   = { name: 'Arial', sz: 10 };
  const fItalic  = { name: 'Arial', sz: 10, italic: true };
  const fColHdr  = { name: 'Arial', sz: 10, bold: true };
  const fData    = { name: 'Arial', sz: 10 };
  const fFooter  = { name: 'Arial', sz: 9, italic: true };
  const fillHdr  = { fgColor: { rgb: 'D9E1F2' }, patternType: 'solid' };
  const aLeft    = { horizontal: 'left',   vertical: 'center', wrapText: true };
  const aCenter  = { horizontal: 'center', vertical: 'center', wrapText: true };

  function setCell(row, col, value, style) {
    ws[XLSX.utils.encode_cell({ r: row, c: col })] = {
      v: value ?? '',
      t: 's',
      s: style,
    };
  }

  function addMerge(r1, c1, r2, c2) {
    merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
  }

  // Fill cells in a column range with an empty string (required for merged regions)
  function fillEmpty(row, c1, c2, style) {
    for (let c = c1; c <= c2; c++) {
      setCell(row, c, '', style);
    }
  }

  for (const { meta, assets } of properties) {
    // ── Row 1: Agreement Number ─────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, 'Agreement Number:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, meta.agreementNumber, { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 2: PM Schedule ──────────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, 'PM Schedule:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, '', { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 3: Customer Name ────────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, '     Customer Name:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, meta.customerName, { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 4: Property (A empty, B = label, C:L = value) ──────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, '', { font: fLabel });
    setCell(r, 1, 'Property:', { font: fLabel, alignment: aLeft });
    setCell(r, 2, meta.propertyName, { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 5: Site Address ─────────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, 'Site Address:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, meta.siteAddress, { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 6: POC ──────────────────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, 'POC:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, '', { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 7: Special Notes ────────────────────────────────────────────────
    rowHeights[r] = { hpt: 23.25 };
    setCell(r, 0, 'Special Notes:', { font: fLabel, alignment: aLeft });
    setCell(r, 1, '', { font: fLabel });
    addMerge(r, 0, r, 1);
    setCell(r, 2, '', { font: fValue, alignment: aLeft });
    fillEmpty(r, 3, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Row 8: italic watermark note ────────────────────────────────────────
    rowHeights[r] = { hpt: 40 };
    setCell(r, 0, '', { font: fValue });
    setCell(r, 1, '(water access, security, etc\u2026)', {
      font: fItalic,
      alignment: aLeft,
    });
    fillEmpty(r, 2, 11, { font: fValue });
    addMerge(r, 2, r, 11);
    r++;

    // ── Column header row ───────────────────────────────────────────────────
    rowHeights[r] = { hpt: 24.75 };
    const colLabels = [
      'Asset ID', 'Asset Type', 'Manufacturer', 'Model #', 'Serial #',
      'Location', 'Area Served', 'Tonnage', 'Qty', 'Filters', 'Qty', 'Belts',
    ];
    for (let c = 0; c < 12; c++) {
      setCell(r, c, colLabels[c], {
        font: fColHdr,
        fill: fillHdr,
        alignment: aCenter,
        border,
      });
    }
    r++;

    // ── Data rows ───────────────────────────────────────────────────────────
    for (const a of assets) {
      rowHeights[r] = { hpt: 18.75 };
      const vals = [
        a.assetId, a.assetType, a.manufacturer, a.modelNo,
        a.serialNo, a.location, a.areaServed, a.tonnage,
        a.filterQty, a.filters, a.beltQty, a.belts,
      ];
      for (let c = 0; c < 12; c++) {
        setCell(r, c, vals[c], { font: fData, alignment: aLeft, border });
      }
      r++;
    }

    // ── Footer note row ─────────────────────────────────────────────────────
    rowHeights[r] = { hpt: 26.25 };
    const footerText =
      '**Technician will acquire any omitted model and/or serial ' +
      'information during the first Preventative Maintenance visit.';
    setCell(r, 0, '', { font: fFooter });
    setCell(r, 1, footerText, { font: fFooter, alignment: aLeft });
    fillEmpty(r, 2, 6, { font: fFooter });
    addMerge(r, 1, r, 6); // B:G
    fillEmpty(r, 7, 11, { font: fFooter });
    r++;

    // ── Blank separator row ─────────────────────────────────────────────────
    r++;
  }

  ws['!ref']    = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(r - 1, 0), c: 11 } });
  ws['!merges'] = merges;
  ws['!rows']   = rowHeights;
  ws['!cols']   = COL_WIDTHS.map(w => ({ wch: w }));

  XLSX.utils.book_append_sheet(wb, ws, 'PMA Worksheet');

  const today    = new Date().toISOString().slice(0, 10);
  const filename = `PMA_Worksheet_${today}.xlsx`;
  const buf      = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  saveAs(new Blob([buf], { type: 'application/octet-stream' }), filename);
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [appData, setAppData]       = useState(null);
  const [fileName, setFileName]     = useState(null);
  const [uploadError, setUploadError] = useState(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef(null);

  function handleFile(file) {
    if (!file) return;
    setUploadError(null);
    const ext = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();

    function onParsed({ headers, rows }) {
      const missing = EXPECTED_COLUMNS.filter(col => !headers.includes(col));
      if (missing.length > 0) {
        setUploadError('columns');
        return;
      }
      const transformed = transformRows(rows);
      setFileName(file.name);
      setAppData(transformed);
      exportToExcel(transformed.properties);
    }

    if (ext === 'csv') {
      reader.onload = (e) => onParsed(parseCSV(e.target.result));
      reader.readAsText(file);
    } else if (ext === 'xlsx' || ext === 'xls') {
      reader.onload = (e) => onParsed(parseXLSX(e.target.result));
      reader.readAsArrayBuffer(file);
    } else {
      setUploadError('filetype');
    }
  }

  function handleReset() {
    setAppData(null);
    setFileName(null);
    setUploadError(null);
    // Reset the file input so the same file can be re-selected
    if (fileInputRef.current) fileInputRef.current.value = '';
  }

  function onDragOver(e) {
    e.preventDefault();
    if (!isDragOver) setIsDragOver(true);
  }

  function onDragLeave(e) {
    if (!e.currentTarget.contains(e.relatedTarget)) {
      setIsDragOver(false);
    }
  }

  function onDrop(e) {
    e.preventDefault();
    setIsDragOver(false);
    handleFile(e.dataTransfer.files[0]);
  }

  function onFileInputChange(e) {
    handleFile(e.target.files[0]);
  }

  function handleExport() {
    if (!appData) return;
    exportToExcel(appData.properties);
  }

  // Build full data list: all properties and assets
  const previewItems = [];
  if (appData) {
    for (const { meta, assets } of appData.properties) {
      previewItems.push({ kind: 'group', label: meta.propertyName });
      for (const asset of assets) {
        previewItems.push({ kind: 'asset', asset });
      }
    }
  }

  const today = new Date().toLocaleDateString('en-US', {
    year: 'numeric', month: 'short', day: 'numeric',
  });

  return (
    <>
      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        html, body, #root {
          width: 100%; height: 100%;
          background: #fff;
          font-family: Arial, sans-serif;
          color: #222;
          display: block;
          place-items: unset;
          text-align: left;
        }
        .pma-btn {
          padding: 6px 16px;
          font-size: 13px;
          font-family: Arial, sans-serif;
          font-weight: bold;
          border-radius: 3px;
          cursor: pointer;
          transition: opacity 0.15s;
        }
        .pma-btn:hover:not(:disabled) { opacity: 0.85; }
        .pma-btn:disabled { cursor: not-allowed; }
        .pma-reset-link {
          background: none;
          border: none;
          padding: 0;
          font-family: Arial, sans-serif;
          font-size: 12px;
          color: #2d5f96;
          cursor: pointer;
          text-decoration: underline;
        }
        .pma-reset-link:hover { color: #1B3A5C; }
      `}</style>

      <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', background: '#fff' }}>

        {/* ── Header bar ── */}
        <header style={{
          background: '#1B3A5C',
          color: '#fff',
          padding: '10px 18px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flexShrink: 0,
        }}>
          <span style={{ fontSize: 15, fontWeight: 'bold', letterSpacing: 0.2 }}>
            PMA Worksheet Export
          </span>
          <button
            className="pma-btn"
            onClick={handleExport}
            disabled={!appData}
            style={{
              background: appData ? '#ffffff' : '#3d5a73',
              color:      appData ? '#1B3A5C' : '#7aa0be',
              border:     appData ? '1px solid #ccc' : '1px solid #2e4d63',
            }}
          >
            Export Excel
          </button>
        </header>

        {/* ── Main content ── */}
        <main style={{ flex: 1, padding: 18 }}>

          {!appData ? (
            /* ── Upload state ── */
            <div style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              paddingTop: 48,
            }}>
              {/* Drop zone */}
              <div
                onDragOver={onDragOver}
                onDragLeave={onDragLeave}
                onDrop={onDrop}
                style={{
                  border: `2px dashed ${isDragOver ? '#1B3A5C' : '#b0bcd8'}`,
                  borderRadius: 6,
                  background: isDragOver ? '#edf1f8' : '#f8f9fb',
                  padding: '44px 64px',
                  textAlign: 'center',
                  maxWidth: 520,
                  width: '100%',
                  transition: 'border-color 0.15s, background 0.15s',
                }}
              >
                <div style={{ fontSize: 15, fontWeight: 'bold', color: '#1B3A5C', marginBottom: 8 }}>
                  Drop or Upload your Assets file here
                </div>
                <div style={{ fontSize: 13, color: '#777', marginBottom: 24, lineHeight: 1.5 }}>
                  Supports CSV and Excel (.xlsx) files
                </div>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  style={{ display: 'none' }}
                  onChange={onFileInputChange}
                />
                <button
                  className="pma-btn"
                  onClick={() => fileInputRef.current?.click()}
                  style={{ background: '#1B3A5C', color: '#fff', border: 'none' }}
                >
                  Choose File
                </button>
              </div>

              {/* Error message */}
              {uploadError && (
                <div style={{
                  marginTop: 14,
                  maxWidth: 520,
                  width: '100%',
                  fontSize: 13,
                  color: '#b03020',
                  background: '#fdf3f2',
                  border: '1px solid #e8c4be',
                  borderRadius: 4,
                  padding: '9px 14px',
                }}>
                  {uploadError === 'filetype'
                    ? 'Please upload a CSV or Excel file.'
                    : 'This file is missing expected columns. Make sure you exported the Asset with Asset Notes table from Sigma.'
                  }
                </div>
              )}

            </div>
          ) : (
            /* ── Ready state ── */
            <>
              {/* Download confirmation banner */}
              <div style={{
                marginBottom: 16,
                padding: '10px 14px',
                background: '#edfaf1',
                border: '1px solid #82c99a',
                borderRadius: 4,
                fontSize: 13,
                color: '#1a5c33',
                display: 'flex',
                alignItems: 'center',
                gap: 10,
              }}>
                <span style={{ fontSize: 17, lineHeight: 1 }}>&#10003;</span>
                <span>
                  Your PMA Worksheet has been downloaded — check your Downloads folder.
                  Use the <strong>Export Excel</strong> button above to download it again.
                </span>
              </div>

              {/* Summary line */}
              <div style={{ marginBottom: 4, fontSize: 13, color: '#555' }}>
                <strong style={{ color: '#1B3A5C' }}>{appData.properties.length}</strong>
                {' '}{appData.properties.length === 1 ? 'property' : 'properties'}&nbsp;&nbsp;·&nbsp;&nbsp;
                <strong style={{ color: '#1B3A5C' }}>{appData.totalAssets}</strong>
                {' '}assets&nbsp;&nbsp;·&nbsp;&nbsp;
                {today}&nbsp;&nbsp;·&nbsp;&nbsp;
                <span style={{ color: '#888' }}>{fileName}</span>
              </div>
              <div style={{ marginBottom: 12 }}>
                <button className="pma-reset-link" onClick={handleReset}>
                  Upload a different file
                </button>
              </div>

              {/* Full data notice */}
              <div style={{
                marginBottom: 10,
                fontSize: 12,
                color: '#666',
                lineHeight: 1.6,
              }}>
                The table below shows all <strong>{appData.totalAssets}</strong> assets
                across all <strong>{appData.properties.length}</strong> {appData.properties.length === 1 ? 'property' : 'properties'} from
                your file — scroll down to review the complete data.
              </div>

              {/* Full data table — fixed-height scrollable window */}
              <div style={{
                overflowX: 'auto',
                overflowY: 'auto',
                maxHeight: 'calc(100vh - 280px)',
                minHeight: 200,
                border: '1px solid #d0d7e2',
                borderRadius: 3,
              }}>
                <table style={{
                  borderCollapse: 'collapse',
                  width: '100%',
                  fontSize: 12,
                  fontFamily: 'Arial, sans-serif',
                  whiteSpace: 'nowrap',
                }}>
                  <thead>
                    <tr>
                      {[
                        'Asset ID', 'Asset Type', 'Manufacturer', 'Model #', 'Serial #',
                        'Location', 'Area Served', 'Tonnage', 'F. Qty', 'Filters', 'B. Qty', 'Belts',
                      ].map(h => (
                        <th key={h} style={{
                          background: '#D9E1F2',
                          border: '1px solid #b0bcd8',
                          padding: '5px 9px',
                          textAlign: 'center',
                          fontWeight: 'bold',
                          position: 'sticky',
                          top: 0,
                          zIndex: 1,
                        }}>
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewItems.map((item, i) => {
                      if (item.kind === 'group') {
                        return (
                          <tr key={i}>
                            <td
                              colSpan={12}
                              style={{
                                background: '#1B3A5C',
                                color: '#fff',
                                padding: '5px 10px',
                                fontWeight: 'bold',
                                fontSize: 12,
                              }}
                            >
                              {item.label}
                            </td>
                          </tr>
                        );
                      }
                      const a = item.asset;
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? '#fff' : '#f5f7fb' }}>
                          {[
                            a.assetId, a.assetType, a.manufacturer, a.modelNo,
                            a.serialNo, a.location, a.areaServed, a.tonnage,
                            a.filterQty, a.filters, a.beltQty, a.belts,
                          ].map((v, ci) => (
                            <td key={ci} style={{
                              border: '1px solid #dde3ee',
                              padding: '4px 8px',
                            }}>
                              {v}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </main>
      </div>
    </>
  );
}
