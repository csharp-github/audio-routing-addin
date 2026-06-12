// ============================================================
// AUDIO ROUTING ADD-IN — taskpane.js
// ============================================================

// Sheet names
const SHEET = {
  DEVICES:     'Devices',
  PORTS:       'Ports',
  CONNECTIONS: 'Connections',
  CONSOLE:     'Console'
};

// Column indices (0-based) per sheet
const COL = {
  // Devices: DeviceID | Name | Type | Notes
  DEV: { ID:0, NAME:1, TYPE:2, NOTES:3 },
  // Ports: PortID | DeviceID | DeviceName | Direction | PortNum | Alias
  PORT: { ID:0, DEVICE_ID:1, DEVICE_NAME:2, DIR:3, NUM:4, ALIAS:5 },
  // Connections: ConnID | SrcPortID | SrcLabel | DstPortID | DstLabel | Notes
  CONN: { ID:0, SRC_ID:1, SRC_LABEL:2, DST_ID:3, DST_LABEL:4, NOTES:5 },
  // Console: Section | ChNum | ChName | MainIn | AltIn | InsertA_Send | InsertA_Ret | InsertB_Send | InsertB_Ret | DirectOut | Output
  CON: { SECTION:0, NUM:1, NAME:2, MAIN_IN:3, ALT_IN:4, INS_A_SND:5, INS_A_RET:6, INS_B_SND:7, INS_B_RET:8, DIRECT_OUT:9, OUTPUT:10 }
};

// Console structure: Quantum 38 mono channels + 4 stereo pairs + groups/auxes/matrices
const CONSOLE_STRUCTURE = {
  channels: { label: 'Channels', count: 38, hasInput: true, hasStereo: true },
  groups:   { label: 'Groups',   count: 16, hasInput: false },
  auxes:    { label: 'Auxes',    count: 16, hasInput: false },
  matrices: { label: 'Matrices', count: 8,  hasInput: false }
};

let state = {
  ready: false,
  devices: [],
  ports: [],
  connections: [],
  consoleRows: []
};

// ============================================================
// INIT
// ============================================================
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    setStatus('Connecting to workbook…');
    try {
      const hasSheets = await checkSheetsExist();
      if (hasSheets) {
        await loadAll();
        switchPanel('devices');
      } else {
        switchPanel('setup');
        setStatus('Ready — click Set Up to begin', false);
      }
    } catch(e) {
      setStatus('Error: ' + e.message, true);
    }
  }
});

async function checkSheetsExist() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    const names = sheets.items.map(s => s.name);
    return Object.values(SHEET).every(n => names.includes(n));
  });
}

// ============================================================
// WORKBOOK SETUP
// ============================================================
async function setupWorkbook() {
  setStatus('Creating sheets…');
  try {
    await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const sheets = wb.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const existing = sheets.items.map(s => s.name);

      // Create sheets if missing
      for (const name of Object.values(SHEET)) {
        if (!existing.includes(name)) {
          sheets.add(name);
        }
      }
      await ctx.sync();

      // Write headers
      const headers = {
        [SHEET.DEVICES]:     [['DeviceID','Name','Type','Notes']],
        [SHEET.PORTS]:       [['PortID','DeviceID','DeviceName','Direction','PortNum','Alias']],
        [SHEET.CONNECTIONS]: [['ConnID','SrcPortID','SrcLabel','DstPortID','DstLabel','Notes']],
        [SHEET.CONSOLE]:     [['Section','ChNum','ChName','MainIn','AltIn','InsertA_Send','InsertA_Ret','InsertB_Send','InsertB_Ret','DirectOut','Output']]
      };

      for (const [name, rows] of Object.entries(headers)) {
        const sh = wb.worksheets.getItem(name);
        const r = sh.getRange('A1').getResizedRange(0, rows[0].length - 1);
        r.values = rows;
        r.format.font.bold = true;
        r.format.fill.color = '#1a1a2e';
        r.format.font.color = '#7dd3fc';
      }

      // Pre-populate Console sheet with structure rows
      await buildConsoleSheetStructure(ctx);

      await ctx.sync();
    });

    await loadAll();
    switchPanel('devices');
    setStatus('Workbook ready');
  } catch(e) {
    setStatus('Setup error: ' + e.message, true);
    console.error(e);
  }
}

async function buildConsoleSheetStructure(ctx) {
  const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
  let row = 2;

  const sections = [
    { key: 'channels', label: 'CHANNELS', count: 38 },
    { key: 'groups',   label: 'GROUPS',   count: 16 },
    { key: 'auxes',    label: 'AUXES',    count: 16 },
    { key: 'matrices', label: 'MATRICES', count: 8  }
  ];

  for (const sec of sections) {
    for (let i = 1; i <= sec.count; i++) {
      const r = sh.getRange(`A${row}`).getResizedRange(0, 10);
      r.values = [[sec.key, i, '', '', '', '', '', '', '', '', '']];
      row++;
    }
  }
  // Don't sync here — caller will sync
}

// ============================================================
// LOAD ALL DATA
// ============================================================
async function loadAll() {
  await Promise.all([
    loadDevicesData(),
    loadPortsData(),
    loadConnectionsData(),
    loadConsoleData()
  ]);
  state.ready = true;
  setStatus(`${state.devices.length} devices · ${state.ports.length} ports · ${state.connections.length} connections`);
  renderAll();
}

async function loadDevicesData() {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
    const range = sh.getUsedRange();
    range.load('values');
    await ctx.sync();
    const rows = range.values.slice(1).filter(r => r[COL.DEV.ID]);
    state.devices = rows.map(r => ({
      id:    r[COL.DEV.ID],
      name:  r[COL.DEV.NAME],
      type:  r[COL.DEV.TYPE],
      notes: r[COL.DEV.NOTES]
    }));
  });
}

async function loadPortsData() {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
    const range = sh.getUsedRange();
    range.load('values');
    await ctx.sync();
    const rows = range.values.slice(1).filter(r => r[COL.PORT.ID]);
    state.ports = rows.map(r => ({
      id:        r[COL.PORT.ID],
      deviceId:  r[COL.PORT.DEVICE_ID],
      deviceName:r[COL.PORT.DEVICE_NAME],
      dir:       r[COL.PORT.DIR],
      num:       r[COL.PORT.NUM],
      alias:     r[COL.PORT.ALIAS]
    }));
  });
}

async function loadConnectionsData() {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
    const range = sh.getUsedRange();
    range.load('values');
    await ctx.sync();
    const rows = range.values.slice(1).filter(r => r[COL.CONN.ID]);
    state.connections = rows.map(r => ({
      id:       r[COL.CONN.ID],
      srcId:    r[COL.CONN.SRC_ID],
      srcLabel: r[COL.CONN.SRC_LABEL],
      dstId:    r[COL.CONN.DST_ID],
      dstLabel: r[COL.CONN.DST_LABEL],
      notes:    r[COL.CONN.NOTES]
    }));
  });
}

async function loadConsoleData() {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
    const range = sh.getUsedRange();
    range.load('values');
    await ctx.sync();
    const rows = range.values.slice(1).filter(r => r[COL.CON.SECTION]);
    state.consoleRows = rows.map((r, i) => ({
      rowIndex: i + 2, // 1-based Excel row
      section:  r[COL.CON.SECTION],
      num:      r[COL.CON.NUM],
      name:     r[COL.CON.NAME],
      mainIn:   r[COL.CON.MAIN_IN],
      altIn:    r[COL.CON.ALT_IN],
      insASnd:  r[COL.CON.INS_A_SND],
      insARet:  r[COL.CON.INS_A_RET],
      insBSnd:  r[COL.CON.INS_B_SND],
      insBRet:  r[COL.CON.INS_B_RET],
      directOut:r[COL.CON.DIRECT_OUT],
      output:   r[COL.CON.OUTPUT]
    }));
  });
}

// ============================================================
// RENDER ALL
// ============================================================
function renderAll() {
  renderDeviceList();
  refreshDeviceDropdowns();
  renderPortList();
  refreshPortDropdowns();
  renderConnectionList();
  renderConsole();
}

// ============================================================
// DEVICES
// ============================================================
async function addDevice() {
  const name  = val('dev-name').trim();
  const type  = val('dev-type');
  const notes = val('dev-notes').trim();
  if (!name) { alert('Device name is required.'); return; }

  const id = 'D' + Date.now();
  try {
    await Excel.run(async (ctx) => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const used = sh.getUsedRange();
      used.load('rowCount');
      await ctx.sync();
      const nextRow = used.rowCount + 1;
      sh.getRange(`A${nextRow}`).getResizedRange(0, 3).values = [[id, name, type, notes]];
      await ctx.sync();
    });
    clearFields(['dev-name', 'dev-notes']);
    await loadDevicesData();
    renderDeviceList();
    refreshDeviceDropdowns();
    setStatus(`Device "${name}" added`);
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

function renderDeviceList() {
  const el = document.getElementById('device-list');
  if (!state.devices.length) {
    el.innerHTML = `<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/></svg>No devices yet</div>`;
    return;
  }
  el.innerHTML = state.devices.map(d => `
    <div class="card">
      <div class="card-row">
        <div class="card-title">${esc(d.name)}</div>
        <span class="badge badge-type">${esc(d.type)||'—'}</span>
        <button class="btn btn-danger btn-sm" onclick="deleteDevice('${d.id}')">✕</button>
      </div>
      ${d.notes ? `<div class="card-sub">${esc(d.notes)}</div>` : ''}
      <div class="card-sub" style="margin-top:4px; font-size:10px; color:var(--text3)">ID: ${d.id}</div>
    </div>
  `).join('');
}

async function deleteDevice(id) {
  if (!confirm('Delete this device? Its ports and connections will remain.')) return;
  try {
    await deleteRowById(SHEET.DEVICES, COL.DEV.ID, id);
    await loadDevicesData();
    renderDeviceList();
    refreshDeviceDropdowns();
    setStatus('Device deleted');
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

async function loadDeviceList() {
  await loadDevicesData();
  renderDeviceList();
  refreshDeviceDropdowns();
}

// ============================================================
// PORTS
// ============================================================
async function addPort() {
  const deviceId = val('port-device');
  const dir      = val('port-dir');
  const num      = parseInt(val('port-num'));
  const alias    = val('port-alias').trim();

  if (!deviceId) { alert('Select a device.'); return; }
  if (!num || num < 1) { alert('Enter a valid port number.'); return; }

  const device = state.devices.find(d => d.id === deviceId);
  // Check duplicate
  const dupe = state.ports.find(p => p.deviceId === deviceId && p.dir === dir && p.num == num);
  if (dupe) { alert(`Port ${dir} ${num} already exists for ${device.name}.`); return; }

  const id = 'P' + Date.now();
  try {
    await Excel.run(async (ctx) => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const used = sh.getUsedRange();
      used.load('rowCount');
      await ctx.sync();
      const nextRow = used.rowCount + 1;
      sh.getRange(`A${nextRow}`).getResizedRange(0, 5).values =
        [[id, deviceId, device.name, dir, num, alias]];
      await ctx.sync();
    });
    clearFields(['port-num', 'port-alias']);
    await loadPortsData();
    renderPortList();
    refreshPortDropdowns();
    setStatus(`Port ${dir} ${num} added for ${device.name}`);
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

function bulkAddPorts() {
  document.getElementById('bulk-wrap').classList.toggle('hidden');
}

async function confirmBulkAdd() {
  const deviceId = val('port-device');
  if (!deviceId) { alert('Select a device first.'); return; }
  const from = parseInt(val('bulk-from'));
  const to   = parseInt(val('bulk-to'));
  const dir  = val('bulk-dir');
  if (from > to) { alert('From must be ≤ To.'); return; }

  const device = state.devices.find(d => d.id === deviceId);
  const rows = [];
  for (let n = from; n <= to; n++) {
    const exists = state.ports.find(p => p.deviceId === deviceId && p.dir === dir && p.num == n);
    if (!exists) rows.push(['P' + Date.now() + '_' + n, deviceId, device.name, dir, n, '']);
  }
  if (!rows.length) { alert('All ports in that range already exist.'); return; }

  try {
    await Excel.run(async (ctx) => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const used = sh.getUsedRange();
      used.load('rowCount');
      await ctx.sync();
      let nextRow = used.rowCount + 1;
      for (const r of rows) {
        sh.getRange(`A${nextRow}`).getResizedRange(0, 5).values = [r];
        nextRow++;
      }
      await ctx.sync();
    });
    document.getElementById('bulk-wrap').classList.add('hidden');
    await loadPortsData();
    renderPortList();
    refreshPortDropdowns();
    setStatus(`${rows.length} ports added`);
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

function portLabel(p) {
  const alias = p.alias ? ` — ${p.alias}` : '';
  return `${p.deviceName} / ${p.dir} ${p.num}${alias}`;
}

function renderPortList() {
  const q = val('port-search').toLowerCase();
  const filterDev = val('port-filter-device');
  const filterDir = val('port-filter-dir');

  let ports = state.ports.filter(p => {
    const label = portLabel(p).toLowerCase();
    return (!q || label.includes(q)) &&
           (!filterDev || p.deviceId === filterDev) &&
           (!filterDir || p.dir === filterDir);
  });

  const el = document.getElementById('port-list');
  if (!ports.length) {
    el.innerHTML = `<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.22 4.22l2.12 2.12M17.66 17.66l2.12 2.12M2 12h3M19 12h3M4.22 19.78l2.12-2.12M17.66 6.34l2.12-2.12"/></svg>No ports</div>`;
    return;
  }

  // Group by device
  const byDevice = {};
  ports.forEach(p => {
    if (!byDevice[p.deviceName]) byDevice[p.deviceName] = [];
    byDevice[p.deviceName].push(p);
  });

  el.innerHTML = Object.entries(byDevice).map(([devName, ps]) => `
    <div style="margin-bottom:10px;">
      <div style="font-size:10px; color:var(--text2); text-transform:uppercase; letter-spacing:.06em; margin-bottom:5px;">${esc(devName)}</div>
      ${ps.map(p => `
        <div class="card" style="padding:7px 10px; margin-bottom:4px;">
          <div class="card-row">
            <span class="badge ${p.dir === 'IN' ? 'badge-in' : 'badge-out'}">${p.dir}</span>
            <span style="font-family:var(--mono); font-size:11px; color:var(--text2); width:28px;">${p.num}</span>
            <span class="card-title" style="font-size:11px;">${p.alias ? esc(p.alias) : '<span style="color:var(--text3)">unassigned</span>'}</span>
            <button class="btn btn-danger btn-sm" onclick="deletePort('${p.id}')">✕</button>
          </div>
        </div>
      `).join('')}
    </div>
  `).join('');
}

function filterPortList() { renderPortList(); }

async function deletePort(id) {
  if (!confirm('Delete this port?')) return;
  try {
    await deleteRowById(SHEET.PORTS, COL.PORT.ID, id);
    await loadPortsData();
    renderPortList();
    refreshPortDropdowns();
    setStatus('Port deleted');
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

async function loadPortList() {
  await loadPortsData();
  renderPortList();
  refreshPortDropdowns();
}

// ============================================================
// CONNECTIONS
// ============================================================
async function addConnection() {
  const srcId = val('conn-src');
  const dstId = val('conn-dst');
  const notes = val('conn-notes').trim();

  if (!srcId || !dstId) { alert('Select both source and destination ports.'); return; }
  if (srcId === dstId)  { alert('Source and destination cannot be the same port.'); return; }

  const src = state.ports.find(p => p.id === srcId);
  const dst = state.ports.find(p => p.id === dstId);

  if (src.dir !== 'OUT') { alert('Source must be an OUTPUT port.'); return; }
  if (dst.dir !== 'IN')  { alert('Destination must be an INPUT port.'); return; }

  // Check duplicate output
  const dupeSrc = state.connections.find(c => c.srcId === srcId);
  if (dupeSrc) {
    if (!confirm(`Warning: ${portLabel(src)} is already connected to ${dupeSrc.dstLabel}. Add another connection anyway?`)) return;
  }

  const id = 'C' + Date.now();
  try {
    await Excel.run(async (ctx) => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
      const used = sh.getUsedRange();
      used.load('rowCount');
      await ctx.sync();
      const nextRow = used.rowCount + 1;
      sh.getRange(`A${nextRow}`).getResizedRange(0, 5).values =
        [[id, srcId, portLabel(src), dstId, portLabel(dst), notes]];
      await ctx.sync();
    });
    clearFields(['conn-notes']);
    await loadConnectionsData();
    renderConnectionList();
    setStatus(`Connection added: ${portLabel(src)} → ${portLabel(dst)}`);
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

function renderConnectionList() {
  const q = val('conn-search').toLowerCase();
  const conns = state.connections.filter(c =>
    !q || c.srcLabel.toLowerCase().includes(q) || c.dstLabel.toLowerCase().includes(q)
  );
  const el = document.getElementById('connection-list');
  if (!conns.length) {
    el.innerHTML = `<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>No connections yet</div>`;
    return;
  }
  el.innerHTML = conns.map(c => `
    <div class="card">
      <div class="card-row">
        <div style="flex:1">
          <div style="font-size:11px; color:var(--orange);">↑ ${esc(c.srcLabel)}</div>
          <div style="font-size:11px; color:var(--green); margin-top:2px;">↓ ${esc(c.dstLabel)}</div>
          ${c.notes ? `<div class="card-sub" style="margin-top:3px;">${esc(c.notes)}</div>` : ''}
        </div>
        <button class="btn btn-danger btn-sm" onclick="deleteConnection('${c.id}')">✕</button>
      </div>
    </div>
  `).join('');
}

function filterConnectionList() { renderConnectionList(); }

async function deleteConnection(id) {
  if (!confirm('Delete this connection?')) return;
  try {
    await deleteRowById(SHEET.CONNECTIONS, COL.CONN.ID, id);
    await loadConnectionsData();
    renderConnectionList();
    setStatus('Connection deleted');
  } catch(e) { setStatus('Error: ' + e.message, true); }
}

async function loadConnectionList() {
  await loadConnectionsData();
  renderConnectionList();
}

// ============================================================
// CONSOLE
// ============================================================
function renderConsole() {
  const el = document.getElementById('console-content');
  if (!state.consoleRows.length) {
    el.innerHTML = `<div class="empty">Console not set up yet. Run workbook setup first.</div>`;
    return;
  }

  // Show info alert on first render
  document.getElementById('console-alert').classList.remove('hidden');

  const sections = ['channels', 'groups', 'auxes', 'matrices'];
  const sectionLabels = { channels: 'Channels', groups: 'Groups', auxes: 'Auxes', matrices: 'Matrices' };

  const allPorts = state.ports;
  const inputPorts  = allPorts.filter(p => p.dir === 'IN');
  const outputPorts = allPorts.filter(p => p.dir === 'OUT');

  function portOption(p) {
    return `<option value="${p.id}">${esc(portLabel(p))}</option>`;
  }

  function selectField(label, currentVal, portList, fieldKey, rowIndex) {
    const hasVal = currentVal && currentVal !== '';
    return `
      <div class="port-field">
        <label>${label}</label>
        <select class="${hasVal ? 'has-value' : ''}"
          onchange="consoleFieldChanged(this, '${fieldKey}', ${rowIndex})">
          <option value="">—</option>
          ${portList.map(p => `<option value="${p.id}" ${currentVal === p.id || currentVal === portLabel(p) ? 'selected' : ''}>${esc(portLabel(p))}</option>`).join('')}
        </select>
      </div>`;
  }

  let html = '';
  for (const sec of sections) {
    const rows = state.consoleRows.filter(r => r.section === sec);
    if (!rows.length) continue;

    html += `<div class="console-section">
      <div class="console-section-title">${sectionLabels[sec]} (${rows.length})</div>`;

    for (const row of rows) {
      const isChannel = sec === 'channels';
      const nameDisplay = row.name || `<span style="color:var(--text3)">${sec.slice(0,-1)} ${row.num}</span>`;

      html += `<div class="channel-row">
        <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
          <span class="channel-num">${row.num}</span>
          <span class="channel-name">${row.name ? esc(row.name) : `<span style="color:var(--text3); font-weight:400; font-size:11px">${sec.slice(0,-1)} ${row.num}</span>`}</span>
        </div>
        <div class="port-grid">`;

      if (isChannel) {
        html += selectField('Main In',       row.mainIn,    inputPorts,  'mainIn',    row.rowIndex);
        html += selectField('Alt In',        row.altIn,     inputPorts,  'altIn',     row.rowIndex);
        html += selectField('InsA Send',     row.insASnd,   outputPorts, 'insASnd',   row.rowIndex);
        html += selectField('InsA Return',   row.insARet,   inputPorts,  'insARet',   row.rowIndex);
        html += selectField('InsB Send',     row.insBSnd,   outputPorts, 'insBSnd',   row.rowIndex);
        html += selectField('InsB Return',   row.insBRet,   inputPorts,  'insBRet',   row.rowIndex);
        html += selectField('Direct Out',    row.directOut, outputPorts, 'directOut', row.rowIndex);
      } else {
        html += selectField('Output',        row.output,    outputPorts, 'output',    row.rowIndex);
      }

      html += `</div></div>`;
    }
    html += `</div>`;
  }

  el.innerHTML = html;
}

async function consoleFieldChanged(selectEl, fieldKey, rowIndex) {
  const portId = selectEl.value;
  const port   = state.ports.find(p => p.id === portId);
  const label  = port ? portLabel(port) : '';

  // Map fieldKey to COL index
  const colMap = {
    mainIn:    COL.CON.MAIN_IN,
    altIn:     COL.CON.ALT_IN,
    insASnd:   COL.CON.INS_A_SND,
    insARet:   COL.CON.INS_A_RET,
    insBSnd:   COL.CON.INS_B_SND,
    insBRet:   COL.CON.INS_B_RET,
    directOut: COL.CON.DIRECT_OUT,
    output:    COL.CON.OUTPUT
  };

  const colIndex = colMap[fieldKey];
  if (colIndex === undefined) return;

  // Write to Console sheet
  try {
    await Excel.run(async (ctx) => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const colLetter = String.fromCharCode(65 + colIndex);
      const cell = sh.getRange(`${colLetter}${rowIndex}`);
      cell.values = [[label]];

      // If this is mainIn, auto-update the channel name column too
      if (fieldKey === 'mainIn' && label) {
        const nameCell = sh.getRange(`C${rowIndex}`);
        nameCell.load('values');
        await ctx.sync();
        // Only auto-name if currently empty
        if (!nameCell.values[0][0]) {
          nameCell.values = [[label]];
        }
      }
      await ctx.sync();
    });

    // Update state
    const stateRow = state.consoleRows.find(r => r.rowIndex === rowIndex);
    if (stateRow) stateRow[fieldKey] = label;

    // Mark the select visually
    selectEl.classList.toggle('has-value', !!portId);

    // Record connection automatically if it makes sense
    if (port) await autoRecordConnection(fieldKey, port, rowIndex);

    setStatus(`Console updated — row ${rowIndex}`);
  } catch(e) { setStatus('Error: ' + e.message, true); console.error(e); }
}

async function autoRecordConnection(fieldKey, port, rowIndex) {
  // For inserts: send is an output, return is an input — those create connections externally
  // For mainIn/altIn: the console itself is the destination
  // We just record whatever was assigned — the connection is implicit in the Console sheet
  // For explicit routing (inserts), auto-add to Connections if both halves exist
  const row = state.consoleRows.find(r => r.rowIndex === rowIndex);
  if (!row) return;

  // Auto-connect insert pairs when both sides are set
  const pairs = [
    { snd: 'insASnd', ret: 'insARet' },
    { snd: 'insBSnd', ret: 'insBRet' }
  ];

  for (const pair of pairs) {
    if ((fieldKey === pair.snd || fieldKey === pair.ret) && row[pair.snd] && row[pair.ret]) {
      const sndPort = state.ports.find(p => portLabel(p) === row[pair.snd]);
      const retPort = state.ports.find(p => portLabel(p) === row[pair.ret]);
      if (sndPort && retPort) {
        const exists = state.connections.find(c => c.srcId === sndPort.id && c.dstId === retPort.id);
        if (!exists) {
          const id = 'C' + Date.now();
          await Excel.run(async (ctx) => {
            const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
            const used = sh.getUsedRange();
            used.load('rowCount');
            await ctx.sync();
            const nextRow = used.rowCount + 1;
            sh.getRange(`A${nextRow}`).getResizedRange(0, 5).values =
              [[id, sndPort.id, portLabel(sndPort), retPort.id, portLabel(retPort), 'Auto (insert pair)']];
            await ctx.sync();
          });
          await loadConnectionsData();
        }
      }
    }
  }
}

// ============================================================
// CHECKS
// ============================================================
async function runChecks() {
  await loadAll();
  const issues = [];
  const ok = [];

  // 1. Outputs used more than once
  const srcCounts = {};
  state.connections.forEach(c => {
    if (!srcCounts[c.srcId]) srcCounts[c.srcId] = [];
    srcCounts[c.srcId].push(c);
  });
  Object.entries(srcCounts).forEach(([portId, conns]) => {
    if (conns.length > 1) {
      issues.push({
        type: 'conflict',
        msg: `Output used ${conns.length}×: <strong>${esc(conns[0].srcLabel)}</strong>`,
        detail: conns.map(c => `→ ${esc(c.dstLabel)}`).join(', ')
      });
    }
  });
  if (!Object.values(srcCounts).find(c => c.length > 1)) ok.push('No duplicate outputs');

  // 2. Direction mismatches
  state.connections.forEach(c => {
    const src = state.ports.find(p => p.id === c.srcId);
    const dst = state.ports.find(p => p.id === c.dstId);
    if (src && src.dir !== 'OUT') {
      issues.push({ type: 'conflict', msg: `Direction mismatch: source <strong>${esc(c.srcLabel)}</strong> is not an OUTPUT` });
    }
    if (dst && dst.dir !== 'IN') {
      issues.push({ type: 'conflict', msg: `Direction mismatch: destination <strong>${esc(c.dstLabel)}</strong> is not an INPUT` });
    }
  });
  if (!issues.find(i => i.msg.includes('Direction mismatch'))) ok.push('No direction mismatches');

  // 3. Orphaned ports (have no device)
  const deviceIds = new Set(state.devices.map(d => d.id));
  const orphans = state.ports.filter(p => !deviceIds.has(p.deviceId));
  if (orphans.length) {
    issues.push({ type: 'warn', msg: `${orphans.length} port(s) reference deleted devices`, detail: orphans.map(p => portLabel(p)).join(', ') });
  } else {
    ok.push('All ports have valid devices');
  }

  // 4. Unconnected outputs
  const connectedOuts = new Set(state.connections.map(c => c.srcId));
  const unconnected = state.ports.filter(p => p.dir === 'OUT' && !connectedOuts.has(p.id));
  if (unconnected.length) {
    issues.push({ type: 'info', msg: `${unconnected.length} output(s) not connected to anything`, detail: unconnected.map(p => portLabel(p)).join(', ') });
  } else if (state.ports.filter(p => p.dir === 'OUT').length > 0) {
    ok.push('All outputs are connected');
  }

  // 5. Stats
  const el = document.getElementById('checks-content');
  let html = '';

  // Summary row
  html += `<div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; margin-bottom:14px;">
    ${[['Devices', state.devices.length], ['Ports', state.ports.length], ['Connections', state.connections.length]].map(([l,n]) => `
      <div class="card" style="text-align:center; padding:10px 6px;">
        <div style="font-size:18px; font-weight:600; color:var(--text)">${n}</div>
        <div style="font-size:10px; color:var(--text2); text-transform:uppercase; letter-spacing:.06em">${l}</div>
      </div>`).join('')}
  </div>`;

  if (ok.length) {
    html += ok.map(msg => `
      <div class="alert alert-success">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
        ${msg}
      </div>`).join('');
  }

  if (issues.length) {
    html += `<div class="divider"></div>`;
    html += issues.map(i => {
      const cls = i.type === 'conflict' ? 'alert-error' : i.type === 'warn' ? 'alert-warn' : 'alert-info';
      const icon = i.type === 'conflict'
        ? `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12" y2="16.01"/></svg>`
        : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12" y2="17.01"/></svg>`;
      return `<div class="alert ${cls}">${icon}<div><div>${i.msg}</div>${i.detail ? `<div style="margin-top:4px; font-size:10px; opacity:.8">${esc(i.detail)}</div>` : ''}</div></div>`;
    }).join('');
  } else {
    html += `<div class="divider"></div><div class="alert alert-success"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg><strong>All checks passed!</strong></div>`;
  }

  el.innerHTML = html;
  setStatus(`Checks complete — ${issues.length} issue(s) found`);
}

// ============================================================
// DROPDOWN HELPERS
// ============================================================
function refreshDeviceDropdowns() {
  const selects = ['port-device', 'port-filter-device'];
  selects.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const current = el.value;
    el.innerHTML = `<option value="">${id === 'port-filter-device' ? 'All devices' : '— select device —'}</option>` +
      state.devices.map(d => `<option value="${d.id}" ${d.id === current ? 'selected':''}>${esc(d.name)}</option>`).join('');
  });
}

function refreshPortDropdowns() {
  const outPorts = state.ports.filter(p => p.dir === 'OUT');
  const inPorts  = state.ports.filter(p => p.dir === 'IN');

  setSelectOptions('conn-src', outPorts, '— select output port —');
  setSelectOptions('conn-dst', inPorts,  '— select input port —');
}

function setSelectOptions(id, ports, placeholder) {
  const el = document.getElementById(id);
  if (!el) return;
  const current = el.value;
  el.innerHTML = `<option value="">${placeholder}</option>` +
    ports.map(p => `<option value="${p.id}" ${p.id === current ? 'selected':''}>${esc(portLabel(p))}</option>`).join('');
}

// ============================================================
// UTILITY — delete a row by matching ID in column
// ============================================================
async function deleteRowById(sheetName, colIndex, targetId) {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(sheetName);
    const used = sh.getUsedRange();
    used.load('values,rowCount');
    await ctx.sync();
    const vals = used.values;
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][colIndex]) === String(targetId)) {
        sh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
        break;
      }
    }
    await ctx.sync();
  });
}

// ============================================================
// UI HELPERS
// ============================================================
function switchPanel(name) {
  document.querySelectorAll('.panel').forEach(p => { p.classList.add('hidden'); p.classList.remove('active'); });
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const panel = document.getElementById('panel-' + name);
  const nav   = document.getElementById('nav-' + name);
  if (panel) { panel.classList.remove('hidden'); panel.classList.add('active'); }
  if (nav)   nav.classList.add('active');
}

function setStatus(msg, isError = false) {
  document.getElementById('status-text').textContent = msg;
  document.getElementById('status-dot').className = 'status-dot' + (isError ? ' error' : '');
}

function val(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function clearFields(ids) {
  ids.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
}

function esc(str) {
  if (!str && str !== 0) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
