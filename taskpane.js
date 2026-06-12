// ============================================================
// AUDIO ROUTING ADD-IN v2 — taskpane.js
// ============================================================

const SHEET = { DEVICES:'Devices', PORTS:'Ports', CONNECTIONS:'Connections', CONSOLE:'Console' };

// Column indices (0-based)
const COL = {
  DEV:  { ID:0, NAME:1, NOTES:2 },
  PORT: { ID:0, DEVICE_ID:1, DEVICE_NAME:2, DIR:3, NUM:4, ALIAS:5 },
  CONN: { ID:0, SRC_PORT_ID:1, SRC_LABEL:2, DST_PORT_ID:3, DST_LABEL:4, NOTES:5 },
  CON:  { ID:0, TYPE:1, WIDTH:2, NAME:3, MAIN_IN:4, ALT_IN:5, INS_A_SND:6, INS_A_RET:7, INS_B_SND:8, INS_B_RET:9, DIRECT_OUT:10, OUTPUT:11 }
};

let S = { ready:false, devices:[], ports:[], connections:[], consoleStrips:[] };

// Matrix expand state: deviceId -> { rowOpen, colOpen }
let matrixState = {};

// ============================================================
// INIT
// ============================================================
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  setStatus('Connecting…');
  try {
    const ok = await checkSheetsExist();
    if (ok) { await loadAll(); switchPanel('devices'); }
    else { switchPanel('setup'); setStatus('Ready — set up workbook to begin', false); }
  } catch(e) { setStatus('Error: ' + e.message, true); }
});

async function checkSheetsExist() {
  return Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    const names = sheets.items.map(s => s.name);
    return Object.values(SHEET).every(n => names.includes(n));
  });
}

// ============================================================
// SETUP
// ============================================================
async function setupWorkbook() {
  setStatus('Creating sheets…');
  try {
    await Excel.run(async ctx => {
      const wb = ctx.workbook;
      const sheets = wb.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const existing = sheets.items.map(s => s.name);
      for (const name of Object.values(SHEET)) {
        if (!existing.includes(name)) sheets.add(name);
      }
      await ctx.sync();
      const hdrs = {
        [SHEET.DEVICES]:     ['DeviceID','Name','Notes'],
        [SHEET.PORTS]:       ['PortID','DeviceID','DeviceName','Direction','PortNum','Alias'],
        [SHEET.CONNECTIONS]: ['ConnID','SrcPortID','SrcLabel','DstPortID','DstLabel','Notes'],
        [SHEET.CONSOLE]:     ['StripID','Type','Width','Name','MainIn','AltIn','InsA_Send','InsA_Ret','InsB_Send','InsB_Ret','DirectOut','Output']
      };
      for (const [name, cols] of Object.entries(hdrs)) {
        const sh = wb.worksheets.getItem(name);
        const r = sh.getRange('A1').getResizedRange(0, cols.length - 1);
        r.values = [cols];
        r.format.font.bold = true;
        r.format.fill.color = '#111827';
        r.format.font.color = '#60a5fa';
      }
      await ctx.sync();
    });
    await loadAll();
    switchPanel('devices');
    setStatus('Workbook ready');
  } catch(e) { setStatus('Setup error: ' + e.message, true); console.error(e); }
}

// ============================================================
// LOAD
// ============================================================
async function loadAll() {
  await Promise.all([loadDevices(), loadPorts(), loadConnections(), loadConsole()]);
  S.ready = true;
  setStatus(`${S.devices.length} devices · ${S.ports.length} ports · ${S.connections.length} connections`);
  renderAll();
}

async function loadDevices() {
  return Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
    const r = sh.getUsedRange(); r.load('values');
    await ctx.sync();
    S.devices = r.values.slice(1).filter(r => r[COL.DEV.ID]).map(r => ({
      id: r[COL.DEV.ID], name: r[COL.DEV.NAME], notes: r[COL.DEV.NOTES]
    }));
  });
}

async function loadPorts() {
  return Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
    const r = sh.getUsedRange(); r.load('values');
    await ctx.sync();
    S.ports = r.values.slice(1).filter(r => r[COL.PORT.ID]).map(r => ({
      id: r[COL.PORT.ID], deviceId: r[COL.PORT.DEVICE_ID], deviceName: r[COL.PORT.DEVICE_NAME],
      dir: r[COL.PORT.DIR], num: r[COL.PORT.NUM], alias: r[COL.PORT.ALIAS] || ''
    }));
  });
}

async function loadConnections() {
  return Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
    const r = sh.getUsedRange(); r.load('values');
    await ctx.sync();
    S.connections = r.values.slice(1).filter(r => r[COL.CONN.ID]).map(r => ({
      id: r[COL.CONN.ID], srcId: r[COL.CONN.SRC_PORT_ID], srcLabel: r[COL.CONN.SRC_LABEL],
      dstId: r[COL.CONN.DST_PORT_ID], dstLabel: r[COL.CONN.DST_LABEL], notes: r[COL.CONN.NOTES] || ''
    }));
  });
}

async function loadConsole() {
  return Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
    const r = sh.getUsedRange(); r.load('values');
    await ctx.sync();
    S.consoleStrips = r.values.slice(1).filter(r => r[COL.CON.ID]).map(r => ({
      id: r[COL.CON.ID], type: r[COL.CON.TYPE], width: r[COL.CON.WIDTH],
      name: r[COL.CON.NAME], mainIn: r[COL.CON.MAIN_IN], altIn: r[COL.CON.ALT_IN],
      insASnd: r[COL.CON.INS_A_SND], insARet: r[COL.CON.INS_A_RET],
      insBSnd: r[COL.CON.INS_B_SND], insBRet: r[COL.CON.INS_B_RET],
      directOut: r[COL.CON.DIRECT_OUT], output: r[COL.CON.OUTPUT]
    }));
  });
}

function renderAll() {
  renderDeviceList();
  refreshDeviceFilter();
  renderPortList();
  renderConsole();
  renderMatrix();
}

// ============================================================
// DEVICES
// ============================================================
async function addDevice() {
  const name    = v('dev-name').trim();
  const inputs  = parseInt(v('dev-inputs'))  || 0;
  const outputs = parseInt(v('dev-outputs')) || 0;
  const notes   = v('dev-notes').trim();
  if (!name) { alert('Device name is required.'); return; }

  const devId = 'D' + Date.now();
  const portRows = [];
  for (let i = 1; i <= inputs;  i++) portRows.push(['P' + Date.now() + '_I' + i, devId, name, 'IN',  i, '']);
  for (let i = 1; i <= outputs; i++) portRows.push(['P' + (Date.now()+1) + '_O' + i, devId, name, 'OUT', i, '']);

  try {
    await Excel.run(async ctx => {
      // Write device
      const dsh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const du = dsh.getUsedRange(); du.load('rowCount'); await ctx.sync();
      dsh.getRange(`A${du.rowCount+1}`).getResizedRange(0, 2).values = [[devId, name, notes]];

      // Write ports
      if (portRows.length) {
        const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
        const pu = psh.getUsedRange(); pu.load('rowCount'); await ctx.sync();
        let nr = pu.rowCount + 1;
        for (const row of portRows) {
          psh.getRange(`A${nr}`).getResizedRange(0, 5).values = [row];
          nr++;
        }
      }
      await ctx.sync();
    });
    clearFields(['dev-name','dev-notes']);
    document.getElementById('dev-inputs').value  = '0';
    document.getElementById('dev-outputs').value = '0';
    await loadDevices(); await loadPorts();
    renderDeviceList(); refreshDeviceFilter(); renderPortList(); renderMatrix();
    setStatus(`"${name}" added with ${inputs} inputs, ${outputs} outputs`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

function renderDeviceList() {
  const el = document.getElementById('device-list');
  if (!S.devices.length) {
    el.innerHTML = `<div class="empty">No devices yet — add one above.</div>`; return;
  }
  el.innerHTML = S.devices.map(d => {
    const ins  = S.ports.filter(p => p.deviceId === d.id && p.dir === 'IN').length;
    const outs = S.ports.filter(p => p.deviceId === d.id && p.dir === 'OUT').length;
    return `
    <div class="device-card" id="dcard-${d.id}">
      <div class="device-card-header" onclick="toggleDeviceCard('${d.id}')">
        <svg class="chevron" id="chev-${d.id}" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M9 18l6-6-6-6"/></svg>
        <span class="card-title">${esc(d.name)}</span>
        <div class="io-summary">
          ${ins  ? `<span class="io-pill io-pill-in">${ins} IN</span>`   : ''}
          ${outs ? `<span class="io-pill io-pill-out">${outs} OUT</span>` : ''}
        </div>
        <button class="btn btn-danger btn-xs" onclick="event.stopPropagation(); confirmDeleteDevice('${d.id}')">✕</button>
      </div>
      <div class="device-card-body" id="dbody-${d.id}">
        <div class="io-row">
          <div class="io-field">
            <label>Inputs</label>
            <input type="number" id="io-in-${d.id}" value="${ins}" min="0" />
          </div>
          <div class="io-field">
            <label>Outputs</label>
            <input type="number" id="io-out-${d.id}" value="${outs}" min="0" />
          </div>
        </div>
        <div class="btn-row">
          <button class="btn btn-primary btn-sm" onclick="updateDeviceIO('${d.id}', '${esc(d.name)}')">Update I/O</button>
        </div>
        ${d.notes ? `<div class="card-sub" style="margin-top:8px;">${esc(d.notes)}</div>` : ''}
      </div>
    </div>`;
  }).join('');
}

function toggleDeviceCard(id) {
  const body  = document.getElementById('dbody-'+id);
  const chev  = document.getElementById('chev-'+id);
  const open  = body.classList.toggle('open');
  chev.classList.toggle('open', open);
}

async function updateDeviceIO(devId, devName) {
  const newIn  = parseInt(document.getElementById('io-in-'+devId).value)  || 0;
  const newOut = parseInt(document.getElementById('io-out-'+devId).value) || 0;

  const curIn  = S.ports.filter(p => p.deviceId === devId && p.dir === 'IN').length;
  const curOut = S.ports.filter(p => p.deviceId === devId && p.dir === 'OUT').length;

  const willRemoveIn  = newIn  < curIn;
  const willRemoveOut = newOut < curOut;

  if (willRemoveIn || willRemoveOut) {
    const msg = `This will remove ${willRemoveIn ? (curIn - newIn) + ' input(s) ' : ''}${willRemoveOut ? (curOut - newOut) + ' output(s)' : ''} from ${devName}.\n\nExisting connections on removed ports will also be deleted.\n\nContinue?`;
    if (!confirm(msg)) return;
  }

  try {
    await Excel.run(async ctx => {
      const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const pu  = psh.getUsedRange(); pu.load('values,rowCount'); await ctx.sync();

      const curInPorts  = S.ports.filter(p => p.deviceId === devId && p.dir === 'IN').sort((a,b) => a.num - b.num);
      const curOutPorts = S.ports.filter(p => p.deviceId === devId && p.dir === 'OUT').sort((a,b) => a.num - b.num);

      // Ports to remove
      const removePortIds = new Set();
      if (newIn < curIn)   curInPorts.slice(newIn).forEach(p  => removePortIds.add(p.id));
      if (newOut < curOut) curOutPorts.slice(newOut).forEach(p => removePortIds.add(p.id));

      // Delete rows in reverse order
      const vals = pu.values;
      for (let i = vals.length - 1; i >= 1; i--) {
        if (removePortIds.has(String(vals[i][COL.PORT.ID]))) {
          psh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
        }
      }
      await ctx.sync();

      // Delete connections for removed ports
      if (removePortIds.size > 0) {
        const csh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
        const cu  = csh.getUsedRange(); cu.load('values'); await ctx.sync();
        const cvals = cu.values;
        for (let i = cvals.length - 1; i >= 1; i--) {
          if (removePortIds.has(String(cvals[i][COL.CONN.SRC_PORT_ID])) ||
              removePortIds.has(String(cvals[i][COL.CONN.DST_PORT_ID]))) {
            csh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
          }
        }
        await ctx.sync();
      }

      // Add new ports
      const rowsToAdd = [];
      if (newIn > curIn) {
        for (let i = curIn + 1; i <= newIn; i++)
          rowsToAdd.push(['P'+(Date.now()+'_I'+i), devId, devName, 'IN', i, '']);
      }
      if (newOut > curOut) {
        for (let i = curOut + 1; i <= newOut; i++)
          rowsToAdd.push(['P'+(Date.now()+1+'_O'+i), devId, devName, 'OUT', i, '']);
      }
      if (rowsToAdd.length) {
        const pu2 = psh.getUsedRange(); pu2.load('rowCount'); await ctx.sync();
        let nr = pu2.rowCount + 1;
        for (const row of rowsToAdd) { psh.getRange(`A${nr}`).getResizedRange(0,5).values=[row]; nr++; }
        await ctx.sync();
      }
    });

    await loadPorts(); await loadConnections();
    renderDeviceList(); renderPortList(); renderMatrix();
    setStatus(`${devName} I/O updated`);
  } catch(e) { setStatus('Error: '+e.message, true); console.error(e); }
}

async function confirmDeleteDevice(id) {
  const dev = S.devices.find(d => d.id === id);
  const portCount = S.ports.filter(p => p.deviceId === id).length;
  const connCount = S.connections.filter(c => {
    const sp = S.ports.find(p => p.id === c.srcId);
    const dp = S.ports.find(p => p.id === c.dstId);
    return (sp && sp.deviceId === id) || (dp && dp.deviceId === id);
  }).length;
  if (!confirm(`Delete "${dev.name}"?\n\nThis will also delete:\n• ${portCount} port(s)\n• ${connCount} connection(s)\n\nThis cannot be undone.`)) return;
  try {
    const portIds = new Set(S.ports.filter(p => p.deviceId === id).map(p => p.id));
    await Excel.run(async ctx => {
      // Delete device row
      const dsh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const du = dsh.getUsedRange(); du.load('values'); await ctx.sync();
      for (let i = du.values.length - 1; i >= 1; i--) {
        if (du.values[i][COL.DEV.ID] === id) { dsh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break; }
      }
      await ctx.sync();
      // Delete port rows
      const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const pu = psh.getUsedRange(); pu.load('values'); await ctx.sync();
      for (let i = pu.values.length - 1; i >= 1; i--) {
        if (portIds.has(String(pu.values[i][COL.PORT.ID]))) psh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
      }
      await ctx.sync();
      // Delete connection rows
      const csh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
      const cu = csh.getUsedRange(); cu.load('values'); await ctx.sync();
      for (let i = cu.values.length - 1; i >= 1; i--) {
        if (portIds.has(String(cu.values[i][COL.CONN.SRC_PORT_ID])) || portIds.has(String(cu.values[i][COL.CONN.DST_PORT_ID])))
          csh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
      }
      await ctx.sync();
    });
    await loadAll(); renderAll();
    setStatus(`"${dev.name}" deleted`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// ============================================================
// PORTS
// ============================================================
function refreshDeviceFilter() {
  const sel = document.getElementById('port-filter-device');
  if (!sel) return;
  const cur = sel.value;
  sel.innerHTML = `<option value="">All devices</option>` +
    S.devices.map(d => `<option value="${d.id}" ${d.id===cur?'selected':''}>${esc(d.name)}</option>`).join('');
}

function renderPortList() {
  const filterDev = v('port-filter-device');
  const filterDir = v('port-filter-dir');
  const q = (v('port-search')||'').toLowerCase();

  let ports = S.ports.filter(p =>
    (!filterDev || p.deviceId === filterDev) &&
    (!filterDir || p.dir === filterDir) &&
    (!q || (p.alias||'').toLowerCase().includes(q) || String(p.num).includes(q) || p.deviceName.toLowerCase().includes(q))
  );

  const el = document.getElementById('port-list');
  if (!ports.length) { el.innerHTML = `<div class="empty">No ports match.</div>`; return; }

  // Group by device
  const byDev = {};
  ports.forEach(p => { if (!byDev[p.deviceId]) byDev[p.deviceId]={name:p.deviceName,ports:[]}; byDev[p.deviceId].ports.push(p); });

  el.innerHTML = Object.values(byDev).map(g => `
    <div style="margin-bottom:12px;">
      <div class="sub-label">${esc(g.name)}</div>
      <div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;">
        ${g.ports.map(p => `
          <div class="port-row">
            <span class="badge ${p.dir==='IN'?'badge-in':'badge-out'}" style="font-size:9px;">${p.dir}</span>
            <span class="port-num">${p.num}</span>
            <input class="port-alias ${p.alias?'':'empty'}" value="${esc(p.alias)}" placeholder="unnamed"
              onblur="savePortAlias('${p.id}', this.value)"
              onfocus="this.classList.remove('empty')"
              onkeydown="if(event.key==='Enter')this.blur()" />
          </div>`).join('')}
      </div>
    </div>`).join('');
}

async function savePortAlias(portId, alias) {
  const port = S.ports.find(p => p.id === portId);
  if (!port || port.alias === alias) return;
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = 1; i < r.values.length; i++) {
        if (r.values[i][COL.PORT.ID] === portId) {
          sh.getRange(`F${i+1}`).values = [[alias]]; break;
        }
      }
      await ctx.sync();
    });
    port.alias = alias;
    renderMatrix(); // refresh matrix labels too
    setStatus(`Port alias saved`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// ============================================================
// MATRIX
// ============================================================
function portLabel(p) {
  return p.alias ? p.alias : `${p.dir} ${p.num}`;
}

function renderMatrix() {
  const container = document.getElementById('matrix-container');
  const countEl   = document.getElementById('matrix-conn-count');

  // Outputs = columns (top), Inputs = rows (left side)
  const devs = S.devices;
  if (!devs.length) { container.innerHTML = `<div class="empty" style="padding:40px">Add devices to see the routing matrix.</div>`; countEl.textContent=''; return; }

  countEl.textContent = `${S.connections.length} connection${S.connections.length!==1?'s':''}`;

  // Build connection lookup: srcPortId -> dstPortId set
  const connMap = {}; // srcId -> Set of dstIds
  const connObjMap = {}; // "srcId|dstId" -> connection obj
  S.connections.forEach(c => {
    if (!connMap[c.srcId]) connMap[c.srcId] = new Set();
    connMap[c.srcId].add(c.dstId);
    connObjMap[c.srcId+'|'+c.dstId] = c;
  });

  // Build table
  // Structure:
  //  - Row 0: corner | col-device headers (spanning their ports, or just 1 cell if collapsed)
  //  - Row 1: empty  | col-port headers (hidden if device collapsed)
  //  - Rows per input device:
  //    - device row: device name | cells (spanning or single)
  //    - port rows: port name | cells (hidden if device collapsed)

  // Collect per-device ports
  const devOuts = {}; // devId -> OUT ports
  const devIns  = {}; // devId -> IN ports
  devs.forEach(d => {
    devOuts[d.id] = S.ports.filter(p => p.deviceId === d.id && p.dir === 'OUT').sort((a,b)=>a.num-b.num);
    devIns[d.id]  = S.ports.filter(p => p.deviceId === d.id && p.dir === 'IN').sort((a,b)=>a.num-b.num);
  });

  // Determine which devices have any ports at all
  const colDevs = devs.filter(d => devOuts[d.id].length > 0); // devices with outputs = columns
  const rowDevs = devs.filter(d => devIns[d.id].length  > 0); // devices with inputs = rows

  if (!colDevs.length || !rowDevs.length) {
    container.innerHTML = `<div class="empty" style="padding:40px">Add ports to devices to see the matrix.</div>`;
    return;
  }

  let html = `<table class="matrix-table" id="matrix-tbl" cellspacing="0">`;

  // ---- HEADER ROW 0: corner + column device headers ----
  html += `<thead><tr>`;
  html += `<th class="m-corner" rowspan="2"><span style="font-size:9px;color:var(--text3)">IN ↓ / OUT →</span></th>`;
  colDevs.forEach(d => {
    const isOpen = !!(matrixState[d.id] && matrixState[d.id].colOpen);
    const count  = devOuts[d.id].length;
    const span   = isOpen ? count : 1;
    html += `<th class="m-col-dev" colspan="${span}" data-dev="${d.id}" onclick="toggleColDevice('${d.id}')" title="${esc(d.name)}">${esc(d.name)} <span style="color:var(--text3);font-weight:400">(${count})</span></th>`;
  });
  html += `</tr>`;

  // ---- HEADER ROW 1: column port headers ----
  html += `<tr>`;
  colDevs.forEach(d => {
    const isOpen = !!(matrixState[d.id] && matrixState[d.id].colOpen);
    if (!isOpen) {
      html += `<th class="m-col-port col-dev-${d.id}-cell" style="font-style:italic;color:var(--text3);font-size:9px;">···</th>`;
    } else {
      devOuts[d.id].forEach(p => {
        html += `<th class="m-col-port col-dev-${d.id}-cell" title="${esc(p.alias||'OUT '+p.num)}">${esc(portLabel(p))}</th>`;
      });
    }
  });
  html += `</tr></thead>`;

  // ---- BODY ROWS ----
  html += `<tbody>`;
  rowDevs.forEach(rowDev => {
    const rowOpen = !!(matrixState[rowDev.id] && matrixState[rowDev.id].rowOpen);
    const rowCount = devIns[rowDev.id].length;

    // Device row
    html += `<tr>`;
    html += `<td class="m-row-dev" rowspan="${rowOpen ? rowCount + 1 : 1}" onclick="toggleRowDevice('${rowDev.id}')" title="${esc(rowDev.name)}">${esc(rowDev.name)} <span style="color:var(--text3);font-size:10px">(${rowCount})</span></td>`;

    // Cells for collapsed row — show aggregate dot if any connection exists
    colDevs.forEach(colDev => {
      const isColOpen = !!(matrixState[colDev.id] && matrixState[colDev.id].colOpen);
      if (!rowOpen && !isColOpen) {
        // Both collapsed: show count of connections between these two devices
        const ins  = devIns[rowDev.id];
        const outs = devOuts[colDev.id];
        let connCount = 0;
        ins.forEach(inp => outs.forEach(outp => { if (connMap[outp.id] && connMap[outp.id].has(inp.id)) connCount++; }));
        html += `<td class="m-cell ${connCount>0?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
          style="${connCount>0?'':'background:var(--bg3);'}"
          title="${connCount} connection(s) between ${esc(rowDev.name)} and ${esc(colDev.name)}"
          onclick="${rowDev.id!==colDev.id ? `toggleDeviceBlock('${rowDev.id}','${colDev.id}')` : ''}">
          ${connCount > 0 ? `<span style="font-size:9px;font-weight:600;color:var(--green)">${connCount}</span>` : ''}
        </td>`;
      } else if (!rowOpen && isColOpen) {
        // Row collapsed, col open: one cell per output port
        devOuts[colDev.id].forEach(outp => {
          const anyConn = devIns[rowDev.id].some(inp => connMap[outp.id] && connMap[outp.id].has(inp.id));
          html += `<td class="m-cell ${anyConn?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
            title="Expand row to route"
            onclick="toggleRowDevice('${rowDev.id}')">
            ${anyConn ? `<span style="font-size:9px;color:var(--green)">·</span>` : ''}
          </td>`;
        });
      } else {
        // Row open, col collapsed: single cell
        if (!isColOpen) {
          const ins  = devIns[rowDev.id];
          const outs = devOuts[colDev.id];
          let connCount = 0;
          ins.forEach(inp => outs.forEach(outp => { if (connMap[outp.id] && connMap[outp.id].has(inp.id)) connCount++; }));
          html += `<td class="m-cell ${connCount>0?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
            title="${connCount} connection(s) — expand column to route"
            onclick="toggleColDevice('${colDev.id}')">
            ${connCount > 0 ? `<span style="font-size:9px;font-weight:600;color:var(--green)">${connCount}</span>` : ''}
          </td>`;
        }
      }
    });
    html += `</tr>`;

    // Port rows (if row open)
    if (rowOpen) {
      devIns[rowDev.id].forEach(inp => {
        html += `<tr>`;
        html += `<td class="m-row-port">${esc(portLabel(inp))}</td>`;
        colDevs.forEach(colDev => {
          const isColOpen = !!(matrixState[colDev.id] && matrixState[colDev.id].colOpen);
          if (!isColOpen) {
            // Col collapsed: single cell, check if any connection from this dev to this inp
            const anyConn = devOuts[colDev.id].some(outp => connMap[outp.id] && connMap[outp.id].has(inp.id));
            html += `<td class="m-cell ${anyConn?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
              title="Expand column to route"
              onclick="toggleColDevice('${colDev.id}')">
              ${anyConn ? `<span style="font-size:9px;color:var(--green)">·</span>` : ''}
            </td>`;
          } else {
            devOuts[colDev.id].forEach(outp => {
              const isConn = connMap[outp.id] && connMap[outp.id].has(inp.id);
              const conn   = isConn ? connObjMap[outp.id+'|'+inp.id] : null;
              const isSame = rowDev.id === colDev.id;
              html += `<td class="m-cell ${isConn?'connected':''} ${isSame?'same-device':''}"
                data-src="${outp.id}" data-dst="${inp.id}"
                data-src-lbl="${esc(portLabel(outp))} (${esc(colDev.name)})"
                data-dst-lbl="${esc(portLabel(inp))} (${esc(rowDev.name)})"
                data-conn-id="${conn?conn.id:''}"
                data-conn-notes="${conn?esc(conn.notes):''}"
                ${!isSame ? `onclick="cellClick(this)" oncontextmenu="cellRightClick(event,this)"` : ''}
                title="${isSame?'Same device':'Click to toggle connection'}">
              </td>`;
            });
          }
        });
        html += `</tr>`;
      });
    }
  });

  html += `</tbody></table>`;
  container.innerHTML = html;
}

function toggleRowDevice(devId) {
  if (!matrixState[devId]) matrixState[devId] = {};
  matrixState[devId].rowOpen = !matrixState[devId].rowOpen;
  renderMatrix();
}

function toggleColDevice(devId) {
  if (!matrixState[devId]) matrixState[devId] = {};
  matrixState[devId].colOpen = !matrixState[devId].colOpen;
  renderMatrix();
}

function toggleDeviceBlock(rowDevId, colDevId) {
  // Expand both to show the connection
  if (!matrixState[rowDevId]) matrixState[rowDevId] = {};
  if (!matrixState[colDevId]) matrixState[colDevId] = {};
  matrixState[rowDevId].rowOpen = true;
  matrixState[colDevId].colOpen = true;
  renderMatrix();
}

function expandAll() {
  S.devices.forEach(d => { matrixState[d.id] = { rowOpen: true, colOpen: true }; });
  renderMatrix();
}

function collapseAll() {
  matrixState = {};
  renderMatrix();
}

async function cellClick(td) {
  const srcId = td.dataset.src;
  const dstId = td.dataset.dst;
  const isConn = td.classList.contains('connected');

  if (isConn) {
    // Remove connection
    const connId = td.dataset.connId;
    await deleteConnectionById(connId);
  } else {
    // Check for duplicate output
    const existingConn = S.connections.find(c => c.srcId === srcId);
    if (existingConn) {
      if (!confirm(`Output "${td.dataset['src-lbl']||srcId}" is already connected to "${existingConn.dstLabel}".\n\nAdd another connection anyway?`)) return;
    }
    await addConnectionDirect(srcId, dstId, td.dataset.srcLbl, td.dataset.dstLbl);
  }
  await loadConnections();
  renderMatrix();
}

async function addConnectionDirect(srcId, dstId, srcLabel, dstLabel) {
  const src = S.ports.find(p => p.id === srcId);
  const dst = S.ports.find(p => p.id === dstId);
  const sl  = src ? `${src.deviceName} / OUT ${src.num}${src.alias?' — '+src.alias:''}` : srcId;
  const dl  = dst ? `${dst.deviceName} / IN ${dst.num}${dst.alias?' — '+dst.alias:''}` : dstId;
  const id  = 'C' + Date.now();
  await Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
    const u = sh.getUsedRange(); u.load('rowCount'); await ctx.sync();
    sh.getRange(`A${u.rowCount+1}`).getResizedRange(0,5).values = [[id,srcId,sl,dstId,dl,'']];
    await ctx.sync();
  });
  setStatus(`Connected: ${sl} → ${dl}`);
}

async function deleteConnectionById(connId) {
  if (!connId) return;
  await Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
    const u = sh.getUsedRange(); u.load('values'); await ctx.sync();
    for (let i = u.values.length - 1; i >= 1; i--) {
      if (u.values[i][COL.CONN.ID] === connId) {
        sh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break;
      }
    }
    await ctx.sync();
  });
  setStatus('Connection removed');
}

function cellRightClick(e, td) {
  e.preventDefault();
  closeCtxMenu();

  const srcId  = td.dataset.src;
  const dstId  = td.dataset.dst;
  const isConn = td.classList.contains('connected');
  const connId = td.dataset.connId;
  const srcLbl = td.dataset.srcLbl || td.getAttribute('data-src-lbl') || srcId;
  const dstLbl = td.dataset.dstLbl || td.getAttribute('data-dst-lbl') || dstId;
  const notes  = td.dataset.connNotes || td.getAttribute('data-conn-notes') || '';

  const menu = document.getElementById('ctx-menu');
  let items = `<div class="ctx-label">${esc(srcLbl)}</div>
    <div class="ctx-label" style="padding-top:0">→ ${esc(dstLbl)}</div>
    <div class="ctx-sep"></div>`;

  if (isConn) {
    items += `
      <div class="ctx-item" onclick="openNoteModal('${connId}','${esc(srcLbl)}','${esc(dstLbl)}','${esc(notes)}'); closeCtxMenu()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        Edit note…
      </div>
      <div class="ctx-item danger" onclick="deleteConnectionById('${connId}').then(()=>loadConnections().then(()=>renderMatrix())); closeCtxMenu()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>
        Remove connection
      </div>`;
  } else {
    items += `
      <div class="ctx-item" onclick="cellClick(document.querySelector('[data-src=\\'${srcId}\\'][data-dst=\\'${dstId}\\']')); closeCtxMenu()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>
        Connect
      </div>`;
  }

  menu.innerHTML = items;
  menu.classList.remove('hidden');
  const x = Math.min(e.clientX, window.innerWidth  - 200);
  const y = Math.min(e.clientY, window.innerHeight - 150);
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  setTimeout(() => document.addEventListener('click', closeCtxMenu, { once: true }), 10);
}

function closeCtxMenu() { document.getElementById('ctx-menu').classList.add('hidden'); }

function openNoteModal(connId, srcLbl, dstLbl, notes) {
  const backdrop = document.getElementById('modal-backdrop');
  const body     = document.getElementById('modal-body');
  body.innerHTML = `
    <h3>Connection Note</h3>
    <div style="font-size:11px;color:var(--text2);margin-bottom:10px;">${esc(srcLbl)} → ${esc(dstLbl)}</div>
    <div class="field"><label>Note</label><input type="text" id="modal-note-input" value="${esc(notes)}" placeholder="e.g. 100ms delay, XLR snake" /></div>
    <div class="modal-actions">
      <button class="btn btn-ghost btn-sm" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary btn-sm" onclick="saveConnectionNote('${connId}')">Save</button>
    </div>`;
  backdrop.classList.remove('hidden');
  setTimeout(() => document.getElementById('modal-note-input').focus(), 50);
}

async function saveConnectionNote(connId) {
  const note = document.getElementById('modal-note-input').value.trim();
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = 1; i < r.values.length; i++) {
        if (r.values[i][COL.CONN.ID] === connId) {
          sh.getRange(`F${i+1}`).values = [[note]]; break;
        }
      }
      await ctx.sync();
    });
    closeModal();
    await loadConnections(); renderMatrix();
    setStatus('Note saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

function closeModal() { document.getElementById('modal-backdrop').classList.add('hidden'); }

// ============================================================
// CONSOLE
// ============================================================
async function addConsoleStrip() {
  const type  = v('con-type');
  const width = v('con-width');
  const name  = v('con-name').trim();

  const id = 'CS' + Date.now();
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const u = sh.getUsedRange(); u.load('rowCount'); await ctx.sync();
      sh.getRange(`A${u.rowCount+1}`).getResizedRange(0,11).values =
        [[id, type, width, name, '', '', '', '', '', '', '', '']];
      await ctx.sync();
    });
    clearFields(['con-name']);
    await loadConsole();
    renderConsole();
    setStatus(`${type} strip added`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

function renderConsole() {
  const el = document.getElementById('console-content');
  if (!S.consoleStrips.length) { el.innerHTML = `<div class="empty">No console strips yet — add one above.</div>`; return; }

  const inPorts  = S.ports.filter(p => p.dir === 'IN');
  const outPorts = S.ports.filter(p => p.dir === 'OUT');

  const sections = ['channel','aux','group','matrix','directout'];
  const secLabels = { channel:'Channels', aux:'Auxes', group:'Groups', matrix:'Matrices', directout:'Direct Outs' };

  let html = '';
  sections.forEach(sec => {
    const strips = S.consoleStrips.filter(s => s.type === sec);
    if (!strips.length) return;
    html += `<div style="margin-bottom:16px;">
      <div class="console-section-hdr">
        <span>${secLabels[sec]} (${strips.length})</span>
        <span style="font-size:9px;color:var(--text2);font-weight:400">click to expand</span>
      </div>`;

    strips.forEach((s, i) => {
      const isChannel = sec === 'channel';
      const numLabel  = i + 1;
      const badge     = s.width === 'stereo'
        ? `<span class="badge badge-stereo" style="font-size:9px;">ST</span>`
        : `<span class="badge badge-mono" style="font-size:9px;">M</span>`;

      const hasData = s.mainIn || s.altIn || s.insASnd || s.output;

      html += `<div class="ch-card">
        <div class="ch-card-hdr" onclick="toggleChCard('${s.id}')">
          <span style="font-size:10px;color:var(--text3);font-family:var(--mono);width:22px;">${numLabel}</span>
          ${badge}
          <span class="card-title" style="font-size:12px;">${s.name || `<span style="color:var(--text3);font-style:italic;font-weight:400">${sec} ${numLabel}</span>`}</span>
          ${hasData ? `<span style="width:6px;height:6px;border-radius:50%;background:var(--green);flex-shrink:0;"></span>` : ''}
          <button class="btn btn-danger btn-xs" onclick="event.stopPropagation();deleteConsoleStrip('${s.id}')">✕</button>
        </div>
        <div class="ch-card-body" id="chbody-${s.id}">
          <div class="field" style="margin-bottom:8px;">
            <label>Name</label>
            <input type="text" value="${esc(s.name)}" placeholder="Strip name"
              onblur="saveConsoleStripName('${s.id}', this.value)"
              onkeydown="if(event.key==='Enter')this.blur()" />
          </div>
          <div class="ch-port-grid">`;

      if (isChannel) {
        html += consolePortField('Main In',     s.mainIn,    inPorts,  s.id, 'mainIn');
        html += consolePortField('Alt In',      s.altIn,     inPorts,  s.id, 'altIn');
        html += consolePortField('InsA Send',   s.insASnd,   outPorts, s.id, 'insASnd');
        html += consolePortField('InsA Return', s.insARet,   inPorts,  s.id, 'insARet');
        html += consolePortField('InsB Send',   s.insBSnd,   outPorts, s.id, 'insBSnd');
        html += consolePortField('InsB Return', s.insBRet,   inPorts,  s.id, 'insBRet');
        html += consolePortField('Direct Out',  s.directOut, outPorts, s.id, 'directOut');
      } else {
        html += consolePortField('Output',      s.output,    outPorts, s.id, 'output');
      }

      html += `</div></div></div>`;
    });
    html += `</div>`;
  });

  el.innerHTML = html;
}

function consolePortField(label, currentVal, ports, stripId, fieldKey) {
  const filled = currentVal && currentVal !== '';
  return `<div class="ch-port-field">
    <label>${label}</label>
    <select class="${filled?'filled':''}" onchange="saveConsolePort('${stripId}','${fieldKey}',this)">
      <option value="">—</option>
      ${ports.map(p => `<option value="${p.id}" ${(currentVal === p.id || currentVal === portLabel(p)) ? 'selected' : ''}>${esc(p.deviceName)} / ${esc(portLabel(p))}</option>`).join('')}
    </select>
  </div>`;
}

function toggleChCard(id) {
  document.getElementById('chbody-'+id).classList.toggle('open');
}

async function saveConsoleStripName(stripId, name) {
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = 1; i < r.values.length; i++) {
        if (r.values[i][COL.CON.ID] === stripId) { sh.getRange(`D${i+1}`).values = [[name]]; break; }
      }
      await ctx.sync();
    });
    const strip = S.consoleStrips.find(s => s.id === stripId);
    if (strip) strip.name = name;
    setStatus('Strip name saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

async function saveConsolePort(stripId, fieldKey, selectEl) {
  const portId = selectEl.value;
  const port   = S.ports.find(p => p.id === portId);
  const label  = port ? portLabel(port) : '';
  const colMap = { mainIn:5, altIn:6, insASnd:7, insARet:8, insBSnd:9, insBRet:10, directOut:11, output:12 };
  const colNum = colMap[fieldKey];
  if (!colNum) return;
  const colLetter = String.fromCharCode(64 + colNum);
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = 1; i < r.values.length; i++) {
        if (r.values[i][COL.CON.ID] === stripId) { sh.getRange(`${colLetter}${i+1}`).values = [[portId]]; break; }
      }
      await ctx.sync();
    });
    selectEl.classList.toggle('filled', !!portId);
    const strip = S.consoleStrips.find(s => s.id === stripId);
    if (strip) strip[fieldKey] = portId;
    setStatus('Console assignment saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

async function deleteConsoleStrip(stripId) {
  if (!confirm('Delete this console strip?')) return;
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = r.values.length - 1; i >= 1; i--) {
        if (r.values[i][COL.CON.ID] === stripId) { sh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break; }
      }
      await ctx.sync();
    });
    await loadConsole(); renderConsole();
    setStatus('Strip deleted');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// ============================================================
// CHECKS
// ============================================================
async function runChecks() {
  await loadAll();
  const issues = [], ok = [];
  const srcCounts = {};
  S.connections.forEach(c => { if (!srcCounts[c.srcId]) srcCounts[c.srcId]=[]; srcCounts[c.srcId].push(c); });
  let dupFound = false;
  Object.values(srcCounts).forEach(conns => {
    if (conns.length > 1) { dupFound = true; issues.push({ type:'conflict', msg:`Output used ${conns.length}×: <strong>${esc(conns[0].srcLabel)}</strong>`, detail: conns.map(c=>'→ '+esc(c.dstLabel)).join(', ') }); }
  });
  if (!dupFound) ok.push('No duplicate outputs');

  let dirOk = true;
  S.connections.forEach(c => {
    const src = S.ports.find(p=>p.id===c.srcId), dst = S.ports.find(p=>p.id===c.dstId);
    if (src && src.dir!=='OUT') { dirOk=false; issues.push({type:'conflict',msg:`Direction mismatch: source is not an output: <strong>${esc(c.srcLabel)}</strong>`}); }
    if (dst && dst.dir!=='IN')  { dirOk=false; issues.push({type:'conflict',msg:`Direction mismatch: destination is not an input: <strong>${esc(c.dstLabel)}</strong>`}); }
  });
  if (dirOk) ok.push('No direction mismatches');

  const devIds = new Set(S.devices.map(d=>d.id));
  const orphans = S.ports.filter(p=>!devIds.has(p.deviceId));
  orphans.length ? issues.push({type:'warn',msg:`${orphans.length} port(s) reference missing devices`}) : ok.push('All ports have valid devices');

  const connectedOuts = new Set(S.connections.map(c=>c.srcId));
  const unconn = S.ports.filter(p=>p.dir==='OUT'&&!connectedOuts.has(p.id));
  if (unconn.length) issues.push({type:'info',msg:`${unconn.length} unconnected output(s)`,detail:unconn.map(p=>esc(portLabel(p))+' ('+esc(p.deviceName)+')').join(', ')});
  else if (S.ports.filter(p=>p.dir==='OUT').length) ok.push('All outputs are connected');

  const el = document.getElementById('checks-content');
  let html = `<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:14px;">
    ${[['Devices',S.devices.length],['Ports',S.ports.length],['Connections',S.connections.length]].map(([l,n])=>`
    <div class="card" style="text-align:center;padding:10px 6px;">
      <div style="font-size:18px;font-weight:600">${n}</div>
      <div style="font-size:10px;color:var(--text2);text-transform:uppercase;letter-spacing:.06em">${l}</div>
    </div>`).join('')}
  </div>`;
  ok.forEach(m => { html+=`<div class="alert alert-success"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>${m}</div>`; });
  if (issues.length) {
    html += `<div class="divider"></div>`;
    issues.forEach(i => {
      const cls = i.type==='conflict'?'alert-error':i.type==='warn'?'alert-warn':'alert-info';
      html+=`<div class="alert ${cls}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12" y2="16.01"/></svg><div><div>${i.msg}</div>${i.detail?`<div style="margin-top:4px;font-size:10px;opacity:.8">${i.detail}</div>`:''}</div></div>`;
    });
  } else if (S.devices.length) {
    html+=`<div class="divider"></div><div class="alert alert-success"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg><strong>All checks passed!</strong></div>`;
  }
  el.innerHTML = html;
  setStatus(`Checks complete — ${issues.length} issue(s)`);
}

// ============================================================
// UI HELPERS
// ============================================================
function switchPanel(name) {
  // Special handling for matrix (flex panel, not in .panels)
  document.querySelectorAll('.panel').forEach(p => { p.classList.remove('active'); p.classList.add('hidden'); });
  document.querySelectorAll('.matrix-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const nav = document.getElementById('nav-'+name);
  if (nav) nav.classList.add('active');
  if (name === 'matrix') {
    document.getElementById('panel-matrix').classList.add('active');
  } else {
    const panel = document.getElementById('panel-'+name);
    if (panel) { panel.classList.remove('hidden'); panel.classList.add('active'); }
  }
}

function setStatus(msg, isError=false) {
  document.getElementById('status-text').textContent = msg;
  document.getElementById('status-dot').className = 'status-dot'+(isError?' error':'');
}

function v(id) { const el=document.getElementById(id); return el?el.value:''; }
function clearFields(ids) { ids.forEach(id=>{const el=document.getElementById(id);if(el)el.value='';}); }
function esc(s) { if(!s&&s!==0) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
