// ============================================================
// AUDIO ROUTING ADD-IN v4 — taskpane.js
// ============================================================
// V4 CHANGES:
//  - Matrix: output port headers rotated CCW (vertical text)
//  - Matrix: row device name sticky + always visible when expanded
//  - Matrix: col device name sticky + always visible when scrolling down
//  - Console: stereo channels store L|R pairs per field
//  - Console: width toggle (M↔ST) on existing strips
//  - Console: searchable port popover replaces plain <select>
//  - Console: 384-input live counter + Checks warning
//  - Console: StripID never shown — channels numbered 1,2,3…
//  - Console: port IDs resolved to readable labels everywhere
// ============================================================

const SHEET = { DEVICES:'Devices', PORTS:'Ports', CONNECTIONS:'Connections', CONSOLE:'Console' };

// Column indices (0-based)
const COL = {
  DEV:  { ID:0, NAME:1, ALIAS:2, NOTES:3 },
  PORT: { ID:0, DEVICE_ID:1, DEVICE_NAME:2, DIR:3, NUM:4, ALIAS:5 },
  CONN: { ID:0, SRC_PORT_ID:1, SRC_LABEL:2, DST_PORT_ID:3, DST_LABEL:4, NOTES:5 },
  // V4: same column count — stereo fields store "portIdL|portIdR" in single cell
  CON:  { ID:0, TYPE:1, WIDTH:2, NAME:3, MAIN_IN:4, ALT_IN:5, INS_A_SND:6, INS_A_RET:7, INS_B_SND:8, INS_B_RET:9, DIRECT_OUT:10, OUTPUT:11 }
};

const MAX_CHANNEL_INPUTS = 384; // mono equivalent

let S = { ready:false, devices:[], ports:[], connections:[], consoleStrips:[] };
let matrixState = {};

// Currently open port popover: { stripId, fieldKey, isStereo, side? }
let _popover = null;

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
        [SHEET.DEVICES]:     ['DeviceID','Name','Alias','Notes'],
        [SHEET.PORTS]:       ['PortID','DeviceID','DeviceName','Direction','PortNum','Name'],
        [SHEET.CONNECTIONS]: ['ConnID','SrcPortID','SrcLabel','DstPortID','DstLabel','Notes'],
        // V4: fields store "portId" (mono) or "portIdL|portIdR" (stereo)
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
      id: r[COL.DEV.ID], name: r[COL.DEV.NAME], alias: r[COL.DEV.ALIAS] || '', notes: r[COL.DEV.NOTES]
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
      name: r[COL.CON.NAME], mainIn: r[COL.CON.MAIN_IN]||'', altIn: r[COL.CON.ALT_IN]||'',
      insASnd: r[COL.CON.INS_A_SND]||'', insARet: r[COL.CON.INS_A_RET]||'',
      insBSnd: r[COL.CON.INS_B_SND]||'', insBRet: r[COL.CON.INS_B_RET]||'',
      directOut: r[COL.CON.DIRECT_OUT]||'', output: r[COL.CON.OUTPUT]||''
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
  const alias   = v('dev-alias').trim();
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
      const dsh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const du = dsh.getUsedRange(); du.load('rowCount'); await ctx.sync();
      dsh.getRange(`A${du.rowCount+1}`).getResizedRange(0, 3).values = [[devId, name, alias, notes]];
      if (portRows.length) {
        const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
        const pu = psh.getUsedRange(); pu.load('rowCount'); await ctx.sync();
        let nr = pu.rowCount + 1;
        for (const row of portRows) { psh.getRange(`A${nr}`).getResizedRange(0, 5).values = [row]; nr++; }
      }
      await ctx.sync();
    });
    clearFields(['dev-name','dev-alias','dev-notes']);
    document.getElementById('dev-inputs').value  = '0';
    document.getElementById('dev-outputs').value = '0';
    await loadDevices(); await loadPorts();
    renderDeviceList(); refreshDeviceFilter(); renderPortList(); renderMatrix();
    setStatus(`"${name}" added with ${inputs} inputs, ${outputs} outputs`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

function renderDeviceList() {
  const el = document.getElementById('device-list');
  if (!S.devices.length) { el.innerHTML = `<div class="empty">No devices yet — add one above.</div>`; return; }
  el.innerHTML = S.devices.map(d => {
    const ins  = S.ports.filter(p => p.deviceId === d.id && p.dir === 'IN').length;
    const outs = S.ports.filter(p => p.deviceId === d.id && p.dir === 'OUT').length;
    return `
    <div class="device-card" id="dcard-${d.id}">
      <div class="device-card-header" onclick="toggleDeviceCard('${d.id}')">
        <svg class="chevron" id="chev-${d.id}" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M9 18l6-6-6-6"/></svg>
        <span class="card-title">${esc(d.name)}${d.alias ? ` <span style="color:var(--text3);font-weight:400;font-size:10px;">[${esc(d.alias)}]</span>` : ''}</span>
        <div class="io-summary">
          ${ins  ? `<span class="io-pill io-pill-in">${ins} IN</span>`   : ''}
          ${outs ? `<span class="io-pill io-pill-out">${outs} OUT</span>` : ''}
        </div>
        <button class="btn btn-danger btn-xs" onclick="event.stopPropagation(); confirmDeleteDevice('${d.id}')">✕</button>
      </div>
      <div class="device-card-body" id="dbody-${d.id}">
        <div class="field" style="margin-bottom:8px;">
          <label>Short Alias <span style="color:var(--text3);font-weight:400">(e.g. OB1, Q, TM)</span></label>
          <input type="text" id="alias-${d.id}" value="${esc(d.alias)}" placeholder="optional short name"
            onblur="saveDeviceAlias('${d.id}', this.value)"
            onkeydown="if(event.key==='Enter')this.blur()" />
        </div>
        <div class="io-row">
          <div class="io-field"><label>Inputs</label><input type="number" id="io-in-${d.id}" value="${ins}" min="0" /></div>
          <div class="io-field"><label>Outputs</label><input type="number" id="io-out-${d.id}" value="${outs}" min="0" /></div>
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
  const body = document.getElementById('dbody-'+id);
  const chev = document.getElementById('chev-'+id);
  const open = body.classList.toggle('open');
  chev.classList.toggle('open', open);
}

async function saveDeviceAlias(devId, alias) {
  const dev = S.devices.find(d => d.id === devId);
  if (!dev || dev.alias === alias) return;
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i = 1; i < r.values.length; i++) {
        if (r.values[i][COL.DEV.ID] === devId) { sh.getRange(`C${i+1}`).values = [[alias]]; break; }
      }
      await ctx.sync();
    });
    dev.alias = alias;
    renderDeviceList(); renderMatrix();
    setStatus(`Alias saved for ${dev.name}`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

async function updateDeviceIO(devId, devName) {
  const newIn  = parseInt(document.getElementById('io-in-'+devId).value)  || 0;
  const newOut = parseInt(document.getElementById('io-out-'+devId).value) || 0;
  const curIn  = S.ports.filter(p => p.deviceId === devId && p.dir === 'IN').length;
  const curOut = S.ports.filter(p => p.deviceId === devId && p.dir === 'OUT').length;
  const willRemoveIn  = newIn  < curIn;
  const willRemoveOut = newOut < curOut;
  if (willRemoveIn || willRemoveOut) {
    const msg = `This will remove ${willRemoveIn ? (curIn-newIn)+' input(s) ' : ''}${willRemoveOut ? (curOut-newOut)+' output(s)' : ''} from ${devName}.\n\nExisting connections on removed ports will also be deleted.\n\nContinue?`;
    if (!confirm(msg)) return;
  }
  try {
    await Excel.run(async ctx => {
      const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const pu  = psh.getUsedRange(); pu.load('values,rowCount'); await ctx.sync();
      const curInPorts  = S.ports.filter(p => p.deviceId===devId && p.dir==='IN').sort((a,b)=>a.num-b.num);
      const curOutPorts = S.ports.filter(p => p.deviceId===devId && p.dir==='OUT').sort((a,b)=>a.num-b.num);
      const removePortIds = new Set();
      if (newIn  < curIn)  curInPorts.slice(newIn).forEach(p  => removePortIds.add(p.id));
      if (newOut < curOut) curOutPorts.slice(newOut).forEach(p => removePortIds.add(p.id));
      const vals = pu.values;
      for (let i = vals.length-1; i >= 1; i--) {
        if (removePortIds.has(String(vals[i][COL.PORT.ID]))) psh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
      }
      await ctx.sync();
      if (removePortIds.size > 0) {
        const csh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
        const cu  = csh.getUsedRange(); cu.load('values'); await ctx.sync();
        for (let i = cu.values.length-1; i >= 1; i--) {
          if (removePortIds.has(String(cu.values[i][COL.CONN.SRC_PORT_ID])) || removePortIds.has(String(cu.values[i][COL.CONN.DST_PORT_ID])))
            csh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
        }
        await ctx.sync();
      }
      const rowsToAdd = [];
      if (newIn  > curIn)  for (let i=curIn+1;  i<=newIn;  i++) rowsToAdd.push(['P'+(Date.now()+'_I'+i),   devId, devName, 'IN',  i, '']);
      if (newOut > curOut) for (let i=curOut+1; i<=newOut; i++) rowsToAdd.push(['P'+(Date.now()+1+'_O'+i), devId, devName, 'OUT', i, '']);
      if (rowsToAdd.length) {
        const pu2 = psh.getUsedRange(); pu2.load('rowCount'); await ctx.sync();
        let nr = pu2.rowCount+1;
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
  const dev = S.devices.find(d => d.id===id);
  const portCount = S.ports.filter(p => p.deviceId===id).length;
  const portIds = new Set(S.ports.filter(p => p.deviceId===id).map(p => p.id));
  const connCount = S.connections.filter(c => {
    const sp = S.ports.find(p => p.id===c.srcId); const dp = S.ports.find(p => p.id===c.dstId);
    return (sp&&sp.deviceId===id)||(dp&&dp.deviceId===id);
  }).length;
  if (!confirm(`Delete "${dev.name}"?\n\nThis will also delete:\n• ${portCount} port(s)\n• ${connCount} connection(s)\n\nThis cannot be undone.`)) return;
  try {
    await Excel.run(async ctx => {
      const dsh = ctx.workbook.worksheets.getItem(SHEET.DEVICES);
      const du = dsh.getUsedRange(); du.load('values'); await ctx.sync();
      for (let i=du.values.length-1; i>=1; i--) { if (du.values[i][COL.DEV.ID]===id) { dsh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break; } }
      await ctx.sync();
      const psh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const pu = psh.getUsedRange(); pu.load('values'); await ctx.sync();
      for (let i=pu.values.length-1; i>=1; i--) { if (portIds.has(String(pu.values[i][COL.PORT.ID]))) psh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); }
      await ctx.sync();
      const csh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
      const cu = csh.getUsedRange(); cu.load('values'); await ctx.sync();
      for (let i=cu.values.length-1; i>=1; i--) {
        if (portIds.has(String(cu.values[i][COL.CONN.SRC_PORT_ID]))||portIds.has(String(cu.values[i][COL.CONN.DST_PORT_ID])))
          csh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up);
      }
      await ctx.sync();
    });
    delete matrixState[id];
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
    S.devices.map(d => `<option value="${d.id}" ${d.id===cur?'selected':''}>${esc(d.name)}${d.alias?' ['+esc(d.alias)+']':''}</option>`).join('');
}

function renderPortList() {
  const filterDev = v('port-filter-device');
  const filterDir = v('port-filter-dir');
  const q = (v('port-search')||'').toLowerCase();
  let ports = S.ports.filter(p =>
    (!filterDev || p.deviceId===filterDev) &&
    (!filterDir || p.dir===filterDir) &&
    (!q || (p.alias||'').toLowerCase().includes(q) || String(p.num).includes(q) || p.deviceName.toLowerCase().includes(q))
  );
  const el = document.getElementById('port-list');
  if (!ports.length) { el.innerHTML = `<div class="empty">No ports match.</div>`; return; }
  const byDev = {};
  ports.forEach(p => { if (!byDev[p.deviceId]) byDev[p.deviceId]={name:p.deviceName,ports:[]}; byDev[p.deviceId].ports.push(p); });
  el.innerHTML = Object.values(byDev).map(g => `
    <div style="margin-bottom:12px;">
      <div class="sub-label">${esc(g.name)}</div>
      <div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;">
        ${g.ports.map(p => `
          <div class="port-row">
            <span class="badge ${p.dir==='IN'?'badge-in':'badge-out'}" style="font-size:9px;flex-shrink:0;">${p.dir}</span>
            <span class="port-num" style="flex-shrink:0;">${p.num}</span>
            <input class="port-alias" value="${esc(p.alias)}" placeholder="port name…"
              onblur="savePortAlias('${p.id}', this.value)"
              onkeydown="if(event.key==='Enter')this.blur()" />
          </div>`).join('')}
      </div>
    </div>`).join('');
}

async function savePortAlias(portId, alias) {
  const port = S.ports.find(p => p.id===portId);
  if (!port || port.alias===alias) return;
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.PORTS);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=1; i<r.values.length; i++) {
        if (r.values[i][COL.PORT.ID]===portId) { sh.getRange(`F${i+1}`).values=[[alias]]; break; }
      }
      await ctx.sync();
    });
    port.alias = alias;
    renderMatrix();
    setStatus('Port name saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// ============================================================
// MATRIX
// ============================================================
function portLabel(p) {
  if (p.alias) return p.alias;
  return `${p.dir==='IN'?'IN':'OUT'} ${p.num}`;
}

function portFullLabel(p) {
  const dev = S.devices.find(d => d.id===p.deviceId);
  const devLabel = (dev && dev.alias) ? dev.alias : p.deviceName;
  const portName = p.alias ? p.alias : `${p.dir} ${p.num}`;
  return `${devLabel} · ${portName}`;
}

function renderMatrix() {
  const container = document.getElementById('matrix-container');
  const countEl   = document.getElementById('matrix-conn-count');
  const devs = S.devices;
  if (!devs.length) { container.innerHTML = `<div class="empty" style="padding:40px">Add devices to see the routing matrix.</div>`; countEl.textContent=''; return; }
  countEl.textContent = `${S.connections.length} connection${S.connections.length!==1?'s':''}`;

  const connMap = {};
  const connObjMap = {};
  S.connections.forEach(c => {
    if (!connMap[c.srcId]) connMap[c.srcId] = new Set();
    connMap[c.srcId].add(c.dstId);
    connObjMap[c.srcId+'|'+c.dstId] = c;
  });

  const devOuts = {};
  const devIns  = {};
  devs.forEach(d => {
    devOuts[d.id] = S.ports.filter(p => p.deviceId===d.id && p.dir==='OUT').sort((a,b)=>a.num-b.num);
    devIns[d.id]  = S.ports.filter(p => p.deviceId===d.id && p.dir==='IN').sort((a,b)=>a.num-b.num);
  });

  const colDevs = devs.filter(d => devOuts[d.id].length > 0);
  const rowDevs = devs.filter(d => devIns[d.id].length  > 0);

  if (!colDevs.length || !rowDevs.length) {
    container.innerHTML = `<div class="empty" style="padding:40px">Add ports to devices to see the matrix.</div>`; return;
  }

  function devLabel(d) {
    return d.alias
      ? `${esc(d.alias)} <span style="color:var(--text3);font-weight:400;font-size:9px">${esc(d.name)}</span>`
      : esc(d.name);
  }

  let html = `<table class="matrix-table" id="matrix-tbl" cellspacing="0">`;

  // ---- HEADER ROW 0: corner + column device headers ----
  // V4: corner spans 3 rows (device row + port-label row + potential spacer)
  html += `<thead><tr>`;
  html += `<th class="m-corner" rowspan="2"><span style="font-size:9px;color:var(--text3)">IN ↓ / OUT →</span></th>`;
  colDevs.forEach(d => {
    const isOpen = !!(matrixState[d.id] && matrixState[d.id].colOpen);
    const count  = devOuts[d.id].length;
    const span   = isOpen ? count : 1;
    // V4: col device header is sticky top:0 (already was), unchanged
    html += `<th class="m-col-dev" colspan="${span}" data-dev="${d.id}" onclick="toggleColDevice('${d.id}')" title="${esc(d.name)}">${devLabel(d)} <span style="color:var(--text3);font-weight:400">(${count})</span></th>`;
  });
  html += `</tr>`;

  // ---- HEADER ROW 1: column port headers (rotated CCW) ----
  html += `<tr>`;
  colDevs.forEach(d => {
    const isOpen = !!(matrixState[d.id] && matrixState[d.id].colOpen);
    if (!isOpen) {
      html += `<th class="m-col-port" style="font-style:italic;color:var(--text3);font-size:9px;">···</th>`;
    } else {
      devOuts[d.id].forEach(p => {
        // V4: writing-mode + rotation handled by .m-col-port CSS (CCW)
        html += `<th class="m-col-port" title="${esc(portFullLabel(p))}">${esc(portLabel(p))}</th>`;
      });
    }
  });
  html += `</tr></thead>`;

  // ---- BODY ROWS ----
  html += `<tbody>`;
  rowDevs.forEach(rowDev => {
    const rowOpen  = !!(matrixState[rowDev.id] && matrixState[rowDev.id].rowOpen);
    const rowCount = devIns[rowDev.id].length;

    // V4: device row — when expanded, device name stays as first row with rowspan
    // The sticky left cell shows device name always (rowspan covers all port rows)
    html += `<tr class="m-dev-row">`;
    html += `<td class="m-row-dev" rowspan="${rowOpen ? rowCount+1 : 1}" onclick="toggleRowDevice('${rowDev.id}')" title="${esc(rowDev.name)}">
      <span class="m-row-dev-label">${devLabel(rowDev)}</span>
      <span style="color:var(--text3);font-size:10px;margin-left:4px">(${rowCount})</span>
      <svg style="flex-shrink:0;margin-left:auto;transition:transform .2s;transform:${rowOpen?'rotate(90deg)':'rotate(0deg)'}" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M9 18l6-6-6-6"/></svg>
    </td>`;

    colDevs.forEach(colDev => {
      const isColOpen = !!(matrixState[colDev.id] && matrixState[colDev.id].colOpen);
      if (!rowOpen && !isColOpen) {
        const ins=devIns[rowDev.id], outs=devOuts[colDev.id];
        let connCount=0;
        ins.forEach(inp => outs.forEach(outp => { if (connMap[outp.id]&&connMap[outp.id].has(inp.id)) connCount++; }));
        html += `<td class="m-cell ${connCount>0?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
          title="${connCount} connection(s) between ${esc(rowDev.name)} and ${esc(colDev.name)}"
          onclick="${rowDev.id!==colDev.id?`toggleDeviceBlock('${rowDev.id}','${colDev.id}')`:''}" >
          ${connCount>0?`<span style="font-size:9px;font-weight:600;color:var(--green)">${connCount}</span>`:''}
        </td>`;
      } else if (!rowOpen && isColOpen) {
        devOuts[colDev.id].forEach(outp => {
          const anyConn = devIns[rowDev.id].some(inp => connMap[outp.id]&&connMap[outp.id].has(inp.id));
          html += `<td class="m-cell ${anyConn?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
            title="Expand row to route" onclick="toggleRowDevice('${rowDev.id}')">
            ${anyConn?`<span style="font-size:9px;color:var(--green)">·</span>`:''}
          </td>`;
        });
      } else if (rowOpen && !isColOpen) {
        const ins=devIns[rowDev.id], outs=devOuts[colDev.id];
        let connCount=0;
        ins.forEach(inp => outs.forEach(outp => { if (connMap[outp.id]&&connMap[outp.id].has(inp.id)) connCount++; }));
        html += `<td class="m-cell ${connCount>0?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
          title="${connCount} connection(s) — expand column to route" onclick="toggleColDevice('${colDev.id}')">
          ${connCount>0?`<span style="font-size:9px;font-weight:600;color:var(--green)">${connCount}</span>`:''}
        </td>`;
      }
      // rowOpen && isColOpen handled in port rows below
    });
    html += `</tr>`;

    if (rowOpen) {
      devIns[rowDev.id].forEach(inp => {
        html += `<tr>`;
        // V4: port-level row — NO sticky left cell here; device name covered by rowspan above
        html += `<td class="m-row-port" title="${esc(portFullLabel(inp))}">${esc(portLabel(inp))}</td>`;
        colDevs.forEach(colDev => {
          const isColOpen = !!(matrixState[colDev.id] && matrixState[colDev.id].colOpen);
          if (!isColOpen) {
            const anyConn = devOuts[colDev.id].some(outp => connMap[outp.id]&&connMap[outp.id].has(inp.id));
            html += `<td class="m-cell ${anyConn?'connected':''} ${rowDev.id===colDev.id?'same-device':''}"
              title="Expand column to route" onclick="toggleColDevice('${colDev.id}')">
              ${anyConn?`<span style="font-size:9px;color:var(--green)">·</span>`:''}
            </td>`;
          } else {
            devOuts[colDev.id].forEach(outp => {
              const isConn = connMap[outp.id] && connMap[outp.id].has(inp.id);
              const conn   = isConn ? connObjMap[outp.id+'|'+inp.id] : null;
              const isSame = rowDev.id===colDev.id;
              const srcLbl = portFullLabel(outp);
              const dstLbl = portFullLabel(inp);
              html += `<td class="m-cell ${isConn?'connected':''} ${isSame?'same-device':''}"
                data-src="${outp.id}" data-dst="${inp.id}"
                data-src-lbl="${esc(srcLbl)}" data-dst-lbl="${esc(dstLbl)}"
                data-conn-id="${conn?conn.id:''}" data-conn-notes="${conn?esc(conn.notes):''}"
                ${!isSame?`onclick="cellClick(this)" oncontextmenu="cellRightClick(event,this)"`:''}
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
  if (!matrixState[rowDevId]) matrixState[rowDevId] = {};
  if (!matrixState[colDevId]) matrixState[colDevId] = {};
  matrixState[rowDevId].rowOpen = true;
  matrixState[colDevId].colOpen = true;
  renderMatrix();
}
function expandAll()  { S.devices.forEach(d => { matrixState[d.id] = { rowOpen:true, colOpen:true }; }); renderMatrix(); }
function collapseAll(){ matrixState = {}; renderMatrix(); }

async function cellClick(td) {
  const srcId = td.dataset.src, dstId = td.dataset.dst;
  const isConn = td.classList.contains('connected');
  if (isConn) {
    await deleteConnectionById(td.dataset.connId);
  } else {
    const existingConn = S.connections.find(c => c.srcId===srcId);
    if (existingConn) {
      const srcPort = S.ports.find(p => p.id===srcId);
      const dstPort = S.ports.find(p => p.id===existingConn.dstId);
      if (!confirm(`Output "${srcPort?portFullLabel(srcPort):srcId}" is already connected to "${dstPort?portFullLabel(dstPort):existingConn.dstLabel}".\n\nAdd another connection anyway?`)) return;
    }
    await addConnectionDirect(srcId, dstId);
  }
  await loadConnections(); renderMatrix();
}

async function addConnectionDirect(srcId, dstId) {
  const src = S.ports.find(p => p.id===srcId), dst = S.ports.find(p => p.id===dstId);
  const sl = src ? portFullLabel(src) : srcId;
  const dl = dst ? portFullLabel(dst) : dstId;
  const id = 'C' + Date.now();
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
    for (let i=u.values.length-1; i>=1; i--) {
      if (u.values[i][COL.CONN.ID]===connId) { sh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break; }
    }
    await ctx.sync();
  });
  setStatus('Connection removed');
}

function cellRightClick(e, td) {
  e.preventDefault(); closeCtxMenu();
  const srcId=td.dataset.src, dstId=td.dataset.dst;
  const isConn=td.classList.contains('connected'), connId=td.dataset.connId;
  const srcPort=S.ports.find(p=>p.id===srcId), dstPort=S.ports.find(p=>p.id===dstId);
  const srcLbl = srcPort ? portFullLabel(srcPort) : (td.getAttribute('data-src-lbl')||srcId);
  const dstLbl = dstPort ? portFullLabel(dstPort) : (td.getAttribute('data-dst-lbl')||dstId);
  const notes  = td.getAttribute('data-conn-notes')||'';
  const menu = document.getElementById('ctx-menu');
  let items = `<div class="ctx-label">${esc(srcLbl)}</div><div class="ctx-label" style="padding-top:0">→ ${esc(dstLbl)}</div><div class="ctx-sep"></div>`;
  if (isConn) {
    items += `
      <div class="ctx-item" onclick="openNoteModal('${connId}','${esc(srcLbl)}','${esc(dstLbl)}','${esc(notes)}');closeCtxMenu()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>Edit note…</div>
      <div class="ctx-item danger" onclick="deleteConnectionById('${connId}').then(()=>loadConnections().then(()=>renderMatrix()));closeCtxMenu()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>Remove connection</div>`;
  } else {
    items += `<div class="ctx-item" onclick="cellClick(document.querySelector('[data-src=\\'${srcId}\\'][data-dst=\\'${dstId}\\']'));closeCtxMenu()">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/></svg>Connect</div>`;
  }
  menu.innerHTML = items;
  menu.classList.remove('hidden');
  menu.style.left = Math.min(e.clientX, window.innerWidth-200)+'px';
  menu.style.top  = Math.min(e.clientY, window.innerHeight-150)+'px';
  setTimeout(() => document.addEventListener('click', closeCtxMenu, {once:true}), 10);
}
function closeCtxMenu() { document.getElementById('ctx-menu').classList.add('hidden'); }

function openNoteModal(connId, srcLbl, dstLbl, notes) {
  const body = document.getElementById('modal-body');
  body.innerHTML = `
    <h3>Connection Note</h3>
    <div style="font-size:11px;color:var(--text2);margin-bottom:10px;">${esc(srcLbl)} → ${esc(dstLbl)}</div>
    <div class="field"><label>Note</label><input type="text" id="modal-note-input" value="${esc(notes)}" placeholder="e.g. 100ms delay, XLR snake" /></div>
    <div class="modal-actions">
      <button class="btn btn-ghost btn-sm" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary btn-sm" onclick="saveConnectionNote('${connId}')">Save</button>
    </div>`;
  document.getElementById('modal-backdrop').classList.remove('hidden');
  setTimeout(() => document.getElementById('modal-note-input').focus(), 50);
}

async function saveConnectionNote(connId) {
  const note = document.getElementById('modal-note-input').value.trim();
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONNECTIONS);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=1; i<r.values.length; i++) {
        if (r.values[i][COL.CONN.ID]===connId) { sh.getRange(`F${i+1}`).values=[[note]]; break; }
      }
      await ctx.sync();
    });
    closeModal(); await loadConnections(); renderMatrix(); setStatus('Note saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}
function closeModal() { document.getElementById('modal-backdrop').classList.add('hidden'); }

// ============================================================
// CONSOLE — V4
// ============================================================

// ---- Stereo helpers ----
// Stored value: "portId" (mono) OR "portIdL|portIdR" (stereo)
function splitPair(val) {
  if (!val) return ['',''];
  const parts = String(val).split('|');
  return [parts[0]||'', parts[1]||''];
}
function joinPair(l, r) {
  if (!l && !r) return '';
  if (!r) return l;
  return `${l}|${r}`;
}
function pairLabel(val, isStereo) {
  if (!val) return '—';
  if (!isStereo) {
    const p = S.ports.find(x => x.id===val);
    return p ? portFullLabel(p) : val;
  }
  const [l,r] = splitPair(val);
  const lp = l ? S.ports.find(x=>x.id===l) : null;
  const rp = r ? S.ports.find(x=>x.id===r) : null;
  if (lp && rp) return `${portFullLabel(lp)} + ${portFullLabel(rp)}`;
  if (lp) return `${portFullLabel(lp)} + ?`;
  if (rp) return `? + ${portFullLabel(rp)}`;
  return '—';
}

// ---- 384-input counter ----
function calcChannelInputs() {
  // Each channel counts: mono=1, stereo=2
  return S.consoleStrips
    .filter(s => s.type==='channel')
    .reduce((sum, s) => sum + (s.width==='stereo' ? 2 : 1), 0);
}
function renderChannelCounter() {
  const el = document.getElementById('con-input-counter');
  if (!el) return;
  const used = calcChannelInputs();
  const pct  = Math.min(used / MAX_CHANNEL_INPUTS * 100, 100);
  const over = used > MAX_CHANNEL_INPUTS;
  el.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
      <span style="font-size:10px;color:${over?'var(--red)':'var(--text2)'};">
        Channel inputs: <strong style="color:${over?'var(--red)':used/MAX_CHANNEL_INPUTS>0.9?'var(--yellow)':'var(--text)'}">${used}</strong> / ${MAX_CHANNEL_INPUTS}
      </span>
      ${over?`<span style="font-size:10px;color:var(--red);font-weight:600;">⚠ OVER LIMIT</span>`:''}
    </div>
    <div style="height:4px;background:var(--bg3);border-radius:2px;overflow:hidden;">
      <div style="height:100%;width:${pct}%;background:${over?'var(--red)':pct>90?'var(--yellow)':'var(--green)'};border-radius:2px;transition:width .3s;"></div>
    </div>`;
}

async function addConsoleStrip() {
  const type  = v('con-type');
  const width = v('con-width');
  const name  = v('con-name').trim();
  const count = Math.max(1, Math.min(384, parseInt(v('con-count'))||1));

  // Pre-check channel limit
  if (type === 'channel') {
    const addedInputs = (width==='stereo' ? 2 : 1) * count;
    const current = calcChannelInputs();
    if (current + addedInputs > MAX_CHANNEL_INPUTS) {
      if (!confirm(`Adding ${count} ${width} channel${count>1?'s':''} would bring you to ${current+addedInputs}/${MAX_CHANNEL_INPUTS} inputs — over the limit.\n\nAdd anyway?`)) return;
    }
  }

  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const u = sh.getUsedRange(); u.load('rowCount'); await ctx.sync();
      let nr = u.rowCount + 1;
      for (let i = 0; i < count; i++) {
        const id = 'CS' + (Date.now() + i);
        // If bulk add with a name, append a number suffix: "Kick 1", "Kick 2"…
        const stripName = count > 1 ? (name ? `${name} ${i+1}` : '') : name;
        sh.getRange(`A${nr}`).getResizedRange(0,11).values =
          [[id, type, width, stripName, '', '', '', '', '', '', '', '']];
        nr++;
      }
      await ctx.sync();
    });
    clearFields(['con-name']);
    document.getElementById('con-count').value = '1';
    await loadConsole();
    renderConsole();
    setStatus(`${count} ${width} ${type} strip${count>1?'s':''} added`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// Toggle mono/stereo on an existing strip
async function toggleStripWidth(stripId) {
  const strip = S.consoleStrips.find(s => s.id===stripId);
  if (!strip) return;
  const newWidth = strip.width==='stereo' ? 'mono' : 'stereo';
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=1; i<r.values.length; i++) {
        if (r.values[i][COL.CON.ID]===stripId) { sh.getRange(`C${i+1}`).values=[[newWidth]]; break; }
      }
      await ctx.sync();
    });
    strip.width = newWidth;
    renderConsole();
    renderChannelCounter();
    setStatus(`Strip width set to ${newWidth}`);
  } catch(e) { setStatus('Error: '+e.message, true); }
}

// ---- Port picker modal dialog ----
// State for what's currently being edited
let _pickerState = null; // { stripId, fieldKey, isStereo }

function openPortPopover(stripId, fieldKey, isStereo) {
  closePortPopover();
  const strip = S.consoleStrips.find(s => s.id===stripId);
  if (!strip) return;

  const isInput = ['mainIn','altIn','insARet','insBRet'].includes(fieldKey);
  const allPorts = S.ports.filter(p => p.dir===(isInput?'IN':'OUT'));
  const currentVal = strip[fieldKey]||'';
  const [curL, curR] = splitPair(currentVal);

  _pickerState = { stripId, fieldKey, isStereo };

  const fieldNames = { mainIn:'Main In', altIn:'Alt In', insASnd:'InsA Send', insARet:'InsA Return', insBSnd:'InsB Send', insBRet:'InsB Return', directOut:'Direct Out', output:'Output' };
  const title = `${fieldNames[fieldKey]||fieldKey}${isStereo?' — Stereo':''}`;

  function sideHtml(side, curVal, label) {
    return `
      <div class="pop-side">
        ${isStereo ? `<div class="pop-side-label">${label}</div>` : ''}
        <div class="pop-search-wrap">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/></svg>
          <input type="text" class="pop-search" id="pop-search-${side}" placeholder="Search ports…" autocomplete="off"
            oninput="filterPopoverList('${side}')" />
        </div>
        <div class="pop-list" id="pop-list-${side}">
          <div class="pop-item ${!curVal?'selected':''}" onclick="selectPopoverPort('${stripId}','${fieldKey}','${side}','')">
            <span style="color:var(--text3)">— none —</span>
          </div>
          ${allPorts.map(p => `
            <div class="pop-item ${p.id===curVal?'selected':''}" data-label="${esc(portFullLabel(p)).toLowerCase()}"
              onclick="selectPopoverPort('${stripId}','${fieldKey}','${side}','${p.id}')">
              <span class="badge ${p.dir==='IN'?'badge-in':'badge-out'}" style="font-size:8px;">${p.dir}</span>
              ${esc(portFullLabel(p))}
            </div>`).join('')}
        </div>
      </div>`;
  }

  const backdrop = document.createElement('div');
  backdrop.className = 'port-modal-backdrop';
  backdrop.id = 'port-modal-backdrop';
  backdrop.onclick = (e) => { if (e.target===backdrop) closePortPopover(); };

  backdrop.innerHTML = `
    <div class="port-modal">
      <div class="port-modal-header">
        <h4>${esc(title)}</h4>
        <button class="btn btn-ghost btn-xs" onclick="closePortPopover()">✕ Close</button>
      </div>
      <div class="port-modal-body${isStereo?' stereo':''}">
        ${isStereo
          ? `${sideHtml('L', curL, 'Left / Ch 1')}<div class="pop-divider"></div>${sideHtml('R', curR, 'Right / Ch 2')}`
          : sideHtml('L', currentVal, '')}
      </div>
    </div>`;

  document.body.appendChild(backdrop);

  setTimeout(() => {
    const inp = document.getElementById('pop-search-L');
    if (inp) inp.focus();
  }, 30);
}

function closePortPopover() {
  const el = document.getElementById('port-modal-backdrop');
  if (el) el.remove();
  _pickerState = null;
}

function filterPopoverList(side) {
  const q = (document.getElementById(`pop-search-${side}`)?.value||'').toLowerCase();
  const list = document.getElementById(`pop-list-${side}`);
  if (!list) return;
  list.querySelectorAll('.pop-item').forEach(item => {
    item.style.display = (!q || (item.dataset.label||'').includes(q)) ? '' : 'none';
  });
}

async function selectPopoverPort(stripId, fieldKey, side, portId) {
  const strip = S.consoleStrips.find(s => s.id===stripId);
  if (!strip) return;
  const isStereo = strip.width==='stereo';
  let newVal;
  if (!isStereo) {
    newVal = portId;
  } else {
    const [curL, curR] = splitPair(strip[fieldKey]||'');
    newVal = side==='L' ? joinPair(portId, curR) : joinPair(curL, portId);
  }
  const colMap = { mainIn:5, altIn:6, insASnd:7, insARet:8, insBSnd:9, insBRet:10, directOut:11, output:12 };
  const colNum = colMap[fieldKey];
  if (!colNum) return;
  const colLetter = String.fromCharCode(64 + colNum);
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=1; i<r.values.length; i++) {
        if (r.values[i][COL.CON.ID]===stripId) { sh.getRange(`${colLetter}${i+1}`).values=[[newVal]]; break; }
      }
      await ctx.sync();
    });
    strip[fieldKey] = newVal;
    // Update the button label in the table without full re-render
    const btn = document.querySelector(`[data-strip="${stripId}"][data-field="${fieldKey}"]`);
    if (btn) { btn.textContent = pairLabel(newVal, isStereo)||'—'; btn.classList.toggle('filled', !!newVal); }
    setStatus('Port assignment saved');
    // For stereo: re-open modal with fresh selection state so both sides stay editable
    if (isStereo) {
      closePortPopover();
      openPortPopover(stripId, fieldKey, true);
    } else {
      closePortPopover();
    }
  } catch(e) { setStatus('Error: '+e.message, true); }
}

function renderConsole() {
  const el = document.getElementById('console-content');
  renderChannelCounter();
  if (!S.consoleStrips.length) { el.innerHTML = `<div class="empty">No console strips yet — add one above.</div>`; return; }

  const inPorts  = S.ports.filter(p => p.dir==='IN');
  const outPorts = S.ports.filter(p => p.dir==='OUT');

  const sections = ['channel','aux','group','matrix','directout'];
  const secLabels = { channel:'Channels', aux:'Auxes', group:'Groups', matrix:'Matrices', directout:'Direct Outs' };

  let html = '';
  sections.forEach(sec => {
    const strips = S.consoleStrips.filter(s => s.type===sec);
    if (!strips.length) return;
    const isChannel = sec==='channel';

    html += `<div style="margin-bottom:20px;">
      <div class="console-section-hdr">${secLabels[sec]} <span style="color:var(--text3);font-weight:400;font-size:10px;">(${strips.length})</span></div>
      <div style="overflow-x:auto;">
        <table class="con-table">
          <thead><tr>
            <th style="width:28px;">#</th>
            <th style="width:40px;">Width</th>
            <th>Name</th>
            <th>Main In</th>`;

    if (isChannel) {
      html += `<th>Alt In</th><th>InsA Snd</th><th>InsA Ret</th><th>InsB Snd</th><th>InsB Ret</th><th>Direct Out</th>`;
    } else {
      html += `<th>Output</th>`;
    }
    html += `<th style="width:26px;"></th></tr></thead><tbody>`;

    strips.forEach((s, i) => {
      const num = i+1;
      const isStereo = s.width==='stereo';
      const widthBtn = `<button class="width-toggle ${isStereo?'stereo':'mono'}" onclick="toggleStripWidth('${s.id}')" title="Click to toggle mono/stereo">${isStereo?'ST':'M'}</button>`;

      function portBtn(fieldKey, val) {
        const filled = !!val;
        const label  = pairLabel(val, isStereo);
        return `<button class="con-port-btn ${filled?'filled':''}" data-strip="${s.id}" data-field="${fieldKey}"
          onclick="openPortPopover('${s.id}','${fieldKey}',${isStereo})">${esc(label)}</button>`;
      }

      html += `<tr class="con-row">
        <td class="con-num">${num}</td>
        <td class="con-w">${widthBtn}</td>
        <td class="con-name-cell">
          <input class="con-name-input" value="${esc(s.name)}" placeholder="name…"
            onblur="saveConsoleStripName('${s.id}', this.value)"
            onkeydown="if(event.key==='Enter')this.blur()" />
        </td>
        <td class="con-port-cell">${portBtn('mainIn', s.mainIn)}</td>`;

      if (isChannel) {
        html += `
        <td class="con-port-cell">${portBtn('altIn',    s.altIn)}</td>
        <td class="con-port-cell">${portBtn('insASnd',  s.insASnd)}</td>
        <td class="con-port-cell">${portBtn('insARet',  s.insARet)}</td>
        <td class="con-port-cell">${portBtn('insBSnd',  s.insBSnd)}</td>
        <td class="con-port-cell">${portBtn('insBRet',  s.insBRet)}</td>
        <td class="con-port-cell">${portBtn('directOut',s.directOut)}</td>`;
      } else {
        html += `<td class="con-port-cell">${portBtn('output', s.output)}</td>`;
      }

      html += `<td class="con-del-cell">
          <button class="btn btn-danger btn-xs" onclick="deleteConsoleStrip('${s.id}')">✕</button>
        </td></tr>`;
    });
    html += `</tbody></table></div></div>`;
  });

  el.innerHTML = html;
}

async function saveConsoleStripName(stripId, name) {
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=1; i<r.values.length; i++) {
        if (r.values[i][COL.CON.ID]===stripId) { sh.getRange(`D${i+1}`).values=[[name]]; break; }
      }
      await ctx.sync();
    });
    const strip = S.consoleStrips.find(s => s.id===stripId);
    if (strip) strip.name = name;
    setStatus('Strip name saved');
  } catch(e) { setStatus('Error: '+e.message, true); }
}

async function deleteConsoleStrip(stripId) {
  if (!confirm('Delete this console strip?')) return;
  try {
    await Excel.run(async ctx => {
      const sh = ctx.workbook.worksheets.getItem(SHEET.CONSOLE);
      const r = sh.getUsedRange(); r.load('values'); await ctx.sync();
      for (let i=r.values.length-1; i>=1; i--) {
        if (r.values[i][COL.CON.ID]===stripId) { sh.getRange(`${i+1}:${i+1}`).delete(Excel.DeleteShiftDirection.up); break; }
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

  // Duplicate outputs
  const srcCounts = {};
  S.connections.forEach(c => { if (!srcCounts[c.srcId]) srcCounts[c.srcId]=[]; srcCounts[c.srcId].push(c); });
  let dupFound = false;
  Object.values(srcCounts).forEach(conns => {
    if (conns.length > 1) {
      dupFound = true;
      const srcPort = S.ports.find(p=>p.id===conns[0].srcId);
      issues.push({ type:'conflict', msg:`Output used ${conns.length}×: <strong>${esc(srcPort?portFullLabel(srcPort):conns[0].srcLabel)}</strong>`,
        detail: conns.map(c => { const dp=S.ports.find(p=>p.id===c.dstId); return '→ '+esc(dp?portFullLabel(dp):c.dstLabel); }).join(', ') });
    }
  });
  if (!dupFound) ok.push('No duplicate outputs');

  // Direction mismatches
  let dirOk = true;
  S.connections.forEach(c => {
    const src=S.ports.find(p=>p.id===c.srcId), dst=S.ports.find(p=>p.id===c.dstId);
    if (src&&src.dir!=='OUT') { dirOk=false; issues.push({type:'conflict',msg:`Direction mismatch: source not an output: <strong>${esc(portFullLabel(src))}</strong>`}); }
    if (dst&&dst.dir!=='IN')  { dirOk=false; issues.push({type:'conflict',msg:`Direction mismatch: destination not an input: <strong>${esc(portFullLabel(dst))}</strong>`}); }
  });
  if (dirOk) ok.push('No direction mismatches');

  // Orphaned ports
  const devIds = new Set(S.devices.map(d=>d.id));
  const orphans = S.ports.filter(p=>!devIds.has(p.deviceId));
  orphans.length ? issues.push({type:'warn',msg:`${orphans.length} port(s) reference missing devices`}) : ok.push('All ports have valid devices');

  // Unconnected outputs
  const connectedOuts = new Set(S.connections.map(c=>c.srcId));
  const unconn = S.ports.filter(p=>p.dir==='OUT'&&!connectedOuts.has(p.id));
  if (unconn.length) issues.push({type:'info',msg:`${unconn.length} unconnected output(s)`,detail:unconn.map(p=>esc(portFullLabel(p))).join(', ')});
  else if (S.ports.filter(p=>p.dir==='OUT').length) ok.push('All outputs are connected');

  // V4: 384-input check
  const usedInputs = calcChannelInputs();
  if (usedInputs > MAX_CHANNEL_INPUTS) {
    issues.push({type:'conflict', msg:`Channel input count <strong>${usedInputs}</strong> exceeds the ${MAX_CHANNEL_INPUTS}-input limit`,
      detail:`Reduce channel count or convert stereo channels to mono to stay within limit.`});
  } else if (usedInputs > MAX_CHANNEL_INPUTS * 0.9) {
    issues.push({type:'warn', msg:`Channel inputs at ${usedInputs}/${MAX_CHANNEL_INPUTS} — approaching limit`});
  } else {
    ok.push(`Channel inputs: ${usedInputs}/${MAX_CHANNEL_INPUTS}`);
  }

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
  document.querySelectorAll('.panel').forEach(p => { p.classList.remove('active'); p.classList.add('hidden'); });
  document.querySelectorAll('.matrix-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  const nav = document.getElementById('nav-'+name);
  if (nav) nav.classList.add('active');
  if (name==='matrix') {
    document.getElementById('panel-matrix').classList.add('active');
  } else {
    const panel = document.getElementById('panel-'+name);
    if (panel) { panel.classList.remove('hidden'); panel.classList.add('active'); }
  }
  closePortPopover();
}

function setStatus(msg, isError=false) {
  document.getElementById('status-text').textContent = msg;
  document.getElementById('status-dot').className = 'status-dot'+(isError?' error':'');
}

function v(id) { const el=document.getElementById(id); return el?el.value:''; }
function clearFields(ids) { ids.forEach(id=>{const el=document.getElementById(id);if(el)el.value='';}); }
function esc(s) { if(!s&&s!==0) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
