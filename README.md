# Audio Routing Add-in for Excel Online

Single source of truth for complex audio system routing.  
Enter information once. View it from the perspective of any device.

---

## What it does

**5 tabs in the sidebar:**

| Tab | Purpose |
|-----|---------|
| **Devices** | Register every device in your system (Nuendo, Orange Box, Quantum, TiMax, etc.) |
| **Ports** | Add every input and output port for each device. Name/alias them when used. |
| **Connect** | Record which outputs connect to which inputs. Validates direction mismatches. |
| **Console** | Quantum-specific view: assign ports to channels, groups, auxes, matrices via dropdown. Connections auto-record. |
| **Checks** | Run system validation: duplicate outputs, direction mismatches, orphaned ports, unconnected outputs. |

---

## How to host on GitHub Pages (free, 5 min)

1. Create a new GitHub repository, e.g. `audio-routing-addin`
2. Upload all files from this folder to the repo root
3. Go to **Settings → Pages → Source → main branch / root** → Save
4. Your add-in will be live at `https://YOUR-USERNAME.github.io/audio-routing-addin/`

---

## Edit the manifest

Open `manifest.xml` and replace every instance of `YOUR-GITHUB-USERNAME` with your actual GitHub username.

Also replace the `<Id>` GUID with a fresh one from https://www.guidgenerator.com

---

## Load into Excel Online

1. Open Excel Online (office.com)
2. Open or create a workbook
3. **Insert → Add-ins → Upload My Add-in**
4. Browse to `manifest.xml` → Upload
5. The "Audio Routing" button appears in the Home tab ribbon

---

## First run

Click **Set Up Workbook Sheets** — this creates 4 sheets:
- `Devices` — your device registry
- `Ports` — all ports for all devices  
- `Connections` — signal routing connections
- `Console` — Quantum console assignment view (pre-populated with 38 channels, 16 groups, 16 auxes, 8 matrices)

**The sidebar is the only place you enter data. The sheets are read-only outputs.**

---

## Workflow

```
1. Devices tab  →  Add all devices (Nuendo, Orange Box 1, Orange Box 2, Quantum, TiMax…)

2. Ports tab    →  For each device, add its ports
                   Use "Bulk Add" for devices with many sequential ports
                   Give aliases to ports as you name them in your system

3. Connect tab  →  Record signal routes (output → input)
                   The add-in warns you if an output is used twice

4. Console tab  →  Assign ports to Quantum channels/groups/auxes/matrices
                   Picking a port writes to Connections automatically

5. Checks tab   →  Run validation any time to catch conflicts
```

---

## The sheets stay clean

The 4 sheets have headers in row 1 and data from row 2 downward.  
**Don't type in them directly** — use the sidebar.  
You can read them, filter them, or build your own FILTER/SORT views on separate sheets.

---

## Extending it

The code is plain HTML + JavaScript (no build step).  
Edit `taskpane.html` and `taskpane.js` directly.

To add a new device-specific view tab (e.g. TiMax-specific):  
- Add a `<button>` to the `.nav` in `taskpane.html`  
- Add a `<div class="panel">` for the content  
- Add a render function in `taskpane.js` that filters `state.connections` by device name

---

## Files

```
audio-routing-addin/
├── taskpane.html     ← UI (sidebar panel)
├── taskpane.js       ← All logic + Office.js calls
├── manifest.xml      ← Add-in registration (edit YOUR-GITHUB-USERNAME)
├── commands.html     ← Required stub
└── README.md         ← This file
```
