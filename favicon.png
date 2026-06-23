// ═══════════════════════════════════════════════════════════════════════════════
// SPIRO BATTERY QMS — PWA App Entry Point
// React + XLSX loaded via CDN in index.html
// ═══════════════════════════════════════════════════════════════════════════════

const { useState, useRef, useCallback, useMemo, useEffect } = React;

// XLSX is loaded globally via CDN script tag
// window.XLSX is available globally


// ═══════════════════════════════════════════════════════════════════════════════
// ████  MICROSOFT GRAPH CONFIG  ████
// Fill these in after completing Azure setup (see SETUP_GUIDE.md)
// ═══════════════════════════════════════════════════════════════════════════════

const GRAPH_CONFIG = {
  // From Azure App Registration → Overview
  clientId:   "YOUR_CLIENT_ID_HERE",
  tenantId:   "YOUR_TENANT_ID_HERE",

  // The Microsoft 365 work account that OWNS the OneDrive files
  // This account must have the 3 Excel files already created on its OneDrive
  serviceAccountEmail: "your.serviceaccount@yourcompany.com",

  // Exact filenames of your 3 Excel files on OneDrive (include .xlsx)
  files: {
    OK:       "Spiro_OK_Charging.xlsx",
    REPAIR:   "Spiro_Incoming_Repair.xlsx",
    WARRANTY: "Spiro_Warranty_Parts.xlsx",
  },

  // Exact name of the Excel Table inside each file (must be the same in all 3)
  // When you create the files, format the data as a Table and name it "BatteryLog"
  tableName: "BatteryLog",
};

// ═══════════════════════════════════════════════════════════════════════════════
// ████  MICROSOFT GRAPH API CLIENT  ████
// Uses MSAL implicit flow — token obtained once, reused for all requests
// ═══════════════════════════════════════════════════════════════════════════════

let _msalToken = null;
let _msalExpiry = 0;

async function getMSALToken() {
  if (_msalToken && Date.now() < _msalExpiry - 60000) return _msalToken;

  // Load MSAL from CDN if not already present
  if (!window.msal) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/microsoft-authentication-library/2.38.0/msal-browser.min.js";
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }

  const msalInstance = new window.msal.PublicClientApplication({
    auth: {
      clientId: GRAPH_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: { cacheLocation: "sessionStorage" },
  });

  await msalInstance.initialize();

  const scopes = ["https://graph.microsoft.com/Files.ReadWrite.All"];

  try {
    // Try silent first (uses cached session)
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const result = await msalInstance.acquireTokenSilent({ scopes, account: accounts[0] });
      _msalToken = result.accessToken;
      _msalExpiry = result.expiresOn.getTime();
      return _msalToken;
    }
  } catch {}

  // Popup login (only happens once per session)
  const result = await msalInstance.loginPopup({ scopes });
  _msalToken = result.accessToken;
  _msalExpiry = result.expiresOn?.getTime() || Date.now() + 3600000;
  return _msalToken;
}

// Find a file on OneDrive by name, return its driveItem ID
async function getFileId(token, filename) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root/children?$filter=name eq '${encodeURIComponent(filename)}'&$select=id,name`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Cannot find file "${filename}" on OneDrive. Check filename in GRAPH_CONFIG.`);
  const data = await res.json();
  if (!data.value?.length) throw new Error(`File "${filename}" not found on OneDrive root. Upload it first.`);
  return data.value[0].id;
}

// Append a row to the named Table in the Excel file
async function appendRowToExcel(token, fileId, tableName, rowValues) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/tables/${tableName}/rows/add`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: [rowValues] }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Excel write failed (HTTP ${res.status}). Check table name "${tableName}".`);
  }
  return await res.json();
}

// Master function: route battery → find correct file → append row
async function syncToOneDrive(entry) {
  const token = await getMSALToken();
  const category = entry.route?.category || "REPAIR";
  const filename = GRAPH_CONFIG.files[category];
  if (!filename) throw new Error(`No file configured for category: ${category}`);

  const fileId = await getFileId(token, filename);

  const f = entry.form || {};
  const cellDelta = f.cellVoltageMin && f.cellVoltageMax
    ? ((parseFloat(f.cellVoltageMax) - parseFloat(f.cellVoltageMin)) * 1000).toFixed(0)
    : "";

  // Row order must match your Excel table column headers exactly
  const rowValues = [
    entry.timestamp,
    entry.serial,
    entry.batteryType,
    entry.technician,
    f.cellVoltageMin || "",
    f.cellVoltageMax || "",
    cellDelta,
    f.packVoltage || "",
    f.t1 || "", f.t2 || "", f.t3 || "", f.t4 || "",
    f.mosfetStatus || "",
    f.afeAlarm?.replace(/_/g, " ") || "",
    f.canCharge || "",
    f.canDischarge || "",
    f.ledPattern?.replace(/_/g, " ") || "",
    (f.missingParts || []).join("; "),
    (f.physicalDamage || []).join("; "),
    f.missingBoltCount || "",
    f.missingCellCount || "",
    (entry.route?.issues || []).join("; "),
    f.observations || "",
    category,
  ];

  await appendRowToExcel(token, fileId, GRAPH_CONFIG.tableName, rowValues);
}

// Column headers — must match what's in your Excel tables
const EXCEL_TABLE_HEADERS = [
  "Timestamp", "Serial Number", "Battery Type", "Technician",
  "Cell V Min (V)", "Cell V Max (V)", "Cell Delta (mV)", "Pack Voltage (V)",
  "T1 (°C)", "T2 (°C)", "T3 (°C)", "T4 (°C)",
  "MOSFET Status", "AFE Alarm", "Can Charge", "Can Discharge", "LED Pattern",
  "Missing Parts", "Physical Damage", "Missing Bolt Count", "Missing Cell Count",
  "Issues Flagged", "Observations", "Route Category",
];

// ═══════════════════════════════════════════════════════════════════════════════
// ████  BATTERY & SOP CONSTANTS  ████
// ═══════════════════════════════════════════════════════════════════════════════

const BATTERY_TYPES = {
  C: { name: "CBAK",     prefix: "C7245AC2AA", totalLen: 18, color: "#1565C0", light: "#E3F2FD" },
  U: { name: "Unique",   prefix: "U7245AU1G1",  totalLen: 18, color: "#2E7D32", light: "#E8F5E9" },
  A: { name: "Ampace",   prefix: "A7246AX1A1",  totalLen: 18, color: "#E65100", light: "#FFF3E0" },
  D: { name: "Greenway", prefix: "D72451LWK",   totalLen: 18, color: "#6A1B9A", light: "#F3E5F5" },
};

const LED_PATTERNS = {
  C: [
    { value: "1_led_low_soc",  label: "1 LED – Low SOC (<25%)" },
    { value: "2_leds",         label: "2 LEDs – SOC 25–50%" },
    { value: "3_leds",         label: "3 LEDs – SOC 50–75%" },
    { value: "4_leds_full",    label: "4 LEDs – Full / Normal" },
    { value: "blinking_fault", label: "Blinking – Fault / Error" },
    { value: "no_display",     label: "No Display" },
  ],
  U: [
    { value: "1_led_low_soc",      label: "1 LED – Low SOC" },
    { value: "2_leds",             label: "2 LEDs – SOC 25–50%" },
    { value: "3_leds",             label: "3 LEDs – SOC 50–75%" },
    { value: "4_leds_full",        label: "4 LEDs – Full / Normal" },
    { value: "1_middle_blink_red", label: "1 Middle LED Blinking Red – BMS Fault" },
    { value: "3_middle_blink_red", label: "3 Middle LEDs Blinking Red – Cell Fault" },
    { value: "5_all_blink_red",    label: "5 All LEDs Blinking Red – Critical Fault" },
    { value: "no_display",         label: "No Display" },
  ],
  A: [
    { value: "1_led_low_soc",       label: "1 LED – Low SOC" },
    { value: "2_leds",              label: "2 LEDs – SOC 25–50%" },
    { value: "3_leds",              label: "3 LEDs – SOC 50–75%" },
    { value: "4_leds",              label: "4 LEDs – SOC 75–100%" },
    { value: "5_leds_full",         label: "5 LEDs – Full" },
    { value: "2_extreme_blinking",  label: "2 Extreme LEDs Blinking – Comm Fault" },
    { value: "5_leds_blinking",     label: "5 LEDs Blinking – Critical Fault" },
    { value: "1_led_not_charging",  label: "1 LED, Not Charging – CHG MOSFET Fault" },
    { value: "no_display",          label: "No Display" },
  ],
  D: [
    { value: "1_led_low_soc",  label: "1 LED – Low SOC" },
    { value: "2_leds",         label: "2 LEDs – SOC 25–50%" },
    { value: "3_leds",         label: "3 LEDs – SOC 50–75%" },
    { value: "4_leds_full",    label: "4 LEDs – Full / Normal" },
    { value: "blinking_fault", label: "Blinking – Fault" },
    { value: "no_display",     label: "No Display" },
  ],
};

const LED_IS_FAULT = (v) => ["blinking_fault","1_middle_blink_red","3_middle_blink_red","5_all_blink_red",
  "2_extreme_blinking","5_leds_blinking","1_led_not_charging","no_display"].includes(v);

const AFE_ALARMS = [
  { value: "none",               label: "None" },
  { value: "cell_overvoltage",   label: "Cell Overvoltage – All LEDs on, charging stops" },
  { value: "cell_undervoltage",  label: "Cell Undervoltage – 1 LED blinking, discharge cut-off" },
  { value: "pack_overvoltage",   label: "Pack Overvoltage – BMS shuts CHG MOSFET" },
  { value: "pack_undervoltage",  label: "Pack Undervoltage – BMS shuts DSG MOSFET" },
  { value: "overtemp_charge",    label: "Overtemp (Charging) – Charging halted, LED blinks" },
  { value: "overtemp_discharge", label: "Overtemp (Discharging) – DSG cut, thermal alarm" },
  { value: "short_circuit",      label: "Short Circuit – Immediate DSG MOSFET trip" },
  { value: "overcurrent_chg",    label: "Overcurrent (Charging) – CHG MOSFET trips" },
  { value: "overcurrent_dsg",    label: "Overcurrent (Discharging) – DSG MOSFET trips" },
  { value: "afe_comm_fault",     label: "AFE Comm Fault – BMS unresponsive, no LED" },
  { value: "balancing_fault",    label: "Cell Balancing Fault – Delta V >500 mV persists" },
];

const MISSING_PARTS = [
  { value: "handles",          label: "Handle(s)" },
  { value: "front_cover",      label: "Front Cover" },
  { value: "back_cover",       label: "Back Cover" },
  { value: "charging_port",    label: "Charging Port" },
  { value: "discharging_port", label: "Discharging Port" },
  { value: "bms_board",        label: "BMS Board" },
  { value: "wire_harness",     label: "Wire Harness" },
  { value: "cells",            label: "Cells (stolen/missing)" },
  { value: "iot_device",       label: "IoT Device" },
  { value: "epoxy_board",      label: "Epoxy Board(s)" },
  { value: "bolts",            label: "Bolts" },
  { value: "ntc_cable",        label: "NTC Cable" },
  { value: "bms_cover",        label: "BMS Cover" },
  { value: "connectors",       label: "Connectors" },
];

const PHYSICAL_DAMAGE = [
  { value: "handle_broken",            label: "Handle – Broken" },
  { value: "handle_bent",              label: "Handle – Bent" },
  { value: "front_cover_broken",       label: "Front Cover – Broken" },
  { value: "front_cover_bent",         label: "Front Cover – Bent" },
  { value: "front_cover_holes",        label: "Front Cover – Has Holes" },
  { value: "front_cover_destroyed",    label: "Front Cover – Destroyed" },
  { value: "back_cover_broken",        label: "Back Cover – Broken" },
  { value: "back_cover_bent",          label: "Back Cover – Bent" },
  { value: "back_cover_holes",         label: "Back Cover – Has Holes" },
  { value: "back_cover_destroyed",     label: "Back Cover – Destroyed" },
  { value: "charging_port_damaged",    label: "Charging Port – Damaged" },
  { value: "discharging_port_damaged", label: "Discharging Port – Damaged" },
  { value: "burn_marks",               label: "Burn Marks" },
  { value: "swelling",                 label: "Case Swelling" },
  { value: "water_damage",             label: "Water / Moisture Damage" },
];

const REPAIR_STATUSES = [
  "Pending Repair","In Repair","Post-Repair Check",
  "Online Test","Offline – Failed","Ready for Dispatch","Dispatched",
];

const TECHNICIANS = [
  { id: "t1", name: "Dennis Mugambwa",  role: "Battery Quality Engineer", pin: "1234" },
  { id: "t2", name: "Technician Alpha", role: "CBAK Rework Tech",         pin: "2222" },
  { id: "t3", name: "Technician Beta",  role: "PDI Inspector",            pin: "3333" },
];

// ═══════════════════════════════════════════════════════════════════════════════
// ████  SOP ROUTING  ████
// ═══════════════════════════════════════════════════════════════════════════════

function routeBattery(form) {
  const issues = [], warranty = [];
  const minV = parseFloat(form.cellVoltageMin), maxV = parseFloat(form.cellVoltageMax);
  if (!isNaN(minV) && minV < 2.6)  issues.push(`Cell min voltage ${minV}V < 2.6V threshold`);
  if (!isNaN(minV) && !isNaN(maxV)) {
    const d = (maxV - minV) * 1000;
    if (d > 500) issues.push(`Cell delta ${d.toFixed(0)}mV exceeds 500mV limit`);
  }
  const pv = parseFloat(form.packVoltage);
  if (!isNaN(pv) && pv < 42)   issues.push("Pack undervoltage (<42V)");
  if (!isNaN(pv) && pv > 54.6) issues.push("Pack overvoltage (>54.6V)");
  ["t1","t2","t3","t4"].forEach((t,i) => {
    const v = parseFloat(form[t]);
    if (!isNaN(v) && v > 45) issues.push(`T${i+1} high temperature: ${v}°C`);
  });
  if (form.mosfetStatus && form.mosfetStatus !== "OK") issues.push(`MOSFET: ${form.mosfetStatus}`);
  if (form.canCharge    === "No") issues.push("Cannot charge");
  if (form.canDischarge === "No") issues.push("Cannot discharge");
  if (form.ledPattern && LED_IS_FAULT(form.ledPattern)) issues.push(`LED fault: ${form.ledPattern.replace(/_/g," ")}`);
  if (form.afeAlarm && form.afeAlarm !== "none") issues.push(`AFE: ${form.afeAlarm.replace(/_/g," ")}`);

  const mp = form.missingParts || [];
  if (mp.includes("cells"))       warranty.push("Stolen/missing cells – warranty claim");
  if (mp.includes("bms_board"))   warranty.push("Missing BMS board");
  if (mp.includes("iot_device"))  warranty.push("Missing IoT device");
  if (mp.includes("wire_harness"))warranty.push("Missing wire harness");

  const pd = form.physicalDamage || [];
  if (pd.some(d => ["front_cover_destroyed","back_cover_destroyed","burn_marks"].includes(d)))
    warranty.push("Severe damage / vandalism");

  if (warranty.length > 0) return { category: "WARRANTY", issues: [...issues, ...warranty] };
  if (issues.length  > 0) return { category: "REPAIR",   issues };
  return                          { category: "OK",        issues: [] };
}

// ═══════════════════════════════════════════════════════════════════════════════
// ████  LOCAL DATABASE  ████
// ═══════════════════════════════════════════════════════════════════════════════

const DB_KEY = "spiro_qms_v4";
function loadDB()  { try { return JSON.parse(localStorage.getItem(DB_KEY) || "{}") || { batteries:[], repairs:[] }; } catch { return { batteries:[], repairs:[] }; } }
function saveDB(d) { try { localStorage.setItem(DB_KEY, JSON.stringify(d)); } catch {} }

let _db = loadDB();
const ldb = {
  all:         ()    => [...(_db.batteries||[])],
  allRepairs:  ()    => [...(_db.repairs||[])],
  insert:      (e)   => { _db.batteries = [e, ...(_db.batteries||[]).filter(b=>b.serial!==e.serial)]; saveDB(_db); },
  addRepair:   (r)   => { _db.repairs   = [...(_db.repairs||[]), r]; saveDB(_db); },
  updateRepair:(id,u)=> { _db.repairs   = (_db.repairs||[]).map(r=>r.id===id?{...r,...u}:r); saveDB(_db); },
  clear:       ()    => { _db = { batteries:[], repairs:[] }; saveDB(_db); },
};

// ═══════════════════════════════════════════════════════════════════════════════
// ████  QR SCANNER  ████
// ═══════════════════════════════════════════════════════════════════════════════

function QRScanner({ onScan, onClose }) {
  const videoRef = useRef(null), canvasRef = useRef(null), streamRef = useRef(null), rafRef = useRef(null);
  const [status, setStatus] = useState("Loading scanner…");

  useEffect(() => {
    let active = true;
    const load = async () => {
      if (!window.jsQR) {
        await new Promise((res,rej) => {
          const s = document.createElement("script");
          s.src = "https://cdnjs.cloudflare.com/ajax/libs/jsQR/1.4.0/jsQR.min.js";
          s.onload = res; s.onerror = rej;
          document.head.appendChild(s);
        });
      }
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode:"environment" } });
        if (!active) { stream.getTracks().forEach(t=>t.stop()); return; }
        streamRef.current = stream;
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
        setStatus("Point at barcode or QR code");
        const scan = () => {
          if (!active) return;
          const v = videoRef.current, c = canvasRef.current;
          if (v && c && v.readyState === v.HAVE_ENOUGH_DATA) {
            c.width = v.videoWidth; c.height = v.videoHeight;
            const ctx = c.getContext("2d"); ctx.drawImage(v,0,0);
            const img = ctx.getImageData(0,0,c.width,c.height);
            const code = window.jsQR(img.data,img.width,img.height,{inversionAttempts:"dontInvert"});
            if (code?.data) { onScan(code.data); return; }
          }
          rafRef.current = requestAnimationFrame(scan);
        };
        rafRef.current = requestAnimationFrame(scan);
      } catch { setStatus("Camera access denied. Use manual entry."); }
    };
    load();
    return () => { active=false; cancelAnimationFrame(rafRef.current); streamRef.current?.getTracks().forEach(t=>t.stop()); };
  }, [onScan]);

  return (
    <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.94)",zIndex:2000,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center" }}>
      <p style={{ color:"#fff",fontSize:13,marginBottom:12,opacity:0.8 }}>{status}</p>
      <div style={{ position:"relative",width:"min(90vw,360px)",aspectRatio:"1",borderRadius:16,overflow:"hidden",border:"3px solid #42A5F5" }}>
        <video ref={videoRef} style={{ width:"100%",height:"100%",objectFit:"cover" }} playsInline muted />
        <canvas ref={canvasRef} style={{ display:"none" }} />
        <div style={{ position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",pointerEvents:"none" }}>
          <div style={{ width:"60%",height:"60%",border:"2.5px solid #42A5F5",borderRadius:10,boxShadow:"0 0 0 9999px rgba(0,0,0,0.5)" }} />
        </div>
      </div>
      <button onClick={onClose} style={{ marginTop:20,padding:"10px 32px",borderRadius:24,background:"#ef5350",color:"#fff",border:"none",fontWeight:700,fontSize:15,cursor:"pointer" }}>Cancel</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ████  SYNC STATUS INDICATOR  ████
// ═══════════════════════════════════════════════════════════════════════════════

const SYNC = { IDLE:"idle", SYNCING:"syncing", OK:"ok", ERROR:"error" };

function SyncBadge({ status, error }) {
  const map = {
    idle:    { bg:"#E3F2FD", color:"#1565C0", icon:"☁️", text:"Ready to sync" },
    syncing: { bg:"#FFF8E1", color:"#E65100", icon:"⏳", text:"Syncing to OneDrive…" },
    ok:      { bg:"#E8F5E9", color:"#1B7A3E", icon:"✅", text:"Saved to OneDrive" },
    error:   { bg:"#FFEBEE", color:"#B71C1C", icon:"❌", text: error || "Sync failed" },
  };
  const s = map[status] || map.idle;
  return (
    <div style={{ display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:20,background:s.bg,border:`1px solid ${s.color}`,fontSize:12,color:s.color,fontWeight:600,maxWidth:"100%" }}>
      <span>{s.icon}</span><span style={{ flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{s.text}</span>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ████  UI HELPERS  ████
// ═══════════════════════════════════════════════════════════════════════════════

function CheckGroup({ options, value=[], onChange, cols=2 }) {
  const toggle = v => onChange(value.includes(v)?value.filter(x=>x!==v):[...value,v]);
  return (
    <div style={{ display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap:"6px 10px" }}>
      {options.map(o=>(
        <label key={o.value} style={{ display:"flex",alignItems:"flex-start",gap:7,fontSize:13,cursor:"pointer",lineHeight:1.35 }}>
          <input type="checkbox" checked={value.includes(o.value)} onChange={()=>toggle(o.value)} style={{ marginTop:2,accentColor:"#1565C0",flexShrink:0 }}/>
          <span>{o.label}</span>
        </label>
      ))}
    </div>
  );
}

// Inline config checker — shown when GRAPH_CONFIG has placeholder values
function ConfigNotice({ onDismiss }) {
  const isConfigured = GRAPH_CONFIG.clientId !== "YOUR_CLIENT_ID_HERE";
  if (isConfigured) return null;
  return (
    <div style={{ background:"#FFF8E1",border:"1.5px solid #F9A825",borderRadius:10,padding:"14px 14px",marginBottom:14 }}>
      <div style={{ fontWeight:700,color:"#E65100",marginBottom:6 }}>⚙️ OneDrive Not Configured</div>
      <div style={{ fontSize:13,color:"#5D4037",lineHeight:1.5 }}>
        The app is running in <strong>offline mode</strong>. Data saves locally only.<br/>
        Follow the <strong>SETUP_GUIDE.md</strong> file included with this app to connect your OneDrive in about 15 minutes.
        Then paste your <code style={{ background:"#fff",padding:"1px 4px",borderRadius:3 }}>clientId</code> and <code style={{ background:"#fff",padding:"1px 4px",borderRadius:3 }}>tenantId</code> into <strong>GRAPH_CONFIG</strong> at the top of the app file.
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ████  STYLE TOKENS  ████
// ═══════════════════════════════════════════════════════════════════════════════

const C = { navy:"#0D1B3E",blue:"#1565C0",sky:"#42A5F5",green:"#1B7A3E",amber:"#E65100",red:"#B71C1C",surface:"#F4F6FB",border:"#DDE3EE",muted:"#6B7280",text:"#1A1F36" };
const ROUTE_S = {
  OK:       { bg:"#E8F5E9",border:"#1B7A3E",text:"#1B7A3E",icon:"✅",label:"Incoming – OK for Charging" },
  REPAIR:   { bg:"#E3F2FD",border:"#1565C0",text:"#1565C0",icon:"🔧",label:"Incoming – Repair" },
  WARRANTY: { bg:"#FFEBEE",border:"#B71C1C",text:"#B71C1C",icon:"⚠️", label:"Warranty – Spare Parts" },
};
const inp  = { width:"100%",padding:"9px 11px",borderRadius:7,border:`1.5px solid ${C.border}`,fontSize:14,fontFamily:"inherit",background:"#fff",boxSizing:"border-box",color:C.text };
const card = { background:"#fff",borderRadius:12,padding:"16px 15px",marginBottom:14,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:`1px solid ${C.border}` };
const sh   = { fontSize:11,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:10,paddingBottom:5,borderBottom:`1px solid ${C.border}` };
const fl   = { display:"block",fontSize:12,fontWeight:600,color:C.muted,marginBottom:4,textTransform:"uppercase",letterSpacing:0.4 };
const btn  = (bg,co="#fff",bd) => ({ padding:"10px 18px",borderRadius:8,border:bd?`1.5px solid ${bd}`:"none",background:bg,color:co,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit" });

// ═══════════════════════════════════════════════════════════════════════════════
// ████  MAIN APP  ████
// ═══════════════════════════════════════════════════════════════════════════════

function SpiroBatteryQMS() {
  const [tech,       setTech]       = useState(null);
  const [loginName,  setLoginName]  = useState("");
  const [loginPin,   setLoginPin]   = useState("");
  const [loginErr,   setLoginErr]   = useState("");
  const [view,       setView]       = useState("scan");
  const [serial,     setSerial]     = useState("");
  const [manualType, setManualType] = useState("");
  const [battery,    setBattery]    = useState(null);
  const [form,       setForm]       = useState({});
  const [submitted,  setSubmitted]  = useState(null);
  const [scanning,   setScanning]   = useState(false);
  const [syncStatus, setSyncStatus] = useState(SYNC.IDLE);
  const [syncError,  setSyncError]  = useState("");
  const [batteries,  setBatteries]  = useState(ldb.all());
  const [repairs,    setRepairs]    = useState(ldb.allRepairs());
  const [repairModal,setRepairModal]= useState(null);
  const [repairForm, setRepairForm] = useState({});
  const [searchQ,    setSearchQ]    = useState("");
  const [filterCat,  setFilterCat]  = useState("ALL");

  const refresh = useCallback(() => { setBatteries(ldb.all()); setRepairs(ldb.allRepairs()); },[]);

  // ── Login ──
  const handleLogin = () => {
    const found = TECHNICIANS.find(t=>t.name.toLowerCase().includes(loginName.toLowerCase())&&t.pin===loginPin);
    found ? (setTech(found),setLoginErr("")) : setLoginErr("Name or PIN not recognised.");
  };

  // ── Serial detection ──
  const detectAndSet = useCallback((raw) => {
    const s = raw.trim().toUpperCase();
    setSerial(s); setSubmitted(null);
    for (const [k,bt] of Object.entries(BATTERY_TYPES)) {
      if (s.startsWith(bt.prefix) && s.length === bt.totalLen) { setBattery({key:k,...bt}); setManualType(""); setForm({}); return; }
    }
    if (manualType && BATTERY_TYPES[manualType] && s.length === BATTERY_TYPES[manualType].totalLen) {
      setBattery({key:manualType,...BATTERY_TYPES[manualType]}); setForm({}); return;
    }
    setBattery(null);
  },[manualType]);

  const handleManualType = k => {
    setManualType(k); setSubmitted(null); setForm({});
    (serial && BATTERY_TYPES[k] && serial.length === BATTERY_TYPES[k].totalLen)
      ? setBattery({key:k,...BATTERY_TYPES[k]}) : setBattery(null);
  };

  const sf = (id,val) => setForm(p=>({...p,[id]:val}));

  // ── Submit ──
  const handleSubmit = async () => {
    if (!battery||!serial) return;
    if (!form.packVoltage) { alert("Pack voltage is required."); return; }

    const route = routeBattery(form);
    const entry = {
      id: Date.now().toString(),
      timestamp: new Date().toLocaleString("en-UG"),
      serial, batteryType: battery.name, batteryKey: battery.key,
      technician: tech?.name || "Unknown",
      form: {...form}, route,
    };

    // 1. Save locally first (always works, even offline)
    ldb.insert(entry); refresh(); setSubmitted(entry);

    // 2. Try OneDrive sync
    const isConfigured = GRAPH_CONFIG.clientId !== "YOUR_CLIENT_ID_HERE";
    if (isConfigured) {
      setSyncStatus(SYNC.SYNCING); setSyncError("");
      try {
        await syncToOneDrive(entry);
        setSyncStatus(SYNC.OK);
      } catch (e) {
        setSyncStatus(SYNC.ERROR);
        setSyncError(e.message);
        // Keep local copy — technician is not blocked
      }
    } else {
      setSyncStatus(SYNC.IDLE);
    }
  };

  const handleNext = () => { setSerial(""); setBattery(null); setForm({}); setSubmitted(null); setManualType(""); setSyncStatus(SYNC.IDLE); };

  const counts = useMemo(()=>({
    OK:      batteries.filter(b=>b.route?.category==="OK").length,
    REPAIR:  batteries.filter(b=>b.route?.category==="REPAIR").length,
    WARRANTY:batteries.filter(b=>b.route?.category==="WARRANTY").length,
  }),[batteries]);

  const typeCounts = useMemo(()=>{ const tc={C:0,U:0,A:0,D:0}; batteries.forEach(b=>{if(tc[b.batteryKey]!==undefined)tc[b.batteryKey]++;}); return tc; },[batteries]);

  const filtered = useMemo(()=>batteries.filter(b=>
    (filterCat==="ALL"||b.route?.category===filterCat)&&
    (!searchQ||b.serial?.includes(searchQ.toUpperCase())||b.batteryType?.toLowerCase().includes(searchQ.toLowerCase())||b.technician?.toLowerCase().includes(searchQ.toLowerCase()))
  ),[batteries,filterCat,searchQ]);

  // Pending sync queue — retry failed entries
  const [pendingSync, setPendingSync] = useState([]);
  const retrySync = async () => {
    if (pendingSync.length===0) return;
    setSyncStatus(SYNC.SYNCING);
    let failed=[];
    for (const entry of pendingSync) {
      try { await syncToOneDrive(entry); }
      catch { failed.push(entry); }
    }
    setPendingSync(failed);
    setSyncStatus(failed.length>0?SYNC.ERROR:SYNC.OK);
    if (failed.length>0) setSyncError(`${failed.length} entr${failed.length>1?"ies":"y"} still pending sync`);
  };

  // ══════════════════════════════════════════════════════════════════════════════
  // LOGIN SCREEN
  // ══════════════════════════════════════════════════════════════════════════════

  if (!tech) return (
    <div style={{ minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter','Segoe UI',sans-serif" }}>
      <div style={{ background:"#fff",borderRadius:16,padding:"36px 28px",width:"min(92vw,360px)",boxShadow:"0 20px 60px rgba(0,0,0,0.4)" }}>
        <div style={{ textAlign:"center",marginBottom:24 }}>
          <div style={{ fontSize:40 }}>⚡</div>
          <div style={{ fontSize:22,fontWeight:800,color:C.navy }}>Spiro Battery QMS</div>
          <div style={{ fontSize:12,color:C.muted,marginTop:3 }}>Battery Quality Management System</div>
        </div>
        <label style={fl}>Your Name</label>
        <input style={{...inp,marginBottom:12}} placeholder="e.g. Dennis" value={loginName} onChange={e=>setLoginName(e.target.value)}/>
        <label style={fl}>PIN</label>
        <input style={{...inp,marginBottom:16}} type="password" placeholder="4-digit PIN" maxLength={4}
          value={loginPin} onChange={e=>setLoginPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()}/>
        {loginErr && <div style={{color:C.red,fontSize:13,marginBottom:10}}>{loginErr}</div>}
        <button style={{...btn(C.navy),width:"100%",padding:13,fontSize:15}} onClick={handleLogin}>Sign In →</button>
        <div style={{marginTop:12,fontSize:11,color:C.muted,textAlign:"center"}}>Dennis→1234 · Alpha→2222 · Beta→3333</div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════════
  // MAIN APP
  // ══════════════════════════════════════════════════════════════════════════════

  const navItems = [{id:"scan",icon:"⚡",label:"Scan"},{id:"dashboard",icon:"📊",label:"Dashboard"},{id:"log",icon:"📋",label:"Log"},{id:"repairs",icon:"🔧",label:"Repairs"},{id:"search",icon:"🔍",label:"Search"}];

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',sans-serif",minHeight:"100vh",background:C.surface,color:C.text,paddingBottom:76 }}>

      {/* Header */}
      <div style={{ background:C.navy,color:"#fff",padding:"11px 15px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 12px rgba(0,0,0,0.3)" }}>
        <div>
          <div style={{ fontSize:16,fontWeight:800,letterSpacing:0.5 }}>⚡ Spiro Battery QMS</div>
          <div style={{ fontSize:11,opacity:0.6,marginTop:1 }}>{tech.name} · {tech.role}</div>
        </div>
        <div style={{ display:"flex",alignItems:"center",gap:8 }}>
          {pendingSync.length > 0 && (
            <button onClick={retrySync} style={{ padding:"4px 10px",borderRadius:16,background:"#F9A825",border:"none",color:"#fff",fontWeight:700,fontSize:11,cursor:"pointer" }}>
              🔄 Retry {pendingSync.length}
            </button>
          )}
          <button onClick={()=>setTech(null)} style={{ background:"rgba(255,255,255,0.12)",border:"none",color:"#fff",padding:"5px 11px",borderRadius:20,fontSize:11,cursor:"pointer" }}>Sign out</button>
        </div>
      </div>

      <div style={{ maxWidth:640,margin:"0 auto",padding:"13px 11px" }}>
        <ConfigNotice />

        {/* ══ SCAN VIEW ══ */}
        {view==="scan" && (
          <>
            {scanning && <QRScanner onScan={s=>{setScanning(false);detectAndSet(s);}} onClose={()=>setScanning(false)}/>}

            {!submitted && (
              <>
                <div style={card}>
                  <div style={sh}>Battery Identification</div>
                  {/* Type buttons */}
                  <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:13 }}>
                    {Object.entries(BATTERY_TYPES).map(([k,bt])=>(
                      <button key={k} onClick={()=>handleManualType(k)} style={{
                        padding:"8px 4px",borderRadius:8,cursor:"pointer",textAlign:"center",fontWeight:700,fontSize:12,
                        border:`2px solid ${(battery?.key===k||manualType===k)?bt.color:C.border}`,
                        background:(battery?.key===k||manualType===k)?bt.light:"#fff",
                        color:(battery?.key===k||manualType===k)?bt.color:C.muted }}>
                        {bt.name}
                      </button>
                    ))}
                  </div>
                  <label style={fl}>Serial Number</label>
                  <div style={{ display:"flex",gap:8 }}>
                    <input style={{...inp,flex:1,fontSize:15,letterSpacing:0.8}} value={serial}
                      placeholder="Scan or type serial number…" onChange={e=>detectAndSet(e.target.value)}/>
                    <button onClick={()=>setScanning(true)} style={{ padding:"9px 13px",borderRadius:8,background:C.sky,border:"none",fontSize:18,cursor:"pointer" }} title="Scan QR / Barcode">📷</button>
                  </div>
                  {serial && !battery && serial.length>5 && (
                    <div style={{ marginTop:8,fontSize:12,color:C.amber,background:"#FFF8E1",padding:"6px 10px",borderRadius:6 }}>
                      ⚠️ Not auto-detected. Select type above. CBAK/Unique/Ampace = 18 chars · Greenway = 18 chars
                    </div>
                  )}
                  {battery && (
                    <div style={{ marginTop:10,display:"flex",alignItems:"center",gap:8 }}>
                      <span style={{ background:battery.light,color:battery.color,border:`1px solid ${battery.color}`,borderRadius:10,padding:"2px 10px",fontSize:12,fontWeight:700 }}>● {battery.name}</span>
                      <span style={{ fontSize:12,color:C.muted }}>{serial.length}/{battery.totalLen} chars</span>
                    </div>
                  )}
                </div>

                {battery && (
                  <div style={card}>
                    {/* ── Voltages ── */}
                    <div style={sh}>Electrical Measurements</div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
                      <div>
                        <label style={fl}>Cell V Min (V)</label>
                        <input style={inp} type="number" step="0.001" placeholder="e.g. 3.280"
                          value={form.cellVoltageMin||""} onChange={e=>sf("cellVoltageMin",e.target.value)}/>
                        {form.cellVoltageMin && parseFloat(form.cellVoltageMin)<2.6 && <div style={{color:C.red,fontSize:11,marginTop:2}}>⚠️ Below 2.6V!</div>}
                      </div>
                      <div>
                        <label style={fl}>Cell V Max (V)</label>
                        <input style={inp} type="number" step="0.001" placeholder="e.g. 3.750"
                          value={form.cellVoltageMax||""} onChange={e=>sf("cellVoltageMax",e.target.value)}/>
                      </div>
                    </div>
                    {form.cellVoltageMin && form.cellVoltageMax && (() => {
                      const d = (parseFloat(form.cellVoltageMax)-parseFloat(form.cellVoltageMin))*1000;
                      return <div style={{ fontSize:12,marginBottom:12,padding:"6px 10px",borderRadius:6,background:d>500?"#FFEBEE":"#E8F5E9",color:d>500?C.red:C.green }}>
                        Cell Δ = {d.toFixed(0)} mV {d>500?"— EXCEEDS 500 mV → REPAIR":"— Within limit ✓"}
                      </div>;
                    })()}
                    <label style={fl}>Pack Voltage (V) *</label>
                    <input style={{...inp,marginBottom:14}} type="number" step="0.1" placeholder="e.g. 48.5"
                      value={form.packVoltage||""} onChange={e=>sf("packVoltage",e.target.value)}/>

                    {/* ── Temperatures ── */}
                    <div style={sh}>🌡 Temperatures (°C)</div>
                    <div style={{ display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7,marginBottom:14 }}>
                      {["t1","t2","t3","t4"].map((t,i)=>(
                        <div key={t}>
                          <label style={{...fl,fontSize:11}}>T{i+1}</label>
                          <input style={{...inp,textAlign:"center"}} type="number" step="0.1"
                            value={form[t]||""} onChange={e=>sf(t,e.target.value)}/>
                          {form[t]&&parseFloat(form[t])>45&&<div style={{color:C.red,fontSize:10}}>HIGH!</div>}
                        </div>
                      ))}
                    </div>

                    {/* ── BMS Status ── */}
                    <div style={sh}>⚙️ BMS / Functional Status</div>
                    <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10 }}>
                      <div>
                        <label style={fl}>MOSFET Status</label>
                        <select style={inp} value={form.mosfetStatus||""} onChange={e=>sf("mosfetStatus",e.target.value)}>
                          <option value="">— select —</option>
                          <option>OK</option><option>Fault – CHG side</option><option>Fault – DSG side</option><option>Fault – Both</option><option>Unknown</option>
                        </select>
                      </div>
                      <div>
                        <label style={fl}>AFE Alarm</label>
                        <select style={inp} value={form.afeAlarm||""} onChange={e=>sf("afeAlarm",e.target.value)}>
                          <option value="">— select —</option>
                          {AFE_ALARMS.map(a=><option key={a.value} value={a.value}>{a.label}</option>)}
                        </select>
                      </div>
                      <div>
                        <label style={fl}>Can Charge?</label>
                        <select style={inp} value={form.canCharge||""} onChange={e=>sf("canCharge",e.target.value)}>
                          <option value="">—</option><option>Yes</option><option>No</option>
                        </select>
                      </div>
                      <div>
                        <label style={fl}>Can Discharge?</label>
                        <select style={inp} value={form.canDischarge||""} onChange={e=>sf("canDischarge",e.target.value)}>
                          <option value="">—</option><option>Yes</option><option>No</option>
                        </select>
                      </div>
                    </div>

                    {/* ── LED ── */}
                    <label style={fl}>LED Pattern – {battery.name}</label>
                    <select style={{...inp,marginBottom:14}} value={form.ledPattern||""} onChange={e=>sf("ledPattern",e.target.value)}>
                      <option value="">— what do you see? —</option>
                      {(LED_PATTERNS[battery.key]||[]).map(p=><option key={p.value} value={p.value}>{p.label}</option>)}
                    </select>

                    {/* ── Physical ── */}
                    <div style={sh}>🔩 Physical Condition</div>
                    <label style={fl}>Missing Parts</label>
                    <CheckGroup options={MISSING_PARTS} value={form.missingParts||[]} onChange={v=>sf("missingParts",v)}/>
                    {(form.missingParts||[]).includes("bolts") && (
                      <div style={{ marginTop:8 }}>
                        <label style={fl}>Missing Bolt Count</label>
                        <input style={inp} type="number" min="0" value={form.missingBoltCount||""} onChange={e=>sf("missingBoltCount",e.target.value)}/>
                      </div>
                    )}
                    {(form.missingParts||[]).includes("cells") && (
                      <div style={{ marginTop:8 }}>
                        <label style={fl}>Stolen / Missing Cell Count</label>
                        <input style={inp} type="number" min="0" value={form.missingCellCount||""} onChange={e=>sf("missingCellCount",e.target.value)}/>
                      </div>
                    )}
                    <label style={{...fl,marginTop:14}}>Physical Damage</label>
                    <CheckGroup options={PHYSICAL_DAMAGE} value={form.physicalDamage||[]} onChange={v=>sf("physicalDamage",v)}/>

                    {/* ── Type-specific ── */}
                    {battery.key==="C" && <>
                      <div style={{...sh,marginTop:14}}>🔵 CBAK-Specific</div>
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                        <div><label style={fl}>BMS FW Version</label>
                          <select style={inp} value={form.bmsVersion||""} onChange={e=>sf("bmsVersion",e.target.value)}>
                            <option value="">—</option>{["B2.3","B2.4","B2.5","B2.6","B2.7"].map(v=><option key={v}>{v}</option>)}
                          </select></div>
                        <div><label style={fl}>Air Leak Test (kPa)</label>
                          <input style={inp} type="number" step="0.1" placeholder="4.1–4.3 OK" value={form.airLeakKpa||""} onChange={e=>sf("airLeakKpa",e.target.value)}/></div>
                      </div>
                    </>}
                    {battery.key==="U" && <>
                      <div style={{...sh,marginTop:14}}>🟢 Unique-Specific</div>
                      <label style={fl}>Cell Configuration</label>
                      <select style={inp} value={form.cellConfig||""} onChange={e=>sf("cellConfig",e.target.value)}>
                        <option value="">—</option><option>14S4P</option><option>14S3P</option><option>Other</option>
                      </select>
                    </>}
                    {battery.key==="A" && <>
                      <div style={{...sh,marginTop:14}}>🟠 Ampace-Specific</div>
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                        <div><label style={fl}>BMS Communication</label>
                          <select style={inp} value={form.bmsComms||""} onChange={e=>sf("bmsComms",e.target.value)}>
                            <option value="">—</option><option>OK</option><option>No Comms</option><option>Intermittent</option>
                          </select></div>
                        <div><label style={fl}>Cell Balancing</label>
                          <select style={inp} value={form.balancing||""} onChange={e=>sf("balancing",e.target.value)}>
                            <option value="">—</option><option>Active</option><option>Inactive</option><option>Unknown</option>
                          </select></div>
                      </div>
                    </>}
                    {battery.key==="D" && <>
                      <div style={{...sh,marginTop:14}}>🟣 Greenway-Specific</div>
                      <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:10 }}>
                        <div><label style={fl}>Case Condition</label>
                          <select style={inp} value={form.caseCondition||""} onChange={e=>sf("caseCondition",e.target.value)}>
                            <option value="">—</option><option>Good</option><option>Cracked</option><option>Dented</option><option>Burnt</option>
                          </select></div>
                        <div><label style={fl}>Label / Sticker</label>
                          <select style={inp} value={form.labelIntact||""} onChange={e=>sf("labelIntact",e.target.value)}>
                            <option value="">—</option><option>Intact</option><option>Faded</option><option>Missing</option>
                          </select></div>
                      </div>
                    </>}

                    {/* ── Observations ── */}
                    <div style={{...sh,marginTop:14}}>📝 Observations</div>
                    <textarea rows={3} style={{...inp,resize:"vertical"}}
                      placeholder="Anything different not captured above…"
                      value={form.observations||""} onChange={e=>sf("observations",e.target.value)}/>

                    <button style={{...btn(C.navy),width:"100%",padding:14,fontSize:15,marginTop:14}} onClick={handleSubmit}>
                      Submit &amp; Sync to OneDrive →
                    </button>
                  </div>
                )}
              </>
            )}

            {/* Result */}
            {submitted && (() => {
              const rs = ROUTE_S[submitted.route.category];
              const bt = BATTERY_TYPES[submitted.batteryKey];
              return (
                <div>
                  <div style={{...card,border:`2.5px solid ${rs.border}`,background:rs.bg}}>
                    <div style={{ fontSize:34,marginBottom:6 }}>{rs.icon}</div>
                    <div style={{ fontSize:20,fontWeight:800,color:rs.text }}>{rs.label}</div>
                    <div style={{ fontSize:13,color:C.muted,marginTop:3 }}>{submitted.serial}</div>
                    <span style={{ background:bt.light,color:bt.color,border:`1px solid ${bt.color}`,borderRadius:10,padding:"2px 10px",fontSize:12,fontWeight:700,display:"inline-block",marginTop:6 }}>{bt.name}</span>

                    {/* Sync status */}
                    <div style={{ marginTop:12 }}>
                      <SyncBadge status={syncStatus} error={syncError}/>
                      {syncStatus===SYNC.ERROR && (
                        <div style={{ marginTop:6,fontSize:12,color:C.amber }}>
                          ℹ️ Data saved locally. Will retry when connection is restored.
                          <button onClick={async()=>{ setSyncStatus(SYNC.SYNCING); try{ await syncToOneDrive(submitted); setSyncStatus(SYNC.OK); }catch(e){ setSyncStatus(SYNC.ERROR); setSyncError(e.message); }}}
                            style={{ marginLeft:8,padding:"2px 8px",borderRadius:6,background:C.amber,border:"none",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer" }}>
                            Retry now
                          </button>
                        </div>
                      )}
                    </div>

                    {submitted.route.issues.length>0 ? (
                      <div style={{ marginTop:12 }}>
                        <div style={{ fontSize:12,fontWeight:700,color:rs.text,marginBottom:6 }}>Issues flagged:</div>
                        {submitted.route.issues.map((iss,i)=>(
                          <div key={i} style={{ fontSize:12,padding:"4px 10px",marginBottom:4,background:"#fff",borderRadius:8,border:`1px solid ${rs.border}`,color:rs.text }}>• {iss}</div>
                        ))}
                      </div>
                    ) : (
                      <div style={{ marginTop:10,fontSize:13,color:C.green }}>All parameters OK. Ready for charging queue. ✓</div>
                    )}
                    <div style={{ marginTop:10,fontSize:11,color:C.muted }}>{submitted.technician} · {submitted.timestamp}</div>
                  </div>
                  <div style={{ display:"flex",gap:8 }}>
                    <button style={{...btn(C.navy),flex:1}} onClick={handleNext}>⚡ Next Battery</button>
                    {["REPAIR","WARRANTY"].includes(submitted.route.category) && (
                      <button style={{...btn(C.amber),flex:1}} onClick={()=>{setRepairModal(submitted);setRepairForm({status:"Pending Repair",partsReplaced:[]})}}>🔧 Repair</button>
                    )}
                  </div>
                </div>
              );
            })()}
          </>
        )}

        {/* ══ DASHBOARD ══ */}
        {view==="dashboard" && (
          <>
            <div style={card}>
              <div style={sh}>Totals</div>
              <div style={{ display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14 }}>
                {[["OK","✅",C.green],["REPAIR","🔧",C.blue],["WARRANTY","⚠️",C.red]].map(([cat,icon,color])=>(
                  <div key={cat} style={{ background:ROUTE_S[cat].bg,borderRadius:10,padding:"12px 6px",textAlign:"center",border:`1.5px solid ${color}` }}>
                    <div style={{ fontSize:20 }}>{icon}</div>
                    <div style={{ fontSize:26,fontWeight:800,color }}>{counts[cat]}</div>
                    <div style={{ fontSize:10,color,fontWeight:600 }}>{cat}</div>
                  </div>
                ))}
              </div>
              <div style={{ textAlign:"center",fontSize:13,color:C.muted }}>Total scanned: <strong style={{color:C.text}}>{batteries.length}</strong></div>
            </div>
            <div style={card}>
              <div style={sh}>By Battery Type</div>
              {Object.entries(BATTERY_TYPES).map(([k,bt])=>(
                <div key={k} style={{ display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:9 }}>
                  <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                    <div style={{ width:10,height:10,borderRadius:"50%",background:bt.color }}/>
                    <span style={{ fontSize:14,fontWeight:600 }}>{bt.name}</span>
                  </div>
                  <div style={{ display:"flex",alignItems:"center",gap:8 }}>
                    <div style={{ height:8,borderRadius:4,background:bt.light,border:`1px solid ${bt.color}`,width:Math.max(4,(typeCounts[k]/Math.max(batteries.length,1))*140) }}/>
                    <span style={{ fontSize:14,fontWeight:700,color:bt.color,minWidth:22,textAlign:"right" }}>{typeCounts[k]}</span>
                  </div>
                </div>
              ))}
            </div>
            <div style={card}>
              <div style={sh}>Repair Pipeline</div>
              {REPAIR_STATUSES.map(s=>{const n=repairs.filter(r=>r.status===s).length;return n>0?(
                <div key={s} style={{ display:"flex",justifyContent:"space-between",fontSize:13,padding:"5px 0",borderBottom:`1px solid ${C.border}` }}>
                  <span>{s}</span><span style={{fontWeight:700,color:C.blue}}>{n}</span>
                </div>
              ):null;})}
              {repairs.length===0&&<div style={{color:C.muted,fontSize:13}}>No repairs yet.</div>}
            </div>
            <div style={{...card,background:"#E8F5E9",border:"1px solid #1B7A3E"}}>
              <div style={{ fontSize:13,fontWeight:700,color:C.green,marginBottom:6 }}>☁️ OneDrive Sync</div>
              <div style={{ fontSize:12,color:C.text,lineHeight:1.6 }}>
                {GRAPH_CONFIG.clientId==="YOUR_CLIENT_ID_HERE"
                  ? "Not configured. See SETUP_GUIDE.md to connect your OneDrive."
                  : <>Connected to: <strong>{GRAPH_CONFIG.serviceAccountEmail}</strong><br/>Files: OK → <em>{GRAPH_CONFIG.files.OK}</em> · Repair → <em>{GRAPH_CONFIG.files.REPAIR}</em> · Warranty → <em>{GRAPH_CONFIG.files.WARRANTY}</em></>}
              </div>
            </div>
            <button style={{...btn(C.red,"#fff"),width:"100%",fontSize:13}} onClick={()=>{if(window.confirm("Clear ALL local data?")) { ldb.clear(); refresh(); }}}>🗑️ Clear Local Data</button>
          </>
        )}

        {/* ══ LOG VIEW ══ */}
        {view==="log" && (
          <>
            <div style={{ display:"flex",gap:6,marginBottom:10 }}>
              {["ALL","OK","REPAIR","WARRANTY"].map(cat=>(
                <button key={cat} onClick={()=>setFilterCat(cat)} style={{ flex:1,padding:"7px 4px",borderRadius:8,cursor:"pointer",
                  border:`1.5px solid ${cat==="ALL"?C.border:ROUTE_S[cat]?.border||C.border}`,
                  background:filterCat===cat?(cat==="ALL"?C.navy:ROUTE_S[cat]?.bg||"#eee"):"#fff",
                  color:filterCat===cat?(cat==="ALL"?"#fff":ROUTE_S[cat]?.text||C.text):C.muted,
                  fontWeight:700,fontSize:11 }}>
                  {cat==="ALL"?"ALL":ROUTE_S[cat]?.icon} {cat}<br/><span style={{fontWeight:400,fontSize:10}}>{cat==="ALL"?batteries.length:counts[cat]}</span>
                </button>
              ))}
            </div>
            <input style={{...inp,fontSize:13,marginBottom:10}} placeholder="Search serial, type, technician…" value={searchQ} onChange={e=>setSearchQ(e.target.value)}/>
            {filtered.length===0&&<div style={{...card,textAlign:"center",color:C.muted}}>No entries found.</div>}
            {filtered.map(b=>{
              const rs=ROUTE_S[b.route?.category]||ROUTE_S.REPAIR, bt=BATTERY_TYPES[b.batteryKey]||{};
              return (
                <div key={b.id} style={{...card,borderLeft:`4px solid ${rs.border}`,padding:"11px 13px"}}>
                  <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start" }}>
                    <div><div style={{fontWeight:700,fontSize:14,letterSpacing:0.5}}>{b.serial}</div>
                      <div style={{fontSize:11,color:C.muted,marginTop:2}}>{b.timestamp} · {b.technician}</div></div>
                    <div style={{ textAlign:"right" }}>
                      <span style={{ background:bt.light,color:bt.color,border:`1px solid ${bt.color}`,borderRadius:8,padding:"1px 8px",fontSize:11,fontWeight:700 }}>{b.batteryType}</span>
                      <div style={{ fontSize:11,color:rs.text,fontWeight:700,marginTop:4 }}>{rs.icon} {b.route?.category}</div>
                    </div>
                  </div>
                  <div style={{ marginTop:7,display:"flex",flexWrap:"wrap",gap:"5px 12px",fontSize:12,color:C.muted }}>
                    {b.form?.cellVoltageMin&&<span>Min:{b.form.cellVoltageMin}V</span>}
                    {b.form?.cellVoltageMax&&<span>Max:{b.form.cellVoltageMax}V</span>}
                    {b.form?.packVoltage&&<span>Pack:{b.form.packVoltage}V</span>}
                    {b.form?.mosfetStatus&&<span>MOSFET:{b.form.mosfetStatus}</span>}
                  </div>
                  {b.route?.issues?.length>0&&<div style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:4}}>
                    {b.route.issues.slice(0,3).map((iss,i)=><span key={i} style={{fontSize:10,padding:"1px 7px",background:rs.bg,color:rs.text,borderRadius:8,border:`1px solid ${rs.border}`}}>{iss}</span>)}
                    {b.route.issues.length>3&&<span style={{fontSize:10,color:C.muted}}>+{b.route.issues.length-3} more</span>}
                  </div>}
                  {["REPAIR","WARRANTY"].includes(b.route?.category)&&(
                    <button onClick={()=>{setRepairModal(b);setRepairForm({status:"Pending Repair",partsReplaced:[]});}}
                      style={{marginTop:8,padding:"5px 11px",borderRadius:6,border:`1px solid ${C.blue}`,background:"#fff",color:C.blue,fontSize:12,fontWeight:600,cursor:"pointer"}}>
                      + Open Repair
                    </button>
                  )}
                </div>
              );
            })}
          </>
        )}

        {/* ══ REPAIRS VIEW ══ */}
        {view==="repairs" && (
          <>
            {repairs.length===0&&<div style={{...card,textAlign:"center",color:C.muted}}>No repairs opened yet.</div>}
            {[...repairs].reverse().map(r=>{
              const sc=r.dispatched?C.green:r.readyForDispatch?C.amber:C.blue;
              return (
                <div key={r.id} style={{...card,borderLeft:`4px solid ${sc}`}}>
                  <div style={{ display:"flex",justifyContent:"space-between",alignItems:"flex-start" }}>
                    <div><div style={{fontWeight:700,fontSize:13}}>{r.serial}</div>
                      <div style={{fontSize:11,color:C.muted}}>{r.batteryType} · {r.openedAt} · {r.openedBy}</div></div>
                    <span style={{ background:r.dispatched?"#E8F5E9":r.readyForDispatch?"#FFF8E1":"#E3F2FD",color:sc,border:`1px solid ${sc}`,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:700 }}>
                      {r.dispatched?"Dispatched":r.readyForDispatch?"Ready to Dispatch":r.status}
                    </span>
                  </div>
                  {r.partsReplaced?.length>0&&<div style={{fontSize:12,marginTop:5,color:C.muted}}>Parts: {r.partsReplaced.join(", ")}</div>}
                  {r.postRepairNotes&&<div style={{fontSize:12,marginTop:3,color:C.text}}>Notes: {r.postRepairNotes}</div>}
                  {r.dispatched&&r.dispatchDate&&<div style={{fontSize:12,marginTop:3,color:C.green}}>Dispatched: {r.dispatchDate}</div>}
                  {!r.dispatched&&(
                    <div style={{ marginTop:9,display:"flex",flexWrap:"wrap",gap:6 }}>
                      {REPAIR_STATUSES.filter(s=>s!=="Dispatched"&&s!=="Ready for Dispatch").map(s=>(
                        <button key={s} onClick={()=>{ldb.updateRepair(r.id,{status:s});refresh();}}
                          style={{ padding:"4px 9px",borderRadius:6,border:`1px solid ${r.status===s?C.blue:C.border}`,background:r.status===s?C.blue:"#fff",color:r.status===s?"#fff":C.muted,fontSize:11,fontWeight:600,cursor:"pointer" }}>
                          {s}
                        </button>
                      ))}
                      {r.status==="Post-Repair Check"&&<button onClick={()=>{ldb.updateRepair(r.id,{onlineStatus:"Online",status:"Online Test"});refresh();}} style={{ padding:"4px 9px",borderRadius:6,background:C.green,border:"none",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer" }}>✓ Online</button>}
                      {r.onlineStatus==="Online"&&!r.readyForDispatch&&<button onClick={()=>{ldb.updateRepair(r.id,{readyForDispatch:true,status:"Ready for Dispatch"});refresh();}} style={{ padding:"4px 9px",borderRadius:6,background:C.amber,border:"none",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer" }}>✓ Ready</button>}
                      {r.readyForDispatch&&!r.dispatched&&<button onClick={()=>{ldb.updateRepair(r.id,{dispatched:true,dispatchDate:new Date().toLocaleString("en-UG"),status:"Dispatched"});refresh();}} style={{ padding:"4px 9px",borderRadius:6,background:C.green,border:"none",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer" }}>🚀 Dispatch</button>}
                    </div>
                  )}
                </div>
              );
            })}
          </>
        )}

        {/* ══ SEARCH VIEW ══ */}
        {view==="search" && (
          <>
            <div style={card}>
              <div style={sh}>Search Battery by Serial</div>
              <input style={inp} placeholder="Type serial or battery type…" value={searchQ} onChange={e=>setSearchQ(e.target.value)} autoFocus/>
            </div>
            {searchQ.length>2&&(()=>{
              const q=searchQ.toUpperCase();
              const found=batteries.filter(b=>b.serial?.includes(q)||b.batteryType?.toUpperCase().includes(q));
              return found.length===0
                ? <div style={{...card,textAlign:"center",color:C.muted}}>No battery found for "{searchQ}"</div>
                : found.map(b=>{
                    const rs=ROUTE_S[b.route?.category]||ROUTE_S.REPAIR;
                    const bReps=repairs.filter(r=>r.serial===b.serial);
                    return (
                      <div key={b.id} style={{...card,borderLeft:`4px solid ${rs.border}`}}>
                        <div style={{fontWeight:700,fontSize:15,letterSpacing:1}}>{b.serial}</div>
                        <div style={{fontSize:12,color:C.muted}}>{b.batteryType} · {b.timestamp} · {b.technician}</div>
                        <span style={{background:rs.bg,color:rs.text,border:`1px solid ${rs.border}`,borderRadius:8,padding:"2px 10px",fontSize:12,fontWeight:700,display:"inline-block",marginTop:6}}>{rs.icon} {rs.label}</span>
                        <div style={{ marginTop:10,display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6,fontSize:12 }}>
                          {[["Cell Min",b.form?.cellVoltageMin,"V"],["Cell Max",b.form?.cellVoltageMax,"V"],["Pack",b.form?.packVoltage,"V"],
                            ["MOSFET",b.form?.mosfetStatus,""],["Charge",b.form?.canCharge,""],["Discharge",b.form?.canDischarge,""]].map(([l,v,u])=>(
                            <div key={l}><span style={{color:C.muted,fontSize:11}}>{l}</span><br/><strong>{v||"—"}{v?u:""}</strong></div>
                          ))}
                        </div>
                        {b.route?.issues?.length>0&&<div style={{marginTop:8}}>{b.route.issues.map((iss,i)=><div key={i} style={{fontSize:11,color:rs.text}}>• {iss}</div>)}</div>}
                        {bReps.length>0&&(
                          <div style={{marginTop:10}}>
                            <div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:4}}>Repair History ({bReps.length})</div>
                            {bReps.map(r=>(
                              <div key={r.id} style={{fontSize:12,padding:"5px 8px",background:C.surface,borderRadius:6,marginBottom:4}}>
                                <strong>{r.id}</strong> · {r.status} · {r.openedAt}
                                {r.dispatched&&<span style={{color:C.green}}> · Dispatched {r.dispatchDate}</span>}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  });
            })()}
          </>
        )}
      </div>

      {/* Bottom Nav */}
      <div style={{ position:"fixed",bottom:0,left:0,right:0,background:C.navy,display:"flex",borderTop:"1px solid rgba(255,255,255,0.08)",zIndex:100 }}>
        {navItems.map(n=>(
          <button key={n.id} onClick={()=>setView(n.id)} style={{ flex:1,padding:"9px 4px 7px",border:"none",background:"transparent",
            color:view===n.id?"#42A5F5":"rgba(255,255,255,0.4)",cursor:"pointer",fontSize:10,fontWeight:600,
            display:"flex",flexDirection:"column",alignItems:"center",gap:3 }}>
            <span style={{fontSize:19}}>{n.icon}</span>{n.label}
          </button>
        ))}
      </div>

      {/* Repair Modal */}
      {repairModal&&(
        <div style={{ position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:500,display:"flex",alignItems:"flex-end" }}>
          <div style={{ background:"#fff",borderRadius:"16px 16px 0 0",padding:"20px 15px",width:"100%",maxHeight:"85vh",overflowY:"auto",boxSizing:"border-box" }}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:3}}>Repair: {repairModal.serial}</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:14}}>{repairModal.batteryType} · {repairModal.route?.category}</div>
            <label style={fl}>Status</label>
            <select style={{...inp,marginBottom:12}} value={repairForm.status||""} onChange={e=>setRepairForm(p=>({...p,status:e.target.value}))}>
              {REPAIR_STATUSES.map(s=><option key={s}>{s}</option>)}
            </select>
            <label style={fl}>Parts to Replace</label>
            <CheckGroup options={MISSING_PARTS} value={repairForm.partsReplaced||[]} onChange={v=>setRepairForm(p=>({...p,partsReplaced:v}))}/>
            <label style={{...fl,marginTop:14}}>Post-Repair Notes / Condition</label>
            <textarea rows={3} style={{...inp,resize:"vertical"}} placeholder="Condition after repair, test results…"
              value={repairForm.postRepairNotes||""} onChange={e=>setRepairForm(p=>({...p,postRepairNotes:e.target.value}))}/>
            <label style={{...fl,marginTop:12}}>Online / Offline Status</label>
            <select style={{...inp,marginBottom:14}} value={repairForm.onlineStatus||""} onChange={e=>setRepairForm(p=>({...p,onlineStatus:e.target.value}))}>
              <option value="">— select —</option>
              <option>Online – Charging</option><option>Online – Discharging</option>
              <option>Offline – Pending retest</option><option>Offline – Failed</option>
            </select>
            <div style={{ display:"flex",gap:10 }}>
              <button style={{...btn(C.blue),flex:1}} onClick={()=>{ldb.addRepair({id:`REP-${Date.now()}`,serial:repairModal.serial,batteryType:repairModal.batteryType,openedBy:tech?.name||"Unknown",openedAt:new Date().toLocaleString("en-UG"),...repairForm});refresh();setRepairModal(null);}}>Save Repair</button>
              <button style={{...btn("#fff",C.muted,C.border),flex:1}} onClick={()=>setRepairModal(null)}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ── Mount the app ─────────────────────────────────────────────────────────────
const container = document.getElementById("root");
const root = ReactDOM.createRoot(container);
root.render(React.createElement(SpiroBatteryQMS));
