import { useState, useRef, useCallback } from "react";

const M365 = [
  { id:"none",      label:"None",                   price:0,    win:false, ems:false },
  { id:"f1",        label:"M365 Frontline F1",       price:2.25, win:false, ems:false },
  { id:"f3",        label:"M365 Frontline F3",       price:8,    win:true,  ems:false },
  { id:"o365e1",    label:"Office 365 E1",           price:10,   win:false, ems:false },
  { id:"biz-basic", label:"M365 Business Basic",     price:7,    win:false, ems:false },
  { id:"biz-std",   label:"M365 Business Standard",  price:12.5, win:false, ems:false },
  { id:"biz-prem",  label:"M365 Business Premium",   price:22,   win:false, ems:true  },
  { id:"o365e3",    label:"Office 365 E3",           price:23,   win:false, ems:false },
  { id:"m365e3",    label:"M365 E3",                 price:36,   win:true,  ems:true  },
  { id:"m365e5",    label:"M365 E5",                 price:57,   win:true,  ems:true, phone:true },
];

const ADDONS = [
  { id:"none",           label:"None",                       price:0   },
  { id:"ems-e3",         label:"EMS E3",                     price:8.8 },
  { id:"ems-e5",         label:"EMS E5",                     price:14.8},
  { id:"m365-e5-sec",    label:"M365 E5 Security",           price:12  },
  { id:"m365-e5-comp",   label:"M365 E5 Compliance",         price:12  },
  { id:"intune-p1",      label:"Intune Plan 1",              price:8   },
  { id:"powerbi-pro",    label:"Power BI Pro",               price:10  },
  { id:"pa-trial",       label:"Power Apps P2 Trial",        price:0,  isTrial:true },
  { id:"fabric-free",    label:"Microsoft Fabric (Free)",    price:0   },
  { id:"pa-free",        label:"Power Automate Free",        price:0   },
  { id:"extra-features", label:"M365 E3 Extra Features",     price:0   },
];

const COPILOT = [
  { id:"none",              label:"None",                       price:0  },
  { id:"copilot-chat",      label:"Copilot Chat (Free)",        price:0  },
  { id:"copilot-biz-promo", label:"Copilot Business Promo",     price:18 },
  { id:"copilot-biz",       label:"Copilot Business",           price:21 },
  { id:"copilot-ent",       label:"M365 Copilot Enterprise",    price:30 },
];

const TEAMS_PHONE = [
  { id:"none",        label:"None",                            price:0  },
  { id:"tp-std",      label:"Teams Phone Standard",            price:10 },
  { id:"tp-payg",     label:"Teams Phone + Pay-As-You-Go",     price:13 },
  { id:"tp-domestic", label:"Teams Phone + Domestic Calling",  price:17 },
  { id:"tp-intl",     label:"Teams Phone + Intl Calling",      price:34 },
];

const D365 = {
  team:       { id:"team",       label:"Team Member",               price:8,   ap:null },
  "bc-ess":   { id:"bc-ess",     label:"BC Essentials",             price:80,  ap:20   },
  "bc-prem":  { id:"bc-prem",    label:"BC Premium",                price:110, ap:20   },
  "sales-pro":{ id:"sales-pro",  label:"Sales Professional",        price:65,  ap:20   },
  "sales-ent":{ id:"sales-ent",  label:"Sales Enterprise",          price:105, ap:20   },
  "cs-pro":   { id:"cs-pro",     label:"Customer Service Pro",      price:50,  ap:20   },
  "cs-ent":   { id:"cs-ent",     label:"Customer Service Ent",      price:95,  ap:20   },
  field:      { id:"field",      label:"Field Service",             price:95,  ap:20   },
  finance:    { id:"finance",    label:"D365 Finance",              price:180, ap:30   },
  scm:        { id:"scm",        label:"Supply Chain Mgmt",         price:180, ap:30   },
  hr:         { id:"hr",         label:"Human Resources",           price:120, ap:30   },
  proj:       { id:"proj",       label:"Project Operations",        price:120, ap:30   },
  commerce:   { id:"commerce",   label:"Commerce",                  price:180, ap:30   },
};

const ATTACH = {
  "bc-ess":   ["cs-pro","cs-ent","field","sales-pro","sales-ent","proj","hr","commerce"],
  "bc-prem":  ["cs-pro","cs-ent","field","sales-pro","sales-ent","proj","hr","commerce"],
  "sales-pro":["cs-pro","cs-ent","field"],
  "sales-ent":["cs-pro","cs-ent","field","finance","hr","proj","commerce","bc-ess","bc-prem"],
  "cs-pro":   ["sales-pro","sales-ent","field"],
  "cs-ent":   ["sales-ent","field","finance","hr","proj","commerce","bc-ess","bc-prem"],
  "field":    ["sales-ent","cs-ent","finance","hr","proj"],
  finance:    ["scm","sales-ent","sales-pro","cs-ent","cs-pro","field","hr","proj","commerce","bc-ess","bc-prem"],
  scm:        ["finance","sales-ent","sales-pro","cs-ent","cs-pro","field","hr","proj","commerce"],
  hr:         ["finance","scm","sales-ent","cs-ent","field"],
  proj:       ["cs-ent","field","finance","hr","scm"],
  commerce:   ["finance","scm","cs-ent","field","sales-ent"],
};

const gm = (list, id) => (Array.isArray(list) ? list : Object.values(list)).find(p => p.id === id) || (Array.isArray(list) ? list[0] : Object.values(list)[0]);

function optimizeD365(ids) {
  if (!ids || !ids.length) return null;
  if (ids.length === 1 && ids[0] === "team") return { rows:[{id:"team",role:"team",price:8}], total:8, sa:8, saves:0, note:"Team Member is correct for read-only/light users." };
  const full = ids.filter(a => a !== "team");
  const hasTeam = ids.includes("team");
  if (!full.length) return null;
  const sa = full.reduce((s,a) => s + (D365[a]?.price||0), 0);
  let best = null;
  full.forEach(cb => {
    const rest = full.filter(a => a !== cb);
    const q = ATTACH[cb] || [];
    let cost = D365[cb]?.price || 0;
    const rows = [{ id:cb, role:"base", price:D365[cb]?.price||0 }];
    rest.forEach(a => {
      const can = q.includes(a);
      const p = can ? (D365[a]?.ap || D365[a]?.price || 0) : (D365[a]?.price || 0);
      cost += p;
      rows.push({ id:a, role:can?"attach":"full", price:p, sp:D365[a]?.price||0 });
    });
    if (!best || cost < best.cost) best = { cost, rows, base:cb };
  });
  if (hasTeam) { best.rows.push({id:"team",role:"team",price:8}); best.cost+=8; }
  const saves = Math.max(0, sa - (best.cost - (hasTeam?8:0)));
  let note = "Base: " + (D365[best.base]?.label||best.base) + " at $" + (D365[best.base]?.price||0) + "/mo (full — highest-cost app must be base). ";
  const att = best.rows.filter(r => r.role === "attach");
  if (att.length) note += att.map(r => (D365[r.id]?.label||r.id) + " attached at $" + r.price + " vs $" + r.sp + " standalone").join("; ") + ". ";
  if (saves > 0) note += "Saves $" + saves.toFixed(0) + "/user/mo vs all-standalone.";
  return { rows:best.rows, total:best.cost, sa:sa+(hasTeam?8:0), saves, note, base:best.base };
}

function getWin(m365id, hasExisting) {
  const p = M365.find(x => x.id === m365id) || M365[0];
  if (hasExisting && p.win) return { risk:"LOW", label:"Dual coverage", note:"Both your M365 plan and an existing Windows license cover this user. Verify the existing license is still needed." };
  if (hasExisting) return { risk:"NONE", label:"Covered — existing license", note:"Covered by existing Windows OEM/SA/VL. OEM licenses are hardware-tied and non-transferable." };
  if (p.win) return { risk:"NONE", label:"Covered via M365", note:"Windows 11 Enterprise included in " + p.label + "." };
  return { risk:"HIGH", label:"UNLICENSED — Windows gap", note:p.label + " does NOT include Windows Enterprise. If devices run Win 10/11 Enterprise they are unlicensed. Fix: upgrade to M365 E3 ($36) which adds Windows + Intune + Azure AD P1, or add Windows Enterprise SA (~$14/user/mo)." };
}

function getM365Rec(m365, addon) {
  if (m365 === "o365e3" && addon === "ems-e3") return { action:"Consolidate to M365 E3", why:"O365 E3 ($23) + EMS E3 ($8.80) = $31.80/user but still no Windows Enterprise. M365 E3 at $36 gives everything plus Windows 11 Enterprise in one SKU.", sev:"danger" };
  if ((m365==="m365e3"||m365==="m365e5"||m365==="biz-prem") && addon==="ems-e3") return { action:"Remove redundant EMS E3", why:"Your M365 plan already includes all EMS E3 features. You are paying $8.80/user/mo for nothing. Remove immediately.", sev:"danger" };
  if (m365 === "m365e5") return { action:"Verify E5 justification", why:"M365 E5 ($57) only justified for Teams Phone, Power BI, Defender P2, or Azure AD P2 users. Otherwise: E3 + E5 Security add-on = $48/user/mo.", sev:"warn" };
  if (m365 === "o365e3" && addon === "none") return { action:"Security gap — no Intune or Azure AD P1", why:"O365 E3 has no device management and no MFA policies. Upgrade to M365 E3 (+$13/user/mo) to close this gap.", sev:"warn" };
  return { action:"M365 looks optimized", why:"No redundancies or gaps detected.", sev:"ok" };
}

const RISK_COLOR = { HIGH:"#f97316", MEDIUM:"#d97706", LOW:"#9ca3af", NONE:"#16a34a" };
const RISK_BG    = { HIGH:"#fff7ed", MEDIUM:"#fffbeb", LOW:"#f9fafb", NONE:"#f0fdf4" };

const css = {
  app:  { fontFamily:"'Segoe UI',system-ui,sans-serif", background:"#f4f4f2", minHeight:"100vh" },
  hdr:  { background:"#fff", borderBottom:"1px solid #eee", padding:"11px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", position:"sticky", top:0, zIndex:20 },
  card: { background:"#fff", borderRadius:12, border:"0.5px solid #e5e5e5", padding:"1.2rem", marginBottom:12 },
  stl:  { fontSize:10, fontWeight:700, color:"#aaa", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:7 },
  inp:  { fontSize:12, height:31, padding:"0 8px", border:"1px solid #ddd", borderRadius:7, background:"#fff", color:"#111", width:"100%", boxSizing:"border-box" },
  sel:  { fontSize:12, height:31, padding:"0 5px", border:"1px solid #ddd", borderRadius:7, background:"#fff", color:"#111", width:"100%", boxSizing:"border-box" },
  btn:  (v) => ({ fontSize:13, fontWeight:600, padding:"7px 13px", borderRadius:8, cursor:"pointer", display:"inline-flex", alignItems:"center", gap:5, border: v==="outline"?"1px solid #ddd":"none", background: v==="primary"?"#111":v==="green"?"#16a34a":v==="outline"?"#fff":"#f3f4f6", color:(v==="primary"||v==="green")?"#fff":"#374151" }),
  badge:(lv) => ({ fontSize:10, fontWeight:700, padding:"2px 7px", borderRadius:5, background:RISK_COLOR[lv]||"#888", color:"#fff", display:"inline-block" }),
  note: { fontSize:11, color:"#888", lineHeight:1.5 },
};

const D365_GROUPS = [
  { label:"ERP — Finance & Ops",  apps:["finance","scm","commerce","hr","proj"] },
  { label:"CRM — Sales & Service", apps:["sales-pro","sales-ent","cs-pro","cs-ent","field"] },
  { label:"SMB ERP",               apps:["bc-ess","bc-prem"] },
  { label:"Light access",          apps:["team"] },
];

function D365Picker({ selected, onChange }) {
  const toggle = id => onChange(selected.includes(id) ? selected.filter(a => a !== id) : [...selected, id]);
  return (
    <div>
      {D365_GROUPS.map(g => (
        <div key={g.label} style={{ marginBottom:7 }}>
          <div style={{ fontSize:10, fontWeight:700, color:"#bbb", textTransform:"uppercase", marginBottom:4 }}>{g.label}</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
            {g.apps.map(id => {
              const info = D365[id];
              const on = selected.includes(id);
              return (
                <button key={id} onClick={() => toggle(id)} style={{ fontSize:11, padding:"3px 9px", borderRadius:6, border:"1px solid "+(on?"#2563eb":"#ddd"), background:on?"#eff6ff":"#fafaf9", color:on?"#1d4ed8":"#555", cursor:"pointer", fontWeight:on?700:400 }}>
                  {info?.label} <span style={{ color:on?"#60a5fa":"#aaa" }}>${info?.price}</span>
                </button>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function Scanner({ onDetected }) {
  const [drag, setDrag] = useState(false);
  const [scanning, setScanning] = useState(false);
  const [result, setResult] = useState(null);
  const [preview, setPreview] = useState(null);
  const [err, setErr] = useState(null);
  const fileRef = useRef();

  const scan = useCallback(async (file) => {
    if (!file?.type.startsWith("image/")) { setErr("Please upload an image file."); return; }
    setErr(null); setResult(null);
    const reader = new FileReader();
    reader.onload = async (e) => {
      setPreview(e.target.result);
      setScanning(true);
      try {
        const res = await fetch("https://api.anthropic.com/v1/messages", {
          method:"POST",
          headers:{ "Content-Type":"application/json" },
          body: JSON.stringify({
            model:"claude-sonnet-4-20250514",
            max_tokens:1200,
            messages:[{ role:"user", content:[
              { type:"image", source:{ type:"base64", media_type:file.type, data:e.target.result.split(",")[1] } },
              { type:"text", text:`You are a Microsoft licensing expert. Analyze this screenshot and return ONLY a raw JSON array (no markdown, no backticks).

Each object in the array:
- name: descriptive role name
- m365: one of: none,f1,f3,o365e1,biz-basic,biz-std,biz-prem,o365e3,m365e3,m365e5
- addon: one of: none,ems-e3,ems-e5,m365-e5-sec,m365-e5-comp,intune-p1,powerbi-pro,pa-trial,fabric-free,pa-free,extra-features
- copilot: one of: none,copilot-chat,copilot-biz-promo,copilot-biz,copilot-ent
- teamsPhone: one of: none,tp-std,tp-payg,tp-domestic,tp-intl
- d365Apps: array from: team,bc-ess,bc-prem,sales-pro,sales-ent,cs-pro,cs-ent,field,finance,scm,hr,proj,commerce
- hasExistingWindows: true if Windows SA or standalone Windows Enterprise visible
- seats: number (default 1)
- isTrial: true if Trial visible
- notes: one sentence

Mappings: "Office 365 E3"=o365e3, "Microsoft 365 E3"=m365e3, "EMS E3"=ems-e3, "Supply Chain"=scm, "Team Member"=team, "Extra Features"=extra-features, "Fabric (Free)"=fabric-free, "Power Automate Free"=pa-free.

Return only the JSON array.` }
            ]}]
          })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error.message);
        const raw = data.content?.find(b => b.type === "text")?.text || "";
        const parsed = JSON.parse(raw.replace(/```json|```/g,"").trim());
        if (!Array.isArray(parsed) || !parsed.length) throw new Error("No licenses detected.");
        setResult(parsed);
      } catch(ex) {
        setErr(ex.message || "Scan failed.");
      } finally {
        setScanning(false);
      }
    };
    reader.readAsDataURL(file);
  }, []);

  return (
    <div style={css.card}>
      <div style={css.stl}>AI License Scanner</div>
      <p style={{ fontSize:12, color:"#555", marginBottom:10, lineHeight:1.6 }}>
        Drop a screenshot from the M365 Admin Center. Claude reads M365, EMS, D365, Windows, and Teams Phone licenses automatically.
      </p>
      <div
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); scan(e.dataTransfer.files[0]); }}
        onClick={() => fileRef.current?.click()}
        style={{ border:"2px dashed "+(drag?"#2563eb":"#ddd"), borderRadius:10, padding:"1.5rem", textAlign:"center", cursor:"pointer", background:drag?"#eff6ff":"#fafaf9", marginBottom:10, transition:"all 0.2s" }}
      >
        <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={e => scan(e.target.files[0])} />
        {preview
          ? <div><img src={preview} alt="" style={{ maxWidth:"100%", maxHeight:140, borderRadius:8, objectFit:"contain" }} /><p style={{ fontSize:11, color:"#aaa", marginTop:5 }}>Click to replace</p></div>
          : <div><div style={{ fontSize:26, marginBottom:6 }}>📋</div><div style={{ fontSize:13, fontWeight:600, color:"#374151" }}>Drop screenshot here</div><div style={{ fontSize:11, color:"#aaa", marginTop:3 }}>Reads M365 · D365 · Windows · Teams Phone</div></div>
        }
      </div>
      {scanning && <div style={{ padding:"9px 11px", background:"#eff6ff", borderRadius:8, fontSize:12, color:"#1e40af", marginBottom:8, display:"flex", gap:7, alignItems:"center" }}><span style={{ animation:"spin 1s linear infinite", display:"inline-block" }}>⟳</span> Claude is reading the screenshot...</div>}
      {err && <div style={{ padding:"9px 11px", background:"#fef2f2", borderRadius:8, fontSize:12, color:"#dc2626", marginBottom:8, border:"0.5px solid #fca5a5" }}>⚠ {err}</div>}
      {result && (
        <div>
          <div style={{ padding:"9px 11px", background:"#f0fdf4", borderRadius:8, border:"0.5px solid #86efac", marginBottom:8 }}>
            <div style={{ fontSize:12, fontWeight:700, color:"#15803d", marginBottom:4 }}>Detected {result.length} group{result.length !== 1 ? "s" : ""}</div>
            {result.map((r,i) => (
              <div key={i} style={{ fontSize:11, color:"#166534", padding:"2px 0", borderBottom:i<result.length-1?"0.5px solid #bbf7d0":"none" }}>
                <strong>{r.name}</strong> — {gm(M365,r.m365).label}
                {r.d365Apps?.length ? " · D365: " + r.d365Apps.map(a => D365[a]?.label||a).join(", ") : ""}
                {r.hasExistingWindows ? " · Windows license" : ""}
                {r.isTrial ? " ⚠ TRIAL" : ""}
              </div>
            ))}
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <button style={css.btn("green")} onClick={() => { onDetected(result); setResult(null); setPreview(null); }}>
              Add {result.length} row{result.length !== 1 ? "s" : ""} to optimizer
            </button>
            <button style={css.btn("outline")} onClick={() => { setResult(null); setPreview(null); }}>Discard</button>
          </div>
        </div>
      )}
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function RowEditor({ rows, setRows }) {
  const [openD365, setOpenD365] = useState({});
  const update = (id, f, v) => setRows(p => p.map(r => r.id === id ? { ...r, [f]: f==="seats" ? Math.max(1,parseInt(v)||1) : v } : r));
  const remove = id => setRows(p => p.filter(r => r.id !== id));
  const add = () => setRows(p => [...p, { id:Date.now(), name:"", m365:"biz-basic", addon:"none", copilot:"none", teamsPhone:"none", d365Apps:[], hasExistingWindows:false, seats:1, isTrial:false }]);

  if (!rows.length) return (
    <div style={{ textAlign:"center", padding:"1.5rem", color:"#aaa", fontSize:13, border:"1px dashed #e5e5e5", borderRadius:10, marginBottom:8 }}>
      No employees added. Upload a screenshot or click "+ Add group".
    </div>
  );

  return (
    <div>
      {rows.map(r => (
        <div key={r.id} style={{ border:"0.5px solid #eee", borderRadius:10, marginBottom:10, overflow:"hidden" }}>
          <div style={{ display:"grid", gridTemplateColumns:"1.3fr 1.2fr 1.1fr 1fr 56px 30px", gap:7, padding:"9px 11px", background:"#fafaf9", borderBottom:"0.5px solid #eee", alignItems:"center" }}>
            <input style={css.inp} value={r.name} placeholder="Name / role" onChange={e => update(r.id,"name",e.target.value)} />
            <select style={css.sel} value={r.m365} onChange={e => update(r.id,"m365",e.target.value)}>
              {M365.map(p => <option key={p.id} value={p.id}>{p.label}{p.price > 0 ? " $"+p.price : ""}</option>)}
            </select>
            <select style={css.sel} value={r.addon} onChange={e => update(r.id,"addon",e.target.value)}>
              {ADDONS.map(p => <option key={p.id} value={p.id}>{p.label}{p.price > 0 ? " $"+p.price : p.id !== "none" ? " Free" : ""}</option>)}
            </select>
            <select style={css.sel} value={r.teamsPhone||"none"} onChange={e => update(r.id,"teamsPhone",e.target.value)}>
              {TEAMS_PHONE.map(p => <option key={p.id} value={p.id}>{p.label}{p.price > 0 ? " $"+p.price : ""}</option>)}
            </select>
            <input type="number" min="1" style={{ ...css.inp, textAlign:"center" }} value={r.seats} onChange={e => update(r.id,"seats",e.target.value)} />
            <button onClick={() => remove(r.id)} style={{ width:28, height:28, borderRadius:7, border:"1px solid #eee", background:"transparent", color:"#aaa", cursor:"pointer", fontSize:15 }}>×</button>
          </div>
          <div style={{ padding:"8px 11px", background:"#fff" }}>
            <div style={{ display:"flex", gap:14, marginBottom:7, flexWrap:"wrap" }}>
              <label style={{ display:"flex", alignItems:"center", gap:5, fontSize:12, color:"#555", cursor:"pointer" }}>
                <input type="checkbox" checked={r.hasExistingWindows||false} onChange={e => update(r.id,"hasExistingWindows",e.target.checked)} style={{ width:13, height:13, accentColor:"#2563eb" }} />
                Existing Windows license
              </label>
              <label style={{ display:"flex", alignItems:"center", gap:5, fontSize:12, color:"#555", cursor:"pointer" }}>
                <input type="checkbox" checked={r.isTrial||false} onChange={e => update(r.id,"isTrial",e.target.checked)} style={{ width:13, height:13, accentColor:"#d97706" }} />
                Trial / expiring
              </label>
              <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                <span style={{ fontSize:12, color:"#555" }}>Copilot:</span>
                <select style={{ ...css.sel, width:"auto", minWidth:170 }} value={r.copilot||"none"} onChange={e => update(r.id,"copilot",e.target.value)}>
                  {COPILOT.map(p => <option key={p.id} value={p.id}>{p.label}{p.price > 0 ? " $"+p.price : ""}</option>)}
                </select>
              </div>
            </div>
            <button onClick={() => setOpenD365(p => ({ ...p, [r.id]:!p[r.id] }))} style={{ fontSize:11, color:"#2563eb", background:"none", border:"1px dashed #93c5fd", borderRadius:6, padding:"3px 9px", cursor:"pointer", marginBottom:openD365[r.id]?8:0 }}>
              {openD365[r.id] ? "▲ Hide" : "▼ Add"} D365 apps {r.d365Apps?.length ? "("+r.d365Apps.length+" selected)" : ""}
            </button>
            {openD365[r.id] && <D365Picker selected={r.d365Apps||[]} onChange={v => update(r.id,"d365Apps",v)} />}
            {r.d365Apps?.length > 0 && !openD365[r.id] && (
              <div style={{ fontSize:11, color:"#555", marginTop:3 }}>D365: {r.d365Apps.map(a => D365[a]?.label||a).join(", ")}</div>
            )}
          </div>
        </div>
      ))}
      <button style={{ ...css.btn("outline"), border:"1px dashed #93c5fd", color:"#2563eb", background:"none" }} onClick={add}>+ Add group</button>
    </div>
  );
}

function ResultCard({ row }) {
  const base   = gm(M365, row.m365);
  const addon  = gm(ADDONS, row.addon);
  const cop    = gm(COPILOT, row.copilot||"none");
  const phone  = gm(TEAMS_PHONE, row.teamsPhone||"none");
  const m365r  = getM365Rec(row.m365, row.addon);
  const win    = getWin(row.m365, row.hasExistingWindows);
  const d365r  = row.d365Apps?.length ? optimizeD365(row.d365Apps) : null;

  const m365Cost = (base.price + addon.price + cop.price + phone.price) * row.seats;
  const d365Cost = (d365r?.total||0) * row.seats;
  const d365SA   = (d365r?.sa||0) * row.seats;
  const total    = m365Cost + d365Cost;

  const topRisk = win.risk==="HIGH" ? "HIGH" : m365r.sev==="danger" ? "HIGH" : m365r.sev==="warn" ? "MEDIUM" : "NONE";

  const sevColor = { danger:"#dc2626", warn:"#d97706", ok:"#16a34a" };
  const sevBg    = { danger:"#fff8f8", warn:"#fffcf0", ok:"#f0fdf4" };

  return (
    <div style={{ border:"0.5px solid #e5e5e5", borderRadius:12, marginBottom:12, overflow:"hidden" }}>
      <div style={{ background:"#fafaf9", borderBottom:"0.5px solid #eee", padding:"9px 13px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div>
          <span style={{ fontWeight:700, fontSize:14 }}>{row.name||"Unnamed"}</span>
          <span style={{ fontSize:12, color:"#888", marginLeft:8 }}>{row.seats} seat{row.seats!==1?"s":""}</span>
          {row.isTrial && <span style={{ marginLeft:6, ...css.badge("MEDIUM"), fontSize:9 }}>TRIAL</span>}
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:13, fontWeight:700 }}>{"$"+total.toFixed(2)+"/mo"}</span>
          <span style={css.badge(topRisk)}>{topRisk}</span>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr" }}>

        {/* Col 1: M365 stack */}
        <div style={{ padding:"11px 13px", borderRight:"0.5px solid #eee" }}>
          <div style={css.stl}>Current M365 Stack</div>
          <div style={{ fontSize:12 }}>{base.label} <span style={css.note}>{"$"+base.price+"/mo"}</span></div>
          {addon.id!=="none" && <div style={{ fontSize:12, marginTop:3 }}>{addon.label} <span style={{ ...css.note, color:addon.isTrial?"#d97706":"#888" }}>{addon.price>0?"$"+addon.price+"/mo":"Free"}{addon.isTrial?" ⚠":""}</span></div>}
          {cop.id!=="none" && <div style={{ fontSize:12, marginTop:3 }}>{cop.label} <span style={css.note}>{cop.price>0?"$"+cop.price+"/mo":"Free"}</span></div>}
          {phone.id!=="none" && <div style={{ fontSize:12, marginTop:3 }}>{phone.label} <span style={css.note}>{"$"+phone.price+"/mo"}</span></div>}
          <div style={{ marginTop:8, padding:"6px 8px", background:sevBg[m365r.sev], borderRadius:6, borderLeft:"2px solid "+sevColor[m365r.sev] }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#111", marginBottom:2 }}>{m365r.action}</div>
            <div style={{ fontSize:11, color:"#555", lineHeight:1.5 }}>{m365r.why}</div>
          </div>
        </div>

        {/* Col 2: Windows + D365 */}
        <div style={{ padding:"11px 13px", borderRight:"0.5px solid #eee" }}>
          <div style={css.stl}>Windows License</div>
          <div style={{ padding:"5px 8px", background:RISK_BG[win.risk]||"#fff", borderRadius:6, borderLeft:"2px solid "+(RISK_COLOR[win.risk]||"#ccc"), marginBottom:10 }}>
            <div style={{ fontSize:11, fontWeight:700, color:RISK_COLOR[win.risk]||"#888", marginBottom:2 }}>{win.label}</div>
            <div style={{ fontSize:11, color:"#555", lineHeight:1.5 }}>{win.note}</div>
          </div>

          {d365r && (
            <>
              <div style={css.stl}>D365 Optimized ({row.seats} seat{row.seats!==1?"s":""})</div>
              {d365r.rows.map((s,i) => (
                <div key={i} style={{ fontSize:12, display:"flex", justifyContent:"space-between", padding:"2px 0", borderBottom:i<d365r.rows.length-1?"0.5px solid #f0f0ee":"none" }}>
                  <span>
                    <span style={{ color:s.role==="base"?"#111":s.role==="attach"?"#2563eb":"#888" }}>{D365[s.id]?.label||s.id}</span>
                    <span style={{ fontSize:10, color:"#aaa", fontStyle:"italic", marginLeft:4 }}>{s.role}</span>
                    {s.role==="attach" && <span style={{ fontSize:10, color:"#16a34a", marginLeft:4 }}>{"saves $"+(s.sp-s.price)+"/mo"}</span>}
                  </span>
                  <span style={{ fontWeight:600 }}>{"$"+s.price+"/mo"}</span>
                </div>
              ))}
              {d365r.saves > 0 && <div style={{ fontSize:11, color:"#16a34a", fontWeight:700, marginTop:5 }}>{"Saves $"+(d365r.saves*row.seats).toFixed(0)+"/mo total"}</div>}
            </>
          )}
          {!d365r && <div style={{ fontSize:12, color:"#aaa" }}>No D365 apps selected</div>}
        </div>

        {/* Col 3: Audit + cost */}
        <div style={{ padding:"11px 13px", background:RISK_BG[topRisk]||"#fff" }}>
          <div style={css.stl}>Audit Risk</div>
          {win.risk!=="NONE" && (
            <div style={{ marginBottom:8 }}>
              <span style={css.badge(win.risk)}>{"Windows: "+win.risk}</span>
              <div style={{ fontSize:11, color:"#555", marginTop:4, lineHeight:1.5 }}>{win.note}</div>
            </div>
          )}
          {d365r && (
            <div style={{ padding:"7px 9px", background:"#fff", borderRadius:6, border:"0.5px solid #e5e5e5", marginBottom:8 }}>
              <div style={{ fontSize:11, fontWeight:700, color:"#111", marginBottom:3 }}>D365 Attach Analysis</div>
              <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, marginBottom:2 }}>
                <span style={{ color:"#888" }}>All standalone:</span>
                <span style={{ color:"#dc2626", fontWeight:600 }}>{"$"+d365SA.toFixed(2)+"/mo"}</span>
              </div>
              <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, marginBottom:2 }}>
                <span style={{ color:"#888" }}>Optimized:</span>
                <span style={{ color:"#16a34a", fontWeight:600 }}>{"$"+d365Cost.toFixed(2)+"/mo"}</span>
              </div>
              {d365r.saves > 0 && (
                <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, borderTop:"0.5px solid #eee", paddingTop:3, marginTop:3 }}>
                  <span style={{ fontWeight:600 }}>Saving:</span>
                  <span style={{ color:"#16a34a", fontWeight:700 }}>{"$"+(d365r.saves*row.seats).toFixed(2)+"/mo"}</span>
                </div>
              )}
              <div style={{ fontSize:10, color:"#888", marginTop:4, lineHeight:1.4 }}>{d365r.note}</div>
            </div>
          )}
          {topRisk==="NONE" && !d365r && <div style={{ fontSize:11, color:"#16a34a" }}>No audit risks detected.</div>}
        </div>

      </div>
    </div>
  );
}

function Summary({ rows }) {
  let m365T=0, d365T=0, d365SA=0;
  rows.forEach(r => {
    const b=gm(M365,r.m365), a=gm(ADDONS,r.addon), c=gm(COPILOT,r.copilot||"none"), ph=gm(TEAMS_PHONE,r.teamsPhone||"none");
    m365T += (b.price+a.price+c.price+ph.price)*r.seats;
    const d = r.d365Apps?.length ? optimizeD365(r.d365Apps) : null;
    if (d) { d365T += d.total*r.seats; d365SA += d.sa*r.seats; }
  });
  const d365Saves = d365SA - d365T;
  const labels = ["M365 + Phone", "D365 (optimized)", "D365 attach savings", "Grand total / mo"];
  const values = ["$"+m365T.toFixed(2), "$"+d365T.toFixed(2), d365Saves>0?"$"+d365Saves.toFixed(2):"—", "$"+(m365T+d365T).toFixed(2)];
  const colors = ["#111","#111",d365Saves>0?"#16a34a":"#aaa","#111"];
  return (
    <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:10, marginBottom:14 }}>
      {labels.map((l,i) => (
        <div key={i} style={{ background:"#f4f4f2", borderRadius:8, padding:"11px 13px" }}>
          <div style={{ fontSize:10, fontWeight:700, color:"#888", textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:4 }}>{l}</div>
          <div style={{ fontSize:20, fontWeight:700, color:colors[i] }}>{values[i]}</div>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  const [rows, setRows] = useState([
    { id:1, name:"Finance / Admin", m365:"o365e3", addon:"ems-e3", copilot:"none", teamsPhone:"none", d365Apps:["finance","scm"], hasExistingWindows:false, seats:5, isTrial:false },
    { id:2, name:"Sales Team",      m365:"biz-std", addon:"none",  copilot:"copilot-biz-promo", teamsPhone:"tp-domestic", d365Apps:["sales-ent","cs-ent"], hasExistingWindows:false, seats:12, isTrial:false },
    { id:3, name:"IT Admin",        m365:"m365e3",  addon:"none",  copilot:"none", teamsPhone:"none", d365Apps:[], hasExistingWindows:false, seats:3, isTrial:false },
  ]);
  const [analyzed, setAnalyzed] = useState(false);
  const [showScanner, setShowScanner] = useState(true);
  const [bgColor, setBgColor] = useState("#f4f4f2");

  const onDetected = detected => {
    setRows(p => [...p, ...detected.map((d,i) => ({
      id: Date.now()+i,
      name: d.name||"Detected",
      m365: d.m365||"none",
      addon: d.addon||"none",
      copilot: d.copilot||"none",
      teamsPhone: d.teamsPhone||"none",
      d365Apps: d.d365Apps||[],
      hasExistingWindows: d.hasExistingWindows||false,
      seats: d.seats||1,
      isTrial: d.isTrial||false,
    }))]);
    setAnalyzed(false);
    setShowScanner(false);
  };

  return (
    <div style={{ ...css.app, background:bgColor }}>
      <div style={css.hdr}>
        <div>
          <div style={{ fontSize:16, fontWeight:700, color:"#111" }}>Microsoft License Optimizer</div>
          <div style={{ fontSize:11, color:"#888" }}>M365 · D365 attach engine · Windows · Teams Phone · Copilot · AI scanner · 2026 pricing</div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <div style={{ display:"flex", alignItems:"center", gap:6, padding:"5px 10px", background:"#f4f4f2", borderRadius:8, border:"1px solid #ddd" }}>
            <span style={{ fontSize:11, color:"#666", fontWeight:600, whiteSpace:"nowrap" }}>Background</span>
            <input
              type="color"
              value={bgColor}
              onChange={e => setBgColor(e.target.value)}
              title="Pick background color"
              style={{ width:28, height:28, border:"none", borderRadius:6, cursor:"pointer", padding:0, background:"transparent" }}
            />
            <button
              onClick={() => setBgColor("#f4f4f2")}
              title="Reset to default"
              style={{ fontSize:11, color:"#aaa", background:"none", border:"none", cursor:"pointer", padding:"0 2px", lineHeight:1 }}
            >↺</button>
          </div>
          <button style={css.btn(showScanner?"primary":"outline")} onClick={() => setShowScanner(s => !s)}>
            {showScanner ? "▲ Hide scanner" : "📷 Scan screenshot"}
          </button>
        </div>
      </div>

      <div style={{ padding:"1.2rem", maxWidth:1100, margin:"0 auto" }}>
        {showScanner && <Scanner onDetected={onDetected} />}

        <div style={css.card}>
          <div style={{ display:"grid", gridTemplateColumns:"1.3fr 1.2fr 1.1fr 1fr 56px 30px", gap:7, marginBottom:6 }}>
            {["Name / role","M365 / O365","Security add-on","Teams Phone","Seats",""].map((h,i) => (
              <div key={i} style={{ fontSize:10, fontWeight:700, color:"#bbb", textTransform:"uppercase", letterSpacing:"0.04em" }}>{h}</div>
            ))}
          </div>
          <RowEditor rows={rows} setRows={setRows} />
          <div style={{ marginTop:12, display:"flex", gap:10, alignItems:"center" }}>
            <button style={css.btn("primary")} onClick={() => setAnalyzed(true)}>Analyze &amp; recommend →</button>
            {analyzed && <span style={{ fontSize:12, color:"#16a34a", fontWeight:600 }}>Report updated</span>}
          </div>
        </div>

        {analyzed && rows.length > 0 && (
          <div style={css.card}>
            <div style={css.stl}>Optimization Report</div>
            <Summary rows={rows} />
            {rows.map(r => <ResultCard key={r.id} row={r} />)}
            <div style={{ padding:"9px 13px", background:"#eff6ff", borderRadius:8, fontSize:12, color:"#1e40af", border:"0.5px solid #bfdbfe", marginTop:4 }}>
              July 1 2026: M365 Business Basic, Standard, E3 and E5 prices increasing — lock in renewals before the cutoff.
              Teams Phone Standard ($10) requires a separate calling plan ($13–$34) for external PSTN calls.
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
