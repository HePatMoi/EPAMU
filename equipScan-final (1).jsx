import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const WEAR_THRESHOLD = 0.85;

const SEED_ITEMS = [
  { id:"I001", name:"Mannequin RCP adulte", family:"MAN-ADU", unit:1, barcode:"MAN-ADU-001", category:"Mannequin", maxUses:500, uses:423, lastUsed:"2026-05-01", location:"Armoire A / Étagère 1", status:"ok", statusNote:"", statusDate:"" },
  { id:"I002", name:"Mannequin RCP adulte", family:"MAN-ADU", unit:2, barcode:"MAN-ADU-002", category:"Mannequin", maxUses:500, uses:198, lastUsed:"2026-04-28", location:"Armoire A / Étagère 1", status:"ok", statusNote:"", statusDate:"" },
  { id:"I003", name:"Mannequin RCP adulte", family:"MAN-ADU", unit:3, barcode:"MAN-ADU-003", category:"Mannequin", maxUses:500, uses:201, lastUsed:"2026-04-15", location:"Armoire A / Étagère 2", status:"hors_service", statusNote:"Valve défectueuse — en réparation", statusDate:"2026-05-10" },
  { id:"I004", name:"Mannequin RCP enfant", family:"MAN-ENF", unit:1, barcode:"MAN-ENF-001", category:"Mannequin", maxUses:500, uses:312, lastUsed:"2026-05-01", location:"Armoire A / Étagère 2", status:"ok", statusNote:"", statusDate:"" },
  { id:"I005", name:"Mannequin RCP enfant", family:"MAN-ENF", unit:2, barcode:"MAN-ENF-002", category:"Mannequin", maxUses:500, uses:87,  lastUsed:"2026-03-10", location:"Armoire A / Étagère 2", status:"ok", statusNote:"", statusDate:"" },
  { id:"I006", name:"Défibrillateur AED",  family:"AED",     unit:1, barcode:"AED-001",     category:"Médical",   maxUses:200, uses:174, lastUsed:"2026-05-01", location:"Local formation / Meuble rouge", status:"ok", statusNote:"", statusDate:"" },
  { id:"I007", name:"Défibrillateur AED",  family:"AED",     unit:2, barcode:"AED-002",     category:"Médical",   maxUses:200, uses:55,  lastUsed:"2026-04-01", location:"Local formation / Meuble rouge", status:"ok", statusNote:"", statusDate:"" },
];

const SEED_CONSUMABLES = [
  { id:"C001", name:"Gants nitrile (paires)", category:"EPI",     stock:18, minStock:30, reorderQty:100, unit:"paires",   deliveryDays:5,  perLearner:1, supplier:{ name:"MedStock Pro", ref:"GNT-L-100", contact:"commandes@medstock.fr" } },
  { id:"C002", name:"Masque bouche-à-bouche", category:"Hygiène", stock:45, minStock:20, reorderQty:50,  unit:"unités",   deliveryDays:7,  perLearner:1, supplier:{ name:"SecourPlus",   ref:"MBB-50",    contact:"supply@secourplus.fr"   } },
  { id:"C003", name:"Compresses stériles",    category:"Soins",   stock:8,  minStock:25, reorderQty:100, unit:"sachets",  deliveryDays:3,  perLearner:2, supplier:{ name:"MedStock Pro", ref:"COMP-10",   contact:"commandes@medstock.fr" } },
  { id:"C004", name:"Bandes de gaze",         category:"Soins",   stock:60, minStock:15, reorderQty:50,  unit:"rouleaux", deliveryDays:3,  perLearner:0, supplier:{ name:"MedStock Pro", ref:"GAZ-R5",    contact:"commandes@medstock.fr" } },
  { id:"C005", name:"Électrodes AED adulte",  category:"Médical", stock:6,  minStock:4,  reorderQty:10,  unit:"paires",   deliveryDays:14, perLearner:0, supplier:{ name:"CardioFrance", ref:"ELEC-AED-A", contact:"pro@cardiofrance.fr"  } },
];

const SEED_TRAININGS = [
  { id:"F001", name:"PSC1 — Gestes de premiers secours", date:"2026-05-14", status:"planned", learners:12,
    requiredFamilies:[{ family:"MAN-ADU", familyName:"Mannequin RCP adulte", quantity:3 },{ family:"AED", familyName:"Défibrillateur AED", quantity:1 }],
    requiredConsumables:[{ consumableId:"C001", perLearner:true },{ consumableId:"C002", perLearner:true }],
    assignedItems:[], consumablesConfirmed:false, scanLog:[], returnLog:[], completedAt:null },
  { id:"F002", name:"Formation secouriste entreprise", date:"2026-05-20", status:"planned", learners:8,
    requiredFamilies:[{ family:"MAN-ADU", familyName:"Mannequin RCP adulte", quantity:2 },{ family:"MAN-ENF", familyName:"Mannequin RCP enfant", quantity:2 }],
    requiredConsumables:[{ consumableId:"C001", perLearner:true },{ consumableId:"C003", perLearner:true }],
    assignedItems:[], consumablesConfirmed:false, scanLog:[], returnLog:[], completedAt:null },
];

// ── Helpers ───────────────────────────────────────────────────────────────
const genId    = p => `${p}${Date.now().toString(36).toUpperCase()}`;
const today    = () => new Date().toISOString().split("T")[0];
const fmtDate  = d => new Date(d+"T12:00:00").toLocaleDateString("fr-FR",{day:"2-digit",month:"short",year:"numeric"});
const usePct   = i => i.maxUses>0 ? Math.min(1,i.uses/i.maxUses) : 0;
const wColor   = p => p>=1?"#ef4444":p>=WEAR_THRESHOLD?"#f59e0b":"#22c55e";
const wLabel   = p => p>=1?"Fin de vie":p>=WEAR_THRESHOLD?"À surveiller":"OK";
const qrUrl    = code => `https://api.qrserver.com/v1/create-qr-code/?size=160x160&data=${encodeURIComponent(code)}&bgcolor=ffffff&color=000000&margin=6`;
const qrUrlRed = code => `https://api.qrserver.com/v1/create-qr-code/?size=160x160&data=${encodeURIComponent(code)}&bgcolor=fff0f0&color=cc0000&margin=6`;

// Items available for a family: ok status + not worn out
const availableItems = (items, family) =>
  items.filter(i => i.family===family && i.status==="ok" && usePct(i)<1);

// Rotation: oldest last-used first for even wear
function pickRotation(items, family, qty) {
  return availableItems(items, family)
    .sort((a,b) => {
      if (!a.lastUsed && !b.lastUsed) return a.uses-b.uses;
      if (!a.lastUsed) return -1;
      if (!b.lastUsed) return 1;
      return new Date(a.lastUsed)-new Date(b.lastUsed);
    })
    .slice(0,qty)
    .map(i=>({ itemId:i.id, barcode:i.barcode, name:i.name, unit:i.unit, location:i.location, uses:i.uses, maxUses:i.maxUses, scanned:false, returned:false }));
}

// ── SMART ALARM ENGINE ────────────────────────────────────────────────────
// For each training, compute feasibility issues accounting for:
// - hors_service units, worn-out units
// - units already committed to same-day trainings
// - consumable stock across all upcoming trainings
function computeAlarms(trainings, items, consumables) {
  const alarms = {}; // { trainingId: [{severity, type, message}] }

  const upcomingPlanned = trainings.filter(t => t.status!=="completed" && t.date>=today());

  upcomingPlanned.forEach(tr => {
    const issues = [];

    // Units committed on same day by OTHER trainings
    const sameDayOthers = upcomingPlanned.filter(t => t.id!==tr.id && t.date===tr.date);
    const committedByFamily = {};
    sameDayOthers.forEach(t =>
      t.requiredFamilies.forEach(rf => {
        committedByFamily[rf.family] = (committedByFamily[rf.family]||0) + rf.quantity;
      })
    );

    // Check tracked items
    tr.requiredFamilies.forEach(rf => {
      const avail  = availableItems(items, rf.family).length;
      const committed = committedByFamily[rf.family]||0;
      const net    = avail - committed;
      const hs     = items.filter(i=>i.family===rf.family && i.status==="hors_service").length;
      const worn   = items.filter(i=>i.family===rf.family && usePct(i)>=1).length;

      if (net < rf.quantity) {
        const shortfall = rf.quantity - net;
        const reasons = [];
        if (hs>0)        reasons.push(`${hs} hors service`);
        if (worn>0)      reasons.push(`${worn} fin de vie`);
        if (committed>0) reasons.push(`${committed} réservé(s) autre formation ce jour`);
        issues.push({
          severity: net<=0?"critical":"warning",
          type:"item",
          message:`${rf.familyName} : ${net}/${rf.quantity} disponible(s)${reasons.length?" ("+reasons.join(", ")+")":""}`,
        });
      }
    });

    // Check consumables — cumulative for all upcoming trainings up to this date
    tr.requiredConsumables.forEach(rc => {
      const c = consumables.find(x=>x.id===rc.consumableId);
      if (!c) return;
      const needed = rc.perLearner ? (c.perLearner||1)*tr.learners : (rc.fixedQty||1);
      // Total needed by all trainings on or before this date
      const totalNeeded = upcomingPlanned
        .filter(t => t.date<=tr.date)
        .reduce((sum,t) => {
          const match = t.requiredConsumables.find(x=>x.consumableId===rc.consumableId);
          if (!match) return sum;
          return sum + (match.perLearner?(c.perLearner||1)*t.learners:(match.fixedQty||1));
        },0);
      if (c.stock < totalNeeded) {
        issues.push({
          severity: c.stock < needed ? "critical":"warning",
          type:"consumable",
          message:`${c.name} : stock ${c.stock} ${c.unit}, besoin cumulé ${totalNeeded} (dont ${needed} pour cette formation)`,
        });
      }
    });

    if (issues.length) alarms[tr.id] = issues;
  });

  return alarms;
}

async function load(k,fb){ try{ const r=await window.storage.get(k); return r?JSON.parse(r.value):fb; }catch{ return fb; } }
async function save(k,v){ try{ await window.storage.set(k,JSON.stringify(v)); }catch{} }

// ── CSS ───────────────────────────────────────────────────────────────────
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
  *{box-sizing:border-box;margin:0;padding:0} input,select,textarea{font-family:inherit}
  ::-webkit-scrollbar{width:5px} ::-webkit-scrollbar-track{background:#1a1d27} ::-webkit-scrollbar-thumb{background:#f59e0b44;border-radius:3px}
  .nav{display:flex;align-items:center;gap:10px;padding:10px 16px;border-radius:8px;cursor:pointer;transition:all .15s;font-size:14px;font-weight:500;color:#9ca3af;border:none;background:none;width:100%;text-align:left}
  .nav:hover{background:#1e2230;color:#e8e8e8} .nav.on{background:#f59e0b15;color:#f59e0b;border-left:3px solid #f59e0b;padding-left:13px}
  .card{background:#161923;border:1px solid #2a2d3a;border-radius:12px;padding:20px}
  .btn{padding:8px 16px;border-radius:8px;border:none;cursor:pointer;font-size:13px;font-weight:600;transition:all .15s;font-family:inherit}
  .bp{background:#f59e0b;color:#0f1117} .bp:hover{background:#fbbf24}
  .bg{background:#1e2230;color:#9ca3af;border:1px solid #2a2d3a} .bg:hover{background:#2a2d3a;color:#e8e8e8}
  .bd{background:#ef444422;color:#ef4444;border:1px solid #ef444444} .bd:hover{background:#ef444433}
  .bs{background:#22c55e22;color:#22c55e;border:1px solid #22c55e44} .bs:hover{background:#22c55e33}
  .bb{background:#60a5fa22;color:#60a5fa;border:1px solid #60a5fa44} .bb:hover{background:#60a5fa33}
  .inp{background:#1e2230;border:1px solid #2a2d3a;border-radius:8px;padding:9px 12px;color:#e8e8e8;font-size:14px;outline:none;width:100%;transition:border .15s}
  .inp:focus{border-color:#f59e0b66}
  .sec{font-size:11px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:#6b7280;margin-bottom:12px}
  .chip{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:6px;font-size:12px;background:#2a2d3a;color:#9ca3af}
  .tag{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600}
  .pbar{height:5px;border-radius:3px;background:#2a2d3a;overflow:hidden}
  .scan-inp{background:#1e2230;border:2px solid #f59e0b;border-radius:12px;padding:14px 18px;color:#f59e0b;font-family:'DM Mono',monospace;font-size:18px;width:100%;outline:none;letter-spacing:1px}
  .scan-inp::placeholder{color:#f59e0b44}
  .logline{padding:7px 12px;border-radius:6px;font-family:'DM Mono',monospace;font-size:11px;border-left:3px solid}
  .overlay{position:fixed;inset:0;background:#000000aa;z-index:100;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(4px)}
  .modal{background:#161923;border:1px solid #2a2d3a;border-radius:16px;padding:28px;width:100%;max-width:580px;max-height:90vh;overflow-y:auto}
  .modal-wide{max-width:860px}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}} .pulse{animation:pulse 2s infinite}
  .tr-card{border-radius:10px;padding:14px;border:1px solid #2a2d3a;background:#161923;transition:border .15s} .tr-card:hover{border-color:#f59e0b44}
  .row:hover{background:#1e2230!important}
  .tab{padding:8px 16px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;border:none;background:none;color:#6b7280;transition:all .15s}
  .tab.on{background:#f59e0b22;color:#f59e0b} .tab:hover{color:#e8e8e8}
  .hs-label{background:white;border:3px solid #dc2626;border-radius:8px;padding:12px;display:flex;flex-direction:column;align-items:center;gap:6px;font-family:'DM Mono',monospace;color:#111;width:180px;flex-shrink:0}
  @media print{.no-print{display:none!important} body{background:white} .print-area{padding:20px}}
`;

// ═════════════════════════════════════════════════════════════════════════
export default function App() {
  const [page,setPage]           = useState("dashboard");
  const [items,setItems]         = useState(SEED_ITEMS);
  const [cons,setCons]           = useState(SEED_CONSUMABLES);
  const [trainings,setTrainings] = useState(SEED_TRAININGS);
  const [scanTrId,setScanTrId]   = useState(null);
  const [scanMode,setScanMode]   = useState("out");
  const [ready,setReady]         = useState(false);

  useEffect(()=>{(async()=>{
    const it=await load("eqs_items_v3",SEED_ITEMS);
    const co=await load("eqs_cons_v3",SEED_CONSUMABLES);
    const tr=await load("eqs_train_v3",SEED_TRAININGS);
    setItems(it); setCons(co); setTrainings(tr); setReady(true);
  })();},[]);
  useEffect(()=>{ if(ready) save("eqs_items_v3",items); },[items,ready]);
  useEffect(()=>{ if(ready) save("eqs_cons_v3",cons); },[cons,ready]);
  useEffect(()=>{ if(ready) save("eqs_train_v3",trainings); },[trainings,ready]);

  const alarms      = computeAlarms(trainings,items,cons);
  const wearAlerts  = items.filter(i=>usePct(i)>=WEAR_THRESHOLD&&i.status==="ok");
  const stockAlerts = cons.filter(c=>c.stock<=c.minStock);
  const totalAlerts = wearAlerts.length + stockAlerts.length + Object.keys(alarms).length;

  const openScan = (id,mode="out") => { setScanTrId(id); setScanMode(mode); setPage("scan"); };

  if(!ready) return <div style={{background:"#0f1117",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{color:"#f59e0b",fontFamily:"monospace"}}>Chargement…</span></div>;

  const navItems=[
    {id:"dashboard", icon:"⬡", label:"Tableau de bord", badge:totalAlerts||null},
    {id:"materiel",  icon:"◈", label:"Matériel"},
    {id:"formations",icon:"◷", label:"Formations"},
    {id:"etiquettes",icon:"⊞", label:"Étiquettes QR"},
  ];
  const isOn = id => page===id||(page==="scan"&&id==="formations");

  return (
    <div style={{display:"flex",minHeight:"100vh",background:"#0f1117",fontFamily:"'DM Sans','Segoe UI',sans-serif",color:"#e8e8e8"}}>
      <style>{CSS}</style>
      <aside style={{width:224,minWidth:224,background:"#0c0e17",borderRight:"1px solid #1e2230",padding:"24px 16px",display:"flex",flexDirection:"column",gap:4,position:"sticky",top:0,height:"100vh"}}>
        <div style={{padding:"0 8px 24px",borderBottom:"1px solid #1e2230",marginBottom:8}}>
          <div style={{fontSize:13,fontWeight:700,color:"#f59e0b",letterSpacing:1,fontFamily:"'DM Mono',monospace"}}>ÉQUIP·SCAN</div>
          <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>Gestion & suivi matériel</div>
        </div>
        {navItems.map(n=>(
          <button key={n.id} className={`nav${isOn(n.id)?" on":""}`} onClick={()=>setPage(n.id)}>
            <span style={{fontSize:15}}>{n.icon}</span>
            <span style={{flex:1}}>{n.label}</span>
            {n.badge?<span style={{background:"#ef444422",color:"#ef4444",borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:700}}>{n.badge}</span>:null}
          </button>
        ))}
        <div style={{marginTop:"auto",padding:"16px 8px 0",borderTop:"1px solid #1e2230"}}>
          <div style={{fontSize:11,color:"#4b5563"}}>{items.filter(i=>i.status==="ok").length} unités actives · {items.filter(i=>i.status==="hors_service").length} HS</div>
        </div>
      </aside>

      <main style={{flex:1,padding:"32px",overflowY:"auto",maxHeight:"100vh"}}>
        {page==="dashboard"  && <Dashboard items={items} cons={cons} trainings={trainings} wearAlerts={wearAlerts} stockAlerts={stockAlerts} alarms={alarms} onOpenScan={openScan}/>}
        {page==="materiel"   && <MatérielPage items={items} setItems={setItems} cons={cons} setCons={setCons}/>}
        {page==="formations" && <FormationsPage items={items} cons={cons} trainings={trainings} setTrainings={setTrainings} alarms={alarms} onOpenScan={openScan}/>}
        {page==="etiquettes" && <ÉtiquettesPage items={items} setItems={setItems}/>}
        {page==="scan"       && <ScanPage trId={scanTrId} mode={scanMode} setMode={setScanMode} items={items} setItems={setItems} cons={cons} setCons={setCons} trainings={trainings} setTrainings={setTrainings} onBack={()=>setPage("formations")}/>}
      </main>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═════════════════════════════════════════════════════════════════════════
function Dashboard({items,cons,trainings,wearAlerts,stockAlerts,alarms,onOpenScan}){
  const td=today();
  const todayTr=trainings.filter(t=>t.date===td&&t.status!=="completed");
  const upcoming=trainings.filter(t=>t.date>td&&t.status!=="completed").slice(0,5);
  const alarmedTr=Object.keys(alarms);

  return(
    <div style={{display:"flex",flexDirection:"column",gap:24}}>
      <div>
        <div style={{fontSize:24,fontWeight:700,marginBottom:4}}>Tableau de bord</div>
        <div style={{color:"#6b7280",fontSize:14}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14}}>
        {[
          {label:"Unités actives",     val:items.filter(i=>i.status==="ok").length,       sub:`${items.filter(i=>i.status==="hors_service").length} hors service`, c:"#60a5fa"},
          {label:"Alertes usure",      val:wearAlerts.length,    sub:"matériel à surveiller", c:wearAlerts.length?"#f59e0b":"#22c55e"},
          {label:"Stocks insuffisants",val:stockAlerts.length,   sub:"sous le seuil de commande", c:stockAlerts.length?"#ef4444":"#22c55e"},
          {label:"Formations à risque",val:alarmedTr.length,     sub:"matériel ou stock insuffisant", c:alarmedTr.length?"#f97316":"#22c55e"},
        ].map(s=>(
          <div key={s.label} className="card" style={{borderLeft:`3px solid ${s.c}`}}>
            <div style={{fontSize:28,fontWeight:700,color:s.c,fontFamily:"'DM Mono',monospace"}}>{s.val}</div>
            <div style={{fontSize:13,fontWeight:600,marginTop:4}}>{s.label}</div>
            <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>{s.sub}</div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
        {/* Formation alarms */}
        <div className="card">
          <div className="sec">🔔 Alertes formations</div>
          {alarmedTr.length===0
            ? <div style={{color:"#22c55e",fontSize:13}}>● Toutes les formations sont faisables</div>
            : alarmedTr.map(tid=>{
              const tr=trainings.find(t=>t.id===tid); if(!tr) return null;
              return(
                <div key={tid} style={{marginBottom:10,padding:"10px 12px",background:"#f9731611",border:"1px solid #f9731633",borderRadius:8}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#fb923c",marginBottom:4}}>{tr.name}</div>
                  <div style={{fontSize:11,color:"#9ca3af",marginBottom:6}}>{fmtDate(tr.date)} · {tr.learners} apprenant{tr.learners>1?"s":""}</div>
                  {alarms[tid].map((iss,i)=>(
                    <div key={i} style={{fontSize:12,color:iss.severity==="critical"?"#ef4444":"#fbbf24",marginBottom:2}}>
                      {iss.severity==="critical"?"⛔":"⚠"} {iss.message}
                    </div>
                  ))}
                </div>
              );
            })}
        </div>

        {/* Stock & wear alerts */}
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div className="card">
            <div className="sec">📦 Consommables à commander</div>
            {stockAlerts.length===0
              ? <div style={{color:"#22c55e",fontSize:13}}>● Tous les stocks OK</div>
              : stockAlerts.map(c=>(
                <div key={c.id} style={{marginBottom:8,padding:"8px 12px",background:c.stock===0?"#ef444411":"#f59e0b11",borderRadius:8,border:`1px solid ${c.stock===0?"#ef444433":"#f59e0b33"}`}}>
                  <div style={{fontSize:12,fontWeight:600}}>{c.name}</div>
                  <div style={{fontSize:11,color:"#9ca3af"}}>{c.stock} {c.unit} en stock · seuil {c.minStock} · délai {c.deliveryDays}j</div>
                  <div style={{fontSize:11,color:"#60a5fa",marginTop:2}}>{c.supplier.name} — {c.supplier.ref}</div>
                </div>
              ))}
          </div>
          <div className="card" style={{flex:1}}>
            <div className="sec">⚠ Usure matériel</div>
            {wearAlerts.length===0
              ? <div style={{color:"#22c55e",fontSize:13}}>● Tout le matériel est en bon état</div>
              : wearAlerts.map(i=>{const p=usePct(i); return(
                <div key={i.id} style={{marginBottom:6,display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:p>=1?"#ef444411":"#f59e0b11",borderRadius:8,border:`1px solid ${p>=1?"#ef444433":"#f59e0b33"}`}}>
                  <span className={p>=1?"pulse":""} style={{color:wColor(p),fontSize:16}}>●</span>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:600}}>{i.name} #{i.unit}</div>
                    <div style={{fontSize:11,color:"#9ca3af"}}>{i.uses}/{i.maxUses} · {i.location}</div>
                  </div>
                  <span style={{fontSize:10,fontWeight:700,color:wColor(p)}}>{wLabel(p)}</span>
                </div>
              );})}
          </div>
        </div>
      </div>

      {/* Today's trainings */}
      {(todayTr.length>0||upcoming.length>0)&&(
        <div className="card">
          <div className="sec">◷ Formations à venir</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:10}}>
            {[...todayTr.map(t=>({...t,_today:true})),...upcoming].map(t=>(
              <div key={t.id} style={{padding:"10px 14px",background:"#1e2230",borderRadius:8,border:`1px solid ${alarms[t.id]?"#f9731644":"#2a2d3a"}`}}>
                <div style={{fontSize:13,fontWeight:600,marginBottom:2}}>{t.name}</div>
                <div style={{fontSize:11,color:t._today?"#22c55e":"#6b7280",marginBottom:6}}>{t._today?"Aujourd'hui":fmtDate(t.date)} · {t.learners} apprenant{t.learners>1?"s":""}</div>
                {alarms[t.id]&&<div style={{fontSize:11,color:"#f97316",marginBottom:6}}>⚠ {alarms[t.id].length} problème{alarms[t.id].length>1?"s":""} détecté{alarms[t.id].length>1?"s":""}</div>}
                {t._today&&<button className="btn bp" style={{fontSize:11,padding:"4px 10px"}} onClick={()=>onOpenScan(t.id,"out")}>▶ Scanner</button>}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// MATÉRIEL PAGE
// ═════════════════════════════════════════════════════════════════════════
function MatérielPage({items,setItems,cons,setCons}){
  const [tab,setTab]         = useState("suivi");
  const [showItemModal,setShowItemModal] = useState(false);
  const [showConsModal,setShowConsModal] = useState(false);
  const [editItem,setEditItem] = useState(null);
  const [editCons,setEditCons] = useState(null);
  const [search,setSearch]   = useState("");

  const saveItem = data => {
    if(data.id) setItems(prev=>prev.map(i=>i.id===data.id?data:i));
    else setItems(prev=>[...prev,{...data,id:genId("I"),uses:0,lastUsed:"",statusDate:""}]);
    setShowItemModal(false);
  };
  const saveCons = data => {
    if(data.id) setCons(prev=>prev.map(c=>c.id===data.id?data:c));
    else setCons(prev=>[...prev,{...data,id:genId("C")}]);
    setShowConsModal(false);
  };
  const toggleStatus = (item) => {
    const newStatus = item.status==="ok"?"hors_service":"ok";
    setItems(prev=>prev.map(i=>i.id===item.id?{...i,status:newStatus,statusDate:today(),statusNote:newStatus==="ok"?"":i.statusNote}:i));
  };

  const filteredItems = items.filter(i=>
    i.name.toLowerCase().includes(search.toLowerCase())||
    i.barcode.toLowerCase().includes(search.toLowerCase())||
    i.family.toLowerCase().includes(search.toLowerCase())
  );
  const filteredCons = cons.filter(c=>c.name.toLowerCase().includes(search.toLowerCase()));

  // Group items by family for display
  const families = [...new Set(filteredItems.map(i=>i.family))];

  return(
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <div style={{flex:1}}>
          <div style={{fontSize:22,fontWeight:700}}>Matériel</div>
          <div style={{color:"#6b7280",fontSize:13,marginTop:2}}>{items.filter(i=>i.status==="ok").length} unités actives · {items.filter(i=>i.status==="hors_service").length} hors service · {cons.length} consommables</div>
        </div>
        <input className="inp" style={{width:220}} placeholder="Rechercher…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <button className="tab on" style={{marginRight:4}} onClick={()=>{setEditItem(null);setShowItemModal(true);}}>+ Unité</button>
        <button className="btn bp" onClick={()=>{setEditCons(null);setShowConsModal(true);}}>+ Consommable</button>
      </div>

      <div style={{display:"flex",gap:8,marginBottom:4}}>
        {["suivi","consommables"].map(t=>(
          <button key={t} className={`tab${tab===t?" on":""}`} onClick={()=>setTab(t)}>
            {t==="suivi"?"Matériel suivi ("+items.length+" unités)":"Consommables ("+cons.length+")"}
          </button>
        ))}
      </div>

      {tab==="suivi" && (
        <div style={{display:"flex",flexDirection:"column",gap:20}}>
          {families.map(fam=>{
            const famItems=filteredItems.filter(i=>i.family===fam);
            const available=famItems.filter(i=>i.status==="ok"&&usePct(i)<1).length;
            const hs=famItems.filter(i=>i.status==="hors_service").length;
            return(
              <div key={fam} className="card" style={{padding:0,overflow:"hidden"}}>
                <div style={{padding:"12px 16px",background:"#1a1d27",borderBottom:"1px solid #2a2d3a",display:"flex",alignItems:"center",gap:12}}>
                  <div style={{flex:1}}>
                    <span style={{fontWeight:700,fontSize:14}}>{famItems[0].name}</span>
                    <span style={{fontSize:11,color:"#6b7280",marginLeft:10,fontFamily:"'DM Mono',monospace"}}>{fam}</span>
                  </div>
                  <span style={{fontSize:12,color:"#22c55e"}}>{available} disponibles</span>
                  {hs>0&&<span style={{fontSize:12,color:"#ef4444"}}>{hs} hors service</span>}
                </div>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead>
                    <tr style={{borderBottom:"1px solid #1e2230"}}>
                      {["#","Code QR","Emplacement","Utilisations","État","Actions"].map(h=>(
                        <th key={h} style={{padding:"9px 14px",textAlign:"left",fontSize:10,letterSpacing:1.5,fontWeight:600,color:"#6b7280",textTransform:"uppercase"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {famItems.map((item,i)=>{
                      const p=usePct(item);
                      const hs=item.status==="hors_service";
                      return(
                        <tr key={item.id} className="row" style={{borderBottom:i<famItems.length-1?"1px solid #1e2230":"none",opacity:hs?0.7:1}}>
                          <td style={{padding:"10px 14px",fontWeight:700,color:"#f59e0b",fontFamily:"'DM Mono',monospace"}}>#{item.unit}</td>
                          <td style={{padding:"10px 14px",fontFamily:"'DM Mono',monospace",fontSize:12,color:"#9ca3af"}}>{item.barcode}</td>
                          <td style={{padding:"10px 14px",fontSize:12,color:"#9ca3af"}}>{item.location||"—"}</td>
                          <td style={{padding:"10px 14px",minWidth:150}}>
                            {hs ? <span style={{fontSize:12,color:"#ef4444"}}>Hors service</span> : (
                              <div style={{display:"flex",alignItems:"center",gap:8}}>
                                <div style={{flex:1}}><div className="pbar"><div style={{height:"100%",width:`${p*100}%`,background:wColor(p),borderRadius:3}}/></div></div>
                                <span style={{fontSize:11,color:"#9ca3af",fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap"}}>{item.uses}/{item.maxUses}</span>
                              </div>
                            )}
                            {hs&&item.statusNote&&<div style={{fontSize:10,color:"#6b7280",marginTop:2}}>{item.statusNote}</div>}
                          </td>
                          <td style={{padding:"10px 14px"}}>
                            <span style={{padding:"3px 8px",borderRadius:4,fontSize:11,fontWeight:700,background:hs?"#ef444422":wColor(p)+"22",color:hs?"#ef4444":wColor(p)}}>
                              {hs?"Hors service":wLabel(p)}
                            </span>
                          </td>
                          <td style={{padding:"10px 14px"}}>
                            <div style={{display:"flex",gap:6}}>
                              <button className="btn bg" style={{fontSize:11,padding:"4px 10px"}} onClick={()=>{setEditItem(item);setShowItemModal(true);}}>Modifier</button>
                              <button className={`btn ${hs?"bs":"bd"}`} style={{fontSize:11,padding:"4px 10px"}} onClick={()=>toggleStatus(item)}>
                                {hs?"✓ Remettre en service":"⊗ Hors service"}
                              </button>
                              <button className="btn" style={{fontSize:11,padding:"4px 10px",background:"#1e2230",color:"#6b7280",border:"1px solid #2a2d3a"}} onClick={()=>setItems(prev=>prev.map(i=>i.id===item.id?{...i,uses:0}:i))} title="Remettre à 0">↺</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          })}
          {filteredItems.length===0&&<div style={{textAlign:"center",color:"#4b5563",padding:40}}>Aucun résultat</div>}
        </div>
      )}

      {tab==="consommables" && (
        <div className="card" style={{padding:0,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead>
              <tr style={{borderBottom:"1px solid #2a2d3a"}}>
                {["Nom","Catégorie","Stock","Seuil commande","Délai livraison","Fournisseur","Actions"].map(h=>(
                  <th key={h} style={{padding:"10px 14px",textAlign:"left",fontSize:10,letterSpacing:1.5,fontWeight:600,color:"#6b7280",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredCons.map((c,i)=>{
                const low=c.stock<=c.minStock;
                return(
                  <tr key={c.id} className="row" style={{borderBottom:i<filteredCons.length-1?"1px solid #1e2230":"none"}}>
                    <td style={{padding:"10px 14px",fontSize:14,fontWeight:500}}>{c.name}</td>
                    <td style={{padding:"10px 14px"}}><span className="chip">{c.category}</span></td>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{fontFamily:"'DM Mono',monospace",fontSize:14,fontWeight:700,color:low?"#ef4444":"#22c55e"}}>{c.stock}</span>
                      <span style={{fontSize:11,color:"#6b7280",marginLeft:4}}>{c.unit}</span>
                      {low&&<span style={{fontSize:10,color:"#ef4444",marginLeft:6}}>⚠ À commander</span>}
                    </td>
                    <td style={{padding:"10px 14px",fontSize:12,color:"#9ca3af"}}>{c.minStock} {c.unit}</td>
                    <td style={{padding:"10px 14px",fontSize:12,color:"#9ca3af"}}>{c.deliveryDays} jours</td>
                    <td style={{padding:"10px 14px",fontSize:12}}>
                      <div style={{fontWeight:600}}>{c.supplier.name}</div>
                      <div style={{color:"#6b7280",fontSize:11}}>{c.supplier.ref} · <a href={`mailto:${c.supplier.contact}`} style={{color:"#60a5fa",textDecoration:"none"}}>{c.supplier.contact}</a></div>
                    </td>
                    <td style={{padding:"10px 14px"}}>
                      <button className="btn bg" style={{fontSize:11,padding:"4px 10px"}} onClick={()=>{setEditCons(c);setShowConsModal(true);}}>Modifier</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {showItemModal&&<ItemModal item={editItem} onSave={saveItem} onClose={()=>setShowItemModal(false)}/>}
      {showConsModal&&<ConsModal cons={editCons} onSave={saveCons} onClose={()=>setShowConsModal(false)}/>}
    </div>
  );
}

// ─── Item Modal ───────────────────────────────────────────────────────────
function ItemModal({item,onSave,onClose}){
  const [f,setF]=useState(item||{name:"",family:"",unit:1,barcode:"",category:"",maxUses:100,location:"",status:"ok",statusNote:""});
  const s=(k,v)=>setF(p=>({...p,[k]:v}));
  const submit=()=>{if(!f.name||!f.barcode) return; onSave({...f,unit:+f.unit,maxUses:+f.maxUses,uses:f.uses||0});};
  return(
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:17,fontWeight:700,marginBottom:20}}>{item?"Modifier l'unité":"Nouvelle unité"}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
          {[["Nom du matériel","name",""],["Famille / Type","family","Ex: MAN-ADU"],["Code-barres / QR","barcode","Ex: MAN-ADU-003"],["Catégorie","category","Ex: Mannequin"],["Emplacement","location","Ex: Armoire A / Étagère 2","",""],].map(([label,key,ph])=>(
            <div key={key} style={{gridColumn:key==="location"?"1 / -1":"auto"}}>
              <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>{label}</div>
              <input className="inp" placeholder={ph} value={f[key]||""} onChange={e=>s(key,e.target.value)}/>
            </div>
          ))}
          <div><div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Numéro d'unité</div><input className="inp" type="number" min={1} value={f.unit||1} onChange={e=>s("unit",e.target.value)}/></div>
          <div><div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Durée de vie (utilisations)</div><input className="inp" type="number" min={1} value={f.maxUses||100} onChange={e=>s("maxUses",e.target.value)}/></div>
        </div>
        <div style={{marginBottom:16}}>
          <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>État</div>
          <select className="inp" value={f.status} onChange={e=>s("status",e.target.value)}>
            <option value="ok">✓ En service</option>
            <option value="hors_service">⊗ Hors service / Défectueux</option>
          </select>
        </div>
        {f.status==="hors_service"&&(
          <div style={{marginBottom:16}}>
            <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Raison (hors service)</div>
            <input className="inp" placeholder="Ex: Valve défectueuse — en réparation" value={f.statusNote||""} onChange={e=>s("statusNote",e.target.value)}/>
          </div>
        )}
        <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:20}}>
          <button className="btn bg" onClick={onClose}>Annuler</button>
          <button className="btn bp" onClick={submit}>Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

// ─── Consumable Modal ─────────────────────────────────────────────────────
function ConsModal({cons,onSave,onClose}){
  const def={name:"",category:"",stock:0,minStock:10,reorderQty:50,unit:"unités",deliveryDays:7,perLearner:0,supplier:{name:"",ref:"",contact:""}};
  const [f,setF]=useState(cons||def);
  const s=(k,v)=>setF(p=>({...p,[k]:v}));
  const ss=(k,v)=>setF(p=>({...p,supplier:{...p.supplier,[k]:v}}));
  const submit=()=>{if(!f.name) return; onSave({...f,stock:+f.stock,minStock:+f.minStock,reorderQty:+f.reorderQty,deliveryDays:+f.deliveryDays,perLearner:+f.perLearner});};
  return(
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:17,fontWeight:700,marginBottom:20}}>{cons?"Modifier le consommable":"Nouveau consommable"}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
          {[["Nom","name",""],["Catégorie","category",""],["Unité","unit","Ex: paires, sachets"],["Qté par apprenant","perLearner","0 = fixe, 1 = 1 par apprenant"]].map(([label,key,ph])=>(
            <div key={key}>
              <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>{label}</div>
              <input className="inp" placeholder={ph} value={f[key]||""} onChange={e=>s(key,e.target.value)}/>
            </div>
          ))}
          {[["Stock actuel","stock"],["Seuil de commande","minStock"],["Qté à commander","reorderQty"],["Délai livraison (jours)","deliveryDays"]].map(([label,key])=>(
            <div key={key}>
              <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>{label}</div>
              <input className="inp" type="number" min={0} value={f[key]||0} onChange={e=>s(key,e.target.value)}/>
            </div>
          ))}
        </div>
        <div style={{fontSize:12,color:"#9ca3af",marginBottom:8,marginTop:4,letterSpacing:1,textTransform:"uppercase"}}>Fournisseur</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:20}}>
          {[["Nom fournisseur","name",""],["Référence produit","ref",""],["Email contact","contact","commandes@…",""]].map(([label,key,ph])=>(
            <div key={key} style={{gridColumn:key==="contact"?"1 / -1":"auto"}}>
              <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>{label}</div>
              <input className="inp" placeholder={ph} value={f.supplier[key]||""} onChange={e=>ss(key,e.target.value)}/>
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
          <button className="btn bg" onClick={onClose}>Annuler</button>
          <button className="btn bp" onClick={submit}>Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// FORMATIONS PAGE
// ═════════════════════════════════════════════════════════════════════════
function FormationsPage({items,cons,trainings,setTrainings,alarms,onOpenScan}){
  const [view,setView]         = useState("list");
  const [showModal,setShow]    = useState(false);
  const [editTr,setEditTr]     = useState(null);
  const [showImport,setImport] = useState(false);
  const [calYear,setCalYear]   = useState(new Date().getFullYear());
  const [calMonth,setCalMonth] = useState(new Date().getMonth());

  const saveTr = data => {
    if(data.id) setTrainings(prev=>prev.map(t=>t.id===data.id?{...t,...data}:t));
    else setTrainings(prev=>[...prev,{...data,id:genId("F"),status:"planned",assignedItems:[],consumablesConfirmed:false,scanLog:[],returnLog:[],completedAt:null}]);
    setShow(false);
  };
  const deleteTr = id => setTrainings(prev=>prev.filter(t=>t.id!==id));

  const sorted=[...trainings].sort((a,b)=>a.date<b.date?-1:1);

  // Calendar
  const daysInMonth = new Date(calYear,calMonth+1,0).getDate();
  const firstDay    = new Date(calYear,calMonth,1).getDay()||7; // Mon=1
  const td=today();
  const trByDate={};
  trainings.forEach(t=>{ if(!trByDate[t.date]) trByDate[t.date]=[]; trByDate[t.date].push(t); });

  // Generate report text for email
  const sendReport = tr => {
    const assigned = tr.assignedItems||[];
    const returned = assigned.filter(i=>i.returned).length;
    const missing  = assigned.filter(i=>i.scanned&&!i.returned);
    let body = `RAPPORT DE FORMATION\n${"=".repeat(40)}\n\n`;
    body += `Formation : ${tr.name}\nDate : ${fmtDate(tr.date)}\nApprenants : ${tr.learners}\nStatut : ${tr.status==="completed"?"Terminée":"En cours"}\n\n`;
    body += `MATÉRIEL SUIVI\n${"-".repeat(30)}\n`;
    assigned.forEach(i=>{ body+=`• ${i.name} #${i.unit} (${i.barcode}) — ${i.returned?"✓ Rentré":"⚠ Non rentré"}\n`; });
    if(missing.length) body+=`\n⚠ MANQUANTS : ${missing.map(i=>`${i.name} #${i.unit}`).join(", ")}\n`;
    body+=`\nMatériel scanné : ${assigned.filter(i=>i.scanned).length}/${assigned.length}\nMatériel rentré : ${returned}/${assigned.length}\n`;
    const sub=encodeURIComponent(`Rapport formation — ${tr.name} — ${fmtDate(tr.date)}`);
    window.open(`mailto:?subject=${sub}&body=${encodeURIComponent(body)}`);
  };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <div style={{flex:1}}>
          <div style={{fontSize:22,fontWeight:700}}>Formations</div>
          <div style={{color:"#6b7280",fontSize:13,marginTop:2}}>{trainings.length} formations · {Object.keys(alarms).length} avec alertes</div>
        </div>
        <div style={{display:"flex",gap:6}}>
          {["list","cal"].map(v=>(
            <button key={v} className={`tab${view===v?" on":""}`} onClick={()=>setView(v)}>{v==="list"?"Liste":"Calendrier"}</button>
          ))}
        </div>
        <button className="btn bb" onClick={()=>setImport(true)}>⬆ Importer Excel</button>
        <button className="btn bp" onClick={()=>{setEditTr(null);setShow(true);}}>+ Nouvelle</button>
      </div>

      {view==="list" && (
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {sorted.map(tr=>{
            const iss=alarms[tr.id];
            const scanned=tr.assignedItems.filter(i=>i.scanned).length;
            const returned=tr.assignedItems.filter(i=>i.returned).length;
            return(
              <div key={tr.id} className="tr-card" style={{borderColor:iss?"#f9731644":tr.status==="completed"?"#22c55e33":"#2a2d3a"}}>
                <div style={{display:"flex",alignItems:"flex-start",gap:14}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:15,fontWeight:700}}>{tr.name}</span>
                      {tr.date===td&&<span className="tag" style={{background:"#22c55e22",color:"#22c55e"}}>● Aujourd'hui</span>}
                      {tr.status==="completed"&&<span className="tag" style={{background:"#22c55e22",color:"#22c55e"}}>✓ Terminée</span>}
                      {iss&&<span className="tag" style={{background:"#f9731622",color:"#f97316"}}>⚠ {iss.length} alerte{iss.length>1?"s":""}</span>}
                    </div>
                    <div style={{fontSize:12,color:"#6b7280",marginBottom:8}}>
                      <span style={{fontFamily:"'DM Mono',monospace",color:"#9ca3af"}}>{tr.id}</span> · {fmtDate(tr.date)} · {tr.learners} apprenant{tr.learners>1?"s":""}
                    </div>
                    {/* Alarm details */}
                    {iss&&(
                      <div style={{marginBottom:10,padding:"8px 12px",background:"#f9731611",borderRadius:8,border:"1px solid #f9731633"}}>
                        {iss.map((issue,i)=>(
                          <div key={i} style={{fontSize:12,color:issue.severity==="critical"?"#ef4444":"#fbbf24",marginBottom:i<iss.length-1?3:0}}>
                            {issue.severity==="critical"?"⛔":"⚠"} {issue.message}
                          </div>
                        ))}
                      </div>
                    )}
                    {/* Equipment chips */}
                    <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                      {tr.requiredFamilies.map(rf=>{
                        const avail=availableItems(items,rf.family).length;
                        const ok=avail>=rf.quantity;
                        return <span key={rf.family} className="chip" style={!ok?{background:"#ef444422",color:"#ef4444"}:{}}>{rf.familyName} ×{rf.quantity}{!ok&&` (${avail} dispo)`}</span>;
                      })}
                      {tr.requiredConsumables.map(rc=>{
                        const c=cons.find(x=>x.id===rc.consumableId); if(!c) return null;
                        const need=rc.perLearner?(c.perLearner||1)*tr.learners:(rc.fixedQty||1);
                        const ok=c.stock>=need;
                        return <span key={rc.consumableId} className="chip" style={!ok?{background:"#f59e0b22",color:"#f59e0b"}:{}}>{c.name} ×{need}</span>;
                      })}
                    </div>
                    {tr.assignedItems.length>0&&(
                      <div style={{fontSize:11,color:"#6b7280",marginTop:6}}>
                        Sorties : {scanned}/{tr.assignedItems.length} · Retours : {returned}/{tr.assignedItems.length}
                      </div>
                    )}
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"flex-end",flexShrink:0}}>
                    {tr.status!=="completed"&&<button className="btn bp" style={{fontSize:12}} onClick={()=>onOpenScan(tr.id,"out")}>▶ Scanner sortie</button>}
                    {tr.assignedItems.some(i=>i.scanned)&&tr.status!=="completed"&&<button className="btn bb" style={{fontSize:12}} onClick={()=>onOpenScan(tr.id,"in")}>↩ Scanner retour</button>}
                    {tr.status==="completed"&&<button className="btn bs" style={{fontSize:12}} onClick={()=>sendReport(tr)}>✉ Rapport email</button>}
                    <button className="btn bg" style={{fontSize:11}} onClick={()=>{setEditTr(tr);setShow(true);}}>Modifier</button>
                    <button className="btn bd" style={{fontSize:11}} onClick={()=>deleteTr(tr.id)}>Supprimer</button>
                  </div>
                </div>
              </div>
            );
          })}
          {sorted.length===0&&<div style={{textAlign:"center",color:"#4b5563",padding:40}}>Aucune formation. Créez-en une !</div>}
        </div>
      )}

      {view==="cal" && (
        <div className="card">
          <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:20}}>
            <button className="btn bg" style={{padding:"6px 12px"}} onClick={()=>{ let m=calMonth-1; let y=calYear; if(m<0){m=11;y--;} setCalMonth(m);setCalYear(y); }}>←</button>
            <div style={{flex:1,textAlign:"center",fontWeight:700,fontSize:16,textTransform:"capitalize"}}>
              {new Date(calYear,calMonth,1).toLocaleDateString("fr-FR",{month:"long",year:"numeric"})}
            </div>
            <button className="btn bg" style={{padding:"6px 12px"}} onClick={()=>{ let m=calMonth+1; let y=calYear; if(m>11){m=0;y++;} setCalMonth(m);setCalYear(y); }}>→</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4,marginBottom:8}}>
            {["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"].map(d=>(
              <div key={d} style={{textAlign:"center",fontSize:11,fontWeight:600,color:"#6b7280",padding:"4px 0"}}>{d}</div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
            {Array.from({length:firstDay-1},(_,i)=><div key={`e${i}`}/>)}
            {Array.from({length:daysInMonth},(_,i)=>{
              const dayNum=i+1;
              const dateStr=`${calYear}-${String(calMonth+1).padStart(2,"0")}-${String(dayNum).padStart(2,"0")}`;
              const dayTr=trByDate[dateStr]||[];
              const isToday=dateStr===td;
              const hasAlarm=dayTr.some(t=>alarms[t.id]);
              return(
                <div key={dayNum} className={`cal-day${dayTr.length?" has":""}${isToday?" today":""}`}>
                  <div style={{fontSize:12,fontWeight:isToday?700:400,color:isToday?"#60a5fa":"#9ca3af",marginBottom:3}}>{dayNum}</div>
                  {dayTr.map(t=>(
                    <div key={t.id} style={{fontSize:10,padding:"2px 4px",borderRadius:3,marginBottom:2,background:alarms[t.id]?"#f9731633":"#f59e0b22",color:alarms[t.id]?"#f97316":"#f59e0b",cursor:"pointer",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}} onClick={()=>onOpenScan(t.id,"out")}>
                      {hasAlarm&&alarms[t.id]?"⚠ ":""}{t.name}
                    </div>
                  ))}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {showModal&&<TrainingModal items={items} cons={cons} tr={editTr} onSave={saveTr} onClose={()=>setShow(false)}/>}
      {showImport&&<ExcelImportModal onImport={rows=>{
        const newTr=rows.map(r=>({
          id:genId("F"), name:r.name, date:r.date, learners:r.learners||10,
          status:"planned", requiredFamilies:[], requiredConsumables:[],
          assignedItems:[], consumablesConfirmed:false, scanLog:[], returnLog:[], completedAt:null
        }));
        setTrainings(prev=>[...prev,...newTr]);
        setImport(false);
      }} onClose={()=>setImport(false)}/>}
    </div>
  );
}

// ─── Training Modal ───────────────────────────────────────────────────────
function TrainingModal({items,cons,tr,onSave,onClose}){
  const [name,setName]=useState(tr?.name||"");
  const [date,setDate]=useState(tr?.date||today());
  const [learners,setLearners]=useState(tr?.learners||10);
  const [reqFam,setReqFam]=useState(tr?.requiredFamilies||[]);
  const [reqCon,setReqCon]=useState(tr?.requiredConsumables||[]);

  const families=[...new Map(items.map(i=>[i.family,{family:i.family,familyName:i.name}])).values()];

  const addFam=(fam,qty=1)=>{
    if(!fam) return;
    setReqFam(prev=>{
      const ex=prev.find(r=>r.family===fam);
      if(ex) return prev.map(r=>r.family===fam?{...r,quantity:r.quantity+qty}:r);
      const f=families.find(x=>x.family===fam);
      return [...prev,{family:fam,familyName:f?.familyName||fam,quantity:qty}];
    });
  };
  const removeFam=fam=>setReqFam(prev=>prev.filter(r=>r.family!==fam));

  const addCon=(cid)=>{
    if(!cid||reqCon.find(r=>r.consumableId===cid)) return;
    const c=cons.find(x=>x.id===cid);
    setReqCon(prev=>[...prev,{consumableId:cid,name:c?.name||cid,perLearner:c?.perLearner>0,fixedQty:c?.perLearner||1}]);
  };
  const removeCon=cid=>setReqCon(prev=>prev.filter(r=>r.consumableId!==cid));

  const [selFam,setSelFam]=useState(""); const [famQty,setFamQty]=useState(1);
  const [selCon,setSelCon]=useState("");

  const submit=()=>{
    if(!name||!date) return;
    onSave({...(tr||{}),name,date,learners:+learners,requiredFamilies:reqFam,requiredConsumables:reqCon});
  };

  return(
    <div className="overlay" onClick={onClose}>
      <div className="modal modal-wide" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:17,fontWeight:700,marginBottom:20}}>{tr?"Modifier la formation":"Nouvelle formation"}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:20}}>
          <div style={{gridColumn:"1 / -1"}}>
            <div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Nom de la formation</div>
            <input className="inp" placeholder="Ex: PSC1 — Gestes de premiers secours" value={name} onChange={e=>setName(e.target.value)}/>
          </div>
          <div><div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Date</div><input className="inp" type="date" value={date} onChange={e=>setDate(e.target.value)}/></div>
          <div><div style={{fontSize:12,color:"#9ca3af",marginBottom:5}}>Nombre d'apprenants</div><input className="inp" type="number" min={1} value={learners} onChange={e=>setLearners(e.target.value)}/></div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
          <div>
            <div className="sec">Matériel suivi requis</div>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <select className="inp" value={selFam} onChange={e=>setSelFam(e.target.value)} style={{flex:2}}>
                <option value="">— Sélectionner —</option>
                {families.map(f=><option key={f.family} value={f.family}>{f.familyName}</option>)}
              </select>
              <input className="inp" type="number" min={1} value={famQty} onChange={e=>setFamQty(+e.target.value)} style={{width:60}}/>
              <button className="btn bp" style={{whiteSpace:"nowrap"}} onClick={()=>{addFam(selFam,famQty);setSelFam("");}}>+</button>
            </div>
            {reqFam.map(r=>(
              <div key={r.family} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",background:"#1e2230",borderRadius:7,marginBottom:6}}>
                <span style={{flex:1,fontSize:13}}>{r.familyName}</span>
                <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#f59e0b"}}>×{r.quantity}</span>
                <button style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",fontSize:16}} onClick={()=>removeFam(r.family)}>×</button>
              </div>
            ))}
          </div>
          <div>
            <div className="sec">Consommables requis</div>
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              <select className="inp" value={selCon} onChange={e=>setSelCon(e.target.value)} style={{flex:1}}>
                <option value="">— Sélectionner —</option>
                {cons.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
              <button className="btn bp" onClick={()=>{addCon(selCon);setSelCon("");}}>+</button>
            </div>
            {reqCon.map(r=>{
              const c=cons.find(x=>x.id===r.consumableId);
              const qty=r.perLearner?(c?.perLearner||1)*learners:(r.fixedQty||1);
              return(
                <div key={r.consumableId} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",background:"#1e2230",borderRadius:7,marginBottom:6}}>
                  <span style={{flex:1,fontSize:13}}>{c?.name||r.consumableId}</span>
                  <span style={{fontSize:11,color:"#9ca3af"}}>{r.perLearner?`${c?.perLearner||1}×apprenant`:"fixe"}</span>
                  <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#f59e0b"}}>={qty}</span>
                  <button style={{background:"none",border:"none",color:"#6b7280",cursor:"pointer",fontSize:16}} onClick={()=>removeCon(r.consumableId)}>×</button>
                </div>
              );
            })}
          </div>
        </div>

        <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:24}}>
          <button className="btn bg" onClick={onClose}>Annuler</button>
          <button className="btn bp" onClick={submit}>Enregistrer</button>
        </div>
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// SCAN PAGE
// ═════════════════════════════════════════════════════════════════════════
function ScanPage({trId,mode,setMode,items,setItems,cons,setCons,trainings,setTrainings,onBack}){
  const training=trainings.find(t=>t.id===trId);
  const scanRef=useRef(null);
  const [inp,setInp]=useState("");
  const [lastResult,setLast]=useState(null);

  useEffect(()=>{ setTimeout(()=>scanRef.current?.focus(),100); },[mode]);

  if(!training) return <div style={{color:"#ef4444",padding:40}}>Formation introuvable.</div>;

  // Assign items on first open (sortie mode)
  const ensureAssigned = () => {
    if(training.assignedItems.length>0) return;
    const assigned=training.requiredFamilies.flatMap(rf=>pickRotation(items,rf.family,rf.quantity));
    setTrainings(prev=>prev.map(t=>t.id===trId?{...t,assignedItems:assigned}:t));
  };
  useEffect(()=>{ if(mode==="out") ensureAssigned(); },[]);

  const updateTr=(fn)=>setTrainings(prev=>prev.map(t=>t.id===trId?fn(t):t));

  const handleScan=()=>{
    const code=inp.trim().toUpperCase(); if(!code) return;
    const ts=new Date().toLocaleTimeString("fr-FR");
    const item=items.find(i=>i.barcode.toUpperCase()===code);
    let log;

    if(!item){
      log={ts,type:"error",msg:`Code inconnu : "${code}"`};
      setLast({ok:false,msg:`❌ Code inconnu : "${code}"`});
    } else if(item.status==="hors_service"){
      log={ts,type:"error",msg:`${item.name} #${item.unit} — HORS SERVICE (${item.statusNote||""})`};
      setLast({ok:false,msg:`⛔ ${item.name} #${item.unit} est hors service`});
    } else {
      const assigned=training.assignedItems.find(a=>a.itemId===item.id);
      if(!assigned){
        log={ts,type:"warning",msg:`${item.name} #${item.unit} — non prévu pour cette formation`};
        setLast({ok:false,msg:`⚠ ${item.name} #${item.unit} n'est pas dans la liste`});
      } else if(mode==="out"){
        if(assigned.scanned){
          log={ts,type:"warning",msg:`${item.name} #${item.unit} — déjà scanné en sortie`};
          setLast({ok:false,msg:`⚠ ${item.name} #${item.unit} déjà scanné`});
        } else {
          log={ts,type:"success",msg:`✓ Sortie confirmée : ${item.name} #${item.unit}`};
          setLast({ok:true,msg:`✓ ${item.name} #${item.unit} — Sortie enregistrée`});
          updateTr(t=>({...t,assignedItems:t.assignedItems.map(a=>a.itemId===item.id?{...a,scanned:true}:a),scanLog:[...(t.scanLog||[]),log]}));
          setInp(""); return;
        }
      } else { // return mode
        if(!assigned.scanned){
          log={ts,type:"warning",msg:`${item.name} #${item.unit} — pas encore scanné en sortie`};
          setLast({ok:false,msg:`⚠ ${item.name} #${item.unit} n'a pas été sorti`});
        } else if(assigned.returned){
          log={ts,type:"warning",msg:`${item.name} #${item.unit} — déjà rentré`};
          setLast({ok:false,msg:`⚠ ${item.name} #${item.unit} déjà enregistré au retour`});
        } else {
          log={ts,type:"success",msg:`✓ Retour confirmé : ${item.name} #${item.unit}`};
          setLast({ok:true,msg:`✓ ${item.name} #${item.unit} — Retour enregistré`});
          // Increment uses on return
          setItems(prev=>prev.map(i=>i.id===item.id?{...i,uses:i.uses+1,lastUsed:today()}:i));
          updateTr(t=>({...t,assignedItems:t.assignedItems.map(a=>a.itemId===item.id?{...a,returned:true}:a),returnLog:[...(t.returnLog||[]),log]}));
          setInp(""); return;
        }
      }
    }
    updateTr(t=>({...t,[mode==="out"?"scanLog":"returnLog"]:[...(t[mode==="out"?"scanLog":"returnLog"]||[]),log]}));
    setInp("");
  };

  const confirmConsumables=()=>{
    // Deduct from stock
    training.requiredConsumables.forEach(rc=>{
      const c=cons.find(x=>x.id===rc.consumableId); if(!c) return;
      const qty=rc.perLearner?(c.perLearner||1)*training.learners:(rc.fixedQty||1);
      setCons(prev=>prev.map(x=>x.id===rc.consumableId?{...x,stock:Math.max(0,x.stock-qty)}:x));
    });
    updateTr(t=>({...t,consumablesConfirmed:true}));
  };

  const completeTraining=()=>{
    updateTr(t=>({...t,status:"completed",completedAt:today()}));
    onBack();
  };

  const assigned=training.assignedItems;
  const allOut=assigned.length>0&&assigned.every(a=>a.scanned);
  const allBack=assigned.every(a=>a.returned);
  const missing=assigned.filter(a=>a.scanned&&!a.returned);
  const log=mode==="out"?training.scanLog:training.returnLog;

  return(
    <div style={{display:"flex",flexDirection:"column",gap:18,maxWidth:960}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <button className="btn bg" onClick={onBack}>← Retour</button>
        <div style={{flex:1}}>
          <div style={{fontSize:18,fontWeight:700}}>{training.name}</div>
          <div style={{color:"#6b7280",fontSize:13}}>{fmtDate(training.date)} · {training.learners} apprenants</div>
        </div>
        <div style={{display:"flex",gap:6}}>
          {[["out","▶ Sortie"],["in","↩ Retour"]].map(([m,label])=>(
            <button key={m} className={`tab${mode===m?" on":""}`} onClick={()=>setMode(m)}>{label}</button>
          ))}
        </div>
        {allBack&&training.consumablesConfirmed&&(
          <button className="btn bs" onClick={completeTraining}>✓ Clôturer</button>
        )}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}>
        {/* Scan zone */}
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div className="card">
            <div className="sec">{mode==="out"?"◈ SCAN SORTIE":"↩ SCAN RETOUR"}</div>
            <input ref={scanRef} className="scan-inp" placeholder="Scanner le code QR…" value={inp} onChange={e=>setInp(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleScan()} autoComplete="off" spellCheck={false}/>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button className="btn bp" style={{flex:1}} onClick={handleScan}>Valider (Entrée)</button>
              <button className="btn bg" onClick={()=>{setInp("");scanRef.current?.focus();}}>Effacer</button>
            </div>
            {lastResult&&(
              <div style={{marginTop:10,padding:"9px 12px",borderRadius:8,background:lastResult.ok?"#22c55e22":"#ef444422",border:`1px solid ${lastResult.ok?"#22c55e44":"#ef444444"}`,color:lastResult.ok?"#22c55e":"#ef4444",fontSize:13,fontWeight:600}}>
                {lastResult.msg}
              </div>
            )}
          </div>
          {/* Consumables */}
          {mode==="out"&&training.requiredConsumables.length>0&&(
            <div className="card">
              <div className="sec">📦 Consommables</div>
              {training.requiredConsumables.map(rc=>{
                const c=cons.find(x=>x.id===rc.consumableId); if(!c) return null;
                const qty=rc.perLearner?(c.perLearner||1)*training.learners:(rc.fixedQty||1);
                const enough=c.stock>=qty;
                return(
                  <div key={rc.consumableId} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:"#1e2230",borderRadius:8,marginBottom:6}}>
                    <div style={{flex:1}}>
                      <div style={{fontSize:13,fontWeight:500}}>{c.name}</div>
                      <div style={{fontSize:11,color:enough?"#9ca3af":"#ef4444"}}>{qty} {c.unit} requis · stock : {c.stock}</div>
                    </div>
                    {training.consumablesConfirmed&&<span style={{color:"#22c55e",fontSize:12,fontWeight:700}}>✓ Pris</span>}
                  </div>
                );
              })}
              {!training.consumablesConfirmed&&(
                <button className="btn bp" style={{width:"100%",marginTop:8}} onClick={confirmConsumables}>
                  ✓ Confirmer les consommables pris
                </button>
              )}
              {training.consumablesConfirmed&&<div style={{color:"#22c55e",fontSize:12,marginTop:6}}>✓ Consommables déduits du stock</div>}
            </div>
          )}
          {/* Log */}
          <div className="card">
            <div className="sec">Journal</div>
            <div style={{maxHeight:180,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
              {(log||[]).length===0?<div style={{color:"#4b5563",fontSize:12}}>Aucun scan</div>
                :[...(log||[])].reverse().map((e,i)=>(
                <div key={i} className="logline" style={{borderColor:e.type==="success"?"#22c55e":e.type==="warning"?"#f59e0b":"#ef4444",background:e.type==="success"?"#22c55e0a":e.type==="warning"?"#f59e0b0a":"#ef44440a",color:e.type==="success"?"#86efac":e.type==="warning"?"#fcd34d":"#fca5a5"}}>
                  <span style={{color:"#4b5563",marginRight:8}}>{e.ts}</span>{e.msg}
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Checklist */}
        <div className="card">
          <div className="sec">Liste de matériel — {mode==="out"?"sortie":"retour"}</div>
          {assigned.length===0&&<div style={{color:"#6b7280",fontSize:13}}>Attribution en cours…</div>}
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {assigned.map(a=>{
              const done=mode==="out"?a.scanned:a.returned;
              const wp=usePct({uses:a.uses,maxUses:a.maxUses});
              return(
                <div key={a.itemId} style={{padding:"11px 14px",borderRadius:10,background:done?"#22c55e11":"#1e2230",border:`1px solid ${done?"#22c55e33":"#2a2d3a"}`,transition:"all .3s"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{fontSize:20,color:done?"#22c55e":"#4b5563"}}>{done?"✓":"○"}</span>
                    <div style={{flex:1}}>
                      <div style={{fontSize:13,fontWeight:700}}>{a.name} <span style={{color:"#f59e0b",fontFamily:"'DM Mono',monospace"}}>#{a.unit}</span></div>
                      <div style={{fontSize:11,color:"#6b7280",fontFamily:"'DM Mono',monospace"}}>{a.barcode}</div>
                      {mode==="out"&&<div style={{fontSize:11,color:"#60a5fa",marginTop:2}}>📍 {a.location}</div>}
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontSize:11,color:wColor(wp)}}>{Math.round(wp*100)}% usé</div>
                      <div style={{fontSize:10,color:"#4b5563"}}>{a.uses}/{a.maxUses}</div>
                    </div>
                  </div>
                  {wp>=WEAR_THRESHOLD&&<div style={{marginTop:6,fontSize:11,color:"#f59e0b"}}>⚠ {wLabel(wp)}</div>}
                </div>
              );
            })}
          </div>
          {mode==="in"&&missing.length>0&&(
            <div style={{marginTop:12,padding:"10px 14px",background:"#ef444411",border:"1px solid #ef444433",borderRadius:8}}>
              <div style={{color:"#ef4444",fontWeight:700,fontSize:13,marginBottom:4}}>⚠ Matériel non rentré</div>
              {missing.map(a=><div key={a.itemId} style={{fontSize:12,color:"#fca5a5"}}>• {a.name} #{a.unit} ({a.barcode})</div>)}
            </div>
          )}
          {mode==="in"&&allBack&&(
            <div style={{marginTop:10,padding:"10px 14px",background:"#22c55e11",border:"1px solid #22c55e33",borderRadius:8,color:"#22c55e",fontSize:13,fontWeight:600}}>
              ✓ Tout le matériel est rentré
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// ÉTIQUETTES QR
// ═════════════════════════════════════════════════════════════════════════
function ÉtiquettesPage({items,setItems}){
  const [filter,setFilter]=useState("all"); // "all" | "ok" | "hors_service"
  const [showHsModal,setShowHsModal]=useState(false);
  const [selectedItem,setSelectedItem]=useState(null);

  const filtered=items.filter(i=>filter==="all"||i.status===filter);

  const openHsLabel=(item)=>{ setSelectedItem(item); setShowHsModal(true); };

  return(
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <div style={{flex:1}}>
          <div style={{fontSize:22,fontWeight:700}}>Étiquettes QR</div>
          <div style={{color:"#6b7280",fontSize:13,marginTop:2}}>Générez et imprimez les étiquettes de chaque unité</div>
        </div>
        <div style={{display:"flex",gap:6}}>
          {[["all","Toutes"],["ok","En service"],["hors_service","Hors service"]].map(([v,l])=>(
            <button key={v} className={`tab${filter===v?" on":""}`} onClick={()=>setFilter(v)}>{l}</button>
          ))}
        </div>
        <button className="btn bp" onClick={()=>window.print()}>🖨 Imprimer</button>
      </div>

      <div className="card no-print" style={{background:"#1e2230",padding:"12px 16px"}}>
        <div style={{fontSize:13,color:"#9ca3af"}}>
          💡 Les étiquettes <span style={{color:"#22c55e"}}>vertes</span> sont pour le matériel en service. Les <span style={{color:"#ef4444"}}>rouges</span> signalent le matériel hors service / défectueux — à coller sur l'objet pour qu'il ne parte pas en formation.
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:16}} className="print-area">
        {filtered.map(item=>{
          const hs=item.status==="hors_service";
          return(
            <div key={item.id} style={{border:`2px solid ${hs?"#dc2626":"#2a2d3a"}`,borderRadius:10,padding:14,background:hs?"#fff0f011":"#161923",display:"flex",flexDirection:"column",alignItems:"center",gap:8,position:"relative"}}>
              <img src={hs?`https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(item.barcode)}&color=cc0000&bgcolor=fff0f0&margin=4`:`https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(item.barcode)}&color=f59e0b&bgcolor=161923&margin=4`}
                alt={item.barcode} style={{width:120,height:120,borderRadius:6}}/>
              <div style={{textAlign:"center"}}>
                <div style={{fontWeight:700,fontSize:13}}>{item.name}</div>
                <div style={{fontFamily:"'DM Mono',monospace",fontSize:12,color:"#f59e0b",marginTop:2}}>#{item.unit} · {item.barcode}</div>
                <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>{item.location}</div>
                {hs&&(
                  <div style={{marginTop:6,padding:"4px 8px",background:"#ef444422",border:"1px solid #ef444455",borderRadius:6}}>
                    <div style={{color:"#ef4444",fontWeight:700,fontSize:11}}>⛔ HORS SERVICE</div>
                    {item.statusNote&&<div style={{color:"#fca5a5",fontSize:10,marginTop:2}}>{item.statusNote}</div>}
                    {item.statusDate&&<div style={{color:"#6b7280",fontSize:10}}>Depuis le {fmtDate(item.statusDate)}</div>}
                  </div>
                )}
                {!hs&&<div style={{fontSize:11,color:wColor(usePct(item)),marginTop:4}}>{item.uses}/{item.maxUses} utilisations · {wLabel(usePct(item))}</div>}
              </div>
              {!hs&&(
                <button className="btn bd no-print" style={{fontSize:10,padding:"3px 8px",width:"100%"}} onClick={()=>openHsLabel(item)}>
                  Générer étiquette HS
                </button>
              )}
            </div>
          );
        })}
      </div>

      {showHsModal&&selectedItem&&(
        <HsLabelModal item={selectedItem} onConfirm={(note)=>{
          setItems(prev=>prev.map(i=>i.id===selectedItem.id?{...i,status:"hors_service",statusNote:note,statusDate:today()}:i));
          setShowHsModal(false);
        }} onClose={()=>setShowHsModal(false)}/>
      )}
    </div>
  );
}

function HsLabelModal({item,onConfirm,onClose}){
  const [note,setNote]=useState("");
  return(
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:17,fontWeight:700,marginBottom:6}}>Mettre hors service</div>
        <div style={{color:"#9ca3af",fontSize:13,marginBottom:20}}>{item.name} #{item.unit} — {item.barcode}</div>

        {/* Preview label */}
        <div style={{display:"flex",justifyContent:"center",marginBottom:20}}>
          <div style={{border:"3px solid #dc2626",borderRadius:10,padding:16,background:"#fff8f8",display:"flex",flexDirection:"column",alignItems:"center",gap:8,width:200}}>
            <img src={`https://api.qrserver.com/v1/create-qr-code/?size=140x140&data=${encodeURIComponent(item.barcode)}&color=cc0000&bgcolor=fff8f8&margin=4`} style={{width:120,height:120,borderRadius:4}} alt="QR"/>
            <div style={{textAlign:"center",fontFamily:"monospace"}}>
              <div style={{fontWeight:800,fontSize:14,color:"#111"}}>{item.name}</div>
              <div style={{fontSize:12,color:"#333"}}>#{item.unit} · {item.barcode}</div>
              <div style={{marginTop:6,padding:"4px 8px",background:"#fee2e2",border:"1px solid #fca5a5",borderRadius:4}}>
                <div style={{color:"#dc2626",fontWeight:800,fontSize:13}}>⛔ HORS SERVICE</div>
                {note&&<div style={{color:"#7f1d1d",fontSize:10,marginTop:2}}>{note}</div>}
                <div style={{color:"#6b7280",fontSize:10}}>Depuis le {fmtDate(today())}</div>
              </div>
            </div>
          </div>
        </div>

        <div style={{marginBottom:20}}>
          <div style={{fontSize:12,color:"#9ca3af",marginBottom:6}}>Raison / Note (apparaîtra sur l'étiquette)</div>
          <input className="inp" placeholder="Ex: Valve défectueuse, envoyé en réparation" value={note} onChange={e=>setNote(e.target.value)}/>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
          <button className="btn bg" onClick={onClose}>Annuler</button>
          <button className="btn bd" onClick={()=>onConfirm(note)}>⛔ Confirmer hors service & imprimer</button>
        </div>
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// EXCEL IMPORT MODAL
// ═════════════════════════════════════════════════════════════════════════

// Parse various date formats into YYYY-MM-DD
function parseDate(val) {
  if (!val) return "";
  // Excel serial number
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
  }
  const s = String(val).trim();
  // Already ISO
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // DD/MM/YYYY or DD-MM-YYYY
  const m1 = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (m1) return `${m1[3]}-${m1[2].padStart(2,"0")}-${m1[1].padStart(2,"0")}`;
  // MM/DD/YYYY
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m2) {
    const y = m2[3].length===2 ? "20"+m2[3] : m2[3];
    return `${y}-${m2[1].padStart(2,"0")}-${m2[2].padStart(2,"0")}`;
  }
  // Try native parse
  const d = new Date(s);
  if (!isNaN(d)) return d.toISOString().split("T")[0];
  return "";
}

// Detect which column likely maps to name/date/learners
function autoDetect(headers) {
  const norm = h => h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  const find = (keys) => headers.find(h => keys.some(k => norm(h).includes(k))) || "";
  return {
    colName:     find(["nom","formation","titre","intitule","description","sujet"]),
    colDate:     find(["date","jour","quand","planning"]),
    colLearners: find(["appren","participant","stagiaire","eleve","etudiant","nombre","nb"]),
  };
}

function ExcelImportModal({onImport, onClose}) {
  const [step, setStep]       = useState("upload"); // upload | map | preview
  const [headers, setHeaders] = useState([]);
  const [rows, setRows]       = useState([]);
  const [cols, setCols]       = useState({colName:"", colDate:"", colLearners:""});
  const [preview, setPreview] = useState([]);
  const [selected, setSelected] = useState([]);
  const [error, setError]     = useState("");
  const [dragging, setDragging] = useState(false);
  const fileRef = useRef();

  const downloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Nom de la formation","Date","Nombre d'apprenants"],
      ["PSC1 — Gestes de premiers secours","14/06/2026",12],
      ["Formation secouriste entreprise","21/06/2026",8],
      ["Initiation premiers secours","05/07/2026",15],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Formations");
    XLSX.writeFile(wb, "modele_formations.xlsx");
  };

  const handleFile = (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv","ods"].includes(ext)) {
      setError("Format non supporté. Utilisez .xlsx, .xls ou .csv");
      return;
    }
    setError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb    = XLSX.read(e.target.result, {type:"array", cellDates:false});
        const ws    = wb.Sheets[wb.SheetNames[0]];
        const data  = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if (data.length < 2) { setError("Le fichier semble vide."); return; }
        const hdrs  = data[0].map(h=>String(h));
        const rws   = data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs);
        setRows(rws);
        const detected = autoDetect(hdrs);
        setCols(detected);
        setStep("map");
      } catch(err) {
        setError("Impossible de lire le fichier : "+err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const buildPreview = () => {
    if (!cols.colName || !cols.colDate) { setError("Veuillez sélectionner au minimum la colonne Nom et la colonne Date."); return; }
    setError("");
    const iName     = headers.indexOf(cols.colName);
    const iDate     = headers.indexOf(cols.colDate);
    const iLearners = cols.colLearners ? headers.indexOf(cols.colLearners) : -1;
    const built = rows.map((r,i) => ({
      _idx: i,
      name:     String(r[iName]||"").trim(),
      date:     parseDate(r[iDate]),
      learners: iLearners>=0 ? (parseInt(r[iLearners])||10) : 10,
    })).filter(r=>r.name && r.date);
    if (built.length===0) { setError("Aucune ligne valide détectée. Vérifiez le mapping des colonnes."); return; }
    setPreview(built);
    setSelected(built.map(r=>r._idx));
    setStep("preview");
  };

  const toggleRow = (idx) => setSelected(prev => prev.includes(idx) ? prev.filter(i=>i!==idx) : [...prev,idx]);
  const toggleAll = () => setSelected(prev => prev.length===preview.length ? [] : preview.map(r=>r._idx));

  const doImport = () => {
    const toImport = preview.filter(r=>selected.includes(r._idx));
    if (toImport.length===0) return;
    onImport(toImport);
  };

  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal modal-wide" onClick={e=>e.stopPropagation()} style={{maxWidth:700}}>

        {/* Header */}
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24}}>
          <div style={{flex:1}}>
            <div style={{fontSize:17,fontWeight:700}}>Importer depuis Excel</div>
            <div style={{fontSize:12,color:"#6b7280",marginTop:2}}>
              Formats acceptés : .xlsx · .xls · .csv
            </div>
          </div>
          {/* Steps indicator */}
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            {[["upload","1. Fichier"],["map","2. Colonnes"],["preview","3. Aperçu"]].map(([s,l],i,arr)=>(
              <div key={s} style={{display:"flex",alignItems:"center",gap:6}}>
                <span style={{fontSize:11,fontWeight:600,color:step===s?"#f59e0b":step>s?"#22c55e":"#4b5563",padding:"3px 8px",borderRadius:4,background:step===s?"#f59e0b22":step>s?"#22c55e22":"transparent"}}>{l}</span>
                {i<arr.length-1&&<span style={{color:"#4b5563"}}>›</span>}
              </div>
            ))}
          </div>
        </div>

        {/* ── STEP 1: UPLOAD ── */}
        {step==="upload" && (
          <div>
            {/* Download template */}
            <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
              <button className="btn bb" style={{fontSize:12}} onClick={downloadTemplate}>⬇ Télécharger le modèle Excel</button>
            </div>

            {/* Drop zone */}
            <div
              onDragOver={e=>{e.preventDefault();setDragging(true);}}
              onDragLeave={()=>setDragging(false)}
              onDrop={e=>{e.preventDefault();setDragging(false);handleFile(e.dataTransfer.files[0]);}}
              onClick={()=>fileRef.current?.click()}
              style={{border:`2px dashed ${dragging?"#f59e0b":"#2a2d3a"}`,borderRadius:12,padding:"48px 24px",textAlign:"center",cursor:"pointer",transition:"all .2s",background:dragging?"#f59e0b08":"transparent"}}>
              <div style={{fontSize:36,marginBottom:12}}>📄</div>
              <div style={{fontWeight:600,marginBottom:6}}>Glissez votre fichier Excel ici</div>
              <div style={{fontSize:13,color:"#6b7280"}}>ou cliquez pour choisir un fichier</div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,.ods" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0])}/>
            </div>

            {/* Template hint */}
            <div style={{marginTop:16,padding:"12px 16px",background:"#1e2230",borderRadius:8,fontSize:12,color:"#9ca3af"}}>
              <div style={{fontWeight:600,color:"#e8e8e8",marginBottom:4}}>Format attendu</div>
              <div>Votre fichier doit contenir au minimum une colonne <span style={{color:"#f59e0b"}}>Nom de la formation</span> et une colonne <span style={{color:"#f59e0b"}}>Date</span>. Une colonne <span style={{color:"#f59e0b"}}>Nombre d'apprenants</span> est optionnelle.</div>
              <div style={{marginTop:6}}>Dates acceptées : <span style={{fontFamily:"'DM Mono',monospace"}}>14/06/2026 · 2026-06-14 · 14-06-2026</span></div>
            </div>

            {error&&<div style={{marginTop:12,color:"#ef4444",fontSize:13}}>⚠ {error}</div>}
          </div>
        )}

        {/* ── STEP 2: MAP COLUMNS ── */}
        {step==="map" && (
          <div>
            <div style={{marginBottom:16,fontSize:13,color:"#9ca3af"}}>
              Fichier chargé — <span style={{color:"#22c55e",fontWeight:600}}>{rows.length} lignes</span> détectées. Vérifiez que les bonnes colonnes sont sélectionnées.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,marginBottom:20}}>
              {[
                {key:"colName",    label:"Nom de la formation", required:true,  color:"#f59e0b"},
                {key:"colDate",    label:"Date",                required:true,  color:"#f59e0b"},
                {key:"colLearners",label:"Nombre d'apprenants", required:false, color:"#60a5fa"},
              ].map(({key,label,required,color})=>(
                <div key={key}>
                  <div style={{fontSize:12,marginBottom:6}}>
                    <span style={{color:color,fontWeight:600}}>{label}</span>
                    {required&&<span style={{color:"#ef4444",marginLeft:4}}>*</span>}
                  </div>
                  <select className="inp" value={cols[key]} onChange={e=>setCols(p=>({...p,[key]:e.target.value}))}>
                    <option value="">— Non utilisé —</option>
                    {headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              ))}
            </div>

            {/* Data preview (first 3 rows) */}
            <div style={{marginBottom:16}}>
              <div className="sec">Aperçu des données (3 premières lignes)</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{borderBottom:"1px solid #2a2d3a"}}>
                      {headers.map(h=>(
                        <th key={h} style={{padding:"7px 10px",textAlign:"left",color:Object.values(cols).includes(h)?"#f59e0b":"#6b7280",fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.slice(0,3).map((r,i)=>(
                      <tr key={i} style={{borderBottom:"1px solid #1e2230"}}>
                        {r.map((c,j)=><td key={j} style={{padding:"6px 10px",color:"#9ca3af"}}>{String(c)}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {error&&<div style={{color:"#ef4444",fontSize:13,marginBottom:12}}>⚠ {error}</div>}
            <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
              <button className="btn bg" onClick={()=>setStep("upload")}>← Retour</button>
              <button className="btn bp" onClick={buildPreview}>Aperçu →</button>
            </div>
          </div>
        )}

        {/* ── STEP 3: PREVIEW & CONFIRM ── */}
        {step==="preview" && (
          <div>
            <div style={{marginBottom:12,display:"flex",alignItems:"center",gap:12}}>
              <div style={{fontSize:13,color:"#9ca3af",flex:1}}>
                <span style={{color:"#22c55e",fontWeight:600}}>{selected.length}</span> / {preview.length} formations sélectionnées à importer
              </div>
              <button className="btn bg" style={{fontSize:11,padding:"4px 10px"}} onClick={toggleAll}>
                {selected.length===preview.length?"Tout décocher":"Tout cocher"}
              </button>
            </div>

            <div style={{maxHeight:320,overflowY:"auto",display:"flex",flexDirection:"column",gap:6,marginBottom:16}}>
              {preview.map(r=>{
                const sel=selected.includes(r._idx);
                const valid=r.name&&r.date;
                return(
                  <div key={r._idx} onClick={()=>valid&&toggleRow(r._idx)}
                    style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:8,border:`1px solid ${sel?"#f59e0b44":"#2a2d3a"}`,background:sel?"#f59e0b08":"#1e2230",cursor:valid?"pointer":"default",opacity:valid?1:0.5}}>
                    <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${sel?"#f59e0b":"#4b5563"}`,background:sel?"#f59e0b":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {sel&&<span style={{color:"#0f1117",fontSize:12,fontWeight:800}}>✓</span>}
                    </div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:13,fontWeight:600,color:valid?"#e8e8e8":"#ef4444"}}>{r.name||"⚠ Nom manquant"}</div>
                      <div style={{fontSize:11,color:"#6b7280",marginTop:2}}>
                        {r.date ? fmtDate(r.date) : <span style={{color:"#ef4444"}}>⚠ Date invalide</span>}
                        {" · "}{r.learners} apprenant{r.learners>1?"s":""}
                      </div>
                    </div>
                    {!valid&&<span style={{fontSize:11,color:"#ef4444"}}>Ligne ignorée</span>}
                  </div>
                );
              })}
            </div>

            <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
              <button className="btn bg" onClick={()=>setStep("map")}>← Retour</button>
              <button className="btn bp" onClick={doImport} disabled={selected.length===0}>
                ⬆ Importer {selected.length} formation{selected.length>1?"s":""}
              </button>
            </div>
          </div>
        )}

        {step!=="upload"&&(
          <div style={{display:"flex",justifyContent:"flex-end",marginTop:16,borderTop:"1px solid #2a2d3a",paddingTop:16}}>
            <button className="btn bg" onClick={onClose}>Annuler</button>
          </div>
        )}
      </div>
    </div>
  );
}
